using Microsoft.Win32;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media.Imaging;
using Tesseract;
using ClosedXML.Excel;
using DuckDB.NET.Data;

namespace AccountingOcrTest
{
    public partial class MainWindow : Window
    {
        private ApiKeyManager _apiKeyManager;
        private ObservableCollection<ProcessingItem> _items = new ObservableCollection<ProcessingItem>();
        private ICollectionView _itemsView;
        private CancellationTokenSource _cts;
        
        private static readonly string ShippersFilePath = "shippers.json";
        private ObservableCollection<string> _shippers = new ObservableCollection<string>();

        // Reference excel fields
        private XLWorkbook? _refWorkbook;
        private ObservableCollection<ReferenceItem> _refItems = new ObservableCollection<ReferenceItem>();

        // Configs and internal copy fields
        private ObservableCollection<ExcelConfigItem> _excelConfigs = new ObservableCollection<ExcelConfigItem>();
        private static readonly string ConfigsFilePath = "excel_configs.json";
        private static readonly string RawMaterialDirName = "raw_material";

        // DuckDB staging variables
        private static readonly string StagingDbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "accounting_staging.db");
        private ObservableCollection<StagingItem> _stagingItems = new ObservableCollection<StagingItem>();
        private static readonly object DbWriteLock = new object();

        // Fix #3: Static HttpClient — avoids socket exhaustion and DNS caching issues
        private static readonly HttpClient _httpClient = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(60)
        };

        // Cached JsonSerializerOptions — avoids rebuilding internal cache every API call
        private static readonly JsonSerializerOptions _jsonOptions = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        };

        // Fix #4: ThreadLocal TesseractEngine cache — each thread gets its own instance
        // (TesseractEngine is NOT thread-safe, so ThreadLocal is the correct approach)
        private static readonly string _tessDataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");
        private static readonly ThreadLocal<TesseractEngine> _tessEngine =
            new ThreadLocal<TesseractEngine>(() =>
            {
                try { return new TesseractEngine(_tessDataPath, "osd", EngineMode.Default); }
                catch { return null; }
            });

        public MainWindow()
        {
            InitializeComponent();
            this.Closing += MainWindow_Closing;
            _apiKeyManager = new ApiKeyManager();
            ItemsListBox.ItemsSource = _items;
            _itemsView = CollectionViewSource.GetDefaultView(_items);
            _itemsView.Filter = SearchFilter;
            
            LoadShippers();
            ShipperSelector.ItemsSource = _shippers;
            
            RefDataGrid.ItemsSource = _refItems;

            // Setup local raw_material dir
            try
            {
                string rawPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, RawMaterialDirName);
                if (!Directory.Exists(rawPath))
                {
                    Directory.CreateDirectory(rawPath);
                }
            }
            catch { }

            // Bind homepage config list and load
            HomeConfigsGrid.ItemsSource = _excelConfigs;
            LoadExcelConfigs();

            // Setup DuckDB Staging Database
            InitStagingDatabase();
            LoadStagingItemsFromDb();
        }

        private bool SearchFilter(object item)
        {
            if (string.IsNullOrWhiteSpace(SearchBox.Text) || SearchBox.Text == "🔍 Tìm kiếm...")
                return true;
            
            var pItem = item as ProcessingItem;
            if (pItem == null) return false;

            return pItem.SmartName.Contains(SearchBox.Text, StringComparison.OrdinalIgnoreCase);
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_itemsView != null)
                _itemsView.Refresh();
        }

        private void SettingsBtn_Click(object sender, RoutedEventArgs e)
        {
            var win = new ApiKeyManagerWindow(_apiKeyManager) { Owner = this };
            win.ShowDialog();
        }

        private void AddImagesBtn_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(ShipperSelector.Text))
            {
                MessageBox.Show("Vui lòng nhập hoặc chọn Tên người giao hàng trước khi thêm ảnh.", "Bắt buộc", MessageBoxButton.OK, MessageBoxImage.Warning);
                ShipperSelector.Focus();
                return;
            }

            string currentShipper = ShipperSelector.Text.Trim();

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Image Files|*.jpg;*.jpeg;*.png;*.webp",
                Title = "Chọn ảnh hóa đơn",
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                ContinueBtn.Visibility = Visibility.Collapsed;
                foreach (var file in openFileDialog.FileNames)
                {
                    _items.Add(new ProcessingItem 
                    { 
                        FilePath = file, 
                        SmartName = $"[{currentShipper}] - {System.IO.Path.GetFileName(file)}",
                        Status = ProcessStatus.Waiting
                    });
                }
            }
        }

        // Fix #1: Async image loading with DecodePixelWidth to prevent UI freeze
        private async void ItemsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ItemsListBox.SelectedItem is ProcessingItem item)
            {
                // Update Preview Image — decode on background thread with size limit
                try
                {
                    var bitmap = await Task.Run(() =>
                    {
                        var bi = new BitmapImage();
                        bi.BeginInit();
                        bi.UriSource = new Uri(item.FilePath);
                        bi.CacheOption = BitmapCacheOption.OnLoad;
                        bi.DecodePixelWidth = 800; // Only decode to preview size, not full resolution
                        bi.EndInit();
                        bi.Freeze(); // Required to pass BitmapImage across threads
                        return bi;
                    });

                    // Verify this item is still selected after the async load completed
                    if (ItemsListBox.SelectedItem == item)
                    {
                        PreviewImage.Source = bitmap;
                    }
                }
                catch { PreviewImage.Source = null; }
            }
            else
            {
                PreviewImage.Source = null;
            }
        }

        private async void StartBtn_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(ShipperSelector.Text))
            {
                MessageBox.Show("Vui lòng nhập hoặc chọn Tên người giao hàng trước khi bắt đầu.", "Bắt buộc", MessageBoxButton.OK, MessageBoxImage.Warning);
                ShipperSelector.Focus();
                return;
            }

            string currentShipper = ShipperSelector.Text.Trim();
            if (!_shippers.Contains(currentShipper))
            {
                _shippers.Add(currentShipper);
                SaveShippers();
            }

            if (_apiKeyManager.Keys.Count == 0)
            {
                MessageBox.Show("Vui lòng cài đặt API Key trước khi bắt đầu.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var waitingItems = _items.Where(x => x.Status == ProcessStatus.Waiting).ToList();
            if (waitingItems.Count == 0)
            {
                MessageBox.Show("Không có ảnh nào đang chờ xử lý.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // UI state: processing mode
            StartBtn.IsEnabled = false;
            StartBtn.Content = "⏳ Đang xử lý...";
            CancelBtn.Visibility = Visibility.Visible;
            CancelBtn.IsEnabled = true;
            CancelBtn.Content = "⏹ Hủy";
            _cts?.Dispose();
            _cts = new CancellationTokenSource();

            string selectedModel = "gemini-3.5-flash";
            if (ModelSelector.SelectedItem is ComboBoxItem comboItem)
            {
                selectedModel = comboItem.Content.ToString();
            }

            int concurrencyLevel = Math.Min(waitingItems.Count, _apiKeyManager.Keys.Count * 2);
            if (concurrencyLevel == 0) concurrencyLevel = 1; // Fallback nếu không có key

            int total = waitingItems.Count;
            int completed = 0;

            // Fix #10: Show progress indicator
            ProgressBar.Visibility = Visibility.Visible;
            ProgressText.Visibility = Visibility.Visible;
            ProgressBar.Maximum = total;
            ProgressBar.Value = 0;
            ProgressText.Text = $"0/{total}";

            Logger.Log($"[Bắt đầu] Đang xử lý {total} hóa đơn bằng {selectedModel} với {concurrencyLevel} luồng...");

            var tasks = new List<Task>();
            var queue = new ConcurrentQueue<ProcessingItem>(waitingItems);

            for (int i = 0; i < concurrencyLevel; i++)
            {
                tasks.Add(Task.Run(async () =>
                {
                    while (queue.TryDequeue(out var item))
                    {
                        if (_cts.Token.IsCancellationRequested) break;

                        // Fix #2: All property changes dispatched to UI thread
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            item.Status = ProcessStatus.Processing;
                        });
                        Logger.Log($"[Tiến trình] Bắt đầu xử lý file: {item.FilePath}");
                        
                        try
                        {
                            int maxRetries = 3;
                            int retryCount = 0;
                            bool success = false;
                            string lastErrorMessage = "";

                            while (retryCount < maxRetries && !success && !_cts.Token.IsCancellationRequested)
                            {
                                string apiKey = await _apiKeyManager.GetNextAvailableKeyAsync(_cts.Token);
                                try
                                {
                                    await ProcessImageAsync(item, apiKey, selectedModel, currentShipper, _cts.Token);
                                    success = true;
                                }
                                catch (Exception ex) when (!_cts.Token.IsCancellationRequested)
                                {
                                    retryCount++;
                                    lastErrorMessage = ex.Message;
                                    
                                    string errText = ex.Message.ToLower();
                                    if (errText.Contains("429") || errText.Contains("toomanyrequests") || errText.Contains("quota"))
                                    {
                                        _apiKeyManager.MarkKeyAsRateLimited(apiKey);
                                        Logger.Log($"[RateLimit] Key {apiKey.Substring(0, 5)}... gặp lỗi giới hạn lượt dùng. Thử lại lần {retryCount}/{maxRetries} với key khác...");
                                    }
                                    else if (errText.Contains("invalid") || errText.Contains("403") || errText.Contains("bad request") || errText.Contains("400"))
                                    {
                                        _apiKeyManager.MarkKeyAsError(apiKey);
                                        Logger.Log($"[ApiKeyError] Key {apiKey.Substring(0, 5)}... gặp lỗi xác thực/cú pháp. Thử lại lần {retryCount}/{maxRetries} với key khác...");
                                    }
                                    else
                                    {
                                        _apiKeyManager.MarkKeyAsError(apiKey);
                                        Logger.Log($"[APIError] Key {apiKey.Substring(0, 5)}... gặp lỗi kết nối: {ex.Message}. Thử lại lần {retryCount}/{maxRetries} với key khác...");
                                    }

                                    if (retryCount < maxRetries)
                                    {
                                        await Task.Delay(1500, _cts.Token);
                                    }
                                }
                            }

                            if (!success)
                            {
                                throw new Exception(lastErrorMessage);
                            }

                            int current = Interlocked.Increment(ref completed);

                            // Fix #2: Update all properties on UI thread
                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                item.Status = ProcessStatus.Success;
                                if (item.Data != null)
                                    item.SmartName = $"[{item.Data.ten_nguoi_giao}] - {item.Data.ngay_giao} - {item.Data.khach_hang} - {item.Data.diem_giao}";

                                // Update progress
                                ProgressBar.Value = current;
                                ProgressText.Text = $"{current}/{total}";
                                // WPF data binding automatically updates the UI
                            });

                            Logger.Log($"[Thành công] Đã trích xuất xong: {item.FilePath} -> {item.SmartName}");
                        }
                        catch (OperationCanceledException)
                        {
                            // Reset item back to Waiting so user can re-run
                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                item.Status = ProcessStatus.Waiting;
                            });
                            Logger.Log($"[Hủy] Đã hủy xử lý file: {item.FilePath}");
                            break;
                        }
                        catch (Exception ex)
                        {
                            int current = Interlocked.Increment(ref completed);

                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                item.Status = ProcessStatus.Error;
                                item.ErrorMessage = ex.Message;
                                ProgressBar.Value = current;
                                ProgressText.Text = $"{current}/{total}";
                            });
                            Logger.Log($"[Lỗi] Lỗi khi xử lý file {item.FilePath}: {ex.ToString()}");
                        }
                    }
                }));
            }

            await Task.WhenAll(tasks);

            // Flush buffered logs before showing completion message
            Logger.Flush();

            // UI state: done
            StartBtn.IsEnabled = true;
            StartBtn.Content = "▶ Bắt đầu xử lý";
            CancelBtn.Visibility = Visibility.Collapsed;
            ProgressBar.Visibility = Visibility.Collapsed;
            ProgressText.Visibility = Visibility.Collapsed;

            if (_cts.Token.IsCancellationRequested)
            {
                MessageBox.Show("Đã hủy xử lý!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                if (_items.Any(x => x.Status == ProcessStatus.Success))
                {
                    ContinueBtn.Visibility = Visibility.Visible;
                }
                MessageBox.Show("Đã hoàn tất xử lý danh sách hóa đơn!", "Thành công", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        // Fix #6: Cancel button handler
        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            _cts?.Cancel();
            CancelBtn.IsEnabled = false;
            CancelBtn.Content = "⏳ Đang hủy...";
        }

        // Fix #7: Resize large images before sending to API (saves bandwidth + latency)
        private static Bitmap ResizeImageForApi(Bitmap original, int maxWidth = 2000, int maxHeight = 2000)
        {
            double ratioX = (double)maxWidth / original.Width;
            double ratioY = (double)maxHeight / original.Height;
            double ratio = Math.Min(ratioX, ratioY);

            int newWidth = (int)(original.Width * ratio);
            int newHeight = (int)(original.Height * ratio);

            var resized = new Bitmap(newWidth, newHeight);
            using (var g = Graphics.FromImage(resized))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                g.DrawImage(original, 0, 0, newWidth, newHeight);
            }
            return resized;
        }

        private static void RotateImageByExif(System.Drawing.Image img)
        {
            try
            {
                if (img.PropertyIdList.Contains(0x0112)) // ExifPropertyTagOrientation
                {
                    var prop = img.GetPropertyItem(0x0112);
                    if (prop != null && prop.Value != null && prop.Value.Length > 0)
                    {
                        int orientation = prop.Value[0];
                        RotateFlipType flip = RotateFlipType.RotateNoneFlipNone;

                        switch (orientation)
                        {
                            case 1:
                                flip = RotateFlipType.RotateNoneFlipNone;
                                break;
                            case 2:
                                flip = RotateFlipType.RotateNoneFlipX;
                                break;
                            case 3:
                                flip = RotateFlipType.Rotate180FlipNone;
                                break;
                            case 4:
                                flip = RotateFlipType.Rotate180FlipX;
                                break;
                            case 5:
                                flip = RotateFlipType.Rotate90FlipX;
                                break;
                            case 6:
                                flip = RotateFlipType.Rotate90FlipNone;
                                break;
                            case 7:
                                flip = RotateFlipType.Rotate270FlipX;
                                break;
                            case 8:
                                flip = RotateFlipType.Rotate270FlipNone;
                                break;
                        }

                        if (flip != RotateFlipType.RotateNoneFlipNone)
                        {
                            img.RotateFlip(flip);
                            try
                            {
                                img.RemovePropertyItem(0x0112);
                            }
                            catch { }
                            Logger.Log($"[EXIF] Đã tự động xoay ảnh theo thẻ EXIF: {flip}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[EXIF Lỗi] Không thể xoay ảnh theo EXIF: {ex.Message}");
            }
        }

        private async Task ProcessImageAsync(ProcessingItem item, string apiKey, string modelName, string shipperName, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            string base64Image = "";
            string mimeType = "image/jpeg";
            if (item.FilePath.ToLower().EndsWith(".png")) mimeType = "image/png";
            if (item.FilePath.ToLower().EndsWith(".webp")) mimeType = "image/webp";

            using (var img = new Bitmap(item.FilePath))
            {
                // Step 1: Rotate by EXIF tag first (extremely common for phone photos)
                RotateImageByExif(img);

                // Step 2: Use cached TesseractEngine to detect any remaining physical rotation
                try 
                {
                    var engine = _tessEngine.Value;
                    if (engine != null)
                    {
                        using (var ms = new MemoryStream())
                        {
                            img.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                            using (var pix = Pix.LoadFromMemory(ms.ToArray()))
                            {
                                using (var page = engine.Process(pix, PageSegMode.OsdOnly))
                                {
                                    var iterator = page.GetIterator();
                                    if (iterator != null)
                                    {
                                        iterator.Begin();
                                        var props = iterator.GetProperties();
                                        Tesseract.Orientation orientation = props.Orientation;
                                        Logger.Log($"[Tesseract OSD] Hướng ảnh phát hiện: {orientation}");

                                        if (orientation == Tesseract.Orientation.PageRight) 
                                        {
                                            img.RotateFlip(RotateFlipType.Rotate270FlipNone);
                                            Logger.Log("[Tesseract OSD] Đã xoay ảnh 270 độ.");
                                        }
                                        else if (orientation == Tesseract.Orientation.PageDown) 
                                        {
                                            img.RotateFlip(RotateFlipType.Rotate180FlipNone);
                                            Logger.Log("[Tesseract OSD] Đã xoay ảnh 180 độ (lộn ngược).");
                                        }
                                        else if (orientation == Tesseract.Orientation.PageLeft) 
                                        {
                                            img.RotateFlip(RotateFlipType.Rotate90FlipNone);
                                            Logger.Log("[Tesseract OSD] Đã xoay ảnh 90 độ.");
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"[Tesseract Lỗi] Không thể chạy OSD: {ex.Message}");
                }

                token.ThrowIfCancellationRequested();

                // Fix #7: Resize large images before encoding for API
                bool needsResize = img.Width > 2000 || img.Height > 2000;
                Bitmap apiImage = needsResize ? ResizeImageForApi(img) : img;

                try
                {
                    using (var ms = new MemoryStream())
                    {
                        var format = mimeType == "image/png" ? System.Drawing.Imaging.ImageFormat.Png : System.Drawing.Imaging.ImageFormat.Jpeg;
                        apiImage.Save(ms, format);
                        // Use GetBuffer() + Length to avoid allocating a copy array on LOH
                        base64Image = Convert.ToBase64String(ms.GetBuffer(), 0, (int)ms.Length);
                    }
                }
                finally
                {
                    // Only dispose the resized copy, not the original (which is managed by the outer using)
                    if (needsResize) apiImage.Dispose();
                }
            }

            item.Data = await CallGeminiApi(base64Image, mimeType, apiKey, modelName, token);
            if (item.Data != null)
            {
                item.Data.ten_nguoi_giao = shipperName;
                item.Data.khach_hang = FuzzyMatcher.FormatKhachHang(item.Data.khach_hang);
                item.Data.diem_giao = FuzzyMatcher.FormatDiemGiao(item.Data.diem_giao);

                if (item.Data.danh_sach_hang_hoa != null)
                {
                    // Remove redundant printed items that have sl_xuat = 0, sl_nhan = 0, and sl_hong = 0
                    item.Data.danh_sach_hang_hoa.RemoveAll(x => x.sl_xuat == 0 && x.sl_nhan == 0 && x.sl_hong == 0);
                }
            }
        }

        private async Task<InvoiceData> CallGeminiApi(string base64Image, string mimeType, string apiKey, string modelName, CancellationToken token)
        {
            string promptText = @"Bạn là một hệ thống OCR và kiểm toán kế toán cao cấp chuyên nghiệp của Việt Nam. Nhiệm vụ của bạn là số hóa hóa đơn này với ĐỘ CHÍNH XÁC TUYỆT ĐỐI về mặt cấu trúc và nội dung. ĐẶC BIỆT CHÚ Ý: Hình ảnh có thể bị xoay ngang, xoay dọc, hoặc lộn ngược. Hãy tự động 'xoay' và định hướng lại hình ảnh trong tư duy trước khi đọc. TUYỆT ĐỐI KHÔNG ĐƯỢC BỎ SÓT BẤT KỲ DÒNG HÀNG HÓA NÀO. Hãy dóng hàng ngang thật cẩn thận để đảm bảo Tên hàng hóa và Số lượng khớp nhau 100%, không bị trôi dòng, lệch dòng hay mất dòng.

HÃY THỰC HIỆN QUY TRÌNH SUY LUẬN (CHAIN OF THOUGHT):
BẮT BUỘC ghi toàn bộ quá trình phân tích của bạn vào trường `quy_trinh_suy_luan` trong file JSON. Đối với phần hàng hóa, hãy phân tích TỪNG DÒNG một cách tường minh, ví dụ: 'Dòng X: bị một đường kẻ dọc/chéo dài cắt qua cột Đơn giá/ĐVT -> Hủy -> sl_nhan = 0', 'Dòng Y: có nét viết tay số 2 -> sl_nhan = 2'. NẾU CÓ MỘT ĐƯỜNG KẺ DÀI KÉO TỪ DÒNG TRÊN XUỐNG DÒNG DƯỚI, BẠN PHẢI GHI NHẬN LÀ TẤT CẢ CÁC DÒNG ĐÓ ĐỀU BỊ GẠCH BỎ.

Bước 1: Phân tích Cấu trúc Bảng và Thông tin chung
- Ngày giao (ngay_giao): Tìm ngày giao/nhận hàng thực tế (Delivery Date / Date Received). TUYỆT ĐỐI BỎ QUA 'Ngày In' (Printed Date / Ngày In / Printed Time, ví dụ: 2026-03-24) và 'Ngày Đặt Hàng' (Order Date).
  BẮT BUỘC TÌM KIẾM THEO ĐÚNG MỨC ĐỘ ƯU TIÊN SAU:
  1. ƯU TIÊN CAO NHẤT: Ngày ghi viết tay trên con dấu (ví dụ: ngày ghi trên con dấu nhận hàng).
  2. ƯU TIÊN 2: Ngày ghi viết tay bằng bút nằm rải rác (gần chữ ký, khu vực ghi chú...).
  3. ƯU TIÊN 3: Ngày được in sẵn ở ô/cột Ngày Giao Hàng hoặc Ngày Nhận Hàng (Delivery Date / Date Received). Ví dụ cột 'DATE RECEIVED' ghi '2026 03 09'.
  
  CẢNH BÁO VỀ SỬA NGÀY (CHO ƯU TIÊN 1 & 2): Đôi khi con dấu đóng ngày (Ưu tiên 1) hoặc ngày viết tay (Ưu tiên 2) bị ghi nhầm theo Ngày In (ví dụ: đóng dấu nhầm ngày '24-03-2026'). Nhân viên nhận hàng phát hiện ra đã dùng bút mực viết tay ghi đè/sửa trực tiếp lên trên con dấu hoặc bên cạnh (ví dụ: khoanh tròn số '24' rồi kéo nét chéo ghi đè số '09' lên trên đầu con dấu, hoặc gạch bỏ số '24' đóng dấu nhầm rồi viết đè số '09').
  Khi thấy có bất kỳ ký hiệu sửa đổi viết tay nào đè lên con dấu hoặc đè lên ngày cũ, bạn BẮT BUỘC phải lấy giá trị đã được sửa bằng tay đó làm ngày nhận thực tế (Ví dụ ở đây số '24' đóng dấu nhầm đã bị gạch/sửa bằng bút viết tay thành '09', kết hợp với ngày nhận hàng in sẵn máy là '03 09' ở Ưu tiên 3 -> Kết luận ngày giao thực tế phải là '09/03/2026').
  
  Trả về định dạng ngày giao chuẩn 'dd/MM/yyyy'. Ví dụ: '09/03/2026'.
- Khách hàng (khach_hang): Tìm tên thương hiệu/tên thương mại chính của khách hàng (ví dụ: Aeon, Winmart, Bigc, B11, Biggreen...). BẮT BUỘC bỏ qua toàn bộ phần tiền tố/hậu tố pháp lý rườm rà như 'CÔNG TY TNHH', 'CÔNG TY CỔ PHẦN', 'CHI NHÁNH', 'MỘT THÀNH VIÊN', 'VIỆT NAM'... Ví dụ: nếu hóa đơn ghi 'CÔNG TY TNHH AEON VIỆT NAM' thì chỉ trích xuất 'Aeon'. Nếu ghi 'CÔNG TY CP THƯƠNG MẠI WINCOMMERCE' thì chỉ trích xuất 'Winmart'.
- Điểm giao (diem_giao): Tên chi nhánh của khách hàng. ĐẶC BIỆT CHÚ Ý: Nếu trên hóa đơn có ghi MÃ CỬA HÀNG đi kèm TÊN CỬA HÀNG (ví dụ như dòng '1708 - WM HNI Lê Văn Thiêm'), bạn PHẢI trích xuất NGUYÊN VẸN toàn bộ chuỗi đó làm điểm giao (tức là lấy đầy đủ cả mã và tên: '1708 - WM HNI Lê Văn Thiêm'). Nếu không có mã cửa hàng, thì lấy tên chi nhánh ngắn gọn như bình thường (ví dụ: 'Xuân Thủy', 'Ciputra'). Thường thông tin này nằm ở phần vị trí/địa chỉ điểm giao.

Bước 2: Phân tích Thị giác Chuyên sâu và Nhận biết Ký hiệu Viết tay (Ink-to-Text & Symbol Analysis) - QUAN TRỌNG NHẤT
- Phân tách rõ ràng giữa mực in máy và mực viết tay (mực màu xanh, đen mờ, đỏ hoặc nét bút viết tay).
- MỖI DÒNG HÀNG HÓA PHẢI ĐƯỢC PHÂN TÍCH ĐỘC LẬP. Không được suy đoán kết quả dòng này dựa trên dòng khác. Phải có BẰNG CHỨNG THỊ GIÁC CỤ THỂ cho từng dòng.
- Quét qua từng dòng dữ liệu hàng hóa và áp dụng QUY TẮC theo THỨ TỰ ƯU TIÊN SAU (rule trên THẮNG rule dưới):

  ƯU TIÊN CAO NHẤT - Rule A: Dấu tích '✓' kèm con số viết tay rõ ràng (ví dụ: '✓ 05', '✓ 5', '✓ 10'...):
     -> ĐÂY LÀ KÝ HIỆU XÁC NHẬN GIAO THỰC TẾ. -> sl_nhan = <con số viết tay đó>.
     -> Rule này LUÔN THẮNG mọi rule khác. Dù dòng đó có bị đường kẻ đi qua, NẾU CÓ ✓ kèm số thì vẫn lấy số đó làm sl_nhan.

  ƯU TIÊN 2 - Rule B: Đường gạch bỏ / hủy giao. BẤT KỲ đường chữ 'Z', nét gạch chéo 'X', nét gạch ngang '-', nét gạch chéo '/', nét kẻ dọc '|', nét mũi tên (->), hoặc đường gạch tay chéo dài. Nét gạch CÓ THỂ vắt qua cột ĐVT, Đơn giá, hoặc Thành tiền. Chỉ cần TRÊN DÒNG ĐÓ bị nét mực gạch xuyên qua VÀ KHÔNG CÓ dấu ✓ kèm số:
     -> sl_nhan = 0.
     CẢNH BÁO CỰC KỲ QUAN TRỌNG VỀ PHẠM VI ĐƯỜNG KẺ: Khi thấy một đường gạch chéo dài, bạn phải XÁC NHẬN CHÍNH XÁC BẰNG THỊ GIÁC rằng đường đó THỰC SỰ ĐI QUA dòng nào. TUYỆT ĐỐI KHÔNG ĐƯỢC SUY ĐOÁN hoặc NGOẠI SUY rằng đường kẻ kéo dài hơn thực tế. Ví dụ: nếu đường kẻ kết thúc tại dòng 100 thì KHÔNG ĐƯỢC gán sl_nhan=0 cho dòng 110, 120, 130. Chỉ những dòng mà mắt thường nhìn thấy có nét mực ĐI XUYÊN QUA mới được đánh dấu hủy.
     ĐẶC BIỆT CHÚ Ý VỚI KÝ HIỆU HỦY VIẾT TAY: Các đường vẽ tay có hình dạng gạch ngang ngắn '-', gạch lượn sóng '~', nét gạch ngang mờ hoặc nét viết tay vẽ ngang/chéo nằm ở cột checkbox/kiểm nhận của dòng (mà không đi kèm số hay dấu tick nào) đều được tính là gạch hủy dòng đó -> sl_nhan = 0.
     LƯU Ý VỀ VỊ TRÍ NÉT MỰC: Nét gạch hủy (đường gạch ngang '-', gạch sóng '~', nét gạch chéo '/') thường chỉ được ký hiệu ngắn gọn ở cột 'Số lượng' hoặc phần khoảng trống kiểm hàng bên trái số lượng, chứ không nhất thiết phải gạch ngang qua toàn bộ tên mặt hàng hay đơn giá. Chỉ cần có bất kỳ nét gạch ngắn, nét gạch chéo hoặc ký hiệu hủy viết tay nào tương tự xuất hiện trên dòng đó (kể cả chỉ ở cột Số lượng hay cột Stt), và không có dấu tick xác nhận nhận hàng, thì bắt buộc phải ghi nhận dòng đó đã bị hủy -> sl_nhan = 0.

  ƯU TIÊN 3 - Rule C (MẶC ĐỊNH): Không có bất kỳ nét bút viết tay nào (chỉ có mực in máy):
     -> Giao đầy đủ. -> sl_nhan = sl_xuat.
     LƯU Ý: Đây là TRƯỜNG HỢP MẶC ĐỊNH. Nếu một dòng không bị gạch bỏ và không có ✓, thì PHẢI giữ nguyên sl_nhan = sl_xuat. KHÔNG ĐƯỢC gán sl_nhan = 0 trừ khi có BẰNG CHỨNG THỊ GIÁC rõ ràng của nét gạch bỏ TRÊN CHÍNH DÒNG ĐÓ.

  Rule D: Nếu số lượng viết tay hoặc in sẵn CÓ DẤU TRỪ (số âm, ví dụ: -2, -5...), điều này có nghĩa là 'nợ hàng'. Ghi nhận CHÍNH XÁC giá trị âm đó. Được phép ghi nhận số âm.

Bước 3: Đối soát tính toán (Tự động tính Số lượng hỏng)
- Với mỗi mặt hàng, bạn phải có sl_xuat (Số lượng in trên hóa đơn) và sl_nhan (Số lượng giao thực tế, tính toán từ Bước 2).
- Tự động tính toán: sl_hong = sl_xuat - sl_nhan. Trả về đúng giá trị toán học này. Chú ý có thể sl_hong > 0 hoặc = 0 hoặc đôi khi sl_nhan lớn hơn dẫn tới âm. Đảm bảo tính toán chính xác.
- LƯU Ý QUAN TRỌNG VỀ HÓA ĐƠN CHỈ CÓ MỘT CỘT SỐ LƯỢNG: Nếu hóa đơn chỉ in một cột số lượng duy nhất (ví dụ: chỉ có một cột Số lượng / Quantity mà không chia riêng cột xuất và cột nhận), thì:
  + Nếu không có ghi chú viết tay hay ký hiệu hủy/chỉnh sửa nào bên cạnh, mặc định coi như giao đủ: sl_xuat = sl_nhan = giá trị số lượng in máy đó (sl_hong = 0).
  + Nếu có nét gạch hủy dòng đó thì sl_nhan = 0 và sl_xuat = giá trị in máy.
  + Nếu có số lượng viết tay điều chỉnh bên cạnh thì sl_nhan = số viết tay đó và sl_xuat = giá trị in máy.";

            var requestBody = new
            {
                contents = new[]
                {
                    new
                    {
                        parts = new object[]
                        {
                            new { text = promptText },
                            new { inlineData = new { mimeType = mimeType, data = base64Image } }
                        }
                    }
                },
                generationConfig = new
                {
                    responseMimeType = "application/json",
                    responseSchema = new
                    {
                        type = "OBJECT",
                        properties = new Dictionary<string, object>
                        {
                            { "quy_trinh_suy_luan", new { type = "STRING", description = "Trình bày chi tiết phân tích từng dòng hàng hóa xem có nét mực gạch bỏ/chỉnh sửa hay không." } },
                            { "ngay_giao", new { type = "STRING" } },
                            { "khach_hang", new { type = "STRING" } },
                            { "diem_giao", new { type = "STRING" } },
                            { "danh_sach_hang_hoa", new {
                                type = "ARRAY",
                                items = new {
                                    type = "OBJECT",
                                    properties = new Dictionary<string, object>
                                    {
                                        { "ten_hang", new { type = "STRING" } },
                                        { "sl_xuat", new { type = "NUMBER" } },
                                        { "sl_nhan", new { type = "NUMBER" } },
                                        { "sl_hong", new { type = "NUMBER" } }
                                    },
                                    required = new[] { "ten_hang", "sl_xuat", "sl_nhan", "sl_hong" }
                                }
                            }}
                        },
                        required = new[] { "quy_trinh_suy_luan", "ngay_giao", "khach_hang", "diem_giao", "danh_sach_hang_hoa" }
                    }
                }
            };

            string jsonBody = JsonSerializer.Serialize(requestBody);

            // Fix #3: Use static HttpClient + pass CancellationToken
            string url = $"https://generativelanguage.googleapis.com/v1beta/models/{modelName}:generateContent?key={apiKey}";
            
            using var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");
            using var response = await _httpClient.PostAsync(url, content, token);
            
            string responseStr = await response.Content.ReadAsStringAsync(token);

            if (!response.IsSuccessStatusCode)
            {
                Logger.Log($"[Gemini API Lỗi] HTTP {response.StatusCode}: {responseStr}");
                throw new Exception($"Gemini API trả về lỗi {response.StatusCode}");
            }

            using (JsonDocument doc = JsonDocument.Parse(responseStr))
            {
                var root = doc.RootElement;
                var candidates = root.GetProperty("candidates");
                if (candidates.GetArrayLength() > 0)
                {
                    var text = candidates[0].GetProperty("content").GetProperty("parts")[0].GetProperty("text").GetString();
                    Logger.Log($"[Gemini API Result]\n{text}\n-----------------------");
                    
                    InvoiceData data = JsonSerializer.Deserialize<InvoiceData>(text, _jsonOptions);
                    return data;
                }
                else
                {
                    throw new Exception("Gemini không trả về kết quả dự kiến.");
                }
            }
        }

        private void NewSessionBtn_Click(object sender, RoutedEventArgs e)
        {
            // Chặn tạo phiên mới nếu batch đang chạy
            if (_cts != null && !_cts.Token.IsCancellationRequested && _items.Any(x => x.Status == ProcessStatus.Processing))
            {
                var cancelResult = MessageBox.Show("Đang có batch xử lý chạy nền. Bạn có muốn HỦY batch hiện tại và tạo phiên mới?", "Batch đang chạy", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (cancelResult == MessageBoxResult.No) return;
                _cts.Cancel();
            }

            if (_items.Any(x => x.Status == ProcessStatus.Success))
            {
                var result = MessageBox.Show("Bạn có hóa đơn đã xử lý thành công nhưng chưa Xuất/Lưu.\n\nBạn có chắc chắn muốn Xóa toàn bộ dữ liệu hiện tại để tạo phiên mới không?", "Cảnh báo mất dữ liệu", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (result == MessageBoxResult.No) return;
            }

            _items.Clear();
            ProgressBar.Visibility = Visibility.Collapsed;
            ProgressText.Visibility = Visibility.Collapsed;
            CancelBtn.Visibility = Visibility.Collapsed;
            StartBtn.IsEnabled = true;
            StartBtn.Content = "▶ Bắt đầu xử lý";
            
            PreviewImage.Source = null;
            NgayGiaoText.Text = "";
            KhachHangText.Text = "";
            DiemGiaoText.Text = "";
            ProductsGrid.ItemsSource = null;
            ContinueBtn.Visibility = Visibility.Collapsed;
            
            ShipperSelector.Text = "";
            ShipperSelector.Focus();
        }

        private void LoadShippers()
        {
            if (File.Exists(ShippersFilePath))
            {
                try
                {
                    string json = File.ReadAllText(ShippersFilePath);
                    var list = JsonSerializer.Deserialize<List<string>>(json);
                    if (list != null)
                    {
                        foreach (var s in list) _shippers.Add(s);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"[App] Lỗi tải file shippers.json: {ex.Message}");
                }
            }
        }

        private void SaveShippers()
        {
            try
            {
                string json = JsonSerializer.Serialize(_shippers.ToList());
                File.WriteAllText(ShippersFilePath, json);
            }
            catch (Exception ex)
            {
                Logger.Log($"[App] Lỗi lưu file shippers.json: {ex.Message}");
            }
        }

        private void LoadExcelConfigs()
        {
            if (File.Exists(ConfigsFilePath))
            {
                try
                {
                    string json = File.ReadAllText(ConfigsFilePath);
                    var list = JsonSerializer.Deserialize<List<ExcelConfigItem>>(json);
                    if (list != null)
                    {
                        _excelConfigs.Clear();
                        foreach (var item in list) _excelConfigs.Add(item);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"[App] Lỗi tải file excel_configs.json: {ex.Message}");
                }
            }
        }

        private void SaveExcelConfigs()
        {
            try
            {
                string json = JsonSerializer.Serialize(_excelConfigs.ToList(), new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(ConfigsFilePath, json);
            }
            catch (Exception ex)
            {
                Logger.Log($"[App] Lỗi lưu file excel_configs.json: {ex.Message}");
            }
        }

        private void ShowRefLoading(string text)
        {
            RefLoadingText.Text = text;
            RefLoadingOverlay.Visibility = Visibility.Visible;
            ChangeSheetsBtn.IsEnabled = false;
        }

        private void HideRefLoading()
        {
            RefLoadingOverlay.Visibility = Visibility.Collapsed;
            ChangeSheetsBtn.IsEnabled = _refWorkbook != null;
        }

        private void SetActiveExcelConfig(string localPath)
        {
            foreach (var cfg in _excelConfigs)
            {
                cfg.IsActive = cfg.LocalPath.Equals(localPath, StringComparison.OrdinalIgnoreCase);
            }
        }

        private async void HomeImportFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xlsm",
                Title = "Chọn File Excel Bán Hàng & Tham Chiếu"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string originalPath = openFileDialog.FileName;
                string fileName = Path.GetFileName(originalPath);
                string targetDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, RawMaterialDirName);
                string targetPath = Path.Combine(targetDir, fileName);

                try
                {
                    // Check if file already exists in raw_material
                    if (File.Exists(targetPath))
                    {
                        var result = MessageBox.Show($"File '{fileName}' đã tồn tại trong thư mục raw_material.\n\nBạn có muốn ghi đè lên file này không? (Chọn 'No' để lưu dưới tên mới với hậu tố thời gian)", "File trùng tên", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
                        
                        if (result == MessageBoxResult.Cancel)
                        {
                            return;
                        }
                        else if (result == MessageBoxResult.No)
                        {
                            string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                            string ext = Path.GetExtension(fileName);
                            string uniqueName = $"{nameWithoutExt}_{DateTime.Now:yyyyMMddHHmmss}{ext}";
                            targetPath = Path.Combine(targetDir, uniqueName);
                        }
                    }

                    // Copy file locally
                    File.Copy(originalPath, targetPath, true);

                    // Now load and configure
                    ShowRefLoading("Đang nạp file Excel...");
                    var workbook = await Task.Run(() => new XLWorkbook(targetPath));

                    _refWorkbook?.Dispose();
                    _refWorkbook = workbook;
                    RefFilePathText.Text = targetPath;

                    var sheetNames = _refWorkbook.Worksheets.Select(x => x.Name).ToList();
                    HideRefLoading();

                    var selectorWin = new SheetSelectorWindow(sheetNames) { Owner = this };
                    if (selectorWin.ShowDialog() == true)
                    {
                        string refSheet = selectorWin.SelectedReferenceSheet;
                        string salesSheet = selectorWin.SelectedSalesSheet;

                        SelectedRefSheetText.Text = refSheet;
                        SelectedSalesSheetText.Text = salesSheet;
                        ChangeSheetsBtn.IsEnabled = true;

                        // Save new config
                        var configItem = new ExcelConfigItem
                        {
                            OriginalPath = originalPath,
                            LocalPath = targetPath,
                            FileName = Path.GetFileName(targetPath),
                            RefSheet = refSheet,
                            SalesSheet = salesSheet,
                            DateAdded = DateTime.Now
                        };
                        
                        // If same local path exists in config list, remove it first
                        var existing = _excelConfigs.FirstOrDefault(x => x.LocalPath.Equals(targetPath, StringComparison.OrdinalIgnoreCase));
                        if (existing != null)
                        {
                            _excelConfigs.Remove(existing);
                        }

                        _excelConfigs.Add(configItem);
                        SaveExcelConfigs();

                        // Set active status
                        SetActiveExcelConfig(targetPath);

                        // Load data
                        await LoadBothSheetsAsync(refSheet, salesSheet);

                        // Switch to the Reference tab (Index 2)
                        MainTabControl.SelectedIndex = 2;
                    }
                    else
                    {
                        // Clean up copied file if they cancelled the initial configuration
                        if (File.Exists(targetPath))
                        {
                            try { File.Delete(targetPath); } catch {}
                        }
                        _refWorkbook?.Dispose();
                        _refWorkbook = null;
                        RefFilePathText.Text = "Chưa chọn file Excel tham chiếu.";
                        ChangeSheetsBtn.IsEnabled = false;
                        SelectedRefSheetText.Text = "Chưa chọn";
                        SelectedSalesSheetText.Text = "Chưa chọn";
                    }
                }
                catch (IOException)
                {
                    MessageBox.Show("Không thể mở file Excel này vì đang được mở ở chương trình khác. Vui lòng đóng file đó lại và thử lại.", "File đang mở", MessageBoxButton.OK, MessageBoxImage.Warning);
                    HideRefLoading();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Lỗi nạp file Excel: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    HideRefLoading();
                }
            }
        }

        private void LoadRefExcel_Click(object sender, RoutedEventArgs e)
        {
            HomeImportFile_Click(sender, e);
        }

        private async void HomeLoadConfigRow_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is ExcelConfigItem config)
            {
                if (!File.Exists(config.LocalPath))
                {
                    // Fallback to original path if local copy is missing
                    if (File.Exists(config.OriginalPath))
                    {
                        var result = MessageBox.Show($"Không tìm thấy file bản sao cục bộ tại raw_material. Bạn có muốn phục hồi bằng cách sao chép lại từ đường dẫn gốc '{config.OriginalPath}'?", "Không tìm thấy file", MessageBoxButton.YesNo, MessageBoxImage.Question);
                        if (result == MessageBoxResult.Yes)
                        {
                            try
                            {
                                string dir = Path.GetDirectoryName(config.LocalPath) ?? "";
                                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                                File.Copy(config.OriginalPath, config.LocalPath, true);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Không thể khôi phục file: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                                return;
                            }
                        }
                        else
                        {
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show($"Tệp Excel không tồn tại ở cả đường dẫn cục bộ và đường dẫn gốc. Vui lòng kiểm tra lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                }

                ShowRefLoading("Đang nạp file Excel từ bản sao cục bộ...");
                try
                {
                    var workbook = await Task.Run(() => new XLWorkbook(config.LocalPath));
                    _refWorkbook?.Dispose();
                    _refWorkbook = workbook;
                    RefFilePathText.Text = config.LocalPath;

                    SelectedRefSheetText.Text = config.RefSheet;
                    SelectedSalesSheetText.Text = config.SalesSheet;
                    ChangeSheetsBtn.IsEnabled = true;

                    // Set active status
                    SetActiveExcelConfig(config.LocalPath);

                    await LoadBothSheetsAsync(config.RefSheet, config.SalesSheet);

                    // Switch to reference tab
                    MainTabControl.SelectedIndex = 2;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Lỗi nạp file Excel: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    HideRefLoading();
                }
            }
        }

        private void HomeDeleteConfigRow_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is ExcelConfigItem config)
            {
                var confirm = MessageBox.Show($"Bạn có chắc chắn muốn xóa cấu hình của file '{config.FileName}'?\n\nChú ý: File bản sao tương ứng trong thư mục raw_material cũng sẽ bị xóa bỏ.", "Xác nhận xóa", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                
                if (confirm == MessageBoxResult.Yes)
                {
                    DeleteConfigAndFile(config);
                    SaveExcelConfigs();
                }
            }
        }

        private void HomeDeleteSelected_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = HomeConfigsGrid.SelectedItems.Cast<ExcelConfigItem>().ToList();
            if (selectedItems.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn ít nhất một cấu hình để xóa.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var confirm = MessageBox.Show($"Bạn có chắc chắn muốn xóa {selectedItems.Count} cấu hình được chọn?\n\nChú ý: Các file bản sao tương ứng trong thư mục raw_material cũng sẽ bị xóa bỏ hoàn toàn.", "Xác nhận xóa nhiều tệp", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            
            if (confirm == MessageBoxResult.Yes)
            {
                foreach (var config in selectedItems)
                {
                    DeleteConfigAndFile(config);
                }
                SaveExcelConfigs();
            }
        }

        private void DeleteConfigAndFile(ExcelConfigItem config)
        {
            // Delete local file
            if (File.Exists(config.LocalPath))
            {
                try
                {
                    // If the workbook is currently open, dispose it first so the file is not locked!
                    if (_refWorkbook != null && RefFilePathText.Text == config.LocalPath)
                    {
                        _refWorkbook.Dispose();
                        _refWorkbook = null;
                        RefFilePathText.Text = "Chưa chọn file Excel tham chiếu.";
                        ChangeSheetsBtn.IsEnabled = false;
                        SelectedRefSheetText.Text = "Chưa chọn";
                        SelectedSalesSheetText.Text = "Chưa chọn";
                        _refItems.Clear();
                        SalesDataGrid.ItemsSource = null;
                        SetActiveExcelConfig(""); // Reset active states
                    }
                    File.Delete(config.LocalPath);
                }
                catch (Exception ex)
                {
                    Logger.Log($"[App] Không thể xóa file cục bộ {config.LocalPath}: {ex.Message}");
                }
            }

            _excelConfigs.Remove(config);
        }

        private async void ChangeSheets_Click(object sender, RoutedEventArgs e)
        {
            if (_refWorkbook == null) return;
            var sheetNames = _refWorkbook.Worksheets.Select(x => x.Name).ToList();
            var selectorWin = new SheetSelectorWindow(sheetNames) { Owner = this };
            
            if (selectorWin.ShowDialog() == true)
            {
                string refSheet = selectorWin.SelectedReferenceSheet;
                string salesSheet = selectorWin.SelectedSalesSheet;

                SelectedRefSheetText.Text = refSheet;
                SelectedSalesSheetText.Text = salesSheet;

                await LoadBothSheetsAsync(refSheet, salesSheet);
            }
        }

        private async Task LoadBothSheetsAsync(string refSheetName, string salesSheetName)
        {
            ShowRefLoading("Đang đọc dữ liệu các sheet...");

            try
            {
                // Unbind ItemsSource and clear collections to prevent UI rendering overhead
                RefDataGrid.ItemsSource = null;
                SalesDataGrid.ItemsSource = null;
                _refItems.Clear();

                // Run Reference and Sales sheet processing sequentially (ClosedXML is NOT thread-safe)
                var items = await Task.Run(() =>
                {
                    var resultList = new List<ReferenceItem>();
                    var worksheet = _refWorkbook.Worksheet(refSheetName);

                    // Find header row in first 20 rows
                    int headerRow = -1;
                    int colAbbrev = -1;
                    int colName = -1;
                    int colUnit = -1;

                    for (int r = 1; r <= 20; r++)
                    {
                        var row = worksheet.Row(r);
                        for (int c = 1; c <= 20; c++)
                        {
                            string val = row.Cell(c).GetString().Trim().ToLower();
                            if (val.Contains("viết tắt") || val == "viet tat" || val == "ma" || val == "mã")
                            {
                                colAbbrev = c;
                            }
                            else if (val.Contains("tên hàng") || val == "ten hang" || val == "sản phẩm" || val == "san pham")
                            {
                                colName = c;
                            }
                            else if (val.Contains("đơn vị") || val == "don vi" || val == "đvt" || val == "dvt")
                            {
                                colUnit = c;
                            }
                        }

                        if (colAbbrev != -1 && colName != -1 && colUnit != -1)
                        {
                            headerRow = r;
                            break;
                        }
                    }

                    if (headerRow == -1)
                    {
                        headerRow = 2; // Default assume row 2
                        colAbbrev = 1;
                        colName = 2;
                        colUnit = 3;
                    }

                    int lastRow = Math.Min(worksheet.LastRowUsed()?.RowNumber() ?? 1000, headerRow + 1000);
                    for (int r = headerRow + 1; r <= lastRow; r++)
                    {
                        string abbrev = worksheet.Cell(r, colAbbrev).GetString().Trim();
                        string name = worksheet.Cell(r, colName).GetString().Trim();
                        string unit = worksheet.Cell(r, colUnit).GetString().Trim();

                        if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(abbrev)) continue;

                        string unitLower = unit.ToLower();
                        if (unitLower == "kg") continue; // Filter out kg, Kg, KG, kG

                        resultList.Add(new ReferenceItem
                        {
                            VietTat = abbrev,
                            TenHang = name,
                            DonViTinh = unit
                        });
                    }
                    return resultList;
                });

                foreach (var item in items)
                {
                    _refItems.Add(item);
                }

                var salesDt = await Task.Run(() =>
                {
                    var dt = new DataTable();
                    var worksheet = _refWorkbook.Worksheet(salesSheetName);
                    var range = worksheet.RangeUsed();
                    if (range != null)
                    {
                        int rowCount = range.RowCount();
                        int colCount = range.ColumnCount();

                        // Row 1: Headers
                        var firstRow = range.FirstRow();
                        for (int col = 1; col <= colCount; col++)
                        {
                            string header = firstRow.Cell(col).GetString().Trim();
                            if (string.IsNullOrEmpty(header))
                            {
                                header = XLHelper.GetColumnLetterFromNumber(col);
                            }

                            string uniqueHeader = header;
                            int counter = 1;
                            while (dt.Columns.Contains(uniqueHeader))
                            {
                                uniqueHeader = $"{header}_{counter++}";
                            }
                            dt.Columns.Add(uniqueHeader);
                        }

                        // Rows 2 to N
                        for (int r = 2; r <= rowCount; r++)
                        {
                            var row = range.Row(r);
                            var dr = dt.NewRow();
                            for (int col = 1; col <= colCount; col++)
                            {
                                var cell = row.Cell(col);
                                var val = cell.HasFormula ? cell.CachedValue : cell.Value;
                                if (val.Type == XLDataType.DateTime)
                                {
                                    var dtValue = val.GetDateTime();
                                    if (dtValue.TimeOfDay == TimeSpan.Zero)
                                    {
                                        dr[col - 1] = dtValue.ToString("dd/MM/yyyy");
                                    }
                                    else
                                    {
                                        dr[col - 1] = dtValue.ToString("dd/MM/yyyy HH:mm:ss");
                                    }
                                }
                                else
                                {
                                    dr[col - 1] = val.ToString();
                                }
                            }
                            dt.Rows.Add(dr);
                        }
                    }
                    return dt;
                });

                SalesDataGrid.ItemsSource = salesDt.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi nạp dữ liệu sheet: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                // Rebind ItemsSource to trigger a single layout render
                RefDataGrid.ItemsSource = _refItems;
                HideRefLoading();
            }
        }

        private void InitStagingDatabase()
        {
            try
            {
                string connStr = $"Data Source={StagingDbPath}";
                using (var connection = new DuckDBConnection(connStr))
                {
                    connection.Open();
                    using (var cmd = connection.CreateCommand())
                    {
                        cmd.CommandText = @"
                            CREATE TABLE IF NOT EXISTS staging_items (
                                id VARCHAR PRIMARY KEY,
                                file_name VARCHAR,
                                ngay_giao VARCHAR,
                                khach_hang VARCHAR,
                                diem_giao VARCHAR,
                                nguoi_giao VARCHAR,
                                ten_hang_goc VARCHAR,
                                ten_hang_khop VARCHAR,
                                hst VARCHAR,
                                don_vi VARCHAR,
                                sl_xuat DOUBLE,
                                sl_nhan DOUBLE,
                                sl_hong DOUBLE,
                                ghi_chu VARCHAR,
                                status VARCHAR DEFAULT 'Success',
                                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                            );";
                        cmd.ExecuteNonQuery();
                    }
                }
                Logger.Log("[DuckDB] Khởi tạo cơ sở dữ liệu DuckDB thành công.");
            }
            catch (Exception ex)
            {
                Logger.Log($"[DuckDB Lỗi] Không thể khởi tạo database: {ex.Message}");
                MessageBox.Show($"Lỗi khởi tạo cơ sở dữ liệu DuckDB: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadStagingItemsFromDb()
        {
            _stagingItems.Clear();
            try
            {
                string connStr = $"Data Source={StagingDbPath}";
                using (var connection = new DuckDBConnection(connStr))
                {
                    connection.Open();
                    using (var cmd = connection.CreateCommand())
                    {
                        cmd.CommandText = "SELECT id, file_name, ngay_giao, khach_hang, diem_giao, nguoi_giao, ten_hang_goc, ten_hang_khop, hst, don_vi, sl_xuat, sl_nhan, sl_hong, ghi_chu FROM staging_items WHERE status != 'Synced' ORDER BY created_at DESC;";
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var item = new StagingItem
                                {
                                    Id = reader.GetString(0),
                                    FileName = reader.GetString(1),
                                    NgayGiao = reader.IsDBNull(2) ? "" : reader.GetString(2),
                                    KhachHang = reader.IsDBNull(3) ? "" : reader.GetString(3),
                                    DiemGiao = reader.IsDBNull(4) ? "" : reader.GetString(4),
                                    NguoiGiao = reader.IsDBNull(5) ? "" : reader.GetString(5),
                                    TenHangGoc = reader.IsDBNull(6) ? "" : reader.GetString(6),
                                    TenHangKhop = reader.IsDBNull(7) ? "" : reader.GetString(7),
                                    Hst = reader.IsDBNull(8) ? "" : reader.GetString(8),
                                    DonVi = reader.IsDBNull(9) ? "" : reader.GetString(9),
                                    SlXuat = reader.IsDBNull(10) ? 0 : reader.GetDouble(10),
                                    SlNhan = reader.IsDBNull(11) ? 0 : reader.GetDouble(11),
                                    SlHong = reader.IsDBNull(12) ? 0 : reader.GetDouble(12),
                                    GhiChu = reader.IsDBNull(13) ? "" : reader.GetString(13)
                                };
                                _stagingItems.Add(item);
                            }
                        }
                    }
                }
                PreviewDataGrid.ItemsSource = _stagingItems;
            }
            catch (Exception ex)
            {
                Logger.Log($"[DuckDB Lỗi] LoadStagingItemsFromDb: {ex.Message}");
                MessageBox.Show($"Lỗi tải dữ liệu xem trước: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateStagingItemInDb(StagingItem item)
        {
            try
            {
                string connStr = $"Data Source={StagingDbPath}";
                using (var connection = new DuckDBConnection(connStr))
                {
                    connection.Open();
                    using (var cmd = connection.CreateCommand())
                    {
                        cmd.CommandText = @"
                            UPDATE staging_items SET 
                                ngay_giao = $ngay_giao, 
                                khach_hang = $khach_hang, 
                                diem_giao = $diem_giao, 
                                nguoi_giao = $nguoi_giao, 
                                ten_hang_khop = $ten_hang_khop, 
                                hst = $hst, 
                                don_vi = $don_vi,
                                sl_xuat = $sl_xuat, 
                                sl_nhan = $sl_nhan, 
                                sl_hong = $sl_hong, 
                                ghi_chu = $ghi_chu 
                            WHERE id = $id;";
                        
                        cmd.Parameters.Add(new DuckDBParameter("ngay_giao", item.NgayGiao ?? ""));
                        cmd.Parameters.Add(new DuckDBParameter("khach_hang", item.KhachHang ?? ""));
                        cmd.Parameters.Add(new DuckDBParameter("diem_giao", item.DiemGiao ?? ""));
                        cmd.Parameters.Add(new DuckDBParameter("nguoi_giao", item.NguoiGiao ?? ""));
                        cmd.Parameters.Add(new DuckDBParameter("ten_hang_khop", item.TenHangKhop ?? ""));
                        cmd.Parameters.Add(new DuckDBParameter("hst", item.Hst ?? ""));
                        cmd.Parameters.Add(new DuckDBParameter("don_vi", item.DonVi ?? ""));
                        cmd.Parameters.Add(new DuckDBParameter("sl_xuat", item.SlXuat));
                        cmd.Parameters.Add(new DuckDBParameter("sl_nhan", item.SlNhan));
                        cmd.Parameters.Add(new DuckDBParameter("sl_hong", item.SlHong));
                        cmd.Parameters.Add(new DuckDBParameter("ghi_chu", item.GhiChu ?? ""));
                        cmd.Parameters.Add(new DuckDBParameter("id", item.Id));

                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[DuckDB Lỗi] UpdateStagingItemInDb: {ex.Message}");
            }
        }

        private async void ContinueBtn_Click(object sender, RoutedEventArgs e)
        {
            var processedSuccessItems = _items.Where(x => x.Status == ProcessStatus.Success && x.Data != null).ToList();
            if (processedSuccessItems.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu hóa đơn nào đã xử lý thành công.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var activeConfig = _excelConfigs.FirstOrDefault(x => x.IsActive);
            if (activeConfig == null)
            {
                MessageBox.Show("Vui lòng chọn hoặc nạp một file cấu hình Excel hoạt động ở Trang chủ trước.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Filter out 'kg' units from reference list
            var targetRefItems = _refItems.Where(x => !x.DonViTinh.Equals("kg", StringComparison.OrdinalIgnoreCase)).ToList();

            ShowRefLoading("Đang đối soát sản phẩm và lưu vào DuckDB...");
            
            try
            {
                await Task.Run(() =>
                {
                    // Phase 1: Fuzzy match parallel (CPU-bound, no DB access)
                    var matchResults = new ConcurrentBag<(string fileName, InvoiceData data, InvoiceItem itemRow,
                        string tenKhop, string hst, string donVi)>();

                    Parallel.ForEach(processedSuccessItems, item =>
                    {
                        string fileName = Path.GetFileName(item.FilePath);
                        var data = item.Data;
                        if (data?.danh_sach_hang_hoa == null) return;

                        foreach (var itemRow in data.danh_sach_hang_hoa)
                        {
                            var matchedItem = FuzzyMatcher.FindBestMatch(itemRow.ten_hang, targetRefItems, out double bestScore);

                            string tenKhop = "";
                            string hst = "";
                            string donVi = "";

                            if (matchedItem != null && bestScore > 0.5)
                            {
                                tenKhop = matchedItem.TenHang;
                                hst = matchedItem.VietTat;
                                donVi = matchedItem.DonViTinh;
                            }
                            matchResults.Add((fileName, data, itemRow, tenKhop, hst, donVi));
                        }
                    });

                    // Phase 2: Batch DuckDB insert — single connection, no lock needed
                    try
                    {
                        using (var conn = new DuckDBConnection($"Data Source={StagingDbPath}"))
                        {
                            conn.Open();
                            foreach (var r in matchResults)
                            {
                                using (var cmd = conn.CreateCommand())
                                {
                                    cmd.CommandText = @"
                                        INSERT INTO staging_items (id, file_name, ngay_giao, khach_hang, diem_giao, nguoi_giao, ten_hang_goc, ten_hang_khop, hst, don_vi, sl_xuat, sl_nhan, sl_hong, ghi_chu, status) 
                                        VALUES ($id, $file_name, $ngay_giao, $khach_hang, $diem_giao, $nguoi_giao, $ten_hang_goc, $ten_hang_khop, $hst, $don_vi, $sl_xuat, $sl_nhan, $sl_hong, $ghi_chu, $status);";

                                    cmd.Parameters.Add(new DuckDBParameter("id", Guid.NewGuid().ToString()));
                                    cmd.Parameters.Add(new DuckDBParameter("file_name", r.fileName));
                                    cmd.Parameters.Add(new DuckDBParameter("ngay_giao", r.data.ngay_giao ?? ""));
                                    cmd.Parameters.Add(new DuckDBParameter("khach_hang", FuzzyMatcher.FormatKhachHang(r.data.khach_hang ?? "")));
                                    cmd.Parameters.Add(new DuckDBParameter("diem_giao", FuzzyMatcher.FormatDiemGiao(r.data.diem_giao ?? "")));
                                    cmd.Parameters.Add(new DuckDBParameter("nguoi_giao", r.data.ten_nguoi_giao ?? ""));
                                    cmd.Parameters.Add(new DuckDBParameter("ten_hang_goc", r.itemRow.ten_hang ?? ""));
                                    cmd.Parameters.Add(new DuckDBParameter("ten_hang_khop", r.tenKhop));
                                    cmd.Parameters.Add(new DuckDBParameter("hst", r.hst));
                                    cmd.Parameters.Add(new DuckDBParameter("don_vi", r.donVi));
                                    cmd.Parameters.Add(new DuckDBParameter("sl_xuat", r.itemRow.sl_xuat));
                                    cmd.Parameters.Add(new DuckDBParameter("sl_nhan", r.itemRow.sl_nhan));
                                    cmd.Parameters.Add(new DuckDBParameter("sl_hong", r.itemRow.sl_hong));
                                    cmd.Parameters.Add(new DuckDBParameter("ghi_chu", ""));
                                    cmd.Parameters.Add(new DuckDBParameter("status", "Success"));

                                    cmd.ExecuteNonQuery();
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"[DuckDB Lỗi] Batch ghi staging_items: {ex.Message}");
                    }
                });

                LoadStagingItemsFromDb();
                ContinueBtn.Visibility = Visibility.Collapsed;
                
                // Switch to Preview Tab (Index 2 in MainTabControl)
                MainTabControl.SelectedIndex = 2;
                MessageBox.Show("Đối soát mờ hoàn tất! Đã chuyển sang tab Preview để kiểm tra dữ liệu.", "Thành công", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi đối soát dữ liệu: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                HideRefLoading();
            }
        }

        private void ReloadPreviewBtn_Click(object sender, RoutedEventArgs e)
        {
            LoadStagingItemsFromDb();
        }

        private void DeletePreviewSelected_Click(object sender, RoutedEventArgs e)
        {
            var selected = _stagingItems.Where(x => x.IsSelected).ToList();
            if (selected.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn các dòng mặt hàng cần xóa.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var confirm = MessageBox.Show($"Bạn có chắc chắn muốn xóa {selected.Count} dòng đang chọn khỏi hàng đợi Preview?\nDữ liệu này sẽ mất vĩnh viễn và không ghi vào Excel.", "Xác nhận xóa", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (confirm == MessageBoxResult.Yes)
            {
                try
                {
                    using (var conn = new DuckDBConnection($"Data Source={StagingDbPath}"))
                    {
                        conn.Open();
                        using (var cmd = conn.CreateCommand())
                        {
                            // Batch delete — single statement with IN clause
                            var idList = string.Join(", ", selected.Select(x => $"'{x.Id}'"));
                            cmd.CommandText = $"DELETE FROM staging_items WHERE id IN ({idList});";
                            cmd.ExecuteNonQuery();
                        }
                    }
                    LoadStagingItemsFromDb();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Lỗi khi xóa: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void SelectAllPreviewChk_Click(object sender, RoutedEventArgs e)
        {
            bool isChecked = SelectAllPreviewChk.IsChecked ?? false;
            foreach (var item in _stagingItems)
            {
                item.IsSelected = isChecked;
            }
        }

        private void PreviewDataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction == DataGridEditAction.Commit && e.Row.Item is StagingItem stagingItem)
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    // Recalculate SL Hỏng
                    stagingItem.SlHong = stagingItem.SlXuat - stagingItem.SlNhan;

                    // Automatically update hst and donVi if name is matched manually
                    if (e.Column.Header.ToString().Contains("Tên khớp"))
                    {
                        var refMatch = _refItems.FirstOrDefault(x => x.TenHang.Equals(stagingItem.TenHangKhop, StringComparison.OrdinalIgnoreCase));
                        if (refMatch != null)
                        {
                            stagingItem.Hst = refMatch.VietTat;
                            stagingItem.DonVi = refMatch.DonViTinh;
                        }
                    }

                    UpdateStagingItemInDb(stagingItem);
                }), System.Windows.Threading.DispatcherPriority.Background);
            }
        }

        private async void CommitToExcelBtn_Click(object sender, RoutedEventArgs e)
        {
            var itemsToCommit = _stagingItems.Where(x => x.IsSelected).ToList();
            if (itemsToCommit.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn ít nhất một dòng để ghi sổ Excel.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var activeConfig = _excelConfigs.FirstOrDefault(x => x.IsActive);
            if (activeConfig == null)
            {
                MessageBox.Show("Vui lòng chọn hoặc nạp một file cấu hình Excel hoạt động ở Trang chủ trước.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Check if file exists
            if (!System.IO.File.Exists(activeConfig.LocalPath))
            {
                MessageBox.Show($"Không tìm thấy tệp Excel tại đường dẫn: {activeConfig.LocalPath}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // Check if file is locked
            try
            {
                using (var stream = new System.IO.FileStream(activeConfig.LocalPath, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None))
                {
                    // File is accessible
                }
            }
            catch (System.IO.IOException ex)
            {
                MessageBox.Show($"Tệp Excel đang được mở hoặc bị khóa bởi ứng dụng khác (ví dụ: Microsoft Excel).\n\nVui lòng đóng tệp Excel tại:\n{activeConfig.LocalPath}\n\nSau đó thử lại!", "Tệp đang bị khóa", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Không thể truy cập tệp Excel: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            ShowRefLoading("Đang ghi dữ liệu đơn luồng vào Excel...");

            try
            {
                // Crucial step: Dispose open memory reference of Excel before modifying it
                _refWorkbook?.Dispose();
                _refWorkbook = null;

                await Task.Run(() =>
                {
                    using (var workbook = new XLWorkbook(activeConfig.LocalPath))
                    {
                        var worksheet = workbook.Worksheet(activeConfig.SalesSheet);

                        // Find column indexes based on header row names (row 1)
                        var firstRow = worksheet.Row(1);
                        int colCount = Math.Max(20, firstRow.LastCellUsed()?.Address.ColumnNumber ?? 20);

                        int colDate = 2;
                        int colCust = 3;
                        int colDeliv = 4;
                        int colShipper = 5;
                        int colHst = 6;
                        int colName = 7;
                        int colUnit = 8;
                        int colXuat = 9;
                        int colHong = 10;
                        int colNhan = 11;
                        int colNote = 12;

                        for (int c = 1; c <= colCount; c++)
                        {
                            string header = firstRow.Cell(c).Value.ToString().Trim().ToLower();
                            if (header == "ngày" || header == "ngay") colDate = c;
                            else if (header == "khách hàng" || header == "khach hang" || header == "khachhang") colCust = c;
                            else if (header == "điểm giao" || header == "diem giao" || header == "diemgiao") colDeliv = c;
                            else if (header == "người giao" || header == "nguoi giao" || header == "nguoigiao") colShipper = c;
                            else if (header == "hst") colHst = c;
                            else if (header == "tên hàng" || header == "ten hang" || header == "tenhang" || header == "sản phẩm") colName = c;
                            else if (header == "đơn vị" || header == "don vi" || header == "đvt" || header == "dvt") colUnit = c;
                            else if (header == "xuất" || header == "xuat") colXuat = c;
                            else if (header == "hỏng" || header == "hong") colHong = c;
                            else if (header == "nhận" || header == "nhan") colNhan = c;
                            else if (header == "ghi chú" || header == "ghi chu") colNote = c;
                        }

                        // Find first empty row (where Date, Customer, and Hst are all blank) to write to, preserving template formulas and format
                        int lastRowNumber = worksheet.LastRowUsed()?.RowNumber() ?? 1;
                        int targetRow = 2; // Start scanning after header row
                        for (int r = 2; r <= lastRowNumber + 1; r++)
                        {
                            var dateVal = worksheet.Cell(r, colDate).Value.ToString().Trim();
                            var custVal = worksheet.Cell(r, colCust).Value.ToString().Trim();
                            var hstVal = worksheet.Cell(r, colHst).Value.ToString().Trim();

                            if (string.IsNullOrEmpty(dateVal) && string.IsNullOrEmpty(custVal) && string.IsNullOrEmpty(hstVal))
                            {
                                targetRow = r;
                                break;
                            }
                        }

                        foreach (var item in itemsToCommit)
                        {
                            var row = worksheet.Row(targetRow);

                            row.Cell(colDate).Value = item.NgayGiao;
                            row.Cell(colCust).Value = FuzzyMatcher.FormatKhachHang(item.KhachHang);
                            row.Cell(colDeliv).Value = FuzzyMatcher.FormatDiemGiao(item.DiemGiao);
                            row.Cell(colShipper).Value = item.NguoiGiao;
                            row.Cell(colHst).Value = item.Hst;
                            row.Cell(colXuat).Value = item.SlXuat;
                            row.Cell(colHong).Value = item.SlHong;

                            targetRow++;
                        }

                        workbook.Save();
                    }

                    // Update status in DuckDB staging to 'Synced' — batch single statement
                    using (var conn = new DuckDBConnection($"Data Source={StagingDbPath}"))
                    {
                        conn.Open();
                        using (var cmd = conn.CreateCommand())
                        {
                            var idList = string.Join(", ", itemsToCommit.Select(x => $"'{x.Id}'"));
                            cmd.CommandText = $"UPDATE staging_items SET status = 'Synced' WHERE id IN ({idList});";
                            cmd.ExecuteNonQuery();
                        }
                    }
                });

                MessageBox.Show($"Đã ghi sổ thành công {itemsToCommit.Count} dòng vào Excel!", "Thành công", MessageBoxButton.OK, MessageBoxImage.Information);
                
                // Reload preview grid (synced items are removed)
                LoadStagingItemsFromDb();

                // Re-open and re-load workbook in application memory
                var workbookCopy = await Task.Run(() => new XLWorkbook(activeConfig.LocalPath));
                _refWorkbook = workbookCopy;

                // Reload sheets on Tham chiếu tab to show the newly committed data!
                await LoadBothSheetsAsync(activeConfig.RefSheet, activeConfig.SalesSheet);

                // Switch to Tab Tham chiếu (Index 3 in MainTabControl)
                MainTabControl.SelectedIndex = 3;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi ghi sổ Excel: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                Logger.Log($"[Excel Ghi sổ Lỗi] CommitToExcelBtn_Click: {ex.Message}");
            }
            finally
            {
                HideRefLoading();
            }
        }

        private void ExportExcelBtn_Click(object sender, RoutedEventArgs e)
        {
            var activeConfig = _excelConfigs.FirstOrDefault(x => x.IsActive);
            if (activeConfig == null)
            {
                MessageBox.Show("Vui lòng chọn hoặc nạp một file cấu hình Excel hoạt động ở Trang chủ trước.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!File.Exists(activeConfig.LocalPath))
            {
                MessageBox.Show($"Không tìm thấy tệp Excel cục bộ tại: {activeConfig.LocalPath}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xlsm",
                Title = "Xuất tệp Excel bán hàng",
                FileName = Path.GetFileNameWithoutExtension(activeConfig.FileName) + "_Exported" + Path.GetExtension(activeConfig.FileName)
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    // Copy local copy to chosen location
                    File.Copy(activeConfig.LocalPath, saveFileDialog.FileName, true);
                    MessageBox.Show($"Đã xuất tệp Excel thành công tại:\n{saveFileDialog.FileName}", "Thành công", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Lỗi khi xuất tệp Excel: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void MainWindow_Closing(object? sender, CancelEventArgs e)
        {
            _cts?.Dispose();
            _refWorkbook?.Dispose();
            _refWorkbook = null;
            if (_tessEngine.IsValueCreated)
            {
                _tessEngine.Value?.Dispose();
            }
            _tessEngine.Dispose();
            Logger.Flush();
        }
    }

    public static class FuzzyMatcher
    {
        public static string RemoveSign4VietnameseString(string str)
        {
            if (string.IsNullOrEmpty(str)) return str;

            string formD = str.Normalize(NormalizationForm.FormD);
            StringBuilder sb = new StringBuilder();

            foreach (char ch in formD)
            {
                System.Globalization.UnicodeCategory uc = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(ch);
                if (uc != System.Globalization.UnicodeCategory.NonSpacingMark)
                {
                    if (ch == 'đ') sb.Append('d');
                    else if (ch == 'Đ') sb.Append('D');
                    else sb.Append(ch);
                }
            }

            string result = sb.ToString().Normalize(NormalizationForm.FormC);
            result = result.Replace("đ", "d").Replace("Đ", "D");
            return result;
        }

        public static double GetTokenMatchScore(string source, string target)
        {
            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(target)) return 0.0;

            var sourceTokens = source.Split(new[] { ' ', '-', '/', '+', '.', ',' }, StringSplitOptions.RemoveEmptyEntries);
            var targetTokens = target.Split(new[] { ' ', '-', '/', '+', '.', ',' }, StringSplitOptions.RemoveEmptyEntries);

            if (sourceTokens.Length == 0 || targetTokens.Length == 0) return 0.0;

            int intersectCount = 0;
            foreach (var sTok in sourceTokens)
            {
                if (targetTokens.Any(tTok => tTok == sTok || tTok.Contains(sTok) || sTok.Contains(tTok)))
                {
                    intersectCount++;
                }
            }

            return (double)intersectCount / Math.Max(sourceTokens.Length, targetTokens.Length);
        }

        public static double JaroWinklerDistance(string s1, string s2)
        {
            if (string.IsNullOrEmpty(s1) || string.IsNullOrEmpty(s2)) return 0.0;
            if (s1 == s2) return 1.0;

            int l1 = s1.Length;
            int l2 = s2.Length;

            int matchDistance = Math.Max(l1, l2) / 2 - 1;

            bool[] s1Matches = new bool[l1];
            bool[] s2Matches = new bool[l2];

            int matches = 0;
            int transpositions = 0;

            for (int i = 0; i < l1; i++)
            {
                int start = Math.Max(0, i - matchDistance);
                int end = Math.Min(i + matchDistance + 1, l2);

                for (int j = start; j < end; j++)
                {
                    if (s2Matches[j]) continue;
                    if (s1[i] != s2[j]) continue;

                    s1Matches[i] = true;
                    s2Matches[j] = true;
                    matches++;
                    break;
                }
            }

            if (matches == 0) return 0.0;

            int k = 0;
            for (int i = 0; i < l1; i++)
            {
                if (!s1Matches[i]) continue;
                while (!s2Matches[k]) k++;
                if (s1[i] != s2[k]) transpositions++;
                k++;
            }

            double jaro = ((double)matches / l1 + (double)matches / l2 + (double)(matches - transpositions / 2.0) / matches) / 3.0;

            double p = 0.1;
            int prefixLength = 0;
            for (int i = 0; i < Math.Min(4, Math.Min(l1, l2)); i++)
            {
                if (s1[i] == s2[i]) prefixLength++;
                else break;
            }

            return jaro + prefixLength * p * (1.0 - jaro);
        }

        public static ReferenceItem? FindBestMatch(string rawName, IEnumerable<ReferenceItem> refItems, out double bestScore)
        {
            bestScore = 0.0;
            ReferenceItem? bestMatch = null;

            string cleanRaw = RemoveSign4VietnameseString(rawName.Trim().ToLower());

            foreach (var refItem in refItems)
            {
                string cleanRef = RemoveSign4VietnameseString(refItem.TenHang.Trim().ToLower());
                string cleanAbbrev = RemoveSign4VietnameseString(refItem.VietTat.Trim().ToLower());

                if (!string.IsNullOrEmpty(cleanAbbrev) && cleanRaw == cleanAbbrev)
                {
                    bestScore = 1.0;
                    return refItem;
                }

                double tokenScore = GetTokenMatchScore(cleanRaw, cleanRef);
                double jwScore = JaroWinklerDistance(cleanRaw, cleanRef);

                double score = (0.6 * tokenScore) + (0.4 * jwScore);

                if (score > bestScore)
                {
                    bestScore = score;
                    bestMatch = refItem;
                }
            }

            return bestMatch;
        }

        public static string FormatDiemGiao(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            return value.Trim().ToLower();
        }

        public static string FormatKhachHang(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            
            string cleaned = value;
            string[] noiseKeywords = new[] 
            {
                "công ty tnhh một thành viên",
                "hộ kinh doanh cá thể",
                "doanh nghiệp tư nhân",
                "thương mại dịch vụ",
                "chi nhánh công ty",
                "công ty cổ phần",
                "công ty tnhh mtv",
                "hộ kinh doanh",
                "tổng công ty",
                "công ty tnhh",
                "thương mại",
                "phát triển",
                "công ty cp",
                "hợp tác xã",
                "chi nhánh",
                "sản xuất",
                "tập đoàn",
                "việt nam",
                "viet nam",
                "vietnam",
                "dịch vụ",
                "đầu tư",
                "dntn",
                "hkd"
            };

            foreach (var noise in noiseKeywords)
            {
                int index;
                while ((index = cleaned.IndexOf(noise, StringComparison.OrdinalIgnoreCase)) != -1)
                {
                    cleaned = cleaned.Remove(index, noise.Length);
                }
            }

            // Clean up extra separators or punctuation
            cleaned = cleaned.Replace("-", " ").Replace(",", " ").Replace("(", " ").Replace(")", " ");
            
            var words = cleaned.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            cleaned = string.Join(" ", words);

            if (string.IsNullOrWhiteSpace(cleaned))
            {
                cleaned = value.Trim();
            }

            if (cleaned.Length == 0) return "";
            if (cleaned.Length == 1) return cleaned.ToUpper();
            
            return char.ToUpper(cleaned[0]) + cleaned.Substring(1).ToLower();
        }
    }
}