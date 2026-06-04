using Microsoft.Win32;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
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

namespace AccountingOcrTest
{
    public partial class MainWindow : Window
    {
        private ApiKeyManager _apiKeyManager;
        private ObservableCollection<ProcessingItem> _items = new ObservableCollection<ProcessingItem>();
        private ICollectionView _itemsView;
        private CancellationTokenSource _cts;

        // Fix #3: Static HttpClient — avoids socket exhaustion and DNS caching issues
        private static readonly HttpClient _httpClient = new HttpClient();

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
            _apiKeyManager = new ApiKeyManager();
            ItemsListBox.ItemsSource = _items;
            _itemsView = CollectionViewSource.GetDefaultView(_items);
            _itemsView.Filter = SearchFilter;
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
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Image Files|*.jpg;*.jpeg;*.png;*.webp",
                Title = "Chọn ảnh hóa đơn",
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                foreach (var file in openFileDialog.FileNames)
                {
                    _items.Add(new ProcessingItem 
                    { 
                        FilePath = file, 
                        SmartName = System.IO.Path.GetFileName(file),
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

                // Update Data
                if (item.Data != null)
                {
                    NgayGiaoText.Text = item.Data.ngay_giao;
                    KhachHangText.Text = item.Data.khach_hang;
                    DiemGiaoText.Text = item.Data.diem_giao;
                    ProductsGrid.ItemsSource = item.Data.danh_sach_hang_hoa;
                }
                else
                {
                    NgayGiaoText.Text = "";
                    KhachHangText.Text = "";
                    DiemGiaoText.Text = "";
                    ProductsGrid.ItemsSource = null;
                }
            }
        }

        private async void StartBtn_Click(object sender, RoutedEventArgs e)
        {
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
            _cts = new CancellationTokenSource();

            string selectedModel = "gemini-3.5-flash";
            if (ModelSelector.SelectedItem is ComboBoxItem comboItem)
            {
                selectedModel = comboItem.Content.ToString();
            }

            int concurrencyLevel = _apiKeyManager.Keys.Count * 2;
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
                            string apiKey = await _apiKeyManager.GetNextAvailableKeyAsync(_cts.Token);
                            try
                            {
                                await ProcessImageAsync(item, apiKey, selectedModel, _cts.Token);
                            }
                            catch (Exception ex) when (!_cts.Token.IsCancellationRequested &&
                                (ex.Message.Contains("TooManyRequests") || ex.Message.Contains("429")))
                            {
                                string fallbackModel = selectedModel == "gemini-3.5-flash" ? "gemini-3-flash-preview" : "gemini-3.5-flash";
                                Logger.Log($"[RateLimit] Model {selectedModel} báo quá tải. Chuyển sang model dự phòng {fallbackModel}...");
                                await ProcessImageAsync(item, apiKey, fallbackModel, _cts.Token);
                            }

                            int current = Interlocked.Increment(ref completed);

                            // Fix #2: Update all properties on UI thread
                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                item.Status = ProcessStatus.Success;
                                item.SmartName = $"{item.Data.ngay_giao} - {item.Data.khach_hang} - {item.Data.diem_giao}";

                                // Update progress
                                ProgressBar.Value = current;
                                ProgressText.Text = $"{current}/{total}";

                                // Re-bind details if this item is currently selected
                                if (ItemsListBox.SelectedItem == item)
                                {
                                    ItemsListBox_SelectionChanged(null, null);
                                }
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

        private async Task ProcessImageAsync(ProcessingItem item, string apiKey, string modelName, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            string base64Image = "";
            string mimeType = "image/jpeg";
            if (item.FilePath.ToLower().EndsWith(".png")) mimeType = "image/png";
            if (item.FilePath.ToLower().EndsWith(".webp")) mimeType = "image/webp";

            using (var img = new Bitmap(item.FilePath))
            {
                // Fix #4: Use cached TesseractEngine instead of creating new one each time
                try 
                {
                    var engine = _tessEngine.Value;
                    if (engine != null)
                    {
                        using (var pix = Pix.LoadFromFile(item.FilePath))
                        {
                            using (var page = engine.Process(pix, PageSegMode.OsdOnly))
                            {
                                var iterator = page.GetIterator();
                                if (iterator != null)
                                {
                                    iterator.Begin();
                                    var props = iterator.GetProperties();
                                    Tesseract.Orientation orientation = props.Orientation;

                                    if (orientation == Tesseract.Orientation.PageRight) img.RotateFlip(RotateFlipType.Rotate270FlipNone);
                                    else if (orientation == Tesseract.Orientation.PageDown) img.RotateFlip(RotateFlipType.Rotate180FlipNone);
                                    else if (orientation == Tesseract.Orientation.PageLeft) img.RotateFlip(RotateFlipType.Rotate90FlipNone);
                                }
                            }
                        }
                    }
                }
                catch { }

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
                        base64Image = Convert.ToBase64String(ms.ToArray());
                    }
                }
                finally
                {
                    // Only dispose the resized copy, not the original (which is managed by the outer using)
                    if (needsResize) apiImage.Dispose();
                }
            }

            item.Data = await CallGeminiApi(base64Image, mimeType, apiKey, modelName, token);
        }

        private async Task<InvoiceData> CallGeminiApi(string base64Image, string mimeType, string apiKey, string modelName, CancellationToken token)
        {
            string promptText = @"Bạn là một hệ thống OCR và kiểm toán kế toán cao cấp chuyên nghiệp của Việt Nam. Nhiệm vụ của bạn là số hóa hóa đơn này với ĐỘ CHÍNH XÁC TUYỆT ĐỐI về mặt cấu trúc và nội dung. ĐẶC BIỆT CHÚ Ý: Hình ảnh có thể bị xoay ngang, xoay dọc, hoặc lộn ngược. Hãy tự động 'xoay' và định hướng lại hình ảnh trong tư duy trước khi đọc. TUYỆT ĐỐI KHÔNG ĐƯỢC BỎ SÓT BẤT KỲ DÒNG HÀNG HÓA NÀO. Hãy dóng hàng ngang thật cẩn thận để đảm bảo Tên hàng hóa và Số lượng khớp nhau 100%, không bị trôi dòng, lệch dòng hay mất dòng.

HÃY THỰC HIỆN QUY TRÌNH SUY LUẬN (CHAIN OF THOUGHT):
BẮT BUỘC ghi toàn bộ quá trình phân tích của bạn vào trường `quy_trinh_suy_luan` trong file JSON. Đối với phần hàng hóa, hãy phân tích TỪNG DÒNG một cách tường minh, ví dụ: 'Dòng X: bị một đường kẻ dọc/chéo dài cắt qua cột Đơn giá/ĐVT -> Hủy -> sl_nhan = 0', 'Dòng Y: có nét viết tay số 2 -> sl_nhan = 2'. NẾU CÓ MỘT ĐƯỜNG KẺ DÀI KÉO TỪ DÒNG TRÊN XUỐNG DÒNG DƯỚI, BẠN PHẢI GHI NHẬN LÀ TẤT CẢ CÁC DÒNG ĐÓ ĐỀU BỊ GẠCH BỎ.

Bước 1: Phân tích Cấu trúc Bảng và Thông tin chung
- Ngày giao (ngay_giao): TÌM NGÀY GIAO HÀNG (Delivery Date / Date Received). TUYỆT ĐỐI BỎ QUA 'Ngày In' (Printed Date) và 'Ngày Đặt Hàng' (Order Date). BẮT BUỘC TÌM KIẾM THEO ĐÚNG MỨC ĐỘ ƯU TIÊN: 1. Ngày ghi viết tay trên con dấu -> 2. NGÀY GHI VIẾT TAY bằng bút nằm rải rác (gần chữ ký, ghi chú...) -> 3. Ngày được in sẵn ở ô Ngày Giao Hàng. ĐẶC BIỆT CHÚ Ý TRƯỜNG HỢP SỬA NGÀY: Đôi khi ngày đóng dấu/in sẵn bị sai (ví dụ 24-03) và người ta dùng bút viết tay một ngày khác to hơn, khoanh tròn, gạch bỏ ngày cũ hoặc viết đè lên bên cạnh (ví dụ khoanh tròn số '09' đè lên số '24'). Nếu thấy có hiện tượng sửa ngày như vậy, bạn BẮT BUỘC phải ghép con số được sửa bằng tay đó vào làm ngày giao (ví dụ kết quả trả về phải là 09/03/2026 chứ không phải 24/03/2026 hay Ngày In).
- Khách hàng (khach_hang): Tìm tên các khách hàng (ví dụ: Bigc, Aeon, Winmart, B11, Biggreen...) thường là tên các chuỗi siêu thị hoặc cửa hàng.
- Điểm giao (diem_giao): Tên chi nhánh của khách hàng. ĐẶC BIỆT CHÚ Ý: Nếu trên hóa đơn có ghi MÃ CỬA HÀNG đi kèm TÊN CỬA HÀNG (ví dụ như dòng '1708 - WM HNI Lê Văn Thiêm'), bạn PHẢI trích xuất NGUYÊN VẸN toàn bộ chuỗi đó làm điểm giao (tức là lấy đầy đủ cả mã và tên: '1708 - WM HNI Lê Văn Thiêm'). Nếu không có mã cửa hàng, thì lấy tên chi nhánh ngắn gọn như bình thường (ví dụ: 'Xuân Thủy', 'Ciputra'). Thường thông tin này nằm ở phần vị trí/địa chỉ điểm giao.

Bước 2: Phân tích Thị giác Chuyên sâu và Nhận biết Ký hiệu Viết tay (Ink-to-Text & Symbol Analysis) - QUAN TRỌNG NHẤT
- Phân tách rõ ràng giữa mực in máy và mực viết tay (mực màu xanh, đen mờ, đỏ hoặc nét bút viết tay).
- Quét qua từng dòng dữ liệu hàng hóa và áp dụng QUY TẮC NHẬN DIỆN KÝ HIỆU VIẾT TAY TRÊN HÓA ĐƠN ĐỐI SOÁT VIỆT NAM:
  1. BẤT KỲ nét gạch ngang '-', nét gạch chéo '/', nét kẻ dọc '|', nét mũi tên (->), hoặc một đường gạch chéo dài kéo qua nhiều dòng. CHÚ Ý CỰC KỲ QUAN TRỌNG: Nét gạch này KHÔNG NHẤT THIẾT phải đè lên cột Số lượng, nó CÓ THỂ vắt qua cột Đơn vị tính (ĐVT), cột Đơn giá, hoặc cột Thành tiền. Tóm lại, chỉ cần TRÊN DÒNG HÀNG HÓA ĐÓ bị nét mực gạch xuyên qua (dù là ngang, dọc, hay chéo) VÀ KHÔNG CÓ con số viết tay nào ghi chú kế bên:
     -> ĐÂY LÀ KÝ HIỆU GẠCH BỎ / HỦY GIAO. -> Số lượng thực nhận (sl_nhan) = 0. Tự tin gán bằng 0 cho TẤT CẢ các dòng bị đường mực kéo qua (kể cả kéo qua ĐVT/Đơn giá).
  2. Dấu tích '✓' kèm con số viết tay rõ ràng chỉnh sửa bên cạnh (ví dụ: '✓ 05', '✓ 5', '✓ 10'...):
     -> ĐÂY LÀ KÝ HIỆU XÁC NHẬN GIAO THỰC TẾ. -> Số lượng thực nhận (sl_nhan) = <con số viết tay đó>.
  3. Không có bất kỳ nét bút viết tay nào (chỉ có mực in máy):
     -> Giao đầy đủ. -> Số lượng thực nhận (sl_nhan) = Số lượng in sẵn gốc (sl_xuat).
  4. Nếu số lượng viết tay hoặc in sẵn CÓ DẤU TRỪ (số âm, ví dụ: -2, -5...), điều này có nghĩa là 'nợ hàng' (khách đặt nhưng không có hàng). Hãy ghi nhận CHÍNH XÁC giá trị âm đó vào `sl_nhan` hoặc `sl_xuat` (tùy thuộc vào số âm đó được viết tay hay in sẵn). Được phép ghi nhận số âm.

Bước 3: Đối soát tính toán (Tự động tính Số lượng hỏng)
- Với mỗi mặt hàng, bạn phải có sl_xuat (Số lượng in trên hóa đơn) và sl_nhan (Số lượng giao thực tế, tính toán từ Bước 2).
- Tự động tính toán: sl_hong = sl_xuat - sl_nhan. Trả về đúng giá trị toán học này. Chú ý có thể sl_hong > 0 hoặc = 0 hoặc đôi khi sl_nhan lớn hơn dẫn tới âm. Đảm bảo tính toán chính xác.";

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
            
            var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");
            var response = await _httpClient.PostAsync(url, content, token);
            
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
                    
                    InvoiceData data = JsonSerializer.Deserialize<InvoiceData>(text, new JsonSerializerOptions 
                    { 
                        PropertyNameCaseInsensitive = true 
                    });
                    return data;
                }
                else
                {
                    throw new Exception("Gemini không trả về kết quả dự kiến.");
                }
            }
        }
    }
}