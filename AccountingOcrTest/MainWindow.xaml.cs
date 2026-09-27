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
using System.Windows.Input;
using Tesseract;
using ClosedXML.Excel;
using DuckDB.NET.Data;
using ExcelDataReader;

namespace AccountingOcrTest
{
    public partial class MainWindow : Window
    {
        private ApiKeyManager _apiKeyManager;
        private ObservableCollection<ProcessingItem> _items = new ObservableCollection<ProcessingItem>();
        private ICollectionView _itemsView;
        private CancellationTokenSource _cts;

        // Zoom/Pan state fields
        private System.Windows.Point _panStartPoint;
        private double _panStartX;
        private double _panStartY;
        private bool _isPanning = false;
        
        private static readonly string ShippersFilePath = "shippers.json";
        private ObservableCollection<string> _shippers = new ObservableCollection<string>();

        // Reference excel fields
        private ObservableCollection<ReferenceItem> _refItems = new ObservableCollection<ReferenceItem>();

        // Configs and internal copy fields
        private ObservableCollection<ExcelConfigItem> _excelConfigs = new ObservableCollection<ExcelConfigItem>();
        private static readonly string ConfigsFilePath = "excel_configs.json";
        private static readonly string RawMaterialDirName = "raw_material";

        // DuckDB staging variables
        private static readonly string StagingDbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "accounting_staging.db");
        private ObservableCollection<StagingItem> _stagingItems = new ObservableCollection<StagingItem>();
        private ObservableCollection<SyncBatchItem> _syncBatches = new ObservableCollection<SyncBatchItem>();
        private ICollectionView _stagingItemsView;
        private ICollectionView _syncBatchesView;
        private bool _isUpdatingPreviewFilters = false;
        private bool _isUpdatingHistoryFilters = false;
        private static readonly object DbWriteLock = new object();

        // OpenAI-compatible provider state
        private static readonly string OpenAiConfigFilePath = "openai_config.json";
        private bool _useOpenAiCompat = false;
        private bool _isInitialized = false;
        private bool _isUpdatingSelectAll = false;
        private bool _isSwitchingTab = false;

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
            }, trackAllValues: true);

        public MainWindow()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
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

            // Auto-restore active Excel config if available and file exists
            var activeConfig = _excelConfigs.FirstOrDefault(x => x.IsActive);
            if (activeConfig != null && File.Exists(activeConfig.LocalPath))
            {
                RefFilePathText.Text = activeConfig.LocalPath;
                SelectedRefSheetText.Text = activeConfig.RefSheet;
                SelectedSalesSheetText.Text = activeConfig.SalesSheet;
                ChangeSheetsBtn.IsEnabled = true;
                _ = LoadBothSheetsAsync(activeConfig.RefSheet, activeConfig.SalesSheet);
            }
            else if (activeConfig != null)
            {
                activeConfig.IsActive = false;
                SaveExcelConfigs();
            }

            // Setup DuckDB Staging Database
            InitStagingDatabase();
            LoadStagingItemsFromDb();
            
            _syncBatchesView = CollectionViewSource.GetDefaultView(_syncBatches);
            _syncBatchesView.Filter = HistoryFilter;
            HistoryDataGrid.ItemsSource = _syncBatchesView;
            LoadSyncHistoryFromDb();

            // Load saved OpenAI-compatible config
            LoadOpenAiConfig();
            _isInitialized = true;
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
            {
                _itemsView.Refresh();
                UpdateBatchActionBar();
            }
        }

        private void SearchBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (SearchBox.Text == "🔍 Tìm kiếm...")
            {
                SearchBox.Text = "";
            }
        }

        private void SearchBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(SearchBox.Text))
            {
                SearchBox.Text = "🔍 Tìm kiếm...";
            }
        }

        private void SettingsBtn_Click(object sender, RoutedEventArgs e)
        {
            var win = new ApiKeyManagerWindow(_apiKeyManager) { Owner = this };
            win.ShowDialog();
        }

        private void AddImagesBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureExcelLoaded()) return;

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
                    try
                    {
                        string targetDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, RawMaterialDirName);
                        string targetPath = Path.Combine(targetDir, System.IO.Path.GetFileName(file));
                        if (!File.Exists(targetPath))
                        {
                            File.Copy(file, targetPath, true);
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"[Sao chép ảnh lỗi] Không thể sao chép {file} vào raw_material: {ex.Message}");
                    }

                    _items.Add(new ProcessingItem 
                    { 
                        FilePath = file, 
                        SmartName = $"[{currentShipper}] - {System.IO.Path.GetFileName(file)}",
                        Status = ProcessStatus.Waiting
                    });
                }
                UpdateBatchActionBar();
            }
        }

        private void SelectAllCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (_isUpdatingSelectAll) return;

            bool isChecked = SelectAllCheckBox.IsChecked == true;
            _isUpdatingSelectAll = true;
            try
            {
                if (_itemsView != null)
                {
                    foreach (var obj in _itemsView)
                    {
                        if (obj is ProcessingItem item)
                        {
                            item.IsChecked = isChecked;
                        }
                    }
                }
                else
                {
                    foreach (var item in _items)
                    {
                        item.IsChecked = isChecked;
                    }
                }
            }
            finally
            {
                _isUpdatingSelectAll = false;
            }

            UpdateBatchActionBar();
        }

        private void ItemCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (_isUpdatingSelectAll) return;
            UpdateBatchActionBar();
        }

        private void UpdateBatchActionBar()
        {
            if (BatchActionBar == null || BatchSelectionText == null) return;

            int totalCount = _items.Count;
            int checkedCount = _items.Count(x => x.IsChecked);

            if (checkedCount > 0)
            {
                BatchActionBar.Visibility = Visibility.Visible;
                BatchSelectionText.Text = $"Đã chọn {checkedCount} ảnh";
            }
            else
            {
                BatchActionBar.Visibility = Visibility.Collapsed;
            }

            if (SelectAllCheckBox != null && !_isUpdatingSelectAll)
            {
                _isUpdatingSelectAll = true;
                try
                {
                    if (totalCount == 0 || checkedCount == 0)
                    {
                        SelectAllCheckBox.IsChecked = false;
                    }
                    else if (checkedCount == totalCount)
                    {
                        SelectAllCheckBox.IsChecked = true;
                    }
                    else
                    {
                        SelectAllCheckBox.IsChecked = null;
                    }
                }
                finally
                {
                    _isUpdatingSelectAll = false;
                }
            }
        }

        private void BatchDeleteBtn_Click(object sender, RoutedEventArgs e)
        {
            var checkedItems = _items.Where(x => x.IsChecked).ToList();
            if (checkedItems.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn ít nhất 1 ảnh để xóa.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (checkedItems.Any(x => x.Status == ProcessStatus.Processing))
            {
                MessageBox.Show("Không thể xóa các ảnh đang trong quá trình xử lý.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var result = MessageBox.Show(
                $"Bạn có chắc chắn muốn xóa {checkedItems.Count} ảnh đã chọn khỏi danh sách không?",
                "Xác nhận xóa hàng loạt",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                if (ItemsListBox.SelectedItem is ProcessingItem selected && checkedItems.Contains(selected))
                {
                    ItemsListBox.SelectedItem = null;
                    PreviewImage.Source = null;
                }

                foreach (var item in checkedItems)
                {
                    _items.Remove(item);
                }

                UpdateBatchActionBar();
            }
        }

        private async void BatchReprocessBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureExcelLoaded()) return;

            var checkedItems = _items.Where(x => x.IsChecked).ToList();
            if (checkedItems.Count == 0)
            {
                MessageBox.Show("Vui lòng tích chọn ít nhất 1 ảnh để xử lý lại.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (checkedItems.Any(x => x.Status == ProcessStatus.Processing))
            {
                MessageBox.Show("Không thể xử lý lại khi có ảnh đang trong quá trình chạy.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            await ExecuteProcessingAsync(checkedItems);
        }

        // Fix #1: Async image loading with DecodePixelWidth to prevent UI freeze
        private async void ItemsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Reset zoom/pan transform on selecting a different item
            if (ImageScaleTransform != null && ImageTranslateTransform != null)
            {
                ImageScaleTransform.ScaleX = 1.0;
                ImageScaleTransform.ScaleY = 1.0;
                ImageTranslateTransform.X = 0.0;
                ImageTranslateTransform.Y = 0.0;
            }

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

        private void PreviewImage_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (PreviewImage.Source == null) return;
            
            System.Windows.Point relative = e.GetPosition(PreviewImage);
            double zoom = e.Delta > 0 ? 1.1 : 0.9;
            
            if (zoom > 1.0 && ImageScaleTransform.ScaleX >= 15.0) return;
            if (zoom < 1.0 && ImageScaleTransform.ScaleX <= 0.2) return;

            double oldScaleX = ImageScaleTransform.ScaleX;
            double oldScaleY = ImageScaleTransform.ScaleY;
            
            ImageScaleTransform.ScaleX *= zoom;
            ImageScaleTransform.ScaleY *= zoom;

            ImageTranslateTransform.X -= (relative.X * ImageScaleTransform.ScaleX - relative.X * oldScaleX);
            ImageTranslateTransform.Y -= (relative.Y * ImageScaleTransform.ScaleY - relative.Y * oldScaleY);
        }

        private void PreviewImage_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (PreviewImage.Source == null) return;

            if (e.ChangedButton == MouseButton.Left)
            {
                if (e.ClickCount == 2)
                {
                    ImageScaleTransform.ScaleX = 1.0;
                    ImageScaleTransform.ScaleY = 1.0;
                    ImageTranslateTransform.X = 0.0;
                    ImageTranslateTransform.Y = 0.0;
                    return;
                }

                var element = sender as UIElement;
                if (element != null)
                {
                    _panStartPoint = e.GetPosition(ImageParentGrid);
                    _panStartX = ImageTranslateTransform.X;
                    _panStartY = ImageTranslateTransform.Y;
                    _isPanning = true;
                    element.CaptureMouse();
                }
            }
        }

        private void PreviewImage_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                _isPanning = false;
                (sender as UIElement)?.ReleaseMouseCapture();
            }
        }

        private void PreviewImage_MouseMove(object sender, MouseEventArgs e)
        {
            var element = sender as UIElement;
            if (_isPanning && element != null && element.IsMouseCaptured)
            {
                System.Windows.Point currentPoint = e.GetPosition(ImageParentGrid);
                double deltaX = currentPoint.X - _panStartPoint.X;
                double deltaY = currentPoint.Y - _panStartPoint.Y;

                ImageTranslateTransform.X = _panStartX + deltaX;
                ImageTranslateTransform.Y = _panStartY + deltaY;
            }
        }

        private async void StartBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureExcelLoaded()) return;

            var waitingItems = _items.Where(x => x.Status == ProcessStatus.Waiting).ToList();
            if (waitingItems.Count == 0)
            {
                MessageBox.Show("Không có ảnh nào đang chờ xử lý.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            await ExecuteProcessingAsync(waitingItems);
        }

        private async Task ExecuteProcessingAsync(List<ProcessingItem> targetItems)
        {
            if (!EnsureExcelLoaded()) return;

            if (targetItems == null || targetItems.Count == 0)
            {
                MessageBox.Show("Không có ảnh nào để xử lý.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

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

            if (_useOpenAiCompat)
            {
                if (string.IsNullOrWhiteSpace(OpenAiEndpointBox.Text))
                {
                    MessageBox.Show("Vui lòng nhập Endpoint URL cho OpenAI-compatible provider.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    OpenAiEndpointBox.Focus();
                    return;
                }
                if (string.IsNullOrWhiteSpace(OpenAiKeyBox.Text))
                {
                    MessageBox.Show("Vui lòng nhập API Key cho OpenAI-compatible provider.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    OpenAiKeyBox.Focus();
                    return;
                }
                SaveOpenAiConfig();
            }
            else
            {
                if (_apiKeyManager.Keys.Count == 0)
                {
                    MessageBox.Show("Vui lòng cài đặt API Key trước khi bắt đầu.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            foreach (var item in targetItems)
            {
                item.Status = ProcessStatus.Waiting;
                item.ErrorMessage = "";
            }

            // UI state: processing mode
            StartBtn.IsEnabled = false;
            StartBtn.Content = "⏳ Đang xử lý...";
            CancelBtn.Visibility = Visibility.Visible;
            CancelBtn.IsEnabled = true;
            CancelBtn.Content = "⏹ Hủy";
            _cts?.Dispose();
            _cts = new CancellationTokenSource();

            ToggleUiLock(true);

            try
            {
                string selectedModel = "gemini-3.5-flash";
                string openAiEndpoint = "";
                string openAiKey = "";
                int concurrencyLevel;

                if (_useOpenAiCompat)
                {
                    openAiEndpoint = OpenAiEndpointBox.Text.Trim();
                    openAiKey = OpenAiKeyBox.Text.Trim();
                    selectedModel = string.IsNullOrWhiteSpace(OpenAiModelBox.Text) ? "gpt-4o" : OpenAiModelBox.Text.Trim();
                    int maxConc = 15;
                    if (OpenAiConcurrencyBox != null && int.TryParse(OpenAiConcurrencyBox.Text?.Trim(), out int c) && c > 0)
                    {
                        maxConc = Math.Clamp(c, 1, 30);
                    }
                    concurrencyLevel = Math.Min(targetItems.Count, maxConc);
                }
                else
                {
                    if (ModelSelector.SelectedItem is ComboBoxItem comboItem)
                    {
                        selectedModel = comboItem.Content.ToString();
                    }
                    concurrencyLevel = Math.Min(targetItems.Count, _apiKeyManager.Keys.Count * 2);
                    if (concurrencyLevel == 0) concurrencyLevel = 1; // Fallback nếu không có key
                }

                int total = targetItems.Count;
                int completed = 0;

                // Fix #10: Show progress indicator
                ProgressBar.Visibility = Visibility.Visible;
                ProgressText.Visibility = Visibility.Visible;
                ProgressBar.Maximum = total;
                ProgressBar.Value = 0;
                ProgressText.Text = $"0/{total}";

                Logger.Log($"[Bắt đầu] Đang xử lý {total} hóa đơn bằng {selectedModel} (Provider: {(_useOpenAiCompat ? "OpenAI-compatible" : "Gemini")}) với {concurrencyLevel} luồng...");

                var tasks = new List<Task>();
                var queue = new ConcurrentQueue<ProcessingItem>(targetItems);

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
                                    if (_useOpenAiCompat)
                                    {
                                        try
                                        {
                                            await ProcessImageAsync(item, openAiKey, selectedModel, currentShipper, _cts.Token, true, openAiEndpoint);
                                            success = true;
                                        }
                                        catch (Exception ex) when (!_cts.Token.IsCancellationRequested)
                                        {
                                            retryCount++;
                                            lastErrorMessage = ex.Message;
                                            Logger.Log($"[OpenAI Lỗi] File {Path.GetFileName(item.FilePath)}: {ex.Message}. Thử lại lần {retryCount}/{maxRetries}...");
                                            if (retryCount < maxRetries)
                                            {
                                                await Task.Delay(2000, _cts.Token);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        string apiKey = await _apiKeyManager.GetNextAvailableKeyAsync(_cts.Token);
                                        try
                                        {
                                            await ProcessImageAsync(item, apiKey, selectedModel, currentShipper, _cts.Token);
                                            _apiKeyManager.RecordSuccess(apiKey);
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
            }
            finally
            {
                // Flush buffered logs before showing completion message
                Logger.Flush();

                // UI state: done
                StartBtn.IsEnabled = true;
                StartBtn.Content = "▶ Bắt đầu xử lý";
                CancelBtn.Visibility = Visibility.Collapsed;
                ProgressBar.Visibility = Visibility.Collapsed;
                ProgressText.Visibility = Visibility.Collapsed;

                ToggleUiLock(false);
                UpdateBatchActionBar();
            }

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

        private void ToggleUiLock(bool lockUi)
        {
            // Lock/unlock other tab items
            for (int i = 0; i < MainTabControl.Items.Count; i++)
            {
                if (MainTabControl.Items[i] is TabItem tabItem)
                {
                    if (i != MainTabControl.SelectedIndex)
                    {
                        tabItem.IsEnabled = !lockUi;
                    }
                }
            }

            // Lock/unlock controls in the active tab (Xử lý OCR)
            SettingsBtn.IsEnabled = !lockUi;
            AddImagesBtn.IsEnabled = !lockUi;
            ShipperSelector.IsEnabled = !lockUi;
            if (ProviderSelector != null) ProviderSelector.IsEnabled = !lockUi;
            ModelSelector.IsEnabled = !lockUi;
            if (OpenAiEndpointBox != null) OpenAiEndpointBox.IsEnabled = !lockUi;
            if (OpenAiKeyBox != null) OpenAiKeyBox.IsEnabled = !lockUi;
            if (OpenAiModelBox != null) OpenAiModelBox.IsEnabled = !lockUi;
            if (OpenAiConcurrencyBox != null) OpenAiConcurrencyBox.IsEnabled = !lockUi;
            if (OpenAiDecConcBtn != null) OpenAiDecConcBtn.IsEnabled = !lockUi;
            if (OpenAiIncConcBtn != null) OpenAiIncConcBtn.IsEnabled = !lockUi;
            NewSessionBtn.IsEnabled = !lockUi;
            StartBtn.IsEnabled = !lockUi;
            ContinueBtn.IsEnabled = !lockUi;
            SearchBox.IsEnabled = !lockUi;
            ItemsListBox.IsEnabled = !lockUi;

            if (lockUi)
            {
                CancelBtn.IsEnabled = true;
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
        private static Bitmap ResizeImageForApi(Bitmap original, int maxWidth = 3072, int maxHeight = 3072)
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

        /// <summary>
        /// Enhance image contrast and sharpness to make faint handwritten pen strokes
        /// (tick marks, strike-throughs, handwritten numbers) more visible to the AI model.
        /// This is critical for invoices where ink marks are light or overlap with printed grid lines.
        /// </summary>
        private static Bitmap EnhanceImageForApi(Bitmap original)
        {
            int w = original.Width;
            int h = original.Height;
            var enhanced = new Bitmap(w, h);

            // Step 1: Increase contrast (factor 1.4) to make faint pen strokes stand out
            float contrastFactor = 1.4f;
            float contrastOffset = (1.0f - contrastFactor) / 2.0f * 255;

            var contrastMatrix = new System.Drawing.Imaging.ColorMatrix(new float[][]
            {
                new float[] { contrastFactor, 0, 0, 0, 0 },
                new float[] { 0, contrastFactor, 0, 0, 0 },
                new float[] { 0, 0, contrastFactor, 0, 0 },
                new float[] { 0, 0, 0, 1, 0 },
                new float[] { contrastOffset / 255f, contrastOffset / 255f, contrastOffset / 255f, 0, 1 }
            });

            using (var g = Graphics.FromImage(enhanced))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                using (var attributes = new System.Drawing.Imaging.ImageAttributes())
                {
                    attributes.SetColorMatrix(contrastMatrix);
                    g.DrawImage(original, new Rectangle(0, 0, w, h), 0, 0, w, h, GraphicsUnit.Pixel, attributes);
                }
            }

            // Step 2: Apply sharpening via 3x3 convolution kernel to enhance ink edges
            try
            {
                var sharpened = new Bitmap(w, h);
                // Sharpen kernel: center=9, neighbors=-1 each
                float[,] kernel = {
                    { 0, -1, 0 },
                    { -1,  5, -1 },
                    { 0, -1, 0 }
                };

                for (int y = 1; y < h - 1; y++)
                {
                    for (int x = 1; x < w - 1; x++)
                    {
                        float rSum = 0, gSum = 0, bSum = 0;
                        for (int ky = -1; ky <= 1; ky++)
                        {
                            for (int kx = -1; kx <= 1; kx++)
                            {
                                var pixel = enhanced.GetPixel(x + kx, y + ky);
                                float k = kernel[ky + 1, kx + 1];
                                rSum += pixel.R * k;
                                gSum += pixel.G * k;
                                bSum += pixel.B * k;
                            }
                        }
                        int r = Math.Clamp((int)rSum, 0, 255);
                        int gVal = Math.Clamp((int)gSum, 0, 255);
                        int b = Math.Clamp((int)bSum, 0, 255);
                        sharpened.SetPixel(x, y, System.Drawing.Color.FromArgb(r, gVal, b));
                    }
                }

                // Copy border pixels as-is
                for (int x = 0; x < w; x++)
                {
                    sharpened.SetPixel(x, 0, enhanced.GetPixel(x, 0));
                    sharpened.SetPixel(x, h - 1, enhanced.GetPixel(x, h - 1));
                }
                for (int y = 0; y < h; y++)
                {
                    sharpened.SetPixel(0, y, enhanced.GetPixel(0, y));
                    sharpened.SetPixel(w - 1, y, enhanced.GetPixel(w - 1, y));
                }

                enhanced.Dispose();
                Logger.Log($"[Image Enhancement] Đã tăng tương phản (x{contrastFactor}) và làm nét ảnh {w}x{h}.");
                return sharpened;
            }
            catch (Exception ex)
            {
                Logger.Log($"[Image Enhancement] Lỗi sharpen, chỉ dùng contrast: {ex.Message}");
                return enhanced; // Fallback: return contrast-only version
            }
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

        private async Task ProcessImageAsync(ProcessingItem item, string apiKey, string modelName, string shipperName, CancellationToken token, bool useOpenAi = false, string openAiEndpoint = "")
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
                                    using (var iterator = page.AnalyseLayout())
                                    {
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
                }
                catch (Exception ex)
                {
                    Logger.Log($"[Tesseract Lỗi] Không thể chạy OSD: {ex.Message}");
                }

                token.ThrowIfCancellationRequested();

                // Save rotated image locally to raw_material folder for subsequent previewing/zooming
                try
                {
                    string targetDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, RawMaterialDirName);
                    string targetPath = Path.Combine(targetDir, Path.GetFileName(item.FilePath));
                    img.Save(targetPath, System.Drawing.Imaging.ImageFormat.Jpeg);
                }
                catch (Exception ex)
                {
                    Logger.Log($"[Save Image Lỗi] Không thể lưu ảnh vào raw_material: {ex.Message}");
                }

                // Fix #7: Resize large images before encoding for API (limit increased to 3072 for visual accuracy of handwritten marks)
                bool needsResize = img.Width > 3072 || img.Height > 3072;
                Bitmap resizedImage = needsResize ? ResizeImageForApi(img, 3072, 3072) : img;

                // Enhance contrast + sharpness to make faint handwritten marks (ticks, strike-throughs) visible to the AI
                Bitmap apiImage = EnhanceImageForApi(resizedImage);

                try
                {
                    using (var ms = new MemoryStream())
                    {
                        if (mimeType == "image/png")
                        {
                            apiImage.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                        }
                        else
                        {
                            // Use high-quality JPEG (95%) to preserve fine ink strokes
                            // Default ImageFormat.Jpeg uses 75% quality which blurs faint handwritten marks
                            var jpegEncoder = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
                                .First(enc => enc.MimeType == "image/jpeg");
                            var encoderParams = new System.Drawing.Imaging.EncoderParameters(1);
                            encoderParams.Param[0] = new System.Drawing.Imaging.EncoderParameter(
                                System.Drawing.Imaging.Encoder.Quality, 95L);
                            apiImage.Save(ms, jpegEncoder, encoderParams);
                        }
                        // Use GetBuffer() + Length to avoid allocating a copy array on LOH
                        base64Image = Convert.ToBase64String(ms.GetBuffer(), 0, (int)ms.Length);
                    }
                }
                finally
                {
                    // Dispose the enhanced copy (always a new bitmap)
                    apiImage.Dispose();
                    // Only dispose the resized copy if it's different from the original
                    if (needsResize) resizedImage.Dispose();
                }
            }

            if (useOpenAi)
            {
                item.Data = await CallOpenAiCompatApi(base64Image, mimeType, openAiEndpoint, apiKey, modelName, token);
            }
            else
            {
                item.Data = await CallGeminiApi(base64Image, mimeType, apiKey, modelName, token);
            }
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

        public static async Task<InvoiceData> CallGeminiApi(string base64Image, string mimeType, string apiKey, string modelName, CancellationToken token)
        {
            string promptText = @"Bạn là một hệ thống OCR và kiểm toán kế toán cao cấp chuyên nghiệp của Việt Nam. Nhiệm vụ của bạn là số hóa hóa đơn này với ĐỘ CHÍNH XÁC TUYỆT ĐỐI về mặt cấu trúc và nội dung. ĐẶC BIỆT CHÚ Ý: Hình ảnh có thể bị xoay ngang, xoay dọc, hoặc lộn ngược. Hãy tự động 'xoay' và định hướng lại hình ảnh trong tư duy trước khi đọc. TUYỆT ĐỐI KHÔNG ĐƯỢC BỎ SÓT BẤT KỲ DÒNG HÀNG HÓA NÀO. Hãy dóng hàng ngang thật cẩn thận để đảm bảo Tên hàng hóa và Số lượng khớp nhau 100%, không bị trôi dòng, lệch dòng hay mất dòng.

HÃY THỰC HIỆN QUY TRÌNH SUY LUẬN (CHAIN OF THOUGHT):
BẮT BUỘC ghi toàn bộ quá trình phân tích của bạn vào trường `quy_trinh_suy_luan` trong file JSON. Đối với phần hàng hóa, bạn phải phân tích TỪNG DÒNG từ trên xuống dưới một cách tường minh, ghi rõ kết quả kiểm tra trực quan cho từng dòng:
- Dòng đó có nét bút viết tay nào (màu xanh, màu đen, hoặc đỏ) vẽ ngang, vẽ chéo, hoặc gạch bỏ đè lên các con số (ở cột SL, Đơn giá, hoặc Thành tiền) hay không? (Lưu ý kiểm tra kỹ xem nét vẽ tay có nằm sát sạt hoặc trùng vào dòng kẻ ngang in sẵn của bảng hay không).
- Dòng đó có dấu tick '✓' hay con số viết tay nào xác nhận số lượng nhận hay không?
Ví dụ ghi chi tiết: 'Dòng 1 (Mầm cải củ trắng): có nét bút gạch ngang đè lên các con số 3 và 36.000 -> Bị hủy -> sl_nhan = 0', 'Dòng 2 (Mầm cải củ đỏ): có dấu tick viết tay ở cột Thực Nhận -> sl_nhan = 2', 'Dòng 3 (Mầm cải ngọt): có nét bút gạch ngang viết tay đè lên số 2 và 28.000 (nằm sát dòng kẻ bảng) -> Bị hủy -> sl_nhan = 0'.
NẾU CÓ MỘT ĐƯỜNG KẺ DÀI KÉO TỪ DÒNG TRÊN XUỐNG DÒNG DƯỚI, BẠN PHẢI GHI NHẬN LÀ TẤT CẢ CÁC DÒNG ĐÓ ĐỀU BỊ GẠCH BỎ.


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

  ĐẶC BIỆT CHÚ Ý ĐỐI VỚI HÓA ĐƠN CÓ CỘT ""THỰC NHẬN"" (như hóa đơn GTech/Easymart):
  - Khi hóa đơn có cột ""THỰC NHẬN"" được kiểm nhận bằng dấu tích:
    + Những dòng được giao thường sẽ có một dấu tích ""✓"" hoặc số lượng thực nhận viết tay rất rõ ở cột ""THỰC NHẬN"".
    + Những dòng BỊ HỦY hoặc KHÔNG GIAO thường sẽ BỊ GẠCH NGANG đè lên số lượng in máy (ở cột SL hoặc cột Thành tiền), và cột ""THỰC NHẬN"" của dòng đó sẽ để TRỐNG (không có dấu tick ✓ hay số viết tay).
    + Bạn phải đối chiếu kỹ lưỡng: nếu một dòng hàng hóa có số lượng xuất in sẵn (sl_xuat > 0) nhưng cột ""THỰC NHẬN"" trống trơn VÀ có bất kỳ nét gạch ngang/gạch chéo nào viết tay đè lên số lượng/đơn giá (kể cả nét gạch cực kỳ mảnh, nằm sát/trùng dòng kẻ bảng), bạn PHẢI xác định dòng đó BỊ HỦY và gán sl_nhan = 0.

  ƯU TIÊN CAO NHẤT - Rule A: Dấu tích '✓' kèm con số viết tay rõ ràng (ví dụ: '✓ 05', '✓ 5', '✓ 10'...):
     -> ĐÂY LÀ KÝ HIỆU XÁC NHẬN GIAO THỰC TẾ. -> sl_nhan = <con số viết tay đó>.
     -> Rule này LUÔN THẮNG mọi rule khác. Dù dòng đó có bị đường kẻ đi qua, NẾU CÓ ✓ kèm số thì vẫn lấy số đó làm sl_nhan.

  ƯU TIÊN 2 - Rule B: Đường gạch bỏ / hủy giao. BẤT KỲ đường chữ 'Z', nét gạch chéo 'X', nét gạch ngang '-', nét gạch chéo '/', nét kẻ dọc '|', nét mũi tên (->), hoặc đường gạch tay chéo dài. Nét gạch CÓ THỂ vắt qua cột ĐVT, Đơn giá, hoặc Thành tiền. Chỉ cần TRÊN DÒNG ĐÓ bị nét mực gạch xuyên qua VÀ KHÔNG CÓ dấu ✓ kèm số:
     -> sl_nhan = 0.
     ĐẶC BIỆT CẢNH BÁO ĐƯỜNG KẺ NGANG TRÙNG DÒNG KẺ BẢNG: Một kiểu gạch hủy dòng phổ biến là vẽ một nét bút mực nằm ngang kéo dài cắt qua các con số ở cột Số lượng (SL), Thực Nhận, Đơn Giá và Thành Tiền (ví dụ: dòng STT 1 'Mầm cải củ trắng' và dòng STT 3 'Mầm cải ngọt' đều có một nét bút kẻ ngang đè lên các con số số lượng, đơn giá và thành tiền của dòng đó). Nét gạch ngang viết tay này có thể nằm sát hoặc trùng với dòng kẻ bảng in sẵn. Hãy kiểm tra thật kỹ các chữ số này, nếu thấy có nét mực viết tay nằm đè lên hoặc cắt ngang qua các con số đó (ngay cả khi nét vẽ rất mảnh hoặc gần trùng dòng kẻ bảng), bạn BẮT BUỘC phải xác định đó là dòng BỊ HỦY BỎ và đặt sl_nhan = 0.
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
  + Nếu có số lượng viết tay điều chỉnh bên cạnh thì sl_nhan = số viết tay đó và sl_xuat = giá trị in máy.

Bước 4: Xác định Tọa độ Vùng chữ (Bounding Box Detection) - BẮT BUỘC
- Với mỗi dòng hàng hóa được trích xuất ở Bước 2, bạn PHẢI xác định tọa độ vùng chứa dòng chữ mặt hàng đó trên hình ảnh gốc (bao gồm cả Tên hàng, Số lượng xuất, Số lượng nhận).
- Tọa độ này phải được chuẩn hóa về hệ tọa độ 0-1000 (tương ứng từ góc trên-trái [0,0] đến góc dưới-phải [1000, 1000]).
- Trả về tọa độ dưới dạng một chuỗi 'y_min,x_min,y_max,x_max' trong thuộc tính `box_coords` của từng mặt hàng. Trong đó:
  + y_min: Tọa độ cạnh trên của dòng chữ (khoảng cách từ mép trên ảnh đến dòng chữ, từ 0-1000).
  + x_min: Tọa độ cạnh trái của dòng chữ (khoảng cách từ mép trái ảnh đến dòng chữ, từ 0-1000).
  + y_max: Tọa độ cạnh dưới của dòng chữ (từ 0-1000).
  + x_max: Tọa độ cạnh phải của dòng chữ (từ 0-1000).
  Ví dụ: '245,80,270,920' (nghĩa là dòng chữ nằm ở y từ 245 đến 270, x từ 80 đến 920). Hãy đảm bảo các số này là số nguyên và phân tách nhau bằng dấu phẩy.

Bước 5: TỰ KIỂM TRA KẾT QUẢ (Self-Verification) - BẮT BUỘC
Sau khi hoàn thành Bước 2 và Bước 3, bạn PHẢI thực hiện bước kiểm tra chéo sau:
- Đếm số dòng hàng hóa mà sl_nhan == sl_xuat (tức là giao đủ 100%).
- NẾU TẤT CẢ các dòng đều có sl_nhan == sl_xuat (tỉ lệ giao đủ = 100%), bạn PHẢI tự hỏi: 'Có thật sự KHÔNG CÓ BẤT KỲ nét viết tay, dấu tick, số viết tay, nét gạch, hoặc ký hiệu bút mực nào trên TOÀN BỘ hóa đơn này không?'
  + Nếu câu trả lời là CÓ (tức là bạn NHÌN THẤY bất kỳ nét mực viết tay nào trên ảnh, dù rất mờ), bạn PHẢI quay lại Bước 2 và phân tích lại TỪNG DÒNG một cách kỹ lưỡng hơn, tập trung vào vùng cột Số lượng và khu vực bên phải mỗi dòng.
  + Chỉ khi bạn XÁC NHẬN CHẮC CHẮN rằng hóa đơn hoàn toàn chỉ có mực in máy (không có bất kỳ nét viết tay nào) thì mới được giữ nguyên kết quả 100% giao đủ.
- Ghi lại kết quả tự kiểm tra này vào cuối trường `quy_trinh_suy_luan`.
";

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
                                        { "sl_hong", new { type = "NUMBER" } },
                                        { "box_coords", new { type = "STRING", description = "Tọa độ hộp giới hạn chứa dòng chữ mặt hàng này trên ảnh gốc dưới dạng 'y_min,x_min,y_max,x_max' (giá trị từ 0-1000)." } }
                                    },
                                    required = new[] { "ten_hang", "sl_xuat", "sl_nhan", "sl_hong", "box_coords" }
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

        public static async Task<InvoiceData> CallOpenAiCompatApi(
            string base64Image,
            string mimeType,
            string endpoint,
            string apiKey,
            string modelName,
            CancellationToken token)
        {
            string promptText = @"Bạn là một hệ thống OCR và kiểm toán kế toán cao cấp chuyên nghiệp của Việt Nam. Nhiệm vụ của bạn là số hóa hóa đơn này với ĐỘ CHÍNH XÁC TUYỆT ĐỐI về mặt cấu trúc và nội dung. ĐẶC BIỆT CHÚ Ý: Hình ảnh có thể bị xoay ngang, xoay dọc, hoặc lộn ngược. Hãy tự động 'xoay' và định hướng lại hình ảnh trong tư duy trước khi đọc. TUYỆT ĐỐI KHÔNG ĐƯỢC BỎ SÓT BẤT KỲ DÒNG HÀNG HÓA NÀO. Hãy dóng hàng ngang thật cẩn thận để đảm bảo Tên hàng hóa và Số lượng khớp nhau 100%, không bị trôi dòng, lệch dòng hay mất dòng.

HÃY THỰC HIỆN QUY TRÌNH SUY LUẬN (CHAIN OF THOUGHT):
BẮT BUỘC ghi toàn bộ quá trình phân tích của bạn vào trường `quy_trinh_suy_luan` trong file JSON. Đối với phần hàng hóa, bạn phải phân tích TỪNG DÒNG từ trên xuống dưới một cách tường minh, ghi rõ kết quả kiểm tra trực quan cho từng dòng:
- Dòng đó có nét bút viết tay nào (màu xanh, màu đen, hoặc đỏ) vẽ ngang, vẽ chéo, hoặc gạch bỏ đè lên các con số (ở cột SL, Đơn giá, hoặc Thành tiền) hay không? (Lưu ý kiểm tra kỹ xem nét vẽ tay có nằm sát sạt hoặc trùng vào dòng kẻ ngang in sẵn của bảng hay không).
- Dòng đó có dấu tick '✓' hay con số viết tay nào xác nhận số lượng nhận hay không?
Ví dụ ghi chi tiết: 'Dòng 1 (Mầm cải củ trắng): có nét bút gạch ngang đè lên các con số 3 và 36.000 -> Bị hủy -> sl_nhan = 0', 'Dòng 2 (Mầm cải củ đỏ): có dấu tick viết tay ở cột Thực Nhận -> sl_nhan = 2', 'Dòng 3 (Mầm cải ngọt): có nét bút gạch ngang viết tay đè lên số 2 và 28.000 (nằm sát dòng kẻ bảng) -> Bị hủy -> sl_nhan = 0'.
NẾU CÓ MỘT ĐƯỜNG KẺ DÀI KÉO TỪ DÒNG TRÊN XUỐNG DÒNG DƯỚI, BẠN PHẢI GHI NHẬN LÀ TẤT CẢ CÁC DÒNG ĐÓ ĐỀU BỊ GẠCH BỎ.


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

  ĐẶC BIỆT CHÚ Ý ĐỐI VỚI HÓA ĐƠN CÓ CỘT ""THỰC NHẬN"" (như hóa đơn GTech/Easymart):
  - Khi hóa đơn có cột ""THỰC NHẬN"" được kiểm nhận bằng dấu tích:
    + Những dòng được giao thường sẽ có một dấu tích ""✓"" hoặc số lượng thực nhận viết tay rất rõ ở cột ""THỰC NHẬN"".
    + Những dòng BỊ HỦY hoặc KHÔNG GIAO thường sẽ BỊ GẠCH NGANG đè lên số lượng in máy (ở cột SL hoặc cột Thành tiền), và cột ""THỰC NHẬN"" của dòng đó sẽ để TRỐNG (không có dấu tick ✓ hay số viết tay).
    + Bạn phải đối chiếu kỹ lưỡng: nếu một dòng hàng hóa có số lượng xuất in sẵn (sl_xuat > 0) nhưng cột ""THỰC NHẬN"" trống trơn VÀ có bất kỳ nét gạch ngang/gạch chéo nào viết tay đè lên số lượng/đơn giá (kể cả nét gạch cực kỳ mảnh, nằm sát/trùng dòng kẻ bảng), bạn PHẢI xác định dòng đó BỊ HỦY và gán sl_nhan = 0.

  ƯU TIÊN CAO NHẤT - Rule A: Dấu tích '✓' kèm con số viết tay rõ ràng (ví dụ: '✓ 05', '✓ 5', '✓ 10'...):
     -> ĐÂY LÀ KÝ HIỆU XÁC NHẬN GIAO THỰC TẾ. -> sl_nhan = <con số viết tay đó>.
     -> Rule này LUÔN THẮNG mọi rule khác. Dù dòng đó có bị đường kẻ đi qua, NẾU CÓ ✓ kèm số thì vẫn lấy số đó làm sl_nhan.

  ƯU TIÊN 2 - Rule B: Đường gạch bỏ / hủy giao. BẤT KỲ đường chữ 'Z', nét gạch chéo 'X', nét gạch ngang '-', nét gạch chéo '/', nét kẻ dọc '|', nét mũi tên (->), hoặc đường gạch tay chéo dài. Nét gạch CÓ THỂ vắt qua cột ĐVT, Đơn giá, hoặc Thành tiền. Chỉ cần TRÊN DÒNG ĐÓ bị nét mực gạch xuyên qua VÀ KHÔNG CÓ dấu ✓ kèm số:
     -> sl_nhan = 0.
     ĐẶC BIỆT CẢNH BÁO ĐƯỜNG KẺ NGANG TRÙNG DÒNG KẺ BẢNG: Một kiểu gạch hủy dòng phổ biến là vẽ một nét bút mực nằm ngang kéo dài cắt qua các con số ở cột Số lượng (SL), Thực Nhận, Đơn Giá và Thành Tiền (ví dụ: dòng STT 1 'Mầm cải củ trắng' và dòng STT 3 'Mầm cải ngọt' đều có một nét bút kẻ ngang đè lên các con số số lượng, đơn giá và thành tiền của dòng đó). Nét gạch ngang viết tay này có thể nằm sát hoặc trùng với dòng kẻ bảng in sẵn. Hãy kiểm tra thật kỹ các chữ số này, nếu thấy có nét mực viết tay nằm đè lên hoặc cắt ngang qua các con số đó (ngay cả khi nét vẽ rất mảnh hoặc gần trùng dòng kẻ bảng), bạn BẮT BUỘC phải xác định đó là dòng BỊ HỦY BỎ và đặt sl_nhan = 0.
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
  + Nếu có số lượng viết tay điều chỉnh bên cạnh thì sl_nhan = số viết tay đó và sl_xuat = giá trị in máy.

Bước 4: Xác định Tọa độ Vùng chữ (Bounding Box Detection) - BẮT BUỘC
- Với mỗi dòng hàng hóa được trích xuất ở Bước 2, bạn PHẢI xác định tọa độ vùng chứa dòng chữ mặt hàng đó trên hình ảnh gốc (bao gồm cả Tên hàng, Số lượng xuất, Số lượng nhận).
- Tọa độ này phải được chuẩn hóa về hệ tọa độ 0-1000 (tương ứng từ góc trên-trái [0,0] đến góc dưới-phải [1000, 1000]).
- Trả về tọa độ dưới dạng một chuỗi 'y_min,x_min,y_max,x_max' trong thuộc tính `box_coords` của từng mặt hàng. Trong đó:
  + y_min: Tọa độ cạnh trên của dòng chữ (khoảng cách từ mép trên ảnh đến dòng chữ, từ 0-1000).
  + x_min: Tọa độ cạnh trái của dòng chữ (khoảng cách từ mép trái ảnh đến dòng chữ, từ 0-1000).
  + y_max: Tọa độ cạnh dưới của dòng chữ (từ 0-1000).
  + x_max: Tọa độ cạnh phải của dòng chữ (từ 0-1000).
  Ví dụ: '245,80,270,920' (nghĩa là dòng chữ nằm ở y từ 245 đến 270, x từ 80 đến 920). Hãy đảm bảo các số này là số nguyên và phân tách nhau bằng dấu phẩy.

Bước 5: TỰ KIỂM TRA KẾT QUẢ (Self-Verification) - BẮT BUỘC
Sau khi hoàn thành Bước 2 và Bước 3, bạn PHẢI thực hiện bước kiểm tra chéo sau:
- Đếm số dòng hàng hóa mà sl_nhan == sl_xuat (tức là giao đủ 100%).
- NẾU TẤT CẢ các dòng đều có sl_nhan == sl_xuat (tỉ lệ giao đủ = 100%), bạn PHẢI tự hỏi: 'Có thật sự KHÔNG CÓ BẤT KỲ nét viết tay, dấu tick, số viết tay, nét gạch, hoặc ký hiệu bút mực nào trên TOÀN BỘ hóa đơn này không?'
  + Nếu câu trả lời là CÓ (tức là bạn NHÌN THẤY bất kỳ nét mực viết tay nào trên ảnh, dù rất mờ), bạn PHẢI quay lại Bước 2 và phân tích lại TỪNG DÒNG một cách kỹ lưỡng hơn, tập trung vào vùng cột Số lượng và khu vực bên phải mỗi dòng.
  + Chỉ khi bạn XÁC NHẬN CHẮC CHẮN rằng hóa đơn hoàn toàn chỉ có mực in máy (không có bất kỳ nét viết tay nào) thì mới được giữ nguyên kết quả 100% giao đủ.
- Ghi lại kết quả tự kiểm tra này vào cuối trường `quy_trinh_suy_luan`.

YÊU CẦU ĐỊNH DẠNG ĐẦU RA (OUTPUT FORMAT):
Bạn PHẢI trả về DUY NHẤT một chuỗi JSON hợp lệ (không kèm markdown ```json hay giải thích ngoài JSON), tuân thủ chính xác cấu trúc sau:
{
  ""quy_trinh_suy_luan"": ""Trình bày chi tiết phân tích từng dòng..."",
  ""ngay_giao"": ""dd/MM/yyyy"",
  ""khach_hang"": ""Tên khách hàng"",
  ""diem_giao"": ""Điểm giao"",
  ""danh_sach_hang_hoa"": [
    {
      ""ten_hang"": ""Tên mặt hàng"",
      ""sl_xuat"": 10,
      ""sl_nhan"": 10,
      ""sl_hong"": 0,
      ""box_coords"": ""y_min,x_min,y_max,x_max""
    }
  ]
}
";

            // Normalize endpoint URL
            string url = endpoint.Trim().TrimEnd('/');
            if (!url.EndsWith("/chat/completions", StringComparison.OrdinalIgnoreCase))
            {
                if (!url.EndsWith("/v1", StringComparison.OrdinalIgnoreCase))
                {
                    url += "/v1";
                }
                url += "/chat/completions";
            }

            var requestBody = new
            {
                model = modelName,
                messages = new object[]
                {
                    new
                    {
                        role = "user",
                        content = new object[]
                        {
                            new { type = "text", text = promptText },
                            new
                            {
                                type = "image_url",
                                image_url = new
                                {
                                    url = $"data:{mimeType};base64,{base64Image}"
                                }
                            }
                        }
                    }
                },
                response_format = new { type = "json_object" },
                temperature = 0.1
            };

            string jsonBody = JsonSerializer.Serialize(requestBody);

            using var request = new HttpRequestMessage(HttpMethod.Post, url);
            request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);
            request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

            using var response = await _httpClient.SendAsync(request, token);
            string responseStr = await response.Content.ReadAsStringAsync(token);

            if (!response.IsSuccessStatusCode)
            {
                Logger.Log($"[OpenAI API Lỗi] HTTP {response.StatusCode}: {responseStr}");
                string errMsg = $"OpenAI API trả về lỗi {(int)response.StatusCode}";
                try
                {
                    using var errDoc = JsonDocument.Parse(responseStr);
                    if (errDoc.RootElement.TryGetProperty("error", out var errObj))
                    {
                        if (errObj.TryGetProperty("message", out var msgElem))
                        {
                            errMsg += $": {msgElem.GetString()}";
                        }
                    }
                }
                catch { }
                throw new Exception(errMsg);
            }

            using (JsonDocument doc = JsonDocument.Parse(responseStr))
            {
                var root = doc.RootElement;
                if (root.TryGetProperty("choices", out var choices) && choices.GetArrayLength() > 0)
                {
                    var firstChoice = choices[0];
                    if (firstChoice.TryGetProperty("message", out var message) && message.TryGetProperty("content", out var contentElem))
                    {
                        string text = contentElem.GetString() ?? "";
                        Logger.Log($"[OpenAI API Result]\n{text}\n-----------------------");

                        // Strip markdown code fences if model returned ```json ... ```
                        string cleanJson = text.Trim();
                        if (cleanJson.StartsWith("```json", StringComparison.OrdinalIgnoreCase))
                        {
                            cleanJson = cleanJson.Substring(7);
                        }
                        else if (cleanJson.StartsWith("```"))
                        {
                            cleanJson = cleanJson.Substring(3);
                        }
                        if (cleanJson.EndsWith("```"))
                        {
                            cleanJson = cleanJson.Substring(0, cleanJson.Length - 3);
                        }
                        cleanJson = cleanJson.Trim();

                        // Extract JSON substring if surrounded by extra text
                        int firstBrace = cleanJson.IndexOf('{');
                        int lastBrace = cleanJson.LastIndexOf('}');
                        if (firstBrace >= 0 && lastBrace > firstBrace)
                        {
                            cleanJson = cleanJson.Substring(firstBrace, lastBrace - firstBrace + 1);
                        }

                        InvoiceData data = JsonSerializer.Deserialize<InvoiceData>(cleanJson, _jsonOptions);
                        return data;
                    }
                    else
                    {
                        throw new Exception("OpenAI API không trả về nội dung tin nhắn (message.content trống).");
                    }
                }
                else
                {
                    throw new Exception("OpenAI API không trả về choices nào.");
                }
            }
        }

        private void ProviderSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (GeminiModelPanel == null || OpenAiConfigPanel == null) return;

            bool isOpenAi = ProviderSelector.SelectedIndex == 1;
            GeminiModelPanel.Visibility = isOpenAi ? Visibility.Collapsed : Visibility.Visible;
            OpenAiConfigPanel.Visibility = isOpenAi ? Visibility.Visible : Visibility.Collapsed;
            if (SettingsBtn != null) SettingsBtn.Visibility = isOpenAi ? Visibility.Collapsed : Visibility.Visible;
            _useOpenAiCompat = isOpenAi;

            if (_isInitialized)
            {
                SaveOpenAiConfig();
            }
        }

        private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.Source != MainTabControl) return;
            if (_isSwitchingTab) return;
            if (!_isInitialized) return;

            // Block navigating to any other tab if no Excel file has been loaded
            if (MainTabControl.SelectedIndex > 0)
            {
                if (!EnsureExcelLoaded())
                {
                    _isSwitchingTab = true;
                    try
                    {
                        MainTabControl.SelectedIndex = 0;
                    }
                    finally
                    {
                        _isSwitchingTab = false;
                    }
                    return;
                }
            }

            // When switching to "Xử lý OCR" tab (index 1), restore saved provider config
            if (MainTabControl.SelectedIndex == 1)
            {
                LoadOpenAiConfig();
            }
        }

        private void LoadOpenAiConfig()
        {
            try
            {
                if (File.Exists(OpenAiConfigFilePath))
                {
                    string json = File.ReadAllText(OpenAiConfigFilePath);
                    var cfg = JsonSerializer.Deserialize<OpenAiConfig>(json, _jsonOptions);
                    if (cfg != null)
                    {
                        if (OpenAiEndpointBox != null) OpenAiEndpointBox.Text = cfg.Endpoint ?? "";
                        if (OpenAiKeyBox != null) OpenAiKeyBox.Text = cfg.ApiKey ?? "";
                        if (OpenAiModelBox != null && !string.IsNullOrWhiteSpace(cfg.ModelName))
                            OpenAiModelBox.Text = cfg.ModelName;
                        if (OpenAiConcurrencyBox != null && cfg.Concurrency > 0)
                            OpenAiConcurrencyBox.Text = cfg.Concurrency.ToString();

                        if (ProviderSelector != null && !string.IsNullOrWhiteSpace(cfg.Provider))
                        {
                            bool matched = false;
                            for (int i = 0; i < ProviderSelector.Items.Count; i++)
                            {
                                if (ProviderSelector.Items[i] is ComboBoxItem item &&
                                    string.Equals(item.Content?.ToString(), cfg.Provider, StringComparison.OrdinalIgnoreCase))
                                {
                                    ProviderSelector.SelectedIndex = i;
                                    matched = true;
                                    break;
                                }
                            }
                            if (!matched)
                            {
                                ProviderSelector.SelectedIndex = 0;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[OpenAI Config] Lỗi tải file {OpenAiConfigFilePath}: {ex.Message}");
            }
        }

        private void SaveOpenAiConfig()
        {
            try
            {
                int concurrency = 15;
                if (OpenAiConcurrencyBox != null && int.TryParse(OpenAiConcurrencyBox.Text?.Trim(), out int c) && c > 0)
                {
                    concurrency = c;
                }

                string currentProvider = (ProviderSelector?.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Gemini";

                var cfg = new OpenAiConfig
                {
                    Provider = currentProvider,
                    Endpoint = OpenAiEndpointBox?.Text?.Trim() ?? "",
                    ApiKey = OpenAiKeyBox?.Text?.Trim() ?? "",
                    ModelName = OpenAiModelBox?.Text?.Trim() ?? "gpt-4o",
                    Concurrency = concurrency
                };
                string json = JsonSerializer.Serialize(cfg, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(OpenAiConfigFilePath, json);
            }
            catch (Exception ex)
            {
                Logger.Log($"[OpenAI Config] Lỗi lưu file {OpenAiConfigFilePath}: {ex.Message}");
            }
        }

        private void OpenAiDecConcBtn_Click(object sender, RoutedEventArgs e)
        {
            if (OpenAiConcurrencyBox == null) return;
            if (int.TryParse(OpenAiConcurrencyBox.Text?.Trim(), out int val))
            {
                if (val > 1) OpenAiConcurrencyBox.Text = (val - 1).ToString();
            }
            else
            {
                OpenAiConcurrencyBox.Text = "15";
            }
        }

        private void OpenAiIncConcBtn_Click(object sender, RoutedEventArgs e)
        {
            if (OpenAiConcurrencyBox == null) return;
            if (int.TryParse(OpenAiConcurrencyBox.Text?.Trim(), out int val))
            {
                if (val < 30) OpenAiConcurrencyBox.Text = (val + 1).ToString();
            }
            else
            {
                OpenAiConcurrencyBox.Text = "15";
            }
        }

        private void OpenAiConcurrencyBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (OpenAiConcurrencyBox == null) return;
            if (int.TryParse(OpenAiConcurrencyBox.Text?.Trim(), out int val))
            {
                OpenAiConcurrencyBox.Text = Math.Clamp(val, 1, 30).ToString();
            }
            else
            {
                OpenAiConcurrencyBox.Text = "15";
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
            UpdateBatchActionBar();
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

        private void ShipperSelector_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete && ShipperSelector.IsDropDownOpen)
            {
                var selectedShipper = ShipperSelector.SelectedItem as string;
                if (!string.IsNullOrEmpty(selectedShipper))
                {
                    var result = MessageBox.Show($"Bạn có chắc chắn muốn xóa người giao hàng '{selectedShipper}' khỏi danh sách gợi ý?", "Xác nhận xóa", MessageBoxButton.YesNo);
                    if (result == MessageBoxResult.Yes)
                    {
                        _shippers.Remove(selectedShipper);
                        SaveShippers();
                        e.Handled = true;
                    }
                }
            }
        }

        private void DeleteShipperItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is string shipper)
            {
                e.Handled = true; // Prevent click from bubbling up and selecting the item/closing dropdown

                var result = MessageBox.Show($"Bạn có chắc chắn muốn xóa người giao hàng '{shipper}' khỏi danh sách gợi ý?", "Xác nhận xóa", MessageBoxButton.YesNo);
                if (result == MessageBoxResult.Yes)
                {
                    _shippers.Remove(shipper);
                    SaveShippers();
                }
            }
        }

        private void LoadExcelConfigs()
        {
            string backupPath = ConfigsFilePath + ".bak";
            bool loadedSuccessfully = false;

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
                        loadedSuccessfully = true;
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"[App] Lỗi tải file excel_configs.json (Thử khôi phục từ backup): {ex.Message}");
                }
            }

            if (!loadedSuccessfully && File.Exists(backupPath))
            {
                try
                {
                    string json = File.ReadAllText(backupPath);
                    var list = JsonSerializer.Deserialize<List<ExcelConfigItem>>(json);
                    if (list != null)
                    {
                        _excelConfigs.Clear();
                        foreach (var item in list) _excelConfigs.Add(item);

                        // Heal the primary file by writing the recovered config back to it
                        File.WriteAllText(ConfigsFilePath, json);
                        Logger.Log("[App] Đã phục hồi thành công file excel_configs.json từ tệp dự phòng (.bak).");
                        loadedSuccessfully = true;
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"[App] Lỗi tải file backup excel_configs.json.bak: {ex.Message}");
                }
            }
        }

        private void SaveExcelConfigs()
        {
            try
            {
                string json = JsonSerializer.Serialize(_excelConfigs.ToList(), new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(ConfigsFilePath, json);

                // Write backup copy
                try
                {
                    string backupPath = ConfigsFilePath + ".bak";
                    File.WriteAllText(backupPath, json);
                }
                catch { }
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
            ChangeSheetsBtn.IsEnabled = !string.IsNullOrEmpty(RefFilePathText.Text) && 
                                        RefFilePathText.Text != "Chưa chọn file Excel tham chiếu." && 
                                        File.Exists(RefFilePathText.Text);
        }

        private void SetActiveExcelConfig(string localPath)
        {
            foreach (var cfg in _excelConfigs)
            {
                cfg.IsActive = cfg.LocalPath.Equals(localPath, StringComparison.OrdinalIgnoreCase);
            }
        }

        private bool EnsureExcelLoaded()
        {
            var activeConfig = _excelConfigs.FirstOrDefault(x => x.IsActive);
            bool isLoaded = activeConfig != null 
                && !string.IsNullOrWhiteSpace(activeConfig.LocalPath) 
                && File.Exists(activeConfig.LocalPath)
                && !string.IsNullOrEmpty(RefFilePathText?.Text) 
                && RefFilePathText.Text != "Chưa chọn file Excel tham chiếu.";

            if (!isLoaded)
            {
                MessageBox.Show(
                    "Vui lòng nạp hoặc chọn một file Excel tại Trang chủ trước khi thực hiện thao tác!",
                    "Chưa nạp file Excel",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return false;
            }
            return true;
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
                    var sheetNames = await Task.Run(() => GetSheetNames(targetPath));
                    RefFilePathText.Text = targetPath;
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

                        // Set active status
                        SetActiveExcelConfig(targetPath);
                        SaveExcelConfigs();

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
                    RefFilePathText.Text = config.LocalPath;

                    SelectedRefSheetText.Text = config.RefSheet;
                    SelectedSalesSheetText.Text = config.SalesSheet;
                    ChangeSheetsBtn.IsEnabled = true;

                    // Set active status
                    SetActiveExcelConfig(config.LocalPath);
                    SaveExcelConfigs();

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
                    // If the configuration being deleted is the active one, reset UI fields
                    if (RefFilePathText.Text == config.LocalPath)
                    {
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

            DeleteAssociatedInvoiceImages(config.FileName);
            _excelConfigs.Remove(config);
        }

        private void DeleteAssociatedInvoiceImages(string excelFileName)
        {
            if (string.IsNullOrEmpty(excelFileName)) return;
            try
            {
                var filesToDelete = new List<string>();
                using (var connection = new DuckDBConnection($"Data Source={StagingDbPath}"))
                {
                    connection.Open();
                    using (var cmd = connection.CreateCommand())
                    {
                        cmd.CommandText = "SELECT DISTINCT file_name FROM staging_items WHERE excel_file_name = $excelName";
                        cmd.Parameters.Add(new DuckDBParameter("$excelName", excelFileName));
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string imgName = reader.GetString(0);
                                if (!string.IsNullOrEmpty(imgName))
                                {
                                    filesToDelete.Add(imgName);
                                }
                            }
                        }
                    }
                }

                string rawDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, RawMaterialDirName);
                foreach (var imgName in filesToDelete)
                {
                    string fullPath = Path.Combine(rawDir, imgName);
                    if (File.Exists(fullPath))
                    {
                        try { File.Delete(fullPath); } catch {}
                    }
                }

                // Delete DB staging records too
                using (var connection = new DuckDBConnection($"Data Source={StagingDbPath}"))
                {
                    connection.Open();
                    using (var cmd = connection.CreateCommand())
                    {
                        cmd.CommandText = "DELETE FROM staging_items WHERE excel_file_name = $excelName";
                        cmd.Parameters.Add(new DuckDBParameter("$excelName", excelFileName));
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[Dọn dẹp ảnh] Lỗi khi dọn dẹp ảnh cho file {excelFileName}: {ex.Message}");
            }
        }

        private async void ChangeSheets_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureExcelLoaded()) return;

            string filePath = RefFilePathText.Text;
            if (string.IsNullOrEmpty(filePath) || filePath == "Chưa chọn file Excel tham chiếu." || !File.Exists(filePath))
                return;

            ShowRefLoading("Đang đọc danh sách sheet...");
            List<string> sheetNames;
            try
            {
                sheetNames = await Task.Run(() => GetSheetNames(filePath));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Không thể đọc file Excel: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                HideRefLoading();
                return;
            }
            HideRefLoading();

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

            string filePath = RefFilePathText.Text;
            if (string.IsNullOrEmpty(filePath) || filePath == "Chưa chọn file Excel tham chiếu." || !File.Exists(filePath))
            {
                MessageBox.Show("Đường dẫn file Excel không hợp lệ hoặc file không tồn tại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                HideRefLoading();
                return;
            }

            try
            {
                // Unbind ItemsSource and clear collections to prevent UI rendering overhead
                RefDataGrid.ItemsSource = null;
                SalesDataGrid.ItemsSource = null;
                _refItems.Clear();

                var result = await Task.Run(() =>
                {
                    DataSet dataset;
                    using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            dataset = reader.AsDataSet();
                        }
                    }

                    var resultList = new List<ReferenceItem>();
                    if (dataset.Tables.Contains(refSheetName))
                    {
                        var refTable = dataset.Tables[refSheetName];
                        if (refTable != null)
                        {
                            // Find header row in first 20 rows
                            int headerRowIndex = -1;
                            int colAbbrevIndex = -1;
                            int colNameIndex = -1;
                            int colUnitIndex = -1;

                            int maxRowsToSearch = Math.Min(20, refTable.Rows.Count);
                            for (int r = 0; r < maxRowsToSearch; r++)
                            {
                                var row = refTable.Rows[r];
                                for (int c = 0; c < refTable.Columns.Count; c++)
                                {
                                    string val = row[c]?.ToString()?.Trim()?.ToLower() ?? "";
                                    if (val.Contains("viết tắt") || val == "viet tat" || val == "ma" || val == "mã")
                                    {
                                        colAbbrevIndex = c;
                                    }
                                    else if (val.Contains("tên hàng") || val == "ten hang" || val == "sản phẩm" || val == "san pham")
                                    {
                                        colNameIndex = c;
                                    }
                                    else if (val.Contains("đơn vị") || val == "don vi" || val == "đvt" || val == "dvt")
                                    {
                                        colUnitIndex = c;
                                    }
                                }

                                if (colAbbrevIndex != -1 && colNameIndex != -1 && colUnitIndex != -1)
                                {
                                    headerRowIndex = r;
                                    break;
                                }
                            }

                            if (headerRowIndex == -1)
                            {
                                headerRowIndex = 1; // Default assume row index 1 (which is row 2)
                                colAbbrevIndex = 0;
                                colNameIndex = 1;
                                colUnitIndex = 2;
                            }

                            int lastRowIndex = Math.Min(refTable.Rows.Count - 1, headerRowIndex + 1000);
                            for (int r = headerRowIndex + 1; r <= lastRowIndex; r++)
                            {
                                var row = refTable.Rows[r];
                                if (colAbbrevIndex >= refTable.Columns.Count || colNameIndex >= refTable.Columns.Count || colUnitIndex >= refTable.Columns.Count)
                                    continue;

                                string abbrev = row[colAbbrevIndex]?.ToString()?.Trim() ?? "";
                                string name = row[colNameIndex]?.ToString()?.Trim() ?? "";
                                string unit = row[colUnitIndex]?.ToString()?.Trim() ?? "";

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
                        }
                    }

                    var dt = new DataTable();
                    if (dataset.Tables.Contains(salesSheetName))
                    {
                        var salesTable = dataset.Tables[salesSheetName];
                        if (salesTable != null && salesTable.Rows.Count > 0)
                        {
                            int colCount = salesTable.Columns.Count;
                            int rowCount = salesTable.Rows.Count;

                            // Row 0 of salesTable is the header row
                            var firstRow = salesTable.Rows[0];
                            for (int col = 0; col < colCount; col++)
                            {
                                string header = firstRow[col]?.ToString()?.Trim() ?? "";
                                if (string.IsNullOrEmpty(header))
                                {
                                    header = GetExcelColumnName(col + 1);
                                }

                                string uniqueHeader = header;
                                int counter = 1;
                                while (dt.Columns.Contains(uniqueHeader))
                                {
                                    uniqueHeader = $"{header}_{counter++}";
                                }
                                dt.Columns.Add(uniqueHeader);
                            }

                            // Rows 1 to N
                            for (int r = 1; r < rowCount; r++)
                            {
                                var row = salesTable.Rows[r];
                                var dr = dt.NewRow();
                                for (int col = 0; col < colCount; col++)
                                {
                                    var val = row[col];
                                    if (val is DateTime dtValue)
                                    {
                                        if (dtValue.TimeOfDay == TimeSpan.Zero)
                                        {
                                            dr[col] = dtValue.ToString("dd/MM/yyyy");
                                        }
                                        else
                                        {
                                            dr[col] = dtValue.ToString("dd/MM/yyyy HH:mm:ss");
                                        }
                                    }
                                    else
                                    {
                                        dr[col] = val?.ToString() ?? "";
                                    }
                                }
                                dt.Rows.Add(dr);
                            }
                        }
                    }

                    return (resultList, dt);
                });

                foreach (var item in result.Item1)
                {
                    _refItems.Add(item);
                }

                SalesDataGrid.ItemsSource = result.Item2.DefaultView;
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
            lock (DbWriteLock)
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
                                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                                    match_score DOUBLE DEFAULT 1.0
                                );";
                            cmd.ExecuteNonQuery();

                            // Migrate existing database to add column if not exists
                            cmd.CommandText = "ALTER TABLE staging_items ADD COLUMN IF NOT EXISTS match_score DOUBLE DEFAULT 1.0;";
                            cmd.ExecuteNonQuery();

                            cmd.CommandText = "ALTER TABLE staging_items ADD COLUMN IF NOT EXISTS sync_batch_id VARCHAR;";
                            cmd.ExecuteNonQuery();

                            cmd.CommandText = "ALTER TABLE staging_items ADD COLUMN IF NOT EXISTS synced_at TIMESTAMP;";
                            cmd.ExecuteNonQuery();

                            cmd.CommandText = "ALTER TABLE staging_items ADD COLUMN IF NOT EXISTS excel_file_name VARCHAR;";
                            cmd.ExecuteNonQuery();

                            cmd.CommandText = "ALTER TABLE staging_items ADD COLUMN IF NOT EXISTS box_coords VARCHAR;";
                            cmd.ExecuteNonQuery();

                            // Clean up items synced > 30 days ago and reclaim physical disk space
                            try
                            {
                                cmd.CommandText = "DELETE FROM staging_items WHERE status = 'Synced' AND created_at < now() - INTERVAL '30 days';";
                                int deletedCount = cmd.ExecuteNonQuery();
                                if (deletedCount > 0)
                                {
                                    Logger.Log($"[DuckDB Dọn Dẹp] Đã xóa {deletedCount} dòng staging đã đồng bộ quá 30 ngày.");
                                    cmd.CommandText = "VACUUM;";
                                    cmd.ExecuteNonQuery();
                                }
                            }
                            catch (Exception dbEx)
                            {
                                Logger.Log($"[DuckDB Dọn Dẹp Lỗi] {dbEx.Message}");
                            }
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
        }

        private void LoadStagingItemsFromDb()
        {
            _stagingItems.Clear();
            try
            {
                lock (DbWriteLock)
                {
                    string connStr = $"Data Source={StagingDbPath}";
                    using (var connection = new DuckDBConnection(connStr))
                    {
                        connection.Open();
                        using (var cmd = connection.CreateCommand())
                        {
                            cmd.CommandText = "SELECT id, file_name, ngay_giao, khach_hang, diem_giao, nguoi_giao, ten_hang_goc, ten_hang_khop, hst, don_vi, sl_xuat, sl_nhan, sl_hong, ghi_chu, match_score, box_coords FROM staging_items WHERE status != 'Synced' ORDER BY created_at DESC;";
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
                                        GhiChu = reader.IsDBNull(13) ? "" : reader.GetString(13),
                                        MatchScore = reader.IsDBNull(14) ? 1.0 : reader.GetDouble(14),
                                        BoxCoords = reader.IsDBNull(15) ? "" : reader.GetString(15)
                                    };
                                    _stagingItems.Add(item);
                                }
                            }
                        }
                    }
                }
                RunAuditRules();
                _stagingItemsView = CollectionViewSource.GetDefaultView(_stagingItems);
                _stagingItemsView.Filter = PreviewStagingFilter;
                PreviewDataGrid.ItemsSource = _stagingItemsView;
                UpdatePreviewFilterLists();
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
                lock (DbWriteLock)
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
                                    ghi_chu = $ghi_chu,
                                    match_score = $match_score,
                                    box_coords = $box_coords
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
                            cmd.Parameters.Add(new DuckDBParameter("match_score", item.MatchScore));
                            cmd.Parameters.Add(new DuckDBParameter("box_coords", item.BoxCoords ?? ""));
                            cmd.Parameters.Add(new DuckDBParameter("id", item.Id));

                            cmd.ExecuteNonQuery();
                        }
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
            if (!EnsureExcelLoaded()) return;

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

            // Cache pre-cleaned reference strings to avoid redundant sign-stripping and allocation loops
            var cachedRefItems = targetRefItems.Select(x => new CachedRefInfo
            {
                Item = x,
                CleanTenHang = FuzzyMatcher.RemoveSign4VietnameseString(x.TenHang.Trim().ToLower()),
                CleanVietTat = FuzzyMatcher.RemoveSign4VietnameseString(x.VietTat.Trim().ToLower())
            }).ToList();

            ShowRefLoading("Đang đối soát sản phẩm và lưu vào DuckDB...");
            
            try
            {
                await Task.Run(() =>
                {
                    // Phase 1: Fuzzy match parallel (CPU-bound, no DB access)
                    var matchResults = new ConcurrentBag<(string fileName, InvoiceData data, InvoiceItem itemRow,
                        string tenKhop, string hst, string donVi, double bestScore)>();

                    Parallel.ForEach(processedSuccessItems, item =>
                    {
                        string fileName = Path.GetFileName(item.FilePath);
                        var data = item.Data;
                        if (data?.danh_sach_hang_hoa == null) return;

                        foreach (var itemRow in data.danh_sach_hang_hoa)
                        {
                            var matchedItem = FuzzyMatcher.FindBestMatchOptimized(itemRow.ten_hang, cachedRefItems, out double bestScore);

                            string tenKhop = "";
                            string hst = "";
                            string donVi = "";

                            if (matchedItem != null && bestScore > 0.5)
                            {
                                tenKhop = matchedItem.TenHang;
                                hst = matchedItem.VietTat;
                                donVi = matchedItem.DonViTinh;
                            }
                            matchResults.Add((fileName, data, itemRow, tenKhop, hst, donVi, bestScore));
                        }
                    });

                    // Phase 2: Batch DuckDB insert under transaction & lock (DbWriteLock)
                    try
                    {
                        lock (DbWriteLock)
                        {
                            using (var conn = new DuckDBConnection($"Data Source={StagingDbPath}"))
                            {
                                conn.Open();
                                using (var tx = conn.BeginTransaction())
                                {
                                    foreach (var r in matchResults)
                                    {
                                        using (var cmd = conn.CreateCommand())
                                        {
                                            cmd.CommandText = @"
                                                INSERT INTO staging_items (id, file_name, ngay_giao, khach_hang, diem_giao, nguoi_giao, ten_hang_goc, ten_hang_khop, hst, don_vi, sl_xuat, sl_nhan, sl_hong, ghi_chu, status, match_score, box_coords) 
                                                VALUES ($id, $file_name, $ngay_giao, $khach_hang, $diem_giao, $nguoi_giao, $ten_hang_goc, $ten_hang_khop, $hst, $don_vi, $sl_xuat, $sl_nhan, $sl_hong, $ghi_chu, $status, $match_score, $box_coords);";

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
                                            cmd.Parameters.Add(new DuckDBParameter("match_score", r.bestScore));
                                            cmd.Parameters.Add(new DuckDBParameter("box_coords", r.itemRow.box_coords ?? ""));

                                            cmd.ExecuteNonQuery();
                                        }
                                    }
                                    tx.Commit();
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
                    lock (DbWriteLock)
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
            string header = e.Column.Header.ToString().ToLower();
            if (header.Contains("sl xuất") || header.Contains("sl nhận"))
            {
                var textBox = e.EditingElement as TextBox;
                if (textBox != null)
                {
                    string newText = textBox.Text.Trim();
                    if (!double.TryParse(newText, out double val) || val < 0)
                    {
                        MessageBox.Show("Số lượng phải là một số không âm hợp lệ.", "Lỗi nhập liệu", MessageBoxButton.OK, MessageBoxImage.Warning);
                        if (e.Row.Item is StagingItem stagingItem)
                        {
                            if (header.Contains("sl xuất"))
                                textBox.Text = stagingItem.SlXuat.ToString();
                            else
                                textBox.Text = stagingItem.SlNhan.ToString();
                        }
                    }
                }
            }

            if (e.EditAction == DataGridEditAction.Commit && e.Row.Item is StagingItem stagingItem2)
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    // Recalculate SL Hỏng
                    stagingItem2.SlHong = stagingItem2.SlXuat - stagingItem2.SlNhan;

                    // Automatically update hst and donVi if name is matched manually
                    if (e.Column.Header.ToString().Contains("Tên khớp"))
                    {
                        var refMatch = _refItems.FirstOrDefault(x => x.TenHang.Equals(stagingItem2.TenHangKhop, StringComparison.OrdinalIgnoreCase));
                        if (refMatch != null)
                        {
                            stagingItem2.Hst = refMatch.VietTat;
                            stagingItem2.DonVi = refMatch.DonViTinh;
                        }
                    }

                    // User manually modified the row -> set MatchScore to 1.0
                    stagingItem2.MatchScore = 1.0;

                    UpdateStagingItemInDb(stagingItem2);
                    RunAuditRules();
                }), System.Windows.Threading.DispatcherPriority.Background);
            }
        }

        private void ProductsGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            string header = e.Column.Header.ToString().ToLower();
            if (header.Contains("sl xuất") || header.Contains("sl nhận"))
            {
                var textBox = e.EditingElement as TextBox;
                if (textBox != null)
                {
                    string newText = textBox.Text.Trim();
                    if (!double.TryParse(newText, out double val) || val < 0)
                    {
                        MessageBox.Show("Số lượng phải là một số không âm hợp lệ.", "Lỗi nhập liệu", MessageBoxButton.OK, MessageBoxImage.Warning);
                        if (e.Row.Item is InvoiceItem item)
                        {
                            if (header.Contains("sl xuất"))
                                textBox.Text = item.sl_xuat.ToString();
                            else
                                textBox.Text = item.sl_nhan.ToString();
                        }
                    }
                }
            }
        }

        private async void CommitToExcelBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureExcelLoaded()) return;

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
            catch (System.IO.IOException)
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

                    // Update status in DuckDB staging to 'Synced' — batch single statement under DbWriteLock
                    lock (DbWriteLock)
                    {
                        using (var conn = new DuckDBConnection($"Data Source={StagingDbPath}"))
                        {
                            conn.Open();
                            using (var cmd = conn.CreateCommand())
                            {
                                var idList = string.Join(", ", itemsToCommit.Select(x => $"'{x.Id}'"));
                                string batchId = Guid.NewGuid().ToString();
                                cmd.CommandText = $"UPDATE staging_items SET status = 'Synced', sync_batch_id = $batchId, synced_at = now(), excel_file_name = $excelFileName WHERE id IN ({idList});";
                                cmd.Parameters.Add(new DuckDBParameter("batchId", batchId));
                                cmd.Parameters.Add(new DuckDBParameter("excelFileName", activeConfig.FileName));
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                });

                MessageBox.Show($"Đã ghi sổ thành công {itemsToCommit.Count} dòng vào Excel!", "Thành công", MessageBoxButton.OK, MessageBoxImage.Information);
                
                // Reload preview grid (synced items are removed)
                LoadStagingItemsFromDb();


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
            if (!EnsureExcelLoaded()) return;

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

                    var exportImagesResult = MessageBox.Show(
                        "Bạn có muốn xuất kèm thư mục chứa hình ảnh hóa đơn đã đối soát (được đổi tên theo hóa đơn) không?",
                        "Xuất kèm ảnh hóa đơn",
                        MessageBoxButton.YesNo
                    );

                    if (exportImagesResult == MessageBoxResult.Yes)
                    {
                        ExportInvoiceImages(activeConfig.FileName, saveFileDialog.FileName);
                        MessageBox.Show(
                            $"Đã xuất tệp Excel và các hình ảnh hóa đơn kèm theo thành công tại:\n{saveFileDialog.FileName}\n\nThư mục ảnh: {Path.GetFileNameWithoutExtension(saveFileDialog.FileName)}_Images",
                            "Thành công",
                            MessageBoxButton.OK
                        );
                    }
                    else
                    {
                        MessageBox.Show($"Đã xuất tệp Excel thành công tại:\n{saveFileDialog.FileName}", "Thành công", MessageBoxButton.OK);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Lỗi khi xuất tệp Excel: {ex.Message}", "Lỗi", MessageBoxButton.OK);
                }
            }
        }

        private void ExportInvoiceImages(string excelFileName, string destExcelPath)
        {
            try
            {
                var imageMapping = new List<(string OriginalFile, string NewName)>();
                using (var connection = new DuckDBConnection($"Data Source={StagingDbPath}"))
                {
                    connection.Open();
                    using (var cmd = connection.CreateCommand())
                    {
                        cmd.CommandText = "SELECT DISTINCT file_name, nguoi_giao, ngay_giao, khach_hang, diem_giao FROM staging_items WHERE excel_file_name = $excelName";
                        cmd.Parameters.Add(new DuckDBParameter("$excelName", excelFileName));
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string origFile = reader.IsDBNull(0) ? "" : reader.GetString(0);
                                string shipper = reader.IsDBNull(1) ? "" : reader.GetString(1);
                                string dateStr = reader.IsDBNull(2) ? "" : reader.GetString(2);
                                string customer = reader.IsDBNull(3) ? "" : reader.GetString(3);
                                string delivery = reader.IsDBNull(4) ? "" : reader.GetString(4);

                                if (!string.IsNullOrEmpty(origFile))
                                {
                                    string newName = $"[{shipper}] - {dateStr} - {customer} - {delivery}";
                                    imageMapping.Add((origFile, newName));
                                }
                            }
                        }
                    }
                }

                if (imageMapping.Count == 0) return;

                string destFolder = Path.Combine(
                    Path.GetDirectoryName(destExcelPath) ?? "",
                    Path.GetFileNameWithoutExtension(destExcelPath) + "_Images"
                );

                if (!Directory.Exists(destFolder))
                {
                    Directory.CreateDirectory(destFolder);
                }

                string rawDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, RawMaterialDirName);
                int copiedCount = 0;

                foreach (var map in imageMapping)
                {
                    string origFullPath = Path.Combine(rawDir, map.OriginalFile);
                    if (File.Exists(origFullPath))
                    {
                        string safeName = map.NewName;
                        foreach (char c in Path.GetInvalidFileNameChars())
                        {
                            safeName = safeName.Replace(c, '_');
                        }
                        
                        string ext = Path.GetExtension(map.OriginalFile);
                        string destFullPath = Path.Combine(destFolder, safeName + ext);

                        try
                        {
                            File.Copy(origFullPath, destFullPath, true);
                            copiedCount++;
                        }
                        catch (Exception ex)
                        {
                            Logger.Log($"[Xuất ảnh] Lỗi sao chép {origFullPath} -> {destFullPath}: {ex.Message}");
                        }
                    }
                }

                Logger.Log($"[Xuất ảnh] Đã xuất {copiedCount} ảnh hóa đơn vào {destFolder}");
            }
            catch (Exception ex)
            {
                Logger.Log($"[Xuất ảnh] Lỗi chung: {ex.Message}");
            }
        }

        private void MainWindow_Closing(object? sender, CancelEventArgs e)
        {
            try
            {
                SaveOpenAiConfig();
            }
            catch { }

            _cts?.Dispose();
            if (_tessEngine.Values != null)
            {
                foreach (var engine in _tessEngine.Values)
                {
                    engine?.Dispose();
                }
            }
            _tessEngine.Dispose();
            Logger.Flush();
        }

        private List<string> GetSheetNames(string filePath)
        {
            var sheetNames = new List<string>();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        sheetNames.Add(reader.Name);
                    } while (reader.NextResult());
                }
            }
            return sheetNames;
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            string columnName = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return columnName;
        }

        private void InlineDeleteBtn_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is ProcessingItem item)
            {
                if (item.Status == ProcessStatus.Processing)
                {
                    MessageBox.Show("Không thể xóa ảnh khi đang trong quá trình xử lý.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var result = MessageBox.Show(
                    $"Bạn có chắc chắn muốn xóa ảnh '{item.SmartName}' khỏi danh sách chờ xử lý không?", 
                    "Xác nhận xóa ảnh", 
                    MessageBoxButton.YesNo, 
                    MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    if (ItemsListBox.SelectedItem == item)
                    {
                        ItemsListBox.SelectedItem = null;
                        PreviewImage.Source = null;
                    }
                    _items.Remove(item);
                    UpdateBatchActionBar();
                }
            }
        }

        private void LoadSyncHistoryFromDb()
        {
            _syncBatches.Clear();
            try
            {
                lock (DbWriteLock)
                {
                    string connStr = $"Data Source={StagingDbPath}";
                    using (var connection = new DuckDBConnection(connStr))
                    {
                        connection.Open();
                        using (var cmd = connection.CreateCommand())
                        {
                            cmd.CommandText = @"
                                SELECT 
                                    sync_batch_id, 
                                    MAX(synced_at) as synced_at, 
                                    COALESCE(MAX(excel_file_name), '') as excel_file_name, 
                                    MAX(nguoi_giao) as nguoi_giao, 
                                    COUNT(*) as row_count,
                                    MAX(status) as status
                                FROM staging_items 
                                WHERE sync_batch_id IS NOT NULL AND sync_batch_id != ''
                                GROUP BY sync_batch_id
                                ORDER BY synced_at DESC;";
                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var batch = new SyncBatchItem
                                    {
                                        SyncBatchId = reader.GetString(0),
                                        SyncedAt = reader.GetDateTime(1),
                                        ConfigName = reader.IsDBNull(2) ? "" : reader.GetString(2),
                                        Shipper = reader.IsDBNull(3) ? "" : reader.GetString(3),
                                        RowCount = reader.GetInt32(4),
                                        Status = reader.IsDBNull(5) ? "" : reader.GetString(5)
                                    };
                                    _syncBatches.Add(batch);
                                }
                            }
                        }
                    }
                }
                UpdateHistoryFilterLists();
            }
            catch (Exception ex)
            {
                Logger.Log($"[DuckDB Lỗi] LoadSyncHistoryFromDb: {ex.Message}");
                MessageBox.Show($"Lỗi tải lịch sử ghi sổ: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ReloadHistoryBtn_Click(object sender, RoutedEventArgs e)
        {
            LoadSyncHistoryFromDb();
        }

        private async void HistoryRollback_Click(object sender, RoutedEventArgs e)
        {
            if (!EnsureExcelLoaded()) return;

            if (sender is Button btn && btn.Tag is SyncBatchItem batch)
            {
                var confirm = MessageBox.Show(
                    $"Bạn có chắc chắn muốn HOÀN TÁC đợt ghi sổ ngày {batch.SyncedAt:dd/MM/yyyy HH:mm:ss} ({batch.RowCount} dòng) không?\n\n" +
                    "Hành động này sẽ:\n" +
                    "1. Tìm và xóa các dòng tương ứng ra khỏi Sheet Bán hàng trong file Excel.\n" +
                    "2. Khôi phục các dòng này về tab 'Xem trước & Đối soát' để bạn có thể sửa đổi hoặc ghi sổ lại.",
                    "Xác nhận hoàn tác",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (confirm != MessageBoxResult.Yes) return;

                var itemsInBatch = new List<StagingItem>();
                try
                {
                    lock (DbWriteLock)
                    {
                        using (var conn = new DuckDBConnection($"Data Source={StagingDbPath}"))
                        {
                            conn.Open();
                            using (var cmd = conn.CreateCommand())
                            {
                                cmd.CommandText = @"
                                    SELECT id, file_name, ngay_giao, khach_hang, diem_giao, nguoi_giao, ten_hang_goc, ten_hang_khop, hst, don_vi, sl_xuat, sl_nhan, sl_hong, ghi_chu, match_score 
                                    FROM staging_items 
                                    WHERE sync_batch_id = $batchId;";
                                cmd.Parameters.Add(new DuckDBParameter("batchId", batch.SyncBatchId));
                                using (var reader = cmd.ExecuteReader())
                                {
                                    while (reader.Read())
                                    {
                                        itemsInBatch.Add(new StagingItem
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
                                            GhiChu = reader.IsDBNull(13) ? "" : reader.GetString(13),
                                            MatchScore = reader.IsDBNull(14) ? 1.0 : reader.GetDouble(14)
                                        });
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception dbEx)
                {
                    MessageBox.Show($"Lỗi khi đọc dữ liệu đợt ghi sổ từ DB: {dbEx.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (itemsInBatch.Count == 0)
                {
                    MessageBox.Show("Không tìm thấy dòng dữ liệu nào thuộc đợt ghi sổ này trong cơ sở dữ liệu.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var config = _excelConfigs.FirstOrDefault(x => x.FileName == batch.ConfigName);
                if (config == null)
                {
                    config = _excelConfigs.FirstOrDefault(x => x.IsActive);
                }

                if (config == null)
                {
                    MessageBox.Show($"Không tìm thấy cấu hình Excel tương ứng với file '{batch.ConfigName}'. Vui lòng kiểm tra lại cấu hình ở Trang chủ.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (!File.Exists(config.LocalPath))
                {
                    MessageBox.Show($"Không tìm thấy file Excel cục bộ tại: {config.LocalPath}\nHủy thao tác hoàn tác.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                ShowRefLoading("Đang thực hiện hoàn tác trên file Excel...");
                bool excelSuccess = false;
                try
                {
                    await Task.Run(() =>
                    {
                        using (var workbook = new ClosedXML.Excel.XLWorkbook(config.LocalPath))
                        {
                            var salesSheet = workbook.Worksheet(config.SalesSheet);
                            if (salesSheet == null)
                            {
                                throw new Exception($"Không tìm thấy sheet '{config.SalesSheet}' trong file Excel.");
                            }

                            var headerRow = salesSheet.Row(1);
                            int colDate = 0, colCustomer = 0, colDiemGiao = 0, colShipper = 0, colHst = 0, colXuat = 0, colHong = 0;
                            int lastCol = salesSheet.LastColumnUsed()?.ColumnNumber() ?? 20;

                            for (int col = 1; col <= lastCol; col++)
                            {
                                string header = headerRow.Cell(col).GetString().Trim().ToLower();
                                if (header.Contains("ngày") || header.Contains("ngay")) colDate = col;
                                else if (header.Contains("khách") || header.Contains("khach")) colCustomer = col;
                                else if (header.Contains("điểm") || header.Contains("diem")) colDiemGiao = col;
                                else if (header.Contains("giao") && (header.Contains("người") || header.Contains("nguoi") || header.Contains("shipper"))) colShipper = col;
                                else if (header.Contains("hst") || header.Contains("mã")) colHst = col;
                                else if (header.Contains("số lượng xuất") || header.Contains("sl xuất") || header.Contains("sl xuat")) colXuat = col;
                                else if (header.Contains("hỏng") || header.Contains("hong")) colHong = col;
                            }

                            if (colDate == 0) colDate = 1;
                            if (colCustomer == 0) colCustomer = 2;
                            if (colDiemGiao == 0) colDiemGiao = 3;
                            if (colShipper == 0) colShipper = 4;
                            if (colHst == 0) colHst = 5;
                            if (colXuat == 0) colXuat = 6;
                            if (colHong == 0) colHong = 7;

                            int lastRow = salesSheet.LastRowUsed()?.RowNumber() ?? 1;
                            var pendingItems = new List<StagingItem>(itemsInBatch);
                            int deleteCount = 0;

                            for (int r = lastRow; r >= 2; r--)
                            {
                                var row = salesSheet.Row(r);

                                string xlDate = row.Cell(colDate).GetString().Trim();
                                string xlCustomer = row.Cell(colCustomer).GetString().Trim();
                                string xlDiemGiao = row.Cell(colDiemGiao).GetString().Trim();
                                string xlShipper = row.Cell(colShipper).GetString().Trim();
                                string xlHst = row.Cell(colHst).GetString().Trim();
                                
                                double xlXuat = 0;
                                double.TryParse(row.Cell(colXuat).GetString(), out xlXuat);
                                
                                double xlHong = 0;
                                double.TryParse(row.Cell(colHong).GetString(), out xlHong);

                                var match = pendingItems.FirstOrDefault(item =>
                                    item.NgayGiao == xlDate &&
                                    item.KhachHang == xlCustomer &&
                                    item.DiemGiao == xlDiemGiao &&
                                    item.NguoiGiao == xlShipper &&
                                    item.Hst == xlHst &&
                                    Math.Abs(item.SlXuat - xlXuat) < 0.001 &&
                                    Math.Abs(item.SlHong - xlHong) < 0.001);

                                if (match != null)
                                {
                                    row.Delete();
                                    pendingItems.Remove(match);
                                    deleteCount++;
                                }
                            }

                            workbook.Save();
                        }
                    });
                    excelSuccess = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Lỗi khi xóa dòng trong file Excel: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally
                {
                    HideRefLoading();
                }

                if (!excelSuccess) return;

                try
                {
                    lock (DbWriteLock)
                    {
                        using (var conn = new DuckDBConnection($"Data Source={StagingDbPath}"))
                        {
                            conn.Open();
                            using (var cmd = conn.CreateCommand())
                            {
                                cmd.CommandText = @"
                                    UPDATE staging_items 
                                    SET status = 'Success', sync_batch_id = NULL, synced_at = NULL 
                                    WHERE sync_batch_id = $batchId;";
                                cmd.Parameters.Add(new DuckDBParameter("batchId", batch.SyncBatchId));
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }

                    MessageBox.Show($"Hoàn tác thành công! Đã xóa các dòng tương ứng trong file Excel và đưa {batch.RowCount} dòng về hàng đợi đối soát.", "Thành công", MessageBoxButton.OK, MessageBoxImage.Information);

                    LoadStagingItemsFromDb();
                    LoadSyncHistoryFromDb();

                    await LoadBothSheetsAsync(config.RefSheet, config.SalesSheet);

                    MainTabControl.SelectedIndex = 2;
                }
                catch (Exception dbEx)
                {
                    MessageBox.Show($"Lỗi cập nhật trạng thái trong DuckDB: {dbEx.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        // ==========================================
        // STAGING PREVIEW & RECONCILE FILTERS
        // ==========================================
        private bool PreviewStagingFilter(object obj)
        {
            if (obj is StagingItem item)
            {
                // 1. Search text filter
                string query = (PreviewSearchBox.Text ?? "").Trim().ToLower();
                if (!string.IsNullOrEmpty(query))
                {
                    bool matchSearch = (item.TenHangGoc ?? "").ToLower().Contains(query) ||
                                       (item.TenHangKhop ?? "").ToLower().Contains(query) ||
                                       (item.Hst ?? "").ToLower().Contains(query);
                    if (!matchSearch) return false;
                }

                // 2. Date filter
                if (PreviewDatePicker.SelectedDate.HasValue)
                {
                    string selectedDateStr = PreviewDatePicker.SelectedDate.Value.ToString("dd/MM/yyyy");
                    if (item.NgayGiao != selectedDateStr) return false;
                }

                // 3. Shipper filter
                if (PreviewShipperCombo.SelectedValue is string shipperVal && !string.IsNullOrEmpty(shipperVal))
                {
                    if (item.NguoiGiao != shipperVal) return false;
                }

                // 4. Customer filter
                if (PreviewCustomerCombo.SelectedValue is string customerVal && !string.IsNullOrEmpty(customerVal))
                {
                    if (item.KhachHang != customerVal) return false;
                }

                // 5. Delivery point filter
                if (PreviewDeliveryCombo.SelectedValue is string deliveryVal && !string.IsNullOrEmpty(deliveryVal))
                {
                    if (item.DiemGiao != deliveryVal) return false;
                }

                return true;
            }
            return false;
        }

        private void RunAuditRules()
        {
            var branchToCustomerMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (var conn = new DuckDBConnection($"Data Source={StagingDbPath}"))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT DISTINCT diem_giao, khach_hang FROM staging_items WHERE status = 'Synced' AND diem_giao IS NOT NULL AND diem_giao != '';";
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string dg = reader.GetString(0);
                                string kh = reader.GetString(1);
                                if (!branchToCustomerMap.ContainsKey(dg))
                                {
                                    branchToCustomerMap[dg] = kh;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"[Audit Engine] Lỗi tải bản đồ chi nhánh - khách hàng: {ex.Message}");
            }

            foreach (var item in _stagingItems)
            {
                var warnings = new List<string>();

                // Rule 1: Excess Delivery
                if (item.SlNhan > item.SlXuat)
                {
                    string gc = (item.GhiChu ?? "").ToLower();
                    if (!gc.Contains("no") && !gc.Contains("tra") && !gc.Contains("bu") && !gc.Contains("muon") && !gc.Contains("gui"))
                    {
                        warnings.Add("Số thực nhận lớn hơn số lượng xuất (có thể ghi nhầm)");
                    }
                }

                // Rule 2: High Damage Rate
                if (item.SlHong > 0 && item.SlXuat > 0 && (item.SlHong / item.SlXuat) >= 0.3)
                {
                    warnings.Add($"Tỉ lệ hàng hỏng quá cao ({ (item.SlHong / item.SlXuat) * 100:F0}% tổng số lượng xuất)");
                }

                // Rule 3: Date Discrepancy
                if (DateTime.TryParseExact(item.NgayGiao, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
                {
                    double diffDays = Math.Abs((parsedDate - DateTime.Now).TotalDays);
                    if (diffDays > 30)
                    {
                        warnings.Add($"Ngày giao ({item.NgayGiao}) lệch quá xa so với ngày hiện tại ({diffDays:F0} ngày)");
                    }
                }
                else if (!string.IsNullOrEmpty(item.NgayGiao))
                {
                    warnings.Add($"Ngày giao không đúng định dạng dd/MM/yyyy: {item.NgayGiao}");
                }

                // Rule 4: Customer Discrepancy
                if (!string.IsNullOrEmpty(item.DiemGiao) && branchToCustomerMap.TryGetValue(item.DiemGiao, out string historicalCustomer))
                {
                    if (!historicalCustomer.Equals(item.KhachHang, StringComparison.OrdinalIgnoreCase))
                    {
                        warnings.Add($"Chi nhánh này lịch sử thuộc về khách hàng '{historicalCustomer}', nhưng hiện tại ghi nhận '{item.KhachHang}'");
                    }
                }

                item.AuditWarning = warnings.Count > 0 ? string.Join("\n• ", warnings.Prepend("Cảnh báo nghiệp vụ:")) : "";
            }
        }

        private void PreviewDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            PreviewInvoiceImage.Source = null;

            // Reset zoom & pan transforms
            PreviewImageScaleTransform.ScaleX = 1.0;
            PreviewImageScaleTransform.ScaleY = 1.0;
            PreviewImageTranslateTransform.X = 0;
            PreviewImageTranslateTransform.Y = 0;

            if (PreviewDataGrid.SelectedItem is StagingItem item)
            {
                string targetDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, RawMaterialDirName);
                string imgPath = Path.Combine(targetDir, item.FileName);

                if (File.Exists(imgPath))
                {
                    try
                    {
                        var bitmap = new System.Windows.Media.Imaging.BitmapImage();
                        bitmap.BeginInit();
                        bitmap.UriSource = new Uri(imgPath);
                        bitmap.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad;
                        bitmap.EndInit();
                        bitmap.Freeze();

                        PreviewInvoiceImage.Source = bitmap;
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"[Preview Ảnh Lỗi] Không thể tải ảnh: {ex.Message}");
                    }
                }
            }
        }

        private System.Windows.Point _previewPanStartPoint;
        private double _previewStartTranslateX;
        private double _previewStartTranslateY;
        private bool _previewIsPanning = false;

        private void ImageContainer_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (PreviewInvoiceImage.Source == null) return;

            double zoomFactor = e.Delta > 0 ? 1.15 : 1.0 / 1.15;
            
            double newScaleX = PreviewImageScaleTransform.ScaleX * zoomFactor;
            double newScaleY = PreviewImageScaleTransform.ScaleY * zoomFactor;

            if (newScaleX >= 0.1 && newScaleX <= 15.0)
            {
                PreviewImageScaleTransform.ScaleX = newScaleX;
                PreviewImageScaleTransform.ScaleY = newScaleY;
            }
            e.Handled = true;
        }

        private void ImageContainer_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (PreviewInvoiceImage.Source == null) return;

            _previewIsPanning = true;
            _previewPanStartPoint = e.GetPosition(ImageContainer);
            _previewStartTranslateX = PreviewImageTranslateTransform.X;
            _previewStartTranslateY = PreviewImageTranslateTransform.Y;
            ImageContainer.CaptureMouse();
            e.Handled = true;
        }

        private void ImageContainer_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_previewIsPanning)
            {
                _previewIsPanning = false;
                ImageContainer.ReleaseMouseCapture();
                e.Handled = true;
            }
        }

        private void ImageContainer_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_previewIsPanning) return;

            System.Windows.Point currentPoint = e.GetPosition(ImageContainer);
            double deltaX = currentPoint.X - _previewPanStartPoint.X;
            double deltaY = currentPoint.Y - _previewPanStartPoint.Y;

            PreviewImageTranslateTransform.X = _previewStartTranslateX + deltaX;
            PreviewImageTranslateTransform.Y = _previewStartTranslateY + deltaY;
            e.Handled = true;
        }

        private void UpdatePreviewFilterLists()
        {
            if (_isUpdatingPreviewFilters) return;
            _isUpdatingPreviewFilters = true;

            try
            {
                string currentShipper = PreviewShipperCombo.SelectedValue as string ?? "";
                string currentCustomer = PreviewCustomerCombo.SelectedValue as string ?? "";
                string currentDelivery = PreviewDeliveryCombo.SelectedValue as string ?? "";

                var query = (PreviewSearchBox.Text ?? "").Trim().ToLower();
                var dateMatchedItems = _stagingItems.Where(item => 
                {
                    if (!string.IsNullOrEmpty(query))
                    {
                        bool matchSearch = (item.TenHangGoc ?? "").ToLower().Contains(query) ||
                                           (item.TenHangKhop ?? "").ToLower().Contains(query) ||
                                           (item.Hst ?? "").ToLower().Contains(query);
                        if (!matchSearch) return false;
                    }
                    if (PreviewDatePicker.SelectedDate.HasValue)
                    {
                        string selectedDateStr = PreviewDatePicker.SelectedDate.Value.ToString("dd/MM/yyyy");
                        return item.NgayGiao == selectedDateStr;
                    }
                    return true;
                }).ToList();

                // 1. Shippers list
                var uniqueShippers = dateMatchedItems
                    .Select(x => x.NguoiGiao)
                    .Where(x => !string.IsNullOrEmpty(x))
                    .Distinct()
                    .OrderBy(x => x)
                    .Select(x => new FilterItem { Value = x, Text = x })
                    .ToList();
                uniqueShippers.Insert(0, new FilterItem { Value = "", Text = "Tất cả người giao" });
                PreviewShipperCombo.ItemsSource = uniqueShippers;

                if (uniqueShippers.Any(x => x.Value == currentShipper))
                    PreviewShipperCombo.SelectedValue = currentShipper;
                else
                    PreviewShipperCombo.SelectedIndex = 0;

                // 2. Customers list
                string activeShipper = PreviewShipperCombo.SelectedValue as string ?? "";
                var shipperMatchedItems = dateMatchedItems.Where(item => 
                    string.IsNullOrEmpty(activeShipper) || item.NguoiGiao == activeShipper
                ).ToList();

                var uniqueCustomers = shipperMatchedItems
                    .Select(x => x.KhachHang)
                    .Where(x => !string.IsNullOrEmpty(x))
                    .Distinct()
                    .OrderBy(x => x)
                    .Select(x => new FilterItem { Value = x, Text = x })
                    .ToList();
                uniqueCustomers.Insert(0, new FilterItem { Value = "", Text = "Tất cả đơn vị" });
                PreviewCustomerCombo.ItemsSource = uniqueCustomers;

                if (uniqueCustomers.Any(x => x.Value == currentCustomer))
                    PreviewCustomerCombo.SelectedValue = currentCustomer;
                else
                    PreviewCustomerCombo.SelectedIndex = 0;

                // 3. Delivery points list
                string activeCustomer = PreviewCustomerCombo.SelectedValue as string ?? "";
                var customerMatchedItems = shipperMatchedItems.Where(item => 
                    string.IsNullOrEmpty(activeCustomer) || item.KhachHang == activeCustomer
                ).ToList();

                var uniqueDeliveries = customerMatchedItems
                    .Select(x => x.DiemGiao)
                    .Where(x => !string.IsNullOrEmpty(x))
                    .Distinct()
                    .OrderBy(x => x)
                    .Select(x => new FilterItem { Value = x, Text = x })
                    .ToList();
                uniqueDeliveries.Insert(0, new FilterItem { Value = "", Text = "Tất cả chi nhánh" });
                PreviewDeliveryCombo.ItemsSource = uniqueDeliveries;

                if (uniqueDeliveries.Any(x => x.Value == currentDelivery))
                    PreviewDeliveryCombo.SelectedValue = currentDelivery;
                else
                    PreviewDeliveryCombo.SelectedIndex = 0;
            }
            finally
            {
                _isUpdatingPreviewFilters = false;
            }
        }

        private void PreviewSearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdatePreviewFilterLists();
            _stagingItemsView?.Refresh();
        }

        private void PreviewDatePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdatePreviewFilterLists();
            _stagingItemsView?.Refresh();
        }

        private void PreviewShipperCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdatingPreviewFilters) return;
            UpdatePreviewFilterLists();
            _stagingItemsView?.Refresh();
        }

        private void PreviewCustomerCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdatingPreviewFilters) return;
            UpdatePreviewFilterLists();
            _stagingItemsView?.Refresh();
        }

        private void PreviewDeliveryCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdatingPreviewFilters) return;
            _stagingItemsView?.Refresh();
        }

        private void ResetPreviewFilters_Click(object sender, RoutedEventArgs e)
        {
            PreviewSearchBox.Text = "";
            PreviewDatePicker.SelectedDate = null;
            PreviewShipperCombo.SelectedIndex = 0;
            PreviewCustomerCombo.SelectedIndex = 0;
            PreviewDeliveryCombo.SelectedIndex = 0;

            UpdatePreviewFilterLists();
            _stagingItemsView?.Refresh();
        }


        // ==========================================
        // SYNC HISTORY FILTERS
        // ==========================================
        private bool HistoryFilter(object obj)
        {
            if (obj is SyncBatchItem batch)
            {
                // 1. Search text filter
                string query = (HistorySearchBox.Text ?? "").Trim().ToLower();
                if (!string.IsNullOrEmpty(query))
                {
                    bool matchSearch = (batch.ConfigName ?? "").ToLower().Contains(query) ||
                                       (batch.Shipper ?? "").ToLower().Contains(query);
                    if (!matchSearch) return false;
                }

                // 2. Date filter
                if (HistoryDatePicker.SelectedDate.HasValue)
                {
                    if (batch.SyncedAt.Date != HistoryDatePicker.SelectedDate.Value.Date) return false;
                }

                // 3. Shipper filter
                if (HistoryShipperCombo.SelectedValue is string shipperVal && !string.IsNullOrEmpty(shipperVal))
                {
                    if (batch.Shipper != shipperVal) return false;
                }

                return true;
            }
            return false;
        }

        private void UpdateHistoryFilterLists()
        {
            if (_isUpdatingHistoryFilters) return;
            _isUpdatingHistoryFilters = true;

            try
            {
                string currentShipper = HistoryShipperCombo.SelectedValue as string ?? "";

                var query = (HistorySearchBox.Text ?? "").Trim().ToLower();
                var dateMatchedBatches = _syncBatches.Where(batch => 
                {
                    if (!string.IsNullOrEmpty(query))
                    {
                        bool matchSearch = (batch.ConfigName ?? "").ToLower().Contains(query) ||
                                           (batch.Shipper ?? "").ToLower().Contains(query);
                        if (!matchSearch) return false;
                    }
                    if (HistoryDatePicker.SelectedDate.HasValue)
                    {
                        return batch.SyncedAt.Date == HistoryDatePicker.SelectedDate.Value.Date;
                    }
                    return true;
                }).ToList();

                // Update Shippers list based on remaining matches
                var uniqueShippers = dateMatchedBatches
                    .Select(x => x.Shipper)
                    .Where(x => !string.IsNullOrEmpty(x))
                    .Distinct()
                    .OrderBy(x => x)
                    .Select(x => new FilterItem { Value = x, Text = x })
                    .ToList();
                uniqueShippers.Insert(0, new FilterItem { Value = "", Text = "Tất cả người giao" });
                HistoryShipperCombo.ItemsSource = uniqueShippers;

                if (uniqueShippers.Any(x => x.Value == currentShipper))
                    HistoryShipperCombo.SelectedValue = currentShipper;
                else
                    HistoryShipperCombo.SelectedIndex = 0;
            }
            finally
            {
                _isUpdatingHistoryFilters = false;
            }
        }

        private void HistorySearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdateHistoryFilterLists();
            _syncBatchesView?.Refresh();
        }

        private void HistoryDatePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateHistoryFilterLists();
            _syncBatchesView?.Refresh();
        }

        private void HistoryShipperCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdatingHistoryFilters) return;
            _syncBatchesView?.Refresh();
        }

        private void ResetHistoryFilters_Click(object sender, RoutedEventArgs e)
        {
            HistorySearchBox.Text = "";
            HistoryDatePicker.SelectedDate = null;
            HistoryShipperCombo.SelectedIndex = 0;

            UpdateHistoryFilterLists();
            _syncBatchesView?.Refresh();
        }
    }

    public class FilterItem
    {
        public string Value { get; set; } = string.Empty;
        public string Text { get; set; } = string.Empty;
    }

    public class SyncBatchItem
    {
        public string SyncBatchId { get; set; } = string.Empty;
        public DateTime SyncedAt { get; set; }
        public string ConfigName { get; set; } = string.Empty;
        public string Shipper { get; set; } = string.Empty;
        public int RowCount { get; set; }
        public string Status { get; set; } = string.Empty;

        public string StatusText => Status == "Synced" ? "Đã ghi sổ" : "Đã hoàn tác";
        public string StatusColor => Status == "Synced" ? "#16a34a" : "#dc2626";
        public bool CanRollback => Status == "Synced";
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
                // Fix: only allow substring contains if length >= 3 to avoid matching single-letter words like "e" (from "lá é") with "green" or "rocket"
                if (targetTokens.Any(tTok => tTok == sTok || (tTok.Length >= 3 && sTok.Length >= 3 && (tTok.Contains(sTok) || sTok.Contains(tTok)))))
                {
                    intersectCount++;
                }
            }

            // Dice Coefficient: 2 * intersect / (len1 + len2)
            return (2.0 * intersectCount) / (sourceTokens.Length + targetTokens.Length);
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

            // Default color rules: "cải mơ" / "mizuna" default to green ("xanh") if neither red ("đỏ") nor green ("xanh") is specified
            if (cleanRaw.Contains("cai mo") && !cleanRaw.Contains("do") && !cleanRaw.Contains("xanh"))
            {
                cleanRaw += " xanh";
            }
            else if (cleanRaw.Contains("mizuna") && !cleanRaw.Contains("do") && !cleanRaw.Contains("xanh"))
            {
                cleanRaw += " xanh";
            }

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

        public static ReferenceItem? FindBestMatchOptimized(string rawName, IEnumerable<CachedRefInfo> cachedItems, out double bestScore)
        {
            bestScore = 0.0;
            ReferenceItem? bestMatch = null;

            string cleanRaw = RemoveSign4VietnameseString(rawName.Trim().ToLower());

            // Default color rules: "cải mơ" / "mizuna" default to green ("xanh") if neither red ("đỏ") nor green ("xanh") is specified
            if (cleanRaw.Contains("cai mo") && !cleanRaw.Contains("do") && !cleanRaw.Contains("xanh"))
            {
                cleanRaw += " xanh";
            }
            else if (cleanRaw.Contains("mizuna") && !cleanRaw.Contains("do") && !cleanRaw.Contains("xanh"))
            {
                cleanRaw += " xanh";
            }

            foreach (var cached in cachedItems)
            {
                if (!string.IsNullOrEmpty(cached.CleanVietTat) && cleanRaw == cached.CleanVietTat)
                {
                    bestScore = 1.0;
                    return cached.Item;
                }

                double tokenScore = GetTokenMatchScore(cleanRaw, cached.CleanTenHang);
                double jwScore = JaroWinklerDistance(cleanRaw, cached.CleanTenHang);

                double score = (0.6 * tokenScore) + (0.4 * jwScore);

                if (score > bestScore)
                {
                    bestScore = score;
                    bestMatch = cached.Item;
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

    public class MatchScoreToBrushConverter : System.Windows.Data.IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value is double score)
            {
                if (score < 0.75)
                {
                    // Light yellow/amber warning background
                    return new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(254, 243, 199)); // #fef3c7 (Amber 100)
                }
            }
            return System.Windows.DependencyProperty.UnsetValue;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
