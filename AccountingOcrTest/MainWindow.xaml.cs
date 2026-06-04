using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using Tesseract;

namespace AccountingOcrTest
{
    public class InvoiceData
    {
        public string ngay_giao { get; set; }
        public string khach_hang { get; set; }
        public string diem_giao { get; set; }
        public List<InvoiceItem> danh_sach_hang_hoa { get; set; }
    }

    public class InvoiceItem
    {
        public string ten_hang { get; set; }
        public double sl_xuat { get; set; }
        public double sl_nhan { get; set; }
        public double sl_hong { get; set; }
    }

    public partial class MainWindow : Window
    {
        private string _selectedImagePath = "";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void SelectImageBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Image Files|*.jpg;*.jpeg;*.png;*.webp",
                Title = "Chọn ảnh hóa đơn"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _selectedImagePath = openFileDialog.FileName;
                FilePathText.Text = _selectedImagePath;
                ExtractBtn.IsEnabled = true;
            }
        }

        private async void ExtractBtn_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(ApiKeyTextBox.Text))
            {
                MessageBox.Show("Vui lòng nhập Gemini API Key trước.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(_selectedImagePath))
            {
                MessageBox.Show("Vui lòng chọn ảnh.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                StatusText.Text = "Đang gửi dữ liệu lên Gemini AI để trích xuất...";
                StatusText.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(59, 130, 246));
                ExtractBtn.IsEnabled = false;

                string base64Image = "";
                string mimeType = "image/jpeg";
                if (_selectedImagePath.ToLower().EndsWith(".png")) mimeType = "image/png";
                if (_selectedImagePath.ToLower().EndsWith(".webp")) mimeType = "image/webp";

                using (var img = new Bitmap(_selectedImagePath))
                {
                    string tessDataPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");
                    try 
                    {
                        using (var engine = new TesseractEngine(tessDataPath, "osd", EngineMode.Default))
                        {
                            using (var pix = Pix.LoadFromFile(_selectedImagePath))
                            {
                                using (var page = engine.Process(pix, PageSegMode.OsdOnly))
                                {
                                    var iterator = page.GetIterator();
                                    if (iterator != null)
                                    {
                                        iterator.Begin();
                                        var props = iterator.GetProperties();
                                        Orientation orientation = props.Orientation;

                                        if (orientation == Orientation.PageRight)
                                        {
                                            img.RotateFlip(RotateFlipType.Rotate270FlipNone);
                                        }
                                        else if (orientation == Orientation.PageDown)
                                        {
                                            img.RotateFlip(RotateFlipType.Rotate180FlipNone);
                                        }
                                        else if (orientation == Orientation.PageLeft)
                                        {
                                            img.RotateFlip(RotateFlipType.Rotate90FlipNone);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch { }

                    using (var ms = new System.IO.MemoryStream())
                    {
                        var format = mimeType == "image/png" ? System.Drawing.Imaging.ImageFormat.Png : System.Drawing.Imaging.ImageFormat.Jpeg;
                        img.Save(ms, format);
                        byte[] imageBytes = ms.ToArray();
                        base64Image = Convert.ToBase64String(imageBytes);
                    }
                }

                InvoiceData result = await CallGeminiApi(base64Image, mimeType, ApiKeyTextBox.Text.Trim());

                NgayGiaoText.Text = result.ngay_giao;
                KhachHangText.Text = result.khach_hang;
                DiemGiaoText.Text = result.diem_giao;
                ProductsGrid.ItemsSource = result.danh_sach_hang_hoa;

                StatusText.Text = "Trích xuất thành công!";
                StatusText.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(16, 185, 129));
            }
            catch (Exception ex)
            {
                StatusText.Text = "Lỗi trích xuất!";
                StatusText.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(239, 68, 68));
                MessageBox.Show($"Lỗi khi gọi API: {ex.Message}\n\nChi tiết: {ex.InnerException?.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                ExtractBtn.IsEnabled = true;
            }
        }

        private async Task<InvoiceData> CallGeminiApi(string base64Image, string mimeType, string apiKey)
        {
            string promptText = @"Bạn là một hệ thống OCR và kiểm toán kế toán cao cấp chuyên nghiệp của Việt Nam. Nhiệm vụ của bạn là số hóa hóa đơn này với ĐỘ CHÍNH XÁC TUYỆT ĐỐI về mặt cấu trúc và nội dung. ĐẶC BIỆT CHÚ Ý: Hình ảnh có thể bị xoay ngang, xoay dọc, hoặc lộn ngược. Hãy tự động 'xoay' và định hướng lại hình ảnh trong tư duy trước khi đọc. TUYỆT ĐỐI KHÔNG ĐƯỢC BỎ SÓT BẤT KỲ DÒNG HÀNG HÓA NÀO. Hãy dóng hàng ngang thật cẩn thận để đảm bảo Tên hàng hóa và Số lượng khớp nhau 100%, không bị trôi dòng, lệch dòng hay mất dòng.

HÃY THỰC HIỆN QUY TRÌNH SUY LUẬN (CHAIN OF THOUGHT):

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
                        required = new[] { "ngay_giao", "khach_hang", "diem_giao", "danh_sach_hang_hoa" }
                    }
                }
            };

            string jsonBody = JsonSerializer.Serialize(requestBody);

            using (var client = new HttpClient())
            {
                string url = $"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite:generateContent?key={apiKey}";
                
                var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");
                var response = await client.PostAsync(url, content);
                
                string responseStr = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    throw new Exception($"Gemini API trả về lỗi {response.StatusCode}: {responseStr}");
                }

                using (JsonDocument doc = JsonDocument.Parse(responseStr))
                {
                    var root = doc.RootElement;
                    var candidates = root.GetProperty("candidates");
                    if (candidates.GetArrayLength() > 0)
                    {
                        var text = candidates[0].GetProperty("content").GetProperty("parts")[0].GetProperty("text").GetString();
                        
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
}