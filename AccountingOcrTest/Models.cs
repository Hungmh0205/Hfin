using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace AccountingOcrTest
{
    public class InvoiceData : INotifyPropertyChanged
    {
        private string _quy_trinh_suy_luan = "";
        private string _ten_nguoi_giao = "";
        private string _ngay_giao = "";
        private string _khach_hang = "";
        private string _diem_giao = "";
        private List<InvoiceItem> _danh_sach_hang_hoa = new List<InvoiceItem>();

        public string quy_trinh_suy_luan
        {
            get => _quy_trinh_suy_luan;
            set { _quy_trinh_suy_luan = value; OnPropertyChanged(); }
        }

        public string ten_nguoi_giao
        {
            get => _ten_nguoi_giao;
            set { _ten_nguoi_giao = value; OnPropertyChanged(); }
        }

        public string ngay_giao
        {
            get => _ngay_giao;
            set { _ngay_giao = value; OnPropertyChanged(); }
        }

        public string khach_hang
        {
            get => _khach_hang;
            set { _khach_hang = FuzzyMatcher.FormatKhachHang(value); OnPropertyChanged(); }
        }

        public string diem_giao
        {
            get => _diem_giao;
            set { _diem_giao = FuzzyMatcher.FormatDiemGiao(value); OnPropertyChanged(); }
        }

        public List<InvoiceItem> danh_sach_hang_hoa
        {
            get => _danh_sach_hang_hoa;
            set { _danh_sach_hang_hoa = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    public class InvoiceItem : INotifyPropertyChanged
    {
        private string _ten_hang;
        private double _sl_xuat;
        private double _sl_nhan;
        private double _sl_hong;

        public string ten_hang
        {
            get => _ten_hang;
            set { _ten_hang = value; OnPropertyChanged(); }
        }

        public double sl_xuat
        {
            get => _sl_xuat;
            set { _sl_xuat = value; OnPropertyChanged(); RecalculateHong(); }
        }

        public double sl_nhan
        {
            get => _sl_nhan;
            set { _sl_nhan = value; OnPropertyChanged(); RecalculateHong(); }
        }

        public double sl_hong
        {
            get => _sl_hong;
            set { _sl_hong = value; OnPropertyChanged(); }
        }

        private void RecalculateHong()
        {
            sl_hong = sl_xuat - sl_nhan;
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
    public enum ProcessStatus
    {
        Waiting,
        Processing,
        Success,
        Error
    }

    public class ProcessingItem : INotifyPropertyChanged
    {
        private string _smartName;
        private ProcessStatus _status;
        private InvoiceData _data;
        private string _errorMessage;

        public string FilePath { get; set; }
        
        public string SmartName
        {
            get => _smartName;
            set { _smartName = value; OnPropertyChanged(); }
        }

        public ProcessStatus Status
        {
            get => _status;
            set { _status = value; OnPropertyChanged(); OnPropertyChanged(nameof(StatusColor)); OnPropertyChanged(nameof(StatusIcon)); }
        }

        public InvoiceData Data
        {
            get => _data;
            set { _data = value; OnPropertyChanged(); }
        }

        public string ErrorMessage
        {
            get => _errorMessage;
            set { _errorMessage = value; OnPropertyChanged(); }
        }

        public string StatusColor
        {
            get
            {
                switch (Status)
                {
                    case ProcessStatus.Waiting: return "Gray";
                    case ProcessStatus.Processing: return "Blue";
                    case ProcessStatus.Success: return "Green";
                    case ProcessStatus.Error: return "Red";
                    default: return "Black";
                }
            }
        }

        public string StatusIcon
        {
            get
            {
                switch (Status)
                {
                    case ProcessStatus.Waiting: return "⏳";
                    case ProcessStatus.Processing: return "🔄";
                    case ProcessStatus.Success: return "✅";
                    case ProcessStatus.Error: return "❌";
                    default: return "";
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    public enum KeyStatus
    {
        Ready,
        RateLimited,
        Error
    }

    public class ApiKeyInfo : INotifyPropertyChanged
    {
        private KeyStatus _status;
        private int _requestsInLastMinute;

        public string Key { get; set; }
        public DateTime LastUsed { get; set; }
        public DateTime WindowStart { get; set; }
        
        public KeyStatus Status
        {
            get => _status;
            set { _status = value; OnPropertyChanged(); OnPropertyChanged(nameof(StatusColor)); }
        }

        public int RequestsInLastMinute
        {
            get => _requestsInLastMinute;
            set { _requestsInLastMinute = value; OnPropertyChanged(); }
        }

        public string StatusColor
        {
            get
            {
                switch (Status)
                {
                    case KeyStatus.Ready: return "Green";
                    case KeyStatus.RateLimited: return "Orange";
                    case KeyStatus.Error: return "Red";
                    default: return "Black";
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    public class ReferenceItem
    {
        public string VietTat { get; set; } = string.Empty;
        public string TenHang { get; set; } = string.Empty;
        public string DonViTinh { get; set; } = string.Empty;
    }

    public class CachedRefInfo
    {
        public ReferenceItem Item { get; set; } = new ReferenceItem();
        public string CleanTenHang { get; set; } = string.Empty;
        public string CleanVietTat { get; set; } = string.Empty;
    }

    public class ExcelConfigItem : INotifyPropertyChanged
    {
        private bool _isActive;

        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string OriginalPath { get; set; } = string.Empty;
        public string LocalPath { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string RefSheet { get; set; } = string.Empty;
        public string SalesSheet { get; set; } = string.Empty;
        public DateTime DateAdded { get; set; } = DateTime.Now;

        public bool IsActive
        {
            get => _isActive;
            set
            {
                _isActive = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(StatusText));
                OnPropertyChanged(nameof(StatusColor));
                OnPropertyChanged(nameof(CanLoad));
            }
        }

        public string StatusText => IsActive ? "Đang dùng" : "Chờ nạp";
        public string StatusColor => IsActive ? "#10b981" : "#64748b"; // Bold green or neutral slate
        public bool CanLoad => !IsActive;

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    public class StagingItem : INotifyPropertyChanged
    {
        private string _id = "";
        private string _fileName = "";
        private string _ngayGiao = "";
        private string _khachHang = "";
        private string _diemGiao = "";
        private string _nguoiGiao = "";
        private string _tenHangGoc = "";
        private string _tenHangKhop = "";
        private string _hst = "";
        private string _donVi = "";
        private double _slXuat;
        private double _slNhan;
        private double _slHong;
        private string _ghiChu = "";
        private bool _isSelected = true;
        private double _matchScore = 1.0;

        public string Id { get => _id; set { _id = value; OnPropertyChanged(); } }
        public string FileName { get => _fileName; set { _fileName = value; OnPropertyChanged(); } }
        public string NgayGiao { get => _ngayGiao; set { _ngayGiao = value; OnPropertyChanged(); } }
        public string KhachHang { get => _khachHang; set { _khachHang = FuzzyMatcher.FormatKhachHang(value); OnPropertyChanged(); } }
        public string DiemGiao { get => _diemGiao; set { _diemGiao = FuzzyMatcher.FormatDiemGiao(value); OnPropertyChanged(); } }
        public string NguoiGiao { get => _nguoiGiao; set { _nguoiGiao = value; OnPropertyChanged(); } }
        public string TenHangGoc { get => _tenHangGoc; set { _tenHangGoc = value; OnPropertyChanged(); } }
        public string TenHangKhop { get => _tenHangKhop; set { _tenHangKhop = value; OnPropertyChanged(); } }
        public string Hst { get => _hst; set { _hst = value; OnPropertyChanged(); } }
        public string DonVi { get => _donVi; set { _donVi = value; OnPropertyChanged(); } }
        public double SlXuat 
        { 
            get => _slXuat; 
            set 
            { 
                _slXuat = value; 
                OnPropertyChanged(); 
                RecalculateHong(); 
            } 
        }
        public double SlNhan 
        { 
            get => _slNhan; 
            set 
            { 
                _slNhan = value; 
                OnPropertyChanged(); 
                RecalculateHong(); 
            } 
        }
        public double SlHong { get => _slHong; set { _slHong = value; OnPropertyChanged(); } }
        public string GhiChu { get => _ghiChu; set { _ghiChu = value; OnPropertyChanged(); } }
        public bool IsSelected { get => _isSelected; set { _isSelected = value; OnPropertyChanged(); } }
        public double MatchScore { get => _matchScore; set { _matchScore = value; OnPropertyChanged(); } }

        private void RecalculateHong()
        {
            SlHong = SlXuat - SlNhan;
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}

