using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace AccountingOcrTest
{
    public class InvoiceData
    {
        public string quy_trinh_suy_luan { get; set; }
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
}
