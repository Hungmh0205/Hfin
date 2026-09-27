using System.Windows;
using System.Windows.Input;

namespace AccountingOcrTest
{
    public partial class CustomMessageBoxWindow : Window
    {
        public MessageBoxResult Result { get; private set; } = MessageBoxResult.None;

        public CustomMessageBoxWindow(string message, string title, MessageBoxButton button)
        {
            InitializeComponent();
            MessageTextBlock.Text = message;
            TitleTextBlock.Text = title;

            // Change icon based on keywords
            string lowerMsg = message.ToLower();
            if (lowerMsg.Contains("lỗi") || lowerMsg.Contains("không thể") || lowerMsg.Contains("thất bại"))
            {
                IconText.Text = "❌";
            }
            else if (lowerMsg.Contains("chắc chắn") || lowerMsg.Contains("xác nhận") || lowerMsg.Contains("có muốn"))
            {
                IconText.Text = "❓";
            }
            else if (lowerMsg.Contains("thành công") || lowerMsg.Contains("hoàn tất") || lowerMsg.Contains("ghi sổ"))
            {
                IconText.Text = "✅";
            }
            else
            {
                IconText.Text = "ℹ️";
            }

            // Configure buttons
            switch (button)
            {
                case MessageBoxButton.OK:
                    OkBtn.Visibility = Visibility.Visible;
                    OkBtn.Content = "OK";
                    NoBtn.Visibility = Visibility.Collapsed;
                    CancelBtn.Visibility = Visibility.Collapsed;
                    break;
                case MessageBoxButton.OKCancel:
                    OkBtn.Visibility = Visibility.Visible;
                    OkBtn.Content = "Đồng ý";
                    CancelBtn.Visibility = Visibility.Visible;
                    CancelBtn.Content = "Hủy";
                    NoBtn.Visibility = Visibility.Collapsed;
                    break;
                case MessageBoxButton.YesNo:
                    OkBtn.Visibility = Visibility.Visible;
                    OkBtn.Content = "Có";
                    NoBtn.Visibility = Visibility.Visible;
                    NoBtn.Content = "Không";
                    CancelBtn.Visibility = Visibility.Collapsed;
                    break;
                case MessageBoxButton.YesNoCancel:
                    OkBtn.Visibility = Visibility.Visible;
                    OkBtn.Content = "Có";
                    NoBtn.Visibility = Visibility.Visible;
                    NoBtn.Content = "Không";
                    CancelBtn.Visibility = Visibility.Visible;
                    CancelBtn.Content = "Hủy";
                    break;
            }
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                DragMove();
            }
            catch { }
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            Result = MessageBoxResult.Cancel;
            Close();
        }

        private void OkBtn_Click(object sender, RoutedEventArgs e)
        {
            if (OkBtn.Content.ToString() == "Có")
                Result = MessageBoxResult.Yes;
            else
                Result = MessageBoxResult.OK;

            DialogResult = true;
            Close();
        }

        private void NoBtn_Click(object sender, RoutedEventArgs e)
        {
            Result = MessageBoxResult.No;
            DialogResult = false;
            Close();
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            Result = MessageBoxResult.Cancel;
            DialogResult = false;
            Close();
        }
    }
}
