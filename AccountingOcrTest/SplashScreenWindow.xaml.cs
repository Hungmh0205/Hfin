using System;
using System.Threading.Tasks;
using System.Windows;

namespace AccountingOcrTest
{
    /// <summary>
    /// Interaction logic for SplashScreenWindow.xaml
    /// </summary>
    public partial class SplashScreenWindow : Window
    {
        public SplashScreenWindow()
        {
            InitializeComponent();
            RunInitialization();
        }

        private async void RunInitialization()
        {
            try
            {
                // Simulate progressive loading steps with nice status updates
                await Task.Delay(600);
                StatusText.Text = "Đang kết nối cơ sở dữ liệu DuckDB...";
                
                await Task.Delay(850);
                StatusText.Text = "Đang nạp cấu hình và người giao hàng...";
                
                await Task.Delay(850);
                StatusText.Text = "Đang kiểm tra API Key và hệ thống...";
                
                await Task.Delay(600);
                StatusText.Text = "Sẵn sàng!";
                
                await Task.Delay(250);

                // Initialize and show MainWindow
                MainWindow main = new MainWindow();
                App.Current.MainWindow = main;
                main.Show();

                // Close the splash window
                this.Close();
            }
            catch (Exception)
            {
                // Fallback to guarantee app launch
                MainWindow main = new MainWindow();
                App.Current.MainWindow = main;
                main.Show();
                this.Close();
            }
        }
    }
}
