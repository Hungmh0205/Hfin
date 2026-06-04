using Microsoft.Win32;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace AccountingOcrTest
{
    public partial class ApiKeyManagerWindow : Window
    {
        private ApiKeyManager _manager;

        public ApiKeyManagerWindow(ApiKeyManager manager)
        {
            InitializeComponent();
            _manager = manager;
            KeysListView.ItemsSource = _manager.Keys;
        }

        private void AddKeyBtn_Click(object sender, RoutedEventArgs e)
        {
            string key = NewKeyTextBox.Text.Trim();
            if (!string.IsNullOrEmpty(key) && !_manager.Keys.Any(k => k.Key == key))
            {
                _manager.Keys.Add(new ApiKeyInfo { Key = key, Status = KeyStatus.Ready });
                _manager.SaveKeys();
                NewKeyTextBox.Clear();
            }
        }

        private void ImportBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Text Files|*.txt",
                Title = "Chọn file chứa API Keys"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                var lines = File.ReadAllLines(openFileDialog.FileName);
                foreach (var line in lines)
                {
                    string key = line.Trim();
                    if (!string.IsNullOrEmpty(key) && !_manager.Keys.Any(k => k.Key == key))
                    {
                        _manager.Keys.Add(new ApiKeyInfo { Key = key, Status = KeyStatus.Ready });
                    }
                }
                _manager.SaveKeys();
            }
        }

        private void DeleteKeyBtn_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is ApiKeyInfo info)
            {
                _manager.Keys.Remove(info);
                _manager.SaveKeys();
            }
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
