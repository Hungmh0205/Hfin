using System.Windows;

namespace AccountingOcrTest
{
    public static class MessageBox
    {
        public static MessageBoxResult Show(string messageBoxText)
        {
            return Show(messageBoxText, "Thông báo", MessageBoxButton.OK);
        }

        public static MessageBoxResult Show(string messageBoxText, string caption)
        {
            return Show(messageBoxText, caption, MessageBoxButton.OK);
        }

        public static MessageBoxResult Show(string messageBoxText, string caption, MessageBoxButton button)
        {
            // Run on UI thread if called from background
            if (Application.Current != null && !Application.Current.Dispatcher.CheckAccess())
            {
                return Application.Current.Dispatcher.Invoke(() => Show(messageBoxText, caption, button));
            }

            var win = new CustomMessageBoxWindow(messageBoxText, caption, button);
            if (Application.Current != null && Application.Current.MainWindow != null && Application.Current.MainWindow.IsVisible)
            {
                win.Owner = Application.Current.MainWindow;
            }
            win.ShowDialog();
            return win.Result;
        }

        public static MessageBoxResult Show(string messageBoxText, string caption, MessageBoxButton button, MessageBoxImage icon)
        {
            return Show(messageBoxText, caption, button);
        }

        public static MessageBoxResult Show(Window owner, string messageBoxText, string caption = "Thông báo", MessageBoxButton button = MessageBoxButton.OK)
        {
            if (Application.Current != null && !Application.Current.Dispatcher.CheckAccess())
            {
                return Application.Current.Dispatcher.Invoke(() => Show(owner, messageBoxText, caption, button));
            }

            var win = new CustomMessageBoxWindow(messageBoxText, caption, button);
            win.Owner = owner;
            win.ShowDialog();
            return win.Result;
        }
    }
}
