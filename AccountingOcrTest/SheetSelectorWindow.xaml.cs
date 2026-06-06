using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace AccountingOcrTest
{
    public partial class SheetSelectorWindow : Window
    {
        public string SelectedReferenceSheet { get; private set; } = "";
        public string SelectedSalesSheet { get; private set; } = "";

        public SheetSelectorWindow(List<string> sheetNames)
        {
            InitializeComponent();

            foreach (var name in sheetNames)
            {
                RefSheetComboBox.Items.Add(name);
                SalesSheetComboBox.Items.Add(name);
            }

            // Smart defaults
            if (sheetNames.Count > 0)
            {
                // Find reference sheet
                string? defaultRef = sheetNames.FirstOrDefault(x => 
                    x.Contains("tham chiếu", StringComparison.OrdinalIgnoreCase) || 
                    x.Contains("tham chieu", StringComparison.OrdinalIgnoreCase) || 
                    x.Contains("danh mục", StringComparison.OrdinalIgnoreCase) || 
                    x.Contains("sp", StringComparison.OrdinalIgnoreCase) || 
                    x.Contains("product", StringComparison.OrdinalIgnoreCase));
                
                RefSheetComboBox.SelectedItem = defaultRef ?? sheetNames[0];

                // Find sales sheet (must be different from ref sheet if possible)
                string? defaultSales = sheetNames.FirstOrDefault(x => 
                    (x.Contains("bán hàng", StringComparison.OrdinalIgnoreCase) || 
                     x.Contains("ban hang", StringComparison.OrdinalIgnoreCase) || 
                     x.Contains("sales", StringComparison.OrdinalIgnoreCase) || 
                     x.Contains("chi tiết", StringComparison.OrdinalIgnoreCase) || 
                     x.Contains("data", StringComparison.OrdinalIgnoreCase)) &&
                    x != RefSheetComboBox.SelectedItem.ToString());

                SalesSheetComboBox.SelectedItem = defaultSales ?? (sheetNames.Count > 1 ? sheetNames[1] : sheetNames[0]);
            }
        }

        private void ConfirmBtn_Click(object sender, RoutedEventArgs e)
        {
            if (RefSheetComboBox.SelectedItem == null || SalesSheetComboBox.SelectedItem == null)
            {
                MessageBox.Show("Vui lòng chọn đầy đủ cả Sheet tham chiếu và Sheet bán hàng.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SelectedReferenceSheet = RefSheetComboBox.SelectedItem.ToString() ?? "";
            SelectedSalesSheet = SalesSheetComboBox.SelectedItem.ToString() ?? "";
            DialogResult = true;
            Close();
        }

        private void CancelBtn_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
