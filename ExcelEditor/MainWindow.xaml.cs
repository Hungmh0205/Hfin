using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using ClosedXML.Excel;
using Microsoft.Win32;

namespace ExcelEditor
{
    public partial class MainWindow : Window
    {
        private string _currentFilePath;
        private XLWorkbook _workbook;
        private IXLWorksheet _currentWorksheet;
        private DataTable _currentDataTable;
        private int _maxCol;
        private int _maxRow;
        private HashSet<(int Row, int Col)> _editedCells = new HashSet<(int, int)>();

        public MainWindow()
        {
            InitializeComponent();
        }

        private async void ImportFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xlsm",
                Title = "Select an Excel File"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _currentFilePath = openFileDialog.FileName;
                FilePathText.Text = _currentFilePath;
                await LoadWorkbookAsync();
            }
        }

        private async Task LoadWorkbookAsync()
        {
            ShowLoading("Opening Workbook...");
            try
            {
                MainDataGrid.ItemsSource = null;
                _currentWorksheet = null;
                SheetsListBox.Items.Clear();
                
                _workbook?.Dispose();

                _workbook = await Task.Run(() => new XLWorkbook(_currentFilePath));

                foreach (var sheet in _workbook.Worksheets)
                {
                    SheetsListBox.Items.Add(sheet.Name);
                }
            }
            catch (IOException)
            {
                MessageBox.Show("The file is currently open in another program. Please close it and try again.", "File in Use", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading workbook: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                HideLoading();
            }
        }

        private async void SheetsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SheetsListBox.SelectedItem == null) return;
            string sheetName = SheetsListBox.SelectedItem.ToString();
            
            _currentWorksheet = await Task.Run(() => _workbook.Worksheet(sheetName));
            await LoadSheetDataAsync();
        }

        private async Task LoadSheetDataAsync()
        {
            if (_currentWorksheet == null) return;

            ShowLoading("Loading Sheet Data...");

            try
            {
                int firstDataRow = 1;
                int firstDataCol = 1;
                bool hasPivot = false;

                var dt = await Task.Run(() => 
                {
                    hasPivot = _currentWorksheet.PivotTables.Any();
                    var dataTable = new DataTable();
                    var lastCell = _currentWorksheet.LastCellUsed();
                    var firstCell = _currentWorksheet.FirstCellUsed();

                    if (firstCell != null)
                    {
                        firstDataRow = firstCell.Address.RowNumber;
                        firstDataCol = firstCell.Address.ColumnNumber;
                    }

                    _maxRow = lastCell != null ? lastCell.Address.RowNumber : 10;
                    _maxCol = lastCell != null ? lastCell.Address.ColumnNumber : 10;

                    // Safety limit to prevent OutOfMemory on huge accidental ranges
                    if (_maxRow > 500000) _maxRow = 500000;
                    if (_maxCol > 1000) _maxCol = 1000;

                    // Ensure minimum visual size to look like Excel
                    if (_maxRow < 20) _maxRow = 20;
                    if (_maxCol < 10) _maxCol = 10;

                    // Create exact Excel columns: A, B, C...
                    for (int c = 1; c <= _maxCol; c++)
                    {
                        string colName = XLHelper.GetColumnLetterFromNumber(c);
                        dataTable.Columns.Add(colName);
                    }

                    // Pre-allocate empty rows (instantly fast compared to ClosedXML cell lookup)
                    for (int r = 1; r <= _maxRow; r++)
                    {
                        dataTable.Rows.Add(dataTable.NewRow());
                    }

                    // Only process ACTUALLY USED cells! (O(U) instead of O(R*C))
                    foreach (var cell in _currentWorksheet.CellsUsed())
                    {
                        int r = cell.Address.RowNumber;
                        int c = cell.Address.ColumnNumber;

                        if (r <= _maxRow && c <= _maxCol)
                        {
                            try
                            {
                                if (cell.HasFormula)
                                {
                                    if (cell.FormulaA1.Contains("["))
                                    {
                                        // External link: ClosedXML cannot evaluate this. Use cache or placeholder.
                                        string cached = cell.CachedValue.ToString();
                                        dataTable.Rows[r - 1][c - 1] = string.IsNullOrWhiteSpace(cached) ? "#EXT_LINK" : cached;
                                    }
                                    else
                                    {
                                        // Local formula: Safe to evaluate. 
                                        // cell.Value automatically evaluates if cache is missing.
                                        dataTable.Rows[r - 1][c - 1] = cell.Value.ToString();
                                    }
                                }
                                else
                                {
                                    dataTable.Rows[r - 1][c - 1] = cell.Value.ToString(); 
                                }
                            }
                            catch
                            {
                                dataTable.Rows[r - 1][c - 1] = "#ERROR"; // Ultimate fallback, safely ignore
                            }
                        }
                    }

                    return dataTable;
                });

                if (dt != null)
                {
                    _currentDataTable = dt;
                    _currentDataTable.AcceptChanges();
                    _editedCells.Clear();

                    _currentDataTable.ColumnChanged += (s, e) => 
                    {
                        int r = _currentDataTable.Rows.IndexOf(e.Row);
                        int c = e.Column.Ordinal;
                        _editedCells.Add((r, c));
                    };

                    MainDataGrid.ItemsSource = _currentDataTable.DefaultView;

                    if (hasPivot)
                    {
                        MainDataGrid.IsReadOnly = true;
                        WarningText.Visibility = Visibility.Visible;
                        SaveBtn.IsEnabled = false;
                        SaveBtn.Opacity = 0.5;
                    }
                    else
                    {
                        MainDataGrid.IsReadOnly = false;
                        WarningText.Visibility = Visibility.Collapsed;
                        SaveBtn.IsEnabled = true;
                        SaveBtn.Opacity = 1.0;
                    }
                    
                    // Auto-scroll to first data cell to handle "unusual layout"
                    if (firstDataRow > 1 || firstDataCol > 1)
                    {
                        MainDataGrid.Dispatcher.InvokeAsync(async () =>
                        {
                            await Task.Delay(100); // Give DataGrid time to render layout
                            if (MainDataGrid.Items.Count >= firstDataRow && MainDataGrid.Columns.Count >= firstDataCol)
                            {
                                var item = MainDataGrid.Items[firstDataRow - 1];
                                var col = MainDataGrid.Columns[firstDataCol - 1];
                                MainDataGrid.ScrollIntoView(item, col);
                                MainDataGrid.SelectedCells.Clear();
                                MainDataGrid.SelectedCells.Add(new DataGridCellInfo(item, col));
                            }
                        }, System.Windows.Threading.DispatcherPriority.ApplicationIdle);
                    }
                }
                else
                {
                    MainDataGrid.ItemsSource = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading sheet: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                HideLoading();
            }
        }

        private void MainDataGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if (_currentWorksheet != null)
            {
                int c = e.Column.DisplayIndex + 1; 
                if (c > 0 && c <= _maxCol)
                {
                    if (_currentWorksheet.Column(c).IsHidden)
                    {
                        e.Column.Visibility = Visibility.Collapsed;
                    }
                }
            }
        }

        private void MainDataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            int rowIndex = e.Row.GetIndex() + 1;
            e.Row.Header = rowIndex.ToString();
            
            if (_currentWorksheet != null && _currentWorksheet.Row(rowIndex).IsHidden)
            {
                e.Row.Visibility = Visibility.Collapsed;
            }
        }

        private async void SaveFile_Click(object sender, RoutedEventArgs e)
        {
            if (_workbook == null || _currentWorksheet == null || _currentDataTable == null)
            {
                MessageBox.Show("No data to save.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                MainDataGrid.CommitEdit(DataGridEditingUnit.Row, true);

                if (_editedCells.Count == 0)
                {
                    MessageBox.Show("Không có thay đổi nào để lưu (No changes to save).", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                ShowLoading("Saving Changes...");

                var editsToProcess = _editedCells.ToList();
                var changes = new List<(int R, int C, string Value)>();
                foreach (var edit in editsToProcess)
                {
                    changes.Add((edit.Row, edit.Col, _currentDataTable.Rows[edit.Row][edit.Col]?.ToString()));
                }

                await Task.Run(() => 
                {
                    foreach (var change in changes)
                    {
                        var cell = _currentWorksheet.Cell(change.R + 1, change.C + 1);
                        string newValue = change.Value;

                        if (!cell.HasFormula)
                        {
                            if (string.IsNullOrWhiteSpace(newValue))
                                cell.Value = Blank.Value;
                            else if (double.TryParse(newValue, out double numVal))
                                cell.Value = numVal;
                            else if (DateTime.TryParse(newValue, out DateTime dateVal))
                                cell.Value = dateVal;
                            else if (bool.TryParse(newValue, out bool boolVal))
                                cell.Value = boolVal;
                            else
                                cell.Value = newValue;
                        }
                    }

                    _workbook.Save();
                });

                _editedCells.Clear();
                MessageBox.Show("Changes saved successfully! Formulas, Macros, and Pivot Tables were preserved.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (IOException)
            {
                MessageBox.Show("Cannot save because the file is open in another application (like Excel). Please close it and try again.", "File in Use", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                HideLoading();
            }
        }

        private void ShowLoading(string text)
        {
            LoadingText.Text = text;
            LoadingOverlay.Visibility = Visibility.Visible;
            MainDataGrid.IsEnabled = false;
            SheetsListBox.IsEnabled = false;
        }

        private void HideLoading()
        {
            LoadingOverlay.Visibility = Visibility.Collapsed;
            MainDataGrid.IsEnabled = true;
            SheetsListBox.IsEnabled = true;
        }
    }
}