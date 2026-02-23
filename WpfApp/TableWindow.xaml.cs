using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using ClosedXML.Excel;

namespace WpfApp
{
    public partial class TableWindow : Window
    {
        public TableWindow(string title, Control content)
        {
            InitializeComponent();
            Title = title;
            ContentHost.Content = content;
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void ExportToCsv_Click(object sender, RoutedEventArgs e)
        {
            if (ContentHost.Content is not DataGrid dg || dg.ItemsSource is not DataView dv)
            {
                MessageBox.Show("No exportable data found.", "Export Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var sfd = new SaveFileDialog
            {
                Filter = "CSV File (*.csv)|*.csv|All Files (*.*)|*.*",
                FileName = $"{Title.Replace(" ", "_")}.csv"
            };

            if (sfd.ShowDialog() == true)
            {
                try
                {
                    ExportDataViewToCsv(dv, sfd.FileName);
                    MessageBox.Show("Data exported successfully!", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error exporting data: {ex.Message}", "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ExportDataViewToCsv(DataView view, string filePath)
        {
            var table = view.Table;
            if (table == null) return;

            var sb = new StringBuilder();
            
            // Use local list separator (e.g., "," for US, ";" for EU)
            string sep = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator;

            // Headers
            for (int i = 0; i < table.Columns.Count; i++)
            {
                sb.Append(FormatCsvField(table.Columns[i].ColumnName, sep));
                if (i < table.Columns.Count - 1)
                    sb.Append(sep);
            }
            sb.AppendLine();

            // Rows (from View to respect Sort/Filter)
            foreach (DataRowView rowView in view)
            {
                var row = rowView.Row;
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    sb.Append(FormatCsvField(row[i]?.ToString(), sep));
                    if (i < table.Columns.Count - 1)
                        sb.Append(sep);
                }
                sb.AppendLine();
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            if (ContentHost.Content is not DataGrid dg || dg.ItemsSource is not DataView dv)
            {
                MessageBox.Show("No exportable data found.", "Export Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var sfd = new SaveFileDialog
            {
                Filter = "Excel Workbook (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                FileName = $"{Title.Replace(" ", "_")}.xlsx"
            };

            if (sfd.ShowDialog() == true)
            {
                try
                {
                    ExportDataViewToExcel(dv, sfd.FileName);
                    MessageBox.Show("Data exported successfully!", "Export", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error exporting data: {ex.Message}", "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ExportDataViewToExcel(DataView view, string filePath)
        {
            // Convert DataView to DataTable to preserver structure
            // NOTE: ToTable() creates a new table with current rows, respecting sort/filter if used correctly
            var dt = view.ToTable();

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Exported Data");
            
            // Insert Data
            // ClosedXML handles DataTable directly
            var tableRange = worksheet.Cell(1, 1).InsertTable(dt);

            // Adjust styles if needed
            tableRange.Theme = XLTableTheme.TableStyleLight8;
            worksheet.Columns().AdjustToContents();

            workbook.SaveAs(filePath);
        }

        private string FormatCsvField(string? field, string separator)
        {
            if (string.IsNullOrEmpty(field)) return "";
            if (field.Contains(separator) || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
            {
                return $"\"{field.Replace("\"", "\"\"")}\"";
            }
            return field;
        }
    }
}
