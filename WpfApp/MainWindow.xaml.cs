
using System;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace WpfApp
{
    public partial class MainWindow : Window
    {
        private readonly ObservableCollection<TouchstoneFileData> _loadedFiles = new();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void MenuItem_OpenTdrCharts(object sender, RoutedEventArgs e)
        {
            if (_loadedFiles.Count == 0)
            {
                MessageBox.Show(this, "Please load at least one Touchstone file first.", "TDR Charts", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var window = new TdrWindow(_loadedFiles)
            {
                Owner = this
            };
            window.Show();
        }

        private void MenuItem_OpenLayoutViewer(object sender, RoutedEventArgs e)
        {
            var dxfWin = new DxfWindow { Owner = this };
            dxfWin.CapacitorFileLoaded += (s, filePath) =>
            {
                try
                {
                    // Avoid loading duplicates
                    if (_loadedFiles.Any(f => string.Equals(f.FilePath, filePath, StringComparison.OrdinalIgnoreCase)))
                        return;
                    var data = TouchstoneParser.Parse(filePath);
                    _loadedFiles.Add(data);
                    AddFileToTablesMenu(data);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, $"Error loading capacitor file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            };
            dxfWin.Show();
        }

        private void MenuItem_OpenFile(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Touchstone (*.s*p;*.spd)|*.s*p;*.spd|All files (*.*)|*.*",
                Multiselect = true
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            foreach (var filePath in dialog.FileNames)
            {
                try
                {
                    var data = TouchstoneParser.Parse(filePath);
                    _loadedFiles.Add(data);
                    AddFileToTablesMenu(data);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, $"Error parsing {filePath}: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void MenuItem_CloseAllFiles(object sender, RoutedEventArgs e)
        {
            if (_loadedFiles.Count == 0) return;

            if (MessageBox.Show("Are you sure you want to close all files and clear data?", "Confirm Close", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                _loadedFiles.Clear();
                TablesMenu.Items.Clear();
                
                // Close any open child windows (Graphs, Tables, TDR)
                foreach (Window window in Application.Current.Windows)
                {
                    if (window != this)
                    {
                        window.Close();
                    }
                }
            }
        }


        private void MenuItem_Exit(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void MenuItem_OpenCharts(object sender, RoutedEventArgs e)
        {
            if (_loadedFiles.Count == 0)
            {
                MessageBox.Show(this, "Please load at least one Touchstone file first.", "Charts", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var window = new GraphWindow(_loadedFiles)
            {
                Owner = this
            };

            window.Show();
        }

        private DataGrid BuildRawGrid(TouchstoneFileData file)
        {
            var table = new DataTable();
            table.Columns.Add("Frequency (Hz)", typeof(double));

            var isRi = string.Equals(file.Format, "RI", StringComparison.OrdinalIgnoreCase);

            foreach (var name in file.ParameterNames)
            {
                if (isRi)
                {
                    table.Columns.Add($"{name} Re", typeof(double));
                    table.Columns.Add($"{name} Im", typeof(double));
                }
                else
                {
                    table.Columns.Add($"{name} Mag", typeof(double));
                    table.Columns.Add($"{name} Phase (deg)", typeof(double));
                }
            }

            foreach (var point in file.Points)
            {
                var row = table.NewRow();
                row[0] = point.FrequencyHz;

                for (var i = 0; i < file.ParameterNames.Count; i++)
                {
                    var param = point.Parameters[i];
                    if (isRi)
                    {
                        row[1 + i * 2] = param.Real;
                        row[1 + i * 2 + 1] = param.Imaginary;
                    }
                    else
                    {
                        row[1 + i * 2] = param.Magnitude;
                        row[1 + i * 2 + 1] = param.PhaseDegrees;
                    }
                }

                table.Rows.Add(row);
            }

            return CreateGrid(table);
        }

        private DataGrid BuildMagnitudeGrid(TouchstoneFileData file)
        {
            var table = new DataTable();
            table.Columns.Add("Frequenza (Hz)", typeof(double));

            foreach (var name in file.ParameterNames)
            {
                table.Columns.Add($"{name} (dB)", typeof(double));
            }

            foreach (var point in file.Points)
            {
                var row = table.NewRow();
                row[0] = point.FrequencyHz;
                for (var i = 0; i < file.ParameterNames.Count; i++)
                {
                    row[1 + i] = ToDb(point.Parameters[i].Magnitude);
                }

                table.Rows.Add(row);
            }

            return CreateGrid(table);
        }

        private static DataGrid CreateGrid(DataTable table)
        {
            return new DataGrid
            {
                ItemsSource = table.DefaultView,
                AutoGenerateColumns = true,
                IsReadOnly = true,
                Margin = new Thickness(8),
                EnableRowVirtualization = true,
                EnableColumnVirtualization = true
            };
        }

        private static double ToDb(double magnitude)
        {
            return magnitude > 0
                ? 20 * Math.Log10(magnitude)
                : double.NegativeInfinity;
        }

        private void AddFileToTablesMenu(TouchstoneFileData file)
        {
            var parent = new MenuItem { Header = file.FileName };

            var dataItem = new MenuItem { Header = "Mostra dati" };
            dataItem.Click += (_, _) => ShowDataTab(file);

            var magnitudeItem = new MenuItem { Header = "Mostra magnitudini" };
            magnitudeItem.Click += (_, _) => ShowMagnitudeTab(file);

            var closeItem = new MenuItem { Header = "Chiudi File" };
            closeItem.Click += (_, _) => CloseFile(file, parent);

            parent.Items.Add(dataItem);
            parent.Items.Add(magnitudeItem);
            parent.Items.Add(new Separator());
            parent.Items.Add(closeItem);

            TablesMenu.Items.Add(parent);
        }

        private void CloseFile(TouchstoneFileData file, MenuItem menuItem)
        {
            if (MessageBox.Show($"Are you sure you want to close '{file.FileName}'?", "Confirm Close", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                _loadedFiles.Remove(file);
                TablesMenu.Items.Remove(menuItem);
            }
        }

        private void ShowDataTab(TouchstoneFileData file)
        {
            var win = new TableWindow($"Dati - {file.FileName}", BuildRawGrid(file))
            {
                Owner = this
            };
            win.Show();
        }

        private void ShowMagnitudeTab(TouchstoneFileData file)
        {
            var win = new TableWindow($"Magnitudini - {file.FileName}", BuildMagnitudeGrid(file))
            {
                Owner = this
            };
            win.Show();
        }
    }
}
