using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using OxyPlot.Legends;

namespace WpfApp
{
    public partial class GraphWindow : Window
    {
        private readonly ObservableCollection<TouchstoneFileData> _files;
        private readonly ObservableCollection<FileSelection> _fileSelections = new();
        private readonly ObservableCollection<GraphLegendItem> _legendItems = new();
        private PlotSettings _settings = PlotSettings.Default();
        private List<(TouchstoneFileData File, string Param, string FileName)> _lastSelection = new();
        private bool _isDraggingLegend;
        private Point _legendOffset;
        private readonly List<Color> _palette = new()
        {
            Colors.DeepSkyBlue,
            Colors.OrangeRed,
            Colors.LimeGreen,
            Colors.Goldenrod,
            Colors.HotPink,
            Colors.MediumPurple,
            Colors.Teal,
            Colors.Crimson,
            Colors.CadetBlue,
            Colors.DarkCyan
        };

        public GraphWindow(ObservableCollection<TouchstoneFileData> files)
        {
            InitializeComponent();
            _files = files;
            SelectionColumns.ItemsSource = _fileSelections;
            LegendList.ItemsSource = _legendItems;
            PopulateSelections();
            _files.CollectionChanged += Files_CollectionChanged;
        }

        private void Files_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            PopulateSelections();

            // Aggiorna il grafico dopo che le selezioni sono state ripopolate
            var selected = GetCurrentSelection();
            if (selected.Count == 0)
            {
                PlotView.Model = null;
                EmptyChartLabel.Visibility = Visibility.Visible;
                _legendItems.Clear();
            }
            else
            {
                DrawChart(selected);
            }
        }

        private void PopulateSelections()
        {
            var previousSelection = _fileSelections
                .ToDictionary(fs => fs.File.FilePath, fs => fs.Parameters.Where(p => p.IsSelected).Select(p => p.Name).ToHashSet(StringComparer.OrdinalIgnoreCase));

            _fileSelections.Clear();
            foreach (var file in _files)
            {
                var selection = new FileSelection(file);
                if (previousSelection.TryGetValue(file.FilePath, out var selectedNames))
                {
                    foreach (var param in selection.Parameters)
                    {
                        param.IsSelected = selectedNames.Contains(param.Name);
                    }
                }
                _fileSelections.Add(selection);
            }
        }

        private void DrawChart(List<(TouchstoneFileData File, string Param, string FileName)> selected)
        {
            if (!_settings.Validate(out var validationError))
            {
                MessageBox.Show(this, validationError, "Settings", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var model = new PlotModel
            {
                Background = OxyColors.White,
                TextColor = OxyColors.Black,
                PlotAreaBorderColor = OxyColors.Gray,
                IsLegendVisible = true
            };

            model.IsLegendVisible = false;

            Axis xAxis = _settings.IsXLog
                ? new LogarithmicAxis { Position = AxisPosition.Bottom, Title = "Frequency (Hz)", MajorGridlineStyle = LineStyle.Solid, MinorGridlineStyle = LineStyle.Dot }
                : new LinearAxis { Position = AxisPosition.Bottom, Title = "Frequency (Hz)", MajorGridlineStyle = LineStyle.Solid, MinorGridlineStyle = LineStyle.Dot, StringFormat = "0.###E+0" };

            Axis yAxis = _settings.IsYLog
                ? new LogarithmicAxis { Position = AxisPosition.Left, Title = "Amplitude (dB)", MajorGridlineStyle = LineStyle.Solid, MinorGridlineStyle = LineStyle.Dot }
                : new LinearAxis { Position = AxisPosition.Left, Title = "Amplitude (dB)", MajorGridlineStyle = LineStyle.Solid, MinorGridlineStyle = LineStyle.Dot };

            if (!_settings.IsXAuto)
            {
                xAxis.Minimum = _settings.XMin!.Value;
                xAxis.Maximum = _settings.XMax!.Value;
            }

            if (!_settings.IsYAuto)
            {
                yAxis.Minimum = _settings.YMin!.Value;
                yAxis.Maximum = _settings.YMax!.Value;
            }

            model.Axes.Add(xAxis);
            model.Axes.Add(yAxis);

            _legendItems.Clear();
            for (var idx = 0; idx < selected.Count; idx++)
            {
                var sel = selected[idx];
                var paramIndex = FindParameterIndex(sel.File, sel.Param);
                if (paramIndex < 0)
                {
                    continue;
                }
                var color = ToOxy(_palette[idx % _palette.Count]);
                var series = new LineSeries
                {
                    Title = $"{sel.FileName} - {sel.Param}",
                    Color = color,
                    StrokeThickness = 2,
                    MarkerType = MarkerType.None
                };

                foreach (var point in sel.File.Points)
                {
                    var magnitudeDb = ToDb(point.Parameters[paramIndex].Magnitude);
                    series.Points.Add(new DataPoint(point.FrequencyHz, magnitudeDb));
                }

                model.Series.Add(series);
                
                _legendItems.Add(new GraphLegendItem
                {
                    Title = $"{sel.FileName} - {sel.Param}",
                    ColorBrush = new SolidColorBrush(_palette[idx % _palette.Count])
                });
            }

            PlotView.Model = model;
            EmptyChartLabel.Visibility = Visibility.Collapsed;
        }

        private void ParameterCheckBox_Toggled(object sender, RoutedEventArgs e)
        {
            // Assicura che il ViewModel rifletta lo stato della CheckBox prima di procedere
            if (sender is CheckBox cb && cb.DataContext is ParameterOption param)
            {
                param.IsSelected = cb.IsChecked ?? false;
            }

            var selected = GetCurrentSelection();
            if (selected.Count == 0)
            {
                _lastSelection = selected;
                PlotView.Model = null;
                EmptyChartLabel.Visibility = Visibility.Visible;
                _legendItems.Clear();
                return;
            }

            _lastSelection = selected;
            EmptyChartLabel.Visibility = Visibility.Collapsed;
            DrawChart(selected);
        }

        private List<(TouchstoneFileData File, string Param, string FileName)> GetCurrentSelection()
        {
            return _fileSelections
                .SelectMany(f => f.Parameters.Where(p => p.IsSelected).Select(p => (File: f.File, Param: p.Name, FileName: f.File.FileName)))
                .ToList();
        }

        private static int FindParameterIndex(TouchstoneFileData file, string parameter)
        {
            for (var i = 0; i < file.ParameterNames.Count; i++)
            {
                if (string.Equals(file.ParameterNames[i], parameter, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        private static double ToDb(double magnitude)
        {
            return magnitude > 0
                ? 20 * Math.Log10(magnitude)
                : double.NegativeInfinity;
        }

        private static OxyColor ToOxy(Color color)
        {
            return OxyColor.FromArgb(color.A, color.R, color.G, color.B);
        }

        private void CustomLegend_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement element)
            {
                _isDraggingLegend = true;
                _legendOffset = e.GetPosition(element);
                
                // If using Canvas.Right initially, switch to Canvas.Left based on current position
                if (double.IsNaN(Canvas.GetLeft(element)))
                {
                    var canvas = element.Parent as Canvas;
                    if (canvas != null)
                    {
                        var positionInCanvas = element.TranslatePoint(new Point(0, 0), canvas);
                        Canvas.SetLeft(element, positionInCanvas.X);
                        Canvas.SetRight(element, double.NaN);
                    }
                }
                
                element.CaptureMouse();
            }
        }

        private void CustomLegend_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (_isDraggingLegend && sender is FrameworkElement element)
            {
                var canvas = element.Parent as Canvas;
                if (canvas != null)
                {
                    var mousePos = e.GetPosition(canvas);
                    Canvas.SetLeft(element, mousePos.X - _legendOffset.X);
                    Canvas.SetTop(element, mousePos.Y - _legendOffset.Y);
                }
            }
        }

        private void CustomLegend_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (_isDraggingLegend && sender is FrameworkElement element)
            {
                _isDraggingLegend = false;
                element.ReleaseMouseCapture();
            }
        }
    }

    public sealed class FileSelection
    {
        public FileSelection(TouchstoneFileData file)
        {
            File = file;
            FileName = file.FileName;
            Parameters = new ObservableCollection<ParameterOption>(file.ParameterNames.Select(n => new ParameterOption { Name = n }));
        }

        public TouchstoneFileData File { get; }
        public string FileName { get; }
        public ObservableCollection<ParameterOption> Parameters { get; }
    }

    public class GraphLegendItem
    {
        public string Title { get; set; }
        public System.Windows.Media.Brush ColorBrush { get; set; }
    }

    public sealed class ParameterOption
    {
        public string Name { get; set; } = string.Empty;
        public bool IsSelected { get; set; }
    }

    public sealed class PlotSettings
    {
        public bool IsXLog { get; set; }
        public bool IsYLog { get; set; }
        public bool IsXAuto { get; set; } = true;
        public bool IsYAuto { get; set; } = true;
        public double? XMin { get; set; }
        public double? XMax { get; set; }
        public double? YMin { get; set; }
        public double? YMax { get; set; }

        public PlotSettings Clone()
        {
            return new PlotSettings
            {
                IsXLog = IsXLog,
                IsYLog = IsYLog,
                IsXAuto = IsXAuto,
                IsYAuto = IsYAuto,
                XMin = XMin,
                XMax = XMax,
                YMin = YMin,
                YMax = YMax
            };
        }

        public static PlotSettings Default() => new();

        public bool Validate(out string error)
        {
            if (!IsXAuto)
            {
                if (XMin == null || XMax == null) { error = "Specify min and max for X axis."; return false; }
                if (XMax <= XMin) { error = "Max X must be greater than Min X."; return false; }
                if (IsXLog && XMin <= 0) { error = "For logarithmic X scale, Min must be > 0."; return false; }
            }

            if (!IsYAuto)
            {
                if (YMin == null || YMax == null) { error = "Specify min and max for Y axis."; return false; }
                if (YMax <= YMin) { error = "Max Y must be greater than Min Y."; return false; }
                if (IsYLog && YMin <= 0) { error = "For logarithmic Y scale, Min must be > 0."; return false; }
            }

            error = string.Empty;
            return true;
        }
    }
}
