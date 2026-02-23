using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using OxyPlot.Legends;

namespace WpfApp
{
    public partial class TdrWindow : Window
    {
        private readonly ObservableCollection<TouchstoneFileData> _files;
        private readonly ObservableCollection<FileSelection> _fileSelections = new();
        private readonly Dictionary<(string FilePath, string ParamName), bool> _selectedParams = new();
        private readonly ObservableCollection<TdrLegendItem> _legendItems = new();
        private bool _isDraggingLegend;
        private Point _legendOffset;

        public TdrWindow(ObservableCollection<TouchstoneFileData> files)
        {
            InitializeComponent();
            _files = files;
            FileSelectionControl.ItemsSource = _fileSelections;
            LegendList.ItemsSource = _legendItems;

            // Init UI
            WindowComboBox.ItemsSource = Enum.GetValues(typeof(WindowType));
            WindowComboBox.SelectedItem = WindowType.Hanning;

            PopulateFileSelection();
            _files.CollectionChanged += Files_CollectionChanged;
        }

        private void Files_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            // Re-populate the list when files are added or removed
            // Preserve existing selections if possible?
            // Simple approach: repopulate fully, but try to keep selected state
            
            PopulateFileSelection();
        }

        private void PopulateFileSelection()
        {
            // Store previous selection state to restore it after refresh
            var previousSelection = new Dictionary<(string FilePath, string Param), bool>();
            foreach (var fs in _fileSelections)
            {
                foreach (var p in fs.Parameters)
                {
                    if (p.IsSelected)
                    {
                        previousSelection[(fs.File.FilePath, p.Name)] = true;
                    }
                }
            }

            _fileSelections.Clear();
            foreach (var file in _files)
            {
                var selection = new FileSelection(file);
                
                // Restore selection or pre-select reflection params if new
                foreach(var p in selection.Parameters)
                {
                    if (previousSelection.TryGetValue((file.FilePath, p.Name), out bool isSelected))
                    {
                        p.IsSelected = isSelected;
                    }
                    else if (p.Name.Length == 3 && p.Name[1] == p.Name[2])
                    {
                        // Optional: auto-select Sxx for new files?
                        // Let's not auto-select to avoid clutter without user intent
                    }
                }
                
                _fileSelections.Add(selection);
            }
        }

        private async void ParameterCheckBox_Checked(object sender, RoutedEventArgs e)
        {
             // Optional: auto-update or wait for button?
             // User has "Aggiorna Grafico" button, so maybe wait.
             // But let's keep track of selection in our internal struct if needed, 
             // or just read from _fileSelections when button is clicked.
        }


        private async void UpdateBtn_Click(object sender, RoutedEventArgs e)
        {
            await RecalculateAndPlot();
        }

        private async Task RecalculateAndPlot()
        {
            LoadingText.Visibility = Visibility.Visible;
            TdrPlot.Model = null;

            try
            {
                // Parse Settings
                if (!double.TryParse(RiseTimeTextBox.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double riseTime))
                {
                    MessageBox.Show("Invalid Rise Time.");
                    return;
                }
                if (!double.TryParse(Z0TextBox.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double z0))
                {
                    MessageBox.Show("Invalid Z0.");
                    return;
                }
                if (!double.TryParse(MaxDurationTextBox.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double tMax))
                {
                    MessageBox.Show("Invalid Max Duration.");
                    return;
                }
                if (!double.TryParse(DelayTextBox.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double delay))
                {
                    MessageBox.Show("Invalid Delay.");
                    return;
                }
                
                var settings = new TdrSettings
                {
                    RiseTime = riseTime,
                    SystemImpedance = z0,
                    Delay = delay,
                    MaxDuration = tMax,
                    Window = (WindowType)WindowComboBox.SelectedItem
                };

                // Filter selected params
                var targets = new List<(TouchstoneFileData File, string Param)>();
                foreach (var selection in _fileSelections)
                {
                    foreach (var p in selection.Parameters)
                    {
                        if (p.IsSelected)
                        {
                            targets.Add((selection.File, p.Name));
                        }
                    }
                }

                if (targets.Count == 0)
                {
                     MessageBox.Show("No parameter selected.");
                     return;
                }

                var model = new PlotModel { Title = "TDR Impedance Profile" };
                
                // Imposta l'asse X per visualizzare da -Delay (o 0 se Delay=0) fino a MaxDuration
                // Se Delay serve per compensare (time shifting), allora l'inizio è -Delay.
                // Se vogliamo visualizzare una durata fissa "MaxDuration", impostiamo min e max.
                // Generalmente si vuole vedere da 0 a MaxDuration, o l'intervallo calcolato.
                // Qui impostiamo Maximum per tagliare la simulazione alla durata richiesta.
                
                model.Axes.Add(new LinearAxis 
                { 
                    Position = AxisPosition.Bottom, 
                    Title = "Time (s)", 
                    Unit = "s",
                    Minimum = 0, 
                    Maximum = settings.MaxDuration,
                    MajorGridlineStyle = LineStyle.Solid,
                    MinorGridlineStyle = LineStyle.Dot
                });
                
                model.Axes.Add(new LinearAxis 
                {
                    Position = AxisPosition.Left, 
                    Title = "Impedance (Ω)", 
                    Unit = "Ohm",
                    MajorGridlineStyle = LineStyle.Solid,
                    MinorGridlineStyle = LineStyle.Dot
                });
                
                model.IsLegendVisible = false;

                // Async calc
                var results = await Task.Run(() => 
                {
                    var list = new List<(string Name, TdrResult Res)>();
                    foreach(var t in targets)
                    {
                        var res = TdrCalculator.Calculate(t.File, t.Param, settings);
                        list.Add(($"{t.File.FileName} - {t.Param}", res));
                    }
                    return list;
                });


                _legendItems.Clear();
                var colorPalette = new[] { 
                    System.Windows.Media.Colors.Red, 
                    System.Windows.Media.Colors.Blue, 
                    System.Windows.Media.Colors.Green, 
                    System.Windows.Media.Colors.Orange, 
                    System.Windows.Media.Colors.Purple, 
                    System.Windows.Media.Colors.Brown, 
                    System.Windows.Media.Colors.Magenta,
                    System.Windows.Media.Colors.Teal
                };
                int idx = 0;

                foreach (var r in results)
                {
                    var color = colorPalette[idx % colorPalette.Length];
                    var series = new LineSeries 
                    { 
                        Title = r.Name,
                        Color = OxyColor.FromRgb(color.R, color.G, color.B)
                    };
                    for (int i = 0; i < r.Res.Time.Length; i++)
                    {
                        series.Points.Add(new DataPoint(r.Res.Time[i], r.Res.Impedance[i]));
                    }
                    model.Series.Add(series);
                    
                    _legendItems.Add(new TdrLegendItem 
                    { 
                        Title = r.Name, 
                        ColorBrush = new System.Windows.Media.SolidColorBrush(color) 
                    });
                    idx++;
                }

                TdrPlot.Model = model;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error in TDR calculation: {ex.Message}");
            }
            finally
            {
                LoadingText.Visibility = Visibility.Collapsed;
            }
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

    public class TdrLegendItem
    {
        public string Title { get; set; }
        public System.Windows.Media.Brush ColorBrush { get; set; }
    }
}
