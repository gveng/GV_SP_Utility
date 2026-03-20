using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Numerics;
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
        
        // RLC Tuning
        private (TouchstoneFileData File, string Param, string FileName)? _rlcTarget;
        private RlcResult _currentRlc;
        private RlcTuningWindow? _tuningWindow;

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

        private void Scale_Changed(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded) return;

            _settings.IsXLog = XLogRadio.IsChecked == true;
            _settings.IsYLog = YLogRadio.IsChecked == true;
            
            // Force Auto range when switching scales to avoid invalid manual ranges (e.g. log <= 0)
            // since the manual range UI is removed.
            _settings.IsXAuto = true;
            _settings.IsYAuto = true;

            var selected = GetCurrentSelection();
            if (selected.Count > 0)
            {
                DrawChart(selected);
            }
        }

        private void Format_Changed(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded) return;

            if (FormatDbRadio.IsChecked == true) _settings.Format = GraphFormat.Db;
            else if (FormatLinearRadio.IsChecked == true) _settings.Format = GraphFormat.Linear;
            else if (FormatImpedanceRadio.IsChecked == true) _settings.Format = GraphFormat.Impedance;
            
            // Force Auto range when switching formats
            _settings.IsYAuto = true;

            var selected = GetCurrentSelection();
            if (selected.Count > 0)
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

            string yTitle = _settings.Format switch
            {
                GraphFormat.Db => "Amplitude (dB)",
                GraphFormat.Linear => "Magnitude (Linear)",
                GraphFormat.Impedance => "Magnitude Impedance (Ω)",
                _ => "Value"
            };
            
            Axis xAxis = _settings.IsXLog
                ? new LogarithmicAxis { Position = AxisPosition.Bottom, Title = "Frequency (Hz)", MajorGridlineStyle = LineStyle.Solid, MinorGridlineStyle = LineStyle.Dot }
                : new LinearAxis { Position = AxisPosition.Bottom, Title = "Frequency (Hz)", MajorGridlineStyle = LineStyle.Solid, MinorGridlineStyle = LineStyle.Dot, StringFormat = "0.###E+0" };

            Axis yAxis = _settings.IsYLog
                ? new LogarithmicAxis { Position = AxisPosition.Left, Title = yTitle, MajorGridlineStyle = LineStyle.Solid, MinorGridlineStyle = LineStyle.Dot }
                : new LinearAxis { Position = AxisPosition.Left, Title = yTitle, MajorGridlineStyle = LineStyle.Solid, MinorGridlineStyle = LineStyle.Dot };

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
                    double val = 0;
                    switch (_settings.Format)
                    {
                        case GraphFormat.Db:
                            val = ToDb(point.Parameters[paramIndex].Magnitude);
                            break;
                        case GraphFormat.Linear:
                            val = point.Parameters[paramIndex].Magnitude;
                            break;
                        case GraphFormat.Impedance:
                            val = CalculateComplexImpedanceMagnitude(point.Parameters[paramIndex], sel.File.ReferenceImpedance);
                            break;
                    }
                    series.Points.Add(new DataPoint(point.FrequencyHz, val));
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



        private static double CalculateComplexImpedanceMagnitude(TouchstoneParameterValue param, double systemImpedance)
        {
            var z = CalculateComplexImpedance(param, systemImpedance);
            return z.Magnitude;
        }

        private static Complex CalculateComplexImpedance(TouchstoneParameterValue param, double systemImpedance)
        {
            var complexParam = new Complex(param.Real, param.Imaginary);

            // Determine calculation method based on parameter name
            // S11/S22 -> Reflection Method (1-port)
            // S21/S12 -> Series-Thru Method (2-port series component)
            bool isTransmission = param.Name.Contains("S21", StringComparison.OrdinalIgnoreCase) || 
                                  param.Name.Contains("S12", StringComparison.OrdinalIgnoreCase);

            if (isTransmission)
            {
                // Series-Thru Method for Series Component (e.g., Capacitor in series)
                // Z = 2 * Z0 * (1 - S21) / S21
                // Check div by zero
                if (complexParam.Magnitude < 1e-12) return new Complex(double.PositiveInfinity, double.PositiveInfinity);
                return 2 * systemImpedance * (1 - complexParam) / complexParam;
            }
            else
            {
                // Reflection Method (1-port Z)
                // Z = Z0 * (1 + S11) / (1 - S11)
                // Check div by zero
                if ((1 - complexParam).Magnitude < 1e-12) return new Complex(double.PositiveInfinity, double.PositiveInfinity);
                
                return systemImpedance * (1 + complexParam) / (1 - complexParam);
            }
        }

        private void CalculateRlc_Click(object sender, RoutedEventArgs e)
        {
            var selected = GetCurrentSelection();
            if (selected.Count == 0)
            {
                MessageBox.Show("Please select a trace first.", "RLC Fit", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            
            // Only process the first selected trace for now or iterate?
            // Let's do first one.
            var sel = selected.First();
            var paramIndex = FindParameterIndex(sel.File, sel.Param);
            if (paramIndex < 0) return;

            // Extract Z data
            var zData = new List<(double Freq, Complex Z)>();
            foreach(var p in sel.File.Points)
            {
                var z = CalculateComplexImpedance(p.Parameters[paramIndex], sel.File.ReferenceImpedance);
                if (!double.IsInfinity(z.Real) && !double.IsInfinity(z.Imaginary) && !double.IsNaN(z.Real) && !double.IsNaN(z.Imaginary))
                {
                    zData.Add((p.FrequencyHz, z));
                }
            }

            if (zData.Count < 3)
            {
                MessageBox.Show("Not enough data points for RLC fit.", "RLC Fit", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                var rlc = FitSeriesRlc(zData);

                // Initialize Tuning UI
                _rlcTarget = (sel.File, sel.Param, sel.FileName);
                _currentRlc = new RlcResult { R = rlc.R, L = rlc.L, C = rlc.C, ResonanceFreq = rlc.ResonanceFreq };

                if (_tuningWindow != null)
                {
                    _tuningWindow.Close();
                    _tuningWindow = null;
                }

                _tuningWindow = new RlcTuningWindow(rlc.R, rlc.L, rlc.C);
                _tuningWindow.Owner = this;
                
                _tuningWindow.OnParametersChanged += (r, l, c) =>
                {
                    if (_currentRlc == null) return;
                    _currentRlc.R = r;
                    _currentRlc.L = l;
                    _currentRlc.C = c;
                    UpdateRlcPlot();
                };

                _tuningWindow.Closed += (s, e) => _tuningWindow = null;
                _tuningWindow.Show();

                UpdateRlcPlot();

                double calculatedResonance = 0;
                if (rlc.L > 0 && rlc.C > 0)
                {
                    calculatedResonance = 1.0 / (2 * Math.PI * Math.Sqrt(rlc.L * rlc.C));
                }

                string msg = $"Initial Fit for {sel.FileName} - {sel.Param}:\n\n" +
                             $"Series RLC Model Fit:\n" +
                             $"R (ESR): {ToEngineeringNotation(rlc.R, "Ω")}\n" +
                             $"L (ESL): {ToEngineeringNotation(rlc.L, "H")}\n" +
                             $"C (Cap): {ToEngineeringNotation(rlc.C, "F")}\n" +
                             $"\n" +
                             $"Measured Resonance (Min |Z|): {ToEngineeringNotation(rlc.ResonanceFreq, "Hz")}\n" +
                             $"Calculated Resonance (1/2π√LC): {ToEngineeringNotation(calculatedResonance, "Hz")}";
                             
                MessageBox.Show(msg, "RLC Calculation", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error calculating RLC: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateRlcPlot()
        {
            if (PlotView.Model == null || _rlcTarget == null || _currentRlc == null) return;
            
            // ... (rest is same)
            var oldSeries = PlotView.Model.Series.Where(s => s.Title != null && s.Title.StartsWith("Fit:")).ToList();
            foreach (var s in oldSeries) PlotView.Model.Series.Remove(s);

            var sel = _rlcTarget.Value;
            var rlc = _currentRlc;

            var fitSeries = new LineSeries
            {
                Title = $"Fit: {sel.FileName} ({sel.Param})",
                Color = OxyColors.Red,
                StrokeThickness = 2,
                LineStyle = LineStyle.Dash
            };

            bool isTransmission = sel.Param.Contains("S21", StringComparison.OrdinalIgnoreCase) || 
                                  sel.Param.Contains("S12", StringComparison.OrdinalIgnoreCase);
            double z0 = sel.File.ReferenceImpedance;

            foreach (var p in sel.File.Points)
            {
                double f = p.FrequencyHz;
                if (f <= 0) continue;

                double w = 2 * Math.PI * f;
                // Z_model = R + j(wL - 1/wC)
                double reactance = (w * rlc.L) - (1.0 / (w * rlc.C));
                Complex zModel = new Complex(rlc.R, reactance);
                
                double plotVal;
                
                if (_settings.Format == GraphFormat.Impedance)
                {
                    plotVal = zModel.Magnitude;
                }
                else
                {
                     Complex sValue;
                     if (isTransmission)
                     {
                         // Series-Thru
                         Complex denom = 2 * z0 + zModel;
                         sValue = (denom.Magnitude < 1e-15) ? 0 : (2 * z0) / denom;
                     }
                     else
                     {
                         // Reflection
                         Complex numer = zModel - z0;
                         Complex denom = zModel + z0;
                         sValue = (denom.Magnitude < 1e-15) ? 0 : numer / denom;
                     }
                     
                     if (_settings.Format == GraphFormat.Db) 
                         plotVal = ToDb(sValue.Magnitude);
                     else 
                         plotVal = sValue.Magnitude;
                }

                fitSeries.Points.Add(new DataPoint(f, plotVal));
            }

            PlotView.Model.Series.Add(fitSeries);
            PlotView.InvalidatePlot(true);
        }

        private static string ToEngineeringNotation(double value, string unit)
        {
            if (value == 0) return $"0 {unit}";
            if (double.IsInfinity(value) || double.IsNaN(value)) return value.ToString();

            double mag = Math.Abs(value);
            int exponent = (int)Math.Floor(Math.Log10(mag));
            int engExponent = (exponent >= 0) ? (exponent / 3) * 3 : ((exponent - 2) / 3) * 3;
            
            double scaledValue = value / Math.Pow(10, engExponent);
            
            string prefix = engExponent switch
            {
                12 => "T",
                9 => "G",
                6 => "M",
                3 => "k",
                0 => "",
                -3 => "m",
                -6 => "µ",
                -9 => "n",
                -12 => "p",
                -15 => "f",
                _ => "?" 
            };

            if (prefix == "?")
            {
                return $"{value:E3} {unit}";
            }

            return $"{scaledValue:F3} {prefix}{unit}";
        }

        private class RlcResult
        {
            public double R { get; set; }
            public double L { get; set; }
            public double C { get; set; }
            public double ResonanceFreq { get; set; }
        }

        private RlcResult FitSeriesRlc(List<(double Freq, Complex Z)> data)
        {
            // Simple heuristic fit
            
            // 1. Find Resonance (Phase crossing 0, or Min Magnitude)
            // Let's use Min Magnitude of Z as resonance point for Series RLC
            var minZ = data.OrderBy(d => d.Z.Magnitude).First();
            double fRes = minZ.Freq;
            double R = minZ.Z.Real; 

            // 2. Estimate L and C constrained by Resonance Frequency
            // Relation: fRes = 1 / (2 * pi * sqrt(L * C))
            // So: L * C = 1 / (2 * pi * fRes)^2 = K
            // We need to find L such that C = K/L gives best fit, or find C such that L = K/C gives best fit.
            
            // Let's estimate C from low frequency slope ( capacitive region )
            // Xc = -1 / (wC)  => C = -1 / (w * X)
            double wRes = 2 * Math.PI * fRes;
            double LC_Product = 1.0 / (wRes * wRes);

            // Collect points for C (Im(Z) < 0 and f < fRes)
            var cPoints = data.Where(d => d.Freq < fRes && d.Z.Imaginary < 0).ToList();
            double C_est = 0;
            if (cPoints.Count > 0)
            {
                // Average C = -1 / ( (2*pi*f) * Im(Z) )
                double sumC = 0;
                int validCount = 0;
                foreach(var p in cPoints)
                {
                    double w = 2 * Math.PI * p.Freq;
                    double val = -1.0 / (w * p.Z.Imaginary);
                    if (val > 0 && val < 1) // Sanity check
                    {
                        sumC += val;
                        validCount++;
                    }
                }
                if (validCount > 0) C_est = sumC / validCount;
            }

            // Calculate L based on fRes and C_est
            double L_calc = 0;
            if (C_est > 0)
            {
                L_calc = LC_Product / C_est;
            }
            
            // If we couldn't estimate C (maybe no low freq points), try estimating L from high freq
            if (C_est == 0)
            {
                 var lPoints = data.Where(d => d.Freq > fRes && d.Z.Imaginary > 0).ToList();
                 double L_est = 0;
                 if (lPoints.Count > 0)
                 {
                     double sumL = 0;
                     int validCount = 0;
                     foreach(var p in lPoints)
                     {
                         double w = 2 * Math.PI * p.Freq;
                         double val = p.Z.Imaginary / w;
                         if (val > 0)
                         {
                             sumL += val;
                             validCount++;
                         }
                     }
                     if (validCount > 0) L_est = sumL / validCount;
                 }
                 
                 if (L_est > 0)
                 {
                     L_calc = L_est;
                     C_est = LC_Product / L_calc;
                 }
            }

            return new RlcResult { R = R, L = L_calc, C = C_est, ResonanceFreq = fRes };
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

    public enum GraphFormat
    {
        Db,
        Linear,
        Impedance
    }

    public sealed class PlotSettings
    {
        public GraphFormat Format { get; set; } = GraphFormat.Db;
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
                Format = Format,
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
