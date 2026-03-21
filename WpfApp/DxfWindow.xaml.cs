using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using IxMilia.Dxf;
using IxMilia.Dxf.Entities;
using ACadSharp;
using ACadSharp.IO;
using Microsoft.Win32;

namespace WpfApp
{
    // ── Data models ────────────────────────────────────────────────
    public class LayerViewModel : System.ComponentModel.INotifyPropertyChanged
    {
        private bool _isVisible = true;
        public string Name { get; set; } = "";
        public bool IsVisible
        {
            get => _isVisible;
            set { _isVisible = value; PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(IsVisible))); }
        }
        public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;
    }

    public class CapacitorAssignment
    {
        public string FileName { get; set; } = string.Empty;
        public string Coordinates { get; set; } = string.Empty;
        public Point Location { get; set; }
        public string FilePath { get; set; } = string.Empty;
        public string RlcInfo { get; set; } = string.Empty;
        public string DomainName { get; set; } = string.Empty;
        public RlcResult? Rlc { get; set; }
    }

    // ── Configuration persistence models ──────────────────────────
    public class DxfSessionConfig
    {
        public string LayoutFilePath { get; set; } = "";
        public Dictionary<string, bool>? LayerStates { get; set; }
        public double ViewCenterX { get; set; }
        public double ViewCenterY { get; set; }
        public double ZoomLevel { get; set; }
        public BoardParameters? BoardParams { get; set; }
        public List<ProbeConfig> Probes { get; set; } = new();
        public List<CapacitorConfig> Capacitors { get; set; } = new();
    }

    public class BoardParameters
    {
        public double PlaneResistivity { get; set; } = 1.72e-8;   // Copper Ω·m
        public double PlaneThickness { get; set; } = 35e-6;       // 35 µm (1 oz Cu)
        public double ProbeSeriesR { get; set; }
        public double ProbeSeriesL { get; set; }
        public double ProbeParallelC { get; set; }
    }

    public class ProbeConfig
    {
        public double X { get; set; }
        public double Y { get; set; }
        public string Domain { get; set; } = "";
    }

    public class CapacitorConfig
    {
        public double X { get; set; }
        public double Y { get; set; }
        public string Name { get; set; } = "";
        public string FilePath { get; set; } = "";
        public double FitR { get; set; }
        public double FitL { get; set; }
        public double FitC { get; set; }
        public double ResonanceFreq { get; set; }
        public List<double> ProbeDistances { get; set; } = new();
    }

    // ── Main window ───────────────────────────────────────────────
    public partial class DxfWindow : Window
    {
        // Events to notify parent (MainWindow) about loaded capacitor files
        public event EventHandler<string>? CapacitorFileLoaded;

        // DXF / DWG data
        private DxfFile? _currentDxfFile;
        private CadDocument? _currentDwgDocument;
        private string _currentLayoutFilePath = "";
        private Rect _contentBounds;

        // Layers
        private readonly ObservableCollection<LayerViewModel> _layers = new();

        // Zoom / Pan
        private double _zoomFactor = 1.0;
        private const double ZoomStep = 1.15;
        private bool _isPanning;
        private Point _panStart;
        private Matrix _panStartMatrix;

        // Probe mode
        private bool _isProbeMode;
        private readonly List<Point> _probePoints = new();
        private readonly List<string> _probeDomains = new();       // domain name per probe
        private string _currentProbeDomain = "";                   // domain being placed

        // Capacitor mode
        private bool _isCapacitorMode;
        private readonly List<CapacitorAssignment> _capacitorAssignments = new();

        // Color palettes
        private static readonly Brush[] DomainColors = new Brush[]
        {
            Brushes.Lime, Brushes.Cyan, Brushes.Yellow, Brushes.Orange,
            Brushes.Magenta, Brushes.Red, Brushes.DeepSkyBlue, Brushes.SpringGreen,
            Brushes.Gold, Brushes.Orchid, Brushes.Coral, Brushes.Aquamarine
        };
        private readonly Dictionary<string, Brush> _domainColorMap = new(StringComparer.OrdinalIgnoreCase);

        private static readonly Brush[] CapFileColors = new Brush[]
        {
            Brushes.DeepSkyBlue, Brushes.Tomato, Brushes.MediumSeaGreen, Brushes.Gold,
            Brushes.Orchid, Brushes.Coral, Brushes.Cyan, Brushes.LimeGreen,
            Brushes.HotPink, Brushes.SandyBrown, Brushes.Turquoise, Brushes.Plum
        };
        private readonly Dictionary<string, Brush> _capFileColorMap = new(StringComparer.OrdinalIgnoreCase);

        // Settings
        private BoardParameters _boardParams = new();

        public DxfWindow()
        {
            InitializeComponent();
            LayersList.ItemsSource = _layers;
        }

        // ─────────────────────── FILE I/O ──────────────────────────
        private void OpenLayout_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "CAD Files (*.dxf;*.dwg)|*.dxf;*.dwg|DXF Files (*.dxf)|*.dxf|DWG Files (*.dwg)|*.dwg|All Files (*.*)|*.*"
            };
            if (dlg.ShowDialog() != true) return;
            LoadLayout(dlg.FileName);
        }

        private void LoadLayout(string path, Dictionary<string, bool>? savedLayerStates = null)
        {
            try
            {
                _currentLayoutFilePath = path;
                _currentDxfFile = null;
                _currentDwgDocument = null;

                string ext = System.IO.Path.GetExtension(path).ToLowerInvariant();
                if (ext == ".dwg")
                {
                    using var reader = new DwgReader(path);
                    _currentDwgDocument = reader.Read();
                }
                else
                {
                    using var fs = File.OpenRead(path);
                    _currentDxfFile = DxfFile.Load(fs);
                }

                RenderLayout(savedLayerStates);
                FitView();
                Title = $"Layout Viewer – {System.IO.Path.GetFileName(path)}";
                StatusText.Text = $"Loaded: {System.IO.Path.GetFileName(path)}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load layout: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ─────────────────────── RENDERING ─────────────────────────
        private void RenderLayout(Dictionary<string, bool>? savedLayerStates = null)
        {
            DxfCanvas.Children.Clear();
            _layers.Clear();

            double minX = double.MaxValue, maxX = double.MinValue;
            double minY = double.MaxValue, maxY = double.MinValue;
            bool hasContent = false;

            IEnumerable<(double x1, double y1, double x2, double y2, string layer, System.Drawing.Color? color)> segments;

            if (_currentDwgDocument != null)
                segments = ExtractDwgSegments();
            else if (_currentDxfFile != null)
                segments = ExtractDxfSegments();
            else
                return;

            var layerNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var seg in segments)
            {
                layerNames.Add(seg.layer);

                var line = new Line
                {
                    X1 = seg.x1,
                    Y1 = -seg.y1,   // flip Y
                    X2 = seg.x2,
                    Y2 = -seg.y2,
                    Stroke = seg.color.HasValue
                        ? new SolidColorBrush(System.Windows.Media.Color.FromRgb(seg.color.Value.R, seg.color.Value.G, seg.color.Value.B))
                        : Brushes.White,
                    StrokeThickness = 2,
                    Tag = seg.layer
                };

                DxfCanvas.Children.Add(line);
                hasContent = true;

                void Expand(double x, double y) { if (x < minX) minX = x; if (x > maxX) maxX = x; if (y < minY) minY = y; if (y > maxY) maxY = y; }
                Expand(seg.x1, -seg.y1);
                Expand(seg.x2, -seg.y2);
            }

            // Populate layers
            foreach (var name in layerNames.OrderBy(n => n))
            {
                bool vis = true;
                if (savedLayerStates != null && savedLayerStates.TryGetValue(name, out bool v)) vis = v;
                _layers.Add(new LayerViewModel { Name = name, IsVisible = vis });
            }

            if (hasContent)
                _contentBounds = new Rect(minX, minY, maxX - minX, maxY - minY);
            else
                _contentBounds = new Rect(0, 0, 100, 100);

            // Apply saved visibility
            if (savedLayerStates != null) ApplyLayerVisibility();
        }

        // ── DXF segment extraction ──
        private IEnumerable<(double, double, double, double, string, System.Drawing.Color?)> ExtractDxfSegments()
        {
            if (_currentDxfFile == null) yield break;

            // Build layer-name → color lookup from file's layer table
            var layerColorMap = new Dictionary<string, System.Drawing.Color?>(StringComparer.OrdinalIgnoreCase);
            foreach (var dxfLayer in _currentDxfFile.Layers)
            {
                try
                {
                    var rgb = dxfLayer.Color.ToRGB();
                    layerColorMap[dxfLayer.Name] = System.Drawing.Color.FromArgb(255, (byte)((rgb >> 16) & 0xFF), (byte)((rgb >> 8) & 0xFF), (byte)(rgb & 0xFF));
                }
                catch { }
            }

            foreach (var entity in _currentDxfFile.Entities)
            {
                string layer = entity.Layer ?? "0";
                System.Drawing.Color? c = null;

                if (!entity.Color.IsByLayer)
                {
                    // Entity has its own explicit color
                    try { var rgb = entity.Color.ToRGB(); c = System.Drawing.Color.FromArgb(255, (byte)((rgb >> 16) & 0xFF), (byte)((rgb >> 8) & 0xFF), (byte)(rgb & 0xFF)); } catch { }
                }
                else
                {
                    // Resolve from layer color
                    if (layerColorMap.TryGetValue(layer, out var lc)) c = lc;
                }

                if (entity is DxfLine ln)
                {
                    yield return (ln.P1.X, ln.P1.Y, ln.P2.X, ln.P2.Y, layer, c);
                }
                else if (entity is DxfCircle circ)
                {
                    int segs = 64;
                    for (int i = 0; i < segs; i++)
                    {
                        double a1 = 2 * Math.PI * i / segs;
                        double a2 = 2 * Math.PI * (i + 1) / segs;
                        yield return (
                            circ.Center.X + circ.Radius * Math.Cos(a1),
                            circ.Center.Y + circ.Radius * Math.Sin(a1),
                            circ.Center.X + circ.Radius * Math.Cos(a2),
                            circ.Center.Y + circ.Radius * Math.Sin(a2),
                            layer, c);
                    }
                }
                else if (entity is DxfArc arc)
                {
                    int segs = 32;
                    double sa = arc.StartAngle * Math.PI / 180;
                    double ea = arc.EndAngle * Math.PI / 180;
                    if (ea <= sa) ea += 2 * Math.PI;
                    for (int i = 0; i < segs; i++)
                    {
                        double a1 = sa + (ea - sa) * i / segs;
                        double a2 = sa + (ea - sa) * (i + 1) / segs;
                        yield return (
                            arc.Center.X + arc.Radius * Math.Cos(a1),
                            arc.Center.Y + arc.Radius * Math.Sin(a1),
                            arc.Center.X + arc.Radius * Math.Cos(a2),
                            arc.Center.Y + arc.Radius * Math.Sin(a2),
                            layer, c);
                    }
                }
                else if (entity is DxfLwPolyline poly)
                {
                    for (int i = 0; i < poly.Vertices.Count - 1; i++)
                    {
                        var v1 = poly.Vertices[i];
                        var v2 = poly.Vertices[i + 1];
                        yield return (v1.X, v1.Y, v2.X, v2.Y, layer, c);
                    }
                    if (poly.IsClosed && poly.Vertices.Count > 2)
                    {
                        var vf = poly.Vertices[0];
                        var vl = poly.Vertices[poly.Vertices.Count - 1];
                        yield return (vl.X, vl.Y, vf.X, vf.Y, layer, c);
                    }
                }
                else if (entity is DxfPolyline pline)
                {
                    var verts = pline.Vertices.ToList();
                    for (int i = 0; i < verts.Count - 1; i++)
                    {
                        yield return (verts[i].Location.X, verts[i].Location.Y,
                                     verts[i + 1].Location.X, verts[i + 1].Location.Y, layer, c);
                    }
                    if (pline.IsClosed && verts.Count > 2)
                        yield return (verts.Last().Location.X, verts.Last().Location.Y,
                                      verts[0].Location.X, verts[0].Location.Y, layer, c);
                }
            }
        }

        // ── DWG segment extraction ──
        private System.Drawing.Color? ResolveDwgEntityColor(ACadSharp.Entities.Entity entity)
        {
            // Check if entity has its own color (not ByLayer)
            var ec = entity.Color;
            if (ec.IsByLayer)
            {
                // Use layer color
                var layerObj = entity.Layer;
                if (layerObj != null)
                {
                    var lc = layerObj.Color;
                    byte r = lc.R, g = lc.G, b = lc.B;
                    if (r != 0 || g != 0 || b != 0)
                        return System.Drawing.Color.FromArgb(255, r, g, b);
                }
                return null;
            }
            // Entity has explicit color
            {
                byte r = ec.R, g = ec.G, b = ec.B;
                if (r != 0 || g != 0 || b != 0)
                    return System.Drawing.Color.FromArgb(255, r, g, b);
            }
            return null;
        }

        private IEnumerable<(double, double, double, double, string, System.Drawing.Color?)> ExtractDwgSegments()
        {
            if (_currentDwgDocument == null) yield break;

            foreach (var entity in _currentDwgDocument.Entities)
            {
                string layer = entity.Layer?.Name ?? "0";
                System.Drawing.Color? c = ResolveDwgEntityColor(entity);

                if (entity is ACadSharp.Entities.Line ln)
                {
                    yield return (ln.StartPoint.X, ln.StartPoint.Y, ln.EndPoint.X, ln.EndPoint.Y, layer, c);
                }
                else if (entity is ACadSharp.Entities.Circle circ)
                {
                    int segs = 64;
                    for (int i = 0; i < segs; i++)
                    {
                        double a1 = 2 * Math.PI * i / segs;
                        double a2 = 2 * Math.PI * (i + 1) / segs;
                        yield return (
                            circ.Center.X + circ.Radius * Math.Cos(a1),
                            circ.Center.Y + circ.Radius * Math.Sin(a1),
                            circ.Center.X + circ.Radius * Math.Cos(a2),
                            circ.Center.Y + circ.Radius * Math.Sin(a2),
                            layer, c);
                    }
                }
                else if (entity is ACadSharp.Entities.Arc arc)
                {
                    int segs = 32;
                    double sa = arc.StartAngle;
                    double ea = arc.EndAngle;
                    if (ea <= sa) ea += 2 * Math.PI;
                    for (int i = 0; i < segs; i++)
                    {
                        double a1 = sa + (ea - sa) * i / segs;
                        double a2 = sa + (ea - sa) * (i + 1) / segs;
                        yield return (
                            arc.Center.X + arc.Radius * Math.Cos(a1),
                            arc.Center.Y + arc.Radius * Math.Sin(a1),
                            arc.Center.X + arc.Radius * Math.Cos(a2),
                            arc.Center.Y + arc.Radius * Math.Sin(a2),
                            layer, c);
                    }
                }
                else if (entity is ACadSharp.Entities.LwPolyline poly)
                {
                    var verts = poly.Vertices.ToList();
                    for (int i = 0; i < verts.Count - 1; i++)
                    {
                        yield return (verts[i].Location.X, verts[i].Location.Y,
                                     verts[i + 1].Location.X, verts[i + 1].Location.Y, layer, c);
                    }
                    if (poly.IsClosed && verts.Count > 2)
                        yield return (verts.Last().Location.X, verts.Last().Location.Y,
                                      verts[0].Location.X, verts[0].Location.Y, layer, c);
                }
            }
        }

        // ─────────────────────── VIEW CONTROLS ─────────────────────
        private void FitView()
        {
            // Compute bounds of only the visible elements (respecting layer visibility)
            var visibleSet = new HashSet<string>(_layers.Where(l => l.IsVisible).Select(l => l.Name), StringComparer.OrdinalIgnoreCase);

            double minX = double.MaxValue, maxX = double.MinValue;
            double minY = double.MaxValue, maxY = double.MinValue;
            bool found = false;

            foreach (UIElement child in DxfCanvas.Children)
            {
                if (child is Line line && line.Tag is string tag &&
                    tag != "__PROBE__" && tag != "__CAP__" &&
                    visibleSet.Contains(tag) && line.Visibility == Visibility.Visible)
                {
                    void Expand(double x, double y)
                    {
                        if (x < minX) minX = x; if (x > maxX) maxX = x;
                        if (y < minY) minY = y; if (y > maxY) maxY = y;
                    }
                    Expand(line.X1, line.Y1);
                    Expand(line.X2, line.Y2);
                    found = true;
                }
            }

            if (!found)
            {
                // Fallback to full content bounds
                if (_contentBounds.Width <= 0 || _contentBounds.Height <= 0) return;
                minX = _contentBounds.X; maxX = _contentBounds.X + _contentBounds.Width;
                minY = _contentBounds.Y; maxY = _contentBounds.Y + _contentBounds.Height;
            }

            double boundsW = maxX - minX;
            double boundsH = maxY - minY;
            if (boundsW <= 0 || boundsH <= 0) return;

            double viewW = MainGrid.ColumnDefinitions[0].ActualWidth;
            double viewH = ActualHeight - 80; // minus menu + status
            if (viewW <= 0 || viewH <= 0) { viewW = 900; viewH = 700; }

            double scaleX = viewW / boundsW;
            double scaleY = viewH / boundsH;
            double scale = Math.Min(scaleX, scaleY) * 0.9;

            double cx = minX + boundsW / 2;
            double cy = minY + boundsH / 2;

            var m = new Matrix();
            m.Translate(-cx, -cy);
            m.Scale(scale, scale);
            m.Translate(viewW / 2, viewH / 2);

            CanvasTransform.Matrix = m;
            _zoomFactor = scale;
            UpdateZoomText();
        }

        private void ResetView_Click(object sender, RoutedEventArgs e) => FitView();

        private void CenterView_Click(object sender, RoutedEventArgs e)
        {
            if (_contentBounds.Width <= 0) return;

            double viewW = MainGrid.ColumnDefinitions[0].ActualWidth;
            double viewH = ActualHeight - 80;
            if (viewW <= 0 || viewH <= 0) return;

            double cx = _contentBounds.X + _contentBounds.Width / 2;
            double cy = _contentBounds.Y + _contentBounds.Height / 2;

            var m = new Matrix();
            m.Translate(-cx, -cy);
            m.Scale(_zoomFactor, _zoomFactor);
            m.Translate(viewW / 2, viewH / 2);

            CanvasTransform.Matrix = m;
            UpdateZoomText();
        }

        private void Canvas_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) != ModifierKeys.Control)
                return;

            // Mouse position in screen (Border) coordinates
            var screenPos = e.GetPosition((UIElement)sender);
            double s = e.Delta > 0 ? ZoomStep : 1.0 / ZoomStep;

            // Transform: screen = M * canvas
            // We want the point under the mouse to stay fixed after scaling.
            // 1. Find the canvas-space point under the mouse
            var m = CanvasTransform.Matrix;
            var inv = m;
            inv.Invert();
            var canvasPoint = inv.Transform(screenPos);

            // 2. Apply scale at that canvas point
            m.ScaleAt(s, s, canvasPoint.X, canvasPoint.Y);

            CanvasTransform.Matrix = m;
            _zoomFactor *= s;
            UpdateZoomText();
            e.Handled = true;
        }

        // Pan with right mouse
        private void Canvas_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            _isPanning = true;
            _panStart = e.GetPosition(this);
            _panStartMatrix = CanvasTransform.Matrix;
            ((UIElement)sender).CaptureMouse();
            e.Handled = true;
        }

        private void Canvas_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            _isPanning = false;
            ((UIElement)sender).ReleaseMouseCapture();
        }

        private void Canvas_MouseMove(object sender, MouseEventArgs e)
        {
            // Update coordinate display
            var canvasPt = e.GetPosition(DxfCanvas);
            CoordsText.Text = $"X: {canvasPt.X:F2}  Y: {(-canvasPt.Y):F2}";

            if (_isPanning && e.RightButton == MouseButtonState.Pressed)
            {
                var cur = e.GetPosition(this);
                double dx = cur.X - _panStart.X;
                double dy = cur.Y - _panStart.Y;
                var m = _panStartMatrix;
                m.Translate(dx, dy);
                CanvasTransform.Matrix = m;
            }
        }

        // Left-click: probe/cap placement
        private void Canvas_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var pt = e.GetPosition(DxfCanvas);
            double cadX = pt.X;
            double cadY = -pt.Y; // flip back to CAD Y

            if (_isProbeMode)
            {
                _probePoints.Add(new Point(cadX, cadY));
                _probeDomains.Add(_currentProbeDomain);
                AddProbeMarker(cadX, cadY, _currentProbeDomain);
                UpdateProbeInfo();
            }
            else if (_isCapacitorMode)
            {
                // Ask user to pick an S-Parameter file
                var dlg = new OpenFileDialog
                {
                    Filter = "Touchstone (*.s*p;*.spd)|*.s*p;*.spd|All files (*.*)|*.*",
                    Title = "Select Capacitor S-Parameter File"
                };
                if (dlg.ShowDialog() != true) return;

                try
                {
                    var fileData = TouchstoneParser.Parse(dlg.FileName);

                    // Calculate RLC fit
                    var rlc = CalculateRlcForFile(fileData);

                    string rlcInfo = $"R={RlcHelper.ToEngineeringNotation(rlc.R, "Ω")}  " +
                                     $"L={RlcHelper.ToEngineeringNotation(rlc.L, "H")}  " +
                                     $"C={RlcHelper.ToEngineeringNotation(rlc.C, "F")}  " +
                                     $"fRes={RlcHelper.ToEngineeringNotation(rlc.ResonanceFreq, "Hz")}";

                    var assignment = new CapacitorAssignment
                    {
                        FileName = fileData.FileName,
                        FilePath = dlg.FileName,
                        Location = new Point(cadX, cadY),
                        Coordinates = $"({cadX:F2}, {cadY:F2})",
                        RlcInfo = rlcInfo,
                        Rlc = rlc
                    };
                    _capacitorAssignments.Add(assignment);

                    AddCapacitorMarker(cadX, cadY, fileData.FileName, dlg.FileName);
                    UpdateCapacitorInfo();

                    // Notify parent to load file in graph
                    CapacitorFileLoaded?.Invoke(this, dlg.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error loading capacitor file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void Canvas_MouseLeftButtonUp(object sender, MouseButtonEventArgs e) { }

        private void UpdateZoomText() => ZoomText.Text = $"Zoom: {_zoomFactor * 100:F0}%";

        // ─────────────────────── LAYERS ────────────────────────────
        private void LayerVisibilityChanged(object sender, RoutedEventArgs e)
        {
            ApplyLayerVisibility();
        }

        private void LayersSelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var l in _layers) l.IsVisible = true;
            ApplyLayerVisibility();
        }

        private void LayersSelectNone_Click(object sender, RoutedEventArgs e)
        {
            foreach (var l in _layers) l.IsVisible = false;
            ApplyLayerVisibility();
        }

        private void ApplyLayerVisibility()
        {
            var visibleSet = new HashSet<string>(_layers.Where(l => l.IsVisible).Select(l => l.Name), StringComparer.OrdinalIgnoreCase);
            foreach (UIElement child in DxfCanvas.Children)
            {
                if (child is FrameworkElement fe && fe.Tag is string layerName)
                {
                    // Markers have special tags
                    if (layerName == "__PROBE__" || layerName == "__CAP__") continue;
                    fe.Visibility = visibleSet.Contains(layerName) ? Visibility.Visible : Visibility.Collapsed;
                }
            }
        }

        // ─────────────────────── TOOLBAR ──────────────────────────
        private void ToolBar_Loaded(object sender, RoutedEventArgs e)
        {
            // Hide the overflow grip on the toolbar
            if (sender is ToolBar tb)
            {
                var overflow = tb.Template.FindName("OverflowGrid", tb) as FrameworkElement;
                if (overflow != null) overflow.Visibility = Visibility.Collapsed;
            }
        }

        private void SyncModeButtons()
        {
            BtnProbeMode.IsChecked = _isProbeMode;
            BtnCapMode.IsChecked = _isCapacitorMode;
            var pickVis = (_isProbeMode || _isCapacitorMode) ? Visibility.Visible : Visibility.Collapsed;
            BtnUndo.Visibility = pickVis;
            BtnClear.Visibility = pickVis;
        }

        private void UndoLast_Click(object sender, RoutedEventArgs e)
        {
            if (_isProbeMode && _probePoints.Count > 0)
            {
                _probePoints.RemoveAt(_probePoints.Count - 1);
                if (_probeDomains.Count > _probePoints.Count)
                    _probeDomains.RemoveAt(_probeDomains.Count - 1);
                RemoveLastMarkerByTag("__PROBE__");
                UpdateProbeInfo();
            }
            else if (_isCapacitorMode && _capacitorAssignments.Count > 0)
            {
                _capacitorAssignments.RemoveAt(_capacitorAssignments.Count - 1);
                RemoveLastMarkerByTag("__CAP__");
                UpdateCapacitorInfo();
            }
            else if (!_isProbeMode && !_isCapacitorMode)
            {
                // Not in a mode: undo from whichever was placed last
                StatusText.Text = "Enter Probe or Capacitor mode first, or use Del key.";
            }
        }

        private void ClearAll_Click(object sender, RoutedEventArgs e)
        {
            if (_isProbeMode)
            {
                _probePoints.Clear();
                _probeDomains.Clear();
                RemoveMarkersByTag("__PROBE__");
                UpdateProbeInfo();
            }
            else if (_isCapacitorMode)
            {
                _capacitorAssignments.Clear();
                _capFileColorMap.Clear();
                RemoveMarkersByTag("__CAP__");
                UpdateCapacitorInfo();
            }
            else
            {
                StatusText.Text = "Enter Probe or Capacitor mode first, or use Shift+C.";
            }
        }

        // ─────────────────────── PROBES ────────────────────────────
        private Brush GetDomainBrush(string domain)
        {
            if (string.IsNullOrWhiteSpace(domain)) domain = "Default";
            if (!_domainColorMap.TryGetValue(domain, out var brush))
            {
                brush = DomainColors[_domainColorMap.Count % DomainColors.Length];
                _domainColorMap[domain] = brush;
            }
            return brush;
        }

        private void DefineProbes_Click(object sender, RoutedEventArgs e)
        {
            if (_isProbeMode)
            {
                // Finishing probe mode – ask which power domain these probes belong to
                _isProbeMode = false;
                _isCapacitorMode = false;
                SyncModeButtons();
                StatusText.Text = "Probe mode OFF";
                return;
            }

            // Starting probe mode – ask for the power domain name first
            var dlg = new PromptDialog("Power Domain", "Enter the power domain name for the probes you are about to place:", _currentProbeDomain);
            dlg.Owner = this;
            if (dlg.ShowDialog() != true || string.IsNullOrWhiteSpace(dlg.ResponseText))
            {
                StatusText.Text = "Probe placement cancelled.";
                return;
            }
            _currentProbeDomain = dlg.ResponseText.Trim();

            _isProbeMode = true;
            _isCapacitorMode = false;
            SyncModeButtons();
            StatusText.Text = $"PROBE MODE [{_currentProbeDomain}]: Left-click to place probes. Press Escape or toggle off to finish.";
            UpdateProbeInfo();
        }

        private void AddProbeMarker(double cadX, double cadY, string domain)
        {
            double size = 10.0 / Math.Max(_zoomFactor, 0.001);
            var fill = GetDomainBrush(domain);
            var marker = new Ellipse
            {
                Width = size,
                Height = size,
                Fill = fill,
                Stroke = Brushes.Black,
                StrokeThickness = size * 0.1,
                Tag = "__PROBE__",
                Opacity = 0.85,
                ToolTip = $"Probe [{domain}] ({cadX:F2}, {cadY:F2})"
            };
            Canvas.SetLeft(marker, cadX - size / 2);
            Canvas.SetTop(marker, -cadY - size / 2);
            Panel.SetZIndex(marker, 100);
            DxfCanvas.Children.Add(marker);
        }

        private void UpdateProbeInfo()
        {
            if (_probePoints.Count == 0)
            {
                InfoText.Text = "No probes defined.";
                return;
            }
            var sb = new System.Text.StringBuilder();
            // Group by domain
            var groups = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < _probePoints.Count; i++)
            {
                string dom = i < _probeDomains.Count ? _probeDomains[i] : "";
                if (!groups.ContainsKey(dom)) groups[dom] = new List<int>();
                groups[dom].Add(i);
            }
            foreach (var kv in groups)
            {
                sb.AppendLine($"Domain: {kv.Key}  ({kv.Value.Count} probes)");
                foreach (int i in kv.Value)
                    sb.AppendLine($"  P{i + 1}: ({_probePoints[i].X:F2}, {_probePoints[i].Y:F2})");
            }
            InfoText.Text = sb.ToString();
        }

        // ─────────────────────── CAPACITORS ────────────────────────
        private void DefineCapacitors_Click(object sender, RoutedEventArgs e)
        {
            if (_isCapacitorMode)
            {
                _isCapacitorMode = false;
                _isProbeMode = false;
                StatusText.Text = "Capacitor mode OFF";
                SyncModeButtons();
                return;
            }

            _isCapacitorMode = true;
            _isProbeMode = false;
            SyncModeButtons();
            StatusText.Text = "CAPACITOR MODE: Left-click to place capacitors (will ask for S-parameter file). Right-click to pan.";
            UpdateCapacitorInfo();
        }

        private Brush GetCapFileBrush(string filePath)
        {
            string key = System.IO.Path.GetFileName(filePath);
            if (!_capFileColorMap.TryGetValue(key, out var brush))
            {
                brush = CapFileColors[_capFileColorMap.Count % CapFileColors.Length];
                _capFileColorMap[key] = brush;
            }
            return brush;
        }

        private void AddCapacitorMarker(double cadX, double cadY, string label, string filePath = "")
        {
            double size = 12.0 / Math.Max(_zoomFactor, 0.001);
            var fill = string.IsNullOrEmpty(filePath) ? Brushes.DeepSkyBlue : GetCapFileBrush(filePath);
            var marker = new Rectangle
            {
                Width = size,
                Height = size,
                Fill = fill,
                Stroke = Brushes.White,
                StrokeThickness = size * 0.08,
                Tag = "__CAP__",
                Opacity = 0.9,
                ToolTip = $"Cap: {label}\n({cadX:F2}, {cadY:F2})"
            };
            Canvas.SetLeft(marker, cadX - size / 2);
            Canvas.SetTop(marker, -cadY - size / 2);
            Panel.SetZIndex(marker, 101);
            DxfCanvas.Children.Add(marker);
        }

        private void UpdateCapacitorInfo()
        {
            if (_capacitorAssignments.Count == 0)
            {
                InfoText.Text = "No capacitors defined.";
                return;
            }
            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"Capacitors: {_capacitorAssignments.Count}");
            foreach (var cap in _capacitorAssignments)
            {
                sb.AppendLine($"  {cap.FileName}");
                sb.AppendLine($"    {cap.Coordinates}");
                sb.AppendLine($"    {cap.RlcInfo}");
            }
            InfoText.Text = sb.ToString();
        }

        private RlcResult CalculateRlcForFile(TouchstoneFileData fileData)
        {
            // Find S21 or fallback to S11
            int paramIndex = -1;
            for (int i = 0; i < fileData.ParameterNames.Count; i++)
            {
                if (fileData.ParameterNames[i].Contains("S21", StringComparison.OrdinalIgnoreCase) ||
                    fileData.ParameterNames[i].Contains("S12", StringComparison.OrdinalIgnoreCase))
                { paramIndex = i; break; }
            }
            if (paramIndex < 0)
            {
                for (int i = 0; i < fileData.ParameterNames.Count; i++)
                {
                    if (fileData.ParameterNames[i].Contains("S11", StringComparison.OrdinalIgnoreCase))
                    { paramIndex = i; break; }
                }
            }
            if (paramIndex < 0 && fileData.ParameterNames.Count > 0) paramIndex = 0;
            if (paramIndex < 0) return new RlcResult();

            var zData = new List<(double Freq, Complex Z)>();
            foreach (var p in fileData.Points)
            {
                var z = RlcHelper.CalculateComplexImpedance(p.Parameters[paramIndex], fileData.ReferenceImpedance);
                if (!double.IsInfinity(z.Real) && !double.IsInfinity(z.Imaginary) &&
                    !double.IsNaN(z.Real) && !double.IsNaN(z.Imaginary))
                    zData.Add((p.FrequencyHz, z));
            }
            return RlcHelper.FitSeriesRlc(zData);
        }

        // ─────────────────────── SETTINGS ──────────────────────────
        private void Settings_Click(object sender, RoutedEventArgs e)
        {
            var win = new SettingsEditorWindow(_boardParams);
            win.Owner = this;
            if (win.ShowDialog() == true)
            {
                _boardParams = win.Result;
            }
        }

        // ─────────────────────── SAVE CONFIG ───────────────────────
        private void SaveConfig_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_currentLayoutFilePath))
            {
                MessageBox.Show("No layout loaded.", "Save", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dlg = new SaveFileDialog { Filter = "Layout Config (*.lconfig)|*.lconfig", FileName = "LayoutConfig.lconfig" };
            if (dlg.ShowDialog() != true) return;

            // Get current view center
            double viewCX = 0, viewCY = 0;
            try
            {
                var mat = CanvasTransform.Matrix;
                if (mat.HasInverse)
                {
                    double vw = MainGrid.ColumnDefinitions[0].ActualWidth;
                    double vh = ActualHeight - 80;
                    var inv = mat; inv.Invert();
                    var cp = inv.Transform(new Point(vw / 2, vh / 2));
                    viewCX = cp.X;
                    viewCY = -cp.Y;
                }
            }
            catch { }

            var config = new DxfSessionConfig
            {
                LayoutFilePath = _currentLayoutFilePath,
                LayerStates = _layers.ToDictionary(l => l.Name, l => l.IsVisible),
                ViewCenterX = viewCX,
                ViewCenterY = viewCY,
                ZoomLevel = _zoomFactor,
                BoardParams = _boardParams
            };

            // Probes
            for (int i = 0; i < _probePoints.Count; i++)
            {
                string dom = i < _probeDomains.Count ? _probeDomains[i] : "";
                config.Probes.Add(new ProbeConfig { X = _probePoints[i].X, Y = _probePoints[i].Y, Domain = dom });
            }

            // Capacitors (with distances to probes)
            foreach (var cap in _capacitorAssignments)
            {
                var cc = new CapacitorConfig
                {
                    X = cap.Location.X,
                    Y = cap.Location.Y,
                    Name = cap.FileName,
                    FilePath = cap.FilePath,
                    FitR = cap.Rlc?.R ?? 0,
                    FitL = cap.Rlc?.L ?? 0,
                    FitC = cap.Rlc?.C ?? 0,
                    ResonanceFreq = cap.Rlc?.ResonanceFreq ?? 0
                };

                // Compute distances from this cap to each probe
                foreach (var probe in _probePoints)
                {
                    double dist = Math.Sqrt(Math.Pow(cap.Location.X - probe.X, 2) + Math.Pow(cap.Location.Y - probe.Y, 2));
                    cc.ProbeDistances.Add(dist);
                }

                config.Capacitors.Add(cc);
            }

            try
            {
                var json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(dlg.FileName, json);
                MessageBox.Show("Configuration saved.", "Save", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ─────────────────────── LOAD CONFIG ───────────────────────
        private void LoadConfig_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "Layout Config (*.lconfig)|*.lconfig" };
            if (dlg.ShowDialog() != true) return;

            try
            {
                string json = File.ReadAllText(dlg.FileName);
                var config = JsonSerializer.Deserialize<DxfSessionConfig>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                if (config == null) return;

                // Board params
                if (config.BoardParams != null) _boardParams = config.BoardParams;

                // Load layout
                if (!string.IsNullOrEmpty(config.LayoutFilePath) && File.Exists(config.LayoutFilePath))
                {
                    LoadLayout(config.LayoutFilePath, config.LayerStates);
                }
                else if (!string.IsNullOrEmpty(config.LayoutFilePath))
                {
                    MessageBox.Show($"Layout file not found:\n{config.LayoutFilePath}", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                }

                // Restore view
                if (config.ZoomLevel > 0)
                {
                    _zoomFactor = config.ZoomLevel;

                    double vw = MainGrid.ColumnDefinitions[0].ActualWidth;
                    double vh = ActualHeight - 80;
                    if (vw <= 0 || vh <= 0) { vw = 900; vh = 700; }

                    var m = new Matrix();
                    m.Translate(-config.ViewCenterX, config.ViewCenterY); // Y is flipped
                    m.Scale(_zoomFactor, _zoomFactor);
                    m.Translate(vw / 2, vh / 2);
                    CanvasTransform.Matrix = m;
                    UpdateZoomText();
                }

                // Restore probes
                _probePoints.Clear();
                _probeDomains.Clear();
                _domainColorMap.Clear();
                RemoveMarkersByTag("__PROBE__");
                foreach (var p in config.Probes)
                {
                    _probePoints.Add(new Point(p.X, p.Y));
                    _probeDomains.Add(p.Domain);
                    AddProbeMarker(p.X, p.Y, p.Domain);
                }
                UpdateProbeInfo();

                // Restore capacitors
                _capacitorAssignments.Clear();
                _capFileColorMap.Clear();
                RemoveMarkersByTag("__CAP__");
                foreach (var cc in config.Capacitors)
                {
                    var assignment = new CapacitorAssignment
                    {
                        FileName = cc.Name,
                        FilePath = cc.FilePath,
                        Location = new Point(cc.X, cc.Y),
                        Coordinates = $"({cc.X:F2}, {cc.Y:F2})",
                        Rlc = new RlcResult { R = cc.FitR, L = cc.FitL, C = cc.FitC, ResonanceFreq = cc.ResonanceFreq },
                        RlcInfo = $"R={RlcHelper.ToEngineeringNotation(cc.FitR, "Ω")}  " +
                                  $"L={RlcHelper.ToEngineeringNotation(cc.FitL, "H")}  " +
                                  $"C={RlcHelper.ToEngineeringNotation(cc.FitC, "F")}  " +
                                  $"fRes={RlcHelper.ToEngineeringNotation(cc.ResonanceFreq, "Hz")}"
                    };
                    _capacitorAssignments.Add(assignment);
                    AddCapacitorMarker(cc.X, cc.Y, cc.Name, cc.FilePath);

                    // Load the S-param file in the graph
                    if (!string.IsNullOrEmpty(cc.FilePath) && File.Exists(cc.FilePath))
                        CapacitorFileLoaded?.Invoke(this, cc.FilePath);
                }
                UpdateCapacitorInfo();

                StatusText.Text = $"Configuration loaded: {_probePoints.Count} probes, {_capacitorAssignments.Count} capacitors";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading config: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RemoveMarkersByTag(string tag)
        {
            var toRemove = DxfCanvas.Children.OfType<FrameworkElement>().Where(fe => fe.Tag is string t && t == tag).ToList();
            foreach (var r in toRemove) DxfCanvas.Children.Remove(r);
        }

        // ─────────────────────── KEYBOARD ──────────────────────────
        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);

            if (e.Key == Key.Escape)
            {
                _isProbeMode = false;
                _isCapacitorMode = false;
                SyncModeButtons();
                StatusText.Text = "Ready";
            }
            else if (e.Key == Key.Delete)
            {
                if (_isProbeMode && _probePoints.Count > 0)
                {
                    _probePoints.RemoveAt(_probePoints.Count - 1);
                    if (_probeDomains.Count > _probePoints.Count)
                        _probeDomains.RemoveAt(_probeDomains.Count - 1);
                    RemoveLastMarkerByTag("__PROBE__");
                    UpdateProbeInfo();
                }
                else if (_isCapacitorMode && _capacitorAssignments.Count > 0)
                {
                    _capacitorAssignments.RemoveAt(_capacitorAssignments.Count - 1);
                    RemoveLastMarkerByTag("__CAP__");
                    UpdateCapacitorInfo();
                }
            }
            else if (e.Key == Key.C && Keyboard.Modifiers == ModifierKeys.Shift)
            {
                // Shift+C: Clear all current mode items
                if (_isProbeMode)
                {
                    _probePoints.Clear();
                    _probeDomains.Clear();
                    RemoveMarkersByTag("__PROBE__");
                    UpdateProbeInfo();
                }
                else if (_isCapacitorMode)
                {
                    _capacitorAssignments.Clear();
                    RemoveMarkersByTag("__CAP__");
                    UpdateCapacitorInfo();
                }
            }
        }

        private void RemoveLastMarkerByTag(string tag)
        {
            for (int i = DxfCanvas.Children.Count - 1; i >= 0; i--)
            {
                if (DxfCanvas.Children[i] is FrameworkElement fe && fe.Tag is string t && t == tag)
                {
                    DxfCanvas.Children.RemoveAt(i);
                    return;
                }
            }
        }

        // ─────────────────────── PUBLIC ACCESS ─────────────────────
        public IReadOnlyList<Point> GetProbePoints() => _probePoints.AsReadOnly();
        public IReadOnlyList<CapacitorAssignment> GetCapacitorAssignments() => _capacitorAssignments.AsReadOnly();
        public BoardParameters GetBoardParameters() => _boardParams;
    }

    // ── Settings Editor Window ────────────────────────────────────
    public class SettingsEditorWindow : Window
    {
        private readonly TextBox _txtResistivity;
        private readonly TextBox _txtThickness;
        private readonly TextBox _txtProbeR;
        private readonly TextBox _txtProbeL;
        private readonly TextBox _txtProbeC;
        public BoardParameters Result { get; private set; }

        public SettingsEditorWindow(BoardParameters current)
        {
            Title = "Board & Probe Settings";
            Width = 380;
            SizeToContent = SizeToContent.Height;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x2D, 0x2D, 0x30));
            ResizeMode = ResizeMode.NoResize;
            Result = current;

            var sp = new StackPanel { Margin = new Thickness(16) };

            sp.Children.Add(MakeLabel("Plane Material Resistivity (Ω·m):"));
            _txtResistivity = MakeTextBox(current.PlaneResistivity.ToString("E4", CultureInfo.InvariantCulture));
            sp.Children.Add(_txtResistivity);

            sp.Children.Add(MakeLabel("Plane Thickness (m):"));
            _txtThickness = MakeTextBox(current.PlaneThickness.ToString("E4", CultureInfo.InvariantCulture));
            sp.Children.Add(_txtThickness);

            sp.Children.Add(MakeLabel("Probe Series Resistance (Ω):"));
            _txtProbeR = MakeTextBox(current.ProbeSeriesR.ToString("G6", CultureInfo.InvariantCulture));
            sp.Children.Add(_txtProbeR);

            sp.Children.Add(MakeLabel("Probe Series Inductance (H):"));
            _txtProbeL = MakeTextBox(current.ProbeSeriesL.ToString("E4", CultureInfo.InvariantCulture));
            sp.Children.Add(_txtProbeL);

            sp.Children.Add(MakeLabel("Probe Parallel Capacitance (F):"));
            _txtProbeC = MakeTextBox(current.ProbeParallelC.ToString("E4", CultureInfo.InvariantCulture));
            sp.Children.Add(_txtProbeC);

            var btnPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 16, 0, 0) };
            var ok = new Button { Content = "OK", Width = 80, Margin = new Thickness(0, 0, 8, 0) };
            ok.Click += Ok_Click;
            var cancel = new Button { Content = "Cancel", Width = 80 };
            cancel.Click += (s, e) => { DialogResult = false; Close(); };
            btnPanel.Children.Add(ok);
            btnPanel.Children.Add(cancel);
            sp.Children.Add(btnPanel);

            Content = sp;
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Result = new BoardParameters
                {
                    PlaneResistivity = double.Parse(_txtResistivity.Text, NumberStyles.Float, CultureInfo.InvariantCulture),
                    PlaneThickness = double.Parse(_txtThickness.Text, NumberStyles.Float, CultureInfo.InvariantCulture),
                    ProbeSeriesR = double.Parse(_txtProbeR.Text, NumberStyles.Float, CultureInfo.InvariantCulture),
                    ProbeSeriesL = double.Parse(_txtProbeL.Text, NumberStyles.Float, CultureInfo.InvariantCulture),
                    ProbeParallelC = double.Parse(_txtProbeC.Text, NumberStyles.Float, CultureInfo.InvariantCulture)
                };
                DialogResult = true;
                Close();
            }
            catch (FormatException)
            {
                MessageBox.Show("Please enter valid numeric values.", "Input Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private static TextBlock MakeLabel(string text) => new TextBlock { Text = text, Foreground = Brushes.White, Margin = new Thickness(0, 8, 0, 2) };
        private static TextBox MakeTextBox(string text) => new TextBox { Text = text, Margin = new Thickness(0, 0, 0, 4) };
    }
}
