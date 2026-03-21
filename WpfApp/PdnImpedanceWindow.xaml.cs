using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Windows;
using System.Windows.Controls;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Legends;
using OxyPlot.Series;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using WStyle = DocumentFormat.OpenXml.Wordprocessing.Style;

namespace WpfApp
{
    public partial class PdnImpedanceWindow : Window
    {
        private readonly List<Point> _probePoints;
        private readonly List<string> _probeDomains;
        private readonly List<CapacitorAssignment> _capacitorAssignments;
        private readonly BoardParameters _boardParams;

        // Cached S-parameter impedance data per cap file: freq[] -> Z[]
        private readonly Dictionary<string, List<(double Freq, Complex Z)>> _capZDataCache = new(StringComparer.OrdinalIgnoreCase);

        private readonly OxyColor[] _probeColors = new[]
        {
            OxyColors.DeepSkyBlue, OxyColors.OrangeRed, OxyColors.LimeGreen,
            OxyColors.Goldenrod, OxyColors.HotPink, OxyColors.MediumPurple,
            OxyColors.Teal, OxyColors.Crimson, OxyColors.CadetBlue, OxyColors.DarkCyan
        };

        public PdnImpedanceWindow(
            List<Point> probePoints,
            List<string> probeDomains,
            List<CapacitorAssignment> capacitorAssignments,
            BoardParameters boardParams)
        {
            InitializeComponent();
            _probePoints = probePoints;
            _probeDomains = probeDomains;
            _capacitorAssignments = capacitorAssignments;
            _boardParams = boardParams;

            PopulateDomains();
        }

        // ----------------------- INIT ------------------------------
        private void PopulateDomains()
        {
            var domains = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < _probeDomains.Count; i++)
                if (!string.IsNullOrWhiteSpace(_probeDomains[i]))
                    domains.Add(_probeDomains[i]);

            foreach (var d in domains)
                CmbDomain.Items.Add(d);

            if (CmbDomain.Items.Count > 0)
                CmbDomain.SelectedIndex = 0;

            // Show sheet resistance
            double rSheet = _boardParams.PlaneResistivity / _boardParams.PlaneThickness;
            TxtRsheet.Text = $"{RlcHelper.ToEngineeringNotation(rSheet, "\u03A9/\u25A1")}";
        }

        // ----------------------- EVENTS ----------------------------
        private void Domain_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (CmbDomain.SelectedItem != null)
                Calculate((string)CmbDomain.SelectedItem);
        }

        private void Recalculate_Click(object sender, RoutedEventArgs e)
        {
            if (CmbDomain.SelectedItem != null)
                Calculate((string)CmbDomain.SelectedItem);
        }

        // ----------------------- CALCULATION -----------------------

        private static bool TryParseDouble(string text, out double value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(text)) return false;
            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value) && value > 0;
        }

        private double GetViaPadRadius()
        {
            if (double.TryParse(TxtViaPadRadius.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double r) && r > 0)
                return r; // mm
            return 0.15; // default
        }

        /// <summary>
        /// Load Z(f) from the S-parameter file of a capacitor (cached).
        /// </summary>
        private List<(double Freq, Complex Z)> GetCapacitorZData(CapacitorAssignment cap)
        {
            if (_capZDataCache.TryGetValue(cap.FilePath, out var cached))
                return cached;

            var result = new List<(double, Complex)>();
            if (string.IsNullOrEmpty(cap.FilePath) || !File.Exists(cap.FilePath))
            {
                _capZDataCache[cap.FilePath] = result;
                return result;
            }

            try
            {
                var fileData = TouchstoneParser.Parse(cap.FilePath);

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
                if (paramIndex < 0) { _capZDataCache[cap.FilePath] = result; return result; }

                foreach (var p in fileData.Points)
                {
                    var z = RlcHelper.CalculateComplexImpedance(p.Parameters[paramIndex], fileData.ReferenceImpedance);
                    if (!double.IsInfinity(z.Real) && !double.IsInfinity(z.Imaginary) &&
                        !double.IsNaN(z.Real) && !double.IsNaN(z.Imaginary))
                        result.Add((p.FrequencyHz, z));
                }
            }
            catch { /* skip on error */ }

            _capZDataCache[cap.FilePath] = result;
            return result;
        }

        /// <summary>
        /// Spreading resistance on a copper plane: R = (R_sheet / 2*pi) * ln(d / r0)
        /// </summary>
        private double SpreadingResistance(double distanceMm, double viaPadRadiusMm)
        {
            double rSheet = _boardParams.PlaneResistivity / _boardParams.PlaneThickness;
            if (distanceMm <= viaPadRadiusMm) distanceMm = viaPadRadiusMm * 1.01; // clamp
            return (rSheet / (2.0 * Math.PI)) * Math.Log(distanceMm / viaPadRadiusMm);
        }

        /// <summary>
        /// Probe parasitic impedance model:
        ///   Z_probe_series = R_s + j*w*L_s   (in series with PDN)
        ///   Z_probe_parallel = 1/(j*w*C_p)   (in parallel, if C_p > 0)
        ///   Z_measured = (Z_pdn + Z_probe_series) || Z_probe_parallel
        /// </summary>
        private Complex ApplyProbeParasitics(Complex zPdn, double freqHz)
        {
            double w = 2.0 * Math.PI * freqHz;

            // Series parasitics
            Complex zSeries = new Complex(_boardParams.ProbeSeriesR, w * _boardParams.ProbeSeriesL);
            Complex zTotal = zPdn + zSeries;

            // Parallel parasitic capacitance
            if (_boardParams.ProbeParallelC > 0 && w > 0)
            {
                Complex zParC = new Complex(0, -1.0 / (w * _boardParams.ProbeParallelC));
                // Parallel combination: (zTotal * zParC) / (zTotal + zParC)
                Complex sum = zTotal + zParC;
                if (sum.Magnitude > 1e-30)
                    zTotal = (zTotal * zParC) / sum;
            }

            return zTotal;
        }

        /// <summary>
        /// Main calculation: for the given domain, compute Z(f) seen from each probe
        /// and also the combined Z(f) from all probes in parallel.
        /// </summary>
        private void Calculate(string domain)
        {
            double r0 = GetViaPadRadius(); // mm

            // Get probes of this domain
            var probeIndices = new List<int>();
            for (int i = 0; i < _probePoints.Count; i++)
            {
                if (i < _probeDomains.Count &&
                    string.Equals(_probeDomains[i], domain, StringComparison.OrdinalIgnoreCase))
                    probeIndices.Add(i);
            }

            // Get capacitors of this domain
            var domainCaps = _capacitorAssignments
                .Where(c => string.Equals(c.DomainName, domain, StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (probeIndices.Count == 0)
            {
                TxtInfo.Text = "No probes found for this domain.";
                PlotView.Model = null;
                return;
            }
            if (domainCaps.Count == 0)
            {
                TxtInfo.Text = "No capacitors assigned to this domain.";
                PlotView.Model = null;
                return;
            }

            // Load Z(f) data for each capacitor
            var capDataList = new List<(CapacitorAssignment Cap, List<(double Freq, Complex Z)> ZData)>();
            foreach (var cap in domainCaps)
            {
                var zData = GetCapacitorZData(cap);
                if (zData.Count > 0)
                    capDataList.Add((cap, zData));
            }

            if (capDataList.Count == 0)
            {
                TxtInfo.Text = "Could not load S-parameter data for any capacitor.";
                PlotView.Model = null;
                return;
            }

            // Use the frequency vector of the first cap (all should be the same or similar)
            var frequencies = capDataList[0].ZData.Select(d => d.Freq).ToArray();

            // Build results: one trace per probe + one combined
            var model = new PlotModel
            {
                Title = $"PDN Impedance \u2014 {domain}",
                TitleColor = OxyColors.White,
                PlotAreaBorderColor = OxyColors.Gray,
                Background = OxyColor.Parse("#1E1E1E"),
                TextColor = OxyColors.White,
                SubtitleColor = OxyColors.LightGray,
                Subtitle = $"{probeIndices.Count} probes, {domainCaps.Count} capacitors"
            };

            // Log-log axes — apply user-defined ranges if provided
            var freqAxis = new LogarithmicAxis
            {
                Position = AxisPosition.Bottom,
                Title = "Frequency (Hz)",
                TitleColor = OxyColors.White,
                TextColor = OxyColors.LightGray,
                MajorGridlineStyle = LineStyle.Solid,
                MajorGridlineColor = OxyColor.FromArgb(40, 255, 255, 255),
                MinorGridlineStyle = LineStyle.Dot,
                MinorGridlineColor = OxyColor.FromArgb(20, 255, 255, 255),
                AxislineColor = OxyColors.Gray,
                TicklineColor = OxyColors.Gray
            };
            if (TryParseDouble(TxtFreqMin.Text, out double fMin)) freqAxis.Minimum = fMin;
            if (TryParseDouble(TxtFreqMax.Text, out double fMax)) freqAxis.Maximum = fMax;
            model.Axes.Add(freqAxis);

            var zAxis = new LogarithmicAxis
            {
                Position = AxisPosition.Left,
                Title = "|Z| (Ω)",
                TitleColor = OxyColors.White,
                TextColor = OxyColors.LightGray,
                MajorGridlineStyle = LineStyle.Solid,
                MajorGridlineColor = OxyColor.FromArgb(40, 255, 255, 255),
                MinorGridlineStyle = LineStyle.Dot,
                MinorGridlineColor = OxyColor.FromArgb(20, 255, 255, 255),
                AxislineColor = OxyColors.Gray,
                TicklineColor = OxyColors.Gray
            };
            if (TryParseDouble(TxtZMin.Text, out double zMin)) zAxis.Minimum = zMin;
            if (TryParseDouble(TxtZMax.Text, out double zMax)) zAxis.Maximum = zMax;
            model.Axes.Add(zAxis);

            // For combined: accumulator across probes (parallel all probes)
            var combinedYinv = new Complex[frequencies.Length]; // sum of 1/Z for each freq

            var infoLines = new List<string>();

            foreach (int pIdx in probeIndices)
            {
                var probePos = _probePoints[pIdx];
                var series = new LineSeries
                {
                    Title = $"Probe P{pIdx + 1} ({probePos.X:F0}, {probePos.Y:F0})",
                    Color = _probeColors[pIdx % _probeColors.Length],
                    StrokeThickness = 2
                };

                for (int fi = 0; fi < frequencies.Length; fi++)
                {
                    double freq = frequencies[fi];

                    // Sum admittances of all capacitor branches
                    Complex yTotal = Complex.Zero;

                    foreach (var (cap, zData) in capDataList)
                    {
                        // Find closest frequency in this cap's data
                        Complex zCap = InterpolateZ(zData, freq);

                        // Spreading resistance from cap to this probe
                        double dist = Math.Sqrt(
                            Math.Pow(cap.Location.X - probePos.X, 2) +
                            Math.Pow(cap.Location.Y - probePos.Y, 2));
                        // dist is in layout units (mils or mm depending on file) - assume mm for now
                        double rSpread = SpreadingResistance(dist, r0);

                        Complex zBranch = zCap + new Complex(rSpread, 0);
                        if (zBranch.Magnitude > 1e-30)
                            yTotal += 1.0 / zBranch;
                    }

                    Complex zPdn = (yTotal.Magnitude > 1e-30) ? 1.0 / yTotal : new Complex(1e12, 0);

                    // Apply probe parasitics
                    Complex zMeasured = ApplyProbeParasitics(zPdn, freq);

                    double mag = zMeasured.Magnitude;
                    if (mag > 0 && !double.IsNaN(mag) && !double.IsInfinity(mag))
                    {
                        series.Points.Add(new DataPoint(freq, mag));
                        combinedYinv[fi] += 1.0 / zMeasured;
                    }
                }

                model.Series.Add(series);

                // Find minimum impedance for info
                if (series.Points.Count > 0)
                {
                    var minPt = series.Points.OrderBy(p => p.Y).First();
                    infoLines.Add($"P{pIdx + 1}: min |Z| = {RlcHelper.ToEngineeringNotation(minPt.Y, "\u03A9")} @ {RlcHelper.ToEngineeringNotation(minPt.X, "Hz")}");
                }
            }

            // Combined: parallel of all probes
            if (probeIndices.Count > 1)
            {
                var combinedSeries = new LineSeries
                {
                    Title = "Combined (all probes)",
                    Color = OxyColors.White,
                    StrokeThickness = 3,
                    LineStyle = LineStyle.Dash
                };

                for (int fi = 0; fi < frequencies.Length; fi++)
                {
                    if (combinedYinv[fi].Magnitude > 1e-30)
                    {
                        Complex zCombined = 1.0 / combinedYinv[fi];
                        double mag = zCombined.Magnitude;
                        if (mag > 0 && !double.IsNaN(mag) && !double.IsInfinity(mag))
                            combinedSeries.Points.Add(new DataPoint(frequencies[fi], mag));
                    }
                }

                model.Series.Add(combinedSeries);

                if (combinedSeries.Points.Count > 0)
                {
                    var minPt = combinedSeries.Points.OrderBy(p => p.Y).First();
                    infoLines.Add($"Combined: min |Z| = {RlcHelper.ToEngineeringNotation(minPt.Y, "\u03A9")} @ {RlcHelper.ToEngineeringNotation(minPt.X, "Hz")}");
                }
            }

            model.Legends.Add(new Legend
            {
                LegendPosition = LegendPosition.RightTop,
                LegendPlacement = LegendPlacement.Inside,
                LegendBackground = OxyColor.FromAColor(180, OxyColors.Black),
                LegendTextColor = OxyColors.White,
                LegendBorder = OxyColors.Gray
            });

            PlotView.Model = model;

            TxtInfo.Text = string.Join("  |  ", infoLines);
        }

        /// <summary>
        /// Linear interpolation of Z at a given frequency from discrete data.
        /// </summary>
        private static Complex InterpolateZ(List<(double Freq, Complex Z)> data, double freq)
        {
            if (data.Count == 0) return Complex.Zero;
            if (freq <= data[0].Freq) return data[0].Z;
            if (freq >= data[^1].Freq) return data[^1].Z;

            // Binary search for bracket
            int lo = 0, hi = data.Count - 1;
            while (hi - lo > 1)
            {
                int mid = (lo + hi) / 2;
                if (data[mid].Freq <= freq) lo = mid;
                else hi = mid;
            }

            double f0 = data[lo].Freq, f1 = data[hi].Freq;
            if (Math.Abs(f1 - f0) < 1e-30) return data[lo].Z;

            double t = (freq - f0) / (f1 - f0);
            return data[lo].Z * (1 - t) + data[hi].Z * t;
        }

        // ======================= REPORT EXPORT ======================

        private void ExportReport_Click(object sender, RoutedEventArgs e)
        {
            if (CmbDomain.SelectedItem == null) { MessageBox.Show("Select a domain first."); return; }
            string domain = (string)CmbDomain.SelectedItem;

            var dlg = new SaveFileDialog
            {
                Filter = "Word Document (*.docx)|*.docx",
                FileName = $"PDN_Report_{domain}_{DateTime.Now:yyyyMMdd_HHmmss}.docx"
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                GenerateReport(dlg.FileName, domain);
                MessageBox.Show($"Report saved:\n{dlg.FileName}", "Report", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export failed:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void GenerateReport(string filePath, string domain)
        {
            double r0 = GetViaPadRadius();
            double rSheet = _boardParams.PlaneResistivity / _boardParams.PlaneThickness;

            var probeIndices = new List<int>();
            for (int i = 0; i < _probePoints.Count; i++)
                if (i < _probeDomains.Count && string.Equals(_probeDomains[i], domain, StringComparison.OrdinalIgnoreCase))
                    probeIndices.Add(i);

            var domainCaps = _capacitorAssignments
                .Where(c => string.Equals(c.DomainName, domain, StringComparison.OrdinalIgnoreCase)).ToList();

            using var doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document);
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;
            AddStyles(mainPart);

            // ── Title ──
            body.Append(WPara("PDN Impedance Analysis Report", "Title"));
            body.Append(WText($"Domain: {domain}"));
            body.Append(WText($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}"));
            body.Append(WText(""));

            // ── 1. Methodology ──
            body.Append(WPara("1. Methodology", "Heading1"));
            body.Append(WText(
                "This analysis computes the Power Distribution Network (PDN) impedance as seen from each probe location " +
                "by modelling each decoupling capacitor as a branch consisting of its measured S-parameter impedance " +
                "in series with the spreading resistance of the copper plane."));
            body.Append(WText(""));

            body.Append(WPara("1.1 Capacitor Impedance from S-Parameters", "Heading2"));
            body.Append(WText(
                "Each capacitor's complex impedance Z_cap(f) is extracted from its Touchstone S-parameter file. " +
                "For a series-thru (S21) measurement the conversion is:"));
            body.Append(WText(""));
            body.Append(WEquation("Z_cap = 2 · Z₀ · (1 − S₂₁) / S₂₁"));
            body.Append(WText(""));
            body.Append(WText("where Z₀ is the reference impedance (typically 50 Ω)."));
            body.Append(WText(""));

            body.Append(WPara("1.2 Spreading Resistance", "Heading2"));
            body.Append(WText(
                "The spreading resistance models the resistive loss in the copper plane between each capacitor " +
                "via/pad and the probe measurement point. Using the analytical solution for a circular current " +
                "spreading on a thin conductive sheet:"));
            body.Append(WText(""));
            body.Append(WEquation("R_spread = (R□ / 2π) · ln(d / r₀)"));
            body.Append(WText(""));
            body.Append(WText("where:"));
            body.Append(WBullet("R□ = ρ / t  is the sheet resistance of the copper plane"));
            body.Append(WBullet("d  is the distance from capacitor to probe (mm)"));
            body.Append(WBullet("r₀ is the effective via/pad radius (mm)"));
            body.Append(WText(""));

            body.Append(WPara("1.3 PDN Impedance Computation", "Heading2"));
            body.Append(WText(
                "For each probe position, all capacitor branches are combined in parallel using superposition " +
                "of admittances. Each branch impedance is:"));
            body.Append(WText(""));
            body.Append(WEquation("Z_branch,i = Z_cap,i(f) + R_spread,i"));
            body.Append(WText(""));
            body.Append(WText("The total PDN impedance seen from the probe is:"));
            body.Append(WText(""));
            body.Append(WEquation("Z_PDN(f) = 1 / Σᵢ (1 / Z_branch,i)"));
            body.Append(WText(""));

            body.Append(WPara("1.4 Probe Parasitic De-embedding", "Heading2"));
            body.Append(WText(
                "The probe introduces series resistance R_s and inductance L_s, plus a parallel parasitic " +
                "capacitance C_p. The measured impedance becomes:"));
            body.Append(WText(""));
            body.Append(WEquation("Z_series = R_s + jωL_s"));
            body.Append(WEquation("Z_measured = (Z_PDN + Z_series) ∥ (1 / jωC_p)"));
            body.Append(WText(""));
            body.Append(WText(
                "When multiple probes are present in the same domain, their combined impedance is computed " +
                "as a parallel combination of all individual probe impedances."));
            body.Append(WText(""));

            // ── 2. Board Parameters ──
            body.Append(WPara("2. Board Parameters", "Heading1"));
            body.Append(WText($"Plane Resistivity (ρ): {_boardParams.PlaneResistivity:E2} Ω·m"));
            body.Append(WText($"Plane Thickness (t): {_boardParams.PlaneThickness * 1e6:F1} µm"));
            body.Append(WText($"Sheet Resistance (R□): {RlcHelper.ToEngineeringNotation(rSheet, "Ω/□")}"));
            body.Append(WText($"Via/Pad Radius (r₀): {r0} mm"));
            body.Append(WText($"Probe Series R: {RlcHelper.ToEngineeringNotation(_boardParams.ProbeSeriesR, "Ω")}"));
            body.Append(WText($"Probe Series L: {RlcHelper.ToEngineeringNotation(_boardParams.ProbeSeriesL, "H")}"));
            body.Append(WText($"Probe Parallel C: {RlcHelper.ToEngineeringNotation(_boardParams.ProbeParallelC, "F")}"));
            body.Append(WText(""));

            // ── 3. Probe Locations ──
            body.Append(WPara("3. Probe Locations", "Heading1"));
            var probeTable = WTable(new[] { "Probe", "X (mm)", "Y (mm)", "Domain" });
            foreach (int idx in probeIndices)
            {
                var p = _probePoints[idx];
                WRow(probeTable, new[] { $"P{idx + 1}", $"{p.X:F2}", $"{p.Y:F2}", _probeDomains[idx] });
            }
            body.Append(probeTable);
            body.Append(WText(""));

            // ── 4. Capacitor Details ──
            body.Append(WPara("4. Capacitor Details", "Heading1"));
            var capHeaders = new List<string> { "#", "Name", "X", "Y", "R", "L", "C", "f₀" };
            foreach (int pi in probeIndices) capHeaders.Add($"d(P{pi + 1})");
            foreach (int pi in probeIndices) capHeaders.Add($"R_sp(P{pi + 1})");

            var capTable = WTable(capHeaders.ToArray());
            for (int ci = 0; ci < domainCaps.Count; ci++)
            {
                var cap = domainCaps[ci];
                var row = new List<string>
                {
                    (ci + 1).ToString(), cap.FileName,
                    $"{cap.Location.X:F2}", $"{cap.Location.Y:F2}",
                    cap.Rlc != null ? RlcHelper.ToEngineeringNotation(cap.Rlc.R, "Ω") : "—",
                    cap.Rlc != null ? RlcHelper.ToEngineeringNotation(cap.Rlc.L, "H") : "—",
                    cap.Rlc != null ? RlcHelper.ToEngineeringNotation(cap.Rlc.C, "F") : "—",
                    cap.Rlc != null ? RlcHelper.ToEngineeringNotation(cap.Rlc.ResonanceFreq, "Hz") : "—"
                };
                foreach (int pi in probeIndices)
                {
                    double dist = Math.Sqrt(Math.Pow(cap.Location.X - _probePoints[pi].X, 2) + Math.Pow(cap.Location.Y - _probePoints[pi].Y, 2));
                    row.Add($"{dist:F3}");
                }
                foreach (int pi in probeIndices)
                {
                    double dist = Math.Sqrt(Math.Pow(cap.Location.X - _probePoints[pi].X, 2) + Math.Pow(cap.Location.Y - _probePoints[pi].Y, 2));
                    row.Add(RlcHelper.ToEngineeringNotation(SpreadingResistance(dist, r0), "Ω"));
                }
                WRow(capTable, row.ToArray());
            }
            body.Append(capTable);
            body.Append(WText(""));

            // ── 5. Impedance Chart ──
            body.Append(WPara("5. Impedance Chart", "Heading1"));
            if (PlotView.Model != null)
            {
                try
                {
                    string tmp = Path.Combine(Path.GetTempPath(), $"pdn_{Guid.NewGuid():N}.png");
                    var png = new OxyPlot.Wpf.PngExporter { Width = 1200, Height = 600 };
                    using (var s = File.Create(tmp)) png.Export(PlotView.Model, s);
                    WImage(mainPart, body, tmp, 1200, 600);
                    try { File.Delete(tmp); } catch { }
                }
                catch { body.Append(WText("[Chart could not be exported]")); }
            }
            body.Append(WText(""));

            // ── 6. Z(f) Data Table ──
            body.Append(WPara("6. Impedance vs Frequency", "Heading1"));
            var capDataList = new List<(CapacitorAssignment Cap, List<(double Freq, Complex Z)> ZData)>();
            foreach (var cap in domainCaps)
            {
                var zd = GetCapacitorZData(cap);
                if (zd.Count > 0) capDataList.Add((cap, zd));
            }
            if (capDataList.Count > 0)
            {
                var freqs = capDataList[0].ZData.Select(d => d.Freq).ToArray();
                var sample = LogSample(freqs, 30);

                var zh = new List<string> { "Frequency" };
                foreach (int pi in probeIndices) zh.Add($"|Z| P{pi + 1}");
                if (probeIndices.Count > 1) zh.Add("|Z| Combined");
                var zt = WTable(zh.ToArray());

                foreach (double f in sample)
                {
                    var rv = new List<string> { RlcHelper.ToEngineeringNotation(f, "Hz") };
                    Complex combY = Complex.Zero;
                    foreach (int pi in probeIndices)
                    {
                        var pp = _probePoints[pi];
                        Complex yT = Complex.Zero;
                        foreach (var (cap, zData) in capDataList)
                        {
                            Complex zC = InterpolateZ(zData, f);
                            double d = Math.Sqrt(Math.Pow(cap.Location.X - pp.X, 2) + Math.Pow(cap.Location.Y - pp.Y, 2));
                            Complex zB = zC + new Complex(SpreadingResistance(d, r0), 0);
                            if (zB.Magnitude > 1e-30) yT += 1.0 / zB;
                        }
                        Complex zP = (yT.Magnitude > 1e-30) ? 1.0 / yT : new Complex(1e12, 0);
                        Complex zM = ApplyProbeParasitics(zP, f);
                        rv.Add(RlcHelper.ToEngineeringNotation(zM.Magnitude, "Ω"));
                        if (zM.Magnitude > 1e-30) combY += 1.0 / zM;
                    }
                    if (probeIndices.Count > 1)
                    {
                        Complex zCmb = (combY.Magnitude > 1e-30) ? 1.0 / combY : new Complex(1e12, 0);
                        rv.Add(RlcHelper.ToEngineeringNotation(zCmb.Magnitude, "Ω"));
                    }
                    WRow(zt, rv.ToArray());
                }
                body.Append(zt);
                body.Append(WText(""));

                // ── 7. Min Impedance Summary ──
                body.Append(WPara("7. Minimum Impedance Summary", "Heading1"));
                foreach (int pi in probeIndices)
                {
                    var pp = _probePoints[pi];
                    double minZ = double.MaxValue, minF = 0;
                    foreach (double f in freqs)
                    {
                        Complex yT = Complex.Zero;
                        foreach (var (cap, zData) in capDataList)
                        {
                            Complex zC = InterpolateZ(zData, f);
                            double d = Math.Sqrt(Math.Pow(cap.Location.X - pp.X, 2) + Math.Pow(cap.Location.Y - pp.Y, 2));
                            Complex zB = zC + new Complex(SpreadingResistance(d, r0), 0);
                            if (zB.Magnitude > 1e-30) yT += 1.0 / zB;
                        }
                        Complex zP = (yT.Magnitude > 1e-30) ? 1.0 / yT : new Complex(1e12, 0);
                        Complex zM = ApplyProbeParasitics(zP, f);
                        double m = zM.Magnitude;
                        if (m < minZ && m > 0 && !double.IsNaN(m) && !double.IsInfinity(m)) { minZ = m; minF = f; }
                    }
                    body.Append(WText($"Probe P{pi + 1}: min |Z| = {RlcHelper.ToEngineeringNotation(minZ, "Ω")} @ {RlcHelper.ToEngineeringNotation(minF, "Hz")}"));
                }
            }

            body.Append(WText(""));
            body.Append(WText("— End of Report —"));
        }

        // --------------- Word helper methods -----------------------

        private static double[] LogSample(double[] all, int n)
        {
            if (all.Length <= n) return all;
            double lo = Math.Log10(all[0]), hi = Math.Log10(all[^1]);
            var r = new List<double>();
            for (int i = 0; i < n; i++)
            {
                double t = Math.Pow(10, lo + (hi - lo) * i / (n - 1));
                double best = all[0]; double bd = Math.Abs(t - best);
                foreach (var f in all) { double d = Math.Abs(f - t); if (d < bd) { best = f; bd = d; } }
                if (r.Count == 0 || Math.Abs(r[^1] - best) > 1e-6) r.Add(best);
            }
            return r.ToArray();
        }

        private static void AddStyles(MainDocumentPart mp)
        {
            var sp = mp.AddNewPart<StyleDefinitionsPart>();
            var styles = new Styles();
            styles.Append(MkStyle("Title", "Title", true, "44", "1F4E79"));
            styles.Append(MkStyle("Heading1", "Heading 1", true, "32", "2E74B5"));
            styles.Append(MkStyle("Heading2", "Heading 2", true, "26", "2E74B5"));
            sp.Styles = styles;
        }

        private static WStyle MkStyle(string id, string name, bool bold, string sz, string color)
        {
            var s = new WStyle { Type = StyleValues.Paragraph, StyleId = id };
            s.Append(new StyleName { Val = name });
            var rp = new StyleRunProperties();
            if (bold) rp.Append(new Bold());
            rp.Append(new FontSize { Val = sz });
            rp.Append(new Color { Val = color });
            s.Append(rp);
            return s;
        }

        private static Paragraph WPara(string text, string styleId)
        {
            var p = new Paragraph();
            p.Append(new ParagraphProperties(new ParagraphStyleId { Val = styleId }));
            p.Append(new Run(new Text(text)));
            return p;
        }

        private static Paragraph WText(string text)
        {
            return new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        }

        private static Paragraph WEquation(string eq)
        {
            var run = new Run(
                new RunProperties(new Bold(), new Italic(), new Color { Val = "2E74B5" }),
                new Text("    " + eq) { Space = SpaceProcessingModeValues.Preserve });
            return new Paragraph(
                new ParagraphProperties(new SpacingBetweenLines { Before = "60", After = "60" }),
                run);
        }

        private static Paragraph WBullet(string text)
        {
            return new Paragraph(
                new ParagraphProperties(new Indentation { Left = "720" }),
                new Run(new Text("• " + text) { Space = SpaceProcessingModeValues.Preserve }));
        }

        private static Table WTable(string[] headers)
        {
            var tbl = new Table();
            tbl.Append(new TableProperties(
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4, Color = "666666" },
                    new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "666666" },
                    new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "666666" },
                    new RightBorder { Val = BorderValues.Single, Size = 4, Color = "666666" },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "AAAAAA" },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "AAAAAA" }),
                new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }));
            var hr = new TableRow();
            foreach (var h in headers)
                hr.Append(new TableCell(
                    new TableCellProperties(new Shading { Val = ShadingPatternValues.Clear, Fill = "2E74B5" }),
                    new Paragraph(new Run(
                        new RunProperties(new Bold(), new Color { Val = "FFFFFF" }, new FontSize { Val = "18" }),
                        new Text(h) { Space = SpaceProcessingModeValues.Preserve }))));
            tbl.Append(hr);
            return tbl;
        }

        private static void WRow(Table tbl, string[] vals)
        {
            var r = new TableRow();
            foreach (var v in vals)
                r.Append(new TableCell(new Paragraph(new Run(
                    new RunProperties(new FontSize { Val = "16" }),
                    new Text(v ?? "—") { Space = SpaceProcessingModeValues.Preserve }))));
            tbl.Append(r);
        }

        private static void WImage(MainDocumentPart mp, Body body, string path, int w, int h)
        {
            var ip = mp.AddImagePart(ImagePartType.Png);
            using (var s = File.OpenRead(path)) ip.FeedData(s);
            string rid = mp.GetIdOfPart(ip);
            long cx = w * 9525L, cy = h * 9525L;
            const long mw = 5486400L;
            if (cx > mw) { double sc = (double)mw / cx; cx = mw; cy = (long)(cy * sc); }
            body.Append(new Paragraph(new Run(new Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = cx, Cy = cy },
                    new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                    new DW.DocProperties { Id = 1U, Name = "Chart" },
                    new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = 0U, Name = "chart.png" },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(new A.Blip { Embed = rid }, new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(new A.Offset { X = 0, Y = 0 }, new A.Extents { Cx = cx, Cy = cy }),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                ) { DistanceFromTop = 0, DistanceFromBottom = 0, DistanceFromLeft = 0, DistanceFromRight = 0 }))));
        }
    }
}

