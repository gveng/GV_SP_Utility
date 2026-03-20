using System;
using System.Windows;
using System.Windows.Controls;

namespace WpfApp
{
    public partial class RlcTuningWindow : Window
    {
        private bool _isUpdating;
        public Action<double, double, double>? OnParametersChanged;

        public RlcTuningWindow(double r, double l, double c)
        {
            InitializeComponent();
            
            _isUpdating = true;

            // Resistance
            RSlider.Minimum = 0;
            RSlider.Maximum = r > 0 ? r * 3 : 1;
            RSlider.Value = r;
            RValueText.Text = FormatValue(r, "Ω");

            // Inductance
            LSlider.Minimum = 0;
            LSlider.Maximum = l > 0 ? l * 3 : 1e-9;
            LSlider.Value = l;
            LValueText.Text = FormatValue(l, "H");

            // Capacitance
            CSlider.Minimum = 0;
            CSlider.Maximum = c > 0 ? c * 3 : 1e-12;
            CSlider.Value = c;
            CValueText.Text = FormatValue(c, "F");

            _isUpdating = false;

             RSlider.ValueChanged += Slider_ValueChanged;
             LSlider.ValueChanged += Slider_ValueChanged;
             CSlider.ValueChanged += Slider_ValueChanged;
        }

        private void Slider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_isUpdating) return;

            double r = RSlider.Value;
            double l = LSlider.Value;
            double c = CSlider.Value;

            RValueText.Text = FormatValue(r, "Ω");
            LValueText.Text = FormatValue(l, "H");
            CValueText.Text = FormatValue(c, "F");

            OnParametersChanged?.Invoke(r, l, c);
        }

        private static string FormatValue(double value, string unit)
        {
             // Simple formatting, could reuse Engineering Notation if made public/shared
             if (value == 0) return $"0 {unit}";
             double mag = Math.Abs(value);
             int exponent = (int)Math.Floor(Math.Log10(mag));
             int engExponent = (exponent >= 0) ? (exponent / 3) * 3 : ((exponent - 2) / 3) * 3;
             
             // Ensure we don't go below femto or above Tera for simplicity if needed, 
             // but standard math works fine.
             
             double scaledValue = value * Math.Pow(10, -engExponent);
             
             string prefix = engExponent switch
             {
                 -15 => "f",
                 -12 => "p",
                 -9 => "n",
                 -6 => "µ",
                 -3 => "m",
                 0 => "",
                 3 => "k",
                 6 => "M",
                 9 => "G",
                 _ => $"e{engExponent} "
             };
             
             return $"{scaledValue:F3} {prefix}{unit}";
        }
    }
}