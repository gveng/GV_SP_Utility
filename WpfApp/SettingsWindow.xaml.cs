using System;
using System.Globalization;
using System.Windows;

namespace WpfApp
{
    public partial class SettingsWindow : Window
    {
        public PlotSettings Settings { get; }

        public SettingsWindow(PlotSettings settings)
        {
            InitializeComponent();
            Settings = settings;
            LoadSettings();
        }

        private void LoadSettings()
        {
            XLinear.IsChecked = !Settings.IsXLog;
            XLog.IsChecked = Settings.IsXLog;
            YLinear.IsChecked = !Settings.IsYLog;
            YLog.IsChecked = Settings.IsYLog;

            XAuto.IsChecked = Settings.IsXAuto;
            XManual.IsChecked = !Settings.IsXAuto;
            YAuto.IsChecked = Settings.IsYAuto;
            YManual.IsChecked = !Settings.IsYAuto;

            XMinBox.IsEnabled = XMaxBox.IsEnabled = !Settings.IsXAuto;
            YMinBox.IsEnabled = YMaxBox.IsEnabled = !Settings.IsYAuto;

            XMinBox.Text = Settings.XMin?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
            XMaxBox.Text = Settings.XMax?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
            YMinBox.Text = Settings.YMin?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
            YMaxBox.Text = Settings.YMax?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            Settings.IsXLog = XLog.IsChecked == true;
            Settings.IsYLog = YLog.IsChecked == true;
            Settings.IsXAuto = XAuto.IsChecked == true;
            Settings.IsYAuto = YAuto.IsChecked == true;

            if (!Settings.IsXAuto)
            {
                if (!TryParseBox(XMinBox, out var xmin) || !TryParseBox(XMaxBox, out var xmax))
                {
                    MessageBox.Show(this, "Inserisci valori numerici validi per Min/Max X.", "Impostazioni", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                Settings.XMin = xmin;
                Settings.XMax = xmax;
            }
            else
            {
                Settings.XMin = Settings.XMax = null;
            }

            if (!Settings.IsYAuto)
            {
                if (!TryParseBox(YMinBox, out var ymin) || !TryParseBox(YMaxBox, out var ymax))
                {
                    MessageBox.Show(this, "Inserisci valori numerici validi per Min/Max Y.", "Impostazioni", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                Settings.YMin = ymin;
                Settings.YMax = ymax;
            }
            else
            {
                Settings.YMin = Settings.YMax = null;
            }

            if (!Settings.Validate(out var error))
            {
                MessageBox.Show(this, error, "Impostazioni", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DialogResult = true;
            Close();
        }

        private static bool TryParseBox(System.Windows.Controls.TextBox box, out double value)
        {
            return double.TryParse(box.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void XAuto_Checked(object sender, RoutedEventArgs e)
        {
            XMinBox.IsEnabled = XMaxBox.IsEnabled = false;
        }

        private void XManual_Checked(object sender, RoutedEventArgs e)
        {
            XMinBox.IsEnabled = XMaxBox.IsEnabled = true;
        }

        private void YAuto_Checked(object sender, RoutedEventArgs e)
        {
            YMinBox.IsEnabled = YMaxBox.IsEnabled = false;
        }

        private void YManual_Checked(object sender, RoutedEventArgs e)
        {
            YMinBox.IsEnabled = YMaxBox.IsEnabled = true;
        }
    }
}
