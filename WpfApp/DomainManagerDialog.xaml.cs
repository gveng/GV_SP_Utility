using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace WpfApp
{
    public partial class DomainManagerDialog : Window
    {
        /// <summary>
        /// After the dialog is closed, contains the domain names the user chose to delete.
        /// </summary>
        public List<string> DomainsToDelete { get; } = new();

        private readonly List<(CheckBox Check, string Domain)> _checkBoxes = new();

        public DomainManagerDialog(
            IEnumerable<(string Domain, int ProbeCount, int CapCount, Brush Color)> domainInfos)
        {
            InitializeComponent();

            foreach (var info in domainInfos)
            {
                var sp = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(4, 3, 4, 3) };

                var cb = new CheckBox
                {
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 8, 0)
                };

                // Colored square matching the domain color
                var swatch = new Border
                {
                    Width = 14,
                    Height = 14,
                    Background = info.Color,
                    CornerRadius = new CornerRadius(2),
                    Margin = new Thickness(0, 0, 8, 0),
                    VerticalAlignment = VerticalAlignment.Center
                };

                var label = new TextBlock
                {
                    Text = $"{info.Domain}   ({info.ProbeCount} probes, {info.CapCount} caps)",
                    Foreground = Brushes.White,
                    FontSize = 13,
                    VerticalAlignment = VerticalAlignment.Center
                };

                sp.Children.Add(cb);
                sp.Children.Add(swatch);
                sp.Children.Add(label);
                DomainPanel.Children.Add(sp);

                _checkBoxes.Add((cb, info.Domain));
            }

            if (_checkBoxes.Count == 0)
            {
                DomainPanel.Children.Add(new TextBlock
                {
                    Text = "No power domains defined.",
                    Foreground = Brushes.Gray,
                    FontStyle = FontStyles.Italic,
                    Margin = new Thickness(8)
                });
            }
        }

        private void DeleteSelected_Click(object sender, RoutedEventArgs e)
        {
            var selected = _checkBoxes.Where(c => c.Check.IsChecked == true).Select(c => c.Domain).ToList();
            if (selected.Count == 0)
            {
                MessageBox.Show("No domains selected.", "Delete Domains",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var msg = $"Delete {selected.Count} domain(s) and all their probes/capacitors?\n\n" +
                      string.Join("\n", selected.Select(d => $"  \u2022 {d}"));
            if (MessageBox.Show(msg, "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
                return;

            DomainsToDelete.AddRange(selected);
            DialogResult = true;
            Close();
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
