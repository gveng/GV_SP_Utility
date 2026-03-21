using System.Collections.Generic;
using System.Windows;

namespace WpfApp
{
    public partial class DomainSelectorDialog : Window
    {
        public string SelectedDomain => cmbDomains.Text?.Trim() ?? string.Empty;

        public DomainSelectorDialog(IEnumerable<string> existingDomains, string defaultSelection = "", bool allowNewDomain = true)
        {
            InitializeComponent();
            cmbDomains.IsEditable = allowNewDomain;
            foreach (var d in existingDomains)
                cmbDomains.Items.Add(d);

            if (!string.IsNullOrWhiteSpace(defaultSelection))
                cmbDomains.Text = defaultSelection;
            else if (cmbDomains.Items.Count > 0)
                cmbDomains.SelectedIndex = 0;

            Loaded += (s, e) => { cmbDomains.Focus(); };
        }

        public DomainSelectorDialog(string message, IEnumerable<string> existingDomains, string defaultSelection = "", bool allowNewDomain = true)
            : this(existingDomains, defaultSelection, allowNewDomain)
        {
            lblMessage.Text = message;
        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(cmbDomains.Text))
            {
                MessageBox.Show("Please enter or select a power domain name.", "Domain Required",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            DialogResult = true;
            Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
