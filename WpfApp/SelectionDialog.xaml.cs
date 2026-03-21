using System.Collections.Generic;
using System.Windows;

namespace WpfApp
{
    public partial class SelectionDialog : Window
    {
        public string SelectedItem => cmbItems.SelectedItem?.ToString() ?? string.Empty;

        public SelectionDialog(string message, IEnumerable<string> items)
        {
            InitializeComponent();
            lblMessage.Text = message;

            foreach (var item in items)
                cmbItems.Items.Add(item);

            if (cmbItems.Items.Count > 0)
                cmbItems.SelectedIndex = 0;

            Loaded += (s, e) => cmbItems.Focus();
        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(SelectedItem))
            {
                MessageBox.Show("Please select an item.", "Notification", MessageBoxButton.OK, MessageBoxImage.Information);
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
