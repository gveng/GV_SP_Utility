using System.Windows;

namespace WpfApp
{
    public partial class PromptDialog : Window
    {
        public string ResponseText => txtInput.Text;

        public PromptDialog(string message, string defaultValue = "")
            : this("Input", message, defaultValue) { }

        public PromptDialog(string title, string message, string defaultValue)
        {
            InitializeComponent();
            Title = title;
            lblMessage.Text = message;
            txtInput.Text = defaultValue;
            Loaded += (s, e) => { txtInput.Focus(); txtInput.SelectAll(); };
        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
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
