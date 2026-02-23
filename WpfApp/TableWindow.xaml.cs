using System.Windows;
using System.Windows.Controls;

namespace WpfApp
{
    public partial class TableWindow : Window
    {
        public TableWindow(string title, Control content)
        {
            InitializeComponent();
            Title = title;
            ContentHost.Content = content;
        }
    }
}
