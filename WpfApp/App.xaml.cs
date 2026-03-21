using System.Windows;

namespace WpfApp
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            DispatcherUnhandledException += App_DispatcherUnhandledException;
        }

        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            var ex = e.Exception;
            string msg = $"Unhandled exception: {ex.GetType().Name}\n{ex.Message}";
            if (ex.InnerException != null)
                msg += $"\n\nInner: {ex.InnerException.GetType().Name}\n{ex.InnerException.Message}";
            msg += $"\n\nStack trace:\n{ex.StackTrace}";
            MessageBox.Show(msg, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            e.Handled = true;
        }
    }
}
