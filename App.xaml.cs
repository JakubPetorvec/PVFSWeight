using System;
using System.IO;
using System.Windows;
using System.Windows.Threading;

namespace VUKVWeightApp
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            ShutdownMode = ShutdownMode.OnMainWindowClose;

            DispatcherUnhandledException += OnUiException;
            AppDomain.CurrentDomain.UnhandledException += OnDomainException;

            base.OnStartup(e);
        }

        private void OnUiException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            ShowCrash(e.Exception);
            e.Handled = true;
        }

        private void OnDomainException(object? sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception ex) ShowCrash(ex);
            else ShowCrash(new Exception("Unknown fatal error"));
        }

        private static void ShowCrash(Exception ex)
        {
            try
            {
                var msg = ex.ToString();
                File.WriteAllText("crashlog.txt", msg);

                MessageBox.Show(
                    msg,
                    "VUKVWeightApp - chyba při startu",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
            }
            catch
            {
                // nic
            }
        }
    }
}
