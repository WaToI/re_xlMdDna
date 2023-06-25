using ExcelDna.Integration;
using System;
using System.Threading.Tasks;
using System.Windows;

namespace xlMdDna {

    public class App {
        //private static Version AppVer = Assembly.GetEntryAssembly().GetName().Version;
        private static Version AppVer = new Version(2023, 6, 22, 0);

        public App() {
            setExceptionHandler();
        }

        void setExceptionHandler() {
            AppDomain.CurrentDomain.FirstChanceException += CurrentDomain_FirstChanceException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;
            //DispatcherUnhandledException += App_DispatcherUnhandledException;
        }

        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e) {
            e.Exception.ErrDump($"⊂( ･∀･) 彡 ｶﾞｯ　{System.Reflection.MethodBase.GetCurrentMethod()}");
            e.Handled = true;
        }

        private void TaskScheduler_UnobservedTaskException(object sender, UnobservedTaskExceptionEventArgs e) {
            e.Exception.ErrDump($"{System.Reflection.MethodBase.GetCurrentMethod()}");
        }

        void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e) {
            dynamic ex = e.ExceptionObject;
            e.ExceptionObject.ErrDump($"{System.Reflection.MethodBase.GetCurrentMethod()}");
        }

        void CurrentDomain_FirstChanceException(object sender, System.Runtime.ExceptionServices.FirstChanceExceptionEventArgs e) {
            e.Exception.ErrDump($"{System.Reflection.MethodBase.GetCurrentMethod()}");
        }

        [ExcelCommand(MenuName = "xlMdDna", MenuText = "About")]
        public static void About() {
            MessageBox.Show($@"xlMdDna.	Ver: {AppVer}

      THIRD-PARTY SOFTWARE NOTICES AND INFORMATION
        https://excel-dna.net/
        https://mermaid.live/
");
        }
    }
}