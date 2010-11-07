using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace WebGear.GoogleContactsSync
{
    static class Program
    {
		private static SettingsForm instance;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
			instance = new SettingsForm();
            Application.Run(instance);
        }

		internal static SettingsForm Instance
		{
			get { return instance; }
		}

        /// <summary>
        /// Fallback. If there is some try/catch missing we will handle it here, just before the application quits unhandled
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception)
                ErrorHandler.Handle((Exception)e.ExceptionObject);
        }
    }
}