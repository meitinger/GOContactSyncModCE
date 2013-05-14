using System;
using System.Threading;
using System.Windows.Forms;

namespace GoContactSyncMod
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            bool isFirstInstance;
            using (new Mutex(false, "ACBBBC09-F76C-4874-AAFF-4F3353A5A5A6", out isFirstInstance))
            {
                if (!isFirstInstance)
                {
                    SettingsForm.ShowRemote();
                    return;
                }
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new SettingsForm());
            }
        }
    }
}
