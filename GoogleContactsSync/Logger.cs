using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace GoContactSyncMod
{
    enum EventType
    {
        Debug,
        Information,
        Warning,
        Error
    }

    static class Logger
    {
        static StreamWriter logWriter = null;

        #region Win32

        const int WM_VSCROLL = 0x0115;
        const int SB_BOTTOM = 7;
        const string WC_EDIT = "Edit";

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern bool PostMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        #endregion

        public static string Folder
        {
            get
            {
                // create the path and ensure that all directories exist
                var folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Google\Google Contacts Sync");
                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);
                return folder;
            }
        }

        public static void Show()
        {
            // specify the start information (use CreateProcess not ShellExecute)
            var psi = new ProcessStartInfo()
            {
                FileName = "notepad.exe",
                Arguments = '"' + Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Google\Google Contacts Sync\Log.txt") + '"',
                UseShellExecute = false,
            };

            // start notepad
            using (var proc = Process.Start(psi))
            {
                // locate the edit window after everthing has been initialized
                proc.WaitForInputIdle();
                if (proc.MainWindowHandle == IntPtr.Zero)
                    return;
                var edit = FindWindowEx(proc.MainWindowHandle, IntPtr.Zero, WC_EDIT, null);
                if (edit == IntPtr.Zero)
                    return;

                // scroll to the end
                PostMessage(edit, WM_VSCROLL, (IntPtr)SB_BOTTOM, IntPtr.Zero);
            }
        }

        public static void Log(string message, EventType eventType)
        {
#if !DEBUG
            // skip all debug messages
            if (eventType == EventType.Debug)
                return;
#endif

            // ensure that a writer has been opened
            if (logWriter == null)
            {
                try { logWriter = new StreamWriter(Path.Combine(Folder, "Log.txt"), true, Encoding.UTF8); }
                catch { return; }
            }

            // append the message and close the writer if there has been an I/O error
            try
            {
                logWriter.Write(String.Format("[{0} | {1}]\t{2}\r\n", DateTime.Now, eventType, message));
                logWriter.Flush();
            }
            catch (IOException)
            {
                try { logWriter.Close(); }
                catch { }
                logWriter = null;
            }
            catch { }
        }
    }
}
