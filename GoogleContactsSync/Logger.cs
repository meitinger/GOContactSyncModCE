using System;
using System.IO;
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

        public static string Folder
        {
            get
            {
                var folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Google\Google Contacts Sync");
                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);
                return folder;
            }
        }

        public static void Log(string message, EventType eventType)
        {
            if (logWriter == null)
            {
                try { logWriter = new StreamWriter(Path.Combine(Folder, "Log.txt"), true, Encoding.UTF8); }
                catch { return; }
            }
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
