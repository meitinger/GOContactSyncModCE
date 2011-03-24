using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace GoContactSyncMod
{
    enum EventType
    {
        Information,
        Warning,
        Error
    }

    struct LogEntry
    {
        public DateTime date;
        public EventType type;
        public string msg;

        public LogEntry(DateTime _date, EventType _type, string _msg)
        {
            date = _date; type = _type;  msg = _msg;
        }
    }

    static class Logger
    {
		public static List<LogEntry> messages = new List<LogEntry>();
		public delegate void LogUpdatedHandler(string Message);
		public static event LogUpdatedHandler LogUpdated;
        private static StreamWriter logwriter;
    
        private static string formatMessage(string message, EventType eventType)
        {
            return String.Format("{0}:{1}{2}", eventType, Environment.NewLine, message);
        }

		private static string GetLogLine(LogEntry entry)
        {
            return String.Format("[{0} | {1}]\t{2}\r\n", entry.date, entry.type, entry.msg);
        }

		public static void Log(string message, EventType eventType)
        {
            
            LogEntry new_logEntry = new LogEntry(DateTime.Now, eventType, message);
            messages.Add(new_logEntry);
            using (logwriter = new StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)+"\\GoContactSyncMOD\\log.txt", true))
            {
                logwriter.Write(GetLogLine(new_logEntry));
                logwriter.Flush();
                logwriter.Close();
            }
            if (LogUpdated != null)
                LogUpdated(GetLogLine(new_logEntry));
        }

        /*
        public void LogUnique(string message, EventType eventType)
        {
            string logMessage = formatMessage(message, eventType);
            if (!messages.ContainsValue(logMessage))
                messages.Add(DateTime.Now, logMessage); //TODO: Outdated, no dictionary used anymore.
        }
        */

		public static void ClearLog()
        {
            messages.Clear();
        }

        /*
        public string GetText()
        {
            StringBuilder output = new StringBuilder();
            foreach (var logitem in messages)
                output.AppendLine(GetLogLine(logitem));

            return output.ToString();
        }
        */
    }
}