using System;
using System.Collections.Generic;
using System.IO;

namespace GoContactSyncMod
{
    enum EventType
    {
        Debug,
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

        public static readonly string Folder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\GoContactSyncMOD\\";

        static Logger()
        {
            if (!Directory.Exists(Folder))
                Directory.CreateDirectory(Folder);
            try
            {
                string logFileName = Folder + "log.txt";
                
                //If log file is bigger than 1 MB, move it to backup file and create new file
                FileInfo logFile = new FileInfo(logFileName);
                if (logFile.Exists && logFile.Length >= 1000000)
                    File.Move(logFileName, logFileName + "_" + DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss"));

                logwriter = new StreamWriter(logFileName, true);
            }
            catch (Exception ex)
            {
                ErrorHandler.Handle(ex);
            }
        }
    
        public static void Close()
        {
            try
            {
                if(logwriter!=null)
                    logwriter.Close();
            }
            catch(Exception e)
            {
                ErrorHandler.Handle(e);
            }
        }

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

            try
            {
                logwriter.Write(GetLogLine(new_logEntry));
                logwriter.Flush();
            }
            catch (Exception)
            {
                //ignore it, because if you handle this error, the handler will again log the message
                //ErrorHandler.Handle(ex);
            }

            //Populate LogMessage to all subscribed Logger-Outputs, but only if not Debug message, Debug messages are only logged to logfile
            if (LogUpdated != null && eventType > EventType.Debug)
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