using System;
using System.Collections.Generic;
using System.Text;

namespace WebGear.GoogleContactsSync
{
    internal interface ILogger
    {
        void Log(string message, EventType eventType);

        void ClearLog();

        void LogUnique(string message, EventType eventType);
    }

    internal enum EventType
    {
        Information,
        Warning,
        Error
    }

    internal class Logger : ILogger
    {
        public string messages;

        public Logger() { }

        #region ILogger Members

        public void Log(string message, EventType eventType)
        {
            message = DateTime.Now + " " + eventType + ": " + message;
            messages += message + Environment.NewLine;
        }

        public void LogUnique(string message, EventType eventType)
        {
            if (!messages.Contains(eventType + ": " + message))
            {
                message = DateTime.Now + " " + eventType + ": " + message;
                messages += message + Environment.NewLine;
            }
        }

        public void ClearLog()
        {
            messages = "";
        }

        #endregion
    }
}
