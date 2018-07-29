using System;
using System.IO;
using System.Threading;

namespace Room17.MeetingDecline.Util
{
    public class Logger
    {
        private static readonly object _lock = new object();
        private static readonly string _LogFile = System.IO.Path.GetTempFileName();
        public static bool DEBUG { get; set; }

        // TODO: overload very log method to have a String params... method with String.Format
        public static void Error(string message)
        {
            Log("ERROR", message);
        }

        public static void Warning(string message)
        {
            Log("WARNING", message);
        }

        public static void Info(string message)
        {
            Log("INFO", message);
        }

        public static void Debug(string message)
        {
            if (DEBUG)
                Log("DEBUG", message);
        }

        private static void Log(string level, string message)
        {
            lock (_lock)
            {
                File.AppendAllText(_LogFile, String.Format("{0} {1} [{2}]: {3}{4}", level, DateTime.Now, Thread.CurrentThread.ManagedThreadId, message, Environment.NewLine));
            }
        }
    }
}