namespace ReadExcel
{
    using System;
    using System.Windows.Media;
    using System.Runtime.CompilerServices;
    using System.Runtime.InteropServices;
    using System.Windows;

    public static class ExceptionExtensions
    {
        public static Exception GetOriginalException(this Exception ex)
        {
            if (ex.InnerException == null) return ex;

            return ex.InnerException.GetOriginalException();
        }
    }

    public enum LogType
    {
        DEBUG,
        INFO,
        NOTE,
        WARNING,
        ERROR
    }

    public class Logging
    {
        public static bool deaf;
        public static string exceptionsPerRun;
        public static bool isRecording;
        public static bool LogConsole = true;
        public static string logPerRun;
        public static DateTime startPerRun;
        public static MainWindow ui;
        public static LogType level = LogType.INFO;
        private delegate void UpdateLogDelegate(string log);

        private static void log(string msg, LogType type = LogType.INFO)
        {
            msg = string.Format("{0:HH:mm:ss.fff}  {1}", DateTime.Now, msg);
            if (((ui != null)) && ui.IsLoaded)
            {
                object[] objArray;
                if (type == LogType.ERROR)
                {
                    objArray = new object[] { msg, (Color)ColorConverter.ConvertFromString("Red") };
                }
                else if (type == LogType.NOTE)
                {
                    objArray = new object[] { msg, (Color)ColorConverter.ConvertFromString("Blue") };
                }
                else if (type == LogType.WARNING)
                {
                    objArray = new object[] { msg, (Color)ColorConverter.ConvertFromString("DarkOrange") };
                }
                else if (type == LogType.INFO)
                {
                    objArray = new object[] { msg, (Color)ColorConverter.ConvertFromString("DarkGreen") };
                }
                else
                {
                    objArray = new object[] { msg, (Color)ColorConverter.ConvertFromString("Black") };
                }
                try
                {
                    ui.Dispatcher.Invoke(ui.updateLogDelegate, objArray);
                }
                catch (Exception exception)
                {
                    logException(exception);
                }
            }
            if (LogConsole)
            {
                Console.WriteLine(msg);
            }
            if (isRecording)
            {
                string str = string.Empty;
                switch (type)
                {
                    case LogType.INFO:
                        str = "__cfg__ ";
                        break;

                    case LogType.ERROR:
                        str = "__cfr__ ";
                        break;

                    case LogType.WARNING:
                        str = "__cfy__ ";
                        break;

                    case LogType.NOTE:
                        str = "__cfb__ ";
                        break;

                    default:
                        str = "__cfg__ ";
                        break;
                }
                logPerRun = logPerRun + str + msg + Environment.NewLine;
            }
        }

        public static void logException(Exception ex)
        {
            Console.WriteLine(ex.ToString());
            exceptionsPerRun = exceptionsPerRun + ex.ToString();
            exceptionsPerRun = exceptionsPerRun + "\r\n===============================================\r\n\r\n";
        }

        public static void logMessage(object obj)
        {
            logMessage(obj.ToString(), LogType.INFO, 0);
        }

        public static void logMessage(string msg, object obj)
        {
            logMessage(msg, LogType.INFO, 0);
            logMessage(obj.ToString(), LogType.INFO, 0);
        }

        public static void logMessage(string msg, LogType type = LogType.INFO, int indent_level = 0)
        {
            if (type < level)
                return;
            if (!deaf)
            {
                string str = string.Empty;
                for (int i = 0; i < indent_level; i++)
                {
                    str = str + "\t";
                }
                msg = str + msg;
                log(msg, type);
            }
        }

        
    }
}

