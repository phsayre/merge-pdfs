using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;

namespace mergeConvertedFolders
{
    /// <summary>
    /// Keeps logs of errors and operations.
    /// Option to turn on or off in app.config.
    /// </summary>
    public static class WriteOut
    {
        /*OPTIONS:
         *  writeSwitch=0 does nothing
         *  writeSwitch=1 writes message to log 
         *  writeSwitch=2 writes message to console
         *  writeSwitch=3 writes message to both log and console
         */

        private static string logPath;
        private static int writeSwitch;  //specified in app.config

        /// <summary>
        /// Decides how to record the log message based on app.config settings.
        /// </summary>
        /// <param name="message">The message to log</param>
        public static void HandleMessage(string message)
        {
            writeSwitch = Int32.Parse(ConfigurationManager.AppSettings["writeSwitch"]);

            switch (writeSwitch)
            {
                case 0:
                    //do nothing
                    break;
                case 1:
                    Log(message);
                    break;
                case 2:
                    ConsoleWrite(message);
                    break;
                case 3:
                    Log(message);
                    ConsoleWrite(message);
                    break;
            }
        }

        /// <summary>
        /// Writes message to the console.
        /// </summary>
        /// <param name="consoleMessage">The message to write.</param>
        public static void ConsoleWrite(string consoleMessage)
        {
            try
            {
                Console.WriteLine(consoleMessage);
            }
            catch
            {
                //do nothing
            }
        }

        /// <summary>
        /// Writes message to the log file.
        /// </summary>
        /// <param name="logMessage">The message to log.</param>
        public static void Log(string logMessage)
        {
            logPath = ConfigurationManager.AppSettings["logPath"];
            StringBuilder sb = new StringBuilder();

            try
            {
                DateTime dt = DateTime.Now;
                sb.Append(String.Format("{0:MM/dd/yyyy hh:mm:ss tt}", dt));
                sb.Append("  |  " + logMessage);
                sb.Append("\r\n");
                File.AppendAllText(logPath, sb.ToString());
                sb.Clear();
            }
            catch
            {
                //do nothing
            }
        }
    }
}
