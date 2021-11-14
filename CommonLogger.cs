using System;
using System.Diagnostics;
using System.IO;

namespace MAUTO
{
    public static class CommonLogger
    {
        public static bool pComLogFlag = false;

        public static void WriteLine(string message)
        {
            string logFile;
            logFile = Process.GetCurrentProcess().MainModule.FileName;
            logFile = Path.GetFileNameWithoutExtension(logFile) + ".log";
            WriteLine(logFile, message);
        }

        public static void WriteLine(string logFile, string message)
        {
            // 
            string logLine;
            try
            {
                // 日付取得
                logLine = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + " {0}";
                logLine = string.Format(logLine, message);
                logLine = logLine + "\r\n";

                // 文言を書き込む
                if (pComLogFlag == true) { File.AppendAllText(logFile, logLine); }
                else { Console.WriteLine(logLine); }
            }
            catch (Exception ex)
            {
                // 日付取得
                logLine = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + " {0}";
                logLine = string.Format(logLine, ex.Message);
                logLine = logLine + "\r\n";
                // 文言を書き込む
                if (pComLogFlag == true) { File.AppendAllText("ErrorLog.log", logLine); }
                else { Console.WriteLine(logLine); }
            }
        }
    }
}
