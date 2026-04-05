#region License statement
/* SnakeTail is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, version 3 of the License.
 */
#endregion

using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;

namespace SnakeTail
{
    /// <summary>应用日志：日常写入 YYYY-MM-DD.log，崩溃写入 YYYY-MM-DD_crash.log；路径相对于当前工作目录下的 log 子目录。</summary>
    internal static class AppLog
    {
        public const string LevelDebug = "DEBUG";
        public const string LevelInfo = "INFO";
        public const string LevelWarn = "WARN";
        public const string LevelErr = "ERR";
        public const string LevelFatal = "FATAL";

        public static string GetLogDirectory()
        {
            return Path.Combine(Directory.GetCurrentDirectory(), "log");
        }

        public static string GetDailyLogFilePath()
        {
            return Path.Combine(GetLogDirectory(), $"{DateTime.Now:yyyy-MM-dd}.log");
        }

        public static string GetCrashLogFilePath()
        {
            return Path.Combine(GetLogDirectory(), $"{DateTime.Now:yyyy-MM-dd}_crash.log");
        }

        public static string FormatLine(string level, string sourceFile, int sourceLine, string message)
        {
            string filePart = string.IsNullOrEmpty(sourceFile) ? "?" : Path.GetFileName(sourceFile);
            return $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] [{filePart}:{sourceLine}] [{message}]";
        }

        public static void AppendDaily(string level, string message, [CallerFilePath] string file = "", [CallerLineNumber] int line = 0)
        {
            try
            {
                Directory.CreateDirectory(GetLogDirectory());
                File.AppendAllText(GetDailyLogFilePath(), FormatLine(level, file, line, message) + Environment.NewLine, Encoding.UTF8);
            }
            catch
            {
                // 日志失败时静默，避免递归异常
            }
        }

        public static void AppendCrash(Exception ex, string extraContext = null, [CallerFilePath] string file = "", [CallerLineNumber] int line = 0)
        {
            try
            {
                if (ex == null)
                    ex = new Exception("Unknown exception (null reference)");

                Directory.CreateDirectory(GetLogDirectory());
                var sb = new StringBuilder();
                sb.AppendLine(FormatLine(LevelFatal, file, line, ex.GetType().FullName + ": " + ex.Message));
                sb.AppendLine(ex.ToString());
                if (!string.IsNullOrEmpty(extraContext))
                {
                    sb.AppendLine(extraContext);
                }
                File.AppendAllText(GetCrashLogFilePath(), sb.ToString(), Encoding.UTF8);
            }
            catch
            {
            }
        }
    }
}
