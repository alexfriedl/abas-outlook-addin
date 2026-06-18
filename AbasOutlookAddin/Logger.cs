using System;
using System.IO;

namespace AbasOutlookAddin
{
    /// <summary>
    /// Einfaches File-Logging für Diagnose und Audit-Trail.
    /// Log-Dateien landen in %LOCALAPPDATA%\AbasOutlookAddin\Logs\
    /// </summary>
    public static class Logger
    {
        private static readonly string LogDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AbasOutlookAddin", "Logs");

        private static readonly string LogFile = Path.Combine(LogDir,
            $"addin_{DateTime.Now:yyyyMMdd}.log");

        private static readonly object _lock = new object();

        static Logger()
        {
            try { Directory.CreateDirectory(LogDir); }
            catch { }
        }

        public static void Log(string message)
        {
            Write("INFO", SanitizeLogMessage(message));
        }

        public static void LogError(string message, Exception ex = null)
        {
            string sanitized = SanitizeLogMessage(message);
            Write("ERROR", ex != null
                ? $"{sanitized}: {SanitizeLogMessage(ex.Message)}\n{ex.StackTrace}"
                : sanitized);
        }

        /// <summary>
        /// Entfernt Newlines und Control-Characters aus Log-Nachrichten,
        /// um Log Injection zu verhindern.
        /// </summary>
        private static string SanitizeLogMessage(string message)
        {
            if (string.IsNullOrEmpty(message))
                return "(empty)";

            // Newlines und Control-Characters ersetzen
            var sanitized = new System.Text.StringBuilder(message.Length);
            foreach (char c in message)
            {
                if (c == '\r' || c == '\n')
                    sanitized.Append(" | ");
                else if (char.IsControl(c))
                    sanitized.Append('_');
                else
                    sanitized.Append(c);
            }

            // Länge begrenzen um Log-Flooding zu verhindern
            const int MaxLogLength = 1000;
            if (sanitized.Length > MaxLogLength)
            {
                sanitized.Length = MaxLogLength;
                sanitized.Append("...(truncated)");
            }

            return sanitized.ToString();
        }

        private static void Write(string level, string message)
        {
            try
            {
                lock (_lock)
                {
                    File.AppendAllText(LogFile,
                        $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}] {message}{Environment.NewLine}");
                }
            }
            catch { /* Logging darf nie crashen */ }
        }
    }
}
