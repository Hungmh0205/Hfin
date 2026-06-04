using System;
using System.Collections.Concurrent;
using System.IO;
using System.Text;
using System.Threading;

namespace AccountingOcrTest
{
    public static class Logger
    {
        private static readonly string LogFile = "app.log";
        private static readonly ConcurrentQueue<string> _logQueue = new ConcurrentQueue<string>();
        private static readonly Timer _flushTimer;

        static Logger()
        {
            _flushTimer = new Timer(FlushLogs, null, 1000, 1000);
            AppDomain.CurrentDomain.ProcessExit += (s, e) => FlushLogs(null);
        }

        public static void Log(string message)
        {
            _logQueue.Enqueue($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}");
        }

        public static void Flush()
        {
            FlushLogs(null);
        }

        private static void FlushLogs(object? state)
        {
            try
            {
                var sb = new StringBuilder();
                while (_logQueue.TryDequeue(out var msg))
                {
                    sb.AppendLine(msg);
                }
                if (sb.Length > 0)
                {
                    File.AppendAllText(LogFile, sb.ToString());
                }
            }
            catch { }
        }
    }
}
