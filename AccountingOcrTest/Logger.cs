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
        private static readonly object LogWriteLock = new object();

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
            lock (LogWriteLock)
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
                        try
                        {
                            if (File.Exists(LogFile))
                            {
                                var info = new FileInfo(LogFile);
                                if (info.Length > 10 * 1024 * 1024) // 10MB limit
                                {
                                    string archiveFile = "app_old.log";
                                    if (File.Exists(archiveFile))
                                    {
                                        File.Delete(archiveFile);
                                    }
                                    File.Move(LogFile, archiveFile);
                                }
                            }
                        }
                        catch
                        {
                            // If move fails, try truncating the file to release space
                            try { File.WriteAllText(LogFile, string.Empty); } catch { }
                        }

                        File.AppendAllText(LogFile, sb.ToString());
                    }
                }
                catch { }
            }
        }
    }
}
