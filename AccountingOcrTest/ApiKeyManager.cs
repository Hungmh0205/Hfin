using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace AccountingOcrTest
{
    public class ApiKeyManager
    {
        private const string KeysFilePath = "apikeys.json";
        private const int MaxRequestsPerMinute = 5;
        private readonly object _lockObj = new object();
        private readonly Timer _cooldownTimer;
        
        public ObservableCollection<ApiKeyInfo> Keys { get; set; } = new ObservableCollection<ApiKeyInfo>();

        public ApiKeyManager()
        {
            LoadKeys();
            _cooldownTimer = new Timer(UpdateCooldowns, null, 1000, 1000);
        }

        private void UpdateCooldowns(object state)
        {
            lock (_lockObj)
            {
                foreach (var k in Keys)
                {
                    if (k.Status == KeyStatus.RateLimited)
                    {
                        if ((DateTime.Now - k.WindowStart).TotalSeconds >= 60)
                        {
                            k.RequestsInLastMinute = 0;
                            k.Status = KeyStatus.Ready;
                            Logger.Log($"[ApiKeyManager] Key {k.Key.Substring(0, 5)}... đã hồi phục trạng thái Ready.");
                        }
                        k.RaiseCooldownChanged();
                    }
                }
            }
        }

        public void LoadKeys()
        {
            if (File.Exists(KeysFilePath))
            {
                try
                {
                    string json = File.ReadAllText(KeysFilePath);
                    var keys = JsonSerializer.Deserialize<List<string>>(json);
                    Keys.Clear();
                    if (keys != null)
                    {
                        foreach (var k in keys)
                        {
                            Keys.Add(new ApiKeyInfo { Key = k, Status = KeyStatus.Ready });
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"[ApiKeyManager] Lỗi đọc file keys: {ex.Message}");
                }
            }
        }

        public void SaveKeys()
        {
            try
            {
                var keys = Keys.Select(k => k.Key).ToList();
                string json = JsonSerializer.Serialize(keys);
                File.WriteAllText(KeysFilePath, json);
            }
            catch (Exception ex)
            {
                Logger.Log($"[ApiKeyManager] Lỗi lưu file keys: {ex.Message}");
            }
        }

        public async Task<string> GetNextAvailableKeyAsync(CancellationToken token = default)
        {
            while (true)
            {
                token.ThrowIfCancellationRequested();
                
                lock (_lockObj)
                {
                    foreach (var k in Keys)
                    {
                        if (k.RequestsInLastMinute > 0 && (DateTime.Now - k.WindowStart).TotalSeconds >= 60)
                        {
                            k.RequestsInLastMinute = 0;
                            if (k.Status == KeyStatus.RateLimited)
                            {
                                k.Status = KeyStatus.Ready;
                                Logger.Log($"[ApiKeyManager] Key {k.Key.Substring(0, 5)}... đã hồi phục trạng thái Ready.");
                            }
                        }
                    }

                    var availableKey = Keys
                        .Where(k => k.Status == KeyStatus.Ready && k.RequestsInLastMinute < MaxRequestsPerMinute)
                        .OrderBy(k => k.LastUsed)
                        .FirstOrDefault();

                    if (availableKey != null)
                    {
                        if (availableKey.RequestsInLastMinute == 0)
                            availableKey.WindowStart = DateTime.Now;

                        availableKey.RequestsInLastMinute++;
                        availableKey.TotalRequests++;
                        availableKey.EstimatedCost += 0.000075; // Baseline Gemini 1.5 Flash cost estimate
                        availableKey.LastUsed = DateTime.Now;
                        Logger.Log($"[ApiKeyManager] Cấp phát Key {availableKey.Key.Substring(0, 5)}... (Lần dùng thứ {availableKey.RequestsInLastMinute}/phút)");
                        
                        if (availableKey.RequestsInLastMinute >= MaxRequestsPerMinute)
                        {
                            availableKey.Status = KeyStatus.RateLimited;
                            Logger.Log($"[ApiKeyManager] Key {availableKey.Key.Substring(0, 5)}... đã chạm ngưỡng (Rate Limited).");
                        }
                        return availableKey.Key;
                    }
                }

                Logger.Log("[ApiKeyManager] Không có Key nào rảnh, đang chờ 2 giây...");
                await Task.Delay(2000, token);
            }
        }

        public void RecordSuccess(string key)
        {
            lock (_lockObj)
            {
                var k = Keys.FirstOrDefault(x => x.Key == key);
                if (k != null)
                {
                    k.SuccessfulRequests++;
                }
            }
        }

        public void MarkKeyAsRateLimited(string key)
        {
            lock (_lockObj)
            {
                var k = Keys.FirstOrDefault(x => x.Key == key);
                if (k != null)
                {
                    k.Status = KeyStatus.RateLimited;
                    k.RequestsInLastMinute = MaxRequestsPerMinute;
                    k.WindowStart = DateTime.Now;
                    k.FailedRequests++;
                    Logger.Log($"[ApiKeyManager] Đánh dấu Key {k.Key.Substring(0, 5)}... gặp lỗi Rate Limit (429/quota). Bắt đầu thời gian chờ 60s.");
                }
            }
        }

        public void MarkKeyAsError(string key)
        {
            lock (_lockObj)
            {
                var k = Keys.FirstOrDefault(x => x.Key == key);
                if (k != null)
                {
                    k.Status = KeyStatus.Error;
                    k.FailedRequests++;
                    Logger.Log($"[ApiKeyManager] Đánh dấu Key {k.Key.Substring(0, 5)}... bị lỗi vĩnh viễn (Error/403/400).");
                }
            }
        }
    }
}
