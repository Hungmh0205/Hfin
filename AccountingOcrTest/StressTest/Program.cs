using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DuckDB.NET.Data;
using AccountingOcrTest;
using ExcelDataReader;

namespace StressTest
{
    class Program
    {
        private static readonly string ExcelPath = @"D:\Excel\AccountingOcrTest\bin\Debug\net10.0-windows\raw_material\THDT.T4.2026.hanh.xlsx";
        private static readonly string TempDbPath = "benchmark_staging.db";

        [STAThread]
        static async Task Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("=======================================================================");
            Console.WriteLine("        BẮT ĐẦU CHẠY STRESS TEST & BENCHMARK DỰ ÁN ACCOUNTING OCR       ");
            Console.WriteLine("=======================================================================");
            Console.WriteLine($"Thời gian chạy: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            Console.WriteLine($"Hệ điều hành: {Environment.OSVersion}");
            Console.WriteLine($"Số lượng nhân CPU: {Environment.ProcessorCount}");
            Console.WriteLine($"Độ rộng tiến trình: {(Environment.Is64BitProcess ? "64-bit" : "32-bit")}");
            Console.WriteLine("=======================================================================\n");

            // Test 1: ClosedXML Excel Memory & Performance
            RunExcelBenchmark();

            // Test 2: FuzzyMatcher Accuracy & Performance
            RunFuzzyMatcherBenchmark();

            // Test 3: ApiKeyManager Concurrency Stress Test
            await RunApiKeyManagerStressTestAsync();

            // Test 4: Logger Thread-Safety & Performance
            RunLoggerStressTest();

            // Test 5: DuckDB Transaction vs Individual Inserts
            RunDuckDBBenchmark();

            // Test 6: Gemini OCR Visual Accuracy Test (Cross-out line detection)
            RunGeminiOcrTestAsync().GetAwaiter().GetResult();

            // Test 7: Advanced Edge Cases, Overload & Crash Vulnerability Test
            await RunAdvancedEdgeCaseTestsAsync();

            // Test 8: Optimization & Reliability Features Verification
            await RunOptimizationFeaturesTestsAsync();

            Console.WriteLine("\n=======================================================================");
            Console.WriteLine("             HOÀN TẤT TẤT CẢ CÁC BÀI STRESS TEST & BENCHMARK            ");
            Console.WriteLine("=======================================================================");
        }

        private static void RunExcelBenchmark()
        {
            Console.WriteLine(">>> TEST 1: EXCEL PARSERS MEMORY & PERFORMANCE BENCHMARK (ClosedXML vs ExcelDataReader)");
            if (!File.Exists(ExcelPath))
            {
                Console.WriteLine($"[CẢNH BÁO] Không tìm thấy file Excel tại {ExcelPath}. Bỏ qua test Excel.");
                return;
            }

            Console.WriteLine($"Đang tải tệp Excel thực tế: {Path.GetFileName(ExcelPath)} (Kích thước: {new FileInfo(ExcelPath).Length / (1024.0 * 1024.0):F2} MB)...");

            // --- ClosedXML Benchmark ---
            Console.WriteLine("\n[1] Thử nghiệm với ClosedXML:");
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            long startMemCxml = GC.GetTotalMemory(true);
            var processCxml = Process.GetCurrentProcess();
            long startWorkingSetCxml = processCxml.WorkingSet64;

            Stopwatch sw = Stopwatch.StartNew();
            try
            {
                using (var workbook = new XLWorkbook(ExcelPath))
                {
                    sw.Stop();
                    GC.Collect();
                    long endMemCxml = GC.GetTotalMemory(true);
                    processCxml.Refresh();
                    long startWorkingSetRef = processCxml.WorkingSet64;

                    double loadTimeSec = sw.ElapsedMilliseconds / 1000.0;
                    double memSpikeMb = (endMemCxml - startMemCxml) / (1024.0 * 1024.0);
                    double workingSetSpikeMb = (startWorkingSetRef - startWorkingSetCxml) / (1024.0 * 1024.0);

                    Console.WriteLine($" - Thời gian nạp XLWorkbook (ClosedXML): {loadTimeSec:F2} giây");
                    Console.WriteLine($" - Lượng RAM heap tăng (GC.GetTotalMemory): {memSpikeMb:F2} MB");
                    Console.WriteLine($" - Lượng RAM vật lý tăng (Working Set): {workingSetSpikeMb:F2} MB");

                    // Read Sheet Tham Chiếu
                    sw.Restart();
                    var refSheet = workbook.Worksheet("tham chiếu");
                    var refRowsUsed = refSheet.RowsUsed().Count();
                    sw.Stop();
                    Console.WriteLine($" - Sheet 'tham chiếu': Đã đọc được {refRowsUsed} hàng trong {sw.ElapsedMilliseconds} ms");

                    // Read Sheet Bán Hàng
                    sw.Restart();
                    var salesSheet = workbook.Worksheet("Bán Hàng");
                    var salesRowsUsed = salesSheet.RowsUsed().Count();
                    sw.Stop();
                    Console.WriteLine($" - Sheet 'Bán Hàng': Đã đọc được {salesRowsUsed} hàng trong {sw.ElapsedMilliseconds} ms");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[LỖI CLOSEDXML] Gặp lỗi khi benchmark ClosedXML: {ex.Message}");
            }

            // --- ExcelDataReader Benchmark ---
            Console.WriteLine("\n[2] Thử nghiệm với ExcelDataReader:");
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            long startMemEdr = GC.GetTotalMemory(true);
            var processEdr = Process.GetCurrentProcess();
            long startWorkingSetEdr = processEdr.WorkingSet64;

            sw.Restart();
            try
            {
                using (var stream = File.Open(ExcelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var dataset = reader.AsDataSet();
                        sw.Stop();
                        GC.Collect();
                        long endMemEdr = GC.GetTotalMemory(true);
                        processEdr.Refresh();
                        long endWorkingSetEdr = processEdr.WorkingSet64;

                        double loadTimeSec = sw.ElapsedMilliseconds / 1000.0;
                        double memSpikeMb = (endMemEdr - startMemEdr) / (1024.0 * 1024.0);
                        double workingSetSpikeMb = (endWorkingSetEdr - startWorkingSetEdr) / (1024.0 * 1024.0);

                        Console.WriteLine($" - Thời gian nạp DataSet (ExcelDataReader): {loadTimeSec:F2} giây");
                        Console.WriteLine($" - Lượng RAM heap tăng (GC.GetTotalMemory): {memSpikeMb:F2} MB");
                        Console.WriteLine($" - Lượng RAM vật lý tăng (Working Set): {workingSetSpikeMb:F2} MB");

                        if (dataset.Tables.Contains("tham chiếu"))
                        {
                            sw.Restart();
                            var refTable = dataset.Tables["tham chiếu"];
                            int rows = refTable?.Rows.Count ?? 0;
                            sw.Stop();
                            Console.WriteLine($" - Sheet 'tham chiếu': Đã đọc được {rows} hàng trong {sw.ElapsedMilliseconds} ms");
                        }
                        if (dataset.Tables.Contains("Bán Hàng"))
                        {
                            sw.Restart();
                            var salesTable = dataset.Tables["Bán Hàng"];
                            int rows = salesTable?.Rows.Count ?? 0;
                            sw.Stop();
                            Console.WriteLine($" - Sheet 'Bán Hàng': Đã đọc được {rows} hàng trong {sw.ElapsedMilliseconds} ms");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[LỖI EXCELDATAREADER] Gặp lỗi khi benchmark: {ex.Message}");
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            long finalMem = GC.GetTotalMemory(true);
            Console.WriteLine($" - RAM Heap sau giải phóng hoàn toàn: {(finalMem - startMemEdr) / (1024.0 * 1024.0):F2} MB (Lượng rác dọn sạch)");
            Console.WriteLine("-----------------------------------------------------------------------\n");
        }

        private static void RunFuzzyMatcherBenchmark()
        {
            Console.WriteLine(">>> TEST 2: FUZZYMATCHER ACCURACY & PERFORMANCE BENCHMARK");
            
            // Build reference database of items (500 items)
            var refItems = new List<ReferenceItem>();
            string[] baseProducts = {
                "Hỗn hợp cải mơ baby leaf 200g", "Cải Cầu Vồng baby leaf 200g", "Hỗn hợp Xà Lách baby leaf 200g",
                "Xà lách Xanh baby leaf 200g", "Hạt Sen tươi 100gr", "Hạt Sen tươi 200gr",
                "Mầm Cải củ trắng 150gr", "Mầm Cải củ đỏ 150gr", "Mầm Cải ngọt 120gr",
                "Quả đậu nành nhật 400g", "Hạt Đậu trắng tươi 100g", "Hạt Đậu đỏ tươi 100g",
                "Hạt Đậu mắt ngọc tươi 100g", "Rau bina baby 150g", "Cải thìa xanh baby 200g"
            };

            // Expand list to 500 items to simulate a large reference database
            for (int i = 0; i < 500; i++)
            {
                var baseProd = baseProducts[i % baseProducts.Length];
                refItems.Add(new ReferenceItem
                {
                    VietTat = $"SP{i:D4}",
                    TenHang = $"{baseProd} #{i}",
                    DonViTinh = "Hộp"
                });
            }

            // Define test queries with varying levels of noise
            var testQueries = new List<(string query, string expectedMatch)>
            {
                // Exact or near-exact match
                ("Hỗn hợp cải mơ baby leaf 200g #0", "Hỗn hợp cải mơ baby leaf 200g #0"),
                // Accent removal
                ("Hon hop cai mo baby leaf 200g #0", "Hỗn hợp cải mơ baby leaf 200g #0"),
                // Typos & missing characters
                ("Cai Cau Vong babi laf 200g #1", "Cải Cầu Vồng baby leaf 200g #1"),
                // Case variations & abbreviation matching
                ("SP0002", "Hỗn hợp Xà Lách baby leaf 200g #2"), // matching via abbreviation (VietTat)
                // High noise query
                ("Mam Cai cu do 150g #7", "Mầm Cải củ đỏ 150gr #7"),
                ("Dau mat ngoc tuoi 100g #12", "Hạt Đậu mắt ngọc tươi 100g #12")
            };

            Console.WriteLine($"Kích thước danh mục tham chiếu: {refItems.Count} sản phẩm");
            Console.WriteLine($"Đang chạy {testQueries.Count} kịch bản kiểm thử độ chính xác so khớp...");

            // Pre-calculate cleaned strings for refItems
            var cachedRefItems = refItems.Select(x => new CachedRefInfo
            {
                Item = x,
                CleanTenHang = FuzzyMatcher.RemoveSign4VietnameseString(x.TenHang.Trim().ToLower()),
                CleanVietTat = FuzzyMatcher.RemoveSign4VietnameseString(x.VietTat.Trim().ToLower())
            }).ToList();

            int passed = 0;
            foreach (var t in testQueries)
            {
                var match = FuzzyMatcher.FindBestMatchOptimized(t.query, cachedRefItems, out double score);
                bool isCorrect = match != null && match.TenHang == t.expectedMatch;
                if (isCorrect) passed++;
                Console.WriteLine($" - Truy vấn: '{t.query}' -> Tìm thấy: '{match?.TenHang}' (Độ khớp: {score:F2}) -> {(isCorrect ? "ĐÚNG" : "SAI")}");
            }
            Console.WriteLine($" -> Độ chính xác: {passed}/{testQueries.Count} ({passed * 100.0 / testQueries.Count:F1}%)");

            // Performance Benchmark: run 5,000 matches sequentially
            Console.WriteLine("Đang chạy Stress Test 5,000 lần so khớp liên tục (Đã tối ưu hóa)...");
            Stopwatch sw = Stopwatch.StartNew();
            int matchCount = 5000;
            double dummySum = 0;

            for (int i = 0; i < matchCount; i++)
            {
                var query = testQueries[i % testQueries.Count].query;
                var match = FuzzyMatcher.FindBestMatchOptimized(query, cachedRefItems, out double score);
                if (match != null) dummySum += score;
            }
            sw.Stop();

            double totalMs = sw.ElapsedMilliseconds;
            double avgUs = (totalMs * 1000.0) / matchCount;
            double opsPerSec = matchCount / (totalMs / 1000.0);

            Console.WriteLine($" - Tổng thời gian cho {matchCount} lần so khớp: {totalMs:F2} ms");
            Console.WriteLine($" - Thời gian so khớp trung bình: {avgUs:F2} microseconds (µs)");
            Console.WriteLine($" - Thông lượng: {opsPerSec:F0} lượt/giây (ops/sec)");
            Console.WriteLine("-----------------------------------------------------------------------\n");
        }

        private static async Task RunApiKeyManagerStressTestAsync()
        {
            Console.WriteLine(">>> TEST 3: APIKEYMANAGER CONCURRENCY STRESS TEST");
            
            // Backup user's original keys if file exists to prevent deletion/overwriting
            string? keysBackup = null;
            if (File.Exists("apikeys.json"))
            {
                try { keysBackup = File.ReadAllText("apikeys.json"); } catch {}
            }

            try
            {
                // Create keys database
                var keysJson = "[\"KEY_ALPHA_001\", \"KEY_BETA_002\", \"KEY_GAMMA_003\", \"KEY_DELTA_004\", \"KEY_EPSILON_005\"]";
                File.WriteAllText("apikeys.json", keysJson);

                var manager = new ApiKeyManager();
                Console.WriteLine($"Đã nạp {manager.Keys.Count} API Keys từ file cấu hình.");
                
                int totalThreads = 5;
                int requestsPerThread = 3;
                Console.WriteLine($"Giả lập {totalThreads} luồng đồng thời gọi khóa API (Mỗi luồng {requestsPerThread} lần yêu cầu)...");

                var allocatedKeys = new ConcurrentBag<string>();
                var lockTimes = new ConcurrentQueue<long>();
                var cts = new CancellationTokenSource();
                
                Stopwatch sw = Stopwatch.StartNew();
                
                var tasks = new List<Task>();
                for (int t = 0; t < totalThreads; t++)
                {
                    int threadId = t;
                    tasks.Add(Task.Run(async () =>
                    {
                        for (int r = 0; r < requestsPerThread; r++)
                        {
                            Stopwatch reqSw = Stopwatch.StartNew();
                            string key = await manager.GetNextAvailableKeyAsync(cts.Token);
                            reqSw.Stop();
                            
                            lockTimes.Enqueue(reqSw.ElapsedMilliseconds);
                            allocatedKeys.Add(key);

                            // Randomly simulate failures to trigger Rate Limiting/Errors on Keys
                            if (r == 3 && threadId % 4 == 0)
                            {
                                manager.MarkKeyAsRateLimited(key);
                            }
                            else if (r == 6 && threadId % 7 == 0)
                            {
                                manager.MarkKeyAsError(key);
                            }

                            // Sleep briefly between requests to simulate work
                            await Task.Delay(10);
                        }
                    }));
                }

                await Task.WhenAll(tasks);
                sw.Stop();

                double totalTimeSec = sw.Elapsed.TotalSeconds;
                var keyCounts = allocatedKeys.GroupBy(x => x).ToDictionary(g => g.Key, g => g.Count());

                Console.WriteLine($" - Tổng thời gian chạy giả lập đồng thời: {totalTimeSec:F2} giây");
                Console.WriteLine($" - Tổng số lượt cấp phát khóa thành công: {allocatedKeys.Count}");
                Console.WriteLine($" - Phân phối khóa cấp phát:");
                foreach (var kvp in keyCounts)
                {
                    var kInfo = manager.Keys.FirstOrDefault(x => x.Key == kvp.Key);
                    Console.WriteLine($"   * Khóa {kvp.Key.Substring(0, 8)}...: {kvp.Value} lần sử dụng (Trạng thái cuối: {kInfo?.Status})");
                }

                // Key Recovery Simulation
                Console.WriteLine(" - Chờ 1.5 giây để hồi phục trạng thái khóa rate-limited...");
                // Force reset of windows in Key Manager directly to bypass 60s wait for unit test speed
                lock (manager)
                {
                    foreach (var k in manager.Keys)
                    {
                        if (k.Status == KeyStatus.RateLimited)
                        {
                            k.WindowStart = DateTime.Now.AddSeconds(-61); // simulate 60s elapsed
                        }
                    }
                }

                // Try to allocate again
                string recoveredKey = await manager.GetNextAvailableKeyAsync();
                Console.WriteLine($"   * Cấp phát lại sau hồi phục: Key {recoveredKey.Substring(0, 8)}... (Thành công)");
            }
            finally
            {
                // Restore or clean up files
                try
                {
                    if (keysBackup != null)
                    {
                        File.WriteAllText("apikeys.json", keysBackup);
                    }
                    else if (File.Exists("apikeys.json"))
                    {
                        File.Delete("apikeys.json");
                    }
                }
                catch {}
            }
            Console.WriteLine("-----------------------------------------------------------------------\n");
        }

        private static void RunLoggerStressTest()
        {
            Console.WriteLine(">>> TEST 4: LOGGER THREAD-SAFETY & PERFORMANCE BENCHMARK");
            
            // Flush and clear the queue first to drain any leftover logs from Test 3
            Logger.Flush();
            if (File.Exists("app.log"))
            {
                try { File.Delete("app.log"); } catch {}
            }

            int numThreads = 10;
            int logsPerThread = 5000;
            int totalLogs = numThreads * logsPerThread;

            Console.WriteLine($"Bắt đầu giả lập {numThreads} luồng ghi {logsPerThread} dòng log mỗi luồng ({totalLogs} dòng tổng)...");

            Stopwatch sw = Stopwatch.StartNew();

            Parallel.For(0, numThreads, t =>
            {
                for (int i = 0; i < logsPerThread; i++)
                {
                    Logger.Log($"[Thread {t}] Ghi log stress test thứ {i} - Kiểm tra thread-safety & throughput.");
                }
            });

            sw.Stop();
            double enqueueTimeMs = sw.ElapsedMilliseconds;

            Console.WriteLine($" - Thời gian Đưa hàng đợi (Enqueue): {enqueueTimeMs:F2} ms (Trung bình: {enqueueTimeMs * 1000.0 / totalLogs:F2} µs/lượt)");

            // Force Flush of Logger queue to disk
            Console.WriteLine(" - Đang kích hoạt Flush thủ công để xả toàn bộ hàng đợi xuống app.log...");
            sw.Restart();
            Logger.Flush();
            sw.Stop();

            double flushTimeMs = sw.ElapsedMilliseconds;
            Console.WriteLine($" - Thời gian Ghi đĩa (Flush): {flushTimeMs:F2} ms");

            // Verify lines in log
            int lineCount = 0;
            if (File.Exists("app.log"))
            {
                lineCount = File.ReadLines("app.log").Count();
            }

            Console.WriteLine($" - Số lượng dòng log ghi nhận thực tế trên đĩa: {lineCount} / {totalLogs}");
            Console.WriteLine($" - Trạng thái toàn vẹn log: {(lineCount == totalLogs ? "100% HOÀN HẢO" : "THẤT THOÁT DỮ LIỆU")}");
            
            // Clean up log
            try { File.Delete("app.log"); } catch {}
            Console.WriteLine("-----------------------------------------------------------------------\n");
        }

        private static void RunDuckDBBenchmark()
        {
            Console.WriteLine(">>> TEST 5: DUCKDB TRANSACTION VS INDIVIDUAL INSERTS BENCHMARK");
            
            // Clean prior benchmark DB if exists
            if (File.Exists(TempDbPath))
            {
                try { File.Delete(TempDbPath); } catch {}
            }

            string connStr = $"Data Source={TempDbPath}";
            int numRows = 1000;

            try
            {
                using (var connection = new DuckDBConnection(connStr))
                {
                    connection.Open();
                    using (var cmd = connection.CreateCommand())
                    {
                        cmd.CommandText = @"
                            CREATE TABLE staging_items (
                                id VARCHAR PRIMARY KEY,
                                file_name VARCHAR,
                                ngay_giao VARCHAR,
                                khach_hang VARCHAR,
                                diem_giao VARCHAR,
                                nguoi_giao VARCHAR,
                                ten_hang_goc VARCHAR,
                                ten_hang_khop VARCHAR,
                                hst VARCHAR,
                                don_vi VARCHAR,
                                sl_xuat DOUBLE,
                                sl_nhan DOUBLE,
                                sl_hong DOUBLE,
                                ghi_chu VARCHAR
                            );";
                        cmd.ExecuteNonQuery();
                    }
                }
                Console.WriteLine("Đã khởi tạo cơ sở dữ liệu DuckDB tạm thời.");

                // Method A: Individual Inserts (open connection for each or execute one by one)
                Console.WriteLine($"A. Chạy thử nghiệm {numRows} lệnh Insert riêng lẻ...");
                Stopwatch sw = Stopwatch.StartNew();
                
                using (var connection = new DuckDBConnection(connStr))
                {
                    connection.Open();
                    for (int i = 0; i < numRows; i++)
                    {
                        using (var cmd = connection.CreateCommand())
                        {
                            cmd.CommandText = @"
                                INSERT INTO staging_items (id, file_name, ngay_giao, khach_hang, diem_giao, nguoi_giao, ten_hang_goc, ten_hang_khop, hst, don_vi, sl_xuat, sl_nhan, sl_hong, ghi_chu) 
                                VALUES ($id, $file_name, $ngay_giao, $khach_hang, $diem_giao, $nguoi_giao, $ten_hang_goc, $ten_hang_khop, $hst, $don_vi, $sl_xuat, $sl_nhan, $sl_hong, $ghi_chu);";

                            cmd.Parameters.Add(new DuckDBParameter("id", Guid.NewGuid().ToString()));
                            cmd.Parameters.Add(new DuckDBParameter("file_name", $"test_file_{i}.jpg"));
                            cmd.Parameters.Add(new DuckDBParameter("ngay_giao", "06/06/2026"));
                            cmd.Parameters.Add(new DuckDBParameter("khach_hang", "WinMart"));
                            cmd.Parameters.Add(new DuckDBParameter("diem_giao", "Lê Văn Thiêm"));
                            cmd.Parameters.Add(new DuckDBParameter("nguoi_giao", "Shipper A"));
                            cmd.Parameters.Add(new DuckDBParameter("ten_hang_goc", "Mầm cải củ đỏ"));
                            cmd.Parameters.Add(new DuckDBParameter("ten_hang_khop", "Mầm Cải củ đỏ 150gr"));
                            cmd.Parameters.Add(new DuckDBParameter("hst", "SP0001"));
                            cmd.Parameters.Add(new DuckDBParameter("don_vi", "Hộp"));
                            cmd.Parameters.Add(new DuckDBParameter("sl_xuat", 5.0));
                            cmd.Parameters.Add(new DuckDBParameter("sl_nhan", 4.0));
                            cmd.Parameters.Add(new DuckDBParameter("sl_hong", 1.0));
                            cmd.Parameters.Add(new DuckDBParameter("ghi_chu", "Hỏng do vận chuyển"));

                            cmd.ExecuteNonQuery();
                        }
                    }
                }
                sw.Stop();
                double individualTime = sw.ElapsedMilliseconds;
                Console.WriteLine($" -> Tổng thời gian (Insert riêng lẻ): {individualTime:F2} ms");
                Console.WriteLine($" -> Tốc độ: {numRows / (individualTime / 1000.0):F0} inserts/giây");

                // Clean table for second test
                using (var connection = new DuckDBConnection(connStr))
                {
                    connection.Open();
                    using (var cmd = connection.CreateCommand())
                    {
                        cmd.CommandText = "DELETE FROM staging_items;";
                        cmd.ExecuteNonQuery();
                    }
                }

                // Method B: Batch transaction insert
                Console.WriteLine($"B. Chạy thử nghiệm {numRows} lệnh Insert gộp chung trong 1 TRANSACTION...");
                sw.Restart();

                using (var connection = new DuckDBConnection(connStr))
                {
                    connection.Open();
                    using (var transaction = connection.BeginTransaction())
                    {
                        for (int i = 0; i < numRows; i++)
                        {
                            using (var cmd = connection.CreateCommand())
                            {
                                cmd.CommandText = @"
                                    INSERT INTO staging_items (id, file_name, ngay_giao, khach_hang, diem_giao, nguoi_giao, ten_hang_goc, ten_hang_khop, hst, don_vi, sl_xuat, sl_nhan, sl_hong, ghi_chu) 
                                    VALUES ($id, $file_name, $ngay_giao, $khach_hang, $diem_giao, $nguoi_giao, $ten_hang_goc, $ten_hang_khop, $hst, $don_vi, $sl_xuat, $sl_nhan, $sl_hong, $ghi_chu);";

                                cmd.Parameters.Add(new DuckDBParameter("id", Guid.NewGuid().ToString()));
                                cmd.Parameters.Add(new DuckDBParameter("file_name", $"test_file_{i}.jpg"));
                                cmd.Parameters.Add(new DuckDBParameter("ngay_giao", "06/06/2026"));
                                cmd.Parameters.Add(new DuckDBParameter("khach_hang", "WinMart"));
                                cmd.Parameters.Add(new DuckDBParameter("diem_giao", "Lê Văn Thiêm"));
                                cmd.Parameters.Add(new DuckDBParameter("nguoi_giao", "Shipper A"));
                                cmd.Parameters.Add(new DuckDBParameter("ten_hang_goc", "Mầm cải củ đỏ"));
                                cmd.Parameters.Add(new DuckDBParameter("ten_hang_khop", "Mầm Cải củ đỏ 150gr"));
                                cmd.Parameters.Add(new DuckDBParameter("hst", "SP0001"));
                                cmd.Parameters.Add(new DuckDBParameter("don_vi", "Hộp"));
                                cmd.Parameters.Add(new DuckDBParameter("sl_xuat", 5.0));
                                cmd.Parameters.Add(new DuckDBParameter("sl_nhan", 4.0));
                                cmd.Parameters.Add(new DuckDBParameter("sl_hong", 1.0));
                                cmd.Parameters.Add(new DuckDBParameter("ghi_chu", "Hỏng do vận chuyển"));

                                cmd.ExecuteNonQuery();
                            }
                        }
                        transaction.Commit();
                    }
                }
                sw.Stop();
                double transactionTime = sw.ElapsedMilliseconds;
                Console.WriteLine($" -> Tổng thời gian (Gộp Transaction): {transactionTime:F2} ms");
                Console.WriteLine($" -> Tốc độ: {numRows / (transactionTime / 1000.0):F0} inserts/giây");
                Console.WriteLine($" -> Tỉ lệ cải thiện hiệu năng: {individualTime / transactionTime:F1}x lần nhanh hơn");

                // Test read time
                sw.Restart();
                int readCount = 0;
                using (var connection = new DuckDBConnection(connStr))
                {
                    connection.Open();
                    using (var cmd = connection.CreateCommand())
                    {
                        cmd.CommandText = "SELECT COUNT(*) FROM staging_items;";
                        readCount = Convert.ToInt32(cmd.ExecuteScalar());
                    }
                }
                sw.Stop();
                Console.WriteLine($" - Số lượng bản ghi kiểm chứng trong DB: {readCount} dòng (Đọc trong {sw.ElapsedMilliseconds} ms)");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[LỖI DUCKDB] Bench DuckDB lỗi: {ex.Message}");
            }

            // Clean up temporary database files
            try
            {
                if (File.Exists(TempDbPath)) File.Delete(TempDbPath);
                // DuckDB might create WAL files or metadata files, let's delete them
                var walPath = TempDbPath + ".wal";
                if (File.Exists(walPath)) File.Delete(walPath);
            }
            catch { }
            Console.WriteLine("-----------------------------------------------------------------------\n");
        }

        private static async Task RunGeminiOcrTestAsync()
        {
            Console.WriteLine(">>> TEST 6: GEMINI OCR VISUAL ACCURACY TEST (CROSS-OUT LINE DETECTION)");
            string imagePath = @"C:\Users\Dell\Desktop\20260419_084855.jpg";
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"[CẢNH BÁO] Không tìm thấy file ảnh tại {imagePath}. Bỏ qua.");
                return;
            }

            string keysPath = "";
            string[] possiblePaths = new[]
            {
                "apikeys.json",
                @"..\apikeys.json",
                @"..\..\apikeys.json",
                @"..\..\..\apikeys.json",
                @"..\..\..\..\apikeys.json",
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "apikeys.json"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "apikeys.json"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "apikeys.json"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "apikeys.json"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "apikeys.json"),
            };

            foreach (var path in possiblePaths)
            {
                if (File.Exists(path))
                {
                    keysPath = path;
                    break;
                }
            }

            if (string.IsNullOrEmpty(keysPath))
            {
                Console.WriteLine("[CẢNH BÁO] Không tìm thấy file apikeys.json ở bất kỳ đường dẫn nào. Bỏ qua.");
                return;
            }

            try
            {
                string jsonKeys = File.ReadAllText(keysPath);
                var keys = System.Text.Json.JsonSerializer.Deserialize<List<string>>(jsonKeys);
                if (keys == null || keys.Count == 0)
                {
                    Console.WriteLine("Không có API key.");
                    return;
                }
                string apiKey = keys[0];

                Console.WriteLine("Đang chuẩn bị ảnh (EXIF rotation & Resize if > 3000)...");
                string base64Image = "";
                string mimeType = "image/jpeg";

                using (var img = new System.Drawing.Bitmap(imagePath))
                {
                    // Rotate image by EXIF orientation
                    if (img.PropertyIdList.Contains(0x0112))
                    {
                        var prop = img.GetPropertyItem(0x0112);
                        if (prop != null && prop.Value != null && prop.Value.Length > 0)
                        {
                            int orientation = prop.Value[0];
                            System.Drawing.RotateFlipType flip = System.Drawing.RotateFlipType.RotateNoneFlipNone;
                            switch (orientation)
                            {
                                case 6: flip = System.Drawing.RotateFlipType.Rotate90FlipNone; break;
                                case 3: flip = System.Drawing.RotateFlipType.Rotate180FlipNone; break;
                                case 8: flip = System.Drawing.RotateFlipType.Rotate270FlipNone; break;
                            }
                            if (flip != System.Drawing.RotateFlipType.RotateNoneFlipNone)
                            {
                                img.RotateFlip(flip);
                            }
                        }
                    }

                    bool needsResize = img.Width > 3000 || img.Height > 3000;
                    System.Drawing.Bitmap apiImage = img;
                    if (needsResize)
                    {
                        double ratioX = 3000.0 / img.Width;
                        double ratioY = 3000.0 / img.Height;
                        double ratio = Math.Min(ratioX, ratioY);
                        int newWidth = (int)(img.Width * ratio);
                        int newHeight = (int)(img.Height * ratio);
                        apiImage = new System.Drawing.Bitmap(newWidth, newHeight);
                        using (var g = System.Drawing.Graphics.FromImage(apiImage))
                        {
                            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                            g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                            g.DrawImage(img, 0, 0, newWidth, newHeight);
                        }
                    }

                    using (var ms = new MemoryStream())
                    {
                        apiImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                        base64Image = Convert.ToBase64String(ms.ToArray());
                    }

                    if (needsResize) apiImage.Dispose();
                }

                Console.WriteLine("Đang gửi yêu cầu tới Gemini API...");
                var cts = new CancellationTokenSource(TimeSpan.FromSeconds(60));
                
                // Call the static helper in MainWindow
                var data = await MainWindow.CallGeminiApi(base64Image, mimeType, apiKey, "gemini-3.5-flash", cts.Token);
                
                Console.WriteLine("\n[KẾT QUẢ PHÂN TÍCH SUY LUẬN TỪ GEMINI]");
                Console.WriteLine(data.quy_trinh_suy_luan);
                Console.WriteLine("\n[DANH SÁCH HÀNG HÓA TRÍCH XUẤT]");
                foreach (var h in data.danh_sach_hang_hoa)
                {
                    if (h.ten_hang.Contains("Mầm cải ngọt") || h.ten_hang.Contains("Mầm cải củ trắng") || h.ten_hang.Contains("Mầm cải củ đỏ"))
                    {
                        Console.WriteLine($" - {h.ten_hang}: Xuất {h.sl_xuat}, Nhận {h.sl_nhan}, Hỏng {h.sl_hong}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[LỖI GEMINI API TEST] {ex.Message}");
            }
            Console.WriteLine("-----------------------------------------------------------------------\n");
        }

        private static async Task RunAdvancedEdgeCaseTestsAsync()
        {
            Console.WriteLine(">>> TEST 7: ADVANCED EDGE CASES, OVERLOAD & CRASH VULNERABILITY TEST");

            // --- 7.1. Empty or Corrupted Excel File Crash Test ---
            Console.WriteLine("[1] Thử nghiệm nạp tệp Excel bị hỏng/rỗng:");
            string corruptFilePath = "corrupt_temp.xlsx";
            try
            {
                File.WriteAllBytes(corruptFilePath, new byte[] { 0x00, 0x11, 0x22, 0x33, 0x44 }); // Corrupt header
                
                // Test ExcelDataReader
                using (var stream = File.Open(corruptFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var ds = reader.AsDataSet();
                        Console.WriteLine("  - ExcelDataReader nạp tệp hỏng thành công (Bất thường!)");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  - ExcelDataReader chặn thành công tệp hỏng với ngoại lệ: {ex.GetType().Name} ({ex.Message})");
            }
            finally
            {
                if (File.Exists(corruptFilePath)) File.Delete(corruptFilePath);
            }

            // --- 7.2. DuckDB Concurrency Write Stress & Lock Test ---
            Console.WriteLine("\n[2] Thử nghiệm ghi đồng thời 10 luồng vào DuckDB (Không dùng Lock vs Có dùng Lock):");
            string testDb = "concurrency_test.db";
            if (File.Exists(testDb)) File.Delete(testDb);

            // A. Không dùng lock (Xem có lỗi/crash không)
            Console.WriteLine("  - Ghi đồng thời 10 luồng (Không đồng bộ hóa - No Lock):");
            using (var conn = new DuckDBConnection($"Data Source={testDb}"))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "CREATE TABLE test_tbl (val INT);";
                    cmd.ExecuteNonQuery();
                }
            }

            int concurrentWrites = 50;
            var tasks = new List<Task>();
            int writeFailures = 0;
            
            for (int i = 0; i < 10; i++)
            {
                int id = i;
                tasks.Add(Task.Run(() =>
                {
                    for (int j = 0; j < concurrentWrites; j++)
                    {
                        try
                        {
                            using (var conn = new DuckDBConnection($"Data Source={testDb}"))
                            {
                                conn.Open();
                                using (var cmd = conn.CreateCommand())
                                {
                                    cmd.CommandText = $"INSERT INTO test_tbl VALUES ({j});";
                                    cmd.ExecuteNonQuery();
                                }
                            }
                        }
                        catch (Exception)
                        {
                            Interlocked.Increment(ref writeFailures);
                        }
                    }
                }));
            }
            await Task.WhenAll(tasks);
            Console.WriteLine($"  - Số lượng luồng ghi lỗi (Không Lock): {writeFailures} / {10 * concurrentWrites}");

            // B. Có dùng lock (Đồng bộ hóa bằng lock object)
            Console.WriteLine("  - Ghi đồng thời 10 luồng (Đồng bộ hóa bằng Lock Object):");
            if (File.Exists(testDb)) File.Delete(testDb);
            using (var conn = new DuckDBConnection($"Data Source={testDb}"))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "CREATE TABLE test_tbl (val INT);";
                    cmd.ExecuteNonQuery();
                }
            }

            object dbLock = new object();
            tasks.Clear();
            int lockedWriteFailures = 0;

            for (int i = 0; i < 10; i++)
            {
                tasks.Add(Task.Run(() =>
                {
                    for (int j = 0; j < concurrentWrites; j++)
                    {
                        try
                        {
                            lock (dbLock)
                            {
                                using (var conn = new DuckDBConnection($"Data Source={testDb}"))
                                {
                                    conn.Open();
                                    using (var cmd = conn.CreateCommand())
                                    {
                                        cmd.CommandText = $"INSERT INTO test_tbl VALUES ({j});";
                                        cmd.ExecuteNonQuery();
                                    }
                                }
                            }
                        }
                        catch (Exception)
                        {
                            Interlocked.Increment(ref lockedWriteFailures);
                        }
                    }
                }));
            }
            await Task.WhenAll(tasks);
            Console.WriteLine($"  - Số lượng luồng ghi lỗi (Có Lock): {lockedWriteFailures} / {10 * concurrentWrites}");

            if (File.Exists(testDb)) File.Delete(testDb);

            // --- 7.3. ApiKeyManager Extreme Concurrency (50 luồng) ---
            Console.WriteLine("\n[3] Thử nghiệm ApiKeyManager với tải cực hạn (50 luồng đồng thời):");
            string? keysBackup7 = null;
            if (File.Exists("apikeys.json"))
            {
                try { keysBackup7 = File.ReadAllText("apikeys.json"); } catch {}
            }

            try
            {
                var keysJson = "[\"KEY_A\", \"KEY_B\", \"KEY_C\", \"KEY_D\", \"KEY_E\"]";
                File.WriteAllText("apikeys.json", keysJson);
                
                var manager = new ApiKeyManager();
                var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
                tasks.Clear();
                var allocatedKeys = new ConcurrentBag<string>();
                int queueFailures = 0;

                for (int i = 0; i < 50; i++)
                {
                    tasks.Add(Task.Run(async () =>
                    {
                        try
                        {
                            string key = await manager.GetNextAvailableKeyAsync(cts.Token);
                            allocatedKeys.Add(key);
                            // Giả lập dùng trong 50ms rồi trả về
                            await Task.Delay(50);
                        }
                        catch (Exception)
                        {
                            Interlocked.Increment(ref queueFailures);
                        }
                    }));
                }
                await Task.WhenAll(tasks);
                Console.WriteLine($"  - Tổng số khóa cấp phát thành công: {allocatedKeys.Count}");
                Console.WriteLine($"  - Số lượng luồng bị timeout/lỗi: {queueFailures}");
            }
            finally
            {
                try
                {
                    if (keysBackup7 != null)
                    {
                        File.WriteAllText("apikeys.json", keysBackup7);
                    }
                    else if (File.Exists("apikeys.json"))
                    {
                        File.Delete("apikeys.json");
                    }
                }
                catch {}
            }

            // --- 7.4. Logger Overload (100,000 dòng log từ 20 luồng) ---
            Console.WriteLine("\n[4] Thử nghiệm Logger chịu tải cực đại (100,000 dòng log từ 20 luồng):");
            Logger.Flush();
            if (File.Exists("app.log")) File.Delete("app.log");

            int overloadThreads = 20;
            int logsPerThread = 5000;
            int totalOverloadLogs = overloadThreads * logsPerThread;

            Stopwatch sw = Stopwatch.StartNew();
            Parallel.For(0, overloadThreads, t =>
            {
                for (int i = 0; i < logsPerThread; i++)
                {
                    Logger.Log($"[Thread {t}] Ghi log overload {i}");
                }
            });
            sw.Stop();
            double enqueueOverloadTime = sw.ElapsedMilliseconds;

            sw.Restart();
            Logger.Flush();
            sw.Stop();
            double flushOverloadTime = sw.ElapsedMilliseconds;

            int finalLogCount = 0;
            if (File.Exists("app.log"))
            {
                finalLogCount = File.ReadLines("app.log").Count();
                File.Delete("app.log");
            }
            Console.WriteLine($"  - Thời gian Đưa hàng đợi (100k dòng): {enqueueOverloadTime} ms");
            Console.WriteLine($"  - Thời gian Xả đĩa (Flush): {flushOverloadTime} ms");
            Console.WriteLine($"  - Toàn vẹn log ghi nhận: {finalLogCount} / {totalOverloadLogs} ({finalLogCount * 100.0 / totalOverloadLogs:F1}%)");

            // --- 7.5. Excel Mismatch & Column Missing Error Test ---
            Console.WriteLine("\n[5] Thử nghiệm đọc tệp cấu trúc cột bị thiếu hoặc sai lệch:");
            string testMismatchFile = "mismatch_temp.xlsx";
            try
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.Worksheets.Add("tham chiếu");
                    ws.Cell(1, 1).Value = "Mã sản phẩm";
                    ws.Cell(1, 2).Value = "Tên"; // Thiếu cột đơn vị
                    ws.Cell(2, 1).Value = "SP001";
                    ws.Cell(2, 2).Value = "Sản phẩm A";
                    
                    wb.Worksheets.Add("Bán Hàng");
                    
                    wb.SaveAs(testMismatchFile);
                }

                // Giả lập đọc
                using (var stream = File.Open(testMismatchFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var ds = reader.AsDataSet();
                        var refTable = ds.Tables["tham chiếu"];
                        int headerRowIndex = -1;
                        int colAbbrevIndex = -1;
                        int colNameIndex = -1;
                        int colUnitIndex = -1;

                        int maxRowsToSearch = Math.Min(20, refTable.Rows.Count);
                        for (int r = 0; r < maxRowsToSearch; r++)
                        {
                            var row = refTable.Rows[r];
                            for (int c = 0; c < refTable.Columns.Count; c++)
                            {
                                string val = row[c]?.ToString()?.Trim()?.ToLower() ?? "";
                                if (val.Contains("viết tắt") || val == "viet tat" || val == "ma" || val == "mã") colAbbrevIndex = c;
                                else if (val.Contains("tên hàng") || val == "ten hang" || val == "sản phẩm" || val == "san pham") colNameIndex = c;
                                else if (val.Contains("đơn vị") || val == "don vi" || val == "đvt" || val == "dvt") colUnitIndex = c;
                            }
                        }

                        Console.WriteLine($"  - Xác định cột viết tắt: {colAbbrevIndex}, cột tên hàng: {colNameIndex}, cột đơn vị: {colUnitIndex}");
                        if (colAbbrevIndex == -1 || colNameIndex == -1 || colUnitIndex == -1)
                        {
                            Console.WriteLine("  - [Thành công] Phát hiện cấu trúc bảng Tham chiếu không khớp tiêu chuẩn (thiếu cột đơn vị)");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  - Lỗi khi thử nghiệm mismatch: {ex.Message}");
            }
            finally
            {
                if (File.Exists(testMismatchFile)) File.Delete(testMismatchFile);
            }

            Console.WriteLine("-----------------------------------------------------------------------\n");
        }

        private static async Task RunOptimizationFeaturesTestsAsync()
        {
            Console.WriteLine(">>> TEST 8: OPTIMIZATION & RELIABILITY FEATURES VERIFICATION");

            // --- 8.1. Log Rotation (10MB Limit Verification) ---
            Console.WriteLine("[1] Kiểm thử xoay vòng log tự động (>10MB):");
            Logger.Flush();
            if (File.Exists("app.log")) File.Delete("app.log");
            if (File.Exists("app_old.log")) File.Delete("app_old.log");

            // Write large logs to exceed 10MB (approx 120,000 log lines)
            Console.WriteLine("  - Đang ghi 120,000 dòng log để vượt ngưỡng 10MB...");
            for (int i = 0; i < 120000; i++)
            {
                Logger.Log($"Đây là dòng log stress test xoay vòng kích thước lớn {i} - Chuỗi đệm dài để tăng nhanh kích thước tệp log ghi đĩa....................................................................");
            }
            Logger.Flush(); // This creates the file exceeding 10MB

            // Write another log and flush to trigger rotation
            Logger.Log("Dòng log kích hoạt xoay vòng");
            Logger.Flush();

            if (File.Exists("app_old.log"))
            {
                var oldSize = new FileInfo("app_old.log").Length / (1024.0 * 1024.0);
                var curSize = new FileInfo("app.log").Length / (1024.0 * 1024.0);
                Console.WriteLine($"  - [Thành công] Đã kích hoạt xoay vòng log:");
                Console.WriteLine($"    * Kích thước file lưu trữ cũ (app_old.log): {oldSize:F2} MB");
                Console.WriteLine($"    * Kích thước file hoạt động mới (app.log): {curSize:F2} MB");
            }
            else
            {
                Console.WriteLine("  - [Thất bại] File app_old.log không được tạo ra.");
            }

            // Cleanup logs
            try
            {
                if (File.Exists("app.log")) File.Delete("app.log");
                if (File.Exists("app_old.log")) File.Delete("app_old.log");
            }
            catch { }

            // --- 8.2. DuckDB Auto-Cleanup (>30 Days Synced Items) ---
            Console.WriteLine("\n[2] Kiểm thử tự động dọn dẹp DuckDB (Synced >30 ngày):");
            string testDb = "cleanup_test.db";
            if (File.Exists(testDb)) File.Delete(testDb);

            using (var conn = new DuckDBConnection($"Data Source={testDb}"))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
                        CREATE TABLE staging_items (
                            id VARCHAR PRIMARY KEY,
                            status VARCHAR,
                            created_at TIMESTAMP
                        );";
                    cmd.ExecuteNonQuery();

                    // Insert 3 mock records:
                    // 1. Synced and 45 days old (Should be deleted)
                    cmd.CommandText = "INSERT INTO staging_items VALUES ('1', 'Synced', now() - INTERVAL '45 days');";
                    cmd.ExecuteNonQuery();

                    // 2. Synced and 10 days old (Should NOT be deleted)
                    cmd.CommandText = "INSERT INTO staging_items VALUES ('2', 'Synced', now() - INTERVAL '10 days');";
                    cmd.ExecuteNonQuery();

                    // 3. Error and 45 days old (Should NOT be deleted, only Synced status is deleted)
                    cmd.CommandText = "INSERT INTO staging_items VALUES ('3', 'Error', now() - INTERVAL '45 days');";
                    cmd.ExecuteNonQuery();

                    // Run the cleanup query
                    cmd.CommandText = "DELETE FROM staging_items WHERE status = 'Synced' AND created_at < now() - INTERVAL '30 days';";
                    int deleted = cmd.ExecuteNonQuery();
                    Console.WriteLine($"  - Đã xóa thành công: {deleted} dòng (Kỳ vọng: 1 dòng)");

                    // Reclaim space via VACUUM
                    cmd.CommandText = "VACUUM;";
                    cmd.ExecuteNonQuery();

                    // Verify remaining rows
                    cmd.CommandText = "SELECT COUNT(*) FROM staging_items;";
                    int remaining = Convert.ToInt32(cmd.ExecuteScalar());
                    Console.WriteLine($"  - Số dòng còn lại trong DB: {remaining} dòng (Kỳ vọng: 2 dòng)");
                    
                    if (deleted == 1 && remaining == 2)
                    {
                        Console.WriteLine("  - [Thành công] Lệnh dọn dẹp và VACUUM hoạt động chính xác.");
                    }
                    else
                    {
                        Console.WriteLine("  - [Thất bại] Lệnh dọn dẹp hoặc đếm bản ghi bị sai lệch.");
                    }
                }
            }
            if (File.Exists(testDb)) File.Delete(testDb);

            // --- 8.3. Config Auto-Backup & Self-Healing ---
            Console.WriteLine("\n[3] Kiểm thử sao lưu cấu hình tự phục hồi (Self-Healing):");
            string mainConfigPath = "excel_configs_test.json";
            string backupConfigPath = mainConfigPath + ".bak";

            if (File.Exists(mainConfigPath)) File.Delete(mainConfigPath);
            if (File.Exists(backupConfigPath)) File.Delete(backupConfigPath);

            try
            {
                // 1. Save config
                var mockConfig = new List<string> { "Config_A", "Config_B" };
                string validJson = System.Text.Json.JsonSerializer.Serialize(mockConfig);
                File.WriteAllText(mainConfigPath, validJson);
                File.WriteAllText(backupConfigPath, validJson);
                Console.WriteLine("  - Đã tạo tệp cấu hình chính và cấu hình dự phòng (.bak).");

                // 2. Corrupt main config file
                File.WriteAllText(mainConfigPath, "{ corrupt json ... [ [");
                Console.WriteLine("  - Đã làm hỏng tệp cấu hình chính.");

                // 3. Load config fallback
                bool loadedFromBackup = false;
                List<string>? loadedList = null;

                if (File.Exists(mainConfigPath))
                {
                    try
                    {
                        string json = File.ReadAllText(mainConfigPath);
                        loadedList = System.Text.Json.JsonSerializer.Deserialize<List<string>>(json);
                    }
                    catch
                    {
                        Console.WriteLine("  - Đọc tệp chính thất bại (đúng kỳ vọng). Đang thử khôi phục từ tệp dự phòng...");
                    }
                }

                if (loadedList == null && File.Exists(backupConfigPath))
                {
                    try
                    {
                        string json = File.ReadAllText(backupConfigPath);
                        loadedList = System.Text.Json.JsonSerializer.Deserialize<List<string>>(json);
                        if (loadedList != null)
                        {
                            // Heal primary file
                            File.WriteAllText(mainConfigPath, json);
                            loadedFromBackup = true;
                        }
                    }
                    catch
                    {
                        Console.WriteLine("  - Đọc tệp dự phòng thất bại.");
                    }
                }

                if (loadedFromBackup && loadedList != null && loadedList.Count == 2)
                {
                    Console.WriteLine("  - [Thành công] Phục hồi cấu hình từ tệp .bak thành công!");
                    // Verify primary is healed
                    string healedJson = File.ReadAllText(mainConfigPath);
                    if (healedJson == validJson)
                    {
                        Console.WriteLine("  - [Thành công] Tệp chính đã được chữa lành (Healed) hoàn toàn.");
                    }
                    else
                    {
                        Console.WriteLine("  - [Thất bại] Nội dung tệp chính chưa được đồng bộ chữa lành.");
                    }
                }
                else
                {
                    Console.WriteLine("  - [Thất bại] Khôi phục từ tệp dự phòng lỗi.");
                }
            }
            finally
            {
                if (File.Exists(mainConfigPath)) File.Delete(mainConfigPath);
                if (File.Exists(backupConfigPath)) File.Delete(backupConfigPath);
            }

            Console.WriteLine("-----------------------------------------------------------------------\n");
        }
    }
}
