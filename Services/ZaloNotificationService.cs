using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;

namespace MPR_Managerment.Services
{
    // =====================================================
    //  MODEL: Cấu hình hàng đợi Zalo (lưu vào file zalo_queue_settings.json)
    // =====================================================
    public class ZaloQueueSettings
    {
        /// <summary>Độ trễ tối thiểu giữa 2 lần gửi (ms), tránh rate-limit</summary>
        public int DelayBetweenMessagesMs { get; set; } = 2000;
        /// <summary>Số lần thử lại khi gặp lỗi trước khi bỏ qua</summary>
        public int MaxRetries { get; set; } = 5;
        /// <summary>Thời gian chờ giữa các lần retry (ms)</summary>
        public int RetryDelayMs { get; set; } = 5000;
    }

    // =====================================================
    //  MODEL: Một thông báo Zalo cần gửi
    // =====================================================
    public class ZaloNotificationItem
    {
        public string GroupName { get; set; } = "";   // Tên nhóm Zalo (dùng cho ZaloHelper)
        public string GroupId { get; set; } = "";   // ID nhóm Zalo (backup)
        public string Message { get; set; } = "";   // Nội dung tin nhắn
        public DateTime QueuedAt { get; set; } = DateTime.Now;
        public string ProjectCode { get; set; } = "";  // Để log / debug
        public int RetryCount { get; set; } = 0;
        public int LogId { get; set; } = 0; // ID của entry trong Notification_Logs
    }

    // =====================================================
    //  SERVICE: Singleton — quản lý hàng đợi gửi Zalo
    // =====================================================
    /// <summary>
    /// Singleton service xử lý hàng đợi thông báo Zalo.
    /// Tất cả thông báo được ghi vào queue và xử lý tuần tự,
    /// đảm bảo không bỏ sót tin nhắn nào kể cả khi nhiều hành động
    /// xảy ra đồng thời hoặc khi đang trong thời gian chờ rate-limit.
    ///
    /// Queue được persist ra file zalo_queue.json để không mất tin nhắn
    /// khi ứng dụng đóng bất ngờ hoặc restart.
    ///
    /// Vì ZaloHelper.SendToGroupAsync dùng WebView2 (chỉ chạy trên UI thread),
    /// service sử dụng SynchronizationContext để đảm bảo mọi gửi tin
    /// đều thực hiện trên UI thread.
    /// </summary>
    public sealed class ZaloNotificationService : IDisposable
    {
        // ── Singleton ──────────────────────────────────────────────────
        private static readonly Lazy<ZaloNotificationService> _instance =
            new(() => new ZaloNotificationService());
        public static ZaloNotificationService Instance => _instance.Value;

        private readonly NotificationLogService _logService = new();

        // ── Hàng đợi bộ nhớ ────────────────────────────────────────────
        private readonly ConcurrentQueue<ZaloNotificationItem> _queue = new();

        // ── UI SynchronizationContext ──────────────────────────────────
        private readonly SynchronizationContext _uiContext;
        private readonly System.Windows.Forms.Timer _processTimer;
        private Task _currentProcessingTask = Task.CompletedTask; // Track async operation
        private DateTime _lastSendTime = DateTime.MinValue;

        // ── Settings ───────────────────────────────────────────────────
        private static readonly string _settingsFile = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "zalo_queue_settings.json");
        private static readonly JsonSerializerOptions _jsonOpt = new() { WriteIndented = true };

        // ── Queue persistence file ─────────────────────────────────────
        private static readonly string _queueFile = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "zalo_queue.json");

        // ── Log ────────────────────────────────────────────────────────
        private static readonly string _logFile = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "zalo_notifications.log");

        // ── Private constructor ────────────────────────────────────────
        private ZaloNotificationService()
        {
            // Capture UI thread SynchronizationContext
            _uiContext = SynchronizationContext.Current;
            if (_uiContext == null)
            {
                // Nếu không có UI context (service được tạo trên background thread),
                // chúng ta cần tìm UI thread. Tạm thời log cảnh báo.
                Log($"[WARN] ZaloNotificationService created on non‑UI thread. " +
                    $"Zalo notifications may fail. Ensure service is initialized from UI thread.");
            }

            // Load pending queue from file (from previous session)
            LoadPendingQueue();

            // Timer chạy trên UI thread, kiểm tra queue mỗi 1 giây
            _processTimer = new System.Windows.Forms.Timer();
            _processTimer.Interval = 1000; // 1 second check interval
            _processTimer.Tick += (s, e) => ProcessNextItem();
            _processTimer.Start();

            Log($"[INIT] ZaloNotificationService started. Pending={_queue.Count} UIContext={_uiContext != null}");
        }

        // ==============================================================
        //  PUBLIC API
        // ==============================================================

        /// <summary>
        /// Đặt một thông báo vào hàng đợi. Thread-safe, non-blocking.
        /// Sử dụng groupName (tên nhóm Zalo) để gửi qua ZaloHelper.
        /// </summary>
        public void Enqueue(string groupName, string message, string projectCode = "")
        {
            if (string.IsNullOrWhiteSpace(groupName) || string.IsNullOrWhiteSpace(message))
                return;

            // 1. Ghi lại lịch sử thông báo với trạng thái Pending trước để lấy LogId
            int logId = _logService.AddLog(new NotificationLog
            {
                Sent_At = DateTime.Now,
                Sent_By = AppSession.CurrentUser?.Full_Name ?? "System",
                Recipient = groupName,
                Type = "Zalo",
                Content = message,
                Status = "Pending",
                Project_Code = projectCode
            });

            // 2. Tạo item và đưa vào queue
            var item = new ZaloNotificationItem
            {
                GroupName = groupName,
                Message = message,
                ProjectCode = projectCode,
                QueuedAt = DateTime.Now,
                LogId = logId
            };

            _queue.Enqueue(item);
            SaveQueue(); // persist to file
            Log($"[QUEUED] Project={projectCode} Group={groupName} QueueSize={_queue.Count} LogId={logId}");
        }

        /// <summary>
        /// Số lượng thông báo đang chờ trong hàng đợi.
        /// </summary>
        public int PendingCount => _queue.Count;

        // ==============================================================
        //  SETTINGS
        // ==============================================================
        public static ZaloQueueSettings LoadQueueSettings()
        {
            try
            {
                if (File.Exists(_settingsFile))
                {
                    string json = File.ReadAllText(_settingsFile);
                    return JsonSerializer.Deserialize<ZaloQueueSettings>(json) ?? new ZaloQueueSettings();
                }
            }
            catch { }
            return new ZaloQueueSettings();
        }

        public static void SaveQueueSettings(ZaloQueueSettings settings)
        {
            string json = JsonSerializer.Serialize(settings, _jsonOpt);
            File.WriteAllText(_settingsFile, json);
        }

        // ── Số lần retry tổng thể tối đa trước khi bỏ qua item ──────────
        private const int MAX_TOTAL_RETRIES = 20;

        // ==============================================================
        //  PROCESS QUEUE — chạy trên UI thread
        // ==============================================================
        private void ProcessNextItem()
        {
            // Nếu đang xử lý hoặc queue trống, return
            if (!_currentProcessingTask.IsCompleted || _queue.IsEmpty)
                return;

            var queueSettings = LoadQueueSettings();
            int delayMs = Math.Max(queueSettings.DelayBetweenMessagesMs, 2000); // Tối thiểu 2s

            // 1. Kiểm tra khoảng cách thời gian giữa 2 lần gửi
            TimeSpan elapsedSinceLast = DateTime.Now - _lastSendTime;
            if (elapsedSinceLast.TotalMilliseconds < delayMs)
            {
                return;
            }

            // 2. Khởi động xử lý item tiếp theo
            if (_uiContext != null && _uiContext != SynchronizationContext.Current)
            {
                try
                {
                    _currentProcessingTask = Task.Run(async () =>
                    {
                        _uiContext.Post(async _ =>
                        {
                            await ProcessNextItemCore();
                        }, null);
                    });
                }
                catch (Exception ex)
                {
                    Log($"[ERROR] Failed to start processing: {ex.Message}");
                }
                return;
            }

            // Đang ở UI thread, xử lý trực tiếp
            _currentProcessingTask = ProcessNextItemCore();
        }

        private async Task ProcessNextItemCore()
        {
            try
            {
                var queueSettings = LoadQueueSettings();

                if (_queue.TryDequeue(out var item))
                {
                    bool sent = false;
                    int attempts = 0;

                    while (!sent && attempts < queueSettings.MaxRetries)
                    {
                        attempts++;
                        try
                        {
                            var zaloCfg = ZaloHelper.LoadSettings();
                            if (!zaloCfg.Enabled)
                            {
                                Log($"[SKIP] Zalo chưa bật — Project={item.ProjectCode}");
                                sent = true;
                                _logService.UpdateLogStatus(item.LogId, "Failed", "Zalo integration is disabled.");
                                break;
                            }

                            Log($"[SENDING] Project={item.ProjectCode} Group={item.GroupName} " +
                                $"attempt={attempts}/{queueSettings.MaxRetries} totalRetried={item.RetryCount}");

                            // Gửi tin nhắn — ZaloHelper sẽ đảm bảo chạy trên UI thread
                            var (ok, err) = await ZaloHelper.SendToGroupAsync(
                                zaloCfg, item.GroupName, item.Message);

                            if (ok)
                            {
                                sent = true;
                                _lastSendTime = DateTime.Now;
                                Log($"[SENT] Project={item.ProjectCode} Group={item.GroupName} " +
                                    $"attempt={attempts} totalRetried={item.RetryCount}");

                                _logService.UpdateLogStatus(item.LogId, "Success");
                            }
                            else
                            {
                                Log($"[FAILED_ATTEMPT {attempts}/{queueSettings.MaxRetries}] " +
                                    $"Project={item.ProjectCode} Err={err}");

                                if (attempts < queueSettings.MaxRetries)
                                {
                                    await Task.Delay(queueSettings.RetryDelayMs);
                                }
                                else
                                {
                                    _RequeueOrDrop(item, $"send failed: {err}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Log($"[ERROR] Project={item.ProjectCode} attempt={attempts} {ex.Message}");
                            if (attempts < queueSettings.MaxRetries)
                            {
                                await Task.Delay(queueSettings.RetryDelayMs);
                            }
                            else
                            {
                                _RequeueOrDrop(item, $"exception: {ex.Message}");
                            }
                        }
                    }

                    // Persist remaining items after each processed item
                    SaveQueue();
                }
            }
            catch (Exception ex)
            {
                Log($"[CRITICAL] ProcessNextItemCore error: {ex.Message}");
            }
        }

        /// <summary>
        /// Đưa item lại vào queue sau khi hết lần retry trong một chu kỳ.
        /// - Tăng RetryCount để track tổng số lần đã thử
        /// - Nếu vượt MAX_TOTAL_RETRIES thì bỏ qua (tránh loop vô tận)
        /// - Cập nhật _lastSendTime để áp dụng cooldown trước lần retry tiếp theo
        /// </summary>
        private void _RequeueOrDrop(ZaloNotificationItem item, string reason)
        {
            item.RetryCount++;

            // Log the failure/drop status to DB
            string status = (item.RetryCount >= MAX_TOTAL_RETRIES) ? "Dropped" : "Pending"; // Vẫn là Pending nếu còn thử lại
            _logService.UpdateLogStatus(item.LogId, status, reason);

            if (item.RetryCount >= MAX_TOTAL_RETRIES)
            {
                Log($"[DROPPED] Project={item.ProjectCode} Group={item.GroupName} " +
                    $"totalRetried={item.RetryCount} — vượt giới hạn {MAX_TOTAL_RETRIES} lần, bỏ qua.");
                return;
            }

            // Cập nhật _lastSendTime khi re-queue để áp dụng cooldown
            _lastSendTime = DateTime.Now;

            _queue.Enqueue(item);
            Log($"[REQUEUE] Project={item.ProjectCode} Group={item.GroupName} " +
                $"totalRetried={item.RetryCount}/{MAX_TOTAL_RETRIES} reason={reason} " +
                $"— sẽ thử lại sau chu kỳ delay");
        }

        // ==============================================================
        //  QUEUE PERSISTENCE
        // ==============================================================
        private void SaveQueue()
        {
            try
            {
                var pending = _queue.ToArray();
                string json = JsonSerializer.Serialize(pending, _jsonOpt);
                File.WriteAllText(_queueFile, json, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Log($"[SAVE_ERROR] {ex.Message}");
            }
        }

        private void LoadPendingQueue()
        {
            try
            {
                if (!File.Exists(_queueFile)) return;

                string json = File.ReadAllText(_queueFile, Encoding.UTF8);
                if (string.IsNullOrWhiteSpace(json)) return;

                var pending = JsonSerializer.Deserialize<List<ZaloNotificationItem>>(json);
                if (pending != null && pending.Count > 0)
                {
                    foreach (var item in pending)
                        _queue.Enqueue(item);

                    Log($"[LOADED] Đã nạp {pending.Count} thông báo chờ từ file zalo_queue.json");
                }
            }
            catch (Exception ex)
            {
                Log($"[LOAD_ERROR] {ex.Message}");
            }
        }

        // ==============================================================
        //  LOG
        // ==============================================================
        private static void Log(string message)
        {
            try
            {
                string line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}{Environment.NewLine}";
                File.AppendAllText(_logFile, line, Encoding.UTF8);
            }
            catch { /* không để lỗi log làm crash app */ }
        }

        // ==============================================================
        //  DISPOSE / SHUTDOWN
        // ==============================================================
        public void Dispose()
        {
            _processTimer.Stop();
            _processTimer.Dispose();

            // Flush log cuối
            if (!_queue.IsEmpty)
                Log($"[SHUTDOWN] Còn {_queue.Count} thông báo chưa gửi trong queue.");

            // Persist remaining items
            SaveQueue();
        }
    }
}