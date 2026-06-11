using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace MPR_Managerment.Services
{
    public class MonitorEvent
    {
        [JsonPropertyName("event")] public string Event { get; set; } = "";
        [JsonPropertyName("time")] public string Time { get; set; } = "";
        [JsonPropertyName("subject")] public string? Subject { get; set; }
        [JsonPropertyName("sender")] public string? Sender { get; set; }
        [JsonPropertyName("files")] public List<string>? Files { get; set; }
        [JsonPropertyName("message")] public string? Message { get; set; }
        [JsonPropertyName("pid")] public int Pid { get; set; }
        [JsonPropertyName("forwarded")] public bool? Forwarded { get; set; }
        [JsonPropertyName("forward_error")] public string? ForwardError { get; set; }
        [JsonPropertyName("marked_read")] public bool? MarkedRead { get; set; }
        [JsonPropertyName("classify_classified")] public int? ClassifyClassified { get; set; }
        [JsonPropertyName("classify_unclassified")] public int? ClassifyUnclassified { get; set; }
        [JsonPropertyName("classify_error")] public string? ClassifyError { get; set; }
    }

    public class SuccessLogEntry
    {
        [JsonPropertyName("time")] public string Time { get; set; } = "";
        [JsonPropertyName("subject")] public string Subject { get; set; } = "";
        [JsonPropertyName("sender")] public string Sender { get; set; } = "";
        [JsonPropertyName("files")] public List<string> Files { get; set; } = new();
        [JsonPropertyName("file_paths")] public List<string> FilePaths { get; set; } = new();
        [JsonPropertyName("forwarded_to")] public List<string> ForwardedTo { get; set; } = new();
        [JsonPropertyName("forwarded")] public bool Forwarded { get; set; }
        [JsonPropertyName("forward_error")] public string? ForwardError { get; set; }
        [JsonPropertyName("marked_read")] public bool MarkedRead { get; set; }
        [JsonPropertyName("classify_classified")] public int ClassifyClassified { get; set; }
        [JsonPropertyName("classify_unclassified")] public int ClassifyUnclassified { get; set; }
        [JsonPropertyName("classify_error")] public string? ClassifyError { get; set; }
    }

    public static class OutlookMonitorService
    {
        private static readonly string ScriptPath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "Scripts", "outlook_invoice_monitor.py");

        private static readonly string PidFile = Path.Combine(
            Path.GetTempPath(), "mpr_outlook_monitor.pid");

        private static readonly string LogFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "MPR_Invoices", "monitor_log.json");

        public static readonly string SuccessLogFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "MPR_Invoices", "invoice_success_log.json");

        private const string TaskName = "MPR_OutlookInvoiceMonitor";
        private const string RegKey = @"SOFTWARE\MPR_Managerment";
        private const string RegValuePersistent = "MonitorPersistent";
        private const string RegValueSaveDir = "MonitorSaveDir";
        private const string RegValueForwardTo = "MonitorForwardTo";

        public static string SaveDir
        {
            get
            {
                try
                {
                    using var key = Registry.CurrentUser.OpenSubKey(RegKey);
                    return key?.GetValue(RegValueSaveDir) as string
                        ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "MPR_Invoices");
                }
                catch { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "MPR_Invoices"); }
            }
            set
            {
                try
                {
                    using var key = Registry.CurrentUser.CreateSubKey(RegKey);
                    key.SetValue(RegValueSaveDir, value);
                }
                catch { }
            }
        }

        /// <summary>Thư mục "Chưa phân loại" — nơi classifier đặt file không tìm được PO.</summary>
        public static string UnclassifiedDir =>
            Path.Combine(SaveDir, "Chua phan loai");

        /// <summary>Đếm số file PDF hiện có trong thư mục Chưa phân loại.</summary>
        public static int CountUnclassified()
        {
            try
            {
                string dir = UnclassifiedDir;
                if (!Directory.Exists(dir)) return 0;
                return Directory.GetFiles(dir, "*.pdf", SearchOption.TopDirectoryOnly).Length;
            }
            catch { return 0; }
        }

        /// <summary>
        /// Danh sách email kế toán để chuyển tiếp, phân cách bằng dấu phẩy.
        /// Lưu vào registry để persist kể cả khi tắt app.
        /// </summary>
        public static string ForwardTo
        {
            get
            {
                try
                {
                    using var key = Registry.CurrentUser.OpenSubKey(RegKey);
                    return key?.GetValue(RegValueForwardTo) as string ?? "";
                }
                catch { return ""; }
            }
            set
            {
                try
                {
                    using var key = Registry.CurrentUser.CreateSubKey(RegKey);
                    key.SetValue(RegValueForwardTo, value ?? "");
                }
                catch { }
            }
        }

        /// <summary>True nếu user đã bật chế độ theo dõi tự động liên tục (kể cả khi tắt app).</summary>
        public static bool IsPersistentEnabled
        {
            get
            {
                try
                {
                    using var key = Registry.CurrentUser.OpenSubKey(RegKey);
                    return key?.GetValue(RegValuePersistent) as string == "1";
                }
                catch { return false; }
            }
            private set
            {
                try
                {
                    using var key = Registry.CurrentUser.CreateSubKey(RegKey);
                    key.SetValue(RegValuePersistent, value ? "1" : "0");
                }
                catch { }
            }
        }

        public static bool IsRunning
        {
            get
            {
                if (!File.Exists(PidFile)) return false;
                try
                {
                    int pid = int.Parse(File.ReadAllText(PidFile).Trim());
                    Process.GetProcessById(pid);
                    return true;
                }
                catch
                {
                    return false;
                }
            }
        }

        public static event Action<MonitorEvent>? OnNewEvent;

        private static Process? _monitorProcess;
        private static CancellationTokenSource? _cts;

        public static async Task<string> StartAsync(string? saveDir = null)
        {
            if (IsRunning) return "Monitor đang chạy.";

            string dir = saveDir ?? SaveDir;
            Directory.CreateDirectory(dir);
            Directory.CreateDirectory(Path.GetDirectoryName(LogFile)!);

            if (!File.Exists(ScriptPath))
                return $"Không tìm thấy script: {ScriptPath}";

            string fwdArg = string.IsNullOrWhiteSpace(ForwardTo) ? "" : $" --forward-to \"{ForwardTo}\"";
            var psi = new ProcessStartInfo
            {
                FileName = "python",
                Arguments = $"\"{ScriptPath}\" --save-dir \"{dir}\" --log-file \"{LogFile}\" --success-log-file \"{SuccessLogFile}\" --pid-file \"{PidFile}\"{fwdArg}",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = System.Text.Encoding.UTF8
            };

            _cts = new CancellationTokenSource();
            _monitorProcess = Process.Start(psi);

            if (_monitorProcess == null) return "Không thể khởi động Python.";

            // C# tự ghi PID ngay — không phụ thuộc Python ghi
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(PidFile)!);
                await File.WriteAllTextAsync(PidFile, _monitorProcess.Id.ToString());
            }
            catch { }

            // Đọc stdout (JSON events) và stderr (lỗi script) liên tục
            _ = Task.Run(async () =>
            {
                try
                {
                    while (!_cts.Token.IsCancellationRequested && !_monitorProcess.HasExited)
                    {
                        string? line = await _monitorProcess.StandardOutput.ReadLineAsync();
                        if (line == null) break;
                        try
                        {
                            var evt = JsonSerializer.Deserialize<MonitorEvent>(line);
                            if (evt != null) OnNewEvent?.Invoke(evt);
                        }
                        catch { }
                    }
                }
                catch { }
            }, _cts.Token);

            _ = Task.Run(async () =>
            {
                try
                {
                    string stderr = await _monitorProcess.StandardError.ReadToEndAsync();
                    if (!string.IsNullOrWhiteSpace(stderr))
                    {
                        string errorLog = Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                            "MPR_Invoices", "monitor_error.log");
                        await File.WriteAllTextAsync(errorLog,
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}]\n{stderr}");
                        // Thông báo lỗi script qua event
                        OnNewEvent?.Invoke(new MonitorEvent
                        {
                            Event = "error",
                            Time = DateTime.Now.ToString("HH:mm:ss"),
                            Message = $"Script lỗi — xem: {errorLog}"
                        });
                    }
                }
                catch { }
            });

            return "Monitor đã khởi động. Đang lắng nghe email mới...";
        }

        public static string Stop()
        {
            _cts?.Cancel();

            if (_monitorProcess != null && !_monitorProcess.HasExited)
            {
                try { _monitorProcess.Kill(entireProcessTree: true); } catch { }
                _monitorProcess = null;
            }

            if (File.Exists(PidFile))
            {
                try
                {
                    int pid = int.Parse(File.ReadAllText(PidFile).Trim());
                    Process.GetProcessById(pid).Kill();
                }
                catch { }
                File.Delete(PidFile);
            }

            return "Monitor đã dừng.";
        }

        public static List<MonitorEvent> ReadLog(int lastN = 50)
        {
            if (!File.Exists(LogFile)) return new();
            try
            {
                var all = JsonSerializer.Deserialize<List<MonitorEvent>>(File.ReadAllText(LogFile))
                    ?? new List<MonitorEvent>();
                int skip = Math.Max(0, all.Count - lastN);
                return all.GetRange(skip, all.Count - skip);
            }
            catch { return new(); }
        }

        /// <summary>
        /// Đọc nhật ký thành công (invoice_success_log.json).
        /// File này KHÔNG bao giờ bị xóa tự động — là lịch sử vĩnh viễn.
        /// </summary>
        public static List<SuccessLogEntry> ReadSuccessLog()
        {
            if (!File.Exists(SuccessLogFile)) return new();
            try
            {
                return JsonSerializer.Deserialize<List<SuccessLogEntry>>(File.ReadAllText(SuccessLogFile))
                    ?? new List<SuccessLogEntry>();
            }
            catch { return new(); }
        }

        // ── Persistent monitoring (Task Scheduler) ────────────────────────────

        /// <summary>
        /// Bật chế độ tự động: đăng ký Task Scheduler chạy khi đăng nhập Windows,
        /// lưu trạng thái vào registry, và khởi động ngay lập tức.
        /// </summary>
        public static async Task<string> EnablePersistentAsync(string saveDir)
        {
            SaveDir = saveDir;
            IsPersistentEnabled = true;
            RegisterScheduledTask(saveDir);
            return await StartAsync(saveDir);
        }

        /// <summary>
        /// Tắt hoàn toàn: dừng process, xóa Task Scheduler, xóa registry flag.
        /// </summary>
        public static string DisablePersistent()
        {
            IsPersistentEnabled = false;
            UnregisterScheduledTask();
            return Stop();
        }

        private static void RegisterScheduledTask(string saveDir)
        {
            try
            {
                // Tìm đường dẫn python thực (pythonw để không hiện cửa sổ console)
                string pythonExe = FindPythonExe();
                string fwdPart = string.IsNullOrWhiteSpace(ForwardTo) ? "" : $" --forward-to \"{ForwardTo}\"";
                string arguments = $"\"{ScriptPath}\" --save-dir \"{saveDir}\" --log-file \"{LogFile}\" --success-log-file \"{SuccessLogFile}\" --pid-file \"{PidFile}\"{fwdPart}";
                string userId = System.Security.Principal.WindowsIdentity.GetCurrent().Name;

                // Dùng XML task definition để tránh vấn đề quoting phức tạp
                string xmlContent = $@"<?xml version=""1.0"" encoding=""UTF-16""?>
<Task version=""1.2"" xmlns=""http://schemas.microsoft.com/windows/2004/02/mit/task"">
  <RegistrationInfo>
    <Description>MPR Outlook Invoice Monitor — tự động theo dõi email hóa đơn</Description>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <UserId>{System.Security.SecurityElement.Escape(userId)}</UserId>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id=""Author"">
      <UserId>{System.Security.SecurityElement.Escape(userId)}</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context=""Author"">
    <Exec>
      <Command>{System.Security.SecurityElement.Escape(pythonExe)}</Command>
      <Arguments>{System.Security.SecurityElement.Escape(arguments)}</Arguments>
      <WorkingDirectory>{System.Security.SecurityElement.Escape(AppDomain.CurrentDomain.BaseDirectory)}</WorkingDirectory>
    </Exec>
  </Actions>
</Task>";

                string xmlFile = Path.Combine(Path.GetTempPath(), "mpr_monitor_task.xml");
                File.WriteAllText(xmlFile, xmlContent, System.Text.Encoding.Unicode);

                var psi = new ProcessStartInfo("schtasks.exe",
                    $"/Create /F /TN \"{TaskName}\" /XML \"{xmlFile}\"")
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                using var p = Process.Start(psi);
                p?.WaitForExit(8000);

                try { File.Delete(xmlFile); } catch { }
            }
            catch { }
        }

        private static string FindPythonExe()
        {
            // Ưu tiên pythonw.exe (không hiện cửa sổ console đen)
            foreach (string candidate in new[] { "pythonw", "python" })
            {
                try
                {
                    var psi = new ProcessStartInfo("where", candidate)
                    {
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardOutput = true
                    };
                    using var p = Process.Start(psi);
                    string? path = p?.StandardOutput.ReadLine()?.Trim();
                    p?.WaitForExit(3000);
                    if (!string.IsNullOrEmpty(path) && File.Exists(path))
                        return path;
                }
                catch { }
            }
            return "pythonw.exe"; // fallback
        }

        private static void UnregisterScheduledTask()
        {
            try
            {
                var psi = new ProcessStartInfo("schtasks.exe", $"/Delete /F /TN \"{TaskName}\"")
                {
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using var p = Process.Start(psi);
                p?.WaitForExit(5000);
            }
            catch { }
        }

        /// <summary>Kiểm tra task scheduler có tồn tại không.</summary>
        public static bool IsScheduledTaskRegistered()
        {
            try
            {
                var psi = new ProcessStartInfo("schtasks.exe", $"/Query /TN \"{TaskName}\"")
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                using var p = Process.Start(psi);
                p?.WaitForExit(3000);
                return p?.ExitCode == 0;
            }
            catch { return false; }
        }
    }
}
