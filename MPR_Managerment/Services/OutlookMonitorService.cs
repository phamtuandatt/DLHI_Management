using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;

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

        public static string SaveDir { get; set; } =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "MPR_Invoices");

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

            var psi = new ProcessStartInfo
            {
                FileName = "python",
                Arguments = $"\"{ScriptPath}\" --save-dir \"{dir}\" --log-file \"{LogFile}\" --pid-file \"{PidFile}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = System.Text.Encoding.UTF8
            };

            _cts = new CancellationTokenSource();
            _monitorProcess = Process.Start(psi);

            if (_monitorProcess == null) return "Không thể khởi động Python.";

            // Đọc stdout từ script liên tục để nhận events
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
    }
}
