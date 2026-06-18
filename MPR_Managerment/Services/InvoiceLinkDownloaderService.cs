using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace MPR_Managerment.Services
{
    public class LinkDownloadResult
    {
        [JsonPropertyName("status")] public string Status { get; set; } = "";
        [JsonPropertyName("subject")] public string? Subject { get; set; }
        [JsonPropertyName("sender")] public string? Sender { get; set; }
        [JsonPropertyName("provider")] public string? Provider { get; set; }
        [JsonPropertyName("lookup_code")] public string? LookupCode { get; set; }
        [JsonPropertyName("file")] public string? File { get; set; }
        [JsonPropertyName("reason")] public string? Reason { get; set; }
    }

    public class LinkDownloadSummary
    {
        [JsonPropertyName("total")] public int Total { get; set; }
        [JsonPropertyName("downloaded")] public int Downloaded { get; set; }
        [JsonPropertyName("pending_manual")] public int PendingManual { get; set; }
    }

    public class LinkDownloadResponse
    {
        [JsonPropertyName("success")] public bool Success { get; set; }
        [JsonPropertyName("error")] public string? Error { get; set; }
        [JsonPropertyName("results")] public List<LinkDownloadResult> Results { get; set; } = new();
        [JsonPropertyName("summary")] public LinkDownloadSummary Summary { get; set; } = new();
    }

    public static class InvoiceLinkDownloaderService
    {
        private static readonly string ScriptPath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "Scripts", "invoice_link_downloader.py");

        public static async Task<LinkDownloadResponse> DownloadFromLinksAsync(
            string saveDir,
            int daysBack = 30,
            IProgress<string>? progress = null,
            string? forwardTo = null)
        {
            if (!File.Exists(ScriptPath))
                return new LinkDownloadResponse { Success = false, Error = $"Không tìm thấy script: {ScriptPath}" };

            progress?.Report("Đang quét email chứa link hóa đơn...");

            string fwdArg = string.IsNullOrWhiteSpace(forwardTo ?? OutlookMonitorService.ForwardTo)
                ? ""
                : $" --forward-to \"{(forwardTo ?? OutlookMonitorService.ForwardTo)}\"";

            var psi = new ProcessStartInfo
            {
                FileName = "python",
                Arguments = $"\"{ScriptPath}\" --save-dir \"{saveDir}\" --days-back {daysBack}{fwdArg}",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = System.Text.Encoding.UTF8
            };

            try
            {
                using var proc = Process.Start(psi)
                    ?? throw new Exception("Không khởi động được Python.");

                string stdout = await proc.StandardOutput.ReadToEndAsync();
                string stderr = await proc.StandardError.ReadToEndAsync();
                await proc.WaitForExitAsync();

                if (string.IsNullOrWhiteSpace(stdout))
                    return new LinkDownloadResponse { Success = false, Error = $"Lỗi script: {stderr}" };

                return JsonSerializer.Deserialize<LinkDownloadResponse>(stdout)
                    ?? new LinkDownloadResponse { Success = false, Error = "Không parse được kết quả." };
            }
            catch (Exception ex)
            {
                return new LinkDownloadResponse { Success = false, Error = ex.Message };
            }
        }
    }
}
