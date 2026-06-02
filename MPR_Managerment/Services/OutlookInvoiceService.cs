using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace MPR_Managerment.Services
{
    public class DownloadedFile
    {
        [JsonPropertyName("file")] public string FilePath { get; set; } = "";
        [JsonPropertyName("subject")] public string Subject { get; set; } = "";
        [JsonPropertyName("sender")] public string Sender { get; set; } = "";
        [JsonPropertyName("received")] public string Received { get; set; } = "";
    }

    public class DownloadResult
    {
        [JsonPropertyName("success")] public bool Success { get; set; }
        [JsonPropertyName("error")] public string? Error { get; set; }
        [JsonPropertyName("emails_processed")] public int EmailsProcessed { get; set; }
        [JsonPropertyName("files_downloaded")] public int FilesDownloaded { get; set; }
        [JsonPropertyName("files")] public List<DownloadedFile> Files { get; set; } = new();
        [JsonPropertyName("errors")] public List<string> Errors { get; set; } = new();
    }

    public static class OutlookInvoiceService
    {
        // Đường dẫn script Python — tương đối so với thư mục app
        private static readonly string ScriptPath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "Scripts", "outlook_invoice_downloader.py");

        // Thư mục mặc định lưu file (có thể override từ settings)
        public static string DefaultSaveDir { get; set; } =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "MPR_Invoices");

        /// <summary>
        /// Chạy script Python để tải file đính kèm từ Outlook.
        /// Trả về DownloadResult với danh sách file đã tải.
        /// </summary>
        public static async Task<DownloadResult> DownloadInvoicesAsync(
            string? saveDir = null,
            int daysBack = 30,
            IProgress<string>? progress = null)
        {
            string dir = saveDir ?? DefaultSaveDir;
            Directory.CreateDirectory(dir);

            if (!File.Exists(ScriptPath))
                return new DownloadResult
                {
                    Success = false,
                    Error = $"Không tìm thấy script tại: {ScriptPath}"
                };

            progress?.Report("Đang kết nối Outlook...");

            var psi = new ProcessStartInfo
            {
                FileName = "python",
                Arguments = $"\"{ScriptPath}\" --save-dir \"{dir}\" --days-back {daysBack}",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = System.Text.Encoding.UTF8
            };

            try
            {
                using var process = Process.Start(psi)
                    ?? throw new Exception("Không thể khởi động Python.");

                string stdout = await process.StandardOutput.ReadToEndAsync();
                string stderr = await process.StandardError.ReadToEndAsync();
                await process.WaitForExitAsync();

                if (string.IsNullOrWhiteSpace(stdout))
                    return new DownloadResult
                    {
                        Success = false,
                        Error = $"Script không trả về kết quả. Lỗi: {stderr}"
                    };

                var result = JsonSerializer.Deserialize<DownloadResult>(stdout)
                    ?? new DownloadResult { Success = false, Error = "Không parse được JSON kết quả." };

                if (result.Success)
                    progress?.Report($"Hoàn thành: {result.FilesDownloaded} file từ {result.EmailsProcessed} email.");

                return result;
            }
            catch (Exception ex)
            {
                return new DownloadResult
                {
                    Success = false,
                    Error = $"Lỗi chạy script: {ex.Message}"
                };
            }
        }

        /// <summary>
        /// Mở thư mục lưu file trong Explorer.
        /// </summary>
        public static void OpenSaveFolder(string? saveDir = null)
        {
            string dir = saveDir ?? DefaultSaveDir;
            if (Directory.Exists(dir))
                Process.Start("explorer.exe", dir);
        }
    }
}
