using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace MPR_Managerment.Services
{
    public class ClassifyResult
    {
        [JsonPropertyName("file")] public string File { get; set; } = "";
        [JsonPropertyName("status")] public string Status { get; set; } = "";
        [JsonPropertyName("po_no")] public string? PoNo { get; set; }
        [JsonPropertyName("project_code")] public string? ProjectCode { get; set; }
        [JsonPropertyName("project_name")] public string? ProjectName { get; set; }
        [JsonPropertyName("reason")] public string? Reason { get; set; }
        [JsonPropertyName("dest")] public string? Dest { get; set; }
        [JsonPropertyName("duplicates_removed")] public int DuplicatesRemoved { get; set; }
        [JsonPropertyName("duplicates")] public List<string>? Duplicates { get; set; }
    }

    public class ClassifySummary
    {
        [JsonPropertyName("total")] public int Total { get; set; }
        [JsonPropertyName("classified")] public int Classified { get; set; }
        [JsonPropertyName("unclassified")] public int Unclassified { get; set; }
    }

    public class ClassifyResponse
    {
        [JsonPropertyName("success")] public bool Success { get; set; }
        [JsonPropertyName("error")] public string? Error { get; set; }
        [JsonPropertyName("results")] public List<ClassifyResult> Results { get; set; } = new();
        [JsonPropertyName("summary")] public ClassifySummary Summary { get; set; } = new();
    }

    public static class InvoiceClassifierService
    {
        private static readonly string ScriptPath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "Scripts", "invoice_classifier.py");

        /// <summary>
        /// Phân loại danh sách file PDF cụ thể.
        /// </summary>
        public static Task<ClassifyResponse> ClassifyFilesAsync(
            IEnumerable<string> pdfFiles,
            string unclassifiedBase,
            IProgress<string>? progress = null)
        {
            string filesArg = string.Join(" ", System.Linq.Enumerable.Select(pdfFiles, f => $"\"{f}\""));
            return RunScriptAsync($"--files {filesArg} --unclassified-base \"{unclassifiedBase}\"", progress);
        }

        /// <summary>
        /// Quét toàn bộ PDF trong thư mục và phân loại.
        /// </summary>
        public static Task<ClassifyResponse> ClassifyFolderAsync(
            string scanDir,
            string unclassifiedBase,
            IProgress<string>? progress = null)
        {
            return RunScriptAsync(
                $"--scan-dir \"{scanDir}\" --unclassified-base \"{unclassifiedBase}\"", progress);
        }

        private static async Task<ClassifyResponse> RunScriptAsync(string args, IProgress<string>? progress)
        {
            if (!File.Exists(ScriptPath))
                return new ClassifyResponse { Success = false, Error = $"Không tìm thấy script: {ScriptPath}" };

            progress?.Report("Đang phân tích PDF và tra cứu database...");

            var psi = new ProcessStartInfo
            {
                FileName = "python",
                Arguments = $"\"{ScriptPath}\" {args}",
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
                    return new ClassifyResponse { Success = false, Error = $"Script lỗi: {stderr}" };

                return JsonSerializer.Deserialize<ClassifyResponse>(stdout)
                    ?? new ClassifyResponse { Success = false, Error = "Không parse được kết quả." };
            }
            catch (Exception ex)
            {
                return new ClassifyResponse { Success = false, Error = ex.Message };
            }
        }
    }
}
