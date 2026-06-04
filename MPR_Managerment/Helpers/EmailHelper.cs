using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text.Json;
using System.Threading.Tasks;

namespace MPR_Managerment.Helpers
{
    public class EmailSettings
    {
        public string Host      { get; set; } = "smtp.gmail.com";
        public int    Port      { get; set; } = 587;
        public string Username  { get; set; } = "";
        public string Password  { get; set; } = "";
        public string FromName  { get; set; } = "MPR Management System";
        public bool   EnableSsl { get; set; } = true;
    }

    /// <summary>
    /// Helper class for sending emails to suppliers
    /// </summary>
    public class EmailHelper
    {
        private static readonly string SettingsFile = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "email_settings.json");

        private static readonly JsonSerializerOptions _jsonOptions = new() { WriteIndented = true };

        public static EmailSettings LoadSettings()
        {
            try
            {
                if (File.Exists(SettingsFile))
                {
                    string json = File.ReadAllText(SettingsFile);
                    return JsonSerializer.Deserialize<EmailSettings>(json) ?? new EmailSettings();
                }
            }
            catch { /* fall through to defaults */ }
            return new EmailSettings();
        }

        public static void SaveSettings(EmailSettings settings)
        {
            string json = JsonSerializer.Serialize(settings, _jsonOptions);
            File.WriteAllText(SettingsFile, json);
        }

        // SMTP Configuration - Update these with your email server settings
        private static readonly string SmtpServer = "smtp.gmail.com"; // Change to your SMTP server
        private static readonly int SmtpPort = 587;
        private static readonly string SenderEmail = "your-email@gmail.com"; // Change to your email
        private static readonly string SenderPassword = "your-app-password"; // Change to your app password
        private static readonly string SenderName = "MPR Management System";

        /// <summary>
        /// Send email to supplier with MPR details
        /// </summary>
        public static async Task<(bool success, string message)> SendMPREmailToSupplier(
            string supplierEmail,
            string supplierName,
            string mprNo,
            string projectName,
            string projectCode,
            string mprDetails)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(supplierEmail))
                    return (false, "Email nhà cung cấp không hợp lệ");

                using (var client = new SmtpClient(SmtpServer, SmtpPort))
                {
                    client.EnableSsl = true;
                    client.Credentials = new NetworkCredential(SenderEmail, SenderPassword);
                    client.Timeout = 10000;

                    var mailMessage = new MailMessage
                    {
                        From = new MailAddress(SenderEmail, SenderName),
                        Subject = $"Phiếu Yêu cầu Mua Vật tư (MPR) - {mprNo}",
                        Body = BuildEmailBody(supplierName, mprNo, projectName, projectCode, mprDetails),
                        IsBodyHtml = true
                    };

                    var emails = (supplierEmail ?? "")
                        .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(e => e.Trim())
                        .Where(e => !string.IsNullOrWhiteSpace(e));

                    foreach (var email in emails)
                    {
                        mailMessage.To.Add(new MailAddress(email, supplierName));
                    }

                    await client.SendMailAsync(mailMessage);
                    return (true, "Gửi email thành công");
                }
            }
            catch (SmtpException ex)
            {
                return (false, $"Lỗi SMTP: {ex.Message}");
            }
            catch (Exception ex)
            {
                return (false, $"Lỗi gửi email: {ex.Message}");
            }
        }

        /// <summary>
        /// Build HTML email body
        /// </summary>
        private static string BuildEmailBody(string supplierName, string mprNo, string projectName, string projectCode, string details)
        {
            return $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <style>
        body {{ font-family: Arial, sans-serif; color: #333; }}
        .container {{ max-width: 600px; margin: 0 auto; padding: 20px; }}
        .header {{ background-color: #0078d4; color: white; padding: 15px; border-radius: 5px; }}
        .content {{ margin: 20px 0; }}
        .field {{ margin: 10px 0; }}
        .label {{ font-weight: bold; color: #0078d4; }}
        .footer {{ margin-top: 30px; padding-top: 15px; border-top: 1px solid #ddd; font-size: 12px; color: #666; }}
        table {{ width: 100%; border-collapse: collapse; margin: 15px 0; }}
        th, td {{ border: 1px solid #ddd; padding: 10px; text-align: left; }}
        th {{ background-color: #f0f0f0; }}
    </style>
</head>
<body>
    <div class='container'>
        <div class='header'>
            <h2>Phiếu Yêu cầu Mua Vật tư (MPR)</h2>
        </div>
        
        <div class='content'>
            <p>Kính gửi <strong>{supplierName}</strong>,</p>
            
            <p>Chúng tôi gửi đến bạn phiếu yêu cầu mua vật tư như chi tiết dưới đây:</p>
            
            <div class='field'>
                <span class='label'>Số MPR:</span> {mprNo}
            </div>
            
            <div class='field'>
                <span class='label'>Tên Dự án:</span> {projectName}
            </div>
            
            <div class='field'>
                <span class='label'>Mã Dự án:</span> {projectCode}
            </div>
            
            <div class='field'>
                <span class='label'>Chi tiết vật tư:</span>
                <div style='margin-top: 10px; background-color: #f9f9f9; padding: 10px; border-radius: 3px;'>
                    {details}
                </div>
            </div>
            
            <p>Vui lòng xác nhận nhận được phiếu này và cung cấp báo giá trong thời gian sớm nhất.</p>
            
            <p>Cảm ơn bạn!</p>
        </div>
        
        <div class='footer'>
            <p>Email này được gửi tự động từ Hệ thống Quản lý MPR.</p>
            <p>Vui lòng không trả lời email này. Nếu có câu hỏi, vui lòng liên hệ với bộ phận mua sắm.</p>
        </div>
    </div>
</body>
</html>";
        }

        /// <summary>
        /// Validate email address format
        /// </summary>
        public static bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }
    }
}