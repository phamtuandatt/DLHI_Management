using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text.Json;

namespace MPR_Managerment.Helpers
{
    public class EmailSettings
    {
        public string Host { get; set; } = "";
        public int Port { get; set; } = 587;
        public string Username { get; set; } = "";
        public string Password { get; set; } = "";
        public string FromName { get; set; } = "ERP System";
        public bool EnableSsl { get; set; } = true;
    }

    public static class EmailHelper
    {
        private static readonly string SettingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "MPR_Managerment", "email_config.json");

        public static EmailSettings LoadSettings()
        {
            try
            {
                if (File.Exists(SettingsPath))
                    return JsonSerializer.Deserialize<EmailSettings>(File.ReadAllText(SettingsPath)) ?? new EmailSettings();
            }
            catch { }
            return new EmailSettings();
        }

        public static void SaveSettings(EmailSettings s)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
            File.WriteAllText(SettingsPath, JsonSerializer.Serialize(s, new JsonSerializerOptions { WriteIndented = true }));
        }

        public static bool IsConfigured(EmailSettings s) =>
            !string.IsNullOrWhiteSpace(s.Host) && !string.IsNullOrWhiteSpace(s.Username);

        public static void SendEmail(EmailSettings settings, string to, string subject, string body, string attachmentPath = null)
        {
            using var client = new SmtpClient(settings.Host, settings.Port)
            {
                EnableSsl = settings.EnableSsl,
                Credentials = new NetworkCredential(settings.Username, settings.Password),
                Timeout = 30000
            };
            using var msg = new MailMessage
            {
                From = new MailAddress(settings.Username, settings.FromName),
                Subject = subject,
                Body = body,
                IsBodyHtml = false
            };
            foreach (var addr in to.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string trimmed = addr.Trim();
                if (!string.IsNullOrEmpty(trimmed)) msg.To.Add(trimmed);
            }
            if (!string.IsNullOrEmpty(attachmentPath) && File.Exists(attachmentPath))
                msg.Attachments.Add(new Attachment(attachmentPath));
            client.Send(msg);
        }
    }
}
