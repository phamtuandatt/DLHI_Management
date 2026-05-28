using System;

namespace MPR_Managerment.Models
{
    public class NotificationLog
    {
        public int Log_ID { get; set; }
        public DateTime Sent_At { get; set; }
        public string Sent_By { get; set; } = "";
        public string Recipient { get; set; } = "";
        public string Type { get; set; } = ""; // "Zalo" or "Email"
        public string Content { get; set; } = "";
        public string Status { get; set; } = ""; // "Success", "Failed", "Dropped"
        public string Error_Message { get; set; } = "";
        public string Project_Code { get; set; } = "";
    }
}