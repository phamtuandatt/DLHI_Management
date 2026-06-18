using System;

namespace MPR_Managerment.Models
{
    public class InternalNotification
    {
        public int Notif_ID { get; set; }
        public string Sender_Username { get; set; } = "";
        public string Sender_FullName { get; set; } = "";
        public string Receiver_Username { get; set; } = "";
        public string Receiver_FullName { get; set; } = "";
        public string Title { get; set; } = "";
        public string Content { get; set; } = "";
        public DateTime Sent_At { get; set; }
        public bool Is_Read { get; set; }
        public DateTime? Read_At { get; set; }
    }
}
