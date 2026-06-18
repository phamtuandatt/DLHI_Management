using System;
using System.Collections.Generic;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;

namespace MPR_Managerment.Services
{
    public class InternalNotificationService
    {
        public InternalNotificationService()
        {
            EnsureTableExists();
        }

        private void EnsureTableExists()
        {
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            const string sql = @"
                IF NOT EXISTS (SELECT * FROM sys.objects
                               WHERE object_id = OBJECT_ID(N'[dbo].[Internal_Notifications]') AND type = N'U')
                BEGIN
                    CREATE TABLE [dbo].[Internal_Notifications] (
                        [Notif_ID]           INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
                        [Sender_Username]    NVARCHAR(100)  NOT NULL,
                        [Sender_FullName]    NVARCHAR(255)  NOT NULL DEFAULT '',
                        [Receiver_Username]  NVARCHAR(100)  NOT NULL,
                        [Receiver_FullName]  NVARCHAR(255)  NOT NULL DEFAULT '',
                        [Title]              NVARCHAR(500)  NOT NULL DEFAULT '',
                        [Content]            NVARCHAR(MAX)  NOT NULL DEFAULT '',
                        [Sent_At]            DATETIME       NOT NULL DEFAULT GETDATE(),
                        [Is_Read]            BIT            NOT NULL DEFAULT 0,
                        [Read_At]            DATETIME       NULL
                    )
                END";
            using var cmd = new SqlCommand(sql, conn);
            cmd.ExecuteNonQuery();
        }

        // Gửi thông báo đến nhiều user (truyền danh sách username)
        public void Send(string senderUsername, string senderFullName,
                         IEnumerable<(string username, string fullName)> receivers,
                         string title, string content)
        {
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            const string sql = @"
                INSERT INTO Internal_Notifications
                    (Sender_Username, Sender_FullName, Receiver_Username, Receiver_FullName, Title, Content, Sent_At)
                VALUES (@sender, @senderFull, @receiver, @receiverFull, @title, @content, GETDATE())";

            foreach (var (username, fullName) in receivers)
            {
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@sender",       senderUsername);
                cmd.Parameters.AddWithValue("@senderFull",   senderFullName);
                cmd.Parameters.AddWithValue("@receiver",     username);
                cmd.Parameters.AddWithValue("@receiverFull", fullName);
                cmd.Parameters.AddWithValue("@title",        title);
                cmd.Parameters.AddWithValue("@content",      content);
                cmd.ExecuteNonQuery();
            }
        }

        // Lấy tất cả thông báo chưa đọc của một user
        public List<InternalNotification> GetUnread(string receiverUsername)
        {
            var list = new List<InternalNotification>();
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            const string sql = @"
                SELECT * FROM Internal_Notifications
                WHERE Receiver_Username = @user AND Is_Read = 0
                ORDER BY Sent_At DESC";
            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@user", receiverUsername);
            using var r = cmd.ExecuteReader();
            while (r.Read()) list.Add(Map(r));
            return list;
        }

        // Đếm thông báo chưa đọc
        public int CountUnread(string receiverUsername)
        {
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            const string sql = @"
                SELECT COUNT(1) FROM Internal_Notifications
                WHERE Receiver_Username = @user AND Is_Read = 0";
            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@user", receiverUsername);
            return Convert.ToInt32(cmd.ExecuteScalar());
        }

        // Đánh dấu đã đọc
        public void MarkRead(int notifId)
        {
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            const string sql = @"
                UPDATE Internal_Notifications
                SET Is_Read = 1, Read_At = GETDATE()
                WHERE Notif_ID = @id";
            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@id", notifId);
            cmd.ExecuteNonQuery();
        }

        // Đánh dấu tất cả đã đọc cho một user
        public void MarkAllRead(string receiverUsername)
        {
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            const string sql = @"
                UPDATE Internal_Notifications
                SET Is_Read = 1, Read_At = GETDATE()
                WHERE Receiver_Username = @user AND Is_Read = 0";
            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@user", receiverUsername);
            cmd.ExecuteNonQuery();
        }

        // Lấy lịch sử (có thể lọc theo sender hoặc receiver)
        public List<InternalNotification> GetHistory(string usernameFilter, DateTime? from, DateTime? to)
        {
            var list = new List<InternalNotification>();
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();

            string sql = "SELECT * FROM Internal_Notifications WHERE 1=1";
            if (!string.IsNullOrWhiteSpace(usernameFilter))
                sql += " AND (Sender_Username LIKE @u OR Receiver_Username LIKE @u OR Sender_FullName LIKE @u OR Receiver_FullName LIKE @u)";
            if (from.HasValue)
                sql += " AND Sent_At >= @from";
            if (to.HasValue)
                sql += " AND Sent_At < @to";
            sql += " ORDER BY Sent_At DESC";

            using var cmd = new SqlCommand(sql, conn);
            if (!string.IsNullOrWhiteSpace(usernameFilter))
                cmd.Parameters.AddWithValue("@u", $"%{usernameFilter}%");
            if (from.HasValue)
                cmd.Parameters.AddWithValue("@from", from.Value);
            if (to.HasValue)
                cmd.Parameters.AddWithValue("@to", to.Value.AddDays(1));

            using var r = cmd.ExecuteReader();
            while (r.Read()) list.Add(Map(r));
            return list;
        }

        private static InternalNotification Map(SqlDataReader r) => new()
        {
            Notif_ID          = Convert.ToInt32(r["Notif_ID"]),
            Sender_Username   = r["Sender_Username"]?.ToString() ?? "",
            Sender_FullName   = r["Sender_FullName"]?.ToString() ?? "",
            Receiver_Username = r["Receiver_Username"]?.ToString() ?? "",
            Receiver_FullName = r["Receiver_FullName"]?.ToString() ?? "",
            Title             = r["Title"]?.ToString() ?? "",
            Content           = r["Content"]?.ToString() ?? "",
            Sent_At           = Convert.ToDateTime(r["Sent_At"]),
            Is_Read           = Convert.ToBoolean(r["Is_Read"]),
            Read_At           = r["Read_At"] == DBNull.Value ? null : (DateTime?)Convert.ToDateTime(r["Read_At"])
        };
    }
}
