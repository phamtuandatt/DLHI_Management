using System;
using System.Collections.Generic;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;

namespace MPR_Managerment.Services
{
    public class NotificationLogService
    {
        public NotificationLogService()
        {
            EnsureTableExists();
        }

        private void EnsureTableExists()
        {
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            string sql = @"
                IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Notification_Logs]') AND type in (N'U'))
                BEGIN
                    CREATE TABLE [dbo].[Notification_Logs](
                        [Log_ID] [int] IDENTITY(1,1) NOT NULL,
                        [Sent_At] [datetime] NOT NULL,
                        [Sent_By] [nvarchar](255) NULL,
                        [Recipient] [nvarchar](255) NULL,
                        [Type] [nvarchar](50) NULL,
                        [Content] [nvarchar](max) NULL,
                        [Status] [nvarchar](50) NULL,
                        [Error_Message] [nvarchar](max) NULL,
                        [Project_Code] [nvarchar](100) NULL,
                        CONSTRAINT [PK_Notification_Logs] PRIMARY KEY CLUSTERED ([Log_ID] ASC)
                    )
                END";
            using var cmd = new SqlCommand(sql, conn);
            cmd.ExecuteNonQuery();
        }

        public void AddLog(NotificationLog log)
        {
            try
            {
                using var conn = DatabaseHelper.GetConnection();
                conn.Open();
                string sql = @"
                    INSERT INTO Notification_Logs (Sent_At, Sent_By, Recipient, Type, Content, Status, Error_Message, Project_Code)
                    VALUES (@sentAt, @sentBy, @recipient, @type, @content, @status, @err, @proj)";
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@sentAt", log.Sent_At);
                cmd.Parameters.AddWithValue("@sentBy", log.Sent_By ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@recipient", log.Recipient ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@type", log.Type ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@content", log.Content ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@status", log.Status ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@err", log.Error_Message ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@proj", log.Project_Code ?? (object)DBNull.Value);
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                // Fallback to file log if DB fails
                System.IO.File.AppendAllText("notification_service_errors.log", $"{DateTime.Now}: {ex.Message}{Environment.NewLine}");
            }
        }

        public List<NotificationLog> GetLogs(string userFilter, DateTime? fromDate, DateTime? toDate)
        {
            var list = new List<NotificationLog>();
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();

            string sql = "SELECT * FROM Notification_Logs WHERE 1=1";
            if (!string.IsNullOrWhiteSpace(userFilter))
            {
                sql += " AND (Sent_By LIKE @user OR Recipient LIKE @user)";
            }
            if (fromDate.HasValue)
            {
                sql += " AND Sent_At >= @from";
            }
            if (toDate.HasValue)
            {
                sql += " AND Sent_At <= @to";
            }
            sql += " ORDER BY Sent_At DESC";

            using var cmd = new SqlCommand(sql, conn);
            if (!string.IsNullOrWhiteSpace(userFilter))
                cmd.Parameters.AddWithValue("@user", $"%{userFilter}%");
            if (fromDate.HasValue)
                cmd.Parameters.AddWithValue("@from", fromDate.Value);
            if (toDate.HasValue)
                cmd.Parameters.AddWithValue("@to", toDate.Value);

            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                list.Add(new NotificationLog
                {
                    Log_ID = Convert.ToInt32(r["Log_ID"]),
                    Sent_At = Convert.ToDateTime(r["Sent_At"]),
                    Sent_By = r["Sent_By"]?.ToString() ?? "",
                    Recipient = r["Recipient"]?.ToString() ?? "",
                    Type = r["Type"]?.ToString() ?? "",
                    Content = r["Content"]?.ToString() ?? "",
                    Status = r["Status"]?.ToString() ?? "",
                    Error_Message = r["Error_Message"]?.ToString() ?? "",
                    Project_Code = r["Project_Code"]?.ToString() ?? ""
                });
            }
            return list;
        }
    }
}