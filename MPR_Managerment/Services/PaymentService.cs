// ============================================================
//  FILE: Services/PaymentService.cs
// ============================================================
using System;
using System.Collections.Generic;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;

namespace MPR_Managerment.Services
{
    public class PaymentService
    {
        // =====================================================
        //  SCHEDULE — Kế hoạch thanh toán
        // =====================================================
        public List<PaymentSchedule> GetSchedules(int poId)
        {
            var list = new List<PaymentSchedule>();
            try
            {
                using var conn = DatabaseHelper.GetConnection();
                conn.Open();
                // Đặt tên alias rõ ràng — tránh conflict cột trùng tên khi SELECT *
                var cmd = new SqlCommand(@"
                    SELECT
                        ps.Schedule_ID, ps.PO_ID, ps.Dot_TT,
                        ps.Payment_Type, ps.Pay_Method, ps.Percent_TT,
                        ps.Amount_Plan,  ps.Due_Date,   ps.Delivery_Ref,
                        ps.Description,  ps.Status,
                        ps.Created_Date, ps.Created_By,
                        h.PONo          AS PONo,
                        h.Project_Name  AS Project_Name
                    FROM PO_Payment_Schedule ps
                    INNER JOIN PO_head h ON h.PO_ID = ps.PO_ID
                    WHERE ps.PO_ID = @poId
                    ORDER BY ps.Dot_TT", conn);
                cmd.Parameters.AddWithValue("@poId", poId);
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    try { list.Add(MapSchedule(r)); }
                    catch { /* bỏ qua dòng lỗi */ }
                }
            }
            catch { }
            return list;
        }

        // Load TẤT CẢ schedules của mọi PO bằng 1 query duy nhất
        public List<PaymentSchedule> GetAllSchedules()
        {
            var list = new List<PaymentSchedule>();
            try
            {
                using var conn = DatabaseHelper.GetConnection();
                conn.Open();
                var cmd = new SqlCommand(@"
                    SELECT
                        ps.Schedule_ID, ps.PO_ID, ps.Dot_TT,
                        ps.Payment_Type, ps.Pay_Method, ps.Percent_TT,
                        ps.Amount_Plan,  ps.Due_Date,   ps.Delivery_Ref,
                        ps.Description,  ps.Status,
                        ps.Created_Date, ps.Created_By,
                        h.PONo          AS PONo,
                        h.Project_Name  AS Project_Name
                    FROM PO_Payment_Schedule ps
                    INNER JOIN PO_head h ON h.PO_ID = ps.PO_ID
                    ORDER BY ps.PO_ID, ps.Dot_TT", conn);
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    try { list.Add(MapSchedule(r)); }
                    catch { /* bỏ qua dòng lỗi, đọc tiếp */ }
                }
            }
            catch { }
            return list;
        }

        public int InsertSchedule(PaymentSchedule s, string createdBy)
        {
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            var cmd = new SqlCommand(@"
                INSERT INTO PO_Payment_Schedule
                    (PO_ID, Dot_TT, Payment_Type, Pay_Method, Percent_TT,
                     Amount_Plan, Due_Date, Delivery_Ref, Description, Status, Created_By)
                VALUES (@poId,@dot,@type,@method,@pct,@amt,@due,@delref,@desc,@status,@by);
                SELECT SCOPE_IDENTITY();", conn);
            cmd.Parameters.AddWithValue("@poId", s.PO_ID);
            cmd.Parameters.AddWithValue("@dot", s.Dot_TT);
            cmd.Parameters.AddWithValue("@type", s.Payment_Type);
            cmd.Parameters.AddWithValue("@method", s.Pay_Method);
            cmd.Parameters.AddWithValue("@pct", s.Percent_TT);
            cmd.Parameters.AddWithValue("@amt", s.Amount_Plan);
            cmd.Parameters.AddWithValue("@due", s.Due_Date.HasValue ? (object)s.Due_Date.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@delref", s.Delivery_Ref ?? "");
            cmd.Parameters.AddWithValue("@desc", s.Description ?? "");
            cmd.Parameters.AddWithValue("@status", s.Status);
            cmd.Parameters.AddWithValue("@by", createdBy);
            return Convert.ToInt32(cmd.ExecuteScalar());
        }

        public void UpdateSchedule(PaymentSchedule s)
        {
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            var cmd = new SqlCommand(@"
                UPDATE PO_Payment_Schedule SET
                    Dot_TT       = @dot,    Payment_Type = @type,
                    Pay_Method   = @method, Percent_TT   = @pct,
                    Amount_Plan  = @amt,    Due_Date     = @due,
                    Delivery_Ref = @delref, Description  = @desc,
                    Status       = @status
                WHERE Schedule_ID = @id", conn);
            cmd.Parameters.AddWithValue("@id", s.Schedule_ID);
            cmd.Parameters.AddWithValue("@dot", s.Dot_TT);
            cmd.Parameters.AddWithValue("@type", s.Payment_Type);
            cmd.Parameters.AddWithValue("@method", s.Pay_Method);
            cmd.Parameters.AddWithValue("@pct", s.Percent_TT);
            cmd.Parameters.AddWithValue("@amt", s.Amount_Plan);
            cmd.Parameters.AddWithValue("@due", s.Due_Date.HasValue ? (object)s.Due_Date.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@delref", s.Delivery_Ref ?? "");
            cmd.Parameters.AddWithValue("@desc", s.Description ?? "");
            cmd.Parameters.AddWithValue("@status", s.Status);
            cmd.ExecuteNonQuery();
        }

        public void DeleteSchedule(int scheduleId)
        {
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            new SqlCommand(
                $"DELETE FROM PO_Payment_Schedule WHERE Schedule_ID = {scheduleId}", conn)
                .ExecuteNonQuery();
        }

        // =====================================================
        //  HISTORY — Lịch sử thanh toán thực tế
        // =====================================================
        public List<PaymentHistory> GetHistories(int poId)
        {
            var list = new List<PaymentHistory>();
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            var cmd = new SqlCommand(@"
                SELECT ph.*, h.PONo, h.Project_Name,
                       ISNULL(s.Company_Name,'') AS Supplier_Name
                FROM PO_Payment_History ph
                INNER JOIN PO_head h ON h.PO_ID = ph.PO_ID
                LEFT JOIN  Suppliers s ON s.Supplier_ID = ph.Supplier_ID
                WHERE ph.PO_ID = @poId ORDER BY ph.Payment_Date DESC", conn);
            cmd.Parameters.AddWithValue("@poId", poId);
            using var r = cmd.ExecuteReader();
            while (r.Read()) list.Add(MapHistory(r));
            return list;
        }

        public int InsertHistory(PaymentHistory p, string createdBy)
        {
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            var cmd = new SqlCommand(@"
                INSERT INTO PO_Payment_History
                    (Schedule_ID, PO_ID, Supplier_ID, Payment_Date, Amount_Paid,
                     Payment_Method, Bank_Name, Transaction_No,
                     Currency, Exchange_Rate, Notes, Created_By)
                VALUES (@sid,@poId,@suppId,@date,@amt,
                        @method,@bank,@transno,
                        @currency,@rate,@notes,@by);
                SELECT SCOPE_IDENTITY();", conn);
            cmd.Parameters.AddWithValue("@sid", p.Schedule_ID.HasValue ? (object)p.Schedule_ID.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@poId", p.PO_ID);
            cmd.Parameters.AddWithValue("@suppId", p.Supplier_ID.HasValue ? (object)p.Supplier_ID.Value : DBNull.Value);
            cmd.Parameters.AddWithValue("@date", p.Payment_Date);
            cmd.Parameters.AddWithValue("@amt", p.Amount_Paid);
            cmd.Parameters.AddWithValue("@method", p.Payment_Method ?? "");
            cmd.Parameters.AddWithValue("@bank", p.Bank_Name ?? "");
            cmd.Parameters.AddWithValue("@transno", p.Transaction_No ?? "");
            cmd.Parameters.AddWithValue("@currency", p.Currency ?? "VND");
            cmd.Parameters.AddWithValue("@rate", p.Exchange_Rate);
            cmd.Parameters.AddWithValue("@notes", p.Notes ?? "");
            cmd.Parameters.AddWithValue("@by", createdBy);
            int newId = Convert.ToInt32(cmd.ExecuteScalar());

            // Tự động cập nhật trạng thái đợt nếu có liên kết
            if (p.Schedule_ID.HasValue)
                AutoUpdateStatus(p.Schedule_ID.Value, conn);

            return newId;
        }

        public void DeleteHistory(int paymentId)
        {
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            var schedCmd = new SqlCommand(
                $"SELECT Schedule_ID FROM PO_Payment_History WHERE Payment_ID={paymentId}", conn);
            var sid = schedCmd.ExecuteScalar();
            new SqlCommand(
                $"DELETE FROM PO_Payment_History WHERE Payment_ID={paymentId}", conn)
                .ExecuteNonQuery();
            if (sid != null && sid != DBNull.Value)
                AutoUpdateStatus(Convert.ToInt32(sid), conn);
        }

        // Gọi SP tự động cập nhật trạng thái (Chưa TT / Một phần / Đã TT đủ)
        private void AutoUpdateStatus(int scheduleId, SqlConnection conn)
        {
            var cmd = new SqlCommand("sp_AutoUpdatePaymentStatus", conn)
            {
                CommandType = System.Data.CommandType.StoredProcedure
            };
            cmd.Parameters.AddWithValue("@Schedule_ID", scheduleId);
            cmd.ExecuteNonQuery();
        }

        // Load summary cho 1 PO cụ thể — dùng sau khi cập nhật schedule
        public POPaymentSummary GetPOSummary(int poId)
        {
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            var cmd = new SqlCommand(
                "SELECT * FROM vw_PO_Payment_Summary WHERE PO_ID = @id", conn);
            cmd.Parameters.AddWithValue("@id", poId);
            using var r = cmd.ExecuteReader();
            return r.Read() ? MapSummary(r) : null;
        }

        // =====================================================
        //  SUMMARY — Tổng hợp thanh toán PO
        // =====================================================
        public List<POPaymentSummary> GetPOSummaries(
            int? supplierId = null,
            string status = null)
        {
            var list = new List<POPaymentSummary>();
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();

            string where = "WHERE 1=1";
            if (supplierId.HasValue)
                where += $" AND Supplier_ID = {supplierId.Value}";
            if (!string.IsNullOrEmpty(status) && status != "Tất cả")
                where += $" AND Payment_Status = N'{status}'";

            var cmd = new SqlCommand(
                $"SELECT * FROM vw_PO_Payment_Summary {where} ORDER BY PO_Date DESC", conn);
            using var r = cmd.ExecuteReader();
            while (r.Read()) list.Add(MapSummary(r));
            return list;
        }

        // =====================================================
        //  SUPPLIER DEBT — Công nợ tổng hợp theo NCC
        // =====================================================
        public List<SupplierDebtSummary> GetSupplierDebt()
        {
            var list = new List<SupplierDebtSummary>();
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            var cmd = new SqlCommand(
                "SELECT * FROM vw_Supplier_Debt_Summary ORDER BY Total_Debt DESC", conn);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                list.Add(new SupplierDebtSummary
                {
                    Supplier_ID = Convert.ToInt32(r["Supplier_ID"]),
                    Supplier_Name = r["Supplier_Name"]?.ToString() ?? "",
                    Supplier_Short = r["Supplier_Short"]?.ToString() ?? "",
                    Phone = r["Phone"]?.ToString() ?? "",
                    Email = r["Email"]?.ToString() ?? "",
                    Total_PO = Convert.ToInt32(r["Total_PO"]),
                    Total_PO_Value = D(r["Total_PO_Value"]),
                    Total_Paid = D(r["Total_Paid"]),
                    Total_Debt = D(r["Total_Debt"]),
                    Overdue_PO_Count = r["Overdue_PO_Count"] != DBNull.Value
                                       ? Convert.ToInt32(r["Overdue_PO_Count"]) : 0
                });
            return list;
        }

        // =====================================================
        //  DEBT REPORT — Báo cáo theo kỳ
        // =====================================================
        public List<DebtReportItem> GetDebtReport(
            DateTime fromDate,
            DateTime toDate,
            int? supplierId = null)
        {
            var list = new List<DebtReportItem>();
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            var cmd = new SqlCommand("sp_GetDebtByDateRange", conn)
            {
                CommandType = System.Data.CommandType.StoredProcedure
            };
            cmd.Parameters.AddWithValue("@FromDate", fromDate.Date);
            cmd.Parameters.AddWithValue("@ToDate", toDate.Date);
            cmd.Parameters.AddWithValue("@SupplierID",
                supplierId.HasValue ? (object)supplierId.Value : DBNull.Value);

            using var r = cmd.ExecuteReader();
            while (r.Read())
                list.Add(new DebtReportItem
                {
                    Supplier_ID = Convert.ToInt32(r["Supplier_ID"]),
                    Supplier_Name = r["Supplier_Name"]?.ToString() ?? "",
                    Supplier_Short = r["Supplier_Short"]?.ToString() ?? "",
                    PO_ID = Convert.ToInt32(r["PO_ID"]),
                    PONo = r["PONo"]?.ToString() ?? "",
                    Project_Name = r["Project_Name"]?.ToString() ?? "",
                    PO_Date = r["PO_Date"] != DBNull.Value ? Convert.ToDateTime(r["PO_Date"]) : null,
                    Total_Amount = D(r["Total_Amount"]),
                    Paid_In_Range = D(r["Paid_In_Range"]),
                    Paid_Before_Range = D(r["Paid_Before_Range"]),
                    Remaining_Debt = D(r["Remaining_Debt"]),
                    Payment_Status = r["Payment_Status"]?.ToString() ?? "",
                    Is_Overdue = r["Is_Overdue"] != DBNull.Value && Convert.ToBoolean(r["Is_Overdue"]),
                    Next_Due_Date = r["Next_Due_Date"] != DBNull.Value ? Convert.ToDateTime(r["Next_Due_Date"]) : null
                });
            return list;
        }

        // =====================================================
        //  MAP HELPERS
        // =====================================================
        private static decimal D(object v) =>
            v != DBNull.Value && v != null ? Convert.ToDecimal(v) : 0;

        private PaymentSchedule MapSchedule(SqlDataReader r) => new PaymentSchedule
        {
            Schedule_ID = r["Schedule_ID"] != DBNull.Value ? Convert.ToInt32(r["Schedule_ID"]) : 0,
            PO_ID = r["PO_ID"] != DBNull.Value ? Convert.ToInt32(r["PO_ID"]) : 0,
            PONo = r["PONo"] != DBNull.Value ? r["PONo"].ToString() : "",
            Project_Name = r["Project_Name"] != DBNull.Value ? r["Project_Name"].ToString() : "",
            Dot_TT = r["Dot_TT"] != DBNull.Value ? Convert.ToInt32(r["Dot_TT"]) : 0,
            Payment_Type = r["Payment_Type"] != DBNull.Value ? r["Payment_Type"].ToString() : "",
            Pay_Method = r["Pay_Method"] != DBNull.Value ? r["Pay_Method"].ToString() : "Full",
            Percent_TT = D(r["Percent_TT"]),
            Amount_Plan = D(r["Amount_Plan"]),
            Due_Date = r["Due_Date"] != DBNull.Value ? Convert.ToDateTime(r["Due_Date"]) : (DateTime?)null,
            Delivery_Ref = r["Delivery_Ref"] != DBNull.Value ? r["Delivery_Ref"].ToString() : "",
            Description = r["Description"] != DBNull.Value ? r["Description"].ToString() : "",
            Status = r["Status"] != DBNull.Value ? r["Status"].ToString() : "Chưa TT",
            Created_Date = r["Created_Date"] != DBNull.Value ? Convert.ToDateTime(r["Created_Date"]) : (DateTime?)null,
            Created_By = r["Created_By"] != DBNull.Value ? r["Created_By"].ToString() : ""
        };

        private PaymentHistory MapHistory(SqlDataReader r) => new PaymentHistory
        {
            Payment_ID = Convert.ToInt32(r["Payment_ID"]),
            Schedule_ID = r["Schedule_ID"] != DBNull.Value ? Convert.ToInt32(r["Schedule_ID"]) : null,
            PO_ID = Convert.ToInt32(r["PO_ID"]),
            PONo = r["PONo"]?.ToString() ?? "",
            Project_Name = r["Project_Name"]?.ToString() ?? "",
            Supplier_ID = r["Supplier_ID"] != DBNull.Value ? Convert.ToInt32(r["Supplier_ID"]) : null,
            Supplier_Name = r["Supplier_Name"]?.ToString() ?? "",
            Payment_Date = Convert.ToDateTime(r["Payment_Date"]),
            Amount_Paid = D(r["Amount_Paid"]),
            Payment_Method = r["Payment_Method"]?.ToString() ?? "",
            Bank_Name = r["Bank_Name"]?.ToString() ?? "",
            Transaction_No = r["Transaction_No"]?.ToString() ?? "",
            Currency = r["Currency"]?.ToString() ?? "VND",
            Exchange_Rate = D(r["Exchange_Rate"]),
            Notes = r["Notes"]?.ToString() ?? "",
            Created_Date = r["Created_Date"] != DBNull.Value ? Convert.ToDateTime(r["Created_Date"]) : null,
            Created_By = r["Created_By"]?.ToString() ?? ""
        };

        private POPaymentSummary MapSummary(SqlDataReader r) => new POPaymentSummary
        {
            PO_ID = Convert.ToInt32(r["PO_ID"]),
            PONo = r["PONo"]?.ToString() ?? "",
            Project_Name = r["Project_Name"]?.ToString() ?? "",
            WorkorderNo = r["WorkorderNo"]?.ToString() ?? "",
            Supplier_ID = r["Supplier_ID"] != DBNull.Value ? Convert.ToInt32(r["Supplier_ID"]) : null,
            Supplier_Name = r["Supplier_Name"]?.ToString() ?? "",
            Supplier_Short = r["Supplier_Short"]?.ToString() ?? "",
            PO_Date = r["PO_Date"] != DBNull.Value ? Convert.ToDateTime(r["PO_Date"]) : null,
            Total_PO_Amount = D(r["Total_PO_Amount"]),
            Total_Plan = D(r["Total_Plan"]),
            Total_Paid = D(r["Total_Paid"]),
            Amount_Remaining = D(r["Amount_Remaining"]),
            Percent_Paid = D(r["Percent_Paid"]),
            Payment_Status = r["Payment_Status"]?.ToString() ?? "",
            Is_Overdue = r["Is_Overdue"] != DBNull.Value && Convert.ToBoolean(r["Is_Overdue"]),
            Last_Payment_Date = r["Last_Payment_Date"] != DBNull.Value ? Convert.ToDateTime(r["Last_Payment_Date"]) : null,
            Next_Due_Date = r["Next_Due_Date"] != DBNull.Value ? Convert.ToDateTime(r["Next_Due_Date"]) : null,
            Total_Dots = r["Total_Dots"] != DBNull.Value ? Convert.ToInt32(r["Total_Dots"]) : 0,
            Done_Dots = r["Done_Dots"] != DBNull.Value ? Convert.ToInt32(r["Done_Dots"]) : 0
        };
    }
}