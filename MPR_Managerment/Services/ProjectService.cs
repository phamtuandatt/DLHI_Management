using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;

namespace MPR_Managerment.Services
{
    public class ProjectService
    {
        public async Task<DataTable> GetProjects()
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                string sql = "SELECT Id, ProjectCode FROM ProjectInfo";
                var cmd = new SqlCommand(sql, conn);
                DataTable dt = new DataTable();
                using (SqlDataReader reader = await cmd.ExecuteReaderAsync()) // Đọc dữ liệu ngầm
                {
                    dt.Load(reader);
                }
                return dt;
            }
        }

        public List<ProjectInfo> GetAll()
        {
            var list = new List<ProjectInfo>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("SELECT * FROM ProjectInfo ORDER BY CreatedDate DESC", conn);
                var r = cmd.ExecuteReader();
                while (r.Read()) list.Add(Map(r));
            }
            return list;
        }

        public ProjectInfo GetByProjectCode(string projectCode)
        {
            var p = new ProjectInfo();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("SELECT TOP 1 * FROM ProjectInfo WHERE ProjectCode = @code", conn);
                cmd.Parameters.AddWithValue("@code", $"{projectCode}");
                var r = cmd.ExecuteReader();
                while (r.Read()) p = (Map(r));
            }
            return p;
        }

        public List<ProjectInfo> Search(string keyword)
        {
            var list = new List<ProjectInfo>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"SELECT * FROM ProjectInfo 
                    WHERE ProjectName LIKE @kw OR ProjectCode LIKE @kw 
                       OR WorkorderNo LIKE @kw OR Customer LIKE @kw
                    ORDER BY CreatedDate DESC", conn);
                cmd.Parameters.AddWithValue("@kw", $"%{keyword}%");
                var r = cmd.ExecuteReader();
                while (r.Read()) list.Add(Map(r));
            }
            return list;
        }

        public int Insert(ProjectInfo p, string createdBy)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    INSERT INTO ProjectInfo 
                        (ProjectName, ProjectCode, WorkorderNo, Customer, PJWeight, PJBudget,
                         POCode, MPRCode, Status, Notes, PO_Link, RIR_link, MPR_link, CreatedDate)
                    VALUES 
                        (@ProjectName, @ProjectCode, @WorkorderNo, @Customer, @PJWeight, @PJBudget,
                         @POCode, @MPRCode, @Status, @Notes, @PO_Link, @RIR_link, @MPR_link, GETDATE());
                    SELECT SCOPE_IDENTITY();", conn);

                cmd.Parameters.AddWithValue("@ProjectName", p.ProjectName);
                cmd.Parameters.AddWithValue("@ProjectCode", p.ProjectCode);
                cmd.Parameters.AddWithValue("@WorkorderNo", p.WorkorderNo ?? "");
                cmd.Parameters.AddWithValue("@Customer", p.Customer ?? "");
                cmd.Parameters.AddWithValue("@PJWeight", p.PJWeight);
                cmd.Parameters.AddWithValue("@PJBudget", p.PJBudget);
                cmd.Parameters.AddWithValue("@POCode", p.POCode ?? "");
                cmd.Parameters.AddWithValue("@MPRCode", p.MPRCode ?? "");
                cmd.Parameters.AddWithValue("@Status", p.Status ?? "Đang thực hiện");
                cmd.Parameters.AddWithValue("@Notes", p.Notes ?? "");
                cmd.Parameters.AddWithValue("@PO_Link", p.PO_Link ?? "");
                cmd.Parameters.AddWithValue("@RIR_link", p.RIR_Link ?? "");
                cmd.Parameters.AddWithValue("@MPR_link", p.MPR_Link ?? "");

                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        public void Update(ProjectInfo p)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    UPDATE ProjectInfo SET
                        ProjectName  = @ProjectName,
                        ProjectCode  = @ProjectCode,
                        WorkorderNo  = @WorkorderNo,
                        Customer     = @Customer,
                        PJWeight     = @PJWeight,
                        PJBudget     = @PJBudget,
                        POCode       = @POCode,
                        MPRCode      = @MPRCode,
                        Status       = @Status,
                        Notes        = @Notes,
                        PO_Link      = @PO_Link,
                        RIR_link     = @RIR_link,
                        MPR_link     = @MPR_link,
                        ModifiedDate = GETDATE()
                    WHERE Id = @Id", conn);

                cmd.Parameters.AddWithValue("@Id", p.Id);
                cmd.Parameters.AddWithValue("@ProjectName", p.ProjectName);
                cmd.Parameters.AddWithValue("@ProjectCode", p.ProjectCode);
                cmd.Parameters.AddWithValue("@WorkorderNo", p.WorkorderNo ?? "");
                cmd.Parameters.AddWithValue("@Customer", p.Customer ?? "");
                cmd.Parameters.AddWithValue("@PJWeight", p.PJWeight);
                cmd.Parameters.AddWithValue("@PJBudget", p.PJBudget);
                cmd.Parameters.AddWithValue("@POCode", p.POCode ?? "");
                cmd.Parameters.AddWithValue("@MPRCode", p.MPRCode ?? "");
                cmd.Parameters.AddWithValue("@Status", p.Status ?? "");
                cmd.Parameters.AddWithValue("@Notes", p.Notes ?? "");
                cmd.Parameters.AddWithValue("@PO_Link", p.PO_Link ?? "");
                cmd.Parameters.AddWithValue("@RIR_link", p.RIR_Link ?? "");
                cmd.Parameters.AddWithValue("@MPR_link", p.MPR_Link ?? "");
                cmd.ExecuteNonQuery();
            }
        }

        public void Delete(int id)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                new SqlCommand($"DELETE FROM ProjectInfo WHERE Id = {id}", conn).ExecuteNonQuery();
            }
        }

        // Lấy thống kê tổng hợp theo dự án
        public List<ProjectInfo> GetWithStats()
        {
            var list = new List<ProjectInfo>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    SELECT p.*,
                        (SELECT COUNT(*) FROM MPR_Header m WHERE m.Project_Code = p.ProjectCode) AS MPR_Count,
                        (SELECT COUNT(*) FROM PO_head po WHERE po.WorkorderNo = p.WorkorderNo)   AS PO_Count,
                        (SELECT COUNT(*) FROM RIR_head r WHERE r.WorkorderNo = p.WorkorderNo)    AS RIR_Count
                    FROM ProjectInfo p
                    ORDER BY p.CreatedDate DESC", conn);
                var r = cmd.ExecuteReader();
                while (r.Read()) list.Add(Map(r));
            }
            return list;
        }

        private ProjectInfo Map(SqlDataReader r)
        {
            return new ProjectInfo
            {
                Id = Convert.ToInt32(r["Id"]),
                ProjectName = r["ProjectName"]?.ToString() ?? "",
                ProjectCode = r["ProjectCode"]?.ToString() ?? "",
                WorkorderNo = r["WorkorderNo"]?.ToString() ?? "",
                Customer = r["Customer"]?.ToString() ?? "",
                PJWeight = r["PJWeight"] != DBNull.Value ? Convert.ToDecimal(r["PJWeight"]) : 0,
                PJBudget = r["PJBudget"] != DBNull.Value ? Convert.ToDecimal(r["PJBudget"]) : 0,
                POCode = r["POCode"]?.ToString() ?? "",
                MPRCode = r["MPRCode"]?.ToString() ?? "",
                Status = r["Status"]?.ToString() ?? "",
                Notes = r["Notes"]?.ToString() ?? "",
                PO_Link = r["PO_Link"]?.ToString() ?? "",
                RIR_Link = r["RIR_link"]?.ToString() ?? "",
                MPR_Link = r["MPR_link"]?.ToString() ?? "",
                CreatedDate = r["CreatedDate"] != DBNull.Value ? Convert.ToDateTime(r["CreatedDate"]) : null,
                ModifiedDate = r["ModifiedDate"] != DBNull.Value ? Convert.ToDateTime(r["ModifiedDate"]) : null,
                PNK_Link = r["PNK_LINK"]?.ToString() ?? "",
            };
        }
    }
}