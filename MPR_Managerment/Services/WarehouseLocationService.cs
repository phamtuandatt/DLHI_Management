using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;

namespace MPR_Managerment.Services
{
    public class WarehouseLocationService
    {
        // ===== GET ALL =====
        public List<Warehouse> GetAll()
        {
            var list = new List<Warehouse>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    SELECT Warehouse_ID, Warehouse_Code, Warehouse_Name,
                           Warehouse_Type, Project_Code, Dept_Abbr,
                           Manager, Notes, IsActive,
                           Created_Date, Created_By,
                           Modified_Date, Modified_By
                    FROM Warehouses
                    WHERE IsActive = 1
                    ORDER BY Project_Code, Warehouse_Type, Warehouse_Code", conn);
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) list.Add(MapWarehouse(r));
            }
            return list;
        }

        // ===== GET BY PROJECT =====
        public List<Warehouse> GetByProject(string projectCode)
        {
            var list = new List<Warehouse>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    SELECT Warehouse_ID, Warehouse_Code, Warehouse_Name,
                           Warehouse_Type, Project_Code, Dept_Abbr,
                           Manager, Notes, IsActive,
                           Created_Date, Created_By,
                           Modified_Date, Modified_By
                    FROM Warehouses
                    WHERE IsActive = 1
                      AND (Project_Code = @code OR Project_Code IS NULL)
                    ORDER BY Warehouse_Type, Warehouse_Code", conn);
                cmd.Parameters.AddWithValue("@code", projectCode);
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) list.Add(MapWarehouse(r));
            }
            return list;
        }

        // ===== GET BY TYPE =====
        public List<Warehouse> GetByType(string warehouseType)
        {
            var list = new List<Warehouse>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    SELECT Warehouse_ID, Warehouse_Code, Warehouse_Name,
                           Warehouse_Type, Project_Code, Dept_Abbr,
                           Manager, Notes, IsActive,
                           Created_Date, Created_By,
                           Modified_Date, Modified_By
                    FROM Warehouses
                    WHERE IsActive = 1
                      AND Warehouse_Type = @type
                    ORDER BY Warehouse_Code", conn);
                cmd.Parameters.AddWithValue("@type", warehouseType);
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) list.Add(MapWarehouse(r));
            }
            return list;
        }

        // ===== GET FOR COMBOBOX =====
        public DataTable GetForCombo(string projectCode = "", string warehouseType = "")
        {
            var dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Code", typeof(string));
            dt.Columns.Add("Name", typeof(string));
            dt.Rows.Add(0, "", "-- Chọn kho --");

            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                string sql = @"
                    SELECT Warehouse_ID, Warehouse_Code, Warehouse_Name
                    FROM Warehouses
                    WHERE IsActive = 1";

                if (!string.IsNullOrEmpty(projectCode))
                    sql += " AND (Project_Code = @code OR Project_Code IS NULL OR Project_Code = '')";
                if (!string.IsNullOrEmpty(warehouseType))
                    sql += " AND Warehouse_Type = @type";
                sql += " ORDER BY Warehouse_Code";

                var cmd = new SqlCommand(sql, conn);
                if (!string.IsNullOrEmpty(projectCode))
                    cmd.Parameters.AddWithValue("@code", projectCode);
                if (!string.IsNullOrEmpty(warehouseType))
                    cmd.Parameters.AddWithValue("@type", warehouseType);

                using (var r = cmd.ExecuteReader())
                    while (r.Read())
                        dt.Rows.Add(
                            Convert.ToInt32(r["Warehouse_ID"]),
                            r["Warehouse_Code"]?.ToString() ?? "",
                            r["Warehouse_Name"]?.ToString() ?? "");
            }
            return dt;
        }

        // ===== GENERATE CODE =====
        public string GenerateCode(string projectCode, string deptAbbr)
        {
            try
            {
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var cmd = new SqlCommand("sp_GenerateWarehouseCode", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@Project_Code", projectCode);
                    cmd.Parameters.AddWithValue("@Dept_Abbr", deptAbbr);
                    var r = cmd.ExecuteReader();
                    if (r.Read()) return r["Warehouse_Code"]?.ToString() ?? "";
                }
            }
            catch { }
            return $"{projectCode}-{deptAbbr}";
        }

        // ===== INSERT =====
        public int Insert(Warehouse w, string createdBy)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("sp_InsertWarehouse", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Warehouse_Code", w.Warehouse_Code);
                cmd.Parameters.AddWithValue("@Warehouse_Name", w.Warehouse_Name);
                cmd.Parameters.AddWithValue("@Warehouse_Type", w.Warehouse_Type);
                cmd.Parameters.AddWithValue("@Project_Code", w.Project_Code ?? "");
                cmd.Parameters.AddWithValue("@Dept_Abbr", w.Dept_Abbr ?? "");
                cmd.Parameters.AddWithValue("@Manager", w.Manager ?? "");
                cmd.Parameters.AddWithValue("@Notes", w.Notes ?? "");
                cmd.Parameters.AddWithValue("@Created_By", createdBy);
                var r = cmd.ExecuteReader();
                if (r.Read()) return Convert.ToInt32(r["NewWarehouse_ID"]);
                return 0;
            }
        }

        // ===== UPDATE =====
        public void Update(Warehouse w, string modifiedBy)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("sp_UpdateWarehouse", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Warehouse_ID", w.Warehouse_ID);
                cmd.Parameters.AddWithValue("@Warehouse_Name", w.Warehouse_Name);
                cmd.Parameters.AddWithValue("@Warehouse_Type", w.Warehouse_Type);
                cmd.Parameters.AddWithValue("@Manager", w.Manager ?? "");
                cmd.Parameters.AddWithValue("@Notes", w.Notes ?? "");
                cmd.Parameters.AddWithValue("@IsActive", w.IsActive);
                cmd.Parameters.AddWithValue("@Modified_By", modifiedBy);
                cmd.ExecuteNonQuery();
            }
        }

        // ===== DELETE (soft) =====
        public void Delete(int warehouseId)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                new SqlCommand(
                    $"UPDATE Warehouses SET IsActive = 0 WHERE Warehouse_ID = {warehouseId}",
                    conn).ExecuteNonQuery();
            }
        }

        // ===== MAP =====
        private Warehouse MapWarehouse(SqlDataReader r)
        {
            return new Warehouse
            {
                Warehouse_ID = Convert.ToInt32(r["Warehouse_ID"]),
                Warehouse_Code = r["Warehouse_Code"]?.ToString() ?? "",
                Warehouse_Name = r["Warehouse_Name"]?.ToString() ?? "",
                Warehouse_Type = r["Warehouse_Type"]?.ToString() ?? "",
                Project_Code = r["Project_Code"]?.ToString() ?? "",
                Dept_Abbr = r["Dept_Abbr"]?.ToString() ?? "",
                Manager = r["Manager"]?.ToString() ?? "",
                Notes = r["Notes"]?.ToString() ?? "",
                IsActive = r["IsActive"] != DBNull.Value && Convert.ToBoolean(r["IsActive"]),
                Created_Date = r["Created_Date"] != DBNull.Value ? Convert.ToDateTime(r["Created_Date"]) : null,
                Created_By = r["Created_By"]?.ToString() ?? "",
                Modified_Date = r["Modified_Date"] != DBNull.Value ? Convert.ToDateTime(r["Modified_Date"]) : null,
                Modified_By = r["Modified_By"]?.ToString() ?? ""
            };
        }
    }
}