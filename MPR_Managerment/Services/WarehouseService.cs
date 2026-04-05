using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;

namespace MPR_Managerment.Services
{
    public class WarehouseService
    {
        public async Task<bool> SaveExportList(DataTable dtSelected, string exportNo, string user)
        {
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                await conn.OpenAsync();
                foreach (DataRow row in dtSelected.Rows)
                {
                    using (SqlCommand cmd = new SqlCommand("sp_InsertWarehouseExport", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.AddWithValue("@Export_No", row["Export_No"]);
                        cmd.Parameters.AddWithValue("@Export_Date", DateTime.Now);
                        cmd.Parameters.AddWithValue("@Import_ID", row["Import_ID"]);
                        cmd.Parameters.AddWithValue("@Item_Name", row["Item_Name"]);
                        cmd.Parameters.AddWithValue("@Material", row["Material"]);
                        cmd.Parameters.AddWithValue("@Size", row["Size"]);
                        cmd.Parameters.AddWithValue("@UNIT", row["UNIT"]);
                        cmd.Parameters.AddWithValue("@Qty_Export", row["Qty_Export"]);
                        cmd.Parameters.AddWithValue("@Weight_kg", row["Weight_kg"]);
                        cmd.Parameters.AddWithValue("@ID_Code", row["ID_Code"]);
                        cmd.Parameters.AddWithValue("@Project_Code", row["Project_Code"]);
                        cmd.Parameters.AddWithValue("@WorkorderNo", row["WorkorderNo"]);
                        cmd.Parameters.AddWithValue("@Export_To", row["Export_To"] ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@Purpose", row["Purpose"] ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@Notes", row["Notes"] ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@Created_By", user);
                        cmd.Parameters.AddWithValue("@Warehouse_ID", row["Warehouse_ID"]);

                        await cmd.ExecuteNonQueryAsync();
                    }
                }
            }
            return true;
        }

        public async Task<DataTable> GetHistoryExportByProject(string projectCode)
        {
            string sqlQuery = string.Format("SELECT *FROM Warehouse_Export WHERE Project_Code = '{0}'", projectCode);
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                SqlCommand cmd = new SqlCommand(sqlQuery, conn);

                DataTable dt = new DataTable();
                await conn.OpenAsync(); // Mở kết nối ngầm

                using (SqlDataReader reader = await cmd.ExecuteReaderAsync()) // Đọc dữ liệu ngầm
                {
                    dt.Load(reader);
                }
                return dt;
            }
        }

        public async Task<DataTable> GetWarehouseImportByPOId(int poID)
        {
            string sqlQuery = string.Format("SELECT *FROM Warehouse_Import WHERE PO_ID = {0}", poID);
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                SqlCommand cmd = new SqlCommand(sqlQuery, conn);

                DataTable dt = new DataTable();
                await conn.OpenAsync(); // Mở kết nối ngầm

                using (SqlDataReader reader = await cmd.ExecuteReaderAsync()) // Đọc dữ liệu ngầm
                {
                    dt.Load(reader);
                }
                return dt;
            }
        }

        // ===== NHẬP KHO =====
        public List<WarehouseImport> GetAllImports()
        {
            var list = new List<WarehouseImport>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("SELECT * FROM Warehouse_Import ORDER BY Import_Date DESC, Import_ID DESC", conn);
                var r = cmd.ExecuteReader();
                while (r.Read()) list.Add(MapImport(r));
            }
            return list;
        }

        public List<WarehouseImport> SearchImports(string keyword, string projectCode = "")
        {
            var list = new List<WarehouseImport>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                string sql = @"SELECT * FROM Warehouse_Import WHERE 1=1";
                if (!string.IsNullOrEmpty(keyword))
                    sql += $" AND (Import_No LIKE N'%{keyword}%' OR Item_Name LIKE N'%{keyword}%' OR ID_Code LIKE N'%{keyword}%')";
                if (!string.IsNullOrEmpty(projectCode))
                    sql += $" AND Project_Code = N'{projectCode}'";
                sql += " ORDER BY Import_Date DESC";
                var cmd = new SqlCommand(sql, conn);
                var r = cmd.ExecuteReader();
                while (r.Read()) list.Add(MapImport(r));
            }
            return list;
        }

        public async Task<DataTable> GetImportRows(string keyword, int poId, string projectCode = "")
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                string sql = @$"SELECT * FROM Warehouse_Import WHERE 1=1 AND PO_ID = {poId}";
                if (!string.IsNullOrEmpty(keyword))
                    sql += $" AND (Import_No LIKE N'%{keyword}%' OR Item_Name LIKE N'%{keyword}%' OR ID_Code LIKE N'%{keyword}%')";
                //if (!string.IsNullOrEmpty(projectCode))
                //    sql += $" AND Project_Code = N'{projectCode}'";
                //sql += " ORDER BY Import_Date DESC";
                var cmd = new SqlCommand(sql, conn);
                DataTable dt = new DataTable();
                using (SqlDataReader reader = await cmd.ExecuteReaderAsync()) // Đọc dữ liệu ngầm
                {
                    dt.Load(reader);
                }
                return dt;
            }
        }

        public async Task<DataTable> GetImportForExport(string projectCode = "")
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                string sql = "SELECT * FROM Warehouse_Import WHERE 1=1";
                if (!string.IsNullOrEmpty(projectCode))
                    sql += $" AND Project_Code = N'{projectCode}'";
                sql += " ORDER BY Import_Date DESC";

                var cmd = new SqlCommand(sql, conn);
                DataTable dt = new DataTable();
                using (SqlDataReader reader = await cmd.ExecuteReaderAsync()) // Đọc dữ liệu ngầm
                {
                    dt.Load(reader);
                }
                return dt;
            }
        }

        public int InsertImport(WarehouseImport imp, string createdBy)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("sp_ImportWarehouse", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Import_No", imp.Import_No);
                cmd.Parameters.AddWithValue("@Import_Date", imp.Import_Date.HasValue ? (object)imp.Import_Date.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@PO_ID", imp.PO_ID.HasValue ? (object)imp.PO_ID.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@PO_Detail_ID", imp.PO_Detail_ID.HasValue ? (object)imp.PO_Detail_ID.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@RIR_ID", imp.RIR_ID.HasValue ? (object)imp.RIR_ID.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@Item_Name", imp.Item_Name ?? "");
                cmd.Parameters.AddWithValue("@Material", imp.Material ?? "");
                cmd.Parameters.AddWithValue("@Size", imp.Size ?? "");
                cmd.Parameters.AddWithValue("@UNIT", imp.UNIT ?? "");
                cmd.Parameters.AddWithValue("@Qty_Import", imp.Qty_Import);
                cmd.Parameters.AddWithValue("@Weight_kg", imp.Weight_kg);
                cmd.Parameters.AddWithValue("@ID_Code", imp.ID_Code ?? "");
                cmd.Parameters.AddWithValue("@MTRno", imp.MTRno ?? "");
                cmd.Parameters.AddWithValue("@Heatno", imp.Heatno ?? "");
                cmd.Parameters.AddWithValue("@Project_Code", imp.Project_Code ?? "");
                cmd.Parameters.AddWithValue("@WorkorderNo", imp.WorkorderNo ?? "");
                cmd.Parameters.AddWithValue("@Location", imp.Location ?? "");
                cmd.Parameters.AddWithValue("@Notes", imp.Notes ?? "");
                cmd.Parameters.AddWithValue("@Created_By", createdBy);
                var r = cmd.ExecuteReader();
                if (r.Read()) return Convert.ToInt32(r["NewImport_ID"]);
                return 0;
            }
        }

        public void DeleteImport(int importId)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                new SqlCommand($"DELETE FROM Warehouse_Export WHERE Import_ID = {importId}; DELETE FROM Warehouse_Import WHERE Import_ID = {importId}", conn).ExecuteNonQuery();
            }
        }

        // ===== XUẤT KHO =====
        public List<WarehouseExport> GetAllExports()
        {
            var list = new List<WarehouseExport>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("SELECT * FROM Warehouse_Export ORDER BY Export_Date DESC, Export_ID DESC", conn);
                var r = cmd.ExecuteReader();
                while (r.Read()) list.Add(MapExport(r));
            }
            return list;
        }

        public List<WarehouseExport> SearchExports(string keyword, string projectCode = "")
        {
            var list = new List<WarehouseExport>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                string sql = "SELECT * FROM Warehouse_Export WHERE 1=1";
                if (!string.IsNullOrEmpty(keyword))
                    sql += $" AND (Export_No LIKE N'%{keyword}%' OR Item_Name LIKE N'%{keyword}%' OR ID_Code LIKE N'%{keyword}%')";
                if (!string.IsNullOrEmpty(projectCode))
                    sql += $" AND Project_Code = N'{projectCode}'";
                sql += " ORDER BY Export_Date DESC";
                var cmd = new SqlCommand(sql, conn);
                var r = cmd.ExecuteReader();
                while (r.Read()) list.Add(MapExport(r));
            }
            return list;
        }

        public int InsertExport(WarehouseExport exp, string createdBy)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("sp_ExportWarehouse", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Export_No", exp.Export_No);
                cmd.Parameters.AddWithValue("@Export_Date", exp.Export_Date.HasValue ? (object)exp.Export_Date.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@Import_ID", exp.Import_ID.HasValue ? (object)exp.Import_ID.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@Qty_Export", exp.Qty_Export);
                cmd.Parameters.AddWithValue("@Weight_kg", exp.Weight_kg);
                cmd.Parameters.AddWithValue("@ID_Code", exp.ID_Code ?? "");
                cmd.Parameters.AddWithValue("@Project_Code", exp.Project_Code ?? "");
                cmd.Parameters.AddWithValue("@WorkorderNo", exp.WorkorderNo ?? "");
                cmd.Parameters.AddWithValue("@Export_To", exp.Export_To ?? "");
                cmd.Parameters.AddWithValue("@Purpose", exp.Purpose ?? "");
                cmd.Parameters.AddWithValue("@Notes", exp.Notes ?? "");
                cmd.Parameters.AddWithValue("@Created_By", createdBy);
                var r = cmd.ExecuteReader();
                if (r.Read()) return Convert.ToInt32(r["NewExport_ID"]);
                return 0;
            }
        }

        public void DeleteExport(int exportId)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                new SqlCommand($"DELETE FROM Warehouse_Export WHERE Export_ID = {exportId}", conn).ExecuteNonQuery();
            }
        }

        // ===== TỒN KHO =====
        public List<WarehouseStock> GetStock(string projectCode = "", string keyword = "")
        {
            var list = new List<WarehouseStock>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                string sql = "SELECT * FROM vw_Warehouse_Stock WHERE 1=1";
                if (!string.IsNullOrEmpty(projectCode))
                    sql += $" AND Project_Code = N'{projectCode}'";
                if (!string.IsNullOrEmpty(keyword))
                    sql += $" AND (Item_Name LIKE N'%{keyword}%' OR ID_Code LIKE N'%{keyword}%' OR PONo LIKE N'%{keyword}%')";
                sql += " ORDER BY Import_Date DESC";
                var cmd = new SqlCommand(sql, conn);
                var r = cmd.ExecuteReader();
                while (r.Read()) list.Add(MapStock(r));
            }
            return list;
        }

        public List<WarehouseStock> GetStockWithRemaining()
        {
            var list = new List<WarehouseStock>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("SELECT * FROM vw_Warehouse_Stock WHERE Qty_Stock > 0 ORDER BY Import_Date DESC", conn);
                var r = cmd.ExecuteReader();
                while (r.Read()) list.Add(MapStock(r));
            }
            return list;
        }

        // ===== IMPORT TỪ PO =====
        public List<WarehouseImport> GetPODetailsForImport(int poId)
        {
            var list = new List<WarehouseImport>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    SELECT
                        d.PO_Detail_ID, d.PO_ID, d.item_name, d.Material,
                        CAST(d.Asize AS NVARCHAR) + 'x' + CAST(d.Bsize AS NVARCHAR) + 'x' + CAST(d.Csize AS NVARCHAR) AS Size,
                        d.UNIT, d.Qty_Per_Sheet, d.Weight_kg, d.MPSNo,
                        h.WorkorderNo, h.MPR_No, h.PONo,
                        ISNULL(imp.Total_Imported, 0) AS Total_Imported,
                        d.Qty_Per_Sheet - ISNULL(imp.Total_Imported, 0) AS Remaining
                    FROM PO_Detail d
                    INNER JOIN PO_head h ON h.PO_ID = d.PO_ID
                    LEFT JOIN (
                        SELECT PO_Detail_ID, SUM(Qty_Import) AS Total_Imported
                        FROM Warehouse_Import
                        GROUP BY PO_Detail_ID
                    ) imp ON imp.PO_Detail_ID = d.PO_Detail_ID
                    WHERE d.PO_ID = @poId
                    ORDER BY d.Item_No", conn);
                cmd.Parameters.AddWithValue("@poId", poId);
                var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    list.Add(new WarehouseImport
                    {
                        PO_ID = Convert.ToInt32(r["PO_ID"]),
                        PO_Detail_ID = Convert.ToInt32(r["PO_Detail_ID"]),
                        Item_Name = r["item_name"]?.ToString() ?? "",
                        Material = r["Material"]?.ToString() ?? "",
                        Size = r["Size"]?.ToString() ?? "",
                        UNIT = r["UNIT"]?.ToString() ?? "",
                        Qty_Import = r["Remaining"] != DBNull.Value ? Convert.ToDecimal(r["Remaining"]) : 0,
                        Weight_kg = r["Weight_kg"] != DBNull.Value ? Convert.ToDecimal(r["Weight_kg"]) : 0,
                        WorkorderNo = r["WorkorderNo"]?.ToString() ?? "",
                        MTRno = r["MPSNo"]?.ToString() ?? ""
                    });
                }
            }
            return list;
        }

        private WarehouseImport MapImport(SqlDataReader r)
        {
            return new WarehouseImport
            {
                Import_ID = Convert.ToInt32(r["Import_ID"]),
                Import_No = r["Import_No"]?.ToString() ?? "",
                Import_Date = r["Import_Date"] != DBNull.Value ? Convert.ToDateTime(r["Import_Date"]) : null,
                PO_ID = r["PO_ID"] != DBNull.Value ? Convert.ToInt32(r["PO_ID"]) : null,
                PO_Detail_ID = r["PO_Detail_ID"] != DBNull.Value ? Convert.ToInt32(r["PO_Detail_ID"]) : null,
                RIR_ID = r["RIR_ID"] != DBNull.Value ? Convert.ToInt32(r["RIR_ID"]) : null,
                Item_Name = r["Item_Name"]?.ToString() ?? "",
                Material = r["Material"]?.ToString() ?? "",
                Size = r["Size"]?.ToString() ?? "",
                UNIT = r["UNIT"]?.ToString() ?? "",
                Qty_Import = r["Qty_Import"] != DBNull.Value ? Convert.ToDecimal(r["Qty_Import"]) : 0,
                Weight_kg = r["Weight_kg"] != DBNull.Value ? Convert.ToDecimal(r["Weight_kg"]) : 0,
                ID_Code = r["ID_Code"]?.ToString() ?? "",
                MTRno = r["MTRno"]?.ToString() ?? "",
                Heatno = r["Heatno"]?.ToString() ?? "",
                Project_Code = r["Project_Code"]?.ToString() ?? "",
                WorkorderNo = r["WorkorderNo"]?.ToString() ?? "",
                Location = r["Location"]?.ToString() ?? "",
                Notes = r["Notes"]?.ToString() ?? "",
                Created_By = r["Created_By"]?.ToString() ?? "",
                Created_Date = r["Created_Date"] != DBNull.Value ? Convert.ToDateTime(r["Created_Date"]) : null
            };
        }

        private WarehouseExport MapExport(SqlDataReader r)
        {
            return new WarehouseExport
            {
                Export_ID = Convert.ToInt32(r["Export_ID"]),
                Export_No = r["Export_No"]?.ToString() ?? "",
                Export_Date = r["Export_Date"] != DBNull.Value ? Convert.ToDateTime(r["Export_Date"]) : null,
                Import_ID = r["Import_ID"] != DBNull.Value ? Convert.ToInt32(r["Import_ID"]) : null,
                Item_Name = r["Item_Name"]?.ToString() ?? "",
                Material = r["Material"]?.ToString() ?? "",
                Size = r["Size"]?.ToString() ?? "",
                UNIT = r["UNIT"]?.ToString() ?? "",
                Qty_Export = r["Qty_Export"] != DBNull.Value ? Convert.ToDecimal(r["Qty_Export"]) : 0,
                Weight_kg = r["Weight_kg"] != DBNull.Value ? Convert.ToDecimal(r["Weight_kg"]) : 0,
                ID_Code = r["ID_Code"]?.ToString() ?? "",
                Project_Code = r["Project_Code"]?.ToString() ?? "",
                WorkorderNo = r["WorkorderNo"]?.ToString() ?? "",
                Export_To = r["Export_To"]?.ToString() ?? "",
                Purpose = r["Purpose"]?.ToString() ?? "",
                Notes = r["Notes"]?.ToString() ?? "",
                Created_By = r["Created_By"]?.ToString() ?? "",
                Created_Date = r["Created_Date"] != DBNull.Value ? Convert.ToDateTime(r["Created_Date"]) : null
            };
        }

        private WarehouseStock MapStock(SqlDataReader r)
        {
            return new WarehouseStock
            {
                Import_ID = Convert.ToInt32(r["Import_ID"]),
                Import_No = r["Import_No"]?.ToString() ?? "",
                Import_Date = r["Import_Date"] != DBNull.Value ? Convert.ToDateTime(r["Import_Date"]) : null,
                Item_Name = r["Item_Name"]?.ToString() ?? "",
                Material = r["Material"]?.ToString() ?? "",
                Size = r["Size"]?.ToString() ?? "",
                UNIT = r["UNIT"]?.ToString() ?? "",
                ID_Code = r["ID_Code"]?.ToString() ?? "",
                MTRno = r["MTRno"]?.ToString() ?? "",
                Heatno = r["Heatno"]?.ToString() ?? "",
                Project_Code = r["Project_Code"]?.ToString() ?? "",
                WorkorderNo = r["WorkorderNo"]?.ToString() ?? "",
                Location = r["Location"]?.ToString() ?? "",
                PONo = r["PONo"]?.ToString() ?? "",
                MPR_No = r["MPR_No"]?.ToString() ?? "",
                Qty_Import = r["Qty_Import"] != DBNull.Value ? Convert.ToDecimal(r["Qty_Import"]) : 0,
                Weight_Import = r["Weight_Import"] != DBNull.Value ? Convert.ToDecimal(r["Weight_Import"]) : 0,
                Qty_Exported = r["Qty_Exported"] != DBNull.Value ? Convert.ToDecimal(r["Qty_Exported"]) : 0,
                Weight_Exported = r["Weight_Exported"] != DBNull.Value ? Convert.ToDecimal(r["Weight_Exported"]) : 0,
                Qty_Stock = r["Qty_Stock"] != DBNull.Value ? Convert.ToDecimal(r["Qty_Stock"]) : 0,
                Weight_Stock = r["Weight_Stock"] != DBNull.Value ? Convert.ToDecimal(r["Weight_Stock"]) : 0
            };
        }
    }
}