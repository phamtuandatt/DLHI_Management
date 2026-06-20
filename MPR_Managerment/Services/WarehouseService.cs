using System;
using System.Collections.Generic;
using System.Data;
using System.Reflection.Emit;
using Microsoft.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;
using static System.Runtime.InteropServices.JavaScript.JSType;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using DataTable = System.Data.DataTable;

namespace MPR_Managerment.Services
{
    public class WarehouseService
    {

        public async Task<DataTable> GetWarehouse_ForExport_V2(string projectCode = "OVPP1")
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                string sql = $"SELECT Import_ID, ID_Code , Item_Name, Material, Size, UNIT FROM Warehouse_Import WHERE Project_Code = '{projectCode}'";
                var cmd = new SqlCommand(sql, conn);
                DataTable dt = new DataTable();
                using (SqlDataReader reader = await cmd.ExecuteReaderAsync()) // Đọc dữ liệu ngầm
                {
                    dt.Load(reader);
                }
                return dt;
            }
        }

        public async Task<DataTable> GetRIROfProject(string keyword)
        {
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                SqlCommand cmd = new SqlCommand(@"
                    SELECT RIR_ID, RIR_No
                    FROM RIR_head
                    WHERE RIR_No       LIKE @kw
                       OR Project_Name LIKE @kw
                       OR WorkorderNo  LIKE @kw
                       OR PONo         LIKE @kw
                    ORDER BY Created_Date DESC", conn);
                cmd.Parameters.AddWithValue("@kw", $"%{keyword}%");
                DataTable dt = new DataTable();
                await conn.OpenAsync(); // Mở kết nối ngầm

                using (SqlDataReader reader = await cmd.ExecuteReaderAsync()) // Đọc dữ liệu ngầm
                {
                    dt.Load(reader);
                }
                return dt;
            }
        }

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

        public async Task<DataTable> GetHistoryExportByProject(string projectCode, string fromDate, string toDate)
        {
            string sqlQuery = string.Format("SET DATEFORMAT DMY SELECT *FROM Warehouse_Export WHERE Project_Code = '{0}' AND Created_Date BETWEEN '{1}' AND '{2}'", projectCode, fromDate, toDate);
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

        public void Update(WarehouseImport wi)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    SET DATEFORMAT DMY
                    UPDATE Warehouse_Import SET
                        InvoiceNo  = @InvoiceNo,
                        InvoiceDate  = @InvoiceDate
                    WHERE Import_ID = @ImportID", conn);

                cmd.Parameters.AddWithValue("@InvoiceNo", wi.InvoiceNo);
                cmd.Parameters.AddWithValue("@InvoiceDate", Convert.ToDateTime(wi.InvoiceDate.ToString()).ToString("dd/MM/yyyy"));
                cmd.Parameters.AddWithValue("@ImportID", wi.Import_ID);

                cmd.ExecuteNonQuery();
            }
        }

        public void UpdateIDCode(WarehouseImport wi)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"UPDATE Warehouse_Import SET QC_Code = @QC_Code WHERE Import_ID = @ImportID", conn);

                cmd.Parameters.AddWithValue("@QC_Code", wi.QC_Code);
                cmd.Parameters.AddWithValue("@ImportID", wi.Import_ID);

                cmd.ExecuteNonQuery();
            }
        }

        public void UpdateMaterialDetail(int material_detail_id, string item_code_existed, string item_number)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                string sql = @$"UPDATE Material_Detail SET";
                if (!string.IsNullOrEmpty(item_code_existed))
                    sql += $" item_code_existed = '{item_code_existed}'";
                if (!string.IsNullOrEmpty(item_number))
                    sql += $" ,material_detail_number = '{item_number}'";
                sql += $" WHERE material_detail_id = {material_detail_id}";

                var cmd = new SqlCommand(sql, conn);

                cmd.ExecuteNonQuery();
            }
        }

        public void DeleteMaterialDetail(int material_detail_id)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"DELETE FROM Material_Detail WHERE material_detail_id = @material_detail_id", conn);

                cmd.Parameters.AddWithValue("@material_detail_id", material_detail_id);

                cmd.ExecuteNonQuery();
            }
        }

        public void ModifyIDCodeOfWarehouse(WarehouseImport wi)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();

                var sqlUpIdCode = $"UPDATE Warehouse_Import SET ID_Code = '{wi.ID_Code}' WHERE Import_ID = {wi.Import_ID}";
                var cmd = new SqlCommand(sqlUpIdCode, conn);
                cmd.ExecuteNonQuery();
            }
        }

        public void ModifyQtyImportOfWarehouseImport(WarehouseImport wi)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var sqlUpQtyImport = $"UPDATE Warehouse_Import SET Qty_Import = {wi.Qty_Import} WHERE Import_ID = {wi.Import_ID}";
                var cmd = new SqlCommand(sqlUpQtyImport, conn);
                cmd.ExecuteNonQuery();
            }
        }

        public void ModifyWeightOfWarehouseImport(WarehouseImport wi)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var sqlUpWeight = $"UPDATE Warehouse_Import SET Weight_kg = {wi.Weight_kg} WHERE Import_ID = {wi.Import_ID}";
                var cmd = new SqlCommand(sqlUpWeight, conn);
                cmd.ExecuteNonQuery();
            }
        }

        public void ModifySizeOfWarehouseImport(WarehouseImport wi)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var sqlUpSize = $"UPDATE Warehouse_Import SET Size = '{wi.Size}' WHERE Import_ID = {wi.Import_ID}";
                var cmd = new SqlCommand(sqlUpSize, conn);
                cmd.ExecuteNonQuery();
            }
        }

        public void ModifyNameOfWarehouseImport(WarehouseImport wi)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var sqlUpSize = $"UPDATE Warehouse_Import SET Item_Name = N'{wi.Item_Name}' WHERE Import_ID = {wi.Import_ID}";
                var cmd = new SqlCommand(sqlUpSize, conn);
                cmd.ExecuteNonQuery();
            }
        }

        public void UpdateStatusPrintPNK(WarehouseImport wi)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var sqlUpSize = $"UPDATE Warehouse_Import SET Is_Printed = 1 WHERE PO_ID = {wi.PO_ID}";
                var cmd = new SqlCommand(sqlUpSize, conn);
                cmd.ExecuteNonQuery();
            }
        }

        public async Task<DataTable> GetWarehouseImportByPOId(int poID)
        {
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                SqlCommand cmd = new SqlCommand("sp_GetDataToFillInvoce", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@POID", poID);

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

        public DataTable GetExportWarehouseForExportExcel(string fromProjectCode = null, DateTime? fromDate = null, DateTime? toDate = null)
        {
            DataTable dataTable = new DataTable();

            // Sử dụng khối using để đảm bảo kết nối và lệnh được giải phóng ngay sau khi dùng
            using (SqlConnection connection = DatabaseHelper.GetConnection())
            {
                using (SqlCommand command = new SqlCommand("[dbo].[sp_GetUnupdatedExportDetails]", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;

                    // 1. Xử lý tham số @From_Project_Code
                    command.Parameters.Add(new SqlParameter("@From_Project_Code", SqlDbType.NVarChar, 50)
                    {
                        Value = string.IsNullOrEmpty(fromProjectCode) ? (object)DBNull.Value : fromProjectCode
                    });

                    // 2. Xử lý tham số @From_Date
                    command.Parameters.Add(new SqlParameter("@From_Date", SqlDbType.Date)
                    {
                        Value = fromDate.HasValue ? (object)fromDate.Value.Date : DBNull.Value
                    });

                    // 3. Xử lý tham số @To_Date
                    command.Parameters.Add(new SqlParameter("@To_Date", SqlDbType.Date)
                    {
                        Value = toDate.HasValue ? (object)toDate.Value.Date : DBNull.Value
                    });

                    try
                    {
                        // Mở kết nối và đổ dữ liệu vào DataTable thông qua SqlDataAdapter
                        connection.Open();
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            adapter.Fill(dataTable);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Bạn có thể log lỗi ở đây hoặc throw ra ngoài tùy kiến trúc ứng dụng
                        throw new Exception("Lỗi khi thực thi Stored Procedure: " + ex.Message, ex);
                    }
                }
            }

            return dataTable;
        }

        public List<ExportDetailModel> ConvertDataTableToList(DataTable dataTable)
        {
            List<ExportDetailModel> list = new List<ExportDetailModel>();

            if (dataTable == null || dataTable.Rows.Count == 0)
                return list;

            foreach (DataRow row in dataTable.Rows)
            {
                ExportDetailModel item = new ExportDetailModel
                {
                    // Ép kiểu an toàn, kiểm tra DBNull trước khi gán dữ liệu
                    Export_Detail_Id = Convert.ToInt32(row["Export_Detail_Id"]),
                    Export_ID = Convert.ToInt32(row["Export_ID"]),

                    Export_No = row["Export_No"] == DBNull.Value ? string.Empty : row["Export_No"].ToString(),
                    Export_Number = row["Export_Number"] == DBNull.Value ? string.Empty : row["Export_Number"].ToString(),
                    From_Project_Code = row["From_Project_Code"] == DBNull.Value ? string.Empty : row["From_Project_Code"].ToString(),
                    From_Project_Name = row["From_Project_Name"] == DBNull.Value ? string.Empty : row["From_Project_Name"].ToString(),
                    To_Project_Name = row["To_Project_Name"] == DBNull.Value ? string.Empty : row["To_Project_Name"].ToString(),

                    Import_ID = Convert.ToInt32(row["Import_ID"]),
                    ID_Code = row["ID_Code"] == DBNull.Value ? string.Empty : row["ID_Code"].ToString(),
                    Item_Name = row["Item_Name"] == DBNull.Value ? string.Empty : row["Item_Name"].ToString(),
                    Material = row["Material"] == DBNull.Value ? string.Empty : row["Material"].ToString(),
                    Size = row["Size"] == DBNull.Value ? string.Empty : row["Size"].ToString(),

                    Qty_Export = row["Qty_Export"] == DBNull.Value ? (decimal?)null : Convert.ToDecimal(row["Qty_Export"]),
                    UNIT = row["UNIT"] == DBNull.Value ? string.Empty : row["UNIT"].ToString(),
                    Notes = row["Notes"] == DBNull.Value ? string.Empty : row["Notes"].ToString(),

                    Create_Date = row["Create_Date"] == DBNull.Value ? (DateTime?)null : Convert.ToDateTime(row["Create_Date"])
                };

                list.Add(item);
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
                string sql = "SELECT * FROM vw_Get_Stock_Export WHERE 1=1";
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

        public int ProjectMaterialTransform(WarehouseImport imp, string oldProjectCode, string newProjectCode, decimal qtyTransform, string note_2)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("sp_ProjectMaterialTransformTransaction", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Old_Value_Location", oldProjectCode);
                cmd.Parameters.AddWithValue("@New_Value_Location", newProjectCode);
                cmd.Parameters.AddWithValue("@Number_Tranform", qtyTransform);
                cmd.Parameters.AddWithValue("@Import_ID", imp.Import_ID);
                cmd.Parameters.AddWithValue("@Import_No", imp.Import_No ?? "");
                cmd.Parameters.AddWithValue("@Import_Date", imp.Import_Date ?? DateTime.Now);
                cmd.Parameters.AddWithValue("@PO_ID", imp.PO_ID);
                cmd.Parameters.AddWithValue("@PO_Detail_ID", imp.PO_Detail_ID);
                cmd.Parameters.AddWithValue("@Item_Name", imp.Item_Name ?? "");
                cmd.Parameters.AddWithValue("@Material", imp.Material ?? "");
                cmd.Parameters.AddWithValue("@Size", imp.Size ?? "");
                cmd.Parameters.AddWithValue("@UNIT", imp.UNIT ?? "");
                cmd.Parameters.AddWithValue("@Qty_Import", imp.Qty_Import);
                cmd.Parameters.AddWithValue("@Weight_kg", imp.Weight_kg);
                cmd.Parameters.AddWithValue("@ID_Code", imp.ID_Code ?? "");
                cmd.Parameters.AddWithValue("@Project_Code", imp.Project_Code ?? "");
                cmd.Parameters.AddWithValue("@WorkorderNo", imp.WorkorderNo ?? "");
                cmd.Parameters.AddWithValue("@Location", imp.Location ?? "");
                cmd.Parameters.AddWithValue("@Notes", imp.Notes);
                var r = cmd.ExecuteReader();

                // Insert import warehouse
                imp.Project_Code = newProjectCode;
                imp.Qty_Import = qtyTransform;
                imp.Notes = note_2;
                var rs = InsertImport(imp, "Admin");

                if (r.Read()) return Convert.ToInt32(r["NewImport_ID"]);
                return 0;
            }
        }


        public WarehouseImport GetImportByImportID(int importID)
        {
            var wh = new WarehouseImport();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand($"SELECT *FROM Warehouse_Import WHERE Import_ID = {importID}", conn);
                var r = cmd.ExecuteReader();
                while (r.Read()) wh = MapImport(r);
            }
            return wh;
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

        public void UpdateStatusExportWarehouseHeader(string status, int export_id)
        {
            string sql = $"UPDATE [dbo].[ExportWarehouseHeader] SET Status = @status, Update_By = @user, Update_Date = GETDATE() WHERE Export_ID = @export_id";

            try
            {
                using (SqlConnection conn = DatabaseHelper.GetConnection())
                {
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@status", status);
                        cmd.Parameters.AddWithValue("@user", AppSession.CurrentUser?.Full_Name ?? Environment.UserName);
                        cmd.Parameters.AddWithValue("@export_id", export_id);
                        conn.Open();
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Lỗi khi cập nhật trạng thái: {ex.Message}");
            }
        }

        public int InsertExportWarehouseHeader(ExportWarehouseDetailModel exp, string export_no)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                try
                {
                    conn.Open();
                    var cmd = new SqlCommand("sp_InsertWarehouseExportFromExportWarehouseHeader", conn);

                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add("@Import_ID", SqlDbType.Int).Value = exp.Import_Id;
                    cmd.Parameters.Add("@Qty_Export", SqlDbType.Decimal).Value = exp.Qty_Export;
                    cmd.Parameters.Add("@Notes", SqlDbType.NVarChar, 550).Value = exp.Notes;
                    cmd.Parameters.Add("@Export_No", SqlDbType.VarChar, 50).Value = export_no;

                    var r = cmd.ExecuteReader();
                    if (r.Read()) return Convert.ToInt32(r["NewExport_ID"]);
                    return 0;
                }
                catch (SqlException ex)
                {
                    throw new Exception(ex.Message);
                }
            }
        }

        public bool SaveFullInvoice(ExportWarehouseHeaderModel header, DataTable dtDetails)
        {
            // 1. Chuẩn bị DataTable đúng định dạng cấu trúc của dbo.ExportDetailType
            DataTable dtForSql = new DataTable();
            dtForSql.Columns.Add("Export_ID", typeof(int)); // Cột này có thể để trống/0 vì SQL sẽ tự ghi đè bằng ID mới
            dtForSql.Columns.Add("Import_ID", typeof(int));
            dtForSql.Columns.Add("Qty_Export", typeof(decimal));
            dtForSql.Columns.Add("Notes", typeof(string));

            foreach (DataRow row in dtDetails.Rows)
            {
                dtForSql.Rows.Add(
                    0, // Tạm thời truyền số 0, SQL Proc sẽ tự xử lý lấy ID thực tế
                    row["Import_ID"],
                    row["Qty_Export"],
                    row["Notes"]
                );
            }

            // 2. Gọi Proc thực thi
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                using (SqlCommand cmd = new SqlCommand("sp_InsertExportWarehouse_AllInOne", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    // Nạp các tham số của Header
                    cmd.Parameters.Add("@Export_No", SqlDbType.NVarChar, 50).Value = header.ExportNo;
                    cmd.Parameters.Add("@Export_Number", SqlDbType.NVarChar, 20).Value = header.ExportNumber;
                    cmd.Parameters.Add("@From_Project_ID", SqlDbType.Int).Value = header.FromProjectID;
                    cmd.Parameters.Add("@From_Project_Code", SqlDbType.NVarChar, 50).Value = header.FromProjectCode;
                    cmd.Parameters.Add("@From_Project_Name", SqlDbType.NVarChar, 250).Value = header.FromProjectName;
                    cmd.Parameters.Add("@To_Project_ID", SqlDbType.Int).Value = header.ToProjectID;
                    cmd.Parameters.Add("@To_Project_Code", SqlDbType.NVarChar, 50).Value = header.ToProjectCode;
                    cmd.Parameters.Add("@To_Project_Name", SqlDbType.NVarChar, 250).Value = header.ToProjectName;
                    cmd.Parameters.Add("@Export_Totals", SqlDbType.Decimal).Value = header.ExportTotals;
                    cmd.Parameters.Add("@Status", SqlDbType.NVarChar, 50).Value = header.Status;
                    cmd.Parameters.Add("@Notes", SqlDbType.NVarChar, -1).Value = string.IsNullOrEmpty(header.Notes) ? DBNull.Value : (object)header.Notes;
                    cmd.Parameters.Add("@Create_By", SqlDbType.NVarChar, 100).Value = header.CreateBy;

                    // Nạp tham số dạng Bảng (Detail)
                    SqlParameter tvpParam = cmd.Parameters.AddWithValue("@DetailTable", dtForSql);
                    tvpParam.SqlDbType = SqlDbType.Structured;
                    tvpParam.TypeName = "dbo.ExportDetailType";

                    try
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        return true; // Thực thi trơn tru không lỗi
                    }
                    catch (Exception ex)
                    {
                        // Nếu lỗi (ví dụ: mất mạng giữa chừng, sai dữ liệu chi tiết...), Transaction sẽ Rollback và nhảy vào đây
                        MessageBox.Show("Lỗi hệ thống khi lưu hóa đơn: " + ex.Message, "Lỗi Transaction", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
            }
        }

        public bool UpdateFullInvoice(int exportId, ExportWarehouseHeaderModel header, DataTable dtDetails)
        {
            // 1. Chuẩn bị DataTable đúng định dạng cấu trúc của dbo.ExportDetailType
            DataTable dtForSql = new DataTable();
            dtForSql.Columns.Add("Export_ID", typeof(int)); // Cột này có thể để trống/0 vì SQL sẽ tự ghi đè bằng ID mới
            dtForSql.Columns.Add("Import_ID", typeof(int));
            dtForSql.Columns.Add("Qty_Export", typeof(decimal));
            dtForSql.Columns.Add("Notes", typeof(string));

            foreach (DataRow row in dtDetails.Rows)
            {
                dtForSql.Rows.Add(
                    exportId, // Tạm thời truyền số 0, SQL Proc sẽ tự xử lý lấy ID thực tế
                    row["Import_ID"],
                    row["Qty_Export"],
                    row["Notes"]
                );
            }

            // 2. Gọi Proc thực thi
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                using (SqlCommand cmd = new SqlCommand("sp_UpdateExportWarehouse_AllInOne", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    // Nạp các tham số của Header
                    cmd.Parameters.Add("@Export_ID", SqlDbType.Int).Value = exportId;
                    cmd.Parameters.Add("@Export_No", SqlDbType.NVarChar, 50).Value = !string.IsNullOrEmpty(header.ExportNo) ? header.ExportNo : DBNull.Value;
                    cmd.Parameters.Add("@Export_Number", SqlDbType.NVarChar, 20).Value = !string.IsNullOrEmpty(header.ExportNumber) ? header.ExportNumber : DBNull.Value;
                    cmd.Parameters.Add("@From_Project_ID", SqlDbType.Int).Value = header.FromProjectID;
                    cmd.Parameters.Add("@From_Project_Code", SqlDbType.NVarChar, 50).Value = !string.IsNullOrEmpty(header.FromProjectCode) ? header.FromProjectCode : DBNull.Value;
                    cmd.Parameters.Add("@From_Project_Name", SqlDbType.NVarChar, 250).Value = !string.IsNullOrEmpty(header.FromProjectName) ? header.FromProjectName : DBNull.Value;
                    cmd.Parameters.Add("@To_Project_ID", SqlDbType.Int).Value = header.ToProjectID;
                    cmd.Parameters.Add("@To_Project_Code", SqlDbType.NVarChar, 50).Value = !string.IsNullOrEmpty(header.ToProjectCode) ? header.ToProjectCode: DBNull.Value;
                    cmd.Parameters.Add("@To_Project_Name", SqlDbType.NVarChar, 250).Value = !string.IsNullOrEmpty(header.ToProjectName) ? header.ToProjectName : DBNull.Value;
                    cmd.Parameters.Add("@Export_Totals", SqlDbType.Decimal).Value = header.ExportTotals;
                    cmd.Parameters.Add("@Status", SqlDbType.NVarChar, 50).Value = !string.IsNullOrEmpty(header.Status) ? header.Status : DBNull.Value;
                    cmd.Parameters.Add("@Notes", SqlDbType.NVarChar, -1).Value = string.IsNullOrEmpty(header.Notes) ? DBNull.Value : (object)header.Notes;
                    cmd.Parameters.Add("@Modified_By", SqlDbType.NVarChar, 100).Value = !string.IsNullOrEmpty(header.UpdateBy) ? header.UpdateBy : DBNull.Value;

                    // Nạp tham số dạng Bảng (Detail)
                    SqlParameter tvpParam = cmd.Parameters.AddWithValue("@DetailTable", dtForSql);
                    tvpParam.SqlDbType = SqlDbType.Structured;
                    tvpParam.TypeName = "dbo.ExportDetailType";

                    try
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        return true; // Thực thi trơn tru không lỗi
                    }
                    catch (Exception ex)
                    {
                        // Nếu lỗi (ví dụ: mất mạng giữa chừng, sai dữ liệu chi tiết...), Transaction sẽ Rollback và nhảy vào đây
                        MessageBox.Show("Lỗi hệ thống khi lưu hóa đơn: " + ex.Message, "Lỗi Transaction", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
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
                string sql = "SELECT * FROM vw_Warehouse_Stock_V2 WHERE 1=1";
                if (!string.IsNullOrEmpty(projectCode))
                    sql += $" AND Project_Code = N'{projectCode}'";
                if (!string.IsNullOrEmpty(keyword))
                    sql += $"AND (Material LIKE N'%{keyword}%' OR Item_Name LIKE N'%{keyword}%' OR ID_Code LIKE N'%{keyword}%' OR PONo LIKE N'%{keyword}%' OR Size LIKE N'%{keyword}%')";
                sql += " ORDER BY Import_Date DESC";
                var cmd = new SqlCommand(sql, conn);
                var r = cmd.ExecuteReader();
                while (r.Read()) list.Add(MapStock(r));
            }
            return list;
        }

        public async Task<DataTable> GetImportPaint(string projectCode = "", string keyword = "")
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                string sql = "SELECT Import_ID, Import_No, Import_Date, PO_Detail_ID, Item_Name, Material, Size, UNIT, Qty_Import, Weight_kg, ID_Code, MTRno, Heatno, QC_Code, Project_Code FROM Warehouse_Import WHERE (Item_Name LIKE '%Paint%' OR Item_Name LIKE '%Sơn%')";
                if (!string.IsNullOrEmpty(projectCode))
                    sql += $" AND Project_Code = N'{projectCode}'";
                var cmd = new SqlCommand(sql, conn);
                DataTable dt = new DataTable();
                using (SqlDataReader reader = await cmd.ExecuteReaderAsync()) // Đọc dữ liệu ngầm
                {
                    dt.Load(reader);
                }
                return dt;
            }
        }

        public async Task<DataTable> GetImportWelding(string projectCode = "", string keyword = "")
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                string sql = "SELECT Import_ID, Import_No, Import_Date, PO_Detail_ID, Item_Name, Material, Size, UNIT, Qty_Import, Weight_kg, ID_Code, MTRno, Heatno, QC_Code, Project_Code FROM Warehouse_Import WHERE (Item_Name LIKE '%Welding%' OR Item_Name LIKE '%Hàn%')";
                if (!string.IsNullOrEmpty(projectCode))
                    sql += $" AND Project_Code = N'{projectCode}'";
                var cmd = new SqlCommand(sql, conn);
                DataTable dt = new DataTable();
                using (SqlDataReader reader = await cmd.ExecuteReaderAsync()) // Đọc dữ liệu ngầm
                {
                    dt.Load(reader);
                }
                return dt;
            }
        }

        public async Task<DataTable> GetStock_V2(string projectCode = "", string keyword = "")
        {
            string sql = "SELECT * FROM vw_Warehouse_Stock WHERE 1=1";
            if (!string.IsNullOrEmpty(projectCode))
                sql += $" AND Project_Code = N'{projectCode}'";
            if (!string.IsNullOrEmpty(keyword))
                sql += $"AND (Material LIKE N'%{keyword}%' OR Item_Name LIKE N'%{keyword}%' OR ID_Code LIKE N'%{keyword}%' OR PONo LIKE N'%{keyword}%' OR Size LIKE N'%{keyword}%')";
            sql += " ORDER BY Import_Date DESC";

            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                SqlCommand cmd = new SqlCommand(sql, conn);

                DataTable dt = new DataTable();
                await conn.OpenAsync(); // Mở kết nối ngầm

                using (SqlDataReader reader = await cmd.ExecuteReaderAsync()) // Đọc dữ liệu ngầm
                {
                    dt.Load(reader);
                }
                return dt;
            }
        }

        public async Task<DataTable> GetInvoices(string projectCode = "", string keyword = "")
        {
            string sql = "EXEC dbo.usp_GetWarehouseStockReport_V3";
            //if (!string.IsNullOrEmpty(projectCode))
            //    sql += $" AND Project_Code = N'{projectCode}'";
            //if (!string.IsNullOrEmpty(keyword))
            //    sql += $"AND (Material LIKE N'%{keyword}%' OR Item_Name LIKE N'%{keyword}%' OR ID_Code LIKE N'%{keyword}%' OR PONo LIKE N'%{keyword}%' OR Size LIKE N'%{keyword}%')";
            //sql += " ORDER BY Import_Date DESC";

            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                SqlCommand cmd = new SqlCommand(sql, conn);

                DataTable dt = new DataTable();
                await conn.OpenAsync(); // Mở kết nối ngầm

                using (SqlDataReader reader = await cmd.ExecuteReaderAsync()) // Đọc dữ liệu ngầm
                {
                    dt.Load(reader);
                }
                return dt;
            }
        }

        public List<WarehouseStock> GetStockWithRemaining(string projectCode)
        {
            var list = new List<WarehouseStock>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand($"SELECT * FROM vw_Warehouse_Stock WHERE Project_Code = '{projectCode}' AND Qty_Stock > 0 ORDER BY Import_Date DESC", conn);
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
            var isPrint = !string.IsNullOrEmpty(r["Is_Printed"].ToString()) ? r["Is_Printed"] : "0";
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
                Created_Date = r["Created_Date"] != DBNull.Value ? Convert.ToDateTime(r["Created_Date"]) : null,
                IsPrint = (isPrint == "0") ? false : true
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
                Weight_Stock = r["Weight_Stock"] != DBNull.Value ? Convert.ToDecimal(r["Weight_Stock"]) : 0,
                Notes = r["Notes"]?.ToString() ?? "",

                QC_Code = r["QC_Code"]?.ToString() ?? "",
                QC_Status = r["QC_Status"]?.ToString() ?? "",
                Remarks = r["Remarks"]?.ToString() ?? ""
            };
        }

        /// <summary>
        /// Save QC Code for Item and update Quantiy of Warehouse Import by HEAT/LOT
        /// </summary>
        public async Task<bool> SaveQCCodeForItemOfWarehouseImportTable(int rir_id, int po_detail_id, decimal newQty, decimal weight_kg, string mtrNo, string heatNo, string qc_code, string qc_status, List<WarehouseImport> lstItemUpdate)
        {
                var wi = new WarehouseImport();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                
                // Get row
                string sql = $"SELECT *FROM Warehouse_Import WHERE PO_Detail_ID = {po_detail_id}";
                var cmd = new SqlCommand(sql, conn);
                var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    wi = MapImport_V2(r);
                };

                // Update Row using transaction to ensure data correct
                DataTable dtInsertList = CreateWarehouseImportStructure();
                wi.RIR_ID = rir_id;
                foreach (var item_r in lstItemUpdate)
                {
                    DataRow dr = dtInsertList.NewRow();
                    dr["Import_No"] = wi.Import_No;
                    dr["Import_Date"] = wi.Import_Date;
                    dr["PO_ID"] = wi.PO_ID;
                    dr["PO_Detail_ID"] = wi.PO_Detail_ID;
                    dr["RIR_ID"] = wi.RIR_ID != null ? Convert.ToInt32(wi.RIR_ID) : null;
                    dr["Item_Name"] = wi.Item_Name;
                    dr["Material"] = wi.Material;
                    dr["Size"] = wi.Size;
                    dr["Qty_Import"] = item_r.Qty_Import; // New
                    dr["Weight_kg"] = (item_r.Weight_kg == 0) ? wi.Weight_kg : item_r.Weight_kg; // New
                    dr["ID_Code"] = wi.ID_Code;
                    dr["MTRno"] = item_r.MTRno; // New
                    dr["Heatno"] = item_r.Heatno; // New
                    dr["Project_Code"] = wi.Project_Code;
                    dr["WorkorderNo"] = wi.WorkorderNo;
                    dr["Location"] = wi.Location;
                    dr["Notes"] = wi.Notes;
                    dr["Created_By"] = wi.Created_By;
                    //dr["Created_Date"] = DateTime.UtcNow;
                    dr["InvoiceNo"] = wi.InvoiceNo;
                    dr["InvoiceDate"] = wi.InvoiceDate;
                    dr["QC_Code"] = item_r.QC_Code; // New
                    dr["QC_Status"] = item_r.QC_Status; // New
                    dr["Is_Printed"] = true;

                    dtInsertList.Rows.Add(dr);
                }

                bool isSuccess = await ExecuteSaveChangeProcedureAsync(po_detail_id, newQty, weight_kg, mtrNo, heatNo, qc_code, qc_status, dtInsertList);
                return isSuccess;
            }
        }

        private WarehouseImport MapImport_V2(SqlDataReader r)
        {
            var isPrint = !string.IsNullOrEmpty(r["Is_Printed"].ToString()) ? r["Is_Printed"] : "0";
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
                Created_Date = r["Created_Date"] != DBNull.Value ? Convert.ToDateTime(r["Created_Date"]) : null,

                InvoiceNo = r["InvoiceNo"]?.ToString() ?? "",
                InvoiceDate = r["InvoiceDate"].ToString() ?? "",
                QC_Code = r["QC_Code"].ToString() ?? "",
                QC_Status = r["QC_Status"].ToString() ?? "",

                IsPrint = (isPrint == "0") ? false : true
            };
        }

        /// <summary>
        /// Hàm khởi tạo cấu trúc DataTable mô phỏng khớp 100% với SQL User-Defined Table Type: WarehouseImportType
        /// </summary>
        private DataTable CreateWarehouseImportStructure()
        {
            DataTable dt = new DataTable();

            // Định nghĩa các cột chính xác như trong file gemini-code-1778942575877.sql
            dt.Columns.Add("Import_No", typeof(string));
            dt.Columns.Add("Import_Date", typeof(DateTime));
            dt.Columns.Add("PO_ID", typeof(int));
            dt.Columns.Add("PO_Detail_ID", typeof(int));
            dt.Columns.Add("RIR_ID", typeof(int));
            dt.Columns.Add("Item_ID", typeof(int));
            dt.Columns.Add("Item_Name", typeof(string));
            dt.Columns.Add("Material", typeof(string));
            dt.Columns.Add("Size", typeof(string));
            dt.Columns.Add("UNIT", typeof(string));
            dt.Columns.Add("Qty_Import", typeof(decimal));
            dt.Columns.Add("Weight_kg", typeof(decimal));
            dt.Columns.Add("ID_Code", typeof(string));
            dt.Columns.Add("MTRno", typeof(string));
            dt.Columns.Add("Heatno", typeof(string));
            dt.Columns.Add("Project_Code", typeof(string));
            dt.Columns.Add("WorkorderNo", typeof(string));
            dt.Columns.Add("Location", typeof(string));
            dt.Columns.Add("Notes", typeof(string));
            dt.Columns.Add("Created_By", typeof(string));
            dt.Columns.Add("Warehouse_ID", typeof(int));
            dt.Columns.Add("Item_Code", typeof(string));
            dt.Columns.Add("InvoiceNo", typeof(string));
            dt.Columns.Add("InvoiceDate", typeof(DateTime));
            dt.Columns.Add("QC_Code", typeof(string));
            dt.Columns.Add("QC_Status", typeof(string));
            dt.Columns.Add("Is_Printed", typeof(bool));

            return dt;
        }

        /// <summary>
        /// Hàm thực thi gọi Stored Procedure xuống SQL Server bằng ADO.NET
        /// </summary>
        private async Task<bool> ExecuteSaveChangeProcedureAsync(int po_detail_id_update, decimal newQty, decimal weight_kg, string mtrNo, string heatNo, string qc_code, string qc_status, DataTable dtItems)
        {
            // Sử dụng từ khóa 'using' để tự động giải phóng tài nguyên kết nối (đóng Connection tự động)
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                using (SqlCommand cmd = new SqlCommand("sp_Warehouse_Import_SaveChange", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    // Thêm các tham số đơn lẻ (Cột đơn)
                    cmd.Parameters.Add("@Update_Import_ID", SqlDbType.Int).Value = po_detail_id_update;
                    cmd.Parameters.Add("@New_Qty_Import", SqlDbType.Decimal).Value = newQty;
                    cmd.Parameters.Add("@New_Weight_kg", SqlDbType.Decimal).Value = weight_kg;
                    cmd.Parameters.Add("@New_MTRno", SqlDbType.NVarChar, 50).Value = mtrNo;
                    cmd.Parameters.Add("@New_Heatno", SqlDbType.NVarChar, 50).Value = heatNo;
                    cmd.Parameters.Add("@New_QC_Code", SqlDbType.NVarChar, 50).Value = qc_code;
                    cmd.Parameters.Add("@New_QC_Status", SqlDbType.NVarChar, 50).Value = qc_status;

                    // QUAN TRỌNG: Thêm tham số dạng Bảng (Table-Valued Parameter)
                    SqlParameter tvpParam = cmd.Parameters.AddWithValue("@InsertItemsList", dtItems);
                    tvpParam.SqlDbType = SqlDbType.Structured; // Chỉ định kiểu dữ liệu là Structured
                    tvpParam.TypeName = "dbo.WarehouseImportType"; // Chỉ định chính xác tên Type trong cơ sở dữ liệu

                    try
                    {
                        await conn.OpenAsync();
                        await cmd.ExecuteNonQueryAsync();
                    }
                    catch (SqlException ex)
                    {

                        throw new Exception(ex.Message);
                    }
                    // Mở kết nối và thực thi bất đồng bộ

                    return true;
                }
            }
        }

        public List<TransferLog> GetTranformHistory()
        {
            var list = new List<TransferLog>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                string sql = "SELECT New_Value_Location, Item_Name, Size, Number_Tranform, Old_Value_Location FROM ProjectMaterialTransformTransaction ORDER BY New_Value_Location ASC";
                var cmd = new SqlCommand(sql, conn);
                var r = cmd.ExecuteReader();
                while (r.Read()) list.Add(MapTranferLog(r));
            }
            return list;
        }


        private TransferLog MapTranferLog(SqlDataReader r)
        {
            return new TransferLog
            {
                NewValueLocation = r["New_Value_Location"].ToString() ?? "",
                ItemName = r["Item_Name"].ToString() ?? "",
                Size= r["Size"].ToString() ?? "",
                NumberTransform = r["Number_Tranform"].ToString() ?? "",
                OldValueLocation = r["Old_Value_Location"].ToString() ?? ""
            };
        }
    }
}