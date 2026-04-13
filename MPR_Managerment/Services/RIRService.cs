using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;

namespace MPR_Managerment.Services
{
    public class RIRService
    {
        // ===== GET ALL =====
        public List<RIRHead> GetAll()
        {
            var list = new List<RIRHead>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    SELECT RIR_ID, RIR_No, Issue_Date, Project_Name,
                           WorkorderNo, MPR_No, Customer, PONo,
                           Created_Date, Created_By
                    FROM RIR_head
                    ORDER BY Created_Date DESC", conn);
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) list.Add(MapHead(r));
            }
            return list;
        }

        // ===== SEARCH =====
        public List<RIRHead> Search(string keyword)
        {
            var list = new List<RIRHead>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    SELECT RIR_ID, RIR_No, Issue_Date, Project_Name,
                           WorkorderNo, MPR_No, Customer, PONo,
                           Created_Date, Created_By
                    FROM RIR_head
                    WHERE RIR_No       LIKE @kw
                       OR Project_Name LIKE @kw
                       OR WorkorderNo  LIKE @kw
                       OR PONo         LIKE @kw
                    ORDER BY Created_Date DESC", conn);
                cmd.Parameters.AddWithValue("@kw", $"%{keyword}%");
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) list.Add(MapHead(r));
            }
            return list;
        }

        // ===== INSERT HEAD =====
        public int InsertHead(RIRHead h, string createdBy)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    INSERT INTO RIR_head
                        (RIR_No, Issue_Date, Project_Name, WorkorderNo,
                         MPR_No, Customer, PONo, Created_By, Created_Date)
                    VALUES
                        (@RIR_No, @Issue_Date, @Project_Name, @WorkorderNo,
                         @MPR_No, @Customer, @PONo, @Created_By, GETDATE());
                    SELECT SCOPE_IDENTITY();", conn);

                cmd.Parameters.AddWithValue("@RIR_No", h.RIR_No);
                cmd.Parameters.AddWithValue("@Issue_Date", h.Issue_Date.HasValue ? (object)h.Issue_Date.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@Project_Name", h.Project_Name ?? "");
                cmd.Parameters.AddWithValue("@WorkorderNo", h.WorkorderNo ?? "");
                cmd.Parameters.AddWithValue("@MPR_No", h.MPR_No ?? "");
                cmd.Parameters.AddWithValue("@Customer", h.Customer ?? "");
                cmd.Parameters.AddWithValue("@PONo", h.PONo ?? "");
                cmd.Parameters.AddWithValue("@Created_By", createdBy);

                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        // ===== UPDATE HEAD =====
        public void UpdateHead(RIRHead h, string modifiedBy)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    UPDATE RIR_head SET
                        RIR_No       = @RIR_No,
                        Issue_Date   = @Issue_Date,
                        Project_Name = @Project_Name,
                        WorkorderNo  = @WorkorderNo,
                        MPR_No       = @MPR_No,
                        Customer     = @Customer,
                        PONo         = @PONo,
                        Modified_By  = @Modified_By,
                        Modified_Date= GETDATE()
                    WHERE RIR_ID = @RIR_ID", conn);

                cmd.Parameters.AddWithValue("@RIR_ID", h.RIR_ID);
                cmd.Parameters.AddWithValue("@RIR_No", h.RIR_No);
                cmd.Parameters.AddWithValue("@Issue_Date", h.Issue_Date.HasValue ? (object)h.Issue_Date.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@Project_Name", h.Project_Name ?? "");
                cmd.Parameters.AddWithValue("@WorkorderNo", h.WorkorderNo ?? "");
                cmd.Parameters.AddWithValue("@MPR_No", h.MPR_No ?? "");
                cmd.Parameters.AddWithValue("@Customer", h.Customer ?? "");
                cmd.Parameters.AddWithValue("@PONo", h.PONo ?? "");
                cmd.Parameters.AddWithValue("@Modified_By", modifiedBy);
                cmd.ExecuteNonQuery();
            }
        }

        // ===== DELETE HEAD =====
        public void DeleteHead(int rirId)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                new SqlCommand($"DELETE FROM RIR_detail WHERE RIR_ID = {rirId}", conn).ExecuteNonQuery();
                new SqlCommand($"DELETE FROM RIR_head   WHERE RIR_ID = {rirId}", conn).ExecuteNonQuery();
            }
        }

        // ===== GET DETAILS =====
        public List<RIRDetail> GetDetails(int rirId)
        {
            var list = new List<RIRDetail>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    SELECT RIR_Detail_ID, RIR_ID, PO_Detail_ID, Item_No,
                           item_name, Material, Size, UNIT,
                           Qty_Per_Sheet, MTRno, Heatno, Created_Date, Qty_Required, Qty_Received, Inspect_Result, ID_Code
                    FROM RIR_detail
                    WHERE RIR_ID = @rirId
                    ORDER BY Item_No", conn);
                cmd.Parameters.AddWithValue("@rirId", rirId);
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) list.Add(MapDetail(r));
            }
            return list;
        }

        public async Task<DataTable> GetDetailsToExport(int rirId)
        {
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                var cmd = new SqlCommand(@"
                    SELECT RIR_Detail_ID, RIR_ID, PO_Detail_ID, Item_No,
                           item_name, Material, Size, UNIT,
                           Qty_Per_Sheet, MTRno, Heatno, Created_Date
                    FROM RIR_detail
                    WHERE RIR_ID = @rirId
                    ORDER BY Item_No", conn);
                cmd.Parameters.AddWithValue("@rirId", rirId);
                //cmd.Parameters.AddWithValue("@catId", cateId);

                DataTable dt = new DataTable();
                await conn.OpenAsync(); // Mở kết nối ngầm

                using (SqlDataReader reader = await cmd.ExecuteReaderAsync()) // Đọc dữ liệu ngầm
                {
                    dt.Load(reader);
                }
                return dt;
            }
        }

        // ===== INSERT DETAIL =====
        public void InsertDetail(RIRDetail d, string createdBy)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    INSERT INTO RIR_detail
                        (RIR_ID, PO_Detail_ID, Item_No, item_name, Material,
                         Size, UNIT, Qty_Per_Sheet, MTRno, Heatno, Created_Date)
                    VALUES
                        (@RIR_ID, @PO_Detail_ID, @Item_No, @item_name, @Material,
                         @Size, @UNIT, @Qty_Per_Sheet, @MTRno, @Heatno, GETDATE())", conn);

                cmd.Parameters.AddWithValue("@RIR_ID", d.RIR_ID);
                cmd.Parameters.AddWithValue("@PO_Detail_ID", d.PO_Detail_ID.HasValue ? (object)d.PO_Detail_ID.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@Item_No", d.Item_No);
                cmd.Parameters.AddWithValue("@item_name", d.Item_Name ?? "");
                cmd.Parameters.AddWithValue("@Material", d.Material ?? "");
                cmd.Parameters.AddWithValue("@Size", d.Size ?? "");
                cmd.Parameters.AddWithValue("@UNIT", d.UNIT ?? "");
                cmd.Parameters.AddWithValue("@Qty_Per_Sheet", d.Qty_Required);
                cmd.Parameters.AddWithValue("@MTRno", d.MTRno ?? "");
                cmd.Parameters.AddWithValue("@Heatno", d.Heatno ?? "");
                cmd.ExecuteNonQuery();
            }
        }

        public async Task<bool> InsertRIRDetailAndUpdateStock(RIRDetail rir)
        {
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                using (SqlCommand cmd = new SqlCommand("sp_InsertRIRDetail_UpdateStock", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    // Thêm các tham số cho Procedure
                    cmd.Parameters.AddWithValue("@RIR_ID", rir.RIR_ID);
                    cmd.Parameters.AddWithValue("@PO_Detail_ID", rir.PO_Detail_ID);
                    cmd.Parameters.AddWithValue("@Item_No", rir.Item_No);
                    cmd.Parameters.AddWithValue("@item_name", rir.Item_Name ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@Material", rir.Material ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@Size", rir.Size ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@UNIT", rir.UNIT ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@Qty_Per_Sheet", rir.Qty_Per_Sheet);
                    cmd.Parameters.AddWithValue("@MTRno", rir.MTRno ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@Heatno", rir.Heatno ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@Qty_Required", rir.Qty_Required);
                    cmd.Parameters.AddWithValue("@Qty_Received", rir.Qty_Received);
                    cmd.Parameters.AddWithValue("@Inspect_Result", rir.Inspect_Result ?? "Accept");
                    cmd.Parameters.AddWithValue("@ID_Code", rir.ID_Code ?? "");

                    try
                    {
                        await conn.OpenAsync();
                        int result = await cmd.ExecuteNonQueryAsync();

                        // Vì Procedure thực hiện cả Insert và Update nên result thường > 1
                        return result > 0;
                    }
                    catch (Exception ex)
                    {
                        // Log lỗi hoặc quăng ngoại lệ ra tầng UI
                        throw new Exception("Lỗi thực thi RIR & Update Stock: " + ex.Message);
                    }
                }
            }
        }

        // ===== UPDATE DETAIL =====
        public async Task<bool> UpdateDetail(RIRDetail d)
        {
            //using (var conn = DatabaseHelper.GetConnection())
            //{
            //    conn.Open();
            //    var cmd = new SqlCommand(@"
            //        UPDATE RIR_detail SET
            //            Item_No      = @Item_No,
            //            item_name    = @item_name,
            //            Material     = @Material,
            //            Size         = @Size,
            //            UNIT         = @UNIT,
            //            Qty_Per_Sheet= @Qty_Per_Sheet,
            //            MTRno        = @MTRno,
            //            Heatno       = @Heatno
            //        WHERE RIR_Detail_ID = @RIR_Detail_ID", conn);

            //    cmd.Parameters.AddWithValue("@RIR_Detail_ID", d.RIR_Detail_ID);
            //    cmd.Parameters.AddWithValue("@Item_No", d.Item_No);
            //    cmd.Parameters.AddWithValue("@item_name", d.Item_Name ?? "");
            //    cmd.Parameters.AddWithValue("@Material", d.Material ?? "");
            //    cmd.Parameters.AddWithValue("@Size", d.Size ?? "");
            //    cmd.Parameters.AddWithValue("@UNIT", d.UNIT ?? "");
            //    cmd.Parameters.AddWithValue("@Qty_Per_Sheet", d.Qty_Required);
            //    cmd.Parameters.AddWithValue("@MTRno", d.MTRno ?? "");
            //    cmd.Parameters.AddWithValue("@Heatno", d.Heatno ?? "");
            //    cmd.ExecuteNonQuery();
            //}

            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                using (SqlCommand cmd = new SqlCommand("sp_UpdateRIRDetail_Warehouse", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    // Thêm các tham số cho Procedure
                    //cmd.Parameters.AddWithValue("@RIR_ID", d.RIR_ID);
                    cmd.Parameters.AddWithValue("@PO_Detail_ID", d.PO_Detail_ID);
                    cmd.Parameters.AddWithValue("@Item_No", d.Item_No);
                    cmd.Parameters.AddWithValue("@item_name", d.Item_Name ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@Material", d.Material ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@Size", d.Size ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@UNIT", d.UNIT ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@Qty_Per_Sheet", d.Qty_Per_Sheet);
                    cmd.Parameters.AddWithValue("@MTRno", d.MTRno ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@Heatno", d.Heatno ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@Qty_Required", d.Qty_Required);
                    cmd.Parameters.AddWithValue("@Qty_Received", d.Qty_Received);
                    cmd.Parameters.AddWithValue("@Inspect_Result", d.Inspect_Result ?? "Accept");
                    cmd.Parameters.AddWithValue("@ID_Code", d.ID_Code ?? "");
                    cmd.Parameters.AddWithValue("@RIR_Detail_ID", d.RIR_Detail_ID);

                    try
                    {
                        await conn.OpenAsync();
                        int result = await cmd.ExecuteNonQueryAsync();

                        // Vì Procedure thực hiện cả Insert và Update nên result thường > 1
                        return result > 0;
                    }
                    catch (Exception ex)
                    {
                        // Log lỗi hoặc quăng ngoại lệ ra tầng UI
                        throw new Exception("Lỗi thực thi RIR & Update Stock: " + ex.Message);
                    }
                }
            }
        }

        // ===== DELETE DETAIL =====
        public void DeleteDetail(int detailId)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                new SqlCommand($"DELETE FROM RIR_detail WHERE RIR_Detail_ID = {detailId}", conn).ExecuteNonQuery();
            }
        }

        // ===== MAP =====
        private RIRHead MapHead(SqlDataReader r)
        {
            return new RIRHead
            {
                RIR_ID = Convert.ToInt32(r["RIR_ID"]),
                RIR_No = r["RIR_No"]?.ToString() ?? "",
                Issue_Date = r["Issue_Date"] != DBNull.Value ? Convert.ToDateTime(r["Issue_Date"]) : null,
                Project_Name = r["Project_Name"]?.ToString() ?? "",
                WorkorderNo = r["WorkorderNo"]?.ToString() ?? "",
                MPR_No = r["MPR_No"]?.ToString() ?? "",
                Customer = r["Customer"]?.ToString() ?? "",
                PONo = r["PONo"]?.ToString() ?? "",
                Created_Date = r["Created_Date"] != DBNull.Value ? Convert.ToDateTime(r["Created_Date"]) : null,
                Created_By = r["Created_By"]?.ToString() ?? ""
            };
        }

        private RIRDetail MapDetail(SqlDataReader r)
        {
            return new RIRDetail
            {
                RIR_Detail_ID = Convert.ToInt32(r["RIR_Detail_ID"]),
                RIR_ID = Convert.ToInt32(r["RIR_ID"]),
                PO_Detail_ID = r["PO_Detail_ID"] != DBNull.Value ? Convert.ToInt32(r["PO_Detail_ID"]) : null,
                Item_No = r["Item_No"] != DBNull.Value ? Convert.ToInt32(r["Item_No"]) : 0,
                Item_Name = r["item_name"]?.ToString() ?? "",
                Material = r["Material"]?.ToString() ?? "",
                Size = r["Size"]?.ToString() ?? "",
                UNIT = r["UNIT"]?.ToString() ?? "",
                Qty_Required = r["Qty_Per_Sheet"] != DBNull.Value ? Convert.ToInt32(r["Qty_Per_Sheet"]) : 0,
                MTRno = r["MTRno"]?.ToString() ?? "",
                Heatno = r["Heatno"]?.ToString() ?? "",
                Created_Date = r["Created_Date"] != DBNull.Value ? Convert.ToDateTime(r["Created_Date"]) : null,
                Inspect_Result = r["Inspect_Result"].ToString() ?? "",
                ID_Code = r["ID_Code"].ToString() ?? "",
                Qty_Received = r["Qty_Received"] != DBNull.Value ? Convert.ToInt32(r["Qty_Received"]) : 0
            };
        }
    }
}