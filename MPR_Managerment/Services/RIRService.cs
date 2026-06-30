using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using System.Windows.Forms;

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
                         MPR_No, Customer, PONo, Created_By, Created_Date, ProjectCode)
                    VALUES
                        (@RIR_No, @Issue_Date, @Project_Name, @WorkorderNo,
                         @MPR_No, @Customer, @PONo, @Created_By, GETDATE(), @ProjectCode);
                    SELECT SCOPE_IDENTITY();", conn);

                cmd.Parameters.AddWithValue("@RIR_No", h.RIR_No);
                cmd.Parameters.AddWithValue("@Issue_Date", h.Issue_Date.HasValue ? (object)h.Issue_Date.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@Project_Name", h.Project_Name ?? "");
                cmd.Parameters.AddWithValue("@WorkorderNo", h.WorkorderNo ?? "");
                cmd.Parameters.AddWithValue("@MPR_No", h.MPR_No ?? "");
                cmd.Parameters.AddWithValue("@Customer", h.Customer ?? "");
                cmd.Parameters.AddWithValue("@PONo", h.PONo ?? "");
                cmd.Parameters.AddWithValue("@Created_By", createdBy);
                cmd.Parameters.AddWithValue("@ProjectCode", h.ProjectCode);

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
                           Qty_Per_Sheet, MTRno, Heatno, Created_Date, Qty_Required, Qty_Received, Inspect_Result, ID_Code, Remarks
                    FROM RIR_detail
                    WHERE RIR_ID = @rirId
                    ORDER BY Item_No", conn);
                cmd.Parameters.AddWithValue("@rirId", rirId);
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) list.Add(MapDetail(r));
            }
            return list;
        }

        // ===== GET DETAILS =====
        public async Task<DataTable> GetRIRMaterialDetail(string projectCode = null,
            string rirNo = null,
            string projectName = null,
            string workorderNo = null,
            string poNo = null)
        {
            DataTable dt = new DataTable();

            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                using (SqlCommand cmd = new SqlCommand("[dbo].[sp_GetRIRDetailsByProjectCode]", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    // Gán các tham số điều kiện (Nếu chuỗi trống hoặc null -> truyền DBNull.Value để SQL xử lý logic mặc định)
                    cmd.Parameters.AddWithValue("@ProjectCode", string.IsNullOrWhiteSpace(projectCode) ? (object)DBNull.Value : projectCode.Trim());
                    cmd.Parameters.AddWithValue("@RIR_No", string.IsNullOrWhiteSpace(rirNo) ? (object)DBNull.Value : rirNo.Trim());
                    cmd.Parameters.AddWithValue("@Project_Name", string.IsNullOrWhiteSpace(projectName) ? (object)DBNull.Value : projectName.Trim());
                    cmd.Parameters.AddWithValue("@WorkorderNo", string.IsNullOrWhiteSpace(workorderNo) ? (object)DBNull.Value : workorderNo.Trim());
                    cmd.Parameters.AddWithValue("@PONo", string.IsNullOrWhiteSpace(poNo) ? (object)DBNull.Value : poNo.Trim());

                    // Sử dụng SqlDataAdapter để tự động mở/đóng kết nối và nạp dữ liệu vào DataTable
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        adapter.Fill(dt);
                    }
                }
            }

            return dt;
        }



        public async Task<DataTable> GetPaintDetails(string projectCode = "")
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                string sql = @"SELECT 
                        d.RIR_Detail_ID, 
                        d.RIR_ID, 
                        d.PO_Detail_ID, 
                        d.Item_No,
                        d.item_name, 
                        d.Material, 
                        d.Size, 
                        d.UNIT,
                        d.Qty_Per_Sheet, 
                        d.MTRno, 
                        d.Heatno, 
                        d.Created_Date, 
                        d.Qty_Required, 
                        d.Qty_Received, 
                        d.Inspect_Result, 
                        d.ID_Code, 
                        d.Remarks,
                        p.ProjectCode,
                        h.RIR_No
                    FROM dbo.RIR_detail d
                    INNER JOIN dbo.RIR_head h ON d.RIR_ID = h.RIR_ID
                    INNER JOIN dbo.PO_head p ON h.PONo = p.PONo
                    WHERE (d.item_name LIKE '%Paint%' OR d.item_name LIKE '%Sơn%')";
                if (!string.IsNullOrEmpty(projectCode))
                {
                    sql += $" AND p.ProjectCode = '{projectCode}'";
                }

                var cmd = new SqlCommand(sql, conn);
                DataTable dt = new DataTable();
                using (SqlDataReader reader = await cmd.ExecuteReaderAsync()) // Đọc dữ liệu ngầm
                {
                    dt.Load(reader);
                }
                return dt;
            }
        }

        public async Task<DataTable> GetRIRNoPaint()
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                const string sql = @"
                    SELECT DISTINCT h.RIR_ID, h.RIR_No
                    FROM dbo.RIR_head h
                    INNER JOIN dbo.PO_head p ON h.PONo = p.PONo
                    INNER JOIN dbo.RIR_detail d ON d.RIR_ID = h.RIR_ID
                    WHERE (d.item_name LIKE '%Paint%' OR d.item_name LIKE N'%Sơn%')
                    ORDER BY h.RIR_No";

                var cmd = new SqlCommand(sql, conn);
                DataTable dt = new DataTable();
                await conn.OpenAsync();

                using (SqlDataReader reader = await cmd.ExecuteReaderAsync())
                {
                    dt.Load(reader);
                }

                return dt;
            }
        }

        public async Task<DataTable> GetRIRNoWelding()
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                const string sql = @"
                    SELECT DISTINCT h.RIR_ID, h.RIR_No
                    FROM dbo.RIR_head h
                    INNER JOIN dbo.PO_head p ON h.PONo = p.PONo
                    INNER JOIN dbo.RIR_detail d ON d.RIR_ID = h.RIR_ID
                    WHERE (d.item_name LIKE '%Welding%' OR d.item_name LIKE N'%Hàn%')
                    ORDER BY h.RIR_No";

                var cmd = new SqlCommand(sql, conn);
                DataTable dt = new DataTable();
                await conn.OpenAsync();

                using (SqlDataReader reader = await cmd.ExecuteReaderAsync())
                {
                    dt.Load(reader);
                }

                return dt;
            }
        }

        public async Task<DataTable> GetRIRHeaderForPaintOrWelding(bool isPaint = false, bool isWelding = false)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                var sql = @"
                    SELECT DISTINCT h.RIR_No
                    FROM dbo.RIR_head h
                    INNER JOIN dbo.PO_head p ON h.PONo = p.PONo
                    INNER JOIN dbo.RIR_detail d ON d.RIR_ID = h.RIR_ID
                    WHERE (d.item_name LIKE '%{0}%' OR d.item_name LIKE N'%{1}%')
                    ORDER BY h.RIR_No";

                if (isPaint)
                {
                    sql = string.Format(sql, "Paint", "Sơn");
                }
                if (isWelding)
                {
                    sql = string.Format(sql, "Welding", "Hàn");
                }

                var cmd = new SqlCommand(sql, conn);
                DataTable dt = new DataTable();
                await conn.OpenAsync();

                using (SqlDataReader reader = await cmd.ExecuteReaderAsync())
                {
                    dt.Load(reader);
                }

                return dt;
            }
        }

        public async Task<DataTable> GetWeldingDetails(string projectCode = "")
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                string sql = @"SELECT 
                        d.RIR_Detail_ID, 
                        d.RIR_ID, 
                        d.PO_Detail_ID, 
                        d.Item_No,
                        d.item_name, 
                        d.Material, 
                        d.Size, 
                        d.UNIT,
                        d.Qty_Per_Sheet, 
                        d.MTRno, 
                        d.Heatno, 
                        d.Created_Date, 
                        d.Qty_Required, 
                        d.Qty_Received, 
                        d.Inspect_Result, 
                        d.ID_Code, 
                        d.Remarks,
                        p.ProjectCode,
                        h.RIR_No
                    FROM dbo.RIR_detail d
                    INNER JOIN dbo.RIR_head h ON d.RIR_ID = h.RIR_ID
                    INNER JOIN dbo.PO_head p ON h.PONo = p.PONo
                    WHERE (d.item_name LIKE '%Welding%' OR d.item_name LIKE '%Hàn%')";
                if (!string.IsNullOrEmpty(projectCode))
                {
                    sql += $" AND p.ProjectCode = '{projectCode}'";
                }

                var cmd = new SqlCommand(sql, conn);
                DataTable dt = new DataTable();
                using (SqlDataReader reader = await cmd.ExecuteReaderAsync()) // Đọc dữ liệu ngầm
                {
                    dt.Load(reader);
                }
                return dt;
            }
        }

        public async Task<DataTable> GetDetailsToExport(int rirId)
        {
            //using (SqlConnection conn = new SqlConnection(connectionString))
            //{
            //    conn.Open();
            //    // Lấy thông tin dự án (Header)
            //    using (SqlCommand cmd = new SqlCommand(headerQuery, conn))
            //    {
            //        cmd.Parameters.AddWithValue("@ProjectId", projectId);
            //        using (SqlDataAdapter da = new SqlDataAdapter(cmd)) { da.Fill(dtHeader); }
            //    }
            //    // Lấy danh sách vật tư chi tiết (Details)
            //    using (SqlCommand cmd = new SqlCommand(detailsQuery, conn))
            //    {
            //        cmd.Parameters.AddWithValue("@ProjectId", projectId);
            //        using (SqlDataAdapter da = new SqlDataAdapter(cmd)) { da.Fill(dtDetails); }
            //    }
            //}
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

        public async Task<RIRHead> GetRIRHeaderByRIRId(int rirId)
        {
            var head = new RIRHead();
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                var cmd = new SqlCommand(@"SELECT TOP 1 *FROM RIR_head WHERE RIR_ID = @rirId", conn);
                cmd.Parameters.AddWithValue("@rirId", rirId);
                await conn.OpenAsync(); // Mở kết nối ngầm

                using (var r = cmd.ExecuteReader())
                    while (r.Read()) head = MapHead(r);
                return head;
            }
        }

        public async Task<DataTable> GetMaterialIDCodeListOfProjectForQC(string project_code = "")
        {
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                string sql = "SELECT Qty_Per_Sheet, Size, Material, Heatno, MTRno, ID_Code FROM RIR_detail INNER JOIN RIR_head on RIR_detail.RIR_ID = RIR_head.RIR_ID WHERE 1=1";
                if (!string.IsNullOrEmpty(project_code))
                    sql += $"AND RIR_head.Project_Name LIKE '%{project_code}%' OR RIR_head.WorkorderNo LIKE '%{project_code}%' OR RIR_head.MPR_No LIKE '%{project_code}%' ";

                var cmd = new SqlCommand(sql, conn);

                DataTable dt = new DataTable();
                await conn.OpenAsync(); // Mở kết nối ngầm

                using (SqlDataReader reader = await cmd.ExecuteReaderAsync()) // Đọc dữ liệu ngầm
                {
                    dt.Load(reader);
                }
                return dt;
            }
        }

        public async Task<DataTable> GetMaterialIDCodeListOfRIRForQC(int rir_id)
        {
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                string sql = $"SELECT Qty_Per_Sheet, Size, Material, Heatno, MTRno, ID_Code FROM RIR_detail INNER JOIN RIR_head on RIR_detail.RIR_ID = RIR_head.RIR_ID WHERE RIR_detail.RIR_ID = {rir_id}";
                var cmd = new SqlCommand(sql, conn);

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
                         Size, UNIT, Qty_Per_Sheet, MTRno, Heatno, Created_Date, Qty_Required, Inspect_Result, ID_Code)
                    VALUES
                        (@RIR_ID, @PO_Detail_ID, @Item_No, @item_name, @Material,
                         @Size, @UNIT, @Qty_Per_Sheet, @MTRno, @Heatno, GETDATE(), @Qty_Required, @Inspect_Result, @ID_Code)", conn);

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
                
                cmd.Parameters.AddWithValue("@Qty_Required", d.Qty_Required);
                cmd.Parameters.AddWithValue("@Inspect_Result", d.Inspect_Result ?? "");
                cmd.Parameters.AddWithValue("@ID_Code", d.ID_Code ?? "");
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

        public async Task<int> UpdateDetailForQC(RIRDetail d)
        {
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                using (SqlCommand cmd = new SqlCommand("[dbo].[sp_SaveRIRDetailForQC]", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    // Hàm bổ trợ xử lý dữ liệu trống thành DBNull.Value
                    object GetDbValue(object value) => value == null || value == DBNull.Value ? DBNull.Value : value;

                    // Ánh xạ an toàn các trường dữ liệu từ RIRDetail object vào Parameter của SQL
                    cmd.Parameters.AddWithValue("@RIR_Detail_ID", GetDbValue(d.RIR_Detail_ID));
                    cmd.Parameters.AddWithValue("@RIR_ID", GetDbValue(d.RIR_ID));
                    cmd.Parameters.AddWithValue("@PO_Detail_ID", GetDbValue(d.PO_Detail_ID));
                    cmd.Parameters.AddWithValue("@Item_No", GetDbValue(d.Item_No));
                    cmd.Parameters.AddWithValue("@Item_Name", GetDbValue(d.Item_Name));
                    cmd.Parameters.AddWithValue("@Material", GetDbValue(d.Material));
                    cmd.Parameters.AddWithValue("@Size", GetDbValue(d.Size));
                    cmd.Parameters.AddWithValue("@UNIT", GetDbValue(d.UNIT));
                    cmd.Parameters.AddWithValue("@Qty_Per_Sheet", GetDbValue(d.Qty_Per_Sheet));
                    cmd.Parameters.AddWithValue("@MTRno", GetDbValue(d.MTRno));
                    cmd.Parameters.AddWithValue("@Heatno", GetDbValue(d.Heatno));
                    cmd.Parameters.AddWithValue("@Qty_Required", GetDbValue(d.Qty_Required));
                    cmd.Parameters.AddWithValue("@Qty_Received", GetDbValue(d.Qty_Received));
                    cmd.Parameters.AddWithValue("@Inspect_Result", GetDbValue(d.Inspect_Result));
                    cmd.Parameters.AddWithValue("@ID_Code", GetDbValue(d.ID_Code));
                    cmd.Parameters.AddWithValue("@Remarks", GetDbValue(d.Remarks));

                    await conn.OpenAsync();
                    // Thực thi và lấy về ID của bản ghi vừa thao tác thành công
                    object result = await cmd.ExecuteScalarAsync();
                    return result != null ? Convert.ToInt32(result) : 0;
                }
            }
        }

        public void UpdateIDFormaterialForQC(WarehouseImport warehouseImport)
        {
            var targetPoDetailId = warehouseImport.PO_Detail_ID;        // ID chi tiết đơn hàng cần cập nhật QC
            var qcCode = warehouseImport.QC_Code;     // Mã định danh kiểm định chất lượng
            var qcStatus = warehouseImport.QC_Status;         // Trạng thái kiểm định chất lượng
            var currentUser = AppSession.CurrentUser?.Full_Name ?? AppSession.CurrentDepartment; // Tài khoản kỹ sư QC thực hiện thay đổi

            Console.WriteLine("=== TIẾN TRÌNH CẬP NHẬT TRẠNG THÁI QC TRÊN HỆ THỐNG KHO ===");

            // 3. Khởi tạo và thực thi kết nối dưới khối using an toàn
            using (SqlConnection connection = DatabaseHelper.GetConnection())
            {
                // Gọi đích danh Stored Procedure xử lý cập nhật theo tập hợp (Set-based)
                using (SqlCommand command = new SqlCommand("dbo.sp_UpdateWarehouseQCStatus", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;

                    // 🎯 KHAI BÁO CHI TIẾT PARAMETER (Khớp 100% kiểu dữ liệu và độ dài trong cấu trúc bảng gốc)
                    command.Parameters.Add("@PO_Detail_ID", SqlDbType.Int).Value = targetPoDetailId;

                    // Xử lý an toàn cho các trường cho phép NULL (Chuỗi rỗng sẽ được gán thành DBNull.Value)
                    command.Parameters.Add("@QC_Code", SqlDbType.VarChar, 50).Value =
                        string.IsNullOrEmpty(qcCode) ? DBNull.Value : qcCode;

                    command.Parameters.Add("@QC_Status", SqlDbType.VarChar, 20).Value =
                        string.IsNullOrEmpty(qcStatus) ? DBNull.Value : qcStatus;

                    command.Parameters.Add("@Modified_By", SqlDbType.NVarChar, 100).Value =
                        string.IsNullOrEmpty(currentUser) ? DBNull.Value : currentUser;

                    try
                    {
                        // Mở băng thông kết nối đến SQL Server
                        connection.Open();
                        Console.WriteLine($"[Hệ thống] Đang truyền lệnh cập nhật thông tin QC cho PO_Detail_ID: {targetPoDetailId}...");

                        // Thực thi câu lệnh tác động ghi (Write I/O) xuống đĩa cứng
                        int rowsAffected = command.ExecuteNonQuery();

                        Console.WriteLine("🎉 [Thành công] Trạng thái và mã định danh QC đã được đồng bộ đồng loạt xuống bảng Warehouse_Import!");
                    }
                    catch (SqlException ex)
                    {
                        // 🛠 BẪY LỖI CHUYÊN SÂU TỪ SQL SERVER (Bắt các lỗi RAISERROR từ Proc ném lên)
                        Console.WriteLine($"❌ [Lỗi SQL Server - Số lỗi {ex.Number}]: {ex.Message}");
                        Console.WriteLine($"Mức độ nghiêm trọng (Severity): {ex.Class}, Trạng thái lỗi (State): {ex.State}");
                    }
                    catch (Exception ex)
                    {
                        // Bẫy các lỗi phát sinh ở môi trường .NET Runtime hệ thống
                        Console.WriteLine($"❌ [Lỗi ứng dụng C#]: {ex.Message}");
                    }
                }
            }
        }

        public void UpdateIDFormaterial(List<int> poDetailIdList)
        {
            // 2. Tham số truyền vào hệ thống (Ví dụ: ID chi tiết đơn đặt hàng cần đồng bộ)
            Console.WriteLine("=== BẮT ĐẦU QUY TRÌNH ĐỒNG BỘ DỮ LIỆU RIR SANG KHO NHẬP ===");

            foreach (var item in poDetailIdList)
            {
                int targetPoDetailId = item;
                string currentUser = AppSession.CurrentUser.Full_Name;
                int currentWarehouseId = 0; // ID kho thực tế tại nhà máy

                // 3. Thực thi kết nối dưới khối using để tự động đóng kết nối (Connection Closure) ngay cả khi xảy ra exception
                using (SqlConnection connection = DatabaseHelper.GetConnection())
                {
                    // Gọi đích danh Stored Procedure xử lý xóa và nạp lại theo tập hợp (Set-based)
                    using (SqlCommand command = new SqlCommand("dbo.sp_SyncRIRToWarehouseImport", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        // 🎯 ĐỊNH NGHĨA CHÍNH XÁC THAM SỐ ĐẦU VÀO (Khớp 100% kiểu dữ liệu dưới DB)
                        command.Parameters.Add("@PO_Detail_ID", SqlDbType.Int).Value = targetPoDetailId;
                        command.Parameters.Add("@Created_By", SqlDbType.NVarChar, 100).Value = currentUser;

                        // Xử lý tham số cho phép NULL: Nếu truyền số <= 0 thì gán DBNull.Value
                        if (currentWarehouseId > 0)
                            command.Parameters.Add("@Warehouse_ID", SqlDbType.Int).Value = currentWarehouseId;
                        else
                            command.Parameters.Add("@Warehouse_ID", SqlDbType.Int).Value = DBNull.Value;

                        try
                        {
                            // Mở băng thông kết nối đến máy chủ cơ sở dữ liệu
                            connection.Open();
                            Console.WriteLine($"[Hệ thống] Đang xóa dữ liệu kho cũ và nạp lại từ biên bản RIR cho PO_Detail_ID: {targetPoDetailId}...");

                            // Thực thi câu lệnh tác động ghi (Write I/O) xuống đĩa cứng
                            int rowsAffected = command.ExecuteNonQuery();

                            Console.WriteLine("🎉 [Thành công] Toàn bộ dữ liệu RIR đã được đồng bộ sang Warehouse_Import hoàn tất!");
                        }
                        catch (SqlException ex)
                        {
                            // 🛠 BẪY LỖI CHUYÊN SÂU TỪ SQL SERVER (Ví dụ: Lỗi RAISERROR từ Proc ném lên)
                            Console.WriteLine($"❌ [Lỗi SQL Server - Mã lỗi {ex.Number}]: {ex.Message}");
                            Console.WriteLine($"Mức độ nghiêm trọng (Severity): {ex.Class}, Trạng thái (State): {ex.State}");
                        }
                        catch (Exception ex)
                        {
                            // Bẫy các lỗi hệ thống khác ở tầng ứng dụng (.NET Runtime Error)
                            Console.WriteLine($"❌ [Lỗi ứng dụng C#]: {ex.Message}");
                        }
                    }
                }
                Console.WriteLine("=== KẾT THÚC TIẾN TRÌNH ===");
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
                Qty_Received = r["Qty_Received"] != DBNull.Value ? Convert.ToInt32(r["Qty_Received"]) : 0,
                Remarks = r["Remarks"].ToString() ?? "",
            };
        }
    }
}