using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;

namespace MPR_Managerment.Services
{
    public class MPRService
    {
        // ===== GET ALL =====
        public List<MPRHeader> GetAll()
        {
            var list = new List<MPRHeader>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    SELECT MPR_ID, MPR_No, Project_Name, Project_Code,
                           Department, Requestor, Rev, Required_Date,
                           Status, Total_Amount, Notes, Created_Date, Created_By
                    FROM MPR_Header
                    ORDER BY Created_Date DESC", conn);
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) list.Add(MapHeader(r));
            }
            return list;
        }

        // ===== SEARCH =====
        public List<MPRHeader> Search(string keyword)
        {
            var list = new List<MPRHeader>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    SELECT MPR_ID, MPR_No, Project_Name, Project_Code,
                           Department, Requestor, Rev, Required_Date,
                           Status, Total_Amount, Notes, Created_Date, Created_By
                    FROM MPR_Header
                    WHERE MPR_No LIKE @kw OR Project_Name LIKE @kw OR Project_Code LIKE @kw
                    ORDER BY Created_Date DESC", conn);
                cmd.Parameters.AddWithValue("@kw", $"%{keyword}%");
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) list.Add(MapHeader(r));
            }
            return list;
        }

        // ===== INSERT HEADER =====
        public int InsertHeader(MPRHeader m, string createdBy)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    INSERT INTO MPR_Header
                        (MPR_No, Project_Name, Project_Code, Department, Requestor,
                         Rev, Required_Date, Status, Notes, Created_By, Created_Date)
                    VALUES
                        (@MPR_No, @Project_Name, @Project_Code, @Department, @Requestor,
                         @Rev, @Required_Date, @Status, @Notes, @Created_By, GETDATE());
                    SELECT SCOPE_IDENTITY();", conn);

                cmd.Parameters.AddWithValue("@MPR_No", m.MPR_No);
                cmd.Parameters.AddWithValue("@Project_Name", m.Project_Name ?? "");
                cmd.Parameters.AddWithValue("@Project_Code", m.Project_Code ?? "");
                cmd.Parameters.AddWithValue("@Department", m.Department ?? "");
                cmd.Parameters.AddWithValue("@Requestor", m.Requestor ?? "");
                cmd.Parameters.AddWithValue("@Rev", m.Rev);
                cmd.Parameters.AddWithValue("@Required_Date", m.Required_Date.HasValue
                    ? (object)m.Required_Date.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@Status", m.Status ?? "Mới");
                cmd.Parameters.AddWithValue("@Notes", m.Notes ?? "");
                cmd.Parameters.AddWithValue("@Created_By", createdBy);

                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        // ===== UPDATE HEADER =====
        public void UpdateHeader(MPRHeader m, string modifiedBy)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    UPDATE MPR_Header SET
                        MPR_No        = @MPR_No,
                        Project_Name  = @Project_Name,
                        Project_Code  = @Project_Code,
                        Department    = @Department,
                        Requestor     = @Requestor,
                        Rev           = @Rev,
                        Required_Date = @Required_Date,
                        Status        = @Status,
                        Notes         = @Notes
                    WHERE MPR_ID = @MPR_ID", conn);

                cmd.Parameters.AddWithValue("@MPR_ID", m.MPR_ID);
                cmd.Parameters.AddWithValue("@MPR_No", m.MPR_No);
                cmd.Parameters.AddWithValue("@Project_Name", m.Project_Name ?? "");
                cmd.Parameters.AddWithValue("@Project_Code", m.Project_Code ?? "");
                cmd.Parameters.AddWithValue("@Department", m.Department ?? "");
                cmd.Parameters.AddWithValue("@Requestor", m.Requestor ?? "");
                cmd.Parameters.AddWithValue("@Rev", m.Rev);
                cmd.Parameters.AddWithValue("@Required_Date", m.Required_Date.HasValue
                    ? (object)m.Required_Date.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@Status", m.Status ?? "Mới");
                cmd.Parameters.AddWithValue("@Notes", m.Notes ?? "");
                cmd.ExecuteNonQuery();
            }
        }

        // ===== DELETE MPR =====
        public void DeleteMPR(int mprId)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                new SqlCommand($"DELETE FROM MPR_Details WHERE MPR_ID = {mprId}", conn).ExecuteNonQuery();
                new SqlCommand($"DELETE FROM MPR_Header  WHERE MPR_ID = {mprId}", conn).ExecuteNonQuery();
            }
        }

        // ===== GET DETAILS =====
        // Trả về TẤT CẢ dòng bao gồm cả Is_Deleted=1 để hiển thị lịch sử
        public List<MPRDetail> GetDetails(int mprId)
        {
            var list = new List<MPRDetail>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    SELECT Detail_ID, MPR_ID, Item_No, Item_Name, Description,
                           Material, Thickness_mm, Depth_mm, C_Width_mm,
                           D_Web_mm, E_Flange_mm, F_Length_mm,
                           UNIT, Qty_Per_Sheet, Weight_kg,
                           MPS_Info, Usage_Location, REV, Remarks, Is_Deleted
                    FROM MPR_Details
                    WHERE MPR_ID = @mprId
                    ORDER BY Item_No", conn);
                cmd.Parameters.AddWithValue("@mprId", mprId);
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) list.Add(MapDetail(r));
            }
            return list;
        }

        // ===== GET DETAILS — CHỈ DÒNG ACTIVE (dùng cho PO, báo cáo, tính toán) =====
        public List<MPRDetail> GetActiveDetails(int mprId)
        {
            var list = new List<MPRDetail>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    SELECT Detail_ID, MPR_ID, Item_No, Item_Name, Description,
                           Material, Thickness_mm, Depth_mm, C_Width_mm,
                           D_Web_mm, E_Flange_mm, F_Length_mm,
                           UNIT, Qty_Per_Sheet, Weight_kg,
                           MPS_Info, Usage_Location, REV, Remarks, Is_Deleted
                    FROM MPR_Details
                    WHERE MPR_ID = @mprId
                      AND Is_Deleted = 0
                    ORDER BY Item_No", conn);
                cmd.Parameters.AddWithValue("@mprId", mprId);
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) list.Add(MapDetail(r));
            }
            return list;
        }

        // ===== INSERT DETAIL =====
        public void InsertDetail(MPRDetail d, string createdBy)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    INSERT INTO MPR_Details
                        (MPR_ID, Item_No, Item_Name, Description, Material,
                         Thickness_mm, Depth_mm, C_Width_mm, D_Web_mm, E_Flange_mm, F_Length_mm,
                         UNIT, Qty_Per_Sheet, Weight_kg, MPS_Info, Usage_Location, REV, Remarks,
                         Is_Deleted, Created_By, Created_Date)
                    VALUES
                        (@MPR_ID, @Item_No, @Item_Name, @Description, @Material,
                         @Thickness_mm, @Depth_mm, @C_Width_mm, @D_Web_mm, @E_Flange_mm, @F_Length_mm,
                         @UNIT, @Qty_Per_Sheet, @Weight_kg, @MPS_Info, @Usage_Location, @REV, @Remarks,
                         @Is_Deleted, @Created_By, GETDATE())", conn);

                AddDetailParams(cmd, d);
                cmd.Parameters.AddWithValue("@Created_By", createdBy);
                cmd.ExecuteNonQuery();
            }
        }

        // ===== UPDATE DETAIL =====
        public void UpdateDetail(MPRDetail d, string modifiedBy)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    UPDATE MPR_Details SET
                        Item_No        = @Item_No,
                        Item_Name      = @Item_Name,
                        Description    = @Description,
                        Material       = @Material,
                        Thickness_mm   = @Thickness_mm,
                        Depth_mm       = @Depth_mm,
                        C_Width_mm     = @C_Width_mm,
                        D_Web_mm       = @D_Web_mm,
                        E_Flange_mm    = @E_Flange_mm,
                        F_Length_mm    = @F_Length_mm,
                        UNIT           = @UNIT,
                        Qty_Per_Sheet  = @Qty_Per_Sheet,
                        Weight_kg      = @Weight_kg,
                        MPS_Info       = @MPS_Info,
                        Usage_Location = @Usage_Location,
                        REV            = @REV,
                        Remarks        = @Remarks,
                        Is_Deleted     = @Is_Deleted
                    WHERE Detail_ID = @MPR_ID", conn);

                AddDetailParams(cmd, d);
                cmd.ExecuteNonQuery();
            }
        }

        // ===== DELETE DETAIL =====
        public void DeleteDetail(int detailId)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                new SqlCommand($"DELETE FROM MPR_Details WHERE Detail_ID = {detailId}", conn).ExecuteNonQuery();
            }
        }

        // ===== HELPERS =====
        private void AddDetailParams(SqlCommand cmd, MPRDetail d)
        {
            cmd.Parameters.AddWithValue("@MPR_ID", d.MPR_ID);
            cmd.Parameters.AddWithValue("@Item_No", d.Item_No);
            cmd.Parameters.AddWithValue("@Item_Name", d.Item_Name ?? "");
            cmd.Parameters.AddWithValue("@Description", d.Description ?? "");
            cmd.Parameters.AddWithValue("@Material", d.Material ?? "");
            cmd.Parameters.AddWithValue("@Thickness_mm", d.Thickness_mm);
            cmd.Parameters.AddWithValue("@Depth_mm", d.Depth_mm);
            cmd.Parameters.AddWithValue("@C_Width_mm", d.C_Width_mm);
            cmd.Parameters.AddWithValue("@D_Web_mm", d.D_Web_mm);
            cmd.Parameters.AddWithValue("@E_Flange_mm", d.E_Flange_mm);
            cmd.Parameters.AddWithValue("@F_Length_mm", d.F_Length_mm);
            cmd.Parameters.AddWithValue("@UNIT", d.UNIT ?? "");
            cmd.Parameters.AddWithValue("@Qty_Per_Sheet", d.Qty_Per_Sheet);
            cmd.Parameters.AddWithValue("@Weight_kg", d.Weight_kg);
            cmd.Parameters.AddWithValue("@MPS_Info", d.MPS_Info ?? "");
            cmd.Parameters.AddWithValue("@Usage_Location", d.Usage_Location ?? "");
            cmd.Parameters.AddWithValue("@REV", d.REV ?? "0");
            cmd.Parameters.AddWithValue("@Remarks", d.Remarks ?? "");
            cmd.Parameters.AddWithValue("@Is_Deleted", d.Is_Deleted ? 1 : 0); // BIT
        }

        private MPRHeader MapHeader(SqlDataReader r)
        {
            return new MPRHeader
            {
                MPR_ID = Convert.ToInt32(r["MPR_ID"]),
                MPR_No = r["MPR_No"]?.ToString() ?? "",
                Project_Name = r["Project_Name"]?.ToString() ?? "",
                Project_Code = r["Project_Code"]?.ToString() ?? "",
                Department = r["Department"]?.ToString() ?? "",
                Requestor = r["Requestor"]?.ToString() ?? "",
                Rev = r["Rev"] != DBNull.Value ? Convert.ToInt32(r["Rev"]) : 0,
                Required_Date = r["Required_Date"] != DBNull.Value ? Convert.ToDateTime(r["Required_Date"]) : null,
                Status = r["Status"]?.ToString() ?? "",
                Total_Amount = r["Total_Amount"] != DBNull.Value ? Convert.ToDecimal(r["Total_Amount"]) : 0,
                Notes = r["Notes"]?.ToString() ?? "",
                Created_Date = r["Created_Date"] != DBNull.Value ? Convert.ToDateTime(r["Created_Date"]) : null,
                Created_By = r["Created_By"]?.ToString() ?? ""
            };
        }

        private MPRDetail MapDetail(SqlDataReader r)
        {
            return new MPRDetail
            {
                Detail_ID = Convert.ToInt32(r["Detail_ID"]),
                MPR_ID = Convert.ToInt32(r["MPR_ID"]),
                Item_No = r["Item_No"] != DBNull.Value ? Convert.ToInt32(r["Item_No"]) : 0,
                Item_Name = r["Item_Name"]?.ToString() ?? "",
                Description = r["Description"]?.ToString() ?? "",
                Material = r["Material"]?.ToString() ?? "",
                Thickness_mm = r["Thickness_mm"] != DBNull.Value ? Convert.ToDecimal(r["Thickness_mm"]) : 0,
                Depth_mm = r["Depth_mm"] != DBNull.Value ? Convert.ToDecimal(r["Depth_mm"]) : 0,
                C_Width_mm = r["C_Width_mm"] != DBNull.Value ? Convert.ToDecimal(r["C_Width_mm"]) : 0,
                D_Web_mm = r["D_Web_mm"] != DBNull.Value ? Convert.ToDecimal(r["D_Web_mm"]) : 0,
                E_Flange_mm = r["E_Flange_mm"] != DBNull.Value ? Convert.ToDecimal(r["E_Flange_mm"]) : 0,
                F_Length_mm = r["F_Length_mm"] != DBNull.Value ? Convert.ToDecimal(r["F_Length_mm"]) : 0,
                UNIT = r["UNIT"]?.ToString() ?? "",
                Qty_Per_Sheet = r["Qty_Per_Sheet"] != DBNull.Value ? Convert.ToDecimal(r["Qty_Per_Sheet"]) : 0,
                Weight_kg = r["Weight_kg"] != DBNull.Value ? Convert.ToDecimal(r["Weight_kg"]) : 0,
                MPS_Info = r["MPS_Info"]?.ToString() ?? "",
                Usage_Location = r["Usage_Location"]?.ToString() ?? "",
                REV = r["REV"]?.ToString() ?? "0",
                Remarks = r["Remarks"]?.ToString() ?? "",
                Is_Deleted = r["Is_Deleted"] != DBNull.Value && Convert.ToBoolean(r["Is_Deleted"])
            };
        }
    }
}