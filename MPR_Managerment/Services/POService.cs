using System;
using System.Collections.Generic;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;

namespace MPR_Managerment.Services
{
    public class POService
    {
        public List<POHead> GetAll()
        {
            var list = new List<POHead>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("SELECT * FROM PO_head ORDER BY Created_Date DESC", conn);
                var r = cmd.ExecuteReader();
                while (r.Read()) list.Add(MapHead(r));
            }
            return list;
        }

        public POHead GetPOByPONo(int poId)
        {
            var poModel = new POHead();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("SELECT TOP 1 * FROM PO_head WHERE PO_ID = @kw", conn);
                cmd.Parameters.AddWithValue("@kw", $"{poId}");
                var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    poModel = MapHead(r);
                    break;
                }
            }
            return poModel;
        }

        public List<POHead> Search(string keyword)
        {
            var list = new List<POHead>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("SELECT * FROM PO_head WHERE PONo LIKE @kw OR Project_Name LIKE @kw OR MPR_No LIKE @kw ORDER BY Created_Date DESC", conn);
                cmd.Parameters.AddWithValue("@kw", $"%{keyword}%");
                var r = cmd.ExecuteReader();
                while (r.Read()) list.Add(MapHead(r));
            }
            return list;
        }

        public List<PODetail> GetDetails(int poId)
        {
            var list = new List<PODetail>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("SELECT * FROM PO_Detail WHERE PO_ID = @id ORDER BY Item_No", conn);
                cmd.Parameters.AddWithValue("@id", poId);
                var r = cmd.ExecuteReader();
                while (r.Read()) list.Add(MapDetail(r));
            }
            return list;
        }

        public int InsertHead(POHead h, string createdBy)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("sp_InsertPO", conn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Project_Name", h.Project_Name ?? "");
                cmd.Parameters.AddWithValue("@WorkorderNo", h.WorkorderNo ?? "");
                cmd.Parameters.AddWithValue("@MPR_No", h.MPR_No ?? "");
                cmd.Parameters.AddWithValue("@PONo", h.PONo);
                cmd.Parameters.AddWithValue("@Prepared", h.Prepared ?? "");
                cmd.Parameters.AddWithValue("@Reviewed", h.Reviewed ?? "");
                cmd.Parameters.AddWithValue("@Agreement", h.Agreement ?? "");
                cmd.Parameters.AddWithValue("@Approved", h.Approved ?? "");
                cmd.Parameters.AddWithValue("@Created_By", createdBy);
                cmd.Parameters.AddWithValue("@SupplierID", h.Supplier_ID);
                cmd.Parameters.AddWithValue("@ProjectCode", h.ProjectCode);
                var r = cmd.ExecuteReader();
                if (r.Read()) return Convert.ToInt32(r["NewPO_ID"]);
                return 0;
            }
        }

        public void UpdateHead(POHead h, string modifiedBy)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("sp_UpdatePOHead", conn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@PO_ID", h.PO_ID);
                cmd.Parameters.AddWithValue("@Project_Name", h.Project_Name ?? "");
                cmd.Parameters.AddWithValue("@WorkorderNo", h.WorkorderNo ?? "");
                cmd.Parameters.AddWithValue("@MPR_No", h.MPR_No ?? "");
                cmd.Parameters.AddWithValue("@PONo", h.PONo);
                cmd.Parameters.AddWithValue("@Prepared", h.Prepared ?? "");
                cmd.Parameters.AddWithValue("@Reviewed", h.Reviewed ?? "");
                cmd.Parameters.AddWithValue("@Agreement", h.Agreement ?? "");
                cmd.Parameters.AddWithValue("@Approved", h.Approved ?? "");
                cmd.Parameters.AddWithValue("@PO_Date", h.PO_Date.HasValue ? (object)h.PO_Date.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@Status", h.Status ?? "");
                cmd.Parameters.AddWithValue("@Notes", h.Notes ?? "");
                cmd.Parameters.AddWithValue("@Modified_By", modifiedBy);
                cmd.Parameters.AddWithValue("@Revise", h.Revise);
                cmd.ExecuteNonQuery();
            }
        }

        public void InsertDetail(PODetail d, int poId)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("sp_InsertPODetail", conn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@PO_ID", poId);
                cmd.Parameters.AddWithValue("@Item_No", d.Item_No);
                cmd.Parameters.AddWithValue("@item_name", d.Item_Name ?? "");
                cmd.Parameters.AddWithValue("@Material", d.Material ?? "");
                cmd.Parameters.AddWithValue("@Asize", d.Asize);
                cmd.Parameters.AddWithValue("@Bsize", d.Bsize);
                cmd.Parameters.AddWithValue("@Csize", d.Csize);
                cmd.Parameters.AddWithValue("@Qty_Per_Sheet", d.Qty_Per_Sheet);
                cmd.Parameters.AddWithValue("@UNIT", d.UNIT ?? "");
                cmd.Parameters.AddWithValue("@Weight_kg", d.Weight_kg);
                cmd.Parameters.AddWithValue("@MPSNo", d.MPSNo ?? "");
                cmd.Parameters.AddWithValue("@RequestDay", d.RequestDay.HasValue ? (object)d.RequestDay.Value : DBNull.Value);
                cmd.Parameters.AddWithValue("@DeliveryLocation", d.DeliveryLocation ?? "");
                cmd.Parameters.AddWithValue("@Price", d.Price);
                cmd.Parameters.AddWithValue("@Amount", d.Price * d.Qty_Per_Sheet);
                cmd.Parameters.AddWithValue("@VAT", d.VAT);
                cmd.Parameters.AddWithValue("@Remarks", d.Remarks ?? "");
                cmd.Parameters.AddWithValue("@MPR_Detail_ID", d.MPR_Detail_ID.HasValue ? (object)d.MPR_Detail_ID.Value : DBNull.Value);
                cmd.ExecuteNonQuery();
            }
        }

        public void DeleteDetail(int detailId)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                new SqlCommand($"DELETE FROM PO_Detail WHERE PO_Detail_ID = {detailId}", conn).ExecuteNonQuery();
            }
        }

        public void DeletePO(int poId)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                new SqlCommand($"DELETE FROM PO_Detail WHERE PO_ID = {poId}; DELETE FROM PO_head WHERE PO_ID = {poId}", conn).ExecuteNonQuery();
            }
        }

        public void UpdateStatus(string poNo, string status, string modifiedBy)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("sp_UpdatePOStatus", conn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@PONo", poNo);
                cmd.Parameters.AddWithValue("@Status", status);
                cmd.Parameters.AddWithValue("@Modified_By", modifiedBy);
                cmd.ExecuteNonQuery();
            }
        }

        private POHead MapHead(SqlDataReader r)
        {
            return new POHead
            {
                PO_ID = r["PO_ID"] != DBNull.Value ? Convert.ToInt32(r["PO_ID"]) : 0,
                Project_Name = r["Project_Name"]?.ToString() ?? "",
                WorkorderNo = r["WorkorderNo"]?.ToString() ?? "",
                MPR_No = r["MPR_No"]?.ToString() ?? "",
                PONo = r["PONo"]?.ToString() ?? "",
                Prepared = r["Prepared"]?.ToString() ?? "",
                Reviewed = r["Reviewed"]?.ToString() ?? "",
                Agreement = r["Agreement"]?.ToString() ?? "",
                Approved = r["Approved"]?.ToString() ?? "",
                PO_Date = r["PO_Date"] != DBNull.Value ? Convert.ToDateTime(r["PO_Date"]) : (DateTime?)null,
                Total_Amount = r["Total_Amount"] != DBNull.Value ? Convert.ToDecimal(r["Total_Amount"]) : 0,
                Status = r["Status"]?.ToString() ?? "",
                Notes = r["Notes"]?.ToString() ?? "",
                Revise = r["Revise"] != DBNull.Value ? Convert.ToInt32(r["Revise"]) : 0,
                Created_Date = r["Created_Date"] != DBNull.Value ? Convert.ToDateTime(r["Created_Date"]) : (DateTime?)null,
                Created_By = r["Created_By"]?.ToString() ?? "",
                // Lỗi chính nằm ở dòng Supplier_ID này (khi PO mới tạo chưa có NCC)
                Supplier_ID = r["Supplier_ID"] != DBNull.Value ? Convert.ToInt32(r["Supplier_ID"]) : 0
            };
        }

        private PODetail MapDetail(SqlDataReader r)
        {
            return new PODetail
            {
                PO_Detail_ID = r["PO_Detail_ID"] != DBNull.Value ? Convert.ToInt32(r["PO_Detail_ID"]) : 0,
                PO_ID = r["PO_ID"] != DBNull.Value ? Convert.ToInt32(r["PO_ID"]) : 0,
                Item_No = r["Item_No"] != DBNull.Value ? Convert.ToInt32(r["Item_No"]) : 0,
                Item_Name = r["item_name"]?.ToString() ?? "",
                Material = r["Material"]?.ToString() ?? "",
                Asize = r["Asize"]?.ToString() ?? "",
                Bsize = r["Bsize"]?.ToString() ?? "",
                Csize = r["Csize"]?.ToString() ?? "",
                Qty_Per_Sheet = r["Qty_Per_Sheet"] != DBNull.Value ? Convert.ToInt32(r["Qty_Per_Sheet"]) : 0,
                UNIT = r["UNIT"]?.ToString() ?? "",
                Weight_kg = r["Weight_kg"] != DBNull.Value ? Convert.ToDecimal(r["Weight_kg"]) : 0,
                MPSNo = r["MPSNo"]?.ToString() ?? "",
                RequestDay = r["RequestDay"] != DBNull.Value ? Convert.ToDateTime(r["RequestDay"]) : (DateTime?)null,
                DeliveryLocation = r["DeliveryLocation"]?.ToString() ?? "",
                Price = r["Price"] != DBNull.Value ? Convert.ToDecimal(r["Price"]) : 0,
                Amount = r["Amount"] != DBNull.Value ? Convert.ToDecimal(r["Amount"]) : 0,
                Received = r["Received"] != DBNull.Value ? Convert.ToInt32(r["Received"]) : 0,
                VAT = r["VAT"] != DBNull.Value ? Convert.ToDecimal(r["VAT"]) : 0,
                Remarks = r["Remarks"]?.ToString() ?? "",
                MPR_Detail_ID = r["MPR_Detail_ID"] != DBNull.Value ? Convert.ToInt32(r["MPR_Detail_ID"]) : (int?)null
            };
        }
    }
}