using System;
using System.Collections.Generic;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;

namespace MPR_Managerment.Services
{
    public class SupplierService
    {
        // ===== GET FOR COMBOBOX =====
        // Dùng trong frmPO — trả về DataTable với cột ID, Name
        public System.Data.DataTable GetForCombo()
        {
            var dt = new System.Data.DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Name", typeof(string));

            dt.Rows.Add(0, "-- Chọn nhà cung cấp --");

            try
            {
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var cmd = new SqlCommand(@"
                        SELECT Supplier_ID, Company_Name
                        FROM Suppliers
                        WHERE IsActive = 1 OR IsActive IS NULL
                        ORDER BY Company_Name", conn);

                    using (var r = cmd.ExecuteReader())
                        while (r.Read())
                            dt.Rows.Add(
                                Convert.ToInt32(r["Supplier_ID"]),
                                r["Company_Name"]?.ToString() ?? "");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("GetForCombo Error: " + ex.Message);
            }

            return dt;
        }

        // ===== GET ALL =====
        public List<Supplier> GetAll()
        {
            var list = new List<Supplier>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(
                    "SELECT * FROM Suppliers ORDER BY Company_Name", conn);
                using (var reader = cmd.ExecuteReader())
                    while (reader.Read())
                        list.Add(MapSupplier(reader));
            }
            return list;
        }

        // ===== GET BY ID =====
        public Supplier GetBySupId(int supId)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(
                    $"SELECT TOP 1 * FROM Suppliers WHERE Supplier_ID = {supId}", conn);
                using (var reader = cmd.ExecuteReader())
                    if (reader.Read()) return MapSupplier(reader);
            }
            return null;
        }

        // ===== SEARCH =====
        public List<Supplier> Search(string keyword)
        {
            var list = new List<Supplier>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("sp_SearchSupplierByName", conn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@SearchTerm", keyword);
                using (var reader = cmd.ExecuteReader())
                    while (reader.Read())
                        list.Add(MapSupplier(reader));
            }
            return list;
        }

        // ===== INSERT =====
        public void Insert(Supplier s, string createdBy)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("sp_InsertSupplier", conn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                AddParams(cmd, s);
                cmd.Parameters.AddWithValue("@IsActive", s.IsActive);
                cmd.Parameters.AddWithValue("@Created_By", createdBy);
                cmd.ExecuteNonQuery();
            }
        }

        // ===== UPDATE =====
        public void Update(Supplier s, string modifiedBy)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("sp_UpdateSupplier", conn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Supplier_ID", s.Supplier_ID);
                AddParams(cmd, s);
                cmd.Parameters.AddWithValue("@IsActive", s.IsActive);
                cmd.Parameters.AddWithValue("@Modified_By", modifiedBy);
                cmd.ExecuteNonQuery();
            }
        }

        // ===== DELETE =====
        public void Delete(int supplierId, string deletedBy)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand("sp_DeleteSupplier", conn);
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Supplier_ID", supplierId);
                cmd.Parameters.AddWithValue("@Deleted_By", deletedBy);
                cmd.ExecuteNonQuery();
            }
        }

        // =====================================================
        //  HELPERS
        // =====================================================
        private void AddParams(SqlCommand cmd, Supplier s)
        {
            cmd.Parameters.AddWithValue("@Company_Name", s.Company_Name ?? "");
            cmd.Parameters.AddWithValue("@Short_Name", s.Short_Name ?? "");
            cmd.Parameters.AddWithValue("@Supplier_Type", s.Supplier_Type ?? "");
            cmd.Parameters.AddWithValue("@Cert", s.Cert ?? "");
            cmd.Parameters.AddWithValue("@Email", s.Email ?? "");
            cmd.Parameters.AddWithValue("@Contact_Person", s.Contact_Person ?? "");
            cmd.Parameters.AddWithValue("@Contact_Phone", s.Contact_Phone ?? "");
            cmd.Parameters.AddWithValue("@Company_Address", s.Company_Address ?? "");
            cmd.Parameters.AddWithValue("@Bank_Account", s.Bank_Account ?? "");
            cmd.Parameters.AddWithValue("@Bank_Name", s.Bank_Name ?? "");
            cmd.Parameters.AddWithValue("@Tax_Code", s.Tax_Code ?? "");
            cmd.Parameters.AddWithValue("@Website", s.Website ?? "");
            cmd.Parameters.AddWithValue("@Notes", s.Notes ?? "");
        }

        private Supplier MapSupplier(SqlDataReader reader) => new Supplier
        {
            Supplier_ID = Convert.ToInt32(reader["Supplier_ID"]),
            Company_Name = reader["Company_Name"]?.ToString() ?? "",
            Short_Name = reader["Short_Name"]?.ToString() ?? "",
            Supplier_Type = reader["Supplier_Type"]?.ToString() ?? "",
            Cert = reader["Cert"]?.ToString() ?? "",
            Email = reader["Email"]?.ToString() ?? "",
            Contact_Person = reader["Contact_Person"]?.ToString() ?? "",
            Contact_Phone = reader["Contact_Phone"]?.ToString() ?? "",
            Company_Address = reader["Company_Address"]?.ToString() ?? "",
            Bank_Account = reader["Bank_Account"]?.ToString() ?? "",
            Bank_Name = reader["Bank_Name"]?.ToString() ?? "",
            Tax_Code = reader["Tax_Code"]?.ToString() ?? "",
            Website = reader["Website"]?.ToString() ?? "",
            Notes = reader["Notes"]?.ToString() ?? "",
            IsActive = reader["IsActive"] != DBNull.Value && Convert.ToBoolean(reader["IsActive"]),
            Created_Date = reader["Created_Date"] != DBNull.Value ? Convert.ToDateTime(reader["Created_Date"]) : null,
            Created_By = reader["Created_By"]?.ToString() ?? ""
        };
    }
}