using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;

namespace MPR_Managerment.Services
{
    public class SupplierServices
    {
        // ===== GET ALL =====
        public List<Supplier> GetAll()
        {
            var list = new List<Supplier>();
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    SELECT Supplier_ID, Company_Name, Short_Name, Supplier_Type,
                           Cert, Email, Contact_Person, Contact_Phone,
                           Company_Address, Bank_Account, Bank_Name,
                           Tax_Code, Website, Notes, IsActive,
                           Created_Date, Created_By, Modified_Date, Modified_By
                    FROM Suppliers
                    WHERE IsActive = 1 OR IsActive IS NULL
                    ORDER BY Company_Name", conn);
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) list.Add(MapSupplier(r));
            }
            return list;
        }

        // ===== GET BY ID =====
        public Supplier GetById(int supplierId)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    SELECT Supplier_ID, Company_Name, Short_Name, Supplier_Type,
                           Cert, Email, Contact_Person, Contact_Phone,
                           Company_Address, Bank_Account, Bank_Name,
                           Tax_Code, Website, Notes, IsActive,
                           Created_Date, Created_By, Modified_Date, Modified_By
                    FROM Suppliers
                    WHERE Supplier_ID = @id", conn);
                cmd.Parameters.AddWithValue("@id", supplierId);
                using (var r = cmd.ExecuteReader())
                    if (r.Read()) return MapSupplier(r);
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
                var cmd = new SqlCommand(@"
                    SELECT Supplier_ID, Company_Name, Short_Name, Supplier_Type,
                           Cert, Email, Contact_Person, Contact_Phone,
                           Company_Address, Bank_Account, Bank_Name,
                           Tax_Code, Website, Notes, IsActive,
                           Created_Date, Created_By, Modified_Date, Modified_By
                    FROM Suppliers
                    WHERE Company_Name   LIKE @kw
                       OR Short_Name     LIKE @kw
                       OR Tax_Code       LIKE @kw
                       OR Contact_Phone  LIKE @kw
                       OR Email          LIKE @kw
                    ORDER BY Company_Name", conn);
                cmd.Parameters.AddWithValue("@kw", $"%{keyword}%");
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) list.Add(MapSupplier(r));
            }
            return list;
        }

        // ===== INSERT =====
        public int Insert(Supplier s, string createdBy)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    INSERT INTO Suppliers
                        (Company_Name, Short_Name, Supplier_Type, Cert,
                         Email, Contact_Person, Contact_Phone, Company_Address,
                         Bank_Account, Bank_Name, Tax_Code, Website, Notes,
                         IsActive, Created_By, Created_Date)
                    VALUES
                        (@Company_Name, @Short_Name, @Supplier_Type, @Cert,
                         @Email, @Contact_Person, @Contact_Phone, @Company_Address,
                         @Bank_Account, @Bank_Name, @Tax_Code, @Website, @Notes,
                         1, @Created_By, GETDATE());
                    SELECT SCOPE_IDENTITY();", conn);

                AddParams(cmd, s);
                cmd.Parameters.AddWithValue("@Created_By", createdBy);
                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        // ===== UPDATE =====
        public void Update(Supplier s, string modifiedBy)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(@"
                    UPDATE Suppliers SET
                        Company_Name    = @Company_Name,
                        Short_Name      = @Short_Name,
                        Supplier_Type   = @Supplier_Type,
                        Cert            = @Cert,
                        Email           = @Email,
                        Contact_Person  = @Contact_Person,
                        Contact_Phone   = @Contact_Phone,
                        Company_Address = @Company_Address,
                        Bank_Account    = @Bank_Account,
                        Bank_Name       = @Bank_Name,
                        Tax_Code        = @Tax_Code,
                        Website         = @Website,
                        Notes           = @Notes,
                        Modified_By     = @Modified_By,
                        Modified_Date   = GETDATE()
                    WHERE Supplier_ID = @Supplier_ID", conn);

                AddParams(cmd, s);
                cmd.Parameters.AddWithValue("@Supplier_ID", s.Supplier_ID);
                cmd.Parameters.AddWithValue("@Modified_By", modifiedBy);
                cmd.ExecuteNonQuery();
            }
        }

        // ===== DELETE (soft delete) =====
        public void Delete(int supplierId)
        {
            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                new SqlCommand(
                    $"UPDATE Suppliers SET IsActive = 0 WHERE Supplier_ID = {supplierId}",
                    conn).ExecuteNonQuery();
            }
        }

        // ===== GET FOR COMBOBOX =====
        public DataTable GetForCombo()
        {
            var dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Name", typeof(string));

            // Dòng mặc định
            dt.Rows.Add(0, "-- Chọn nhà cung cấp --");

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
            return dt;
        }

        // ===== HELPERS =====
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

        private Supplier MapSupplier(SqlDataReader r)
        {
            return new Supplier
            {
                Supplier_ID = Convert.ToInt32(r["Supplier_ID"]),
                Company_Name = r["Company_Name"]?.ToString() ?? "",
                Short_Name = r["Short_Name"]?.ToString() ?? "",
                Supplier_Type = r["Supplier_Type"]?.ToString() ?? "",
                Cert = r["Cert"]?.ToString() ?? "",
                Email = r["Email"]?.ToString() ?? "",
                Contact_Person = r["Contact_Person"]?.ToString() ?? "",
                Contact_Phone = r["Contact_Phone"]?.ToString() ?? "",
                Company_Address = r["Company_Address"]?.ToString() ?? "",
                Bank_Account = r["Bank_Account"]?.ToString() ?? "",
                Bank_Name = r["Bank_Name"]?.ToString() ?? "",
                Tax_Code = r["Tax_Code"]?.ToString() ?? "",
                Website = r["Website"]?.ToString() ?? "",
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