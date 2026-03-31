// ============================================================
//  FILE: Models/Supplier.cs
//  Thêm property Supplier_Name (alias của Company_Name)
//  để các form cũ (frmPayment, frmPO...) dùng được
//  mà không cần sửa code
// ============================================================
using System;

namespace MPR_Managerment.Models
{
    public class Supplier
    {
        public int Supplier_ID { get; set; }

        // Tên cột thực tế trong database
        public string Company_Name { get; set; } = "";
        public string Short_Name { get; set; } = "";

        // *** Alias để tương thích với code cũ dùng Supplier_Name ***
        // frmPayment, frmPO dùng s.Supplier_Name → trỏ về Company_Name
        public string Supplier_Name
        {
            get => Company_Name;
            set => Company_Name = value;
        }

        public string Supplier_Type { get; set; } = "";
        public string Cert { get; set; } = "";
        public string Email { get; set; } = "";
        public string Contact_Person { get; set; } = "";
        public string Contact_Phone { get; set; } = "";
        public string Company_Address { get; set; } = "";
        public string Bank_Account { get; set; } = "";
        public string Bank_Name { get; set; } = "";
        public string Tax_Code { get; set; } = "";
        public string Website { get; set; } = "";
        public string Notes { get; set; } = "";
        public bool IsActive { get; set; } = true;
        public DateTime? Created_Date { get; set; }
        public string Created_By { get; set; } = "";
        public DateTime? Modified_Date { get; set; }
        public string Modified_By { get; set; } = "";
    }
}