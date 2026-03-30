using System;

namespace MPR_Managerment.Models
{
    public class Warehouse
    {
        public int Warehouse_ID { get; set; }
        public string Warehouse_Code { get; set; } = "";
        public string Warehouse_Name { get; set; } = "";
        public string Warehouse_Type { get; set; } = "";
        public string Project_Code { get; set; } = "";
        public string Dept_Abbr { get; set; } = "";
        public string Manager { get; set; } = "";
        public string Notes { get; set; } = "";
        public bool IsActive { get; set; } = true;
        public DateTime? Created_Date { get; set; }
        public string Created_By { get; set; } = "";
        public DateTime? Modified_Date { get; set; }
        public string Modified_By { get; set; } = "";
    }
}