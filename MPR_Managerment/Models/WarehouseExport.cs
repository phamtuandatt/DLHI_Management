using System;

namespace MPR_Managerment.Models
{
    public class WarehouseExport
    {
        public int Export_ID { get; set; }
        public string Export_No { get; set; } = "";
        public DateTime? Export_Date { get; set; }
        public int? Import_ID { get; set; }
        public string Item_Name { get; set; } = "";
        public string Material { get; set; } = "";
        public string Size { get; set; } = "";
        public string UNIT { get; set; } = "";
        public decimal Qty_Export { get; set; }
        public decimal Weight_kg { get; set; }
        public string ID_Code { get; set; } = "";
        public string Project_Code { get; set; } = "";
        public string WorkorderNo { get; set; } = "";
        public string Export_To { get; set; } = "";
        public string Purpose { get; set; } = "";
        public string Notes { get; set; } = "";
        public string Created_By { get; set; } = "";
        public DateTime? Created_Date { get; set; }
    }
}