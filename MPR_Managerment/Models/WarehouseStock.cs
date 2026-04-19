using System;

namespace MPR_Managerment.Models
{
    public class WarehouseStock
    {
        public int Import_ID { get; set; }
        public string Import_No { get; set; } = "";
        public DateTime? Import_Date { get; set; }
        public string Item_Name { get; set; } = "";
        public string Material { get; set; } = "";
        public string Size { get; set; } = "";
        public string UNIT { get; set; } = "";
        public string ID_Code { get; set; } = "";
        public string MTRno { get; set; } = "";
        public string Heatno { get; set; } = "";
        public string Project_Code { get; set; } = "";
        public string WorkorderNo { get; set; } = "";
        public string Location { get; set; } = "";
        public string PONo { get; set; } = "";
        public string MPR_No { get; set; } = "";
        public decimal Qty_Import { get; set; }
        public decimal Weight_Import { get; set; }
        public decimal Qty_Exported { get; set; }
        public decimal Weight_Exported { get; set; }
        public decimal Qty_Stock { get; set; }
        public decimal Weight_Stock { get; set; }


        public string QC_Code { get; set; } = "";
        public string QC_Status { get; set; } = "";
    }
}