using System;

namespace MPR_Managerment.Models
{
    public class WarehouseImport
    {
        public int Import_ID { get; set; }
        public string Import_No { get; set; } = "";
        public DateTime? Import_Date { get; set; }
        public int? PO_ID { get; set; }
        public int? PO_Detail_ID { get; set; }
        public int? RIR_ID { get; set; }
        public string Item_Name { get; set; } = "";
        public string Material { get; set; } = "";
        public string Size { get; set; } = "";
        public string UNIT { get; set; } = "";
        public decimal Qty_Import { get; set; }
        public decimal Weight_kg { get; set; }
        public string ID_Code { get; set; } = "";
        public string MTRno { get; set; } = "";
        public string Heatno { get; set; } = "";
        public string Project_Code { get; set; } = "";
        public string WorkorderNo { get; set; } = "";
        public string Location { get; set; } = "";
        public string Notes { get; set; } = "";
        public string Created_By { get; set; } = "";
        public DateTime? Created_Date { get; set; }

        public string Material_Detail_ID { get; set; }
        public string Material_Detail_Number { get; set; }

        public string InvoiceNo { get; set; }
        public string InvoiceDate { get; set; }
    }
}