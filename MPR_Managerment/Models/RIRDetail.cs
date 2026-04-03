using System;

namespace MPR_Managerment.Models
{
    public class RIRDetail
    {
        public int RIR_Detail_ID { get; set; }
        public int RIR_ID { get; set; }
        public int? PO_Detail_ID { get; set; }
        public int Item_No { get; set; }
        public string Item_Name { get; set; } = "";
        public string Material { get; set; } = "";
        public string Size { get; set; } = "";
        public string UNIT { get; set; } = "";
        public decimal Qty_Required { get; set; }
        public decimal Qty_Received { get; set; }
        public string MTRno { get; set; } = "";
        public string Heatno { get; set; } = "";
        public string ID_Code { get; set; } = "";
        public string Inspect_Result { get; set; } = "";
        public string Remarks { get; set; } = "";
        public DateTime? Created_Date { get; set; }
        public decimal Qty_Per_Sheet { get; set; }
    }
}