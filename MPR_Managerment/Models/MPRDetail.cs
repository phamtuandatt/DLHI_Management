using System;

namespace MPR_Managerment.Models
{
    public class MPRDetail
    {
        public int Detail_ID { get; set; }
        public int MPR_ID { get; set; }
        public int Item_No { get; set; }
        public string Item_Name { get; set; } = "";
        public string Description { get; set; } = "";
        public string Material { get; set; } = "";
        public decimal Thickness_mm { get; set; }
        public decimal Depth_mm { get; set; }
        public decimal C_Width_mm { get; set; }
        public decimal D_Web_mm { get; set; }
        public decimal E_Flange_mm { get; set; }
        public decimal F_Length_mm { get; set; }
        public string Usage_Location { get; set; } = "";
        public string MPS_Info { get; set; } = "";
        public string REV { get; set; } = "";
        public DateTime? DWG_BOQ_Receive_Date { get; set; }
        public DateTime? Issue_Date { get; set; }
        public string UNIT { get; set; } = "";
        public decimal Qty_Per_Sheet { get; set; }
        public decimal Weight_kg { get; set; }
        public string Remarks { get; set; } = "";
    }
}