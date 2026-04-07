using System;

namespace MPR_Managerment.Models
{
    public class POHead
    {
        public int PO_ID { get; set; }
        public string Project_Name { get; set; } = "";
        public string WorkorderNo { get; set; } = "";
        public string MPR_No { get; set; } = "";
        public string PONo { get; set; } = "";
        public string Prepared { get; set; } = "";
        public string Reviewed { get; set; } = "";
        public string Agreement { get; set; } = "";
        public string Approved { get; set; } = "";
        public DateTime? PO_Date { get; set; }
        public decimal Total_Amount { get; set; }
        public string Status { get; set; } = "";
        public string Notes { get; set; } = "";
        public int Revise { get; set; }
        public DateTime? Created_Date { get; set; }
        public string Created_By { get; set; } = "";

        public int Supplier_ID { get; set; }    
        public string ProjectCode {  get; set; }

        public bool IsImported { get; set; } = false;
        public DateTime? ImportedDate { get; set; }
    }
}