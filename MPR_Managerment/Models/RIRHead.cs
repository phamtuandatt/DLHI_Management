using System;

namespace MPR_Managerment.Models
{
    public class RIRHead
    {
        public int RIR_ID { get; set; }
        public string RIR_No { get; set; } = "";
        public DateTime? Issue_Date { get; set; }
        public string Project_Name { get; set; } = "";
        public string WorkorderNo { get; set; } = "";
        public string MPR_No { get; set; } = "";
        public string Customer { get; set; } = "";
        public string PONo { get; set; } = "";
        public string Status { get; set; } = "Chờ kiểm tra";
        public DateTime? Created_Date { get; set; }
        public string Created_By { get; set; } = "";
    }
}