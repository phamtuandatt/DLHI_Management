using System;

namespace MPR_Managerment.Models
{
    public class MPRHeader
    {
        public int MPR_ID { get; set; }
        public string MPR_No { get; set; } = "";
        public string Project_Name { get; set; } = "";
        public string Project_Code { get; set; } = "";
        public string Department { get; set; } = "";
        public string Requestor { get; set; } = "";
        public int Rev { get; set; } = 0;
        public DateTime? Required_Date { get; set; }
        public string Status { get; set; } = "";
        public decimal Total_Amount { get; set; } = 0;
        public string Notes { get; set; } = "";
        public DateTime? Created_Date { get; set; }
        public string Created_By { get; set; } = "";
        public DateTime? Modified_Date { get; set; }
        public string Modified_By { get; set; } = "";
    }
}