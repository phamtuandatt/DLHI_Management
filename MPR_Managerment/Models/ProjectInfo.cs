using System;

namespace MPR_Managerment.Models
{
    public class ProjectInfo
    {
        public int Id { get; set; }
        public string ProjectName { get; set; } = "";
        public string ProjectCode { get; set; } = "";
        public string WorkorderNo { get; set; } = "";
        public string Customer { get; set; } = "";
        public decimal PJWeight { get; set; }
        public decimal PJBudget { get; set; }
        public string POCode { get; set; } = "";
        public string MPRCode { get; set; } = "";
        public string Status { get; set; } = "";
        public string Notes { get; set; } = "";
        public string PO_Link { get; set; } = "";
        public string RIR_Link { get; set; } = "";
        public string MPR_Link { get; set; } = "";
        public DateTime? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
    }
}