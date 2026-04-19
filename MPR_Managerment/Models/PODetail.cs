using System;

namespace MPR_Managerment.Models
{
    public class PODetail
    {
        public int PO_Detail_ID { get; set; }
        public int PO_ID { get; set; }
        public int Item_No { get; set; }
        public string Item_Name { get; set; } = "";
        public string Material { get; set; } = "";
        public string Asize { get; set; }
        public string Bsize { get; set; }
        public string Csize { get; set; }
        public decimal Qty_Per_Sheet { get; set; }
        public string UNIT { get; set; } = "";
        public decimal Weight_kg { get; set; }
        public string MPSNo { get; set; } = "";
        public DateTime? RequestDay { get; set; }
        public string DeliveryLocation { get; set; } = "";
        public decimal Price { get; set; }
        public decimal Amount { get; set; }
        public int Received { get; set; }
        public decimal VAT { get; set; }
        public string Remarks { get; set; } = "";
        public int? MPR_Detail_ID { get; set; }

        public bool Status_Delivery { get; set; }
        public decimal Received_Qty { get; set; }
    }
}