// ============================================================
//  FILE: Models/PaymentModels.cs
// ============================================================
using System;

namespace MPR_Managerment.Models
{
    // Kế hoạch từng đợt thanh toán
    public class PaymentSchedule
    {
        public int Schedule_ID { get; set; }
        public int PO_ID { get; set; }
        public string PONo { get; set; } = "";
        public string Project_Name { get; set; } = "";
        public int Dot_TT { get; set; } = 1;
        public string Payment_Type { get; set; } = "Chuyển khoản"; // Hình thức chuyển tiền
        public string Pay_Method { get; set; } = "Full";          // Full | Partial | Percent | ByDelivery
        public decimal Percent_TT { get; set; } = 0;
        public decimal Amount_Plan { get; set; } = 0;
        public DateTime? Due_Date { get; set; }
        public string Delivery_Ref { get; set; } = "";  // Mã lô hàng
        public string Description { get; set; } = "";
        // 3 trạng thái: Chưa TT | Một phần | Đã TT đủ
        public string Status { get; set; } = "Chưa TT";
        public DateTime? Created_Date { get; set; }
        public string Created_By { get; set; } = "";
    }

    // Lịch sử thanh toán thực tế
    public class PaymentHistory
    {
        public int Payment_ID { get; set; }
        public int? Schedule_ID { get; set; }
        public int PO_ID { get; set; }
        public string PONo { get; set; } = "";
        public string Project_Name { get; set; } = "";
        public int? Supplier_ID { get; set; }
        public string Supplier_Name { get; set; } = "";
        public DateTime Payment_Date { get; set; } = DateTime.Today;
        public decimal Amount_Paid { get; set; } = 0;
        public string Payment_Method { get; set; } = "Chuyển khoản";
        public string Bank_Name { get; set; } = "";
        public string Transaction_No { get; set; } = "";
        public string Currency { get; set; } = "VND";
        public decimal Exchange_Rate { get; set; } = 1;
        public string Notes { get; set; } = "";
        public DateTime? Created_Date { get; set; }
        public string Created_By { get; set; } = "";
    }

    // Tổng hợp thanh toán của 1 PO
    public class POPaymentSummary
    {
        public int PO_ID { get; set; }
        public string PONo { get; set; } = "";
        public string Project_Name { get; set; } = "";
        public string WorkorderNo { get; set; } = "";
        public int? Supplier_ID { get; set; }
        public string Supplier_Name { get; set; } = "";
        public string Supplier_Short { get; set; } = "";

        public DateTime? PO_Date { get; set; }
        public decimal Total_PO_Amount { get; set; }
        public decimal Total_Plan { get; set; }
        public decimal Total_Paid { get; set; }
        public decimal Amount_Remaining { get; set; }
        public decimal Percent_Paid { get; set; }
        // Chưa TT | Một phần | Đã TT đủ
        public string Payment_Status { get; set; } = "";
        public bool Is_Overdue { get; set; }
        public DateTime? Last_Payment_Date { get; set; }
        public DateTime? Next_Due_Date { get; set; }
        public int Total_Dots { get; set; }
        public int Done_Dots { get; set; }
    }

    // Tổng hợp công nợ theo NCC
    public class SupplierDebtSummary
    {
        public int Supplier_ID { get; set; }
        public string Supplier_Name { get; set; } = "";
        public string Supplier_Short { get; set; } = "";

        public string Phone { get; set; } = "";
        public string Email { get; set; } = "";
        public int Total_PO { get; set; }
        public decimal Total_PO_Value { get; set; }
        public decimal Total_Paid { get; set; }
        public decimal Total_Debt { get; set; }
        public int Overdue_PO_Count { get; set; }
    }

    // Item báo cáo công nợ theo kỳ
    public class DebtReportItem
    {
        public int Supplier_ID { get; set; }
        public string Supplier_Name { get; set; } = "";
        public string Supplier_Short { get; set; } = "";

        public int PO_ID { get; set; }
        public string PONo { get; set; } = "";
        public string Project_Name { get; set; } = "";
        public DateTime? PO_Date { get; set; }
        public decimal Total_Amount { get; set; }
        public decimal Paid_In_Range { get; set; }
        public decimal Paid_Before_Range { get; set; }
        public decimal Remaining_Debt { get; set; }
        // Chưa TT | Một phần | Đã TT đủ
        public string Payment_Status { get; set; } = "";
        public bool Is_Overdue { get; set; }
        public DateTime? Next_Due_Date { get; set; }
    }
}
