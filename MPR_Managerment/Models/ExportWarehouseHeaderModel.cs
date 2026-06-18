using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;

namespace MPR_Managerment.Models
{
    public class ExportWarehouseHeaderModel
    {
        public int ExportID { get; set; }

        // Các thông tin mã phiếu xuất
        public string ExportNo { get; set; } = string.Empty;
        public string ExportNumber { get; set; } = string.Empty;

        // Dự án xuất đi
        public int FromProjectID { get; set; }
        public string FromProjectCode { get; set; } = string.Empty;
        public string FromProjectName { get; set; } = string.Empty;

        // Dự án nhận đến
        public int ToProjectID { get; set; }
        public string ToProjectCode { get; set; } = string.Empty;
        public string ToProjectName { get; set; } = string.Empty;

        // Tổng tiền (Trong SQL là DECIMAL/MONEY, tương ứng với decimal trong C#)
        public decimal ExportTotals { get; set; }

        // Trạng thái phiếu (Thường là INT: 1-Mới, 2-Đã xuất, v.v...)
        public string Status { get; set; }

        // Ghi chú (Cho phép Null nếu trong DB không bắt buộc nhập)
        public string? Notes { get; set; }

        // Thông tin khởi tạo
        public string CreateBy { get; set; } = string.Empty;
        public DateTime CreateDate { get; set; }

        // Thông tin cập nhật (Cho phép Null vì khi mới tạo chưa có dữ liệu này)
        public string? UpdateBy { get; set; }
        public DateTime? UpdateDate { get; set; }
    }
}
