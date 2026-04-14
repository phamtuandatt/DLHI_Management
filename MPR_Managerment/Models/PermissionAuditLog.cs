// ============================================================
//  THÊM VÀO Models/ (tạo file mới: Models/PermissionAuditLog.cs)
// ============================================================
namespace MPR_Managerment.Models
{
    public class PermissionAuditLog
    {
        public DateTime Changed_At { get; set; }
        public string Changed_By { get; set; } = "";
        public string Action_Detail { get; set; } = "";
    }
}