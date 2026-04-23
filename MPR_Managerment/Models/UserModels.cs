// ============================================================
//  FILE: Models/UserModels.cs
//  Chứa toàn bộ Models liên quan đến User & Permission
// ============================================================
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace MPR_Managerment.Models
{
    // ===== USER =====
    public class AppUser
    {
        public int User_ID { get; set; }
        public string Username { get; set; } = "";
        public string Full_Name { get; set; } = "";
        public string Email { get; set; } = "";
        public string Phone { get; set; } = "";
        public string Department { get; set; } = "";
        public int Role_ID { get; set; }
        public string Role_Name { get; set; } = "";
        public bool Is_Active { get; set; } = true;
        public bool Must_Change_Password { get; set; } = false;
        public DateTime? Last_Login { get; set; }
        public DateTime? Created_Date { get; set; }
        public string Created_By { get; set; } = "";
    }

    // ===== ROLE =====
    public class Role
    {
        public int Role_ID { get; set; }
        public string Role_Name { get; set; } = "";
        public string Description { get; set; } = "";
        public bool Is_Active { get; set; } = true;
    }

    // ===== MODULE =====
    public class AppModule
    {
        public int Module_ID { get; set; }
        public string Module_Code { get; set; } = "";
        public string Module_Name { get; set; } = "";
        public int Sort_Order { get; set; }
    }

    // ===== PERMISSION (quyền hiệu lực cuối) =====
    public class UserPermission
    {
        public int Module_ID { get; set; }
        public string Module_Code { get; set; } = "";
        public string Module_Name { get; set; } = "";
        public bool Can_View { get; set; }
        public bool Can_Create { get; set; }
        public bool Can_Edit { get; set; }
        public bool Can_Delete { get; set; }
        public bool Can_Export { get; set; }
        public bool Is_Custom_Override { get; set; }
    }

    // ===== LOGIN RESULT =====
    public class LoginResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = "";
        public AppUser? User { get; set; }
        public List<UserPermission> Permissions { get; set; } = new();

        // Quyền chi tiết dạng key-value: "MODULE:action" = true/false
        // Được nạp từ UserService.Login() → dùng bởi PermissionHelper.Check/Apply
        public Dictionary<string, bool> DetailedPermissions { get; set; }
            = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
    }

    // ===== SESSION — lưu thông tin user đang đăng nhập =====
    public static class AppSession
    {
        // ── User hiện tại ─────────────────────────────────────────────────────
        public static AppUser? CurrentUser { get; set; }

        // ── Quyền dạng List<UserPermission> — dùng cho CanView/CanEdit/... ───
        public static List<UserPermission> Permissions { get; set; } = new();

        // ── Quyền chi tiết dạng Dictionary — dùng cho HasPermission() ────────
        // Key: "MODULE_CODE:action"  VD: "MPR:Tạo MPR", "PO:Xóa PO"
        // Được gán từ frmLogin hoặc PermissionHelper.LoadPermissions(userId)
        private static Dictionary<string, bool> _detailedPermissions
            = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

        public static Dictionary<string, bool> DetailedPermissions
        {
            get => _detailedPermissions;
            set
            {
                // ⚡ Admin bypass: không lưu DetailedPermissions cho admin
                // Admin luôn có mọi quyền qua IsAdmin — không cần kiểm tra từng action
                if (IsAdmin) return;
                _detailedPermissions = value != null
                    ? new Dictionary<string, bool>(value, StringComparer.OrdinalIgnoreCase)
                    : new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            }
        }

        public static bool IsLoggedIn => CurrentUser != null;

        // ── Admin bypass ──────────────────────────────────────────────────────
        // Admin được xác định HOÀN TOÀN từ thông tin User — không phụ thuộc DB quyền
        public static bool IsAdmin =>
            CurrentUser != null
            && (CurrentUser.Role_ID == 1
             || string.Equals(CurrentUser.Role_Name, "Admin", StringComparison.OrdinalIgnoreCase)
             || string.Equals(CurrentUser.Role_Name, "Administrator", StringComparison.OrdinalIgnoreCase)
             || string.Equals(CurrentUser.Username, "admin", StringComparison.OrdinalIgnoreCase));

        // ── Reset toàn bộ phân quyền về trạng thái "Admin có tất cả" ─────────
        /// <summary>
        /// Gọi khi phát hiện admin bị mất quyền.
        /// Xóa hết DetailedPermissions để IsAdmin bypass hoàn toàn.
        /// </summary>
        public static void EnsureAdminBypass()
        {
            if (!IsAdmin) return;
            // Admin không cần DetailedPermissions — xóa đi để tránh conflict
            _detailedPermissions.Clear();
            Permissions.Clear();
        }

        // =====================================================================
        //  LẤY QUYỀN MODULE (dùng List<UserPermission>)
        // =====================================================================
        public static UserPermission? GetPermission(string moduleCode)
            => Permissions.Find(p => string.Equals(p.Module_Code, moduleCode,
                                                    StringComparison.OrdinalIgnoreCase));

        // =====================================================================
        //  KIỂM TRA QUYỀN — dùng cho menu frmMain và các form
        //  Ưu tiên DetailedPermissions (key "MODULE:Xem") vì đây là nguồn
        //  dữ liệu thực tế được lưu từ frmUserManagement.
        //  Fallback về List<UserPermission>.Can_View nếu hệ thống cũ populate nó.
        // =====================================================================
        public static bool CanView(string moduleCode)
        {
            if (IsAdmin) return true;
            if (_detailedPermissions.Count > 0)
            {
                // Kiểm tra key "MODULE:Xem" trước
                if (_detailedPermissions.TryGetValue(moduleCode + ":Xem", out bool v))
                    return v;
                // Key không có → fallback về Can_View từ UserPermission
            }
            return GetPermission(moduleCode)?.Can_View ?? false;
        }

        public static bool CanDelete(string moduleCode)
        {
            if (IsAdmin) return true;
            if (_detailedPermissions.Count > 0)
                return HasPermission(moduleCode, "Xóa")
                    || HasPermission(moduleCode, "Xóa PO")
                    || HasPermission(moduleCode, "Xóa MPR")
                    || HasPermission(moduleCode, "Xóa RIR")
                    || HasPermission(moduleCode, "Xóa dòng");
            return GetPermission(moduleCode)?.Can_Delete ?? false;
        }

        public static bool CanExport(string moduleCode)
        {
            if (IsAdmin) return true;
            if (_detailedPermissions.Count > 0)
                return HasPermission(moduleCode, "Xuất Excel")
                    || HasPermission(moduleCode, "Xuất tồn kho");
            return GetPermission(moduleCode)?.Can_Export ?? false;
        }

        // =====================================================================
        //  KIỂM TRA QUYỀN CHI TIẾT — dùng cho PermissionHelper.Check/Apply
        // =====================================================================

        /// <summary>
        /// Kiểm tra quyền theo "MODULE:action" trong DetailedPermissions.
        /// Admin luôn trả về true.
        /// Đây là method cốt lõi — PermissionHelper.Check và Apply đều gọi vào đây.
        /// </summary>
        public static bool HasPermission(string moduleCode, string action)
        {
            if (IsAdmin) return true;
            string key = moduleCode + ":" + action;
            return _detailedPermissions.TryGetValue(key, out bool val) && val;
        }

        /// <summary>
        /// Kiểm tra bất kỳ quyền ghi nào của module — dùng trong frmPayment.
        /// </summary>
        public static bool CanEdit(string moduleCode)
        {
            if (IsAdmin) return true;
            return HasPermission(moduleCode, "Lưu")
                || HasPermission(moduleCode, "Lưu chi tiết")
                || HasPermission(moduleCode, "Lưu PO")
                || HasPermission(moduleCode, "Lưu Header")
                || (GetPermission(moduleCode)?.Can_Edit ?? false);
        }

        /// <summary>
        /// Kiểm tra bất kỳ quyền tạo mới nào của module — dùng trong frmPayment.
        /// </summary>
        public static bool CanCreate(string moduleCode)
        {
            if (IsAdmin) return true;
            return HasPermission(moduleCode, "Tạo PO")
                || HasPermission(moduleCode, "Tạo MPR")
                || HasPermission(moduleCode, "Tạo RIR")
                || HasPermission(moduleCode, "Thêm mới")
                || HasPermission(moduleCode, "Tạo user")
                || (GetPermission(moduleCode)?.Can_Create ?? false);
        }

        // =====================================================================
        //  CLEAR — gọi khi logout
        // =====================================================================
        public static void Clear()
        {
            CurrentUser = null;
            Permissions.Clear();
            _detailedPermissions.Clear();
        }
    }
}