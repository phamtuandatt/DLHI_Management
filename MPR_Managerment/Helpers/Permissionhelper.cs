// ============================================================
//  FILE: Helpers/PermissionHelper.cs
//  Helper kiểm tra quyền từ AppSession — dùng chung mọi form
//  Admin (Role_ID=1) được bypass toàn bộ phân quyền
// ============================================================
using System;
using System.Collections.Generic;
using MPR_Managerment.Models;
using MPR_Managerment.Services;

namespace MPR_Managerment.Helpers
{
    public static class PermissionHelper
    {
        // Cache quyền cho session hiện tại (load 1 lần khi login)
        private static Dictionary<string, bool> _cache = new Dictionary<string, bool>();
        private static int _cachedUserId = -1;
        private static bool _isAdmin = false; // Admin bypass toàn bộ phân quyền

        // Gọi hàm này sau khi login thành công để nạp quyền vào cache
        public static void LoadPermissions(int userId)
        {
            _cachedUserId = userId;

            // Kiểm tra Admin qua AppSession (Role_Name hoặc Role_ID)
            var currentUser = AppSession.CurrentUser;
            _isAdmin = currentUser != null &&
                       (currentUser.Role_Name?.Equals("Admin", StringComparison.OrdinalIgnoreCase) == true
                        || currentUser.Role_ID == 1);

            // Admin không cần load cache — bypass hết
            if (_isAdmin) return;

            var svc = new UserService();
            _cache = svc.GetDetailedPermissions(userId);
        }

        // Kiểm tra quyền: HasPermission("PO", "Tạo PO")
        public static bool HasPermission(string moduleCode, string action)
        {
            // Admin luôn có quyền
            if (_isAdmin) return true;

            string key = moduleCode + ":" + action;
            return _cache.ContainsKey(key) && _cache[key];
        }

        // Ẩn/disable button nếu không có quyền
        public static void Apply(System.Windows.Forms.Control control, string moduleCode, string action)
        {
            bool allowed = HasPermission(moduleCode, action);
            control.Enabled = allowed;
            control.Visible = allowed;
        }

        // Chỉ disable (vẫn hiện nhưng không bấm được)
        public static void ApplyDisableOnly(System.Windows.Forms.Control control, string moduleCode, string action)
        {
            control.Enabled = HasPermission(moduleCode, action);
        }

        // Kiểm tra và hiện MessageBox nếu không có quyền — dùng trong Click handler
        public static bool Check(string moduleCode, string action, string actionLabel = "")
        {
            if (HasPermission(moduleCode, action)) return true;
            System.Windows.Forms.MessageBox.Show(
                $"Bạn không có quyền thực hiện chức năng này{(string.IsNullOrEmpty(actionLabel) ? "" : $": {actionLabel}")}!",
                "Không có quyền",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Warning);
            return false;
        }

        // Trả về true nếu user hiện tại là Admin
        public static bool IsAdmin => _isAdmin;

        // Xóa cache khi logout
        public static void Clear()
        {
            _cache.Clear();
            _cachedUserId = -1;
            _isAdmin = false;
        }
    }
}