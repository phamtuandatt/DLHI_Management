// ============================================================
//  FILE: Helpers/PermissionHelper.cs
//  Helper kiểm tra quyền từ AppSession — dùng chung mọi form
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

        // Gọi hàm này sau khi login thành công để nạp quyền vào cache
        public static void LoadPermissions(int userId)
        {
            _cachedUserId = userId;
            var svc = new UserService();
            _cache = svc.GetDetailedPermissions(userId);
        }

        // Kiểm tra quyền: HasPermission("PO", "Tạo PO")
        public static bool HasPermission(string moduleCode, string action)
        {
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

        // Xóa cache khi logout
        public static void Clear()
        {
            _cache.Clear();
            _cachedUserId = -1;
        }
    }
}