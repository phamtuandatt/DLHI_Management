// ============================================================
//  FILE: Helpers/PermissionHelper.cs
//  Helper kiểm tra quyền — dùng chung mọi form
//  Admin bypass hoàn toàn mọi giới hạn phân quyền
// ============================================================
using System;
using System.Collections.Generic;
using MPR_Managerment.Models;
using MPR_Managerment.Services;

namespace MPR_Managerment.Helpers
{
    public static class PermissionHelper
    {
        // Cache DetailedPermissions (Dictionary) cho các form dùng Check() trực tiếp
        private static Dictionary<string, bool> _cache = new Dictionary<string, bool>();
        private static int _cachedUserId = -1;

        // ── Gọi sau login để nạp cache ───────────────────────────────────────
        public static void LoadPermissions(int userId)
        {
            _cachedUserId = userId;

            // Admin không cần cache — bypass hết
            if (AppSession.IsAdmin) return;

            var svc = new UserService();
            _cache = svc.GetDetailedPermissions(userId);
        }

        // ── Kiểm tra quyền (dùng cả AppSession.List lẫn cache Dictionary) ──
        public static bool HasPermission(string moduleCode, string action)
        {
            // Admin luôn có mọi quyền
            if (AppSession.IsAdmin) return true;

            // Kiểm tra qua List<UserPermission> trong AppSession (nguồn chính)
            var perm = AppSession.GetPermission(moduleCode);
            if (perm != null)
            {
                return action switch
                {
                    "Xem" => perm.Can_View,
                    "Xuất Excel" => perm.Can_Export,
                    "Xóa" => perm.Can_Delete,
                    "Xóa PO" => perm.Can_Delete,
                    "Xóa MPR" => perm.Can_Delete,
                    "Xóa RIR" => perm.Can_Delete,
                    "Xóa dòng" => perm.Can_Delete,
                    "Xóa user" => perm.Can_Delete,
                    _ when action.StartsWith("Tạo") || action.StartsWith("Thêm")
                                 => perm.Can_Create,
                    _ when action.StartsWith("Lưu") || action.StartsWith("Cập nhật")
                                 => perm.Can_Edit,
                    _ => perm.Can_View  // fallback: nếu có quyền xem thì cho phép
                };
            }

            // Fallback: dùng cache Dictionary (GetDetailedPermissions)
            string key = moduleCode + ":" + action;
            return _cache.ContainsKey(key) && _cache[key];
        }

        // ── Ẩn/disable control nếu không có quyền ───────────────────────────
        public static void Apply(System.Windows.Forms.Control control, string moduleCode, string action)
        {
            bool allowed = HasPermission(moduleCode, action);
            control.Enabled = allowed;
            control.Visible = allowed;
        }

        // ── Chỉ disable (vẫn hiện) ───────────────────────────────────────────
        public static void ApplyDisableOnly(System.Windows.Forms.Control control, string moduleCode, string action)
        {
            control.Enabled = HasPermission(moduleCode, action);
        }

        // ── Kiểm tra và hiện MessageBox nếu không có quyền ──────────────────
        public static bool Check(string moduleCode, string action, string actionLabel = "")
        {
            if (HasPermission(moduleCode, action)) return true;
            System.Windows.Forms.MessageBox.Show(
                $"Bạn không có quyền thực hiện chức năng này" +
                (string.IsNullOrEmpty(actionLabel) ? "" : $": {actionLabel}") + "!",
                "Không có quyền",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Warning);
            return false;
        }

        // ── Property IsAdmin công khai ───────────────────────────────────────
        public static bool IsAdmin => AppSession.IsAdmin;

        // ── Xóa cache khi logout ─────────────────────────────────────────────
        public static void Clear()
        {
            _cache.Clear();
            _cachedUserId = -1;
        }
    }
}