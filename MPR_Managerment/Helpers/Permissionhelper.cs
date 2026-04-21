// ============================================================
//  FILE: Helpers/PermissionHelper.cs
//  Dùng trong từng Form để kiểm tra & áp quyền lên UI controls
//  Toàn bộ logic thực sự nằm trong AppSession.HasPermission()
// ============================================================
using System.Windows.Forms;
using MPR_Managerment.Models;

namespace MPR_Managerment.Helpers
{
    public static class PermissionHelper
    {
        // =====================================================================
        //  LOAD PERMISSIONS TỪ DB
        //  Gọi từ Program.cs và frmLogin làm fallback
        //  khi LoginResult.DetailedPermissions chưa được nạp
        // =====================================================================
        public static void LoadPermissions(int userId)
        {
            // ⚡ Admin không cần load quyền từ DB — bypass hoàn toàn
            // Điều này đảm bảo admin KHÔNG BAO GIỜ bị khóa dù DB có data sai
            if (AppSession.IsAdmin)
            {
                AppSession.EnsureAdminBypass();
                return;
            }

            try
            {
                var svc = new Services.UserService();
                // GetDetailedPermissions trả Dictionary<string,bool>
                // key format: "MODULE_CODE:action"
                var perms = svc.GetDetailedPermissions(userId);
                AppSession.DetailedPermissions = perms;
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[PermissionHelper] LoadPermissions error: " + ex.Message);
            }
        }

        // =====================================================================
        //  APPLY — Disable control nếu không có quyền
        //  Gọi trong ApplyPermissions() khi Form Load
        // =====================================================================
        /// <summary>
        /// Nếu user KHÔNG có quyền → control.Enabled = false (xám, không click được).
        /// Nếu user CÓ quyền       → control.Enabled = true.
        /// Admin luôn được Enabled = true.
        /// </summary>
        public static void Apply(Control control, string moduleCode, string action)
        {
            if (control == null) return;
            control.Enabled = AppSession.HasPermission(moduleCode, action);
        }

        // =====================================================================
        //  CHECK — Dùng trong event handler trước khi thực thi logic
        // =====================================================================
        /// <summary>
        /// Nếu KHÔNG có quyền → hiện MessageBox + return false.
        /// Nếu CÓ quyền       → return true.
        /// Dùng: if (!PermissionHelper.Check("MPR", "Tạo MPR", "Tạo MPR")) return;
        /// </summary>
        public static bool Check(string moduleCode, string action, string actionLabel = null)
        {
            if (AppSession.HasPermission(moduleCode, action)) return true;

            string label = actionLabel ?? action;
            MessageBox.Show(
                $"Bạn không có quyền thực hiện chức năng:\n\n  \"{label}\"\n\n" +
                $"Vui lòng liên hệ quản trị viên để được cấp quyền.",
                "⛔ Không có quyền truy cập",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return false;
        }
    }
}