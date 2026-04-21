// ============================================================
//  FILE: Program.cs — Entry point WinForms
// ============================================================
using System;
using System.Windows.Forms;

namespace MPR_Managerment
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();

            // Bước 1: Hiện form đăng nhập
            var loginForm = new Forms.frmLogin();
            if (loginForm.ShowDialog() != DialogResult.OK)
                return;

            // Bước 2: Nạp quyền chi tiết nếu frmLogin chưa set
            // (frmLogin đã gán AppSession.DetailedPermissions từ result.DetailedPermissions)
            // Nếu rỗng (UserService.Login chưa trả DetailedPermissions) → load từ DB
            var user = Models.AppSession.CurrentUser;
            if (user != null && Models.AppSession.DetailedPermissions.Count == 0)
                Helpers.PermissionHelper.LoadPermissions(user.User_ID);

            // Bước 3: Mở màn hình chính
            Application.Run(new Forms.frmMain());
        }
    }
}