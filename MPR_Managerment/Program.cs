// ============================================================
//  FILE: Program.cs  — Entry point WinForms
//  Namespace: khớp với project của bạn
//  ⚠ Sửa "YourNamespace" thành namespace thực tế của project
// ============================================================
using System;
using System.Windows.Forms;

// ── Thêm đúng using cho project của bạn ──
// Ví dụ: using MPR_Managerment.Forms;
// Ví dụ: using ERP_Management.Forms;

namespace MPR_Managerment   // ← ĐỔI thành namespace đúng của project
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();

            var loginForm = new Forms.frmLogin();
            if (loginForm.ShowDialog() != DialogResult.OK)
                return;

            // Nạp quyền sau login
            var user = Models.AppSession.CurrentUser;
            if (user != null)
                Helpers.PermissionHelper.LoadPermissions(user.User_ID);

            Application.Run(new Forms.frmMain());
        }
    }
}