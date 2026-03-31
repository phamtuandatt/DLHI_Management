// ============================================================
//  BƯỚC 1: Sửa Program.cs — thêm màn hình đăng nhập
// ============================================================
// Mở Program.cs → thay toàn bộ bằng:

using System;
using System.Windows.Forms;
using MPR_Managerment.Forms;
using MPR_Managerment.Forms.ItemCodeGUI;

namespace MPR_Managerment
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();

            // Hiển thị form đăng nhập trước
            var login = new frmLogin();
            if (login.ShowDialog() != DialogResult.OK)
            {
                Application.Exit();
                return;
            }

            // Đăng nhập thành công → mở Main
            Application.Run(new frmMain());
            //Application.Run(new frmCreateItemCode());
        }
    }
}