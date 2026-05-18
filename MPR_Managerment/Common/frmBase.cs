using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MPR_Managerment.Common
{
    public partial class frmBase : Form
    {
        public frmBase()
        {
            try
            {
                // Đảm bảo file icon.ico nằm cùng thư mục với file chạy .exe (thư mục bin\Debug)
                string iconPath = System.IO.Path.Combine(Application.StartupPath, "icon.ico");
                if (System.IO.File.Exists(iconPath))
                {
                    this.Icon = new Icon(iconPath);
                }
            }
            catch
            {
                // Tránh crash ứng dụng nếu file icon bị mất ngoài ý muốn
            }
        }
    }
}
