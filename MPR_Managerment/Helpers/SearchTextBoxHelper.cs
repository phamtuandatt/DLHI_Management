using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace MPR_Managerment.Helpers
{
    public class SearchTextBoxHelper : Form
    {
        // Gọi các hàm thư viện Windows API gốc để tùy chỉnh Margin
        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wp, IntPtr lp);

        private const int EM_SETMARGINS = 0xd3;
        private const int EC_LEFTMARGIN = 1;

        // Khởi tạo hình ảnh kính lúp mặc định
        private Bitmap searchIcon;

        public SearchTextBoxHelper()
        {
            // Bạn có thể thay thế bằng icon tùy chọn từ file hoặc Resources của bạn
            searchIcon = new Bitmap(16, 16);
            using (Graphics g = Graphics.FromImage(searchIcon))
            {
                // Vẽ một icon kính lúp tạm thời bằng code nếu chưa có sẵn file ảnh
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                Pen pen = new Pen(Color.Gray, 2);
                g.DrawEllipse(pen, 2, 2, 8, 8);
                g.DrawLine(pen, 9, 9, 14, 14);
            }

            // Tạo khoảng trống bên trái cho TextBox để chứa icon
            // Chiều rộng là 24 pixel (đủ cho icon kích cỡ 16x16 + khoảng đệm)
            SendMessage(this.Handle, EM_SETMARGINS, (IntPtr)EC_LEFTMARGIN, (IntPtr)(24 << 0));
        }

        // Lắng nghe thông điệp vẽ (WM_PAINT) để tự vẽ icon lên góc trái
        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            // Mã lệnh 15 tương ứng với sự kiện vẽ lại giao diện WM_PAINT
            if (m.Msg == 15 && searchIcon != null)
            {
                using (Graphics g = Graphics.FromHwnd(this.Handle))
                {
                    // Căn lề cho icon nằm giữa theo chiều dọc của TextBox
                    int yPos = (this.ClientSize.Height - searchIcon.Height) / 2;
                    g.DrawImage(searchIcon, 4, yPos);
                }
            }
        }
    }
}
