using System;
using System.Drawing;
using System.Windows.Forms;

namespace MPR_Managerment.Helpers
{
    /// <summary>
    /// Tạo và quản lý thông báo nổi (toast) trên Form.
    /// Dùng: var toast = ToastHelper.Attach(form);
    ///       toast.Show("⏳ Đang tải...");
    ///       toast.Hide();
    /// </summary>
    public class ToastHelper
    {
        private readonly Label _label;
        private readonly Control _host;

        private ToastHelper(Control host, Label label)
        {
            _host = host;
            _label = label;

            host.Resize += (s, e) => Center();
            host.Controls.Add(_label);
            _label.BringToFront();
        }

        /// <summary>Gắn toast vào form/control chỉ định.</summary>
        public static ToastHelper Attach(Control host)
        {
            var lbl = new Label
            {
                AutoSize = false,
                Size = new Size(340, 50),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(220, 40, 44, 56),
                BorderStyle = BorderStyle.None,
                Visible = false,
                Anchor = AnchorStyles.None,
                Padding = new Padding(8)
            };
            var toast = new ToastHelper(host, lbl);
            toast.Center();
            return toast;
        }

        /// <summary>Hiển thị toast với nội dung text.</summary>
        public void Show(string message)
        {
            _label.Text = message;
            Center();
            _label.Visible = true;
            _label.BringToFront();
        }

        /// <summary>Ẩn toast.</summary>
        public new void Hide() => _label.Visible = false;

        /// <summary>True nếu toast đang hiện.</summary>
        public bool IsVisible => _label.Visible;

        private void Center()
        {
            _label.Location = new Point(
                (_host.ClientSize.Width  - _label.Width)  / 2,
                (_host.ClientSize.Height - _label.Height) / 2);
        }
    }
}
