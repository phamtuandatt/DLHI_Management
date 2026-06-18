// ============================================================
//  FILE: Forms/frmNotificationToast.cs
//  Toast cảnh báo nổi — hiện khi nhận thông báo nội bộ mới
// ============================================================
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using MPR_Managerment.Models;

namespace MPR_Managerment.Forms
{
    /// <summary>
    /// Popup cảnh báo nổi ở góc dưới-phải màn hình.
    /// Dùng ShowToasts() để hiện một hoặc nhiều thông báo liên tiếp.
    /// </summary>
    public class frmNotificationToast : Form
    {
        // ── Queue tĩnh để xếp hàng — mỗi lúc chỉ 1 toast hiển thị ──────────
        private static readonly Queue<InternalNotification> _queue = new();
        private static bool _showing = false;

        /// <summary>Đưa danh sách thông báo vào queue và bắt đầu hiển thị nếu chưa có toast đang chạy.</summary>
        public static void ShowToasts(IEnumerable<InternalNotification> notifications, Form owner = null)
        {
            foreach (var n in notifications)
                _queue.Enqueue(n);
            if (!_showing)
                ShowNext(owner);
        }

        private static void ShowNext(Form owner)
        {
            if (_queue.Count == 0) { _showing = false; return; }
            _showing = true;
            var n = _queue.Dequeue();
            var toast = new frmNotificationToast(n, owner);
            toast.Show();
        }

        // ── Instance ─────────────────────────────────────────────────────────
        private readonly InternalNotification _notif;
        private readonly Form _owner;
        private System.Windows.Forms.Timer _autoCloseTimer;
        private int _secondsLeft = 10;
        private Label _lblCountdown;

        private frmNotificationToast(InternalNotification notif, Form owner)
        {
            _notif = notif;
            _owner = owner;
            BuildUI();
            PositionOnScreen();
        }

        private void BuildUI()
        {
            // Form properties
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            Size = new Size(380, 160);
            BackColor = Color.White;

            // Hiệu ứng fade-in
            Opacity = 0;

            // ── Thanh tiêu đề đỏ (cảnh báo) ──────────────────────────────────
            var pnlTitle = new Panel
            {
                Dock = DockStyle.Top,
                Height = 42,
                BackColor = Color.FromArgb(198, 40, 40)
            };

            var lblIcon = new Label
            {
                Text = "🔔",
                Font = new Font("Segoe UI", 14),
                ForeColor = Color.White,
                Location = new Point(10, 8),
                Size = new Size(32, 28),
                AutoSize = false
            };
            pnlTitle.Controls.Add(lblIcon);

            var lblTitle = new Label
            {
                Text = TruncateText(_notif.Title, 38),
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(46, 5),
                Size = new Size(270, 32),
                AutoSize = false
            };
            pnlTitle.Controls.Add(lblTitle);

            var btnClose = new Label
            {
                Text = "✕",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(350, 0),
                Size = new Size(30, 42),
                TextAlign = ContentAlignment.MiddleCenter,
                Cursor = Cursors.Hand
            };
            btnClose.Click += (s, e) => Close();
            btnClose.MouseEnter += (s, e) => ((Label)s).BackColor = Color.FromArgb(220, 60, 60);
            btnClose.MouseLeave += (s, e) => ((Label)s).BackColor = Color.Transparent;
            pnlTitle.Controls.Add(btnClose);

            // ── Nội dung ──────────────────────────────────────────────────────
            var pnlBody = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(10, 8, 10, 8)
            };

            var lblSender = new Label
            {
                Text = $"Từ: {_notif.Sender_FullName}  ({_notif.Sent_At:dd/MM/yyyy HH:mm})",
                Font = new Font("Segoe UI", 8, FontStyle.Italic),
                ForeColor = Color.Gray,
                Location = new Point(10, 50),
                Size = new Size(356, 18),
                AutoSize = false
            };

            var lblContent = new Label
            {
                Text = TruncateText(_notif.Content, 90),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.FromArgb(40, 40, 40),
                Location = new Point(10, 72),
                Size = new Size(356, 36),
                AutoSize = false
            };

            _lblCountdown = new Label
            {
                Text = $"Tự đóng sau {_secondsLeft}s",
                Font = new Font("Segoe UI", 7.5f),
                ForeColor = Color.Gray,
                Location = new Point(10, 112),
                Size = new Size(200, 18),
                AutoSize = false
            };

            var btnRead = new Button
            {
                Text = "Xem chi tiết",
                Location = new Point(236, 108),
                Size = new Size(130, 26),
                BackColor = Color.FromArgb(198, 40, 40),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnRead.FlatAppearance.BorderSize = 0;
            btnRead.Click += BtnRead_Click;

            pnlBody.Controls.Add(lblSender);
            pnlBody.Controls.Add(lblContent);
            pnlBody.Controls.Add(_lblCountdown);
            pnlBody.Controls.Add(btnRead);

            // Thứ tự add: Fill trước Top
            Controls.Add(pnlBody);
            Controls.Add(pnlTitle);

            // Bo viền
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 10, 10));
        }

        private void PositionOnScreen()
        {
            var screen = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1920, 1080);
            Left = screen.Right - Width - 20;
            Top  = screen.Bottom - Height - 20;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            // Fade-in
            var fadeTimer = new System.Windows.Forms.Timer { Interval = 30 };
            fadeTimer.Tick += (s, _) =>
            {
                Opacity += 0.1;
                if (Opacity >= 1.0) { Opacity = 1.0; fadeTimer.Stop(); }
            };
            fadeTimer.Start();

            // Auto-close countdown
            _autoCloseTimer = new System.Windows.Forms.Timer { Interval = 1000 };
            _autoCloseTimer.Tick += (s, _) =>
            {
                _secondsLeft--;
                if (_secondsLeft <= 0) { _autoCloseTimer.Stop(); CloseAndNext(); }
                else _lblCountdown.Text = $"Tự đóng sau {_secondsLeft}s";
            };
            _autoCloseTimer.Start();
        }

        private void BtnRead_Click(object sender, EventArgs e)
        {
            _autoCloseTimer?.Stop();
            // Mở dialog chi tiết
            ShowDetailDialog();
            CloseAndNext();
        }

        private void ShowDetailDialog()
        {
            using var dlg = new Form
            {
                Text = "📢 Thông báo nội bộ",
                Size = new Size(520, 400),
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                TopMost = true,
                BackColor = Color.White
            };

            var pnlTop = new Panel
            {
                Dock = DockStyle.Top,
                Height = 55,
                BackColor = Color.FromArgb(198, 40, 40)
            };
            pnlTop.Controls.Add(new Label
            {
                Text = $"🔔  {_notif.Title}",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(12, 12),
                Size = new Size(480, 30),
                AutoSize = false
            });

            var pnlInfo = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(16),
                BackColor = Color.White
            };

            var lblMeta = new Label
            {
                Text = $"Từ: {_notif.Sender_FullName}  •  {_notif.Sent_At:dd/MM/yyyy HH:mm:ss}",
                Font = new Font("Segoe UI", 9, FontStyle.Italic),
                ForeColor = Color.Gray,
                Dock = DockStyle.Top,
                Height = 24
            };

            var sep = new Panel
            {
                Dock = DockStyle.Top,
                Height = 1,
                BackColor = Color.FromArgb(220, 220, 220)
            };

            var rtb = new RichTextBox
            {
                Dock = DockStyle.Fill,
                Text = _notif.Content,
                Font = new Font("Segoe UI", 10),
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                BackColor = Color.White,
                ScrollBars = RichTextBoxScrollBars.Vertical
            };

            var pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 48,
                BackColor = Color.FromArgb(245, 245, 245)
            };
            var btnOk = new Button
            {
                Text = "Đã đọc  ✓",
                Size = new Size(120, 32),
                Location = new Point(12, 8),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.OK
            };
            btnOk.FlatAppearance.BorderSize = 0;
            pnlBottom.Controls.Add(btnOk);

            pnlInfo.Controls.Add(rtb);
            pnlInfo.Controls.Add(sep);
            pnlInfo.Controls.Add(lblMeta);

            dlg.Controls.Add(pnlInfo);
            dlg.Controls.Add(pnlBottom);
            dlg.Controls.Add(pnlTop);

            dlg.ShowDialog(_owner);

            // Đánh dấu đã đọc trên DB
            try
            {
                var svc = new Services.InternalNotificationService();
                svc.MarkRead(_notif.Notif_ID);
            }
            catch { }
        }

        private void CloseAndNext()
        {
            _autoCloseTimer?.Stop();
            // Fade-out rồi đóng
            var fade = new System.Windows.Forms.Timer { Interval = 30 };
            fade.Tick += (s, _) =>
            {
                Opacity -= 0.15;
                if (Opacity <= 0)
                {
                    fade.Stop();
                    Close();
                    ShowNext(_owner);
                }
            };
            fade.Start();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);
            _autoCloseTimer?.Stop();
        }

        private static string TruncateText(string text, int maxLen)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return text.Length <= maxLen ? text : text.Substring(0, maxLen - 3) + "...";
        }

        [System.Runtime.InteropServices.DllImport("Gdi32.dll")]
        private static extern IntPtr CreateRoundRectRgn(int nLeftRect, int nTopRect, int nRightRect, int nBottomRect, int nWidthEllipse, int nHeightEllipse);
    }
}
