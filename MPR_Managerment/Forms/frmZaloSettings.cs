using MPR_Managerment.Services;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace MPR_Managerment.Forms
{
    /// <summary>
    /// Form cấu hình hàng đợi Zalo (Delay, Retry) và kiểm tra gửi tin.
    /// </summary>
    public class frmZaloSettings : Form
    {
        private NumericUpDown nudDelay, nudMaxRetries, nudRetryDelay;
        private Button btnSave, btnCancel, btnTest;
        private Label lblQueueStatus;
        private System.Windows.Forms.Timer _statusTimer;

        public frmZaloSettings()
        {
            // Gắn AI Trợ lý — hiển thị nút floating + chat panel
            frmAIChat.Attach(this);
            BuildUI();
            LoadSettings();
            StartStatusTimer();
        }

        private void BuildUI()
        {
            this.Text = "⚙ Cấu hình Hàng đợi Zalo Notification";
            this.Size = new Size(620, 320);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.BackColor = Color.FromArgb(245, 245, 245);

            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20)
            };
            this.Controls.Add(panel);

            int y = 20;

            // ── Tiêu đề ──────────────────────────────────────
            panel.Controls.Add(new Label
            {
                Text = "CẤU HÌNH HÀNG ĐỢI GỬI TIN",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 150, 60),
                Location = new Point(20, y),
                Size = new Size(560, 28)
            });

            y += 40;

            // ── Delay giữa tin ────────────────────────────────
            AddLbl(panel, "Delay giữa tin (ms):", 20, y);
            nudDelay = new NumericUpDown
            {
                Location = new Point(165, y),
                Size = new Size(120, 25),
                Font = new Font("Segoe UI", 9),
                Minimum = 1000,
                Maximum = 60000,
                Value = 2000,
                Increment = 500
            };
            panel.Controls.Add(nudDelay);

            panel.Controls.Add(new Label
            {
                Text = "ms (tránh Zalo chặn spam, khuyến nghị ≥ 2000)",
                Location = new Point(295, y + 3),
                Size = new Size(350, 20),
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = Color.Gray
            });

            y += 38;

            // ── Max Retries ────────────────────────────────────
            AddLbl(panel, "Số lần thử lại:", 20, y);
            nudMaxRetries = new NumericUpDown
            {
                Location = new Point(165, y),
                Size = new Size(80, 25),
                Font = new Font("Segoe UI", 9),
                Minimum = 0,
                Maximum = 20,
                Value = 5
            };
            panel.Controls.Add(nudMaxRetries);

            AddLbl(panel, "Chờ thử lại (ms):", 265, y);
            nudRetryDelay = new NumericUpDown
            {
                Location = new Point(380, y),
                Size = new Size(100, 25),
                Font = new Font("Segoe UI", 9),
                Minimum = 1000,
                Maximum = 60000,
                Value = 5000,
                Increment = 1000
            };
            panel.Controls.Add(nudRetryDelay);

            y += 50;

            // ── Trạng thái queue ─────────────────────────────
            lblQueueStatus = new Label
            {
                Location = new Point(20, y),
                Size = new Size(560, 22),
                Font = new Font("Segoe UI", 9, FontStyle.Italic),
                ForeColor = Color.FromArgb(0, 120, 212),
                Text = "🔄 Đang tải trạng thái..."
            };
            panel.Controls.Add(lblQueueStatus);

            y += 40;

            // ── Buttons ──────────────────────────────────────
            btnTest = new Button
            {
                Text = "🧪 Gửi test",
                Location = new Point(20, y),
                Size = new Size(110, 32),
                BackColor = Color.FromArgb(255, 140, 0),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnTest.FlatAppearance.BorderSize = 0;
            btnTest.Click += BtnTest_Click;
            panel.Controls.Add(btnTest);

            btnSave = new Button
            {
                Text = "💾 Lưu cài đặt",
                Location = new Point(350, y),
                Size = new Size(120, 32),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Click += BtnSave_Click;
            panel.Controls.Add(btnSave);

            btnCancel = new Button
            {
                Text = "Đóng",
                Location = new Point(480, y),
                Size = new Size(90, 32),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9),
                Cursor = Cursors.Hand
            };
            btnCancel.FlatAppearance.BorderSize = 0;
            btnCancel.Click += (s, e) => this.Close();
            panel.Controls.Add(btnCancel);
        }

        private void AddLbl(Panel p, string text, int x, int y)
        {
            p.Controls.Add(new Label
            {
                Text = text,
                Location = new Point(x, y + 3),
                Size = new Size(140, 20),
                Font = new Font("Segoe UI", 9)
            });
        }

        private void LoadSettings()
        {
            var s = ZaloNotificationService.LoadQueueSettings();
            nudDelay.Value = Math.Max(nudDelay.Minimum, Math.Min(nudDelay.Maximum, s.DelayBetweenMessagesMs));
            nudMaxRetries.Value = Math.Max(0, Math.Min(20, s.MaxRetries));
            nudRetryDelay.Value = Math.Max(nudRetryDelay.Minimum, Math.Min(nudRetryDelay.Maximum, s.RetryDelayMs));
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            var s = new ZaloQueueSettings
            {
                DelayBetweenMessagesMs = (int)nudDelay.Value,
                MaxRetries = (int)nudMaxRetries.Value,
                RetryDelayMs = (int)nudRetryDelay.Value
            };
            ZaloNotificationService.SaveQueueSettings(s);
            MessageBox.Show(this, "Đã lưu cài đặt hàng đợi Zalo!", "Thành công",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnTest_Click(object sender, EventArgs e)
        {
            // Hỏi tên nhóm để test
            string groupName = Microsoft.VisualBasic.Interaction.InputBox(
                "Nhập Tên nhóm Zalo chính xác để gửi tin thử:", "Test Zalo Queue", "");

            if (string.IsNullOrWhiteSpace(groupName)) return;

            // Lưu settings tạm rồi enqueue
            BtnSave_Click(null, null);

            ZaloNotificationService.Instance.Enqueue(
                groupName,
                $"🧪 [TEST QUEUE] Đây là tin nhắn thử từ hệ thống MPR\n🕐 {DateTime.Now:dd/MM/yyyy HH:mm:ss}",
                "TEST");

            MessageBox.Show(this, $"Đã đặt tin nhắn test vào hàng đợi cho nhóm '{groupName}'!\nQueue hiện có: {ZaloNotificationService.Instance.PendingCount} tin.",
                "Đã enqueue", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void StartStatusTimer()
        {
            _statusTimer = new System.Windows.Forms.Timer { Interval = 1000 };
            _statusTimer.Tick += (s, e) =>
            {
                int cnt = ZaloNotificationService.Instance.PendingCount;
                lblQueueStatus.Text = cnt == 0
                    ? "✅ Hàng đợi trống — tất cả thông báo đã gửi"
                    : $"⏳ Đang chờ gửi: {cnt} thông báo trong hàng đợi...";
                lblQueueStatus.ForeColor = cnt == 0
                    ? Color.FromArgb(40, 167, 69)
                    : Color.FromArgb(255, 140, 0);
            };
            _statusTimer.Start();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _statusTimer?.Stop();
            _statusTimer?.Dispose();
            base.OnFormClosed(e);
        }
    }
}