// ============================================================
//  FILE: Forms/frmSendNotification.cs
//  Admin gửi thông báo nội bộ đến một hoặc nhiều user khác
// ============================================================
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using MPR_Managerment.Common;
using MPR_Managerment.Models;
using MPR_Managerment.Services;

namespace MPR_Managerment.Forms
{
    public class frmSendNotification : frmBase
    {
        private readonly InternalNotificationService _svc = new();
        private readonly UserService _userSvc = new();
        private List<AppUser> _allUsers = new();

        // Controls
        private CheckedListBox clbReceivers;
        private TextBox txtSearch;
        private TextBox txtTitle;
        private RichTextBox rtbContent;
        private Button btnSend;
        private Button btnSelectAll;
        private Button btnClearAll;
        private Label lblReceiverCount;

        public frmSendNotification()
        {
            InitializeComponent();
            frmAIChat.Attach(this);
            LoadUsers();
        }

        private void InitializeComponent()
        {
            this.Text = "Gửi thông báo nội bộ";
            this.Size = new Size(880, 620);
            this.MinimumSize = new Size(750, 550);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(245, 245, 245);

            // ── Header banner ─────────────────────────────────────────────────
            var pnlHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 55,
                BackColor = Color.FromArgb(0, 120, 212)
            };
            pnlHeader.Controls.Add(new Label
            {
                Text = "📢  Gửi thông báo nội bộ",
                Font = new Font("Segoe UI", 13, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(15, 12),
                Size = new Size(500, 32)
            });
            // ── Panel chính (bên dưới header) ────────────────────────────────
            // Phải add Fill trước Top để WinForms dock đúng thứ tự
            var pnlMain = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(245, 245, 245),
                Padding = new Padding(12)
            };
            this.Controls.Add(pnlMain);
            this.Controls.Add(pnlHeader);

            // ── Cột trái: danh sách người nhận ────────────────────────────────
            var pnlLeft = new Panel
            {
                Location = new Point(12, 12),
                Size = new Size(290, 490),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom
            };
            pnlMain.Controls.Add(pnlLeft);

            pnlLeft.Controls.Add(new Label
            {
                Text = "NGƯỜI NHẬN",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 10),
                Size = new Size(270, 20)
            });

            txtSearch = new TextBox
            {
                Location = new Point(10, 36),
                Size = new Size(268, 26),
                Font = new Font("Segoe UI", 9),
                PlaceholderText = "🔍 Tìm tên / username..."
            };
            txtSearch.TextChanged += TxtSearch_TextChanged;
            pnlLeft.Controls.Add(txtSearch);

            // Nút Chọn tất cả / Bỏ chọn
            btnSelectAll = new Button
            {
                Text = "✅ Chọn tất cả",
                Location = new Point(10, 68),
                Size = new Size(130, 26),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnSelectAll.FlatAppearance.BorderSize = 0;
            btnSelectAll.Click += (s, e) => SetAllChecked(true);
            pnlLeft.Controls.Add(btnSelectAll);

            btnClearAll = new Button
            {
                Text = "❌ Bỏ chọn",
                Location = new Point(148, 68),
                Size = new Size(130, 26),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnClearAll.FlatAppearance.BorderSize = 0;
            btnClearAll.Click += (s, e) => SetAllChecked(false);
            pnlLeft.Controls.Add(btnClearAll);

            clbReceivers = new CheckedListBox
            {
                Location = new Point(10, 100),
                Size = new Size(268, 360),
                Font = new Font("Segoe UI", 9),
                CheckOnClick = true,
                BorderStyle = BorderStyle.None,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom | AnchorStyles.Right
            };
            clbReceivers.ItemCheck += (s, e) => UpdateReceiverCount();
            pnlLeft.Controls.Add(clbReceivers);

            lblReceiverCount = new Label
            {
                Text = "Đã chọn: 0 người",
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.Gray,
                Location = new Point(10, 466),
                Size = new Size(268, 18),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            pnlLeft.Controls.Add(lblReceiverCount);

            // ── Cột phải: soạn thông báo ──────────────────────────────────────
            var pnlRight = new Panel
            {
                Location = new Point(314, 12),
                Size = new Size(530, 490),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            pnlMain.Controls.Add(pnlRight);

            pnlRight.Controls.Add(new Label
            {
                Text = "NỘI DUNG THÔNG BÁO",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 10),
                Size = new Size(500, 20)
            });

            // Người gửi (readonly)
            pnlRight.Controls.Add(new Label
            {
                Text = "Người gửi:",
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.Gray,
                Location = new Point(10, 42),
                Size = new Size(80, 18)
            });
            string senderText = AppSession.CurrentUser != null
                ? $"{AppSession.CurrentUser.Full_Name}  ({AppSession.CurrentUser.Username})"
                : "Admin";
            pnlRight.Controls.Add(new TextBox
            {
                Text = senderText,
                Location = new Point(92, 39),
                Size = new Size(420, 26),
                Font = new Font("Segoe UI", 9),
                ReadOnly = true,
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.FromArgb(60, 60, 60)
            });

            // Tiêu đề
            pnlRight.Controls.Add(new Label
            {
                Text = "Tiêu đề (*):",
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.Gray,
                Location = new Point(10, 78),
                Size = new Size(80, 18)
            });
            txtTitle = new TextBox
            {
                Location = new Point(92, 75),
                Size = new Size(420, 26),
                Font = new Font("Segoe UI", 10),
                MaxLength = 500
            };
            pnlRight.Controls.Add(txtTitle);

            // Nội dung
            pnlRight.Controls.Add(new Label
            {
                Text = "Nội dung (*):",
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.Gray,
                Location = new Point(10, 115),
                Size = new Size(100, 18)
            });
            rtbContent = new RichTextBox
            {
                Location = new Point(10, 136),
                Size = new Size(502, 300),
                Font = new Font("Segoe UI", 10),
                BorderStyle = BorderStyle.FixedSingle,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            pnlRight.Controls.Add(rtbContent);

            // Nút gửi
            btnSend = new Button
            {
                Text = "📤  Gửi thông báo",
                Location = new Point(10, 450),
                Size = new Size(180, 36),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            btnSend.FlatAppearance.BorderSize = 0;
            btnSend.Click += BtnSend_Click;
            pnlRight.Controls.Add(btnSend);

            // Nút xóa form
            var btnClear = new Button
            {
                Text = "🔄 Xóa nội dung",
                Location = new Point(200, 450),
                Size = new Size(145, 36),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            btnClear.FlatAppearance.BorderSize = 0;
            btnClear.Click += (s, e) => { txtTitle.Clear(); rtbContent.Clear(); };
            pnlRight.Controls.Add(btnClear);
        }

        // ─────────────────────────────────────────────────────────────────────
        //  NẠP DANH SÁCH USER — bỏ bản thân admin gửi khỏi danh sách
        // ─────────────────────────────────────────────────────────────────────
        private void LoadUsers()
        {
            _allUsers = _userSvc.GetAll()
                .Where(u => u.Is_Active)
                .OrderBy(u => u.Full_Name)
                .ToList();
            PopulateCheckedList(_allUsers);
        }

        private void PopulateCheckedList(List<AppUser> users)
        {
            clbReceivers.Items.Clear();
            foreach (var u in users)
            {
                string display = string.IsNullOrWhiteSpace(u.Full_Name)
                    ? u.Username
                    : $"{u.Full_Name}  ({u.Username})";
                clbReceivers.Items.Add(new UserItem(u, display));
            }
            UpdateReceiverCount();
        }

        private void TxtSearch_TextChanged(object sender, EventArgs e)
        {
            string kw = txtSearch.Text.Trim();
            var filtered = string.IsNullOrEmpty(kw)
                ? _allUsers
                : _allUsers.Where(u =>
                    u.Username.Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                    (u.Full_Name ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase)).ToList();
            PopulateCheckedList(filtered);
        }

        private void SetAllChecked(bool check)
        {
            for (int i = 0; i < clbReceivers.Items.Count; i++)
                clbReceivers.SetItemChecked(i, check);
            UpdateReceiverCount();
        }

        private void UpdateReceiverCount()
        {
            int count = clbReceivers.CheckedItems.Count;
            lblReceiverCount.Text = $"Đã chọn: {count} người nhận";
            lblReceiverCount.ForeColor = count > 0 ? Color.FromArgb(40, 167, 69) : Color.Gray;
        }

        // ─────────────────────────────────────────────────────────────────────
        //  GỬI THÔNG BÁO
        // ─────────────────────────────────────────────────────────────────────
        private void BtnSend_Click(object sender, EventArgs e)
        {
            if (clbReceivers.CheckedItems.Count == 0)
            {
                MessageBox.Show(this, "Vui lòng chọn ít nhất một người nhận!", "Thiếu thông tin",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string title = txtTitle.Text.Trim();
            string content = rtbContent.Text.Trim();

            if (string.IsNullOrWhiteSpace(title))
            {
                MessageBox.Show(this, "Vui lòng nhập tiêu đề thông báo!", "Thiếu thông tin",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTitle.Focus();
                return;
            }
            if (string.IsNullOrWhiteSpace(content))
            {
                MessageBox.Show(this, "Vui lòng nhập nội dung thông báo!", "Thiếu thông tin",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                rtbContent.Focus();
                return;
            }

            var receivers = clbReceivers.CheckedItems
                .Cast<UserItem>()
                .Select(ui => (ui.User.Username, ui.User.Full_Name ?? ""))
                .ToList();

            string senderUser = AppSession.CurrentUser?.Username ?? "admin";
            string senderFull = AppSession.CurrentUser?.Full_Name ?? "Admin";

            // Xác nhận trước khi gửi
            int cnt = receivers.Count;
            string preview = cnt <= 5
                ? string.Join(", ", receivers.Select(r => r.Username))
                : string.Join(", ", receivers.Take(5).Select(r => r.Username)) + $" ... và {cnt - 5} người khác";

            if (MessageBox.Show(this,
                $"Xác nhận gửi thông báo:\n\n" +
                $"📌 Tiêu đề: {title}\n" +
                $"👥 Đến: {preview}\n" +
                $"📊 Tổng: {cnt} người nhận\n\n" +
                $"Tiếp tục?",
                "Xác nhận gửi",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

            try
            {
                btnSend.Enabled = false;
                btnSend.Text = "⏳ Đang gửi...";

                _svc.Send(senderUser, senderFull, receivers, title, content);

                MessageBox.Show(this,
                    $"✅ Đã gửi thông báo thành công!\n\n" +
                    $"  • Tiêu đề: {title}\n" +
                    $"  • Số người nhận: {cnt}",
                    "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);

                txtTitle.Clear();
                rtbContent.Clear();
                SetAllChecked(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Lỗi khi gửi thông báo:\n" + ex.Message,
                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnSend.Enabled = true;
                btnSend.Text = "📤  Gửi thông báo";
            }
        }

        // Wrapper lưu cả AppUser lẫn text hiển thị
        private class UserItem
        {
            public AppUser User { get; }
            private readonly string _display;
            public UserItem(AppUser user, string display) { User = user; _display = display; }
            public override string ToString() => _display;
        }
    }
}
