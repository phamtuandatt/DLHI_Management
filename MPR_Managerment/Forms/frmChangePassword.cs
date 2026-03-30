

// ============================================================
//  FILE: Forms/frmChangePassword.cs
// ============================================================
using System;
using System.Drawing;
using System.Windows.Forms;
using MPR_Managerment.Services;

namespace MPR_Managerment.Forms
{
    public partial class frmChangePassword : Form
    {
        private readonly UserService _svc = new UserService();
        private readonly int _userId;
        private readonly bool _isForced;

        private TextBox txtOld, txtNew, txtConfirm;
        private Label lblError;
        private Button btnSave, btnCancel;

        public frmChangePassword(int userId, bool isForced = false)
        {
            _userId = userId;
            _isForced = isForced;
            InitializeComponent();
            BuildUI();
        }

        private void BuildUI()
        {
            this.Text = _isForced ? "Đổi mật khẩu (bắt buộc)" : "Đổi mật khẩu";
            this.Size = new Size(400, 340);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.White;

            if (_isForced)
                this.Controls.Add(new Label
                {
                    Text = "⚠ Bạn cần đổi mật khẩu trước khi sử dụng hệ thống!",
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    ForeColor = Color.FromArgb(255, 140, 0),
                    Location = new Point(15, 15),
                    Size = new Size(360, 40)
                });

            int startY = _isForced ? 60 : 20;

            AddRow("Mật khẩu hiện tại:", startY, out txtOld);
            AddRow("Mật khẩu mới:", startY + 65, out txtNew);
            AddRow("Xác nhận mới:", startY + 130, out txtConfirm);
            txtOld.PasswordChar = txtNew.PasswordChar = txtConfirm.PasswordChar = '●';

            lblError = new Label
            {
                Location = new Point(15, startY + 200),
                Size = new Size(360, 22),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.Red
            };
            this.Controls.Add(lblError);

            btnSave = new Button
            {
                Text = "💾 Lưu mật khẩu",
                Location = new Point(15, startY + 228),
                Size = new Size(175, 36),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Click += BtnSave_Click;
            this.Controls.Add(btnSave);

            if (!_isForced)
            {
                btnCancel = new Button
                {
                    Text = "Hủy",
                    Location = new Point(200, startY + 228),
                    Size = new Size(175, 36),
                    Font = new Font("Segoe UI", 9),
                    BackColor = Color.FromArgb(108, 117, 125),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                btnCancel.FlatAppearance.BorderSize = 0;
                btnCancel.Click += (s, e) => this.Close();
                this.Controls.Add(btnCancel);
            }
        }

        private void AddRow(string label, int y, out TextBox txt)
        {
            this.Controls.Add(new Label
            {
                Text = label,
                Location = new Point(15, y),
                Size = new Size(200, 22),
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            });
            txt = new TextBox
            {
                Location = new Point(15, y + 25),
                Size = new Size(360, 28),
                Font = new Font("Segoe UI", 10)
            };
            this.Controls.Add(txt);
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            lblError.Text = "";

            if (string.IsNullOrEmpty(txtOld.Text) || string.IsNullOrEmpty(txtNew.Text))
            {
                lblError.Text = "Vui lòng nhập đầy đủ thông tin!";
                return;
            }
            if (txtNew.Text != txtConfirm.Text)
            {
                lblError.Text = "Mật khẩu mới không khớp!";
                return;
            }
            if (txtNew.Text.Length < 6)
            {
                lblError.Text = "Mật khẩu tối thiểu 6 ký tự!";
                return;
            }

            var (success, message) = _svc.ChangePassword(_userId, txtOld.Text, txtNew.Text);
            if (success)
            {
                MessageBox.Show("✅ " + message, "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                lblError.Text = "✗ " + message;
            }
        }
    }
}