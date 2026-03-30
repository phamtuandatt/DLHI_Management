// ============================================================
//  FILE: Forms/frmLogin.cs
// ============================================================
using System;
using System.Drawing;
using System.Windows.Forms;
using MPR_Managerment.Models;
using MPR_Managerment.Services;

namespace MPR_Managerment.Forms
{
    public partial class frmLogin : Form
    {
        private readonly UserService _userService = new UserService();

        private TextBox txtUsername, txtPassword;
        private Button btnLogin, btnExit;
        private Label lblError;
        private CheckBox chkShowPwd;
        private int _failCount = 0;

        public frmLogin()
        {
            InitializeComponent();
            BuildUI();
        }

        private void BuildUI()
        {
            this.Text = "Đăng nhập — MPR Management System";
            this.Size = new Size(440, 380);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.White;

            // Header
            var pHeader = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(440, 80),
                BackColor = Color.FromArgb(0, 120, 212)
            };
            pHeader.Controls.Add(new Label
            {
                Text = "⚙ MPR MANAGEMENT SYSTEM",
                Font = new Font("Segoe UI", 13, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(15, 12),
                Size = new Size(410, 30)
            });
            pHeader.Controls.Add(new Label
            {
                Text = "Vui lòng đăng nhập để tiếp tục",
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.FromArgb(200, 230, 255),
                Location = new Point(15, 45),
                Size = new Size(400, 22)
            });
            this.Controls.Add(pHeader);

            // Username
            AddFormLabel("Tên đăng nhập:", 100);
            txtUsername = new TextBox
            {
                Location = new Point(30, 125),
                Size = new Size(370, 30),
                Font = new Font("Segoe UI", 11),
                PlaceholderText = "Nhập username..."
            };
            this.Controls.Add(txtUsername);

            // Password
            AddFormLabel("Mật khẩu:", 165);
            txtPassword = new TextBox
            {
                Location = new Point(30, 190),
                Size = new Size(370, 30),
                Font = new Font("Segoe UI", 11),
                PasswordChar = '●',
                PlaceholderText = "Nhập mật khẩu..."
            };
            txtPassword.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) BtnLogin_Click(null, null); };
            this.Controls.Add(txtPassword);

            // Show password
            chkShowPwd = new CheckBox
            {
                Text = "Hiện mật khẩu",
                Location = new Point(30, 228),
                Size = new Size(140, 22),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.Gray
            };
            chkShowPwd.CheckedChanged += (s, e) =>
                txtPassword.PasswordChar = chkShowPwd.Checked ? '\0' : '●';
            this.Controls.Add(chkShowPwd);

            // Error label
            lblError = new Label
            {
                Location = new Point(30, 258),
                Size = new Size(370, 22),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(220, 53, 69),
                Text = ""
            };
            this.Controls.Add(lblError);

            // Buttons
            btnLogin = new Button
            {
                Text = "ĐĂNG NHẬP",
                Location = new Point(30, 290),
                Size = new Size(200, 38),
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnLogin.FlatAppearance.BorderSize = 0;
            btnLogin.Click += BtnLogin_Click;
            this.Controls.Add(btnLogin);

            btnExit = new Button
            {
                Text = "Thoát",
                Location = new Point(245, 290),
                Size = new Size(155, 38),
                Font = new Font("Segoe UI", 10),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnExit.FlatAppearance.BorderSize = 0;
            btnExit.Click += (s, e) => Application.Exit();
            this.Controls.Add(btnExit);

            this.AcceptButton = btnLogin;
            txtUsername.Focus();
        }

        private void AddFormLabel(string text, int y)
        {
            this.Controls.Add(new Label
            {
                Text = text,
                Location = new Point(30, y),
                Size = new Size(200, 22),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(50, 50, 50)
            });
        }

        private void BtnLogin_Click(object sender, EventArgs e)
        {
            lblError.Text = "";
            string username = txtUsername.Text.Trim();
            string password = txtPassword.Text;

            if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
            {
                lblError.Text = "⚠ Vui lòng nhập đầy đủ thông tin!";
                return;
            }

            btnLogin.Enabled = false;
            btnLogin.Text = "Đang kiểm tra...";

            try
            {
                var result = _userService.Login(username, password);

                if (!result.Success)
                {
                    _failCount++;
                    lblError.Text = $"✗ {result.Message}";

                    if (_failCount >= 5)
                    {
                        lblError.Text = "✗ Quá nhiều lần đăng nhập sai. Vui lòng liên hệ Admin!";
                        btnLogin.Enabled = false;
                        return;
                    }

                    txtPassword.Clear();
                    txtPassword.Focus();
                    return;
                }

                // Lưu session
                AppSession.CurrentUser = result.User;
                AppSession.Permissions = result.Permissions;

                // Cần đổi mật khẩu lần đầu?
                if (result.User!.Must_Change_Password)
                {
                    var changePwd = new frmChangePassword(result.User.User_ID, isForced: true);
                    changePwd.ShowDialog(this);
                }

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                lblError.Text = "✗ Lỗi kết nối: " + ex.Message;
            }
            finally
            {
                btnLogin.Enabled = true;
                btnLogin.Text = "ĐĂNG NHẬP";
            }
        }
    }
}
