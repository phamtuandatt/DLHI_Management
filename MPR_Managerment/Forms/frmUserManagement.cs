// ============================================================
//  FILE: Forms/frmUserManagement.cs
//  Quản lý User + phân quyền theo module
// ============================================================
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using MPR_Managerment.Models;
using MPR_Managerment.Services;

namespace MPR_Managerment.Forms
{
    public partial class frmUserManagement : Form
    {
        private readonly UserService _svc = new UserService();
        private List<AppUser> _users = new List<AppUser>();
        private List<Role> _roles = new List<Role>();
        private List<AppModule> _modules = new List<AppModule>();
        private int _selectedUserId = 0;

        // Controls
        private DataGridView dgvUsers;
        private DataGridView dgvPermissions;
        private TextBox txtUsername, txtFullName, txtEmail, txtPhone, txtDept, txtSearch;
        private TextBox txtNewPwd;
        private ComboBox cboRole;
        private CheckBox chkActive, chkMustChange;
        private Button btnNew, btnSave, btnDeactivate, btnResetPwd, btnSavePerms, btnResetPerms;
        private Label lblStatus;
        private Panel panelTop, panelForm, panelPerm;

        public frmUserManagement()
        {
            InitializeComponent();
            BuildUI();
            LoadData();
            this.Resize += (s, e) => ResizeControls();
        }

        // =====================================================================
        //  BUILD UI
        // =====================================================================
        private void BuildUI()
        {
            this.Text = "Quản lý Người dùng & Phân quyền";
            this.BackColor = Color.FromArgb(245, 245, 245);

            // ===== PANEL TOP: danh sách user =====
            panelTop = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(1360, 220),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(panelTop);

            panelTop.Controls.Add(new Label
            {
                Text = "DANH SÁCH NGƯỜI DÙNG",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 10),
                Size = new Size(400, 28)
            });

            txtSearch = new TextBox { Location = new Point(10, 48), Size = new Size(250, 28), Font = new Font("Segoe UI", 10), PlaceholderText = "Tìm username / tên..." };
            txtSearch.TextChanged += (s, e) => FilterUsers();
            panelTop.Controls.Add(txtSearch);

            btnNew = Btn("➕ Tạo user", Color.FromArgb(40, 167, 69), new Point(270, 47), 110, 30);
            btnNew.Click += BtnNew_Click;
            panelTop.Controls.Add(btnNew);

            lblStatus = new Label { Location = new Point(395, 52), Size = new Size(400, 22), Font = new Font("Segoe UI", 9), ForeColor = Color.Gray };
            panelTop.Controls.Add(lblStatus);

            dgvUsers = BuildGrid(panelTop, 85, 125);
            dgvUsers.ReadOnly = true;
            dgvUsers.SelectionChanged += DgvUsers_SelectionChanged;
            dgvUsers.CellFormatting += DgvUsers_CellFormatting;

            // ===== PANEL FORM: thông tin user =====
            panelForm = new Panel
            {
                Location = new Point(10, 240),
                Size = new Size(680, 230),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            this.Controls.Add(panelForm);

            panelForm.Controls.Add(new Label
            {
                Text = "THÔNG TIN NGƯỜI DÙNG",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 8),
                Size = new Size(400, 25)
            });

            int y = 38;
            AddLbl(panelForm, "Username (*):", 10, y); txtUsername = AddTxt(panelForm, 130, y, 160);
            AddLbl(panelForm, "Họ tên (*):", 310, y); txtFullName = AddTxt(panelForm, 430, y, 220);

            y += 38;
            AddLbl(panelForm, "Email:", 10, y); txtEmail = AddTxt(panelForm, 130, y, 200);
            AddLbl(panelForm, "SĐT:", 345, y); txtPhone = AddTxt(panelForm, 430, y, 100);
            AddLbl(panelForm, "Mật khẩu mới:", 540, y); txtNewPwd = AddTxt(panelForm, 640, y, 0);
            // txtNewPwd sẽ ẩn khi edit, hiện khi tạo mới — xử lý ở BtnNew_Click

            y += 38;
            AddLbl(panelForm, "Phòng ban:", 10, y); txtDept = AddTxt(panelForm, 130, y, 200);
            AddLbl(panelForm, "Role (*):", 345, y);
            cboRole = new ComboBox { Location = new Point(430, y), Size = new Size(160, 25), Font = new Font("Segoe UI", 9), DropDownStyle = ComboBoxStyle.DropDownList };
            panelForm.Controls.Add(cboRole);

            y += 38;
            chkActive = new CheckBox { Text = "Tài khoản Active", Location = new Point(10, y), Size = new Size(170, 22), Font = new Font("Segoe UI", 9), Checked = true };
            panelForm.Controls.Add(chkActive);
            chkMustChange = new CheckBox { Text = "Bắt buộc đổi mật khẩu", Location = new Point(190, y), Size = new Size(200, 22), Font = new Font("Segoe UI", 9) };
            panelForm.Controls.Add(chkMustChange);

            y += 40;
            btnSave = Btn("💾 Lưu", Color.FromArgb(0, 120, 212), new Point(10, y), 100, 32);
            btnResetPwd = Btn("🔑 Reset Password", Color.FromArgb(255, 140, 0), new Point(120, y), 145, 32);
            btnDeactivate = Btn("🚫 Vô hiệu hóa", Color.FromArgb(220, 53, 69), new Point(275, y), 140, 32);

            btnSave.Click += BtnSave_Click;
            btnResetPwd.Click += BtnResetPwd_Click;
            btnDeactivate.Click += BtnDeactivate_Click;
            panelForm.Controls.AddRange(new Control[] { btnSave, btnResetPwd, btnDeactivate });

            // ===== PANEL PERMISSION: bảng quyền =====
            panelPerm = new Panel
            {
                Location = new Point(700, 240),
                Size = new Size(670, 230),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(panelPerm);

            panelPerm.Controls.Add(new Label
            {
                Text = "PHÂN QUYỀN MODULE (Override riêng cho user)",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 8),
                Size = new Size(600, 25)
            });

            panelPerm.Controls.Add(new Label
            {
                Text = "⚠ Nếu để mặc định (không tick Override) → dùng quyền theo Role",
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.Gray,
                Location = new Point(10, 35),
                Size = new Size(600, 18)
            });

            dgvPermissions = BuildGrid(panelPerm, 58, 130);
            BuildPermissionColumns();

            btnSavePerms = Btn("💾 Lưu phân quyền", Color.FromArgb(0, 120, 212), new Point(10, 195), 155, 30);
            btnResetPerms = Btn("🔄 Reset về Role", Color.FromArgb(108, 117, 125), new Point(175, 195), 145, 30);
            btnSavePerms.Click += BtnSavePerms_Click;
            btnResetPerms.Click += BtnResetPerms_Click;
            panelPerm.Controls.AddRange(new Control[] { btnSavePerms, btnResetPerms });
        }

        private void BuildPermissionColumns()
        {
            dgvPermissions.Columns.Clear();
            dgvPermissions.Columns.Add(new DataGridViewTextBoxColumn { Name = "Module_ID", Visible = false });
            dgvPermissions.Columns.Add(new DataGridViewTextBoxColumn { Name = "Module_Name", HeaderText = "Module", Width = 140, ReadOnly = true });
            dgvPermissions.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Override", HeaderText = "Override", Width = 70 });
            dgvPermissions.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Can_View", HeaderText = "Xem", Width = 55 });
            dgvPermissions.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Can_Create", HeaderText = "Tạo", Width = 55 });
            dgvPermissions.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Can_Edit", HeaderText = "Sửa", Width = 55 });
            dgvPermissions.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Can_Delete", HeaderText = "Xóa", Width = 55 });
            dgvPermissions.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Can_Export", HeaderText = "Xuất", Width = 55 });
        }

        // =====================================================================
        //  LOAD DATA
        // =====================================================================
        private void LoadData()
        {
            _roles = _svc.GetRoles();
            _modules = _svc.GetModules();

            cboRole.Items.Clear();
            foreach (var r in _roles) cboRole.Items.Add(r.Role_Name);
            if (cboRole.Items.Count > 0) cboRole.SelectedIndex = 0;

            LoadUsers();
        }

        private void LoadUsers()
        {
            _users = _svc.GetAll();
            BindUserGrid(_users);
            lblStatus.Text = $"Tổng: {_users.Count} tài khoản";
        }

        private void BindUserGrid(List<AppUser> list)
        {
            dgvUsers.DataSource = list.ConvertAll(u => new
            {
                ID = u.User_ID,
                Username = u.Username,
                Ho_Ten = u.Full_Name,
                Role = u.Role_Name,
                Phong_Ban = u.Department,
                Email = u.Email,
                Trang_Thai = u.Is_Active ? "Active" : "Disabled",
                Dang_Nhap = u.Last_Login.HasValue ? u.Last_Login.Value.ToString("dd/MM/yyyy HH:mm") : "Chưa đăng nhập"
            });
            if (dgvUsers.Columns.Contains("ID")) dgvUsers.Columns["ID"].Visible = false;
        }

        private void FilterUsers()
        {
            string kw = txtSearch.Text.Trim();
            var filtered = string.IsNullOrEmpty(kw) ? _users
                : _users.FindAll(u =>
                    u.Username.Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                    u.Full_Name.Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                    u.Role_Name.Contains(kw, StringComparison.OrdinalIgnoreCase));
            BindUserGrid(filtered);
        }

        private void LoadPermissions(int userId)
        {
            var perms = _svc.GetEffectivePermissions(userId);
            dgvPermissions.Rows.Clear();

            foreach (var m in _modules)
            {
                var p = perms.Find(x => x.Module_ID == m.Module_ID);
                int idx = dgvPermissions.Rows.Add();
                var row = dgvPermissions.Rows[idx];
                row.Cells["Module_ID"].Value = m.Module_ID;
                row.Cells["Module_Name"].Value = m.Module_Name;
                row.Cells["Override"].Value = p?.Is_Custom_Override ?? false;
                row.Cells["Can_View"].Value = p?.Can_View ?? false;
                row.Cells["Can_Create"].Value = p?.Can_Create ?? false;
                row.Cells["Can_Edit"].Value = p?.Can_Edit ?? false;
                row.Cells["Can_Delete"].Value = p?.Can_Delete ?? false;
                row.Cells["Can_Export"].Value = p?.Can_Export ?? false;
            }
        }

        // =====================================================================
        //  EVENTS
        // =====================================================================
        private void DgvUsers_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvUsers.SelectedRows.Count == 0) return;
            _selectedUserId = Convert.ToInt32(dgvUsers.SelectedRows[0].Cells["ID"].Value);
            var u = _users.Find(x => x.User_ID == _selectedUserId);
            if (u == null) return;

            txtUsername.Text = u.Username;
            txtUsername.ReadOnly = true; // không cho sửa username
            txtFullName.Text = u.Full_Name;
            txtEmail.Text = u.Email;
            txtPhone.Text = u.Phone;
            txtDept.Text = u.Department;
            chkActive.Checked = u.Is_Active;
            chkMustChange.Checked = u.Must_Change_Password;
            txtNewPwd.Text = "";
            txtNewPwd.Visible = false;

            int roleIdx = _roles.FindIndex(r => r.Role_ID == u.Role_ID);
            cboRole.SelectedIndex = roleIdx >= 0 ? roleIdx : 0;

            LoadPermissions(_selectedUserId);
        }

        private void DgvUsers_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (dgvUsers.Columns[e.ColumnIndex].Name == "Trang_Thai")
            {
                e.CellStyle.ForeColor = e.Value?.ToString() == "Active"
                    ? Color.FromArgb(40, 167, 69) : Color.FromArgb(220, 53, 69);
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
        }

        private void BtnNew_Click(object sender, EventArgs e)
        {
            _selectedUserId = 0;
            txtUsername.Text = "";
            txtUsername.ReadOnly = false;
            txtFullName.Text = "";
            txtEmail.Text = "";
            txtPhone.Text = "";
            txtDept.Text = "";
            chkActive.Checked = true;
            chkMustChange.Checked = true;
            cboRole.SelectedIndex = 2; // User
            txtNewPwd.Visible = true;
            txtNewPwd.Text = "";
            txtNewPwd.PlaceholderText = "Mật khẩu ban đầu...";
            dgvPermissions.Rows.Clear();
            txtUsername.Focus();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtUsername.Text) || string.IsNullOrWhiteSpace(txtFullName.Text))
            {
                MessageBox.Show("Vui lòng nhập Username và Họ tên!", "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int roleId = _roles[cboRole.SelectedIndex].Role_ID;

            var user = new AppUser
            {
                User_ID = _selectedUserId,
                Username = txtUsername.Text.Trim(),
                Full_Name = txtFullName.Text.Trim(),
                Email = txtEmail.Text.Trim(),
                Phone = txtPhone.Text.Trim(),
                Department = txtDept.Text.Trim(),
                Role_ID = roleId,
                Is_Active = chkActive.Checked,
                Must_Change_Password = chkMustChange.Checked
            };

            try
            {
                if (_selectedUserId == 0)
                {
                    if (string.IsNullOrWhiteSpace(txtNewPwd.Text))
                    {
                        MessageBox.Show("Vui lòng nhập mật khẩu ban đầu!", "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    _selectedUserId = _svc.InsertUser(user, txtNewPwd.Text, AppSession.CurrentUser?.Username ?? "Admin");
                    MessageBox.Show($"✅ Tạo tài khoản '{user.Username}' thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    _svc.UpdateUser(user, AppSession.CurrentUser?.Username ?? "Admin");
                    MessageBox.Show("✅ Cập nhật thông tin thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                LoadUsers();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnResetPwd_Click(object sender, EventArgs e)
        {
            if (_selectedUserId == 0) { MessageBox.Show("Vui lòng chọn user!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            string newPwd = Microsoft.VisualBasic.Interaction.InputBox(
                "Nhập mật khẩu mới cho user này:", "Reset Password", "Admin@123");
            if (string.IsNullOrWhiteSpace(newPwd)) return;

            _svc.ResetPassword(_selectedUserId, newPwd, AppSession.CurrentUser?.Username ?? "Admin");
            MessageBox.Show($"✅ Đã reset mật khẩu thành công!\nUser sẽ phải đổi mật khẩu khi đăng nhập lần tiếp.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnDeactivate_Click(object sender, EventArgs e)
        {
            if (_selectedUserId == 0) { MessageBox.Show("Vui lòng chọn user!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (_selectedUserId == AppSession.CurrentUser?.User_ID) { MessageBox.Show("Không thể vô hiệu hóa tài khoản đang đăng nhập!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            if (MessageBox.Show("Vô hiệu hóa tài khoản này?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                _svc.DeactivateUser(_selectedUserId, AppSession.CurrentUser?.Username ?? "Admin");
                MessageBox.Show("✅ Đã vô hiệu hóa tài khoản!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadUsers();
            }
        }

        private void BtnSavePerms_Click(object sender, EventArgs e)
        {
            if (_selectedUserId == 0) { MessageBox.Show("Vui lòng chọn user trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            var perms = new List<UserPermission>();
            foreach (DataGridViewRow row in dgvPermissions.Rows)
            {
                bool isOverride = Convert.ToBoolean(row.Cells["Override"].Value ?? false);
                if (!isOverride) continue; // chỉ lưu những dòng được tick Override

                perms.Add(new UserPermission
                {
                    Module_ID = Convert.ToInt32(row.Cells["Module_ID"].Value),
                    Can_View = Convert.ToBoolean(row.Cells["Can_View"].Value ?? false),
                    Can_Create = Convert.ToBoolean(row.Cells["Can_Create"].Value ?? false),
                    Can_Edit = Convert.ToBoolean(row.Cells["Can_Edit"].Value ?? false),
                    Can_Delete = Convert.ToBoolean(row.Cells["Can_Delete"].Value ?? false),
                    Can_Export = Convert.ToBoolean(row.Cells["Can_Export"].Value ?? false)
                });
            }

            _svc.SaveUserPermissions(_selectedUserId, perms);
            MessageBox.Show("✅ Đã lưu phân quyền tùy chỉnh!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            LoadPermissions(_selectedUserId);
        }

        private void BtnResetPerms_Click(object sender, EventArgs e)
        {
            if (_selectedUserId == 0) return;
            if (MessageBox.Show("Reset về quyền mặc định của Role?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                _svc.ResetUserPermissions(_selectedUserId);
                MessageBox.Show("✅ Đã reset về quyền mặc định!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadPermissions(_selectedUserId);
            }
        }

        // =====================================================================
        //  HELPERS
        // =====================================================================
        private DataGridView BuildGrid(Panel parent, int top, int height)
        {
            var dgv = new DataGridView
            {
                Location = new Point(10, top),
                Size = new Size(parent.Width - 20, height),
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.EnableHeadersVisualStyles = false;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            parent.Controls.Add(dgv);
            return dgv;
        }

        private Button Btn(string text, Color color, Point loc, int w, int h)
        {
            var btn = new Button { Text = text, Location = loc, Size = new Size(w, h), BackColor = color, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Cursor = Cursors.Hand };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        private void AddLbl(Panel p, string text, int x, int y) =>
            p.Controls.Add(new Label { Text = text, Location = new Point(x, y + 3), Size = new Size(115, 20), Font = new Font("Segoe UI", 9) });

        private TextBox AddTxt(Panel p, int x, int y, int w)
        {
            var t = new TextBox { Location = new Point(x, y), Size = new Size(w > 0 ? w : 150, 25), Font = new Font("Segoe UI", 9) };
            p.Controls.Add(t);
            return t;
        }

        private void ResizeControls()
        {
            try
            {
                int w = this.ClientSize.Width - 20;
                panelTop.Width = w;
                dgvUsers.Width = panelTop.Width - 20;
                panelPerm.Width = w - panelForm.Width - 10;
                panelPerm.Left = panelForm.Right + 10;
                dgvPermissions.Width = panelPerm.Width - 20;
            }
            catch { }
        }
    }
}