// ============================================================
//  FILE: Forms/frmUserManagement.cs
//  Quản lý User + phân quyền theo module & chức năng button
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

        // ── Danh sách module & permission tương ứng ──────────────────────────
        // Mỗi module định nghĩa: Key, Tên hiển thị, và danh sách cột permission
        // Key phải khớp với Module_Code trong DB
        private static readonly List<ModuleDef> ModuleDefs = new List<ModuleDef>
        {
            new ModuleDef("PROJECT",  "Dự án (Project)",
                new[]{ "Xem","Thêm mới","Lưu","Xóa","Làm mới" }),

            new ModuleDef("MPR",      "Yêu cầu MH (MPR)",
                new[]{ "Xem","Tạo MPR","Lưu Header","Xóa MPR","Thêm dòng","Lưu chi tiết","Xóa dòng","Tạo PO","Check All Items","Xuất Excel" }),

            new ModuleDef("PO",       "Đơn đặt hàng (PO)",
                new[]{ "Xem","Tạo PO","Lưu PO","Xóa PO","Import MPR","Thêm dòng","Lưu chi tiết","Xóa dòng","Payment","Revise History","Tìm theo NCC","Check by size","Xuất Excel" }),

            new ModuleDef("PAYMENT",  "Thanh toán (Payment)",
                new[]{ "Xem","Thêm đợt","Lưu","Xóa","Request to EC","In Request","Ghi nhận TT","Xuất Excel","Xem báo cáo" }),

            new ModuleDef("RIR",      "Nhận hàng (RIR)",
                new[]{ "Xem","Tạo RIR","Lưu Header","Xóa RIR","In RIR","Import Phiếu Nhập","Thêm dòng","Lưu chi tiết","Xóa dòng" }),

            new ModuleDef("WAREHOUSE","Kho (Warehouse)",
                new[]{ "Xem","Lưu hóa đơn","Xuất tồn kho" }),

            new ModuleDef("USER_MGT", "Quản lý User",
                new[]{ "Xem","Tạo user","Lưu user","Vô hiệu hóa","Reset Password","Phân quyền" }),
        };

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
            this.Size = new Size(1400, 820);
            this.MinimumSize = new Size(1100, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
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

            // ===== PANEL PERMISSION: bảng phân quyền chi tiết =====
            panelPerm = new Panel
            {
                Location = new Point(700, 240),
                Size = new Size(670, 530),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            this.Controls.Add(panelPerm);

            panelPerm.Controls.Add(new Label
            {
                Text = "PHÂN QUYỀN THEO MODULE & CHỨC NĂNG",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 8),
                Size = new Size(600, 25)
            });

            panelPerm.Controls.Add(new Label
            {
                Text = "⚠ Tick ✔ vào ô tương ứng để cấp quyền. Bỏ tick = từ chối quyền đó.",
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.Gray,
                Location = new Point(10, 35),
                Size = new Size(640, 18)
            });

            // Nút Cấp tất cả / Thu hồi tất cả
            var btnGrantAll = Btn("✅ Cấp tất cả", Color.FromArgb(40, 167, 69), new Point(10, 55), 125, 26);
            var btnRevokeAll = Btn("❌ Thu hồi tất cả", Color.FromArgb(220, 53, 69), new Point(143, 55), 135, 26);
            var btnGrantModule = Btn("☑ Cấp theo Role mẫu", Color.FromArgb(102, 51, 153), new Point(286, 55), 155, 26);
            btnGrantAll.Click += (s, e) => SetAllPermissions(true);
            btnRevokeAll.Click += (s, e) => SetAllPermissions(false);
            btnGrantModule.Click += BtnApplyRoleTemplate_Click;
            panelPerm.Controls.AddRange(new Control[] { btnGrantAll, btnRevokeAll, btnGrantModule });

            dgvPermissions = new DataGridView
            {
                Location = new Point(10, 88),
                Size = new Size(panelPerm.Width - 20, panelPerm.Height - 130),
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                MultiSelect = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                ScrollBars = ScrollBars.Both
            };
            dgvPermissions.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvPermissions.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPermissions.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvPermissions.EnableHeadersVisualStyles = false;
            dgvPermissions.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvPermissions.RowTemplate.Height = 26;
            panelPerm.Controls.Add(dgvPermissions);

            BuildPermissionColumns();

            int btnY = panelPerm.Height - 38;
            btnSavePerms = Btn("💾 Lưu phân quyền", Color.FromArgb(0, 120, 212), new Point(10, btnY), 155, 30);
            btnResetPerms = Btn("🔄 Reset về Role", Color.FromArgb(108, 117, 125), new Point(175, btnY), 145, 30);
            btnSavePerms.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnResetPerms.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnSavePerms.Click += BtnSavePerms_Click;
            btnResetPerms.Click += BtnResetPerms_Click;
            panelPerm.Controls.AddRange(new Control[] { btnSavePerms, btnResetPerms });

            // ===== PANEL FORM mở rộng xuống theo chiều cao còn lại =====
            panelForm.Size = new Size(680, 530);
            panelForm.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom;

            // Bảng lịch sử phân quyền (phía dưới panelForm)
            panelForm.Controls.Add(new Label
            {
                Text = "LỊCH SỬ THAY ĐỔI QUYỀN",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(102, 51, 153),
                Location = new Point(10, 245),
                Size = new Size(300, 22)
            });

            var dgvAudit = new DataGridView
            {
                Name = "dgvAudit",
                Location = new Point(10, 270),
                Size = new Size(655, 245),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 8),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            dgvAudit.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(102, 51, 153);
            dgvAudit.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvAudit.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 8, FontStyle.Bold);
            dgvAudit.EnableHeadersVisualStyles = false;
            dgvAudit.Columns.Add(new DataGridViewTextBoxColumn { Name = "Thoi_Gian", HeaderText = "Thời gian", Width = 130 });
            dgvAudit.Columns.Add(new DataGridViewTextBoxColumn { Name = "Nguoi_Thay_Doi", HeaderText = "Người thay đổi", Width = 130 });
            dgvAudit.Columns.Add(new DataGridViewTextBoxColumn { Name = "Hanh_Dong", HeaderText = "Hành động", Width = 395 });
            panelForm.Controls.Add(dgvAudit);
        }

        // ─────────────────────────────────────────────────────────────────────
        //  XÂY DỰNG CỘT PERMISSION THEO TỪNG MODULE & CHỨC NĂNG
        // ─────────────────────────────────────────────────────────────────────
        private void BuildPermissionColumns()
        {
            dgvPermissions.Columns.Clear();

            // Cột cố định
            dgvPermissions.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Module_Code",
                Visible = false
            });
            dgvPermissions.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Module_Name",
                HeaderText = "Module / Chức năng",
                Width = 180,
                ReadOnly = true
            });

            // Cột permission động theo từng module
            // Dùng cột chung: mỗi cột = 1 permission key "MODULE:Action"
            // Để bảng không quá rộng, ta dùng cột CheckBox với HeaderText ngắn
            // và tooltip đầy đủ

            // Tập hợp TẤT CẢ action từ mọi module (union)
            var allActions = new List<string>();
            foreach (var mod in ModuleDefs)
                foreach (var act in mod.Actions)
                    if (!allActions.Contains(act)) allActions.Add(act);

            foreach (var act in allActions)
            {
                var col = new DataGridViewCheckBoxColumn
                {
                    Name = "ACT_" + act.Replace(" ", "_"),
                    HeaderText = act,
                    Width = 75,
                    ToolTipText = act
                };
                dgvPermissions.Columns.Add(col);
            }
        }

        // ─────────────────────────────────────────────────────────────────────
        //  NẠP DỮ LIỆU PERMISSION VÀO LƯỚI
        // ─────────────────────────────────────────────────────────────────────
        private void LoadPermissions(int userId)
        {
            // Lấy permission hiệu lực của user từ DB
            var perms = _svc.GetDetailedPermissions(userId); // Dictionary<"MODULE:Action", bool>

            dgvPermissions.Rows.Clear();

            foreach (var mod in ModuleDefs)
            {
                int idx = dgvPermissions.Rows.Add();
                var row = dgvPermissions.Rows[idx];

                row.Cells["Module_Code"].Value = mod.Code;
                row.Cells["Module_Name"].Value = mod.DisplayName;

                // Tô màu header row
                row.DefaultCellStyle.BackColor = Color.FromArgb(220, 235, 252);
                row.DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);

                // Gán giá trị checkbox cho từng action
                foreach (DataGridViewColumn col in dgvPermissions.Columns)
                {
                    if (!col.Name.StartsWith("ACT_")) continue;
                    string actionName = col.HeaderText;

                    if (mod.Actions.Contains(actionName))
                    {
                        // Action này thuộc module → cho phép tương tác
                        string key = mod.Code + ":" + actionName;
                        bool granted = perms.ContainsKey(key) && perms[key];
                        row.Cells[col.Name].Value = granted;
                        row.Cells[col.Name].ReadOnly = false;
                    }
                    else
                    {
                        // Action không thuộc module → ẩn / không áp dụng
                        row.Cells[col.Name].Value = false;
                        row.Cells[col.Name].ReadOnly = true;
                        row.Cells[col.Name].Style.BackColor = Color.FromArgb(220, 220, 220);
                        row.Cells[col.Name].Style.ForeColor = Color.FromArgb(180, 180, 180);
                    }
                }
            }

            // Load lịch sử audit
            LoadAuditLog(userId);
        }

        private void LoadAuditLog(int userId)
        {
            var dgvAudit = panelForm.Controls["dgvAudit"] as DataGridView;
            if (dgvAudit == null) return;
            dgvAudit.Rows.Clear();

            var logs = _svc.GetPermissionAuditLog(userId);
            foreach (var log in logs)
            {
                int i = dgvAudit.Rows.Add();
                dgvAudit.Rows[i].Cells["Thoi_Gian"].Value = log.Changed_At.ToString("dd/MM/yyyy HH:mm");
                dgvAudit.Rows[i].Cells["Nguoi_Thay_Doi"].Value = log.Changed_By;
                dgvAudit.Rows[i].Cells["Hanh_Dong"].Value = log.Action_Detail;
            }
        }

        // ─────────────────────────────────────────────────────────────────────
        //  CẤP / THU HỒI TẤT CẢ QUYỀN
        // ─────────────────────────────────────────────────────────────────────
        private void SetAllPermissions(bool grant)
        {
            foreach (DataGridViewRow row in dgvPermissions.Rows)
            {
                foreach (DataGridViewColumn col in dgvPermissions.Columns)
                {
                    if (!col.Name.StartsWith("ACT_")) continue;
                    if (!row.Cells[col.Name].ReadOnly)
                        row.Cells[col.Name].Value = grant;
                }
            }
        }

        // ─────────────────────────────────────────────────────────────────────
        //  ÁP DỤNG TEMPLATE QUYỀN THEO ROLE
        // ─────────────────────────────────────────────────────────────────────
        private void BtnApplyRoleTemplate_Click(object sender, EventArgs e)
        {
            if (_selectedUserId == 0) { MessageBox.Show("Vui lòng chọn user trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            var u = _users.Find(x => x.User_ID == _selectedUserId);
            if (u == null) return;

            // Lấy template theo Role hiện tại
            var template = _svc.GetRolePermissionTemplate(u.Role_ID); // Dictionary<"MODULE:Action", bool>

            foreach (DataGridViewRow row in dgvPermissions.Rows)
            {
                string modCode = row.Cells["Module_Code"].Value?.ToString() ?? "";
                foreach (DataGridViewColumn col in dgvPermissions.Columns)
                {
                    if (!col.Name.StartsWith("ACT_")) continue;
                    if (row.Cells[col.Name].ReadOnly) continue;
                    string key = modCode + ":" + col.HeaderText;
                    row.Cells[col.Name].Value = template.ContainsKey(key) && template[key];
                }
            }

            lblStatus.Text = $"Đã áp dụng template quyền theo Role: {u.Role_Name}";
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
            txtUsername.ReadOnly = true;
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

            // Thu thập tất cả permission từ lưới
            var perms = new Dictionary<string, bool>(); // key = "MODULE:Action"

            foreach (DataGridViewRow row in dgvPermissions.Rows)
            {
                string modCode = row.Cells["Module_Code"].Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(modCode)) continue;

                foreach (DataGridViewColumn col in dgvPermissions.Columns)
                {
                    if (!col.Name.StartsWith("ACT_")) continue;
                    if (row.Cells[col.Name].ReadOnly) continue; // ô xám = không áp dụng

                    string key = modCode + ":" + col.HeaderText;
                    bool granted = Convert.ToBoolean(row.Cells[col.Name].Value ?? false);
                    perms[key] = granted;
                }
            }

            _svc.SaveDetailedPermissions(_selectedUserId, perms, AppSession.CurrentUser?.Username ?? "Admin");
            MessageBox.Show("✅ Đã lưu phân quyền chi tiết!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
            dgv.DefaultCellStyle.SelectionForeColor = Color.Black;
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
                panelPerm.Height = this.ClientSize.Height - panelPerm.Top - 10;
                panelForm.Height = this.ClientSize.Height - panelForm.Top - 10;
                dgvPermissions.Width = panelPerm.Width - 20;
                dgvPermissions.Height = panelPerm.Height - 130;
                btnSavePerms.Top = panelPerm.Height - 38;
                btnResetPerms.Top = panelPerm.Height - 38;
            }
            catch { }
        }
    }

    // =========================================================================
    //  HELPER CLASS: Định nghĩa module & danh sách action
    // =========================================================================
    public class ModuleDef
    {
        public string Code { get; }
        public string DisplayName { get; }
        public List<string> Actions { get; }

        public ModuleDef(string code, string displayName, string[] actions)
        {
            Code = code;
            DisplayName = displayName;
            Actions = new List<string>(actions);
        }
    }
}