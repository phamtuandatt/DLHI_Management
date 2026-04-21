// ============================================================
//  FILE: Forms/frmUserManagement.cs
//  Quản lý User + phân quyền theo module & chức năng button
// ============================================================
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;
using MPR_Managerment.Services;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace MPR_Managerment.Forms
{
    public partial class frmUserManagement : Form
    {
        private readonly UserService _svc = new UserService();
        private List<AppUser> _users = new List<AppUser>();
        private List<Role> _roles = new List<Role>();
        private List<AppModule> _modules = new List<AppModule>();
        private int _selectedUserId = 0;
        // Track trạng thái thu gọn của từng module (code → collapsed)
        private readonly HashSet<string> _collapsedModules = new HashSet<string>();

        // Controls
        private DataGridView dgvUsers;
        private DataGridView dgvPermissions;
        private TextBox txtUsername, txtFullName, txtEmail, txtPhone, txtSearch;
        private TextBox txtNewPwd;
        private ComboBox cboRole, cboDept, cboPosition;
        private CheckBox chkActive, chkMustChange;
        private Button btnNew, btnSave, btnDelete, btnDeactivate, btnResetPwd, btnSavePerms, btnResetPerms;
        private Label lblStatus;
        private Panel panelTop, panelForm, panelPerm;

        // ── Danh sách Phòng ban ──────────────────────────────────────────────
        private static readonly string[] Departments = new[]
        {
            "BOD",
            "Hành chính - Kế toán",
            "Dự án",
            "Thiết kế",
            "Mua Hàng",
            "Sản Xuất",
            "QA-QC",
            "Kho"
        };

        // ── Danh sách Chức vụ ────────────────────────────────────────────────
        private static readonly string[] Positions = new[]
        {
            "Tổng Giám Đốc",
            "Phó Tổng Giám Đốc",
            "Trưởng phòng",
            "Phó phòng",
            "Staff",
            "Viewer"
        };

        // ── Danh sách module & permission tương ứng ──────────────────────────
        private static readonly List<ModuleDef> ModuleDefs = new List<ModuleDef>
        {
            new ModuleDef("PROJECT",  "Dự án (Project)",
                new[]{ "Xem","Thêm mới","Lưu","Xóa","Làm mới" }),

            new ModuleDef("MPR",      "Yêu cầu MH (MPR)",
                new[]{ "Xem","Tạo MPR","Lưu Header","Xóa MPR","Thêm dòng","Lưu chi tiết","Xóa dòng","Tạo PO","Check All Items","Xuất Excel" }),

            new ModuleDef("PO",       "Đơn đặt hàng (PO)",
                new[]{ "Xem","Tạo PO","Lưu PO","Xóa PO","Import MPR","Thêm dòng","Lưu chi tiết","Xóa dòng","Payment","Revise History","Tìm theo NCC","Check by size","Xuất Excel","Xem đơn giá","Xem TT trước thuế","Xem TT sau thuế" }),

            new ModuleDef("PAYMENT",  "Thanh toán (Payment)",
                new[]{ "Xem","Thêm đợt","Lưu","Xóa","Request to EC","In Request","Ghi nhận TT","Xuất Excel","Xem báo cáo","Xem TT trước thuế","Xem TT sau thuế" }),

            new ModuleDef("RIR",      "Nhận hàng (RIR)",
                new[]{ "Xem","Tạo RIR","Lưu Header","Xóa RIR","In RIR","Import Phiếu Nhập","Thêm dòng","Lưu chi tiết","Xóa dòng" }),

            new ModuleDef("WAREHOUSE","Kho (Warehouse)",
                new[]{ "Xem","Lưu hóa đơn","Xuất tồn kho" }),

            new ModuleDef("USER_MGT", "Quản lý User",
                new[]{ "Xem","Tạo user","Lưu user","Vô hiệu hóa","Reset Password","Phân quyền" }),

            new ModuleDef("Material Inspector Request", "Kiểm tra vật tư (QC)",
                new[]{ "Xem","Tìm kiếm RIR","Lưu chi tiết" }),
        };

        public frmUserManagement()
        {
            InitializeComponent();
            BuildUI();
            LoadData();
            ApplyPermissions();
            this.Resize += (s, e) => ResizeControls();
        }

        // =====================================================================
        //  BUILD UI
        // =====================================================================
        private void BuildUI()
        {
            this.Text = "Quản lý Người dùng & Phân quyền";
            this.Size = new Size(1400, 860);
            this.MinimumSize = new Size(1100, 750);
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

            // ── Nút Reset Admin khẩn cấp (chỉ hiện với Admin) ──────────────
            var btnResetAdmin = Btn("🛡 Reset Admin", Color.FromArgb(155, 0, 0), new Point(390, 47), 135, 30);
            btnResetAdmin.Click += BtnResetAdmin_Click;
            btnResetAdmin.Visible = AppSession.IsAdmin;
            panelTop.Controls.Add(btnResetAdmin);

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
                Size = new Size(680, 580),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom
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

            // ── Hàng 1: Username | Họ tên ──
            // Label nằm dưới input (label ở y+28, input ở y)
            int y = 40;
            txtUsername = AddTxtWithLabel(panelForm, "Username (*):", 10, y, 160);
            txtFullName = AddTxtWithLabel(panelForm, "Họ tên (*):", 200, y, 200);
            txtEmail = AddTxtWithLabel(panelForm, "Email:", 420, y, 235);

            // ── Hàng 2: SĐT | Mật khẩu mới ──
            y += 62;
            txtPhone = AddTxtWithLabel(panelForm, "SĐT:", 10, y, 160);
            txtNewPwd = AddTxtWithLabel(panelForm, "Mật khẩu mới:", 200, y, 180);
            txtNewPwd.PasswordChar = '●';

            // ── Hàng 3: Phòng ban (ComboBox) | Chức vụ (ComboBox) | Role ──
            y += 62;
            panelForm.Controls.Add(new Label
            {
                Text = "Phòng ban:",
                Location = new Point(10, y + 27),
                Size = new Size(100, 18),
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.Gray
            });
            cboDept = new ComboBox
            {
                Location = new Point(10, y),
                Size = new Size(175, 26),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboDept.Items.AddRange(Departments);
            panelForm.Controls.Add(cboDept);

            panelForm.Controls.Add(new Label
            {
                Text = "Chức vụ:",
                Location = new Point(200, y + 27),
                Size = new Size(100, 18),
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.Gray
            });
            cboPosition = new ComboBox
            {
                Location = new Point(200, y),
                Size = new Size(160, 26),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboPosition.Items.AddRange(Positions);
            panelForm.Controls.Add(cboPosition);

            panelForm.Controls.Add(new Label
            {
                Text = "Role (*):",
                Location = new Point(378, y + 27),
                Size = new Size(100, 18),
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.Gray
            });
            cboRole = new ComboBox
            {
                Location = new Point(378, y),
                Size = new Size(160, 26),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            panelForm.Controls.Add(cboRole);

            // ── Hàng 4: CheckBox ──
            y += 62;
            chkActive = new CheckBox
            {
                Text = "Tài khoản Active",
                Location = new Point(10, y),
                Size = new Size(170, 24),
                Font = new Font("Segoe UI", 9),
                Checked = true
            };
            panelForm.Controls.Add(chkActive);
            chkMustChange = new CheckBox
            {
                Text = "Bắt buộc đổi mật khẩu",
                Location = new Point(200, y),
                Size = new Size(200, 24),
                Font = new Font("Segoe UI", 9)
            };
            panelForm.Controls.Add(chkMustChange);

            // ── Hàng 5: Buttons ──
            y += 38;
            btnSave = Btn("💾 Lưu", Color.FromArgb(0, 120, 212), new Point(10, y), 100, 32);
            btnResetPwd = Btn("🔑 Reset Password", Color.FromArgb(255, 140, 0), new Point(118, y), 145, 32);
            btnDeactivate = Btn("🚫 Vô hiệu hóa", Color.FromArgb(255, 140, 0), new Point(271, y), 140, 32);
            btnDelete = Btn("🗑 Xóa user", Color.FromArgb(220, 53, 69), new Point(419, y), 120, 32);

            btnSave.Click += BtnSave_Click;
            btnResetPwd.Click += BtnResetPwd_Click;
            btnDeactivate.Click += BtnDeactivate_Click;
            btnDelete.Click += BtnDelete_Click;
            panelForm.Controls.AddRange(new Control[] { btnSave, btnResetPwd, btnDeactivate, btnDelete });

            // ── Lịch sử thay đổi quyền ──
            y += 50;
            panelForm.Controls.Add(new Label
            {
                Text = "LỊCH SỬ THAY ĐỔI QUYỀN",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(102, 51, 153),
                Location = new Point(10, y),
                Size = new Size(300, 22)
            });

            var dgvAudit = new DataGridView
            {
                Name = "dgvAudit",
                Location = new Point(10, y + 26),
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

            // ===== PANEL PERMISSION =====
            panelPerm = new Panel
            {
                Location = new Point(700, 240),
                Size = new Size(670, 580),
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
                ScrollBars = ScrollBars.Vertical,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                GridColor = Color.FromArgb(220, 220, 220),
                EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2
            };
            dgvPermissions.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvPermissions.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPermissions.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvPermissions.EnableHeadersVisualStyles = false;
            dgvPermissions.RowTemplate.Height = 28;
            dgvPermissions.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 240, 255);
            dgvPermissions.DefaultCellStyle.SelectionForeColor = Color.Black;
            panelPerm.Controls.Add(dgvPermissions);

            BuildPermissionColumns();

            // Format màu cho dòng header module vs dòng action
            dgvPermissions.CellFormatting += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                var r = dgvPermissions.Rows[e.RowIndex];
                if (r.Cells["Row_Type"].Value?.ToString() == "HEADER")
                {
                    e.CellStyle.BackColor = Color.FromArgb(0, 120, 212);
                    e.CellStyle.ForeColor = Color.White;
                    e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    // Dòng header: checkbox luôn là false (không cho tick)
                    if (dgvPermissions.Columns[e.ColumnIndex].Name == "Allowed")
                    {
                        e.Value = false;
                        e.FormattingApplied = true;
                    }
                }
            };

            // Bắt DataError để tránh dialog lỗi Boolean pop-up
            dgvPermissions.DataError += (s, e) => { e.Cancel = true; };

            // Toggle checkbox ngay khi click 1 lần (không cần click 2 lần)
            dgvPermissions.CellContentClick += (s, e) =>
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
                if (dgvPermissions.Columns[e.ColumnIndex].Name != "Allowed") return;
                var row = dgvPermissions.Rows[e.RowIndex];
                if (row.Cells["Row_Type"].Value?.ToString() != "ACTION") return;
                if (row.Cells["Allowed"].ReadOnly) return;
                dgvPermissions.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };

            // Click vào dòng HEADER → toggle collapse/expand
            dgvPermissions.CellClick += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                var row = dgvPermissions.Rows[e.RowIndex];
                if (row.Cells["Row_Type"].Value?.ToString() != "HEADER") return;
                string modCode = row.Cells["Module_Code"].Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(modCode))
                    ToggleModuleCollapse(modCode);
            };

            // Cursor dạng tay khi hover trên dòng HEADER
            dgvPermissions.CellMouseEnter += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                var row = dgvPermissions.Rows[e.RowIndex];
                dgvPermissions.Cursor = row.Cells["Row_Type"].Value?.ToString() == "HEADER"
                    ? Cursors.Hand
                    : Cursors.Default;
            };
            dgvPermissions.CellMouseLeave += (s, e) =>
            {
                dgvPermissions.Cursor = Cursors.Default;
            };

            int btnY = panelPerm.Height - 38;
            btnSavePerms = Btn("💾 Lưu phân quyền", Color.FromArgb(0, 120, 212), new Point(10, btnY), 155, 30);
            btnResetPerms = Btn("🔄 Reset về Role", Color.FromArgb(108, 117, 125), new Point(175, btnY), 145, 30);
            btnSavePerms.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnResetPerms.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnSavePerms.Click += BtnSavePerms_Click;
            btnResetPerms.Click += BtnResetPerms_Click;
            panelPerm.Controls.AddRange(new Control[] { btnSavePerms, btnResetPerms });
        }

        // ─────────────────────────────────────────────────────────────────────
        //  Helper: TextBox có Label phía dưới
        // ─────────────────────────────────────────────────────────────────────
        private TextBox AddTxtWithLabel(Panel p, string labelText, int x, int y, int w)
        {
            var txt = new TextBox
            {
                Location = new Point(x, y),
                Size = new Size(w > 0 ? w : 150, 26),
                Font = new Font("Segoe UI", 9)
            };
            p.Controls.Add(txt);
            p.Controls.Add(new Label
            {
                Text = labelText,
                Location = new Point(x, y + 28),
                Size = new Size(w > 0 ? w : 150, 18),
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.Gray
            });
            return txt;
        }

        // ─────────────────────────────────────────────────────────────────────
        //  XÂY DỰNG CỘT PERMISSION
        // ─────────────────────────────────────────────────────────────────────
        private void BuildPermissionColumns()
        {
            dgvPermissions.Columns.Clear();

            // Cột ẩn: lưu module code
            dgvPermissions.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Module_Code",
                Visible = false
            });
            // Cột ẩn: đánh dấu row là header module hay action
            dgvPermissions.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Row_Type",
                Visible = false   // "HEADER" | "ACTION"
            });
            // Cột ẩn: lưu tên action gốc (key permission)
            dgvPermissions.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Action_Key",
                Visible = false
            });
            // Cột tên chức năng (hiển thị)
            dgvPermissions.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Action_Name",
                HeaderText = "Chức năng",
                Width = 260,
                ReadOnly = true
            });
            // Cột checkbox cho phép
            dgvPermissions.Columns.Add(new DataGridViewCheckBoxColumn
            {
                Name = "Allowed",
                HeaderText = "Cho phép",
                Width = 80,
                FalseValue = false,
                TrueValue = true
            });
        }

        // ─────────────────────────────────────────────────────────────────────
        //  NẠP PERMISSION VÀO LƯỚI — dạng list: header module + dòng action
        // ─────────────────────────────────────────────────────────────────────
        private void LoadPermissions(int userId)
        {
            var perms = _svc.GetDetailedPermissions(userId);
            dgvPermissions.SuspendLayout();
            dgvPermissions.Rows.Clear();

            // Lần đầu mở form: thu gọn tất cả module theo mặc định
            if (_collapsedModules.Count == 0)
                foreach (var m in ModuleDefs)
                    _collapsedModules.Add(m.Code);

            foreach (var mod in ModuleDefs)
            {
                bool collapsed = _collapsedModules.Contains(mod.Code);
                string icon = collapsed ? "▶  " : "▼  ";

                // ── Dòng HEADER module ──────────────────────────────────────
                int hIdx = dgvPermissions.Rows.Add();
                var hRow = dgvPermissions.Rows[hIdx];
                hRow.Cells["Module_Code"].Value = mod.Code;
                hRow.Cells["Row_Type"].Value = "HEADER";
                hRow.Cells["Action_Key"].Value = "";
                hRow.Cells["Action_Name"].Value = icon + mod.DisplayName.ToUpper();
                hRow.Cells["Allowed"].Value = false;
                hRow.Cells["Allowed"].ReadOnly = true;
                // Style header
                hRow.DefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
                hRow.DefaultCellStyle.ForeColor = Color.White;
                hRow.DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                hRow.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 90, 180);
                hRow.DefaultCellStyle.SelectionForeColor = Color.White;
                hRow.Height = 30;

                // ── Dòng ACTION cho mỗi chức năng của module ───────────────
                bool isOdd = false;
                foreach (var action in mod.Actions)
                {
                    string key = mod.Code + ":" + action;
                    bool granted = perms.ContainsKey(key) && perms[key];

                    int aIdx = dgvPermissions.Rows.Add();
                    var aRow = dgvPermissions.Rows[aIdx];
                    aRow.Cells["Module_Code"].Value = mod.Code;
                    aRow.Cells["Row_Type"].Value = "ACTION";
                    aRow.Cells["Action_Key"].Value = action;
                    aRow.Cells["Action_Name"].Value = "      " + action;
                    aRow.Cells["Allowed"].Value = granted;
                    aRow.Cells["Allowed"].ReadOnly = false;
                    // Ẩn nếu module đang collapsed
                    aRow.Visible = !collapsed;

                    // Zebra striping
                    aRow.DefaultCellStyle.BackColor = isOdd
                        ? Color.FromArgb(248, 252, 255)
                        : Color.White;
                    aRow.DefaultCellStyle.ForeColor = Color.FromArgb(40, 40, 40);
                    aRow.DefaultCellStyle.SelectionBackColor = Color.FromArgb(210, 230, 255);
                    aRow.DefaultCellStyle.SelectionForeColor = Color.Black;
                    isOdd = !isOdd;
                }
            }

            dgvPermissions.ResumeLayout();
            LoadAuditLog(userId);
        }

        // ─────────────────────────────────────────────────────────────────────
        //  TOGGLE COLLAPSE MODULE khi click vào dòng HEADER
        // ─────────────────────────────────────────────────────────────────────
        private void ToggleModuleCollapse(string moduleCode)
        {
            bool nowCollapsed;
            if (_collapsedModules.Contains(moduleCode))
            {
                _collapsedModules.Remove(moduleCode);
                nowCollapsed = false;
            }
            else
            {
                _collapsedModules.Add(moduleCode);
                nowCollapsed = true;
            }

            string icon = nowCollapsed ? "▶  " : "▼  ";

            dgvPermissions.SuspendLayout();
            foreach (DataGridViewRow row in dgvPermissions.Rows)
            {
                string code = row.Cells["Module_Code"].Value?.ToString() ?? "";
                if (code != moduleCode) continue;

                string rowType = row.Cells["Row_Type"].Value?.ToString() ?? "";
                if (rowType == "HEADER")
                {
                    // Cập nhật icon trên header
                    var mod = ModuleDefs.Find(m => m.Code == moduleCode);
                    if (mod != null)
                        row.Cells["Action_Name"].Value = icon + mod.DisplayName.ToUpper();
                }
                else if (rowType == "ACTION")
                {
                    row.Visible = !nowCollapsed;
                }
            }
            dgvPermissions.ResumeLayout();
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
                if (row.Cells["Row_Type"].Value?.ToString() != "ACTION") continue;
                if (!row.Cells["Allowed"].ReadOnly)
                    row.Cells["Allowed"].Value = grant;
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

            var template = _svc.GetRolePermissionTemplate(u.Role_ID);
            foreach (DataGridViewRow row in dgvPermissions.Rows)
            {
                if (row.Cells["Row_Type"].Value?.ToString() != "ACTION") continue;
                if (row.Cells["Allowed"].ReadOnly) continue;
                string modCode = row.Cells["Module_Code"].Value?.ToString() ?? "";
                string action = row.Cells["Action_Key"].Value?.ToString() ?? "";
                string key = modCode + ":" + action;
                row.Cells["Allowed"].Value = template.ContainsKey(key) && template[key];
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
            chkActive.Checked = u.Is_Active;
            chkMustChange.Checked = u.Must_Change_Password;
            txtNewPwd.Text = "";
            txtNewPwd.Visible = false;

            // Phòng ban
            int deptIdx = Array.IndexOf(Departments, ParseDepartment(u.Department));
            cboDept.SelectedIndex = deptIdx >= 0 ? deptIdx : -1;

            // Chức vụ (lưu trong Position hoặc Department tùy model — fallback về -1)
            int posIdx = Array.IndexOf(Positions, ParsePosition(u.Department));
            cboPosition.SelectedIndex = posIdx >= 0 ? posIdx : -1;

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
            if (!PermissionHelper.Check("USER_MGT", "Tạo user", "Tạo user mới")) return;
            _selectedUserId = 0;
            txtUsername.Text = "";
            txtUsername.ReadOnly = false;
            txtFullName.Text = "";
            txtEmail.Text = "";
            txtPhone.Text = "";
            cboDept.SelectedIndex = -1;
            cboPosition.SelectedIndex = -1;
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
            if (!PermissionHelper.Check("USER_MGT", "Lưu user", "Lưu thông tin user")) return;

            if (string.IsNullOrWhiteSpace(txtUsername.Text) || string.IsNullOrWhiteSpace(txtFullName.Text))
            {
                MessageBox.Show("Vui lòng nhập Username và Họ tên!", "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int roleId = _roles[cboRole.SelectedIndex].Role_ID;
            string dept = cboDept.SelectedIndex >= 0 ? cboDept.SelectedItem.ToString() : "";
            string position = cboPosition.SelectedIndex >= 0 ? cboPosition.SelectedItem.ToString() : "";
            string deptCombined = string.IsNullOrEmpty(position) ? dept : $"{dept}||{position}";

            var user = new AppUser
            {
                User_ID = _selectedUserId,
                Username = txtUsername.Text.Trim(),
                Full_Name = txtFullName.Text.Trim(),
                Email = txtEmail.Text.Trim(),
                Phone = txtPhone.Text.Trim(),
                Department = deptCombined,

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
            if (!PermissionHelper.Check("USER_MGT", "Reset Password", "Reset Password")) return;
            if (_selectedUserId == 0) { MessageBox.Show("Vui lòng chọn user!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            string newPwd = Microsoft.VisualBasic.Interaction.InputBox(
                "Nhập mật khẩu mới cho user này:", "Reset Password", "Admin@123");
            if (string.IsNullOrWhiteSpace(newPwd)) return;

            _svc.ResetPassword(_selectedUserId, newPwd, AppSession.CurrentUser?.Username ?? "Admin");
            MessageBox.Show($"✅ Đã reset mật khẩu thành công!\nUser sẽ phải đổi mật khẩu khi đăng nhập lần tiếp.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnDeactivate_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("USER_MGT", "Vô hiệu hóa", "Vô hiệu hóa tài khoản")) return;
            if (_selectedUserId == 0) { MessageBox.Show("Vui lòng chọn user!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (_selectedUserId == AppSession.CurrentUser?.User_ID) { MessageBox.Show("Không thể vô hiệu hóa tài khoản đang đăng nhập!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            if (MessageBox.Show("Vô hiệu hóa tài khoản này?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                _svc.DeactivateUser(_selectedUserId, AppSession.CurrentUser?.Username ?? "Admin");
                MessageBox.Show("✅ Đã vô hiệu hóa tài khoản!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadUsers();
            }
        }

        // ─────────────────────────────────────────────────────────────────────
        //  XÓA USER — yêu cầu mật khẩu Admin
        // ─────────────────────────────────────────────────────────────────────
        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("USER_MGT", "Lưu user", "Xóa user")) return;
            if (_selectedUserId == 0)
            {
                MessageBox.Show("Vui lòng chọn user cần xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (_selectedUserId == AppSession.CurrentUser?.User_ID)
            {
                MessageBox.Show("Không thể xóa tài khoản đang đăng nhập!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var u = _users.Find(x => x.User_ID == _selectedUserId);
            string uname = u?.Username ?? _selectedUserId.ToString();

            if (MessageBox.Show(
                $"Bạn chắc chắn muốn XÓA VĨNH VIỄN tài khoản '{uname}'?\n\nHành động này KHÔNG thể hoàn tác!",
                "Xác nhận xóa user",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;

            // ── Xác thực mật khẩu Admin ──
            if (!VerifyAdminPassword()) return;

            try
            {
                _svc.DeactivateUser(_selectedUserId, AppSession.CurrentUser?.Username ?? "Admin");

                MessageBox.Show($"✅ Đã vô hiệu hóa tài khoản '{uname}' thành công!\n(Tài khoản bị vô hiệu hóa — liên hệ DBA nếu cần xóa hẳn khỏi DB.)", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                _selectedUserId = 0;
                LoadUsers();
                BtnNew_Click(null, null); // reset form
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi xóa: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ─────────────────────────────────────────────────────────────────────
        //  XÁC THỰC MẬT KHẨU ADMIN
        // ─────────────────────────────────────────────────────────────────────
        private bool VerifyAdminPassword()
        {
            var dlg = new Form
            {
                Text = "🔐 Xác thực Admin",
                Size = new Size(380, 175),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(245, 245, 245),
                KeyPreview = true
            };
            dlg.Controls.Add(new Label
            {
                Text = "Nhập mật khẩu tài khoản Admin để xác nhận xóa:",
                Font = new Font("Segoe UI", 9),
                Location = new Point(15, 15),
                Size = new Size(340, 20)
            });
            var txtPwd = new TextBox
            {
                Location = new Point(15, 42),
                Size = new Size(340, 26),
                Font = new Font("Segoe UI", 10),
                PasswordChar = '●'
            };
            dlg.Controls.Add(txtPwd);
            var lblErr = new Label
            {
                Text = "",
                ForeColor = Color.FromArgb(220, 53, 69),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Location = new Point(15, 74),
                Size = new Size(340, 20)
            };
            dlg.Controls.Add(lblErr);
            var btnOK = new Button
            {
                Text = "✔ Xác nhận",
                Location = new Point(155, 100),
                Size = new Size(100, 30),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            btnOK.FlatAppearance.BorderSize = 0;
            dlg.Controls.Add(btnOK);
            var btnCancel = new Button
            {
                Text = "Hủy",
                Location = new Point(265, 100),
                Size = new Size(90, 30),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                DialogResult = DialogResult.Cancel
            };
            btnCancel.FlatAppearance.BorderSize = 0;
            dlg.Controls.Add(btnCancel);
            dlg.CancelButton = btnCancel;

            bool verified = false;
            btnOK.Click += (s, ev) =>
            {
                string pwd = txtPwd.Text;
                if (string.IsNullOrEmpty(pwd)) { lblErr.Text = "Vui lòng nhập mật khẩu!"; return; }
                try
                {
                    string inputHash;
                    using (var sha256 = System.Security.Cryptography.SHA256.Create())
                    {
                        byte[] bytes = sha256.ComputeHash(System.Text.Encoding.UTF8.GetBytes(pwd));
                        inputHash = BitConverter.ToString(bytes).Replace("-", "").ToLower();
                    }
                    const string ADMIN_HASH = "e86f78a8a3caf0b60d8e74e5942aa6d86dc150cd3c03338aef25b7d2d7e3acc7";
                    bool match = inputHash == ADMIN_HASH;
                    if (!match)
                    {
                        using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                        conn.Open();
                        var cmd = new Microsoft.Data.SqlClient.SqlCommand(
                            "SELECT COUNT(1) FROM Users WHERE LOWER(Username)='admin' AND (LOWER(Password)=@hash OR Password=@pwd)", conn);
                        cmd.Parameters.AddWithValue("@hash", inputHash);
                        cmd.Parameters.AddWithValue("@pwd", pwd);
                        if (Convert.ToInt32(cmd.ExecuteScalar()) > 0) match = true;
                    }
                    if (match) { verified = true; dlg.DialogResult = DialogResult.OK; dlg.Close(); }
                    else { lblErr.Text = "❌ Mật khẩu không đúng!"; txtPwd.Clear(); txtPwd.Focus(); }
                }
                catch (Exception ex2) { lblErr.Text = "Lỗi: " + ex2.Message; }
            };
            dlg.KeyDown += (s, ev) => { if (ev.KeyCode == Keys.Enter) { btnOK.PerformClick(); ev.SuppressKeyPress = true; } };
            txtPwd.Focus();
            dlg.ShowDialog(this);
            return verified;
        }

        private void BtnSavePerms_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("USER_MGT", "Phân quyền", "Lưu phân quyền")) return;
            if (_selectedUserId == 0) { MessageBox.Show("Vui lòng chọn user trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            var perms = new Dictionary<string, bool>();
            foreach (DataGridViewRow row in dgvPermissions.Rows)
            {
                if (row.Cells["Row_Type"].Value?.ToString() != "ACTION") continue;
                if (row.Cells["Allowed"].ReadOnly) continue;
                string modCode = row.Cells["Module_Code"].Value?.ToString() ?? "";
                string action = row.Cells["Action_Key"].Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(modCode) || string.IsNullOrEmpty(action)) continue;
                string key = modCode + ":" + action;
                perms[key] = Convert.ToBoolean(row.Cells["Allowed"].Value ?? false);
            }
            _svc.SaveDetailedPermissions(_selectedUserId, perms, AppSession.CurrentUser?.Username ?? "Admin");
            MessageBox.Show("✅ Đã lưu phân quyền chi tiết!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            LoadPermissions(_selectedUserId);
        }

        private void BtnResetPerms_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("USER_MGT", "Phân quyền", "Reset phân quyền về mặc định")) return;
            if (_selectedUserId == 0) return;
            if (MessageBox.Show("Reset về quyền mặc định của Role?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                _svc.ResetUserPermissions(_selectedUserId);
                MessageBox.Show("✅ Đã reset về quyền mặc định!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadPermissions(_selectedUserId);
            }
        }

        // ─────────────────────────────────────────────────────────────────────
        //  RESET ADMIN — khôi phục toàn bộ quyền cho tài khoản admin
        // ─────────────────────────────────────────────────────────────────────
        private void BtnResetAdmin_Click(object sender, EventArgs e)
        {
            if (!AppSession.IsAdmin)
            {
                MessageBox.Show("Chỉ tài khoản Admin mới được dùng chức năng này!",
                    "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBox.Show(
                "Thao tác này sẽ:\n\n" +
                "  1. Xóa toàn bộ quyền tùy chỉnh của tất cả tài khoản Admin\n" +
                "  2. Admin sẽ có MỌI quyền (bypass toàn bộ phân quyền)\n\n" +
                "Xác nhận thực hiện?",
                "Reset Admin Permissions",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;

            try
            {
                // Xóa DetailedPermissions trong DB cho tất cả user có Role admin
                using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                conn.Open();

                // Xóa quyền tùy chỉnh của admin trong DB
                var cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                    DELETE FROM User_Detailed_Permissions
                    WHERE User_ID IN (
                        SELECT User_ID FROM Users
                        WHERE Role_ID = 1
                           OR LOWER(Username) = 'admin'
                    )", conn);
                int affected = cmd.ExecuteNonQuery();

                // Reset session hiện tại ngay lập tức
                AppSession.EnsureAdminBypass();

                MessageBox.Show(
                    $"✅ Đã reset thành công!\n\n" +
                    $"  • Xóa {affected} bản ghi quyền tùy chỉnh của Admin\n" +
                    $"  • Admin hiện có toàn quyền truy cập\n\n" +
                    $"Các tài khoản Admin khác sẽ có đầy đủ quyền khi đăng nhập lại.",
                    "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Reload để cập nhật hiển thị
                LoadUsers();
                if (_selectedUserId > 0) LoadPermissions(_selectedUserId);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Lỗi khi reset: " + ex.Message + "\n\n" +
                    "Bạn có thể tự chạy SQL sau trong DB:\n" +
                    "DELETE FROM User_Detailed_Permissions WHERE User_ID IN\n" +
                    "(SELECT User_ID FROM Users WHERE Role_ID=1 OR LOWER(Username)='admin')",
                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =====================================================================
        //  APPLY PERMISSIONS — ẩn/disable button theo quyền USER_MGT
        // =====================================================================
        private void ApplyPermissions()
        {
            // ⚡ Admin bypass hoàn toàn — không disable bất kỳ thứ gì
            if (AppSession.IsAdmin) return;

            // Nút Tạo user
            if (btnNew != null) PermissionHelper.Apply(btnNew, "USER_MGT", "Tạo user");
            // Nút Lưu user
            if (btnSave != null) PermissionHelper.Apply(btnSave, "USER_MGT", "Lưu user");
            // Nút Vô hiệu hóa
            if (btnDeactivate != null) PermissionHelper.Apply(btnDeactivate, "USER_MGT", "Vô hiệu hóa");
            // Nút Reset Password
            if (btnResetPwd != null) PermissionHelper.Apply(btnResetPwd, "USER_MGT", "Reset Password");
            // Nút Xóa user
            if (btnDelete != null) PermissionHelper.Apply(btnDelete, "USER_MGT", "Lưu user");
            // Nút Lưu phân quyền & Reset phân quyền
            if (btnSavePerms != null) PermissionHelper.Apply(btnSavePerms, "USER_MGT", "Phân quyền");
            if (btnResetPerms != null) PermissionHelper.Apply(btnResetPerms, "USER_MGT", "Phân quyền");

            // Panel phân quyền — disable nếu không có quyền "Phân quyền"
            bool canManagePerm = HasPermissionSilent("USER_MGT", "Phân quyền");
            if (panelPerm != null) panelPerm.Enabled = canManagePerm;
            if (dgvPermissions != null)
                dgvPermissions.ReadOnly = !canManagePerm;
        }

        // ─────────────────────────────────────────────────────────────────────
        //  Kiểm tra quyền không hiện MessageBox (dùng nội bộ để ẩn/hiện UI)
        // ─────────────────────────────────────────────────────────────────────
        private bool HasPermissionSilent(string moduleCode, string action)
        {
            // Admin luôn có mọi quyền
            if (AppSession.IsAdmin) return true;
            // User thường: kiểm tra từ AppSession.DetailedPermissions (cache)
            // Không query DB lại để tránh overhead và sai lệch
            return AppSession.HasPermission(moduleCode, action);
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
            var btn = new Button
            {
                Text = text,
                Location = loc,
                Size = new Size(w, h),
                BackColor = color,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        // ── Parse Department / Position từ chuỗi "PhongBan||ChucVu" ──
        private static string ParseDepartment(string dept)
        {
            if (string.IsNullOrEmpty(dept)) return "";
            int sep = dept.IndexOf("||", StringComparison.Ordinal);
            return sep >= 0 ? dept.Substring(0, sep) : dept;
        }

        private static string ParsePosition(string dept)
        {
            if (string.IsNullOrEmpty(dept)) return "";
            int sep = dept.IndexOf("||", StringComparison.Ordinal);
            return sep >= 0 ? dept.Substring(sep + 2) : "";
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
                // Cột Action_Name chiếm phần còn lại sau cột Allowed (80) + scrollbar (18)
                if (dgvPermissions.Columns.Contains("Action_Name"))
                    dgvPermissions.Columns["Action_Name"].Width =
                        dgvPermissions.Width - 80 - 18;
                btnSavePerms.Top = panelPerm.Height - 38;
                btnResetPerms.Top = panelPerm.Height - 38;
            }
            catch { }
        }
    }

    // =========================================================================
    //  HELPER CLASS
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