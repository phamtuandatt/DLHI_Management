using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using MPR_Managerment.Models;
using MPR_Managerment.Services;

namespace MPR_Managerment.Forms
{
    public partial class frmWarehouseSetup : Form
    {
        private WarehouseLocationService _service = new WarehouseLocationService();
        private List<Warehouse> _warehouses = new List<Warehouse>();
        private int _selectedID = 0;
        private string _currentUser = "Admin";

        private DataGridView dgvWarehouses;
        private TextBox txtSearch, txtWarehouseName, txtDeptAbbr, txtManager, txtNotes;
        private ComboBox cboProjectCode, cboWarehouseType;
        private Label lblWarehouseCode, lblStatus;
        private Button btnNew, btnSave, btnDelete, btnClear, btnSearch;
        private Panel panelList, panelForm;

        public frmWarehouseSetup()
        {
            InitializeComponent();
            BuildUI();
            LoadProjectCombo();
            LoadWarehouses();
            this.Resize += FrmWarehouseSetup_Resize;
        }
        

        private void BuildUI()
        {
            this.Text = "Quản lý Danh Mục Kho";
            this.BackColor = Color.FromArgb(245, 245, 245);

            // ===== PANEL LIST =====
            panelList = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(1200, 280),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(panelList);

            panelList.Controls.Add(new Label
            {
                Text = "DANH SÁCH KHO",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 10),
                Size = new Size(300, 30)
            });

            txtSearch = new TextBox
            {
                Location = new Point(10, 48),
                Size = new Size(250, 28),
                Font = new Font("Segoe UI", 10),
                PlaceholderText = "Tìm mã kho, tên kho..."
            };
            panelList.Controls.Add(txtSearch);
            txtSearch.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) LoadWarehouses(); };

            btnSearch = CreateBtn("🔍 Tìm", Color.FromArgb(0, 120, 212), new Point(268, 47), 80, 30);
            btnSearch.Click += (s, e) => LoadWarehouses();
            panelList.Controls.Add(btnSearch);

            btnNew = CreateBtn("➕ Tạo kho", Color.FromArgb(40, 167, 69), new Point(358, 47), 110, 30);
            btnNew.Click += BtnNew_Click;
            panelList.Controls.Add(btnNew);

            btnDelete = CreateBtn("🗑 Xóa kho", Color.FromArgb(220, 53, 69), new Point(478, 47), 110, 30);
            btnDelete.Click += BtnDelete_Click;
            panelList.Controls.Add(btnDelete);

            lblStatus = new Label
            {
                Location = new Point(600, 52),
                Size = new Size(400, 22),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.Gray
            };
            panelList.Controls.Add(lblStatus);

            dgvWarehouses = new DataGridView
            {
                Location = new Point(10, 85),
                Size = new Size(1175, 185),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            dgvWarehouses.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvWarehouses.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvWarehouses.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvWarehouses.EnableHeadersVisualStyles = false;
            dgvWarehouses.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvWarehouses.SelectionChanged += DgvWarehouses_SelectionChanged;
            dgvWarehouses.CellFormatting += DgvWarehouses_CellFormatting;
            panelList.Controls.Add(dgvWarehouses);

            // ===== PANEL FORM =====
            panelForm = new Panel
            {
                Location = new Point(10, 300),
                Size = new Size(1200, 340),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            this.Controls.Add(panelForm);

            panelForm.Controls.Add(new Label
            {
                Text = "THÔNG TIN KHO",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 8),
                Size = new Size(300, 25)
            });

            // Hiển thị mã kho preview
            panelForm.Controls.Add(new Label
            {
                Text = "Mã kho:",
                Font = new Font("Segoe UI", 9),
                Location = new Point(10, 40),
                Size = new Size(60, 20)
            });
            lblWarehouseCode = new Label
            {
                Text = "(chưa có — chọn dự án và nhập viết tắt bộ phận)",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.Gray,
                Location = new Point(75, 40),
                Size = new Size(700, 20)
            };
            panelForm.Controls.Add(lblWarehouseCode);

            // Row 1
            int y = 70;
            AddLbl(panelForm, "Dự án (*):", 10, y);
            cboProjectCode = new ComboBox
            {
                Location = new Point(110, y),
                Size = new Size(200, 25),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboProjectCode.Items.Add("-- Chọn dự án --");
            cboProjectCode.SelectedIndex = 0;
            cboProjectCode.SelectedIndexChanged += CboProjectCode_Changed;
            panelForm.Controls.Add(cboProjectCode);

            AddLbl(panelForm, "Loại kho (*):", 325, y);
            cboWarehouseType = new ComboBox
            {
                Location = new Point(425, y),
                Size = new Size(200, 25),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboWarehouseType.Items.AddRange(new[]
            {
                "-- Chọn loại kho --",
                "RM - Nguyên vật liệu",
                "SUP - Vật tư phụ",
                "EQ - Thiết bị"
            });
            cboWarehouseType.SelectedIndex = 0;
            cboWarehouseType.SelectedIndexChanged += CboWarehouseType_Changed;
            panelForm.Controls.Add(cboWarehouseType);

            AddLbl(panelForm, "Viết tắt BP (*):", 640, y);
            txtDeptAbbr = AddTb(panelForm, 760, y, 100);
            txtDeptAbbr.PlaceholderText = "VD: RM, WH";
            txtDeptAbbr.TextChanged += TxtDeptAbbr_Changed;

            AddLbl(panelForm, "Người phụ trách:", 875, y);
            txtManager = AddTb(panelForm, 1000, y, 170);

            // Row 2
            y += 40;
            AddLbl(panelForm, "Tên kho (*):", 10, y);
            txtWarehouseName = AddTb(panelForm, 110, y, 500);
            txtWarehouseName.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            // Row 3
            y += 40;
            AddLbl(panelForm, "Ghi chú:", 10, y);
            txtNotes = AddTb(panelForm, 110, y, 900);
            txtNotes.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            // Buttons
            y += 50;
            btnSave = CreateBtn("💾 Lưu kho", Color.FromArgb(0, 120, 212), new Point(10, y), 120, 34);
            btnSave.Click += BtnSave_Click;
            panelForm.Controls.Add(btnSave);

            btnClear = CreateBtn("🔄 Làm mới", Color.FromArgb(108, 117, 125), new Point(140, y), 110, 34);
            btnClear.Click += (s, e) => ClearForm();
            panelForm.Controls.Add(btnClear);

            panelForm.Controls.Add(new Label
            {
                Text = "💡 Mã kho tự động: [Mã dự án]-[Viết tắt bộ phận]   Ví dụ: 25G10-DEC-RM",
                Font = new Font("Segoe UI", 9, FontStyle.Italic),
                ForeColor = Color.FromArgb(255, 140, 0),
                Location = new Point(10, y + 42),
                Size = new Size(900, 20)
            });
            // Đưa tất cả input controls lên trên labels
            foreach (Control c in panelForm.Controls)
            {
                if (c is TextBox || c is ComboBox || c is DateTimePicker || c is NumericUpDown)
                    c.BringToFront();
            }
        }

        // ===== HELPERS =====
        private void AddLbl(Panel p, string text, int x, int y)
        {
            var lbl = new Label
            {
                Text = text,
                Location = new Point(x, y + 3),
                Size = new Size(120, 20),
                Font = new Font("Segoe UI", 9),
                BackColor = Color.Transparent
            };
            p.Controls.Add(lbl);
            lbl.SendToBack();
        }

        private TextBox AddTb(Panel p, int x, int y, int w)
        {
            var txt = new TextBox { Location = new Point(x, y), Size = new Size(w, 25), Font = new Font("Segoe UI", 9) };
            p.Controls.Add(txt);
            return txt;
        }

        private Button CreateBtn(string text, Color color, Point loc, int w, int h)
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

        private void LoadProjectCombo()
        {
            try
            {
                var projects = new ProjectService().GetAll();
                cboProjectCode.Items.Clear();
                cboProjectCode.Items.Add("-- Chọn dự án --");
                foreach (var p in projects)
                    cboProjectCode.Items.Add($"{p.ProjectCode} — {p.ProjectName}");
                cboProjectCode.SelectedIndex = 0;
            }
            catch { }
        }

        private string GetProjectCode()
        {
            if (cboProjectCode.SelectedIndex <= 0) return "";
            return cboProjectCode.SelectedItem.ToString().Split('—')[0].Trim();
        }

        private string GetWarehouseTypeCode()
        {
            if (cboWarehouseType.SelectedIndex <= 0) return "";
            return cboWarehouseType.SelectedItem.ToString().Split('-')[0].Trim();
        }

        private void UpdateWarehouseCodePreview()
        {
            string proj = GetProjectCode();
            string dept = txtDeptAbbr?.Text.Trim().ToUpper() ?? "";
            string type = GetWarehouseTypeCode();

            if (!string.IsNullOrEmpty(proj) && !string.IsNullOrEmpty(dept))
            {
                string preview = $"{proj}-{dept}";
                lblWarehouseCode.Text = $"✅  {preview}";
                lblWarehouseCode.ForeColor = Color.FromArgb(40, 167, 69);

                // Tự động gợi ý tên kho nếu chưa nhập
                if (string.IsNullOrEmpty(txtWarehouseName?.Text) && !string.IsNullOrEmpty(type))
                {
                    string typeName =
                        type == "RM" ? "Kho nguyên vật liệu" :
                        type == "SUP" ? "Kho vật tư phụ" :
                        type == "EQ" ? "Kho thiết bị" : "Kho";
                    txtWarehouseName.Text = $"{typeName} — {proj}";
                }
            }
            else
            {
                lblWarehouseCode.Text = "(chưa có — chọn dự án và nhập viết tắt bộ phận)";
                lblWarehouseCode.ForeColor = Color.Gray;
            }
        }// ===== RESIZE =====
        private void FrmWarehouseSetup_Resize(object sender, EventArgs e)
        {
            try
            {
                int w = this.ClientSize.Width - 20;
                int h = this.ClientSize.Height;

                panelList.Width = w;
                panelForm.Width = w;
                panelForm.Height = h - panelForm.Top - 10;

                dgvWarehouses.Width = panelList.Width - 20;
                txtWarehouseName.Width = panelForm.Width - txtWarehouseName.Left - 20;
                txtNotes.Width = panelForm.Width - txtNotes.Left - 20;
                txtManager.Width = panelForm.Width - txtManager.Left - 20;
            }
            catch { }
        }

        // ===== LOAD =====
        private void LoadWarehouses()
        {
            try
            {
                string kw = txtSearch?.Text.Trim() ?? "";
                var all = _service.GetAll();

                if (!string.IsNullOrEmpty(kw))
                    all = all.FindAll(w =>
                        w.Warehouse_Code.Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                        w.Warehouse_Name.Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                        w.Project_Code.Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                        w.Manager.Contains(kw, StringComparison.OrdinalIgnoreCase));

                _warehouses = all;

                dgvWarehouses.DataSource = _warehouses.ConvertAll(w => new
                {
                    ID = w.Warehouse_ID,
                    Ma_Kho = w.Warehouse_Code,
                    Ten_Kho = w.Warehouse_Name,
                    Loai_Kho = w.Warehouse_Type,
                    Ma_Du_An = w.Project_Code,
                    Bo_Phan = w.Dept_Abbr,
                    Phu_Trach = w.Manager,
                    Trang_Thai = w.IsActive ? "Hoạt động" : "Tạm dừng",
                    Ngay_Tao = w.Created_Date.HasValue
                                 ? w.Created_Date.Value.ToString("dd/MM/yyyy") : "",
                    Ghi_Chu = w.Notes
                });

                if (dgvWarehouses.Columns.Contains("ID"))
                    dgvWarehouses.Columns["ID"].Visible = false;

                lblStatus.Text = $"Tổng: {_warehouses.Count} kho";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải danh sách kho: " + ex.Message, "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ===== SỰ KIỆN =====
        private void CboProjectCode_Changed(object sender, EventArgs e)
        {
            txtWarehouseName.Text = "";
            UpdateWarehouseCodePreview();
        }

        private void CboWarehouseType_Changed(object sender, EventArgs e)
        {
            txtWarehouseName.Text = "";
            UpdateWarehouseCodePreview();
        }

        private void TxtDeptAbbr_Changed(object sender, EventArgs e)
        {
            // Tự động viết hoa
            int pos = txtDeptAbbr.SelectionStart;
            txtDeptAbbr.TextChanged -= TxtDeptAbbr_Changed;
            txtDeptAbbr.Text = txtDeptAbbr.Text.ToUpper();
            txtDeptAbbr.SelectionStart = pos;
            txtDeptAbbr.TextChanged += TxtDeptAbbr_Changed;

            UpdateWarehouseCodePreview();
        }

        private void DgvWarehouses_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvWarehouses.SelectedRows.Count == 0) return;
            var row = dgvWarehouses.SelectedRows[0];
            _selectedID = Convert.ToInt32(row.Cells["ID"].Value);

            var w = _warehouses.Find(x => x.Warehouse_ID == _selectedID);
            if (w == null) return;

            // Điền thông tin vào form
            lblWarehouseCode.Text = $"✅  {w.Warehouse_Code}";
            lblWarehouseCode.ForeColor = Color.FromArgb(40, 167, 69);
            txtWarehouseName.Text = w.Warehouse_Name;
            txtDeptAbbr.Text = w.Dept_Abbr;
            txtManager.Text = w.Manager;
            txtNotes.Text = w.Notes;

            // Chọn loại kho tương ứng
            for (int i = 0; i < cboWarehouseType.Items.Count; i++)
            {
                if (cboWarehouseType.Items[i].ToString().StartsWith(w.Warehouse_Type))
                {
                    cboWarehouseType.SelectedIndex = i;
                    break;
                }
            }

            // Chọn dự án tương ứng
            for (int i = 0; i < cboProjectCode.Items.Count; i++)
            {
                if (cboProjectCode.Items[i].ToString().StartsWith(w.Project_Code))
                {
                    cboProjectCode.SelectedIndex = i;
                    break;
                }
            }
        }

        private void DgvWarehouses_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string col = dgvWarehouses.Columns[e.ColumnIndex].Name;

            if (col == "Trang_Thai")
            {
                e.CellStyle.ForeColor = e.Value?.ToString() == "Hoạt động"
                    ? Color.FromArgb(40, 167, 69)
                    : Color.FromArgb(220, 53, 69);
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }

            if (col == "Loai_Kho")
            {
                string val = e.Value?.ToString() ?? "";
                e.CellStyle.ForeColor =
                    val == "RM" ? Color.FromArgb(0, 120, 212) :
                    val == "SUP" ? Color.FromArgb(255, 140, 0) :
                    val == "EQ" ? Color.FromArgb(102, 51, 153) :
                                   Color.Black;
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
        }

        private void BtnNew_Click(object sender, EventArgs e)
        {
            _selectedID = 0;
            ClearForm();
            LoadProjectCombo();
            cboProjectCode.Focus();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            // Validate
            if (cboProjectCode.SelectedIndex <= 0)
            {
                MessageBox.Show("Vui lòng chọn Dự án!", "Thiếu thông tin",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cboProjectCode.Focus();
                return;
            }
            if (cboWarehouseType.SelectedIndex <= 0)
            {
                MessageBox.Show("Vui lòng chọn Loại kho!", "Thiếu thông tin",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cboWarehouseType.Focus();
                return;
            }
            if (string.IsNullOrWhiteSpace(txtDeptAbbr.Text))
            {
                MessageBox.Show("Vui lòng nhập Viết tắt bộ phận!", "Thiếu thông tin",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtDeptAbbr.Focus();
                return;
            }
            if (string.IsNullOrWhiteSpace(txtWarehouseName.Text))
            {
                MessageBox.Show("Vui lòng nhập Tên kho!", "Thiếu thông tin",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtWarehouseName.Focus();
                return;
            }

            try
            {
                string projectCode = GetProjectCode();
                string typeCode = GetWarehouseTypeCode();
                string deptAbbr = txtDeptAbbr.Text.Trim().ToUpper();

                if (_selectedID == 0)
                {
                    // Tạo mới
                    string warehouseCode = _service.GenerateCode(projectCode, deptAbbr);

                    var w = new Warehouse
                    {
                        Warehouse_Code = warehouseCode,
                        Warehouse_Name = txtWarehouseName.Text.Trim(),
                        Warehouse_Type = typeCode,
                        Project_Code = projectCode,
                        Dept_Abbr = deptAbbr,
                        Manager = txtManager.Text.Trim(),
                        Notes = txtNotes.Text.Trim(),
                        IsActive = true
                    };

                    _service.Insert(w, _currentUser);
                    MessageBox.Show(
                        $"✅ Tạo kho thành công!\n" +
                        $"Mã kho: {warehouseCode}\n" +
                        $"Tên kho: {w.Warehouse_Name}",
                        "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    // Cập nhật
                    var w = new Warehouse
                    {
                        Warehouse_ID = _selectedID,
                        Warehouse_Name = txtWarehouseName.Text.Trim(),
                        Warehouse_Type = typeCode,
                        Project_Code = projectCode,
                        Dept_Abbr = deptAbbr,
                        Manager = txtManager.Text.Trim(),
                        Notes = txtNotes.Text.Trim(),
                        IsActive = true
                    };

                    _service.Update(w, _currentUser);
                    MessageBox.Show("✅ Cập nhật kho thành công!",
                        "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                LoadWarehouses();
                ClearForm();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi lưu kho: " + ex.Message, "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (_selectedID == 0)
            {
                MessageBox.Show("Vui lòng chọn kho cần xóa!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var w = _warehouses.Find(x => x.Warehouse_ID == _selectedID);
            if (MessageBox.Show(
                $"Vô hiệu hóa kho:\n{w?.Warehouse_Code} — {w?.Warehouse_Name}?\n\n(Kho sẽ bị ẩn khỏi danh sách chọn)",
                "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    _service.Delete(_selectedID);
                    MessageBox.Show("Đã vô hiệu hóa kho!", "Thông báo",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    _selectedID = 0;
                    ClearForm();
                    LoadWarehouses();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.Message, "Lỗi",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void ClearForm()
        {
            _selectedID = 0;
            cboProjectCode.SelectedIndex = 0;
            cboWarehouseType.SelectedIndex = 0;
            txtDeptAbbr.Text = "";
            txtWarehouseName.Text = "";
            txtManager.Text = "";
            txtNotes.Text = "";
            lblWarehouseCode.Text = "(chưa có — chọn dự án và nhập viết tắt bộ phận)";
            lblWarehouseCode.ForeColor = Color.Gray;
        }
    }
}