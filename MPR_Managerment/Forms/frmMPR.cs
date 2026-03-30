using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using MPR_Managerment.Models;
using MPR_Managerment.Services;

namespace MPR_Managerment.Forms
{
    public partial class frmMPR : Form
    {
        private MPRService _service = new MPRService();
        private List<MPRHeader> _mprList = new List<MPRHeader>();
        private List<MPRDetail> _details = new List<MPRDetail>();
        private int _selectedMPR_ID = 0;
        private string _currentUser = "Admin";

        private DataGridView dgvMPR;
        private TextBox txtSearch;
        private Button btnSearch, btnNewMPR, btnSaveHeader, btnDeleteMPR, btnClearHeader;
        private Label lblStatus;

        private TextBox txtMPRNo, txtProjectName, txtProjectCode, txtDepartment, txtRequestor, txtRev, txtNotes;
        private DateTimePicker dtpRequiredDate;
        private ComboBox cboStatus;

        private DataGridView dgvDetails;
        private Button btnAddDetail, btnDeleteDetail, btnSaveDetail;

        private Panel panelTop, panelHeader, panelDetail;

        public frmMPR()
        {
            InitializeComponent();
            BuildUI();
            LoadMPR();
            this.Resize += FrmMPR_Resize;
            this.WindowState = FormWindowState.Maximized;
        }

        private void BuildUI()
        {
            this.Text = "Quản lý Phiếu Yêu Cầu Mua Hàng (MPR)";
            this.BackColor = Color.FromArgb(245, 245, 245);

            // ===== PANEL TOP =====
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
                Text = "DANH SÁCH PHIẾU MPR",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 10),
                Size = new Size(400, 30)
            });

            txtSearch = new TextBox
            {
                Location = new Point(10, 48),
                Size = new Size(300, 28),
                Font = new Font("Segoe UI", 10),
                PlaceholderText = "Tìm theo MPR No hoặc tên dự án..."
            };
            panelTop.Controls.Add(txtSearch);
            txtSearch.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) BtnSearch_Click(null, null); };

            btnSearch = CreateButton("🔍 Tìm", Color.FromArgb(0, 120, 212), new Point(320, 47), 85, 30);
            btnSearch.Click += BtnSearch_Click;
            panelTop.Controls.Add(btnSearch);

            btnNewMPR = CreateButton("➕ Tạo MPR", Color.FromArgb(40, 167, 69), new Point(415, 47), 110, 30);
            btnNewMPR.Click += BtnNewMPR_Click;
            panelTop.Controls.Add(btnNewMPR);

            btnDeleteMPR = CreateButton("🗑 Xóa MPR", Color.FromArgb(220, 53, 69), new Point(535, 47), 110, 30);
            btnDeleteMPR.Click += BtnDeleteMPR_Click;
            panelTop.Controls.Add(btnDeleteMPR);

            lblStatus = new Label
            {
                Location = new Point(660, 52),
                Size = new Size(500, 25),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.Gray
            };
            panelTop.Controls.Add(lblStatus);

            dgvMPR = new DataGridView
            {
                Location = new Point(10, 85),
                Size = new Size(1335, 125),
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
            dgvMPR.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvMPR.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvMPR.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvMPR.EnableHeadersVisualStyles = false;
            dgvMPR.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvMPR.SelectionChanged += DgvMPR_SelectionChanged;
            panelTop.Controls.Add(dgvMPR);

            // ===== PANEL HEADER =====
            panelHeader = new Panel
            {
                Location = new Point(10, 240),
                Size = new Size(1360, 185),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(panelHeader);

            panelHeader.Controls.Add(new Label
            {
                Text = "THÔNG TIN PHIẾU MPR",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 8),
                Size = new Size(300, 25)
            });

            int y = 38;
            AddLabel(panelHeader, "MPR No (*):", 10, y);
            txtMPRNo = AddTextBox(panelHeader, 110, y, 160);

            AddLabel(panelHeader, "Tên dự án:", 285, y);
            txtProjectName = AddTextBox(panelHeader, 375, y, 250);

            AddLabel(panelHeader, "Mã dự án:", 640, y);
            txtProjectCode = AddTextBox(panelHeader, 720, y, 150);

            AddLabel(panelHeader, "Trạng thái:", 885, y);
            cboStatus = new ComboBox
            {
                Location = new Point(965, y),
                Size = new Size(160, 25),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboStatus.Items.AddRange(new[] { "Mới", "Đang xử lý", "Đã duyệt", "Hoàn thành", "Hủy" });
            cboStatus.SelectedIndex = 0;
            panelHeader.Controls.Add(cboStatus);

            y += 38;
            AddLabel(panelHeader, "Phòng ban:", 10, y);
            txtDepartment = AddTextBox(panelHeader, 110, y, 160);

            AddLabel(panelHeader, "Người yêu cầu:", 285, y);
            txtRequestor = AddTextBox(panelHeader, 390, y, 175);

            AddLabel(panelHeader, "Ngày cần:", 580, y);
            dtpRequiredDate = new DateTimePicker
            {
                Location = new Point(650, y),
                Size = new Size(150, 25),
                Font = new Font("Segoe UI", 9),
                Format = DateTimePickerFormat.Short
            };
            panelHeader.Controls.Add(dtpRequiredDate);

            AddLabel(panelHeader, "Rev:", 815, y);
            txtRev = AddTextBox(panelHeader, 848, y, 60);
            txtRev.Text = "0";

            AddLabel(panelHeader, "Ghi chú:", 925, y);
            txtNotes = AddTextBox(panelHeader, 990, y, 240);
            txtNotes.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            y += 45;
            btnSaveHeader = CreateButton("💾 Lưu Header", Color.FromArgb(0, 120, 212), new Point(10, y), 130, 32);
            btnSaveHeader.Click += BtnSaveHeader_Click;
            panelHeader.Controls.Add(btnSaveHeader);

            btnClearHeader = CreateButton("🔄 Làm mới", Color.FromArgb(108, 117, 125), new Point(150, y), 110, 32);
            btnClearHeader.Click += BtnClearHeader_Click;
            panelHeader.Controls.Add(btnClearHeader);

            // ===== PANEL DETAIL =====
            panelDetail = new Panel
            {
                Location = new Point(10, 435),
                Size = new Size(1360, 345),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            this.Controls.Add(panelDetail);

            panelDetail.Controls.Add(new Label
            {
                Text = "CHI TIẾT VẬT TƯ",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 8),
                Size = new Size(300, 25)
            });

            btnAddDetail = CreateButton("➕ Thêm dòng", Color.FromArgb(40, 167, 69), new Point(10, 38), 120, 30);
            btnAddDetail.Click += BtnAddDetail_Click;
            panelDetail.Controls.Add(btnAddDetail);

            btnDeleteDetail = CreateButton("🗑 Xóa dòng", Color.FromArgb(220, 53, 69), new Point(140, 38), 110, 30);
            btnDeleteDetail.Click += BtnDeleteDetail_Click;
            panelDetail.Controls.Add(btnDeleteDetail);

            btnSaveDetail = CreateButton("💾 Lưu chi tiết", Color.FromArgb(0, 120, 212), new Point(260, 38), 130, 30);
            btnSaveDetail.Click += BtnSaveDetail_Click;
            panelDetail.Controls.Add(btnSaveDetail);

            dgvDetails = new DataGridView
            {
                Location = new Point(10, 75),
                Size = new Size(1335, 260),
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            dgvDetails.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvDetails.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvDetails.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvDetails.EnableHeadersVisualStyles = false;
            dgvDetails.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            BuildDetailColumns();
            panelDetail.Controls.Add(dgvDetails);
            // Đưa tất cả TextBox và ComboBox lên trên Label
            //foreach (Panel panel in new[] { panelHead, panelTop, panelDetail })
            //{
            //    foreach (Control c in panel.Controls)
            //    {
            //        if (c is TextBox || c is ComboBox || c is DateTimePicker || c is NumericUpDown)
            //            c.BringToFront();
            //    }
            //}
        }

        private void BuildDetailColumns()
        {
            dgvDetails.Columns.Clear();

            // Hidden ID
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            { Name = "Detail_ID", HeaderText = "ID", Visible = false });

            // Visible columns
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            { Name = "Item_No", HeaderText = "STT", Width = 45, ReadOnly = true });

            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            { Name = "Item_Name", HeaderText = "Tên vật tư", Width = 180 });

            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            { Name = "Description", HeaderText = "Mô tả", Width = 100 });

            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            { Name = "Material", HeaderText = "Vật liệu", Width = 85 });

            // ===== 6 CỘT KÍCH THƯỚC =====
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            { Name = "Thickness_mm", HeaderText = "A-Dày(mm)", Width = 75 });

            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            { Name = "Depth_mm", HeaderText = "B-Sâu(mm)", Width = 75 });

            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            { Name = "C_Width_mm", HeaderText = "C-Rộng(mm)", Width = 80 });

            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            { Name = "D_Web_mm", HeaderText = "D-Bụng(mm)", Width = 80 });

            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            { Name = "E_Flange_mm", HeaderText = "E-Cánh(mm)", Width = 80 });

            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            { Name = "F_Length_mm", HeaderText = "F-Dài(mm)", Width = 75 });
            // ================================

            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            { Name = "UNIT", HeaderText = "ĐVT", Width = 50 });

            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            { Name = "Qty", HeaderText = "SL", Width = 50 });

            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            { Name = "Weight", HeaderText = "KG", Width = 55 });

            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            { Name = "MPS_Info", HeaderText = "MPS Info", Width = 100 });

            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            { Name = "Usage_Location", HeaderText = "Vị trí dùng", Width = 110 });

            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            { Name = "REV", HeaderText = "REV", Width = 45 });

            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            { Name = "Remarks", HeaderText = "Ghi chú", FillWeight = 100 });
        }

        private void AddLabel(Panel panel, string text, int x, int y)
        {
            panel.Controls.Add(new Label
            {
                Text = text,
                Location = new Point(x, y + 3),
                Size = new Size(110, 20),
                Font = new Font("Segoe UI", 9)
            });
        }

        private TextBox AddTextBox(Panel panel, int x, int y, int width)
        {
            var txt = new TextBox
            {
                Location = new Point(x, y),
                Size = new Size(width, 25),
                Font = new Font("Segoe UI", 9)
            };
            panel.Controls.Add(txt);
            return txt;
        }

        private Button CreateButton(string text, Color color, Point location, int w, int h)
        {
            var btn = new Button
            {
                Text = text,
                Location = location,
                Size = new Size(w, h),
                BackColor = color,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }// ===== RESIZE =====
        private void FrmMPR_Resize(object sender, EventArgs e)
        {
            try
            {
                int w = this.ClientSize.Width - 20;
                int h = this.ClientSize.Height;

                panelTop.Width = w;
                panelHeader.Width = w;
                panelDetail.Width = w;
                panelDetail.Height = h - panelDetail.Top - 10;

                dgvMPR.Width = panelTop.Width - 20;
                dgvDetails.Width = panelDetail.Width - 20;
                dgvDetails.Height = panelDetail.Height - 85;

                if (txtNotes != null && panelHeader != null)
                    txtNotes.Width = panelHeader.Width - txtNotes.Left - 20;
            }
            catch { }
        }

        // ===== LOAD MPR =====
        private void LoadMPR()
        {
            try
            {
                _mprList = _service.GetAll();
                BindMPRGrid(_mprList);
                lblStatus.Text = $"Tổng: {_mprList.Count} phiếu MPR";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải MPR: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BindMPRGrid(List<MPRHeader> list)
        {
            dgvMPR.DataSource = list.ConvertAll(m => new
            {
                ID = m.MPR_ID,
                MPR_No = m.MPR_No,
                Ten_Du_An = m.Project_Name,
                Ma_Du_An = m.Project_Code,
                Phong_Ban = m.Department,
                Nguoi_YC = m.Requestor,
                Ngay_Can = m.Required_Date.HasValue ? m.Required_Date.Value.ToString("dd/MM/yyyy") : "",
                Rev = m.Rev,
                Trang_Thai = m.Status,
                Ngay_Tao = m.Created_Date.HasValue ? m.Created_Date.Value.ToString("dd/MM/yyyy") : ""
            });
            if (dgvMPR.Columns.Contains("ID"))
                dgvMPR.Columns["ID"].Visible = false;
        }

        // ===== LOAD DETAILS =====
        private void LoadDetails(int mprId)
        {
            try
            {
                _details = _service.GetDetails(mprId);
                dgvDetails.Rows.Clear();

                foreach (var d in _details)
                {
                    int idx = dgvDetails.Rows.Add();
                    var row = dgvDetails.Rows[idx];

                    row.Cells["Detail_ID"].Value = d.Detail_ID;
                    row.Cells["Item_No"].Value = d.Item_No;
                    row.Cells["Item_Name"].Value = d.Item_Name;
                    row.Cells["Description"].Value = d.Description;
                    row.Cells["Material"].Value = d.Material;
                    row.Cells["Thickness_mm"].Value = d.Thickness_mm;
                    row.Cells["Depth_mm"].Value = d.Depth_mm;
                    row.Cells["C_Width_mm"].Value = d.C_Width_mm;
                    row.Cells["D_Web_mm"].Value = d.D_Web_mm;
                    row.Cells["E_Flange_mm"].Value = d.E_Flange_mm;
                    row.Cells["F_Length_mm"].Value = d.F_Length_mm;
                    row.Cells["UNIT"].Value = d.UNIT;
                    row.Cells["Qty"].Value = d.Qty_Per_Sheet;
                    row.Cells["Weight"].Value = d.Weight_kg;
                    row.Cells["MPS_Info"].Value = d.MPS_Info;
                    row.Cells["Usage_Location"].Value = d.Usage_Location;
                    row.Cells["REV"].Value = d.REV;
                    row.Cells["Remarks"].Value = d.Remarks;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải chi tiết: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ===== SỰ KIỆN =====
        private void DgvMPR_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvMPR.SelectedRows.Count == 0) return;
            var row = dgvMPR.SelectedRows[0];
            _selectedMPR_ID = Convert.ToInt32(row.Cells["ID"].Value);

            var m = _mprList.Find(x => x.MPR_ID == _selectedMPR_ID);
            if (m == null) return;

            txtMPRNo.Text = m.MPR_No;
            txtProjectName.Text = m.Project_Name;
            txtProjectCode.Text = m.Project_Code;
            txtDepartment.Text = m.Department;
            txtRequestor.Text = m.Requestor;
            txtRev.Text = m.Rev.ToString();
            txtNotes.Text = m.Notes;
            dtpRequiredDate.Value = m.Required_Date ?? DateTime.Today;

            int idx = cboStatus.Items.IndexOf(m.Status);
            cboStatus.SelectedIndex = idx >= 0 ? idx : 0;

            LoadDetails(_selectedMPR_ID);
        }

        private void BtnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                string kw = txtSearch.Text.Trim();
                _mprList = string.IsNullOrEmpty(kw)
                    ? _service.GetAll()
                    : _service.GetAll().FindAll(m =>
                        (m.MPR_No ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                        (m.Project_Name ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                        (m.Project_Code ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase));

                BindMPRGrid(_mprList);
                lblStatus.Text = $"Tìm thấy: {_mprList.Count} phiếu";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tìm kiếm: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnNewMPR_Click(object sender, EventArgs e)
        {
            _selectedMPR_ID = 0;
            ClearHeader();
            dgvDetails.Rows.Clear();
            _details.Clear();
            txtMPRNo.Focus();
        }

        private void BtnSaveHeader_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtMPRNo.Text))
            {
                MessageBox.Show("Vui lòng nhập MPR No!", "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtMPRNo.Focus();
                return;
            }

            try
            {
                var m = new MPRHeader
                {
                    MPR_ID = _selectedMPR_ID,
                    MPR_No = txtMPRNo.Text.Trim(),
                    Project_Name = txtProjectName.Text.Trim(),
                    Project_Code = txtProjectCode.Text.Trim(),
                    Department = txtDepartment.Text.Trim(),
                    Requestor = txtRequestor.Text.Trim(),
                    Required_Date = dtpRequiredDate.Value,
                    Rev = int.TryParse(txtRev.Text, out int rev) ? rev : 0,
                    Status = cboStatus.SelectedItem?.ToString() ?? "Mới",
                    Notes = txtNotes.Text.Trim()
                };

                if (_selectedMPR_ID == 0)
                {
                    _selectedMPR_ID = _service.InsertHeader(m, _currentUser);
                    MessageBox.Show("Tạo phiếu MPR thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    _service.UpdateHeader(m, _currentUser);
                    MessageBox.Show("Cập nhật phiếu MPR thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                LoadMPR();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi lưu header: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnDeleteMPR_Click(object sender, EventArgs e)
        {
            if (_selectedMPR_ID == 0)
            {
                MessageBox.Show("Vui lòng chọn phiếu MPR cần xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show("Xóa phiếu MPR này và toàn bộ chi tiết?", "Xác nhận",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    _service.DeleteMPR(_selectedMPR_ID);
                    MessageBox.Show("Xóa thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    _selectedMPR_ID = 0;
                    ClearHeader();
                    dgvDetails.Rows.Clear();
                    LoadMPR();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnClearHeader_Click(object sender, EventArgs e)
        {
            _selectedMPR_ID = 0;
            ClearHeader();
            dgvDetails.Rows.Clear();
            _details.Clear();
        }

        private void BtnAddDetail_Click(object sender, EventArgs e)
        {
            if (_selectedMPR_ID == 0)
            {
                MessageBox.Show("Vui lòng chọn hoặc lưu phiếu MPR trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int nextNo = dgvDetails.Rows.Count + 1;
            int newIdx = dgvDetails.Rows.Add();
            var newRow = dgvDetails.Rows[newIdx];

            newRow.Cells["Detail_ID"].Value = 0;
            newRow.Cells["Item_No"].Value = nextNo;
            newRow.Cells["Item_Name"].Value = "";
            newRow.Cells["Description"].Value = "";
            newRow.Cells["Material"].Value = "";
            newRow.Cells["Thickness_mm"].Value = 0;
            newRow.Cells["Depth_mm"].Value = 0;
            newRow.Cells["C_Width_mm"].Value = 0;
            newRow.Cells["D_Web_mm"].Value = 0;
            newRow.Cells["E_Flange_mm"].Value = 0;
            newRow.Cells["F_Length_mm"].Value = 0;
            newRow.Cells["UNIT"].Value = "cái";
            newRow.Cells["Qty"].Value = 0;
            newRow.Cells["Weight"].Value = 0;
            newRow.Cells["MPS_Info"].Value = "";
            newRow.Cells["Usage_Location"].Value = "";
            newRow.Cells["REV"].Value = "0";
            newRow.Cells["Remarks"].Value = "";

            dgvDetails.CurrentCell = dgvDetails.Rows[newIdx].Cells["Item_Name"];
            dgvDetails.BeginEdit(true);
        }

        private void BtnDeleteDetail_Click(object sender, EventArgs e)
        {
            if (dgvDetails.SelectedRows.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn dòng cần xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var row = dgvDetails.SelectedRows[0];
            int detailId = Convert.ToInt32(row.Cells["Detail_ID"].Value ?? 0);

            if (detailId > 0)
            {
                if (MessageBox.Show("Xóa dòng vật tư này?", "Xác nhận",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    try
                    {
                        _service.DeleteDetail(detailId);
                        dgvDetails.Rows.Remove(row);
                        MessageBox.Show("Đã xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
                dgvDetails.Rows.Remove(row);
            }
        }

        private void BtnSaveDetail_Click(object sender, EventArgs e)
        {
            if (_selectedMPR_ID == 0)
            {
                MessageBox.Show("Vui lòng lưu header MPR trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (dgvDetails.Rows.Count == 0)
            {
                MessageBox.Show("Không có dòng nào để lưu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                int saved = 0;
                foreach (DataGridViewRow row in dgvDetails.Rows)
                {
                    string itemName = row.Cells["Item_Name"].Value?.ToString() ?? "";
                    if (string.IsNullOrWhiteSpace(itemName)) continue;

                    var d = new MPRDetail
                    {
                        Detail_ID = Convert.ToInt32(row.Cells["Detail_ID"].Value ?? 0),
                        MPR_ID = _selectedMPR_ID,
                        Item_No = Convert.ToInt32(row.Cells["Item_No"].Value ?? 0),
                        Item_Name = itemName,
                        Description = row.Cells["Description"].Value?.ToString() ?? "",
                        Material = row.Cells["Material"].Value?.ToString() ?? "",
                        Thickness_mm = DecimalVal(row.Cells["Thickness_mm"].Value),
                        Depth_mm = DecimalVal(row.Cells["Depth_mm"].Value),
                        C_Width_mm = DecimalVal(row.Cells["C_Width_mm"].Value),
                        D_Web_mm = DecimalVal(row.Cells["D_Web_mm"].Value),
                        E_Flange_mm = DecimalVal(row.Cells["E_Flange_mm"].Value),
                        F_Length_mm = DecimalVal(row.Cells["F_Length_mm"].Value),
                        UNIT = row.Cells["UNIT"].Value?.ToString() ?? "",
                        Qty_Per_Sheet = (int)DecimalVal(row.Cells["Qty"].Value),
                        Weight_kg = DecimalVal(row.Cells["Weight"].Value),
                        MPS_Info = row.Cells["MPS_Info"].Value?.ToString() ?? "",
                        Usage_Location = row.Cells["Usage_Location"].Value?.ToString() ?? "",
                        REV = row.Cells["REV"].Value?.ToString() ?? "0",
                        Remarks = row.Cells["Remarks"].Value?.ToString() ?? ""
                    };

                    if (d.Detail_ID == 0)
                        _service.InsertDetail(d, _currentUser);
                    else
                        _service.UpdateDetail(d, _currentUser);

                    saved++;
                }

                MessageBox.Show($"Đã lưu {saved} dòng chi tiết thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDetails(_selectedMPR_ID);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi lưu chi tiết: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ===== HELPERS =====
        private decimal DecimalVal(object val)
        {
            if (val == null || val == DBNull.Value) return 0;
            return decimal.TryParse(val.ToString(), out decimal d) ? d : 0;
        }

        private void ClearHeader()
        {
            txtMPRNo.Text = "";
            txtProjectName.Text = "";
            txtProjectCode.Text = "";
            txtDepartment.Text = "";
            txtRequestor.Text = "";
            txtRev.Text = "0";
            txtNotes.Text = "";
            dtpRequiredDate.Value = DateTime.Today;
            cboStatus.SelectedIndex = 0;
        }
    }
}