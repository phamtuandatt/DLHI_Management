using MPR_Managerment.Helpers;
using MPR_Managerment.Models;
using MPR_Managerment.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using OfficeOpenXml; // Đã bổ sung thư viện SQL

namespace MPR_Managerment.Forms
{
    public partial class frmRIR : Form
    {
        private RIRService _service = new RIRService();
        private POService _poService = new POService();
        private ProjectService _projectServices = new ProjectService();
        private SupplierService _supplierServices = new SupplierService();

        private List<RIRHead> _rirList = new List<RIRHead>();
        private List<RIRDetail> _details = new List<RIRDetail>();
        private int _selectedRIR_ID = 0;
        private string _currentUser = "Admin";

        // ===== CONTROLS =====
        private DataGridView dgvRIR, dgvDetails;
        private TextBox txtSearch, txtRIRNo, txtProjectName, txtWorkorderNo;
        private TextBox txtMPRNo, txtCustomer, txtPONo;
        private DateTimePicker dtpIssueDate;
        private ComboBox cboStatus;
        private Button btnSearch, btnNewRIR, btnSaveHead, btnDeleteRIR;
        private Button btnAddDetail, btnDeleteDetail, btnSaveDetail, btnImportPO;
        private Label lblStatus;
        private Panel panelTop, panelHead, panelDetail;
        private Button btnExportRIR;

        public frmRIR()
        {
            InitializeComponent();
            BuildUI();
            LoadRIR();
            this.Resize += FrmRIR_Resize;
        }

        private void BuildUI()
        {
            this.Text = "Quản lý Phiếu Kiểm Tra Hàng Nhập (RIR)";
            this.BackColor = Color.FromArgb(245, 245, 245);

            // ===== PANEL TOP =====
            panelTop = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(1360, 210),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(panelTop);

            panelTop.Controls.Add(new Label
            {
                Text = "DANH SÁCH PHIẾU RIR",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 10),
                Size = new Size(400, 30)
            });

            txtSearch = new TextBox
            {
                Location = new Point(10, 48),
                Size = new Size(280, 28),
                Font = new Font("Segoe UI", 10),
                PlaceholderText = "Tìm RIR No, PO No, Workorder..."
            };
            panelTop.Controls.Add(txtSearch);

            txtSearch.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) BtnSearch_Click(null, null); };

            btnSearch = CreateBtn("🔍 Tìm", Color.FromArgb(0, 120, 212), new Point(300, 47), 80, 30);
            btnSearch.Click += BtnSearch_Click;
            panelTop.Controls.Add(btnSearch);

            btnNewRIR = CreateBtn("➕ Tạo RIR", Color.FromArgb(40, 167, 69), new Point(390, 47), 110, 30);
            btnNewRIR.Click += BtnNewRIR_Click;
            panelTop.Controls.Add(btnNewRIR);

            btnDeleteRIR = CreateBtn("🗑 Xóa RIR", Color.FromArgb(220, 53, 69), new Point(510, 47), 110, 30);
            btnDeleteRIR.Click += BtnDeleteRIR_Click;
            panelTop.Controls.Add(btnDeleteRIR);

            lblStatus = new Label
            {
                Location = new Point(635, 52),
                Size = new Size(500, 25),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.Gray
            };
            panelTop.Controls.Add(lblStatus);

            btnExportRIR = CreateBtn("🗑 Xóa RIR", Color.FromArgb(220, 53, 69), new Point(790, 47), 110, 30);
            btnExportRIR.Click += BtnExportRIR_Click;
            btnExportRIR.BringToFront();
            panelTop.Controls.Add(btnExportRIR);

            dgvRIR = new DataGridView
            {
                Location = new Point(10, 85),
                Size = new Size(1335, 115),
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
            dgvRIR.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvRIR.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvRIR.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvRIR.EnableHeadersVisualStyles = false;
            dgvRIR.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);

            dgvRIR.SelectionChanged += DgvRIR_SelectionChanged;
            dgvRIR.CellFormatting += DgvRIR_CellFormatting;
            panelTop.Controls.Add(dgvRIR);

            // ===== PANEL HEAD =====
            panelHead = new Panel
            {
                Location = new Point(10, 230),
                Size = new Size(1360, 175),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(panelHead);

            panelHead.Controls.Add(new Label
            {
                Text = "THÔNG TIN PHIẾU RIR",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 8),
                Size = new Size(300, 25)
            });

            // Row 1
            int y = 38;
            AddLbl(panelHead, "RIR No (*):", 10, y);
            txtRIRNo = AddTb(panelHead, 120, y, 150);

            AddLbl(panelHead, "Ngày phát hành:", 270, y);
            dtpIssueDate = new DateTimePicker
            {
                Location = new Point(390, y),
                Size = new Size(140, 25),
                Font = new Font("Segoe UI", 9),
                Format = DateTimePickerFormat.Short
            };
            panelHead.Controls.Add(dtpIssueDate);

            AddLbl(panelHead, "Tên dự án:", 500, y);
            txtProjectName = AddTb(panelHead, 625, y, 220);

            AddLbl(panelHead, "Trạng thái:", 860, y);
            cboStatus = new ComboBox
            {
                Location = new Point(940, y),
                Size = new Size(180, 25),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboStatus.Items.AddRange(new[] { "Chờ kiểm tra", "Đang kiểm tra", "Hoàn thành", "Từ chối" });
            cboStatus.SelectedIndex = 0;
            panelHead.Controls.Add(cboStatus);

            // Row 2
            y += 38;
            AddLbl(panelHead, "Workorder No:", 10, y);
            txtWorkorderNo = AddTb(panelHead, 120, y, 150);

            AddLbl(panelHead, "MPR No:", 270, y);
            txtMPRNo = AddTb(panelHead, 400, y, 140);

            AddLbl(panelHead, "PO No:", 545, y);
            txtPONo = AddTb(panelHead, 625, y, 140);

            AddLbl(panelHead, "Khách hàng:", 780, y);
            txtCustomer = AddTb(panelHead, 860, y, 260);
            txtCustomer.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            // Buttons Row
            y += 45;
            btnSaveHead = CreateBtn("💾 Lưu Header", Color.FromArgb(0, 120, 212), new Point(10, y), 130, 32);
            btnSaveHead.Click += BtnSaveHead_Click;
            panelHead.Controls.Add(btnSaveHead);

            var btnClear = CreateBtn("🔄 Làm mới", Color.FromArgb(108, 117, 125), new Point(150, y), 110, 32);
            btnClear.Click += (s, e) => { _selectedRIR_ID = 0; ClearHead(); dgvDetails.Rows.Clear(); };
            panelHead.Controls.Add(btnClear);

            // ===== PANEL DETAIL =====
            panelDetail = new Panel
            {
                Location = new Point(10, 415),
                Size = new Size(1360, 355),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            this.Controls.Add(panelDetail);

            panelDetail.Controls.Add(new Label
            {
                Text = "CHI TIẾT VẬT TƯ KIỂM TRA",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 8),
                Size = new Size(400, 25)
            });

            // Đổi nút Import PO thành Import PNK (Kho)
            btnImportPO = CreateBtn("📦 Import từ Phiếu Nhập Kho", Color.FromArgb(255, 140, 0), new Point(10, 38), 210, 30);
            btnImportPO.Click += BtnImportPO_Click;
            panelDetail.Controls.Add(btnImportPO);

            btnAddDetail = CreateBtn("➕ Thêm dòng", Color.FromArgb(40, 167, 69), new Point(230, 38), 120, 30);
            btnAddDetail.Click += BtnAddDetail_Click;
            panelDetail.Controls.Add(btnAddDetail);

            btnDeleteDetail = CreateBtn("🗑 Xóa dòng", Color.FromArgb(220, 53, 69), new Point(360, 38), 110, 30);
            btnDeleteDetail.Click += BtnDeleteDetail_Click;
            panelDetail.Controls.Add(btnDeleteDetail);

            btnSaveDetail = CreateBtn("💾 Lưu chi tiết", Color.FromArgb(0, 120, 212), new Point(480, 38), 130, 30);
            btnSaveDetail.Click += BtnSaveDetail_Click;
            panelDetail.Controls.Add(btnSaveDetail);

            dgvDetails = new DataGridView
            {
                Location = new Point(10, 75),
                Size = new Size(1335, 270),
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
            dgvDetails.CellFormatting += DgvDetails_CellFormatting;
            BuildDetailColumns();
            panelDetail.Controls.Add(dgvDetails);

            // Đưa tất cả TextBox và ComboBox lên trên Label
            foreach (Panel panel in new[] { panelHead, panelTop, panelDetail })
            {
                foreach (Control c in panel.Controls)
                {
                    if (c is TextBox || c is ComboBox || c is DateTimePicker || c is NumericUpDown)
                        c.BringToFront();
                }
            }
        }

        private void BtnExportRIR_Click(object? sender, EventArgs e)
        {
            PrintRIR();
        }

        private void BuildDetailColumns()
        {
            dgvDetails.Columns.Clear();
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "RIR_Detail_ID", HeaderText = "ID", Visible = false });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Item_No", HeaderText = "STT", Width = 45, ReadOnly = true });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Item_Name", HeaderText = "Tên vật tư", Width = 200 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Material", HeaderText = "Vật liệu", Width = 90 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Size", HeaderText = "Kích thước", Width = 110 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "UNIT", HeaderText = "ĐVT", Width = 55 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Qty_Required", HeaderText = "SL Yêu cầu", Width = 80 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Qty_Received", HeaderText = "SL Thực nhận", Width = 85 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "MTRno", HeaderText = "MTR No", Width = 100 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Heatno", HeaderText = "Heat No", Width = 90 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "ID_Code", HeaderText = "ID Code", Width = 100 });

            var cboResult = new DataGridViewComboBoxColumn
            {
                Name = "Inspect_Result",
                HeaderText = "Kết quả KT",
                Width = 100,
                FlatStyle = FlatStyle.Flat
            };
            cboResult.Items.AddRange(new[] { "", "Pass", "Fail", "Hold" });
            dgvDetails.Columns.Add(cboResult);

            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Remarks", HeaderText = "Ghi chú", FillWeight = 100 });
        }

        private void AddLbl(Panel p, string text, int x, int y)
        {
            p.Controls.Add(new Label
            {
                Text = text,
                Location = new Point(x, y + 3),
                Size = new Size(110, 20),
                Font = new Font("Segoe UI", 9)
            });
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

        // ===== RESIZE =====
        private void FrmRIR_Resize(object sender, EventArgs e)
        {
            try
            {
                int w = this.ClientSize.Width - 20;
                int h = this.ClientSize.Height;

                panelTop.Width = w;
                panelHead.Width = w;
                panelDetail.Width = w;
                panelDetail.Height = h - panelDetail.Top - 10;

                dgvRIR.Width = panelTop.Width - 20;
                dgvDetails.Width = panelDetail.Width - 20;
                dgvDetails.Height = panelDetail.Height - 85;

                if (txtCustomer != null && panelHead != null)
                    txtCustomer.Width = panelHead.Width - txtCustomer.Left - 20;
            }
            catch { }
        }

        // ===== LOAD RIR =====
        private void LoadRIR()
        {
            try
            {
                _rirList = _service.GetAll();
                BindRIRGrid(_rirList);
                lblStatus.Text = $"Tổng: {_rirList.Count} phiếu RIR";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải RIR: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BindRIRGrid(List<RIRHead> list)
        {
            dgvRIR.DataSource = list.ConvertAll(r => new
            {
                ID = r.RIR_ID,
                RIR_No = r.RIR_No,
                Ngay_PH = r.Issue_Date.HasValue ? r.Issue_Date.Value.ToString("dd/MM/yyyy") : "",
                Ten_Du_An = r.Project_Name,
                Workorder = r.WorkorderNo,
                MPR_No = r.MPR_No,
                PO_No = r.PONo,
                Khach_Hang = r.Customer,
                Trang_Thai = r.Status,
                Ngay_Tao = r.Created_Date.HasValue ? r.Created_Date.Value.ToString("dd/MM/yyyy") : ""
            });

            if (dgvRIR.Columns.Contains("ID"))
                dgvRIR.Columns["ID"].Visible = false;
        }

        // ===== LOAD DETAILS =====
        private void LoadDetails(int rirId)
        {
            try
            {
                _details = _service.GetDetails(rirId);
                dgvDetails.Rows.Clear();

                foreach (var d in _details)
                {
                    int idx = dgvDetails.Rows.Add();
                    var row = dgvDetails.Rows[idx];

                    row.Cells["RIR_Detail_ID"].Value = d.RIR_Detail_ID;
                    row.Cells["Item_No"].Value = d.Item_No;
                    row.Cells["Item_Name"].Value = d.Item_Name;
                    row.Cells["Material"].Value = d.Material;
                    row.Cells["Size"].Value = d.Size;
                    row.Cells["UNIT"].Value = d.UNIT;
                    row.Cells["Qty_Required"].Value = d.Qty_Required;
                    row.Cells["Qty_Received"].Value = d.Qty_Received;
                    row.Cells["MTRno"].Value = d.MTRno;
                    row.Cells["Heatno"].Value = d.Heatno;
                    row.Cells["ID_Code"].Value = d.ID_Code;
                    row.Cells["Inspect_Result"].Value = d.Inspect_Result;
                    row.Cells["Remarks"].Value = d.Remarks ?? "";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải chi tiết: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ===== SỰ KIỆN =====
        private void DgvRIR_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvRIR.SelectedRows.Count == 0) return;
            var row = dgvRIR.SelectedRows[0];
            _selectedRIR_ID = Convert.ToInt32(row.Cells["ID"].Value);

            var h = _rirList.Find(x => x.RIR_ID == _selectedRIR_ID);
            if (h == null) return;

            txtRIRNo.Text = h.RIR_No;
            txtProjectName.Text = h.Project_Name;
            txtWorkorderNo.Text = h.WorkorderNo;
            txtMPRNo.Text = h.MPR_No;
            txtPONo.Text = h.PONo;
            txtCustomer.Text = h.Customer;
            dtpIssueDate.Value = h.Issue_Date ?? DateTime.Today;

            int idx = cboStatus.Items.IndexOf(h.Status);
            cboStatus.SelectedIndex = idx >= 0 ? idx : 0;

            LoadDetails(_selectedRIR_ID);
        }

        private void DgvRIR_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;

            if (dgvRIR.Columns[e.ColumnIndex].Name == "Trang_Thai")
            {
                string val = e.Value?.ToString() ?? "";
                e.CellStyle.ForeColor =
                    val == "Hoàn thành" ? Color.FromArgb(40, 167, 69) :
                    val == "Đang kiểm tra" ? Color.FromArgb(0, 120, 212) :
                    val == "Từ chối" ? Color.FromArgb(220, 53, 69) :
                                              Color.FromArgb(255, 140, 0);
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
        }

        private void DgvDetails_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;

            if (dgvDetails.Columns[e.ColumnIndex].Name == "Inspect_Result")
            {
                string val = e.Value?.ToString() ?? "";
                e.CellStyle.ForeColor =
                    val == "Pass" ? Color.FromArgb(40, 167, 69) :
                    val == "Fail" ? Color.FromArgb(220, 53, 69) :
                    val == "Hold" ? Color.FromArgb(255, 140, 0) :
                                    Color.Black;

                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
        }

        private void BtnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                string kw = txtSearch.Text.Trim();
                _rirList = string.IsNullOrEmpty(kw) ? _service.GetAll() : _service.Search(kw);
                BindRIRGrid(_rirList);
                lblStatus.Text = $"Tìm thấy: {_rirList.Count} phiếu";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnNewRIR_Click(object sender, EventArgs e)
        {
            _selectedRIR_ID = 0;
            ClearHead();
            dgvDetails.Rows.Clear();
            _details.Clear();
            txtRIRNo.Focus();
        }

        private void BtnSaveHead_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtRIRNo.Text))
            {
                MessageBox.Show("Vui lòng nhập RIR No!", "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtRIRNo.Focus();
                return;
            }
            try
            {
                var h = new RIRHead
                {
                    RIR_ID = _selectedRIR_ID,
                    RIR_No = txtRIRNo.Text.Trim(),
                    Issue_Date = dtpIssueDate.Value,
                    Project_Name = txtProjectName.Text.Trim(),
                    WorkorderNo = txtWorkorderNo.Text.Trim(),
                    MPR_No = txtMPRNo.Text.Trim(),
                    PONo = txtPONo.Text.Trim(),
                    Customer = txtCustomer.Text.Trim(),
                    Status = cboStatus.SelectedItem?.ToString() ?? "Chờ kiểm tra"
                };

                if (_selectedRIR_ID == 0)
                {
                    _selectedRIR_ID = _service.InsertHead(h, _currentUser);
                    MessageBox.Show("Tạo phiếu RIR thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    _service.UpdateHead(h, _currentUser);
                    MessageBox.Show("Cập nhật phiếu RIR thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                LoadRIR();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi lưu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnDeleteRIR_Click(object sender, EventArgs e)
        {
            if (_selectedRIR_ID == 0)
            {
                MessageBox.Show("Vui lòng chọn phiếu RIR cần xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show("Xóa phiếu RIR và toàn bộ chi tiết?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    _service.DeleteHead(_selectedRIR_ID);
                    MessageBox.Show("Xóa thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    _selectedRIR_ID = 0;
                    ClearHead();
                    dgvDetails.Rows.Clear();
                    LoadRIR();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // =========================================================================
        // ĐÃ CẬP NHẬT: IMPORT TỪ PHIẾU NHẬP KHO (WAREHOUSE_IMPORT)
        // =========================================================================
        private void BtnImportPO_Click(object sender, EventArgs e)
        {
            try
            {
                using (var dlg = new Form())
                {
                    dlg.Text = "Chọn Phiếu Nhập Kho (PNK) để tạo RIR";
                    dlg.Size = new Size(1000, 480);
                    dlg.StartPosition = FormStartPosition.CenterParent;
                    dlg.BackColor = Color.White;

                    dlg.Controls.Add(new Label
                    {
                        Text = "Danh sách Phiếu nhập kho đã hoàn tất:",
                        Font = new Font("Segoe UI", 10, FontStyle.Bold),
                        ForeColor = Color.FromArgb(0, 120, 212),
                        Location = new Point(10, 10),
                        Size = new Size(500, 25)
                    });

                    // TÌM KIẾM
                    var txtSearchPNK = new TextBox
                    {
                        Location = new Point(10, 42),
                        Size = new Size(300, 25),
                        Font = new Font("Segoe UI", 9),
                        PlaceholderText = "Tìm theo mã PNK, PO No, Dự án..."
                    };
                    dlg.Controls.Add(txtSearchPNK);

                    var btnFilter = new Button
                    {
                        Text = "🔍 Lọc",
                        Location = new Point(320, 41),
                        Size = new Size(80, 28),
                        BackColor = Color.FromArgb(0, 120, 212),
                        ForeColor = Color.White,
                        FlatStyle = FlatStyle.Flat,
                        Font = new Font("Segoe UI", 9, FontStyle.Bold)
                    };
                    btnFilter.FlatAppearance.BorderSize = 0;
                    dlg.Controls.Add(btnFilter);

                    // GRID DANH SÁCH PNK
                    var dgvPNK = new DataGridView
                    {
                        Location = new Point(10, 78),
                        Size = new Size(960, 300),
                        ReadOnly = true,
                        AllowUserToAddRows = false,
                        SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                        BackgroundColor = Color.White,
                        BorderStyle = BorderStyle.FixedSingle,
                        RowHeadersVisible = false,
                        Font = new Font("Segoe UI", 9),
                        AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                    };
                    dgvPNK.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
                    dgvPNK.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                    dgvPNK.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    dgvPNK.EnableHeadersVisualStyles = false;
                    dgvPNK.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
                    dlg.Controls.Add(dgvPNK);

                    // HÀM TẢI DANH SÁCH PHIẾU NHẬP KHO TRỰC TIẾP TỪ CSDL
                    Action loadPNK = () =>
                    {
                        string kw = txtSearchPNK.Text.Trim();
                        string sql = @"
                            SELECT 
                                wi.Import_No AS [Mã phiếu], 
                                MAX(wi.Import_Date) AS [Ngày nhập], 
                                ph.PONo AS [PO No], 
                                MAX(ph.Project_Name) AS [Dự án],
                                MAX(ph.WorkorderNo) AS [Workorder],
                                MAX(ph.MPR_No) AS [MPR No],
                                COUNT(wi.Import_ID) AS [Số vật tư]
                            FROM Warehouse_Import wi
                            LEFT JOIN PO_head ph ON wi.PO_ID = ph.PO_ID
                            WHERE wi.Import_No LIKE N'%' + @kw + '%' 
                               OR ph.PONo LIKE N'%' + @kw + '%'
                               OR ph.Project_Name LIKE N'%' + @kw + '%'
                            GROUP BY wi.Import_No, ph.PONo
                            ORDER BY MAX(wi.Import_Date) DESC";

                        using (var conn = DatabaseHelper.GetConnection())
                        {
                            conn.Open();
                            var cmd = new SqlCommand(sql, conn);
                            cmd.Parameters.AddWithValue("@kw", kw);
                            var dt = new DataTable();
                            dt.Load(cmd.ExecuteReader());
                            dgvPNK.DataSource = dt;

                            // Cân chỉnh format ngày tháng cho đẹp
                            if (dgvPNK.Columns.Contains("Ngày nhập"))
                            {
                                dgvPNK.Columns["Ngày nhập"].DefaultCellStyle.Format = "dd/MM/yyyy HH:mm";
                            }
                        }
                    };

                    // Gọi tải dữ liệu lần đầu
                    loadPNK();
                    btnFilter.Click += (s2, e2) => loadPNK();
                    txtSearchPNK.KeyDown += (s2, e2) => { if (e2.KeyCode == Keys.Enter) loadPNK(); };

                    // BUTTONS CHỌN/HỦY
                    var btnChon = new Button
                    {
                        Text = "✔ Chọn phiếu này",
                        Location = new Point(10, 390),
                        Size = new Size(160, 32),
                        BackColor = Color.FromArgb(40, 167, 69),
                        ForeColor = Color.White,
                        FlatStyle = FlatStyle.Flat,
                        Font = new Font("Segoe UI", 9, FontStyle.Bold),
                        DialogResult = DialogResult.OK
                    };
                    btnChon.FlatAppearance.BorderSize = 0;
                    dlg.Controls.Add(btnChon);

                    var btnHuy = new Button
                    {
                        Text = "Hủy",
                        Location = new Point(180, 390),
                        Size = new Size(80, 32),
                        BackColor = Color.FromArgb(108, 117, 125),
                        ForeColor = Color.White,
                        FlatStyle = FlatStyle.Flat,
                        Font = new Font("Segoe UI", 9, FontStyle.Bold),
                        DialogResult = DialogResult.Cancel
                    };
                    btnHuy.FlatAppearance.BorderSize = 0;
                    dlg.Controls.Add(btnHuy);

                    dlg.AcceptButton = btnChon;
                    dlg.CancelButton = btnHuy;

                    // KIỂM TRA ĐIỀU KIỆN SAU KHI ĐÓNG FORM
                    if (dlg.ShowDialog() != DialogResult.OK) return;
                    if (dgvPNK.SelectedRows.Count == 0) return;

                    var selRow = dgvPNK.SelectedRows[0];
                    string pnkNo = selRow.Cells["Mã phiếu"].Value.ToString();
                    string poNo = selRow.Cells["PO No"].Value.ToString();
                    string projectName = selRow.Cells["Dự án"].Value.ToString();
                    string woNo = selRow.Cells["Workorder"].Value.ToString();
                    string mprNo = selRow.Cells["MPR No"].Value.ToString();

                    // TẠO MÃ RIR TỰ ĐỘNG
                    string autoRIRNo = GenerateRIRNo(poNo, woNo);

                    // ĐIỀN HEADER RIR
                    txtRIRNo.Text = autoRIRNo;
                    txtPONo.Text = poNo;
                    txtMPRNo.Text = mprNo;
                    txtProjectName.Text = projectName;
                    txtWorkorderNo.Text = woNo;
                    dtpIssueDate.Value = DateTime.Today;
                    cboStatus.SelectedIndex = 0;

                    // Lấy tên khách hàng từ Dự án
                    try
                    {
                        var projects = new ProjectService().GetAll();
                        var proj2 = projects.Find(p => p.WorkorderNo == woNo);
                        txtCustomer.Text = proj2?.Customer ?? "";
                    }
                    catch { txtCustomer.Text = ""; }

                    // TẢI CHI TIẾT TỪ WAREHOUSE_IMPORT VÀO LƯỚI CHI TIẾT
                    string sqlDetails = @"
                        SELECT 
                            wi.Item_Name, 
                            wi.Material, 
                            wi.Size, 
                            wi.UNIT, 
                            wi.Qty_Import, 
                            wi.ID_Code, 
                            ISNULL(wi.MTRno, '') AS MTRno 
                        FROM Warehouse_Import wi
                        WHERE wi.Import_No = @pnkNo";

                    int countItems = 0;
                    using (var conn = DatabaseHelper.GetConnection())
                    {
                        conn.Open();
                        var cmd = new SqlCommand(sqlDetails, conn);
                        cmd.Parameters.AddWithValue("@pnkNo", pnkNo);
                        using (var reader = cmd.ExecuteReader())
                        {
                            dgvDetails.Rows.Clear();
                            int itemNo = 1;
                            while (reader.Read())
                            {
                                int idx = dgvDetails.Rows.Add();
                                var row = dgvDetails.Rows[idx];
                                row.Cells["RIR_Detail_ID"].Value = 0;
                                row.Cells["Item_No"].Value = itemNo++;
                                row.Cells["Item_Name"].Value = reader["Item_Name"]?.ToString() ?? "";
                                row.Cells["Material"].Value = reader["Material"]?.ToString() ?? "";
                                row.Cells["Size"].Value = reader["Size"]?.ToString() ?? "";
                                row.Cells["UNIT"].Value = reader["UNIT"]?.ToString() ?? "";

                                // Gán SL nhập kho thực tế thành SL Yêu cầu kiểm tra
                                row.Cells["Qty_Required"].Value = reader["Qty_Import"] != DBNull.Value ? Convert.ToDecimal(reader["Qty_Import"]) : 0;
                                row.Cells["Qty_Received"].Value = 0; // Để QC tự nhập kết quả

                                row.Cells["MTRno"].Value = reader["MTRno"]?.ToString() ?? "";
                                row.Cells["Heatno"].Value = "";
                                row.Cells["ID_Code"].Value = reader["ID_Code"]?.ToString() ?? "";
                                row.Cells["Inspect_Result"].Value = "";
                                row.Cells["Remarks"].Value = "";
                                countItems++;
                            }
                        }
                    }

                    if (countItems == 0)
                    {
                        MessageBox.Show("Phiếu nhập kho này không có chi tiết vật tư!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        MessageBox.Show(
                            $"✅ Đã kéo dữ liệu từ phiếu nhập: {pnkNo}\n" +
                            $"Mã RIR tạo mới: {autoRIRNo}\n" +
                            $"Số lượng vật tư: {countItems} mục\n\n" +
                            $"Nhấn 'Lưu Header' và 'Lưu chi tiết' để hoàn tất.",
                            "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi trong quá trình Import: " + ex.Message, "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GenerateRIRNo(string poNo, string workorderNo = "")
        {
            try
            {
                // Lấy mã PO từ ProjectInfo theo WorkorderNo
                string projectPOCode = "";
                if (!string.IsNullOrEmpty(workorderNo))
                {
                    try
                    {
                        var projects = new ProjectService().GetAll();
                        var proj = projects.Find(p => p.WorkorderNo == workorderNo);
                        projectPOCode = proj?.POCode ?? "";
                    }
                    catch { }
                }

                // Nếu không lấy được từ Project thì dùng PONo
                string baseCode = string.IsNullOrEmpty(projectPOCode) ? poNo : projectPOCode;
                string prefix = $"DV-{baseCode}-RIR-";

                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var cmd = new Microsoft.Data.SqlClient.SqlCommand(
                        "SELECT COUNT(*) FROM RIR_head WHERE RIR_No LIKE @prefix", conn);
                    cmd.Parameters.AddWithValue("@prefix", prefix + "%");
                    int count = Convert.ToInt32(cmd.ExecuteScalar());
                    return $"{prefix}{count + 1:D3}";
                }
            }
            catch { return $"DV-{poNo}-RIR-001"; }
        }

        private void BtnAddDetail_Click(object sender, EventArgs e)
        {
            if (_selectedRIR_ID == 0)
            {
                MessageBox.Show("Vui lòng lưu header RIR trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            int nextNo = dgvDetails.Rows.Count + 1;
            int idx = dgvDetails.Rows.Add();
            var row = dgvDetails.Rows[idx];
            row.Cells["RIR_Detail_ID"].Value = 0;
            row.Cells["Item_No"].Value = nextNo;
            row.Cells["Item_Name"].Value = "";
            row.Cells["Material"].Value = "";
            row.Cells["Size"].Value = "";
            row.Cells["UNIT"].Value = "cái";
            row.Cells["Qty_Required"].Value = 0;
            row.Cells["Qty_Received"].Value = 0;
            row.Cells["MTRno"].Value = "";
            row.Cells["Heatno"].Value = "";
            row.Cells["ID_Code"].Value = "";
            row.Cells["Inspect_Result"].Value = "";
            row.Cells["Remarks"].Value = "";
            dgvDetails.CurrentCell = dgvDetails.Rows[idx].Cells["Item_Name"];
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
            int detailId = Convert.ToInt32(row.Cells["RIR_Detail_ID"].Value ?? 0);
            if (detailId > 0)
            {
                if (MessageBox.Show("Xóa dòng này?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    try
                    {
                        _service.DeleteDetail(detailId);
                        dgvDetails.Rows.Remove(row);
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
            if (_selectedRIR_ID == 0)
            {
                MessageBox.Show("Vui lòng lưu header RIR trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

                    var d = new RIRDetail
                    {
                        RIR_Detail_ID = Convert.ToInt32(row.Cells["RIR_Detail_ID"].Value ?? 0),
                        RIR_ID = _selectedRIR_ID,
                        Item_No = Convert.ToInt32(row.Cells["Item_No"].Value ?? 0),
                        Item_Name = itemName,
                        Material = row.Cells["Material"].Value?.ToString() ?? "",
                        Size = row.Cells["Size"].Value?.ToString() ?? "",
                        UNIT = row.Cells["UNIT"].Value?.ToString() ?? "",
                        Qty_Required = IntVal(row.Cells["Qty_Required"].Value),
                        Qty_Received = IntVal(row.Cells["Qty_Received"].Value),
                        MTRno = row.Cells["MTRno"].Value?.ToString() ?? "",
                        Heatno = row.Cells["Heatno"].Value?.ToString() ?? "",
                        ID_Code = row.Cells["ID_Code"].Value?.ToString() ?? "",
                        Inspect_Result = row.Cells["Inspect_Result"].Value?.ToString() ?? "",
                        Remarks = row.Cells["Remarks"].Value?.ToString() ?? ""
                    };

                    if (d.RIR_Detail_ID == 0)
                        _service.InsertDetail(d, _currentUser);
                    else
                        _service.UpdateDetail(d);

                    saved++;
                }
                MessageBox.Show($"Đã lưu {saved} dòng thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDetails(_selectedRIR_ID);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi lưu chi tiết: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ===== HELPERS =====
        private int IntVal(object val)
        {
            if (val == null || val == DBNull.Value) return 0;
            return int.TryParse(val.ToString(), out int i) ? i : 0;
        }

        private void ClearHead()
        {
            txtRIRNo.Text = "";
            txtProjectName.Text = "";
            txtWorkorderNo.Text = "";
            txtMPRNo.Text = "";
            txtPONo.Text = "";
            txtCustomer.Text = "";
            dtpIssueDate.Value = DateTime.Today;
            cboStatus.SelectedIndex = 0;
        }

        private async void PrintRIR()
        {
            if (dgvRIR.Rows.Count <= 0) return;
            int rsl = dgvRIR.CurrentRow.Index;
            var rirId = int.Parse(dgvRIR.Rows[rsl].Cells[0].Value.ToString());
            var rirNO = (dgvRIR.Rows[rsl].Cells[1].Value.ToString());
            var poNO = dgvRIR.Rows[rsl].Cells[6].Value.ToString();

            // lấy data PO
            var poModel = await _poService.GetPOAsync(poNO);

            // Lấy data project
            var projectMode = _projectServices.GetByProjectCode(poModel.ProjectCode);

            // Lấy data supplier
            var supplierModel = _supplierServices.GetBySupId(poModel.Supplier_ID);

            // Lấy data RIR DEtail
            var dtImports = await _service.GetDetailsToExport(rirId);
            PrintBill(dtImports, projectMode, rirNO, poModel);
        }

        public void PrintBill(DataTable dtDetails, ProjectInfo projects, string RIRNo, POHead po)
        {
            try
            {
                if (dgvRIR.CurrentRow == null)
                {
                    MessageBox.Show("Vui lòng chọn một phiếu nhập kho để in!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "rir_template.xlsx");
                string exportFolder = projects.PNK_Link;
                if (!Directory.Exists(exportFolder))
                {
                    Directory.CreateDirectory(exportFolder);
                }

                string fileName = $"{RIRNo}_{DateTime.Now:ddMMyyyy}.xlsx";
                string actualSavePath = Path.Combine(exportFolder, fileName);

                if (!File.Exists(templatePath))
                {
                    MessageBox.Show("Không tìm thấy file template tại: " + templatePath, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 2. Xử lý Excel
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                FileInfo templateFile = new FileInfo(templatePath);
                FileInfo newFile = new FileInfo(actualSavePath);
                using (ExcelPackage package = new ExcelPackage(newFile, templateFile))
                {
                    ExcelWorksheet ws = package.Workbook.Worksheets[0];

                    // --- ĐIỀN THÔNG TIN HEADER (Dựa trên các Placeholder trong file của bạn) ---
                    // Lưu ý: Tùy vào vị trí ô trong file xlsx thực tế, bạn có thể dùng Replace hoặc gán trực tiếp

                    // Tìm và thay thế các từ khóa trong vùng Header (giả định từ dòng 1 đến dòng 12)
                    var headerCells = ws.Cells["A1:AW12"];
                    foreach (var cell in headerCells)
                    {
                        if (cell.Value == null) continue;
                        string txt = cell.Value.ToString();

                        if (txt.Contains("<<RIR-NO>>")) cell.Value = txt.Replace("<<RIR-NO>>", RIRNo);
                        if (txt.Contains("<<DATE>>")) cell.Value = txt.Replace("<<DATE>>", DateTime.Now.ToString("dd/MM/yyyy"));
                        if (txt.Contains("<<WO-NO>>")) cell.Value = txt.Replace("<<WO-NO>>", projects.WorkorderNo ?? "");
                        if (txt.Contains("<<MPR-NO>>")) cell.Value = txt.Replace("<<MPR-NO>>", po.MPR_No ?? "");
                        if (txt.Contains("<<PROJECT-NAME>>")) cell.Value = txt.Replace("<<PROJECT-NAME>>", projects.ProjectName ?? "");
                        if (txt.Contains("<<MPS-NO>>")) cell.Value = txt.Replace("<<MPS-NO>>", "");
                        if (txt.Contains("<<CLIENT>>")) cell.Value = txt.Replace("<<CLIENT>>", projects.Customer ?? "");
                        if (txt.Contains("<<PO-NO>>")) cell.Value = txt.Replace("<<PO-NO>>", po.PONo ?? "");
                        if (txt.Contains("<<USER-CREATE>>")) cell.Value = txt.Replace("<<USER-CREATE>>", _currentUser);
                    }

                    // --- ĐIỀN CHI TIẾT VẬT TƯ (Bắt đầu từ sau dòng tiêu đề STT) ---
                    // Theo file CSV bạn gửi, dòng tiêu đề cột bắt đầu khoảng dòng 13-14
                    int startRow = 7;

                    int count = dtDetails.Rows.Count;
                    if (count > 1)
                    {
                        // Chèn thêm dòng nếu danh sách vật tư nhiều hơn 1 (để giữ định dạng Footer bên dưới)
                        ws.InsertRow(startRow + 1, count - 1);
                        for (int i = 1; i < count; i++)
                        {
                            ws.Cells[startRow, 1, startRow, 30].Copy(ws.Cells[startRow + i, 1]);
                        }
                    }

                    for (int i = 0; i < count; i++)
                    {
                        DataRow dr = dtDetails.Rows[i];
                        int curr = startRow + i;

                        ws.Row(curr).Height = 25; // Độ cao dòng chuẩn cho RIR

                        // Điền dữ liệu theo cấu trúc template RIR
                        ws.Cells[curr, 1].Value = i + 1;                             // No.
                        ws.Cells[curr, 3].Value = dr["item_name"];                  // Item Name
                        ws.Cells[curr, 10].Value = dr["Material"];                  // Material Spec
                        ws.Cells[curr, 18].Value = dr["Size"];                      // Size
                        ws.Cells[curr, 25].Value = dr["UNIT"];                      // Unit
                        ws.Cells[curr, 27].Value = dr["Qty_Per_Sheet"];                // Qty
                        ws.Cells[curr, 29].Value = dr["MTRno"];                     // MTR No.
                        ws.Cells[curr, 35].Value = dr["Heatno"];                    // Heat/Lot No.
                        ws.Cells[curr, 41].Value = "";                        // Result (Mặc định)
                        ws.Cells[curr, 46].Value =  "";                     // Remarks

                        //RIR_Detail_ID, RIR_ID, PO_Detail_ID, Item_No,
                        //   item_name, Material, Size, UNIT,
                        //   Qty_Per_Sheet, MTRno, Heatno, Created_Date
                        // Căn lề và Border
                        using (var range = ws.Cells[curr, 1, curr, 50])
                        {
                            range.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            range.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        }

                        // Cột Tên và Size nên căn Trái
                        ws.Cells[curr, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        ws.Cells[curr, 18].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                        if (i > 0)
                        {
                            for (int col = 1; col <= 16; col++)
                            {
                                ws.Cells[curr, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                ws.Cells[curr, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                ws.Cells[curr, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                ws.Cells[curr, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                ws.Cells[curr, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                ws.Cells[curr, col].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                ws.Cells[curr, col].Style.Font.Name = "Times New Roman";
                                ws.Cells[curr, col].Style.Font.Size = 9;
                                ws.Cells[curr, col].Style.Font.Italic = false;
                            }
                        }
                    }

                    // Tự động lưu
                    package.Save();
                }

                var result = MessageBox.Show(
                    $"✅ Xuất phiếu nhập kho thành công!\nFile: {actualSavePath}\n\nBạn có muốn mở file ngay không?",
                    "Thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (result == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = actualSavePath,
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi in phiếu: " + ex.Message, "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}