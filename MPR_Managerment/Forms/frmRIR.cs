using MPR_Managerment.Helpers;
using MPR_Managerment.Models;
using MPR_Managerment.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
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
        private string _currentUser = AppSession.CurrentUser.Full_Name ?? "Admin";

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

        // ===== RIR LINK FOLDER BROWSER (góc phải panelDetail) =====
        private ComboBox cboProjectRIR;         // Chọn dự án theo ProjectCode
        private DataGridView dgvRIRFolders;     // Danh sách thư mục trong RIR_Link
        private DataTable _projectTable;        // DataTable: ProjectCode | RIR_Link
        private bool _isProjectSearching = false;

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

            btnExportRIR = CreateBtn("🖨 In RIR", Color.FromArgb(220, 53, 69), new Point(1200, 47), 110, 30);
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

            // Hướng dẫn paste
            panelDetail.Controls.Add(new Label
            {
                Text = "💡 Ctrl+V để dán dữ liệu từ Excel",
                Font = new Font("Segoe UI", 8, FontStyle.Italic),
                ForeColor = Color.FromArgb(102, 102, 102),
                Location = new Point(740, 12),
                Size = new Size(230, 18)
            });

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

            // =====================================================================
            // RIR LINK FOLDER BROWSER — Góc phải trên của panelDetail
            // =====================================================================
            const int folderPanelW = 340;
            const int folderPanelRight = 10; // khoảng cách từ mép phải panelDetail

            // Label tiêu đề
            var lblRIRFolder = new Label
            {
                Text = "📁 Thư mục RIR theo dự án",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(102, 51, 153),
                Size = new Size(folderPanelW, 18),
                // Location sẽ được tính lại trong Resize
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Name = "lblRIRFolderTitle"
            };
            panelDetail.Controls.Add(lblRIRFolder);

            // ComboBox chọn dự án — có tìm kiếm nhanh
            cboProjectRIR = new ComboBox
            {
                Size = new Size(folderPanelW, 25),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteMode = AutoCompleteMode.None,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Name = "cboProjectRIR"
            };
            panelDetail.Controls.Add(cboProjectRIR);
            cboProjectRIR.BringToFront();

            // DataGridView danh sách thư mục
            dgvRIRFolders = new DataGridView
            {
                Size = new Size(folderPanelW, 260),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom,
                Name = "dgvRIRFolders",
                Cursor = Cursors.Hand
            };
            dgvRIRFolders.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(102, 51, 153);
            dgvRIRFolders.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvRIRFolders.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvRIRFolders.EnableHeadersVisualStyles = false;
            dgvRIRFolders.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 240, 255);

            // Thêm cột thư mục
            dgvRIRFolders.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "FolderPath",
                HeaderText = "FolderPath",
                Visible = false
            });
            dgvRIRFolders.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "FolderName",
                HeaderText = "📂  Tên thư mục",
                FillWeight = 100
            });

            // Double click → mở thư mục
            dgvRIRFolders.CellDoubleClick += DgvRIRFolders_CellDoubleClick;
            panelDetail.Controls.Add(dgvRIRFolders);

            // Nạp dữ liệu dự án vào DataTable và bind ComboBox
            LoadProjectTable();
            BindProjectRIRCombo(_projectTable);

            cboProjectRIR.TextChanged += CboProjectRIR_TextChanged;
            cboProjectRIR.SelectedIndexChanged += CboProjectRIR_SelectedIndexChanged;
            cboProjectRIR.KeyDown += CboProjectRIR_KeyDown;

            // Vị trí ban đầu (sẽ tính lại trong Resize)
            PositionRIRFolderControls();
            // =====================================================================

            dgvDetails = new DataGridView
            {
                Location = new Point(10, 75),
                Size = new Size(1335 - folderPanelW - 15, 270),
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

            // ── Đăng ký sự kiện Ctrl+V paste từ Excel ──
            dgvDetails.KeyDown += DgvDetails_KeyDown;

            BuildDetailColumns();
            panelDetail.Controls.Add(dgvDetails);

            // Đưa tất cả TextBox và ComboBox lên trên Label
            foreach (Panel panel in new[] { panelHead, panelTop, panelDetail })
                foreach (Control c in panel.Controls)
                    if (c is TextBox || c is ComboBox || c is DateTimePicker || c is NumericUpDown)
                        c.BringToFront();
        }

        // =====================================================================
        // TÍNH VỊ TRÍ CÁC CONTROL FOLDER BROWSER (gọi khi init và resize)
        // =====================================================================
        private void PositionRIRFolderControls()
        {
            if (panelDetail == null || cboProjectRIR == null) return;
            int pw = panelDetail.ClientSize.Width;
            const int folderPanelW = 340;
            int left = pw - folderPanelW - 10;
            if (left < 500) left = 500; // không che bảng chi tiết

            var lbl = panelDetail.Controls["lblRIRFolderTitle"] as Label;
            if (lbl != null) { lbl.Left = left; lbl.Top = 8; lbl.Width = folderPanelW; }

            cboProjectRIR.Left = left;
            cboProjectRIR.Top = 30;
            cboProjectRIR.Width = folderPanelW;

            dgvRIRFolders.Left = left;
            dgvRIRFolders.Top = 60;
            dgvRIRFolders.Width = folderPanelW;
            // Chiều cao sẽ cập nhật trong FrmRIR_Resize
        }

        // =====================================================================
        // COMBOBOX DỰ ÁN — DÙNG ProjectCode, TÌM KIẾM NHANH
        // DataTable có 4 cột: ProjectCode | RIR_Link | WorkorderNo | ProjectName
        // =====================================================================
        private void LoadProjectTable()
        {
            _projectTable = new DataTable();
            _projectTable.Columns.Add("ProjectCode", typeof(string));
            _projectTable.Columns.Add("RIR_Link", typeof(string));
            _projectTable.Columns.Add("WorkorderNo", typeof(string));
            _projectTable.Columns.Add("ProjectName", typeof(string));

            try
            {
                var all = _projectServices.GetAll();
                foreach (var p in all)
                    _projectTable.Rows.Add(
                        p.ProjectCode ?? "",
                        p.RIR_Link ?? "",
                        p.WorkorderNo ?? "",
                        p.ProjectName ?? "");
            }
            catch { }
        }

        private void BindProjectRIRCombo(DataTable dt)
        {
            _isProjectSearching = true;
            string cur = cboProjectRIR.Text;
            cboProjectRIR.DataSource = null;
            cboProjectRIR.DataSource = dt;
            cboProjectRIR.DisplayMember = "ProjectCode";
            cboProjectRIR.ValueMember = "RIR_Link";
            cboProjectRIR.Text = cur;
            _isProjectSearching = false;
        }

        private void CboProjectRIR_TextChanged(object sender, EventArgs e)
        {
            if (_isProjectSearching) return;
            string kw = cboProjectRIR.Text.Trim();

            if (string.IsNullOrEmpty(kw))
            {
                BindProjectRIRCombo(_projectTable);
                cboProjectRIR.DroppedDown = false;
                return;
            }

            // Lọc DataTable theo ProjectCode chứa từ khoá
            DataTable filtered = _projectTable.Clone();
            foreach (DataRow row in _projectTable.Rows)
            {
                string code = row["ProjectCode"]?.ToString() ?? "";
                if (code.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0)
                    filtered.ImportRow(row);
            }

            _isProjectSearching = true;
            cboProjectRIR.DataSource = null;
            cboProjectRIR.DataSource = filtered.Rows.Count > 0 ? filtered : _projectTable;
            cboProjectRIR.DisplayMember = "ProjectCode";
            cboProjectRIR.ValueMember = "RIR_Link";
            cboProjectRIR.Text = kw;
            cboProjectRIR.SelectionStart = kw.Length;
            cboProjectRIR.DroppedDown = true;
            _isProjectSearching = false;
        }

        private void CboProjectRIR_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_isProjectSearching) return;
            // ValueMember = "RIR_Link" nên SelectedValue chính là đường dẫn
            string rirLink = cboProjectRIR.SelectedValue?.ToString() ?? "";
            LoadRIRFolders(rirLink);
        }

        private void CboProjectRIR_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                _isProjectSearching = true;
                BindProjectRIRCombo(_projectTable);
                cboProjectRIR.Text = "";
                cboProjectRIR.DroppedDown = false;
                _isProjectSearching = false;
                dgvRIRFolders.Rows.Clear();
            }
        }

        // =====================================================================
        // LOAD DANH SÁCH THƯ MỤC TRONG RIR_LINK
        // =====================================================================
        private void LoadRIRFolders(string rirLink)
        {
            dgvRIRFolders.Rows.Clear();
            if (string.IsNullOrWhiteSpace(rirLink))
            {
                dgvRIRFolders.Rows.Add("", "⚠ Dự án chưa cấu hình RIR Link");
                dgvRIRFolders.Rows[0].DefaultCellStyle.ForeColor = Color.Gray;
                return;
            }
            if (!Directory.Exists(rirLink))
            {
                dgvRIRFolders.Rows.Add("", $"⚠ Không tìm thấy thư mục:\n{rirLink}");
                dgvRIRFolders.Rows[0].DefaultCellStyle.ForeColor = Color.FromArgb(220, 53, 69);
                return;
            }
            try
            {
                var dirs = Directory.GetDirectories(rirLink);
                if (dirs.Length == 0)
                {
                    dgvRIRFolders.Rows.Add("", "📭 Chưa có thư mục nào");
                    dgvRIRFolders.Rows[0].DefaultCellStyle.ForeColor = Color.Gray;
                    return;
                }
                foreach (string dir in dirs)
                {
                    string name = Path.GetFileName(dir);
                    int idx = dgvRIRFolders.Rows.Add(dir, "📂  " + name);
                    dgvRIRFolders.Rows[idx].DefaultCellStyle.ForeColor = Color.FromArgb(0, 80, 160);
                }
            }
            catch (Exception ex)
            {
                dgvRIRFolders.Rows.Add("", "⚠ Lỗi đọc thư mục: " + ex.Message);
                dgvRIRFolders.Rows[0].DefaultCellStyle.ForeColor = Color.FromArgb(220, 53, 69);
            }
        }

        // =====================================================================
        // DOUBLE CLICK → MỞ THƯ MỤC
        // =====================================================================
        private void DgvRIRFolders_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string path = dgvRIRFolders.Rows[e.RowIndex].Cells["FolderPath"].Value?.ToString() ?? "";
            if (string.IsNullOrEmpty(path)) return;
            if (!Directory.Exists(path))
            {
                MessageBox.Show("Thư mục không tồn tại:\n" + path, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = path,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể mở thư mục: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "PO_Detail_ID", HeaderText = "PO Detail No", Width = 100 }); // Add column PO_Detail_ID

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

        // =========================================================================
        // PASTE TỪ EXCEL — Ctrl+V vào bảng Chi tiết vật tư kiểm tra
        // =========================================================================
        private void DgvDetails_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V)
            {
                PasteFromExcel();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void PasteFromExcel()
        {
            try
            {
                string clipText = Clipboard.GetText();
                if (string.IsNullOrEmpty(clipText)) return;

                string[] lines = clipText.Split(
                    new[] { "\r\n", "\r", "\n" },
                    StringSplitOptions.RemoveEmptyEntries);
                if (lines.Length == 0) return;

                int startRowIndex = dgvDetails.CurrentCell?.RowIndex ?? dgvDetails.Rows.Count;
                int startColIndex = dgvDetails.CurrentCell?.ColumnIndex ?? 0;

                int addedRows = 0;

                foreach (string line in lines)
                {
                    string[] cells = line.Split('\t');

                    if (startRowIndex >= dgvDetails.Rows.Count)
                    {
                        int nextNo = dgvDetails.Rows.Count + 1;
                        int newIdx = dgvDetails.Rows.Add();
                        var newRow = dgvDetails.Rows[newIdx];
                        newRow.Cells["RIR_Detail_ID"].Value = 0;
                        newRow.Cells["Item_No"].Value = nextNo;
                        newRow.Cells["Qty_Required"].Value = 0;
                        newRow.Cells["Qty_Received"].Value = 0;
                        newRow.Cells["Inspect_Result"].Value = "";
                    }

                    var gridRow = dgvDetails.Rows[startRowIndex];
                    int colIndex = startColIndex;

                    foreach (string cellValue in cells)
                    {
                        while (colIndex < dgvDetails.Columns.Count &&
                               (!dgvDetails.Columns[colIndex].Visible ||
                                dgvDetails.Columns[colIndex].ReadOnly))
                            colIndex++;

                        if (colIndex >= dgvDetails.Columns.Count) break;

                        string colName = dgvDetails.Columns[colIndex].Name;
                        string value = cellValue.Trim();

                        if (colName == "Inspect_Result")
                            gridRow.Cells[colName].Value = NormalizeInspectResult(value);
                        else if (colName == "Qty_Required" || colName == "Qty_Received")
                            gridRow.Cells[colName].Value = int.TryParse(value, out int num) ? num : 0;
                        else
                            gridRow.Cells[colName].Value = value;

                        colIndex++;
                    }

                    startRowIndex++;
                    addedRows++;
                }

                RenumberItems();
                lblStatus.Text = $"✅ Đã dán {addedRows} dòng từ Excel. Nhấn 'Lưu chi tiết' để lưu vào CSDL.";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi dán dữ liệu từ Excel:\n" + ex.Message,
                    "Lỗi Paste", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string NormalizeInspectResult(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return "";
            string upper = raw.Trim().ToUpper();
            if (upper == "PASS") return "Pass";
            if (upper == "FAIL") return "Fail";
            if (upper == "HOLD") return "Hold";
            return "";
        }

        private void RenumberItems()
        {
            int no = 1;
            foreach (DataGridViewRow row in dgvDetails.Rows)
                if (!row.IsNewRow)
                    row.Cells["Item_No"].Value = no++;
        }

        // =========================================================================
        // HELPERS
        // =========================================================================
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

        private int IntVal(object val)
        {
            if (val == null || val == DBNull.Value) return 0;
            return int.TryParse(val.ToString(), out int i) ? i : 0;
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

                // Cập nhật vị trí folder browser
                PositionRIRFolderControls();

                // dgvDetails chiều rộng = panelDetail - folder panel - lề
                const int folderPanelW = 340;
                int detailW = panelDetail.ClientSize.Width - folderPanelW - 30;
                dgvDetails.Width = Math.Max(200, detailW);
                dgvDetails.Height = panelDetail.Height - 85;

                // Cập nhật chiều cao dgvRIRFolders
                if (dgvRIRFolders != null)
                    dgvRIRFolders.Height = panelDetail.Height - 70;

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
            dgvRIR.SelectionChanged -= DgvRIR_SelectionChanged; // tạm ngắt event
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

            dgvRIR.ClearSelection(); // không tự chọn dòng nào khi load
            dgvRIR.SelectionChanged += DgvRIR_SelectionChanged; // kết nối lại event
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

                    row.Cells["PO_Detail_ID"].Value = d.PO_Detail_ID;
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

            if ((dgvDetails.Columns[e.ColumnIndex].Name == "Qty_Required" && e.Value != null)
                && (dgvDetails.Columns[e.ColumnIndex].Name == "Qty_Received" && e.Value != null))
            {
                if (decimal.TryParse(e.Value.ToString(), out decimal qty))
                {
                    e.Value = qty.ToString("N0");
                    e.FormattingApplied = true;
                }
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

        private async void BtnSaveHead_Click(object sender, EventArgs e)
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
                    if (_selectedRIR_ID > 0)
                    {
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
                                    Qty_Required = (int)Math.Round(Convert.ToDecimal(row.Cells["Qty_Required"].Value ?? 0)),
                                    Qty_Received = (int)Math.Round(Convert.ToDecimal(row.Cells["Qty_Received"].Value ?? 0)),
                                    MTRno = row.Cells["MTRno"].Value?.ToString() ?? "",
                                    Heatno = row.Cells["Heatno"].Value?.ToString() ?? "",
                                    ID_Code = row.Cells["ID_Code"].Value?.ToString() ?? "",
                                    Inspect_Result = row.Cells["Inspect_Result"].Value?.ToString() ?? "",
                                    Remarks = row.Cells["Remarks"].Value?.ToString() ?? "",
                                    Qty_Per_Sheet = (int)Math.Round(Convert.ToDecimal(row.Cells["Qty_Required"].Value ?? 0)),
                                    PO_Detail_ID = Convert.ToInt32(row.Cells["PO_Detail_ID"].Value.ToString())
                                };

                                if (d.RIR_Detail_ID == 0)
                                {
                                    var rs = await _service.InsertRIRDetailAndUpdateStock(d);
                                }
                                else
                                {
                                    _service.UpdateDetail(d);
                                }
                                saved++;
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Lỗi lưu chi tiết: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }

                    // ── Tạo thư mục theo RIR No trong RIR_Link của dự án ──
                    TryCreateRIRFolder(h.RIR_No);

                    MessageBox.Show("Tạo phiếu RIR thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    _service.UpdateHead(h, _currentUser);

                    // ── Tạo thư mục nếu chưa có (khi cập nhật cũng kiểm tra) ──
                    TryCreateRIRFolder(h.RIR_No);

                    MessageBox.Show("Cập nhật phiếu RIR thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                LoadRIR();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi lưu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ── Tạo thư mục RIR No trong RIR_Link của dự án hiện tại ──
        private void TryCreateRIRFolder(string rirNo)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(rirNo)) return;

                // Ưu tiên lấy RIR_Link từ combo nếu đang chọn
                string rirLink = cboProjectRIR.SelectedValue?.ToString() ?? "";

                // Nếu combo chưa có → tra cứu qua WorkorderNo trước, rồi ProjectName
                if (string.IsNullOrWhiteSpace(rirLink))
                {
                    string wo = txtWorkorderNo.Text.Trim();
                    string projName = txtProjectName.Text.Trim();

                    foreach (DataRow row in _projectTable.Rows)
                    {
                        string rowWO = row["WorkorderNo"]?.ToString() ?? "";
                        string rowPN = row["ProjectName"]?.ToString() ?? "";
                        string rowCode = row["ProjectCode"]?.ToString() ?? "";
                        string rowLink = row["RIR_Link"]?.ToString() ?? "";

                        // Match theo WorkorderNo (chính xác nhất)
                        if (!string.IsNullOrEmpty(wo) && rowWO.Equals(wo, StringComparison.OrdinalIgnoreCase))
                        { rirLink = rowLink; break; }

                        // Fallback: match theo ProjectName hoặc ProjectCode
                        if (!string.IsNullOrEmpty(projName) &&
                            (rowPN.Equals(projName, StringComparison.OrdinalIgnoreCase) ||
                             rowCode.Equals(projName, StringComparison.OrdinalIgnoreCase)))
                        { rirLink = rowLink; break; }
                    }
                }

                if (string.IsNullOrWhiteSpace(rirLink))
                {
                    MessageBox.Show(
                        $"⚠ Không tìm thấy đường dẫn RIR Link của dự án.\n" +
                        $"Workorder: {txtWorkorderNo.Text.Trim()}\n" +
                        $"Dự án: {txtProjectName.Text.Trim()}\n\n" +
                        "Vui lòng kiểm tra lại thông tin dự án trong phần Quản lý Dự án.",
                        "Không tìm thấy RIR Link", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!System.IO.Directory.Exists(rirLink))
                {
                    MessageBox.Show(
                        $"⚠ Thư mục RIR Link không tồn tại trên ổ đĩa:\n{rirLink}\nThư mục không được tạo.",
                        "Thư mục không tồn tại", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Tên thư mục = RIR No (làm sạch ký tự không hợp lệ)
                string safeName = string.Join("_", rirNo.Split(System.IO.Path.GetInvalidFileNameChars()));
                string newFolderPath = System.IO.Path.Combine(rirLink, safeName);

                if (System.IO.Directory.Exists(newFolderPath))
                {
                    MessageBox.Show(
                        $"ℹ Thư mục đã tồn tại:\n{newFolderPath}",
                        "Thư mục đã tồn tại", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                System.IO.Directory.CreateDirectory(newFolderPath);
                MessageBox.Show(
                    $"✅ Đã tạo thư mục thành công:\n{newFolderPath}",
                    "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Refresh danh sách thư mục trong panel
                LoadRIRFolders(rirLink);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tạo thư mục: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
        // IMPORT TỪ PHIẾU NHẬP KHO (WAREHOUSE_IMPORT)
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

                    Action loadPNK = () =>
                    {
                        string kw = txtSearchPNK.Text.Trim();
                        string sql = @"
                            SELECT 
                                wi.Import_No            AS [Mã phiếu], 
                                MAX(wi.Import_Date)     AS [Ngày nhập], 
                                ph.PONo                 AS [PO No], 
                                MAX(ph.Project_Name)    AS [Dự án],
                                MAX(ph.WorkorderNo)     AS [Workorder],
                                MAX(ph.MPR_No)          AS [MPR No],
                                COUNT(wi.Import_ID)     AS [Số vật tư]
                            FROM Warehouse_Import wi
                            LEFT JOIN PO_head ph ON wi.PO_ID = ph.PO_ID
                            WHERE wi.Import_No    LIKE N'%' + @kw + '%' 
                               OR ph.PONo         LIKE N'%' + @kw + '%'
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
                            if (dgvPNK.Columns.Contains("Ngày nhập"))
                                dgvPNK.Columns["Ngày nhập"].DefaultCellStyle.Format = "dd/MM/yyyy HH:mm";
                        }
                    };

                    loadPNK();
                    btnFilter.Click += (s2, e2) => loadPNK();
                    txtSearchPNK.KeyDown += (s2, e2) => { if (e2.KeyCode == Keys.Enter) loadPNK(); };

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

                    if (dlg.ShowDialog() != DialogResult.OK) return;
                    if (dgvPNK.SelectedRows.Count == 0) return;

                    var selRow = dgvPNK.SelectedRows[0];
                    string pnkNo = selRow.Cells["Mã phiếu"].Value.ToString();
                    string poNo = selRow.Cells["PO No"].Value.ToString();
                    string projName = selRow.Cells["Dự án"].Value.ToString();
                    string woNo = selRow.Cells["Workorder"].Value.ToString();
                    string mprNo = selRow.Cells["MPR No"].Value.ToString();

                    string autoRIRNo = GenerateRIRNo(poNo, woNo);

                    txtRIRNo.Text = autoRIRNo;
                    txtPONo.Text = poNo;
                    txtMPRNo.Text = mprNo;
                    txtProjectName.Text = projName;
                    txtWorkorderNo.Text = woNo;
                    dtpIssueDate.Value = DateTime.Today;
                    cboStatus.SelectedIndex = 0;

                    try
                    {
                        var projects = new ProjectService().GetAll();
                        var proj2 = projects.Find(p => p.WorkorderNo == woNo);
                        txtCustomer.Text = proj2?.Customer ?? "";
                    }
                    catch { txtCustomer.Text = ""; }

                    string sqlDetails = @"
                        SELECT 
                            wi.Item_Name, 
                            wi.Material, 
                            wi.Size, 
                            wi.UNIT, 
                            wi.Qty_Import, 
                            wi.ID_Code, 
                            ISNULL(wi.MTRno, '') AS MTRno,
                            PO_Detail_ID
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
                                row.Cells["Qty_Required"].Value = reader["Qty_Import"] != DBNull.Value
                                                                        ? Convert.ToDecimal(reader["Qty_Import"]) : 0;
                                row.Cells["Qty_Received"].Value = 0;
                                row.Cells["MTRno"].Value = reader["MTRno"]?.ToString() ?? "";
                                row.Cells["Heatno"].Value = "";
                                row.Cells["ID_Code"].Value = "";
                                row.Cells["Inspect_Result"].Value = "";
                                row.Cells["Remarks"].Value = "";
                                row.Cells["PO_Detail_ID"].Value = reader["PO_Detail_ID"].ToString() ?? "";
                                countItems++;
                            }
                        }
                    }

                    if (countItems == 0)
                        MessageBox.Show("Phiếu nhập kho này không có chi tiết vật tư!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    else
                        MessageBox.Show(
                            $"✅ Đã kéo dữ liệu từ phiếu nhập: {pnkNo}\n" +
                            $"Mã RIR tạo mới: {autoRIRNo}\n" +
                            $"Số lượng vật tư: {countItems} mục\n\n" +
                            $"Nhấn 'Lưu Header' và 'Lưu chi tiết' để hoàn tất.",
                            "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            row.Cells["PO_Detail_ID"].Value = "";

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
                        RenumberItems();
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
                RenumberItems();
            }
        }

        private async void BtnSaveDetail_Click(object sender, EventArgs e)
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
                        Qty_Required = (int)Math.Round(Convert.ToDecimal(row.Cells["Qty_Required"].Value ?? 0)),
                        Qty_Received = (int)Math.Round(Convert.ToDecimal(row.Cells["Qty_Received"].Value ?? 0)),
                        MTRno = row.Cells["MTRno"].Value?.ToString() ?? "",
                        Heatno = row.Cells["Heatno"].Value?.ToString() ?? "",
                        ID_Code = row.Cells["ID_Code"].Value?.ToString() ?? "",
                        Inspect_Result = row.Cells["Inspect_Result"].Value?.ToString() ?? "",
                        Remarks = row.Cells["Remarks"].Value?.ToString() ?? ""
                    };

                    if (d.RIR_Detail_ID == 0)
                    {
                        var rs = await _service.InsertRIRDetailAndUpdateStock(d);
                    }
                    else
                    {
                        _service.UpdateDetail(d);
                    }
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

            var poModel = await _poService.GetPOAsync(poNO);
            var projectMode = _projectServices.GetByProjectCode(poModel.ProjectCode);
            var supplierModel = _supplierServices.GetBySupId(poModel.Supplier_ID);
            var dtImports = await _service.GetDetailsToExport(rirId);
            PrintBill(dtImports, projectMode, rirNO, poModel);
        }

        public void PrintBill(DataTable dtDetails, ProjectInfo projects, string RIRNo, POHead po)
        {
            try
            {
                if (dgvRIR.CurrentRow == null)
                {
                    MessageBox.Show("Vui lòng chọn một phiếu để in!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "rir_template_v2.xlsx");
                string exportFolder = projects.RIR_Link;
                if (!Directory.Exists(exportFolder)) Directory.CreateDirectory(exportFolder);

                string fileName = $"{RIRNo}_{DateTime.Now:ddMMyyyy_HHmmss}.xlsx";
                string actualSavePath = Path.Combine(exportFolder, fileName);

                if (!File.Exists(templatePath))
                {
                    MessageBox.Show("Không tìm thấy file template!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // --- BƯỚC FIX LỖI: Copy file mẫu ra file mới trước ---
                File.Copy(templatePath, actualSavePath, true);

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                FileInfo newFile = new FileInfo(actualSavePath);

                // Mở trực tiếp file mới để thao tác
                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    ExcelWorksheet ws = package.Workbook.Worksheets[0];

                    // 1. Thay thế Header (A1 đến AZ15)
                    var headerCells = ws.Cells["A1:AZ15"];
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

                    // 2. Điền chi tiết vật tư (Dòng 11 theo template v2)
                    int startRow = 7;
                    int count = dtDetails.Rows.Count;

                    if (count > 1)
                    {
                        ws.InsertRow(startRow + 1, count - 1);
                        // Copy format từ dòng gốc cho các dòng mới
                        ws.Cells[startRow, 1, startRow, 50].Copy(ws.Cells[startRow + 1, 1, startRow + count - 1, 50]);
                    }

                    for (int i = 0; i < count; i++)
                    {
                        DataRow dr = dtDetails.Rows[i];
                        int curr = startRow + i;

                        ws.Row(curr).Height = 28; // Chiều cao dòng

                        // Gán giá trị (Cột theo template v2 của bạn)
                        ws.Cells[curr, 1].Value = i + 1;
                        ws.Cells[curr, 2].Value = dr["item_name"];
                        ws.Cells[curr, 4].Value = dr["Material"];
                        ws.Cells[curr, 5].Value = dr["Size"];
                        ws.Cells[curr, 6].Value = dr["UNIT"];
                        ws.Cells[curr, 7].Value = dr["Qty_Per_Sheet"];
                        ws.Cells[curr, 8].Value = dr["MTRno"];
                        ws.Cells[curr, 9].Value = dr["Heatno"];
                        ws.Cells[curr, 11].Value = "";
                        ws.Cells[curr, 12].Value = "";

                        // --- UPDATE ĐỊNH DẠNG ĐỒNG NHẤT ---
                        using (var range = ws.Cells[curr, 1, curr, 12])
                        {
                            range.Style.Font.Name = "Times New Roman";
                            range.Style.Font.Size = 11;
                            range.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                            range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        }

                        // Căn trái riêng cho cột tên (B)
                        ws.Cells[curr, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    }

                    // Lưu file - Sẽ không còn lỗi "Closed file"
                    package.Save();
                }

                if (MessageBox.Show("✅ Xuất phiếu RIR thành công! Mở file ngay?", "Thành công",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(actualSavePath) { UseShellExecute = true });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //public void PrintBill(DataTable dtDetails, ProjectInfo projects, string RIRNo, POHead po)
        //{
        //    try
        //    {
        //        if (dgvRIR.CurrentRow == null)
        //        {
        //            MessageBox.Show("Vui lòng chọn một phiếu nhập kho để in!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //            return;
        //        }

        //        string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "rir_template.xlsx");
        //        string exportFolder = projects.RIR_Link;
        //        if (!Directory.Exists(exportFolder))
        //            Directory.CreateDirectory(exportFolder);

        //        string fileName = $"{RIRNo}_{DateTime.Now:ddMMyyyy}.xlsx";
        //        string actualSavePath = Path.Combine(exportFolder, fileName);

        //        if (!File.Exists(templatePath))
        //        {
        //            MessageBox.Show("Không tìm thấy file template tại: " + templatePath, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //            return;
        //        }

        //        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        //        FileInfo templateFile = new FileInfo(templatePath);
        //        FileInfo newFile = new FileInfo(actualSavePath);
        //        using (ExcelPackage package = new ExcelPackage(newFile, templateFile))
        //        {
        //            ExcelWorksheet ws = package.Workbook.Worksheets[0];

        //            var headerCells = ws.Cells["A1:AW12"];
        //            foreach (var cell in headerCells)
        //            {
        //                if (cell.Value == null) continue;
        //                string txt = cell.Value.ToString();

        //                if (txt.Contains("<<RIR-NO>>")) cell.Value = txt.Replace("<<RIR-NO>>", RIRNo);
        //                if (txt.Contains("<<DATE>>")) cell.Value = txt.Replace("<<DATE>>", DateTime.Now.ToString("dd/MM/yyyy"));
        //                if (txt.Contains("<<WO-NO>>")) cell.Value = txt.Replace("<<WO-NO>>", projects.WorkorderNo ?? "");
        //                if (txt.Contains("<<MPR-NO>>")) cell.Value = txt.Replace("<<MPR-NO>>", po.MPR_No ?? "");
        //                if (txt.Contains("<<PROJECT-NAME>>")) cell.Value = txt.Replace("<<PROJECT-NAME>>", projects.ProjectName ?? "");
        //                if (txt.Contains("<<MPS-NO>>")) cell.Value = txt.Replace("<<MPS-NO>>", "");
        //                if (txt.Contains("<<CLIENT>>")) cell.Value = txt.Replace("<<CLIENT>>", projects.Customer ?? "");
        //                if (txt.Contains("<<PO-NO>>")) cell.Value = txt.Replace("<<PO-NO>>", po.PONo ?? "");
        //                if (txt.Contains("<<USER-CREATE>>")) cell.Value = txt.Replace("<<USER-CREATE>>", _currentUser);
        //            }

        //            int startRow = 7;
        //            int count = dtDetails.Rows.Count;
        //            if (count > 1)
        //            {
        //                ws.InsertRow(startRow + 1, count - 1);
        //                for (int i = 1; i < count; i++)
        //                    ws.Cells[startRow, 1, startRow, 30].Copy(ws.Cells[startRow + i, 1]);
        //            }

        //            for (int i = 0; i < count; i++)
        //            {
        //                DataRow dr = dtDetails.Rows[i];
        //                int curr = startRow + i;
        //                ws.Row(curr).Height = 25;

        //                ws.Cells[curr, 1].Value = i + 1;
        //                ws.Cells[curr, 3].Value = dr["item_name"];
        //                ws.Cells[curr, 10].Value = dr["Material"];
        //                ws.Cells[curr, 18].Value = dr["Size"];
        //                ws.Cells[curr, 25].Value = dr["UNIT"];
        //                ws.Cells[curr, 27].Value = dr["Qty_Per_Sheet"];
        //                ws.Cells[curr, 29].Value = dr["MTRno"];
        //                ws.Cells[curr, 35].Value = dr["Heatno"];
        //                ws.Cells[curr, 41].Value = "";
        //                ws.Cells[curr, 46].Value = "";

        //                using (var range = ws.Cells[curr, 1, curr, 50])
        //                {
        //                    range.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
        //                    range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //                    range.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
        //                }

        //                ws.Cells[curr, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
        //                ws.Cells[curr, 18].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

        //                if (i > 0)
        //                {
        //                    for (int col = 1; col <= 16; col++)
        //                    {
        //                        ws.Cells[curr, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        //                        ws.Cells[curr, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        //                        ws.Cells[curr, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        //                        ws.Cells[curr, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        //                        ws.Cells[curr, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //                        ws.Cells[curr, col].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
        //                        ws.Cells[curr, col].Style.Font.Name = "Times New Roman";
        //                        ws.Cells[curr, col].Style.Font.Size = 9;
        //                        ws.Cells[curr, col].Style.Font.Italic = false;
        //                    }
        //                }
        //            }

        //            package.Save();
        //        }

        //        var result = MessageBox.Show(
        //            $"✅ Xuất phiếu nhập kho thành công!\nFile: {actualSavePath}\n\nBạn có muốn mở file ngay không?",
        //            "Thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
        //        if (result == DialogResult.Yes)
        //        {
        //            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
        //            {
        //                FileName = actualSavePath,
        //                UseShellExecute = true
        //            });
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Lỗi khi in phiếu: " + ex.Message, "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
    }
}