using Microsoft.Data.SqlClient;
using MPR_Managerment.Forms.MPRGUI;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;
using MPR_Managerment.Services;
using OfficeOpenXml;
using Syncfusion.XlsIO.Implementation.XmlSerialization;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace MPR_Managerment.Forms
{
    public partial class frmMPR : Form
    {
        private MPRService _service = new MPRService();
        private List<MPRHeader> _mprList = new List<MPRHeader>();
        private List<MPRDetail> _details = new List<MPRDetail>();
        private int _selectedMPR_ID = 0;
        private string _currentUser = "Admin";

        // Thêm biến để lưu ID truyền từ Dashboard sang
        private int _targetMprId = 0;

        private DataGridView dgvMPR;
        private TextBox txtSearch;
        private Button btnSearch, btnNewMPR, btnSaveHeader, btnDeleteMPR, btnClearHeader;
        private Label lblStatus;

        private TextBox txtMPRNo, txtProjectName, txtProjectCode, txtDepartment, txtRequestor, txtRev, txtNotes;
        private DateTimePicker dtpRequiredDate;
        private ComboBox cboStatus;

        // BẢNG MỚI: Danh sách file đính kèm
        private DataGridView dgvFiles;

        private DataGridView dgvDetails;
        private Button btnAddDetail, btnDeleteDetail, btnSaveDetail;

        // BẢNG: Tiến độ PO
        private DataGridView dgvPOProgress;
        private Label lblPOProgressTitle;

        private Panel panelTop, panelHeader, panelDetail;
        private ComboBox _cboFilterPO;    // Loc theo Da len PO
        private Button _btnExportDetail; // Xuat Excel chi tiet

        // Khai báo phía trên cùng của Class Form
        private System.Diagnostics.Process _excelProcess = null;


        public frmMPR(int mprId = 0)
        {
            _targetMprId = mprId;
            InitializeComponent();
            BuildUI();
            ApplyPermissions();
            LoadMPR();
            this.Resize += FrmMPR_Resize;
            this.WindowState = FormWindowState.Maximized;

            if (_targetMprId > 0)
            {
                SelectMPRById(_targetMprId);
            }
        }

        private void SelectMPRById(int id)
        {
            var targetMPR = _mprList.Find(m => m.MPR_ID == id);
            if (targetMPR != null)
            {
                txtSearch.Text = targetMPR.MPR_No;
                BtnSearch_Click(null, null);
            }

            foreach (DataGridViewRow row in dgvMPR.Rows)
            {
                if (Convert.ToInt32(row.Cells["ID"].Value) == id)
                {
                    dgvMPR.ClearSelection();
                    row.Selected = true;

                    if (row.Index >= 0)
                        dgvMPR.FirstDisplayedScrollingRowIndex = row.Index;
                    break;
                }
            }
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

            // ── "➕ Tạo MPR" mới: mở popup tạo MPR ──
            var btnCreateMPR = CreateButton("➕ Tạo MPR", Color.FromArgb(40, 167, 69), new Point(415, 47), 110, 30);
            btnCreateMPR.Click += BtnCreateMPR_Click;
            panelTop.Controls.Add(btnCreateMPR);

            // ── "Update from Excel": chức năng cũ của btnNewMPR ──
            btnNewMPR = CreateButton("📥 Update from Excel", Color.FromArgb(0, 140, 120), new Point(533, 47), 155, 30);
            btnNewMPR.Click += BtnNewMPR_Click;
            panelTop.Controls.Add(btnNewMPR);

            btnDeleteMPR = CreateButton("🗑 Xóa MPR", Color.FromArgb(220, 53, 69), new Point(696, 47), 110, 30);
            btnDeleteMPR.Click += BtnDeleteMPR_Click;
            panelTop.Controls.Add(btnDeleteMPR);

            var btnPrint = new Button
            {
                Text = "🖨 In MPR",
                Location = new Point(860, 47),
                Size = new Size(110, 30),
                BackColor = Color.FromArgb(33, 115, 70),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            btnPrint.FlatAppearance.BorderSize = 0;
            btnPrint.Click += BtnPrint_Click;
            panelTop.Controls.Add(btnPrint);

            lblStatus = new Label
            {
                Location = new Point(820, 52),
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
            dgvMPR.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
            dgvMPR.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgvMPR.SelectionChanged += DgvMPR_SelectionChanged;
            panelTop.Controls.Add(dgvMPR);

            // ===== PANEL HEADER =====
            panelHeader = new Panel
            {
                Location = new Point(10, 240),
                Size = new Size(1360, 160),
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

            // === BẢNG FILE ĐÍNH KÈM ===
            int gridFilesWidth = 450;
            int filesLeft = panelHeader.Width - gridFilesWidth - 10;
            dgvFiles = new DataGridView
            {
                Location = new Point(filesLeft, 10),
                Size = new Size(gridFilesWidth, panelHeader.Height - 20),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom
            };
            dgvFiles.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(108, 117, 125);
            dgvFiles.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvFiles.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvFiles.EnableHeadersVisualStyles = false;
            dgvFiles.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvFiles.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
            dgvFiles.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgvFiles.Columns.Add("FileName", "Tệp đính kèm (MPR Link)");
            dgvFiles.Columns.Add("FullPath", "FullPath");
            dgvFiles.Columns["FullPath"].Visible = false;
            dgvFiles.CellDoubleClick += DgvFiles_CellDoubleClick;
            panelHeader.Controls.Add(dgvFiles);

            // === CÁC CONTROL NHẬP LIỆU BÊN TRÁI ===
            int y = 38;

            // Hàng 1
            AddLabel(panelHeader, "MPR No (*):", 10, y);
            txtMPRNo = AddTextBox(panelHeader, 90, y, 140);

            AddLabel(panelHeader, "Tên dự án:", 240, y);
            txtProjectName = AddTextBox(panelHeader, 320, y, 200);

            AddLabel(panelHeader, "Mã dự án:", 530, y);
            txtProjectCode = AddTextBox(panelHeader, 610, y, 130);

            AddLabel(panelHeader, "Trạng thái:", 750, y);
            cboStatus = new ComboBox
            {
                Location = new Point(830, y),
                Size = new Size(60, 25),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboStatus.Items.AddRange(new[] { "Mới", "Đang xử lý", "Đã duyệt", "Hoàn thành", "Hủy" });
            cboStatus.SelectedIndex = 0;
            panelHeader.Controls.Add(cboStatus);

            // Hàng 2
            y += 38;
            AddLabel(panelHeader, "Phòng ban:", 10, y);
            txtDepartment = AddTextBox(panelHeader, 90, y, 140);

            AddLabel(panelHeader, "Người YC:", 240, y);
            txtRequestor = AddTextBox(panelHeader, 320, y, 200);

            AddLabel(panelHeader, "Ngày cần:", 530, y);
            dtpRequiredDate = new DateTimePicker
            {
                Location = new Point(610, y),
                Size = new Size(130, 25),
                Font = new Font("Segoe UI", 9),
                Format = DateTimePickerFormat.Short
            };
            panelHeader.Controls.Add(dtpRequiredDate);

            AddLabel(panelHeader, "Rev:", 750, y);
            txtRev = AddTextBox(panelHeader, 790, y, 100);
            txtRev.Text = "0";

            // Hàng 3 (Buttons & Notes)
            y += 38;
            btnSaveHeader = CreateButton("💾 Lưu Header", Color.FromArgb(0, 120, 212), new Point(10, y), 130, 32);
            btnSaveHeader.Click += BtnSaveHeader_Click;
            panelHeader.Controls.Add(btnSaveHeader);

            btnClearHeader = CreateButton("🔄 Làm mới", Color.FromArgb(108, 117, 125), new Point(150, y), 110, 32);
            btnClearHeader.Click += BtnClearHeader_Click;
            panelHeader.Controls.Add(btnClearHeader);

            AddLabel(panelHeader, "Ghi chú:", 270, y + 5);
            txtNotes = AddTextBox(panelHeader, 340, y + 2, filesLeft - 340 - 15);
            txtNotes.Anchor = AnchorStyles.Top | AnchorStyles.Left;

            // ===== PANEL DETAIL =====
            panelDetail = new Panel
            {
                Location = new Point(10, panelHeader.Bottom + 10),
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

            var btnCreatePO = CreateButton("🛒 Tạo PO", Color.FromArgb(255, 140, 0), new Point(400, 38), 120, 30);
            btnCreatePO.Click += BtnCreatePO_Click;
            panelDetail.Controls.Add(btnCreatePO);

            var btnCheckAll = CreateButton("🔎 Check All Items", Color.FromArgb(102, 51, 153), new Point(530, 38), 150, 30);
            btnCheckAll.Click += BtnCheckAllItems_Click;
            panelDetail.Controls.Add(btnCheckAll);

            // ── Loc theo Da len PO ──
            panelDetail.Controls.Add(new Label
            {
                Text = "Da len PO:",
                Location = new Point(695, 45),
                Size = new Size(72, 22),
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            });
            _cboFilterPO = new ComboBox
            {
                Location = new Point(770, 44),
                Size = new Size(120, 24),
                Font = new Font("Segoe UI", 8),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cboFilterPO.Items.Add("(Tat ca)"); // placeholder, se duoc load dong
            _cboFilterPO.SelectedIndex = 0;
            _cboFilterPO.SelectedIndexChanged += (s, ev) => FilterDetailByPO();
            panelDetail.Controls.Add(_cboFilterPO);

            // ── Xuat Excel ──
            _btnExportDetail = CreateButton("📥 Xuat Excel", Color.FromArgb(0, 150, 100), new Point(898, 38), 120, 30);
            _btnExportDetail.Click += BtnExportDetail_Click;
            panelDetail.Controls.Add(_btnExportDetail);

            dgvDetails = new DataGridView
            {
                Location = new Point(10, 75),
                Size = new Size(900, 260),
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom
            };
            dgvDetails.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvDetails.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvDetails.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvDetails.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvDetails.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvDetails.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvDetails.EnableHeadersVisualStyles = false;
            dgvDetails.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvDetails.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
            dgvDetails.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgvDetails.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvDetails.CellFormatting += DgvDetails_CellFormatting;
            dgvDetails.KeyDown += DgvDetails_GridKeyDown;
            // Cho phep copy nhieu o
            dgvDetails.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;

            BuildDetailColumns();
            panelDetail.Controls.Add(dgvDetails);

            lblPOProgressTitle = new Label
            {
                Text = "TỔNG HỢP TIẾN ĐỘ PO ĐÃ ĐẶT",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(255, 140, 0),
                Location = new Point(930, 48),
                Size = new Size(300, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            panelDetail.Controls.Add(lblPOProgressTitle);

            dgvPOProgress = new DataGridView
            {
                Location = new Point(930, 75),
                Size = new Size(415, 260),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom
            };
            dgvPOProgress.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 140, 0);
            dgvPOProgress.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPOProgress.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvPOProgress.EnableHeadersVisualStyles = false;
            dgvPOProgress.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(255, 248, 235);
            dgvPOProgress.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
            dgvPOProgress.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgvPOProgress.CellFormatting += DgvPOProgress_CellFormatting;
            dgvPOProgress.CellDoubleClick += DgvPOProgress_CellDoubleClick;
            panelDetail.Controls.Add(dgvPOProgress);

            Common.Common.AutoBringToFontControl(new[] { panelTop, panelHeader, panelDetail });
        }

        private void BtnPrint_Click(object? sender, EventArgs e)
        {
            if (dgvMPR.Rows.Count <= 0) return;
            int rsl = dgvMPR.CurrentRow.Index;
            int mprId = Convert.ToInt32(dgvMPR.Rows[rsl].Cells["ID"].Value.ToString().Trim());

            var projectName = dgvMPR.CurrentRow.Cells["Ten_Du_An"].Value.ToString().Trim();
            var mprNo = dgvMPR.CurrentRow.Cells["MPR_No"].Value.ToString().Trim();
            var mpr_detail = _service.GetActiveDetails(mprId);

            var mpr_header = new MPRHeader() { MPR_ID = mprId, MPR_No = mprNo, Project_Name = projectName };

            ExportMPRToExcel(mpr_header, mpr_detail);
        }

        // =====================================================================
        // SỰ KIỆN DOUBLE CLICK VÀO FILE TRONG BẢNG ĐÍNH KÈM
        // =====================================================================
        private void DgvFiles_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string path = dgvFiles.Rows[e.RowIndex].Cells["FullPath"].Value?.ToString();

            if (!string.IsNullOrEmpty(path) && File.Exists(path))
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = path,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Không thể mở file: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else if (!string.IsNullOrEmpty(path))
            {
                MessageBox.Show("File không tồn tại hoặc đã bị xóa / di chuyển khỏi thư mục!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // =====================================================================
        // HÀM QUÉT THƯ MỤC LẤY DANH SÁCH FILE CỦA DỰ ÁN ĐANG CHỌN
        // =====================================================================
        private void LoadFiles(string projectName)
        {
            dgvFiles.Rows.Clear();
            if (string.IsNullOrEmpty(projectName)) return;

            try
            {
                var projects = new ProjectService().GetAll();
                var prj = projects.Find(p =>
                    !string.IsNullOrEmpty(p.ProjectName) &&
                    p.ProjectName.Equals(projectName, StringComparison.OrdinalIgnoreCase));

                if (prj == null)
                {
                    prj = projects.Find(p =>
                        !string.IsNullOrEmpty(p.ProjectName) &&
                        p.ProjectName.IndexOf(projectName, StringComparison.OrdinalIgnoreCase) >= 0);
                }

                if (prj != null && !string.IsNullOrEmpty(prj.MPR_Link) && Directory.Exists(prj.MPR_Link))
                {
                    var files = Directory.GetFiles(prj.MPR_Link);
                    foreach (var f in files)
                    {
                        dgvFiles.Rows.Add(Path.GetFileName(f), f);
                    }
                }
                else if (prj != null && !string.IsNullOrEmpty(prj.MPR_Link))
                {
                    dgvFiles.Rows.Add("(Thư mục không tồn tại)", "");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Lỗi load files: " + ex.Message);
            }
        }

        private void DgvPOProgress_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string poNo = dgvPOProgress.Rows[e.RowIndex].Cells["PO No"].Value?.ToString() ?? "";

            if (!string.IsNullOrEmpty(poNo))
            {
                var frm = new frmPO(poNo);
                frm.Show();
            }
        }

        private void BuildDetailColumns()
        {
            dgvDetails.Columns.Clear();
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Detail_ID", HeaderText = "ID", Visible = false });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Item_No",
                HeaderText = "STT",
                Width = 45,
                ReadOnly = true,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Item_Name",
                HeaderText = "Tên vật tư",
                Width = 180,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleLeft }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Description",
                HeaderText = "Mô tả",
                Width = 100,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleLeft }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Material",
                HeaderText = "Vật liệu",
                Width = 85,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Thickness_mm",
                HeaderText = "A-Dày(mm)",
                Width = 45,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Depth_mm",
                HeaderText = "B-Sâu(mm)",
                Width = 45,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "C_Width_mm",
                HeaderText = "C-Rộng(mm)",
                Width = 45,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "D_Web_mm",
                HeaderText = "D-Bụng(mm)",
                Width = 45,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "E_Flange_mm",
                HeaderText = "E-Cánh(mm)",
                Width = 45,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "F_Length_mm",
                HeaderText = "F-Dài(mm)",
                Width = 75,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "UNIT",
                HeaderText = "ĐVT",
                Width = 50,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Qty",
                HeaderText = "SL",
                Width = 50,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Weight",
                HeaderText = "KG",
                Width = 55,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "MPS_Info",
                HeaderText = "MPS Info",
                Width = 100,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Usage_Location",
                HeaderText = "Vị trí dùng",
                Width = 110,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleLeft }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "REV",
                HeaderText = "REV",
                Width = 45,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Remarks",
                HeaderText = "Ghi chú",
                FillWeight = 100,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleLeft }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "PO_No",
                HeaderText = "Đã lên PO",
                Width = 120,
                ReadOnly = true,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
        }

        private void DgvDetails_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string colName = dgvDetails.Columns[e.ColumnIndex].Name;

            if (colName == "Thickness_mm" || colName == "Depth_mm" ||
                colName == "C_Width_mm" || colName == "D_Web_mm" ||
                colName == "E_Flange_mm" || colName == "F_Length_mm" ||
                colName == "Qty" || colName == "Weight")
            {
                if (e.Value != null && decimal.TryParse(e.Value.ToString(), out decimal num) && num == 0)
                {
                    e.Value = "";
                    e.FormattingApplied = true;
                }
                e.CellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            // Cột PO_No — màu xanh bold
            if (colName == "PO_No")
            {
                string val = e.Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(val))
                {
                    e.CellStyle.ForeColor = Color.FromArgb(40, 167, 69);
                    e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                }
            }
        }

        private void DgvPOProgress_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string colName = dgvPOProgress.Columns[e.ColumnIndex].Name;

            if (colName == "% Giao")
            {
                if (decimal.TryParse(e.Value?.ToString(), out decimal pct))
                {
                    e.CellStyle.ForeColor = pct >= 100 ? Color.FromArgb(40, 167, 69) : pct >= 50 ? Color.FromArgb(255, 140, 0) : Color.FromArgb(220, 53, 69);
                    e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    e.Value = $"{pct}%";
                    e.FormattingApplied = true;
                }
            }
            if (colName == "Ngày PO")
            {
                if (e.Value != null && e.Value != DBNull.Value)
                {
                    e.Value = Convert.ToDateTime(e.Value).ToString("dd/MM/yyyy");
                    e.FormattingApplied = true;
                }
            }
        }

        private void AddLabel(Panel panel, string text, int x, int y)
        {
            panel.Controls.Add(new Label
            {
                Text = text,
                Location = new Point(x, y + 3),
                Size = new Size(80, 20),
                Font = new Font("Segoe UI", 9),
                Margin = new Padding(0)
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
        }

        // ===== RESIZE =====
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

                int rightWidth = 420;
                dgvPOProgress.Width = rightWidth;
                dgvPOProgress.Left = panelDetail.Width - rightWidth - 10;
                dgvPOProgress.Height = panelDetail.Height - 85;

                lblPOProgressTitle.Left = dgvPOProgress.Left;

                dgvDetails.Width = dgvPOProgress.Left - 20;
                dgvDetails.Height = panelDetail.Height - 85;
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

        // ===== LOAD TỔNG HỢP TIẾN ĐỘ PO =====
        private void LoadPOProgress(string mprNo)
        {
            if (string.IsNullOrEmpty(mprNo))
            {
                dgvPOProgress.DataSource = null;
                return;
            }

            try
            {
                string sql = @"
                    SELECT
                        h.PONo AS [PO No],
                        h.PO_Date AS [Ngày PO],
                        h.Status AS [Trạng thái],
                        CASE
                            WHEN ISNULL(SUM(d.Qty_Per_Sheet), 0) = 0 THEN 0
                            ELSE CAST(
                                ISNULL(
                                    (SELECT SUM(Qty_Import) FROM Warehouse_Import wi WHERE wi.PO_ID = h.PO_ID), 
                                    ISNULL(SUM(d.Received), 0)
                                ) * 100.0 / SUM(d.Qty_Per_Sheet) 
                            AS DECIMAL(5,1))
                        END AS [% Giao]
                    FROM PO_head h
                    LEFT JOIN PO_Detail d ON h.PO_ID = d.PO_ID
                    WHERE h.MPR_No = @mprNo
                    GROUP BY h.PO_ID, h.PONo, h.PO_Date, h.Status
                    ORDER BY h.PO_Date DESC";
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@mprNo", mprNo);
                    var dt = new DataTable();
                    dt.Load(cmd.ExecuteReader());
                    dgvPOProgress.DataSource = dt;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Lỗi tải PO Progress: " + ex.Message);
            }
        }

        // =====================================================================
        // LẤY PO MAPPING CHO MPR — HỖ TRỢ CẢ MPR REVISE
        // Logic:
        //   Bước 1: Tìm PO liên kết trực tiếp qua MPR_Detail_ID (Detail_ID hiện tại)
        //   Bước 2: Với các dòng chưa có PO, tìm sang các phiên bản MPR khác cùng MPR_No
        //           (revise) khớp theo Item_No + Item_Name + Material để lấy PO đã đặt
        //           từ phiên bản cũ — rồi điền vào dòng tương ứng của phiên bản mới
        // =====================================================================
        // Trả về dict: Detail_ID → danh sách PONo đã đặt cho từng vật tư
        // Chỉ lấy PO qua PO_Detail.MPR_Detail_ID = MPR_Details.Detail_ID
        // (đúng 1 vật tư MPR → 1 hoặc nhiều dòng PO_Detail → 1 hoặc nhiều PO)
        private Dictionary<int, string> GetPoMappingForMpr(int mprId)
        {
            var dict = new Dictionary<int, string>();
            if (mprId <= 0) return dict;
            try
            {
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();

                    // Join chính xác: 1 dòng MPR_Details → nhiều dòng PO_Detail → nhiều PO_head
                    // GROUP BY Detail_ID để gộp nhiều PO của cùng 1 vật tư
                    string sql = @"
                        SELECT   pod.MPR_Detail_ID  AS Detail_ID,
                                 poh.PONo
                        FROM     PO_Detail pod
                        INNER JOIN PO_head poh ON poh.PO_ID = pod.PO_ID
                        WHERE    pod.MPR_Detail_ID IN (
                                     SELECT Detail_ID
                                     FROM   MPR_Details
                                     WHERE  MPR_ID = @mprId
                                 )
                        ORDER BY pod.MPR_Detail_ID, poh.PONo";

                    using (var cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@mprId", mprId);
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                if (reader["Detail_ID"] == DBNull.Value) continue;
                                int detailId = Convert.ToInt32(reader["Detail_ID"]);
                                string poNo = reader["PONo"]?.ToString()?.Trim() ?? "";
                                if (string.IsNullOrEmpty(poNo)) continue;

                                if (dict.ContainsKey(detailId))
                                {
                                    // Tránh trùng PO (1 PO có nhiều dòng cùng vật tư)
                                    var existing = dict[detailId].Split(new[] { ", " },
                                        StringSplitOptions.RemoveEmptyEntries);
                                    if (!Array.Exists(existing, p => p == poNo))
                                        dict[detailId] += ", " + poNo;
                                }
                                else
                                {
                                    dict[detailId] = poNo;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Lỗi lấy PO Mapping: " + ex.Message);
            }
            return dict;
        }

        private void LoadDetails(int mprId)
        {
            try
            {
                _details = _service.GetDetails(mprId);
                dgvDetails.Rows.Clear();

                var poMapping = GetPoMappingForMpr(mprId);

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
                    row.Cells["PO_No"].Value = poMapping.ContainsKey(d.Detail_ID) ? poMapping[d.Detail_ID] : "";
                }
                // Populate combobox filter voi cac gia tri thuc te tu cot PO_No
                RefreshPOFilterCombo();
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
            LoadPOProgress(m.MPR_No);
            LoadFiles(m.Project_Name);
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
            //if (!PermissionHelper.Check("MPR", "Tạo MPR", "Tạo MPR")) return;
            //_selectedMPR_ID = 0;
            //ClearHeader();
            //dgvDetails.Rows.Clear();
            //dgvPOProgress.DataSource = null;
            //dgvFiles.Rows.Clear();
            //_details.Clear();
            //txtMPRNo.Focus();

            //string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "SQLTesting-Template.xlsm");

            //ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            //FileInfo newFile = new FileInfo(templatePath);
            //System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(templatePath) { UseShellExecute = true });

            ////frmExcelPreview frm = new frmExcelPreview(templatePath, "Xem trước biểu mẫu");
            ////frm.Owner = this; // Rất quan trọng: Khi tắt chương trình (Form chính), Form này tắt theo
            ////frm.Show();

            if (!PermissionHelper.Check("MPR", "Tạo MPR", "Tạo MPR")) return;

            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "SQLTesting-Template.xlsm");
            if (!File.Exists(templatePath)) { MessageBox.Show("Không tìm thấy file!"); return; }

            Form mainForm = this.ParentForm ?? this.FindForm();

            try
            {
                if (mainForm != null) mainForm.Hide();

                // 1. Khởi chạy file
                ProcessStartInfo startInfo = new ProcessStartInfo(templatePath) { UseShellExecute = true };
                Process p = Process.Start(startInfo);

                // 2. Chờ một chút để Excel kịp load file
                System.Threading.Thread.Sleep(2000);

                // 3. Tìm tiến trình thực sự đang giữ file Excel đó
                // (Vì Excel thường gom các file vào 1 tiến trình duy nhất "EXCEL")
                Process actualProcess = null;
                string fileName = Path.GetFileNameWithoutExtension(templatePath);

                // Lặp lại việc tìm kiếm cho đến khi Excel thực sự đóng
                bool isExcelRunning = true;
                while (isExcelRunning)
                {
                    // Kiểm tra xem có tiến trình Excel nào đang mở file của mình không
                    // Lưu ý: MainWindowTitle của Excel thường có dạng "Tên_File - Excel"
                    var processes = Process.GetProcessesByName("EXCEL");
                    isExcelRunning = false;

                    foreach (var proc in processes)
                    {
                        if (proc.MainWindowTitle.Contains(fileName))
                        {
                            isExcelRunning = true;
                            break;
                        }
                    }

                    if (isExcelRunning)
                    {
                        System.Threading.Thread.Sleep(1000); // Đợi 1 giây rồi kiểm tra lại
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
            finally
            {
                // 4. Luôn đảm bảo hiện lại Form khi kết thúc vòng lặp
                if (mainForm != null)
                {
                    mainForm.Show();
                    mainForm.BringToFront();
                }

                _selectedMPR_ID = 0;
                ClearHeader();
                txtMPRNo.Focus();
            }
        }

        // ── Popup Tạo MPR mới ─────────────────────────────────────────────────
        private void BtnCreateMPR_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("MPR", "Tạo MPR", "Tạo MPR")) return;
            ShowCreateMPRPopup();
        }

        private void ShowCreateMPRPopup()
        {
            var dlg = new Form
            {
                Text = "➕ Tạo MPR mới",
                Size = new Size(1100, 680),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.Sizable,
                MinimizeBox = false,
                BackColor = Color.FromArgb(245, 245, 245)
            };

            // ── Bảng tìm dự án (trái) ──────────────────────────────────────────
            dlg.Controls.Add(new Label { Text = "🔍 Tìm dự án:", Location = new Point(10, 10), Size = new Size(90, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold) });
            var txtSearch = new TextBox { Location = new Point(102, 8), Size = new Size(180, 24), Font = new Font("Segoe UI", 9), PlaceholderText = "Mã/tên dự án..." };
            dlg.Controls.Add(txtSearch);

            var dgvProj = new DataGridView
            {
                Location = new Point(10, 38),
                Size = new Size(272, 200),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            dgvProj.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvProj.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvProj.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvProj.EnableHeadersVisualStyles = false;
            dgvProj.Columns.Add(new DataGridViewTextBoxColumn { Name = "ProjCode", HeaderText = "Mã DA", Width = 100 });
            dgvProj.Columns.Add(new DataGridViewTextBoxColumn { Name = "ProjName", HeaderText = "Tên DA", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            dlg.Controls.Add(dgvProj);

            // Load và lọc projects
            var allProjects = new List<MPR_Managerment.Models.ProjectInfo>();
            try { allProjects = new ProjectService().GetAll(); } catch { }
            var projDict = new Dictionary<int, MPR_Managerment.Models.ProjectInfo>();

            void FilterProjects()
            {
                string kw = txtSearch.Text.Trim().ToLower();
                dgvProj.Rows.Clear();
                projDict.Clear();
                foreach (var p in allProjects)
                {
                    if (string.IsNullOrEmpty(kw)
                        || (p.ProjectCode ?? "").ToLower().Contains(kw)
                        || (p.ProjectName ?? "").ToLower().Contains(kw))
                    {
                        int r = dgvProj.Rows.Add();
                        dgvProj.Rows[r].Cells["ProjCode"].Value = p.ProjectCode;
                        dgvProj.Rows[r].Cells["ProjName"].Value = p.ProjectName;
                        projDict[r] = p;
                    }
                }
            }
            FilterProjects();
            // Chỉ lọc khi nhấn Enter
            // ApplyProject sẽ được gọi qua SelectionChanged (định nghĩa sau)
            txtSearch.KeyDown += (s, ev) =>
            {
                if (ev.KeyCode != Keys.Enter) return;
                ev.SuppressKeyPress = true;
                FilterProjects();
                // Nếu đúng 1 kết quả → chọn ngay, SelectionChanged tự gọi ApplyProject
                if (dgvProj.Rows.Count == 1)
                {
                    dgvProj.ClearSelection();
                    dgvProj.Rows[0].Selected = true;
                    dgvProj.CurrentCell = dgvProj.Rows[0].Cells[0];
                }
                // Nếu nhiều kết quả → chọn row đầu tiên để user thấy ngay
                else if (dgvProj.Rows.Count > 1)
                {
                    dgvProj.ClearSelection();
                    dgvProj.Rows[0].Selected = true;
                    dgvProj.CurrentCell = dgvProj.Rows[0].Cells[0];
                }
            };

            // ── Thông tin dự án (chỉ đọc) ──────────────────────────────────────
            int xInfo = 292;
            dlg.Controls.Add(new Label { Text = "THÔNG TIN DỰ ÁN", Location = new Point(xInfo, 10), Size = new Size(300, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212) });

            Label LblF(string t2, int y) => new Label { Text = t2, Location = new Point(xInfo, y), Size = new Size(105, 20), Font = new Font("Segoe UI", 9) };
            TextBox TxtR(int y) => new TextBox { Location = new Point(xInfo + 108, y), Size = new Size(280, 24), Font = new Font("Segoe UI", 9), ReadOnly = true, BackColor = Color.FromArgb(240, 240, 240) };

            var tProjCode = TxtR(34); var tProjName = TxtR(62); var tDept = TxtR(90); var tReq = TxtR(118);
            dlg.Controls.Add(LblF("Mã dự án:", 34)); dlg.Controls.Add(tProjCode);
            dlg.Controls.Add(LblF("Tên dự án:", 62)); dlg.Controls.Add(tProjName);
            dlg.Controls.Add(LblF("Department:", 90)); dlg.Controls.Add(tDept);
            dlg.Controls.Add(LblF("Requestor:", 118)); dlg.Controls.Add(tReq);

            // MPR No (có thể chỉnh sửa)
            dlg.Controls.Add(new Label { Text = "MPR No (*):", Location = new Point(xInfo, 148), Size = new Size(105, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold) });
            var tMPRNo = new TextBox { Location = new Point(xInfo + 108, 146), Size = new Size(280, 24), Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            dlg.Controls.Add(tMPRNo);

            dlg.Controls.Add(new Label { Text = "Required Date:", Location = new Point(xInfo, 176), Size = new Size(105, 20), Font = new Font("Segoe UI", 9) });
            var dtp = new DateTimePicker { Location = new Point(xInfo + 108, 174), Size = new Size(180, 24), Font = new Font("Segoe UI", 9), Value = DateTime.Today.AddDays(30) };
            dlg.Controls.Add(dtp);

            dlg.Controls.Add(new Label { Text = "Notes:", Location = new Point(xInfo, 204), Size = new Size(105, 20), Font = new Font("Segoe UI", 9) });
            var tNotes = new TextBox { Location = new Point(xInfo + 108, 202), Size = new Size(280, 24), Font = new Font("Segoe UI", 9) };
            dlg.Controls.Add(tNotes);

            // Hàm điền thông tin dự án được chọn
            void ApplyProject(MPR_Managerment.Models.ProjectInfo prj)
            {
                if (prj == null) return;
                tProjCode.Text = prj.ProjectCode ?? "";
                tProjName.Text = prj.ProjectName ?? "";
                tDept.Text = "";
                tReq.Text = "";
                // Tính MPR No tiếp theo: MPRCode-001, MPRCode-002, ...
                try
                {
                    string mprPrefix = (prj.MPRCode ?? prj.ProjectCode ?? "").TrimEnd('-');
                    int nextSeq = 1;
                    using var conn = DatabaseHelper.GetConnection();
                    conn.Open();
                    var cmd = new SqlCommand(
                        "SELECT MAX(CAST(SUBSTRING(MPR_No, LEN(@prefix)+2, 10) AS INT)) FROM MPR_Header WHERE MPR_No LIKE @like",
                        conn);
                    cmd.Parameters.AddWithValue("@prefix", mprPrefix);
                    cmd.Parameters.AddWithValue("@like", mprPrefix + "-%");
                    var res = cmd.ExecuteScalar();
                    if (res != DBNull.Value && res != null) nextSeq = Convert.ToInt32(res) + 1;
                    tMPRNo.Text = $"{mprPrefix}-{nextSeq:D3}";
                }
                catch { tMPRNo.Text = (prj.MPRCode ?? prj.ProjectCode ?? "").TrimEnd('-') + "-001"; }
            }
            // Click chuột chọn row
            dgvProj.CellClick += (s, ev) =>
            {
                if (ev.RowIndex < 0) return;
                if (projDict.TryGetValue(ev.RowIndex, out var prj2)) ApplyProject(prj2);
            };
            // Dùng bàn phím
            dgvProj.SelectionChanged += (s, ev) =>
            {
                if (dgvProj.SelectedRows.Count == 0) return;
                int ri = dgvProj.SelectedRows[0].Index;
                if (projDict.TryGetValue(ri, out var prj3)) ApplyProject(prj3);
            };

            // ── Bảng MPR Details ────────────────────────────────────────────────
            dlg.Controls.Add(new Label { Text = "CHI TIẾT VẬT TƯ (MPR_Details)", Location = new Point(10, 248), Size = new Size(400, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212) });

            var dgvDet = new DataGridView
            {
                Location = new Point(10, 270),
                Size = new Size(dlg.ClientSize.Width - 20, 300),
                AllowUserToAddRows = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 8),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText
            };
            dgvDet.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvDet.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvDet.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 8, FontStyle.Bold);
            dgvDet.EnableHeadersVisualStyles = false;

            dgvDet.CellDoubleClick += (e, s) =>
            {
                frmSelectItem frmSelectItem = new frmSelectItem();
                frmSelectItem.ShowDialog();

                if (frmSelectItem.selectedItems.Count <= 0) return;
                int startRow = dgvDet.CurrentCell?.RowIndex ?? 0;
                foreach (var item in frmSelectItem.selectedItems)
                {
                    dgvDet.Rows.Add();
                    dgvDet.Rows[startRow].Cells["item_name"].Value = item.Name;
                    dgvDet.Rows[startRow].Cells["Description"].Value = item.Des2;
                    dgvDet.Rows[startRow].Cells["Material"].Value = item.ProdMaterialCode;
                    dgvDet.Rows[startRow].Cells["Thickness_mm"].Value = item.A_Thickness;
                    dgvDet.Rows[startRow].Cells["Depth_mm"].Value = item.B_Depth;
                    dgvDet.Rows[startRow].Cells["C_Width_mm"].Value = item.C_Width;
                    dgvDet.Rows[startRow].Cells["D_Web_mm"].Value = item.D_Web;
                    dgvDet.Rows[startRow].Cells["E_Flange_mm"].Value = item.E_Flag;
                    dgvDet.Rows[startRow].Cells["F_Length_mm"].Value = item.F_Length;
                    dgvDet.Rows[startRow].Cells["UNIT"].Value = "";
                    dgvDet.Rows[startRow].Cells["Weight_kg"].Value = item.G_Weight;

                    startRow++;
                }
            };

            dgvDet.ColumnHeaderMouseClick += (s, e) =>
            {
                if (e.ColumnIndex < 0) return;
                string colName = dgvDet.Columns[e.ColumnIndex].Name;
                if (colName == "item_name")
                {
                    int rowIndex = dgvDet.Rows.Add();

                    // 2. Tùy chọn: Focus vào ô đầu tiên của dòng mới để người dùng nhập liệu ngay
                    dgvDet.CurrentCell = dgvDet.Rows[rowIndex].Cells[0];
                    dgvDet.BeginEdit(true);
                }
            };

            // Các cột theo MPR_Details
            void AddCol(string name, string hdr, int w, bool ro = false)
            {
                dgvDet.Columns.Add(new DataGridViewTextBoxColumn
                {
                    Name = name,
                    HeaderText = hdr,
                    Width = w,
                    ReadOnly = ro,
                    DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleLeft }
                });
            }
            AddCol("Item_No", "STT", 40);
            AddCol("item_name", "Tên hàng", 140);
            AddCol("Description", "Mô tả", 120);
            AddCol("Material", "Vật liệu", 80);
            AddCol("Thickness_mm", "T(mm)", 55);
            AddCol("Depth_mm", "D(mm)", 55);
            AddCol("C_Width_mm", "W(mm)", 55);
            AddCol("D_Web_mm", "Web(mm)", 60);
            AddCol("E_Flange_mm", "Flange(mm)", 70);
            AddCol("F_Length_mm", "L(mm)", 60);
            AddCol("Usage_Location", "Vị trí", 90);
            AddCol("MPS_Info", "MPS", 70);
            AddCol("REV", "REV", 45);
            AddCol("DWG_BOQ_Receive_Date", "DWG Date", 80);
            AddCol("Issue_Date", "Issue Date", 80);
            AddCol("UNIT", "Đơn vị", 55);
            AddCol("Qty_Per_Sheet", "SL", 45);
            AddCol("Weight_kg", "KG", 55);
            AddCol("Remarks", "Ghi chú", 100);
            dlg.Controls.Add(dgvDet);

            // Paste từ Excel — 1 lần, xử lý đúng new row placeholder
            dgvDet.KeyDown += (s, ev) =>
            {
                if (ev.Control && ev.KeyCode == Keys.V)
                {
                    ev.Handled = true;
                    ev.SuppressKeyPress = true;
                    string clip = Clipboard.GetText();
                    if (string.IsNullOrEmpty(clip)) return;
                    dgvDet.EndEdit();
                    string[] rows2 = clip.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    if (rows2.Length == 0) return;

                    // Xác định ô bắt đầu — nếu CurrentCell là new row hoặc null thì về row 0
                    int startCol = dgvDet.CurrentCell?.ColumnIndex ?? 0;
                    int curRow = dgvDet.CurrentCell?.RowIndex ?? 0;
                    // Số dòng data thực (không tính new row placeholder)
                    int dataRowCount = dgvDet.Rows.Count - (dgvDet.AllowUserToAddRows ? 1 : 0);
                    // Nếu đang đứng ở new row placeholder → paste từ cuối data
                    int startRow = (curRow >= dataRowCount) ? dataRowCount : curRow;

                    dgvDet.SuspendLayout();
                    foreach (var row2 in rows2)
                    {
                        string[] cells = row2.Split('\t');
                        // Thêm dòng mới nếu cần (trước new row placeholder)
                        int dataRows = dgvDet.Rows.Count - (dgvDet.AllowUserToAddRows ? 1 : 0);
                        if (startRow >= dataRows)
                        {
                            dgvDet.Rows.Insert(dataRows, 1);
                        }
                        for (int c = 0; c < cells.Length && startCol + c < dgvDet.Columns.Count; c++)
                            if (!dgvDet.Columns[startCol + c].ReadOnly)
                                dgvDet.Rows[startRow].Cells[startCol + c].Value = cells[c].Trim();
                        startRow++;
                    }
                    dgvDet.ResumeLayout();
                    dgvDet.Refresh();
                }
            };

            // ── Buttons ─────────────────────────────────────────────────────────
            var lblErr2 = new Label { Location = new Point(10, dlg.ClientSize.Height - 78), Size = new Size(500, 20), ForeColor = Color.Red, Font = new Font("Segoe UI", 8), Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            dlg.Controls.Add(lblErr2);

            var btnCreate = new Button
            {
                Text = "✅ Tạo MPR",
                Location = new Point(dlg.ClientSize.Width - 460, dlg.ClientSize.Height - 50),
                Size = new Size(110, 34),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            btnCreate.FlatAppearance.BorderSize = 0;

            var btnRevise = new Button
            {
                Text = "🔄 Revise MPR",
                Location = new Point(dlg.ClientSize.Width - 342, dlg.ClientSize.Height - 50),
                Size = new Size(125, 34),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            btnRevise.FlatAppearance.BorderSize = 0;

            var btnAdmin = new Button
            {
                Text = "🔑 For Admin",
                Location = new Point(dlg.ClientSize.Width - 210, dlg.ClientSize.Height - 50),
                Size = new Size(110, 34),
                BackColor = Color.FromArgb(128, 0, 128),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Visible = AppSession.IsAdmin
            };
            btnAdmin.FlatAppearance.BorderSize = 0;

            var btnClose2 = new Button
            {
                Text = "Đóng",
                Location = new Point(dlg.ClientSize.Width - 92, dlg.ClientSize.Height - 50),
                Size = new Size(80, 34),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                DialogResult = DialogResult.Cancel
            };
            btnClose2.FlatAppearance.BorderSize = 0;
            dlg.Controls.AddRange(new Control[] { btnCreate, btnRevise, btnAdmin, btnClose2 });
            dlg.CancelButton = btnClose2;

            // Resize handler
            dlg.Resize += (s, ev) =>
            {
                dgvDet.Width = dlg.ClientSize.Width - 20;
                dgvDet.Height = dlg.ClientSize.Height - 270 - 62;
                lblErr2.Top = dlg.ClientSize.Height - 78;
            };

            // ── Tạo MPR ─────────────────────────────────────────────────────────
            btnCreate.Click += (s, ev) =>
            {
                if (string.IsNullOrWhiteSpace(tMPRNo.Text)) { lblErr2.Text = "⚠ Nhập MPR No!"; return; }
                if (string.IsNullOrWhiteSpace(tProjCode.Text)) { lblErr2.Text = "⚠ Chọn dự án!"; return; }
                // Kiểm tra MPR_No đã tồn tại chưa
                try
                {
                    using var conn = DatabaseHelper.GetConnection();
                    conn.Open();
                    var chk = new SqlCommand("SELECT COUNT(1) FROM MPR_Header WHERE MPR_No=@no", conn);
                    chk.Parameters.AddWithValue("@no", tMPRNo.Text.Trim());
                    if (Convert.ToInt32(chk.ExecuteScalar()) > 0)
                    { lblErr2.Text = "❌ MPR No đã tồn tại! Vui lòng đổi số MPR."; return; }
                }
                catch (Exception ex) { lblErr2.Text = "❌ " + ex.Message; return; }

                try
                {
                    var header = new MPRHeader
                    {
                        MPR_ID = 0,
                        MPR_No = tMPRNo.Text.Trim(),
                        Project_Name = tProjName.Text.Trim(),
                        Project_Code = tProjCode.Text.Trim(),
                        Department = tDept.Text.Trim(),
                        Requestor = tReq.Text.Trim(),
                        Required_Date = dtp.Value,
                        Rev = 0,
                        Status = "Mới",
                        Notes = tNotes.Text.Trim()
                    };
                    int newId = _service.InsertHeader(header, _currentUser);

                    // Lưu details
                    // Insert details dùng MPRService.InsertDetail
                    int stt = 1;
                    foreach (DataGridViewRow row2 in dgvDet.Rows)
                    {
                        if (row2.IsNewRow) continue;
                        string nm = row2.Cells["item_name"].Value?.ToString() ?? "";
                        if (string.IsNullOrWhiteSpace(nm)) continue;
                        _service.InsertDetail(new MPRDetail
                        {
                            MPR_ID = newId,
                            Item_No = stt++,
                            Item_Name = nm,
                            Description = row2.Cells["Description"].Value?.ToString() ?? "",
                            Material = row2.Cells["Material"].Value?.ToString() ?? "",
                            Thickness_mm = decimal.TryParse(row2.Cells["Thickness_mm"].Value?.ToString(), out decimal th) ? th : 0,
                            Depth_mm = decimal.TryParse(row2.Cells["Depth_mm"].Value?.ToString(), out decimal dp) ? dp : 0,
                            C_Width_mm = decimal.TryParse(row2.Cells["C_Width_mm"].Value?.ToString(), out decimal cw) ? cw : 0,
                            D_Web_mm = decimal.TryParse(row2.Cells["D_Web_mm"].Value?.ToString(), out decimal dw) ? dw : 0,
                            E_Flange_mm = decimal.TryParse(row2.Cells["E_Flange_mm"].Value?.ToString(), out decimal ef) ? ef : 0,
                            F_Length_mm = decimal.TryParse(row2.Cells["F_Length_mm"].Value?.ToString(), out decimal fl) ? fl : 0,
                            UNIT = row2.Cells["UNIT"].Value?.ToString() ?? "",
                            Qty_Per_Sheet = decimal.TryParse(row2.Cells["Qty_Per_Sheet"].Value?.ToString(), out decimal qs) ? qs : 0,
                            Weight_kg = decimal.TryParse(row2.Cells["Weight_kg"].Value?.ToString(), out decimal wk) ? wk : 0,
                            MPS_Info = row2.Cells["MPS_Info"].Value?.ToString() ?? "",
                            Usage_Location = row2.Cells["Usage_Location"].Value?.ToString() ?? "",
                            REV = "0",
                            Remarks = row2.Cells["Remarks"].Value?.ToString() ?? ""
                        }, _currentUser);
                    }

                    MessageBox.Show($"✅ Tạo MPR '{header.MPR_No}' thành công ({stt - 1} vật tư)!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    dlg.DialogResult = DialogResult.OK;
                    dlg.Close();
                    LoadMPR();
                    _selectedMPR_ID = newId;
                    foreach (DataGridViewRow row2 in dgvMPR.Rows)
                        if (Convert.ToInt32(row2.Cells["MPR_ID"]?.Value ?? 0) == newId)
                        { row2.Selected = true; dgvMPR.CurrentCell = row2.Cells[1]; break; }
                }
                catch (Exception ex) { lblErr2.Text = "❌ " + ex.Message; }
            };

            // ── Revise MPR ───────────────────────────────────────────────────────
            btnRevise.Click += (s, ev) =>
            {
                if (string.IsNullOrWhiteSpace(tProjCode.Text)) { lblErr2.Text = "⚠ Chọn dự án trước!"; return; }
                ShowReviseMPRPopup(tProjCode.Text, tProjName.Text, false);
            };

            // ── For Admin ────────────────────────────────────────────────────────
            btnAdmin.Click += (s, ev) =>
            {
                if (string.IsNullOrWhiteSpace(tProjCode.Text)) { lblErr2.Text = "⚠ Chọn dự án trước!"; return; }
                ShowReviseMPRPopup(tProjCode.Text, tProjName.Text, true);
            };

            dlg.Owner = this.FindForm();
            dlg.ShowDialog();
        }

        public void ExportMPRToExcel(MPRHeader header, List<MPRDetail> details)
        {
            // 1. Kiểm tra template
            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "mpr_template.xlsx");
            if (!File.Exists(templatePath))
            {
                MessageBox.Show("Không tìm thấy file template mpr_template.xlsx!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 2. Chọn nơi lưu file
            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"MPR_{header.MPR_No}_{DateTime.Now:ddMMyyyy_HHmm}.xlsx",
                Title = "Lưu file MPR"
            };

            if (saveDialog.ShowDialog() != DialogResult.OK) return;

            try
            {
                // Copy từ template
                File.Copy(templatePath, saveDialog.FileName, true);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(saveDialog.FileName)))
                {
                    var ws = package.Workbook.Worksheets[0]; // Giả định dữ liệu ở Sheet đầu tiên

                    // --- PHẦN 1: THAY THẾ HEADER <<...>> ---
                    // Dựa vào snippet của bạn, các giá trị này nằm ở các dòng đầu (dòng 1 và 2)
                    ReplaceText(ws, "<<MPR-NO>>", header.MPR_No);
                    ReplaceText(ws, "<<PROJECT-NAME>>", header.Project_Name);
                    ReplaceText(ws, "<<WO-NO>>", ""); // Hoặc WO_No của bạn
                    ReplaceText(ws, "<<REV>>", header.Rev.ToString() ?? "0");
                    ReplaceText(ws, "<<DATE>>", DateTime.Now.ToString("dd/MM/yyyy"));

                    // --- PHẦN 2: ĐIỀN DỮ LIỆU CHI TIẾT ---
                    int startRow = 4; // Dòng bắt đầu điền item đầu tiên (Dưới tiêu đề cột)
                    int detailCount = details.Count;

                    // Nếu có nhiều hơn 1 dòng, chèn thêm dòng để giữ định dạng và không đè phần Footer
                    if (detailCount > 1)
                    {
                        // Chèn thêm (n-1) dòng, copy format từ dòng startRow
                        ws.InsertRow(startRow + 1, detailCount - 1, startRow);
                    }

                    decimal totalQty = 0;
                    decimal totalWeight = 0;

                    for (int i = 0; i < detailCount; i++)
                    {
                        var d = details[i];
                        int currentRow = startRow + i;

                        // Cột dựa theo cấu trúc: NO(A), DESCRIPTION(B,C), MATERIAL(D), A(E), B(F)...
                        ws.Cells[currentRow, 1].Value = i + 1; // NO
                        ws.Cells[currentRow, 2].Value = d.Item_Name; // Cột B
                        ws.Cells[currentRow, 3].Value = d.Description; // Cột C (nếu có)
                        ws.Cells[currentRow, 4].Value = d.Material; // MATERIAL
                        ws.Cells[currentRow, 5].Value = d.Thickness_mm; // A
                        ws.Cells[currentRow, 6].Value = d.Depth_mm; // B
                        ws.Cells[currentRow, 7].Value = d.C_Width_mm; // C
                        ws.Cells[currentRow, 8].Value = d.D_Web_mm; // D
                        ws.Cells[currentRow, 9].Value = d.E_Flange_mm; // E
                        ws.Cells[currentRow, 10].Value = d.F_Length_mm; // F

                        ws.Cells[currentRow, 11].Value = d.Usage_Location; // F
                        ws.Cells[currentRow, 12].Value = d.MPS_Info; // F
                        ws.Cells[currentRow, 13].Value = d.REV; // F
                        ws.Cells[currentRow, 14].Value = d.DWG_BOQ_Receive_Date; // F
                        ws.Cells[currentRow, 15].Value = d.Issue_Date; // F

                        // Các cột khác...
                        ws.Cells[currentRow, 16].Value = d.UNIT; // UNIT (Cột P)
                        ws.Cells[currentRow, 17].Value = d.Qty_Per_Sheet; // Q'ty / Sh't (Cột Q)
                        ws.Cells[currentRow, 18].Value = d.Weight_kg; // Weight(kg) (Cột R)

                        totalQty += d.Qty_Per_Sheet > 0 ? d.Qty_Per_Sheet : 0;
                        totalWeight += d.Weight_kg > 0 ? d.Weight_kg : 0;
                    }

                    // --- PHẦN 3: TÍNH TỔNG (GRAND-TOTAL) ---
                    // Tìm dòng có chữ "GRAND-TOTAL" để điền tổng
                    int grandTotalRow = -1;
                    // Tìm trong phạm vi 10 dòng sau khi kết thúc dữ liệu
                    for (int r = startRow + detailCount; r < startRow + detailCount + 10; r++)
                    {
                        if (ws.Cells[r, 2].Text.Contains("GRAND-TOTAL"))
                        {
                            grandTotalRow = r;
                            break;
                        }
                    }

                    if (grandTotalRow != -1)
                    {
                        ws.Cells[grandTotalRow, 17].Value = totalQty; // Cột Q
                        ws.Cells[grandTotalRow, 18].Value = totalWeight; // Cột R
                        ws.Cells[grandTotalRow, 17, grandTotalRow, 18].Style.Font.Bold = true;
                    }

                    package.Save();
                }

                if (MessageBox.Show($"✅ Xuất phiếu MPR thành công!\nBạn có muốn mở file không?", "Thành công",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(saveDialog.FileName) { UseShellExecute = true });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xuất Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Hàm bổ trợ thay thế text trong Header
        private void ReplaceText(ExcelWorksheet ws, string findText, string replaceValue)
        {
            // Quét vùng Header (ví dụ dòng 1 đến 3)
            for (int r = 1; r <= 5; r++)
            {
                for (int c = 1; c <= 20; c++)
                {
                    if (ws.Cells[r, c].Text.Contains(findText))
                    {
                        ws.Cells[r, c].Value = ws.Cells[r, c].Text.Replace(findText, replaceValue ?? "");
                    }
                }
            }
        }

        // ── Revise MPR Popup ─────────────────────────────────────────────────────
        private void ShowReviseMPRPopup(string projCode, string projName, bool isAdmin)
        {
            var dlg = new Form
            {
                Text = isAdmin ? "🔑 Admin Edit MPR" : "🔄 Revise MPR",
                Size = new Size(1100, 700),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.Sizable,
                MinimizeBox = false,
                BackColor = Color.FromArgb(245, 245, 245)
            };

            // Danh sách MPR của dự án
            dlg.Controls.Add(new Label { Text = $"MPR của dự án: {projCode}", Location = new Point(10, 8), Size = new Size(400, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212) });
            var dgvMPRList = new DataGridView
            {
                Location = new Point(10, 32),
                Size = new Size(260, 580),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom
            };
            dgvMPRList.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvMPRList.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvMPRList.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvMPRList.EnableHeadersVisualStyles = false;
            dgvMPRList.Columns.Add(new DataGridViewTextBoxColumn { Name = "RMprId", Visible = false });
            dgvMPRList.Columns.Add(new DataGridViewTextBoxColumn { Name = "RMprNo", HeaderText = "MPR No", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            dgvMPRList.Columns.Add(new DataGridViewTextBoxColumn { Name = "RMprRev", HeaderText = "Rev", Width = 45 });
            dlg.Controls.Add(dgvMPRList);

            // Load MPR theo project
            try
            {
                using var conn = DatabaseHelper.GetConnection();
                conn.Open();
                var cmd = new SqlCommand("SELECT MPR_ID, MPR_No, Rev FROM MPR_Header WHERE Project_Code=@code ORDER BY MPR_No", conn);
                cmd.Parameters.AddWithValue("@code", projCode);
                using var rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    int r = dgvMPRList.Rows.Add();
                    dgvMPRList.Rows[r].Cells["RMprId"].Value = rdr["MPR_ID"];
                    dgvMPRList.Rows[r].Cells["RMprNo"].Value = rdr["MPR_No"];
                    dgvMPRList.Rows[r].Cells["RMprRev"].Value = rdr["Rev"];
                }
            }
            catch { }

            // Bảng chi tiết
            dlg.Controls.Add(new Label { Text = "CHI TIẾT VẬT TƯ", Location = new Point(278, 8), Size = new Size(400, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212) });
            var dgvRevDet = new DataGridView
            {
                Location = new Point(278, 32),
                Size = new Size(dlg.ClientSize.Width - 288, 530),
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 8),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2
            };
            dgvRevDet.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvRevDet.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvRevDet.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 8, FontStyle.Bold);
            dgvRevDet.EnableHeadersVisualStyles = false;
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RDetId", Visible = false });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RDeleted", Visible = false }); // "1"=xóa mềm
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RItem_No", HeaderText = "STT", Width = 40, ReadOnly = true });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ritem_name", HeaderText = "Tên hàng", Width = 140 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RDesc", HeaderText = "Mô tả", Width = 110 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RMaterial", HeaderText = "Vật liệu", Width = 80 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RT_mm", HeaderText = "T(mm)", Width = 50 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RD_mm", HeaderText = "D(mm)", Width = 50 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RW_mm", HeaderText = "W(mm)", Width = 50 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RWeb_mm", HeaderText = "Web", Width = 50 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RFlange_mm", HeaderText = "Flange", Width = 55 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RL_mm", HeaderText = "L(mm)", Width = 55 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RUNIT", HeaderText = "ĐVT", Width = 50 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RQty", HeaderText = "SL", Width = 45 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RKG", HeaderText = "KG", Width = 55 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RRemarks", HeaderText = "Ghi chú", Width = 100 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RREV", HeaderText = "REV", Width = 40, ReadOnly = true });
            dlg.Controls.Add(dgvRevDet);

            // Hiệu ứng: dòng xóa mềm = xám gạch ngang
            dgvRevDet.CellFormatting += (s2, ev2) =>
            {
                if (ev2.RowIndex < 0) return;
                string del = dgvRevDet.Rows[ev2.RowIndex].Cells["RDeleted"].Value?.ToString() ?? "";
                if (del == "1") { ev2.CellStyle.ForeColor = Color.FromArgb(180, 180, 180); ev2.CellStyle.Font = new Font("Segoe UI", 8, FontStyle.Strikeout); }
            };

            int selMprId = 0;
            // Khi chọn MPR → load details
            dgvMPRList.SelectionChanged += (s2, ev2) =>
            {
                if (dgvMPRList.SelectedRows.Count == 0) return;
                selMprId = Convert.ToInt32(dgvMPRList.SelectedRows[0].Cells["RMprId"].Value ?? 0);
                dgvRevDet.Rows.Clear();
                try
                {
                    using var conn = DatabaseHelper.GetConnection();
                    conn.Open();
                    var cmd = new SqlCommand("SELECT * FROM MPR_Details WHERE MPR_ID=@id ORDER BY Item_No", conn);
                    cmd.Parameters.AddWithValue("@id", selMprId);
                    using var rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        int r = dgvRevDet.Rows.Add();
                        dgvRevDet.Rows[r].Cells["RDetId"].Value = rdr["Detail_ID"];
                        dgvRevDet.Rows[r].Cells["RDeleted"].Value = "";
                        dgvRevDet.Rows[r].Cells["RItem_No"].Value = rdr["Item_No"];
                        dgvRevDet.Rows[r].Cells["Ritem_name"].Value = rdr["item_name"];
                        dgvRevDet.Rows[r].Cells["RDesc"].Value = rdr["Description"];
                        dgvRevDet.Rows[r].Cells["RMaterial"].Value = rdr["Material"];
                        dgvRevDet.Rows[r].Cells["RT_mm"].Value = rdr["Thickness_mm"];
                        dgvRevDet.Rows[r].Cells["RD_mm"].Value = rdr["Depth_mm"];
                        dgvRevDet.Rows[r].Cells["RW_mm"].Value = rdr["C_Width_mm"];
                        dgvRevDet.Rows[r].Cells["RWeb_mm"].Value = rdr["D_Web_mm"];
                        dgvRevDet.Rows[r].Cells["RFlange_mm"].Value = rdr["E_Flange_mm"];
                        dgvRevDet.Rows[r].Cells["RL_mm"].Value = rdr["F_Length_mm"];
                        dgvRevDet.Rows[r].Cells["RUNIT"].Value = rdr["UNIT"];
                        dgvRevDet.Rows[r].Cells["RQty"].Value = rdr["Qty_Per_Sheet"];
                        dgvRevDet.Rows[r].Cells["RKG"].Value = rdr["Weight_kg"];
                        dgvRevDet.Rows[r].Cells["RRemarks"].Value = rdr["Remarks"];
                        dgvRevDet.Rows[r].Cells["RREV"].Value = rdr["REV"];
                    }
                }
                catch { }
            };

            // Buttons Thêm / Xóa vật tư
            var btnAddRow = new Button
            {
                Text = "+ Thêm vật tư",
                Location = new Point(278, dlg.ClientSize.Height - 58),
                Size = new Size(120, 30),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            btnAddRow.FlatAppearance.BorderSize = 0;
            btnAddRow.Click += (s2, ev2) => { dgvRevDet.AllowUserToAddRows = true; dgvRevDet.Rows.Add(); dgvRevDet.AllowUserToAddRows = false; };

            var btnDelRow = new Button
            {
                Text = "🗑 Xóa vật tư",
                Location = new Point(406, dlg.ClientSize.Height - 58),
                Size = new Size(120, 30),
                BackColor = Color.FromArgb(220, 53, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            btnDelRow.FlatAppearance.BorderSize = 0;
            // Xóa mềm: đánh dấu "1" vào cột RDeleted
            btnDelRow.Click += (s2, ev2) =>
            {
                foreach (DataGridViewRow row2 in dgvRevDet.SelectedRows)
                {
                    if (row2.IsNewRow) continue;
                    string cur = row2.Cells["RDeleted"].Value?.ToString() ?? "";
                    row2.Cells["RDeleted"].Value = cur == "1" ? "" : "1"; // toggle
                }
                dgvRevDet.Refresh();
            };

            var lblSave = new Label { Location = new Point(10, dlg.ClientSize.Height - 30), Size = new Size(500, 22), ForeColor = Color.Red, Font = new Font("Segoe UI", 8), Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            dlg.Controls.Add(lblSave);

            var btnSaveMPR = new Button
            {
                Text = "💾 Lưu MPR",
                Location = new Point(dlg.ClientSize.Width - 312, dlg.ClientSize.Height - 58),
                Size = new Size(120, 30),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            btnSaveMPR.FlatAppearance.BorderSize = 0;

            var btnCloseRev = new Button
            {
                Text = "Đóng",
                Location = new Point(dlg.ClientSize.Width - 184, dlg.ClientSize.Height - 58),
                Size = new Size(80, 30),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                DialogResult = DialogResult.Cancel
            };
            btnCloseRev.FlatAppearance.BorderSize = 0;
            dlg.Controls.AddRange(new Control[] { btnAddRow, btnDelRow, btnSaveMPR, btnCloseRev });
            dlg.CancelButton = btnCloseRev;

            dlg.Resize += (s2, ev2) =>
            {
                int btnY = dlg.ClientSize.Height - 50;
                int gridH = Math.Max(80, btnY - 36 - 8);
                dgvMPRList.Height = gridH;
                dgvRevDet.Width = dlg.ClientSize.Width - 288;
                dgvRevDet.Height = gridH;
                btnAddRow.Top = btnDelRow.Top = btnSaveMPR.Top = btnCloseRev.Top = btnY;
                btnSaveMPR.Left = dlg.ClientSize.Width - 312;
                btnCloseRev.Left = dlg.ClientSize.Width - 184;
                lblSave.Top = dlg.ClientSize.Height - 26;
            };

            // Lưu MPR Revise
            btnSaveMPR.Click += (s2, ev2) =>
            {
                if (selMprId == 0) { lblSave.Text = "⚠ Chọn MPR cần Revise!"; return; }
                try
                {
                    using var conn = DatabaseHelper.GetConnection();
                    conn.Open();
                    // Tính REV max hiện tại
                    var cmdMaxRev = new SqlCommand("SELECT ISNULL(MAX(REV),0) FROM MPR_Details WHERE MPR_ID=@id", conn);
                    cmdMaxRev.Parameters.AddWithValue("@id", selMprId);
                    int maxRev = Convert.ToInt32(cmdMaxRev.ExecuteScalar());
                    int nextRev = isAdmin ? maxRev : maxRev + 1;

                    // MPR_No mới
                    string oldMprNo = dgvMPRList.SelectedRows[0].Cells["RMprNo"].Value?.ToString() ?? "";
                    string baseMprNo = oldMprNo.Contains("_Rev.")
                        ? oldMprNo.Substring(0, oldMprNo.IndexOf("_Rev."))
                        : oldMprNo;
                    string newMprNo = isAdmin ? oldMprNo : $"{baseMprNo}_Rev.{nextRev}";

                    // Nếu không phải Admin → tạo bản Revise mới
                    int targetId = selMprId;
                    if (!isAdmin)
                    {
                        // Lấy header cũ
                        var cmdH = new SqlCommand("SELECT * FROM MPR_Header WHERE MPR_ID=@id", conn);
                        cmdH.Parameters.AddWithValue("@id", selMprId);
                        using var rdrH = cmdH.ExecuteReader();
                        rdrH.Read();
                        var newHeader = new MPRHeader
                        {
                            MPR_ID = 0,
                            MPR_No = newMprNo,
                            Project_Name = rdrH["Project_Name"]?.ToString() ?? "",
                            Project_Code = rdrH["Project_Code"]?.ToString() ?? "",
                            Department = rdrH["Department"]?.ToString() ?? "",
                            Requestor = rdrH["Requestor"]?.ToString() ?? "",
                            Required_Date = rdrH["Required_Date"] is DateTime dt ? dt : DateTime.Today,
                            Rev = nextRev,
                            Status = rdrH["Status"]?.ToString() ?? "Mới",
                            Notes = rdrH["Notes"]?.ToString() ?? ""
                        };
                        rdrH.Close();
                        targetId = _service.InsertHeader(newHeader, _currentUser);
                    }
                    else
                    {
                        // Admin: xóa toàn bộ details cũ (hard delete) rồi insert mới
                        var cmdDel = new SqlCommand("DELETE FROM MPR_Details WHERE MPR_ID=@id", conn);
                        cmdDel.Parameters.AddWithValue("@id", targetId);
                        cmdDel.ExecuteNonQuery();
                    }

                    // Insert details (bỏ dòng xóa mềm)
                    int stt = 1;
                    foreach (DataGridViewRow row2 in dgvRevDet.Rows)
                    {
                        if (row2.IsNewRow) continue;
                        if (row2.Cells["RDeleted"].Value?.ToString() == "1") continue;
                        string nm = row2.Cells["Ritem_name"].Value?.ToString() ?? "";
                        if (string.IsNullOrWhiteSpace(nm)) continue;
                        int detRev = isAdmin
                            ? (int.TryParse(row2.Cells["RREV"].Value?.ToString(), out int rv) ? rv : 0)
                            : nextRev; // REV field is string, convert below
                        _service.InsertDetail(new MPRDetail
                        {
                            MPR_ID = targetId,
                            Item_No = stt++,
                            Item_Name = nm,
                            Description = row2.Cells["RDesc"].Value?.ToString() ?? "",
                            Material = row2.Cells["RMaterial"].Value?.ToString() ?? "",
                            Thickness_mm = decimal.TryParse(row2.Cells["RT_mm"].Value?.ToString(), out decimal rTh) ? rTh : 0,
                            Depth_mm = decimal.TryParse(row2.Cells["RD_mm"].Value?.ToString(), out decimal rDp) ? rDp : 0,
                            C_Width_mm = decimal.TryParse(row2.Cells["RW_mm"].Value?.ToString(), out decimal rCw) ? rCw : 0,
                            D_Web_mm = decimal.TryParse(row2.Cells["RWeb_mm"].Value?.ToString(), out decimal rDw) ? rDw : 0,
                            E_Flange_mm = decimal.TryParse(row2.Cells["RFlange_mm"].Value?.ToString(), out decimal rEf) ? rEf : 0,
                            F_Length_mm = decimal.TryParse(row2.Cells["RL_mm"].Value?.ToString(), out decimal rFl) ? rFl : 0,
                            UNIT = row2.Cells["RUNIT"].Value?.ToString() ?? "",
                            Qty_Per_Sheet = decimal.TryParse(row2.Cells["RQty"].Value?.ToString(), out decimal rQs) ? rQs : 0,
                            Weight_kg = decimal.TryParse(row2.Cells["RKG"].Value?.ToString(), out decimal rWk) ? rWk : 0,
                            Remarks = row2.Cells["RRemarks"].Value?.ToString() ?? "",
                            REV = detRev.ToString()
                        }, _currentUser);
                    }

                    string msg = isAdmin
                        ? $"✅ Admin đã cập nhật MPR '{oldMprNo}' (giữ nguyên REV)!"
                        : $"✅ Đã tạo Revise '{newMprNo}' (REV={nextRev})!";
                    MessageBox.Show(msg, "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadMPR();
                    dlg.Close();
                }
                catch (Exception ex) { lblSave.Text = "❌ " + ex.Message; }
            };

            dlg.Owner = this.FindForm();
            dlg.ShowDialog();
        }

        private void BtnSaveHeader_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("MPR", "Lưu Header", "Lưu Header")) return;
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

        // =====================================================================
        // XÓA MPR — CASCADE DELETE ĐÚNG THỨ TỰ TRONG TRANSACTION
        // Thứ tự: PO_Detail (FK) → MPR_Details → MPR_Header
        // =====================================================================
        private void BtnDeleteMPR_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("MPR", "Xóa MPR", "Xóa MPR")) return;
            if (_selectedMPR_ID == 0)
            {
                MessageBox.Show("Vui lòng chọn phiếu MPR cần xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Lấy tên MPR để hiển thị xác nhận
            var mprToDelete = _mprList.Find(m => m.MPR_ID == _selectedMPR_ID);
            string mprNoDisplay = mprToDelete?.MPR_No ?? _selectedMPR_ID.ToString();

            string confirmMsg =
                $"Bạn có chắc chắn muốn xóa phiếu MPR: [{mprNoDisplay}] ?\n\n" +
                $"⚠ Thao tác này sẽ xóa toàn bộ:\n" +
                $"   • Chi tiết vật tư của phiếu MPR này\n" +
                $"   • Liên kết PO_Detail đến các dòng vật tư\n\n" +
                $"Dữ liệu sẽ KHÔNG thể khôi phục!";

            if (MessageBox.Show(confirmMsg, "⚠ Xác nhận xóa MPR",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2) != DialogResult.Yes)
                return;

            try
            {
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    using (var tran = conn.BeginTransaction())
                    {
                        try
                        {
                            // Bước 1: Xóa PO_Detail liên kết đến các dòng MPR_Details của MPR này
                            // (giải phóng FK_PO_Detail_MPR_Details trước khi xóa MPR_Details)
                            var cmd1 = new SqlCommand(@"
                                DELETE pod
                                FROM dbo.PO_Detail pod
                                INNER JOIN dbo.MPR_Details md ON pod.MPR_Detail_ID = md.Detail_ID
                                WHERE md.MPR_ID = @mprId", conn, tran);
                            cmd1.Parameters.AddWithValue("@mprId", _selectedMPR_ID);
                            int poDetailDeleted = cmd1.ExecuteNonQuery();

                            // Bước 2: Xóa toàn bộ MPR_Details của MPR này
                            var cmd2 = new SqlCommand(
                                "DELETE FROM dbo.MPR_Details WHERE MPR_ID = @mprId", conn, tran);
                            cmd2.Parameters.AddWithValue("@mprId", _selectedMPR_ID);
                            int detailDeleted = cmd2.ExecuteNonQuery();

                            // Bước 3: Xóa MPR Header
                            var cmd3 = new SqlCommand(
                                "DELETE FROM dbo.MPR_Header WHERE MPR_ID = @mprId", conn, tran);
                            cmd3.Parameters.AddWithValue("@mprId", _selectedMPR_ID);
                            cmd3.ExecuteNonQuery();

                            tran.Commit();

                            string resultMsg = $"✅ Xóa phiếu MPR [{mprNoDisplay}] thành công!\n\n" +
                                               $"   • {detailDeleted} dòng vật tư đã xóa\n" +
                                               $"   • {poDetailDeleted} liên kết PO_Detail đã xóa";
                            MessageBox.Show(resultMsg, "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch
                        {
                            tran.Rollback();
                            throw;
                        }
                    }
                }

                _selectedMPR_ID = 0;
                ClearHeader();
                dgvDetails.Rows.Clear();
                dgvPOProgress.DataSource = null;
                dgvFiles.Rows.Clear();
                LoadMPR();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi xóa MPR: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnClearHeader_Click(object sender, EventArgs e)
        {
            _selectedMPR_ID = 0;
            ClearHeader();
            dgvDetails.Rows.Clear();
            dgvPOProgress.DataSource = null;
            dgvFiles.Rows.Clear();
            _details.Clear();
        }

        private void BtnCreatePO_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("MPR", "Tạo PO", "Tạo PO từ MPR")) return;
            if (_selectedMPR_ID == 0)
            {
                MessageBox.Show("Vui lòng chọn một MPR trước!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var mpr = _mprList.Find(m => m.MPR_ID == _selectedMPR_ID);
            if (mpr == null) return;

            string mprNo = mpr.MPR_No;
            var frm = new frmPO(mprNo, true);
            frm.Show();
        }

        private void BtnAddDetail_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("MPR", "Thêm dòng", "Thêm dòng")) return;
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
            newRow.Cells["PO_No"].Value = "";

            dgvDetails.CurrentCell = dgvDetails.Rows[newIdx].Cells["Item_Name"];
            dgvDetails.BeginEdit(true);
        }

        private void BtnDeleteDetail_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("MPR", "Xóa dòng", "Xóa dòng")) return;
            if (dgvDetails.SelectedRows.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn dòng cần xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string msg = dgvDetails.SelectedRows.Count == 1
                ? "Bạn có chắc chắn muốn xóa dòng này?"
                : $"Bạn có chắc chắn muốn xóa {dgvDetails.SelectedRows.Count} dòng đã chọn?";
            if (MessageBox.Show(msg, "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    var rowsToDelete = new List<DataGridViewRow>();
                    foreach (DataGridViewRow row in dgvDetails.SelectedRows)
                    {
                        if (!row.IsNewRow)
                        {
                            rowsToDelete.Add(row);
                        }
                    }

                    foreach (var row in rowsToDelete)
                    {
                        dgvDetails.Rows.Remove(row);
                    }

                    int itemNo = 1;
                    foreach (DataGridViewRow row in dgvDetails.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            row.Cells["Item_No"].Value = itemNo++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi xóa: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnSaveDetail_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("MPR", "Lưu chi tiết", "Lưu chi tiết")) return;
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

            // ── Xác thực mật khẩu Admin trước khi lưu ──
            if (!VerifyAdminPassword()) return;

            try
            {
                int saved = 0;
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    foreach (DataGridViewRow row in dgvDetails.Rows)
                    {
                        if (row.IsNewRow) continue;
                        string itemName = row.Cells["Item_Name"].Value?.ToString() ?? "";
                        if (string.IsNullOrWhiteSpace(itemName)) continue;

                        int detailId = Convert.ToInt32(row.Cells["Detail_ID"].Value ?? 0);
                        int itemNo = Convert.ToInt32(row.Cells["Item_No"].Value ?? 0);
                        string desc = row.Cells["Description"].Value?.ToString() ?? "";
                        string material = row.Cells["Material"].Value?.ToString() ?? "";
                        decimal thickMm = DecimalVal(row.Cells["Thickness_mm"].Value);
                        decimal depthMm = DecimalVal(row.Cells["Depth_mm"].Value);
                        decimal cWidthMm = DecimalVal(row.Cells["C_Width_mm"].Value);
                        decimal dWebMm = DecimalVal(row.Cells["D_Web_mm"].Value);
                        decimal eFlangeMm = DecimalVal(row.Cells["E_Flange_mm"].Value);
                        decimal fLengthMm = DecimalVal(row.Cells["F_Length_mm"].Value);
                        string unit = row.Cells["UNIT"].Value?.ToString() ?? "";
                        int qty = (int)DecimalVal(row.Cells["Qty"].Value);
                        decimal weight = DecimalVal(row.Cells["Weight"].Value);
                        string mpsInfo = row.Cells["MPS_Info"].Value?.ToString() ?? "";
                        string usageLoc = row.Cells["Usage_Location"].Value?.ToString() ?? "";
                        string rev = row.Cells["REV"].Value?.ToString() ?? "0";
                        string remarks = row.Cells["Remarks"].Value?.ToString() ?? "";
                        string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                        Microsoft.Data.SqlClient.SqlCommand cmd;

                        if (detailId == 0)
                        {
                            cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                                INSERT INTO MPR_Details
                                    (MPR_ID, Item_No, item_name, Description, Material,
                                     Thickness_mm, Depth_mm, C_Width_mm, D_Web_mm, E_Flange_mm, F_Length_mm,
                                     UNIT, Qty_Per_Sheet, Weight_kg, MPS_Info, Usage_Location, REV, Remarks,
                                     Created_Date, Created_By, Modified_Date, Modified_By)
                                VALUES
                                    (@mprId, @itemNo, @itemName, @desc, @material,
                                     @thick, @depth, @cWidth, @dWeb, @eFlange, @fLen,
                                     @unit, @qty, @weight, @mps, @usage, @rev, @remarks,
                                     @now, @user, @now, @user);
                                SELECT SCOPE_IDENTITY();", conn);
                        }
                        else
                        {
                            cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                                UPDATE MPR_Details SET
                                    Item_No       = @itemNo,
                                    item_name     = @itemName,
                                    Description   = @desc,
                                    Material      = @material,
                                    Thickness_mm  = @thick,
                                    Depth_mm      = @depth,
                                    C_Width_mm    = @cWidth,
                                    D_Web_mm      = @dWeb,
                                    E_Flange_mm   = @eFlange,
                                    F_Length_mm   = @fLen,
                                    UNIT          = @unit,
                                    Qty_Per_Sheet = @qty,
                                    Weight_kg     = @weight,
                                    MPS_Info      = @mps,
                                    Usage_Location= @usage,
                                    REV           = @rev,
                                    Remarks       = @remarks,
                                    Modified_Date = @now,
                                    Modified_By   = @user
                                WHERE Detail_ID = @detailId", conn);
                            cmd.Parameters.AddWithValue("@detailId", detailId);
                        }

                        cmd.Parameters.AddWithValue("@mprId", _selectedMPR_ID);
                        cmd.Parameters.AddWithValue("@itemNo", itemNo);
                        cmd.Parameters.AddWithValue("@itemName", itemName);
                        cmd.Parameters.AddWithValue("@desc", desc);
                        cmd.Parameters.AddWithValue("@material", material);
                        cmd.Parameters.AddWithValue("@thick", thickMm);
                        cmd.Parameters.AddWithValue("@depth", depthMm);
                        cmd.Parameters.AddWithValue("@cWidth", cWidthMm);
                        cmd.Parameters.AddWithValue("@dWeb", dWebMm);
                        cmd.Parameters.AddWithValue("@eFlange", eFlangeMm);
                        cmd.Parameters.AddWithValue("@fLen", fLengthMm);
                        cmd.Parameters.AddWithValue("@unit", unit);
                        cmd.Parameters.AddWithValue("@qty", qty);
                        cmd.Parameters.AddWithValue("@weight", weight);
                        cmd.Parameters.AddWithValue("@mps", mpsInfo);
                        cmd.Parameters.AddWithValue("@usage", usageLoc);
                        cmd.Parameters.AddWithValue("@rev", rev);
                        cmd.Parameters.AddWithValue("@remarks", remarks);
                        cmd.Parameters.AddWithValue("@now", now);
                        cmd.Parameters.AddWithValue("@user", _currentUser ?? "Admin");

                        if (detailId == 0)
                        {
                            var newId = cmd.ExecuteScalar();
                            if (newId != null && newId != DBNull.Value)
                                row.Cells["Detail_ID"].Value = Convert.ToInt32(newId);
                        }
                        else
                        {
                            cmd.ExecuteNonQuery();
                        }

                        saved++;
                    }
                }

                MessageBox.Show($"✅ Đã lưu {saved} dòng chi tiết thành công!", "Thành công",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDetails(_selectedMPR_ID);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi lưu chi tiết: " + ex.Message, "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =========================================================================
        // XÁC THỰC MẬT KHẨU ADMIN
        // =========================================================================
        private bool VerifyAdminPassword()
        {
            var dlg = new Form
            {
                Text = "🔐 Xác thực Admin",
                Size = new Size(380, 170),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(245, 245, 245),
                KeyPreview = true
            };

            dlg.Controls.Add(new Label
            {
                Text = "Nhập mật khẩu tài khoản Admin để xác nhận lưu:",
                Font = new Font("Segoe UI", 9),
                Location = new Point(15, 15),
                Size = new Size(340, 20)
            });

            var txtPwd = new TextBox
            {
                Location = new Point(15, 40),
                Size = new Size(340, 26),
                Font = new Font("Segoe UI", 10),
                PasswordChar = '●',
                UseSystemPasswordChar = false
            };
            dlg.Controls.Add(txtPwd);

            var lblErr = new Label
            {
                Text = "",
                ForeColor = Color.FromArgb(220, 53, 69),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Location = new Point(15, 72),
                Size = new Size(340, 20)
            };
            dlg.Controls.Add(lblErr);

            var btnOK = new Button
            {
                Text = "✔ Xác nhận",
                Location = new Point(155, 98),
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
                Location = new Point(265, 98),
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
                if (string.IsNullOrEmpty(pwd))
                { lblErr.Text = "Vui lòng nhập mật khẩu!"; return; }

                try
                {
                    string inputHash;
                    using (var sha256 = System.Security.Cryptography.SHA256.Create())
                    {
                        byte[] bytes = sha256.ComputeHash(System.Text.Encoding.UTF8.GetBytes(pwd));
                        inputHash = BitConverter.ToString(bytes).Replace("-", "").ToLower();
                    }

                    const string ADMIN_HASH = "e86f78a8a3caf0b60d8e74e5942aa6d86dc150cd3c03338aef25b7d2d7e3acc7";

                    bool match = false;

                    if (inputHash == ADMIN_HASH)
                    {
                        match = true;
                    }
                    else
                    {
                        using var conn = DatabaseHelper.GetConnection();
                        conn.Open();

                        var cmd1 = new Microsoft.Data.SqlClient.SqlCommand(
                            @"SELECT COUNT(1) FROM Users
                              WHERE LOWER(Username) = 'admin'
                                AND (LOWER(Password) = @hash
                                  OR Password = @hashUpper)", conn);
                        cmd1.Parameters.AddWithValue("@hash", inputHash);
                        cmd1.Parameters.AddWithValue("@hashUpper", inputHash.ToUpper());
                        if (Convert.ToInt32(cmd1.ExecuteScalar()) > 0)
                            match = true;

                        if (!match)
                        {
                            var cmd2 = new Microsoft.Data.SqlClient.SqlCommand(
                                @"SELECT COUNT(1) FROM Users
                                  WHERE LOWER(Username) = 'admin'
                                    AND Password = @pwd", conn);
                            cmd2.Parameters.AddWithValue("@pwd", pwd);
                            if (Convert.ToInt32(cmd2.ExecuteScalar()) > 0)
                                match = true;
                        }
                    }

                    if (match)
                    {
                        verified = true;
                        dlg.DialogResult = DialogResult.OK;
                        dlg.Close();
                    }
                    else
                    {
                        lblErr.Text = "❌ Mật khẩu không đúng!";
                        txtPwd.Clear();
                        txtPwd.Focus();
                    }
                }
                catch (Exception ex)
                {
                    lblErr.Text = "Lỗi xác thực: " + ex.Message;
                }
            };

            dlg.KeyDown += (s, ev) =>
            {
                if (ev.KeyCode == Keys.Enter) { btnOK.PerformClick(); ev.SuppressKeyPress = true; }
            };

            txtPwd.Focus();
            dlg.ShowDialog(this);
            return verified;
        }

        // =========================================================================
        // CHECK ALL ITEMS — Popup tổng hợp toàn bộ MPR Detail + RIR kết quả KT
        // =========================================================================
        private void BtnCheckAllItems_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("MPR", "Check All Items", "Check All Items")) return;
            ShowCheckAllItemsPopup();
        }

        private void ShowCheckAllItemsPopup()
        {
            try
            {
                // SQL tối ưu: CTE pre-aggregate, tương thích SQL Server 2012+
                const string SQL = @"
                    WITH
                    cte_PO AS (
                        SELECT pod.MPR_Detail_ID,
                               ISNULL(STUFF((
                                   SELECT DISTINCT N', ' + ph2.PONo
                                   FROM   PO_Detail pod2
                                   INNER JOIN PO_head ph2 ON ph2.PO_ID = pod2.PO_ID
                                   WHERE  pod2.MPR_Detail_ID = pod.MPR_Detail_ID
                                   FOR XML PATH(''), TYPE).value('.','NVARCHAR(MAX)'),1,2,N''),N'') AS POList
                        FROM   PO_Detail pod
                        WHERE  pod.MPR_Detail_ID IS NOT NULL
                        GROUP BY pod.MPR_Detail_ID
                    ),
                    cte_Heat AS (
                        SELECT pod.MPR_Detail_ID,
                               ISNULL(STUFF((
                                   SELECT DISTINCT N', ' + rd2.Heatno
                                   FROM   PO_Detail pod2
                                   INNER JOIN PO_head    ph2 ON ph2.PO_ID  = pod2.PO_ID
                                   INNER JOIN RIR_head   rh2 ON rh2.PONo   = ph2.PONo
                                   INNER JOIN RIR_detail rd2 ON rd2.RIR_ID = rh2.RIR_ID
                                   WHERE  pod2.MPR_Detail_ID = pod.MPR_Detail_ID
                                     AND  ISNULL(rd2.Heatno,N'') != N''
                                   FOR XML PATH(''), TYPE).value('.','NVARCHAR(MAX)'),1,2,N''),N'') AS HeatList
                        FROM   PO_Detail pod
                        WHERE  pod.MPR_Detail_ID IS NOT NULL
                        GROUP BY pod.MPR_Detail_ID
                    ),
                    cte_KT AS (
                        SELECT pod.MPR_Detail_ID,
                               MIN(CASE rd.Inspect_Result
                                   WHEN N'Fail' THEN 1 WHEN N'Hold' THEN 2
                                   WHEN N'Pass' THEN 3 ELSE 4 END) AS KT_Rank
                        FROM   PO_Detail pod
                        INNER JOIN PO_head    ph ON ph.PO_ID  = pod.PO_ID
                        INNER JOIN RIR_head   rh ON rh.PONo   = ph.PONo
                        INNER JOIN RIR_detail rd ON rd.RIR_ID = rh.RIR_ID
                        WHERE  pod.MPR_Detail_ID IS NOT NULL
                        GROUP BY pod.MPR_Detail_ID
                    ),
                    cte_RIR AS (
                        SELECT pod.MPR_Detail_ID,
                               ISNULL(STUFF((
                                   SELECT DISTINCT N', ' + rh2.RIR_No
                                   FROM   PO_Detail pod2
                                   INNER JOIN PO_head  ph2 ON ph2.PO_ID = pod2.PO_ID
                                   INNER JOIN RIR_head rh2 ON rh2.PONo  = ph2.PONo
                                   WHERE  pod2.MPR_Detail_ID = pod.MPR_Detail_ID
                                     AND  ISNULL(rh2.RIR_No,N'') != N''
                                   FOR XML PATH(''), TYPE).value('.','NVARCHAR(MAX)'),1,2,N''),N'') AS RIRList,
                               (SELECT TOP 1 rh3.Status
                                FROM   PO_Detail pod3
                                INNER JOIN PO_head  ph3 ON ph3.PO_ID = pod3.PO_ID
                                INNER JOIN RIR_head rh3 ON rh3.PONo  = ph3.PONo
                                WHERE  pod3.MPR_Detail_ID = pod.MPR_Detail_ID
                                  AND  ISNULL(rh3.Status,N'') != N''
                                ORDER BY rh3.Issue_Date DESC) AS RIR_Status
                        FROM   PO_Detail pod
                        WHERE  pod.MPR_Detail_ID IS NOT NULL
                        GROUP BY pod.MPR_Detail_ID
                    )
                    SELECT
                        ISNULL(pi.ProjectCode, N'')                         AS [Mã dự án],
                        h.MPR_No                                            AS [MPR No],
                        h.Rev                                               AS [Rev],
                        d.Item_No                                           AS [Item No],
                        d.item_name                                         AS [Tên vật tư],
                        d.Material                                          AS [Vật liệu],
                        ISNULL(CAST(NULLIF(d.Thickness_mm,0) AS NVARCHAR),N'') AS [A-Dày(mm)],
                        ISNULL(CAST(NULLIF(d.Depth_mm,    0) AS NVARCHAR),N'') AS [B-Sâu(mm)],
                        ISNULL(CAST(NULLIF(d.C_Width_mm,  0) AS NVARCHAR),N'') AS [C-Rộng(mm)],
                        ISNULL(CAST(NULLIF(d.D_Web_mm,    0) AS NVARCHAR),N'') AS [D-Bụng(mm)],
                        ISNULL(CAST(NULLIF(d.E_Flange_mm, 0) AS NVARCHAR),N'') AS [E-Cánh(mm)],
                        ISNULL(CAST(NULLIF(d.F_Length_mm, 0) AS NVARCHAR),N'') AS [F-Dài(mm)],
                        ISNULL(d.UNIT,       N'')                           AS [ĐVT],
                        d.Qty_Per_Sheet                                     AS [SL],
                        ISNULL(d.Weight_kg,  0)                             AS [KG],
                        ISNULL(cp.POList,    N'')                           AS [PO No],
                        ISNULL(ch.HeatList,  N'')                           AS [Heat No],
                        ISNULL(CASE ck.KT_Rank
                            WHEN 1 THEN N'Fail' WHEN 2 THEN N'Hold'
                            WHEN 3 THEN N'Pass' ELSE N'Chưa KT' END, N'Chưa KT') AS [Kết quả KT],
                        ISNULL(cr.RIRList,   N'')                           AS [RIR No],
                        ISNULL(cr.RIR_Status,N'')                           AS [Trạng thái RIR]
                    FROM MPR_Header h
                    INNER JOIN MPR_Details d  ON d.MPR_ID       = h.MPR_ID
                    LEFT  JOIN ProjectInfo pi ON pi.ProjectCode = h.Project_Code
                    LEFT  JOIN cte_PO  cp ON cp.MPR_Detail_ID   = d.Detail_ID
                    LEFT  JOIN cte_Heat ch ON ch.MPR_Detail_ID  = d.Detail_ID
                    LEFT  JOIN cte_KT  ck ON ck.MPR_Detail_ID   = d.Detail_ID
                    LEFT  JOIN cte_RIR cr ON cr.MPR_Detail_ID   = d.Detail_ID
                    ORDER BY pi.ProjectCode, h.MPR_No, d.Item_No";


                DataTable dtFull;
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    dtFull = new DataTable();
                    dtFull.Load(new SqlCommand(SQL, conn).ExecuteReader());
                }

                // ── Popup ──
                var popup = new Form
                {
                    Text = "🔎 Check All Items — Tổng hợp toàn bộ vật tư MPR",
                    Size = new Size(1400, 700),
                    StartPosition = FormStartPosition.CenterParent,
                    BackColor = Color.FromArgb(245, 245, 245),
                    MinimumSize = new Size(1100, 500),
                    KeyPreview = true
                };

                popup.Controls.Add(new Label
                {
                    Text = "🔎  CHECK ALL ITEMS — Tổng hợp vật tư tất cả MPR & kết quả kiểm tra RIR  |  💡 Double click → mở MPR",
                    Font = new Font("Segoe UI", 10, FontStyle.Bold),
                    ForeColor = Color.FromArgb(102, 51, 153),
                    Location = new Point(10, 8),
                    Size = new Size(900, 24)
                });

                int total = dtFull.Rows.Count;
                int pass = 0, fail = 0, hold = 0, notYet = 0;
                foreach (DataRow r in dtFull.Rows)
                {
                    string kt = r["Kết quả KT"]?.ToString() ?? "";
                    if (kt == "Pass") pass++;
                    else if (kt == "Fail") fail++;
                    else if (kt == "Hold") hold++;
                    else notYet++;
                }
                var lblStat = new Label
                {
                    Text = $"Tổng: {total}  |  ✅ Pass: {pass}  |  ❌ Fail: {fail}  |  ⏸ Hold: {hold}  |  ⏳ Chưa KT: {notYet}",
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    ForeColor = Color.FromArgb(0, 120, 212),
                    Location = new Point(10, 36),
                    Size = new Size(1360, 20),
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
                };
                popup.Controls.Add(lblStat);

                var pFilter = new Panel
                {
                    Location = new Point(10, 62),
                    Size = new Size(popup.ClientSize.Width - 20, 72),
                    BackColor = Color.White,
                    BorderStyle = BorderStyle.FixedSingle,
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
                };
                popup.Controls.Add(pFilter);

                Action<string, int, int, int> addFL = (txt, x, y2, w) =>
                    pFilter.Controls.Add(new Label
                    {
                        Text = txt,
                        Location = new Point(x, y2 + 3),
                        Size = new Size(w, 18),
                        Font = new Font("Segoe UI", 8, FontStyle.Bold),
                        ForeColor = Color.FromArgb(60, 60, 60)
                    });

                int row1Y = 6;
                int x1 = 6;

                addFL("Mã DA:", x1, row1Y, 48);

                // ── Checked-combo dropdown cho Mã DA (chọn nhiều) ──
                var daProjectList = dtFull.AsEnumerable()
                    .Select(r => r["Mã dự án"].ToString())
                    .Where(v => !string.IsNullOrEmpty(v))
                    .Distinct().OrderBy(v => v).ToList();

                var clbDA = new CheckedListBox
                {
                    Font = new Font("Segoe UI", 9),
                    CheckOnClick = true,
                    BorderStyle = BorderStyle.None,
                    BackColor = Color.White,
                    IntegralHeight = false,
                    Width = 220
                };
                foreach (var p in daProjectList) clbDA.Items.Add(p, false);
                clbDA.Height = Math.Min(clbDA.Items.Count, 10) * 17 + 4;

                var panelDropDA = new Panel
                {
                    Size = new Size(224, clbDA.Height + 2),
                    BackColor = Color.White,
                    BorderStyle = BorderStyle.FixedSingle,
                    Visible = false
                };
                popup.Controls.Add(panelDropDA);
                panelDropDA.Controls.Add(clbDA);
                panelDropDA.BringToFront();

                var btnDropDA = new Button
                {
                    Location = new Point(x1 + 50, row1Y),
                    Size = new Size(120, 22),
                    Text = "(Tất cả)  ▼",
                    TextAlign = ContentAlignment.MiddleLeft,
                    Font = new Font("Segoe UI", 8),
                    BackColor = Color.White,
                    ForeColor = Color.FromArgb(30, 30, 30),
                    FlatStyle = FlatStyle.Flat,
                    Cursor = Cursors.Hand,
                    Padding = new Padding(3, 0, 0, 0)
                };
                btnDropDA.FlatAppearance.BorderColor = Color.FromArgb(171, 173, 179);
                btnDropDA.FlatAppearance.BorderSize = 1;
                btnDropDA.Paint += (s, ev) =>
                {
                    int ax = btnDropDA.Width - 14, ay = btnDropDA.Height / 2;
                    ev.Graphics.FillPolygon(Brushes.DimGray, new[]
                    {
                        new Point(ax, ay - 3), new Point(ax + 7, ay - 3), new Point(ax + 3, ay + 3)
                    });
                };
                pFilter.Controls.Add(btnDropDA);

                Action updateBtnDA = () =>
                {
                    var sel = clbDA.CheckedItems.Cast<string>().ToList();
                    btnDropDA.Text = sel.Count == 0 ? "(Tất cả)  ▼" :
                                     sel.Count == 1 ? sel[0] + "  ▼" :
                                     $"({sel.Count} DA)  ▼";
                    btnDropDA.ForeColor = sel.Count > 0 ? Color.FromArgb(102, 51, 153) : Color.FromArgb(30, 30, 30);
                    btnDropDA.Font = new Font("Segoe UI", 8, sel.Count > 0 ? FontStyle.Bold : FontStyle.Regular);
                };

                btnDropDA.Click += (s, ev) =>
                {
                    if (panelDropDA.Visible) { panelDropDA.Visible = false; return; }
                    var pt = popup.PointToClient(btnDropDA.Parent.PointToScreen(
                        new Point(btnDropDA.Left, btnDropDA.Bottom + 2)));
                    panelDropDA.Location = pt;
                    panelDropDA.BringToFront();
                    panelDropDA.Visible = true;
                    clbDA.Focus();
                };

                popup.MouseDown += (s, ev) =>
                {
                    if (panelDropDA.Visible && !panelDropDA.Bounds.Contains(ev.Location))
                        panelDropDA.Visible = false;
                };

                x1 += 180;

                addFL("Tên VT:", x1, row1Y, 50);
                var txtFName = new TextBox { Location = new Point(x1 + 52, row1Y), Size = new Size(140, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "Tên vật tư..." };
                pFilter.Controls.Add(txtFName);
                x1 += 202;

                addFL("A(mm):", x1, row1Y, 48);
                var txtFA = new TextBox { Location = new Point(x1 + 50, row1Y), Size = new Size(44, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "A" };
                pFilter.Controls.Add(txtFA); x1 += 102;

                addFL("B(mm):", x1, row1Y, 48);
                var txtFB = new TextBox { Location = new Point(x1 + 50, row1Y), Size = new Size(44, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "B" };
                pFilter.Controls.Add(txtFB); x1 += 102;

                addFL("C(mm):", x1, row1Y, 48);
                var txtFC = new TextBox { Location = new Point(x1 + 50, row1Y), Size = new Size(44, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "C" };
                pFilter.Controls.Add(txtFC); x1 += 102;

                addFL("D(mm):", x1, row1Y, 48);
                var txtFD = new TextBox { Location = new Point(x1 + 50, row1Y), Size = new Size(44, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "D" };
                pFilter.Controls.Add(txtFD); x1 += 102;

                addFL("E(mm):", x1, row1Y, 48);
                var txtFE = new TextBox { Location = new Point(x1 + 50, row1Y), Size = new Size(44, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "E" };
                pFilter.Controls.Add(txtFE); x1 += 102;

                addFL("F(mm):", x1, row1Y, 48);
                var txtFF = new TextBox { Location = new Point(x1 + 50, row1Y), Size = new Size(44, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "F" };
                pFilter.Controls.Add(txtFF);

                int row2Y = 38;
                int x2 = 6;

                addFL("Heat No:", x2, row2Y, 55);
                var txtFHeat = new TextBox { Location = new Point(x2 + 57, row2Y), Size = new Size(100, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "Heat No..." };
                pFilter.Controls.Add(txtFHeat);
                x2 += 167;

                addFL("KQ KT:", x2, row2Y, 48);
                var cboFKQ = new ComboBox
                {
                    Location = new Point(x2 + 50, row2Y),
                    Size = new Size(100, 22),
                    Font = new Font("Segoe UI", 9),
                    DropDownStyle = ComboBoxStyle.DropDownList
                };
                cboFKQ.Items.AddRange(new[] { "Tất cả", "Pass", "Fail", "Hold", "Chưa KT" });
                cboFKQ.SelectedIndex = 0;
                pFilter.Controls.Add(cboFKQ);
                x2 += 160;

                addFL("Vật liệu:", x2, row2Y, 58);
                var txtFMat = new TextBox
                {
                    Location = new Point(x2 + 60, row2Y),
                    Size = new Size(110, 22),
                    Font = new Font("Segoe UI", 9),
                    PlaceholderText = "Vật liệu..."
                };
                pFilter.Controls.Add(txtFMat);
                x2 += 178;

                var btnFSearch = new Button
                {
                    Text = "🔍 Lọc",
                    Location = new Point(x2, row2Y - 2),
                    Size = new Size(75, 26),
                    BackColor = Color.FromArgb(0, 120, 212),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 8, FontStyle.Bold)
                };
                btnFSearch.FlatAppearance.BorderSize = 0;
                pFilter.Controls.Add(btnFSearch);

                var btnFClear = new Button
                {
                    Text = "✖ Xóa lọc",
                    Location = new Point(x2 + 79, row2Y - 2),
                    Size = new Size(85, 26),
                    BackColor = Color.FromArgb(108, 117, 125),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 8, FontStyle.Bold)
                };
                btnFClear.FlatAppearance.BorderSize = 0;
                pFilter.Controls.Add(btnFClear);

                foreach (Control c in pFilter.Controls)
                    if (c is TextBox || c is ComboBox || c is Button) c.BringToFront();
                // Dropdown panel phải luôn nằm trên cùng khi hiện
                panelDropDA.BringToFront();

                var dgv = new DataGridView
                {
                    Location = new Point(10, 142),
                    Size = new Size(popup.ClientSize.Width - 20, popup.ClientSize.Height - 190),
                    ReadOnly = true,
                    AllowUserToAddRows = false,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                    BackgroundColor = Color.White,
                    BorderStyle = BorderStyle.FixedSingle,
                    RowHeadersVisible = false,
                    Font = new Font("Segoe UI", 9),
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
                };
                dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(102, 51, 153);
                dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                dgv.EnableHeadersVisualStyles = false;
                dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 240, 255);
                dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
                dgv.DefaultCellStyle.SelectionForeColor = Color.Black;
                popup.Controls.Add(dgv);

                dgv.CellFormatting += (s, ev) =>
                {
                    if (ev.RowIndex < 0) return;
                    string col = dgv.Columns[ev.ColumnIndex].Name;
                    if (col == "Kết quả KT")
                    {
                        string v = ev.Value?.ToString() ?? "";
                        ev.CellStyle.ForeColor =
                            v == "Pass" ? Color.FromArgb(40, 167, 69) :
                            v == "Fail" ? Color.FromArgb(220, 53, 69) :
                            v == "Hold" ? Color.FromArgb(255, 140, 0) :
                            Color.FromArgb(108, 117, 125);
                        ev.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    }
                    if (col == "Trạng thái RIR")
                    {
                        string v = ev.Value?.ToString() ?? "";
                        ev.CellStyle.ForeColor =
                            v == "Hoàn thành" ? Color.FromArgb(40, 167, 69) :
                            v == "Đang kiểm tra" ? Color.FromArgb(255, 140, 0) :
                            string.IsNullOrEmpty(v) ? Color.FromArgb(180, 180, 180) :
                            Color.FromArgb(0, 120, 212);
                        ev.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    }
                };

                dgv.RowPrePaint += (s, ev) =>
                {
                    if (ev.RowIndex < 0 || dgv.Rows[ev.RowIndex].IsNewRow) return;
                    string kt = dgv.Rows[ev.RowIndex].Cells["Kết quả KT"].Value?.ToString() ?? "";
                    dgv.Rows[ev.RowIndex].DefaultCellStyle.BackColor =
                        kt == "Pass" ? Color.FromArgb(235, 255, 235) :
                        kt == "Fail" ? Color.FromArgb(255, 235, 235) :
                        kt == "Hold" ? Color.FromArgb(255, 248, 230) :
                        Color.White;
                };

                Action applyFilter = () =>
                {
                    var selDA = clbDA.CheckedItems.Cast<string>().ToList();
                    string kName = txtFName.Text.Trim().ToLower();
                    string kA = txtFA.Text.Trim();
                    string kB = txtFB.Text.Trim();
                    string kC = txtFC.Text.Trim();
                    string kD = txtFD.Text.Trim();
                    string kE = txtFE.Text.Trim();
                    string kF = txtFF.Text.Trim();
                    string kHeat = txtFHeat.Text.Trim().ToLower();
                    string kKQ = cboFKQ.SelectedItem?.ToString() ?? "Tất cả";
                    string kMat = txtFMat.Text.Trim().ToLower();

                    var rows = dtFull.AsEnumerable().Where(r =>
                    {
                        if (selDA.Count > 0 && !selDA.Contains(r["Mã dự án"].ToString())) return false;
                        if (!string.IsNullOrEmpty(kName) && !r["Tên vật tư"].ToString().ToLower().Contains(kName)) return false;
                        if (!string.IsNullOrEmpty(kMat) && !r["Vật liệu"].ToString().ToLower().Contains(kMat)) return false;
                        if (!string.IsNullOrEmpty(kA) && !r["A-Dày(mm)"].ToString().Contains(kA)) return false;
                        if (!string.IsNullOrEmpty(kB) && !r["B-Sâu(mm)"].ToString().Contains(kB)) return false;
                        if (!string.IsNullOrEmpty(kC) && !r["C-Rộng(mm)"].ToString().Contains(kC)) return false;
                        if (!string.IsNullOrEmpty(kD) && !r["D-Bụng(mm)"].ToString().Contains(kD)) return false;
                        if (!string.IsNullOrEmpty(kE) && !r["E-Cánh(mm)"].ToString().Contains(kE)) return false;
                        if (!string.IsNullOrEmpty(kF) && !r["F-Dài(mm)"].ToString().Contains(kF)) return false;
                        if (!string.IsNullOrEmpty(kHeat) && !r["Heat No"].ToString().ToLower().Contains(kHeat)) return false;
                        if (kKQ != "Tất cả" && r["Kết quả KT"].ToString() != kKQ) return false;
                        return true;
                    });

                    DataTable dtView = rows.Any() ? rows.CopyToDataTable() : dtFull.Clone();
                    dgv.DataSource = dtView;

                    foreach (string colName in new[] { "MPR No", "PO No" })
                    {
                        if (!dgv.Columns.Contains(colName)) continue;
                        var col = dgv.Columns[colName];
                        col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                        int measuredW = col.GetPreferredWidth(DataGridViewAutoSizeColumnMode.AllCells, true);
                        col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                        col.Width = Math.Min(measuredW, 400);
                    }

                    int t = dtView.Rows.Count, p = 0, f = 0, h2 = 0, n = 0;
                    foreach (DataRow r in dtView.Rows)
                    {
                        string kt = r["Kết quả KT"]?.ToString() ?? "";
                        if (kt == "Pass") p++;
                        else if (kt == "Fail") f++;
                        else if (kt == "Hold") h2++;
                        else n++;
                    }
                    lblStat.Text = $"Hiển thị: {t}  |  ✅ Pass: {p}  |  ❌ Fail: {f}  |  ⏸ Hold: {h2}  |  ⏳ Chưa KT: {n}";
                };

                dgv.DataSource = null;
                lblStat.Text = "Nhấn [🔍 Lọc] để tìm kiếm dữ liệu.";

                dgv.DataBindingComplete += (s, ev) =>
                {
                    foreach (string colName in new[] { "MPR No", "PO No" })
                    {
                        if (!dgv.Columns.Contains(colName)) continue;
                        var col = dgv.Columns[colName];
                        col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                        int measuredW = col.GetPreferredWidth(DataGridViewAutoSizeColumnMode.AllCells, true);
                        col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                        col.Width = Math.Min(measuredW, 400);
                    }
                };

                // ItemCheck: update button text + refilter (after applyFilter to avoid CS0841)
                clbDA.ItemCheck += (s, ev) =>
                {
                    clbDA.BeginInvoke(new Action(() =>
                    {
                        updateBtnDA();
                        applyFilter();
                    }));
                };

                btnFSearch.Click += (s, ev) => applyFilter();
                btnFClear.Click += (s, ev) =>
                {
                    for (int i = 0; i < clbDA.Items.Count; i++) clbDA.SetItemChecked(i, false);
                    updateBtnDA();
                    panelDropDA.Visible = false;
                    txtFName.Text = "";
                    txtFMat.Text = "";
                    txtFA.Text = ""; txtFB.Text = ""; txtFC.Text = "";
                    txtFD.Text = ""; txtFE.Text = ""; txtFF.Text = "";
                    txtFHeat.Text = ""; cboFKQ.SelectedIndex = 0;
                    dgv.DataSource = null;
                    lblStat.Text = "Nhấn [🔍 Lọc] để tìm kiếm dữ liệu.";
                };

                dgv.CellDoubleClick += (s, ev) =>
                {
                    if (ev.RowIndex < 0) return;
                    string mprNo = dgv.Rows[ev.RowIndex].Cells["MPR No"].Value?.ToString() ?? "";
                    if (string.IsNullOrEmpty(mprNo)) return;

                    var target = _mprList.Find(m => m.MPR_No == mprNo);
                    if (target == null)
                    {
                        MessageBox.Show($"Không tìm thấy MPR: {mprNo}", "Thông báo",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    popup.Close();
                    SelectMPRById(target.MPR_ID);
                };

                var ttip = new ToolTip();
                ttip.SetToolTip(dgv, "Double click vào dòng để mở MPR tương ứng");

                popup.KeyDown += (s, ev) =>
                {
                    if (ev.KeyCode == Keys.Enter)
                    {
                        applyFilter();
                        ev.Handled = true;
                        ev.SuppressKeyPress = true;
                    }
                    if (ev.KeyCode == Keys.Escape) popup.Close();
                };

                var btnExport = new Button
                {
                    Text = "📥 Xuất Excel",
                    Size = new Size(120, 30),
                    BackColor = Color.FromArgb(0, 150, 100),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Right
                };
                btnExport.FlatAppearance.BorderSize = 0;
                btnExport.Location = new Point(popup.ClientSize.Width - 245, popup.ClientSize.Height - 40);
                popup.Controls.Add(btnExport);

                btnExport.Click += (s, ev) =>
                {
                    var dtExport = dgv.DataSource as DataTable;
                    if (dtExport == null || dtExport.Rows.Count == 0)
                    {
                        MessageBox.Show("Không có dữ liệu để xuất!", "Thông báo",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    var sfd = new SaveFileDialog
                    {
                        Title = "Lưu file Excel",
                        Filter = "Excel Files|*.xlsx",
                        FileName = $"CheckAllItems_{DateTime.Now:yyyyMMdd_HHmm}",
                        InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                    };
                    if (sfd.ShowDialog() != DialogResult.OK) return;

                    try
                    {
                        OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                        using (var pkg = new OfficeOpenXml.ExcelPackage())
                        {
                            var ws = pkg.Workbook.Worksheets.Add("Check All Items");

                            ws.Cells[1, 1].Value = "CHECK ALL ITEMS — Tổng hợp vật tư MPR & kết quả kiểm tra RIR";
                            ws.Cells[1, 1, 1, dtExport.Columns.Count].Merge = true;
                            ws.Cells[1, 1].Style.Font.Size = 13;
                            ws.Cells[1, 1].Style.Font.Bold = true;
                            ws.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            ws.Cells[1, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(102, 51, 153));
                            ws.Cells[1, 1].Style.Font.Color.SetColor(Color.White);

                            ws.Cells[2, 1].Value = $"Xuất ngày: {DateTime.Now:dd/MM/yyyy HH:mm}  |  Tổng: {dtExport.Rows.Count} dòng";
                            ws.Cells[2, 1, 2, dtExport.Columns.Count].Merge = true;
                            ws.Cells[2, 1].Style.Font.Italic = true;
                            ws.Cells[2, 1].Style.Font.Size = 9;

                            for (int c = 0; c < dtExport.Columns.Count; c++)
                            {
                                var cell = ws.Cells[3, c + 1];
                                cell.Value = dtExport.Columns[c].ColumnName;
                                cell.Style.Font.Bold = true;
                                cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(102, 51, 153));
                                cell.Style.Font.Color.SetColor(Color.White);
                                cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            }

                            for (int row = 0; row < dtExport.Rows.Count; row++)
                            {
                                var dr = dtExport.Rows[row];
                                bool isAlt = row % 2 == 1;

                                for (int c = 0; c < dtExport.Columns.Count; c++)
                                {
                                    var cell = ws.Cells[row + 4, c + 1];
                                    string colName = dtExport.Columns[c].ColumnName;

                                    if (colName == "KG")
                                    {
                                        if (dr[c] != DBNull.Value && decimal.TryParse(dr[c]?.ToString(), out decimal kg))
                                        {
                                            cell.Value = Math.Round(kg, 2);
                                            cell.Style.Numberformat.Format = "#,##0.00";
                                            cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                                        }
                                        else
                                        {
                                            cell.Value = 0.00m;
                                            cell.Style.Numberformat.Format = "#,##0.00";
                                        }
                                    }
                                    else
                                    {
                                        cell.Value = dr[c]?.ToString() ?? "";
                                    }

                                    cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                    cell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                                    if (isAlt)
                                    {
                                        cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                        cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(248, 240, 255));
                                    }

                                    if (colName == "Kết quả KT")
                                    {
                                        string kt = dr[c]?.ToString() ?? "";
                                        cell.Style.Font.Bold = true;
                                        cell.Style.Font.Color.SetColor(
                                            kt == "Pass" ? Color.FromArgb(40, 167, 69) :
                                            kt == "Fail" ? Color.FromArgb(220, 53, 69) :
                                            kt == "Hold" ? Color.FromArgb(255, 140, 0) :
                                            Color.FromArgb(108, 117, 125));
                                    }
                                }
                            }

                            ws.Cells[ws.Dimension.Address].AutoFitColumns();
                            ws.View.FreezePanes(4, 1);

                            pkg.SaveAs(new System.IO.FileInfo(sfd.FileName));
                        }

                        var res = MessageBox.Show(
                            $"✅ Xuất Excel thành công!\nFile: {sfd.FileName}\n\nBạn có muốn mở file không?",
                            "Thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                        if (res == DialogResult.Yes)
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                            { FileName = sfd.FileName, UseShellExecute = true });
                    }
                    catch (Exception exExport)
                    {
                        MessageBox.Show("Lỗi xuất Excel: " + exExport.Message, "Lỗi",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                };

                var btnClose = new Button
                {
                    Text = "Đóng",
                    Size = new Size(100, 30),
                    BackColor = Color.FromArgb(108, 117, 125),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                    DialogResult = DialogResult.OK
                };
                btnClose.FlatAppearance.BorderSize = 0;
                btnClose.Location = new Point(popup.ClientSize.Width - 115, popup.ClientSize.Height - 40);
                popup.Controls.Add(btnClose);
                popup.AcceptButton = btnFSearch;
                popup.CancelButton = btnClose;

                popup.Resize += (s, ev) =>
                {
                    pFilter.Width = popup.ClientSize.Width - 20;
                    btnExport.Location = new Point(popup.ClientSize.Width - 245, popup.ClientSize.Height - 40);
                    btnClose.Location = new Point(popup.ClientSize.Width - 115, popup.ClientSize.Height - 40);
                };

                popup.Owner = this;
                popup.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        // =====================================================
        //  ÁP DỤNG PHÂN QUYỀN
        // =====================================================
        private void ApplyPermissions()
        {
            if (btnNewMPR != null) PermissionHelper.Apply(btnNewMPR, "MPR", "Tạo MPR");
            // btnCreateMPR dùng cùng quyền "Tạo MPR"
            if (btnDeleteMPR != null) PermissionHelper.Apply(btnDeleteMPR, "MPR", "Xóa MPR");
            if (btnSaveHeader != null) PermissionHelper.Apply(btnSaveHeader, "MPR", "Lưu Header");
            if (btnAddDetail != null) PermissionHelper.Apply(btnAddDetail, "MPR", "Thêm dòng");
            if (btnDeleteDetail != null) PermissionHelper.Apply(btnDeleteDetail, "MPR", "Xóa dòng");
            if (btnSaveDetail != null) PermissionHelper.Apply(btnSaveDetail, "MPR", "Lưu chi tiết");
            foreach (var c in this.Controls.Find("btnCreatePO", true))
                PermissionHelper.Apply(c, "MPR", "Tạo PO");
            foreach (var c in this.Controls.Find("btnCheckAll", true))
                PermissionHelper.Apply(c, "MPR", "Check All Items");
        }

        // =====================================================================
        //  LOC DETAIL THEO DA LEN PO
        // =====================================================================
        // Load dong cac gia tri PO_No vao combobox filter
        private void RefreshPOFilterCombo()
        {
            if (_cboFilterPO == null) return;
            _cboFilterPO.SelectedIndexChanged -= (s, ev) => FilterDetailByPO();
            _cboFilterPO.Items.Clear();
            _cboFilterPO.Items.Add("(Tat ca)"); // mac dinh

            // Lay cac gia tri duy nhat tu cot PO_No
            var seen = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
            bool hasEmpty = false;
            foreach (DataGridViewRow row in dgvDetails.Rows)
            {
                if (row.IsNewRow) continue;
                string val = row.Cells["PO_No"]?.Value?.ToString() ?? "";
                if (string.IsNullOrWhiteSpace(val)) hasEmpty = true;
                else if (seen.Add(val)) _cboFilterPO.Items.Add(val);
            }
            // Them muc loc rong neu co dong chua len PO
            if (hasEmpty) _cboFilterPO.Items.Add("(Chua len PO)");
            _cboFilterPO.SelectedIndex = 0; // chon "(Tat ca)"
            _cboFilterPO.SelectedIndexChanged += (s, ev) => FilterDetailByPO();
        }

        private void FilterDetailByPO()
        {
            if (_cboFilterPO == null || dgvDetails == null) return;
            string sel = _cboFilterPO.SelectedItem?.ToString() ?? "(Tat ca)";

            foreach (DataGridViewRow row in dgvDetails.Rows)
            {
                if (row.IsNewRow) continue;
                string poVal = row.Cells["PO_No"]?.Value?.ToString() ?? "";

                if (sel == "(Tat ca)")
                    row.Visible = true;
                else if (sel == "(Chua len PO)")
                    row.Visible = string.IsNullOrWhiteSpace(poVal);
                else
                    row.Visible = string.Equals(poVal, sel, StringComparison.OrdinalIgnoreCase);
            }
        }

        // =====================================================================
        //  XUAT EXCEL CHI TIET VAT TU
        // =====================================================================
        private void BtnExportDetail_Click(object sender, EventArgs e)
        {
            if (dgvDetails.Rows.Count == 0)
            {
                MessageBox.Show("Khong co du lieu de xuat!", "Thong bao",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using var sfd = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = "MPR_ChiTiet_" + (txtMPRNo?.Text.Trim() ?? "export")
                           + "_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".xlsx",
                Title = "Luu file Excel"
            };
            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using var pkg = new OfficeOpenXml.ExcelPackage();
                var ws = pkg.Workbook.Worksheets.Add("Chi tiet vat tu");

                // Header row
                var visCols = new System.Collections.Generic.List<DataGridViewColumn>();
                foreach (DataGridViewColumn col in dgvDetails.Columns)
                    if (col.Visible && col.Name != "Detail_ID")
                        visCols.Add(col);

                // Style header
                for (int ci = 0; ci < visCols.Count; ci++)
                {
                    var cell = ws.Cells[1, ci + 1];
                    cell.Value = visCols[ci].HeaderText;
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 120, 212));
                    cell.Style.Font.Color.SetColor(System.Drawing.Color.White);
                    cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                }

                // Lay gia tri truc tiep tu _details (so nguyen/decimal thuc)
                // de tranh bi format string boi CellFormatting
                var poMap2 = GetPoMappingForMpr(_selectedMPR_ID);
                // Build map Detail_ID -> row.Visible de biet dong nao dang hien
                var visibleIds = new System.Collections.Generic.HashSet<int>();
                foreach (DataGridViewRow dgvRow in dgvDetails.Rows)
                {
                    if (dgvRow.IsNewRow || !dgvRow.Visible) continue;
                    if (dgvRow.Cells["Detail_ID"]?.Value is int rid) visibleIds.Add(rid);
                    else if (int.TryParse(dgvRow.Cells["Detail_ID"]?.Value?.ToString(), out int rid2))
                        visibleIds.Add(rid2);
                }

                int excelRow = 2;
                foreach (var d in _details)
                {
                    if (!visibleIds.Contains(d.Detail_ID)) continue;

                    string poNo = poMap2.ContainsKey(d.Detail_ID) ? poMap2[d.Detail_ID] : "";

                    // Map gia tri theo ten cot
                    var rowData = new System.Collections.Generic.Dictionary<string, object>
                    {
                        ["Item_No"] = d.Item_No,
                        ["Item_Name"] = d.Item_Name ?? "",
                        ["Description"] = d.Description ?? "",
                        ["Material"] = d.Material ?? "",
                        ["Thickness_mm"] = d.Thickness_mm,
                        ["Depth_mm"] = d.Depth_mm,
                        ["C_Width_mm"] = d.C_Width_mm,
                        ["D_Web_mm"] = d.D_Web_mm,
                        ["E_Flange_mm"] = d.E_Flange_mm,
                        ["F_Length_mm"] = d.F_Length_mm,
                        ["UNIT"] = d.UNIT ?? "",
                        ["Qty"] = d.Qty_Per_Sheet,
                        ["Weight"] = d.Weight_kg,
                        ["MPS_Info"] = d.MPS_Info ?? "",
                        ["Usage_Location"] = d.Usage_Location ?? "",
                        ["REV"] = d.REV,
                        ["Remarks"] = d.Remarks ?? "",
                        ["PO_No"] = poNo
                    };

                    var numericSet = new System.Collections.Generic.HashSet<string>
                        { "Thickness_mm","Depth_mm","C_Width_mm","D_Web_mm",
                          "E_Flange_mm","F_Length_mm","Qty","Weight" };

                    for (int ci = 0; ci < visCols.Count; ci++)
                    {
                        string colN = visCols[ci].Name;
                        var cell = ws.Cells[excelRow, ci + 1];

                        if (!rowData.ContainsKey(colN)) { cell.Value = ""; continue; }

                        object rawVal = rowData[colN];

                        if (numericSet.Contains(colN))
                        {
                            // Set so thuc truc tiep - KHONG qua string
                            double dbl = Convert.ToDouble(rawVal);
                            cell.Value = dbl;
                            // Format: hien so thap phan neu co, khong co separator
                            cell.Style.Numberformat.Format = "0.##";
                        }
                        else
                            cell.Value = rawVal;

                        // Mau dong xen ke
                        if (excelRow % 2 == 0)
                        {
                            cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            cell.Style.Fill.BackgroundColor.SetColor(
                                System.Drawing.Color.FromArgb(240, 248, 255));
                        }

                        // Mau cot Da len PO
                        if (colN == "PO_No" && !string.IsNullOrWhiteSpace(poNo))
                        {
                            cell.Style.Font.Bold = true;
                            cell.Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(40, 167, 69));
                        }

                        cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Hair);
                    }
                    excelRow++;
                }

                // Title MPR info
                ws.InsertRow(1, 2);
                ws.Cells[1, 1].Value = "BANG CHI TIET VAT TU MPR";
                ws.Cells[1, 1].Style.Font.Bold = true;
                ws.Cells[1, 1].Style.Font.Size = 13;
                ws.Cells[1, 1, 1, visCols.Count].Merge = true;

                ws.Cells[2, 1].Value = "MPR No: " + (txtMPRNo?.Text ?? "") +
                                       "   Du an: " + (txtProjectName?.Text ?? "") +
                                       "   Ngay: " + DateTime.Now.ToString("dd/MM/yyyy");
                ws.Cells[2, 1, 2, visCols.Count].Merge = true;

                // Filter info
                string filterInfo = _cboFilterPO?.SelectedItem?.ToString() ?? "Tat ca";
                ws.Cells[2, 1].Value += "   Loc: " + filterInfo;

                // Auto fit
                ws.Cells[ws.Dimension.Address].AutoFitColumns(8, 50);

                pkg.SaveAs(new System.IO.FileInfo(sfd.FileName));

                if (MessageBox.Show("Xuat Excel thanh cong!\nMo file ngay?", "Thanh cong",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    System.Diagnostics.Process.Start(
                        new System.Diagnostics.ProcessStartInfo
                        { FileName = sfd.FileName, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Loi xuat Excel: " + ex.Message, "Loi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // =====================================================================
        //  COPY BANG SANG CLIPBOARD (Ctrl+C / Ctrl+Shift+C)
        // =====================================================================
        private void DgvDetails_GridKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C)
            {
                CopyGridToClipboard(e.Shift); // Shift = kem header
                e.Handled = true;
            }
        }

        private void CopyGridToClipboard(bool withHeader)
        {
            try
            {
                // Chi copy cac dong DANG DUOC CHON (selected rows)
                var selectedRows = dgvDetails.SelectedRows
                    .Cast<DataGridViewRow>()
                    .Where(r => !r.IsNewRow)
                    .OrderBy(r => r.Index)
                    .ToList();

                if (selectedRows.Count == 0)
                {
                    MessageBox.Show("Vui long chon it nhat mot dong de copy!",
                        "Thong bao", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Lay cot hien thi
                var visCols = dgvDetails.Columns
                    .Cast<DataGridViewColumn>()
                    .Where(c => c.Visible && c.Name != "Detail_ID")
                    .OrderBy(c => c.DisplayIndex)
                    .ToList();

                var numSet = new System.Collections.Generic.HashSet<string>
                    { "Qty","Weight","Thickness_mm","Depth_mm",
                      "C_Width_mm","D_Web_mm","E_Flange_mm","F_Length_mm" };

                var sb = new System.Text.StringBuilder();

                // Header neu can
                if (withHeader)
                    sb.AppendLine(string.Join("\t", visCols.Select(c => c.HeaderText)));

                // Doc gia tri tu _details theo Detail_ID de dam bao la so thuc
                var poMap = GetPoMappingForMpr(_selectedMPR_ID);

                foreach (var row in selectedRows)
                {
                    int detId = 0;
                    int.TryParse(row.Cells["Detail_ID"]?.Value?.ToString(), out detId);
                    var d = _details.Find(x => x.Detail_ID == detId);

                    var parts = new System.Collections.Generic.List<string>();
                    foreach (var col in visCols)
                    {
                        string colN = col.Name;
                        string cellVal;

                        if (d != null && numSet.Contains(colN))
                        {
                            // Doc so thuc tu _details, xuat dang InvariantCulture
                            double dbl = colN switch
                            {
                                "Qty" => (double)d.Qty_Per_Sheet,
                                "Weight" => (double)d.Weight_kg,
                                "Thickness_mm" => (double)d.Thickness_mm,
                                "Depth_mm" => (double)d.Depth_mm,
                                "C_Width_mm" => (double)d.C_Width_mm,
                                "D_Web_mm" => (double)d.D_Web_mm,
                                "E_Flange_mm" => (double)d.E_Flange_mm,
                                "F_Length_mm" => (double)d.F_Length_mm,
                                _ => 0
                            };
                            // Xuat: so nguyen thi khong co thap phan
                            cellVal = (dbl == Math.Floor(dbl))
                                ? ((long)dbl).ToString()
                                : dbl.ToString(System.Globalization.CultureInfo.InvariantCulture);
                        }
                        else if (colN == "PO_No")
                            cellVal = (d != null && poMap.ContainsKey(d.Detail_ID))
                                ? poMap[d.Detail_ID] : "";
                        else
                            cellVal = row.Cells[colN]?.Value?.ToString() ?? "";

                        parts.Add(cellVal);
                    }
                    sb.AppendLine(string.Join("\t", parts));
                }

                Clipboard.SetText(sb.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Loi copy: " + ex.Message, "Loi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


    }
}