using MPR_Managerment.Models;
using MPR_Managerment.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Diagnostics;
using System.Windows.Forms;

namespace MPR_Managerment.Forms
{
    public partial class frmPO : Form
    {
        private POService _service = new POService();
        private List<POHead> _poList = new List<POHead>();
        private List<PODetail> _details = new List<PODetail>();
        private int _selectedPO_ID = 0;
        private string _currentUser = "Admin";

        private string _targetPoNo = "";
        private string _importMprNo = "";

        private DataGridView dgvPO;
        private TextBox txtSearch;
        private Button btnSearch, btnNewPO, btnDeletePO, btnClearHeader, btnExport, btnSavePO;
        private Button btnSearchBySupp;
        private Label lblStatus;

        private TextBox txtPONo, txtProjectName, txtWorkorderNo, txtMPRNo;
        private TextBox txtPrepared, txtReviewed, txtAgreement, txtApproved, txtNotes;
        private DateTimePicker dtpPODate;
        private ComboBox cboStatus;
        private NumericUpDown nudRevise;

        // BẢNG MỚI: Tệp đính kèm
        private DataGridView dgvFiles;

        // BẢNG THEO DÕI GIAO HÀNG (Delivery Tracking)
        private DataGridView dgvDelivery;
        private System.Windows.Forms.Timer _deliveryTimer;

        private DataGridView dgvDetails;
        private Button btnAddDetail, btnDeleteDetail;
        private Label lblTotal, lblSubTotal;
        private Panel panelTop, panelHeader, panelDetail;
        private ComboBox cboSupplier;
        private System.Data.DataTable _supplierTable;
        private bool _isSearching = false;

        private string _projectCodeImport = string.Empty;

        public frmPO(string poNo = "")
        {
            _targetPoNo = poNo;
            InitializeComponent();
            BuildUI();
            LoadPO();
            LoadDeliveries();
            this.Resize += FrmPO_Resize;
            if (!string.IsNullOrEmpty(_targetPoNo))
                SelectPOByNo(_targetPoNo);
        }

        public frmPO(string mprNo, bool importMode)
        {
            _importMprNo = mprNo;
            InitializeComponent();
            BuildUI();
            LoadPO();
            this.Resize += FrmPO_Resize;
            this.Shown += (s, e) => ImportMPRByNo(_importMprNo);
        }

        private void SelectPOByNo(string poNo)
        {
            var targetPO = _poList.Find(p => p.PONo == poNo);
            if (targetPO != null) { txtSearch.Text = targetPO.PONo; BtnSearch_Click(null, null); }
            foreach (DataGridViewRow row in dgvPO.Rows)
            {
                if (row.Cells["PO_No"].Value?.ToString() == poNo)
                {
                    dgvPO.ClearSelection();
                    row.Selected = true;
                    if (row.Index >= 0) dgvPO.FirstDisplayedScrollingRowIndex = row.Index;
                    break;
                }
            }
        }

        private void BuildUI()
        {
            this.Text = "Quản lý Đơn Đặt Hàng (PO)";
            this.Size = new Size(1300, 780);
            this.MinimumSize = new Size(1000, 650);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(245, 245, 245);

            // ===== PANEL TOP =====
            panelTop = new Panel { Location = new Point(10, 10), Size = new Size(1260, 210), BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            this.Controls.Add(panelTop);
            panelTop.Controls.Add(new Label { Text = "DANH SÁCH ĐƠN ĐẶT HÀNG (PO)", Font = new Font("Segoe UI", 11, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212), Location = new Point(10, 10), Size = new Size(450, 30) });

            txtSearch = new TextBox { Location = new Point(10, 48), Size = new Size(300, 28), Font = new Font("Segoe UI", 10), PlaceholderText = "Tìm theo PO No, MPR No, tên dự án..." };
            panelTop.Controls.Add(txtSearch);
            txtSearch.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) BtnSearch_Click(null, null); };

            btnSearch = CreateButton("Tìm", Color.FromArgb(0, 120, 212), new Point(320, 47), 70, 30);
            btnNewPO = CreateButton("+ Tạo PO", Color.FromArgb(40, 167, 69), new Point(400, 47), 100, 30);
            btnDeletePO = CreateButton("Xóa PO", Color.FromArgb(220, 53, 69), new Point(510, 47), 90, 30);
            btnSearchBySupp = CreateButton("🔍 Tìm theo NCC", Color.FromArgb(102, 51, 153), new Point(610, 47), 130, 30);

            btnSearch.Click += BtnSearch_Click; btnNewPO.Click += BtnNewPO_Click;
            btnDeletePO.Click += BtnDeletePO_Click; btnSearchBySupp.Click += BtnSearchBySupp_Click;
            panelTop.Controls.Add(btnSearch); panelTop.Controls.Add(btnNewPO);
            panelTop.Controls.Add(btnDeletePO); panelTop.Controls.Add(btnSearchBySupp);

            lblStatus = new Label { Location = new Point(750, 52), Size = new Size(400, 25), Font = new Font("Segoe UI", 9), ForeColor = Color.Gray };
            panelTop.Controls.Add(lblStatus);

            dgvPO = new DataGridView
            {
                Location = new Point(10, 85),
                Size = new Size(1235, 115),
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
            dgvPO.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvPO.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPO.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvPO.EnableHeadersVisualStyles = false;
            dgvPO.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvPO.SelectionChanged += DgvPO_SelectionChanged;

            // Click chuột trái vào bất kỳ ô nào → copy số PO vào clipboard
            dgvPO.CellClick += (s, ev) =>
            {
                if (ev.RowIndex < 0) return;
                string poNo = dgvPO.Rows[ev.RowIndex].Cells["PO_No"].Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(poNo))
                {
                    Clipboard.SetText(poNo);
                    lblStatus.Text = $"✔ Đã copy: {poNo}";
                }
            };

            panelTop.Controls.Add(dgvPO);

            // Button Payment — mở frmPayment với filter theo PO đang chọn
            var btnPayment = CreateButton("💳 Payment", Color.FromArgb(0, 150, 100), new Point(850, 47), 110, 30);
            btnPayment.Click += (s, ev) =>
            {
                string poNo = "";
                if (dgvPO.SelectedRows.Count > 0)
                    poNo = dgvPO.SelectedRows[0].Cells["PO_No"].Value?.ToString() ?? "";

                var frm = new frmPayment(_currentUser, poNo);
                frm.Show();
            };

            panelTop.Controls.Add(btnPayment);

            // 🔥 ĐẢM BẢO NẰM TRÊN LABEL
            btnPayment.BringToFront();

            // ===== PANEL HEADER =====
            panelHeader = new Panel { Location = new Point(10, 230), Size = new Size(1260, 245), BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            this.Controls.Add(panelHeader);
            panelHeader.Controls.Add(new Label { Text = "THÔNG TIN ĐƠN ĐẶT HÀNG", Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212), Location = new Point(10, 8), Size = new Size(350, 25) });

            // BẢNG FILE ĐÍNH KÈM (Bên Phải cùng)
            int gridFilesWidth = 200;
            dgvFiles = new DataGridView
            {
                Location = new Point(panelHeader.Width - gridFilesWidth - 10, 10),
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
            dgvFiles.Columns.Add("FileName", "Tệp đính kèm (PO Link)");
            dgvFiles.Columns.Add("FullPath", "FullPath");
            dgvFiles.Columns["FullPath"].Visible = false;
            dgvFiles.CellDoubleClick += DgvFiles_CellDoubleClick;
            panelHeader.Controls.Add(dgvFiles);

            // BẢNG THEO DÕI GIAO HÀNG — bọc trong Panel con để tọa độ nội bộ luôn chính xác
            const int delivW = 550;
            const int delivGap = 6;
            int delivLeft = (panelHeader.Width - gridFilesWidth - 10) - delivW - delivGap;

            // Panel con — có Anchor Right, các control bên trong dùng tọa độ (0,0)
            var panelDelivery = new Panel
            {
                Location = new Point(delivLeft, 8),
                Size = new Size(delivW, panelHeader.Height - 16),
                BackColor = Color.White,
                Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom
            };
            panelHeader.Controls.Add(panelDelivery);

            // ── Label + Buttons — tọa độ tương đối với panelDelivery ──
            panelDelivery.Controls.Add(new Label
            {
                Text = "📦 Theo dõi giao hàng",
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 150, 100),
                Location = new Point(0, 4),
                Size = new Size(155, 18)
            });

            var btnDelivAdd = new Button
            {
                Text = "＋ Add",
                Location = new Point(158, 1),
                Size = new Size(65, 22),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnDelivAdd.FlatAppearance.BorderSize = 0;
            panelDelivery.Controls.Add(btnDelivAdd);

            var btnDelivDone = new Button
            {
                Text = "✔ Done",
                Location = new Point(227, 1),
                Size = new Size(70, 22),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnDelivDone.FlatAppearance.BorderSize = 0;
            panelDelivery.Controls.Add(btnDelivDone);

            var btnDelivHistory = new Button
            {
                Text = "📋 History",
                Location = new Point(301, 1),
                Size = new Size(80, 22),
                BackColor = Color.FromArgb(0, 150, 100),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnDelivHistory.FlatAppearance.BorderSize = 0;
            btnDelivHistory.Click += BtnReceivedHistory_Click;
            panelDelivery.Controls.Add(btnDelivHistory);

            var btnDelivDel = new Button
            {
                Text = "✖",
                Location = new Point(385, 1),
                Size = new Size(30, 22),
                BackColor = Color.FromArgb(220, 53, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnDelivDel.FlatAppearance.BorderSize = 0;
            panelDelivery.Controls.Add(btnDelivDel);

            // ── dgvDelivery — tọa độ tương đối với panelDelivery ──
            dgvDelivery = new DataGridView
            {
                Location = new Point(0, 26),
                Size = new Size(delivW, panelDelivery.Height - 26),
                ReadOnly = false,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                Name = "dgvDelivery"
            };
            dgvDelivery.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 150, 100);
            dgvDelivery.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvDelivery.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 8, FontStyle.Bold);
            dgvDelivery.EnableHeadersVisualStyles = false;
            dgvDelivery.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(235, 255, 245);
            // Cột ẩn
            dgvDelivery.Columns.Add(new DataGridViewTextBoxColumn { Name = "TrackID", HeaderText = "ID", Visible = false });
            // Cột hiển thị
            dgvDelivery.Columns.Add(new DataGridViewTextBoxColumn { Name = "PONo", HeaderText = "PO No", ReadOnly = true, FillWeight = 25 });
            dgvDelivery.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaDuAn", HeaderText = "Mã DA", ReadOnly = true, FillWeight = 20 });
            dgvDelivery.Columns.Add(new DataGridViewTextBoxColumn { Name = "ExpDelivery", HeaderText = "Exp.Deliv", ReadOnly = true, FillWeight = 22 });
            dgvDelivery.Columns.Add(new DataGridViewTextBoxColumn { Name = "GhiChu", HeaderText = "Ghi chú", ReadOnly = true, FillWeight = 20 });
            dgvDelivery.Columns.Add(new DataGridViewTextBoxColumn { Name = "Status", HeaderText = "T.Thái", ReadOnly = true, FillWeight = 13 });
            dgvDelivery.Columns.Add(new DataGridViewTextBoxColumn { Name = "ReceiverNote", HeaderText = "Receiver", ReadOnly = false, FillWeight = 20 });

            // Màu sắc trạng thái
            dgvDelivery.CellFormatting += (s, ev) =>
            {
                if (ev.RowIndex < 0) return;
                if (dgvDelivery.Columns[ev.ColumnIndex].Name == "Status")
                {
                    string v = ev.Value?.ToString() ?? "";
                    ev.CellStyle.ForeColor =
                        v == "Done" ? Color.FromArgb(40, 167, 69) :
                        v == "Overdue" ? Color.FromArgb(220, 53, 69) :
                        Color.FromArgb(255, 140, 0);
                    ev.CellStyle.Font = new Font("Segoe UI", 8, FontStyle.Bold);
                }
            };
            dgvDelivery.RowPrePaint += (s, ev) =>
            {
                if (ev.RowIndex < 0 || dgvDelivery.Rows[ev.RowIndex].IsNewRow) return;
                string st = dgvDelivery.Rows[ev.RowIndex].Cells["Status"].Value?.ToString() ?? "";
                dgvDelivery.Rows[ev.RowIndex].DefaultCellStyle.BackColor =
                    st == "Done" ? Color.FromArgb(235, 255, 235) :
                    st == "Overdue" ? Color.FromArgb(255, 235, 235) :
                    Color.White;
            };
            panelDelivery.Controls.Add(dgvDelivery);

            // Double-click → hiện chi tiết vật tư PO
            dgvDelivery.CellDoubleClick += (s, ev) =>
            {
                if (ev.RowIndex < 0) return;
                string poNo = dgvDelivery.Rows[ev.RowIndex].Cells["PONo"].Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(poNo)) ShowDeliveryDetailPopup(poNo);
            };

            // Sự kiện các nút Delivery
            btnDelivAdd.Click += (s, e) => ShowDeliveryAddPopup();
            btnDelivDone.Click += (s, e) => MarkDeliveryDone();
            btnDelivDel.Click += (s, e) => DeleteDeliveryRow();

            // Timer tự xóa dòng quá hạn (kiểm tra mỗi giờ)
            _deliveryTimer = new System.Windows.Forms.Timer { Interval = 3_600_000 };
            _deliveryTimer.Tick += (s, e) => { UpdateOverdueDeliveries(); LoadDeliveries(); };
            _deliveryTimer.Start();

            // QUY HOẠCH CÁC Ô NHẬP LIỆU BÊN TRÁI (Tối đa width = 790px)
            int y = 38;

            // Row 1
            AddLabel(panelHeader, "PO No (*):", 10, y); txtPONo = AddTxt(panelHeader, 80, y, 100);
            AddLabel(panelHeader, "Tên dự án:", 190, y); txtProjectName = AddTxt(panelHeader, 260, y, 160);
            AddLabel(panelHeader, "Workorder:", 430, y); txtWorkorderNo = AddTxt(panelHeader, 505, y, 110);
            AddLabel(panelHeader, "MPR No:", 625, y); txtMPRNo = AddTxt(panelHeader, 685, y, 105);

            // Row 2
            y += 38;
            AddLabel(panelHeader, "Nhà CC:", 10, y);
            cboSupplier = new ComboBox { Location = new Point(80, y), Size = new Size(260, 25), Font = new Font("Segoe UI", 9), DropDownStyle = ComboBoxStyle.DropDown, AutoCompleteMode = AutoCompleteMode.None };
            panelHeader.Controls.Add(cboSupplier);
            cboSupplier.Validating += CboSupplier_Validating; cboSupplier.SelectedIndexChanged += CboSupplier_SelectedIndexChanged;
            cboSupplier.TextChanged += CboSupplier_TextChanged; cboSupplier.KeyDown += CboSupplier_KeyDown;
            LoadSupplierCombo();

            AddLabel(panelHeader, "Ngày PO:", 350, y); dtpPODate = new DateTimePicker { Location = new Point(410, y), Size = new Size(100, 25), Font = new Font("Segoe UI", 9), Format = DateTimePickerFormat.Short };
            panelHeader.Controls.Add(dtpPODate);
            AddLabel(panelHeader, "Trạng thái:", 520, y);
            cboStatus = new ComboBox { Location = new Point(590, y), Size = new Size(90, 25), Font = new Font("Segoe UI", 9), DropDownStyle = ComboBoxStyle.DropDownList };
            cboStatus.Items.AddRange(new[] { "Draft", "Pending", "Approved", "In Progress", "Completed", "Cancelled" });
            cboStatus.SelectedIndex = 0; panelHeader.Controls.Add(cboStatus);
            AddLabelCus(panelHeader, "Revise:", 690, y, 45, 20);
            nudRevise = new NumericUpDown { Location = new Point(740, y), Size = new Size(50, 25), Font = new Font("Segoe UI", 9), Minimum = 0, Maximum = 99 };
            nudRevise.BringToFront(); panelHeader.Controls.Add(nudRevise);

            // Row 3
            y += 38;
            AddLabel(panelHeader, "Prepared:", 10, y); txtPrepared = AddTxt(panelHeader, 80, y, 100);
            AddLabel(panelHeader, "Reviewed:", 190, y); txtReviewed = AddTxt(panelHeader, 260, y, 110);
            AddLabel(panelHeader, "Agreement:", 380, y); txtAgreement = AddTxt(panelHeader, 455, y, 110);
            AddLabel(panelHeader, "Approved:", 575, y); txtApproved = AddTxt(panelHeader, 645, y, 145);

            // Row 4
            y += 38;
            AddLabel(panelHeader, "Ghi chú:", 10, y);
            txtNotes = AddTxt(panelHeader, 80, y, 200);
            txtNotes.Anchor = AnchorStyles.Top | AnchorStyles.Left;

            // Row 5 (Buttons)
            y += 45;
            btnSavePO = CreateButton("💾 Lưu Toàn Bộ PO", Color.FromArgb(0, 120, 212), new Point(10, y), 150, 32);
            btnSavePO.Click += BtnSavePO_Click; panelHeader.Controls.Add(btnSavePO);
            btnClearHeader = CreateButton("Làm mới", Color.FromArgb(108, 117, 125), new Point(170, y), 100, 32);
            btnClearHeader.Click += BtnClearHeader_Click; panelHeader.Controls.Add(btnClearHeader);
            var btnImportMPR = CreateButton("Import MPR", Color.FromArgb(255, 140, 0), new Point(280, y), 120, 32);
            btnImportMPR.Click += BtnImportMPR_Click; panelHeader.Controls.Add(btnImportMPR);
            var btnHistory = CreateButton("Revise History", Color.FromArgb(102, 51, 153), new Point(410, y), 130, 32);
            btnHistory.Click += (s, e) => { if (string.IsNullOrEmpty(txtPONo.Text)) { MessageBox.Show("Vui lòng chọn một PO trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; } new frmReviseHistory(txtPONo.Text).ShowDialog(); };
            panelHeader.Controls.Add(btnHistory);

            // ===== PANEL DETAIL =====
            panelDetail = new Panel { Location = new Point(10, 500), Size = new Size(1260, 285), BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom };
            this.Controls.Add(panelDetail);
            panelDetail.Controls.Add(new Label { Text = "CHI TIẾT ĐƠN HÀNG", Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212), Location = new Point(10, 8), Size = new Size(300, 25) });

            btnAddDetail = CreateButton("+ Thêm dòng", Color.FromArgb(40, 167, 69), new Point(10, 38), 120, 30);
            btnDeleteDetail = CreateButton("Xóa dòng", Color.FromArgb(220, 53, 69), new Point(140, 38), 100, 30);
            var btnSaveDetail = CreateButton("💾 Lưu chi tiết", Color.FromArgb(0, 120, 212), new Point(250, 38), 130, 30);
            btnExport = CreateButton("📄 Xuất Excel", Color.FromArgb(0, 150, 100), new Point(390, 38), 130, 30);
            var btnCheckBySize = CreateButton("🔍 Check by size", Color.FromArgb(102, 51, 153), new Point(530, 38), 145, 30);

            btnAddDetail.Click += BtnAddDetail_Click;
            btnDeleteDetail.Click += BtnDeleteDetail_Click;
            btnSaveDetail.Click += BtnSaveDetail_Click;
            btnExport.Click += BtnExport_Click;
            btnCheckBySize.Click += BtnCheckBySize_Click;

            panelDetail.Controls.Add(btnAddDetail); panelDetail.Controls.Add(btnDeleteDetail);
            panelDetail.Controls.Add(btnSaveDetail); panelDetail.Controls.Add(btnExport);
            panelDetail.Controls.Add(btnCheckBySize);

            // lblSubTotal và lblTotal căn mép PHẢI panelDetail — không bị button che
            lblSubTotal = new Label
            {
                Size = new Size(220, 22),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(220, 53, 69),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            lblSubTotal.Location = new Point(panelDetail.Width - 220 - 280 - 10, 45);
            panelDetail.Controls.Add(lblSubTotal);

            lblTotal = new Label
            {
                Size = new Size(270, 22),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            lblTotal.Location = new Point(panelDetail.Width - 270 - 10, 45);
            panelDetail.Controls.Add(lblTotal);

            dgvDetails = new DataGridView
            {
                Location = new Point(10, 75),
                Size = new Size(1235, 195),
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
            dgvDetails.CellEndEdit += DgvDetails_CellEndEdit; dgvDetails.KeyDown += DgvDetails_KeyDown;
            dgvDetails.CurrentCellDirtyStateChanged += DgvDetails_CurrentCellDirtyStateChanged;
            dgvDetails.CellValueChanged += DgvDetails_CellValueChanged;
            dgvDetails.CellFormatting += DgvDetails_CellFormatting;
            dgvDetails.DataError += DgvDetails_DataError;

            BuildDetailColumns(); panelDetail.Controls.Add(dgvDetails);

            // ── Đảm bảo tất cả TextBox, ComboBox, DateTimePicker, NumericUpDown
            //    trong mọi panel đều hiển thị trên Label ──
            BringInputsToFront(panelTop);
            BringInputsToFront(panelHeader);
            BringInputsToFront(panelDetail);
        }

        // =====================================================================
        // HELPER — Đưa tất cả input controls lên trên label trong một panel
        // =====================================================================
        private static void BringInputsToFront(Control parent)
        {
            foreach (Control c in parent.Controls)
                if (c is TextBox || c is ComboBox || c is DateTimePicker || c is NumericUpDown || c is CheckBox)
                    c.BringToFront();
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
        // HÀM QUÉT THƯ MỤC LẤY DANH SÁCH FILE TỪ PO LINK
        // =====================================================================
        private void LoadFiles(string workorderNo, string projectName)
        {
            dgvFiles.Rows.Clear();
            if (string.IsNullOrEmpty(workorderNo) && string.IsNullOrEmpty(projectName)) return;

            try
            {
                var projects = new ProjectService().GetAll();
                var prj = projects.Find(p =>
                    (!string.IsNullOrEmpty(p.WorkorderNo) && p.WorkorderNo.Equals(workorderNo, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(p.ProjectName) && p.ProjectName.Equals(projectName, StringComparison.OrdinalIgnoreCase))
                );

                if (prj != null && !string.IsNullOrEmpty(prj.PO_Link) && Directory.Exists(prj.PO_Link))
                {
                    var files = Directory.GetFiles(prj.PO_Link);
                    foreach (var f in files)
                    {
                        dgvFiles.Rows.Add(Path.GetFileName(f), f);
                    }
                }
                else if (prj != null && !string.IsNullOrEmpty(prj.PO_Link))
                {
                    dgvFiles.Rows.Add("(Thư mục không tồn tại)", "");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Lỗi load files: " + ex.Message);
            }
        }

        // ĐÃ KHẮC PHỤC LỖI CS7036: Chèn đúng danh sách _poList
        private void BtnSearchBySupp_Click(object sender, EventArgs e)
        {
            try
            {
                var suppliers = new SupplierService().GetAll();
                var dt = new System.Data.DataTable();
                dt.Columns.Add("PO_ID", typeof(int));
                dt.Columns.Add("PO No", typeof(string));
                dt.Columns.Add("NCC", typeof(string));
                dt.Columns.Add("Dự án", typeof(string));
                dt.Columns.Add("MPR No", typeof(string));
                dt.Columns.Add("Workorder", typeof(string));
                dt.Columns.Add("Ngày PO", typeof(string));
                dt.Columns.Add("Trạng thái", typeof(string));
                dt.Columns.Add("Tổng tiền", typeof(string));
                foreach (var h in _poList)
                {
                    var supp = suppliers.Find(s => s.Supplier_ID == h.Supplier_ID);
                    dt.Rows.Add(h.PO_ID, h.PONo, supp?.Company_Name ?? supp?.Short_Name ?? "", h.Project_Name, h.MPR_No, h.WorkorderNo,
                        h.PO_Date.HasValue ? h.PO_Date.Value.ToString("dd/MM/yyyy") : "", h.Status, h.Total_Amount.ToString("N0"));
                }
                var dtFull = dt.Copy();
                System.Data.DataTable dtCurrent = dtFull.Copy();
                string selectedPONo = null;

                var popup = new Form { Text = "🔍 Tìm theo NCC", Size = new Size(1100, 640), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(245, 245, 245), MinimumSize = new Size(800, 450) };
                popup.Controls.Add(new Label { Text = "🔍  TÌM KIẾM PO THEO NHÀ CUNG CẤP", Font = new Font("Segoe UI", 11, FontStyle.Bold), ForeColor = Color.FromArgb(102, 51, 153), Location = new Point(10, 8), Size = new Size(700, 26) });

                var pF = new Panel { Location = new Point(10, 38), Size = new Size(popup.ClientSize.Width - 20, 46), BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
                popup.Controls.Add(pF);
                pF.Controls.Add(new Label { Text = "NCC:", Location = new Point(8, 12), Size = new Size(35, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold) });
                var txtNCC = new TextBox { Location = new Point(43, 8), Size = new Size(200, 26), Font = new Font("Segoe UI", 9), PlaceholderText = "Tên nhà cung cấp..." };
                pF.Controls.Add(txtNCC);
                pF.Controls.Add(new Label { Text = "Dự án:", Location = new Point(258, 12), Size = new Size(45, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold) });
                var txtDA = new TextBox { Location = new Point(303, 8), Size = new Size(160, 26), Font = new Font("Segoe UI", 9), PlaceholderText = "Tên dự án..." };
                pF.Controls.Add(txtDA);
                pF.Controls.Add(new Label { Text = "T.Thái:", Location = new Point(475, 12), Size = new Size(48, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold) });
                var cboTT = new ComboBox { Location = new Point(523, 8), Size = new Size(115, 26), Font = new Font("Segoe UI", 9), DropDownStyle = ComboBoxStyle.DropDownList };
                cboTT.Items.AddRange(new[] { "Tất cả", "Draft", "Pending", "Approved", "In Progress", "Completed", "Cancelled" });
                cboTT.SelectedIndex = 0;
                pF.Controls.Add(cboTT);
                var btnF = new Button { Text = "🔍 Tìm", Location = new Point(648, 8), Size = new Size(80, 28), BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Cursor = Cursors.Hand };
                btnF.FlatAppearance.BorderSize = 0; pF.Controls.Add(btnF);
                var btnClear = new Button { Text = "✖ Xóa", Location = new Point(736, 8), Size = new Size(75, 28), BackColor = Color.FromArgb(108, 117, 125), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Cursor = Cursors.Hand };
                btnClear.FlatAppearance.BorderSize = 0; pF.Controls.Add(btnClear);

                var lblCount = new Label { Text = "", Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212), Location = new Point(10, 90), Size = new Size(500, 20), Anchor = AnchorStyles.Top | AnchorStyles.Left };
                popup.Controls.Add(lblCount);

                var dgv = new DataGridView { Location = new Point(10, 114), Size = new Size(popup.ClientSize.Width - 20, popup.ClientSize.Height - 165), ReadOnly = true, AllowUserToAddRows = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, BackgroundColor = Color.White, BorderStyle = BorderStyle.FixedSingle, RowHeadersVisible = false, Font = new Font("Segoe UI", 9), AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom };
                dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(102, 51, 153); dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White; dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold); dgv.EnableHeadersVisualStyles = false;
                dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 240, 255);
                popup.Controls.Add(dgv);

                Action applyFilter = () =>
                {
                    string kNCC = txtNCC.Text.Trim().ToLower(), kDA = txtDA.Text.Trim().ToLower(), kTT = cboTT.SelectedItem?.ToString() ?? "Tất cả";
                    var rows = dtFull.AsEnumerable().Where(r =>
                    {
                        if (!string.IsNullOrEmpty(kNCC) && !r["NCC"].ToString().ToLower().Contains(kNCC)) return false;
                        if (!string.IsNullOrEmpty(kDA) && !r["Dự án"].ToString().ToLower().Contains(kDA)) return false;
                        if (kTT != "Tất cả" && r["Trạng thái"].ToString() != kTT) return false;
                        return true;
                    });
                    dtCurrent = rows.Any() ? rows.CopyToDataTable() : dtFull.Clone();
                    dgv.DataSource = dtCurrent;
                    if (dgv.Columns.Contains("PO_ID")) dgv.Columns["PO_ID"].Visible = false;
                    lblCount.Text = $"Hiển thị: {dtCurrent.Rows.Count} / {dtFull.Rows.Count} PO";
                };
                applyFilter();
                btnF.Click += (s, ev) => applyFilter();
                btnClear.Click += (s, ev) => { txtNCC.Text = ""; txtDA.Text = ""; cboTT.SelectedIndex = 0; applyFilter(); };
                popup.KeyPreview = true;
                popup.KeyDown += (s, ev) => { if (ev.KeyCode == Keys.Enter) { applyFilter(); ev.SuppressKeyPress = true; } };
                cboTT.SelectedIndexChanged += (s, ev) => applyFilter();
                dgv.CellDoubleClick += (s, ev) => { if (ev.RowIndex < 0) return; selectedPONo = dgv.Rows[ev.RowIndex].Cells["PO No"].Value?.ToString(); popup.DialogResult = DialogResult.OK; popup.Close(); };

                int btnY = popup.ClientSize.Height - 42;
                var btnSel = new Button { Text = "✔ Chọn PO này", Location = new Point(10, btnY), Size = new Size(130, 32), BackColor = Color.FromArgb(40, 167, 69), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
                btnSel.FlatAppearance.BorderSize = 0;
                btnSel.Click += (s, ev) => { if (dgv.SelectedRows.Count == 0) { MessageBox.Show("Vui lòng chọn một dòng!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; } selectedPONo = dgv.SelectedRows[0].Cells["PO No"].Value?.ToString(); popup.DialogResult = DialogResult.OK; popup.Close(); };
                popup.Controls.Add(btnSel);

                var btnExp = new Button { Text = "📥 Xuất Excel", Location = new Point(150, btnY), Size = new Size(130, 32), BackColor = Color.FromArgb(0, 150, 100), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
                btnExp.FlatAppearance.BorderSize = 0;
                btnExp.Click += (s, ev) =>
                {
                    if (dtCurrent == null || dtCurrent.Rows.Count == 0) { MessageBox.Show("Không có dữ liệu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                    using var sfd = new SaveFileDialog { Title = "Xuất chi tiết PO theo NCC", Filter = "Excel Files|*.xlsx", FileName = $"PO_ChiTiet_{DateTime.Now:yyyyMMdd_HHmm}", InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) };
                    if (sfd.ShowDialog() != DialogResult.OK) return;
                    try
                    {
                        OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                        using var pkg = new OfficeOpenXml.ExcelPackage();
                        int totalRows = 0;
                        foreach (System.Data.DataRow dr in dtCurrent.Rows)
                        {
                            int poId = Convert.ToInt32(dr["PO_ID"]);
                            string poNo = dr["PO No"]?.ToString() ?? "";
                            string sheetName = System.Text.RegularExpressions.Regex.Replace(poNo, @"[\\\/\?\*\[\]:]", "_");
                            if (sheetName.Length > 31) sheetName = sheetName.Substring(0, 31);
                            var ws = pkg.Workbook.Worksheets.Add(sheetName);

                            ws.Cells[1, 1].Value = $"CHI TIẾT ĐƠN HÀNG — {poNo}";
                            ws.Cells[1, 1, 1, 11].Merge = true;
                            ws.Cells[1, 1].Style.Font.Bold = true; ws.Cells[1, 1].Style.Font.Size = 13;
                            ws.Cells[1, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(102, 51, 153));
                            ws.Cells[1, 1].Style.Font.Color.SetColor(Color.White);
                            ws.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                            void WI(int r, string lbl, string val) { ws.Cells[r, 1].Value = lbl; ws.Cells[r, 1].Style.Font.Bold = true; ws.Cells[r, 2, r, 5].Merge = true; ws.Cells[r, 2].Value = val; }
                            WI(2, "PO No:", dr["PO No"]?.ToString() ?? ""); WI(3, "NCC:", dr["NCC"]?.ToString() ?? "");
                            WI(4, "Dự án:", dr["Dự án"]?.ToString() ?? ""); WI(5, "MPR No:", dr["MPR No"]?.ToString() ?? "");
                            WI(6, "Workorder:", dr["Workorder"]?.ToString() ?? ""); WI(7, "Ngày PO:", dr["Ngày PO"]?.ToString() ?? "");
                            WI(8, "Trạng thái:", dr["Trạng thái"]?.ToString() ?? "");

                            string[] hdrs = { "STT", "Tên hàng", "Vật liệu", "A(mm)", "B(mm)", "C(mm)", "SL", "ĐVT", "KG", "Đơn giá", "VAT(%)", "Thành tiền" };
                            for (int c = 0; c < hdrs.Length; c++) { ws.Cells[10, c + 1].Value = hdrs[c]; ws.Cells[10, c + 1].Style.Font.Bold = true; ws.Cells[10, c + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid; ws.Cells[10, c + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 120, 212)); ws.Cells[10, c + 1].Style.Font.Color.SetColor(Color.White); }

                            var dets = _service.GetDetails(poId); int dRow = 11; decimal sub = 0;
                            for (int i = 0; i < dets.Count; i++)
                            {
                                var d = dets[i]; decimal rp = d.Price, bv = d.Qty_Per_Sheet;
                                if ((d.Remarks ?? "").Contains("[CALC:KG]") && d.Weight_kg > 0 && d.Qty_Per_Sheet > 0) { rp = Math.Round((d.Price * d.Qty_Per_Sheet) / d.Weight_kg, 0); bv = d.Weight_kg; }
                                decimal amt = Math.Round(bv * rp, 0); sub += amt;
                                ws.Cells[dRow, 1].Value = i + 1; ws.Cells[dRow, 2].Value = d.Item_Name ?? ""; ws.Cells[dRow, 3].Value = d.Material ?? "";
                                ws.Cells[dRow, 4].Value = d.Asize; ws.Cells[dRow, 5].Value = d.Bsize; ws.Cells[dRow, 6].Value = d.Csize;
                                ws.Cells[dRow, 7].Value = d.Qty_Per_Sheet; ws.Cells[dRow, 8].Value = d.UNIT ?? ""; ws.Cells[dRow, 9].Value = d.Weight_kg;
                                ws.Cells[dRow, 10].Value = rp; ws.Cells[dRow, 10].Style.Numberformat.Format = "#,##0";
                                ws.Cells[dRow, 11].Value = d.VAT; ws.Cells[dRow, 12].Value = amt; ws.Cells[dRow, 12].Style.Numberformat.Format = "#,##0";
                                if (i % 2 == 1) for (int c = 1; c <= 12; c++) { ws.Cells[dRow, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid; ws.Cells[dRow, c].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(240, 248, 255)); }
                                dRow++; totalRows++;
                            }
                            decimal vat = Math.Round(sub * 0.1m, 0);
                            ws.Cells[dRow, 11].Value = "Sub-Total:"; ws.Cells[dRow, 11].Style.Font.Bold = true; ws.Cells[dRow, 12].Value = sub; ws.Cells[dRow, 12].Style.Numberformat.Format = "#,##0"; ws.Cells[dRow, 12].Style.Font.Bold = true; dRow++;
                            ws.Cells[dRow, 11].Value = "VAT (10%):"; ws.Cells[dRow, 12].Value = vat; ws.Cells[dRow, 12].Style.Numberformat.Format = "#,##0"; dRow++;
                            ws.Cells[dRow, 10, dRow, 11].Merge = true; ws.Cells[dRow, 10].Value = "TOTAL (incl. VAT):"; ws.Cells[dRow, 10].Style.Font.Bold = true;
                            ws.Cells[dRow, 12].Value = sub + vat; ws.Cells[dRow, 12].Style.Numberformat.Format = "#,##0"; ws.Cells[dRow, 12].Style.Font.Bold = true;
                            ws.Cells[ws.Dimension.Address].AutoFitColumns();
                        }
                        pkg.SaveAs(new System.IO.FileInfo(sfd.FileName));
                        MessageBox.Show($"✅ Đã xuất {dtCurrent.Rows.Count} PO với {totalRows} dòng chi tiết!\nMỗi PO = 1 sheet riêng.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex2) { MessageBox.Show("Lỗi xuất Excel: " + ex2.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                };
                popup.Controls.Add(btnExp);

                var btnClose = new Button { Text = "Đóng", Size = new Size(100, 32), BackColor = Color.FromArgb(108, 117, 125), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Anchor = AnchorStyles.Bottom | AnchorStyles.Right, DialogResult = DialogResult.Cancel };
                btnClose.FlatAppearance.BorderSize = 0; btnClose.Location = new Point(popup.ClientSize.Width - 115, btnY);
                popup.Controls.Add(btnClose); popup.CancelButton = btnClose;

                popup.Resize += (s, ev) => { pF.Width = popup.ClientSize.Width - 20; dgv.Size = new Size(popup.ClientSize.Width - 20, popup.ClientSize.Height - 165); btnSel.Location = new Point(10, popup.ClientSize.Height - 42); btnExp.Location = new Point(150, popup.ClientSize.Height - 42); btnClose.Location = new Point(popup.ClientSize.Width - 115, popup.ClientSize.Height - 42); };
                if (popup.ShowDialog(this) == DialogResult.OK && !string.IsNullOrEmpty(selectedPONo)) SelectPOByNo(selectedPONo);
            }
            catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void AutoAdjustColumnWidths()
        {
            if (dgvDetails.Columns.Count == 0) return;
            dgvDetails.SuspendLayout();
            dgvDetails.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            using (Graphics g = dgvDetails.CreateGraphics())
            {
                Font headerFont = dgvDetails.ColumnHeadersDefaultCellStyle.Font ?? dgvDetails.Font;
                foreach (DataGridViewColumn col in dgvDetails.Columns)
                {
                    if (!col.Visible) continue;
                    col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    SizeF headerSize = g.MeasureString(col.HeaderText, headerFont);
                    int minWidth = (int)Math.Ceiling(headerSize.Width) + 20; int maxWidth = minWidth;
                    foreach (DataGridViewRow row in dgvDetails.Rows)
                    {
                        if (row.IsNewRow) continue;
                        string cellValue = row.Cells[col.Index].Value?.ToString() ?? "";
                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            SizeF size = g.MeasureString(cellValue, dgvDetails.Font); int cw = (int)Math.Ceiling(size.Width) + 15;
                            if (cw > maxWidth) maxWidth = cw;
                        }
                    }
                    if (maxWidth > 200) { col.Width = 200; col.DefaultCellStyle.WrapMode = DataGridViewTriState.True; }
                    else { col.Width = maxWidth; col.DefaultCellStyle.WrapMode = DataGridViewTriState.False; }
                }
            }
            dgvDetails.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            dgvDetails.ResumeLayout();
        }

        private void DgvDetails_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (!dgvDetails.IsCurrentCellDirty) return;
            string col = dgvDetails.CurrentCell?.OwningColumn.Name ?? "";
            // Commit ngay khi chọn giá trị ComboBox (VAT và Calc_Method)
            if (col == "Calc_Method" || col == "VAT")
                dgvDetails.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void DgvDetails_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            // Bỏ qua lỗi ComboBox value not valid — xảy ra khi giá trị chưa khớp Items
            // Tự động sửa bằng cách map về giá trị hợp lệ
            if (e.ColumnIndex >= 0 && e.RowIndex >= 0)
            {
                string col = dgvDetails.Columns[e.ColumnIndex].Name;
                if (col == "VAT")
                {
                    var cell = dgvDetails.Rows[e.RowIndex].Cells["VAT"];
                    string cur = cell.Value?.ToString() ?? "";
                    // Map về "8" hoặc "10", mặc định "10"
                    cell.Value = cur == "8" ? "8" : "10";
                    e.ThrowException = false;
                    return;
                }
                if (col == "Calc_Method")
                {
                    var cell = dgvDetails.Rows[e.RowIndex].Cells["Calc_Method"];
                    string cur = cell.Value?.ToString() ?? "";
                    cell.Value = cur == "Theo SL" ? "Theo SL" : "Theo KG";
                    e.ThrowException = false;
                    return;
                }
            }
            e.ThrowException = false; // Suppress mọi DataError không mong muốn
        }

        private void DgvDetails_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        { if (e.RowIndex >= 0 && dgvDetails.Columns[e.ColumnIndex].Name == "Calc_Method") RecalculateAmount(e.RowIndex); }

        private void DgvDetails_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string colName = dgvDetails.Columns[e.ColumnIndex].Name;

            // Cột Đơn giá và Thành tiền — định dạng số có dấu phân cách, căn phải
            if (colName == "Price" || colName == "Amount")
            {
                if (e.Value != null && decimal.TryParse(e.Value.ToString(), out decimal num))
                {
                    e.Value = num.ToString("N0");
                    e.FormattingApplied = true;
                }
                e.CellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }

            if (colName == "Ordered_PO")
            {
                string val = e.Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(val))
                {
                    e.CellStyle.ForeColor = Color.FromArgb(220, 53, 69);
                    e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                }
            }
        }

        public static void AutoCompleteComboboxValidating(ComboBox sender, CancelEventArgs e)
        {
            var cb = sender as ComboBox;
            string typedText = cb.Text?.Trim();
            if (string.IsNullOrEmpty(typedText)) { cb.SelectedIndex = 0; return; }
            foreach (var item in cb.Items)
            {
                if (item is DataRowView drv)
                {
                    string value = drv[cb.DisplayMember]?.ToString();
                    if (value != null && value.Equals(typedText, StringComparison.OrdinalIgnoreCase)) { cb.SelectedItem = item; return; }
                }
            }
            cb.SelectedIndex = 0;
        }

        private void LoadSupplierCombo()
        { try { _supplierTable = new SupplierService().GetForCombo(); BindSupplierCombo(_supplierTable); } catch (Exception ex) { MessageBox.Show("Lỗi tải nhà cung cấp: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); } }

        private void BindSupplierCombo(System.Data.DataTable dt)
        { _isSearching = true; string ct = cboSupplier.Text; cboSupplier.DataSource = null; cboSupplier.DataSource = dt; cboSupplier.DisplayMember = "Name"; cboSupplier.ValueMember = "ID"; cboSupplier.Text = ct; _isSearching = false; }

        private void CboSupplier_TextChanged(object sender, EventArgs e)
        {
            if (_isSearching) return;
            // KHÔNG dùng .Trim() để giữ khoảng trắng khi tìm kiếm
            string keyword = cboSupplier.Text;
            if (string.IsNullOrEmpty(keyword)) { BindSupplierCombo(_supplierTable); cboSupplier.DroppedDown = false; return; }
            string kn = RemoveDiacritics(keyword).ToLower();
            var filtered = new System.Data.DataTable(); filtered.Columns.Add("ID", typeof(int)); filtered.Columns.Add("Name", typeof(string));
            foreach (System.Data.DataRow row in _supplierTable.Rows)
            {
                string name = row["Name"].ToString();
                if (RemoveDiacritics(name).ToLower().Contains(kn) || name.ToLower().Contains(keyword.ToLower())) filtered.Rows.Add(row["ID"], row["Name"]);
            }
            if (filtered.Rows.Count == 0)
            {
                var empty = new System.Data.DataTable();
                empty.Columns.Add("ID", typeof(int)); empty.Columns.Add("Name", typeof(string)); empty.Rows.Add(0, "-- Không tìm thấy --"); BindSupplierCombo(empty);
            }
            else BindSupplierCombo(filtered);
            _isSearching = true; cboSupplier.Text = keyword;
            cboSupplier.SelectionStart = keyword.Length; cboSupplier.DroppedDown = true; _isSearching = false;
        }

        private string RemoveDiacritics(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            try
            {
                string n = text.Normalize(System.Text.NormalizationForm.FormD); var sb = new System.Text.StringBuilder();
                foreach (char c in n) if (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c) != System.Globalization.UnicodeCategory.NonSpacingMark) sb.Append(c); return sb.ToString().Normalize(System.Text.NormalizationForm.FormC);
            }
            catch { return text; }
        }

        private void CboSupplier_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                if (cboSupplier.DroppedDown && cboSupplier.Items.Count > 0)
                {
                    if (cboSupplier.SelectedIndex >= 0)
                    {
                        int sid = Convert.ToInt32(cboSupplier.SelectedValue ?? 0);
                        if (sid > 0)
                        {
                            cboSupplier.DroppedDown = false; _isSearching = true; BindSupplierCombo(_supplierTable); cboSupplier.SelectedValue = sid; _isSearching = false;
                            cboSupplier.BackColor = Color.White; e.SuppressKeyPress = true; e.Handled = true; return;
                        }
                    }
                    string kw = cboSupplier.Text;
                    string kwn = RemoveDiacritics(kw).ToLower(); int matchId = 0;
                    foreach (System.Data.DataRowView drv in cboSupplier.Items)
                    {
                        string name = drv["Name"].ToString();
                        int id = Convert.ToInt32(drv["ID"]); if (id > 0 && (RemoveDiacritics(name).ToLower().Contains(kwn) || name.ToLower().Contains(kw.ToLower()))) { matchId = id; break; }
                    }
                    if (matchId > 0)
                    {
                        cboSupplier.DroppedDown = false;
                        _isSearching = true; BindSupplierCombo(_supplierTable); cboSupplier.SelectedValue = matchId; _isSearching = false; cboSupplier.BackColor = Color.White;
                    }
                    else cboSupplier.BackColor = Color.FromArgb(255, 230, 230);
                }
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
            if (e.KeyCode == Keys.Escape)
            {
                _isSearching = true;
                BindSupplierCombo(_supplierTable); cboSupplier.Text = ""; cboSupplier.DroppedDown = false; cboSupplier.BackColor = Color.White; _isSearching = false;
            }
        }

        private void CboSupplier_Validating(object sender, System.ComponentModel.CancelEventArgs e) => AutoCompleteComboboxValidating(sender as ComboBox, e);
        private void CboSupplier_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_isSearching) return;
            if (cboSupplier.SelectedValue == null) return;
            int supplierId = Convert.ToInt32(cboSupplier.SelectedValue);
            if (supplierId == 0) { cboSupplier.BackColor = Color.White; return; }
            try
            {
                cboSupplier.BackColor = Color.White; _isSearching = true;
                BindSupplierCombo(_supplierTable); cboSupplier.SelectedValue = supplierId; _isSearching = false;
            }
            catch { }
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (_selectedPO_ID == 0) { MessageBox.Show("Vui lòng chọn PO cần xuất!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            try
            {
                var po = _poList.Find(p => p.PO_ID == _selectedPO_ID);
                var details = _service.GetDetails(_selectedPO_ID); if (po == null) return;
                var suppliers = new SupplierService().GetAll();
                var supplier = suppliers.Find(s => s.Supplier_ID == Convert.ToInt32(cboSupplier.SelectedValue?.ToString() ?? "0"));
                var projects = new ProjectService().GetAll();
                var project = projects.Find(p => p.WorkorderNo == po.WorkorderNo);
                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "po_template.xlsx");
                if (!File.Exists(templatePath)) { MessageBox.Show($"Lỗi: Không tìm thấy file template!\nĐường dẫn dự kiến: {templatePath}", "Lỗi hệ thống", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                var saveDialog = new SaveFileDialog { Title = "Lưu file PO", Filter = "Excel Files|*.xlsx", FileName = $"PO_{po.PONo}_{DateTime.Now:ddMMyyyy}", InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) };
                if (saveDialog.ShowDialog() != DialogResult.OK) return;
                File.Copy(templatePath, saveDialog.FileName, true);
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(saveDialog.FileName)))
                {
                    var ws = package.Workbook.Worksheets[0];
                    ReplaceCell(ws, "<<PROJECT_NAME>>", project?.ProjectName ?? po.Project_Name ?? ""); ReplaceCell(ws, "<<WO-NO>>", po.WorkorderNo ?? ""); ReplaceCell(ws, "<<REV.NUM>>", po.Revise.ToString() ?? "0");
                    ReplaceCell(ws, "<<DATE>>", po.PO_Date.HasValue ? po.PO_Date.Value.ToString("dd/MM/yyyy") : DateTime.Today.ToString("dd/MM/yyyy")); ReplaceCell(ws, "<<MPR-NO>>", po.MPR_No ?? ""); ReplaceCell(ws, "<<PO-NO>>", po.PONo ?? "");
                    string supplierInfo = supplier != null ? $"{supplier.Company_Name}\nCert: {supplier.Cert ?? ""}\nEmail: {supplier.Email}" : "";
                    ReplaceCell(ws, "<<SUPPLIER-INFO>>", supplierInfo);
                    int startRow = 8; int detailCount = details.Count;
                    if (detailCount > 1) ws.InsertRow(startRow + 1, detailCount - 1);
                    for (int i = 0; i < detailCount; i++)
                    {
                        var d = details[i];
                        int row = startRow + i;
                        decimal q = d.Qty_Per_Sheet; decimal wk = d.Weight_kg; decimal realPrice = d.Price;
                        string rem = d.Remarks ?? "";
                        if (rem.Contains("[CALC:KG]"))
                        {
                            rem = rem.Replace("[CALC:KG]", "").Trim();
                            if (wk > 0 && q > 0) realPrice = (d.Price * q) / wk;
                        }
                        else if (rem.Contains("[CALC:SL]")) rem = rem.Replace("[CALC:SL]", "").Trim();
                        ws.Cells[row, 1].Value = i + 1; ws.Cells[row, 2].Value = d.Item_Name ?? ""; ws.Cells[row, 3].Value = d.Material ?? "";
                        ws.Cells[row, 4].Value = d.Asize; ws.Cells[row, 5].Value = d.Bsize; ws.Cells[row, 6].Value = d.Csize;
                        ws.Cells[row, 7].Value = d.Qty_Per_Sheet;
                        ws.Cells[row, 8].Value = d.UNIT ?? ""; ws.Cells[row, 9].Value = d.Weight_kg;
                        ws.Cells[row, 10].Value = d.MPSNo ?? ""; ws.Cells[row, 11].Value = d.RequestDay;
                        ws.Cells[row, 12].Value = "Kho DLHI";
                        ws.Cells[row, 13].Value = Math.Round(realPrice, 0); ws.Cells[row, 14].Value = d.Amount; ws.Cells[row, 16].Value = rem;
                        if (i > 0)
                        {
                            for (int col = 1; col <= 16; col++)
                            {
                                ws.Cells[row, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                ws.Cells[row, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                ws.Cells[row, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; ws.Cells[row, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                ws.Cells[row, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; ws.Cells[row, col].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                ws.Cells[row, col].Style.Font.Name = "Arial"; ws.Cells[row, col].Style.Font.Size = 9;
                            }
                            ws.Cells[row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        }
                    }
                    int subTotalRow = startRow + detailCount;
                    int vatRow = subTotalRow + 1;
                    ws.Cells[subTotalRow, 3].Value = "SUB-TOTAL"; ws.Cells[subTotalRow, 9].Value = details.Sum(d => (double)d.Weight_kg);
                    ws.Cells[subTotalRow, 14].Formula = $"=SUM(N{startRow}:N{startRow + detailCount - 1})";
                    ws.Cells[vatRow, 3].Value = "Final Price Requested (Included 10% VAT)";
                    ws.Cells[vatRow, 14].Formula = $"=N{subTotalRow}*1.1";
                    for (int r = 1; r <= ws.Dimension.End.Row; r++) for (int c = 1; c <= ws.Dimension.End.Column; c++) if (ws.Cells[r, c].Value?.ToString() == "<<DATE>>") ws.Cells[r, c].Value = DateTime.Today.ToString("dd/MM/yyyy");
                    package.Save();
                }
                var result = MessageBox.Show($"✅ Xuất Excel thành công!\nFile: {saveDialog.FileName}\n\nBạn có muốn mở file không?", "Thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (result == DialogResult.Yes) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = saveDialog.FileName, UseShellExecute = true });
            }
            catch (Exception ex) { MessageBox.Show("Lỗi xuất Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void ReplaceCell(OfficeOpenXml.ExcelWorksheet ws, string placeholder, string value)
        { for (int r = 1; r <= ws.Dimension.End.Row; r++) for (int c = 1; c <= ws.Dimension.End.Column; c++) if (ws.Cells[r, c].Value?.ToString() == placeholder) ws.Cells[r, c].Value = value; }

        private void BuildDetailColumns()
        {
            dgvDetails.Columns.Clear();
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "DeliveryLocation", HeaderText = "Nơi giao" });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Item_No", HeaderText = "STT", ReadOnly = true });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Item_Name", HeaderText = "Tên hàng" });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Material", HeaderText = "Vật liệu" });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Asize", HeaderText = "A(mm)" });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Bsize", HeaderText = "B(mm)" });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Csize", HeaderText = "C(mm)" });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Qty", HeaderText = "SL" });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "UNIT", HeaderText = "ĐVT" });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Weight", HeaderText = "KG" });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Price", HeaderText = "Đơn giá" });

            // VAT — ComboBox chọn 8 hoặc 10, mặc định 10
            var colVAT = new DataGridViewComboBoxColumn
            {
                Name = "VAT",
                HeaderText = "VAT(%)",
                Width = 70,
                FlatStyle = FlatStyle.Flat
            };
            colVAT.Items.AddRange("10", "8");
            dgvDetails.Columns.Add(colVAT);

            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Amount", HeaderText = "Thành tiền", ReadOnly = true });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Received", HeaderText = "Đã nhận" });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "MPSNo", HeaderText = "MPS No" });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Remarks", HeaderText = "Ghi chú" });

            // Cách tính — mặc định Theo KG
            var colCalc = new DataGridViewComboBoxColumn { Name = "Calc_Method", HeaderText = "Cách tính" };
            colCalc.Items.AddRange("Theo KG", "Theo SL");
            dgvDetails.Columns.Add(colCalc);

            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ordered_PO", HeaderText = "Đã lên PO", ReadOnly = true });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "PO_Detail_ID", HeaderText = "PO_ID", Visible = false });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "MPR_Detail_ID", HeaderText = "MPR_Detail_ID", Visible = false });
            foreach (DataGridViewColumn col in dgvDetails.Columns) col.Width = 60;
            dgvDetails.Columns["Item_No"].Width = 40; dgvDetails.Columns["Item_Name"].Width = 150;
            AutoAdjustColumnWidths();
        }

        private void AddLabel(Panel p, string text, int x, int y) => p.Controls.Add(new Label { Text = text, Location = new Point(x, y + 3), Size = new Size(80, 20), Font = new Font("Segoe UI", 9) });
        private void AddLabelCus(Panel p, string text, int x, int y, int w, int h) => p.Controls.Add(new Label { Text = text, Location = new Point(x, y + 3), Size = new Size(w, h), Font = new Font("Segoe UI", 9) });
        private TextBox AddTxt(Panel p, int x, int y, int width)
        {
            var txt = new TextBox { Location = new Point(x, y), Size = new Size(width, 25), Font = new Font("Segoe UI", 9) };
            p.Controls.Add(txt); return txt;
        }
        private Button CreateButton(string text, Color color, Point loc, int w, int h) => new Button { Text = text, Location = loc, Size = new Size(w, h), BackColor = color, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Cursor = Cursors.Hand };
        private void FrmPO_Resize(object sender, EventArgs e)
        {
            int w = this.ClientSize.Width - 20;
            int h = this.ClientSize.Height;
            panelTop.Width = w; panelHeader.Width = w; panelDetail.Width = w;
            panelHeader.Top = panelTop.Bottom + 10;
            panelDetail.Top = panelHeader.Bottom + 10; panelDetail.Height = h - panelDetail.Top - 10;
            dgvPO.Width = panelTop.Width - 20;
            dgvDetails.Width = panelDetail.Width - 20; dgvDetails.Height = panelDetail.Height - 80;
            // txtNotes.Width được cố định 200px — KHÔNG resize theo form

            // Giữ label tổng tiền luôn sát mép phải panelDetail
            if (lblTotal != null && lblSubTotal != null && panelDetail != null)
            {
                lblTotal.Left = panelDetail.Width - lblTotal.Width - 10;
                lblSubTotal.Left = lblTotal.Left - lblSubTotal.Width - 10;
            }
        }

        private void LoadPO()
        {
            try
            {
                _poList = _service.GetAll();
                BindPOGrid(_poList); lblStatus.Text = $"Tổng: {_poList.Count} đơn PO";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =========================================================================
        // BIND PO GRID — sắp xếp A-Z theo PO No
        // =========================================================================
        private void BindPOGrid(List<POHead> list)
        {
            var suppliers = new SupplierService().GetAll();
            var sorted = list.OrderBy(h => h.PONo, StringComparer.OrdinalIgnoreCase).ToList();
            dgvPO.DataSource = sorted.ConvertAll(h =>
            {
                var supplier = suppliers.Find(s => s.Supplier_ID == h.Supplier_ID);
                return new
                {
                    ID = h.PO_ID,
                    PO_No = h.PONo,
                    NCC = supplier?.Short_Name ?? "",
                    Du_An = h.Project_Name,
                    MPR_No = h.MPR_No,
                    Workorder = h.WorkorderNo,
                    Ngay_PO = h.PO_Date.HasValue ? h.PO_Date.Value.ToString("dd/MM/yyyy") : "",
                    Trang_Thai = h.Status,
                    Tong_Tien = h.Total_Amount.ToString("N0"),
                    Revise = h.Revise,
                    Ngay_Tao = h.Created_Date.HasValue ? h.Created_Date.Value.ToString("dd/MM/yyyy") : ""
                };
            });
            if (dgvPO.Columns.Contains("ID")) dgvPO.Columns["ID"].Visible = false;
        }

        private void LoadDetails(int poId)
        {
            try
            {
                _details = new POService().GetDetails(poId);
                dgvDetails.CellValueChanged -= DgvDetails_CellValueChanged;
                dgvDetails.Rows.Clear();
                foreach (var d in _details)
                {
                    int idx = dgvDetails.Rows.Add();
                    var row = dgvDetails.Rows[idx];
                    string remarks = d.Remarks ?? ""; string calcMethod = "Theo KG"; decimal realPrice = d.Price;
                    decimal q = d.Qty_Per_Sheet; decimal wk = d.Weight_kg;
                    if (remarks.Contains("[CALC:KG]"))
                    {
                        calcMethod = "Theo KG"; remarks = remarks.Replace("[CALC:KG]", "").Trim();
                        if (wk > 0 && q > 0) realPrice = Math.Round((d.Price * q) / wk, 2);
                    }
                    else if (remarks.Contains("[CALC:SL]"))
                    {
                        calcMethod = "Theo SL";
                        remarks = remarks.Replace("[CALC:SL]", "").Trim();
                    }
                    else
                    {
                        decimal amtByKG = wk * d.Price * (1 + d.VAT / 100);
                        decimal amtBySL = q * d.Price * (1 + d.VAT / 100);
                        if (Math.Abs(d.Amount - amtByKG) < Math.Abs(d.Amount - amtBySL)) calcMethod = "Theo KG";
                    }
                    row.Cells["PO_Detail_ID"].Value = d.PO_Detail_ID;
                    row.Cells["MPR_Detail_ID"].Value = d.MPR_Detail_ID.HasValue ? (object)d.MPR_Detail_ID.Value : DBNull.Value;
                    row.Cells["Item_No"].Value = d.Item_No; row.Cells["Item_Name"].Value = d.Item_Name; row.Cells["Material"].Value = d.Material;
                    row.Cells["Asize"].Value = d.Asize;
                    row.Cells["Bsize"].Value = d.Bsize; row.Cells["Csize"].Value = d.Csize;
                    row.Cells["Qty"].Value = d.Qty_Per_Sheet; row.Cells["UNIT"].Value = d.UNIT; row.Cells["Weight"].Value = d.Weight_kg;
                    row.Cells["Price"].Value = realPrice;
                    // VAT ComboBox chỉ chấp nhận "8" hoặc "10"
                    string vatStr = ((int)Math.Round(d.VAT)).ToString();
                    row.Cells["VAT"].Value = vatStr == "8" ? "8" : "10";
                    row.Cells["Amount"].Value = d.Amount;
                    row.Cells["Received"].Value = d.Received; row.Cells["MPSNo"].Value = d.MPSNo; row.Cells["DeliveryLocation"].Value = d.DeliveryLocation;
                    row.Cells["Remarks"].Value = remarks;
                    row.Cells["Calc_Method"].Value = calcMethod; row.Cells["Ordered_PO"].Value = "";
                }
                dgvDetails.CellValueChanged += DgvDetails_CellValueChanged;
                UpdateTotal(); AutoAdjustColumnWidths();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải chi tiết PO: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateTotal()
        {
            decimal total = 0;
            decimal subTotal = 0;
            foreach (DataGridViewRow row in dgvDetails.Rows)
            {
                if (row.IsNewRow) continue;
                decimal.TryParse(row.Cells["Qty"].Value?.ToString(), out decimal qty); decimal.TryParse(row.Cells["Weight"].Value?.ToString(), out decimal weight);
                decimal.TryParse(row.Cells["Price"].Value?.ToString(), out decimal price); decimal.TryParse(row.Cells["VAT"].Value?.ToString(), out decimal vat);
                string calcMethod = row.Cells["Calc_Method"].Value?.ToString() ?? "Theo KG"; decimal baseValue = (calcMethod == "Theo KG") ? weight : qty;
                decimal rowSubTotal = Math.Round(baseValue * price, 0); decimal rowTotalAmount = Math.Round(rowSubTotal * (1 + vat / 100), 0);
                subTotal += rowSubTotal; total += rowTotalAmount;
            }
            if (lblSubTotal != null) lblSubTotal.Text = $"Trước VAT: {subTotal:N0} VND";
            if (lblTotal != null) lblTotal.Text = $"Sau VAT: {total:N0} VND";
        }

        private void DgvDetails_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V)
            {
                PasteFromExcel();
                e.Handled = true;
            }
        }

        private void PasteFromExcel()
        {
            try
            {
                string copiedData = Clipboard.GetText();
                if (string.IsNullOrEmpty(copiedData)) return;
                string[] lines = copiedData.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries); if (lines.Length == 0) return;
                int startRow = dgvDetails.CurrentCell?.RowIndex ?? 0; int startCol = dgvDetails.CurrentCell?.ColumnIndex ?? 0;
                foreach (string line in lines)
                {
                    string[] cells = line.Split('\t');
                    if (startRow >= dgvDetails.Rows.Count)
                    {
                        int nextItem = dgvDetails.Rows.Count + 1;
                        int newIdx = dgvDetails.Rows.Add(); var r = dgvDetails.Rows[newIdx];
                        r.Cells["Item_No"].Value = nextItem; r.Cells["PO_Detail_ID"].Value = 0; r.Cells["UNIT"].Value = "PCS";
                        r.Cells["Qty"].Value = 0;
                        r.Cells["Weight"].Value = 0; r.Cells["Price"].Value = 0; r.Cells["VAT"].Value = "10";
                        r.Cells["Amount"].Value = 0; r.Cells["Received"].Value = 0; r.Cells["Calc_Method"].Value = "Theo KG";
                        r.Cells["Ordered_PO"].Value = "";
                    }
                    int colIndex = startCol;
                    for (int i = 0; i < cells.Length; i++)
                    {
                        while (colIndex < dgvDetails.Columns.Count && (!dgvDetails.Columns[colIndex].Visible || dgvDetails.Columns[colIndex].ReadOnly)) colIndex++;
                        if (colIndex >= dgvDetails.Columns.Count) break;
                        if (!(dgvDetails.Columns[colIndex] is DataGridViewComboBoxColumn)) dgvDetails.Rows[startRow].Cells[colIndex].Value = cells[i].Trim();
                        else i--;
                        colIndex++;
                    }
                    RecalculateAmount(startRow);
                    startRow++;
                }
                AutoAdjustColumnWidths();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi dán dữ liệu: " + ex.Message, "Lỗi Paste", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RecalculateAmount(int rowIndex)
        {
            var row = dgvDetails.Rows[rowIndex];
            decimal.TryParse(row.Cells["Qty"].Value?.ToString(), out decimal qty);
            decimal.TryParse(row.Cells["Weight"].Value?.ToString(), out decimal weight);
            // Price có thể đang ở dạng "1,000,000" — cần bỏ dấu phẩy trước parse
            decimal.TryParse((row.Cells["Price"].Value?.ToString() ?? "0").Replace(",", ""), out decimal price);
            decimal.TryParse(row.Cells["VAT"].Value?.ToString(), out decimal vat);
            string calcMethod = row.Cells["Calc_Method"].Value?.ToString() ?? "Theo KG";
            decimal baseValue = (calcMethod == "Theo KG") ? weight : qty;
            row.Cells["Amount"].Value = Math.Round(baseValue * price * (1 + vat / 100), 0);
            UpdateTotal();
        }

        private void DgvDetails_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            RecalculateAmount(e.RowIndex); AutoAdjustColumnWidths();
        }

        private void BtnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtSearch.Text)) LoadPO();
                else
                {
                    var result = _service.Search(txtSearch.Text.Trim()); BindPOGrid(result); lblStatus.Text = $"Tìm thấy: {result.Count} đơn PO";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tìm kiếm: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DgvPO_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvPO.SelectedRows.Count == 0) return;
            var row = dgvPO.SelectedRows[0]; _selectedPO_ID = Convert.ToInt32(row.Cells["ID"].Value);
            var h = _poList.Find(x => x.PO_ID == _selectedPO_ID); if (h == null) return;
            txtPONo.Text = h.PONo; txtProjectName.Text = h.Project_Name; txtWorkorderNo.Text = h.WorkorderNo; txtMPRNo.Text = h.MPR_No;
            txtPrepared.Text = h.Prepared; txtReviewed.Text = h.Reviewed;
            txtAgreement.Text = h.Agreement; txtApproved.Text = h.Approved;
            txtNotes.Text = h.Notes; nudRevise.Value = h.Revise;
            if (h.PO_Date.HasValue) dtpPODate.Value = h.PO_Date.Value;
            var idx = cboStatus.Items.IndexOf(h.Status); cboStatus.SelectedIndex = idx >= 0 ? idx : 0;
            LoadDetails(_selectedPO_ID);

            // GỌI HÀM LOAD FILES KHI CHỌN PO
            LoadFiles(h.WorkorderNo, h.Project_Name);

            // Load delivery tracking
            LoadDeliveries();
        }

        private void BtnNewPO_Click(object sender, EventArgs e)
        {
            ClearHeader(); _selectedPO_ID = 0; dgvDetails.Rows.Clear(); dgvFiles.Rows.Clear();
            dgvDelivery.Rows.Clear();
            UpdateTotal(); txtPONo.Focus(); lblStatus.Text = "Đang tạo đơn PO mới...";
        }

        // =========================================================================
        // LƯU TOÀN BỘ PO (Header + Detail)
        // =========================================================================
        private void BtnSavePO_Click(object sender, EventArgs e)
        {
            dgvDetails.EndEdit();
            if (string.IsNullOrWhiteSpace(txtPONo.Text))
            {
                MessageBox.Show("Vui lòng nhập PO No!", "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Warning); txtPONo.Focus(); return;
            }
            if (dgvDetails.Rows.Count == 0 && MessageBox.Show("Đơn hàng này chưa có chi tiết vật tư nào.\nBạn có chắc chắn muốn lưu chỉ với Header không?", "Cảnh báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No) return;
            try
            {
                string basePONo = txtPONo.Text.Trim();
                int revIdx = basePONo.LastIndexOf("_Rev"); if (revIdx > 0) basePONo = basePONo.Substring(0, revIdx);
                string finalPONo = basePONo;
                bool isBaseDuplicate = _poList.Exists(p => p.PONo == basePONo && p.PO_ID != _selectedPO_ID);
                if (isBaseDuplicate || nudRevise.Value > 0)
                {
                    if (nudRevise.Value == 0)
                    {
                        MessageBox.Show("Số PO này đã tồn tại!\nVui lòng tăng số Revise để tạo bản sửa đổi.", "Trùng lặp", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        nudRevise.Focus(); return;
                    }
                    finalPONo = $"{basePONo}_Rev{nudRevise.Value}";
                    if (_poList.Exists(p => p.PONo == finalPONo && p.PO_ID != _selectedPO_ID))
                    {
                        MessageBox.Show($"Bản '{finalPONo}' cũng đã tồn tại!\nVui lòng tăng Revise lên mức cao hơn.", "Trùng lặp", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        nudRevise.Focus(); return;
                    }
                }
                var h = new POHead
                {
                    PO_ID = _selectedPO_ID,
                    PONo = finalPONo,
                    Project_Name = txtProjectName.Text.Trim(),
                    WorkorderNo = txtWorkorderNo.Text.Trim(),
                    MPR_No = txtMPRNo.Text.Trim(),
                    Prepared = txtPrepared.Text.Trim(),
                    Reviewed = txtReviewed.Text.Trim(),
                    Agreement = txtAgreement.Text.Trim(),
                    Approved = txtApproved.Text.Trim(),
                    Notes = txtNotes.Text.Trim(),
                    PO_Date = dtpPODate.Value,
                    Status = cboStatus.SelectedItem?.ToString() ?? "Draft",
                    Revise = (int)nudRevise.Value,
                    Supplier_ID = Convert.ToInt32(cboSupplier.SelectedValue ?? 0),
                    ProjectCode = _projectCodeImport
                };
                if (_selectedPO_ID == 0) _selectedPO_ID = _service.InsertHead(h, _currentUser);
                else _service.UpdateHead(h, _currentUser);
                txtPONo.Text = finalPONo;
                SaveDetailsToDb();
                MessageBox.Show($"Đã lưu toàn bộ PO thành công!\n- Số PO: {finalPONo}\n- Số dòng vật tư: {dgvDetails.Rows.Count}", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);

                int savedId = _selectedPO_ID;
                LoadPO();
                foreach (DataGridViewRow row in dgvPO.Rows)
                {
                    if (Convert.ToInt32(row.Cells["ID"].Value ?? 0) == savedId)
                    {
                        dgvPO.ClearSelection();
                        row.Selected = true;
                        if (row.Index >= 0) dgvPO.FirstDisplayedScrollingRowIndex = row.Index;
                        break;
                    }
                }
                if (_selectedPO_ID == savedId) LoadDetails(savedId);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi lưu PO: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =========================================================================
        // LƯU CHI TIẾT — chỉ lưu detail, không đụng Header
        // =========================================================================
        private void BtnSaveDetail_Click(object sender, EventArgs e)
        {
            dgvDetails.EndEdit();
            if (_selectedPO_ID == 0)
            {
                MessageBox.Show("Vui lòng chọn hoặc lưu Header PO trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }
            if (dgvDetails.Rows.Count == 0)
            {
                MessageBox.Show("Không có dòng nào để lưu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                SaveDetailsToDb();
                MessageBox.Show($"✅ Đã lưu {dgvDetails.Rows.Count} dòng chi tiết thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDetails(_selectedPO_ID);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi lưu chi tiết: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =========================================================================
        // HÀM CHUNG lưu detail — dùng bởi cả BtnSavePO và BtnSaveDetail
        // =========================================================================
        private void SaveDetailsToDb()
        {
            var oldDetails = _service.GetDetails(_selectedPO_ID);
            foreach (var d in oldDetails) _service.DeleteDetail(d.PO_Detail_ID);
            int itemNo = 1;
            foreach (DataGridViewRow row in dgvDetails.Rows)
            {
                if (row.IsNewRow) continue;
                decimal q = decimal.TryParse(row.Cells["Qty"].Value?.ToString(), out decimal _q) ? _q : 0;
                decimal wk = decimal.TryParse(row.Cells["Weight"].Value?.ToString(), out decimal _wk) ?
                _wk : 0;
                decimal p = decimal.TryParse(row.Cells["Price"].Value?.ToString(), out decimal _p) ? _p : 0;
                decimal vat = decimal.TryParse(row.Cells["VAT"].Value?.ToString(), out decimal _vat) ? _vat : 0;
                string calcMethod = row.Cells["Calc_Method"].Value?.ToString() ?? "Theo KG";
                string remarks = row.Cells["Remarks"].Value?.ToString() ?? "";
                remarks = remarks.Replace("[CALC:KG]", "").Replace("[CALC:SL]", "").Trim();
                decimal dbPrice = p;
                if (calcMethod == "Theo KG")
                {
                    remarks += " [CALC:KG]";
                    if (q > 0 && wk > 0) dbPrice = (wk * p) / q;
                }
                else remarks += " [CALC:SL]";
                int? mprDetailId = null;
                if (dgvDetails.Columns.Contains("MPR_Detail_ID") && row.Cells["MPR_Detail_ID"].Value != null)
                    if (int.TryParse(row.Cells["MPR_Detail_ID"].Value.ToString(), out int mdi) && mdi > 0) mprDetailId = mdi;
                var detail = new PODetail
                {
                    Item_No = itemNo++,
                    Item_Name = row.Cells["Item_Name"].Value?.ToString() ??
                    "",
                    Material = row.Cells["Material"].Value?.ToString() ??
                    "",
                    Asize = row.Cells["Asize"].Value?.ToString() ??
                    "",
                    Bsize = row.Cells["Bsize"].Value?.ToString() ??
                    "",
                    Csize = row.Cells["Csize"].Value?.ToString() ??
                    "",
                    Qty_Per_Sheet = (int)q,
                    UNIT = row.Cells["UNIT"].Value?.ToString() ??
                    "",
                    Weight_kg = wk,
                    Price = dbPrice,
                    VAT = vat,
                    Amount = 0,
                    Received = int.TryParse(row.Cells["Received"].Value?.ToString(), out int rec) ?
                    rec : 0,
                    MPSNo = row.Cells["MPSNo"].Value?.ToString() ??
                    "",
                    DeliveryLocation = row.Cells["DeliveryLocation"].Value?.ToString() ??
                    "",
                    Remarks = remarks.Trim(),
                    MPR_Detail_ID = mprDetailId
                };
                _service.InsertDetail(detail, _selectedPO_ID);
            }
        }

        private void BtnDeletePO_Click(object sender, EventArgs e)
        {
            if (dgvPO.SelectedRows.Count == 0 || _selectedPO_ID == 0)
            {
                MessageBox.Show("Vui lòng chọn một đơn PO trong 'Danh sách đơn đặt hàng' để xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string poNo = dgvPO.SelectedRows[0].Cells["PO_No"].Value?.ToString() ?? "";
            if (MessageBox.Show($"Bạn có chắc chắn muốn xóa đơn PO '{poNo}' và toàn bộ chi tiết vật tư bên trong không?\nHành động này không thể hoàn tác!", "Xác nhận xóa PO", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            {
                try
                {
                    _service.DeletePO(_selectedPO_ID);
                    MessageBox.Show($"Đã xóa thành công đơn PO '{poNo}'!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information); ClearHeader(); dgvDetails.Rows.Clear(); dgvFiles.Rows.Clear(); UpdateTotal(); _selectedPO_ID = 0; LoadPO();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi xóa PO: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnAddDetail_Click(object sender, EventArgs e)
        {
            int nextItem = dgvDetails.Rows.Count + 1;
            int newIdx = dgvDetails.Rows.Add(); var r = dgvDetails.Rows[newIdx];
            r.Cells["DeliveryLocation"].Value = ""; r.Cells["Item_No"].Value = nextItem; r.Cells["Item_Name"].Value = ""; r.Cells["Material"].Value = "";
            r.Cells["Asize"].Value = ""; r.Cells["Bsize"].Value = ""; r.Cells["Csize"].Value = "";
            r.Cells["Qty"].Value = 0; r.Cells["UNIT"].Value = "PCS"; r.Cells["Weight"].Value = 0;
            r.Cells["Price"].Value = 0; r.Cells["VAT"].Value = "10";
            r.Cells["Amount"].Value = 0; r.Cells["Received"].Value = 0; r.Cells["MPSNo"].Value = ""; r.Cells["Remarks"].Value = "";
            r.Cells["Calc_Method"].Value = "Theo KG"; r.Cells["Ordered_PO"].Value = ""; r.Cells["PO_Detail_ID"].Value = 0;
            dgvDetails.CurrentCell = r.Cells["Item_Name"]; AutoAdjustColumnWidths();
        }

        private void BtnDeleteDetail_Click(object sender, EventArgs e)
        {
            if (dgvDetails.SelectedRows.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn ít nhất một dòng trong danh sách vật tư để xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string msg = dgvDetails.SelectedRows.Count == 1 ?
            "Bạn có chắc chắn muốn xóa dòng này?" : $"Bạn có chắc chắn muốn xóa {dgvDetails.SelectedRows.Count} dòng đã chọn?";
            if (MessageBox.Show(msg, "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    var rowsToDelete = new List<DataGridViewRow>();
                    foreach (DataGridViewRow row in dgvDetails.SelectedRows) if (!row.IsNewRow) rowsToDelete.Add(row);
                    foreach (var row in rowsToDelete) dgvDetails.Rows.Remove(row);
                    int itemNo = 1;
                    foreach (DataGridViewRow row in dgvDetails.Rows) if (!row.IsNewRow) row.Cells["Item_No"].Value = itemNo++;
                    UpdateTotal(); AutoAdjustColumnWidths();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi xóa: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private Dictionary<int, string> GetPoMappingForMpr(int mprId)
        {
            var dict = new Dictionary<int, string>();
            if (mprId <= 0) return dict;
            try
            {
                using (var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    string sql = @"SELECT pod.MPR_Detail_ID, poh.PONo FROM PO_Detail pod INNER JOIN PO_head poh ON pod.PO_ID = poh.PO_ID WHERE pod.MPR_Detail_ID IS NOT NULL AND pod.MPR_Detail_ID IN (SELECT Detail_ID FROM MPR_Details WHERE MPR_ID = @mprId)";
                    using (var cmd = new Microsoft.Data.SqlClient.SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@mprId", mprId);
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                if (reader["MPR_Detail_ID"] != DBNull.Value)
                                {
                                    int detailId = Convert.ToInt32(reader["MPR_Detail_ID"]);
                                    string poNo = reader["PONo"]?.ToString() ?? ""; if (dict.ContainsKey(detailId))
                                    {
                                        if (!dict[detailId].Contains(poNo)) dict[detailId] += ", " + poNo;
                                    }
                                    else dict[detailId] = poNo;
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

        private void BtnImportMPR_Click(object sender, EventArgs e)
        {
            using (var dlg = new frmSelectMPR())
            {
                if (dlg.ShowDialog() == DialogResult.OK && dlg.SelectedMPR != null)
                {
                    ClearHeader(); _selectedPO_ID = 0; _details.Clear(); dgvDetails.Rows.Clear(); dgvFiles.Rows.Clear();
                    var mpr = dlg.SelectedMPR;
                    var details = dlg.SelectedDetails; var poMapping = GetPoMappingForMpr(mpr.MPR_ID);
                    txtProjectName.Text = mpr.Project_Name; txtMPRNo.Text = mpr.MPR_No;
                    try
                    {
                        var projects = new ProjectService().GetAll();
                        var project = projects.Find(p => !string.IsNullOrEmpty(p.ProjectName) && p.ProjectName.Equals(mpr.Project_Name, StringComparison.OrdinalIgnoreCase));
                        if (project == null) project = projects.Find(p => !string.IsNullOrEmpty(p.ProjectName) && (p.ProjectName.Contains(mpr.Project_Name, StringComparison.OrdinalIgnoreCase) || mpr.Project_Name.Contains(p.ProjectName, StringComparison.OrdinalIgnoreCase)));
                        if (project != null)
                        {
                            txtWorkorderNo.Text = project.WorkorderNo ?? ""; _projectCodeImport = project.ProjectCode; txtPONo.Text = GenerateAutoPoNo(project.POCode ?? project.ProjectCode ?? "");
                        }
                        else MessageBox.Show($"Không tìm thấy dự án khớp với tên \"{mpr.Project_Name}\".\nVui lòng kiểm tra lại thông tin Workorder và PO No.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("Lỗi tìm project: " + ex.Message);
                    }
                    int itemNo = 1;
                    foreach (var d in details)
                    {
                        string orderedPo = poMapping.ContainsKey(d.Detail_ID) ?
                        poMapping[d.Detail_ID] : "";
                        string aSize = d.Thickness_mm > 0 ? d.Thickness_mm.ToString() : (d.Depth_mm > 0 ? d.Depth_mm.ToString() : "");
                        string bSize = d.C_Width_mm > 0 ? d.C_Width_mm.ToString() : "";
                        string cSize = (d.D_Web_mm == 0 && d.E_Flange_mm == 0) ? (d.F_Length_mm > 0 ? d.F_Length_mm.ToString() : "") : $"{d.D_Web_mm}x{d.E_Flange_mm}x{d.F_Length_mm}";
                        int idx = dgvDetails.Rows.Add(); var r = dgvDetails.Rows[idx];
                        r.Cells["DeliveryLocation"].Value = d.Usage_Location; r.Cells["Item_No"].Value = itemNo++; r.Cells["Item_Name"].Value = d.Item_Name; r.Cells["Material"].Value = d.Material;
                        r.Cells["Asize"].Value = aSize; r.Cells["Bsize"].Value = bSize; r.Cells["Csize"].Value = cSize; r.Cells["Qty"].Value = d.Qty_Per_Sheet; r.Cells["UNIT"].Value = d.UNIT; r.Cells["Weight"].Value = d.Weight_kg;
                        r.Cells["Price"].Value = 0; r.Cells["VAT"].Value = "10"; r.Cells["Amount"].Value = 0; r.Cells["Received"].Value = 0; r.Cells["MPSNo"].Value = d.MPS_Info; r.Cells["Remarks"].Value = d.Remarks;
                        r.Cells["Calc_Method"].Value = "Theo KG"; r.Cells["Ordered_PO"].Value = orderedPo; r.Cells["PO_Detail_ID"].Value = 0; r.Cells["MPR_Detail_ID"].Value = d.Detail_ID;
                    }
                    UpdateTotal();
                    AutoAdjustColumnWidths();
                    MessageBox.Show($"✅ Đã import {details.Count} dòng từ MPR {mpr.MPR_No}!\nPO No dự kiến: {txtPONo.Text}\nWorkorder: {txtWorkorderNo.Text}", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        public void ImportMPRByNo(string mprNo)
        {
            if (string.IsNullOrEmpty(mprNo)) return;
            try
            {
                var mprService = new MPR_Managerment.Services.MPRService();
                var mpr = mprService.GetAll().Find(m => m.MPR_No == mprNo);
                if (mpr == null)
                {
                    MessageBox.Show($"Không tìm thấy MPR: {mprNo}", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                var details = mprService.GetDetails(mpr.MPR_ID);
                if (details == null || details.Count == 0)
                {
                    MessageBox.Show($"MPR {mprNo} chưa có chi tiết vật tư!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                ClearHeader();
                _selectedPO_ID = 0; _details.Clear(); dgvDetails.Rows.Clear(); dgvFiles.Rows.Clear();
                var poMapping = GetPoMappingForMpr(mpr.MPR_ID); txtProjectName.Text = mpr.Project_Name; txtMPRNo.Text = mpr.MPR_No;
                try
                {
                    var projects = new MPR_Managerment.Services.ProjectService().GetAll();
                    var project = projects.Find(p => !string.IsNullOrEmpty(p.ProjectName) && p.ProjectName.Equals(mpr.Project_Name, StringComparison.OrdinalIgnoreCase));
                    if (project == null) project = projects.Find(p => !string.IsNullOrEmpty(p.ProjectName) && (p.ProjectName.Contains(mpr.Project_Name, StringComparison.OrdinalIgnoreCase) || mpr.Project_Name.Contains(p.ProjectName, StringComparison.OrdinalIgnoreCase)));
                    if (project != null)
                    {
                        txtWorkorderNo.Text = project.WorkorderNo ?? ""; _projectCodeImport = project.ProjectCode; txtPONo.Text = GenerateAutoPoNo(project.POCode ?? project.ProjectCode ?? "");
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Lỗi tìm project: " + ex.Message);
                }
                int itemNo = 1;
                foreach (var d in details)
                {
                    string orderedPo = poMapping.ContainsKey(d.Detail_ID) ?
                    poMapping[d.Detail_ID] : "";
                    string aSize = d.Thickness_mm > 0 ? d.Thickness_mm.ToString() : (d.Depth_mm > 0 ? d.Depth_mm.ToString() : "");
                    string bSize = d.C_Width_mm > 0 ? d.C_Width_mm.ToString() : "";
                    string cSize = (d.D_Web_mm == 0 && d.E_Flange_mm == 0) ? (d.F_Length_mm > 0 ? d.F_Length_mm.ToString() : "") : $"{d.D_Web_mm}x{d.E_Flange_mm}x{d.F_Length_mm}";
                    int idx = dgvDetails.Rows.Add(); var r = dgvDetails.Rows[idx];
                    r.Cells["DeliveryLocation"].Value = d.Usage_Location; r.Cells["Item_No"].Value = itemNo++; r.Cells["Item_Name"].Value = d.Item_Name; r.Cells["Material"].Value = d.Material;
                    r.Cells["Asize"].Value = aSize; r.Cells["Bsize"].Value = bSize; r.Cells["Csize"].Value = cSize; r.Cells["Qty"].Value = d.Qty_Per_Sheet; r.Cells["UNIT"].Value = d.UNIT; r.Cells["Weight"].Value = d.Weight_kg;
                    r.Cells["Price"].Value = 0; r.Cells["VAT"].Value = "10"; r.Cells["Amount"].Value = 0; r.Cells["Received"].Value = 0; r.Cells["MPSNo"].Value = d.MPS_Info; r.Cells["Remarks"].Value = d.Remarks;
                    r.Cells["Calc_Method"].Value = "Theo KG"; r.Cells["Ordered_PO"].Value = orderedPo; r.Cells["PO_Detail_ID"].Value = 0; r.Cells["MPR_Detail_ID"].Value = d.Detail_ID;
                }
                UpdateTotal(); AutoAdjustColumnWidths();
                MessageBox.Show($"✅ Đã import {details.Count} dòng từ MPR {mpr.MPR_No}!\nPO No: {txtPONo.Text}\nWorkorder: {txtWorkorderNo.Text}", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi import MPR: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GenerateAutoPoNo(string poCode)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(poCode)) poCode = "PRJ";
                string prefix = $"DV-{poCode}-PC-";

                using (var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection())
                {
                    conn.Open();

                    // Lấy tất cả PO No có prefix này, rồi tìm số thứ tự lớn nhất
                    var cmd = new Microsoft.Data.SqlClient.SqlCommand(
                        "SELECT PONo FROM PO_head WHERE PONo LIKE @prefix", conn);
                    cmd.Parameters.AddWithValue("@prefix", prefix + "%");

                    int maxNo = 0;
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string poNo = reader["PONo"]?.ToString() ?? "";
                            // Lấy phần số sau prefix, bỏ qua _RevX nếu có
                            string suffix = poNo.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                                ? poNo.Substring(prefix.Length) : "";
                            // Bỏ phần _Rev nếu có
                            int revIdx = suffix.IndexOf("_Rev", StringComparison.OrdinalIgnoreCase);
                            if (revIdx > 0) suffix = suffix.Substring(0, revIdx);
                            if (int.TryParse(suffix, out int num))
                                if (num > maxNo) maxNo = num;
                        }
                    }

                    // Kiểm tra thêm trong _poList (dữ liệu chưa lưu DB)
                    foreach (var p in _poList)
                    {
                        string poNo = p.PONo ?? "";
                        if (!poNo.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) continue;
                        string suffix = poNo.Substring(prefix.Length);
                        int revIdx = suffix.IndexOf("_Rev", StringComparison.OrdinalIgnoreCase);
                        if (revIdx > 0) suffix = suffix.Substring(0, revIdx);
                        if (int.TryParse(suffix, out int num))
                            if (num > maxNo) maxNo = num;
                    }

                    return $"{prefix}{maxNo + 1:D3}";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("GenerateAutoPoNo error: " + ex.Message);
                return $"DV-{poCode}-PC-{DateTime.Now:ddMMHH}";
            }
        }

        // =========================================================================
        // DELIVERY TRACKING — Load, Add popup, Done, Delete, Auto-clean
        // =========================================================================

        private void LoadDeliveries()
        {
            dgvDelivery.Rows.Clear();
            try
            {
                UpdateOverdueDeliveries();

                string sql = @"
                    SELECT dt.TrackID, dt.PONo, ISNULL(pi.ProjectCode,'') AS MaDuAn,
                           CONVERT(NVARCHAR(10), dt.ExpDelivery, 103) AS ExpDelivery,
                           ISNULL(dt.GhiChu,'') AS GhiChu,
                           ISNULL(dt.Status,'Pending') AS Status,
                           ISNULL(dt.ReceiverNote,'') AS ReceiverNote
                    FROM PO_DeliveryTracking dt
                    LEFT JOIN PO_head ph ON ph.PONo = dt.PONo
                    LEFT JOIN ProjectInfo pi ON pi.WorkorderNo = ph.WorkorderNo
                    WHERE ISNULL(dt.Status,'Pending') != 'Done'
                    ORDER BY dt.ExpDelivery ASC";

                using (var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var dt = new System.Data.DataTable();
                    dt.Load(new Microsoft.Data.SqlClient.SqlCommand(sql, conn).ExecuteReader());
                    foreach (System.Data.DataRow r in dt.Rows)
                    {
                        dgvDelivery.Rows.Add(
                            r["TrackID"], r["PONo"], r["MaDuAn"],
                            r["ExpDelivery"], r["GhiChu"],
                            r["Status"], r["ReceiverNote"]);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("LoadDeliveries: " + ex.Message);
            }
        }

        private void UpdateOverdueDeliveries()
        {
            try
            {
                string sql = @"
                    UPDATE PO_DeliveryTracking
                    SET Status = 'Overdue'
                    WHERE ExpDelivery < CAST(GETDATE() AS DATE)
                      AND ISNULL(Status,'Pending') NOT IN ('Done','Overdue')";
                using (var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    new Microsoft.Data.SqlClient.SqlCommand(sql, conn).ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("UpdateOverdueDeliveries: " + ex.Message);
            }
        }

        private void MarkDeliveryDone()
        {
            if (dgvDelivery.SelectedRows.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn một dòng!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            int trackId = Convert.ToInt32(dgvDelivery.SelectedRows[0].Cells["TrackID"].Value ?? 0);
            if (trackId == 0) return;

            // Lưu ReceiverNote trước khi Done
            dgvDelivery.EndEdit();
            string receiverNote = dgvDelivery.SelectedRows[0].Cells["ReceiverNote"].Value?.ToString() ?? "";

            try
            {
                string sql = "UPDATE PO_DeliveryTracking SET Status = 'Done', ReceiverNote = @note WHERE TrackID = @id";
                using (var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var cmd = new Microsoft.Data.SqlClient.SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@note", receiverNote);
                    cmd.Parameters.AddWithValue("@id", trackId);
                    cmd.ExecuteNonQuery();
                }
                LoadDeliveries();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DeleteDeliveryRow()
        {
            if (dgvDelivery.SelectedRows.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn một dòng!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show("Xóa dòng theo dõi này?", "Xác nhận",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

            int trackId = Convert.ToInt32(dgvDelivery.SelectedRows[0].Cells["TrackID"].Value ?? 0);
            if (trackId > 0)
            {
                try
                {
                    string sql = "DELETE FROM PO_DeliveryTracking WHERE TrackID = @id";
                    using (var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection())
                    {
                        conn.Open();
                        var cmd = new Microsoft.Data.SqlClient.SqlCommand(sql, conn);
                        cmd.Parameters.AddWithValue("@id", trackId);
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            LoadDeliveries();
        }

        private void ShowDeliveryDetailPopup(string poNo)
        {
            try
            {
                var po = _poList.Find(p => p.PONo == poNo);
                if (po == null) { MessageBox.Show($"Không tìm thấy PO: {poNo}", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

                var details = _service.GetDetails(po.PO_ID);
                string suppName = "";
                try { var s = new SupplierService().GetAll().Find(x => x.Supplier_ID == po.Supplier_ID); suppName = s?.Company_Name ?? s?.Short_Name ?? ""; } catch { }

                var popup = new Form
                {
                    Text = $"📦  Chi tiết vật tư  —  {poNo}  |  MPR: {po.MPR_No ?? "—"}",
                    Size = new Size(1000, 540),
                    StartPosition = FormStartPosition.CenterParent,
                    BackColor = Color.FromArgb(245, 245, 245),
                    MinimumSize = new Size(750, 380),
                    KeyPreview = true
                };
                popup.Controls.Add(new Label { Text = $"📦  {poNo}", Font = new Font("Segoe UI", 12, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212), Location = new Point(10, 8), Size = new Size(600, 28) });
                popup.Controls.Add(new Label
                {
                    Text = $"MPR No: {po.MPR_No ?? "—"}    |    Dự án: {po.Project_Name ?? "—"}    |    NCC: {suppName}    |    Ngày PO: {(po.PO_Date.HasValue ? po.PO_Date.Value.ToString("dd/MM/yyyy") : "—")}    |    Tổng: {po.Total_Amount:N0} VNĐ",
                    Font = new Font("Segoe UI", 9),
                    ForeColor = Color.FromArgb(80, 80, 80),
                    Location = new Point(10, 38),
                    Size = new Size(960, 20),
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
                });
                popup.Controls.Add(new Label { Text = $"Tổng: {details.Count} dòng vật tư", Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.FromArgb(0, 150, 100), Location = new Point(10, 62), Size = new Size(300, 20) });

                var dgv = new DataGridView
                {
                    Location = new Point(10, 88),
                    Size = new Size(popup.ClientSize.Width - 20, popup.ClientSize.Height - 140),
                    ReadOnly = true,
                    AllowUserToAddRows = false,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                    BackgroundColor = Color.White,
                    BorderStyle = BorderStyle.FixedSingle,
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
                dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "STT", HeaderText = "STT", Width = 42 });
                dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "TenHang", HeaderText = "Tên hàng", Width = 200 });
                dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "VatLieu", HeaderText = "Vật liệu", Width = 120 });
                dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Amm", HeaderText = "A(mm)", Width = 70 });
                dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Bmm", HeaderText = "B(mm)", Width = 70 });
                dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Cmm", HeaderText = "C(mm)", Width = 70 });
                dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "SoLuong", HeaderText = "Số lượng", Width = 80 });
                dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "DVT", HeaderText = "ĐVT", Width = 60 });
                dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "MPRNo", HeaderText = "MPR No", Width = 150 });
                popup.Controls.Add(dgv);

                for (int i = 0; i < details.Count; i++)
                {
                    var d = details[i];
                    int idx = dgv.Rows.Add();
                    dgv.Rows[idx].Cells["STT"].Value = i + 1;
                    dgv.Rows[idx].Cells["TenHang"].Value = d.Item_Name ?? "";
                    dgv.Rows[idx].Cells["VatLieu"].Value = d.Material ?? "";
                    dgv.Rows[idx].Cells["Amm"].Value = d.Asize;
                    dgv.Rows[idx].Cells["Bmm"].Value = d.Bsize;
                    dgv.Rows[idx].Cells["Cmm"].Value = d.Csize;
                    dgv.Rows[idx].Cells["SoLuong"].Value = d.Qty_Per_Sheet;
                    dgv.Rows[idx].Cells["DVT"].Value = d.UNIT ?? "";
                    dgv.Rows[idx].Cells["MPRNo"].Value = po.MPR_No ?? "";
                }
                dgv.CellFormatting += (s, ev) =>
                {
                    if (ev.RowIndex < 0) return;
                    string col = dgv.Columns[ev.ColumnIndex].Name;
                    if (col == "Amm" || col == "Bmm" || col == "Cmm" || col == "SoLuong")
                        ev.CellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    if (col == "STT") ev.CellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                };

                var btnClose = new Button { Text = "Đóng", Location = new Point(popup.ClientSize.Width - 110, popup.ClientSize.Height - 42), Size = new Size(100, 32), BackColor = Color.FromArgb(108, 117, 125), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Anchor = AnchorStyles.Bottom | AnchorStyles.Right, DialogResult = DialogResult.Cancel };
                btnClose.FlatAppearance.BorderSize = 0;
                popup.Controls.Add(btnClose);
                popup.CancelButton = btnClose;
                popup.Resize += (s, ev) =>
                {
                    dgv.Size = new Size(popup.ClientSize.Width - 20, popup.ClientSize.Height - 140);
                    btnClose.Location = new Point(popup.ClientSize.Width - 110, popup.ClientSize.Height - 42);
                };
                popup.ShowDialog(this);
            }
            catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void ShowDeliveryAddPopup()
        {
            try
            {
                // ── Load danh sách PO chưa Complete ──
                const string SQL_PO = @"
                    SELECT ph.PONo, ISNULL(pi.ProjectCode,'') AS MaDuAn,
                           ph.MPR_No, ph.Status
                    FROM PO_head ph
                    LEFT JOIN ProjectInfo pi ON pi.WorkorderNo = ph.WorkorderNo
                    WHERE ph.Status NOT IN ('Completed','Cancelled')
                    ORDER BY ph.PONo";

                System.Data.DataTable dtPO;
                using (var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    dtPO = new System.Data.DataTable();
                    dtPO.Load(new Microsoft.Data.SqlClient.SqlCommand(SQL_PO, conn).ExecuteReader());
                }

                // ── Popup ──
                var dlg = new Form
                {
                    Text = "➕ Thêm theo dõi giao hàng",
                    Size = new Size(900, 560),
                    StartPosition = FormStartPosition.CenterParent,
                    BackColor = Color.FromArgb(245, 245, 245),
                    MinimumSize = new Size(750, 480),
                    KeyPreview = true
                };

                // ── Tiêu đề ──
                dlg.Controls.Add(new Label
                {
                    Text = "Chọn PO cần theo dõi giao hàng",
                    Font = new Font("Segoe UI", 10, FontStyle.Bold),
                    ForeColor = Color.FromArgb(0, 150, 100),
                    Location = new Point(10, 8),
                    Size = new Size(500, 24)
                });

                // ── PANEL BỘ LỌC ──
                var pFilter = new Panel
                {
                    Location = new Point(10, 36),
                    Size = new Size(860, 38),
                    BackColor = Color.White,
                    BorderStyle = BorderStyle.FixedSingle,
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
                };
                dlg.Controls.Add(pFilter);

                pFilter.Controls.Add(new Label { Text = "Mã DA:", Location = new Point(6, 10), Size = new Size(45, 18), Font = new Font("Segoe UI", 8, FontStyle.Bold) });
                var cboDaFilter = new ComboBox
                {
                    Location = new Point(52, 7),
                    Size = new Size(130, 24),
                    Font = new Font("Segoe UI", 9),
                    DropDownStyle = ComboBoxStyle.DropDownList
                };
                cboDaFilter.Items.Add("Tất cả");
                dtPO.AsEnumerable().Select(r => r["MaDuAn"].ToString())
                    .Where(v => !string.IsNullOrEmpty(v)).Distinct().OrderBy(v => v)
                    .ToList().ForEach(v => cboDaFilter.Items.Add(v));
                cboDaFilter.SelectedIndex = 0;
                pFilter.Controls.Add(cboDaFilter);

                pFilter.Controls.Add(new Label { Text = "MPR No:", Location = new Point(196, 10), Size = new Size(52, 18), Font = new Font("Segoe UI", 8, FontStyle.Bold) });
                var txtMprFilter = new TextBox
                {
                    Location = new Point(250, 7),
                    Size = new Size(120, 24),
                    Font = new Font("Segoe UI", 9),
                    PlaceholderText = "MPR No..."
                };
                pFilter.Controls.Add(txtMprFilter);

                var btnDlgFilter = new Button
                {
                    Text = "🔍 Lọc",
                    Location = new Point(382, 6),
                    Size = new Size(70, 26),
                    BackColor = Color.FromArgb(0, 120, 212),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 8, FontStyle.Bold)
                };
                btnDlgFilter.FlatAppearance.BorderSize = 0;
                pFilter.Controls.Add(btnDlgFilter);

                var btnDlgClear = new Button
                {
                    Text = "✖",
                    Location = new Point(458, 6),
                    Size = new Size(32, 26),
                    BackColor = Color.FromArgb(108, 117, 125),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 8, FontStyle.Bold)
                };
                btnDlgClear.FlatAppearance.BorderSize = 0;
                pFilter.Controls.Add(btnDlgClear);

                BringInputsToFront(pFilter);

                // ── BẢNG PO ──
                var dgvDlg = new DataGridView
                {
                    Location = new Point(10, 82),
                    Size = new Size(860, 290),
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
                dgvDlg.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 150, 100);
                dgvDlg.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dgvDlg.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                dgvDlg.EnableHeadersVisualStyles = false;
                dgvDlg.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(235, 255, 245);
                dlg.Controls.Add(dgvDlg);

                // Hàm bind bảng PO theo filter
                Action bindDlgGrid = () =>
                {
                    string selDa = cboDaFilter.SelectedItem?.ToString() ?? "Tất cả";
                    string selMpr = txtMprFilter.Text.Trim().ToLower();
                    var rows = dtPO.AsEnumerable().Where(r =>
                    {
                        if (selDa != "Tất cả" && r["MaDuAn"].ToString() != selDa) return false;
                        if (!string.IsNullOrEmpty(selMpr) && !r["MPR_No"].ToString().ToLower().Contains(selMpr)) return false;
                        return true;
                    });
                    dgvDlg.DataSource = rows.Any() ? rows.CopyToDataTable() : dtPO.Clone();
                };
                bindDlgGrid();

                btnDlgFilter.Click += (s, ev) => bindDlgGrid();
                btnDlgClear.Click += (s, ev) => { cboDaFilter.SelectedIndex = 0; txtMprFilter.Text = ""; bindDlgGrid(); };
                cboDaFilter.SelectedIndexChanged += (s, ev) => bindDlgGrid();
                txtMprFilter.KeyDown += (s, ev) => { if (ev.KeyCode == Keys.Enter) bindDlgGrid(); };

                // ── KHU VỰC NHẬP THÊM THÔNG TIN ──
                int iy = dlg.ClientSize.Height - 140;

                // Expect Delivery
                dlg.Controls.Add(new Label
                {
                    Text = "Expect Delivery:",
                    Location = new Point(10, iy + 3),
                    Size = new Size(110, 20),
                    Font = new Font("Segoe UI", 9, FontStyle.Bold)
                });

                var cboExpDelivery = new ComboBox
                {
                    Location = new Point(125, iy),
                    Size = new Size(180, 25),
                    Font = new Font("Segoe UI", 9),
                    DropDownStyle = ComboBoxStyle.DropDownList
                };
                // Gợi ý: hôm nay + 7/14/30/60/90 ngày
                var today = DateTime.Today;
                new[] { 7, 14, 30, 60, 90 }.ToList().ForEach(d =>
                    cboExpDelivery.Items.Add(today.AddDays(d).ToString("dd/MM/yyyy") + $"  (+{d} ngày)"));
                // Thêm tùy chọn DateTimePicker
                cboExpDelivery.Items.Add("-- Chọn ngày khác --");
                cboExpDelivery.SelectedIndex = 0;
                dlg.Controls.Add(cboExpDelivery);

                var dtpCustomDate = new DateTimePicker
                {
                    Location = new Point(315, iy),
                    Size = new Size(130, 25),
                    Font = new Font("Segoe UI", 9),
                    Format = DateTimePickerFormat.Short,
                    Visible = false
                };
                dlg.Controls.Add(dtpCustomDate);
                cboExpDelivery.SelectedIndexChanged += (s, ev) =>
                    dtpCustomDate.Visible = cboExpDelivery.SelectedItem?.ToString().StartsWith("--") == true;

                // Ghi chú
                dlg.Controls.Add(new Label
                {
                    Text = "Ghi chú:",
                    Location = new Point(460, iy + 3),
                    Size = new Size(60, 20),
                    Font = new Font("Segoe UI", 9, FontStyle.Bold)
                });
                var txtDlgNote = new TextBox
                {
                    Location = new Point(525, iy),
                    Size = new Size(345, 25),
                    Font = new Font("Segoe UI", 9),
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
                };
                dlg.Controls.Add(txtDlgNote);

                // ── NÚT OK & HỦY ──
                var btnOK = new Button
                {
                    Text = "✔ OK",
                    Location = new Point(dlg.ClientSize.Width - 220, dlg.ClientSize.Height - 45),
                    Size = new Size(90, 32),
                    BackColor = Color.FromArgb(40, 167, 69),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Right
                };
                btnOK.FlatAppearance.BorderSize = 0;
                dlg.Controls.Add(btnOK);

                var btnCancel = new Button
                {
                    Text = "Hủy",
                    Location = new Point(dlg.ClientSize.Width - 120, dlg.ClientSize.Height - 45),
                    Size = new Size(85, 32),
                    BackColor = Color.FromArgb(108, 117, 125),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                    DialogResult = DialogResult.Cancel
                };
                btnCancel.FlatAppearance.BorderSize = 0;
                dlg.Controls.Add(btnCancel);
                dlg.CancelButton = btnCancel;

                // Resize sync
                dlg.Resize += (s, ev) =>
                {
                    int newIy = dlg.ClientSize.Height - 140;
                    cboExpDelivery.Top = newIy; dtpCustomDate.Top = newIy;
                    txtDlgNote.Top = newIy;
                    foreach (Control c in dlg.Controls)
                        if (c is Label lbl && lbl.Text == "Expect Delivery:") { lbl.Top = newIy + 3; break; }
                    btnOK.Location = new Point(dlg.ClientSize.Width - 220, dlg.ClientSize.Height - 45);
                    btnCancel.Location = new Point(dlg.ClientSize.Width - 120, dlg.ClientSize.Height - 45);
                    pFilter.Width = dlg.ClientSize.Width - 20;
                    dgvDlg.Width = dlg.ClientSize.Width - 20;
                    dgvDlg.Height = dlg.ClientSize.Height - 82 - 145;
                };

                BringInputsToFront(dlg);

                // ── XỬ LÝ OK ──
                btnOK.Click += (s, ev) =>
                {
                    if (dgvDlg.SelectedRows.Count == 0)
                    {
                        MessageBox.Show("Vui lòng chọn một PO!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    string selPONo = dgvDlg.SelectedRows[0].Cells["PONo"]?.Value?.ToString() ?? "";
                    string selDaAn = dgvDlg.SelectedRows[0].Cells["MaDuAn"]?.Value?.ToString() ?? "";

                    // Lấy ngày giao hàng
                    DateTime expDate;
                    if (dtpCustomDate.Visible)
                        expDate = dtpCustomDate.Value.Date;
                    else
                    {
                        string datePart = cboExpDelivery.SelectedItem?.ToString().Split(' ')[0] ?? "";
                        if (!DateTime.TryParseExact(datePart, "dd/MM/yyyy",
                            System.Globalization.CultureInfo.InvariantCulture,
                            System.Globalization.DateTimeStyles.None, out expDate))
                            expDate = DateTime.Today.AddDays(30);
                    }

                    string note = txtDlgNote.Text.Trim();

                    try
                    {
                        string sqlIns = @"
                            IF NOT EXISTS (SELECT 1 FROM PO_DeliveryTracking WHERE PONo = @poNo AND ExpDelivery = @exp)
                            INSERT INTO PO_DeliveryTracking (PONo, ExpDelivery, GhiChu, Status, ReceiverNote, Created_Date)
                            VALUES (@poNo, @exp, @note, 'Pending', '', GETDATE())";
                        using (var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection())
                        {
                            conn.Open();
                            var cmd = new Microsoft.Data.SqlClient.SqlCommand(sqlIns, conn);
                            cmd.Parameters.AddWithValue("@poNo", selPONo);
                            cmd.Parameters.AddWithValue("@exp", expDate);
                            cmd.Parameters.AddWithValue("@note", note);
                            cmd.ExecuteNonQuery();
                        }
                        LoadDeliveries();
                        dlg.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Lỗi lưu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                };

                dlg.KeyDown += (s, ev) => { if (ev.KeyCode == Keys.Enter) { btnOK.PerformClick(); ev.Handled = true; } };
                dlg.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi mở popup: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnClearHeader_Click(object sender, EventArgs e)
        {
            ClearHeader(); dgvDetails.Rows.Clear(); dgvFiles.Rows.Clear();
            dgvDelivery.Rows.Clear(); UpdateTotal(); _selectedPO_ID = 0; LoadPO();
        }

        // =========================================================================
        // RECEIVED HISTORY — Lịch sử nhận hàng theo PO
        // =========================================================================
        private void BtnReceivedHistory_Click(object sender, EventArgs e)
        {
            ShowDeliveryHistoryPopup();
        }

        private void ShowDeliveryHistoryPopup()
        {
            try
            {
                // Load TOÀN BỘ lịch sử từ PO_DeliveryTracking (kể cả đã Done và quá hạn)
                string sql = @"
                    SELECT
                        dt.PONo                                                 AS [PO No],
                        ISNULL(pi.ProjectCode, N'')                             AS [Mã dự án],
                        CONVERT(NVARCHAR(10), dt.ExpDelivery, 103)              AS [Exp. Delivery],
                        ISNULL(dt.GhiChu, N'')                                  AS [Ghi chú],
                        ISNULL(dt.Status, N'Pending')                           AS [Trạng thái],
                        ISNULL(dt.ReceiverNote, N'')                            AS [Receiver Note],
                        CONVERT(NVARCHAR(16), dt.Created_Date, 103)             AS [Ngày tạo]
                    FROM PO_DeliveryTracking dt
                    LEFT JOIN PO_head ph ON ph.PONo = dt.PONo
                    LEFT JOIN ProjectInfo pi ON pi.WorkorderNo = ph.WorkorderNo
                    ORDER BY dt.Created_Date DESC";

                System.Data.DataTable dt;
                using (var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    dt = new System.Data.DataTable();
                    dt.Load(new Microsoft.Data.SqlClient.SqlCommand(sql, conn).ExecuteReader());
                }

                // ── Popup ──
                var popup = new Form
                {
                    Text = "📋 Delivery Tracking History",
                    Size = new Size(900, 520),
                    StartPosition = FormStartPosition.CenterParent,
                    BackColor = Color.FromArgb(245, 245, 245),
                    MinimumSize = new Size(700, 400)
                };

                popup.Controls.Add(new Label
                {
                    Text = "📋  Lịch sử theo dõi giao hàng",
                    Font = new Font("Segoe UI", 10, FontStyle.Bold),
                    ForeColor = Color.FromArgb(0, 150, 100),
                    Location = new Point(10, 8),
                    Size = new Size(500, 24)
                });

                // Thống kê nhanh
                int total = dt.Rows.Count;
                int done = 0, pending = 0, overdue = 0;
                foreach (System.Data.DataRow r in dt.Rows)
                {
                    string st = r["Trạng thái"]?.ToString() ?? "";
                    if (st == "Done") done++;
                    else if (st == "Overdue") overdue++;
                    else pending++;
                }
                popup.Controls.Add(new Label
                {
                    Text = $"Tổng: {total}  |  ✅ Done: {done}  |  ⏳ Pending: {pending}  |  ⚠ Overdue: {overdue}",
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    ForeColor = Color.FromArgb(0, 120, 212),
                    Location = new Point(10, 36),
                    Size = new Size(860, 20),
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
                });

                // DataGridView
                var dgv = new DataGridView
                {
                    Location = new Point(10, 64),
                    Size = new Size(popup.ClientSize.Width - 20, popup.ClientSize.Height - 110),
                    ReadOnly = true,
                    AllowUserToAddRows = false,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                    BackgroundColor = Color.White,
                    BorderStyle = BorderStyle.FixedSingle,
                    RowHeadersVisible = false,
                    Font = new Font("Segoe UI", 9),
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                    DataSource = dt
                };
                dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 150, 100);
                dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                dgv.EnableHeadersVisualStyles = false;
                dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(235, 255, 245);

                dgv.CellFormatting += (s, ev) =>
                {
                    if (ev.RowIndex < 0) return;
                    if (dgv.Columns[ev.ColumnIndex].Name == "Trạng thái")
                    {
                        string val = ev.Value?.ToString() ?? "";
                        ev.CellStyle.ForeColor =
                            val == "Done" ? Color.FromArgb(40, 167, 69) :
                            val == "Overdue" ? Color.FromArgb(220, 53, 69) :
                            Color.FromArgb(255, 140, 0);
                        ev.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    }
                };
                dgv.RowPrePaint += (s, ev) =>
                {
                    if (ev.RowIndex < 0 || dgv.Rows[ev.RowIndex].IsNewRow) return;
                    string st = dgv.Rows[ev.RowIndex].Cells["Trạng thái"].Value?.ToString() ?? "";
                    dgv.Rows[ev.RowIndex].DefaultCellStyle.BackColor =
                        st == "Done" ? Color.FromArgb(235, 255, 235) :
                        st == "Overdue" ? Color.FromArgb(255, 235, 235) :
                        Color.White;
                };
                popup.Controls.Add(dgv);

                // Nút đóng
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
                popup.AcceptButton = btnClose;
                popup.CancelButton = btnClose;
                popup.Resize += (s, ev) =>
                {
                    btnClose.Location = new Point(popup.ClientSize.Width - 115, popup.ClientSize.Height - 40);
                };

                popup.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải lịch sử: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ClearHeader()
        {
            txtPONo.Text = ""; txtProjectName.Text = ""; txtWorkorderNo.Text = ""; txtMPRNo.Text = "";
            txtPrepared.Text = ""; txtReviewed.Text = ""; txtAgreement.Text = "";
            txtApproved.Text = ""; txtNotes.Text = "";
            nudRevise.Value = 0; dtpPODate.Value = DateTime.Today; cboStatus.SelectedIndex = 0;
        }

        // =========================================================================
        // CHECK BY SIZE — Popup load TOÀN BỘ dữ liệu + bộ lọc
        // =========================================================================
        private void BtnCheckBySize_Click(object sender, EventArgs e)
        {
            ShowCheckBySizePopup();
        }

        private void ShowCheckBySizePopup()
        {
            try
            {
                // ── Query load TOÀN BỘ dữ liệu tất cả PO ──
                const string SQL_ALL = @"
                    SELECT
                        ISNULL(pi.ProjectCode, N'')                         AS [Mã dự án],
                        ph.PONo                                             AS [PO No],
                        pod.Item_Name                                       AS [Tên vật tư],
                        ISNULL(pod.Asize, N'')                              AS [A(mm)],
                        ISNULL(pod.Bsize, N'')                              AS [B(mm)],
                        ISNULL(pod.Csize, N'')                              AS [C(mm)],
                        pod.Qty_Per_Sheet                                   AS [SL đặt],
                        ISNULL(rd.Qty_Required, 0)                          AS [SL YC kiểm],
                        ISNULL(rd.Qty_Received, 0)                          AS [SL đã nhận],
                        ISNULL(rd.Heatno,  N'')                             AS [Heat No],
                        ISNULL(rd.MTRno,   N'')                             AS [MTR No],
                        ISNULL(rd.Inspect_Result, N'Chưa KT')               AS [Trạng thái KT],
                        ISNULL(rh.RIR_No,  N'')                             AS [RIR No],
                        ISNULL(CONVERT(NVARCHAR(10), rh.Issue_Date, 103), N'') AS [Ngày RIR],
                        ISNULL(rh.Status,  N'')                             AS [Trạng thái RIR]
                    FROM PO_head ph
                    INNER JOIN PO_Detail pod  ON pod.PO_ID = ph.PO_ID
                    LEFT  JOIN ProjectInfo pi ON pi.WorkorderNo = ph.WorkorderNo
                    LEFT  JOIN RIR_head rh    ON rh.PONo = ph.PONo
                    LEFT  JOIN RIR_detail rd  ON rd.RIR_ID = rh.RIR_ID
                        AND (
                            ISNULL(rd.Size, N'') LIKE N'%' + ISNULL(pod.Asize, N'') + N'%'
                            OR pod.Asize IS NULL OR pod.Asize = N''
                        )
                    ORDER BY ph.PONo, pod.Item_No, rh.RIR_No";

                DataTable dtFull;
                using (var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    dtFull = new DataTable();
                    dtFull.Load(new Microsoft.Data.SqlClient.SqlCommand(SQL_ALL, conn).ExecuteReader());
                }

                // ── Tạo popup ──
                var popup = new Form
                {
                    Text = "🔍 Check by Size — Toàn bộ dữ liệu RIR",
                    Size = new Size(1350, 720),
                    StartPosition = FormStartPosition.CenterParent,
                    BackColor = Color.FromArgb(245, 245, 245),
                    MinimumSize = new Size(1000, 500)
                };

                // ── TIÊU ĐỀ ──
                popup.Controls.Add(new Label
                {
                    Text = "🔍  CHECK BY SIZE  —  Tra cứu kết quả kiểm tra RIR theo kích thước vật tư",
                    Font = new Font("Segoe UI", 11, FontStyle.Bold),
                    ForeColor = Color.FromArgb(102, 51, 153),
                    Location = new Point(10, 8),
                    Size = new Size(900, 26),
                    Anchor = AnchorStyles.Top | AnchorStyles.Left
                });

                // ── PANEL BỘ LỌC ──
                var panelFilter = new Panel
                {
                    Location = new Point(10, 40),
                    Size = new Size(popup.ClientSize.Width - 20, 60),
                    BackColor = Color.White,
                    BorderStyle = BorderStyle.FixedSingle,
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
                };
                popup.Controls.Add(panelFilter);

                // Helper tạo label nhỏ trong filter panel
                Action<string, int> addFLbl = (txt, x) =>
                    panelFilter.Controls.Add(new Label
                    {
                        Text = txt,
                        Location = new Point(x, 8),
                        Size = new Size(70, 18),
                        Font = new Font("Segoe UI", 8, FontStyle.Bold),
                        ForeColor = Color.FromArgb(80, 80, 80)
                    });

                // ComboBox — Mã dự án
                addFLbl("Mã dự án:", 8);
                var cboDuAn = new ComboBox
                {
                    Location = new Point(78, 5),
                    Size = new Size(140, 25),
                    Font = new Font("Segoe UI", 9),
                    DropDownStyle = ComboBoxStyle.DropDownList
                };
                cboDuAn.Items.Add("Tất cả");
                var distinctProjects = dtFull.AsEnumerable()
                    .Select(r => r["Mã dự án"].ToString())
                    .Where(v => !string.IsNullOrEmpty(v))
                    .Distinct().OrderBy(v => v).ToList();
                foreach (var p in distinctProjects) cboDuAn.Items.Add(p);
                cboDuAn.SelectedIndex = 0;
                panelFilter.Controls.Add(cboDuAn);

                // TextBox — Tên vật tư
                addFLbl("Tên:", 232);
                var txtFilterName = new TextBox
                {
                    Location = new Point(260, 5),
                    Size = new Size(170, 25),
                    Font = new Font("Segoe UI", 9),
                    PlaceholderText = "Tên vật tư..."
                };
                panelFilter.Controls.Add(txtFilterName);

                // TextBox — A(mm)
                addFLbl("A(mm):", 445);
                var txtFilterA = new TextBox
                {
                    Location = new Point(490, 5),
                    Size = new Size(80, 25),
                    Font = new Font("Segoe UI", 9),
                    PlaceholderText = "A..."
                };
                panelFilter.Controls.Add(txtFilterA);

                // TextBox — B(mm)
                addFLbl("B(mm):", 583);
                var txtFilterB = new TextBox
                {
                    Location = new Point(628, 5),
                    Size = new Size(80, 25),
                    Font = new Font("Segoe UI", 9),
                    PlaceholderText = "B..."
                };
                panelFilter.Controls.Add(txtFilterB);

                // TextBox — C(mm)
                addFLbl("C(mm):", 721);
                var txtFilterC = new TextBox
                {
                    Location = new Point(766, 5),
                    Size = new Size(80, 25),
                    Font = new Font("Segoe UI", 9),
                    PlaceholderText = "C..."
                };
                panelFilter.Controls.Add(txtFilterC);

                // Nút Tìm
                var btnFilter = new Button
                {
                    Text = "🔍 Tìm",
                    Location = new Point(858, 4),
                    Size = new Size(80, 28),
                    BackColor = Color.FromArgb(0, 120, 212),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    Cursor = Cursors.Hand
                };
                btnFilter.FlatAppearance.BorderSize = 0;
                panelFilter.Controls.Add(btnFilter);

                // Nút Xóa lọc
                var btnClearFilter = new Button
                {
                    Text = "✖ Xóa lọc",
                    Location = new Point(948, 4),
                    Size = new Size(85, 28),
                    BackColor = Color.FromArgb(108, 117, 125),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    Cursor = Cursors.Hand
                };
                btnClearFilter.FlatAppearance.BorderSize = 0;
                panelFilter.Controls.Add(btnClearFilter);

                // ── Đưa tất cả input controls trong filter panel lên trên label ──
                BringInputsToFront(panelFilter);

                // ── LABEL THỐNG KÊ ──
                var lblStat = new Label
                {
                    Text = "",
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    ForeColor = Color.FromArgb(0, 120, 212),
                    Location = new Point(10, 108),
                    Size = new Size(1300, 20),
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
                };
                popup.Controls.Add(lblStat);

                // ── DATAGRIDVIEW ──
                var dgv = new DataGridView
                {
                    Location = new Point(10, 132),
                    Size = new Size(popup.ClientSize.Width - 20, popup.ClientSize.Height - 180),
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
                popup.Controls.Add(dgv);

                // ── CellFormatting ──
                dgv.CellFormatting += (s, ev) =>
                {
                    if (ev.RowIndex < 0) return;
                    string colName = dgv.Columns[ev.ColumnIndex].Name;
                    if (colName == "Trạng thái KT")
                    {
                        string val = ev.Value?.ToString() ?? "";
                        ev.CellStyle.ForeColor =
                            val == "Pass" ? Color.FromArgb(40, 167, 69) :
                            val == "Fail" ? Color.FromArgb(220, 53, 69) :
                            val == "Hold" ? Color.FromArgb(255, 140, 0) :
                            Color.FromArgb(108, 117, 125);
                        ev.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    }
                    if (colName == "Trạng thái RIR")
                    {
                        string val = ev.Value?.ToString() ?? "";
                        ev.CellStyle.ForeColor =
                            val == "Hoàn thành" ? Color.FromArgb(40, 167, 69) :
                            val == "Đang kiểm tra" ? Color.FromArgb(255, 140, 0) :
                            string.IsNullOrEmpty(val) ? Color.FromArgb(180, 180, 180) :
                            Color.FromArgb(0, 120, 212);
                        ev.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    }
                };

                // ── RowPrePaint ──
                dgv.RowPrePaint += (s, ev) =>
                {
                    if (ev.RowIndex < 0 || dgv.Rows[ev.RowIndex].IsNewRow) return;
                    string kt = dgv.Rows[ev.RowIndex].Cells["Trạng thái KT"].Value?.ToString() ?? "";
                    dgv.Rows[ev.RowIndex].DefaultCellStyle.BackColor =
                        kt == "Pass" ? Color.FromArgb(235, 255, 235) :
                        kt == "Fail" ? Color.FromArgb(255, 235, 235) :
                        kt == "Hold" ? Color.FromArgb(255, 248, 230) :
                        Color.White;
                };

                // ── HÀM ÁP DỤNG BỘ LỌC CLIENT-SIDE ──
                Action applyFilter = () =>
                {
                    string selProject = cboDuAn.SelectedItem?.ToString() ?? "Tất cả";
                    string kName = txtFilterName.Text.Trim().ToLower();
                    string kA = txtFilterA.Text.Trim();
                    string kB = txtFilterB.Text.Trim();
                    string kC = txtFilterC.Text.Trim();

                    var filtered = dtFull.AsEnumerable().Where(r =>
                    {
                        if (selProject != "Tất cả" && r["Mã dự án"].ToString() != selProject) return false;
                        if (!string.IsNullOrEmpty(kName) && !r["Tên vật tư"].ToString().ToLower().Contains(kName)) return false;
                        if (!string.IsNullOrEmpty(kA) && !r["A(mm)"].ToString().Contains(kA)) return false;
                        if (!string.IsNullOrEmpty(kB) && !r["B(mm)"].ToString().Contains(kB)) return false;
                        if (!string.IsNullOrEmpty(kC) && !r["C(mm)"].ToString().Contains(kC)) return false;
                        return true;
                    });

                    DataTable dtView = filtered.Any() ? filtered.CopyToDataTable() : dtFull.Clone();
                    dgv.DataSource = dtView;

                    // Cập nhật thống kê
                    int total = dtView.Rows.Count, pass = 0, fail = 0, hold = 0, notYet = 0;
                    foreach (DataRow dr in dtView.Rows)
                    {
                        string kt = dr["Trạng thái KT"]?.ToString() ?? "";
                        if (kt == "Pass") pass++;
                        else if (kt == "Fail") fail++;
                        else if (kt == "Hold") hold++;
                        else notYet++;
                    }
                    lblStat.Text = $"Hiển thị: {total} dòng  |  ✅ Pass: {pass}  |  ❌ Fail: {fail}  |  ⏸ Hold: {hold}  |  ⏳ Chưa KT: {notYet}";
                };

                // Áp dụng lọc ngay khi mở (hiển thị toàn bộ)
                applyFilter();

                // Kết nối sự kiện
                btnFilter.Click += (s, ev) => applyFilter();
                btnClearFilter.Click += (s, ev) =>
                {
                    cboDuAn.SelectedIndex = 0;
                    txtFilterName.Text = "";
                    txtFilterA.Text = "";
                    txtFilterB.Text = "";
                    txtFilterC.Text = "";
                    applyFilter();
                };

                // Enter trong bất kỳ TextBox nào → tìm kiếm (dùng KeyPreview ở cấp Form)
                popup.KeyPreview = true;
                popup.KeyDown += (s, ev) =>
                {
                    if (ev.KeyCode == Keys.Enter)
                    {
                        applyFilter();
                        ev.Handled = true;
                        ev.SuppressKeyPress = true;
                    }
                };

                cboDuAn.SelectedIndexChanged += (s, ev) => applyFilter();

                // ── NÚT ĐÓNG ──
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
                popup.AcceptButton = btnFilter;   // Enter → Tìm kiếm
                popup.CancelButton = btnClose;    // Escape → Đóng

                // Resize handler
                popup.Resize += (s, ev) =>
                {
                    btnClose.Location = new Point(popup.ClientSize.Width - 115, popup.ClientSize.Height - 40);
                    panelFilter.Width = popup.ClientSize.Width - 20;
                };

                popup.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tra cứu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}