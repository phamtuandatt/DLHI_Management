using MPR_Managerment.Models;
using MPR_Managerment.Helpers;
using MPR_Managerment.Services;
using System;
using System.Globalization;
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
        private string _currentUser = AppSession.CurrentUser.Full_Name ?? "Admin";

        private string _targetPoNo = "";
        private string _importMprNo = "";

        private DataGridView dgvPO;
        private TextBox txtSearch;
        private Button btnSearch, btnNewPO, btnDeletePO, btnClearHeader, btnExport, btnSavePO;
        private Button btnSearchBySupp;
        private Label lblStatus;

        private TextBox txtPONo, txtProjectName, txtWorkorderNo, txtMPRNo;
        private TextBox txtPrepared, txtReviewed, txtAgreement, txtApproved, txtNotes;
        private ComboBox cboPaymentTerm;
        private DateTimePicker dtpPODate, dtpPOExpectDelivery;
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
        // Culture dinh dang so: dau "." ngan hang nghin, dau "," thap phan
        private static readonly System.Globalization.CultureInfo _numCulture =
            new System.Globalization.CultureInfo("vi-VN");
        private Panel panelTop, panelHeader, panelDetail;
        private DataGridView dgvMPRFiles; // Bảng file MPR Link bên phải chi tiết
        private ComboBox cboSupplier;
        private System.Data.DataTable _supplierTable;
        private bool _isSearching = false;

        private string _projectCodeImport = string.Empty;

        public frmPO(string poNo = "")
        {
            _targetPoNo = poNo;
            InitializeComponent();
            BuildUI();
            ApplyPermissions();
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
            dgvPO.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
            dgvPO.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgvPO.SelectionChanged += DgvPO_SelectionChanged;
            dgvPO.DataBindingComplete += (s, ev) => dgvPO.ClearSelection();

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
            var btnPayment = CreateButton("💳 Payment", Color.FromArgb(0, 150, 100), new Point(850, 48), 110, 30);
            btnPayment.Click += (s, ev) =>

            {
                string poNo = "";
                if (dgvPO.SelectedRows.Count > 0)
                    poNo = dgvPO.SelectedRows[0].Cells["PO_No"].Value?.ToString() ?? "";

                var frm = new frmPayment(_currentUser, poNo);
                frm.Show();
            };
            panelTop.Controls.Add(btnPayment);
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
            dgvFiles.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
            dgvFiles.DefaultCellStyle.SelectionForeColor = Color.Black;
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
            dgvDelivery.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
            dgvDelivery.DefaultCellStyle.SelectionForeColor = Color.Black;
            // Cột ẩn
            dgvDelivery.Columns.Add(new DataGridViewTextBoxColumn { Name = "TrackID", HeaderText = "ID", Visible = false });
            // Cột hiển thị
            dgvDelivery.Columns.Add(new DataGridViewTextBoxColumn { Name = "PONo", HeaderText = "PO No", ReadOnly = true, FillWeight = 25 });
            dgvDelivery.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaDuAn", HeaderText = "Mã DA", ReadOnly = true, FillWeight = 15 });
            dgvDelivery.Columns.Add(new DataGridViewTextBoxColumn { Name = "NCC", HeaderText = "Nhà CC", ReadOnly = true, FillWeight = 22 });
            dgvDelivery.Columns.Add(new DataGridViewTextBoxColumn { Name = "ExpDelivery", HeaderText = "Exp.Deliv", ReadOnly = true, FillWeight = 22 });
            dgvDelivery.Columns.Add(new DataGridViewTextBoxColumn { Name = "GhiChu", HeaderText = "Ghi chú", ReadOnly = true, FillWeight = 20 });
            dgvDelivery.Columns.Add(new DataGridViewTextBoxColumn { Name = "Status", HeaderText = "T.Thái", ReadOnly = true, FillWeight = 13 });
            dgvDelivery.Columns.Add(new DataGridViewTextBoxColumn { Name = "ReceiverNote", HeaderText = "Ghi chú nhận", ReadOnly = false, FillWeight = 20 });

            // Màu sắc trạng thái
            dgvDelivery.CellFormatting += (s, ev) =>
            {
                if (ev.RowIndex < 0) return;
                string colName = dgvDelivery.Columns[ev.ColumnIndex].Name;

                if (colName == "Status")
                {
                    string v = ev.Value?.ToString() ?? "";
                    ev.CellStyle.ForeColor =
                        v == "Done" ? Color.FromArgb(40, 167, 69) :
                        v == "Overdue" ? Color.FromArgb(220, 53, 69) :
                        Color.FromArgb(255, 140, 0);
                    ev.CellStyle.Font = new Font("Segoe UI", 8, FontStyle.Bold);
                }

                // Tô màu vàng ô Exp.Deliv nếu trùng ngày hôm nay
                if (colName == "ExpDelivery")
                {
                    string dateStr = ev.Value?.ToString() ?? "";
                    if (DateTime.TryParseExact(dateStr, new[] { "dd/MM/yyyy", "yyyy-MM-dd", "M/d/yyyy", "d/M/yyyy" },
                        System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None, out DateTime expDate))
                    {
                        if (expDate.Date == DateTime.Today)
                        {
                            ev.CellStyle.BackColor = Color.Yellow;
                            ev.CellStyle.ForeColor = Color.FromArgb(100, 60, 0);
                            ev.CellStyle.Font = new Font("Segoe UI", 8, FontStyle.Bold);
                        }
                    }
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
            // Load MPR Files khi MPR No thay đổi (người dùng nhập xong và rời ô)
            txtMPRNo.Leave += (s, e) => LoadMPRFiles();

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
            AddLabel(panelHeader, "Payment Term:", 295, y);
            cboPaymentTerm = new ComboBox
            {
                Location = new Point(390, y),
                Size = new Size(200, 25),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboPaymentTerm.Items.AddRange(new[]
            {
                "T/T 100% within 7days after delivery",
                "T/T 30% Advance + 70% before shipment",
                "T/T 50% Advance + 50% before shipment",
                "T/T 100% Advance",
                "T/T 30 days after shipment",
                "T/T 45 days after shipment",
                "T/T 60 days after shipment",
                "L/C at sight",
                "L/C 30 days",
                "L/C 60 days",
                "D/P at sight",
                "D/A 30 days",
                "Cash on Delivery",
                "Net 30",
                "Net 45",
                "Net 60"
            });
            cboPaymentTerm.SelectedIndex = 0;
            cboPaymentTerm.BringToFront();
            panelHeader.Controls.Add(cboPaymentTerm);
            AddLabel(panelHeader, "Ngày Giao hàng:", 610, y);
            dtpPOExpectDelivery = new DateTimePicker { Location = new Point(680, y), Size = new Size(100, 25), Font = new Font("Segoe UI", 9), Format = DateTimePickerFormat.Short };
            panelHeader.Controls.Add(dtpPOExpectDelivery);

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
            panelDetail = new Panel { Location = new Point(10, 500), Size = new Size(980, 285), BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom };
            this.Controls.Add(panelDetail);

            // ── Panel MPR Files (bên phải panelDetail) ──
            var panelMPRFiles = new Panel
            {
                Location = new Point(panelDetail.Right + 10, panelDetail.Top),
                Size = new Size(280, 285),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom
            };
            this.Controls.Add(panelMPRFiles);

            panelMPRFiles.Controls.Add(new Label
            {
                Text = "📁 MPR Files",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(102, 51, 153),
                Location = new Point(6, 8),
                Size = new Size(260, 20)
            });

            dgvMPRFiles = new DataGridView
            {
                Location = new Point(4, 30),
                Size = new Size(272, 248),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                Cursor = Cursors.Hand
            };
            dgvMPRFiles.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(102, 51, 153);
            dgvMPRFiles.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvMPRFiles.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvMPRFiles.EnableHeadersVisualStyles = false;
            dgvMPRFiles.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 240, 255);
            dgvMPRFiles.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
            dgvMPRFiles.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgvMPRFiles.Columns.Add(new DataGridViewTextBoxColumn { Name = "FileName", HeaderText = "Tên file" });
            dgvMPRFiles.Columns.Add(new DataGridViewTextBoxColumn { Name = "FullPath", HeaderText = "Path", Visible = false });

            // Double click → mở file
            dgvMPRFiles.CellDoubleClick += (s, ev) =>
            {
                if (ev.RowIndex < 0) return;
                string path = dgvMPRFiles.Rows[ev.RowIndex].Cells["FullPath"].Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(path) && System.IO.File.Exists(path))
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    { FileName = path, UseShellExecute = true });
                else
                    MessageBox.Show("Không tìm thấy file!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            };
            panelMPRFiles.Controls.Add(dgvMPRFiles);
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

            // ── Panel phai: chua labels TruocVAT/SauVAT + apply VAT + apply Calc ──
            // Anchor Right de tu can le phai khi resize
            var panelRight = new Panel
            {
                Size = new Size(590, 68),
                BackColor = Color.Transparent,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            // Vi tri: sat le phai panelDetail, dong voi hang buttons
            panelRight.Location = new Point(panelDetail.Width - 590 - 5, 2);
            panelDetail.Resize += (s, ev) =>
                panelRight.Location = new Point(panelDetail.Width - panelRight.Width - 5, 2);

            // -- Cot 1: Truoc VAT + Sau VAT (x=0, w=185) --
            lblSubTotal = new Label
            {
                Location = new Point(0, 4),
                Size = new Size(185, 20),
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                ForeColor = Color.FromArgb(200, 53, 53),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };
            panelRight.Controls.Add(lblSubTotal);

            lblTotal = new Label
            {
                Location = new Point(0, 28),
                Size = new Size(185, 20),
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 100, 200),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };
            panelRight.Controls.Add(lblTotal);

            // -- Cot 2: Apply VAT (x=190) --
            int xVAT = 190;
            panelRight.Controls.Add(new Label
            {
                Text = "VAT(%):",
                Location = new Point(xVAT, 4),
                Size = new Size(50, 20),
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            });
            var cboApplyVAT = new ComboBox
            {
                Location = new Point(xVAT, 26),
                Size = new Size(72, 22),
                Font = new Font("Segoe UI", 8),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboApplyVAT.Items.AddRange(new object[] { "10%", "8%", "Khong VAT" });
            cboApplyVAT.SelectedIndex = 0;
            panelRight.Controls.Add(cboApplyVAT);
            var btnApplyVAT = CreateButton("Ap dung", Color.FromArgb(255, 140, 0),
                new Point(xVAT + 76, 26), 62, 22);
            btnApplyVAT.Font = new Font("Segoe UI", 7, FontStyle.Bold);
            btnApplyVAT.Click += (s, ev) =>
            {
                string vatVal = cboApplyVAT.SelectedItem?.ToString() == "Khong VAT" ? "0" :
                    cboApplyVAT.SelectedItem?.ToString()?.Replace("%", "") ?? "10";
                foreach (DataGridViewRow row in dgvDetails.Rows)
                    if (!row.IsNewRow && row.Tag?.ToString() != "TOTAL")
                    { row.Cells["VAT"].Value = vatVal; RecalculateAmount(row.Index); }
                UpdateTotal();
            };
            panelRight.Controls.Add(btnApplyVAT);

            // -- Cot 3: Apply Calc_Method (x=340) --
            int xCalc = 340;
            panelRight.Controls.Add(new Label
            {
                Text = "Cach tinh:",
                Location = new Point(xCalc, 4),
                Size = new Size(65, 20),
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            });
            var cboApplyCalc = new ComboBox
            {
                Location = new Point(xCalc, 26),
                Size = new Size(72, 22),
                Font = new Font("Segoe UI", 8),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboApplyCalc.Items.AddRange(new object[] { "Theo KG", "Theo SL" });
            cboApplyCalc.SelectedIndex = 0;
            panelRight.Controls.Add(cboApplyCalc);
            var btnApplyCalc = CreateButton("Ap dung", Color.FromArgb(255, 140, 0),
                new Point(xCalc + 76, 26), 62, 22);
            btnApplyCalc.Font = new Font("Segoe UI", 7, FontStyle.Bold);
            btnApplyCalc.Click += (s, ev) =>
            {
                string calcVal = cboApplyCalc.SelectedItem?.ToString() ?? "Theo KG";
                foreach (DataGridViewRow row in dgvDetails.Rows)
                    if (!row.IsNewRow && row.Tag?.ToString() != "TOTAL")
                    { row.Cells["Calc_Method"].Value = calcVal; RecalculateAmount(row.Index); }
                UpdateTotal();
            };
            panelRight.Controls.Add(btnApplyCalc);

            panelDetail.Controls.Add(panelRight);
            panelRight.BringToFront();

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
            dgvDetails.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
            dgvDetails.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgvDetails.CellEndEdit += DgvDetails_CellEndEdit; dgvDetails.KeyDown += DgvDetails_KeyDown;
            dgvDetails.CellParsing += DgvDetails_CellParsing;
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
                        h.PO_Date.HasValue ? h.PO_Date.Value.ToString("dd/MM/yyyy") : "", h.Status, h.Total_Amount.ToString("N2", _numCulture));
                }
                var dtFull = dt.Copy();
                System.Data.DataTable dtCurrent = dtFull.Copy();
                string selectedPONo = null;

                var popup = new Form { Text = "🔍 Tìm theo NCC", Size = new Size(1100, 680), StartPosition = FormStartPosition.CenterParent, BackColor = Color.FromArgb(245, 245, 245), MinimumSize = new Size(800, 500) };
                popup.Controls.Add(new Label { Text = "🔍  TÌM KIẾM PO THEO NHÀ CUNG CẤP", Font = new Font("Segoe UI", 11, FontStyle.Bold), ForeColor = Color.FromArgb(102, 51, 153), Location = new Point(10, 8), Size = new Size(700, 26) });

                // ── Filter panel — 1 hàng: NCC | Dự án (dropdown multi-select) | T.Thái | Tìm | Xóa ──
                const int PF_H = 46;
                var pF = new Panel { Location = new Point(10, 38), Size = new Size(popup.ClientSize.Width - 20, PF_H), BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
                popup.Controls.Add(pF);
                popup.Resize += (s, ev) => pF.Width = popup.ClientSize.Width - 20;

                // NCC
                pF.Controls.Add(new Label { Text = "NCC:", Location = new Point(8, 12), Size = new Size(35, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold) });
                var txtNCC = new TextBox { Location = new Point(43, 8), Size = new Size(200, 26), Font = new Font("Segoe UI", 9), PlaceholderText = "Tên nhà cung cấp..." };
                pF.Controls.Add(txtNCC);

                // Dự án — button giả dropdown
                pF.Controls.Add(new Label { Text = "Dự án:", Location = new Point(255, 12), Size = new Size(45, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold) });
                var btnDropDA = new Button
                {
                    Location = new Point(302, 8),
                    Size = new Size(220, 26),
                    Text = "Tất cả dự án  ▼",
                    TextAlign = ContentAlignment.MiddleLeft,
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.White,
                    ForeColor = Color.FromArgb(50, 50, 50),
                    Font = new Font("Segoe UI", 9),
                    Cursor = Cursors.Hand
                };
                btnDropDA.FlatAppearance.BorderColor = Color.FromArgb(180, 180, 180);
                btnDropDA.FlatAppearance.BorderSize = 1;
                pF.Controls.Add(btnDropDA);

                // T.Thái
                pF.Controls.Add(new Label { Text = "T.Thái:", Location = new Point(534, 12), Size = new Size(48, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold) });
                var cboTT = new ComboBox { Location = new Point(584, 8), Size = new Size(125, 26), Font = new Font("Segoe UI", 9), DropDownStyle = ComboBoxStyle.DropDownList };
                cboTT.Items.AddRange(new[] { "Tất cả", "Draft", "Pending", "Approved", "In Progress", "Completed", "Cancelled" });
                cboTT.SelectedIndex = 0;
                pF.Controls.Add(cboTT);

                var btnF = new Button { Text = "🔍 Tìm", Location = new Point(720, 8), Size = new Size(80, 28), BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Cursor = Cursors.Hand };
                btnF.FlatAppearance.BorderSize = 0; pF.Controls.Add(btnF);
                var btnClear = new Button { Text = "✖ Xóa lọc", Location = new Point(808, 8), Size = new Size(85, 28), BackColor = Color.FromArgb(108, 117, 125), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Cursor = Cursors.Hand };
                btnClear.FlatAppearance.BorderSize = 0; pF.Controls.Add(btnClear);

                // ── Danh sách dự án unique ──
                var projectList = dtFull.AsEnumerable()
                    .Select(r => r["Dự án"]?.ToString() ?? "")
                    .Where(v => !string.IsNullOrWhiteSpace(v))
                    .Distinct()
                    .OrderBy(v => v)
                    .ToList();

                // ── Panel nổi chứa CheckedListBox — hiện/ẩn khi click btnDropDA ──
                int clbRowH = Math.Min(projectList.Count, 10) * 18 + 8;
                var panelDropDA = new Panel
                {
                    Size = new Size(btnDropDA.Width + 60, clbRowH + 2),
                    BackColor = Color.White,
                    BorderStyle = BorderStyle.FixedSingle,
                    Visible = false
                };
                // Đặt vị trí ngay dưới btnDropDA — tính lại khi show
                popup.Controls.Add(panelDropDA);
                panelDropDA.BringToFront();

                var clbDA = new CheckedListBox
                {
                    Dock = DockStyle.Fill,
                    Font = new Font("Segoe UI", 9),
                    CheckOnClick = true,
                    BorderStyle = BorderStyle.None,
                    BackColor = Color.White
                };
                foreach (var p in projectList) clbDA.Items.Add(p, false);
                panelDropDA.Controls.Add(clbDA);

                // Hàm cập nhật text trên button dropdown
                Action updateDropBtn = () =>
                {
                    var sel = clbDA.CheckedItems.Cast<string>().ToList();
                    btnDropDA.Text = sel.Count == 0
                        ? "Tất cả dự án  ▼"
                        : (sel.Count == 1 ? sel[0] + "  ▼" : $"{sel.Count} dự án đã chọn  ▼");
                    btnDropDA.ForeColor = sel.Count == 0
                        ? Color.FromArgb(50, 50, 50)
                        : Color.FromArgb(102, 51, 153);
                    btnDropDA.Font = new Font("Segoe UI", 9, sel.Count > 0 ? FontStyle.Bold : FontStyle.Regular);
                };

                // Mở/đóng dropdown
                btnDropDA.Click += (s, ev) =>
                {
                    if (panelDropDA.Visible) { panelDropDA.Visible = false; return; }
                    // Tính vị trí tuyệt đối của btnDropDA trong popup
                    var btnPos = popup.PointToClient(btnDropDA.Parent.PointToScreen(btnDropDA.Location));
                    panelDropDA.Location = new Point(btnPos.X, btnPos.Y + btnDropDA.Height + 2);
                    panelDropDA.Width = Math.Max(btnDropDA.Width + 60, 260);
                    panelDropDA.BringToFront();
                    panelDropDA.Visible = true;
                    clbDA.Focus();
                };

                // Đóng dropdown khi click ra ngoài
                popup.MouseDown += (s, ev) =>
                {
                    if (panelDropDA.Visible && !panelDropDA.Bounds.Contains(ev.Location))
                        panelDropDA.Visible = false;
                };

                var lblCount = new Label { Text = "", Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212), Location = new Point(10, 38 + PF_H + 4), Size = new Size(700, 20), Anchor = AnchorStyles.Top | AnchorStyles.Left };
                popup.Controls.Add(lblCount);

                int DGV_TOP = 38 + PF_H + 28;
                var dgv = new DataGridView { Location = new Point(10, DGV_TOP), Size = new Size(popup.ClientSize.Width - 20, popup.ClientSize.Height - DGV_TOP - 50), ReadOnly = true, AllowUserToAddRows = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, BackgroundColor = Color.White, BorderStyle = BorderStyle.FixedSingle, RowHeadersVisible = false, Font = new Font("Segoe UI", 9), AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom };
                dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(102, 51, 153); dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White; dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold); dgv.EnableHeadersVisualStyles = false;
                dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 240, 255);
                dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
                dgv.DefaultCellStyle.SelectionForeColor = Color.Black;
                popup.Controls.Add(dgv);

                popup.Resize += (s, ev) =>
                {
                    dgv.Width = popup.ClientSize.Width - 20;
                    dgv.Height = popup.ClientSize.Height - DGV_TOP - 50;
                    lblCount.Width = popup.ClientSize.Width - 20;
                };

                Action applyFilter = () =>
                {
                    string kNCC = txtNCC.Text.Trim().ToLower();
                    string kTT = cboTT.SelectedItem?.ToString() ?? "Tất cả";
                    var checkedProjects = clbDA.CheckedItems.Cast<string>().ToList();

                    var rows = dtFull.AsEnumerable().Where(r =>
                    {
                        if (!string.IsNullOrEmpty(kNCC) && !r["NCC"].ToString().ToLower().Contains(kNCC)) return false;
                        if (kTT != "Tất cả" && r["Trạng thái"].ToString() != kTT) return false;
                        // Lọc dự án: nếu có chọn thì chỉ lấy dự án được tick
                        if (checkedProjects.Count > 0 && !checkedProjects.Contains(r["Dự án"].ToString())) return false;
                        return true;
                    });
                    dtCurrent = rows.Any() ? rows.CopyToDataTable() : dtFull.Clone();
                    dgv.DataSource = dtCurrent;
                    if (dgv.Columns.Contains("PO_ID")) dgv.Columns["PO_ID"].Visible = false;

                    int total = dtFull.Rows.Count, shown = dtCurrent.Rows.Count;
                    string daInfo = checkedProjects.Count == 0 ? "tất cả dự án" : $"{checkedProjects.Count} dự án";
                    lblCount.Text = $"Hiển thị: {shown} / {total} PO  ({daInfo})";
                };
                // Cập nhật button text + lọc khi check/uncheck dự án
                // (đặt sau khai báo applyFilter để tránh CS0841)
                clbDA.ItemCheck += (s, ev) =>
                {
                    // ItemCheck fires before state changes — dùng BeginInvoke để đọc sau khi state đã update
                    clbDA.BeginInvoke(new Action(() =>
                    {
                        updateDropBtn();
                        applyFilter();
                    }));
                };
                applyFilter();
                btnF.Click += (s, ev) => applyFilter();
                btnClear.Click += (s, ev) =>
                {
                    txtNCC.Text = "";
                    cboTT.SelectedIndex = 0;
                    for (int i = 0; i < clbDA.Items.Count; i++) clbDA.SetItemChecked(i, false);
                    updateDropBtn();
                    panelDropDA.Visible = false;
                    applyFilter();
                };
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

                        // ── 1 SHEET DUY NHẤT chứa toàn bộ detail tất cả PO ──
                        var ws = pkg.Workbook.Worksheets.Add("Chi tiết PO theo NCC");
                        int totalRows = 0;
                        const int TOTAL_COLS = 12;

                        // ── Dòng 1: Tiêu đề ──
                        string nccTitle = txtNCC.Text.Trim();
                        if (string.IsNullOrEmpty(nccTitle)) nccTitle = "Tất cả NCC";
                        ws.Cells[1, 1].Value = $"CHI TIẾT ĐẶT HÀNG THEO NHÀ CUNG CẤP — {nccTitle}";
                        ws.Cells[1, 1, 1, TOTAL_COLS].Merge = true;
                        ws.Cells[1, 1].Style.Font.Bold = true; ws.Cells[1, 1].Style.Font.Size = 13;
                        ws.Cells[1, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(102, 51, 153));
                        ws.Cells[1, 1].Style.Font.Color.SetColor(Color.White);
                        ws.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                        // ── Dòng 2: Thông tin xuất ──
                        ws.Cells[2, 1].Value = $"Tổng PO: {dtCurrent.Rows.Count}   |   Xuất ngày: {DateTime.Now:dd/MM/yyyy HH:mm}";
                        ws.Cells[2, 1, 2, TOTAL_COLS].Merge = true;
                        ws.Cells[2, 1].Style.Font.Italic = true; ws.Cells[2, 1].Style.Font.Size = 9;
                        ws.Cells[2, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        ws.Cells[2, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(240, 230, 255));

                        // ── Dòng 3: Header cột ──
                        string[] hdrs = { "STT", "Tên hàng", "Vật liệu", "A(mm)", "B(mm)", "C(mm)", "SL", "ĐVT", "KG", "Đơn giá", "VAT(%)", "Thành tiền" };
                        for (int c = 0; c < hdrs.Length; c++)
                        {
                            ws.Cells[3, c + 1].Value = hdrs[c];
                            ws.Cells[3, c + 1].Style.Font.Bold = true;
                            ws.Cells[3, c + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            ws.Cells[3, c + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 120, 212));
                            ws.Cells[3, c + 1].Style.Font.Color.SetColor(Color.White);
                            ws.Cells[3, c + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            ws.Cells[3, c + 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        }
                        ws.Row(3).Height = 22;
                        ws.View.FreezePanes(4, 1);

                        int curRow = 4;

                        foreach (System.Data.DataRow dr in dtCurrent.Rows)
                        {
                            int poId = Convert.ToInt32(dr["PO_ID"]);
                            string poNo = dr["PO No"]?.ToString() ?? "";

                            // ── Dòng nhóm PO (màu tím) — Center Across Selection (không Merge) ──
                            ws.Cells[curRow, 1].Value =
                                $"  🛒  PO: {poNo}   |   NCC: {dr["NCC"]}   |   Dự án: {dr["Dự án"]}   |   MPR: {dr["MPR No"]}   |   Ngày PO: {dr["Ngày PO"]}   |   Trạng thái: {dr["Trạng thái"]}";
                            var poHdrRange = ws.Cells[curRow, 1, curRow, TOTAL_COLS];
                            poHdrRange.Style.Font.Bold = true;
                            poHdrRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            poHdrRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(102, 51, 153));
                            poHdrRange.Style.Font.Color.SetColor(Color.White);
                            poHdrRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
                            ws.Row(curRow).Height = 18;
                            curRow++;

                            // ── Các dòng vật tư của PO này ──
                            var dets = _service.GetDetails(poId);
                            decimal sub = 0;
                            for (int i = 0; i < dets.Count; i++)
                            {
                                var d = dets[i];
                                decimal rp = d.Price, bv = d.Qty_Per_Sheet;
                                if ((d.Remarks ?? "").Contains("[CALC:KG]") && d.Weight_kg > 0 && d.Qty_Per_Sheet > 0)
                                { rp = Math.Round((d.Price * d.Qty_Per_Sheet) / d.Weight_kg, 2); bv = d.Weight_kg; }
                                decimal amt = Math.Round(bv * rp, 2); sub += amt;

                                ws.Cells[curRow, 1].Value = i + 1;
                                ws.Cells[curRow, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                ws.Cells[curRow, 2].Value = d.Item_Name ?? "";
                                ws.Cells[curRow, 3].Value = d.Material ?? "";
                                // A/B/C (mm) — tối đa 2 chữ số thập phân, bỏ số 0 thừa
                                ws.Cells[curRow, 4].Value = d.Asize; ws.Cells[curRow, 4].Style.Numberformat.Format = "#,##0.##";
                                ws.Cells[curRow, 5].Value = d.Bsize; ws.Cells[curRow, 5].Style.Numberformat.Format = "#,##0.##";
                                ws.Cells[curRow, 6].Value = d.Csize; ws.Cells[curRow, 6].Style.Numberformat.Format = "#,##0.##";
                                // SL — số nguyên
                                ws.Cells[curRow, 7].Value = d.Qty_Per_Sheet; ws.Cells[curRow, 7].Style.Numberformat.Format = "#,##0";
                                ws.Cells[curRow, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                ws.Cells[curRow, 8].Value = d.UNIT ?? "";
                                ws.Cells[curRow, 8].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                // KG — luôn hiển thị 2 chữ số thập phân
                                ws.Cells[curRow, 9].Value = d.Weight_kg; ws.Cells[curRow, 9].Style.Numberformat.Format = "#,##0.00";
                                ws.Cells[curRow, 9].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                                // Đơn giá — luôn hiển thị 2 chữ số thập phân
                                ws.Cells[curRow, 10].Value = rp; ws.Cells[curRow, 10].Style.Numberformat.Format = "#,##0.00";
                                ws.Cells[curRow, 10].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                                // VAT (%) — số nguyên
                                ws.Cells[curRow, 11].Value = d.VAT; ws.Cells[curRow, 11].Style.Numberformat.Format = "#,##0";
                                ws.Cells[curRow, 11].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                // Thành tiền — luôn hiển thị 2 chữ số thập phân
                                ws.Cells[curRow, 12].Value = amt; ws.Cells[curRow, 12].Style.Numberformat.Format = "#,##0.00";
                                ws.Cells[curRow, 12].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                                if (i % 2 == 1)
                                    for (int c = 1; c <= TOTAL_COLS; c++)
                                    { ws.Cells[curRow, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid; ws.Cells[curRow, c].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(240, 248, 255)); }

                                ws.Cells[curRow, 1, curRow, TOTAL_COLS].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Hair);
                                curRow++; totalRows++;
                            }

                            // ── Sub-Total / VAT / Total của PO này ──
                            decimal vat = Math.Round(sub * 0.1m, 2);
                            ws.Cells[curRow, 11].Value = "Sub-Total:"; ws.Cells[curRow, 11].Style.Font.Bold = true;
                            ws.Cells[curRow, 12].Value = sub; ws.Cells[curRow, 12].Style.Numberformat.Format = "#,##0.00"; ws.Cells[curRow, 12].Style.Font.Bold = true;
                            ws.Cells[curRow, 12].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right; curRow++;
                            ws.Cells[curRow, 11].Value = "VAT (10%):";
                            ws.Cells[curRow, 12].Value = vat; ws.Cells[curRow, 12].Style.Numberformat.Format = "#,##0.00";
                            ws.Cells[curRow, 12].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right; curRow++;
                            ws.Cells[curRow, 10, curRow, 11].Merge = true; ws.Cells[curRow, 10].Value = "TOTAL (incl. VAT):"; ws.Cells[curRow, 10].Style.Font.Bold = true;
                            ws.Cells[curRow, 12].Value = sub + vat; ws.Cells[curRow, 12].Style.Numberformat.Format = "#,##0.00"; ws.Cells[curRow, 12].Style.Font.Bold = true;
                            ws.Cells[curRow, 12].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                            curRow += 2; // dòng trống giữa các PO
                        }

                        ws.Cells[ws.Dimension.Address].AutoFitColumns();
                        ws.Column(2).Width = Math.Min(ws.Column(2).Width, 45);

                        pkg.SaveAs(new System.IO.FileInfo(sfd.FileName));
                        MessageBox.Show($"✅ Đã xuất {dtCurrent.Rows.Count} PO với {totalRows} dòng chi tiết!\nTất cả gộp vào 1 sheet duy nhất.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                        if (row.IsNewRow || row.Tag?.ToString() == "TOTAL") continue;
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
            if (colName == "Price" || colName == "Amount" || colName == "SubAmount")
            {
                if (e.Value != null && decimal.TryParse(e.Value.ToString(), out decimal num))
                {
                    e.Value = num.ToString("N2", _numCulture);
                    e.FormattingApplied = true;
                }
                e.CellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }

            // Cột Mô tả (MPR) — chỉ xem, nền xanh nhạt, chữ nghiêng khi trống
            if (colName == "MPR_Desc")
            {
                e.CellStyle.BackColor = Color.FromArgb(240, 248, 255);
                string val = e.Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(val))
                {
                    e.Value = "(không có)";
                    e.CellStyle.ForeColor = Color.FromArgb(180, 180, 180);
                    e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Italic);
                    e.FormattingApplied = true;
                }
                else
                {
                    e.CellStyle.ForeColor = Color.FromArgb(0, 100, 180);
                }
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
        {
            _isSearching = true;
            int curId = 0;
            if (cboSupplier.SelectedValue != null)
                int.TryParse(cboSupplier.SelectedValue.ToString(), out curId);
            string curText = cboSupplier.Text;
            cboSupplier.DataSource = null;
            cboSupplier.DataSource = dt;
            cboSupplier.DisplayMember = "Name";
            cboSupplier.ValueMember = "ID";
            // Restore: nếu đang tìm kiếm thì giữ text, nếu đã chọn thì restore ID
            if (curId > 0)
                cboSupplier.SelectedValue = curId;
            else
                cboSupplier.Text = curText;
            _isSearching = false;
        }


        private void CboSupplier_TextChanged(object sender, EventArgs e)
        {
            if (_isSearching) return;
            // Chỉ xử lý khi combobox đang có focus (user đang gõ trực tiếp)
            if (!cboSupplier.Focused) return;
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

        //private void BtnExport_Click(object sender, EventArgs e)
        //{
        //    if (_selectedPO_ID == 0) { MessageBox.Show("Vui lòng chọn PO cần xuất!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        //    try
        //    {
        //        var po = _poList.Find(p => p.PO_ID == _selectedPO_ID);
        //        var details = _service.GetDetails(_selectedPO_ID); if (po == null) return;
        //        var suppliers = new SupplierService().GetAll();
        //        var supplier = suppliers.Find(s => s.Supplier_ID == Convert.ToInt32(cboSupplier.SelectedValue?.ToString() ?? "0"));
        //        var projects = new ProjectService().GetAll();
        //        var project = projects.Find(p => p.WorkorderNo == po.WorkorderNo);
        //        string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "po_template.xlsx");
        //        if (!File.Exists(templatePath)) { MessageBox.Show($"Lỗi: Không tìm thấy file template!\nĐường dẫn dự kiến: {templatePath}", "Lỗi hệ thống", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
        //        var saveDialog = new SaveFileDialog { Title = "Lưu file PO", Filter = "Excel Files|*.xlsx", FileName = $"PO_{po.PONo}_{DateTime.Now:ddMMyyyy}", InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) };
        //        if (saveDialog.ShowDialog() != DialogResult.OK) return;
        //        File.Copy(templatePath, saveDialog.FileName, true);
        //        OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        //        using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(saveDialog.FileName)))
        //        {
        //            var ws = package.Workbook.Worksheets[0];
        //            ReplaceCell(ws, "<<PROJECT_NAME>>", project?.ProjectName ?? po.Project_Name ?? ""); ReplaceCell(ws, "<<WO-NO>>", po.WorkorderNo ?? ""); ReplaceCell(ws, "<<REV.NUM>>", po.Revise.ToString() ?? "0");
        //            ReplaceCell(ws, "<<DATE>>", po.PO_Date.HasValue ? po.PO_Date.Value.ToString("dd/MM/yyyy") : DateTime.Today.ToString("dd/MM/yyyy")); ReplaceCell(ws, "<<MPR-NO>>", po.MPR_No ?? ""); ReplaceCell(ws, "<<PO-NO>>", po.PONo ?? "");
        //            string supplierInfo = supplier != null ? $"{supplier.Company_Name}\nCert: {supplier.Cert ?? ""}\nEmail: {supplier.Email}" : "";
        //            ReplaceCell(ws, "<<SUPPLIER-INFO>>", supplierInfo);
        //            int startRow = 8; int detailCount = details.Count;
        //            if (detailCount > 1) ws.InsertRow(startRow + 1, detailCount - 1);
        //            for (int i = 0; i < detailCount; i++)
        //            {
        //                var d = details[i];
        //                int row = startRow + i;
        //                decimal q = d.Qty_Per_Sheet; decimal wk = d.Weight_kg; decimal realPrice = d.Price;
        //                string rem = d.Remarks ?? "";
        //                if (rem.Contains("[CALC:KG]"))
        //                {
        //                    rem = rem.Replace("[CALC:KG]", "").Trim();
        //                    if (wk > 0 && q > 0) realPrice = (d.Price * q) / wk;
        //                }
        //                else if (rem.Contains("[CALC:SL]")) rem = rem.Replace("[CALC:SL]", "").Trim();
        //                ws.Cells[row, 1].Value = i + 1; ws.Cells[row, 2].Value = d.Item_Name ?? ""; ws.Cells[row, 3].Value = d.Material ?? "";
        //                ws.Cells[row, 4].Value = d.Asize; ws.Cells[row, 5].Value = d.Bsize; ws.Cells[row, 6].Value = d.Csize;
        //                ws.Cells[row, 7].Value = d.Qty_Per_Sheet;
        //                ws.Cells[row, 8].Value = d.UNIT ?? ""; ws.Cells[row, 9].Value = d.Weight_kg;
        //                ws.Cells[row, 10].Value = d.MPSNo ?? ""; ws.Cells[row, 11].Value = d.RequestDay;
        //                ws.Cells[row, 12].Value = "Kho DLHI";
        //                ws.Cells[row, 13].Value = Math.Round(realPrice, 0); ws.Cells[row, 14].Value = d.Amount; ws.Cells[row, 16].Value = rem;
        //                if (i > 0)
        //                {
        //                    for (int col = 1; col <= 16; col++)
        //                    {
        //                        ws.Cells[row, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        //                        ws.Cells[row, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        //                        ws.Cells[row, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; 
        //                        ws.Cells[row, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        //                        ws.Cells[row, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; 
        //                        ws.Cells[row, col].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
        //                        ws.Cells[row, col].Style.Font.Name = "Arial"; ws.Cells[row, col].Style.Font.Size = 9;
        //                    }
        //                    ws.Cells[row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
        //                }
        //            }
        //            int subTotalRow = startRow + detailCount;
        //            int vatRow = subTotalRow + 1;
        //            ws.Cells[subTotalRow, 3].Value = "SUB-TOTAL"; ws.Cells[subTotalRow, 9].Value = details.Sum(d => (double)d.Weight_kg);
        //            ws.Cells[subTotalRow, 14].Formula = $"=SUM(N{startRow}:N{startRow + detailCount - 1})";
        //            ws.Cells[vatRow, 3].Value = "Final Price Requested (Included 10% VAT)";
        //            ws.Cells[vatRow, 14].Formula = $"=N{subTotalRow}*1.1";
        //            for (int r = 1; r <= ws.Dimension.End.Row; r++) for (int c = 1; c <= ws.Dimension.End.Column; c++) if (ws.Cells[r, c].Value?.ToString() == "<<DATE>>") ws.Cells[r, c].Value = DateTime.Today.ToString("dd/MM/yyyy");
        //            package.Save();
        //        }
        //        var result = MessageBox.Show($"✅ Xuất Excel thành công!\nFile: {saveDialog.FileName}\n\nBạn có muốn mở file không?", "Thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
        //        if (result == DialogResult.Yes) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = saveDialog.FileName, UseShellExecute = true });
        //    }
        //    catch (Exception ex) { MessageBox.Show("Lỗi xuất Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        //}

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("PO", "Xuất Excel", "Xuất Excel")) return;
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
                if (!File.Exists(templatePath)) { MessageBox.Show($"Lỗi: Không tìm thấy template!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var saveDialog = new SaveFileDialog { Title = "Lưu file PO", Filter = "Excel Files|*.xlsx", FileName = $"PO_{po.PONo}_{DateTime.Now:ddMMyyyy}" };
                if (saveDialog.ShowDialog() != DialogResult.OK) return;

                File.Copy(templatePath, saveDialog.FileName, true);
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(saveDialog.FileName)))
                {
                    var ws = package.Workbook.Worksheets[0];

                    // 1. Thay thế Header Tags
                    ReplaceCell(ws, "<<PROJECT_NAME>>", project?.ProjectName ?? po.Project_Name ?? "");
                    ReplaceCell(ws, "<<WO-NO>>", po.WorkorderNo ?? "");
                    ReplaceCell(ws, "<<REV.NUM>>", po.Revise.ToString() ?? "0");
                    ReplaceCell(ws, "<<DATE>>", po.PO_Date.HasValue ? po.PO_Date.Value.ToString("dd/MM/yyyy") : DateTime.Today.ToString("dd/MM/yyyy"));
                    ReplaceCell(ws, "<<MPR-NO>>", po.MPR_No ?? "");
                    ReplaceCell(ws, "<<PO-NO>>", po.PONo ?? "");
                    string supplierInfo = supplier != null ? $"{supplier.Company_Name}\nCert: {supplier.Cert ?? ""}\nEmail: {supplier.Email}" : "";
                    ReplaceCell(ws, "<<SUPPLIER-INFO>>", supplierInfo);

                    // Payment Term — set trực tiếp vào O5 (merged O5:Q5)
                    ws.Cells[5, 15].Value = !string.IsNullOrEmpty(po.Payment_Term)
                                             ? po.Payment_Term
                                             : "Within 7 days after delivery";
                    // Expect day = ngày hiện tại + 7 — set K8 trước InsertRow
                    ws.Cells[8, 11].Value = DateTime.Today.AddDays(7).ToString("dd/MM/yyyy");

                    int startRow = 8;
                    int detailCount = details.Count;

                    // 2. XỬ LÝ CHÈN DÒNG VÀ GIỮ ĐỊNH DẠNG (MERGED CELLS)
                    if (detailCount > 1)
                    {
                        // Chèn thêm dòng
                        ws.InsertRow(startRow + 1, detailCount - 1);

                        // Copy định dạng và Merged Cells từ dòng 8 xuống các dòng mới chèn
                        for (int i = 1; i < detailCount; i++)
                        {
                            // Copy từ dòng startRow sang dòng startRow + i
                            ws.Cells[startRow, 1, startRow, 16].Copy(ws.Cells[startRow + i, 1]);
                        }
                    }

                    // 3. ĐIỀN DỮ LIỆU VÀO CÁC DÒNG
                    for (int i = 0; i < detailCount; i++)
                    {
                        var d = details[i];
                        int row = startRow + i;

                        decimal q = d.Qty_Per_Sheet;
                        decimal wk = d.Weight_kg;
                        decimal realPrice = d.Price;
                        string rem = d.Remarks ?? "";

                        if (rem.Contains("[CALC:KG]"))
                        {
                            rem = rem.Replace("[CALC:KG]", "").Trim();
                            if (wk > 0 && q > 0) realPrice = (d.Price * q) / wk;
                        }

                        // Gán giá trị
                        ws.Cells[row, 1].Value = i + 1;
                        ws.Cells[row, 2].Value = d.Item_Name ?? "";
                        ws.Cells[row, 3].Value = d.Material ?? "";

                        // Lưu ý: Trong template của bạn, cột Size (D, E, F) có thể gộp hoặc chia nhỏ A, B, C
                        ws.Cells[row, 4].Value = d.Asize;
                        ws.Cells[row, 5].Value = d.Bsize;
                        ws.Cells[row, 6].Value = d.Csize;

                        ws.Cells[row, 7].Value = d.Qty_Per_Sheet;
                        ws.Cells[row, 8].Value = d.UNIT ?? "";
                        ws.Cells[row, 9].Value = d.Weight_kg;
                        ws.Cells[row, 10].Value = d.MPSNo ?? "";
                        ws.Cells[row, 11].Value = po.Expected_Delivery;
                        ws.Cells[row, 12].Value = "Kho DLHI";
                        //ws.Cells[row, 13].Value = Math.Round(realPrice, 0);
                        ws.Cells[row, 13].Value = realPrice;
                        ws.Cells[row, 14].Value = d.Amount;
                        //ws.Cells[row, 16].Value = rem;

                        // --- CỦNG CỐ ĐỊNH DẠNG MERGE CHO REMARKS ---
                        // Nếu trong template cột Remarks gộp từ cột 16 (P) và 17 (Q)
                        if (!ws.Cells[row, 16].Merge)
                        {
                            ws.Cells[row, 16, row, 17].Merge = true;
                        }

                        // Thiết lập Style cho hàng
                        using (var range = ws.Cells[row, 1, row, 17])
                        {
                            range.Style.Font.Name = "Times New Roman";
                            range.Style.Font.Size = 11;
                            range.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                            // Kẻ khung
                            range.Style.Border.Top.Style = range.Style.Border.Bottom.Style =
                            range.Style.Border.Left.Style = range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        }
                        ws.Cells[row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        ws.Cells[row, 16].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                        ws.Cells[row, 11].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        ws.Cells[row, 11].Style.Numberformat.Format = "dd/MM/yyyy";

                        ws.Cells[row, 13].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        ws.Cells[row, 13].Style.Numberformat.Format = "#,##0.00";

                        ws.Cells[row, 14].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        ws.Cells[row, 14].Style.Numberformat.Format = "#,##0.00";
                    }

                    //// 4. TÍNH TỔNG (Dòng Sum sẽ bị đẩy xuống do lệnh InsertRow ở trên)
                    //int subTotalRow = startRow + detailCount;
                    //int vatRow = subTotalRow + 1;


                    //ws.Cells[subTotalRow, 9].Formula = $"=SUM(I{startRow}:I{startRow + detailCount - 1})";
                    //ws.Cells[subTotalRow, 14].Formula = $"=SUM(N{startRow}:N{startRow + detailCount - 1})";

                    //ws.Cells[vatRow, 14].Formula = $"=N{subTotalRow}*1.1";

                    // 1. Tính tổng tiền sau thuế từ DataTable
                    decimal totalAfterVAT = 0;
                    foreach (var dr in details)
                    {
                        // Sử dụng hàm SafeParse chúng ta đã xây dựng để tránh lỗi định dạng số
                        decimal amount = dr.Amount; // Thành tiền chưa thuế
                        decimal vatPercent = dr.VAT; // Thuế suất (ví dụ: 8 hoặc 10)

                        // Tính thành tiền sau thuế của từng dòng và cộng dồn
                        totalAfterVAT += amount * (1 + vatPercent / 100);
                    }

                    // 2. Xác định vị trí các dòng tổng kết
                    int subTotalRow = startRow + detailCount;
                    int vatRow = subTotalRow + 1;

                    // Gán nhãn và tính tổng chưa thuế (Sub-total) bằng Formula Excel
                    ws.Cells[subTotalRow, 3].Value = "SUB-TOTAL";
                    ws.Cells[subTotalRow, 14].Formula = $"=SUM(N{startRow}:N{startRow + detailCount - 1})";

                    // Gán nhãn cho dòng tổng thanh toán (đã bao gồm các mức thuế khác nhau)
                    ws.Cells[vatRow, 3].Value = "Final Price Requested (Included VAT)";

                    // 3. CẬP NHẬT: Gán giá trị tổng sau thuế đã tính toán vào ô kết quả
                    // Chúng ta gán giá trị số trực tiếp và định dạng hiển thị cho chuyên nghiệp
                    ws.Cells[vatRow, 14].Value = totalAfterVAT;
                    ws.Cells[vatRow, 14].Style.Numberformat.Format = "#,##0.##"; // Định dạng dấu phân cách hàng nghìn
                    ws.Cells[vatRow, 14].Style.Font.Bold = true;

                    package.Save();
                }

                if (MessageBox.Show("✅ Xuất Excel thành công! Bạn có muốn mở file?", "Thành công", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = saveDialog.FileName, UseShellExecute = true });
            }
            catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message); }
        }

        //private void BtnExport_Click(object sender, EventArgs e)
        //{
        //    if (_selectedPO_ID == 0) { MessageBox.Show("Vui lòng chọn PO cần xuất!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        //    try
        //    {
        //        var po = _poList.Find(p => p.PO_ID == _selectedPO_ID);
        //        var details = _service.GetDetails(_selectedPO_ID); if (po == null) return;
        //        var suppliers = new SupplierService().GetAll();
        //        var supplier = suppliers.Find(s => s.Supplier_ID == Convert.ToInt32(cboSupplier.SelectedValue?.ToString() ?? "0"));
        //        var projects = new ProjectService().GetAll();
        //        var project = projects.Find(p => p.WorkorderNo == po.WorkorderNo);
        //        string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "po_template.xlsx");

        //        if (!File.Exists(templatePath)) { MessageBox.Show($"Lỗi: Không tìm thấy file template!\nĐường dẫn dự kiến: {templatePath}", "Lỗi hệ thống", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

        //        var saveDialog = new SaveFileDialog { Title = "Lưu file PO", Filter = "Excel Files|*.xlsx", FileName = $"PO_{po.PONo}_{DateTime.Now:ddMMyyyy}", InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) };
        //        if (saveDialog.ShowDialog() != DialogResult.OK) return;

        //        File.Copy(templatePath, saveDialog.FileName, true);
        //        OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

        //        using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(saveDialog.FileName)))
        //        {
        //            var ws = package.Workbook.Worksheets[0];

        //            // Thay thế Header
        //            ReplaceCell(ws, "<<PROJECT_NAME>>", project?.ProjectName ?? po.Project_Name ?? "");
        //            ReplaceCell(ws, "<<WO-NO>>", po.WorkorderNo ?? "");
        //            ReplaceCell(ws, "<<REV.NUM>>", po.Revise.ToString() ?? "0");
        //            ReplaceCell(ws, "<<DATE>>", po.PO_Date.HasValue ? po.PO_Date.Value.ToString("dd/MM/yyyy") : DateTime.Today.ToString("dd/MM/yyyy"));
        //            ReplaceCell(ws, "<<MPR-NO>>", po.MPR_No ?? "");
        //            ReplaceCell(ws, "<<PO-NO>>", po.PONo ?? "");

        //            string supplierInfo = supplier != null ? $"{supplier.Company_Name}\nCert: {supplier.Cert ?? ""}\nEmail: {supplier.Email}" : "";
        //            ReplaceCell(ws, "<<SUPPLIER-INFO>>", supplierInfo);

        //            int startRow = 8;
        //            int detailCount = details.Count;

        //            // Chèn dòng nếu nhiều hơn 1 item
        //            if (detailCount > 1) ws.InsertRow(startRow + 1, detailCount - 1);

        //            for (int i = 0; i < detailCount; i++)
        //            {
        //                var d = details[i];
        //                int row = startRow + i;

        //                decimal q = d.Qty_Per_Sheet;
        //                decimal wk = d.Weight_kg;
        //                decimal realPrice = d.Price;
        //                string rem = d.Remarks ?? "";

        //                if (rem.Contains("[CALC:KG]"))
        //                {
        //                    rem = rem.Replace("[CALC:KG]", "").Trim();
        //                    if (wk > 0 && q > 0) realPrice = (d.Price * q) / wk;
        //                }
        //                else if (rem.Contains("[CALC:SL]")) rem = rem.Replace("[CALC:SL]", "").Trim();

        //                // Gán giá trị vào các ô
        //                ws.Cells[row, 1].Value = i + 1;
        //                ws.Cells[row, 2].Value = d.Item_Name ?? "";
        //                ws.Cells[row, 3].Value = d.Material ?? "";
        //                ws.Cells[row, 4].Value = d.Asize;
        //                ws.Cells[row, 5].Value = d.Bsize;
        //                ws.Cells[row, 6].Value = d.Csize;
        //                ws.Cells[row, 7].Value = d.Qty_Per_Sheet;
        //                ws.Cells[row, 8].Value = d.UNIT ?? "";
        //                ws.Cells[row, 9].Value = d.Weight_kg;
        //                ws.Cells[row, 10].Value = d.MPSNo ?? "";
        //                ws.Cells[row, 11].Value = d.RequestDay;
        //                ws.Cells[row, 12].Value = "Kho DLHI";
        //                ws.Cells[row, 13].Value = Math.Round(realPrice, 0);
        //                ws.Cells[row, 14].Value = d.Amount;
        //                ws.Cells[row, 16].Value = rem;

        //                // --- ĐỊNH DẠNG TOÀN BỘ CÁC DÒNG DỮ LIỆU ---
        //                using (var range = ws.Cells[row, 1, row, 16])
        //                {
        //                    range.Style.Font.Name = "Times New Roman";
        //                    range.Style.Font.Size = 11;
        //                    range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //                    range.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

        //                    // Kẻ khung
        //                    range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        //                    range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        //                    range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        //                    range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        //                }

        //                // Riêng cột Tên vật tư (Cột 2) cho phép căn trái để dễ đọc hơn (Nếu bạn vẫn muốn giữa hết thì xóa 2 dòng dưới)
        //                ws.Cells[row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
        //            }

        //            int subTotalRow = startRow + detailCount;
        //            int vatRow = subTotalRow + 1;

        //            ws.Cells[subTotalRow, 3].Value = "SUB-TOTAL";
        //            ws.Cells[subTotalRow, 9].Value = details.Sum(d => (double)d.Weight_kg);
        //            ws.Cells[subTotalRow, 14].Formula = $"=SUM(N{startRow}:N{startRow + detailCount - 1})";

        //            ws.Cells[vatRow, 3].Value = "Final Price Requested (Included 10% VAT)";
        //            ws.Cells[vatRow, 14].Formula = $"=N{subTotalRow}*1.1";

        //            // Quét lại toàn bộ file để thay thế các tag ngày tháng còn sót
        //            for (int r = 1; r <= ws.Dimension.End.Row; r++)
        //                for (int c = 1; c <= ws.Dimension.End.Column; c++)
        //                    if (ws.Cells[r, c].Value?.ToString() == "<<DATE>>")
        //                        ws.Cells[r, c].Value = DateTime.Today.ToString("dd/MM/yyyy");

        //            package.Save();
        //        }

        //        var result = MessageBox.Show($"✅ Xuất Excel thành công!\nFile: {saveDialog.FileName}\n\nBạn có muốn mở file không?", "Thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
        //        if (result == DialogResult.Yes) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = saveDialog.FileName, UseShellExecute = true });
        //    }
        //    catch (Exception ex) { MessageBox.Show("Lỗi xuất Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        //}

        private void ReplaceCell(OfficeOpenXml.ExcelWorksheet ws, string placeholder, string value)
        { for (int r = 1; r <= ws.Dimension.End.Row; r++) for (int c = 1; c <= ws.Dimension.End.Column; c++) if (ws.Cells[r, c].Value?.ToString() == placeholder) ws.Cells[r, c].Value = value; }

        private void BuildDetailColumns()
        {
            dgvDetails.Columns.Clear();
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "DeliveryLocation", HeaderText = "Nơi giao" });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Item_No", HeaderText = "STT", ReadOnly = true });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Item_Name", HeaderText = "Tên hàng" });
            // Cột Description từ MPR — chỉ hiển thị, không lưu
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "MPR_Desc",
                HeaderText = "Mô tả (MPR)",
                ReadOnly = true,
                DefaultCellStyle = { ForeColor = Color.FromArgb(0, 120, 212), BackColor = Color.FromArgb(240, 248, 255) }
            });
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
            colVAT.Items.AddRange("10", "8", "0");
            dgvDetails.Columns.Add(colVAT);

            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "SubAmount", HeaderText = "TT trước thuế", ReadOnly = true });
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
            panelTop.Width = w; panelHeader.Width = w;
            panelHeader.Top = panelTop.Bottom + 10;
            panelDetail.Top = panelHeader.Bottom + 10;
            panelDetail.Height = h - panelDetail.Top - 10;

            // panelMPRFiles: 280px cố định bên phải
            var panelMPRFiles = this.Controls.OfType<Panel>()
                .FirstOrDefault(p => p.Controls.OfType<DataGridView>()
                    .Any(d => d.Name == "dgvMPRFiles" || (dgvMPRFiles != null && d == dgvMPRFiles)));
            int mprW = 280;
            int detailW = w - mprW - 10;
            panelDetail.Width = detailW;
            if (panelMPRFiles != null)
            {
                panelMPRFiles.Location = new Point(panelDetail.Right + 10, panelDetail.Top);
                panelMPRFiles.Width = mprW;
                panelMPRFiles.Height = panelDetail.Height;
            }

            dgvPO.Width = panelTop.Width - 20;
            dgvDetails.Width = panelDetail.Width - 20;
            dgvDetails.Height = panelDetail.Height - 80;

        }

        // Load danh sách file từ MPR_Link — tìm theo MPR No tại "Thông tin đơn đặt hàng"
        private void LoadMPRFiles(string workorderNo = "")
        {
            if (dgvMPRFiles == null) return;
            dgvMPRFiles.Rows.Clear();
            try
            {
                // Lấy MPR No từ txtMPRNo trong bảng "Thông tin đơn đặt hàng"
                string mprNo = txtMPRNo?.Text?.Trim() ?? "";
                if (string.IsNullOrEmpty(mprNo)) return;

                // Tìm MPR_Header theo MPR_No → lấy Project_Name
                string projectName = "";
                using (var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var cmd = new Microsoft.Data.SqlClient.SqlCommand(
                        "SELECT TOP 1 Project_Name FROM MPR_Header WHERE MPR_No = @mprNo", conn);
                    cmd.Parameters.AddWithValue("@mprNo", mprNo);
                    var result = cmd.ExecuteScalar();
                    projectName = result?.ToString() ?? "";
                }

                if (string.IsNullOrEmpty(projectName)) return;

                // Tìm ProjectInfo theo Project_Name → lấy MPR_Link
                var proj = new MPR_Managerment.Services.ProjectService().GetAll()
                    .Find(p =>
                        !string.IsNullOrEmpty(p.ProjectName) &&
                        (p.ProjectName.Equals(projectName, StringComparison.OrdinalIgnoreCase) ||
                         p.ProjectName.IndexOf(projectName, StringComparison.OrdinalIgnoreCase) >= 0));

                string mprLink = proj?.MPR_Link?.Trim() ?? "";
                if (string.IsNullOrEmpty(mprLink) || !System.IO.Directory.Exists(mprLink))
                    return;

                foreach (var f in System.IO.Directory.GetFiles(mprLink).OrderBy(x => x))
                {
                    int idx = dgvMPRFiles.Rows.Add();
                    dgvMPRFiles.Rows[idx].Cells["FileName"].Value = System.IO.Path.GetFileName(f);
                    dgvMPRFiles.Rows[idx].Cells["FullPath"].Value = f;
                }
            }
            catch { }
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
        // BIND PO GRID — sắp xếp theo Ngày tạo mới nhất
        // =========================================================================
        private void BindPOGrid(List<POHead> list)
        {
            var suppliers = new SupplierService().GetAll();
            // Sắp xếp theo Created_Date giảm dần (mới nhất lên đầu)
            var sorted = list
                .OrderByDescending(h => h.Created_Date ?? DateTime.MinValue)
                .ThenBy(h => h.PONo, StringComparer.OrdinalIgnoreCase)
                .ToList();
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
                    Tong_Tien = h.Total_Amount.ToString("N2", _numCulture),
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

                // Load Description từ MPR_Details cho tất cả dòng có MPR_Detail_ID
                var mprDescMap = new Dictionary<int, string>();
                try
                {
                    var mprDetailIds = _details
                        .Where(d => d.MPR_Detail_ID.HasValue && d.MPR_Detail_ID.Value > 0)
                        .Select(d => d.MPR_Detail_ID.Value)
                        .Distinct().ToList();

                    if (mprDetailIds.Count > 0)
                    {
                        string ids = string.Join(",", mprDetailIds);
                        using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                        conn.Open();
                        var cmd = new Microsoft.Data.SqlClient.SqlCommand(
                            $"SELECT Detail_ID, ISNULL(Description,'') AS Description FROM MPR_Details WHERE Detail_ID IN ({ids})", conn);
                        using var reader = cmd.ExecuteReader();
                        while (reader.Read())
                            mprDescMap[Convert.ToInt32(reader["Detail_ID"])] = reader["Description"]?.ToString() ?? "";
                    }
                }
                catch { /* Nếu lỗi query Description thì bỏ qua, vẫn load bình thường */ }

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

                    // Fill Description từ MPR (chỉ hiển thị)
                    string mprDesc = "";
                    if (d.MPR_Detail_ID.HasValue && d.MPR_Detail_ID.Value > 0)
                        mprDescMap.TryGetValue(d.MPR_Detail_ID.Value, out mprDesc);
                    row.Cells["MPR_Desc"].Value = mprDesc ?? "";

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
            decimal totalQty = 0, totalKg = 0, totalAmount = 0;

            foreach (DataGridViewRow row in dgvDetails.Rows)
            {
                if (row.IsNewRow || row.Tag?.ToString() == "TOTAL") continue;
                decimal qty = ParseDecimalRaw(row.Cells["Qty"].Value?.ToString() ?? "");
                decimal weight = ParseDecimalRaw(row.Cells["Weight"].Value?.ToString() ?? "");
                decimal price = ParseDecimalRaw(row.Cells["Price"].Value?.ToString() ?? "");
                decimal.TryParse(row.Cells["VAT"].Value?.ToString(), out decimal vat);
                decimal amount = ParseDecimalRaw(row.Cells["Amount"].Value?.ToString() ?? "");
                string calcMethod = row.Cells["Calc_Method"].Value?.ToString() ?? "Theo KG";
                decimal baseValue = (calcMethod == "Theo KG") ? weight : qty;
                decimal rowSubTotal = baseValue * price;
                decimal rowTotalAmount = rowSubTotal * (1 + vat / 100);
                subTotal += rowSubTotal; total += rowTotalAmount;
                totalQty += qty;
                totalKg += weight;
                totalAmount += amount > 0 ? amount : rowTotalAmount;
            }

            if (lblSubTotal != null) lblSubTotal.Text = "Truoc VAT: " + subTotal.ToString("N2", _numCulture) + " VND";
            if (lblTotal != null) lblTotal.Text = "Sau VAT: " + total.ToString("N2", _numCulture) + " VND";

            // ── Cập nhật dòng Total ở cuối dgvDetails ──
            // Xóa dòng total cũ nếu có
            for (int i = dgvDetails.Rows.Count - 1; i >= 0; i--)
                if (dgvDetails.Rows[i].Tag?.ToString() == "TOTAL")
                { dgvDetails.Rows.RemoveAt(i); break; }

            if (dgvDetails.Rows.Count > 0)
            {
                int idx = dgvDetails.Rows.Add();
                var tr = dgvDetails.Rows[idx];
                tr.Tag = "TOTAL";
                tr.ReadOnly = true;

                // Đặt giá trị tổng
                if (dgvDetails.Columns.Contains("Item_Name")) tr.Cells["Item_Name"].Value = "TỔNG CỘNG";
                if (dgvDetails.Columns.Contains("Qty")) tr.Cells["Qty"].Value = totalQty > 0 ? (object)totalQty : "";
                if (dgvDetails.Columns.Contains("Weight")) tr.Cells["Weight"].Value = totalKg > 0 ? (object)Math.Round(totalKg, 2) : "";
                if (dgvDetails.Columns.Contains("Amount")) tr.Cells["Amount"].Value = totalAmount > 0 ? (object)totalAmount : "";

                // Style dòng total
                tr.DefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
                tr.DefaultCellStyle.ForeColor = Color.White;
                tr.DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                tr.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
        }

        private void DgvDetails_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V)
            {
                PasteFromExcel();
                e.Handled = true;
            }
        }

        // Cac cot so: parse voi ParseDecimalRaw
        private static readonly System.Collections.Generic.HashSet<string> _numericCols
            = new System.Collections.Generic.HashSet<string>
              { "Qty", "Weight", "Price", "Received" };

        private void PasteFromExcel()
        {
            try
            {
                string copiedData = Clipboard.GetText();
                if (string.IsNullOrEmpty(copiedData)) return;

                // Tach thanh cac dong (ho tro CR, LF, CRLF)
                string[] rows = copiedData.Split(
                    new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

                // Bo dong cuoi neu rong (Excel thuong them \r\n cuoi)
                if (rows.Length > 0 && string.IsNullOrEmpty(rows[rows.Length - 1]))
                    rows = rows.Take(rows.Length - 1).ToArray();

                if (rows.Length == 0) return;

                int startRow = dgvDetails.CurrentCell?.RowIndex ?? 0;
                int startCol = dgvDetails.CurrentCell?.ColumnIndex ?? 0;

                // Build danh sach cot hien (visible, editable, khong hidden)
                // De paste dung cot tuong ung voi vi tri clipboard
                var pasteableCols = new System.Collections.Generic.List<int>();
                for (int c = startCol; c < dgvDetails.Columns.Count; c++)
                {
                    var col = dgvDetails.Columns[c];
                    if (!col.Visible) continue;
                    if (col.Name == "PO_Detail_ID" || col.Name == "MPR_Detail_ID") continue;
                    pasteableCols.Add(c);
                }

                dgvDetails.SuspendLayout();
                for (int ri = 0; ri < rows.Length; ri++)
                {
                    string[] cells = rows[ri].Split('\t');
                    int curRow = startRow + ri;

                    // Them dong moi neu can
                    if (curRow >= dgvDetails.Rows.Count)
                    {
                        int newIdx = dgvDetails.Rows.Add();
                        var r = dgvDetails.Rows[newIdx];
                        r.Cells["Item_No"].Value = newIdx + 1;
                        r.Cells["PO_Detail_ID"].Value = 0;
                        r.Cells["UNIT"].Value = "PCS";
                        r.Cells["Qty"].Value = 0m;
                        r.Cells["Weight"].Value = 0m;
                        r.Cells["Price"].Value = 0m;
                        r.Cells["VAT"].Value = "10";
                        r.Cells["Amount"].Value = 0m;
                        r.Cells["Received"].Value = 0m;
                        r.Cells["Calc_Method"].Value = "Theo KG";
                        r.Cells["Ordered_PO"].Value = "";
                    }

                    var dgvRow = dgvDetails.Rows[curRow];
                    if (dgvRow.IsNewRow || dgvRow.Tag?.ToString() == "TOTAL") continue;

                    for (int ci = 0; ci < cells.Length && ci < pasteableCols.Count; ci++)
                    {
                        int colIdx = pasteableCols[ci];
                        var col = dgvDetails.Columns[colIdx];
                        string val = cells[ci].Trim();

                        if (col is DataGridViewComboBoxColumn cboCol)
                        {
                            // VAT: map gia tri hop le
                            if (col.Name == "VAT")
                            {
                                string v = val.Replace("%", "").Trim();
                                dgvRow.Cells[colIdx].Value =
                                    v == "8" ? "8" : v == "0" ? "0" : "10";
                            }
                            else if (col.Name == "Calc_Method")
                            {
                                dgvRow.Cells[colIdx].Value =
                                    val.Contains("SL") ? "Theo SL" : "Theo KG";
                            }
                        }
                        else if (col.ReadOnly)
                        {
                            // Bo qua cot ReadOnly (Amount, SubAmount, Ordered_PO...)
                            continue;
                        }
                        else if (_numericCols.Contains(col.Name))
                        {
                            // Parse so voi ParseDecimalRaw
                            dgvRow.Cells[colIdx].Value = ParseDecimalRaw(val);
                        }
                        else
                        {
                            dgvRow.Cells[colIdx].Value = val;
                        }
                    }
                    RecalculateAmount(curRow);
                }
                dgvDetails.ResumeLayout();
                AutoAdjustColumnWidths();
                UpdateTotal();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Loi khi dan du lieu: " + ex.Message,
                    "Loi Paste", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RecalculateAmount(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= dgvDetails.Rows.Count) return;
            var row = dgvDetails.Rows[rowIndex];
            if (row.IsNewRow || row.Tag?.ToString() == "TOTAL") return;
            // Dung ParseDecimalRaw de xu ly moi dinh dang nhap/paste
            decimal qty = ParseDecimalRaw(row.Cells["Qty"].Value?.ToString() ?? "");
            decimal weight = ParseDecimalRaw(row.Cells["Weight"].Value?.ToString() ?? "");
            decimal price = ParseDecimalRaw(row.Cells["Price"].Value?.ToString() ?? "");
            decimal.TryParse(row.Cells["VAT"].Value?.ToString(), out decimal vat);
            string calcMethod = row.Cells["Calc_Method"].Value?.ToString() ?? "Theo KG";
            decimal baseValue = (calcMethod == "Theo KG") ? weight : qty;
            decimal subAmt = baseValue * price;
            if (dgvDetails.Columns.Contains("SubAmount")) row.Cells["SubAmount"].Value = subAmt;
            row.Cells["Amount"].Value = subAmt * (1 + vat / 100);
            UpdateTotal();
        }

        private decimal ParseDecimal(object value)
        {
            if (value == null) return 0;
            string input = value.ToString().Trim();
            if (string.IsNullOrEmpty(input)) return 0;

            // Dung ParseDecimalRaw de xu ly thong nhat
            return ParseDecimalRaw(input);
        }

        // Intercept gia tri khi nguoi dung nhap/paste — parse dung vi-VN
        private void DgvDetails_CellParsing(object sender, DataGridViewCellParsingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            string colName = dgvDetails.Columns[e.ColumnIndex].Name;

            // Chi xu ly cac cot so
            if (colName != "Price" && colName != "Qty" && colName != "Weight"
                && colName != "Amount" && colName != "SubAmount" && colName != "Received") return;

            string raw = e.Value?.ToString() ?? "";
            if (string.IsNullOrWhiteSpace(raw)) return;

            decimal parsed = ParseDecimalRaw(raw);
            // Luu gia tri decimal thuc vao cell, tranh WinForms parse lai sai
            e.Value = parsed;
            e.ParsingApplied = true;
            // Dong thoi ghi thang vao cell de dam bao
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
                dgvDetails.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = parsed;
        }

        // Parse so tu chuoi bat ky: xu ly ca vi-VN (. ngan , thap phan)
        // lan InvariantCulture (, ngan . thap phan)
        private decimal ParseDecimalRaw(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return 0;
            raw = raw.Trim();

            // Co ca "." va ","
            if (raw.Contains(".") && raw.Contains(","))
            {
                if (raw.IndexOf(".") < raw.IndexOf(","))
                {
                    // "1.234,56" -> vi-VN -> bo . -> "1234,56"
                    raw = raw.Replace(".", "");
                }
                else
                {
                    // "1,234.56" -> InvariantCulture -> bo , -> "1234.56" -> doi . thanh ,
                    raw = raw.Replace(",", "").Replace(".", ",");
                }
            }
            else if (raw.Contains(".") && !raw.Contains(","))
            {
                // Chi co "."
                var parts = raw.Split('.');
                // Neu tat ca phan sau dau . deu co 3 chu so -> day la ngan separator
                bool allThousand = parts.Length > 1 &&
                    parts.Skip(1).All(p => p.Length == 3);
                if (allThousand)
                    raw = raw.Replace(".", "");        // bo ngan
                else
                    raw = raw.Replace(".", ",");        // doi . -> , (thap phan vi-VN)
            }
            // Chi co "," hoac so nguyen: vi-VN hieu "," la thap phan -> giu nguyen

            decimal.TryParse(raw,
                System.Globalization.NumberStyles.Number,
                _numCulture, out decimal result);
            return result;
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
            LoadMPRFiles();
            txtPrepared.Text = h.Prepared; txtReviewed.Text = h.Reviewed;
            txtAgreement.Text = h.Agreement; txtApproved.Text = h.Approved;
            txtNotes.Text = h.Notes; nudRevise.Value = h.Revise;
            if (h.PO_Date.HasValue) dtpPODate.Value = h.PO_Date.Value;
            var ptIdx = cboPaymentTerm.Items.IndexOf(h.Payment_Term ?? "");
            cboPaymentTerm.SelectedIndex = ptIdx > 0 ? ptIdx : 0;
            var idx = cboStatus.Items.IndexOf(h.Status); cboStatus.SelectedIndex = idx >= 0 ? idx : 0;

            // Khôi phục nhà cung cấp
            if (h.Supplier_ID > 0)
            {
                _isSearching = true;
                BindSupplierCombo(_supplierTable);
                cboSupplier.SelectedValue = h.Supplier_ID;
                _isSearching = false;
            }
            else
            {
                _isSearching = true;
                cboSupplier.Text = "";
                _isSearching = false;
            }

            LoadDetails(_selectedPO_ID);
            LoadFiles(h.WorkorderNo, h.Project_Name);
            LoadDeliveries();
        }

        private void BtnNewPO_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("PO", "Tạo PO", "Tạo PO")) return;
            ClearHeader(); _selectedPO_ID = 0; dgvDetails.Rows.Clear(); dgvFiles.Rows.Clear();
            dgvDelivery.Rows.Clear();
            UpdateTotal(); txtPONo.Focus(); lblStatus.Text = "Đang tạo đơn PO mới...";
        }

        // =========================================================================
        // LƯU TOÀN BỘ PO (Header + Detail)
        // =========================================================================
        private void BtnSavePO_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("PO", "Lưu PO", "Lưu PO")) return;
            dgvDetails.EndEdit();
            if (string.IsNullOrWhiteSpace(txtPONo.Text))
            {
                MessageBox.Show("Vui lòng nhập PO No!", "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Warning); txtPONo.Focus(); return;
            }
            // Bat buoc chon Nha Cung Cap
            int _suppId = Convert.ToInt32(cboSupplier.SelectedValue ?? 0);
            if (_suppId <= 0 || string.IsNullOrWhiteSpace(cboSupplier.Text))
            {
                MessageBox.Show("Vui long chon Nha Cung Cap truoc khi luu PO!",
                    "Thieu thong tin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cboSupplier.Focus(); return;
            }
            if (dgvDetails.Rows.Count == 0 && MessageBox.Show("Don hang nay chua co chi tiet vat tu nao.\nBan co chac chan muon luu chi voi Header khong?", "Canh bao", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No) return;
            try
            {
                string basePONo = txtPONo.Text.Trim();
                int revIdx = basePONo.LastIndexOf("_Rev"); if (revIdx > 0) basePONo = basePONo.Substring(0, revIdx);
                string finalPONo = basePONo;

                if (_selectedPO_ID == 0)
                {
                    // TAO MOI: kiem tra trung so PO
                    bool isBaseDuplicate = _poList.Exists(p => p.PONo == basePONo && p.PO_ID != _selectedPO_ID);
                    if (isBaseDuplicate || nudRevise.Value > 0)
                    {
                        if (nudRevise.Value == 0)
                        {
                            MessageBox.Show("So PO nay da ton tai!\nVui long tang so Revise de tao ban sua doi.",
                                "Trung lap", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            nudRevise.Focus(); return;
                        }
                        finalPONo = $"{basePONo}_Rev{nudRevise.Value}";
                        if (_poList.Exists(p => p.PONo == finalPONo && p.PO_ID != _selectedPO_ID))
                        {
                            MessageBox.Show($"Ban '{finalPONo}' cung da ton tai!\nVui long tang Revise len cao hon.",
                                "Trung lap", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            nudRevise.Focus(); return;
                        }
                    }
                }
                else
                {
                    // UPDATE PO CU: giu nguyen so PO, khong check trung
                    finalPONo = _poList.Find(p => p.PO_ID == _selectedPO_ID)?.PONo ?? basePONo;
                    txtPONo.Text = finalPONo;
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
                    Payment_Term = cboPaymentTerm.SelectedIndex > 0
                                   ? cboPaymentTerm.SelectedItem.ToString()
                                   : "",
                    PO_Date = dtpPODate.Value,
                    Status = cboStatus.SelectedItem?.ToString() ?? "Draft",
                    Revise = (int)nudRevise.Value,
                    Supplier_ID = Convert.ToInt32(cboSupplier.SelectedValue ?? 0),
                    ProjectCode = _projectCodeImport,
                    Expected_Delivery = dtpPOExpectDelivery.Value
                };
                if (_selectedPO_ID == 0) _selectedPO_ID = _service.InsertHead(h, _currentUser);
                else _service.UpdateHead(h, _currentUser);
                txtPONo.Text = finalPONo;
                SaveDetailsToDb();
                MessageBox.Show($"Đã lưu toàn bộ PO thành công!\n- Số PO: {finalPONo}\n- Số dòng vật tư: {dgvDetails.Rows.Count}", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Tự động xuất Excel vào thư mục PO Link của dự án
                AutoSavePOExcelToPOLink(h, _selectedPO_ID);

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
        // TỰ ĐỘNG XUẤT EXCEL PO VÀO THƯ MỤC PO LINK CỦA DỰ ÁN
        // =========================================================================
        private void AutoSavePOExcelToPOLink(POHead poHead, int poId)
        {
            try
            {
                // 1. Tìm dự án hiện tại để lấy PO_Link
                var projects = new ProjectService().GetAll();
                var prj = projects.Find(p =>
                    (!string.IsNullOrEmpty(p.WorkorderNo) && p.WorkorderNo.Equals(poHead.WorkorderNo, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(p.ProjectName) && p.ProjectName.Equals(poHead.Project_Name, StringComparison.OrdinalIgnoreCase)));

                if (prj == null || string.IsNullOrEmpty(prj.PO_Link))
                {
                    MessageBox.Show(
                        "⚠️ Không tìm thấy đường dẫn PO Link cho dự án này!\n\n" +
                        "Vui lòng cập nhật 'PO Link' trong mục Quản lý Dự án trước khi lưu file Excel tự động.",
                        "Thiếu PO Link", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string poLinkDir = prj.PO_Link.Trim();
                if (!Directory.Exists(poLinkDir))
                {
                    var createDir = MessageBox.Show(
                        $"⚠️ Thư mục PO Link không tồn tại:\n{poLinkDir}\n\nBạn có muốn tạo thư mục này không?",
                        "Thư mục không tồn tại", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (createDir == DialogResult.Yes)
                    {
                        try { Directory.CreateDirectory(poLinkDir); }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"❌ Không thể tạo thư mục:\n{poLinkDir}\n\nLỗi: {ex.Message}",
                                "Lỗi tạo thư mục", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                    else return;
                }

                // 2. Tìm template
                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "po_template.xlsx");
                if (!File.Exists(templatePath))
                {
                    MessageBox.Show(
                        $"❌ Không tìm thấy file template Excel!\n\nĐường dẫn dự kiến:\n{templatePath}\n\n" +
                        "Vui lòng đảm bảo file 'po_template.xlsx' tồn tại trong thư mục Templates của ứng dụng.",
                        "Thiếu Template", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 3. Đường dẫn file đích
                string safePoNo = poHead.PONo?.Replace("/", "-").Replace("\\", "-").Replace(":", "-") ?? "PO";
                string fileName = $"PO_{safePoNo}_{DateTime.Now:ddMMyyyy}.xlsx";
                string destPath = Path.Combine(poLinkDir, fileName);

                // 4. Hỏi xác nhận trước khi lưu
                bool fileExists = File.Exists(destPath);
                string confirmMsg = fileExists
                    ? $"📁 File Excel đã tồn tại, bạn có muốn GHI ĐÈ không?\n\n" +
                      $"📄 File: {fileName}\n📂 Thư mục: {poLinkDir}"
                    : $"💾 Xác nhận lưu file Excel PO vào thư mục dự án?\n\n" +
                      $"📄 File: {fileName}\n📂 Thư mục: {poLinkDir}";

                var confirm = MessageBox.Show(confirmMsg,
                    fileExists ? "Xác nhận ghi đè" : "Xác nhận lưu file",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (confirm != DialogResult.Yes) return;

                // 5. Copy template và điền dữ liệu
                File.Copy(templatePath, destPath, true);

                var details = _service.GetDetails(poId);
                var suppliers = new SupplierService().GetAll();
                var supplier = suppliers.Find(s => s.Supplier_ID == poHead.Supplier_ID);

                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(destPath)))
                {
                    var ws = package.Workbook.Worksheets[0];

                    // Header tags
                    ReplaceCell(ws, "<<PROJECT_NAME>>", prj.ProjectName ?? poHead.Project_Name ?? "");
                    ReplaceCell(ws, "<<WO-NO>>", poHead.WorkorderNo ?? "");
                    ReplaceCell(ws, "<<REV.NUM>>", poHead.Revise.ToString());
                    ReplaceCell(ws, "<<DATE>>", poHead.PO_Date.HasValue ? poHead.PO_Date.Value.ToString("dd/MM/yyyy") : DateTime.Today.ToString("dd/MM/yyyy"));
                    ReplaceCell(ws, "<<MPR-NO>>", poHead.MPR_No ?? "");
                    ReplaceCell(ws, "<<PO-NO>>", poHead.PONo ?? "");
                    string supplierInfo = supplier != null ? $"{supplier.Company_Name}\nCert: {supplier.Cert ?? ""}\nEmail: {supplier.Email}" : "";
                    ReplaceCell(ws, "<<SUPPLIER-INFO>>", supplierInfo);

                    ws.Cells[5, 15].Value = !string.IsNullOrEmpty(poHead.Payment_Term)
                        ? poHead.Payment_Term : "Within 7 days after delivery";
                    ws.Cells[8, 11].Value = DateTime.Today.AddDays(7).ToString("dd/MM/yyyy");

                    int startRow = 8;
                    int detailCount = details.Count;

                    if (detailCount > 1)
                    {
                        ws.InsertRow(startRow + 1, detailCount - 1);
                        for (int i = 1; i < detailCount; i++)
                            ws.Cells[startRow, 1, startRow, 16].Copy(ws.Cells[startRow + i, 1]);
                    }

                    decimal totalAfterVAT = 0;
                    for (int i = 0; i < detailCount; i++)
                    {
                        var d = details[i];
                        int row = startRow + i;
                        decimal q = d.Qty_Per_Sheet;
                        decimal wk = d.Weight_kg;
                        decimal realPrice = d.Price;
                        string rem = d.Remarks ?? "";

                        if (rem.Contains("[CALC:KG]"))
                        {
                            rem = rem.Replace("[CALC:KG]", "").Trim();
                            if (wk > 0 && q > 0) realPrice = (d.Price * q) / wk;
                        }

                        ws.Cells[row, 1].Value = i + 1;
                        ws.Cells[row, 2].Value = d.Item_Name ?? "";
                        ws.Cells[row, 3].Value = d.Material ?? "";
                        ws.Cells[row, 4].Value = d.Asize;
                        ws.Cells[row, 5].Value = d.Bsize;
                        ws.Cells[row, 6].Value = d.Csize;
                        ws.Cells[row, 7].Value = d.Qty_Per_Sheet;
                        ws.Cells[row, 8].Value = d.UNIT ?? "";
                        ws.Cells[row, 9].Value = d.Weight_kg;
                        ws.Cells[row, 10].Value = d.MPSNo ?? "";
                        ws.Cells[row, 11].Value = d.RequestDay;
                        ws.Cells[row, 12].Value = "Kho DLHI";
                        ws.Cells[row, 13].Value = Math.Round(realPrice, 0);
                        ws.Cells[row, 14].Value = d.Amount;
                        ws.Cells[row, 16].Value = rem;

                        totalAfterVAT += d.Amount * (1 + d.VAT / 100);
                    }

                    int subTotalRow = startRow + detailCount;
                    int vatRow = subTotalRow + 1;
                    ws.Cells[subTotalRow, 3].Value = "SUB-TOTAL";
                    ws.Cells[subTotalRow, 14].Formula = $"=SUM(N{startRow}:N{startRow + detailCount - 1})";
                    ws.Cells[vatRow, 3].Value = "Final Price Requested (Included VAT)";
                    ws.Cells[vatRow, 14].Value = totalAfterVAT;
                    ws.Cells[vatRow, 14].Style.Numberformat.Format = "#,##0.##";
                    ws.Cells[vatRow, 14].Style.Font.Bold = true;

                    package.Save();
                }

                // 6. Refresh danh sách file đính kèm
                LoadFiles(poHead.WorkorderNo, poHead.Project_Name);

                // 7. Thông báo thành công và tự động mở file
                MessageBox.Show(
                    $"✅ Đã lưu file Excel thành công!\n\n📄 File: {fileName}\n📂 Thư mục: {poLinkDir}",
                    "Lưu file thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);

                Process.Start(new ProcessStartInfo { FileName = destPath, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Lỗi khi tự động lưu file Excel vào PO Link:\n\n{ex.Message}",
                    "Lỗi lưu file", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =========================================================================
        // LƯU CHI TIẾT — chỉ lưu detail, không đụng Header
        // =========================================================================
        private void BtnSaveDetail_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("PO", "Lưu chi tiết", "Lưu chi tiết")) return;
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

            // ── Xác thực mật khẩu Admin trước khi lưu ──
            if (!VerifyAdminPassword()) return;

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
            dlg.Controls.Add(new Label { Text = "Nhập mật khẩu tài khoản Admin để xác nhận lưu:", Font = new Font("Segoe UI", 9), Location = new Point(15, 15), Size = new Size(340, 20) });
            var txtPwd = new TextBox { Location = new Point(15, 40), Size = new Size(340, 26), Font = new Font("Segoe UI", 10), PasswordChar = '●' };
            dlg.Controls.Add(txtPwd);
            var lblErr = new Label { Text = "", ForeColor = Color.FromArgb(220, 53, 69), Font = new Font("Segoe UI", 9, FontStyle.Bold), Location = new Point(15, 72), Size = new Size(340, 20) };
            dlg.Controls.Add(lblErr);
            var btnOK = new Button { Text = "✔ Xác nhận", Location = new Point(155, 98), Size = new Size(100, 30), BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            btnOK.FlatAppearance.BorderSize = 0;
            dlg.Controls.Add(btnOK);
            var btnCancel = new Button { Text = "Hủy", Location = new Point(265, 98), Size = new Size(90, 30), BackColor = Color.FromArgb(108, 117, 125), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), DialogResult = DialogResult.Cancel };
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
                catch (Exception ex) { lblErr.Text = "Lỗi: " + ex.Message; }
            };
            dlg.KeyDown += (s, ev) => { if (ev.KeyCode == Keys.Enter) { btnOK.PerformClick(); ev.SuppressKeyPress = true; } };
            txtPwd.Focus();
            dlg.ShowDialog(this);
            return verified;
        }

        // =========================================================================
        // HÀM CHUNG lưu detail — dùng bởi cả BtnSavePO và BtnSaveDetail
        // =========================================================================
        private void SaveDetailsToDb()
        {
            var oldDetails = _service.GetDetails(_selectedPO_ID);

            // ── Thu thập dữ liệu mới từ grid ──
            var newRows = new System.Collections.Generic.List<PODetail>();
            int itemNo = 1;
            foreach (DataGridViewRow row in dgvDetails.Rows)
            {
                if (row.IsNewRow || row.Tag?.ToString() == "TOTAL") continue;
                decimal q = decimal.TryParse(row.Cells["Qty"].Value?.ToString(), out decimal _q) ? _q : 0;
                decimal wk = decimal.TryParse(row.Cells["Weight"].Value?.ToString(), out decimal _wk) ? _wk : 0;
                decimal p = decimal.TryParse((row.Cells["Price"].Value?.ToString() ?? "0").Replace(",", ""), out decimal _p) ? _p : 0;
                decimal vat = decimal.TryParse(row.Cells["VAT"].Value?.ToString(), out decimal _vt) ? _vt : 0;
                string calcMethod = row.Cells["Calc_Method"].Value?.ToString() ?? "Theo KG";
                string remarks = row.Cells["Remarks"].Value?.ToString() ?? "";
                remarks = remarks.Replace("[CALC:KG]", "").Replace("[CALC:SL]", "").Trim();
                decimal dbPrice = p;
                if (calcMethod == "Theo KG") { remarks += " [CALC:KG]"; if (q > 0 && wk > 0) dbPrice = (wk * p) / q; }
                else remarks += " [CALC:SL]";
                int? mprDetailId = null;
                if (dgvDetails.Columns.Contains("MPR_Detail_ID") && row.Cells["MPR_Detail_ID"].Value != null)
                    if (int.TryParse(row.Cells["MPR_Detail_ID"].Value.ToString(), out int mdi) && mdi > 0) mprDetailId = mdi;
                newRows.Add(new PODetail
                {
                    Item_No = itemNo++,
                    Item_Name = row.Cells["Item_Name"].Value?.ToString() ?? "",
                    Material = row.Cells["Material"].Value?.ToString() ?? "",
                    Asize = row.Cells["Asize"].Value?.ToString() ?? "",
                    Bsize = row.Cells["Bsize"].Value?.ToString() ?? "",
                    Csize = row.Cells["Csize"].Value?.ToString() ?? "",
                    Qty_Per_Sheet = q,
                    UNIT = row.Cells["UNIT"].Value?.ToString() ?? "",
                    Weight_kg = wk,
                    Price = dbPrice,
                    VAT = vat,
                    Amount = 0,
                    Received = int.TryParse(row.Cells["Received"].Value?.ToString(), out int rec) ? rec : 0,
                    MPSNo = row.Cells["MPSNo"].Value?.ToString() ?? "",
                    DeliveryLocation = row.Cells["DeliveryLocation"].Value?.ToString() ?? "",
                    Remarks = remarks.Trim(),
                    MPR_Detail_ID = mprDetailId
                });
            }

            // ── So sánh cũ vs mới → ghi Revise History ──
            var reviseLogs = new System.Collections.Generic.List<(int ino, string col, string oldV, string newV)>();
            foreach (var n in newRows)
            {
                var o = oldDetails.Find(x => x.Item_No == n.Item_No);
                if (o == null) { reviseLogs.Add((n.Item_No, "Item mới", "", n.Item_Name)); continue; }
                void Chk(string c, string ov, string nv) { if ((ov ?? "").Trim() != (nv ?? "").Trim()) reviseLogs.Add((n.Item_No, c, ov ?? "", nv ?? "")); }
                Chk("Tên hàng", o.Item_Name, n.Item_Name);
                Chk("Vật liệu", o.Material, n.Material);
                Chk("A(mm)", o.Asize, n.Asize);
                Chk("B(mm)", o.Bsize, n.Bsize);
                Chk("C(mm)", o.Csize, n.Csize);
                Chk("SL", o.Qty_Per_Sheet.ToString(), n.Qty_Per_Sheet.ToString());
                Chk("ĐVT", o.UNIT, n.UNIT);
                Chk("KG", o.Weight_kg.ToString("0.##"), n.Weight_kg.ToString("0.##"));
                Chk("Đơn giá", o.Price.ToString("0.##"), n.Price.ToString("0.##"));
                Chk("VAT", o.VAT.ToString("0.##"), n.VAT.ToString("0.##"));
                Chk("Đã nhận", o.Received.ToString(), n.Received.ToString());
                Chk("Ghi chú", o.Remarks, n.Remarks);
                Chk("Nơi giao", o.DeliveryLocation, n.DeliveryLocation);
                Chk("MPS No", o.MPSNo, n.MPSNo);
            }
            foreach (var o in oldDetails)
                if (!newRows.Exists(n => n.Item_No == o.Item_No))
                    reviseLogs.Add((o.Item_No, "Item xóa", o.Item_Name, ""));

            // ── Xóa cũ, insert mới ──
            foreach (var d in oldDetails) _service.DeleteDetail(d.PO_Detail_ID);
            foreach (var d in newRows) _service.InsertDetail(d, _selectedPO_ID);

            // ── Ghi Revise History ──
            if (reviseLogs.Count > 0 && _selectedPO_ID > 0)
            {
                try
                {
                    using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                    conn.Open();
                    foreach (var (ino, col, oldV, newV) in reviseLogs)
                    {
                        var cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                            INSERT INTO PO_Revise_Transactions (po_id, item_no, column_name_change, old_value, new_value, trans_date)
                            VALUES (@poId, @itemNo, @col, @oldVal, @newVal, GETDATE())", conn);
                        cmd.Parameters.AddWithValue("@poId", _selectedPO_ID);
                        cmd.Parameters.AddWithValue("@itemNo", ino);
                        cmd.Parameters.AddWithValue("@col", col);
                        cmd.Parameters.AddWithValue("@oldVal", oldV);
                        cmd.Parameters.AddWithValue("@newVal", newV);
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex) { System.Diagnostics.Debug.WriteLine("ReviseLog: " + ex.Message); }
            }
        }

        private void BtnDeletePO_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("PO", "Xóa PO", "Xóa PO")) return;
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
                    // Xóa Revise History trước để tránh lỗi FK_PO_Transaction_PO
                    using (var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection())
                    {
                        conn.Open();
                        var cmd = new Microsoft.Data.SqlClient.SqlCommand(
                            "DELETE FROM PO_Revise_Transactions WHERE po_id = @poId", conn);
                        cmd.Parameters.AddWithValue("@poId", _selectedPO_ID);
                        cmd.ExecuteNonQuery();
                    }

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
            if (!PermissionHelper.Check("PO", "Thêm dòng", "Thêm dòng chi tiết")) return;
            // Tính số thứ tự — không đếm dòng TOTAL
            int dataRowCount = dgvDetails.Rows.Cast<DataGridViewRow>()
                .Count(r => !r.IsNewRow && r.Tag?.ToString() != "TOTAL");
            int nextItem = dataRowCount + 1;

            // Chèn dòng mới TRƯỚC dòng TOTAL (nếu có)
            int insertIdx = dgvDetails.Rows.Count;
            for (int i = dgvDetails.Rows.Count - 1; i >= 0; i--)
                if (dgvDetails.Rows[i].Tag?.ToString() == "TOTAL") { insertIdx = i; break; }

            int newIdx;
            if (insertIdx < dgvDetails.Rows.Count)
            {
                dgvDetails.Rows.Insert(insertIdx, 1);
                newIdx = insertIdx;
            }
            else
                newIdx = dgvDetails.Rows.Add();

            var r = dgvDetails.Rows[newIdx];
            r.Cells["DeliveryLocation"].Value = ""; r.Cells["Item_No"].Value = nextItem;
            r.Cells["Item_Name"].Value = ""; r.Cells["Material"].Value = "";
            r.Cells["Asize"].Value = ""; r.Cells["Bsize"].Value = ""; r.Cells["Csize"].Value = "";
            r.Cells["Qty"].Value = 0; r.Cells["UNIT"].Value = "PCS"; r.Cells["Weight"].Value = 0;
            r.Cells["Price"].Value = 0; r.Cells["VAT"].Value = "10";
            r.Cells["Amount"].Value = 0; r.Cells["Received"].Value = 0;
            r.Cells["MPSNo"].Value = ""; r.Cells["Remarks"].Value = "";
            r.Cells["Calc_Method"].Value = "Theo KG"; r.Cells["Ordered_PO"].Value = "";
            r.Cells["PO_Detail_ID"].Value = 0;
            dgvDetails.CurrentCell = r.Cells["Item_Name"];
            AutoAdjustColumnWidths();
        }

        private void BtnDeleteDetail_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("PO", "Xóa dòng", "Xóa dòng chi tiết")) return;
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
                    foreach (DataGridViewRow row in dgvDetails.SelectedRows)
                        if (!row.IsNewRow && row.Tag?.ToString() != "TOTAL") rowsToDelete.Add(row);
                    foreach (var row in rowsToDelete) dgvDetails.Rows.Remove(row);
                    int itemNo = 1;
                    foreach (DataGridViewRow row in dgvDetails.Rows)
                        if (!row.IsNewRow && row.Tag?.ToString() != "TOTAL") row.Cells["Item_No"].Value = itemNo++;
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
            if (!PermissionHelper.Check("PO", "Import MPR", "Import MPR")) return;
            using (var dlg = new frmSelectMPR())
            {
                if (dlg.ShowDialog() == DialogResult.OK && dlg.SelectedMPR != null)
                {
                    ClearHeader(); _selectedPO_ID = 0; _details.Clear(); dgvDetails.Rows.Clear(); dgvFiles.Rows.Clear();
                    var mpr = dlg.SelectedMPR;
                    var details = dlg.SelectedDetails; var poMapping = GetPoMappingForMpr(mpr.MPR_ID);
                    txtProjectName.Text = mpr.Project_Name; txtMPRNo.Text = mpr.MPR_No; LoadMPRFiles();
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
                var poMapping = GetPoMappingForMpr(mpr.MPR_ID); txtProjectName.Text = mpr.Project_Name; txtMPRNo.Text = mpr.MPR_No; LoadMPRFiles();
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
                           ISNULL(s.Short_Name, ISNULL(s.Company_Name,'')) AS NCC,
                           CONVERT(NVARCHAR(10), dt.ExpDelivery, 103) AS ExpDelivery,
                           ISNULL(dt.GhiChu,'') AS GhiChu,
                           ISNULL(dt.Status,'Pending') AS Status,
                           ISNULL(dt.ReceiverNote,'') AS ReceiverNote
                    FROM PO_DeliveryTracking dt
                    LEFT JOIN PO_head     ph ON ph.PONo        = dt.PONo
                    LEFT JOIN ProjectInfo pi ON pi.WorkorderNo = ph.WorkorderNo
                    LEFT JOIN Suppliers   s  ON s.Supplier_ID  = ph.Supplier_ID
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
                            r["TrackID"], r["PONo"], r["MaDuAn"], r["NCC"],
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
                    Text = $"MPR No: {po.MPR_No ?? "—"}    |    Dự án: {po.Project_Name ?? "—"}    |    NCC: {suppName}    |    Ngày PO: {(po.PO_Date.HasValue ? po.PO_Date.Value.ToString("dd/MM/yyyy") : "—")}    |    Tổng: " + po.Total_Amount.ToString("N2", _numCulture) + " VNĐ",
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
                dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
                dgv.DefaultCellStyle.SelectionForeColor = Color.Black;
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
                    SELECT ph.PONo,
                           ISNULL(pi.ProjectCode,'')                        AS MaDuAn,
                           ph.MPR_No,
                           ph.Status,
                           ISNULL(s.Short_Name, ISNULL(s.Company_Name,''))  AS NCC
                    FROM PO_head ph
                    LEFT JOIN ProjectInfo pi ON pi.WorkorderNo = ph.WorkorderNo
                    LEFT JOIN Suppliers   s  ON s.Supplier_ID  = ph.Supplier_ID
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
                var pFilter = new FlowLayoutPanel
                {
                    Location = new Point(10, 36),
                    Size = new Size(dlg.ClientSize.Width - 20, 38),
                    BackColor = Color.White,
                    BorderStyle = BorderStyle.FixedSingle,
                    WrapContents = false,
                    FlowDirection = FlowDirection.LeftToRight,
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
                };
                dlg.Controls.Add(pFilter);

                pFilter.Controls.Add(new Label { Text = "Mã DA:", Size = new Size(45, 28), TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Segoe UI", 8, FontStyle.Bold) });
                var cboDaFilter = new ComboBox { Size = new Size(115, 26), Font = new Font("Segoe UI", 9), DropDownStyle = ComboBoxStyle.DropDownList };
                cboDaFilter.Items.Add("Tất cả");
                dtPO.AsEnumerable().Select(r => r["MaDuAn"].ToString())
                    .Where(v => !string.IsNullOrEmpty(v)).Distinct().OrderBy(v => v)
                    .ToList().ForEach(v => cboDaFilter.Items.Add(v));
                cboDaFilter.SelectedIndex = 0;
                pFilter.Controls.Add(cboDaFilter);

                pFilter.Controls.Add(new Label { Text = "PO No:", Size = new Size(45, 28), TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Segoe UI", 8, FontStyle.Bold) });
                var txtPoFilter = new TextBox { Size = new Size(110, 26), Font = new Font("Segoe UI", 9), PlaceholderText = "PO No..." };
                pFilter.Controls.Add(txtPoFilter);

                pFilter.Controls.Add(new Label { Text = "MPR No:", Size = new Size(52, 28), TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Segoe UI", 8, FontStyle.Bold) });
                var txtMprFilter = new TextBox { Size = new Size(105, 26), Font = new Font("Segoe UI", 9), PlaceholderText = "MPR No..." };
                pFilter.Controls.Add(txtMprFilter);

                pFilter.Controls.Add(new Label { Text = "NCC:", Size = new Size(35, 28), TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Segoe UI", 8, FontStyle.Bold) });
                var txtNccFilter = new TextBox { Size = new Size(120, 26), Font = new Font("Segoe UI", 9), PlaceholderText = "Nhà cung cấp..." };
                pFilter.Controls.Add(txtNccFilter);

                var btnDlgFilter = new Button { Text = "🔍 Lọc", Size = new Size(70, 26), BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8, FontStyle.Bold) };
                btnDlgFilter.FlatAppearance.BorderSize = 0;
                pFilter.Controls.Add(btnDlgFilter);

                var btnDlgClear = new Button { Text = "✖", Size = new Size(32, 26), BackColor = Color.FromArgb(108, 117, 125), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8, FontStyle.Bold) };
                btnDlgClear.FlatAppearance.BorderSize = 0;
                pFilter.Controls.Add(btnDlgClear);

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
                dgvDlg.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
                dgvDlg.DefaultCellStyle.SelectionForeColor = Color.Black;
                dlg.Controls.Add(dgvDlg);

                // Sau khi bind: đặt header, căn chỉnh width
                dgvDlg.DataBindingComplete += (s, ev) =>
                {
                    if (dgvDlg.Columns.Contains("PONo")) { dgvDlg.Columns["PONo"].HeaderText = "PO No"; dgvDlg.Columns["PONo"].FillWeight = 28; }
                    if (dgvDlg.Columns.Contains("MaDuAn")) { dgvDlg.Columns["MaDuAn"].HeaderText = "Mã DA"; dgvDlg.Columns["MaDuAn"].FillWeight = 15; }
                    if (dgvDlg.Columns.Contains("NCC")) { dgvDlg.Columns["NCC"].HeaderText = "NCC"; dgvDlg.Columns["NCC"].FillWeight = 25; }
                    if (dgvDlg.Columns.Contains("MPR_No")) { dgvDlg.Columns["MPR_No"].HeaderText = "MPR No"; dgvDlg.Columns["MPR_No"].FillWeight = 20; }
                    if (dgvDlg.Columns.Contains("Status")) { dgvDlg.Columns["Status"].HeaderText = "Trạng thái"; dgvDlg.Columns["Status"].FillWeight = 12; }
                };

                // Hàm bind bảng PO theo filter
                Action bindDlgGrid = () =>
                {
                    string selDa = cboDaFilter.SelectedItem?.ToString() ?? "Tất cả";
                    string selPo = txtPoFilter.Text.Trim().ToLower();
                    string selMpr = txtMprFilter.Text.Trim().ToLower();
                    string selNcc = txtNccFilter.Text.Trim().ToLower();
                    var rows = dtPO.AsEnumerable().Where(r =>
                    {
                        if (selDa != "Tất cả" && r["MaDuAn"].ToString() != selDa) return false;
                        if (!string.IsNullOrEmpty(selPo) && !r["PONo"].ToString().ToLower().Contains(selPo)) return false;
                        if (!string.IsNullOrEmpty(selMpr) && !r["MPR_No"].ToString().ToLower().Contains(selMpr)) return false;
                        if (!string.IsNullOrEmpty(selNcc) && !r["NCC"].ToString().ToLower().Contains(selNcc)) return false;
                        return true;
                    });
                    dgvDlg.DataSource = rows.Any() ? rows.CopyToDataTable() : dtPO.Clone();
                };
                bindDlgGrid();

                btnDlgFilter.Click += (s, ev) => bindDlgGrid();
                btnDlgClear.Click += (s, ev) =>
                {
                    cboDaFilter.SelectedIndex = 0;
                    txtPoFilter.Text = "";
                    txtMprFilter.Text = "";
                    txtNccFilter.Text = "";
                    bindDlgGrid();
                };
                cboDaFilter.SelectedIndexChanged += (s, ev) => bindDlgGrid();
                txtPoFilter.KeyDown += (s, ev) => { if (ev.KeyCode == Keys.Enter) { bindDlgGrid(); ev.SuppressKeyPress = true; } };
                txtMprFilter.KeyDown += (s, ev) => { if (ev.KeyCode == Keys.Enter) { bindDlgGrid(); ev.SuppressKeyPress = true; } };
                txtNccFilter.KeyDown += (s, ev) => { if (ev.KeyCode == Keys.Enter) { bindDlgGrid(); ev.SuppressKeyPress = true; } };

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

                dlg.KeyDown += (s, ev) =>
                {
                    if (ev.KeyCode != Keys.Enter) return;
                    var focused = dlg.ActiveControl;
                    if (focused == txtPoFilter || focused == txtMprFilter || focused == txtNccFilter)
                    {
                        bindDlgGrid();
                        ev.Handled = true;
                        ev.SuppressKeyPress = true;
                        return;
                    }
                    btnOK.PerformClick();
                    ev.Handled = true;
                };
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
                dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
                dgv.DefaultCellStyle.SelectionForeColor = Color.Black;

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
            cboPaymentTerm.SelectedIndex = 0;
            // Reset Nha CC ve rong
            _isSearching = true; cboSupplier.Text = ""; cboSupplier.SelectedIndex = -1;
            BindSupplierCombo(_supplierTable); _isSearching = false;
            cboSupplier.BackColor = Color.White;
        }

        // =========================================================================
        // CHECK BY SIZE — Popup load TOÀN BỘ dữ liệu + bộ lọc
        // =========================================================================
        private void BtnCheckBySize_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("PO", "Check by size", "Check by size")) return;
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
                dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
                dgv.DefaultCellStyle.SelectionForeColor = Color.Black;
                popup.Controls.Add(dgv);

                // Sau khi DataSource được set lần đầu → chỉnh width cột PO No
                dgv.DataBindingComplete += (s, ev) =>
                {
                    if (dgv.Columns.Contains("PO No"))
                    {
                        dgv.Columns["PO No"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                        dgv.Columns["PO No"].Width = 150;
                    }
                };

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

                // ── NÚT XUẤT EXCEL ──
                var btnExport = new Button
                {
                    Text = "📥 Xuất Excel",
                    Size = new Size(120, 30),
                    BackColor = Color.FromArgb(0, 150, 100),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
                    Location = new Point(10, popup.ClientSize.Height - 40)
                };
                btnExport.FlatAppearance.BorderSize = 0;
                btnExport.Click += (s, ev) =>
                {
                    if (dgv.Rows.Count == 0)
                    { MessageBox.Show("Không có dữ liệu để xuất!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

                    using var sfd = new SaveFileDialog
                    {
                        Title = "Xuất Check by Size",
                        Filter = "Excel Files|*.xlsx",
                        FileName = $"CheckBySize_{DateTime.Now:yyyyMMdd_HHmm}",
                        InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                    };
                    if (sfd.ShowDialog() != DialogResult.OK) return;
                    try
                    {
                        OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                        using var pkg = new OfficeOpenXml.ExcelPackage();
                        var ws = pkg.Workbook.Worksheets.Add("Check by Size");

                        // Header
                        int colCount = dgv.Columns.Count;
                        for (int c = 0; c < colCount; c++)
                        {
                            ws.Cells[1, c + 1].Value = dgv.Columns[c].HeaderText;
                            ws.Cells[1, c + 1].Style.Font.Bold = true;
                            ws.Cells[1, c + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            ws.Cells[1, c + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(102, 51, 153));
                            ws.Cells[1, c + 1].Style.Font.Color.SetColor(Color.White);
                            ws.Cells[1, c + 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            ws.Cells[1, c + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        }

                        // Data — lấy từ DataTable hiện tại của dgv
                        var dtExport = dgv.DataSource as DataTable;
                        if (dtExport != null)
                        {
                            for (int r = 0; r < dtExport.Rows.Count; r++)
                            {
                                for (int c = 0; c < dtExport.Columns.Count; c++)
                                {
                                    var val = dtExport.Rows[r][c];
                                    ws.Cells[r + 2, c + 1].Value = val != DBNull.Value ? val : null;
                                }
                                // Tô màu xen kẽ
                                if (r % 2 == 1)
                                    for (int c = 0; c < dtExport.Columns.Count; c++)
                                    {
                                        ws.Cells[r + 2, c + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                        ws.Cells[r + 2, c + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(248, 240, 255));
                                    }
                            }
                        }

                        // Border và AutoFit
                        if (ws.Dimension != null)
                        {
                            ws.Cells[ws.Dimension.Address].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            ws.Cells[ws.Dimension.Address].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            ws.Cells[ws.Dimension.Address].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            ws.Cells[ws.Dimension.Address].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        }
                        ws.Cells[ws.Dimension.Address].AutoFitColumns();
                        ws.View.FreezePanes(2, 1);

                        pkg.SaveAs(new System.IO.FileInfo(sfd.FileName));
                        MessageBox.Show($"✅ Đã xuất {dgv.Rows.Count} dòng thành công!", "Thành công",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        { FileName = sfd.FileName, UseShellExecute = true });
                    }
                    catch (Exception ex2)
                    {
                        MessageBox.Show("Lỗi xuất Excel: " + ex2.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                };
                popup.Controls.Add(btnExport);

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
                popup.AcceptButton = btnFilter;
                popup.CancelButton = btnClose;

                // Resize handler
                popup.Resize += (s, ev) =>
                {
                    btnClose.Location = new Point(popup.ClientSize.Width - 115, popup.ClientSize.Height - 40);
                    btnExport.Location = new Point(10, popup.ClientSize.Height - 40);
                    panelFilter.Width = popup.ClientSize.Width - 20;
                    dgv.Size = new Size(popup.ClientSize.Width - 20, popup.ClientSize.Height - 180);
                };

                popup.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tra cứu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // =====================================================
        //  ÁP DỤNG PHÂN QUYỀN
        // =====================================================
        private void ApplyPermissions()
        {
            if (btnNewPO != null) PermissionHelper.Apply(btnNewPO, "PO", "Tạo PO");
            if (btnDeletePO != null) PermissionHelper.Apply(btnDeletePO, "PO", "Xóa PO");
            if (btnSavePO != null) PermissionHelper.Apply(btnSavePO, "PO", "Lưu PO");
            if (btnExport != null) PermissionHelper.Apply(btnExport, "PO", "Xuất Excel");
            if (btnAddDetail != null) PermissionHelper.Apply(btnAddDetail, "PO", "Thêm dòng");
            if (btnDeleteDetail != null) PermissionHelper.Apply(btnDeleteDetail, "PO", "Xóa dòng");
            foreach (var c in this.Controls.Find("btnImportMPR", true))
                PermissionHelper.Apply(c, "PO", "Import MPR");
            foreach (var c in this.Controls.Find("btnSaveDetail", true))
                PermissionHelper.Apply(c, "PO", "Lưu chi tiết");
            foreach (var c in this.Controls.Find("btnCheckBySize", true))
                PermissionHelper.Apply(c, "PO", "Check by size");
        }

    }
}