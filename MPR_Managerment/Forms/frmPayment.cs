//  FILE: Forms/frmPayment.cs
//  Tab 1: Tiến độ thanh toán từng PO
//  Tab 2: Báo cáo tổng hợp công nợ NCC theo kỳ
// ============================================================
using MPR_Managerment.Common;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;
using MPR_Managerment.Services;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace MPR_Managerment.Forms
{
    public partial class frmPayment : Form, IRefreshable // Implement IRefreshable interface
    {
        private readonly PaymentService _svc = new PaymentService();
        private readonly POService _poSvc = new POService();
        private readonly SupplierService _suppSvc = new SupplierService();

        // State
        private List<POPaymentSummary> _poSummaries = new List<POPaymentSummary>();
        private List<PaymentSchedule> _schedules = new List<PaymentSchedule>();
        private List<PaymentHistory> _histories = new List<PaymentHistory>();
        private List<DebtReportItem> _debtReport = new List<DebtReportItem>();
        private List<SupplierDebtSummary> _suppDebt = new List<SupplierDebtSummary>();
        private List<Supplier> _allSuppliers = new List<Supplier>();
        private int _selectedPO_ID = 0;
        private int _selectedSchedID = 0;
        private int _selectedHistID = 0;
        private string _currentUser = AppSession.CurrentUser?.Username ?? "Admin";
        private Dictionary<int, List<PaymentSchedule>> _allSchedulesCache
            = new Dictionary<int, List<PaymentSchedule>>();
        // Cache tổng số tiền đã thanh toán (sau thuế) từ Zalo_PaymentImport theo PONo (toàn thời gian)
        private Dictionary<string, decimal> _zaloPaidCache = new Dictionary<string, decimal>();
        // Cache tổng tiền đã TT trong kỳ được chọn (dùng cho tab Báo cáo công nợ)
        private Dictionary<string, decimal> _zaloInRangeCache = new Dictionary<string, decimal>();

        // Controls chính
        private TabControl tabs;
        private TabPage tabPO, tabDebt;

        // Tab PO
        private TextBox txtSearchPO;
        private ComboBox cboStatusFilter;
        private DataGridView dgvPO, dgvSchedule, dgvHistory;
        private Label lblPOName, lblPOAmount, lblPOPaid, lblPORemain, lblPOStatus, lblPOProgress;
        private Panel panelTop, panelInfo, panelSched, panelHist;
        private DataGridView dgvPaid;
        private ComboBox _cboHistStatus;   // Bộ lọc Status trong panelHist
        private DateTimePicker _paidFrom, _paidTo; // Bộ lọc thời gian trong popup History Paid
        private DataGridView dgvDoc;
        private Panel panelPrintHistory;   // Danh sách PO đã in Request
        private DataGridView dgvPrintHistory;
        private DateTimePicker _phDateFrom, _phDateTo; // Bộ lọc thời gian
        private TextBox _txtPhNCC; // Bộ lọc NCC
        private DateTimePicker _schedDtp;              // DTP overlay cho cột Đến hạn
        private int _schedDtpRow = -1;                 // Row đang được DTP overlay
        private ProgressBar progressPO;

        // Tab Debt
        private DateTimePicker dtpFrom, dtpTo;
        private ComboBox cboSuppFilter;
        private ComboBox cboDebtStatus;          // Lọc trạng thái TT
        private CheckBox chkOverdueOnly;         // Chỉ hiện quá hạn
        private TextBox txtDebtSearch;           // Tìm PO / Dự án
        private Panel _pDebtCards;              // Panel cards — cần dịch chuyển
        private DataGridView dgvDebtSupp, dgvDebtDetail;
        private Label lblSumValue, lblSumPaid, lblSumDebt, lblSumOverdue;
        private Button btnExportDebt;
        private Panel _pNCC, _pDet;   // Panels tab Debt — dùng trong ResizeAll

        private Button btnRefreshPO;

        // Preview file (like frmPO)
        private System.Windows.Forms.Timer _previewTimer;
        private string _previewFilePath = null;
        private DataGridViewCell _previewCell = null;
        private System.Windows.Forms.Timer _previewCloseTimer;
        private Form _previewForm;

        // =====================================================================
        public frmPayment()
        {
            InitializeComponent();
            BuildUI();
            // Deferred loading: Load data sau khi form hiển thị để tránh blocking UI
            this.Shown += FrmPayment_Shown;
            this.Resize += (s, e) => ResizeAll();
            frmAIChat.Attach(this);
        }

        private async void FrmPayment_Shown(object sender, EventArgs e)
        {
            this.Shown -= FrmPayment_Shown;  // Chỉ chạy 1 lần
            var toast = ToastHelper.Attach(this);
            toast.Show("⏳ Đang tải dữ liệu, vui lòng chờ...");
            try { await LoadDataAsync(); }
            finally { toast.Hide(); }
        }

        // Mở với filter sẵn theo PO No (gọi từ frmPO)
        public frmPayment(string currentUser, string initPoNo = "") : this()
        {
            if (!string.IsNullOrEmpty(currentUser))
                _currentUser = currentUser;
            if (!string.IsNullOrEmpty(initPoNo))
            {
                txtSearchPO.Text = initPoNo;
                FilterAndBind();
                // Tự động chọn dòng đầu nếu tìm thấy đúng 1 PO
                if (dgvPO.Rows.Count == 1)
                {
                    dgvPO.ClearSelection();
                    dgvPO.Rows[0].Selected = true;
                }
            }
        }

        // =====================================================================
        //  BUILD UI
        // =====================================================================
        private void BuildUI()
        {
            this.Text = "💳  Quản lý Thanh toán & Công nợ";
            this.BackColor = Color.FromArgb(245, 245, 245);

            tabs = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };
            this.Controls.Add(tabs);

            tabPO = new TabPage("💳  Tiến độ thanh toán PO");
            tabDebt = new TabPage("📊  Báo cáo công nợ NCC");
            tabs.TabPages.AddRange(new[] { tabPO, tabDebt });
            // Gọi ResizeAll khi chuyển tab để layout đúng kích thước
            tabs.SelectedIndexChanged += (s, e) => ResizeAll();

            tabPO.BackColor = tabDebt.BackColor = Color.FromArgb(245, 245, 245);

            BuildTabPO();
            BuildTabDebt();
        }

        private void BuildTabPO()
        {
            var pFilter = P(tabPO, 5, 5, 0, 42, Color.White);
            pFilter.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            Lbl(pFilter, "Tìm:", 6, 12, 40, 20);
            txtSearchPO = Txt(pFilter, 46, 8, 220);
            txtSearchPO.PlaceholderText = "PO No / Dự án / NCC... (Enter để tìm)";
            txtSearchPO.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; FilterAndBind(); } };

            Lbl(pFilter, "Trạng thái:", 278, 12, 85, 20);
            cboStatusFilter = Cbo(pFilter, 363, 8, 180,
                new[] { "Tất cả", "Pending", "Thanh toán 1 phần", "Đã thanh toán", "⚠ Quá hạn" });
            cboStatusFilter.SelectedIndexChanged += (s, e) => FilterAndBind();

            btnRefreshPO = Btn("🔄 Làm mới", Color.FromArgb(0, 120, 212), 555, 8, 105, 26);
            btnRefreshPO.Click += (s, e) => LoadPOSummary();
            pFilter.Controls.Add(btnRefreshPO);

            panelTop = P(tabPO, 5, 52, 0, 190, Color.White);
            panelTop.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Lbl(panelTop, "DANH SÁCH ĐƠN PO", 8, 5, 350, 20, true, Color.FromArgb(0, 120, 212));
            dgvPO = Grid(panelTop, 28, 156);
            dgvPO.SelectionChanged += DgvPO_SelectionChanged;
            dgvPO.CellFormatting += DgvPO_CellFormatting;
            dgvPO.CellMouseEnter += DgvPO_CellMouseEnter;
            dgvPO.CellMouseLeave += DgvPO_CellMouseLeave;
            dgvPO.ColumnHeadersHeight = 60;
            dgvPO.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            BuildPOGridCols();

            panelInfo = new Panel
            {
                Location = new Point(5, 247),
                Size = new Size(0, 65),
                BackColor = Color.FromArgb(0, 120, 212),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            tabPO.Controls.Add(panelInfo);

            lblPOName = InfoLbl(panelInfo, "", 8, 5, 700, 20, 10, true);
            lblPOStatus = InfoLbl(panelInfo, "", 0, 5, 200, 20, 10, true);
            lblPOStatus.TextAlign = ContentAlignment.MiddleRight;

            lblPOAmount = InfoLbl(panelInfo, "Tổng PO: —", 8, 30, 200, 18, 9, false);
            lblPOPaid = InfoLbl(panelInfo, "Đã TT: —", 215, 30, 200, 18, 9, false);
            lblPORemain = InfoLbl(panelInfo, "Còn nợ: —", 422, 30, 220, 18, 9, false);
            lblPOProgress = InfoLbl(panelInfo, "", 650, 30, 100, 18, 9, false);

            progressPO = new ProgressBar
            {
                Location = new Point(640, 32),
                Size = new Size(180, 14),
                Minimum = 0,
                Maximum = 100,
                Value = 0,
                Style = ProgressBarStyle.Continuous
            };
            panelInfo.Controls.Add(progressPO);

            panelSched = P(tabPO, 5, 317, 0, 200, Color.White);
            panelSched.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            Lbl(panelSched, "📅  KẾ HOẠCH THANH TOÁN", 8, 5, 300, 20, true, Color.FromArgb(0, 120, 212));

            bool canEdit = AppSession.CanEdit("PO") || AppSession.CanCreate("PO");
            if (canEdit)
            {
                var bAdd = Btn("+ Thêm", Color.FromArgb(40, 167, 69), 8, 28, 72, 24);
                var bSave = Btn("💾 Lưu", Color.FromArgb(0, 120, 212), 84, 28, 65, 24);
                var bReq = Btn("📄 Eccount", Color.FromArgb(102, 51, 153), 153, 28, 88, 24);
                var bDel = Btn("Xóa", Color.FromArgb(220, 53, 69), 245, 28, 48, 24);
                var bPrint = Btn("🖨 In Req", Color.FromArgb(0, 150, 100), 297, 28, 78, 24);
                var bPrintDoc = Btn("🖨 In tài liệu", Color.FromArgb(102, 51, 153), 379, 28, 100, 24);

                bAdd.Click += BtnAddSched_Click;
                bDel.Click += BtnDelSched_Click;
                bSave.Click += BtnSaveSched_Click;
                bReq.Click += BtnPaymentRequest_Click;
                bPrint.Click += BtnPrintRequest_Click;
                bPrintDoc.Click += BtnPrintDocs_Click;

                panelSched.Controls.AddRange(new Control[] { bAdd, bDel, bSave, bReq, bPrint, bPrintDoc });
            }

            dgvSchedule = Grid(panelSched, 57, 0);
            dgvSchedule.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom;
            dgvSchedule.ColumnHeadersHeight = 50;
            dgvSchedule.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvSchedule.SelectionChanged += (s, e) =>
            {
                if (dgvSchedule.SelectedRows.Count > 0)
                    _selectedSchedID = Convert.ToInt32(dgvSchedule.SelectedRows[0].Cells["S_ID"].Value ?? 0);
            };
            dgvSchedule.CellFormatting += DgvSched_CellFormatting;
            dgvSchedule.CellEndEdit += DgvSchedule_CellEndEdit;
            BuildSchedCols();

            // ── Label + DataGridView Document — 200px bên phải trong panelSched ──
            const int docPanelW = 200;
            Lbl(panelSched, "📎 Document", panelSched.Width - docPanelW, 5, docPanelW - 5, 18, true, Color.FromArgb(0, 120, 212));

            dgvDoc = new DataGridView
            {
                Location = new Point(panelSched.Width - docPanelW, 28),
                Size = new Size(docPanelW - 5, panelSched.Height - 33),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 8),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom,
                Name = "dgvDoc"
            };
            dgvDoc.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvDoc.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvDoc.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 8, FontStyle.Bold);
            dgvDoc.EnableHeadersVisualStyles = false;
            dgvDoc.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvDoc.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
            dgvDoc.DefaultCellStyle.SelectionForeColor = Color.Black;

            dgvDoc.Columns.Add(new DataGridViewTextBoxColumn { Name = "Doc_Path", Visible = false });
            dgvDoc.Columns.Add(new DataGridViewTextBoxColumn { Name = "Doc_Name", HeaderText = "Tên file", FillWeight = 100, ReadOnly = true });

            dgvDoc.CellFormatting += (s, ev) =>
            {
                if (ev.RowIndex < 0) return;
                string docPath = dgvDoc.Rows[ev.RowIndex].Cells["Doc_Path"].Value?.ToString() ?? "";
                bool isInvoice = System.IO.Path.GetFileName(docPath).StartsWith("INV_", StringComparison.OrdinalIgnoreCase);
                ev.CellStyle.ForeColor = isInvoice ? Color.FromArgb(0, 120, 212) : Color.FromArgb(40, 167, 69);
            };

            // Double-click → mở file
            dgvDoc.CellDoubleClick += (s, ev) =>
            {
                if (ev.RowIndex < 0) return;
                string path = dgvDoc.Rows[ev.RowIndex].Cells["Doc_Path"].Value?.ToString() ?? "";
                if (System.IO.File.Exists(path))
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
            };

            panelSched.Controls.Add(dgvDoc);

            // ── Danh sách PO đã in Request ──
            panelPrintHistory = P(tabPO, 5, 317 + 200 + 5, 0, 0, Color.White);
            panelPrintHistory.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom;
            Lbl(panelPrintHistory, "🖨  DANH SÁCH PO ĐÃ IN REQUEST", 8, 5, 350, 20, true, Color.FromArgb(0, 150, 100));

            // ── Toolbar lọc theo thời gian ──
            Lbl(panelPrintHistory, "Từ:", 8, 30, 25, 20);
            _phDateFrom = new DateTimePicker
            {
                Location = new Point(33, 27),
                Size = new Size(115, 24),
                Font = new Font("Segoe UI", 9),
                Format = DateTimePickerFormat.Short,
                Value = DateTime.Today.AddYears(-2)   // mặc định 2 năm để không mất dữ liệu cũ
            };
            panelPrintHistory.Controls.Add(_phDateFrom);

            Lbl(panelPrintHistory, "Đến:", 155, 30, 30, 20);
            _phDateTo = new DateTimePicker
            {
                Location = new Point(185, 27),
                Size = new Size(115, 24),
                Font = new Font("Segoe UI", 9),
                Format = DateTimePickerFormat.Short,
                Value = DateTime.Today
            };
            panelPrintHistory.Controls.Add(_phDateTo);

            var btnPhSearch = Btn("🔍 Lọc", Color.FromArgb(0, 120, 212), 308, 26, 70, 26);
            btnPhSearch.Click += (s, ev) => LoadPrintHistory(_phDateFrom.Value.Date, _phDateTo.Value.Date.AddDays(1).AddSeconds(-1));
            panelPrintHistory.Controls.Add(btnPhSearch);

            // Nút "Tất cả" — load toàn bộ lịch sử không giới hạn ngày
            var btnPhAll = Btn("📋 Tất cả", Color.FromArgb(0, 150, 100), 384, 26, 85, 26);
            btnPhAll.Click += (s, ev) =>
            {
                _phDateFrom.Value = new DateTime(2000, 1, 1);
                _phDateTo.Value = DateTime.Today;
                LoadPrintHistory(new DateTime(2000, 1, 1), DateTime.Today.AddDays(1).AddSeconds(-1));
            };
            panelPrintHistory.Controls.Add(btnPhAll);

            var btnPhReset = Btn("✖ Reset", Color.FromArgb(108, 117, 125), 475, 26, 70, 26);
            btnPhReset.Click += (s, ev) =>
            {
                _phDateFrom.Value = DateTime.Today.AddYears(-2);
                _phDateTo.Value = DateTime.Today;
                _txtPhNCC.Text = "";
                LoadPrintHistory(_phDateFrom.Value.Date, _phDateTo.Value.Date.AddDays(1).AddSeconds(-1));
            };
            panelPrintHistory.Controls.Add(btnPhReset);

            // ── Bộ lọc NCC ──
            Lbl(panelPrintHistory, "NCC:", 553, 30, 30, 20);
            _txtPhNCC = new TextBox
            {
                Location = new Point(583, 27),
                Size = new Size(160, 24),
                Font = new Font("Segoe UI", 9),
                PlaceholderText = "Tìm theo NCC..."
            };
            _txtPhNCC.TextChanged += (s, ev) =>
                LoadPrintHistory(_phDateFrom.Value.Date, _phDateTo.Value.Date.AddDays(1).AddSeconds(-1));
            panelPrintHistory.Controls.Add(_txtPhNCC);

            var btnPhDel = Btn("🗑 Xóa dòng", Color.FromArgb(220, 53, 69), 751, 26, 100, 26);
            btnPhDel.Click += BtnDeletePrintHistory_Click;
            panelPrintHistory.Controls.Add(btnPhDel);

            // ── Grid — top=58 để có chỗ cho toolbar ──
            dgvPrintHistory = Grid(panelPrintHistory, 58, 0);
            dgvPrintHistory.ReadOnly = true;
            dgvPrintHistory.Columns.Clear();
            dgvPrintHistory.AutoGenerateColumns = false;
            // Cột PH_ID ẩn để xóa DB
            dgvPrintHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "PH_ID", HeaderText = "ID", Visible = false });
            dgvPrintHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "PH_PONo", HeaderText = "PO No", Width = 150, ReadOnly = true });
            dgvPrintHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "PH_Supp", HeaderText = "NCC", Width = 100, ReadOnly = true });
            dgvPrintHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "PH_Project", HeaderText = "Dự án", Width = 150, ReadOnly = true });
            dgvPrintHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "PH_Dot", HeaderText = "Đợt in", Width = 60, ReadOnly = true });
            dgvPrintHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "PH_Net", HeaderText = "Số tiền (Net)", Width = 120, ReadOnly = true });
            dgvPrintHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "PH_Vat", HeaderText = "VAT", Width = 100, ReadOnly = true });
            dgvPrintHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "PH_Total", HeaderText = "Tổng sau VAT", Width = 120, ReadOnly = true });
            dgvPrintHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "PH_Date", HeaderText = "Ngày in ▼", Width = 130, ReadOnly = true });
            foreach (DataGridViewColumn col in dgvPrintHistory.Columns)
                col.SortMode = DataGridViewColumnSortMode.Programmatic;
            dgvPrintHistory.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 150, 100);
            dgvPrintHistory.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPrintHistory.EnableHeadersVisualStyles = false;
            dgvPrintHistory.CellFormatting += (s, ev) =>
            {
                if (ev.RowIndex < 0) return;
                string col = dgvPrintHistory.Columns[ev.ColumnIndex].Name;
                if (col == "PH_Net" || col == "PH_Vat" || col == "PH_Total")
                    ev.CellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            };

            panelHist = P(tabPO, 0, 317, 0, 200, Color.White);
            panelHist.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Lbl(panelHist, "📋  PAYMENT REQUEST PROGRESSING", 8, 5, 400, 20, true, Color.FromArgb(102, 51, 153));

            var bHistPaid = Btn("📋 History Paid", Color.FromArgb(40, 167, 69), 8, 28, 130, 24);
            bHistPaid.Click += (s, e) => ShowHistoryPaidPopup();
            panelHist.Controls.Add(bHistPaid);

            dgvHistory = Grid(panelHist, 57, 0);
            dgvHistory.ColumnHeadersHeight = 50;
            dgvHistory.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            BuildHistCols();
            dgvHistory.CellDoubleClick += DgvHistory_CellDoubleClick;

            // ── dgvPaid: khởi tạo standalone, dùng trong popup History Paid ──
            _paidFrom = new DateTimePicker { Format = DateTimePickerFormat.Short, Value = DateTime.Today.AddMonths(-3) };
            _paidTo = new DateTimePicker { Format = DateTimePickerFormat.Short, Value = DateTime.Today };

            dgvPaid = new DataGridView
            {
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            dgvPaid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(40, 167, 69);
            dgvPaid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPaid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvPaid.EnableHeadersVisualStyles = false;
            dgvPaid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(235, 255, 235);
            dgvPaid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
            dgvPaid.DefaultCellStyle.SelectionForeColor = Color.Black;

            dgvPaid.Columns.Add(new DataGridViewTextBoxColumn { Name = "HP_ID", HeaderText = "ID", Visible = false });
            dgvPaid.Columns.Add(new DataGridViewTextBoxColumn { Name = "HP_PONo", HeaderText = "PO No", Width = 140 });
            dgvPaid.Columns.Add(new DataGridViewTextBoxColumn { Name = "HP_Total", HeaderText = "Tổng sau VAT", Width = 120 });
            dgvPaid.Columns.Add(new DataGridViewTextBoxColumn { Name = "HP_Note", HeaderText = "Ghi chú", Width = 150 });
            dgvPaid.Columns.Add(new DataGridViewTextBoxColumn { Name = "HP_INV", HeaderText = "INV", Width = 160 });
            dgvPaid.Columns.Add(new DataGridViewTextBoxColumn { Name = "HP_Delivery", HeaderText = "Delivery Note", Width = 160 });
            dgvPaid.Columns.Add(new DataGridViewTextBoxColumn { Name = "HP_PaidAt", HeaderText = "Thời gian TT", Width = 140 });

            dgvPaid.CellFormatting += (s, ev) =>
            {
                if (ev.RowIndex < 0) return;
                if (dgvPaid.Columns[ev.ColumnIndex].Name == "HP_Total")
                    ev.CellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            };
        }

        private void BuildPOGridCols()
        {
            dgvPO.Columns.Clear();
            dgvPO.AutoGenerateColumns = false;
            dgvPO.Columns.Add(new DataGridViewTextBoxColumn { Name = "ID", DataPropertyName = "ID", Visible = false });
            dgvPO.Columns.Add(new DataGridViewTextBoxColumn { Name = "PO_No", DataPropertyName = "PO_No", HeaderText = "PO No", Width = 200, ReadOnly = true });
            dgvPO.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ngay_PO", DataPropertyName = "Ngay_PO", HeaderText = "Ngày PO", Width = 85, ReadOnly = true });
            dgvPO.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ten_DA", DataPropertyName = "Ten_DA", HeaderText = "Dự án", Width = 160, ReadOnly = true });
            dgvPO.Columns.Add(new DataGridViewTextBoxColumn { Name = "NCC", DataPropertyName = "NCC", HeaderText = "Nhà CC", Width = 130, ReadOnly = true });
            dgvPO.Columns.Add(new DataGridViewTextBoxColumn { Name = "Tong_PO", DataPropertyName = "Tong_PO", HeaderText = "Tổng PO", Width = 100, ReadOnly = true });
            dgvPO.Columns.Add(new DataGridViewTextBoxColumn { Name = "Da_TT", DataPropertyName = "Da_TT", HeaderText = "Đã TT", Width = 100, ReadOnly = true });
            dgvPO.Columns.Add(new DataGridViewTextBoxColumn { Name = "Con_No", DataPropertyName = "Con_No", HeaderText = "Còn nợ", Width = 100, ReadOnly = true });
            dgvPO.Columns.Add(new DataGridViewTextBoxColumn { Name = "Pct", DataPropertyName = "Pct", HeaderText = "%", Width = 55, ReadOnly = true });
            dgvPO.Columns.Add(new DataGridViewTextBoxColumn { Name = "TT_Status", DataPropertyName = "TT_Status", HeaderText = "Trạng thái", Width = 110, ReadOnly = true });
            dgvPO.Columns.Add(new DataGridViewTextBoxColumn { Name = "Den_Han", DataPropertyName = "Den_Han", HeaderText = "Đến hạn", Width = 85, ReadOnly = true });
            dgvPO.Columns.Add(new DataGridViewTextBoxColumn { Name = "Qua_Han", DataPropertyName = "Qua_Han", HeaderText = "Quá hạn", Width = 70, ReadOnly = true });
            // ── Cột kế hoạch TT từng đợt (tối đa 5 đợt) ──
            for (int i = 1; i <= 5; i++)
            {
                dgvPO.Columns.Add(new DataGridViewTextBoxColumn { Name = $"Dot{i}_Amount", DataPropertyName = $"Dot{i}_Amount", HeaderText = $"Đợt {i} - Số tiền", Width = 110, ReadOnly = true });
                dgvPO.Columns.Add(new DataGridViewTextBoxColumn { Name = $"Dot{i}_Status", DataPropertyName = $"Dot{i}_Status", HeaderText = $"Đợt {i} - T.Thái", Width = 95, ReadOnly = true });
            }
        }

        private void BuildSchedCols()
        {
            dgvSchedule.Columns.Clear();
            dgvSchedule.Columns.Add(new DataGridViewTextBoxColumn { Name = "S_ID", Visible = false });
            dgvSchedule.Columns.Add(new DataGridViewTextBoxColumn { Name = "Dot_TT", HeaderText = "Đợt", Width = 42 });
            var cboMethod = new DataGridViewComboBoxColumn { Name = "Pay_Method", HeaderText = "Kiểu TT", Width = 100, FlatStyle = FlatStyle.Flat };
            cboMethod.Items.AddRange(new[] { "Full", "Partial", "Percent", "ByDelivery" });
            dgvSchedule.Columns.Add(cboMethod);
            dgvSchedule.Columns.Add(new DataGridViewTextBoxColumn { Name = "Payment_Type", HeaderText = "Hình thức", Width = 110 });
            dgvSchedule.Columns.Add(new DataGridViewTextBoxColumn { Name = "Percent_TT", HeaderText = "%", Width = 48 });
            dgvSchedule.Columns.Add(new DataGridViewTextBoxColumn { Name = "Amount_Plan", HeaderText = "Số tiền KH", Width = 105 });
            dgvSchedule.Columns.Add(new DataGridViewTextBoxColumn { Name = "Due_Date", HeaderText = "Đến hạn 📅", Width = 105 });
            dgvSchedule.Columns.Add(new DataGridViewTextBoxColumn { Name = "Description", HeaderText = "Điều kiện", FillWeight = 100 });
            var cboStatus = new DataGridViewComboBoxColumn { Name = "S_Status", HeaderText = "Trạng thái", Width = 100, FlatStyle = FlatStyle.Flat };
            cboStatus.Items.AddRange(new[] { "Chưa TT", "Một phần", "Đã TT đủ" });
            dgvSchedule.Columns.Add(cboStatus);

            // ── DateTimePicker ẩn — hiện khi click vào ô Due_Date ──
            _schedDtp = new DateTimePicker
            {
                Format = DateTimePickerFormat.Short,
                Font = new Font("Segoe UI", 9),
                Visible = false,
                MinDate = new DateTime(2000, 1, 1)
            };

            // Thêm DTP vào panel cha của dgvSchedule
            panelSched.Controls.Add(_schedDtp);
            _schedDtp.BringToFront();

            // Hiện DTP khi click vào cột Due_Date
            dgvSchedule.CellClick += (s, ev) =>
            {
                if (ev.RowIndex < 0 || dgvSchedule.Columns[ev.ColumnIndex].Name != "Due_Date") return;

                _schedDtpRow = ev.RowIndex;

                // Parse giá trị hiện tại
                string cur = dgvSchedule.Rows[ev.RowIndex].Cells["Due_Date"].Value?.ToString() ?? "";
                _schedDtp.Value = DateTime.TryParseExact(cur, "dd/MM/yyyy",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out DateTime parsed)
                    ? parsed : DateTime.Today;

                // Tính tọa độ của cell trong panelSched
                var cellRect = dgvSchedule.GetCellDisplayRectangle(ev.ColumnIndex, ev.RowIndex, true);
                var cellPos = dgvSchedule.PointToScreen(new Point(cellRect.Left, cellRect.Top));
                var panelPos = panelSched.PointToClient(cellPos);

                _schedDtp.Location = new Point(panelPos.X, panelPos.Y);
                _schedDtp.Width = cellRect.Width;
                _schedDtp.Height = cellRect.Height;
                _schedDtp.Visible = true;
                _schedDtp.Focus();
            };

            // Ẩn DTP khi click ra ngoài
            dgvSchedule.CellClick += (s, ev) =>
            {
                if (ev.ColumnIndex >= 0 && dgvSchedule.Columns[ev.ColumnIndex].Name != "Due_Date")
                {
                    CommitSchedDtp();
                    _schedDtp.Visible = false;
                }
            };
            dgvSchedule.Scroll += (s, ev) => { CommitSchedDtp(); _schedDtp.Visible = false; };

            // Khi chọn ngày → ghi vào cell ngay lập tức
            _schedDtp.ValueChanged += (s, ev) =>
            {
                if (_schedDtpRow < 0 || !_schedDtp.Visible) return;
                dgvSchedule.Rows[_schedDtpRow].Cells["Due_Date"].Value = _schedDtp.Value.ToString("dd/MM/yyyy");
            };

            _schedDtp.Leave += (s, ev) => { CommitSchedDtp(); _schedDtp.Visible = false; };

            _schedDtp.KeyDown += (s, ev) =>
            {
                if (ev.KeyCode == Keys.Escape || ev.KeyCode == Keys.Enter)
                {
                    CommitSchedDtp();
                    _schedDtp.Visible = false;
                    dgvSchedule.Focus();
                }
            };
        }

        private void BuildHistCols()
        {
            dgvHistory.Columns.Clear();
            dgvHistory.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvHistory.AllowUserToResizeColumns = true;
            dgvHistory.ColumnHeadersHeight = 35;
            dgvHistory.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

            // Cột trái cố định
            dgvHistory.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "H_PONo", HeaderText = "PO No",
                Width = 160, AutoSizeMode = DataGridViewAutoSizeColumnMode.None, ReadOnly = true
            });
            // Cột giữa — fill theo tỷ lệ
            dgvHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "H_PreTax",   HeaderText = "Trước thuế",    FillWeight = 13, ReadOnly = true });
            dgvHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "H_Total",    HeaderText = "Thành tiền",    FillWeight = 15, ReadOnly = true });
            dgvHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "H_ECStatus", HeaderText = "EC status",     FillWeight =  9, ReadOnly = true });
            dgvHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "H_Paid",     HeaderText = "Đã thanh toán", FillWeight = 13, ReadOnly = true });
            dgvHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "H_PaidDate", HeaderText = "Paid Date",     FillWeight = 11, ReadOnly = true });
            dgvHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "H_Remain",   HeaderText = "Còn lại",       FillWeight = 13, ReadOnly = true });
            dgvHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "H_SuppShort",HeaderText = "Tên NCC",       FillWeight = 15, ReadOnly = true });
            // Cột phải cố định
            dgvHistory.Columns.Add(new DataGridViewButtonColumn
            {
                Name = "H_Print",
                HeaderText = "In thanh toán",
                Width = 90,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                FlatStyle = FlatStyle.Flat,
                Text = "🖨 In",
                UseColumnTextForButtonValue = true
            });

            dgvHistory.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(102, 51, 153);
            dgvHistory.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvHistory.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvHistory.EnableHeadersVisualStyles = false;
            dgvHistory.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 240, 255);

            dgvHistory.CellFormatting += (s, ev) =>
            {
                if (ev.RowIndex < 0) return;
                string col = dgvHistory.Columns[ev.ColumnIndex].Name;
                if (col == "H_PreTax" || col == "H_Total" || col == "H_Paid" || col == "H_Remain")
                    ev.CellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                if (col == "H_ECStatus")
                {
                    string ecv = ev.Value?.ToString() ?? "";
                    if (ecv == "상신") // 상신
                    {
                        ev.CellStyle.ForeColor = Color.Red;
                    }
                    else if (ecv == "종결") // 종결
                    {
                        ev.CellStyle.ForeColor = Color.FromArgb(180, 0, 0);
                        ev.CellStyle.BackColor = Color.FromArgb(255, 220, 220);
                        ev.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    }
                    else if (ecv == "대기중") // 대기중
                    {
                        ev.CellStyle.ForeColor = Color.Red;
                        ev.CellStyle.BackColor = Color.FromArgb(255, 200, 100);
                    }
                }
                if (col == "H_Paid" && !string.IsNullOrEmpty(ev.Value?.ToString()))
                {
                    ev.CellStyle.BackColor = Color.FromArgb(230, 255, 230);
                    ev.CellStyle.ForeColor = Color.FromArgb(40, 130, 40);
                    ev.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                }
                if (col == "H_Remain")
                {
                    if (decimal.TryParse((ev.Value?.ToString() ?? "").Replace(",", ""), out decimal remain) && remain <= 0)
                    {
                        ev.CellStyle.ForeColor = Color.FromArgb(40, 167, 69);
                        ev.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    }
                }
            };

             dgvHistory.CellContentClick += (s, ev) =>
             {
                 if (ev.RowIndex < 0) return;
                 if (dgvHistory.Columns[ev.ColumnIndex].Name != "H_Print") return;
                 string poNo = dgvHistory.Rows[ev.RowIndex].Cells["H_PONo"].Value?.ToString() ?? "";
                 var po = _poSummaries.Find(p => (p.PONo ?? "") == poNo);
                 if (po == null) { Warn($"Không tìm thấy PO {poNo} trong danh sách!"); return; }
                 _selectedPO_ID = po.PO_ID;
                 foreach (DataGridViewRow row in dgvPO.Rows)
                 {
                     if (Convert.ToInt32(row.Cells["ID"].Value ?? 0) == po.PO_ID)
                     { dgvPO.ClearSelection(); row.Selected = true; break; }
                 }
                 ShowPrintHistoryAndOptions(poNo);
             };
        }
        private void BuildTabDebt()
        {
            // ── Hàng 1: Date range + quick-date buttons + NCC ──
            var pF = P(tabDebt, 5, 5, 0, 92, Color.White);
            pF.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            // Row 1
            Lbl(pF, "Từ ngày:", 6, 13, 62, 20);
            dtpFrom = new DateTimePicker
            {
                Location = new Point(68, 9),
                Size = new Size(115, 26),
                Font = new Font("Segoe UI", 9),
                Format = DateTimePickerFormat.Short,
                Value = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1)
            };
            pF.Controls.Add(dtpFrom);

            Lbl(pF, "Đến:", 190, 13, 38, 20);
            dtpTo = new DateTimePicker
            {
                Location = new Point(228, 9),
                Size = new Size(115, 26),
                Font = new Font("Segoe UI", 9),
                Format = DateTimePickerFormat.Short,
                Value = DateTime.Today
            };
            pF.Controls.Add(dtpTo);

            // Quick date buttons
            int qx = 354;
            foreach (var (label, color, tag) in new (string, Color, string)[]
            {
                ("Tháng này",   Color.FromArgb(0,120,212),   "thisMonth"),
                ("Tháng trước", Color.FromArgb(100,100,180), "lastMonth"),
                ("Quý này",     Color.FromArgb(0,150,100),   "thisQuarter"),
                ("Năm nay",     Color.FromArgb(80,80,80),    "thisYear"),
            })
            {
                var b = new Button
                {
                    Text = label, Tag = tag,
                    Location = new Point(qx, 8),
                    Size = new Size(88, 26),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = color,
                    ForeColor = Color.White,
                    Font = new Font("Segoe UI", 8, FontStyle.Bold),
                    Cursor = Cursors.Hand
                };
                b.FlatAppearance.BorderSize = 0;
                b.Click += (s, _) =>
                {
                    var today = DateTime.Today;
                    switch (((Button)s).Tag?.ToString())
                    {
                        case "thisMonth":
                            dtpFrom.Value = new DateTime(today.Year, today.Month, 1);
                            dtpTo.Value = today; break;
                        case "lastMonth":
                            var lm = today.AddMonths(-1);
                            dtpFrom.Value = new DateTime(lm.Year, lm.Month, 1);
                            dtpTo.Value = new DateTime(lm.Year, lm.Month, DateTime.DaysInMonth(lm.Year, lm.Month)); break;
                        case "thisQuarter":
                            int q = (today.Month - 1) / 3;
                            dtpFrom.Value = new DateTime(today.Year, q * 3 + 1, 1);
                            dtpTo.Value = today; break;
                        case "thisYear":
                            dtpFrom.Value = new DateTime(today.Year, 1, 1);
                            dtpTo.Value = today; break;
                    }
                };
                pF.Controls.Add(b);
                qx += 92;
            }

            Lbl(pF, "NCC:", qx + 4, 13, 38, 20);
            cboSuppFilter = new ComboBox
            {
                Location = new Point(qx + 42, 9),
                Size = new Size(200, 26),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboSuppFilter.Items.Add("Tất cả nhà cung cấp");
            cboSuppFilter.SelectedIndex = 0;
            pF.Controls.Add(cboSuppFilter);

            var bView = Btn("🔍 Xem báo cáo", Color.FromArgb(0, 120, 212), qx + 250, 8, 140, 28);
            bView.Click += BtnViewDebt_Click;
            pF.Controls.Add(bView);

            btnExportDebt = Btn("📥 Xuất Excel", Color.FromArgb(0, 150, 100), qx + 398, 8, 120, 28);
            btnExportDebt.Click += BtnExportDebt_Click;
            pF.Controls.Add(btnExportDebt);

            // Row 2: Status filter + overdue checkbox + search
            Lbl(pF, "Trạng thái:", 6, 55, 78, 20);
            cboDebtStatus = new ComboBox
            {
                Location = new Point(84, 51),
                Size = new Size(155, 26),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboDebtStatus.Items.AddRange(new object[] { "Tất cả", "Pending", "Thanh toán 1 phần", "Đã thanh toán", "⚠ Quá hạn" });
            cboDebtStatus.SelectedIndex = 0;
            cboDebtStatus.SelectedIndexChanged += (s, e) => FilterAndBindDebt();
            pF.Controls.Add(cboDebtStatus);

            chkOverdueOnly = new CheckBox
            {
                Text = "⚠ Chỉ quá hạn",
                Location = new Point(250, 52),
                Size = new Size(130, 24),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(220, 53, 69),
                Cursor = Cursors.Hand
            };
            chkOverdueOnly.CheckedChanged += (s, e) => FilterAndBindDebt();
            pF.Controls.Add(chkOverdueOnly);

            Lbl(pF, "Tìm PO / Dự án:", 392, 55, 118, 20);
            txtDebtSearch = new TextBox
            {
                Location = new Point(510, 51),
                Size = new Size(220, 26),
                Font = new Font("Segoe UI", 9),
                PlaceholderText = "PO No hoặc tên dự án..."
            };
            txtDebtSearch.TextChanged += (s, e) => FilterAndBindDebt();
            pF.Controls.Add(txtDebtSearch);

            // ── Cards tổng kết ──
            _pDebtCards = P(tabDebt, 5, 102, 0, 72, Color.White);
            _pDebtCards.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            lblSumValue   = Card(_pDebtCards, 10,  "Tổng giá trị PO", Color.FromArgb(0, 120, 212));
            lblSumPaid    = Card(_pDebtCards, 225, "Đã thanh toán",   Color.FromArgb(40, 167, 69));
            lblSumDebt    = Card(_pDebtCards, 440, "Còn nợ",          Color.FromArgb(255, 140, 0));
            lblSumOverdue = Card(_pDebtCards, 655, "Quá hạn (PO)",    Color.FromArgb(220, 53, 69));

            // ── Panel NCC (trái) ──
            _pNCC = P(tabDebt, 5, 179, 380, 0, Color.White);
            var pNCC = _pNCC;
            pNCC.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom;
            Lbl(pNCC, "TỔNG HỢP THEO NHÀ CUNG CẤP", 8, 5, 360, 20, true, Color.FromArgb(0, 120, 212));
            dgvDebtSupp = Grid(pNCC, 28, 0);
            dgvDebtSupp.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom | AnchorStyles.Right;
            dgvDebtSupp.ColumnHeadersHeight = 50;
            dgvDebtSupp.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvDebtSupp.SelectionChanged += DgvDebtSupp_SelectionChanged;
            dgvDebtSupp.CellFormatting += DgvDebtSupp_CellFormatting;
            BuildDebtSuppCols();

            // ── Panel Chi tiết PO (phải) ──
            _pDet = P(tabDebt, 390, 179, 0, 0, Color.White);
            var pDet = _pDet;
            pDet.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            Lbl(pDet, "CHI TIẾT TỪNG ĐƠN PO", 8, 5, 400, 20, true, Color.FromArgb(0, 120, 212));
            dgvDebtDetail = Grid(pDet, 28, 0);
            dgvDebtDetail.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            dgvDebtDetail.ColumnHeadersHeight = 50;
            dgvDebtDetail.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvDebtDetail.CellFormatting += DgvDebtDetail_CellFormatting;
            dgvDebtDetail.CellDoubleClick += DgvDebtDetail_CellDoubleClick;
            BuildDebtDetailCols();
        }

        private void BuildDebtSuppCols()
        {
            dgvDebtSupp.Columns.Clear();
            dgvDebtSupp.Columns.Add(new DataGridViewTextBoxColumn { Name = "D_SuppID",  Visible = false });
            dgvDebtSupp.Columns.Add(new DataGridViewTextBoxColumn { Name = "D_Name",    HeaderText = "Nhà cung cấp",   Width = 200, ReadOnly = true });
            dgvDebtSupp.Columns.Add(new DataGridViewTextBoxColumn { Name = "D_TotalPO", HeaderText = "Số PO",          Width = 50,  ReadOnly = true });
            dgvDebtSupp.Columns.Add(new DataGridViewTextBoxColumn { Name = "D_Value",   HeaderText = "Tổng PO",        Width = 105, ReadOnly = true });
            dgvDebtSupp.Columns.Add(new DataGridViewTextBoxColumn { Name = "D_Paid",    HeaderText = "Đã thanh toán",  Width = 105, ReadOnly = true });
            dgvDebtSupp.Columns.Add(new DataGridViewTextBoxColumn { Name = "D_Debt",    HeaderText = "Còn nợ",         Width = 105, ReadOnly = true });
            dgvDebtSupp.Columns.Add(new DataGridViewTextBoxColumn { Name = "D_Overdue", HeaderText = "Quá hạn\n(PO)", Width = 65,  ReadOnly = true });
        }

        private void BuildDebtDetailCols()
        {
            dgvDebtDetail.Columns.Clear();
            dgvDebtDetail.Columns.Add(new DataGridViewTextBoxColumn { Name = "DD_POID",    Visible = false });
            dgvDebtDetail.Columns.Add(new DataGridViewTextBoxColumn { Name = "DD_PONo",    HeaderText = "PO No",        Width = 110, ReadOnly = true });
            dgvDebtDetail.Columns.Add(new DataGridViewTextBoxColumn { Name = "DD_Project", HeaderText = "Mã dự án",     FillWeight = 100, ReadOnly = true });
            dgvDebtDetail.Columns.Add(new DataGridViewTextBoxColumn { Name = "DD_PODate",  HeaderText = "Ngày PO",      Width = 85,  ReadOnly = true });
            dgvDebtDetail.Columns.Add(new DataGridViewTextBoxColumn { Name = "DD_Total",   HeaderText = "Giá trị PO",   Width = 100, ReadOnly = true });
            dgvDebtDetail.Columns.Add(new DataGridViewTextBoxColumn { Name = "DD_Before",  HeaderText = "TT trước kỳ", Width = 100, ReadOnly = true });
            dgvDebtDetail.Columns.Add(new DataGridViewTextBoxColumn { Name = "DD_InRange", HeaderText = "TT trong kỳ", Width = 100, ReadOnly = true });
            dgvDebtDetail.Columns.Add(new DataGridViewTextBoxColumn { Name = "DD_Remain",  HeaderText = "Còn nợ",       Width = 100, ReadOnly = true });
            dgvDebtDetail.Columns.Add(new DataGridViewTextBoxColumn { Name = "DD_Percent", HeaderText = "% TT",         Width = 60,  ReadOnly = true });
            dgvDebtDetail.Columns.Add(new DataGridViewTextBoxColumn { Name = "DD_Status",  HeaderText = "Trạng thái",   Width = 95,  ReadOnly = true });
            dgvDebtDetail.Columns.Add(new DataGridViewTextBoxColumn { Name = "DD_Due", HeaderText = "Đến hạn", Width = 85, ReadOnly = true });
        }

        // Async loading với parallel queries để tăng tốc
        private async System.Threading.Tasks.Task LoadDataAsync()
        {
            btnRefreshPO.Enabled = false;

            try
            {
                // Chạy 3 nguồn dữ liệu song song
                await System.Threading.Tasks.Task.WhenAll(
                    LoadSuppliersAsync(),
                    LoadPOSummaryAsync(),
                    LoadPrintHistoryAsync(DateTime.Today.AddYears(-2), DateTime.Today.AddDays(1).AddSeconds(-1))
                );
            }
            catch (Exception ex)
            {
                Err($"Lỗi tải dữ liệu: {ex.Message}");
            }
            finally
            {
                btnRefreshPO.Enabled = true;
                btnRefreshPO.Text = "🔄 Làm mới";
            }
        }

        // Load suppliers async
        private async System.Threading.Tasks.Task LoadSuppliersAsync()
        {
            await System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    _allSuppliers = _suppSvc.GetAll();
                }
                catch { _allSuppliers = new List<Supplier>(); }
            }).ConfigureAwait(false);

            // Update UI on main thread
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() =>
                {
                    cboSuppFilter.Items.Clear();
                    cboSuppFilter.Items.Add("Tất cả nhà cung cấp");
                    foreach (var s in _allSuppliers)
                        cboSuppFilter.Items.Add(s.Company_Name ?? s.Supplier_Name);
                    if (cboSuppFilter.Items.Count > 0)
                        cboSuppFilter.SelectedIndex = 0;
                }));
            }
            else
            {
                cboSuppFilter.Items.Clear();
                cboSuppFilter.Items.Add("Tất cả nhà cung cấp");
                foreach (var s in _allSuppliers)
                    cboSuppFilter.Items.Add(s.Company_Name ?? s.Supplier_Name);
                if (cboSuppFilter.Items.Count > 0)
                    cboSuppFilter.SelectedIndex = 0;
            }
        }

        // Load PO summary async
        private async System.Threading.Tasks.Task LoadPOSummaryAsync()
        {
            var result = await System.Threading.Tasks.Task.Run(() =>
            {
                var summaries = _svc.GetPOSummaries();
                var allScheds = _svc.GetAllSchedules();
                var cache = allScheds
                    .GroupBy(s => s.PO_ID)
                    .ToDictionary(g => g.Key, g => g.ToList());
                var zaloPaid = LoadZaloPaidCache();
                return (summaries, cache, zaloPaid);
            }).ConfigureAwait(false);

            _poSummaries = result.summaries;
            _allSchedulesCache = result.cache;
            _zaloPaidCache = result.zaloPaid;

            // Update grid on main thread
            if (this.InvokeRequired)
                this.Invoke(new Action(FilterAndBind));
            else
                FilterAndBind();
        }

        // Load print history async
        private async System.Threading.Tasks.Task LoadPrintHistoryAsync(DateTime from, DateTime to)
        {
            await System.Threading.Tasks.Task.Run(() =>
            {
                if (this.InvokeRequired)
                    this.Invoke(new Action(() => LoadPrintHistory(from, to)));
                else
                    LoadPrintHistory(from, to);
            }).ConfigureAwait(false);
        }

        // Synchronous wrapper for backward compatibility
        private void LoadData()
        {
            try
            {
                _allSuppliers = _suppSvc.GetAll();
                cboSuppFilter.Items.Clear();
                cboSuppFilter.Items.Add("Tất cả nhà cung cấp");
                foreach (var s in _allSuppliers)
                    cboSuppFilter.Items.Add(s.Company_Name ?? s.Supplier_Name);
                cboSuppFilter.SelectedIndex = 0;
            }
            catch { }
            LoadPOSummary();
            LoadPrintHistory(DateTime.Today.AddYears(-2), DateTime.Today.AddDays(1).AddSeconds(-1));
        }

        private async void LoadPOSummary()
        {
            var toastRefresh = ToastHelper.Attach(this);
            toastRefresh.Show("⏳ Đang tải dữ liệu, vui lòng chờ...");
            btnRefreshPO.Enabled = false;
            try
            {
                var result = await System.Threading.Tasks.Task.Run(() =>
                {
                    var summaries = _svc.GetPOSummaries();
                    var allScheds = _svc.GetAllSchedules();
                    var cache = allScheds
                        .GroupBy(s => s.PO_ID)
                        .ToDictionary(g => g.Key, g => g.ToList());
                    var zaloPaid = LoadZaloPaidCache();
                    return (summaries, cache, zaloPaid);
                });
                _poSummaries = result.summaries;
                _allSchedulesCache = result.cache;
                _zaloPaidCache = result.zaloPaid;
                FilterAndBind();
            }
            catch (Exception ex) { Err(ex.Message); }
            finally
            {
                toastRefresh.Hide();
                btnRefreshPO.Enabled = true;
            }
        }

        private void FilterAndBind()
        {
            string kw = txtSearchPO.Text.Trim();
            string status = cboStatusFilter.SelectedItem?.ToString() ?? "Tất cả";

            var list = _poSummaries;
            if (!string.IsNullOrEmpty(kw))
            {
                list = list.FindAll(p =>
                    (p.PONo ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                    (p.Project_Name ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                    (p.Supplier_Name ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase));
            }

            var displayList = list.ConvertAll(p =>
            {
                decimal totalPO = p.Total_PO_Amount;

                // Lấy số tiền đã thanh toán (sau thuế) từ Zalo_PaymentImport (cột "Đã thanh toán")
                string poKey = (p.PONo ?? "").ToUpperInvariant()
                    .Replace(" ", "").Replace("\t", "").Replace("\r", "").Replace("\n", "");
                _zaloPaidCache.TryGetValue(poKey, out decimal totalPaid);

                decimal remain = totalPO - totalPaid;
                if (remain < 0) remain = 0;

                decimal pct = totalPO > 0 ? (totalPaid / totalPO) * 100 : 0;
                if (pct > 100) pct = 100;

                string realStatus = GetZaloStatus(p.PONo, totalPO);

                bool isNew = p.PO_Date.HasValue && (DateTime.Now - p.PO_Date.Value).TotalDays <= 3;
                string poDisplayObj = isNew ? $"🔥 {p.PONo} (Mới)" : p.PONo;

                // ── Schedules từng đợt từ cache ──
                var scheds = _allSchedulesCache.ContainsKey(p.PO_ID)
                    ? _allSchedulesCache[p.PO_ID]
                    : new List<PaymentSchedule>();
                string d1a = "", d1s = "", d2a = "", d2s = "", d3a = "", d3s = "", d4a = "", d4s = "", d5a = "", d5s = "";
                for (int idx = 0; idx < scheds.Count && idx < 5; idx++)
                {
                    string a = FormatAmt(scheds[idx].Amount_Plan);
                    string t = scheds[idx].Status ?? "Chưa TT";
                    switch (idx)
                    {
                        case 0: d1a = a; d1s = t; break;
                        case 1: d2a = a; d2s = t; break;
                        case 2: d3a = a; d3s = t; break;
                        case 3: d4a = a; d4s = t; break;
                        case 4: d5a = a; d5s = t; break;
                    }
                }

                return new
                {
                    ID = p.PO_ID,
                    PO_No = poDisplayObj,
                    Ngay_PO = p.PO_Date.HasValue ? p.PO_Date.Value.ToString("dd/MM/yyyy") : "",
                    Ten_DA = p.Project_Name,
                    NCC = p.Supplier_Name,
                    Tong_PO = FormatAmt(totalPO),
                    Da_TT = FormatAmt(totalPaid),
                    Con_No = FormatAmt(remain),
                    Pct = pct.ToString("N1") + "%",
                    TT_Status = realStatus,
                    Den_Han = p.Next_Due_Date.HasValue ? p.Next_Due_Date.Value.ToString("dd/MM/yyyy") : "—",
                    Qua_Han = p.Is_Overdue ? "⚠ Quá hạn" : "",
                    Is_Overdue = p.Is_Overdue,
                    Dot1_Amount = d1a,
                    Dot1_Status = d1s,
                    Dot2_Amount = d2a,
                    Dot2_Status = d2s,
                    Dot3_Amount = d3a,
                    Dot3_Status = d3s,
                    Dot4_Amount = d4a,
                    Dot4_Status = d4s,
                    Dot5_Amount = d5a,
                    Dot5_Status = d5s,
                };
            });

            if (status == "⚠ Quá hạn")
                displayList = displayList.FindAll(p => p.Is_Overdue);
            else if (status != "Tất cả")
                displayList = displayList.FindAll(p => p.TT_Status == status);

            // Grid binding suspension để tránh event firing quá nhiều
            dgvPO.SuspendLayout();
            try
            {
                // Tách event handlers trước khi bind
                dgvPO.SelectionChanged -= DgvPO_SelectionChanged;
                dgvPO.CellFormatting -= DgvPO_CellFormatting;

                // Bind data
                dgvPO.DataSource = displayList;

                // Gắn lại event handlers
                dgvPO.CellFormatting += DgvPO_CellFormatting;
                dgvPO.SelectionChanged += DgvPO_SelectionChanged;

                // Kích thủ công SelectionChanged cho dòng đầu tiên:
                // khi bind DataSource, WinForms auto-select row[0] nhưng event đã bị tách
                // nên DgvPO_SelectionChanged chưa bao giờ chạy → _selectedPO_ID vẫn = 0.
                if (dgvPO.Rows.Count > 0)
                {
                    dgvPO.ClearSelection();
                    dgvPO.Rows[0].Selected = true;   // fires SelectionChanged → load schedule/doc
                }

                // Force redraw
                dgvPO.Invalidate();
            }
            finally
            {
                dgvPO.ResumeLayout(true);
            }
        }

        private void LoadSchedHist()
        {
            if (_selectedPO_ID == 0) return;
            try
            {
                _schedules = _svc.GetSchedules(_selectedPO_ID);

                // Cập nhật cache để grid PO phản ánh đợt mới nhất
                _allSchedulesCache[_selectedPO_ID] = _schedules;

                dgvSchedule.Rows.Clear();
                foreach (var s in _schedules)
                {
                    int i = dgvSchedule.Rows.Add();
                    var r = dgvSchedule.Rows[i];
                    r.Cells["S_ID"].Value = s.Schedule_ID;
                    r.Cells["Dot_TT"].Value = s.Dot_TT;
                    r.Cells["Pay_Method"].Value = s.Pay_Method;
                    r.Cells["Payment_Type"].Value = s.Payment_Type;
                    r.Cells["Percent_TT"].Value = s.Percent_TT;
                    r.Cells["Amount_Plan"].Value = FormatAmt(s.Amount_Plan);
                    r.Cells["Due_Date"].Value = s.Due_Date.HasValue ? s.Due_Date.Value.ToString("dd/MM/yyyy") : "";
                    r.Cells["Description"].Value = s.Description;
                    r.Cells["S_Status"].Value = s.Status;
                }

                // Reload Payment Request Progressing theo project của PO đang chọn
                var selPo = _poSummaries.Find(x => x.PO_ID == _selectedPO_ID);
                LoadPaymentProgress(selPo?.Project_Name);
            }
            catch (Exception ex) { Err(ex.Message); }
        }


        // ── Lưu một dòng vào PO_PaymentProgress (UPSERT, chỉ INSERT nếu chưa có) ──
        private void SavePaymentProgressToDB(int printId, string poNo, string totalStr,
            string invPath, string delPath)
        {
            if (printId <= 0) return;
            decimal.TryParse((totalStr ?? "").Replace(",", ""), out decimal total);
            try
            {
                using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                conn.Open();
                // Tạo bảng nếu chưa có
                new Microsoft.Data.SqlClient.SqlCommand(@"
                    IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='PO_PaymentProgress')
                    CREATE TABLE PO_PaymentProgress (
                        Progress_ID  INT IDENTITY(1,1) PRIMARY KEY,
                        Print_ID     INT NOT NULL UNIQUE,
                        PONo         NVARCHAR(100) NULL,
                        Amount_Total DECIMAL(18,2) NULL,
                        PR_Status    NVARCHAR(50)  DEFAULT 'Pending',
                        PR_Note      NVARCHAR(500) NULL,
                        PR_Paid      BIT           DEFAULT 0,
                        INV_Path     NVARCHAR(500) NULL,
                        Del_Path     NVARCHAR(500) NULL,
                        Created_At   DATETIME      DEFAULT GETDATE(),
                        Updated_At   DATETIME      DEFAULT GETDATE()
                    );
                    IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='PO_PaymentProgress' AND COLUMN_NAME='INV_Path')
                        ALTER TABLE PO_PaymentProgress ADD INV_Path NVARCHAR(500) NULL;
                    IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='PO_PaymentProgress' AND COLUMN_NAME='Del_Path')
                        ALTER TABLE PO_PaymentProgress ADD Del_Path NVARCHAR(500) NULL;
                    IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='PO_PaymentProgress' AND COLUMN_NAME='PONo')
                        ALTER TABLE PO_PaymentProgress ADD PONo NVARCHAR(100) NULL;
                    IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='PO_PaymentProgress' AND COLUMN_NAME='Amount_Total')
                        ALTER TABLE PO_PaymentProgress ADD Amount_Total DECIMAL(18,2) NULL;
                    IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='PO_PaymentProgress' AND COLUMN_NAME='Dot_TT')
                        ALTER TABLE PO_PaymentProgress ADD Dot_TT NVARCHAR(50) NULL;
                    IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='PO_PaymentProgress' AND COLUMN_NAME='EC_Status')
                        ALTER TABLE PO_PaymentProgress ADD EC_Status NVARCHAR(50) NULL;", conn).ExecuteNonQuery();

                // Chỉ INSERT nếu chưa tồn tại — giữ nguyên Status/Note/Paid nếu đã có
                var cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                    IF NOT EXISTS (SELECT 1 FROM PO_PaymentProgress WHERE Print_ID = @pid)
                        INSERT INTO PO_PaymentProgress (Print_ID, PONo, Amount_Total, INV_Path, Del_Path)
                        VALUES (@pid, @poNo, @total, @invPath, @delPath);", conn);
                cmd.Parameters.AddWithValue("@pid", printId);
                cmd.Parameters.AddWithValue("@poNo", poNo);
                cmd.Parameters.AddWithValue("@total", total);
                cmd.Parameters.AddWithValue("@invPath", invPath ?? "");
                cmd.Parameters.AddWithValue("@delPath", delPath ?? "");
                cmd.ExecuteNonQuery();
            }
            catch { }
        }

        // ── Load Payment Request Progressing: tất cả PO của dự án đang chọn, join Zalo data ──
        private void LoadPaymentProgress(string projectName = null)
        {
            if (dgvHistory == null) return;
            dgvHistory.Rows.Clear();
            if (string.IsNullOrEmpty(projectName)) return;
            try
            {
                Services.ZaloImportService.EnsureTable();

                using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                conn.Open();

                var cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                    SELECT
                        ph.PONo,
                        ph.PO_ID,
                        (SELECT ISNULL(SUM(d.Amount), 0) FROM PO_Detail d WHERE d.PO_ID = ph.PO_ID) AS Amount_Net,
                        ISNULL(ph.Total_Amount, 0) AS Total_Amount,
                        ISNULL(z.Progress_Status, '') AS EC_Status,
                        z.Paid_Date AS Paid_Date,
                        ISNULL(z.Final_Amount, 0) AS Paid_Amount,
                        ISNULL(z.Note, '') AS ZNote
                    FROM PO_head ph
                    LEFT JOIN (
                        SELECT UPPER(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(PO_No,''))),CHAR(160),''),CHAR(9),''),CHAR(13),''),CHAR(10),''),' ','')) AS PO_No_Key,
                               Progress_Status, Paid_Date, Final_Amount, Note,
                               ROW_NUMBER() OVER (PARTITION BY UPPER(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(PO_No,''))),CHAR(160),''),CHAR(9),''),CHAR(13),''),CHAR(10),''),' ',''))
                                                  ORDER BY CASE WHEN Paid_Date IS NOT NULL THEN 0 ELSE 1 END,
                                                           File_Date DESC, Import_ID DESC) AS rn
                        FROM Zalo_PaymentImport
                        WHERE REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(PO_No,''))),CHAR(160),''),CHAR(9),''),CHAR(13),''),CHAR(10),''),' ','') <> ''
                    ) z ON z.PO_No_Key = UPPER(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(ph.PONo,''))),CHAR(160),''),CHAR(9),''),CHAR(13),''),CHAR(10),''),' ','')) AND z.rn = 1
                    WHERE ph.Project_Name = @proj
                    ORDER BY ph.PONo", conn);
                cmd.Parameters.AddWithValue("@proj", projectName);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    decimal total  = Convert.ToDecimal(reader["Total_Amount"]);
                    decimal net    = Convert.ToDecimal(reader["Amount_Net"]);
                    decimal paid   = reader["Paid_Amount"] != DBNull.Value ? Convert.ToDecimal(reader["Paid_Amount"]) : 0;
                    bool hasPaid   = reader["Paid_Date"] != DBNull.Value;
                    // Hiển thị: total = giá trị hợp đồng, remain = total - paid
                    decimal remain = Math.Max(total - paid, 0);

                    int i = dgvHistory.Rows.Add();
                    dgvHistory.Rows[i].Cells["H_PONo"].Value     = reader["PONo"]?.ToString() ?? "";
                    dgvHistory.Rows[i].Cells["H_PreTax"].Value   = net > 0    ? FormatAmt(net)   : "";
                    dgvHistory.Rows[i].Cells["H_Total"].Value    = total > 0  ? FormatAmt(total) : "";
                    dgvHistory.Rows[i].Cells["H_ECStatus"].Value = reader["EC_Status"]?.ToString() ?? "";
                    dgvHistory.Rows[i].Cells["H_Paid"].Value     = paid > 0   ? FormatAmt(paid)  : "";
                    dgvHistory.Rows[i].Cells["H_PaidDate"].Value = hasPaid
                        ? Convert.ToDateTime(reader["Paid_Date"]).ToString("dd/MM/yyyy")
                        : "";
                    dgvHistory.Rows[i].Cells["H_Remain"].Value   = FormatAmt(remain);
                    
                    // ── Lấy Tên NCC (Supplier Short Name) từ PO_ID ──
                    string suppShort = "";
                    try
                    {
                        int poId = Convert.ToInt32(reader["PO_ID"]);
                        var poSummary = _poSummaries.Find(p => p.PO_ID == poId);
                        if (poSummary != null && !string.IsNullOrEmpty(poSummary.Supplier_Short))
                            suppShort = poSummary.Supplier_Short;
                    }
                    catch { }
                    dgvHistory.Rows[i].Cells["H_SuppShort"].Value = suppShort;
                }
            }
            catch (Exception ex) { Err("LoadPaymentProgress: " + ex.Message); }
        }
        private void LoadDocuments()
        {
            if (dgvDoc == null) return;
            if (dgvDoc.InvokeRequired)
            {
                dgvDoc.Invoke(new Action(LoadDocuments));
                return;
            }
            dgvDoc.Rows.Clear();
            if (_selectedPO_ID == 0) return;

            var po = _poSummaries.Find(x => x.PO_ID == _selectedPO_ID);
            if (po == null) return;

            ProjectInfo proj = null;
            try
            {
                var projSvc = new ProjectService();
                var all = projSvc.GetAll();
                proj = all.Find(p => (p.ProjectName ?? "").Equals(po.Project_Name, StringComparison.OrdinalIgnoreCase)
                                  || (p.ProjectCode ?? "").Equals(po.Project_Name, StringComparison.OrdinalIgnoreCase));
            }
            catch { }

            string poNo = po.PONo ?? "";
            ScanFolderToGrid(proj?.INV_Link ?? "", $"INV_{poNo}", "Invoice");
            ScanFolderToGrid(proj?.DeliveryNote_Link ?? "", $"Delivery_{poNo}", "Delivery Note");
        }

        private void ScanFolderToGrid(string folder, string prefix, string docType)
        {
            if (string.IsNullOrWhiteSpace(folder) || !System.IO.Directory.Exists(folder)) return;
            try
            {
                var files = System.IO.Directory.GetFiles(folder, $"{prefix}*",
                    System.IO.SearchOption.TopDirectoryOnly);
                foreach (var f in files)
                {
                    int i = dgvDoc.Rows.Add();
                    dgvDoc.Rows[i].Cells["Doc_Path"].Value = f;
                    dgvDoc.Rows[i].Cells["Doc_Name"].Value = System.IO.Path.GetFileName(f);
                }
            }
            catch { }
        }

        // Lấy tên file đầu tiên khớp prefix trong folder
        private string GetFirstFileName(string folder, string prefix)
        {
            if (string.IsNullOrWhiteSpace(folder) || !System.IO.Directory.Exists(folder)) return "";
            try
            {
                var files = System.IO.Directory.GetFiles(folder, $"{prefix}*",
                    System.IO.SearchOption.TopDirectoryOnly);
                return files.Length > 0 ? System.IO.Path.GetFileName(files[0]) : "";
            }
            catch { return ""; }
        }

        // Lấy full path file đầu tiên khớp prefix trong folder
        private string GetFirstFilePath(string folder, string prefix)
        {
            if (string.IsNullOrWhiteSpace(folder) || !System.IO.Directory.Exists(folder)) return "";
            try
            {
                var files = System.IO.Directory.GetFiles(folder, $"{prefix}*",
                    System.IO.SearchOption.TopDirectoryOnly);
                return files.Length > 0 ? files[0] : "";
            }
            catch { return ""; }
        }

        private void FilterHistoryGrid()
        {
            // Filter đã được thay bằng load theo project — method giữ lại để tránh lỗi tham chiếu
        }

        // Tính trạng thái thanh toán dựa trên Zalo_PaymentImport (dùng chung cho cả 2 tab)
        private string GetZaloStatus(string poNo, decimal totalPOAmount)
        {
            string key = (poNo ?? "").ToUpperInvariant()
                .Replace(" ", "").Replace("\t", "").Replace("\r", "").Replace("\n", "");
            _zaloPaidCache.TryGetValue(key, out decimal zaloPaid);

            if (zaloPaid <= 0) return "Pending";
            decimal remain = totalPOAmount - zaloPaid;
            return remain <= 0 ? "Đã thanh toán" : "Thanh toán 1 phần";
        }

        // Load tiền đã TT trong kỳ từ Zalo_PaymentImport.
        // Mỗi PONo chỉ lấy 1 dòng đại diện (rn=1, ưu tiên dòng có Paid_Date mới nhất)
        // rồi lọc dòng đó có Paid_Date nằm trong kỳ.
        private Dictionary<string, decimal> LoadZaloInRangeCache(DateTime from, DateTime to)
        {
            var cache = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                conn.Open();
                var chk = new Microsoft.Data.SqlClient.SqlCommand(
                    "SELECT COUNT(1) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='Zalo_PaymentImport'", conn);
                if ((int)chk.ExecuteScalar() == 0) return cache;

                var cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                    SELECT PO_Key, ISNULL(Final_Amount, 0) AS Paid_In_Range
                    FROM (
                        SELECT
                            UPPER(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(PO_No,''))),
                                CHAR(160),''),CHAR(9),''),CHAR(13),''),CHAR(10),''),' ','')) AS PO_Key,
                            Final_Amount,
                            Paid_Date,
                            ROW_NUMBER() OVER (
                                PARTITION BY UPPER(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(PO_No,''))),
                                    CHAR(160),''),CHAR(9),''),CHAR(13),''),CHAR(10),''),' ',''))
                                ORDER BY CASE WHEN Paid_Date IS NOT NULL THEN 0 ELSE 1 END,
                                         File_Date DESC, Import_ID DESC
                            ) AS rn
                        FROM Zalo_PaymentImport
                        WHERE REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(PO_No,''))),
                                  CHAR(160),''),CHAR(9),''),CHAR(13),''),CHAR(10),''),' ','') <> ''
                    ) z
                    WHERE rn = 1
                      AND Paid_Date >= @from AND Paid_Date <= @to", conn);
                cmd.Parameters.AddWithValue("@from", from.Date);
                cmd.Parameters.AddWithValue("@to", to.Date);
                using var rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    string key = rdr["PO_Key"]?.ToString() ?? "";
                    decimal paid = rdr["Paid_In_Range"] != DBNull.Value ? Convert.ToDecimal(rdr["Paid_In_Range"]) : 0;
                    if (!string.IsNullOrEmpty(key))
                        cache[key] = paid;
                }
            }
            catch { }
            return cache;
        }

        // Load tiền đã TT từ Zalo_PaymentImport.
        // Lấy Final_Amount của dòng mới nhất theo PO_No (ưu tiên dòng có Paid_Date).
        // Không yêu cầu Paid_Date IS NOT NULL vì cột này có thể NULL do import từ
        // file Excel sai vị trí cột (đã được fix bởi header detection).
        private Dictionary<string, decimal> LoadZaloPaidCache()
        {
            var cache = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                conn.Open();
                var chk = new Microsoft.Data.SqlClient.SqlCommand(
                    "SELECT COUNT(1) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='Zalo_PaymentImport'", conn);
                if ((int)chk.ExecuteScalar() == 0) return cache;

                var cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                    SELECT PO_Key, ISNULL(Final_Amount, 0) AS Da_TT
                    FROM (
                        SELECT
                            UPPER(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(PO_No,''))),
                                CHAR(160),''),CHAR(9),''),CHAR(13),''),CHAR(10),''),' ','')) AS PO_Key,
                            Final_Amount,
                            Paid_Date,
                            ROW_NUMBER() OVER (
                                PARTITION BY UPPER(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(PO_No,''))),
                                    CHAR(160),''),CHAR(9),''),CHAR(13),''),CHAR(10),''),' ',''))
                                ORDER BY CASE WHEN Paid_Date IS NOT NULL THEN 0 ELSE 1 END,
                                         File_Date DESC, Import_ID DESC
                            ) AS rn
                        FROM Zalo_PaymentImport
                        WHERE REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(PO_No,''))),
                                  CHAR(160),''),CHAR(9),''),CHAR(13),''),CHAR(10),''),' ','') <> ''
                          AND ISNULL(Final_Amount, 0) > 0
                    ) z
                    WHERE rn = 1", conn);

                using var rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    string key = rdr["PO_Key"]?.ToString() ?? "";
                    decimal paid = rdr["Da_TT"] != DBNull.Value ? Convert.ToDecimal(rdr["Da_TT"]) : 0;
                    if (!string.IsNullOrEmpty(key) && paid > 0)
                        cache[key] = paid;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[LoadZaloPaidCache] Error: " + ex.Message);
            }
            return cache;
        }


        // =====================================================================
        //  CẬP NHẬT TRẠNG THÁI THANH TOÁN TỰ ĐỘNG
        //  Logic:
        //    Tổng đã TT >= Tổng kế hoạch  → "Đã TT đủ"
        //    Tổng đã TT > 0               → "Một phần"
        //    Tổng đã TT = 0               → "Chưa TT"
        //  Cập nhật vào:
        //    - PO_Payment_Schedule.Status  (từng đợt liên quan)
        //    - PO_head.Payment_Status      (tổng trạng thái PO)
        // =====================================================================
        private void UpdatePaymentStatusByPONo(string poNo)
        {
            if (string.IsNullOrEmpty(poNo)) return;

            using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
            conn.Open();

            // ── 1. Tổng đã TT từ PO_HistoryPaid theo PONo ────────────────────
            // PO_HistoryPaid có cột PONo trực tiếp — không cần join phức tạp
            var cmdPaid = new Microsoft.Data.SqlClient.SqlCommand(
                "SELECT ISNULL(SUM(Amount_Total), 0) AS Total_Paid FROM PO_HistoryPaid WHERE PONo = @poNo",
                conn);
            cmdPaid.Parameters.AddWithValue("@poNo", poNo);
            decimal totalPaid = Convert.ToDecimal(cmdPaid.ExecuteScalar() ?? 0);

            // ── 2. Tổng kế hoạch từ PO_Payment_Schedule ──────────────────────
            var cmdPlan = new Microsoft.Data.SqlClient.SqlCommand(@"
                SELECT ISNULL(SUM(ps.Amount_Plan), 0) AS Total_Plan
                FROM PO_Payment_Schedule ps
                INNER JOIN PO_head ph ON ph.PO_ID = ps.PO_ID
                WHERE ph.PONo = @poNo2", conn);
            cmdPlan.Parameters.AddWithValue("@poNo2", poNo);
            decimal totalPlan = Convert.ToDecimal(cmdPlan.ExecuteScalar() ?? 0);

            // ── 3. Xác định trạng thái tổng ──────────────────────────────────
            string newPoStatus = totalPaid <= 0 ? "Chưa TT"
                               : totalPaid >= totalPlan ? "Đã TT đủ"
                               : "Một phần";

            // ── 4. Cập nhật từng đợt trong PO_Payment_Schedule ───────────────
            // Tính paid theo từng Dot_TT qua PO_PrintRequestHistory.Dot_TT
            var cmdDots = new Microsoft.Data.SqlClient.SqlCommand(@"
                SELECT ps.Schedule_ID,
                       ps.Amount_Plan,
                       ISNULL(paid.Paid_Dot, 0) AS Paid_Dot
                FROM PO_Payment_Schedule ps
                INNER JOIN PO_head ph ON ph.PO_ID = ps.PO_ID
                LEFT JOIN (
                    SELECT prh.Dot_TT,
                           SUM(hp.Amount_Total) AS Paid_Dot
                    FROM PO_PrintRequestHistory prh
                    INNER JOIN PO_HistoryPaid hp ON hp.Print_ID = prh.Print_ID
                    WHERE prh.PONo = @poNo3
                    GROUP BY prh.Dot_TT
                ) paid ON paid.Dot_TT = ps.Dot_TT
                WHERE ph.PONo = @poNo4", conn);
            cmdDots.Parameters.AddWithValue("@poNo3", poNo);
            cmdDots.Parameters.AddWithValue("@poNo4", poNo);

            var dotUpdates = new List<(int schedId, string status)>();
            using (var rdr = cmdDots.ExecuteReader())
                while (rdr.Read())
                {
                    int sid = Convert.ToInt32(rdr["Schedule_ID"]);
                    decimal plan = rdr["Amount_Plan"] != DBNull.Value ? Convert.ToDecimal(rdr["Amount_Plan"]) : 0;
                    decimal paid = rdr["Paid_Dot"] != DBNull.Value ? Convert.ToDecimal(rdr["Paid_Dot"]) : 0;
                    string dotSt = paid <= 0 ? "Chưa TT"
                                    : paid >= plan ? "Đã TT đủ"
                                    : "Một phần";
                    dotUpdates.Add((sid, dotSt));
                }

            foreach (var (sid, st) in dotUpdates)
            {
                var c = new Microsoft.Data.SqlClient.SqlCommand(
                    "UPDATE PO_Payment_Schedule SET Status = @st WHERE Schedule_ID = @sid", conn);
                c.Parameters.AddWithValue("@st", st);
                c.Parameters.AddWithValue("@sid", sid);
                c.ExecuteNonQuery();
            }

            // ── 5. Cập nhật PO_head.Payment_Status ───────────────────────────
            var cmdPO = new Microsoft.Data.SqlClient.SqlCommand(
                "UPDATE PO_head SET Status = @st WHERE PONo = @poNo5", conn);
            cmdPO.Parameters.AddWithValue("@st", newPoStatus);
            cmdPO.Parameters.AddWithValue("@poNo5", poNo);
            int rows = cmdPO.ExecuteNonQuery();

            // Debug
            System.Diagnostics.Debug.WriteLine(
                $"[UpdatePaymentStatus] {poNo} → {newPoStatus} " +
                $"(paid={totalPaid:N2} / plan={totalPlan:N2}, PO_head rows={rows})");
        }
        private void LoadHistoryPaid(DateTime from, DateTime to)
        {
            if (dgvPaid == null) return;
            dgvPaid.Rows.Clear();
            try
            {
                using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                conn.Open();

                // Bảng chưa tồn tại → không load
                var chk = new Microsoft.Data.SqlClient.SqlCommand(
                    "SELECT COUNT(1) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='PO_HistoryPaid'", conn);
                if ((int)chk.ExecuteScalar() == 0) return;

                var cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                    SELECT HP_ID, PONo, Amount_Total, PR_Note, INV_File, Delivery_File,
                           CONVERT(NVARCHAR(16), Paid_At, 103) + ' ' +
                           SUBSTRING(CONVERT(NVARCHAR(8), Paid_At, 108), 1, 5) AS Paid_At
                    FROM PO_HistoryPaid
                    WHERE Paid_At BETWEEN @from AND @to
                    ORDER BY Paid_At DESC", conn);
                cmd.Parameters.AddWithValue("@from", from);
                cmd.Parameters.AddWithValue("@to", to);
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    int i = dgvPaid.Rows.Add();
                    dgvPaid.Rows[i].Cells["HP_ID"].Value = reader["HP_ID"];
                    dgvPaid.Rows[i].Cells["HP_PONo"].Value = reader["PONo"]?.ToString() ?? "";
                    dgvPaid.Rows[i].Cells["HP_Total"].Value =
                        reader["Amount_Total"] != DBNull.Value
                        ? FormatAmt(Convert.ToDecimal(reader["Amount_Total"])) : "";
                    dgvPaid.Rows[i].Cells["HP_Note"].Value = reader["PR_Note"]?.ToString() ?? "";
                    dgvPaid.Rows[i].Cells["HP_INV"].Value = reader["INV_File"]?.ToString() ?? "";
                    dgvPaid.Rows[i].Cells["HP_Delivery"].Value = reader["Delivery_File"]?.ToString() ?? "";
                    dgvPaid.Rows[i].Cells["HP_PaidAt"].Value = reader["Paid_At"]?.ToString() ?? "";
                }
            }
            catch { }
        }

        // Helper: lấy danh sách file tài liệu (Invoice/Delivery) tồn tại trên đĩa
        private List<string> GetDocFilesToPrint()
        {
            if (_selectedPO_ID == 0) return new List<string>();
            LoadDocuments();
            if (dgvDoc == null || dgvDoc.Rows.Count == 0) return new List<string>();
            var files = new List<string>();
            foreach (DataGridViewRow row in dgvDoc.Rows)
            {
                string path = row.Cells["Doc_Path"].Value?.ToString() ?? "";
                if (System.IO.File.Exists(path))
                    files.Add(path);
            }
            return files;
        }

        // Helper: gửi lệnh in danh sách file tài liệu
        private void PrintDocFiles(List<string> filesToPrint)
        {
            if (filesToPrint == null || filesToPrint.Count == 0) return;
            int ok = 0, fail = 0;
            foreach (var f in filesToPrint)
            {
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = f,
                        Verb = "print",
                        UseShellExecute = true,
                        WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
                    });
                    ok++;
                }
                catch
                {
                    try
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        { FileName = f, UseShellExecute = true });
                        ok++;
                    }
                    catch { fail++; }
                }
            }
            string msg = $"✅ Đã gửi lệnh in {ok} file tài liệu.";
            if (fail > 0) msg += $"\n⚠ {fail} file không thể in.";
            Info(msg, "Hoàn tất");
        }

        // ─────────────────────────────────────────────────────────────────────
        //  IN REQUEST — Fill payment_template.xlsx rồi hiện Print Preview
        // ─────────────────────────────────────────────────────────────────────
        /// <returns>true = đã in/mở file; false = user nhấn Hủy hoặc lỗi sớm</returns>
        private bool PrintPaymentRequest()
        {
            try
            {
                var po = _poSummaries.Find(x => x.PO_ID == _selectedPO_ID);
                if (po == null) { Warn("Không tìm thấy thông tin PO!"); return false; }

                var poHead = _poSvc.GetAll().Find(x => x.PO_ID == _selectedPO_ID);

                // Luôn lấy schedule mới nhất từ DB để tránh cache cũ/thiếu dữ liệu khi in request
                var scheds = _svc.GetSchedules(_selectedPO_ID) ?? new List<PaymentSchedule>();

                // Chuẩn hóa thứ tự đợt để map đúng vào các placeholder đợt 1..5
                scheds = scheds
                    .OrderBy(s => s.Dot_TT)
                    .ThenBy(s => s.Due_Date ?? DateTime.MaxValue)
                    .ToList();

                // Đồng bộ lại cache sau khi lấy mới
                _allSchedulesCache[_selectedPO_ID] = scheds;

                // Tìm Supplier
                Supplier supp = null;
                if (poHead != null)
                    supp = _allSuppliers.Find(s => s.Supplier_ID == poHead.Supplier_ID);
                supp = supp ?? new Supplier();

                // Đường dẫn template
                string templatePath = System.IO.Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory, "Templates", "payment_template.xlsx");
                if (!System.IO.File.Exists(templatePath))
                {
                    Warn($"Không tìm thấy file template!\nĐường dẫn: {templatePath}");
                    return false;
                }

                // Tạo file tạm để fill dữ liệu
                string tempPath = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(),
                    $"PaymentRequest_{po.PONo}_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
                System.IO.File.Copy(templatePath, tempPath, true);

                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (var pkg = new OfficeOpenXml.ExcelPackage(new System.IO.FileInfo(tempPath)))
                {
                    var ws = pkg.Workbook.Worksheets[0];

                    // ── Tính toán ──
                    decimal totalBeforeVat = po.Total_PO_Amount;
                    decimal totalPaid = po.Total_Paid;
                    int dotCount = scheds.Count;

                    // Tính VAT thực tế từ các dòng PO — group theo từng mức thuế
                    var poDetails = _poSvc.GetDetails(_selectedPO_ID);
                    var vatGroups = poDetails
                        .GroupBy(d => d.VAT)
                        .Select(g => new {
                            Rate     = g.Key,
                            SubTotal = g.Sum(d => d.VAT > 0 ? d.Amount / (1 + d.VAT / 100) : d.Amount)
                        })
                        .OrderBy(g => g.Rate)
                        .ToList();
                    decimal detailSubTotal = vatGroups.Sum(g => g.SubTotal);
                    decimal detailVatTotal = vatGroups.Sum(g => g.SubTotal * g.Rate / 100);
                    decimal vatRate = detailSubTotal > 0 ? detailVatTotal / detailSubTotal : 0.1m;
                    // Mixed = có từ 2 mức thuế khác nhau trở lên
                    bool isMixedVat = vatGroups.Select(g => g.Rate).Distinct().Count() > 1;

                    // A1 — (N)th Payment Request
                    int paidDots = scheds.Count(s => s.Status == "Đã TT đủ");
                    string ordinal = (paidDots + 1) switch { 1 => "1st", 2 => "2nd", 3 => "3rd", _ => $"{paidDots + 1}th" };
                    ReplaceCell(ws, "(   )th  Payment Request", $"({ordinal}) Payment Request");

                    // A3 — Project Name (ô C3 trống, điền tên dự án sau dấu ":")
                    FillNextCell(ws, "A3", "Project Name", po.Project_Name ?? "");

                    // C5 — W/O No, M5 — PO No
                    ReplaceCell(ws, "<<WO-NO>>", poHead?.WorkorderNo ?? "");
                    ReplaceCell(ws, "<<PO-NO>>", po.PONo ?? "");

                    // A6 Contract date — lấy PO_Date
                    string contractDate = po.PO_Date.HasValue ? po.PO_Date.Value.ToString("dd/MM/yyyy") : "";
                    FillNextCell(ws, "A6", "Contract date", contractDate);

                    // I6 Payment date — thứ 4 tuần sau
                    DateTime nextWed = GetNextWednesday();
                    string paymentDate = nextWed.ToString("dd/MM/yyyy");
                    FillRightCell(ws, "I6", "Payment date", paymentDate);

                    // C7 — Contract amount (tổng trước VAT)
                    ReplaceCell(ws, "<<Tổng số tiền trước thuế>>", FormatAmt(totalBeforeVat));

                    // C8 — Requested amount (tổng đợt chưa TT)
                    decimal reqAmt = scheds.Where(s => s.Status != "Đã TT đủ").Sum(s => s.Amount_Plan);
                    ReplaceCell(ws, "<<Số tiền theo đợt>>", FormatAmt(reqAmt));

                    // ── Lấy ngày thanh toán thực tế của các đợt đã TT ──
                    var actualPayDates = new Dictionary<int, string>();
                    try
                    {
                        using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                        conn.Open();
                        var cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                            SELECT ps.Dot_TT,
                                   MAX(ph.Payment_Date) AS Last_Payment_Date
                            FROM PO_Payment_Schedule ps
                            INNER JOIN PO_Payment_History ph ON ph.Schedule_ID = ps.Schedule_ID
                            WHERE ps.PO_ID = @poId
                            GROUP BY ps.Dot_TT", conn);
                        cmd.Parameters.AddWithValue("@poId", _selectedPO_ID);
                        using var reader = cmd.ExecuteReader();
                        while (reader.Read())
                        {
                            int dot = Convert.ToInt32(reader["Dot_TT"]);
                            string datePaid = reader["Last_Payment_Date"] != DBNull.Value
                                ? Convert.ToDateTime(reader["Last_Payment_Date"]).ToString("dd/MM/yyyy")
                                : "";
                            actualPayDates[dot] = datePaid;
                        }
                    }
                    catch { }

                    // ── Rows 12-16: từng đợt ──
                    decimal sumNet = 0, sumVat = 0, sumTotal = 0;
                    for (int i = 0; i < 5; i++)
                    {
                        int excelRow = 12 + i; // Rows 12..16
                        if (i < dotCount)
                        {
                            var s = scheds[i];
                            decimal net = s.Amount_Plan;
                            decimal vat = Math.Round(net * vatRate, 0);
                            decimal tot = Math.Round(net + vat, 0);
                            sumNet += net;
                            sumVat += vat;
                            sumTotal += tot;

                            string dateValue;
                            if (s.Status == "Đã TT đủ")
                            {
                                actualPayDates.TryGetValue(s.Dot_TT, out dateValue);
                                dateValue = dateValue ?? (s.Due_Date.HasValue ? s.Due_Date.Value.ToString("dd/MM/yyyy") : "");
                            }
                            else if (tot > 0 && i == scheds.FindIndex(x => x.Status != "Đã TT đủ"))
                            {
                                dateValue = GetNextWednesday().ToString("dd/MM/yyyy");
                            }
                            else
                            {
                                dateValue = "";
                            }

                            ReplaceCell(ws, $"<<Số tiền đợt {i + 1}>>", FormatAmt(net));
                            ReplaceCell(ws, $"<<Số tiền thuế lần {i + 1}>>", FormatAmt0(vat));
                            ReplaceCell(ws, $"<<Số tiền sau thuế lần {i + 1}>>", FormatAmt0(tot));
                            ReplaceCell(ws, $"<<Ngày yêu cầu lần {i + 1}>>", dateValue);

                            // Khi PO có nhiều mức VAT khác nhau: ghi breakdown vào cột Remarks (O)
                            // Hiển thị tổng sau VAT của từng mức, prorated theo tỉ lệ subtotal
                            if (isMixedVat && detailSubTotal > 0)
                            {
                                var parts = vatGroups
                                    .Where(g => g.Rate > 0)
                                    .Select(g =>
                                    {
                                        decimal groupSubNet = Math.Round(net * (g.SubTotal / detailSubTotal), 0);
                                        decimal groupAfterVat = Math.Round(groupSubNet * (1 + g.Rate / 100), 0);
                                        return $"VAT {g.Rate:0}%: {FormatAmt0(groupAfterVat)}";
                                    });
                                ws.Cells[excelRow, 15].Value = string.Join(" | ", parts);
                            }
                        }
                        else
                        {
                            ReplaceCell(ws, $"<<Số tiền đợt {i + 1}>>", "");
                            ReplaceCell(ws, $"<<Số tiền thuế lần {i + 1}>>", "");
                            ReplaceCell(ws, $"<<Số tiền sau thuế lần {i + 1}>>", "");
                            ReplaceCell(ws, $"<<Ngày yêu cầu lần {i + 1}>>", "");
                            if (isMixedVat) ws.Cells[excelRow, 15].Value = "";
                        }
                    }

                    ReplaceCellAll(ws, "<<Sum>>", new[] { FormatAmt(sumNet), FormatAmt0(sumVat), FormatAmt0(sumTotal) });

                    decimal balNet = Math.Max(totalBeforeVat - sumNet, 0);
                    decimal balTotal = Math.Round(Math.Max(totalBeforeVat * (1 + vatRate) - sumTotal - totalPaid, 0), 0);
                    ReplaceCell(ws, "<<Tổng số tiền trước thuế còn lại>>", FormatAmt(balNet));
                    ReplaceCell(ws, "<<Tổng số tiền sau thuế còn lại>>", FormatAmt0(balTotal));

                    ReplaceCell(ws, "<<Ngày yêu cầu>>", DateTime.Today.ToString("dd/MM/yyyy"));

                    string suppName = supp.Company_Name ?? supp.Supplier_Name ?? "";
                    string suppAddress = GetSupplierProp(supp, "Company_Address", "Address") ?? "";
                    ReplaceCell(ws, "<<Tên nhà cung cấp>>", suppName);
                    ReplaceCell(ws, "<<Địa chỉ Nhà cung cấp>>", suppAddress);

                    pkg.Save();
                }

                // ── Hỏi người dùng cách in ──
                var printOpt = ShowPrintOptionDialog(po.PONo);
                if (printOpt == PrintOption.Cancel) return false;   // ← Hủy → báo caller biết

                if (printOpt == PrintOption.OpenFirst)
                {
                    try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = tempPath, UseShellExecute = true }); }
                    catch (Exception ex2) { Err("Không thể mở file:\n" + ex2.Message); }
                }
                else // PrintOption.DirectPrint
                {
                    PrintExcelSilently(tempPath);
                }
                return true;
            }
            catch (Exception ex) { Err("Lỗi tạo file in: " + ex.Message); return false; }
        }

        private enum PrintOption { Cancel, OpenFirst, DirectPrint }

        private PrintOption ShowPrintOptionDialog(string poNo)
        {
            var result = PrintOption.Cancel;

            var dlg = new Form
            {
                Text            = "🖨 Tùy chọn in Payment Request",
                Size            = new Size(420, 200),
                MinimumSize     = new Size(420, 200),
                MaximumSize     = new Size(420, 200),
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition   = FormStartPosition.CenterParent,
                MaximizeBox     = false,
                MinimizeBox     = false,
                BackColor       = Color.FromArgb(245, 247, 250)
            };

            var lbl = new Label
            {
                Text      = $"PO: {poNo}\nChọn cách xử lý file Payment Request:",
                Location  = new Point(14, 14),
                Size      = new Size(390, 40),
                Font      = new Font("Segoe UI", 9.5f),
                ForeColor = Color.FromArgb(40, 60, 90)
            };
            dlg.Controls.Add(lbl);

            var btnOpen = new Button
            {
                Text      = "📂  Mở file trước khi in",
                Location  = new Point(14, 68),
                Size      = new Size(185, 48),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Segoe UI", 9.5f),
                Cursor    = Cursors.Hand
            };
            btnOpen.FlatAppearance.BorderSize = 0;
            btnOpen.Click += (s, e) => { result = PrintOption.OpenFirst; dlg.Close(); };
            dlg.Controls.Add(btnOpen);

            var btnPrint = new Button
            {
                Text      = "🖨  In ngay (không mở)",
                Location  = new Point(210, 68),
                Size      = new Size(185, 48),
                BackColor = Color.FromArgb(0, 150, 100),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font      = new Font("Segoe UI", 9.5f),
                Cursor    = Cursors.Hand
            };
            btnPrint.FlatAppearance.BorderSize = 0;
            btnPrint.Click += (s, e) => { result = PrintOption.DirectPrint; dlg.Close(); };
            dlg.Controls.Add(btnPrint);

            var btnCancel = new Button
            {
                Text         = "Hủy",
                Location     = new Point(315, 135),
                Size         = new Size(80, 28),
                BackColor    = Color.FromArgb(220, 225, 235),
                ForeColor    = Color.FromArgb(80, 90, 110),
                FlatStyle    = FlatStyle.Flat,
                Font         = new Font("Segoe UI", 8.5f),
                Cursor       = Cursors.Hand,
                DialogResult = DialogResult.Cancel
            };
            btnCancel.FlatAppearance.BorderSize = 0;
            dlg.CancelButton = btnCancel;
            dlg.Controls.Add(btnCancel);

            dlg.ShowDialog(this);
            return result;
        }

        // ─────────────────────────────────────────────────────────────────────
        //  In ẩn qua Excel Interop — không hỏi save, tự đóng sau khi in xong
        // ─────────────────────────────────────────────────────────────────────
        private void PrintExcelSilently(string filePath)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Microsoft.Office.Interop.Excel.Workbook    wb        = null;
            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application
                {
                    Visible       = false,
                    DisplayAlerts = false
                };

                wb = excelApp.Workbooks.Open(
                    filePath,
                    UpdateLinks:  false,
                    ReadOnly:     true);

                // In toàn bộ worksheet, không preview, không hỏi lưu
                wb.PrintOut(
                    Preview:     false,
                    PrintToFile: false,
                    Collate:     true);
            }
            catch (Exception ex)
            {
                // Fallback: mở file để user tự in — thông báo lý do
                Err($"Không thể in ngầm:\n{ex.Message}\n\nHệ thống sẽ mở file để bạn in thủ công.");
                try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = filePath, UseShellExecute = true }); }
                catch { }
            }
            finally
            {
                // Đóng không save, giải phóng COM object
                try { wb?.Close(SaveChanges: false); } catch { }
                try
                {
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
                catch { }
                try
                {
                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                    }
                }
                catch { }
            }
        }

        private void ReplaceCell(OfficeOpenXml.ExcelWorksheet ws, string find, string replace)
        {
            if (ws.Dimension == null) return;
            foreach (var cell in ws.Cells[ws.Dimension.Address])
            {
                if (cell.IsRichText)
                {
                    foreach (var run in cell.RichText)
                        if (run.Text.Contains(find, StringComparison.OrdinalIgnoreCase))
                            run.Text = run.Text.Replace(find, replace, StringComparison.OrdinalIgnoreCase);
                }
                else
                {
                    string v = cell.Value?.ToString() ?? "";
                    if (v.Contains(find, StringComparison.OrdinalIgnoreCase))
                        cell.Value = v.Replace(find, replace, StringComparison.OrdinalIgnoreCase);
                }
            }
        }

        private void ReplaceCellAll(OfficeOpenXml.ExcelWorksheet ws, string find, string[] replacements)
        {
            if (ws.Dimension == null) return;
            int foundIdx = 0;
            foreach (var cell in ws.Cells[ws.Dimension.Address])
            {
                string v = cell.IsRichText
                    ? string.Concat(cell.RichText.Select(r => r.Text))
                    : cell.Value?.ToString() ?? "";
                if (!v.Contains(find, StringComparison.OrdinalIgnoreCase)) continue;
                if (foundIdx >= replacements.Length) break;

                string rep = replacements[foundIdx++];
                if (cell.IsRichText)
                {
                    foreach (var run in cell.RichText)
                        if (run.Text.Contains(find, StringComparison.OrdinalIgnoreCase))
                            run.Text = run.Text.Replace(find, rep, StringComparison.OrdinalIgnoreCase);
                }
                else
                {
                    cell.Value = v.Replace(find, rep, StringComparison.OrdinalIgnoreCase);
                }
            }
        }

        private static string GetCellText(OfficeOpenXml.ExcelRange cell)
        {
            if (cell.IsRichText)
                return string.Concat(cell.RichText.Select(r => r.Text));
            return cell.Value?.ToString() ?? "";
        }

        private void FillNextCell(OfficeOpenXml.ExcelWorksheet ws, string anchorAddr, string labelFind, string value)
        {
            var anchor = ws.Cells[anchorAddr];
            if (GetCellText(anchor).Contains(labelFind, StringComparison.OrdinalIgnoreCase))
            {
                if (anchorAddr == "A3") ws.Cells["C3"].Value = value;
                else if (anchorAddr == "A6") ws.Cells["C6"].Value = value;
            }
        }

        private void FillRightCell(OfficeOpenXml.ExcelWorksheet ws, string anchorAddr, string labelFind, string value)
        {
            var anchor = ws.Cells[anchorAddr];
            if (GetCellText(anchor).Contains(labelFind, StringComparison.OrdinalIgnoreCase))
                ws.Cells[anchor.Start.Row, anchor.Start.Column + 1].Value = value;
        }

        private string FormatAmt0(decimal v) => v == 0 ? "0" : v.ToString("#,##0");

        private static DateTime GetNextWednesday()
        {
            DateTime today = DateTime.Today;
            // Tìm thứ 2 của tuần hiện tại (tuần bắt đầu thứ 2, Chủ nhật = cuối tuần)
            int dow = (int)today.DayOfWeek; // Sun=0, Mon=1..Sat=6
            int daysToThisMonday = dow == 0 ? -6 : 1 - dow;
            DateTime thisMonday = today.AddDays(daysToThisMonday);
            // Thứ 4 tuần sau = thứ 2 tuần này + 7 ngày + 2 ngày
            return thisMonday.AddDays(9);
        }

        private string GetSupplierProp(Supplier s, params string[] props)
        {
            if (s == null) return "";
            var t = s.GetType();
            foreach (var pName in props)
            {
                var p = t.GetProperty(pName);
                if (p != null) return p.GetValue(s)?.ToString() ?? "";
            }
            return "";
        }

        private void BtnPrintDocs_Click(object sender, EventArgs e)
        {
            if (_selectedPO_ID == 0) { Warn("Vui lòng chọn một PO trước!"); return; }

            // Đảm bảo grid đã load mới nhất
            LoadDocuments();

            if (dgvDoc == null || dgvDoc.Rows.Count == 0)
            {
                Warn($"Không tìm thấy file Invoice hoặc Delivery Note nào cho PO này.\nVui lòng kiểm tra thư mục INV_Link và DeliveryNote_Link của dự án!");
                return;
            }

            var filesToPrint = new List<string>();
            foreach (DataGridViewRow row in dgvDoc.Rows)
            {
                string path = row.Cells["Doc_Path"].Value?.ToString() ?? "";
                if (System.IO.File.Exists(path))
                    filesToPrint.Add(path);
            }

            if (filesToPrint.Count == 0)
            {
                Warn("Không tìm thấy file nào để in. Vui lòng kiểm tra đường dẫn!");
                return;
            }

            int ok = 0, fail = 0;
            foreach (var f in filesToPrint)
            {
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = f,
                        Verb = "print",
                        UseShellExecute = true,
                        WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
                    });
                    ok++;
                }
                catch
                {
                    try
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        { FileName = f, UseShellExecute = true });
                        ok++;
                    }
                    catch { fail++; }
                }
            }

            string msg = $"✅ Đã gửi lệnh in {ok} file.";
            if (fail > 0) msg += $"\n⚠ {fail} file không thể in.";
            Info(msg, "Hoàn tất");
        }



        private void DgvPO_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvPO.SelectedRows.Count == 0) return;
            _selectedPO_ID = Convert.ToInt32(dgvPO.SelectedRows[0].Cells["ID"].Value);
            var p = _poSummaries.Find(x => x.PO_ID == _selectedPO_ID);
            if (p == null) return;

            lblPOName.Text = $"PO: {p.PONo}  —  {p.Project_Name}  |  NCC: {p.Supplier_Name}";
            lblPOAmount.Text = $"Tổng PO: {FormatAmt(p.Total_PO_Amount)} VNĐ";
            lblPOPaid.Text = $"Đã TT: {FormatAmt(p.Total_Paid)} VNĐ";
            lblPORemain.Text = $"Còn nợ: {FormatAmt(p.Amount_Remaining)} VNĐ";
            lblPOStatus.Text = p.Is_Overdue ? "⚠ QUÁ HẠN" : p.Payment_Status;
            lblPOStatus.ForeColor =
                p.Is_Overdue ? Color.FromArgb(255, 100, 100) :
                p.Payment_Status == "Đã TT đủ" ? Color.FromArgb(144, 238, 144) :
                p.Payment_Status == "Một phần" ? Color.FromArgb(255, 200, 100) :
                                                   Color.White;

            int pct = (int)Math.Min(p.Percent_Paid, 100);
            progressPO.Value = pct;
            lblPOProgress.Text = $"{pct}%";

            // Load trực tiếp trên UI thread để tránh cross-thread khi cập nhật grid/label
            LoadSchedHist();
            LoadDocuments();
        }

        private void DgvPO_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string col = dgvPO.Columns[e.ColumnIndex].Name;
            if (col == "TT_Status")
            {
                string v = e.Value?.ToString() ?? "";
                e.CellStyle.ForeColor =
                    v == "Đã thanh toán"    ? Color.FromArgb(40, 167, 69) :
                    v == "Thanh toán 1 phần" ? Color.FromArgb(255, 140, 0) :
                                               Color.FromArgb(0, 120, 212);   // Pending
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
            if (col == "Qua_Han" && e.Value?.ToString() != "")
            {
                e.CellStyle.ForeColor = Color.FromArgb(220, 53, 69);
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
            if (col == "Con_No")
            {
                e.CellStyle.ForeColor = Color.FromArgb(220, 53, 69);
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
            if (col.StartsWith("Dot") && col.EndsWith("_Status"))
            {
                string v = e.Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(v))
                {
                    e.CellStyle.ForeColor =
                        v == "Đã TT đủ" ? Color.FromArgb(40, 167, 69) :
                        v == "Một phần" ? Color.FromArgb(255, 140, 0) :
                                           Color.FromArgb(0, 120, 212);
                    e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                }
            }
            if (col.StartsWith("Dot") && col.EndsWith("_Amount") && string.IsNullOrEmpty(e.Value?.ToString()))
                e.CellStyle.BackColor = Color.FromArgb(245, 245, 245);
        }

        private void DgvSched_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (dgvSchedule.Columns[e.ColumnIndex].Name == "S_Status")
            {
                string v = e.Value?.ToString() ?? "";
                e.CellStyle.ForeColor =
                    v == "Đã TT đủ" ? Color.FromArgb(40, 167, 69) :
                    v == "Một phần" ? Color.FromArgb(255, 140, 0) :
                                      Color.FromArgb(0, 120, 212);
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
        }

private void BtnAddSched_Click(object sender, EventArgs e)
{
    if (!PermissionHelper.Check("PAYMENT", "Thêm đợt", "Thêm đợt thanh toán")) return;
    if (_selectedPO_ID == 0) { Warn("Vui lòng chọn PO!"); return; }

    // Kiểm tra trạng thái Email của PO, chỉ cho phép thêm khi Email = "done"
    var poInfo = _poSvc.GetPOByPONo(_selectedPO_ID);
    if (poInfo == null || poInfo.PO_ID == 0) { Warn("Không tìm thấy thông tin chi tiết PO!"); return; }
    
    var emailStatus = poInfo.Email_Status ?? "";
    if (!emailStatus.Equals("done", StringComparison.OrdinalIgnoreCase))
    {
        Warn("Không thể thêm đợt thanh toán Vui lòng gửi Email cho nhà cung cấp trước khi thực hiện thanh toán.");
        return;
    }

            // Lấy tổng PO sau thuế của PO đang chọn
            var po = _poSummaries.Find(x => x.PO_ID == _selectedPO_ID);
            decimal poTotalAfterVat = po?.Total_PO_Amount ?? 0;

            // % mặc định: 100% nếu chưa có đợt nào, ngược lại tính phần còn lại
            decimal usedPct = 0;
            foreach (DataGridViewRow row in dgvSchedule.Rows)
                if (decimal.TryParse(row.Cells["Percent_TT"].Value?.ToString(), out decimal rp)) usedPct += rp;
            decimal defaultPct = Math.Max(0, 100 - usedPct);
            decimal defaultAmt = poTotalAfterVat > 0 ? Math.Round(poTotalAfterVat * defaultPct / 100, 2) : 0;

            int i = dgvSchedule.Rows.Add();
            var r = dgvSchedule.Rows[i];
            r.Cells["S_ID"].Value = 0;
            r.Cells["Dot_TT"].Value = _schedules.Count + dgvSchedule.Rows.Count;
            r.Cells["Pay_Method"].Value = "Full";
            r.Cells["Payment_Type"].Value = "Chuyển khoản";
            r.Cells["Percent_TT"].Value = defaultPct;
            r.Cells["Amount_Plan"].Value = FormatAmt(defaultAmt);
            r.Cells["S_Status"].Value = "Chưa TT";
            dgvSchedule.CurrentCell = dgvSchedule.Rows[i].Cells["Percent_TT"];
            dgvSchedule.BeginEdit(true);
        }

        // Tự động tính lại Amount_Plan khi user sửa cột Percent_TT
        private void DgvSchedule_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (dgvSchedule.Columns[e.ColumnIndex].Name != "Percent_TT") return;

            var po = _poSummaries.Find(x => x.PO_ID == _selectedPO_ID);
            decimal poTotal = po?.Total_PO_Amount ?? 0;
            if (poTotal <= 0) return;

            var row = dgvSchedule.Rows[e.RowIndex];
            if (decimal.TryParse(row.Cells["Percent_TT"].Value?.ToString(), out decimal pct))
            {
                decimal amt = Math.Round(poTotal * pct / 100, 2);
                row.Cells["Amount_Plan"].Value = FormatAmt(amt);
            }
        }

        private void BtnDelSched_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("PAYMENT", "Xóa", "Xóa đợt thanh toán")) return;
            // Lấy dòng đang chọn (ưu tiên SelectedRows, fallback CurrentRow)
            DataGridViewRow selRow = null;
            if (dgvSchedule.SelectedRows.Count > 0)
                selRow = dgvSchedule.SelectedRows[0];
            else if (dgvSchedule.CurrentRow != null)
                selRow = dgvSchedule.CurrentRow;

            if (selRow == null) { Warn("Vui lòng chọn đợt cần xóa!"); return; }

            int schedId = 0;
            try { schedId = Convert.ToInt32(selRow.Cells["S_ID"].Value ?? 0); } catch { }

            string dotLabel = selRow.Cells["Dot_TT"].Value?.ToString() ?? "";
            if (!Ask($"Xóa đợt thanh toán {dotLabel} này?\n(Thao tác này không thể hoàn tác)")) return;

            try
            {
                // Nếu đã có ID trong DB → xóa DB trước
                if (schedId > 0)
                {
                    _svc.DeleteSchedule(schedId);
                    _selectedSchedID = 0;
                }

                // Xóa dòng khỏi grid
                dgvSchedule.Rows.Remove(selRow);

                // Cập nhật lại cache và summary
                if (_allSchedulesCache.ContainsKey(_selectedPO_ID))
                    _allSchedulesCache[_selectedPO_ID].RemoveAll(s => s.Schedule_ID == schedId);

                LoadPOSummary();

                MessageBox.Show(TopOwner, "✅ Đã xóa đợt thanh toán thành công!", "Thành công",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { Err(ex.Message); }
        }

        // Commit giá trị DTP vào cell trước khi lưu hoặc chuyển focus
        private void CommitSchedDtp()
        {
            if (_schedDtp == null || !_schedDtp.Visible || _schedDtpRow < 0) return;
            if (_schedDtpRow < dgvSchedule.Rows.Count)
                dgvSchedule.Rows[_schedDtpRow].Cells["Due_Date"].Value =
                    _schedDtp.Value.ToString("dd/MM/yyyy");
        }

        private async void BtnSaveSched_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("PAYMENT", "Lưu", "Lưu thanh toán")) return;
            if (_selectedPO_ID == 0) return;

            // Force commit DTP nếu đang mở
            CommitSchedDtp();
            if (_schedDtp != null) _schedDtp.Visible = false;
            try
            {
                int saved = 0;
                foreach (DataGridViewRow row in dgvSchedule.Rows)
                {
                    var s = new PaymentSchedule
                    {
                        Schedule_ID = Convert.ToInt32(row.Cells["S_ID"].Value ?? 0),
                        PO_ID = _selectedPO_ID,
                        Dot_TT = Convert.ToInt32(row.Cells["Dot_TT"].Value ?? 1),
                        Pay_Method = row.Cells["Pay_Method"].Value?.ToString() ?? "Full",
                        Payment_Type = row.Cells["Payment_Type"].Value?.ToString() ?? "Chuyển khoản",
                        Percent_TT = decimal.TryParse(row.Cells["Percent_TT"].Value?.ToString(), out decimal pct) ? pct : 0,
                        Amount_Plan = decimal.TryParse((row.Cells["Amount_Plan"].Value?.ToString() ?? "0").Replace(",", ""), out decimal amt) ? amt : 0,
                        Due_Date = (
                            DateTime.TryParseExact(
                                row.Cells["Due_Date"].Value?.ToString() ?? "",
                                new[] { "dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd", "M/d/yyyy" },
                                System.Globalization.CultureInfo.InvariantCulture,
                                System.Globalization.DateTimeStyles.None,
                                out DateTime dd)
                            ||
                            DateTime.TryParse(
                                row.Cells["Due_Date"].Value?.ToString() ?? "",
                                out dd)
                        ) ? dd : (DateTime?)null,
                        Delivery_Ref = "",
                        Description = row.Cells["Description"].Value?.ToString() ?? "",
                        Status = row.Cells["S_Status"].Value?.ToString() ?? "Chưa TT"
                    };
                    if (s.Schedule_ID == 0) _svc.InsertSchedule(s, _currentUser);
                    else _svc.UpdateSchedule(s);
                    saved++;
                }
                MessageBox.Show(TopOwner, $"✅ Đã lưu {saved} đợt thanh toán!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Ghi nhớ PO đang chọn để giữ hiển thị sau khi refresh
                int savedPoId = _selectedPO_ID;

                // Chỉ reload schedule/history của PO này — không reload toàn bộ grid PO
                LoadSchedHist();

                // Reload dữ liệu từ DB trên background thread, sau đó cập nhật UI trên main thread
                var reloadResult = await System.Threading.Tasks.Task.Run(() =>
                {
                    try
                    {
                        var newScheds = _svc.GetSchedules(savedPoId);
                        var freshSummary = _svc.GetPOSummary(savedPoId);
                        return (newScheds, freshSummary);
                    }
                    catch
                    {
                        return ((List<PaymentSchedule>)null, (POPaymentSummary)null);
                    }
                });

                if (reloadResult.Item1 != null)
                    _allSchedulesCache[savedPoId] = reloadResult.Item1;

                if (reloadResult.Item2 != null)
                {
                    int idx = _poSummaries.FindIndex(p => p.PO_ID == savedPoId);
                    if (idx >= 0) _poSummaries[idx] = reloadResult.Item2;
                }

                // Refresh grid PO nhưng giữ nguyên dòng đang chọn
                FilterAndBind();
                foreach (DataGridViewRow row in dgvPO.Rows)
                {
                    if (Convert.ToInt32(row.Cells["ID"].Value ?? 0) == savedPoId)
                    {
                        dgvPO.ClearSelection();
                        row.Selected = true;
                        break;
                    }
                }
            }
            catch (Exception ex) { Err(ex.Message); }
        }


        // Implement IRefreshable interface method
        public async void RefreshData()
        {
            // Reload data and rebind grids
            await LoadDataAsync();
            FilterAndBind();
            LoadSchedHist(); // Ensure schedules and history are also reloaded

            // Need to re-select the current PO if it's still valid
            if (_selectedPO_ID > 0)
            {
                var currentPO = _poSummaries.Find(p => p.PO_ID == _selectedPO_ID);
                if (currentPO != null)
                {
                    // Re-select the row in the PO grid
                    foreach (DataGridViewRow row in dgvPO.Rows)
                    {
                        if (Convert.ToInt32(row.Cells["ID"].Value ?? 0) == _selectedPO_ID)
                        {
                            dgvPO.ClearSelection();
                            row.Selected = true;
                            dgvPO.CurrentCell = row.Cells["PO_No"];
                            break;
                        }
                    }
                    LoadSchedHist();
                }
                else
                {
                    _selectedPO_ID = 0;
                    ClearDetailViews();
                }
            }
            else
            {
                ClearDetailViews();
            }
        }

        private void ClearDetailViews()
        {
            dgvSchedule?.Rows.Clear();
            dgvHistory?.Rows.Clear();
            _schedules?.Clear();
            _histories?.Clear();
        }

        private void BtnDelPayment_Click(object sender, EventArgs e)
        {
            // Lấy Print_ID trực tiếp từ dòng đang chọn — tránh lỗi khi dòng bị ẩn bởi filter
            int printId = 0;
            DataGridViewRow targetRow = null;

            if (dgvHistory.SelectedRows.Count > 0)
                targetRow = dgvHistory.SelectedRows[0];
            else if (dgvHistory.CurrentRow != null)
                targetRow = dgvHistory.CurrentRow;

            if (targetRow != null)
                printId = Convert.ToInt32(targetRow.Cells["H_ID"].Value ?? 0);

            if (printId == 0) { Warn("Vui lòng chọn bản ghi cần xóa!"); return; }

            string poNo = targetRow?.Cells["H_PONo"].Value?.ToString() ?? "";
            if (!Ask($"Xóa dòng này khỏi Payment Request Progressing?\n\nPO: {poNo}\n(Lịch sử in Request gốc vẫn được giữ lại)")) return;
            if (!VerifyAdminPassword()) return;
            try
            {
                using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                conn.Open();
                var cmd = new Microsoft.Data.SqlClient.SqlCommand(
                    "DELETE FROM PO_PaymentProgress WHERE Print_ID = @id", conn);
                cmd.Parameters.AddWithValue("@id", printId);
                int affected = cmd.ExecuteNonQuery();
                if (affected == 0)
                    Warn("Không tìm thấy bản ghi trong DB. Có thể đã bị xóa trước đó.");
                else
                {
                    var selPo2 = _poSummaries.Find(x => x.PO_ID == _selectedPO_ID);
                    LoadPaymentProgress(selPo2?.Project_Name);
                }
            }
            catch (Exception ex) { Err(ex.Message); }
        }

        private void BtnPaymentRequest_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("PAYMENT", "Request to EC", "Request to EC")) return;
            if (_selectedPO_ID == 0)
            {
                Warn("Vui lòng chọn một PO trong danh sách để tạo yêu cầu!");
                return;
            }

            var po = _poSummaries.Find(p => p.PO_ID == _selectedPO_ID);
            var poHead = _poSvc.GetAll().Find(p => p.PO_ID == _selectedPO_ID);
            string mprNo = poHead?.MPR_No ?? "";

            // Kiểm tra PO đã từng được tạo Eccount trước đó chưa
            try
            {
                using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                conn.Open();
                using var cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                    SELECT TOP 1 Printed_Date, ISNULL(Dot_Label,'') AS Dot_Label
                    FROM PO_PrintRequestHistory
                    WHERE PONo = @poNo
                    ORDER BY Printed_Date DESC", conn);
                cmd.Parameters.AddWithValue("@poNo", po?.PONo ?? "");
                using var reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    string printedDate = reader["Printed_Date"] != DBNull.Value
                        ? Convert.ToDateTime(reader["Printed_Date"]).ToString("dd/MM/yyyy HH:mm")
                        : "";
                    string dotLabel = reader["Dot_Label"]?.ToString() ?? "";
                    string dotInfo = string.IsNullOrEmpty(dotLabel) ? "" : $" ({dotLabel})";

                    var result = MessageBox.Show(
                        $"⚠️ PO {po?.PONo} đã được tạo Eccount trước đó{dotInfo} vào lúc {printedDate}.\n\n" +
                        $"Vui lòng kiểm tra lại trước khi tạo mới!\n\n" +
                        $"Bấm [OK] để tiếp tục tạo Eccount, [Cancel] để hủy.",
                        "Kiểm tra lại",
                        MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Warning);

                    if (result != DialogResult.OK) return;
                }
            }
            catch { }

            var details = _poSvc.GetDetails(_selectedPO_ID);

            Supplier supp = null;
            if (poHead != null)
            {
                supp = _allSuppliers.Find(s => s.Supplier_ID == poHead.Supplier_ID);
            }
            if (supp == null) supp = new Supplier();

            // Truyền schedules để popup dùng Amount_Plan thay vì tính lại
            var schedules = _allSchedulesCache.ContainsKey(_selectedPO_ID)
                ? _allSchedulesCache[_selectedPO_ID]
                : new List<PaymentSchedule>();

            using var dlg = new frmPaymentRequestPreview(po, mprNo, details, supp, schedules);
            dlg.ShowDialog();
        }

        // =====================================================================
        //  EVENTS — Tab Debt
        // =====================================================================
        private void BtnViewDebt_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("PAYMENT", "Xem báo cáo", "Xem báo cáo")) return;
            try
            {
                int? suppId = null;
                if (cboSuppFilter.SelectedIndex > 0)
                {
                    var name = cboSuppFilter.SelectedItem.ToString();
                    var s = _allSuppliers.Find(x => (x.Company_Name ?? x.Supplier_Name) == name);
                    if (s != null) suppId = s.Supplier_ID;
                }

                _debtReport       = _svc.GetDebtReport(dtpFrom.Value, dtpTo.Value, suppId);
                _suppDebt         = _svc.GetSupplierDebt();
                _zaloPaidCache    = LoadZaloPaidCache();
                _zaloInRangeCache = LoadZaloInRangeCache(dtpFrom.Value, dtpTo.Value);

                // Lọc NCC theo suppId rồi bind + cập nhật cards
                FilterAndBindDebt();
            }
            catch (Exception ex) { Err(ex.Message); }
        }

        // Lọc + bind lại cả 2 grid theo các bộ lọc hiện tại (không query DB lại)
        private void FilterAndBindDebt()
        {
            if (_suppDebt == null || _debtReport == null) return;

            int? suppId = null;
            if (cboSuppFilter?.SelectedIndex > 0)
            {
                var name = cboSuppFilter.SelectedItem.ToString();
                var s = _allSuppliers.Find(x => (x.Company_Name ?? x.Supplier_Name) == name);
                if (s != null) suppId = s.Supplier_ID;
            }

            string statusFilter = cboDebtStatus?.SelectedItem?.ToString() ?? "Tất cả";
            bool overdueOnly    = chkOverdueOnly?.Checked ?? false;
            string search       = txtDebtSearch?.Text.Trim().ToLower() ?? "";

            // ── Lọc _debtReport để bind Detail ──
            var filteredDetail = _debtReport.FindAll(d =>
            {
                if (suppId.HasValue && d.Supplier_ID != suppId.Value) return false;
                if (overdueOnly && !d.Is_Overdue) return false;
                if (statusFilter != "Tất cả")
                {
                    string zaloSt = d.Is_Overdue ? "⚠ Quá hạn" : GetZaloStatus(d.PONo, d.Total_Amount);
                    if (zaloSt != statusFilter) return false;
                }
                if (!string.IsNullOrEmpty(search))
                {
                    string wono = _poSummaries.Find(p => p.PO_ID == d.PO_ID)?.WorkorderNo ?? "";
                    if (!d.PONo.ToLower().Contains(search) &&
                        !wono.ToLower().Contains(search) &&
                        !(d.Project_Name ?? "").ToLower().Contains(search)) return false;
                }
                return true;
            });

            // ── Lọc _suppDebt để bind NCC ──
            // Lấy tập Supplier_ID còn xuất hiện trong filteredDetail
            var activeSuppIds = new HashSet<int>(filteredDetail.Select(d => d.Supplier_ID));
            var filteredSupp = _suppDebt.FindAll(s =>
            {
                if (!activeSuppIds.Contains(s.Supplier_ID)) return false;
                if (suppId.HasValue && s.Supplier_ID != suppId.Value) return false;
                if (overdueOnly && s.Overdue_PO_Count == 0) return false;
                return true;
            });

            // ── Bind NCC grid — tính Đã TT và Còn nợ từ Zalo cache ──
            dgvDebtSupp.Rows.Clear();
            decimal tVal = 0, tPaid = 0, tDebt = 0; int tOver = 0;
            foreach (var s in filteredSupp)
            {
                // Tổng tiền đã TT (Zalo) cho tất cả PO của NCC này trong filteredDetail
                decimal suppZaloPaid = filteredDetail
                    .Where(d => d.Supplier_ID == s.Supplier_ID)
                    .Sum(d =>
                    {
                        string k = (d.PONo ?? "").ToUpperInvariant()
                            .Replace(" ", "").Replace("\t", "").Replace("\r", "").Replace("\n", "");
                        _zaloPaidCache.TryGetValue(k, out decimal p);
                        return p;
                    });
                decimal suppZaloDebt = filteredDetail
                    .Where(d => d.Supplier_ID == s.Supplier_ID)
                    .Sum(d =>
                    {
                        string k = (d.PONo ?? "").ToUpperInvariant()
                            .Replace(" ", "").Replace("\t", "").Replace("\r", "").Replace("\n", "");
                        _zaloPaidCache.TryGetValue(k, out decimal p);
                        return Math.Max(d.Total_Amount - p, 0);
                    });

                int i = dgvDebtSupp.Rows.Add();
                var r = dgvDebtSupp.Rows[i];
                r.Cells["D_SuppID"].Value  = s.Supplier_ID;
                r.Cells["D_Name"].Value    = s.Supplier_Name;
                r.Cells["D_TotalPO"].Value = s.Total_PO;
                r.Cells["D_Value"].Value   = FormatAmt(s.Total_PO_Value);
                r.Cells["D_Paid"].Value    = FormatAmt(suppZaloPaid);
                r.Cells["D_Debt"].Value    = FormatAmt(suppZaloDebt);
                r.Cells["D_Overdue"].Value = s.Overdue_PO_Count > 0 ? $"⚠ {s.Overdue_PO_Count}" : "—";

                tVal  += s.Total_PO_Value;
                tPaid += suppZaloPaid;
                tDebt += suppZaloDebt;
                tOver += s.Overdue_PO_Count;
            }

            // Cards tổng kết
            if (lblSumValue  != null && lblSumValue.Visible)  lblSumValue.Text  = $"{FormatAmt(tVal)} VNĐ";
            if (lblSumPaid   != null && lblSumPaid.Visible)   lblSumPaid.Text   = $"{FormatAmt(tPaid)} VNĐ";
            if (lblSumDebt   != null && lblSumDebt.Visible)   lblSumDebt.Text   = $"{FormatAmt(tDebt)} VNĐ";
            if (lblSumOverdue != null)                         lblSumOverdue.Text = $"{tOver} PO";

            // ── Bind Detail grid ──
            BindDebtDetail(filteredDetail);
        }

        private void DgvDebtSupp_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvDebtSupp.SelectedRows.Count == 0) return;
            int sid = Convert.ToInt32(dgvDebtSupp.SelectedRows[0].Cells["D_SuppID"].Value);

            // Giữ nguyên bộ lọc status/overdue/search khi click NCC
            string statusFilter = cboDebtStatus?.SelectedItem?.ToString() ?? "Tất cả";
            bool overdueOnly    = chkOverdueOnly?.Checked ?? false;
            string search       = txtDebtSearch?.Text.Trim().ToLower() ?? "";

            var items = _debtReport.FindAll(d =>
            {
                if (d.Supplier_ID != sid) return false;
                if (overdueOnly && !d.Is_Overdue) return false;
                if (statusFilter != "Tất cả")
                {
                    string zaloSt = d.Is_Overdue ? "⚠ Quá hạn" : GetZaloStatus(d.PONo, d.Total_Amount);
                    if (zaloSt != statusFilter) return false;
                }
                if (!string.IsNullOrEmpty(search))
                {
                    string wono = _poSummaries.Find(p => p.PO_ID == d.PO_ID)?.WorkorderNo ?? "";
                    if (!d.PONo.ToLower().Contains(search) &&
                        !wono.ToLower().Contains(search) &&
                        !(d.Project_Name ?? "").ToLower().Contains(search)) return false;
                }
                return true;
            });
            BindDebtDetail(items);
        }

        private void DgvDebtSupp_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string col = dgvDebtSupp.Columns[e.ColumnIndex].Name;
            if (col == "D_Paid")
            { e.CellStyle.ForeColor = Color.FromArgb(40, 167, 69); e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold); }
            if (col == "D_Debt")
            { e.CellStyle.ForeColor = Color.FromArgb(220, 53, 69); e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold); }
            if (col == "D_Overdue" && e.Value?.ToString() != "—")
            { e.CellStyle.ForeColor = Color.FromArgb(220, 53, 69); e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold); }
        }

        private void BindDebtDetail(List<DebtReportItem> items)
        {
            dgvDebtDetail.Rows.Clear();
            foreach (var d in items)
            {
                int i = dgvDebtDetail.Rows.Add();
                var r = dgvDebtDetail.Rows[i];
                // Lấy mã dự án (WorkorderNo) từ _poSummaries; fallback về Project_Name
                string workorderNo = _poSummaries.Find(p => p.PO_ID == d.PO_ID)?.WorkorderNo ?? d.Project_Name;
                r.Cells["DD_POID"].Value    = d.PO_ID;
                r.Cells["DD_PONo"].Value    = d.PONo;
                r.Cells["DD_Project"].Value = workorderNo;
                r.Cells["DD_PODate"].Value  = d.PO_Date?.ToString("dd/MM/yyyy") ?? "";
                // Lấy số tiền đã TT từ Zalo: toàn thời gian và trong kỳ
                string poKey = (d.PONo ?? "").ToUpperInvariant()
                    .Replace(" ", "").Replace("\t", "").Replace("\r", "").Replace("\n", "");
                _zaloPaidCache.TryGetValue(poKey, out decimal zaloPaidTotal);
                _zaloInRangeCache.TryGetValue(poKey, out decimal zaloInRange);
                decimal zaloBefore = Math.Max(zaloPaidTotal - zaloInRange, 0);
                decimal zaloRemain = Math.Max(d.Total_Amount - zaloPaidTotal, 0);

                r.Cells["DD_Total"].Value   = FormatAmt(d.Total_Amount);
                r.Cells["DD_Before"].Value  = FormatAmt(zaloBefore);
                r.Cells["DD_InRange"].Value = FormatAmt(zaloInRange);
                r.Cells["DD_Remain"].Value  = FormatAmt(zaloRemain);

                // % đã thanh toán (dựa trên tổng Zalo)
                decimal pct = d.Total_Amount > 0 ? Math.Round(zaloPaidTotal / d.Total_Amount * 100, 1) : 0;
                r.Cells["DD_Percent"].Value = $"{pct}%";

                string ddStatus = d.Is_Overdue
                    ? "⚠ Quá hạn"
                    : GetZaloStatus(d.PONo, d.Total_Amount);
                r.Cells["DD_Status"].Value  = ddStatus;
                r.Cells["DD_Due"].Value     = d.Next_Due_Date?.ToString("dd/MM/yyyy") ?? "—";
            }
        }

        private void DgvDebtDetail_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string col = dgvDebtDetail.Columns[e.ColumnIndex].Name;
            if (col == "DD_Status")
            {
                string v = e.Value?.ToString() ?? "";
                e.CellStyle.ForeColor =
                    v.Contains("Quá hạn")       ? Color.FromArgb(220, 53, 69) :
                    v == "Đã thanh toán"         ? Color.FromArgb(40, 167, 69) :
                    v == "Thanh toán 1 phần"     ? Color.FromArgb(255, 140, 0) :
                                                   Color.FromArgb(0, 120, 212);   // Pending
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
            if (col == "DD_Remain")
            {
                e.CellStyle.ForeColor = Color.FromArgb(220, 53, 69);
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
            if (col == "DD_Percent")
            {
                // Màu theo % thanh toán: xanh ≥100%, cam 1-99%, đỏ 0%
                string pctStr = e.Value?.ToString()?.TrimEnd('%') ?? "0";
                if (decimal.TryParse(pctStr, out decimal pct))
                {
                    e.CellStyle.ForeColor = pct >= 100 ? Color.FromArgb(40, 167, 69) :
                                            pct > 0    ? Color.FromArgb(255, 140, 0) :
                                                         Color.FromArgb(220, 53, 69);
                    e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                }
            }
            if (col == "DD_Due")
            {
                // Tô đỏ nếu đến hạn đã qua
                string dateStr = e.Value?.ToString() ?? "";
                if (DateTime.TryParseExact(dateStr, "dd/MM/yyyy",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var due) && due < DateTime.Today)
                {
                    e.CellStyle.ForeColor = Color.FromArgb(220, 53, 69);
                    e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                }
            }
        }

        private void DgvDebtDetail_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string poNo = dgvDebtDetail.Rows[e.RowIndex].Cells["DD_PONo"].Value?.ToString() ?? "";
            if (string.IsNullOrEmpty(poNo)) return;

            // Chuyển sang Tab PO và lọc theo PO No
            tabs.SelectedTab = tabPO;
            txtSearchPO.Text = poNo;
            FilterAndBind();

            // Tự động chọn dòng đầu nếu tìm thấy đúng 1 PO
            if (dgvPO.Rows.Count == 1)
            {
                dgvPO.ClearSelection();
                dgvPO.Rows[0].Selected = true;
            }
        }

        // =====================================================================
        //  XUẤT EXCEL
        // =====================================================================
        private void BtnExportDebt_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("PAYMENT", "Xuất Excel", "Xuất Excel")) return;
            if (_debtReport.Count == 0) { Warn("Vui lòng xem báo cáo trước!"); return; }
            using var sfd = new SaveFileDialog
            {
                Title = "Lưu báo cáo công nợ",
                Filter = "Excel|*.xlsx",
                FileName = $"CongNo_{dtpFrom.Value:yyyyMMdd}_{dtpTo.Value:yyyyMMdd}",
                InitialDirectory = Directory.Exists(@"D:\RÁC") ? @"D:\RÁC" : Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };
            if (sfd.ShowDialog() != DialogResult.OK) return;
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using var pkg = new ExcelPackage();
                var ws = pkg.Workbook.Worksheets.Add("Công nợ NCC");

                ws.Cells[1, 1].Value = "BÁO CÁO CÔNG NỢ NHÀ CUNG CẤP";
                ws.Cells[1, 1, 1, 9].Merge = true;
                ws.Cells[1, 1].Style.Font.Size = 14;
                ws.Cells[1, 1].Style.Font.Bold = true;
                ws.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                ws.Cells[2, 1].Value = $"Kỳ: {dtpFrom.Value:dd/MM/yyyy} — {dtpTo.Value:dd/MM/yyyy}";
                ws.Cells[2, 1, 2, 9].Merge = true;

                string[] hdrs = { "Nhà cung cấp", "PO No", "Dự án", "Ngày PO",
                                   "Giá trị PO", "TT trước kỳ", "TT trong kỳ", "Còn nợ", "Trạng thái" };
                for (int c = 0; c < hdrs.Length; c++)
                {
                    ws.Cells[4, c + 1].Value = hdrs[c];
                    ws.Cells[4, c + 1].Style.Font.Bold = true;
                    ws.Cells[4, c + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    ws.Cells[4, c + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 120, 212));
                    ws.Cells[4, c + 1].Style.Font.Color.SetColor(Color.White);
                }

                int row = 5;
                foreach (var d in _debtReport)
                {
                    ws.Cells[row, 1].Value = d.Supplier_Name;
                    ws.Cells[row, 2].Value = d.PONo;
                    ws.Cells[row, 3].Value = d.Project_Name;
                    ws.Cells[row, 4].Value = d.PO_Date?.ToString("dd/MM/yyyy") ?? "";
                    ws.Cells[row, 5].Value = d.Total_Amount; ws.Cells[row, 5].Style.Numberformat.Format = "#,##0.##";
                    ws.Cells[row, 6].Value = d.Paid_Before_Range; ws.Cells[row, 6].Style.Numberformat.Format = "#,##0.##";
                    ws.Cells[row, 7].Value = d.Paid_In_Range; ws.Cells[row, 7].Style.Numberformat.Format = "#,##0.##";
                    ws.Cells[row, 8].Value = d.Remaining_Debt; ws.Cells[row, 8].Style.Numberformat.Format = "#,##0.##";
                    ws.Cells[row, 9].Value = d.Is_Overdue ? "⚠ Quá hạn" : d.Payment_Status;
                    if (d.Is_Overdue)
                    {
                        ws.Cells[row, 1, row, 9].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        ws.Cells[row, 1, row, 9].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 230, 230));
                    }
                    row++;
                }

                ws.Cells[ws.Dimension.Address].AutoFitColumns();
                pkg.SaveAs(new System.IO.FileInfo(sfd.FileName));
                MessageBox.Show(TopOwner, "✅ Xuất Excel thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = sfd.FileName, UseShellExecute = true });
            }
            catch (Exception ex) { Err(ex.Message); }
        }

private void ResizeAll()
{
    try
    {
        // Use tabs.ClientSize for accurate dimensions of both tabs
        int w = tabs.ClientSize.Width;
        int h = tabs.ClientSize.Height;

        if (panelTop != null) panelTop.Width = w - 10;
        if (panelInfo != null) { panelInfo.Width = w - 10; lblPOStatus.Left = panelInfo.Width - 205; }

        int leftW = w / 2 - 8;

        // panelSched fixed height 200, full width on the left
        if (panelSched != null)
        {
            const int docW = 200;
            panelSched.Width = leftW;
            panelSched.Height = 200;
            dgvSchedule.Width = panelSched.Width - docW - 15;
            dgvSchedule.Height = panelSched.Height - 62;
            if (dgvDoc != null)
            {
                dgvDoc.Left = panelSched.Width - docW - 5;
                dgvDoc.Width = docW - 2;
                dgvDoc.Height = panelSched.Height - 30;
                foreach (Control c in panelSched.Controls)
                    if (c is Label lbl && lbl.Text.Contains("Document"))
                        lbl.Left = dgvDoc.Left;
            }
        }

        // panelPrintHistory moved to below panelSched (left side)
        if (panelPrintHistory != null)
        {
            panelPrintHistory.Top = (panelSched?.Bottom ?? 0) + 5;
            panelPrintHistory.Left = 5;
            panelPrintHistory.Width = leftW;
            panelPrintHistory.Height = Math.Max(100, h - panelPrintHistory.Top - 10);
            dgvPrintHistory.Width = panelPrintHistory.Width - 10;
            dgvPrintHistory.Height = panelPrintHistory.Height - 63;
            if (_phDateTo != null)
                _phDateTo.Width = Math.Min(115, (panelPrintHistory.Width - 470) / 2);
        }

        // panelHist now sits on the right side, occupying the full height
        if (panelHist != null)
        {
            panelHist.Left = w / 2 + 3;
            panelHist.Top = 317;
            panelHist.Width = w / 2 - 8;
            panelHist.Height = Math.Max(200, h - panelHist.Top - 10);
            dgvHistory.Width = panelHist.Width - 10;
            dgvHistory.Height = panelHist.Height - 62;
        }

                // ── Tab Debt: pNCC chiếm 50% width, pDet chiếm phần còn lại ──
                if (_pNCC != null && _pDet != null)
                {
                    int nccW = (int)(w * 0.50);
                    int detLeft = nccW + 10;
                    int detW = w - detLeft - 5;
                    int panelTop = 179;
                    int panelH = Math.Max(100, h - panelTop - 5);

                    _pNCC.Left = 5;
                    _pNCC.Width = nccW;
                    _pNCC.Height = panelH;

                    _pDet.Left = detLeft;
                    _pDet.Width = detW;
                    _pDet.Height = panelH;

                    dgvDebtSupp.Width = _pNCC.Width - 10;
                    dgvDebtSupp.Height = _pNCC.Height - 33;
                    dgvDebtDetail.Width = _pDet.Width - 10;
                    dgvDebtDetail.Height = _pDet.Height - 33;
                }
            }
            catch { }
        }

        private DataGridView Grid(Panel parent, int top, int height)
        {
            var dgv = new DataGridView
            {
                Location = new Point(5, top),
                Size = new Size(parent.Width - 10, height > 0 ? height : parent.Height - top - 5),
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
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

        private Panel P(Control parent, int x, int y, int w, int h, Color bg)
        {
            var p = new Panel
            {
                Location = new Point(x, y),
                Size = new Size(w > 0 ? w : parent.ClientSize.Width - x - 5,
                                       h > 0 ? h : parent.ClientSize.Height - y - 5),
                BackColor = bg,
                BorderStyle = BorderStyle.FixedSingle
            };
            parent.Controls.Add(p);
            return p;
        }

        private void Lbl(Control parent, string text, int x, int y, int w, int h,
                          bool bold = false, Color? color = null)
        {
            parent.Controls.Add(new Label
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(w, h),
                Font = new Font("Segoe UI", 9, bold ? FontStyle.Bold : FontStyle.Regular),
                ForeColor = color ?? Color.FromArgb(50, 50, 50)
            });
        }

        private TextBox Txt(Control parent, int x, int y, int w)
        {
            var t = new TextBox { Location = new Point(x, y), Size = new Size(w, 26), Font = new Font("Segoe UI", 9) };
            parent.Controls.Add(t);
            return t;
        }

        private ComboBox Cbo(Control parent, int x, int y, int w, string[] items)
        {
            var c = new ComboBox
            {
                Location = new Point(x, y),
                Size = new Size(w, 26),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            c.Items.AddRange(items);
            c.SelectedIndex = 0;
            parent.Controls.Add(c);
            return c;
        }

        private Button Btn(string text, Color color, int x, int y, int w, int h)
        {
            var b = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(w, h),
                BackColor = color,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            b.FlatAppearance.BorderSize = 0;
            return b;
        }

        private Label InfoLbl(Panel p, string text, int x, int y, int w, int h, float size, bool bold)
        {
            var l = new Label
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(w, h),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", size, bold ? FontStyle.Bold : FontStyle.Regular)
            };
            p.Controls.Add(l);
            return l;
        }

        private Label Card(Panel parent, int x, string title, Color color)
        {
            var card = new Panel { Location = new Point(x, 5), Size = new Size(210, 60), BackColor = color };
            parent.Controls.Add(card);
            card.Controls.Add(new Label
            {
                Text = title,
                Location = new Point(5, 3),
                Size = new Size(200, 18),
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                ForeColor = Color.FromArgb(210, 255, 255, 255),
                TextAlign = ContentAlignment.MiddleCenter
            });
            var val = new Label
            {
                Text = "—",
                Location = new Point(5, 22),
                Size = new Size(200, 32),
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleCenter
            };
            card.Controls.Add(val);
            return val;
        }

        private Form TopOwner
        {
            get
            {
                if (this.InvokeRequired)
                {
                    return (Form)this.Invoke(new Func<Form>(() => TopOwner));
                }
                var f = (this.TopLevelControl as Form) ?? this;
                if (!f.IsDisposed)
                {
                    f.BringToFront();
                    f.Activate();
                }
                return f;
            }
        }
        private void Warn(string msg)
        {
            if (this.InvokeRequired) { this.Invoke(new Action(() => Warn(msg))); return; }
            var f = TopOwner; MessageBox.Show(f, msg, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        private void Err(string msg)
        {
            if (this.InvokeRequired) { this.Invoke(new Action(() => Err(msg))); return; }
            var f = TopOwner; MessageBox.Show(f, "Lỗi: " + msg, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        private bool Ask(string msg)
        {
            if (this.InvokeRequired) return (bool)this.Invoke(new Func<bool>(() => Ask(msg)));
            var f = TopOwner; return MessageBox.Show(f, msg, "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
        }
        private void Info(string msg, string title = "Thông báo")
        {
            if (this.InvokeRequired) { this.Invoke(new Action(() => Info(msg, title))); return; }
            var f = TopOwner; MessageBox.Show(f, msg, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private bool AskYN(string msg, string title = "Xác nhận")
        {
            if (this.InvokeRequired) return (bool)this.Invoke(new Func<bool>(() => AskYN(msg, title)));
            var f = TopOwner; return MessageBox.Show(f, msg, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
        }

        // Được gọi từ frmPrintPreview khi user chọn OK cập nhật lịch sử
        public void AddPrintHistory(string poNo, string project, List<PaymentSchedule> scheds, string docNames = "")
        {
            if (dgvPrintHistory == null) return;
            string dateStr = DateTime.Now.ToString("dd/MM/yyyy HH:mm");

            foreach (var s in scheds)
            {
                decimal net = s.Amount_Plan;
                decimal vat = Math.Round(net * 0.1m, 0);
                decimal total = Math.Round(net + vat, 0);
                string dot = s.Dot_TT == 1 ? "1st" : s.Dot_TT == 2 ? "2nd" :
                                s.Dot_TT == 3 ? "3rd" : $"{s.Dot_TT}th";

                // ── Lấy Short_Name NCC (dùng cho cả DB và grid) ──
                string suppShortDb = "";
                try
                {
                    var pSum = _poSummaries.Find(p => p.PONo == poNo);
                    if (pSum != null)
                    {
                        var suppObj = _allSuppliers.Find(x => x.Supplier_ID == pSum.Supplier_ID);
                        suppShortDb = suppObj?.Short_Name ?? suppObj?.Company_Name ?? suppObj?.Supplier_Name ?? "";
                    }
                }
                catch { }

                // ── Lưu vào DB ──
                try
                {
                    using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                    conn.Open();
                    var cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                        INSERT INTO PO_PrintRequestHistory
                            (PONo, Project_Name, Dot_TT, Dot_Label,
                             Amount_Net, Amount_VAT, Amount_Total,
                             Supplier_Short, Printed_By, Printed_Date, Documents_List)
                        VALUES
                            (@poNo, @proj, @dot, @dotLabel,
                             @net, @vat, @total,
                             @suppShort, @by, GETDATE(), @docs)", conn);
                    cmd.Parameters.AddWithValue("@poNo", poNo);
                    cmd.Parameters.AddWithValue("@proj", project ?? "");
                    cmd.Parameters.AddWithValue("@dot", s.Dot_TT);
                    cmd.Parameters.AddWithValue("@dotLabel", dot);
                    cmd.Parameters.AddWithValue("@net", net);
                    cmd.Parameters.AddWithValue("@vat", vat);
                    cmd.Parameters.AddWithValue("@total", total);
                    cmd.Parameters.AddWithValue("@suppShort", suppShortDb);
                    cmd.Parameters.AddWithValue("@by", _currentUser ?? "");
                    cmd.Parameters.AddWithValue("@docs", docNames ?? "");
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("AddPrintHistory DB error: " + ex.Message);
                }

                // ── Thêm vào đầu grid (mới nhất lên trên) ──
                dgvPrintHistory.Rows.Insert(0);
                dgvPrintHistory.Rows[0].Cells["PH_ID"].Value = DBNull.Value;
                dgvPrintHistory.Rows[0].Cells["PH_PONo"].Value = poNo;
                dgvPrintHistory.Rows[0].Cells["PH_Supp"].Value = suppShortDb;
                dgvPrintHistory.Rows[0].Cells["PH_Project"].Value = project;
                dgvPrintHistory.Rows[0].Cells["PH_Dot"].Value = dot;
                dgvPrintHistory.Rows[0].Cells["PH_Net"].Value = FormatAmt(net);
                dgvPrintHistory.Rows[0].Cells["PH_Vat"].Value = FormatAmt(vat);
                dgvPrintHistory.Rows[0].Cells["PH_Total"].Value = FormatAmt(total);
                dgvPrintHistory.Rows[0].Cells["PH_Date"].Value = dateStr;
            }

            if (dgvPrintHistory.Rows.Count > 0)
                dgvPrintHistory.FirstDisplayedScrollingRowIndex = 0; // cuộn lên đầu — mới nhất
        }

        // Load lịch sử 3 tháng gần nhất từ DB
        private void LoadPrintHistory(DateTime? from = null, DateTime? to = null)
        {
            if (dgvPrintHistory == null) return;
            dgvPrintHistory.Rows.Clear();
            DateTime dtFrom = from ?? DateTime.Today.AddYears(-2);
            DateTime dtTo = to ?? DateTime.Today.AddDays(1).AddSeconds(-1);

            // Từ khoá lọc NCC (không phân biệt hoa/thường)
            string nccFilter = _txtPhNCC?.Text.Trim() ?? "";

            try
            {
                using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                conn.Open();
                var cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                    SELECT Print_ID, PONo,
                           ISNULL(Supplier_Short, '') AS Supplier_Short,
                           ISNULL(Project_Name,   '') AS Project_Name,
                           ISNULL(Dot_Label,      '') AS Dot_Label,
                           ISNULL(Amount_Net,     0)  AS Amount_Net,
                           ISNULL(Amount_VAT,     0)  AS Amount_VAT,
                           ISNULL(Amount_Total,   0)  AS Amount_Total,
                           ISNULL(Printed_By,     '') AS Printed_By,
                           CONVERT(NVARCHAR(16), Printed_Date, 103) + ' '
                           + SUBSTRING(CONVERT(NVARCHAR(8), Printed_Date, 108), 1, 5) AS Printed_Date
                    FROM PO_PrintRequestHistory
                    WHERE Printed_Date BETWEEN @from AND @to
                    ORDER BY Printed_Date DESC", conn);
                cmd.Parameters.AddWithValue("@from", dtFrom);
                cmd.Parameters.AddWithValue("@to", dtTo);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    decimal net = Convert.ToDecimal(reader["Amount_Net"]);
                    decimal vat = Convert.ToDecimal(reader["Amount_VAT"]);
                    decimal total = Convert.ToDecimal(reader["Amount_Total"]);

                    // Nếu Supplier_Short rỗng, thử tra từ _allSuppliers qua PONo
                    string suppShort = reader["Supplier_Short"]?.ToString() ?? "";
                    if (string.IsNullOrEmpty(suppShort))
                    {
                        string poNo = reader["PONo"]?.ToString() ?? "";
                        var pSum = _poSummaries.Find(p => p.PONo == poNo);
                        if (pSum != null)
                        {
                            var suppObj = _allSuppliers.Find(x => x.Supplier_ID == pSum.Supplier_ID);
                            suppShort = suppObj?.Short_Name ?? suppObj?.Company_Name ?? suppObj?.Supplier_Name ?? "";
                        }
                    }

                    // ── Lọc theo NCC ──
                    if (!string.IsNullOrEmpty(nccFilter) &&
                        !suppShort.Contains(nccFilter, StringComparison.OrdinalIgnoreCase))
                        continue;

                    int i = dgvPrintHistory.Rows.Add();
                    dgvPrintHistory.Rows[i].Cells["PH_ID"].Value = reader["Print_ID"];
                    dgvPrintHistory.Rows[i].Cells["PH_PONo"].Value = reader["PONo"]?.ToString() ?? "";
                    dgvPrintHistory.Rows[i].Cells["PH_Supp"].Value = suppShort;
                    dgvPrintHistory.Rows[i].Cells["PH_Project"].Value = reader["Project_Name"]?.ToString() ?? "";
                    dgvPrintHistory.Rows[i].Cells["PH_Dot"].Value = reader["Dot_Label"]?.ToString() ?? "";
                    dgvPrintHistory.Rows[i].Cells["PH_Net"].Value = FormatAmt(net);
                    dgvPrintHistory.Rows[i].Cells["PH_Vat"].Value = FormatAmt(vat);
                    dgvPrintHistory.Rows[i].Cells["PH_Total"].Value = FormatAmt(total);
                    dgvPrintHistory.Rows[i].Cells["PH_Date"].Value = reader["Printed_Date"]?.ToString() ?? "";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(TopOwner,
                    "Lỗi tải lịch sử in Request:\n" + ex.Message,
                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnDeletePrintHistory_Click(object sender, EventArgs e)
        {
            if (dgvPrintHistory.SelectedRows.Count == 0 && dgvPrintHistory.CurrentRow == null)
            { Warn("Vui lòng chọn dòng cần xóa!"); return; }

            var row = dgvPrintHistory.SelectedRows.Count > 0
                ? dgvPrintHistory.SelectedRows[0]
                : dgvPrintHistory.CurrentRow;

            string poNo = row.Cells["PH_PONo"].Value?.ToString() ?? "";
            string date = row.Cells["PH_Date"].Value?.ToString() ?? "";
            int printId = row.Cells["PH_ID"].Value != null &&
                             row.Cells["PH_ID"].Value != DBNull.Value
                             ? Convert.ToInt32(row.Cells["PH_ID"].Value) : 0;

            if (MessageBox.Show(TopOwner,
                $"Xóa lịch sử in Request này?\n\nPO: {poNo}\nNgày in: {date}",
                "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                return;

            try
            {
                if (printId > 0)
                {
                    using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                    conn.Open();
                    var cmd = new Microsoft.Data.SqlClient.SqlCommand(
                        "DELETE FROM PO_PrintRequestHistory WHERE Print_ID = @id", conn);
                    cmd.Parameters.AddWithValue("@id", printId);
                    cmd.ExecuteNonQuery();
                }
                // Xóa khỏi grid dù có ID hay không
                dgvPrintHistory.Rows.Remove(row);
                MessageBox.Show(TopOwner, "✅ Đã xóa thành công!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { Err("Lỗi xóa: " + ex.Message); }
        }

        // Kiểm tra PO đã in request chưa (trong 3 tháng gần nhất)
        private bool CheckAlreadyPrinted(string poNo, out string lastPrintDate)
        {
            lastPrintDate = "";
            try
            {
                using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                conn.Open();
                var cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                    SELECT TOP 1
                        CONVERT(NVARCHAR(16), Printed_Date, 103) + ' '
                        + SUBSTRING(CONVERT(NVARCHAR(8), Printed_Date, 108), 1, 5) AS LastDate,
                        Printed_By
                    FROM PO_PrintRequestHistory
                    WHERE PONo = @poNo
                      AND Printed_Date >= DATEADD(MONTH, -3, GETDATE())
                    ORDER BY Printed_Date DESC", conn);
                cmd.Parameters.AddWithValue("@poNo", poNo);
                using var reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    lastPrintDate = $"{reader["LastDate"]}  (bởi: {reader["Printed_By"]})";
                    return true;
                }
            }
            catch { }
            return false;
        }

        // ─────────────────────────────────────────────────────────────────────
        //  Chuẩn hoá số PO — bỏ prefix emoji và suffix ghi chú hiển thị
        //  Ví dụ: "🔥 DV-001-002-003 (Mới)" → "DV-001-002-003"
        // ─────────────────────────────────────────────────────────────────────
        private static string CleanPONo(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return "";
            string s = raw.Trim();

            // Bỏ ký tự đầu không phải chữ/số (emoji, dấu cách…)
            while (s.Length > 0 && !char.IsLetterOrDigit(s[0]))
                s = s.Substring(1).TrimStart();

            // Bỏ hậu tố từ dấu '(' trở đi — format PO (DV-xxx) không chứa ngoặc đơn
            int paren = s.IndexOf('(');
            if (paren > 0)
                s = s.Substring(0, paren).TrimEnd();

            return s;
        }

        // =====================================================================
        //  SHOW PRINT HISTORY AND OPTIONS — Khi bấn "In" từ dgvHistory
        // =====================================================================
        private void ShowPrintHistoryAndOptions(string poNo)
        {
            // Chuẩn hoá: loại bỏ decoration "(Mới)", emoji… nếu có
            poNo = CleanPONo(poNo);

            var popup = new Form
            {
                Text = $"🖨 Lịch sử in & Tùy chọn in — PO: {poNo}",
                Size = new Size(1100, 700),
                MinimumSize = new Size(900, 550),
                StartPosition = FormStartPosition.CenterParent,
                BackColor = Color.FromArgb(245, 245, 245)
            };

            var pTop = new Panel
            {
                Dock = DockStyle.Top,
                Height = 50,
                BackColor = Color.White,
                Padding = new Padding(10)
            };
            popup.Controls.Add(pTop);

            Lbl(pTop, $"Lịch sử in của PO: {poNo}", 10, 10, 500, 20, true, Color.FromArgb(0, 120, 212));

var btnPrintAll = Btn("🖨 In Tất cả", Color.FromArgb(0, 120, 212), 600, 10, 120, 30);
btnPrintAll.Click += (s, e) =>
{
    // ── Bước 1: Chọn đúng PO & nạp tài liệu đính kèm ──
    var po = _poSummaries.FirstOrDefault(p =>
        string.Equals(p.PONo?.Trim(), poNo?.Trim(), StringComparison.OrdinalIgnoreCase));
    if (po != null)
    {
        _selectedPO_ID = po.PO_ID;
        LoadDocuments();
    }

    // Kiểm tra file tồn tại
    bool anyFileMissing = false;
    var docs = new List<string>();
    if (dgvDoc == null || dgvDoc.Rows.Count == 0)
    {
        anyFileMissing = true;
    }
    else
    {
        foreach (DataGridViewRow row in dgvDoc.Rows)
        {
            string path = row.Cells["Doc_Path"].Value?.ToString() ?? "";
            if (System.IO.File.Exists(path))
                docs.Add(path);
            else
                anyFileMissing = true;
        }
    }

    if (anyFileMissing)
    {
        MessageBox.Show(
            "Một số file tài liệu (Invoice / Delivery Note) không tồn tại trên hệ thống.\nHệ thống sẽ tiếp tục in Payment Request.",
            "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    // ── Bước 2: In Payment Request (nếu Hủy → dừng toàn bộ) ──
    if (!PrintPaymentRequest()) return;

    // ── Bước 3: In tài liệu đính kèm (nếu có) ──
    if (docs.Count > 0)
        PrintDocFiles(docs);

    // ── Bước 4: Lưu lịch sử in ──
    if (po != null)
    {
        var scheds = _allSchedulesCache.ContainsKey(po.PO_ID)
            ? _allSchedulesCache[po.PO_ID]
            : new List<PaymentSchedule>();
        string docNames = docs.Count > 0
            ? string.Join(", ", docs.Select(System.IO.Path.GetFileName))
            : "";
        AddPrintHistory(poNo, po.Project_Name, scheds, docNames);
    }
    popup.Close();
};
            pTop.Controls.Add(btnPrintAll);

            var btnPrintReq = Btn("🖨 In Req", Color.FromArgb(0, 150, 100), 730, 10, 120, 30);
            btnPrintReq.Click += (s, e) =>
            {
                var po = _poSummaries.FirstOrDefault(p =>
                    string.Equals(p.PONo?.Trim(), CleanPONo(poNo).Trim(), StringComparison.OrdinalIgnoreCase));
                if (po != null)
                {
                    _selectedPO_ID = po.PO_ID;

                    // Logic in Request gốc
                    if (!PermissionHelper.Check("PAYMENT", "In Request", "In Request")) return;
                    if (CheckAlreadyPrinted(po.PONo, out string lastDate))
                    {
                        var ans = MessageBox.Show(popup,
                            $"⚠ PO \"{po.PONo}\" đã được in Request trước đó.\n" +
                            $"Lần in gần nhất: {lastDate}\n\n" +
                            "Bạn có muốn in lại không?",
                            "Đã in trước đó",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Warning);
                        if (ans != DialogResult.Yes) return;
                    }
                    if (PrintPaymentRequest())
                        popup.Close();
                }
            };
pTop.Controls.Add(btnPrintReq);

var btnPrintDoc = Btn("📎 In tài liệu", Color.FromArgb(40, 167, 69), 860, 10, 120, 30);
btnPrintDoc.Click += (s, e) =>
{
    // Ensure correct PO selection
    var po = _poSummaries.FirstOrDefault(p => 
        string.Equals(p.PONo?.Trim(), poNo?.Trim(), StringComparison.OrdinalIgnoreCase));
    if (po != null)
    {
        _selectedPO_ID = po.PO_ID;
        // Load latest documents for this PO
        LoadDocuments();
    }

    bool anyFileMissing = false;
    var docs = new List<string>();
    if (dgvDoc == null || dgvDoc.Rows.Count == 0)
    {
        anyFileMissing = true;
    }
    else
    {
        foreach (DataGridViewRow row in dgvDoc.Rows)
        {
            string path = row.Cells["Doc_Path"].Value?.ToString() ?? "";
            if (System.IO.File.Exists(path))
                docs.Add(path);
            else
                anyFileMissing = true;
        }
    }

    if (anyFileMissing)
    {
        MessageBox.Show(
            "Một số file tài liệu (Invoice / Delivery Note) không tồn tại trên hệ thống.\nHệ thống sẽ tiếp tục in Payment Request.",
            "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);

        // Nếu Hủy dialog in → không in docs còn lại
        if (!PrintPaymentRequest()) return;
        if (docs.Count > 0) PrintDocFiles(docs);
    }
    else
    {
        PrintDocFiles(docs);
    }
    popup.Close();
};
            pTop.Controls.Add(btnPrintDoc);

            var dgvHist = new DataGridView
            {
                Location = new Point(10, 70),
                Size = new Size(1070, 580),
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
            dgvHist.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvHist.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvHist.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvHist.EnableHeadersVisualStyles = false;
            dgvHist.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);

            dgvHist.Columns.Add(new DataGridViewTextBoxColumn { Name = "PH_Date", HeaderText = "Ngày in", Width = 200 });
            dgvHist.Columns.Add(new DataGridViewTextBoxColumn { Name = "PH_Dot", HeaderText = "Đợt", Width = 70 });
            dgvHist.Columns.Add(new DataGridViewTextBoxColumn { Name = "PH_Net", HeaderText = "Số tiền (Net)", Width = 120 });
            dgvHist.Columns.Add(new DataGridViewTextBoxColumn { Name = "PH_Total", HeaderText = "Tổng sau VAT", Width = 120 });
            dgvHist.Columns.Add(new DataGridViewTextBoxColumn { Name = "PH_Docs", HeaderText = "Tài liệu đã in", Width = 560 });

            dgvHist.CellFormatting += (s, ev) => {
                if (ev.RowIndex < 0) return;
                string col = dgvHist.Columns[ev.ColumnIndex].Name;
                if (col == "PH_Net" || col == "PH_Total") ev.CellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            };

            popup.Controls.Add(dgvHist);

            try {
                using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                conn.Open();
                var cmd = new Microsoft.Data.SqlClient.SqlCommand(
                    "SELECT Printed_Date, Dot_Label, Amount_Net, Amount_Total, Documents_List FROM PO_PrintRequestHistory WHERE PONo = @poNo ORDER BY Printed_Date DESC", conn);
                cmd.Parameters.AddWithValue("@poNo", poNo);
                using var reader = cmd.ExecuteReader();
                while (reader.Read()) {
                    int i = dgvHist.Rows.Add();
                    dgvHist.Rows[i].Cells["PH_Date"].Value = Convert.ToDateTime(reader["Printed_Date"]).ToString("dd/MM/yyyy HH:mm");
                    dgvHist.Rows[i].Cells["PH_Dot"].Value = reader["Dot_Label"];
                    dgvHist.Rows[i].Cells["PH_Net"].Value = FormatAmt(Convert.ToDecimal(reader["Amount_Net"]));
                    dgvHist.Rows[i].Cells["PH_Total"].Value = FormatAmt(Convert.ToDecimal(reader["Amount_Total"]));
                    dgvHist.Rows[i].Cells["PH_Docs"].Value = reader["Documents_List"];
                }
            } catch { }

            popup.ShowDialog();
        }

        private void btnPrintReq_Popup_Click(string poNo)
        {
            var po = _poSummaries.Find(p => (p.PONo ?? "") == poNo);
            if (po == null) return;

            // Load documents for current PO to check existence
            _selectedPO_ID = po.PO_ID;
            LoadDocuments();

            bool anyFileMissing = false;
            if (dgvDoc == null || dgvDoc.Rows.Count == 0)
            {
                anyFileMissing = true;
            }
            else
            {
                foreach (DataGridViewRow row in dgvDoc.Rows)
                {
                    string path = row.Cells["Doc_Path"].Value?.ToString() ?? "";
                    if (!System.IO.File.Exists(path))
                    {
                        anyFileMissing = true;
                        break;
                    }
                }
            }

            if (anyFileMissing)
            {
                MessageBox.Show(
                    "Một số file tài liệu (Invoice / Delivery Note) không tồn tại trên hệ thống.\nHệ thống sẽ tiếp tục in Payment Request.",
                    "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            var poHead = _poSvc.GetAll().Find(p => p.PO_ID == po.PO_ID);
            string mprNo = poHead?.MPR_No ?? "";
            var details = _poSvc.GetDetails(po.PO_ID);
            Supplier supp = _allSuppliers.Find(s => s.Supplier_ID == poHead?.Supplier_ID) ?? new Supplier();
            var schedules = _allSchedulesCache.ContainsKey(po.PO_ID) ? _allSchedulesCache[po.PO_ID] : new List<PaymentSchedule>();

            PrintPaymentRequest(); // nếu Hủy chỉ đóng dialog, không có doc nào cần bỏ qua
        }

        // ═══════════════════════════════════════════
        // CÁC PHƯƠNG THỨC THIẾU (pre-existing)
        // ═══════════════════════════════════════════

        private string FormatAmt(decimal v)
        {
            return v.ToString("#,##0");
        }

        private void DgvPO_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && dgvPO.Columns.Contains("In"))
                dgvPO.Cursor = Cursors.Hand;
        }

        private void DgvPO_CellMouseLeave(object sender, DataGridViewCellEventArgs e)
        {
            dgvPO.Cursor = Cursors.Default;
        }

        private void BtnPrintRequest_Click(object sender, EventArgs e)
        {
            if (dgvPO.CurrentRow == null) return;
            string poNo = dgvPO.CurrentRow.Cells["PO_No"]?.Value?.ToString() ?? "";
            if (string.IsNullOrWhiteSpace(poNo)) return;
            ShowPrintHistoryAndOptions(poNo);
        }

        private void ShowHistoryPaidPopup()
        {
            var popup = new Form
            {
                Text = "Lịch sử thanh toán",
                Size = new Size(800, 500),
                StartPosition = FormStartPosition.CenterParent,
                BackColor = Color.White
            };

            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Color.White
            };
            dgv.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn { Name = "HP_Date", HeaderText = "Ngày TT", Width = 100 },
                new DataGridViewTextBoxColumn { Name = "HP_PONo", HeaderText = "PO No.", Width = 120 },
                new DataGridViewTextBoxColumn { Name = "HP_Amount", HeaderText = "Số tiền", Width = 120 },
                new DataGridViewTextBoxColumn { Name = "HP_Method", HeaderText = "Phương thức", Width = 100 },
                new DataGridViewTextBoxColumn { Name = "HP_Note", HeaderText = "Ghi chú", Width = 200 }
            });

            // Load history data
            try
            {
                using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                conn.Open();
                using var cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                    SELECT TOP 200 Payment_Date, PONo, Amount, Payment_Method, Note
                    FROM Payment_History ORDER BY Payment_Date DESC", conn);
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    dgv.Rows.Add(
                        reader["Payment_Date"] != DBNull.Value ? Convert.ToDateTime(reader["Payment_Date"]).ToString("dd/MM/yyyy") : "",
                        reader["PONo"]?.ToString() ?? "",
                        reader["Amount"] != DBNull.Value ? FormatAmt(Convert.ToDecimal(reader["Amount"])) : "0",
                        reader["Payment_Method"]?.ToString() ?? "",
                        reader["Note"]?.ToString() ?? "");
                }
            }
            catch { }

            popup.Controls.Add(dgv);
            popup.ShowDialog(this);
        }

        private void DgvHistory_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var row = dgvHistory.Rows[e.RowIndex];
            string poNo = row.Cells["H_PONo"]?.Value?.ToString() ?? "";
            if (!string.IsNullOrWhiteSpace(poNo))
                ShowPrintHistoryAndOptions(poNo);
        }

        private bool VerifyAdminPassword()
        {
            var input = Microsoft.VisualBasic.Interaction.InputBox("Nhập mật khẩu quản trị:", "Xác nhận Admin", "");
            if (string.IsNullOrWhiteSpace(input)) return false;
            // Simple check - in production, use proper auth
            return input == "admin123" || AppSession.IsAdmin;
        }
    }
}

public class frmPaymentRequestPreview : Form
{
    private readonly POPaymentSummary _po;
    private readonly string _mprNo;
    private readonly List<PODetail> _details;
    private readonly Supplier _supp;
    private readonly List<PaymentSchedule> _schedules;

    private DateTimePicker dtpDate;
    private TextBox txtBenef, txtBankAcc, txtBankName;
    private ComboBox cboDot;          // Chọn đợt thanh toán
    private RichTextBox rtbPreview;
    private Form TopOwner => (this.TopLevelControl as Form) ?? this;

    public frmPaymentRequestPreview(POPaymentSummary po, string mprNo,
        List<PODetail> details, Supplier supp,
        List<PaymentSchedule> schedules = null)
    {
        _po = po;
        _mprNo = mprNo;
        _details = details;
        _supp = supp ?? new Supplier();
        _schedules = schedules ?? new List<PaymentSchedule>();
        BuildUI();
        GeneratePreview();
    }

    private string GetPropValue(object obj, params string[] propNames)
    {
        if (obj == null) return "";
        var type = obj.GetType();
        foreach (var name in propNames)
        {
            var prop = type.GetProperty(name);
            if (prop != null)
            {
                return prop.GetValue(obj, null)?.ToString() ?? "";
            }
        }
        return "";
    }

    private void BuildUI()
    {
        this.Text = "📄 Trích xuất Payment Request";
        this.Size = new Size(1100, 700);
        this.StartPosition = FormStartPosition.CenterParent;
        this.BackColor = Color.White;

        var pLeft = new Panel { Location = new Point(10, 10), Size = new Size(300, 630), BorderStyle = BorderStyle.FixedSingle, Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left };
        this.Controls.Add(pLeft);

        var lbl1 = new Label { Text = "THÔNG TIN THANH TOÁN", Location = new Point(10, 10), Size = new Size(280, 20), Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212) };
        pLeft.Controls.Add(lbl1);

        DateTime createdDate = _po.PO_Date ?? DateTime.Today;
        int y = 40;
        pLeft.Controls.Add(new Label { Text = "Ngày dự kiến TT (+7):", Location = new Point(10, y), Size = new Size(280, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold) });
        dtpDate = new DateTimePicker { Location = new Point(10, y + 22), Size = new Size(270, 25), Font = new Font("Segoe UI", 9), Format = DateTimePickerFormat.Short, Value = createdDate.AddDays(7) };
        pLeft.Controls.Add(dtpDate);

        // ── Chọn đợt thanh toán → lấy Amount_Plan ──
        y += 60;
        pLeft.Controls.Add(new Label { Text = "Đợt thanh toán (Final amount):", Location = new Point(10, y), Size = new Size(280, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold) });
        cboDot = new ComboBox { Location = new Point(10, y + 22), Size = new Size(270, 25), Font = new Font("Segoe UI", 9), DropDownStyle = ComboBoxStyle.DropDownList };
        cboDot.Items.Add("— Tính từ chi tiết PO (tổng VAT) —");
        foreach (var s in _schedules)
            cboDot.Items.Add($"Đợt {s.Dot_TT}: {FormatAmt(s.Amount_Plan)} VNĐ  [{s.Status}]");
        cboDot.SelectedIndex = _schedules.Count > 0 ? 1 : 0;
        cboDot.SelectedIndexChanged += (s, ev) => GeneratePreview();
        pLeft.Controls.Add(cboDot);

        string fullName = GetPropValue(_supp, "Company_Name", "CompanyName", "FullName");
        if (string.IsNullOrEmpty(fullName)) fullName = _po.Supplier_Name;

        y += 60;
        pLeft.Controls.Add(new Label { Text = "Người thụ hưởng (Beneficiary):", Location = new Point(10, y), Size = new Size(280, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold) });
        txtBenef = new TextBox { Location = new Point(10, y + 22), Size = new Size(270, 25), Font = new Font("Segoe UI", 9), Text = fullName };
        pLeft.Controls.Add(txtBenef);

        string bankAcc = GetPropValue(_supp, "Bank_Account", "BankAccount", "Account_No");
        string bankName = GetPropValue(_supp, "Bank_Name", "BankName", "Bank");

        y += 60;
        pLeft.Controls.Add(new Label { Text = "Số tài khoản (Bank Account):", Location = new Point(10, y), Size = new Size(280, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold) });
        txtBankAcc = new TextBox { Location = new Point(10, y + 22), Size = new Size(270, 25), Font = new Font("Segoe UI", 9), Text = bankAcc };
        pLeft.Controls.Add(txtBankAcc);

        y += 60;
        pLeft.Controls.Add(new Label { Text = "Ngân hàng (Bank Name):", Location = new Point(10, y), Size = new Size(280, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold) });
        txtBankName = new TextBox { Location = new Point(10, y + 22), Size = new Size(270, 25), Font = new Font("Segoe UI", 9), Text = bankName };
        pLeft.Controls.Add(txtBankName);

        y += 60;
        var btnUpdate = new Button { Text = "🔄 Cập nhật văn bản", Location = new Point(10, y), Size = new Size(270, 35), BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Cursor = Cursors.Hand };
        btnUpdate.FlatAppearance.BorderSize = 0;
        btnUpdate.Click += (s, e) => GeneratePreview();
        pLeft.Controls.Add(btnUpdate);

        var lblNote = new Label { Text = "Lưu ý: Màn hình này hiển thị dạng Tab (khoảng trắng) để bạn dễ xem và sửa nội dung. Khi bấm Copy, code sẽ tự bọc Bảng HTML kẻ ô để dán ra Word/Excel cực chuẩn.", Location = new Point(10, y + 50), Size = new Size(270, 100), Font = new Font("Segoe UI", 8, FontStyle.Italic), ForeColor = Color.Gray };
        pLeft.Controls.Add(lblNote);

        var pRight = new Panel { Location = new Point(320, 10), Size = new Size(750, 630), BorderStyle = BorderStyle.FixedSingle, Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right };
        this.Controls.Add(pRight);

        var lbl2 = new Label { Text = "NỘI DUNG VĂN BẢN (Có thể chỉnh sửa trực tiếp)", Location = new Point(10, 10), Size = new Size(400, 20), Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.FromArgb(40, 167, 69) };
        pRight.Controls.Add(lbl2);

        var btnCopy = new Button { Text = "📋 Copy sang Bảng tạm", Location = new Point(590, 5), Size = new Size(150, 30), BackColor = Color.FromArgb(40, 167, 69), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Cursor = Cursors.Hand, Anchor = AnchorStyles.Top | AnchorStyles.Right };
        btnCopy.FlatAppearance.BorderSize = 0;
        btnCopy.Click += BtnCopy_Click;
        pRight.Controls.Add(btnCopy);

        rtbPreview = new RichTextBox { Location = new Point(10, 40), Size = new Size(730, 580), Font = new Font("Times New Roman", 11), Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right };
        rtbPreview.WordWrap = false;
        pRight.Controls.Add(rtbPreview);
    }

    private void GeneratePreview()
    {
        var sb = new System.Text.StringBuilder();

        // ── Dòng 1: dùng Short_Name (Viết tắt) của NCC ──
        string suppShort = GetPropValue(_supp, "Short_Name", "ShortName", "Supplier_Name", "SupplierName");
        if (string.IsNullOrEmpty(suppShort)) suppShort = _po.Supplier_Name;

        sb.AppendLine($"1. Please transfer for Request payment for PO {_po.PONo} to {suppShort} of {_mprNo}");
        sb.AppendLine();
        sb.AppendLine("2. Description");
        sb.AppendLine();

        // Header bảng — 11 cột
        sb.AppendLine("STT\tTên hàng\tVật Liệu\tA(mm)\tB(mm)\tC(mm)\tSL\tĐVT\tKG\tĐơn giá\tThành tiền");

        decimal subTotal = 0, finalTotal = 0;
        decimal vatPct = 0;
        int stt = 1;
        foreach (var d in _details)
        {
            decimal q = d.Qty_Per_Sheet;
            decimal wk = d.Weight_kg;
            decimal p = d.Price;
            decimal v = d.VAT;
            if (v > vatPct) vatPct = v; // lấy VAT cao nhất để hiển thị

            string calcMethod = (d.Remarks ?? "").Contains("[CALC:KG]") ? "Theo KG" : "Theo SL";
            decimal baseVal = calcMethod == "Theo KG" ? wk : q;
            decimal realPrice = p;
            if (calcMethod == "Theo KG" && wk > 0 && q > 0) realPrice = (p * q) / wk;
            decimal amtBeforeVat = Math.Round(baseVal * realPrice, 0);
            decimal amtAfterVat = Math.Round(amtBeforeVat * (1 + v / 100), 0);
            subTotal += amtBeforeVat;
            finalTotal += amtAfterVat;

            // Làm sạch các field — thay \r\n, \n thành space để không vỡ bảng
            string itemName = (d.Item_Name ?? "").Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ").Trim();
            string material = (d.Material ?? "").Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ").Trim();

            sb.AppendLine($"{stt++}\t{itemName}\t{material}\t{d.Asize}\t{d.Bsize}\t{d.Csize}\t{q}\t{d.UNIT}\t{wk}\t{FormatAmt(realPrice)}\t{FormatAmt(amtAfterVat)}");
        }

        sb.AppendLine($"\t\t\t\t\t\t\t\t\tSUB-TOTAL\t{FormatAmt(subTotal)}");
        sb.AppendLine($"\t\t\t\t\t\t\t\t\tFinal Price Requested (Included {vatPct:N0}% VAT)\t{FormatAmt(finalTotal)}");
        sb.AppendLine();
        sb.AppendLine("3. Amount");
        sb.AppendLine();
        sb.AppendLine($"Total Amount:\t\t{FormatAmt(subTotal)} VNĐ (excluded VAT)");
        sb.AppendLine();

        // ── Final amount: luôn là số tiền SAU thuế, làm tròn không lấy số thập phân ──
        decimal finalAmt = Math.Round(finalTotal, 0); // mặc định = tổng sau VAT
        string dotLabel = "";
        if (cboDot != null && cboDot.SelectedIndex > 0)
        {
            var sched = _schedules[cboDot.SelectedIndex - 1];
            // Amount_Plan là số tiền kế hoạch — nhân VAT để ra số tiền sau thuế
            finalAmt = Math.Round(sched.Amount_Plan * (1 + vatPct / 100), 0);
            dotLabel = $"  (Đợt {sched.Dot_TT} — {sched.Percent_TT}%)";
        }

        // VAT amount = finalAmt - (finalAmt / (1 + vatPct/100))
        decimal baseBeforeVat = vatPct > 0 ? Math.Round(finalAmt / (1 + vatPct / 100), 0) : finalAmt;
        decimal vatAmount = finalAmt - baseBeforeVat;

        sb.AppendLine("4. Payment information");
        sb.AppendLine();
        sb.AppendLine($"Final amount :\t\t{FormatAmt(finalAmt)} VNĐ included {vatPct:N0}% VAT ({FormatAmt(vatAmount)} VNĐ){dotLabel}");
        sb.AppendLine($"Expect payment date:\t{dtpDate.Value:dd/MM/yyyy}");
        sb.AppendLine($"Name of beneficiary:\t{txtBenef.Text}");
        sb.AppendLine($"Bank account of beneficiary:\t{txtBankAcc.Text}");
        sb.AppendLine($"Bank name of beneficiary:\t{txtBankName.Text}");
        sb.AppendLine();
        sb.AppendLine("5. Remarks");

        rtbPreview.Text = sb.ToString();
    }

    private void BtnCopy_Click(object sender, EventArgs e)
    {
        if (string.IsNullOrEmpty(rtbPreview.Text)) return;

        var sbHtml = new StringBuilder();
        string[] lines = rtbPreview.Text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
        bool inTable = false;

        foreach (string line in lines)
        {
            // Làm sạch ký tự newline ẩn trong từng ô trước khi split
            string cleanLine = line.Replace("\r", " ").Replace("\n", " ");
            string[] cells = cleanLine.Split('\t');

            if (cells.Length >= 5)
            {
                if (!inTable)
                {
                    // Bảng KHÔNG dùng width:100% để giữ chiều rộng cột cố định
                    sbHtml.Append("<table border='1' cellspacing='0' cellpadding='5' style='" +
                        "border-collapse:collapse; font-family:\"Times New Roman\",serif; " +
                        "font-size:11pt; border:1px solid black; margin-bottom:10px; table-layout:fixed;'>");
                    // Cố định chiều rộng từng cột — cột Thành tiền (cột 11) giữ nguyên
                    sbHtml.Append("<colgroup>" +
                        "<col style='width:35px;'/>" +   // STT
                        "<col style='width:160px;'/>" +   // Tên hàng
                        "<col style='width:80px;'/>" +   // Vật liệu
                        "<col style='width:55px;'/>" +   // A(mm)
                        "<col style='width:55px;'/>" +   // B(mm)
                        "<col style='width:55px;'/>" +   // C(mm)
                        "<col style='width:40px;'/>" +   // SL
                        "<col style='width:40px;'/>" +   // ĐVT
                        "<col style='width:55px;'/>" +   // KG
                        "<col style='width:90px;'/>" +   // Đơn giá
                        "<col style='width:110px;'/>" +   // Thành tiền — CỐ ĐỊNH
                        "</colgroup>");
                    inTable = true;
                }
                sbHtml.Append("<tr>");
                bool isHeader = (cells[0].Trim() == "STT");

                if (line.Contains("SUB-TOTAL") || line.Contains("Final Price Requested"))
                {
                    string textLabel = cells.FirstOrDefault(c => c.Contains("SUB-TOTAL") || c.Contains("Final Price Requested"))?.Trim() ?? "";
                    string amountVal = cells.LastOrDefault()?.Trim() ?? "";
                    sbHtml.Append($"<td colspan='9' style='border:1px solid black; padding:5px; font-weight:bold; text-align:center;'>{textLabel}</td>");
                    sbHtml.Append("<td style='border:1px solid black;'></td>");
                    sbHtml.Append($"<td style='border:1px solid black; padding:5px; font-weight:bold; text-align:right;'>{amountVal}</td>");
                }
                else
                {
                    foreach (string cell in cells)
                    {
                        string cellVal = cell.Trim();
                        if (isHeader)
                        {
                            sbHtml.Append($"<th style='background-color:#d9d9d9; border:1px solid black; padding:5px; text-align:center; overflow:hidden;'>{cellVal}</th>");
                        }
                        else
                        {
                            bool isNumber = decimal.TryParse(cellVal.Replace(",", ""), out _) && cellVal.Length > 0;
                            bool isSTT = cellVal.Length <= 3 && cellVal.All(char.IsDigit) && cellVal.Length > 0;
                            string align = isSTT ? "center" : isNumber ? "right" : "left";
                            sbHtml.Append($"<td style='border:1px solid black; padding:5px; text-align:{align}; overflow:hidden; word-break:break-word;'>{cellVal}</td>");
                        }
                    }
                }
                sbHtml.Append("</tr>");
            }
            else
            {
                if (inTable) { sbHtml.Append("</table><br/>"); inTable = false; }
                string normalLine = cleanLine.Replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;");
                if (string.IsNullOrWhiteSpace(normalLine))
                    sbHtml.Append("<br/>");
                else
                {
                    bool isSection = normalLine.TrimStart().StartsWith("1.") || normalLine.TrimStart().StartsWith("2.") ||
                                     normalLine.TrimStart().StartsWith("3.") || normalLine.TrimStart().StartsWith("4.") ||
                                     normalLine.TrimStart().StartsWith("5.");
                    if (isSection)
                        sbHtml.Append($"<div style='margin-top:10px; margin-bottom:5px;'><b>{normalLine}</b></div>");
                    else
                        sbHtml.Append($"<div style='margin-bottom:5px;'>{normalLine}</div>");
                }
            }
        }
        if (inTable) sbHtml.Append("</table>");

        CopyToClipboardAsHtml(sbHtml.ToString(), rtbPreview.Text);
        MessageBox.Show(TopOwner, "✅ Đã copy nội dung vào Bảng tạm!\nDán (Ctrl+V) vào Word hoặc Outlook sẽ hiển thị bảng kẻ ô chuẩn, font Times New Roman.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
        this.Close();
    }

    // =====================================================================
    // THUẬT TOÁN ĐẨY HTML LÊN CLIPBOARD BẰNG BYTE OFFSET (CHỐNG LỖI UTF-8)
    // =====================================================================
    // HELPER: Format số tiền — luôn làm tròn, không lấy số thập phân (VNĐ)
    private static string FormatAmt(decimal value)
    {
        return Math.Round(value, 0).ToString("N0");
    }

    private void CopyToClipboardAsHtml(string htmlFragment, string plainText)
    {
        string startHtml = "<html><body style=\"font-family:'Times New Roman', serif; font-size:11pt;\">\r\n\r\n";
        string endHtml = "\r\n\r\n</body></html>";
        string htmlContext = startHtml + htmlFragment + endHtml;

        string headerTemplate =
            "Version:0.9\r\n" +
            "StartHTML:{0:D8}\r\n" +
            "EndHTML:{1:D8}\r\n" +
            "StartFragment:{2:D8}\r\n" +
            "EndFragment:{3:D8}\r\n";

        int headerLength = Encoding.UTF8.GetByteCount(string.Format(headerTemplate, 0, 0, 0, 0));
        int htmlContextLength = Encoding.UTF8.GetByteCount(htmlContext);

        int startHtmlOffset = headerLength;
        int startFragmentOffset = headerLength + Encoding.UTF8.GetByteCount(startHtml);
        int endFragmentOffset = startFragmentOffset + Encoding.UTF8.GetByteCount(htmlFragment);
        int endHtmlOffset = headerLength + htmlContextLength;

        string header = string.Format(headerTemplate, startHtmlOffset, endHtmlOffset, startFragmentOffset, endFragmentOffset);
        string cfHtml = header + htmlContext;

        DataObject obj = new DataObject();
        obj.SetData(DataFormats.Html, cfHtml);
        obj.SetData(DataFormats.UnicodeText, plainText);

        for (int attempt = 0; attempt < 5; attempt++)
        {
            try
            {
                Clipboard.SetDataObject(obj, true);
                return;
            }
            catch (System.Runtime.InteropServices.ExternalException)
            {
                if (attempt == 4) throw;
                System.Threading.Thread.Sleep(50);
            }
        }
    }
}
