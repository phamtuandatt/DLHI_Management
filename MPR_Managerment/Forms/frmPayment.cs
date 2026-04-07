// ============================================================
//  FILE: Forms/frmPayment.cs
//  Tab 1: Tiến độ thanh toán từng PO
//  Tab 2: Báo cáo tổng hợp công nợ NCC theo kỳ
// ============================================================
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MPR_Managerment.Models;
using MPR_Managerment.Services;
using OfficeOpenXml;

namespace MPR_Managerment.Forms
{
    public partial class frmPayment : Form
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

        // Controls chính
        private TabControl tabs;
        private TabPage tabPO, tabDebt;

        // Tab PO
        private TextBox txtSearchPO;
        private ComboBox cboStatusFilter;
        private DataGridView dgvPO, dgvSchedule, dgvHistory;
        private Panel panelDoc;        // Bảng Document (Invoice + Delivery)
        private DataGridView dgvDoc;
        private Label lblPOName, lblPOAmount, lblPOPaid, lblPORemain, lblPOStatus, lblPOProgress;
        private Panel panelTop, panelInfo, panelSched, panelHist;
        private Panel panelPrintHistory;   // Danh sách PO đã in Request
        private DataGridView dgvPrintHistory;
        private DateTimePicker _phDateFrom, _phDateTo; // Bộ lọc thời gian
        private ProgressBar progressPO;

        // Tab Debt
        private DateTimePicker dtpFrom, dtpTo;
        private ComboBox cboSuppFilter;
        private DataGridView dgvDebtSupp, dgvDebtDetail;
        private Label lblSumValue, lblSumPaid, lblSumDebt, lblSumOverdue;
        private Button btnExportDebt;

        private Button btnRefreshPO;

        // =====================================================================
        public frmPayment()
        {
            InitializeComponent();
            BuildUI();
            LoadData();
            this.Resize += (s, e) => ResizeAll();
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
            txtSearchPO.PlaceholderText = "PO No / Dự án / NCC...";
            txtSearchPO.TextChanged += (s, e) => FilterAndBind();

            Lbl(pFilter, "Trạng thái:", 278, 12, 85, 20);
            cboStatusFilter = Cbo(pFilter, 363, 8, 180,
                new[] { "Tất cả", "Chưa TT", "Một phần", "Đã TT đủ", "⚠ Quá hạn" });
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
                var bAdd = Btn("+ Thêm đợt", Color.FromArgb(40, 167, 69), 8, 28, 100, 26);
                var bDel = Btn("Xóa", Color.FromArgb(220, 53, 69), 114, 28, 65, 26);
                var bSave = Btn("💾 Lưu", Color.FromArgb(0, 120, 212), 185, 28, 80, 26);
                var bReq = Btn("📄 Request to EC", Color.FromArgb(102, 51, 153), 271, 28, 150, 26);
                var bPrint = Btn("🖨 In Request", Color.FromArgb(0, 150, 100), 427, 28, 120, 26);

                bAdd.Click += BtnAddSched_Click;
                bDel.Click += BtnDelSched_Click;
                bSave.Click += BtnSaveSched_Click;
                bReq.Click += BtnPaymentRequest_Click;
                bPrint.Click += BtnPrintRequest_Click;

                panelSched.Controls.AddRange(new Control[] { bAdd, bDel, bSave, bReq, bPrint });
            }

            dgvSchedule = Grid(panelSched, 60, 0);
            dgvSchedule.SelectionChanged += (s, e) =>
            {
                if (dgvSchedule.SelectedRows.Count > 0)
                    _selectedSchedID = Convert.ToInt32(dgvSchedule.SelectedRows[0].Cells["S_ID"].Value ?? 0);
            };
            dgvSchedule.CellFormatting += DgvSched_CellFormatting;
            dgvSchedule.CellEndEdit += DgvSchedule_CellEndEdit;
            BuildSchedCols();

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
                Value = DateTime.Today.AddMonths(-3)
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

            var btnPhReset = Btn("✖ Reset", Color.FromArgb(108, 117, 125), 384, 26, 70, 26);
            btnPhReset.Click += (s, ev) =>
            {
                _phDateFrom.Value = DateTime.Today.AddMonths(-3);
                _phDateTo.Value = DateTime.Today;
                LoadPrintHistory(_phDateFrom.Value.Date, _phDateTo.Value.Date.AddDays(1).AddSeconds(-1));
            };
            panelPrintHistory.Controls.Add(btnPhReset);

            var btnPhDel = Btn("🗑 Xóa dòng", Color.FromArgb(220, 53, 69), 462, 26, 100, 26);
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
            Lbl(panelHist, "💰  LỊCH SỬ THANH TOÁN THỰC TẾ", 8, 5, 350, 20, true, Color.FromArgb(40, 167, 69));

            if (canEdit)
            {
                var bPay = Btn("+ Ghi nhận TT", Color.FromArgb(40, 167, 69), 8, 28, 125, 26);
                var bDel = Btn("Xóa", Color.FromArgb(220, 53, 69), 139, 28, 65, 26);
                bPay.Click += BtnAddPayment_Click;
                bDel.Click += BtnDelPayment_Click;
                panelHist.Controls.AddRange(new Control[] { bPay, bDel });
            }

            dgvHistory = Grid(panelHist, 58, 137);   // 200 - 58 - 5 = 137
            dgvHistory.ReadOnly = true;
            dgvHistory.SelectionChanged += (s, e) =>
            {
                if (dgvHistory.SelectedRows.Count > 0)
                    _selectedHistID = Convert.ToInt32(dgvHistory.SelectedRows[0].Cells["H_ID"].Value ?? 0);
                else if (dgvHistory.CurrentRow != null)
                    _selectedHistID = Convert.ToInt32(dgvHistory.CurrentRow.Cells["H_ID"].Value ?? 0);
            };
            BuildHistCols();

            // ── PANEL: Document (Invoice + Delivery Note) ──
            panelDoc = P(tabPO, 0, 317 + 200 + 5, 0, 0, Color.White);
            panelDoc.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            Lbl(panelDoc, "📎  DOCUMENT", 8, 5, 200, 20, true, Color.FromArgb(0, 120, 212));

            var bPrintDoc = Btn("🖨 In tài liệu", Color.FromArgb(102, 51, 153), 220, 3, 130, 26);
            bPrintDoc.Click += BtnPrintDocs_Click;
            panelDoc.Controls.Add(bPrintDoc);

            var bRefreshDoc = Btn("🔄 Tải lại", Color.FromArgb(0, 150, 150), 360, 3, 90, 26);
            bRefreshDoc.Click += (s, e) => LoadDocuments();
            panelDoc.Controls.Add(bRefreshDoc);

            dgvDoc = Grid(panelDoc, 32, 0);
            dgvDoc.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvDoc.MultiSelect = true;
            dgvDoc.ReadOnly = true;
            dgvDoc.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Doc_Chk", HeaderText = "", Width = 30 });
            dgvDoc.Columns.Add(new DataGridViewTextBoxColumn { Name = "Doc_Type", HeaderText = "Loại", Width = 100, ReadOnly = true });
            dgvDoc.Columns.Add(new DataGridViewTextBoxColumn { Name = "Doc_Name", HeaderText = "Tên file", Width = 280, ReadOnly = true });
            dgvDoc.Columns.Add(new DataGridViewTextBoxColumn { Name = "Doc_Path", HeaderText = "Đường dẫn", FillWeight = 100, ReadOnly = true });
            dgvDoc.Columns["Doc_Path"].Visible = false; // ẩn đường dẫn đầy đủ
            dgvDoc.CellFormatting += (s, ev) =>
            {
                if (ev.RowIndex < 0) return;
                string col = dgvDoc.Columns[ev.ColumnIndex].Name;
                string type = dgvDoc.Rows[ev.RowIndex].Cells["Doc_Type"].Value?.ToString() ?? "";
                if (col == "Doc_Type")
                {
                    ev.CellStyle.ForeColor = type == "Invoice"
                        ? Color.FromArgb(0, 120, 212)
                        : Color.FromArgb(40, 167, 69);
                    ev.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                }
            };
            // Tick checkbox khi click dòng
            dgvDoc.CellClick += (s, ev) =>
            {
                if (ev.RowIndex < 0 || ev.ColumnIndex != dgvDoc.Columns["Doc_Chk"].Index) return;
                var cell = dgvDoc.Rows[ev.RowIndex].Cells["Doc_Chk"] as DataGridViewCheckBoxCell;
                if (cell != null) cell.Value = !(bool)(cell.Value ?? false);
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
            dgvSchedule.Columns.Add(new DataGridViewTextBoxColumn { Name = "Delivery_Ref", HeaderText = "Lô hàng", Width = 90 });
            dgvSchedule.Columns.Add(new DataGridViewTextBoxColumn { Name = "Description", HeaderText = "Điều kiện", FillWeight = 100 });
            var cboStatus = new DataGridViewComboBoxColumn { Name = "S_Status", HeaderText = "Trạng thái", Width = 100, FlatStyle = FlatStyle.Flat };
            cboStatus.Items.AddRange(new[] { "Chưa TT", "Một phần", "Đã TT đủ" });
            dgvSchedule.Columns.Add(cboStatus);

            // ── DateTimePicker ẩn — hiện khi click vào ô Due_Date ──
            var dtp = new DateTimePicker
            {
                Format = DateTimePickerFormat.Short,
                Font = new Font("Segoe UI", 9),
                Visible = false,
                MinDate = new DateTime(2000, 1, 1)
            };

            // Thêm DTP vào panel cha của dgvSchedule
            panelSched.Controls.Add(dtp);
            dtp.BringToFront();

            int _dtpRow = -1, _dtpCol = -1;

            // Hiện DTP khi click vào cột Due_Date
            dgvSchedule.CellClick += (s, ev) =>
            {
                if (ev.RowIndex < 0 || dgvSchedule.Columns[ev.ColumnIndex].Name != "Due_Date") return;

                _dtpRow = ev.RowIndex;
                _dtpCol = ev.ColumnIndex;

                // Parse giá trị hiện tại
                string cur = dgvSchedule.Rows[ev.RowIndex].Cells["Due_Date"].Value?.ToString() ?? "";
                dtp.Value = DateTime.TryParseExact(cur, "dd/MM/yyyy",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out DateTime parsed)
                    ? parsed : DateTime.Today;

                // Tính tọa độ của cell trong panelSched
                var cellRect = dgvSchedule.GetCellDisplayRectangle(ev.ColumnIndex, ev.RowIndex, true);
                var cellPos = dgvSchedule.PointToScreen(new Point(cellRect.Left, cellRect.Top));
                var panelPos = panelSched.PointToClient(cellPos);

                dtp.Location = new Point(panelPos.X, panelPos.Y);
                dtp.Width = cellRect.Width;
                dtp.Height = cellRect.Height;
                dtp.Visible = true;
                dtp.Focus();
            };

            // Ẩn DTP khi click ra ngoài
            dgvSchedule.CellClick += (s, ev) =>
            {
                if (ev.ColumnIndex >= 0 && dgvSchedule.Columns[ev.ColumnIndex].Name != "Due_Date")
                    dtp.Visible = false;
            };
            dgvSchedule.Scroll += (s, ev) => dtp.Visible = false;

            // Khi chọn ngày → ghi vào cell và ẩn DTP
            dtp.ValueChanged += (s, ev) =>
            {
                if (_dtpRow < 0 || !dtp.Visible) return;
                dgvSchedule.Rows[_dtpRow].Cells["Due_Date"].Value = dtp.Value.ToString("dd/MM/yyyy");
            };

            dtp.Leave += (s, ev) => dtp.Visible = false;

            dtp.KeyDown += (s, ev) =>
            {
                if (ev.KeyCode == Keys.Escape || ev.KeyCode == Keys.Enter)
                {
                    if (ev.KeyCode == Keys.Enter && _dtpRow >= 0)
                        dgvSchedule.Rows[_dtpRow].Cells["Due_Date"].Value = dtp.Value.ToString("dd/MM/yyyy");
                    dtp.Visible = false;
                    dgvSchedule.Focus();
                }
            };
        }

        private void BuildHistCols()
        {
            dgvHistory.Columns.Clear();
            dgvHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "H_ID", Visible = false });
            dgvHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "Pay_Date", HeaderText = "Ngày TT", Width = 90, ReadOnly = true });
            dgvHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "Amount_Paid", HeaderText = "Số tiền", Width = 110, ReadOnly = true });
            dgvHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "Pay_Method", HeaderText = "Hình thức", Width = 100, ReadOnly = true });
            dgvHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "Bank_Name", HeaderText = "Ngân hàng", Width = 100, ReadOnly = true });
            dgvHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "Transaction_No", HeaderText = "Số CT", Width = 100, ReadOnly = true });
            dgvHistory.Columns.Add(new DataGridViewTextBoxColumn { Name = "Notes", HeaderText = "Ghi chú", FillWeight = 100, ReadOnly = true });
        }

        private void BuildTabDebt()
        {
            var pF = P(tabDebt, 5, 5, 0, 45, Color.White);
            pF.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            Lbl(pF, "Từ ngày:", 6, 13, 65, 20);
            dtpFrom = new DateTimePicker
            {
                Location = new Point(71, 9),
                Size = new Size(125, 26),
                Font = new Font("Segoe UI", 9),
                Format = DateTimePickerFormat.Short,
                Value = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1)
            };
            pF.Controls.Add(dtpFrom);

            Lbl(pF, "Đến ngày:", 205, 13, 70, 20);
            dtpTo = new DateTimePicker
            {
                Location = new Point(275, 9),
                Size = new Size(125, 26),
                Font = new Font("Segoe UI", 9),
                Format = DateTimePickerFormat.Short,
                Value = DateTime.Today
            };
            pF.Controls.Add(dtpTo);

            Lbl(pF, "Nhà cung cấp:", 410, 13, 100, 20);
            cboSuppFilter = new ComboBox
            {
                Location = new Point(510, 9),
                Size = new Size(220, 26),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboSuppFilter.Items.Add("Tất cả nhà cung cấp");
            cboSuppFilter.SelectedIndex = 0;
            pF.Controls.Add(cboSuppFilter);

            var bView = Btn("🔍 Xem báo cáo", Color.FromArgb(0, 120, 212), 745, 8, 145, 30);
            bView.Click += BtnViewDebt_Click;
            pF.Controls.Add(bView);

            btnExportDebt = Btn("📥 Xuất Excel", Color.FromArgb(0, 150, 100), 900, 8, 125, 30);
            btnExportDebt.Click += BtnExportDebt_Click;
            pF.Controls.Add(btnExportDebt);

            var pCards = P(tabDebt, 5, 55, 0, 72, Color.White);
            pCards.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            lblSumValue = Card(pCards, 10, "Tổng giá trị PO", Color.FromArgb(0, 120, 212));
            lblSumPaid = Card(pCards, 225, "Đã thanh toán", Color.FromArgb(40, 167, 69));
            lblSumDebt = Card(pCards, 440, "Còn nợ", Color.FromArgb(255, 140, 0));
            lblSumOverdue = Card(pCards, 655, "Quá hạn (PO)", Color.FromArgb(220, 53, 69));

            var pNCC = P(tabDebt, 5, 132, 380, 0, Color.White);
            pNCC.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom;
            Lbl(pNCC, "TỔNG HỢP THEO NHÀ CUNG CẤP", 8, 5, 360, 20, true, Color.FromArgb(0, 120, 212));
            dgvDebtSupp = Grid(pNCC, 28, 0);
            dgvDebtSupp.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom | AnchorStyles.Right;
            dgvDebtSupp.SelectionChanged += DgvDebtSupp_SelectionChanged;
            dgvDebtSupp.CellFormatting += DgvDebtSupp_CellFormatting;
            BuildDebtSuppCols();

            var pDet = P(tabDebt, 390, 132, 0, 0, Color.White);
            pDet.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            Lbl(pDet, "CHI TIẾT TỪNG ĐƠN PO", 8, 5, 400, 20, true, Color.FromArgb(0, 120, 212));
            dgvDebtDetail = Grid(pDet, 28, 0);
            dgvDebtDetail.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            dgvDebtDetail.CellFormatting += DgvDebtDetail_CellFormatting;
            BuildDebtDetailCols();
        }

        private void BuildDebtSuppCols()
        {
            dgvDebtSupp.Columns.Clear();
            dgvDebtSupp.Columns.Add(new DataGridViewTextBoxColumn { Name = "D_SuppID", Visible = false });
            dgvDebtSupp.Columns.Add(new DataGridViewTextBoxColumn { Name = "D_Name", HeaderText = "Nhà cung cấp", FillWeight = 100, ReadOnly = true });
            dgvDebtSupp.Columns.Add(new DataGridViewTextBoxColumn { Name = "D_TotalPO", HeaderText = "Số PO", Width = 55, ReadOnly = true });
            dgvDebtSupp.Columns.Add(new DataGridViewTextBoxColumn { Name = "D_Value", HeaderText = "Tổng PO", Width = 105, ReadOnly = true });
            dgvDebtSupp.Columns.Add(new DataGridViewTextBoxColumn { Name = "D_Debt", HeaderText = "Còn nợ", Width = 105, ReadOnly = true });
            dgvDebtSupp.Columns.Add(new DataGridViewTextBoxColumn { Name = "D_Overdue", HeaderText = "Quá hạn", Width = 65, ReadOnly = true });
        }

        private void BuildDebtDetailCols()
        {
            dgvDebtDetail.Columns.Clear();
            dgvDebtDetail.Columns.Add(new DataGridViewTextBoxColumn { Name = "DD_PONo", HeaderText = "PO No", Width = 110, ReadOnly = true });
            dgvDebtDetail.Columns.Add(new DataGridViewTextBoxColumn { Name = "DD_Project", HeaderText = "Dự án", FillWeight = 100, ReadOnly = true });
            dgvDebtDetail.Columns.Add(new DataGridViewTextBoxColumn { Name = "DD_PODate", HeaderText = "Ngày PO", Width = 85, ReadOnly = true });
            dgvDebtDetail.Columns.Add(new DataGridViewTextBoxColumn { Name = "DD_Total", HeaderText = "Giá trị PO", Width = 100, ReadOnly = true });
            dgvDebtDetail.Columns.Add(new DataGridViewTextBoxColumn { Name = "DD_Before", HeaderText = "TT trước kỳ", Width = 100, ReadOnly = true });
            dgvDebtDetail.Columns.Add(new DataGridViewTextBoxColumn { Name = "DD_InRange", HeaderText = "TT trong kỳ", Width = 100, ReadOnly = true });
            dgvDebtDetail.Columns.Add(new DataGridViewTextBoxColumn { Name = "DD_Remain", HeaderText = "Còn nợ", Width = 100, ReadOnly = true });
            dgvDebtDetail.Columns.Add(new DataGridViewTextBoxColumn { Name = "DD_Status", HeaderText = "Trạng thái", Width = 95, ReadOnly = true });
            dgvDebtDetail.Columns.Add(new DataGridViewTextBoxColumn { Name = "DD_Due", HeaderText = "Đến hạn", Width = 85, ReadOnly = true });
        }

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
            LoadPrintHistory(DateTime.Today.AddMonths(-3), DateTime.Today.AddDays(1).AddSeconds(-1));
        }

        private async void LoadPOSummary()
        {
            btnRefreshPO.Enabled = false;
            btnRefreshPO.Text = "⏳ Đang tải...";
            try
            {
                var result = await System.Threading.Tasks.Task.Run(() =>
                {
                    var summaries = _svc.GetPOSummaries();
                    var allScheds = _svc.GetAllSchedules();
                    var cache = allScheds
                        .GroupBy(s => s.PO_ID)
                        .ToDictionary(g => g.Key, g => g.ToList());
                    return (summaries, cache);
                });
                _poSummaries = result.summaries;
                _allSchedulesCache = result.cache;
                FilterAndBind();
            }
            catch (Exception ex) { Err(ex.Message); }
            finally
            {
                btnRefreshPO.Enabled = true;
                btnRefreshPO.Text = "🔄 Làm mới";
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
                decimal totalPaid = p.Total_Paid;
                decimal remain = totalPO - totalPaid;
                if (remain < 0) remain = 0;

                decimal pct = totalPO > 0 ? (totalPaid / totalPO) * 100 : 0;
                if (pct > 100) pct = 100;

                string realStatus = "Chưa TT";
                if (totalPaid >= totalPO && totalPO > 0) realStatus = "Đã TT đủ";
                else if (totalPaid > 0) realStatus = "Một phần";

                bool isNew = p.PO_Date.HasValue && (DateTime.Now - p.PO_Date.Value).TotalDays <= 3;
                string poDisplayObj = isNew ? $"🔥 {p.PONo} (Mới)" : p.PONo;

                // ── Schedules từng đợt từ cache ──
                var scheds = _allSchedulesCache.ContainsKey(p.PO_ID)
                    ? _allSchedulesCache[p.PO_ID]
                    : new List<PaymentSchedule>();
                string d1a = "", d1s = "", d2a = "", d2s = "", d3a = "", d3s = "", d4a = "", d4s = "", d5a = "", d5s = "";
                for (int idx = 0; idx < scheds.Count && idx < 5; idx++)
                {
                    string a = scheds[idx].Amount_Plan.ToString("N0");
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
                    Tong_PO = totalPO.ToString("N0"),
                    Da_TT = totalPaid.ToString("N0"),
                    Con_No = remain.ToString("N0"),
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

            dgvPO.DataSource = displayList;
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
                    r.Cells["Amount_Plan"].Value = s.Amount_Plan.ToString("N0");
                    r.Cells["Due_Date"].Value = s.Due_Date.HasValue ? s.Due_Date.Value.ToString("dd/MM/yyyy") : "";
                    r.Cells["Delivery_Ref"].Value = s.Delivery_Ref;
                    r.Cells["Description"].Value = s.Description;
                    r.Cells["S_Status"].Value = s.Status;
                }

                _histories = _svc.GetHistories(_selectedPO_ID);
                dgvHistory.Rows.Clear();
                foreach (var h in _histories)
                {
                    int i = dgvHistory.Rows.Add();
                    var r = dgvHistory.Rows[i];
                    r.Cells["H_ID"].Value = h.Payment_ID;
                    r.Cells["Pay_Date"].Value = h.Payment_Date.ToString("dd/MM/yyyy");
                    r.Cells["Amount_Paid"].Value = h.Amount_Paid.ToString("N0");
                    r.Cells["Pay_Method"].Value = h.Payment_Method;
                    r.Cells["Bank_Name"].Value = h.Bank_Name;
                    r.Cells["Transaction_No"].Value = h.Transaction_No;
                    r.Cells["Notes"].Value = h.Notes;
                }
            }
            catch (Exception ex) { Err(ex.Message); }
        }

        // ── Load Invoice + Delivery Note cho PO đang chọn ──
        private void LoadDocuments()
        {
            if (dgvDoc == null) return;
            dgvDoc.Rows.Clear();
            if (_selectedPO_ID == 0) return;

            var po = _poSummaries.Find(x => x.PO_ID == _selectedPO_ID);
            if (po == null) return;

            // Lấy ProjectInfo để có INV_Link và DeliveryNote_Link
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
            string invFolder = proj?.INV_Link ?? "";
            string delFolder = proj?.DeliveryNote_Link ?? "";

            // ── Tìm file Invoice: INV_<PONo>* trong INV_Link ──
            ScanFolder(invFolder, $"INV_{poNo}", "Invoice");
            // ── Tìm file Delivery Note: Delivery_<PONo>* trong DeliveryNote_Link ──
            ScanFolder(delFolder, $"Delivery_{poNo}", "Delivery Note");
        }

        private void ScanFolder(string folder, string prefix, string docType)
        {
            if (string.IsNullOrWhiteSpace(folder)) return;
            if (!System.IO.Directory.Exists(folder)) return;
            try
            {
                var files = System.IO.Directory.GetFiles(folder, $"{prefix}*",
                    System.IO.SearchOption.TopDirectoryOnly);
                foreach (var f in files)
                {
                    int i = dgvDoc.Rows.Add();
                    dgvDoc.Rows[i].Cells["Doc_Chk"].Value = false;
                    dgvDoc.Rows[i].Cells["Doc_Type"].Value = docType;
                    dgvDoc.Rows[i].Cells["Doc_Name"].Value = System.IO.Path.GetFileName(f);
                    dgvDoc.Rows[i].Cells["Doc_Path"].Value = f;
                }
            }
            catch { }
        }

        private void BtnPrintDocs_Click(object sender, EventArgs e)
        {
            if (dgvDoc == null || dgvDoc.Rows.Count == 0)
            {
                Warn("Không có tài liệu nào trong bảng để in!");
                return;
            }

            // Thu thập các file được tick — nếu không có file nào được tick thì in tất cả
            var filesToPrint = new List<string>();
            foreach (DataGridViewRow row in dgvDoc.Rows)
            {
                bool chk = row.Cells["Doc_Chk"].Value is bool b && b;
                if (chk)
                    filesToPrint.Add(row.Cells["Doc_Path"].Value?.ToString() ?? "");
            }
            if (filesToPrint.Count == 0)
            {
                // Mặc định: in tất cả
                foreach (DataGridViewRow row in dgvDoc.Rows)
                    filesToPrint.Add(row.Cells["Doc_Path"].Value?.ToString() ?? "");
            }

            filesToPrint = filesToPrint.Where(f => !string.IsNullOrEmpty(f)
                && System.IO.File.Exists(f)).ToList();

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
                    // Fallback: mở file bình thường
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
            MessageBox.Show(msg, "Hoàn tất", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void DgvPO_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvPO.SelectedRows.Count == 0) return;
            _selectedPO_ID = Convert.ToInt32(dgvPO.SelectedRows[0].Cells["ID"].Value);
            var p = _poSummaries.Find(x => x.PO_ID == _selectedPO_ID);
            if (p == null) return;

            lblPOName.Text = $"PO: {p.PONo}  —  {p.Project_Name}  |  NCC: {p.Supplier_Name}";
            lblPOAmount.Text = $"Tổng PO: {p.Total_PO_Amount:N0} VNĐ";
            lblPOPaid.Text = $"Đã TT: {p.Total_Paid:N0} VNĐ";
            lblPORemain.Text = $"Còn nợ: {p.Amount_Remaining:N0} VNĐ";
            lblPOStatus.Text = p.Is_Overdue ? "⚠ QUÁ HẠN" : p.Payment_Status;
            lblPOStatus.ForeColor =
                p.Is_Overdue ? Color.FromArgb(255, 100, 100) :
                p.Payment_Status == "Đã TT đủ" ? Color.FromArgb(144, 238, 144) :
                p.Payment_Status == "Một phần" ? Color.FromArgb(255, 200, 100) :
                                                   Color.White;

            int pct = (int)Math.Min(p.Percent_Paid, 100);
            progressPO.Value = pct;
            lblPOProgress.Text = $"{pct}%";

            LoadSchedHist();
            LoadDocuments(); // Load Invoice + Delivery Note
        }

        private void DgvPO_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string col = dgvPO.Columns[e.ColumnIndex].Name;
            if (col == "TT_Status")
            {
                string v = e.Value?.ToString() ?? "";
                e.CellStyle.ForeColor =
                    v == "Đã TT đủ" ? Color.FromArgb(40, 167, 69) :
                    v == "Một phần" ? Color.FromArgb(255, 140, 0) :
                                      Color.FromArgb(0, 120, 212);
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
            if (_selectedPO_ID == 0) { Warn("Vui lòng chọn PO!"); return; }

            // Lấy tổng PO sau thuế của PO đang chọn
            var po = _poSummaries.Find(x => x.PO_ID == _selectedPO_ID);
            decimal poTotalAfterVat = po?.Total_PO_Amount ?? 0;

            // % mặc định: 100% nếu chưa có đợt nào, ngược lại tính phần còn lại
            decimal usedPct = 0;
            foreach (DataGridViewRow row in dgvSchedule.Rows)
                if (decimal.TryParse(row.Cells["Percent_TT"].Value?.ToString(), out decimal rp)) usedPct += rp;
            decimal defaultPct = Math.Max(0, 100 - usedPct);
            decimal defaultAmt = poTotalAfterVat > 0 ? Math.Round(poTotalAfterVat * defaultPct / 100, 0) : 0;

            int i = dgvSchedule.Rows.Add();
            var r = dgvSchedule.Rows[i];
            r.Cells["S_ID"].Value = 0;
            r.Cells["Dot_TT"].Value = _schedules.Count + dgvSchedule.Rows.Count;
            r.Cells["Pay_Method"].Value = "Full";
            r.Cells["Payment_Type"].Value = "Chuyển khoản";
            r.Cells["Percent_TT"].Value = defaultPct;
            r.Cells["Amount_Plan"].Value = defaultAmt.ToString("N0");
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
                decimal amt = Math.Round(poTotal * pct / 100, 0);
                row.Cells["Amount_Plan"].Value = amt.ToString("N0");
            }
        }

        private void BtnDelSched_Click(object sender, EventArgs e)
        {
            if (_selectedSchedID == 0) { Warn("Vui lòng chọn đợt cần xóa!"); return; }
            if (Ask("Xóa đợt thanh toán này?"))
            {
                try { _svc.DeleteSchedule(_selectedSchedID); LoadSchedHist(); LoadPOSummary(); }
                catch (Exception ex) { Err(ex.Message); }
            }
        }

        private async void BtnSaveSched_Click(object sender, EventArgs e)
        {
            if (_selectedPO_ID == 0) return;
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
                        Due_Date = DateTime.TryParse(row.Cells["Due_Date"].Value?.ToString(), out DateTime dd) ? dd : (DateTime?)null,
                        Delivery_Ref = row.Cells["Delivery_Ref"].Value?.ToString() ?? "",
                        Description = row.Cells["Description"].Value?.ToString() ?? "",
                        Status = row.Cells["S_Status"].Value?.ToString() ?? "Chưa TT"
                    };
                    if (s.Schedule_ID == 0) _svc.InsertSchedule(s, _currentUser);
                    else _svc.UpdateSchedule(s);
                    saved++;
                }
                MessageBox.Show($"✅ Đã lưu {saved} đợt thanh toán!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Ghi nhớ PO đang chọn để giữ hiển thị sau khi refresh
                int savedPoId = _selectedPO_ID;

                // Chỉ reload schedule/history của PO này — không reload toàn bộ grid PO
                LoadSchedHist();

                // Reload dữ liệu từ DB trên background thread
                await System.Threading.Tasks.Task.Run(() =>
                {
                    try
                    {
                        // Cập nhật cache schedules
                        var newScheds = _svc.GetSchedules(savedPoId);
                        _allSchedulesCache[savedPoId] = newScheds;

                        // Reload lại summary của PO này từ DB để lấy Next_Due_Date mới nhất
                        var freshSummary = _svc.GetPOSummary(savedPoId);
                        if (freshSummary != null)
                        {
                            int idx = _poSummaries.FindIndex(p => p.PO_ID == savedPoId);
                            if (idx >= 0) _poSummaries[idx] = freshSummary;
                        }
                    }
                    catch { }
                });

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

        private void BtnAddPayment_Click(object sender, EventArgs e)
        {
            if (_selectedPO_ID == 0) { Warn("Vui lòng chọn PO!"); return; }
            using var dlg = new frmAddPayment(_selectedPO_ID, _schedules, _currentUser);
            if (dlg.ShowDialog() == DialogResult.OK) { LoadSchedHist(); LoadPOSummary(); }
        }

        private void BtnDelPayment_Click(object sender, EventArgs e)
        {
            if (_selectedHistID == 0) { Warn("Vui lòng chọn bản ghi!"); return; }
            if (Ask("Xóa lịch sử thanh toán này?"))
            {
                try { _svc.DeleteHistory(_selectedHistID); LoadSchedHist(); LoadPOSummary(); }
                catch (Exception ex) { Err(ex.Message); }
            }
        }

        private void BtnPaymentRequest_Click(object sender, EventArgs e)
        {
            if (_selectedPO_ID == 0)
            {
                Warn("Vui lòng chọn một PO trong danh sách để tạo yêu cầu!");
                return;
            }

            var po = _poSummaries.Find(p => p.PO_ID == _selectedPO_ID);
            var poHead = _poSvc.GetAll().Find(p => p.PO_ID == _selectedPO_ID);
            string mprNo = poHead?.MPR_No ?? "";

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
            try
            {
                int? suppId = null;
                if (cboSuppFilter.SelectedIndex > 0)
                {
                    var name = cboSuppFilter.SelectedItem.ToString();
                    var s = _allSuppliers.Find(x => x.Company_Name == name);
                    if (s != null) suppId = s.Supplier_ID;
                }

                _debtReport = _svc.GetDebtReport(dtpFrom.Value, dtpTo.Value, suppId);
                _suppDebt = _svc.GetSupplierDebt();

                dgvDebtSupp.Rows.Clear();
                decimal tVal = 0, tPaid = 0, tDebt = 0; int tOver = 0;
                foreach (var s in _suppDebt)
                {
                    int i = dgvDebtSupp.Rows.Add();
                    var r = dgvDebtSupp.Rows[i];
                    r.Cells["D_SuppID"].Value = s.Supplier_ID;
                    r.Cells["D_Name"].Value = s.Supplier_Name;
                    r.Cells["D_TotalPO"].Value = s.Total_PO;
                    r.Cells["D_Value"].Value = s.Total_PO_Value.ToString("N0");
                    r.Cells["D_Debt"].Value = s.Total_Debt.ToString("N0");
                    r.Cells["D_Overdue"].Value = s.Overdue_PO_Count > 0 ? $"⚠ {s.Overdue_PO_Count}" : "—";

                    tVal += s.Total_PO_Value;
                    tPaid += s.Total_Paid;
                    tDebt += s.Total_Debt;
                    tOver += s.Overdue_PO_Count;
                }

                lblSumValue.Text = $"{tVal:N0} VNĐ";
                lblSumPaid.Text = $"{tPaid:N0} VNĐ";
                lblSumDebt.Text = $"{tDebt:N0} VNĐ";
                lblSumOverdue.Text = $"{tOver} PO";

                BindDebtDetail(_debtReport);
            }
            catch (Exception ex) { Err(ex.Message); }
        }

        private void DgvDebtSupp_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvDebtSupp.SelectedRows.Count == 0) return;
            int sid = Convert.ToInt32(dgvDebtSupp.SelectedRows[0].Cells["D_SuppID"].Value);
            BindDebtDetail(_debtReport.FindAll(r => r.Supplier_ID == sid));
        }

        private void DgvDebtSupp_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (dgvDebtSupp.Columns[e.ColumnIndex].Name == "D_Debt")
            { e.CellStyle.ForeColor = Color.FromArgb(220, 53, 69); e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold); }
            if (dgvDebtSupp.Columns[e.ColumnIndex].Name == "D_Overdue" && e.Value?.ToString() != "—")
            { e.CellStyle.ForeColor = Color.FromArgb(220, 53, 69); e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold); }
        }

        private void BindDebtDetail(List<DebtReportItem> items)
        {
            dgvDebtDetail.Rows.Clear();
            foreach (var d in items)
            {
                int i = dgvDebtDetail.Rows.Add();
                var r = dgvDebtDetail.Rows[i];
                r.Cells["DD_PONo"].Value = d.PONo;
                r.Cells["DD_Project"].Value = d.Project_Name;
                r.Cells["DD_PODate"].Value = d.PO_Date?.ToString("dd/MM/yyyy") ?? "";
                r.Cells["DD_Total"].Value = d.Total_Amount.ToString("N0");
                r.Cells["DD_Before"].Value = d.Paid_Before_Range.ToString("N0");
                r.Cells["DD_InRange"].Value = d.Paid_In_Range.ToString("N0");
                r.Cells["DD_Remain"].Value = d.Remaining_Debt.ToString("N0");
                r.Cells["DD_Status"].Value = d.Is_Overdue ? "⚠ Quá hạn" : d.Payment_Status;
                r.Cells["DD_Due"].Value = d.Next_Due_Date?.ToString("dd/MM/yyyy") ?? "—";
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
                    v.Contains("Quá hạn") ? Color.FromArgb(220, 53, 69) :
                    v == "Đã TT đủ" ? Color.FromArgb(40, 167, 69) :
                    v == "Một phần" ? Color.FromArgb(255, 140, 0) :
                                            Color.FromArgb(0, 120, 212);
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
            if (col == "DD_Remain")
            {
                e.CellStyle.ForeColor = Color.FromArgb(220, 53, 69);
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
        }

        // =====================================================================
        //  XUẤT EXCEL
        // =====================================================================
        private void BtnExportDebt_Click(object sender, EventArgs e)
        {
            if (_debtReport.Count == 0) { Warn("Vui lòng xem báo cáo trước!"); return; }
            using var sfd = new SaveFileDialog
            {
                Title = "Lưu báo cáo công nợ",
                Filter = "Excel|*.xlsx",
                FileName = $"CongNo_{dtpFrom.Value:yyyyMMdd}_{dtpTo.Value:yyyyMMdd}"
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
                    ws.Cells[row, 5].Value = d.Total_Amount; ws.Cells[row, 5].Style.Numberformat.Format = "#,##0";
                    ws.Cells[row, 6].Value = d.Paid_Before_Range; ws.Cells[row, 6].Style.Numberformat.Format = "#,##0";
                    ws.Cells[row, 7].Value = d.Paid_In_Range; ws.Cells[row, 7].Style.Numberformat.Format = "#,##0";
                    ws.Cells[row, 8].Value = d.Remaining_Debt; ws.Cells[row, 8].Style.Numberformat.Format = "#,##0";
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
                MessageBox.Show("✅ Xuất Excel thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = sfd.FileName, UseShellExecute = true });
            }
            catch (Exception ex) { Err(ex.Message); }
        }

        private void ResizeAll()
        {
            try
            {
                int w = tabPO.ClientSize.Width;
                int h = tabPO.ClientSize.Height;

                if (panelTop != null) panelTop.Width = w - 10;
                if (panelInfo != null) { panelInfo.Width = w - 10; lblPOStatus.Left = panelInfo.Width - 205; }

                int leftW = w / 2 - 8;

                // panelSched cố định chiều cao 200
                if (panelSched != null)
                {
                    panelSched.Width = leftW;
                    panelSched.Height = 200;
                    dgvSchedule.Width = panelSched.Width - 10;
                    dgvSchedule.Height = panelSched.Height - 65;
                }

                // panelPrintHistory chiếm phần còn lại bên dưới panelSched
                if (panelPrintHistory != null)
                {
                    int printTop = panelSched.Bottom + 5;
                    panelPrintHistory.Top = printTop;
                    panelPrintHistory.Width = leftW;
                    panelPrintHistory.Height = Math.Max(100, h - printTop - 10);
                    dgvPrintHistory.Width = panelPrintHistory.Width - 10;
                    dgvPrintHistory.Height = panelPrintHistory.Height - 63; // 58 toolbar + 5
                    // Resize bộ lọc theo chiều rộng
                    if (_phDateTo != null) _phDateTo.Width = Math.Min(115, (panelPrintHistory.Width - 470) / 2);
                }

                if (panelHist != null)
                {
                    panelHist.Left = w / 2 + 3;
                    panelHist.Width = w / 2 - 8;
                    panelHist.Height = 200; // cố định
                    dgvHistory.Width = panelHist.Width - 10;
                }

                // panelDoc bên dưới panelHist, co giãn
                if (panelDoc != null)
                {
                    int docTop = panelHist.Bottom + 5;
                    panelDoc.Left = w / 2 + 3;
                    panelDoc.Top = docTop;
                    panelDoc.Width = w / 2 - 8;
                    panelDoc.Height = Math.Max(80, h - docTop - 10);
                    dgvDoc.Width = panelDoc.Width - 10;
                    dgvDoc.Height = panelDoc.Height - 37;
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
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
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

        private void Warn(string msg) =>
            MessageBox.Show(msg, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        private void Err(string msg) =>
            MessageBox.Show("Lỗi: " + msg, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        private bool Ask(string msg) =>
            MessageBox.Show(msg, "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;

        // Được gọi từ frmPrintPreview khi user chọn OK cập nhật lịch sử
        public void AddPrintHistory(string poNo, string project, List<PaymentSchedule> scheds)
        {
            if (dgvPrintHistory == null) return;
            string dateStr = DateTime.Now.ToString("dd/MM/yyyy HH:mm");

            foreach (var s in scheds)
            {
                decimal net = s.Amount_Plan;
                decimal vat = Math.Round(net * 0.1m, 0);
                decimal total = net + vat;
                string dot = s.Dot_TT == 1 ? "1st" : s.Dot_TT == 2 ? "2nd" :
                                s.Dot_TT == 3 ? "3rd" : $"{s.Dot_TT}th";

                // ── Lưu vào DB ──
                try
                {
                    using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                    conn.Open();
                    var cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                        INSERT INTO PO_PrintRequestHistory
                            (PONo, Project_Name, Dot_TT, Dot_Label,
                             Amount_Net, Amount_VAT, Amount_Total,
                             Printed_By, Printed_Date)
                        VALUES
                            (@poNo, @proj, @dot, @dotLabel,
                             @net, @vat, @total,
                             @by, GETDATE())", conn);
                    cmd.Parameters.AddWithValue("@poNo", poNo);
                    cmd.Parameters.AddWithValue("@proj", project ?? "");
                    cmd.Parameters.AddWithValue("@dot", s.Dot_TT);
                    cmd.Parameters.AddWithValue("@dotLabel", dot);
                    cmd.Parameters.AddWithValue("@net", net);
                    cmd.Parameters.AddWithValue("@vat", vat);
                    cmd.Parameters.AddWithValue("@total", total);
                    cmd.Parameters.AddWithValue("@by", _currentUser ?? "");
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("AddPrintHistory DB error: " + ex.Message);
                }

                // ── Thêm vào đầu grid (mới nhất lên trên) ──
                dgvPrintHistory.Rows.Insert(0);
                dgvPrintHistory.Rows[0].Cells["PH_ID"].Value = DBNull.Value; // sẽ có ID sau khi reload
                dgvPrintHistory.Rows[0].Cells["PH_PONo"].Value = poNo;
                dgvPrintHistory.Rows[0].Cells["PH_Project"].Value = project;
                dgvPrintHistory.Rows[0].Cells["PH_Dot"].Value = dot;
                dgvPrintHistory.Rows[0].Cells["PH_Net"].Value = net.ToString("N0");
                dgvPrintHistory.Rows[0].Cells["PH_Vat"].Value = vat.ToString("N0");
                dgvPrintHistory.Rows[0].Cells["PH_Total"].Value = total.ToString("N0");
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
            DateTime dtFrom = from ?? DateTime.Today.AddMonths(-3);
            DateTime dtTo = to ?? DateTime.Today.AddDays(1).AddSeconds(-1);
            try
            {
                using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                conn.Open();
                var cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                    SELECT Print_ID, PONo, Project_Name, Dot_Label,
                           Amount_Net, Amount_VAT, Amount_Total,
                           Printed_By,
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
                    decimal net = reader["Amount_Net"] != DBNull.Value ? Convert.ToDecimal(reader["Amount_Net"]) : 0;
                    decimal vat = reader["Amount_VAT"] != DBNull.Value ? Convert.ToDecimal(reader["Amount_VAT"]) : 0;
                    decimal total = reader["Amount_Total"] != DBNull.Value ? Convert.ToDecimal(reader["Amount_Total"]) : 0;

                    int i = dgvPrintHistory.Rows.Add();
                    dgvPrintHistory.Rows[i].Cells["PH_ID"].Value = reader["Print_ID"];
                    dgvPrintHistory.Rows[i].Cells["PH_PONo"].Value = reader["PONo"]?.ToString() ?? "";
                    dgvPrintHistory.Rows[i].Cells["PH_Project"].Value = reader["Project_Name"]?.ToString() ?? "";
                    dgvPrintHistory.Rows[i].Cells["PH_Dot"].Value = reader["Dot_Label"]?.ToString() ?? "";
                    dgvPrintHistory.Rows[i].Cells["PH_Net"].Value = net.ToString("N0");
                    dgvPrintHistory.Rows[i].Cells["PH_Vat"].Value = vat.ToString("N0");
                    dgvPrintHistory.Rows[i].Cells["PH_Total"].Value = total.ToString("N0");
                    dgvPrintHistory.Rows[i].Cells["PH_Date"].Value = reader["Printed_Date"]?.ToString() ?? "";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("LoadPrintHistory error: " + ex.Message);
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

            if (MessageBox.Show(
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
                MessageBox.Show("✅ Đã xóa thành công!", "Thông báo",
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

        // =====================================================================
        //  IN REQUEST — Fill payment_template.xlsx rồi hiện Print Preview
        // =====================================================================
        private void BtnPrintRequest_Click(object sender, EventArgs e)
        {
            if (_selectedPO_ID == 0) { Warn("Vui lòng chọn một PO trước!"); return; }

            // Kiểm tra PO đã được in trong 3 tháng gần nhất chưa
            var po = _poSummaries.Find(x => x.PO_ID == _selectedPO_ID);
            if (po != null && CheckAlreadyPrinted(po.PONo, out string lastDate))
            {
                var ans = MessageBox.Show(
                    $"⚠ PO \"{po.PONo}\" đã được in Request trước đó.\n" +
                    $"Lần in gần nhất: {lastDate}\n\n" +
                    "Bạn có muốn in lại không?",
                    "Đã in trước đó",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (ans != DialogResult.Yes) return;
            }

            PrintPaymentRequest();
        }

        private void PrintPaymentRequest()
        {
            try
            {
                var po = _poSummaries.Find(x => x.PO_ID == _selectedPO_ID);
                if (po == null) { Warn("Không tìm thấy thông tin PO!"); return; }

                var poHead = _poSvc.GetAll().Find(x => x.PO_ID == _selectedPO_ID);
                var scheds = _allSchedulesCache.ContainsKey(_selectedPO_ID)
                              ? _allSchedulesCache[_selectedPO_ID]
                              : _svc.GetSchedules(_selectedPO_ID);

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
                    return;
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

                    // A1 — (N)th Payment Request
                    int paidDots = scheds.Count(s => s.Status == "Đã TT đủ");
                    string ordinal = (paidDots + 1) switch { 1 => "1st", 2 => "2nd", 3 => "3rd", _ => $"{paidDots + 1}th" };
                    ReplaceCell(ws, "(   )th  Payment Request", $"({ordinal}) Payment Request");

                    // A3 — Project Name (ô C3 trống, điền tên dự án sau dấu ":")
                    // Tìm ô C3 hoặc ô cạnh B3 để điền
                    FillNextCell(ws, "A3", "Project Name", po.Project_Name ?? "");

                    // C5 — W/O No, M5 — PO No
                    ReplaceCell(ws, "<<WO-NO>>", poHead?.WorkorderNo ?? "");
                    ReplaceCell(ws, "<<PO-NO>>", po.PONo ?? "");

                    // A6 Contract date — lấy PO_Date
                    string contractDate = po.PO_Date.HasValue ? po.PO_Date.Value.ToString("dd/MM/yyyy") : "";
                    FillNextCell(ws, "A6", "Contract date", contractDate);

                    // I6 Payment date — ngày hôm nay
                    string paymentDate = DateTime.Today.ToString("dd/MM/yyyy");
                    FillRightCell(ws, "I6", "Payment date", paymentDate);

                    // C7 — Contract amount (tổng trước VAT)
                    ReplaceCell(ws, "<<Tổng số tiền trước thuế>>", totalBeforeVat.ToString("N0"));

                    // C8 — Requested amount (tổng đợt chưa TT)
                    decimal reqAmt = scheds.Where(s => s.Status != "Đã TT đủ").Sum(s => s.Amount_Plan);
                    ReplaceCell(ws, "<<Số tiền theo đợt>>", reqAmt.ToString("N0"));

                    // ── Lấy ngày thanh toán thực tế của các đợt đã TT ──
                    // Key = Dot_TT (số đợt), Value = ngày thanh toán thực tế
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
                        if (i < dotCount)
                        {
                            var s = scheds[i];
                            decimal net = s.Amount_Plan;
                            decimal vat = Math.Round(net * 0.1m, 0);
                            decimal tot = net + vat;
                            sumNet += net;
                            sumVat += vat;
                            sumTotal += tot;

                            // Ngày thanh toán:
                            // - Đã TT đủ → lấy ngày thực tế từ PaymentHistory
                            // - Đợt hiện tại (chưa TT) → Due_Date hoặc hôm nay
                            // - Đợt chưa đến → để trống
                            string dateValue;
                            if (s.Status == "Đã TT đủ")
                            {
                                // Đợt đã thanh toán: hiển thị ngày TT thực tế
                                actualPayDates.TryGetValue(s.Dot_TT, out dateValue);
                                dateValue = dateValue ?? (s.Due_Date.HasValue
                                    ? s.Due_Date.Value.ToString("dd/MM/yyyy") : "");
                            }
                            else if (i == scheds.FindIndex(x => x.Status != "Đã TT đủ"))
                            {
                                // Đây là đợt đang request (đợt chưa TT đầu tiên)
                                // → hiển thị Due_Date (ngày dự kiến TT)
                                dateValue = s.Due_Date.HasValue
                                    ? s.Due_Date.Value.ToString("dd/MM/yyyy")
                                    : DateTime.Today.ToString("dd/MM/yyyy");
                            }
                            else
                            {
                                // Đợt chưa đến (chưa TT và không phải đợt hiện tại) → để trống
                                dateValue = "";
                            }

                            ReplaceCell(ws, $"<<Số tiền đợt {i + 1}>>", net.ToString("N0"));
                            ReplaceCell(ws, $"<<Số tiền thuế lần {i + 1}>>", vat.ToString("N0"));
                            ReplaceCell(ws, $"<<Số tiền sau thuế lần {i + 1}>>", tot.ToString("N0"));
                            ReplaceCell(ws, $"<<Ngày yêu cầu lần {i + 1}>>", dateValue);
                        }
                        else
                        {
                            ReplaceCell(ws, $"<<Số tiền đợt {i + 1}>>", "");
                            ReplaceCell(ws, $"<<Số tiền thuế lần {i + 1}>>", "");
                            ReplaceCell(ws, $"<<Số tiền sau thuế lần {i + 1}>>", "");
                            ReplaceCell(ws, $"<<Ngày yêu cầu lần {i + 1}>>", "");
                        }
                    }

                    // Row 18 — Sum (3 ô đều có placeholder <<Sum>>)
                    ReplaceCellAll(ws, "<<Sum>>", new[]
                    {
                        sumNet.ToString("N0"),
                        sumVat.ToString("N0"),
                        sumTotal.ToString("N0")
                    });

                    // Row 19 — Balance
                    decimal balNet = Math.Max(totalBeforeVat - sumNet, 0);
                    decimal balTotal = Math.Max(totalBeforeVat * 1.1m - sumTotal - totalPaid, 0);
                    ReplaceCell(ws, "<<Tổng số tiền trước thuế còn lại>>", balNet.ToString("N0"));
                    ReplaceCell(ws, "<<Tổng số tiền sau thuế còn lại>>", balTotal.ToString("N0"));

                    // A26 — Ngày yêu cầu (ngày ký)
                    ReplaceCell(ws, "<<Ngày yêu cầu>>", DateTime.Today.ToString("dd/MM/yyyy"));

                    // Supplier info
                    string suppName = supp.Company_Name ?? supp.Supplier_Name ?? "";
                    string suppAddress = GetSupplierProp(supp, "Company_Address", "Address") ?? "";
                    ReplaceCell(ws, "<<Tên nhà cung cấp>>", suppName);
                    ReplaceCell(ws, "<<Địa chỉ Nhà cung cấp>>", suppAddress);

                    pkg.Save();
                }

                // ── Mở Print Preview ──
                var dlg = new frmPrintPreview(tempPath, po.PONo, po.Project_Name, scheds, this);
                dlg.ShowDialog(this);

                // Dọn file tạm sau khi đóng
                try { if (System.IO.File.Exists(tempPath)) System.IO.File.Delete(tempPath); } catch { }
            }
            catch (Exception ex) { Err("Lỗi tạo file in: " + ex.Message); }
        }

        // Fill ô ngay bên phải label (dùng cho Contract date, Project Name...)
        private void FillNextCell(OfficeOpenXml.ExcelWorksheet ws, string cellAddr, string label, string value)
        {
            try
            {
                var cell = ws.Cells[cellAddr];
                if (cell.Value?.ToString()?.Contains(label) == true)
                {
                    // Điền vào ô C cùng hàng (sau cột B ": ")
                    int row = cell.Start.Row;
                    ws.Cells[row, 3].Value = value;
                }
            }
            catch { }
        }

        // Fill ô bên phải tiêu đề bên phải (Payment date ở I6 → fill vào M6 hoặc N6)
        private void FillRightCell(OfficeOpenXml.ExcelWorksheet ws, string cellAddr, string label, string value)
        {
            try
            {
                var cell = ws.Cells[cellAddr];
                if (cell.Value?.ToString()?.Contains(label) == true)
                {
                    int row = cell.Start.Row;
                    ws.Cells[row, 13].Value = value; // cột M
                }
            }
            catch { }
        }

        private void ReplaceCell(OfficeOpenXml.ExcelWorksheet ws, string placeholder, string value)
        {
            foreach (var cell in ws.Cells[ws.Dimension.Address])
                if (cell.Value?.ToString() == placeholder)
                    cell.Value = value;
        }

        // Thay nhiều ô có cùng placeholder bằng các giá trị khác nhau theo thứ tự
        private void ReplaceCellAll(OfficeOpenXml.ExcelWorksheet ws, string placeholder, string[] values)
        {
            int idx = 0;
            foreach (var cell in ws.Cells[ws.Dimension.Address])
                if (cell.Value?.ToString() == placeholder && idx < values.Length)
                    cell.Value = values[idx++];
        }

        private string GetSupplierProp(Supplier supp, params string[] names)
        {
            if (supp == null) return "";
            var type = supp.GetType();
            foreach (var name in names)
            {
                var prop = type.GetProperty(name);
                if (prop != null) return prop.GetValue(supp)?.ToString() ?? "";
            }
            return "";
        }
    }

    // =========================================================================
    //  frmPrintPreview — Hiển thị preview và nút xác nhận in
    // =========================================================================
    public class frmPrintPreview : Form
    {
        private readonly string _filePath;
        private readonly string _poNo;
        private readonly string _project;
        private readonly List<PaymentSchedule> _scheds;
        private readonly frmPayment _owner;

        public frmPrintPreview(string filePath, string poNo, string project,
            List<PaymentSchedule> scheds, frmPayment owner)
        {
            _filePath = filePath;
            _poNo = poNo;
            _project = project ?? "";
            _scheds = scheds ?? new List<PaymentSchedule>();
            _owner = owner;
            BuildUI();
            // Hỏi cập nhật lịch sử khi form đóng
            this.FormClosing += FrmPrintPreview_FormClosing;
        }

        private bool _historyUpdated = false;

        private void FrmPrintPreview_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_historyUpdated) return; // Đã cập nhật rồi, không hỏi lại
            if (_owner == null || _scheds.Count == 0) return;

            var ans = MessageBox.Show(
                "Cập nhật thông tin vào lịch sử in Request?",
                "Xác nhận",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Question);

            if (ans == DialogResult.OK)
            {
                _owner.AddPrintHistory(_poNo, _project, _scheds);
                _historyUpdated = true;
            }
        }

        private void BuildUI()
        {
            this.Text = $"🖨 Print Preview — Payment Request  |  PO: {_poNo}";
            this.Size = new Size(900, 680);
            this.StartPosition = FormStartPosition.CenterParent;
            this.BackColor = Color.FromArgb(245, 245, 245);
            this.MinimumSize = new Size(700, 500);

            // Tiêu đề
            this.Controls.Add(new Label
            {
                Text = $"📋  Payment Request — PO: {_poNo}",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 10),
                Size = new Size(700, 28)
            });

            // Thông báo
            this.Controls.Add(new Label
            {
                Text = "File đã được tạo thành công. Bấm \"🖨 In ngay\" để mở hộp thoại in, hoặc \"📂 Mở file\" để xem chi tiết trước.",
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.FromArgb(80, 80, 80),
                Location = new Point(10, 44),
                Size = new Size(860, 40),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            });

            // Panel thông tin file
            var pInfo = new Panel
            {
                Location = new Point(10, 90),
                Size = new Size(860, 50),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            pInfo.Controls.Add(new Label
            {
                Text = "📄  " + System.IO.Path.GetFileName(_filePath),
                Font = new Font("Segoe UI", 10),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 14),
                Size = new Size(700, 22)
            });
            this.Controls.Add(pInfo);

            // Hướng dẫn
            var pGuide = new Panel
            {
                Location = new Point(10, 150),
                Size = new Size(860, 120),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            pGuide.Controls.Add(new Label
            {
                Text = "📌  Hướng dẫn:",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Location = new Point(10, 8),
                Size = new Size(200, 18)
            });
            string guide = "1. Bấm \"📂 Mở file\" để xem nội dung Payment Request trong Excel trước khi in.\r\n" +
                           "2. Bấm \"🖨 In ngay\" để gửi thẳng đến máy in (Excel sẽ mở và in tự động).\r\n" +
                           "3. Bấm \"💾 Lưu về máy\" để chọn nơi lưu file trước khi in.";
            pGuide.Controls.Add(new Label
            {
                Text = guide,
                Font = new Font("Segoe UI", 9),
                Location = new Point(10, 30),
                Size = new Size(840, 80)
            });
            this.Controls.Add(pGuide);

            // Buttons
            int btnY = this.ClientSize.Height - 50;

            var btnOpen = new Button
            {
                Text = "📂 Mở file",
                Location = new Point(10, btnY),
                Size = new Size(130, 36),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            btnOpen.FlatAppearance.BorderSize = 0;
            btnOpen.Click += (s, ev) =>
            {
                try
                {
                    // Mở bằng PowerShell ở chế độ ReadOnly → không hỏi lưu khi đóng
                    string psOpen = $@"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open('{_filePath.Replace("'", "''")}', $false, $true)
";
                    string psFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                        $"open_{System.IO.Path.GetFileNameWithoutExtension(_filePath)}.ps1");
                    System.IO.File.WriteAllText(psFile, psOpen, System.Text.Encoding.UTF8);
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = $"-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"{psFile}\"",
                        UseShellExecute = false,
                        CreateNoWindow = true
                    });
                    System.Threading.Tasks.Task.Delay(15000).ContinueWith(_ =>
                    {
                        try { if (System.IO.File.Exists(psFile)) System.IO.File.Delete(psFile); } catch { }
                    });
                }
                catch
                {
                    // Fallback thông thường
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    { FileName = _filePath, UseShellExecute = true });
                }
            };
            this.Controls.Add(btnOpen);

            var btnSave = new Button
            {
                Text = "💾 Lưu về máy",
                Location = new Point(150, btnY),
                Size = new Size(130, 36),
                BackColor = Color.FromArgb(0, 150, 100),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Click += (s, ev) =>
            {
                using var sfd = new SaveFileDialog
                {
                    Title = "Lưu Payment Request",
                    Filter = "Excel Files|*.xlsx",
                    FileName = $"PaymentRequest_{_poNo}_{DateTime.Now:yyyyMMdd}",
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                };
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    System.IO.File.Copy(_filePath, sfd.FileName, true);
                    MessageBox.Show("✅ Đã lưu file thành công!", "Thành công",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };
            this.Controls.Add(btnSave);

            var btnPrint = new Button
            {
                Text = "🖨 In ngay",
                Location = new Point(290, btnY),
                Size = new Size(120, 36),
                BackColor = Color.FromArgb(102, 51, 153),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            btnPrint.FlatAppearance.BorderSize = 0;
            btnPrint.Click += (s, ev) =>
            {
                try
                {
                    // Dùng PowerShell để in Excel ẩn — không hiện hộp thoại lưu file
                    string psScript = $@"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open('{_filePath.Replace("'", "''")}')
$wb.PrintOut()
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
";
                    string psFile = System.IO.Path.Combine(
                        System.IO.Path.GetTempPath(),
                        $"print_{System.IO.Path.GetFileNameWithoutExtension(_filePath)}.ps1");
                    System.IO.File.WriteAllText(psFile, psScript, System.Text.Encoding.UTF8);

                    var psi = new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = $"-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"{psFile}\"",
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
                    };
                    var proc = System.Diagnostics.Process.Start(psi);

                    // Không chờ process kết thúc để UI không bị block
                    MessageBox.Show("✅ Đã gửi lệnh in!\nFile sẽ được in mà không cần lưu lại.",
                        "In thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();

                    // Dọn file PS tạm sau 30s
                    System.Threading.Tasks.Task.Delay(30000).ContinueWith(_ =>
                    {
                        try { if (System.IO.File.Exists(psFile)) System.IO.File.Delete(psFile); } catch { }
                    });
                }
                catch (Exception ex)
                {
                    // Fallback: mở file bình thường nếu PowerShell không khả dụng
                    var ans = MessageBox.Show(
                        $"Không thể in tự động: {ex.Message}\n\nBấm OK để mở file và in thủ công.",
                        "Lỗi in", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (ans == DialogResult.OK)
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        { FileName = _filePath, UseShellExecute = true });
                }
            };
            this.Controls.Add(btnPrint);

            var btnClose = new Button
            {
                Text = "Đóng",
                Location = new Point(this.ClientSize.Width - 110, btnY),
                Size = new Size(100, 36),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                DialogResult = DialogResult.Cancel
            };
            btnClose.FlatAppearance.BorderSize = 0;
            this.Controls.Add(btnClose);
            this.CancelButton = btnClose;

            this.Resize += (s, ev) =>
            {
                btnClose.Location = new Point(this.ClientSize.Width - 110, this.ClientSize.Height - 50);
                pInfo.Width = this.ClientSize.Width - 20;
                pGuide.Width = this.ClientSize.Width - 20;
            };
        }
    }

    public class frmAddPayment : Form
    {
        private readonly int _poId;
        private readonly List<PaymentSchedule> _scheds;
        private readonly string _user;
        private readonly PaymentService _svc = new PaymentService();

        private ComboBox cboSched, cboMethod;
        private DateTimePicker dtpDate;
        private TextBox txtAmount, txtBank, txtTransNo, txtNotes;
        private Label lblErr;

        public frmAddPayment(int poId, List<PaymentSchedule> scheds, string user)
        {
            _poId = poId;
            _scheds = scheds;
            _user = user;
            BuildUI();
        }

        private void BuildUI()
        {
            this.Text = "Ghi nhận thanh toán";
            this.Size = new Size(470, 420);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.BackColor = Color.White;

            int y = 12;
            Row("Liên kết đợt TT:", y);
            cboSched = new ComboBox
            {
                Location = new Point(160, y),
                Size = new Size(280, 26),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboSched.Items.Add("— Không liên kết —");
            foreach (var s in _scheds)
                cboSched.Items.Add($"Đợt {s.Dot_TT}: {s.Amount_Plan:N0} VNĐ  [{s.Status}]");
            cboSched.SelectedIndex = 0;
            this.Controls.Add(cboSched);

            y += 42; Row("Ngày thanh toán (*):", y);
            dtpDate = new DateTimePicker
            {
                Location = new Point(160, y),
                Size = new Size(150, 26),
                Font = new Font("Segoe UI", 9),
                Format = DateTimePickerFormat.Short,
                Value = DateTime.Today
            };
            this.Controls.Add(dtpDate);

            y += 42; Row("Số tiền (*) VNĐ:", y);
            txtAmount = new TextBox
            {
                Location = new Point(160, y),
                Size = new Size(200, 26),
                Font = new Font("Segoe UI", 9),
                PlaceholderText = "Ví dụ: 50000000"
            };
            this.Controls.Add(txtAmount);

            y += 42; Row("Hình thức TT:", y);
            cboMethod = new ComboBox
            {
                Location = new Point(160, y),
                Size = new Size(180, 26),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboMethod.Items.AddRange(new[] { "Chuyển khoản", "Tiền mặt", "LC", "TT" });
            cboMethod.SelectedIndex = 0;
            this.Controls.Add(cboMethod);

            y += 42; Row("Ngân hàng:", y);
            txtBank = new TextBox { Location = new Point(160, y), Size = new Size(280, 26), Font = new Font("Segoe UI", 9) };
            this.Controls.Add(txtBank);

            y += 42; Row("Số chứng từ:", y);
            txtTransNo = new TextBox { Location = new Point(160, y), Size = new Size(280, 26), Font = new Font("Segoe UI", 9) };
            this.Controls.Add(txtTransNo);

            y += 42; Row("Ghi chú:", y);
            txtNotes = new TextBox { Location = new Point(160, y), Size = new Size(280, 26), Font = new Font("Segoe UI", 9) };
            this.Controls.Add(txtNotes);

            y += 40;
            lblErr = new Label
            {
                Location = new Point(12, y),
                Size = new Size(440, 22),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.Red
            };
            this.Controls.Add(lblErr);

            var bOK = new Button
            {
                Text = "✔ Ghi nhận",
                Location = new Point(12, y + 28),
                Size = new Size(155, 35),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };
            bOK.FlatAppearance.BorderSize = 0;
            bOK.Click += BtnOK_Click;
            this.Controls.Add(bOK);

            var bCan = new Button
            {
                Text = "Hủy",
                Location = new Point(177, y + 28),
                Size = new Size(100, 35),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10)
            };
            bCan.FlatAppearance.BorderSize = 0;
            bCan.Click += (s, ev) => this.Close();
            this.Controls.Add(bCan);

            this.Height = y + 100;
        }

        private void Row(string lbl, int y) =>
            this.Controls.Add(new Label
            {
                Text = lbl,
                Location = new Point(12, y + 4),
                Size = new Size(148, 22),
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            });

        private void BtnOK_Click(object sender, EventArgs e)
        {
            lblErr.Text = "";
            if (!decimal.TryParse(txtAmount.Text.Replace(",", ""), out decimal amt) || amt <= 0)
            { lblErr.Text = "Vui lòng nhập số tiền hợp lệ!"; return; }

            try
            {
                int? schedId = cboSched.SelectedIndex > 0
                    ? _scheds[cboSched.SelectedIndex - 1].Schedule_ID
                    : (int?)null;

                var po = new POService().GetAll().Find(p => p.PO_ID == _poId);

                _svc.InsertHistory(new PaymentHistory
                {
                    PO_ID = _poId,
                    Schedule_ID = schedId,
                    Supplier_ID = po?.Supplier_ID,
                    Payment_Date = dtpDate.Value,
                    Amount_Paid = amt,
                    Payment_Method = cboMethod.SelectedItem?.ToString() ?? "Chuyển khoản",
                    Bank_Name = txtBank.Text.Trim(),
                    Transaction_No = txtTransNo.Text.Trim(),
                    Notes = txtNotes.Text.Trim()
                }, _user);

                MessageBox.Show($"✅ Đã ghi nhận {amt:N0} VNĐ!", "Thành công",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                lblErr.Text = "Lỗi: " + ex.Message;
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
                cboDot.Items.Add($"Đợt {s.Dot_TT}: {s.Amount_Plan:N0} VNĐ  [{s.Status}]");
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

                sb.AppendLine($"{stt++}\t{itemName}\t{material}\t{d.Asize}\t{d.Bsize}\t{d.Csize}\t{q}\t{d.UNIT}\t{wk}\t{realPrice:N0}\t{amtAfterVat:N0}");
            }

            sb.AppendLine($"\t\t\t\t\t\t\t\t\tSUB-TOTAL\t{subTotal:N0}");
            sb.AppendLine($"\t\t\t\t\t\t\t\t\tFinal Price Requested (Included {vatPct:N0}% VAT)\t{finalTotal:N0}");
            sb.AppendLine();
            sb.AppendLine("3. Amount");
            sb.AppendLine();
            sb.AppendLine($"Total Amount:\t\t{subTotal:N0} VNĐ (excluded VAT)");
            sb.AppendLine();

            // ── Final amount: luôn là số tiền SAU thuế ──
            decimal finalAmt = finalTotal; // mặc định = tổng sau VAT
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
            sb.AppendLine($"Final amount :\t\t{finalAmt:N0} VNĐ included {vatPct:N0}% VAT ({vatAmount:N0} VNĐ){dotLabel}");
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
            MessageBox.Show("✅ Đã copy nội dung vào Bảng tạm!\nDán (Ctrl+V) vào Word hoặc Outlook sẽ hiển thị bảng kẻ ô chuẩn, font Times New Roman.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // =====================================================================
        // THUẬT TOÁN ĐẨY HTML LÊN CLIPBOARD BẰNG BYTE OFFSET (CHỐNG LỖI UTF-8)
        // =====================================================================
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
            Clipboard.SetDataObject(obj, true);
        }
    }
}