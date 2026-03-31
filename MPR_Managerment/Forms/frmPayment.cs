// ============================================================
//  FILE: Forms/frmPayment.cs
//  Tab 1: Tiến độ thanh toán từng PO
//  Tab 2: Báo cáo tổng hợp công nợ NCC theo kỳ
// ============================================================
using System;
using System.Collections.Generic;
using System.Drawing;
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

        // Controls chính
        private TabControl tabs;
        private TabPage tabPO, tabDebt;

        // Tab PO
        private TextBox txtSearchPO;
        private ComboBox cboStatusFilter;
        private DataGridView dgvPO, dgvSchedule, dgvHistory;
        private Label lblPOName, lblPOAmount, lblPOPaid, lblPORemain, lblPOStatus, lblPOProgress;
        private Panel panelTop, panelInfo, panelSched, panelHist;
        private ProgressBar progressPO;

        // Tab Debt
        private DateTimePicker dtpFrom, dtpTo;
        private ComboBox cboSuppFilter;
        private DataGridView dgvDebtSupp, dgvDebtDetail;
        private Label lblSumValue, lblSumPaid, lblSumDebt, lblSumOverdue;
        private Button btnExportDebt;
        // Làm mới dữ liệu PO mỗi 5 phút
        private Button btnRefreshPO;
        

        // =====================================================================
        public frmPayment()
        {
            InitializeComponent();
            BuildUI();
            LoadData();
            this.Resize += (s, e) => ResizeAll();
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

        // =====================================================================
        //  TAB 1 — TIẾN ĐỘ THANH TOÁN PO
        // =====================================================================
        private void BuildTabPO()
        {
            // --- Thanh lọc ---
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



            // --- Grid PO list ---
            panelTop = P(tabPO, 5, 52, 0, 190, Color.White);
            panelTop.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Lbl(panelTop, "DANH SÁCH ĐƠN PO", 8, 5, 350, 20, true, Color.FromArgb(0, 120, 212));
            dgvPO = Grid(panelTop, 28, 156);
            dgvPO.SelectionChanged += DgvPO_SelectionChanged;
            dgvPO.CellFormatting += DgvPO_CellFormatting;
            BuildPOGridCols();

            // --- Panel thông tin PO (header xanh) ---
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

            // --- Panel kế hoạch TT (trái) ---
            panelSched = P(tabPO, 5, 317, 0, 0, Color.White);
            panelSched.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom;
            Lbl(panelSched, "📅  KẾ HOẠCH THANH TOÁN", 8, 5, 300, 20, true, Color.FromArgb(0, 120, 212));

            bool canEdit = AppSession.CanEdit("PO") || AppSession.CanCreate("PO");
            if (canEdit)
            {
                var bAdd = Btn("+ Thêm đợt", Color.FromArgb(40, 167, 69), 8, 28, 100, 26);
                var bDel = Btn("Xóa", Color.FromArgb(220, 53, 69), 114, 28, 65, 26);
                var bSave = Btn("💾 Lưu", Color.FromArgb(0, 120, 212), 185, 28, 80, 26);
                bAdd.Click += BtnAddSched_Click;
                bDel.Click += BtnDelSched_Click;
                bSave.Click += BtnSaveSched_Click;
                panelSched.Controls.AddRange(new Control[] { bAdd, bDel, bSave });
            }

            dgvSchedule = Grid(panelSched, 60, 0);
            dgvSchedule.SelectionChanged += (s, e) =>
            {
                if (dgvSchedule.SelectedRows.Count > 0)
                    _selectedSchedID = Convert.ToInt32(dgvSchedule.SelectedRows[0].Cells["S_ID"].Value ?? 0);
            };
            dgvSchedule.CellFormatting += DgvSched_CellFormatting;
            BuildSchedCols();

            // --- Panel lịch sử TT (phải) ---
            panelHist = P(tabPO, 0, 317, 0, 0, Color.White);
            panelHist.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            Lbl(panelHist, "💰  LỊCH SỬ THANH TOÁN THỰC TẾ", 8, 5, 350, 20, true, Color.FromArgb(40, 167, 69));

            if (canEdit)
            {
                var bPay = Btn("+ Ghi nhận TT", Color.FromArgb(40, 167, 69), 8, 28, 125, 26);
                var bDel = Btn("Xóa", Color.FromArgb(220, 53, 69), 139, 28, 65, 26);
                bPay.Click += BtnAddPayment_Click;
                bDel.Click += BtnDelPayment_Click;
                panelHist.Controls.AddRange(new Control[] { bPay, bDel });
            }

            dgvHistory = Grid(panelHist, 60, 0);
            dgvHistory.SelectionChanged += (s, e) =>
            {
                if (dgvHistory.SelectedRows.Count > 0)
                    _selectedHistID = Convert.ToInt32(dgvHistory.SelectedRows[0].Cells["H_ID"].Value ?? 0);
            };
            BuildHistCols();
        }

        private void BuildPOGridCols()
        {
            dgvPO.Columns.Clear();

            // 1. Tắt tính năng tự động sinh cột dư thừa
            dgvPO.AutoGenerateColumns = false;

            // 2. Thêm DataPropertyName vào từng cột để nối đúng với dữ liệu
            dgvPO.Columns.Add(new DataGridViewTextBoxColumn { Name = "ID", DataPropertyName = "ID", Visible = false });
            dgvPO.Columns.Add(new DataGridViewTextBoxColumn { Name = "PO_No", DataPropertyName = "PO_No", HeaderText = "PO No", Width = 110, ReadOnly = true });
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
            dgvSchedule.Columns.Add(new DataGridViewTextBoxColumn { Name = "Due_Date", HeaderText = "Đến hạn", Width = 90 });
            dgvSchedule.Columns.Add(new DataGridViewTextBoxColumn { Name = "Delivery_Ref", HeaderText = "Lô hàng", Width = 90 });
            dgvSchedule.Columns.Add(new DataGridViewTextBoxColumn { Name = "Description", HeaderText = "Điều kiện", FillWeight = 100 });
            var cboStatus = new DataGridViewComboBoxColumn { Name = "S_Status", HeaderText = "Trạng thái", Width = 100, FlatStyle = FlatStyle.Flat };
            cboStatus.Items.AddRange(new[] { "Chưa TT", "Một phần", "Đã TT đủ" });
            dgvSchedule.Columns.Add(cboStatus);
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

        // =====================================================================
        //  TAB 2 — BÁO CÁO CÔNG NỢ NCC
        // =====================================================================
        private void BuildTabDebt()
        {
            // --- Filter bar ---
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

            // --- Summary Cards ---
            var pCards = P(tabDebt, 5, 55, 0, 72, Color.White);
            pCards.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            lblSumValue = Card(pCards, 10, "Tổng giá trị PO", Color.FromArgb(0, 120, 212));
            lblSumPaid = Card(pCards, 225, "Đã thanh toán", Color.FromArgb(40, 167, 69));
            lblSumDebt = Card(pCards, 440, "Còn nợ", Color.FromArgb(255, 140, 0));
            lblSumOverdue = Card(pCards, 655, "Quá hạn (PO)", Color.FromArgb(220, 53, 69));

            // --- Grid NCC (trái) ---
            var pNCC = P(tabDebt, 5, 132, 380, 0, Color.White);
            pNCC.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom;
            Lbl(pNCC, "TỔNG HỢP THEO NHÀ CUNG CẤP", 8, 5, 360, 20, true, Color.FromArgb(0, 120, 212));
            dgvDebtSupp = Grid(pNCC, 28, 0);
            dgvDebtSupp.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom | AnchorStyles.Right;
            dgvDebtSupp.SelectionChanged += DgvDebtSupp_SelectionChanged;
            dgvDebtSupp.CellFormatting += DgvDebtSupp_CellFormatting;
            BuildDebtSuppCols();

            // --- Grid chi tiết (phải) ---
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

        // =====================================================================
        //  LOAD / BIND DATA
        // =====================================================================
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
        }

        private void LoadPOSummary()
        {
            try
            {
                _poSummaries = _svc.GetPOSummaries();
                FilterAndBind(); // Đã mở comment để tự áp dụng lại bộ lọc đang có
            }
            catch (Exception ex)
            {
                Err(ex.Message);
            }
        }

        private void FilterAndBind()
        {
            string kw = txtSearchPO.Text.Trim();
            string status = cboStatusFilter.SelectedItem?.ToString() ?? "Tất cả";

            var list = _poSummaries;

            // 1. Lọc theo từ khóa tìm kiếm trước
            if (!string.IsNullOrEmpty(kw))
            {
                list = list.FindAll(p =>
                    (p.PONo ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                    (p.Project_Name ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                    (p.Supplier_Name ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase));
            }

            // 2. Khởi tạo danh sách hiển thị và TÍNH TOÁN LẠI các con số thực tế
            // 2. Khởi tạo danh sách hiển thị và TÍNH TOÁN LẠI các con số thực tế
            var displayList = list.ConvertAll(p =>
            {
                // --- Logic tính toán số dư và % (như đã làm ở bước trước) ---
                decimal totalPO = p.Total_PO_Amount;
                decimal totalPaid = p.Total_Paid;
                decimal remain = totalPO - totalPaid;
                if (remain < 0) remain = 0;

                decimal pct = totalPO > 0 ? (totalPaid / totalPO) * 100 : 0;
                if (pct > 100) pct = 100;

                string realStatus = "Chưa TT";
                if (totalPaid >= totalPO && totalPO > 0) realStatus = "Đã TT đủ";
                else if (totalPaid > 0) realStatus = "Một phần";

                // --- LOGIC MỚI: Kiểm tra xem PO có được tạo trong 3 ngày gần nhất không ---
                bool isNew = p.PO_Date.HasValue && (DateTime.Now - p.PO_Date.Value).TotalDays <= 3;
                string poDisplayObj = isNew ? $"🔥 {p.PONo} (Mới)" : p.PONo;

                return new
                {
                    ID = p.PO_ID,
                    PO_No = poDisplayObj,                // Hiển thị thêm chữ (Mới) nếu <= 3 ngày
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
                    Is_Overdue = p.Is_Overdue
                };
            });

            // 3. Lọc theo Combobox Trạng thái (áp dụng trên trạng thái THỰC TẾ vừa tính)
            if (status == "⚠ Quá hạn")
                displayList = displayList.FindAll(p => p.Is_Overdue);
            else if (status != "Tất cả")
                displayList = displayList.FindAll(p => p.TT_Status == status);

            // 4. Đổ dữ liệu lên Grid
            dgvPO.DataSource = displayList;
        }

        private void LoadSchedHist()
        {
            if (_selectedPO_ID == 0) return;
            try
            {
                _schedules = _svc.GetSchedules(_selectedPO_ID);
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

        // =====================================================================
        //  EVENTS — Tab PO
        // =====================================================================
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
            int i = dgvSchedule.Rows.Add();
            var r = dgvSchedule.Rows[i];
            r.Cells["S_ID"].Value = 0;
            r.Cells["Dot_TT"].Value = _schedules.Count + 1;
            r.Cells["Pay_Method"].Value = "Full";
            r.Cells["Payment_Type"].Value = "Chuyển khoản";
            r.Cells["Percent_TT"].Value = 0;
            r.Cells["Amount_Plan"].Value = "0";
            r.Cells["S_Status"].Value = "Chưa TT";
            dgvSchedule.CurrentCell = dgvSchedule.Rows[i].Cells["Payment_Type"];
            dgvSchedule.BeginEdit(true);
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

        private void BtnSaveSched_Click(object sender, EventArgs e)
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
                MessageBox.Show($"✅ Đã lưu {saved} đợt thanh toán!", "Thành công",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadSchedHist(); LoadPOSummary();
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

                // Bind grid NCC
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

                // Tiêu đề
                ws.Cells[1, 1].Value = "BÁO CÁO CÔNG NỢ NHÀ CUNG CẤP";
                ws.Cells[1, 1, 1, 9].Merge = true;
                ws.Cells[1, 1].Style.Font.Size = 14;
                ws.Cells[1, 1].Style.Font.Bold = true;
                ws.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[2, 1].Value = $"Kỳ: {dtpFrom.Value:dd/MM/yyyy} — {dtpTo.Value:dd/MM/yyyy}";
                ws.Cells[2, 1, 2, 9].Merge = true;

                // Header row
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

                // Data rows
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
                MessageBox.Show("✅ Xuất Excel thành công!", "Thành công",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = sfd.FileName,
                    UseShellExecute = true
                });
            }
            catch (Exception ex) { Err(ex.Message); }
        }

        // =====================================================================
        //  RESIZE
        // =====================================================================
        private void ResizeAll()
        {
            try
            {
                int w = tabPO.ClientSize.Width;
                int h = tabPO.ClientSize.Height;

                if (panelTop != null) panelTop.Width = w - 10;
                if (panelInfo != null) { panelInfo.Width = w - 10; lblPOStatus.Left = panelInfo.Width - 205; }
                if (panelSched != null)
                {
                    panelSched.Width = w / 2 - 8;
                    panelSched.Height = h - 322;
                    dgvSchedule.Height = panelSched.Height - 65;
                }
                if (panelHist != null)
                {
                    panelHist.Left = w / 2 + 3;
                    panelHist.Width = w / 2 - 8;
                    panelHist.Height = h - 322;
                    dgvHistory.Height = panelHist.Height - 65;
                }
            }
            catch { }
        }

        // =====================================================================
        //  UI HELPERS
        // =====================================================================
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
    }

    // =====================================================================
    //  DIALOG GHI NHẬN THANH TOÁN
    // =====================================================================
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
}