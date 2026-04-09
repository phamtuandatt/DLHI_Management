using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using MPR_Managerment.Models;
using MPR_Managerment.Services;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;

namespace MPR_Managerment.Forms
{
    public partial class frmWarehouse : Form
    {
        private WarehouseService _service = new WarehouseService();
        private POService _poService = new POService();
        private WarehouseLocationService _warehouseService = new WarehouseLocationService();
        private string _currentUser = "Admin";

        private TabControl tabMain;
        private TabPage tabImport, tabExport, tabStock;

        // ===== NHẬP KHO =====
        private DataGridView dgvImport, dgvImportQueue;
        private TextBox txtSearchImport, txtItemName, txtMaterial;
        private TextBox txtSize, txtUnit, txtIDCode, txtMTRno, txtHeatno;
        private TextBox txtProjectCode, txtWorkorderNo, txtLocation, txtNotesImport;
        private DateTimePicker dtpImportDate;
        private NumericUpDown nudQtyImport, nudWeightImport;
        private Button btnSaveImport, btnDeleteImport, btnDeleteImportItem, btnClearImport, btnRemoveQueue;
        private Label lblImportStatus, lblQueueStatus, lblCurrentBatch;
        private Panel panelImportForm, panelImportList, panelImportQueue;
        private ComboBox cboPOFilter, cboProjectImportFilter;
        private List<WarehouseImport> _imports = new List<WarehouseImport>();
        private List<WarehouseImport> _importQueue = new List<WarehouseImport>();
        private int _selectedImportID = 0;
        private int _pendingPO_ID = 0;
        private string _currentBatchNo = "";

        // ===== XUẤT KHO =====
        private DataGridView dgvExport, dgvStockForExport;
        private TextBox txtSearchExport, txtExportNo, txtExportTo, txtPurpose, txtNotesExport;
        private TextBox txtExportProjectCode, txtExportWorkorderNo;
        private DateTimePicker dtpExportDate, dtpExportFrom, dtpExportTo;
        private NumericUpDown nudQtyExport, nudWeightExport;
        private Button btnSaveExport, btnDeleteExport;
        private Label lblExportStatus, lblStockInfo, lblExportHistoryStatus;
        private Panel panelExportForm, panelExportList, panelStockSelect;
        private ComboBox cboProjectExportFilter, cboPOExportFilter, cboWarehouseExport, cboExportDateRange;
        private List<WarehouseExport> _exports = new List<WarehouseExport>();
        private int _selectedExportID = 0;
        private int _selectedStockImportID = 0;
        private decimal _currentStockQty = 0;

        // ===== TỒN KHO =====
        private DataGridView dgvStock;
        private TextBox txtSearchStock;
        private ComboBox cboProjectFilter;
        private Label lblStockTotal, lblStockQty, lblStockWeight;
        private Panel panelStockSummary;

        public frmWarehouse()
        {
            InitializeComponent();
            BuildUI();
            this.Resize += FrmWarehouse_Resize;
            this.Load += FrmWarehouse_Load;
        }

        private void FrmWarehouse_Load(object sender, EventArgs e) => LoadAll();

        // ===== TẠO MÃ TỰ ĐỘNG =====
        private string GenerateImportNo(string poNo)
        {
            try
            {
                string baseNo = $"PNK-{poNo}";
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var cmd = new SqlCommand(
                        "SELECT COUNT(DISTINCT Import_No) FROM Warehouse_Import WHERE Import_No LIKE @base", conn);
                    cmd.Parameters.AddWithValue("@base", baseNo + "%");
                    int count = Convert.ToInt32(cmd.ExecuteScalar());
                    var uniqueQ = new HashSet<string>();
                    foreach (var q in _importQueue)
                        if (q.Import_No.StartsWith(baseNo)) uniqueQ.Add(q.Import_No);
                    int total = count + uniqueQ.Count;
                    return total == 0 ? baseNo : $"{baseNo}_{total + 1}";
                }
            }
            catch { return $"PNK-{poNo}-{DateTime.Now:ddMMHHmm}"; }
        }

        private string GenerateExportNo(string projectCode)
        {
            try
            {
                string prefix = $"PXK-{projectCode}-";
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var cmd = new SqlCommand(
                        "SELECT COUNT(*) FROM Warehouse_Export WHERE Export_No LIKE @prefix", conn);
                    cmd.Parameters.AddWithValue("@prefix", prefix + "%");
                    int count = Convert.ToInt32(cmd.ExecuteScalar());
                    return $"{prefix}{count + 1:D3}";
                }
            }
            catch { return $"PXK-{projectCode}-001"; }
        }

        private void BuildUI()
        {
            this.Text = "Quản lý Xuất Nhập Kho";
            this.BackColor = Color.FromArgb(245, 245, 245);

            tabMain = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 10),
                Padding = new Point(20, 5)
            };
            tabImport = new TabPage("📥  Nhập kho")
            {
                AutoScroll = true,
            };
            tabExport = new TabPage("📤  Xuất kho");
            tabStock = new TabPage("📦  Tồn kho");
            tabMain.TabPages.Add(tabImport);
            tabMain.TabPages.Add(tabExport);
            tabMain.TabPages.Add(tabStock);
            this.Controls.Add(tabMain);

            BuildImportTab();
            BuildExportTab();
            BuildStockTab();
        }

        // ==================== NHẬP KHO ====================
        private void BuildImportTab()
        {
            tabImport.BackColor = Color.FromArgb(245, 245, 245);

            var panelFilter = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(1200, 45),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            tabImport.Controls.Add(panelFilter);

            panelFilter.Controls.Add(new Label { Text = "Dự án:", Location = new Point(8, 12), Size = new Size(48, 20), Font = new Font("Segoe UI", 9) });
            cboProjectImportFilter = new ComboBox { Location = new Point(58, 9), Size = new Size(160, 25), Font = new Font("Segoe UI", 9), DropDownStyle = ComboBoxStyle.DropDownList };
            cboProjectImportFilter.Items.Add("Tất cả dự án");
            cboProjectImportFilter.SelectedIndex = 0;
            cboProjectImportFilter.SelectedIndexChanged += CboProjectImport_Changed;
            panelFilter.Controls.Add(cboProjectImportFilter);

            panelFilter.Controls.Add(new Label { Text = "PO No:", Location = new Point(228, 12), Size = new Size(48, 20), Font = new Font("Segoe UI", 9) });
            cboPOFilter = new ComboBox { Location = new Point(278, 9), Size = new Size(160, 25), Font = new Font("Segoe UI", 9), DropDownStyle = ComboBoxStyle.DropDownList };
            cboPOFilter.Items.Add("-- Chọn PO --");
            cboPOFilter.SelectedIndex = 0;
            cboPOFilter.SelectedIndexChanged += CboPOFilter_Changed;
            panelFilter.Controls.Add(cboPOFilter);

            panelFilter.Controls.Add(new Label { Text = "Tìm:", Location = new Point(448, 12), Size = new Size(38, 20), Font = new Font("Segoe UI", 9) });
            txtSearchImport = new TextBox { Location = new Point(488, 9), Size = new Size(175, 25), Font = new Font("Segoe UI", 9), PlaceholderText = "Tìm tên vật tư, mã nhập..." };
            panelFilter.Controls.Add(txtSearchImport);
            txtSearchImport.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) LoadImports(); };

            var btnSrch = CreateBtn("🔍 Tìm", Color.FromArgb(0, 120, 212), new Point(671, 8), 75, 28);
            btnSrch.Click += (s, e) => LoadImports();
            panelFilter.Controls.Add(btnSrch);

            var btnClearF = CreateBtn("✖ Xóa lọc", Color.FromArgb(108, 117, 125), new Point(756, 8), 90, 28);
            btnClearF.Click += (s, e) => { txtSearchImport.Text = ""; cboProjectImportFilter.SelectedIndex = 0; cboPOFilter.SelectedIndex = 0; LoadImports(); };
            panelFilter.Controls.Add(btnClearF);

            panelImportForm = new Panel
            {
                Location = new Point(10, 65),
                Size = new Size(1200, 255),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            tabImport.Controls.Add(panelImportForm);

            panelImportForm.Controls.Add(new Label { Text = "THÔNG TIN NHẬP KHO", Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212), Location = new Point(10, 8), Size = new Size(300, 25) });

            lblCurrentBatch = new Label { Text = "Mã phiếu: (chưa có)", Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.FromArgb(40, 167, 69), Location = new Point(320, 12), Size = new Size(500, 18) };
            panelImportForm.Controls.Add(lblCurrentBatch);

            int y = 38;
            AddLbl(panelImportForm, "Ngày nhập:", 10, y);
            dtpImportDate = new DateTimePicker { Location = new Point(100, y), Size = new Size(140, 25), Font = new Font("Segoe UI", 9), Format = DateTimePickerFormat.Short };
            panelImportForm.Controls.Add(dtpImportDate);

            AddLbl(panelImportForm, "Mã dự án:", 255, y);
            txtProjectCode = AddTb(panelImportForm, 335, y, 130);

            AddLbl(panelImportForm, "Workorder:", 480, y);
            txtWorkorderNo = AddTb(panelImportForm, 555, y, 150);

            AddLbl(panelImportForm, "Vị trí kho:", 720, y);
            txtLocation = AddTb(panelImportForm, 800, y, 180);

            y += 35;
            AddLbl(panelImportForm, "Tên vật tư (*):", 10, y);
            txtItemName = AddTb(panelImportForm, 110, y, 240);

            AddLbl(panelImportForm, "Vật liệu:", 365, y);
            txtMaterial = AddTb(panelImportForm, 430, y, 110);

            AddLbl(panelImportForm, "Kích thước:", 555, y);
            txtSize = AddTb(panelImportForm, 635, y, 140);

            AddLbl(panelImportForm, "ĐVT:", 790, y);
            txtUnit = AddTb(panelImportForm, 825, y, 70);

            y += 35;
            AddLbl(panelImportForm, "Số lượng NK:", 10, y);
            nudQtyImport = new NumericUpDown { Location = new Point(110, y), Size = new Size(110, 25), Font = new Font("Segoe UI", 9), Maximum = 999999, DecimalPlaces = 2 };
            panelImportForm.Controls.Add(nudQtyImport);

            AddLbl(panelImportForm, "Trọng lượng(kg):", 235, y);
            nudWeightImport = new NumericUpDown { Location = new Point(360, y), Size = new Size(110, 25), Font = new Font("Segoe UI", 9), Maximum = 9999999, DecimalPlaces = 2 };
            panelImportForm.Controls.Add(nudWeightImport);

            AddLbl(panelImportForm, "ID Code:", 490, y);
            txtIDCode = AddTb(panelImportForm, 555, y, 120);

            AddLbl(panelImportForm, "MTR No:", 690, y);
            txtMTRno = AddTb(panelImportForm, 748, y, 110);

            AddLbl(panelImportForm, "Heat No:", 873, y);
            txtHeatno = AddTb(panelImportForm, 930, y, 100);

            y += 35;
            AddLbl(panelImportForm, "Ghi chú:", 10, y);
            txtNotesImport = AddTb(panelImportForm, 110, y, 860);
            txtNotesImport.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            y += 38;
            var btnAddQ = CreateBtn("➕ Thêm vào phiếu", Color.FromArgb(255, 140, 0), new Point(10, y), 155, 32);
            btnAddQ.Click += BtnAddToQueue_Click;
            panelImportForm.Controls.Add(btnAddQ);

            btnSaveImport = CreateBtn("💾 Lưu phiếu nhập", Color.FromArgb(0, 120, 212), new Point(175, y), 155, 32);
            btnSaveImport.Click += BtnSaveImport_Click;
            panelImportForm.Controls.Add(btnSaveImport);

            btnClearImport = CreateBtn("🆕 Phiếu mới", Color.FromArgb(108, 117, 125), new Point(340, y), 120, 32);
            btnClearImport.Click += BtnNewBatch_Click;
            panelImportForm.Controls.Add(btnClearImport);

            foreach (Control c in panelImportForm.Controls)
                if (c is TextBox || c is ComboBox || c is DateTimePicker || c is NumericUpDown)
                    c.BringToFront();

            panelImportQueue = new Panel
            {
                Location = new Point(10, 330),
                Size = new Size(1200, 160),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            tabImport.Controls.Add(panelImportQueue);

            panelImportQueue.Controls.Add(new Label { Text = "DANH SÁCH VẬT TƯ TRONG PHIẾU — Double click SL/KG để sửa", Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.FromArgb(255, 140, 0), Location = new Point(10, 8), Size = new Size(550, 22) });

            btnRemoveQueue = CreateBtn("🗑 Xóa dòng chọn", Color.FromArgb(220, 53, 69), new Point(650, 5), 140, 28);
            btnRemoveQueue.Click += BtnRemoveQueue_Click;
            panelImportQueue.Controls.Add(btnRemoveQueue);

            lblQueueStatus = new Label { Location = new Point(800, 10), Size = new Size(380, 22), Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.FromArgb(255, 140, 0) };
            panelImportQueue.Controls.Add(lblQueueStatus);

            dgvImportQueue = new DataGridView
            {
                Location = new Point(10, 35),
                Size = new Size(1175, 115),
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            dgvImportQueue.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 140, 0);
            dgvImportQueue.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvImportQueue.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvImportQueue.EnableHeadersVisualStyles = false;
            dgvImportQueue.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(255, 248, 235);
            dgvImportQueue.CellDoubleClick += DgvImportQueue_CellDoubleClick;
            panelImportQueue.Controls.Add(dgvImportQueue);
            BuildQueueColumns();

            panelImportList = new Panel
            {
                Location = new Point(10, 500),
                Size = new Size(1200, 200),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            tabImport.Controls.Add(panelImportList);

            panelImportList.Controls.Add(new Label { Text = "LỊCH SỬ NHẬP KHO — Chọn dòng để xóa item nhập sai", Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212), Location = new Point(10, 8), Size = new Size(500, 25) });

            btnDeleteImportItem = CreateBtn("🗑 Xóa item đã chọn", Color.FromArgb(220, 53, 69), new Point(10, 38), 160, 28);
            btnDeleteImportItem.Click += BtnDeleteImportItem_Click;
            panelImportList.Controls.Add(btnDeleteImportItem);

            btnDeleteImport = CreateBtn("🗑 Xóa cả phiếu", Color.FromArgb(180, 30, 30), new Point(180, 38), 140, 28);
            btnDeleteImport.Click += BtnDeleteImport_Click;
            panelImportList.Controls.Add(btnDeleteImport);

            lblImportStatus = new Label { Location = new Point(335, 42), Size = new Size(400, 22), Font = new Font("Segoe UI", 9), ForeColor = Color.Gray };
            panelImportList.Controls.Add(lblImportStatus);

            dgvImport = BuildGrid(panelImportList, 72, 120);
            dgvImport.SelectionChanged += DgvImport_SelectionChanged;
            dgvImport.CellFormatting += DgvImport_CellFormatting;
        }

        private void BuildQueueColumns()
        {
            dgvImportQueue.Columns.Clear();
            dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "QIdx", HeaderText = "#", Width = 35, ReadOnly = true });
            dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "Item_Name", HeaderText = "Tên vật tư", Width = 220, ReadOnly = true });
            dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "Material", HeaderText = "Vật liệu", Width = 90, ReadOnly = true });
            dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "Size", HeaderText = "Kích thước", Width = 110, ReadOnly = true });
            dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "UNIT", HeaderText = "ĐVT", Width = 55, ReadOnly = true });
            dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "Qty_Import", HeaderText = "SL nhập", Width = 80 });
            dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "Weight_kg", HeaderText = "KG", Width = 75 });
            dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "ID_Code", HeaderText = "ID Code", Width = 100, ReadOnly = true });
            dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ma_Phieu", HeaderText = "Mã phiếu", Width = 160, ReadOnly = true });
        }

        // ==================== XUẤT KHO ====================
        private void BuildExportTab()
        {
            tabExport.BackColor = Color.FromArgb(245, 245, 245);

            panelStockSelect = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(1200, 230),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            tabExport.Controls.Add(panelStockSelect);

            panelStockSelect.Controls.Add(new Label { Text = "CHỌN VẬT TƯ TỪ KHO ĐỂ XUẤT — click vào dòng muốn xuất", Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212), Location = new Point(10, 8), Size = new Size(700, 22) });

            panelStockSelect.Controls.Add(new Label { Text = "Dự án:", Location = new Point(10, 36), Size = new Size(50, 20), Font = new Font("Segoe UI", 9) });
            cboProjectExportFilter = new ComboBox { Location = new Point(62, 33), Size = new Size(160, 25), Font = new Font("Segoe UI", 9), DropDownStyle = ComboBoxStyle.DropDownList };
            cboProjectExportFilter.Items.Add("Tất cả dự án");
            cboProjectExportFilter.SelectedIndex = 0;
            cboProjectExportFilter.SelectedIndexChanged += CboProjectExport_Changed;
            panelStockSelect.Controls.Add(cboProjectExportFilter);

            panelStockSelect.Controls.Add(new Label { Text = "PO No:", Location = new Point(232, 36), Size = new Size(48, 20), Font = new Font("Segoe UI", 9) });
            cboPOExportFilter = new ComboBox { Location = new Point(282, 33), Size = new Size(160, 25), Font = new Font("Segoe UI", 9), DropDownStyle = ComboBoxStyle.DropDownList };
            cboPOExportFilter.Items.Add("-- Chọn PO --");
            cboPOExportFilter.SelectedIndex = 0;
            cboPOExportFilter.SelectedIndexChanged += CboPOExport_Changed;
            panelStockSelect.Controls.Add(cboPOExportFilter);

            var btnClearExFilter = CreateBtn("✖ Xóa lọc", Color.FromArgb(108, 117, 125), new Point(452, 32), 90, 28);
            btnClearExFilter.Click += (s, e) => { cboProjectExportFilter.SelectedIndex = 0; cboPOExportFilter.SelectedIndex = 0; LoadStockForExport(); };
            panelStockSelect.Controls.Add(btnClearExFilter);

            lblStockInfo = new Label { Location = new Point(10, 65), Size = new Size(1160, 22), Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.FromArgb(40, 167, 69) };
            panelStockSelect.Controls.Add(lblStockInfo);

            dgvStockForExport = BuildGrid(panelStockSelect, 92, 128);
            dgvStockForExport.SelectionChanged += DgvStockForExport_SelectionChanged;

            // Panel form xuất
            panelExportForm = new Panel
            {
                Location = new Point(10, 250),
                Size = new Size(1200, 220),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            tabExport.Controls.Add(panelExportForm);

            panelExportForm.Controls.Add(new Label { Text = "THÔNG TIN XUẤT KHO", Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.FromArgb(220, 53, 69), Location = new Point(10, 8), Size = new Size(300, 25) });

            int y = 40;
            AddLbl(panelExportForm, "Mã xuất:", 10, y);
            txtExportNo = AddTb(panelExportForm, 90, y, 180);
            txtExportNo.BackColor = Color.FromArgb(255, 245, 245);
            txtExportNo.ReadOnly = true;

            AddLbl(panelExportForm, "Ngày xuất:", 285, y);
            dtpExportDate = new DateTimePicker { Location = new Point(360, y), Size = new Size(140, 25), Font = new Font("Segoe UI", 9), Format = DateTimePickerFormat.Short };
            panelExportForm.Controls.Add(dtpExportDate);

            AddLbl(panelExportForm, "Mã dự án:", 515, y);
            txtExportProjectCode = AddTb(panelExportForm, 590, y, 130);

            AddLbl(panelExportForm, "Workorder:", 735, y);
            txtExportWorkorderNo = AddTb(panelExportForm, 813, y, 150);

            y += 35;
            AddLbl(panelExportForm, "Số lượng XK (*):", 10, y);
            nudQtyExport = new NumericUpDown { Location = new Point(130, y), Size = new Size(110, 25), Font = new Font("Segoe UI", 9), Maximum = 999999, DecimalPlaces = 2 };
            panelExportForm.Controls.Add(nudQtyExport);

            AddLbl(panelExportForm, "Trọng lượng(kg):", 260, y);
            nudWeightExport = new NumericUpDown { Location = new Point(380, y), Size = new Size(110, 25), Font = new Font("Segoe UI", 9), Maximum = 9999999, DecimalPlaces = 2 };
            panelExportForm.Controls.Add(nudWeightExport);

            AddLbl(panelExportForm, "Xuất cho:", 510, y);
            txtExportTo = AddTb(panelExportForm, 575, y, 200);

            AddLbl(panelExportForm, "Mục đích:", 790, y);
            txtPurpose = AddTb(panelExportForm, 855, y, 200);

            y += 35;
            AddLbl(panelExportForm, "Kho xuất (*):", 10, y);
            cboWarehouseExport = new ComboBox { Location = new Point(110, y), Size = new Size(280, 25), Font = new Font("Segoe UI", 9), DropDownStyle = ComboBoxStyle.DropDownList };
            cboWarehouseExport.SelectedIndexChanged += (s, e) =>
            {
                try
                {
                    int id = GetWarehouseExportID();
                    cboWarehouseExport.BackColor = id > 0 ? Color.White : Color.FromArgb(255, 248, 220);
                }
                catch { }
            };
            panelExportForm.Controls.Add(cboWarehouseExport);

            AddLbl(panelExportForm, "Ghi chú:", 410, y);
            txtNotesExport = AddTb(panelExportForm, 470, y, 500);
            txtNotesExport.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            y += 42;
            btnSaveExport = CreateBtn("💾 Lưu xuất kho", Color.FromArgb(220, 53, 69), new Point(10, y), 140, 32);
            btnSaveExport.Click += BtnSaveExport_Click;
            panelExportForm.Controls.Add(btnSaveExport);

            var btnClearEx = CreateBtn("🔄 Xóa form", Color.FromArgb(108, 117, 125), new Point(160, y), 110, 32);
            btnClearEx.Click += (s, e) => ClearExportForm();
            panelExportForm.Controls.Add(btnClearEx);

            lblExportStatus = new Label { Location = new Point(285, y + 8), Size = new Size(400, 22), Font = new Font("Segoe UI", 9), ForeColor = Color.Gray };
            panelExportForm.Controls.Add(lblExportStatus);

            foreach (Control c in panelExportForm.Controls)
                if (c is TextBox || c is ComboBox || c is DateTimePicker || c is NumericUpDown)
                    c.BringToFront();

            // ===== PANEL LỊCH SỬ XUẤT KHO =====
            panelExportList = new Panel
            {
                Location = new Point(10, 480),
                Size = new Size(1200, 200),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            tabExport.Controls.Add(panelExportList);

            panelExportList.Controls.Add(new Label { Text = "LỊCH SỬ XUẤT KHO", Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.FromArgb(220, 53, 69), Location = new Point(10, 8), Size = new Size(200, 25) });

            // Bộ lọc thời gian
            panelExportList.Controls.Add(new Label { Text = "Từ ngày:", Location = new Point(220, 12), Size = new Size(58, 20), Font = new Font("Segoe UI", 9) });
            dtpExportFrom = new DateTimePicker { Location = new Point(280, 9), Size = new Size(120, 25), Font = new Font("Segoe UI", 9), Format = DateTimePickerFormat.Short, Value = DateTime.Today.AddMonths(-1) };
            panelExportList.Controls.Add(dtpExportFrom);

            panelExportList.Controls.Add(new Label { Text = "Đến ngày:", Location = new Point(410, 12), Size = new Size(65, 20), Font = new Font("Segoe UI", 9) });
            dtpExportTo = new DateTimePicker { Location = new Point(477, 9), Size = new Size(120, 25), Font = new Font("Segoe UI", 9), Format = DateTimePickerFormat.Short, Value = DateTime.Today };
            panelExportList.Controls.Add(dtpExportTo);

            cboExportDateRange = new ComboBox { Location = new Point(607, 9), Size = new Size(140, 25), Font = new Font("Segoe UI", 9), DropDownStyle = ComboBoxStyle.DropDownList };
            cboExportDateRange.Items.AddRange(new[] { "-- Chọn nhanh --", "Hôm nay", "7 ngày qua", "Tháng này", "Tháng trước", "3 tháng qua", "Năm nay", "Tất cả" });
            cboExportDateRange.SelectedIndex = 0;
            cboExportDateRange.SelectedIndexChanged += (s, e) =>
            {
                DateTime today = DateTime.Today;
                switch (cboExportDateRange.SelectedIndex)
                {
                    case 1: dtpExportFrom.Value = today; dtpExportTo.Value = today; break;
                    case 2: dtpExportFrom.Value = today.AddDays(-7); dtpExportTo.Value = today; break;
                    case 3: dtpExportFrom.Value = new DateTime(today.Year, today.Month, 1); dtpExportTo.Value = today; break;
                    case 4: dtpExportFrom.Value = new DateTime(today.Year, today.Month, 1).AddMonths(-1); dtpExportTo.Value = new DateTime(today.Year, today.Month, 1).AddDays(-1); break;
                    case 5: dtpExportFrom.Value = today.AddMonths(-3); dtpExportTo.Value = today; break;
                    case 6: dtpExportFrom.Value = new DateTime(today.Year, 1, 1); dtpExportTo.Value = today; break;
                    case 7: dtpExportFrom.Value = new DateTime(2000, 1, 1); dtpExportTo.Value = today; break;
                }
                if (cboExportDateRange.SelectedIndex > 0) LoadExports();
            };
            panelExportList.Controls.Add(cboExportDateRange);

            txtSearchExport = new TextBox { Location = new Point(757, 9), Size = new Size(160, 25), Font = new Font("Segoe UI", 9), PlaceholderText = "Tìm mã xuất, tên vật tư..." };
            panelExportList.Controls.Add(txtSearchExport);
            txtSearchExport.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) LoadExports(); };

            var btnSrchEx = CreateBtn("🔍 Tìm", Color.FromArgb(0, 120, 212), new Point(925, 8), 75, 28);
            btnSrchEx.Click += (s, e) => LoadExports();
            panelExportList.Controls.Add(btnSrchEx);

            btnDeleteExport = CreateBtn("🗑 Xóa", Color.FromArgb(220, 53, 69), new Point(1008, 8), 75, 28);
            btnDeleteExport.Click += BtnDeleteExport_Click;
            panelExportList.Controls.Add(btnDeleteExport);

            lblExportHistoryStatus = new Label { Location = new Point(10, 40), Size = new Size(700, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212) };
            panelExportList.Controls.Add(lblExportHistoryStatus);

            dgvExport = BuildGrid(panelExportList, 62, 128);
            dgvExport.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
        }

        // ==================== TỒN KHO ====================
        private void BuildStockTab()
        {
            tabStock.BackColor = Color.FromArgb(245, 245, 245);

            panelStockSummary = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(1200, 60),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            tabStock.Controls.Add(panelStockSummary);

            lblStockTotal = AddStatLbl(panelStockSummary, "Tổng mục:", "0 mục", Color.FromArgb(0, 120, 212), 10);
            lblStockQty = AddStatLbl(panelStockSummary, "Tổng SL tồn:", "0", Color.FromArgb(40, 167, 69), 250);
            lblStockWeight = AddStatLbl(panelStockSummary, "Tổng KG tồn:", "0 kg", Color.FromArgb(255, 140, 0), 490);

            int fy = 80;
            tabStock.Controls.Add(new Label { Text = "Tìm kiếm:", Location = new Point(10, fy + 3), Size = new Size(70, 20), Font = new Font("Segoe UI", 9) });
            txtSearchStock = new TextBox { Location = new Point(83, fy), Size = new Size(200, 25), Font = new Font("Segoe UI", 9), PlaceholderText = "Tìm tên, ID Code, PO No..." };
            tabStock.Controls.Add(txtSearchStock);
            txtSearchStock.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) LoadStock(); };

            tabStock.Controls.Add(new Label { Text = "Dự án:", Location = new Point(295, fy + 3), Size = new Size(50, 20), Font = new Font("Segoe UI", 9) });
            cboProjectFilter = new ComboBox { Location = new Point(347, fy), Size = new Size(180, 25), Font = new Font("Segoe UI", 9), DropDownStyle = ComboBoxStyle.DropDownList };
            cboProjectFilter.Items.Add("Tất cả dự án");
            cboProjectFilter.SelectedIndex = 0;
            cboProjectFilter.SelectedIndexChanged += (s, e) => LoadStock();
            tabStock.Controls.Add(cboProjectFilter);

            var b1 = CreateBtn("🔍 Tìm", Color.FromArgb(0, 120, 212), new Point(537, fy - 1), 80, 28);
            var b2 = CreateBtn("📦 Chỉ còn tồn", Color.FromArgb(40, 167, 69), new Point(627, fy - 1), 130, 28);
            var b3 = CreateBtn("🔄 Làm mới", Color.FromArgb(108, 117, 125), new Point(767, fy - 1), 100, 28);
            b1.Click += (s, e) => LoadStock();
            b2.Click += (s, e) => LoadStockOnly();
            b3.Click += (s, e) => LoadStock();
            tabStock.Controls.Add(b1);
            tabStock.Controls.Add(b2);
            tabStock.Controls.Add(b3);

            dgvStock = new DataGridView
            {
                Location = new Point(10, 115),
                Size = new Size(1200, 400),
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
            dgvStock.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvStock.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvStock.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvStock.EnableHeadersVisualStyles = false;
            dgvStock.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvStock.CellFormatting += DgvStock_CellFormatting;
            tabStock.Controls.Add(dgvStock);
        }

        // ===== HELPERS =====
        private Label AddStatLbl(Panel p, string title, string value, Color color, int x)
        {
            var card = new Panel { Location = new Point(x, 8), Size = new Size(220, 42), BackColor = color };
            p.Controls.Add(card);
            card.Controls.Add(new Label { Text = title, Font = new Font("Segoe UI", 8, FontStyle.Bold), ForeColor = Color.White, Location = new Point(6, 3), Size = new Size(208, 18) });
            var lbl = new Label { Text = value, Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.White, Location = new Point(6, 22), Size = new Size(208, 18) };
            card.Controls.Add(lbl);
            return lbl;
        }

        private DataGridView BuildGrid(Panel parent, int top, int height)
        {
            var dgv = new DataGridView
            {
                Location = new Point(10, top),
                Size = new Size(parent.Width - 20, height),
                ReadOnly = true,
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

        private void AddLbl(Panel p, string text, int x, int y)
        {
            var lbl = new Label { Text = text, Location = new Point(x, y + 3), Size = new Size(100, 20), Font = new Font("Segoe UI", 9), BackColor = Color.Transparent };
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
            var btn = new Button { Text = text, Location = loc, Size = new Size(w, h), BackColor = color, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Cursor = Cursors.Hand };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        // ===== LOAD FILTERS =====
        private void LoadPOFilterByProject(string projectCode)
        {
            try
            {
                var allPO = _poService.GetAll();
                if (string.IsNullOrEmpty(projectCode))
                {
                    cboPOFilter.Items.Clear();
                    cboPOFilter.Items.Add("-- Chọn PO --");
                    foreach (var po in allPO) cboPOFilter.Items.Add(po.PONo);
                    cboPOFilter.SelectedIndex = 0;
                    return;
                }
                var projects = new ProjectService().GetAll();
                var proj = projects.Find(p => p.ProjectCode == projectCode);
                List<POHead> filtered;
                if (proj != null)
                    filtered = allPO.FindAll(p =>
                        (!string.IsNullOrEmpty(proj.WorkorderNo) && (p.WorkorderNo ?? "").Equals(proj.WorkorderNo, StringComparison.OrdinalIgnoreCase)) ||
                        (!string.IsNullOrEmpty(proj.MPRCode) && (p.MPR_No ?? "").Contains(proj.MPRCode, StringComparison.OrdinalIgnoreCase)) ||
                        (!string.IsNullOrEmpty(proj.ProjectCode) && (p.WorkorderNo ?? "").Contains(proj.ProjectCode, StringComparison.OrdinalIgnoreCase)));
                else
                    filtered = allPO.FindAll(p =>
                        (p.WorkorderNo ?? "").Contains(projectCode, StringComparison.OrdinalIgnoreCase) ||
                        (p.MPR_No ?? "").Contains(projectCode, StringComparison.OrdinalIgnoreCase));

                cboPOFilter.Items.Clear();
                cboPOFilter.Items.Add("-- Chọn PO --");
                if (filtered.Count == 0) { cboPOFilter.Items.Add("(Không có PO)"); cboPOFilter.SelectedIndex = 0; return; }
                foreach (var po in filtered) cboPOFilter.Items.Add(po.PONo);
                cboPOFilter.SelectedIndex = 0;
            }
            catch (Exception ex) { MessageBox.Show("Lỗi lọc PO: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void LoadPOExportFilterByProject(string projectCode)
        {
            try
            {
                var allPO = _poService.GetAll();
                if (string.IsNullOrEmpty(projectCode))
                {
                    cboPOExportFilter.Items.Clear();
                    cboPOExportFilter.Items.Add("-- Chọn PO --");
                    foreach (var po in allPO) cboPOExportFilter.Items.Add(po.PONo);
                    cboPOExportFilter.SelectedIndex = 0;
                    return;
                }
                var projects = new ProjectService().GetAll();
                var proj = projects.Find(p => p.ProjectCode == projectCode);
                List<POHead> filtered;
                if (proj != null)
                    filtered = allPO.FindAll(p =>
                        (!string.IsNullOrEmpty(proj.WorkorderNo) && (p.WorkorderNo ?? "").Equals(proj.WorkorderNo, StringComparison.OrdinalIgnoreCase)) ||
                        (!string.IsNullOrEmpty(proj.MPRCode) && (p.MPR_No ?? "").Contains(proj.MPRCode, StringComparison.OrdinalIgnoreCase)));
                else
                    filtered = allPO.FindAll(p =>
                        (p.WorkorderNo ?? "").Contains(projectCode, StringComparison.OrdinalIgnoreCase) ||
                        (p.MPR_No ?? "").Contains(projectCode, StringComparison.OrdinalIgnoreCase));

                cboPOExportFilter.Items.Clear();
                cboPOExportFilter.Items.Add("-- Chọn PO --");
                if (filtered.Count == 0) { cboPOExportFilter.Items.Add("(Không có PO)"); cboPOExportFilter.SelectedIndex = 0; return; }
                foreach (var po in filtered) cboPOExportFilter.Items.Add(po.PONo);
                cboPOExportFilter.SelectedIndex = 0;
            }
            catch { }
        }

        private void LoadProjectImportFilter()
        {
            try { cboProjectImportFilter.Items.Clear(); cboProjectImportFilter.Items.Add("Tất cả dự án"); foreach (var p in new ProjectService().GetAll()) cboProjectImportFilter.Items.Add(p.ProjectCode); cboProjectImportFilter.SelectedIndex = 0; }
            catch { }
        }

        private void LoadProjectExportFilter()
        {
            try { cboProjectExportFilter.Items.Clear(); cboProjectExportFilter.Items.Add("Tất cả dự án"); foreach (var p in new ProjectService().GetAll()) cboProjectExportFilter.Items.Add(p.ProjectCode); cboProjectExportFilter.SelectedIndex = 0; }
            catch { }
        }

        private void LoadProjectFilter()
        {
            try { cboProjectFilter.Items.Clear(); cboProjectFilter.Items.Add("Tất cả dự án"); foreach (var p in new ProjectService().GetAll()) cboProjectFilter.Items.Add(p.ProjectCode); cboProjectFilter.SelectedIndex = 0; }
            catch { }
        }

        private int GetWarehouseExportID()
        {
            try
            {
                if (cboWarehouseExport?.DataSource == null) return 0;
                if (cboWarehouseExport.SelectedItem == null) return 0;
                var row = cboWarehouseExport.SelectedItem as System.Data.DataRowView;
                if (row == null) return 0;
                return Convert.ToInt32(row["ID"]);
            }
            catch { return 0; }
        }

        private void LoadWarehouseExportCombo(string projectCode = "")
        {
            try
            {
                var dt = _warehouseService.GetForCombo(projectCode);
                cboWarehouseExport.DataSource = dt;
                cboWarehouseExport.DisplayMember = "Name";
                cboWarehouseExport.ValueMember = "ID";
                cboWarehouseExport.SelectedIndex = 0;
            }
            catch { }
        }

        // ===== RESIZE =====
        private void FrmWarehouse_Resize(object sender, EventArgs e)
        {
            try
            {
                int wI = tabImport.ClientSize.Width - 20;
                int wE = tabExport.ClientSize.Width - 20;
                int wS = tabStock.ClientSize.Width - 20;
                int hI = tabImport.ClientSize.Height;
                int hE = tabExport.ClientSize.Height;
                int hS = tabStock.ClientSize.Height;

                foreach (Control c in tabImport.Controls)
                    if (c is Panel p) p.Width = wI;
                if (dgvImport != null) dgvImport.Width = wI - 20;
                if (dgvImportQueue != null) dgvImportQueue.Width = wI - 20;
                if (txtNotesImport != null && panelImportForm != null)
                    txtNotesImport.Width = panelImportForm.Width - txtNotesImport.Left - 20;
                if (panelImportList != null)
                {
                    panelImportList.Height = hI - panelImportList.Top - 10;
                    if (dgvImport != null) dgvImport.Height = panelImportList.Height - 82;
                }

                if (panelStockSelect != null) panelStockSelect.Width = wE;
                if (panelExportForm != null) panelExportForm.Width = wE;
                if (dgvExport != null)
                {
                    dgvExport.Width = wE - 20;
                    dgvExport.Height = (panelExportList?.Height ?? 250) - 72;
                }
                if (dgvStockForExport != null) dgvStockForExport.Width = wE - 20;
                if (dgvExport != null)
                {
                    dgvExport.Width = wE - 20;
                    dgvExport.Height = (panelExportList?.Height ?? 220) - 75;
                }
                if (panelStockSummary != null) panelStockSummary.Width = wS;
                if (dgvStock != null) { dgvStock.Width = wS; dgvStock.Height = hS - 125; }
                if (txtNotesExport != null && panelExportForm != null)
                    txtNotesExport.Width = panelExportForm.Width - txtNotesExport.Left - 20;
            }
            catch { }
        }

        // ===== LOAD =====
        private void LoadAll()
        {
            LoadProjectImportFilter();
            LoadProjectExportFilter();
            LoadProjectFilter();
            LoadPOFilterByProject("");
            LoadPOExportFilterByProject("");
            LoadWarehouseExportCombo();
            LoadImports();
            LoadExports();
            LoadStock();
            LoadStockForExport();

            // Đảm bảo dgvExport hiển thị đúng
            if (dgvExport != null && panelExportList != null)
            {
                dgvExport.Width = panelExportList.Width - 20;
                dgvExport.Height = panelExportList.Height - 72;
                dgvExport.BringToFront();
            }
        }

        private void LoadImports()
        {
            try
            {
                if (dgvImport == null) return;
                string kw = txtSearchImport?.Text.Trim() ?? "";
                string poNo = (cboPOFilter != null && cboPOFilter.SelectedIndex > 0) ? cboPOFilter.SelectedItem.ToString() : "";
                string project = (cboProjectImportFilter != null && cboProjectImportFilter.SelectedIndex > 0) ? cboProjectImportFilter.SelectedItem.ToString() : "";

                var all = _service.GetAllImports();
                if (!string.IsNullOrEmpty(kw))
                    all = all.FindAll(i =>
                        i.Item_Name.Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                        i.Import_No.Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                        i.ID_Code.Contains(kw, StringComparison.OrdinalIgnoreCase));
                if (!string.IsNullOrEmpty(poNo))
                {
                    var po = _poService.GetAll().Find(p => p.PONo == poNo);
                    all = po != null ? all.FindAll(i => i.PO_ID == po.PO_ID) : new List<WarehouseImport>();
                }
                if (!string.IsNullOrEmpty(project))
                    all = all.FindAll(i => i.Project_Code == project);

                _imports = all;
                dgvImport.DataSource = _imports.ConvertAll(i => new
                {
                    ID = i.Import_ID,
                    Ma_Phieu = i.Import_No,
                    Ngay_Nhap = i.Import_Date.HasValue ? i.Import_Date.Value.ToString("dd/MM/yyyy") : "",
                    Ten_Vat_Tu = i.Item_Name,
                    Vat_Lieu = i.Material,
                    Kich_Thuoc = i.Size,
                    DVT = i.UNIT,
                    SL_Nhap = i.Qty_Import,
                    KG_Nhap = i.Weight_kg,
                    ID_Code = i.ID_Code,
                    MTR_No = i.MTRno,
                    Ma_DA = i.Project_Code,
                    Vi_Tri = i.Location
                });
                if (dgvImport.Columns.Contains("ID")) dgvImport.Columns["ID"].Visible = false;
                if (lblImportStatus != null) lblImportStatus.Text = $"Tổng: {_imports.Count} bản ghi";
            }
            catch (Exception ex) 
            { 
                MessageBox.Show("Lỗi tải nhập kho: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void LoadExports()
        {
            try
            {
                if (dgvExport == null) return;

                DateTime fromDate = dtpExportFrom?.Value.Date ?? DateTime.Today.AddMonths(-1);
                DateTime toDate = (dtpExportTo?.Value.Date ?? DateTime.Today).AddDays(1);
                string kw = txtSearchExport?.Text.Trim() ?? "";

                var all = _service.GetAllExports();

                // Lọc theo khoảng ngày
                all = all.FindAll(e =>
                    e.Export_Date.HasValue &&
                    e.Export_Date.Value.Date >= fromDate &&
                    e.Export_Date.Value.Date < toDate);

                // Lọc theo từ khóa
                if (!string.IsNullOrEmpty(kw))
                    all = all.FindAll(e =>
                        (e.Export_No ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                        (e.Item_Name ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                        (e.Export_To ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                        (e.Project_Code ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase));

                _exports = all;
                dgvExport.DataSource = _exports.ConvertAll(e => new
                {
                    ID = e.Export_ID,
                    Ma_Xuat = e.Export_No,
                    Ngay_Xuat = e.Export_Date.HasValue ? e.Export_Date.Value.ToString("dd/MM/yyyy") : "",
                    Ten_Vat_Tu = e.Item_Name,
                    DVT = e.UNIT,
                    SL_Xuat = e.Qty_Export,
                    KG_Xuat = e.Weight_kg,
                    ID_Code = e.ID_Code,
                    Kho_Xuat = e.Export_To,
                    Ma_DA = e.Project_Code,
                    Muc_Dich = e.Purpose
                });
                if (dgvExport.Columns.Contains("ID")) dgvExport.Columns["ID"].Visible = false;

                decimal totalQty = 0, totalKg = 0;
                foreach (var ex in _exports) { totalQty += ex.Qty_Export; totalKg += ex.Weight_kg; }
                if (lblExportHistoryStatus != null)
                    lblExportHistoryStatus.Text = $"📋 Tổng: {_exports.Count} phiếu  |  SL xuất: {totalQty:N2}  |  KG xuất: {totalKg:N2}";
                // Debug: force dgvExport visible và bring to front
                dgvExport.Visible = true;
                dgvExport.BringToFront();
                dgvExport.Refresh();
            }
            catch (Exception ex) { MessageBox.Show("Lỗi tải xuất kho: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void LoadStock()
        {
            try
            {
                if (dgvStock == null) return;
                string kw = txtSearchStock?.Text.Trim() ?? "";
                string project = (cboProjectFilter != null && cboProjectFilter.SelectedIndex > 0) ? cboProjectFilter.SelectedItem.ToString() : "";
                BindStockGrid(_service.GetStock(project, kw));
            }
            catch (Exception ex) { MessageBox.Show("Lỗi tải tồn kho: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void LoadStockOnly()
        {
            try { if (dgvStock != null) BindStockGrid(_service.GetStockWithRemaining(cboProjectFilter.SelectedText)); }
            catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void BindStockGrid(List<WarehouseStock> stocks)
        {
            dgvStock.DataSource = stocks.ConvertAll(s => new
            {
                Import_ID = s.Import_ID,
                Ma_Phieu = s.Import_No,
                Ngay_Nhap = s.Import_Date.HasValue ? s.Import_Date.Value.ToString("dd/MM/yyyy") : "",
                Ten_Vat_Tu = s.Item_Name,
                Vat_Lieu = s.Material,
                Kich_Thuoc = s.Size,
                DVT = s.UNIT,
                ID_Code = s.ID_Code,
                PO_No = s.PONo,
                Ma_DA = s.Project_Code,
                Vi_Tri = s.Location,
                SL_Nhap = s.Qty_Import,
                SL_Xuat = s.Qty_Exported,
                SL_Ton = s.Qty_Stock,
                KG_Nhap = s.Weight_Import,
                KG_Xuat = s.Weight_Exported,
                KG_Ton = s.Weight_Stock
            });
            if (dgvStock.Columns.Contains("Import_ID")) dgvStock.Columns["Import_ID"].Visible = false;
            decimal tQ = 0, tW = 0;
            foreach (var s in stocks) { tQ += s.Qty_Stock; tW += s.Weight_Stock; }
            if (lblStockTotal != null) lblStockTotal.Text = $"{stocks.Count} mục";
            if (lblStockQty != null) lblStockQty.Text = tQ.ToString("N2");
            if (lblStockWeight != null) lblStockWeight.Text = tW.ToString("N2") + " kg";
        }

        private void LoadStockForExport(string projectCode = "", string poNo = "")
        {
            try
            {
                if (dgvStockForExport == null) return;
                var stocks = _service.GetStockWithRemaining(cboProjectExportFilter.SelectedText);
                if (!string.IsNullOrEmpty(projectCode))
                    stocks = stocks.FindAll(s => (s.Project_Code ?? "").Equals(projectCode, StringComparison.OrdinalIgnoreCase));
                if (!string.IsNullOrEmpty(poNo))
                    stocks = stocks.FindAll(s => (s.PONo ?? "").Equals(poNo, StringComparison.OrdinalIgnoreCase));

                dgvStockForExport.DataSource = stocks.ConvertAll(s => new
                {
                    Import_ID = s.Import_ID,
                    Ten_Vat_Tu = s.Item_Name,
                    Vat_Lieu = s.Material,
                    Kich_Thuoc = s.Size,
                    DVT = s.UNIT,
                    ID_Code = s.ID_Code,
                    PO_No = s.PONo,
                    Ma_DA = s.Project_Code,
                    SL_Ton = s.Qty_Stock,
                    KG_Ton = s.Weight_Stock,
                    Vi_Tri = s.Location
                });
                if (dgvStockForExport.Columns.Contains("Import_ID"))
                    dgvStockForExport.Columns["Import_ID"].Visible = false;
            }
            catch { }
        }

        private void RefreshQueueGrid()
        {
            dgvImportQueue.Rows.Clear();
            for (int i = 0; i < _importQueue.Count; i++)
            {
                var item = _importQueue[i];
                dgvImportQueue.Rows.Add(i + 1, item.Item_Name, item.Material, item.Size, item.UNIT, item.Qty_Import, item.Weight_kg, item.ID_Code, item.Import_No);
            }
            int count = _importQueue.Count;
            if (lblQueueStatus != null) lblQueueStatus.Text = count > 0 ? $"📋 Phiếu: {_currentBatchNo}  |  {count} vật tư" : "";
            if (lblCurrentBatch != null) lblCurrentBatch.Text = string.IsNullOrEmpty(_currentBatchNo)
                ? "Mã phiếu: (chưa có — chọn PO hoặc thêm thủ công)"
                : $"✅ Mã phiếu: {_currentBatchNo}  ({count} items)";
        }

        // ===== SỰ KIỆN LỌC =====
        private void CboProjectImport_Changed(object sender, EventArgs e)
        {
            try
            {
                string project = (cboProjectImportFilter != null && cboProjectImportFilter.SelectedIndex > 0) ? cboProjectImportFilter.SelectedItem.ToString() : "";
                cboPOFilter.SelectedIndexChanged -= CboPOFilter_Changed;
                LoadPOFilterByProject(project);
                cboPOFilter.SelectedIndexChanged += CboPOFilter_Changed;
                LoadImports();
            }
            catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void CboProjectExport_Changed(object sender, EventArgs e)
        {
            try
            {
                string project = cboProjectExportFilter.SelectedIndex > 0 ? cboProjectExportFilter.SelectedItem.ToString() : "";
                cboPOExportFilter.SelectedIndexChanged -= CboPOExport_Changed;
                LoadPOExportFilterByProject(project);
                cboPOExportFilter.SelectedIndexChanged += CboPOExport_Changed;
                LoadWarehouseExportCombo(project);
                LoadStockForExport(project, "");
            }
            catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void CboPOExport_Changed(object sender, EventArgs e)
        {
            try
            {
                string project = cboProjectExportFilter.SelectedIndex > 0 ? cboProjectExportFilter.SelectedItem.ToString() : "";
                string poNo = (cboPOExportFilter.SelectedIndex > 0 && !cboPOExportFilter.SelectedItem.ToString().StartsWith("("))
                    ? cboPOExportFilter.SelectedItem.ToString() : "";
                LoadStockForExport(project, poNo);
            }
            catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void CboPOFilter_Changed(object sender, EventArgs e)
        {
            LoadImports();
            if (cboPOFilter.SelectedIndex <= 0) return;
            try
            {
                string poNo = cboPOFilter.SelectedItem.ToString();
                var po = _poService.GetAll().Find(p => p.PONo == poNo);
                if (po == null) return;
                var details = _poService.GetDetails(po.PO_ID);
                if (details.Count == 0) { MessageBox.Show("PO này chưa có chi tiết vật tư!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }

                using (var dlg = new Form())
                {
                    dlg.Text = $"Chọn vật tư nhập kho từ PO: {poNo}";
                    dlg.Size = new Size(1100, 510);
                    dlg.StartPosition = FormStartPosition.CenterParent;
                    dlg.BackColor = Color.White;

                    dlg.Controls.Add(new Label { Text = $"PO: {poNo}  —  {po.Project_Name}  —  Tick chọn vật tư, sửa SL nếu cần:", Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212), Location = new Point(10, 10), Size = new Size(900, 25) });

                    var dgv = new DataGridView { Location = new Point(10, 45), Size = new Size(1060, 350), AllowUserToAddRows = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, BackgroundColor = Color.White, BorderStyle = BorderStyle.FixedSingle, RowHeadersVisible = false, Font = new Font("Segoe UI", 9), AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill };
                    dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
                    dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                    dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    dgv.EnableHeadersVisualStyles = false;
                    dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
                    dlg.Controls.Add(dgv);

                    dgv.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Chon", HeaderText = "Chọn", Width = 50 });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "PO_Detail_ID", HeaderText = "ID", Visible = false });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "STT", HeaderText = "STT", Width = 40, ReadOnly = true });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ten_Hang", HeaderText = "Tên hàng", Width = 210, ReadOnly = true });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Vat_Lieu", HeaderText = "Vật liệu", Width = 80, ReadOnly = true });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "A_mm", HeaderText = "A(mm)", Width = 60, ReadOnly = true });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "B_mm", HeaderText = "B(mm)", Width = 60, ReadOnly = true });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "C_mm", HeaderText = "C(mm)", Width = 60, ReadOnly = true });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "DVT", HeaderText = "ĐVT", Width = 50, ReadOnly = true });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "SL_NK", HeaderText = "SL nhập", Width = 75 });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "KG", HeaderText = "KG", Width = 65, ReadOnly = true });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "MPS_No", HeaderText = "MPS No", Width = 90, ReadOnly = true });

                    foreach (var d in details)
                        dgv.Rows.Add(false, d.PO_Detail_ID, d.Item_No, d.Item_Name, d.Material, d.Asize, d.Bsize, d.Csize, d.UNIT, d.Qty_Per_Sheet, d.Weight_kg, d.MPSNo);

                    var btnAll = new Button { Text = "☑ Chọn tất cả", Location = new Point(10, 405), Size = new Size(120, 32), BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
                    btnAll.FlatAppearance.BorderSize = 0;
                    btnAll.Click += (s2, e2) => { foreach (DataGridViewRow r in dgv.Rows) r.Cells["Chon"].Value = true; };
                    dlg.Controls.Add(btnAll);

                    var btnAdd = new Button { Text = "✔ Thêm vào phiếu", Location = new Point(140, 405), Size = new Size(160, 32), BackColor = Color.FromArgb(40, 167, 69), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), DialogResult = DialogResult.OK };
                    btnAdd.FlatAppearance.BorderSize = 0;
                    dlg.Controls.Add(btnAdd);

                    var btnCan = new Button { Text = "Hủy", Location = new Point(310, 405), Size = new Size(80, 32), BackColor = Color.FromArgb(108, 117, 125), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), DialogResult = DialogResult.Cancel };
                    btnCan.FlatAppearance.BorderSize = 0;
                    dlg.Controls.Add(btnCan);
                    dlg.AcceptButton = btnAdd;
                    dlg.CancelButton = btnCan;

                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        if (string.IsNullOrEmpty(_currentBatchNo)) _currentBatchNo = GenerateImportNo(poNo);
                        int addedCount = 0;
                        foreach (DataGridViewRow row in dgv.Rows)
                        {
                            bool ticked = row.Cells["Chon"].Value != null && Convert.ToBoolean(row.Cells["Chon"].Value);
                            if (!ticked) continue;
                            int pdId = Convert.ToInt32(row.Cells["PO_Detail_ID"].Value);
                            var detail = details.Find(d => d.PO_Detail_ID == pdId);
                            if (detail == null) continue;
                            decimal qty = decimal.TryParse(row.Cells["SL_NK"].Value?.ToString(), out decimal q) ? q : detail.Qty_Per_Sheet;

                            string projectCode = "";
                            if (cboProjectImportFilter != null && cboProjectImportFilter.SelectedIndex > 0)
                                projectCode = cboProjectImportFilter.SelectedItem.ToString();
                            else
                            {
                                try { var pjs = new ProjectService().GetAll(); projectCode = pjs.Find(p => p.WorkorderNo == po.WorkorderNo)?.ProjectCode ?? po.MPR_No ?? ""; }
                                catch { projectCode = po.MPR_No ?? ""; }
                            }

                            _importQueue.Add(new WarehouseImport
                            {
                                Import_No = _currentBatchNo,
                                Import_Date = dtpImportDate.Value,
                                PO_ID = po.PO_ID,
                                PO_Detail_ID = detail.PO_Detail_ID,
                                Item_Name = detail.Item_Name ?? "",
                                Material = detail.Material ?? "",
                                Size = $"{detail.Asize}x{detail.Bsize}x{detail.Csize}",
                                UNIT = detail.UNIT ?? "",
                                Qty_Import = qty,
                                Weight_kg = detail.Weight_kg,
                                Project_Code = projectCode,
                                WorkorderNo = po.WorkorderNo ?? "",
                                Location = txtLocation.Text.Trim()
                            });
                            addedCount++;
                        }
                        RefreshQueueGrid();
                        if (addedCount > 0)
                            MessageBox.Show($"✅ Đã thêm {addedCount} vật tư vào phiếu: {_currentBatchNo}\nTổng: {_importQueue.Count} items — Nhấn 'Lưu phiếu nhập' để hoàn tất.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void DgvImportQueue_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _importQueue.Count) return;
            string colName = dgvImportQueue.Columns[e.ColumnIndex].Name;
            if (colName != "Qty_Import" && colName != "Weight_kg") return;
            var item = _importQueue[e.RowIndex];
            string fld = colName == "Qty_Import" ? "Số lượng nhập" : "Trọng lượng (kg)";
            decimal cur = colName == "Qty_Import" ? item.Qty_Import : item.Weight_kg;
            string input = Microsoft.VisualBasic.Interaction.InputBox($"Nhập {fld} mới cho:\n{item.Item_Name}", $"Sửa {fld}", cur.ToString("N2"));
            if (string.IsNullOrWhiteSpace(input)) return;
            if (!decimal.TryParse(input, out decimal newVal) || newVal < 0) { MessageBox.Show("Giá trị không hợp lệ!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (colName == "Qty_Import") item.Qty_Import = newVal; else item.Weight_kg = newVal;
            RefreshQueueGrid();
        }

        private void BtnAddToQueue_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtItemName.Text)) { MessageBox.Show("Vui lòng nhập Tên vật tư!", "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (nudQtyImport.Value <= 0) { MessageBox.Show("Vui lòng nhập Số lượng!", "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (string.IsNullOrEmpty(_currentBatchNo))
            {
                string proj = txtProjectCode.Text.Trim();
                _currentBatchNo = GenerateImportNo(string.IsNullOrEmpty(proj) ? "MANUAL" : proj);
            }
            _importQueue.Add(new WarehouseImport
            {
                Import_No = _currentBatchNo,
                Import_Date = dtpImportDate.Value,
                PO_ID = _pendingPO_ID > 0 ? _pendingPO_ID : (int?)null,
                Item_Name = txtItemName.Text.Trim(),
                Material = txtMaterial.Text.Trim(),
                Size = txtSize.Text.Trim(),
                UNIT = txtUnit.Text.Trim(),
                Qty_Import = nudQtyImport.Value,
                Weight_kg = nudWeightImport.Value,
                ID_Code = txtIDCode.Text.Trim(),
                MTRno = txtMTRno.Text.Trim(),
                Heatno = txtHeatno.Text.Trim(),
                Project_Code = txtProjectCode.Text.Trim(),
                WorkorderNo = txtWorkorderNo.Text.Trim(),
                Location = txtLocation.Text.Trim(),
                Notes = txtNotesImport.Text.Trim()
            });
            RefreshQueueGrid();
            ClearImportItemForm();
        }

        private void BtnRemoveQueue_Click(object sender, EventArgs e)
        {
            if (dgvImportQueue.SelectedRows.Count == 0) return;
            int idx = dgvImportQueue.SelectedRows[0].Index;
            if (idx >= 0 && idx < _importQueue.Count)
            {
                _importQueue.RemoveAt(idx);
                if (_importQueue.Count == 0) _currentBatchNo = "";
                RefreshQueueGrid();
            }
        }

        private void BtnNewBatch_Click(object sender, EventArgs e)
        {
            if (_importQueue.Count > 0)
                if (MessageBox.Show($"Bạn có {_importQueue.Count} items chưa lưu. Tạo phiếu mới sẽ xóa danh sách. Tiếp tục?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            _importQueue.Clear(); _currentBatchNo = ""; _pendingPO_ID = 0;
            ClearImportItemForm(); RefreshQueueGrid();
        }

        private void DgvImport_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvImport.SelectedRows.Count == 0) return;
            _selectedImportID = Convert.ToInt32(dgvImport.SelectedRows[0].Cells["ID"].Value);
        }

        private void DgvImport_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (dgvImport.Columns[e.ColumnIndex].Name == "SL_Nhap")
            { e.CellStyle.ForeColor = Color.FromArgb(0, 120, 212); e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold); }
        }

        private void DgvStockForExport_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvStockForExport.SelectedRows.Count == 0) return;
            var row = dgvStockForExport.SelectedRows[0];
            _selectedStockImportID = Convert.ToInt32(row.Cells["Import_ID"].Value);
            _currentStockQty = Convert.ToDecimal(row.Cells["SL_Ton"].Value);
            string project = row.Cells["Ma_DA"].Value?.ToString() ?? "";

            if (lblStockInfo != null)
                lblStockInfo.Text = $"✅ {row.Cells["Ten_Vat_Tu"].Value}  |  ID: {row.Cells["ID_Code"].Value}  |  Tồn: {_currentStockQty}  |  DA: {project}";

            nudQtyExport.Maximum = _currentStockQty;
            txtExportProjectCode.Text = project;
            txtExportWorkorderNo.Text = row.Cells["PO_No"].Value?.ToString() ?? "";

            if (!string.IsNullOrEmpty(project))
            {
                txtExportNo.Text = GenerateExportNo(project);
                LoadWarehouseExportCombo(project);
            }
            if (cboWarehouseExport != null && GetWarehouseExportID() == 0)
                cboWarehouseExport.BackColor = Color.FromArgb(255, 248, 220);
        }

        private void DgvExport_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvExport.SelectedRows.Count == 0) return;
            _selectedExportID = Convert.ToInt32(dgvExport.SelectedRows[0].Cells["ID"].Value);
        }

        private void DgvStock_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string col = dgvStock.Columns[e.ColumnIndex].Name;
            if (col == "SL_Ton" || col == "KG_Ton")
            {
                decimal val = e.Value != null ? Convert.ToDecimal(e.Value) : 0;
                e.CellStyle.ForeColor = val > 0 ? Color.FromArgb(40, 167, 69) : Color.FromArgb(220, 53, 69);
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
        }

        private void BtnSaveImport_Click(object sender, EventArgs e)
        {
            if (_importQueue.Count == 0) { MessageBox.Show("Danh sách phiếu đang trống!\nHãy thêm vật tư trước.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            try
            {
                int saved = 0;
                foreach (var imp in _importQueue) { imp.Import_Date = dtpImportDate.Value; _service.InsertImport(imp, _currentUser); saved++; }
                MessageBox.Show($"✅ Lưu phiếu nhập kho thành công!\nMã phiếu: {_currentBatchNo}\nSố vật tư: {saved} items", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                _importQueue.Clear(); _currentBatchNo = ""; _pendingPO_ID = 0;
                RefreshQueueGrid(); LoadAll(); ClearImportItemForm();
            }
            catch (Exception ex) { MessageBox.Show("Lỗi nhập kho: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void BtnDeleteImportItem_Click(object sender, EventArgs e)
        {
            if (_selectedImportID == 0) { MessageBox.Show("Vui lòng chọn item nhập cần xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            var imp = _imports.Find(i => i.Import_ID == _selectedImportID);
            string info = imp != null ? $"\nVật tư: {imp.Item_Name}\nMã phiếu: {imp.Import_No}" : "";
            if (MessageBox.Show($"Xóa item nhập này?{info}", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    using (var conn = DatabaseHelper.GetConnection())
                    {
                        conn.Open();
                        new SqlCommand($"DELETE FROM Warehouse_Export WHERE Import_ID = {_selectedImportID}", conn).ExecuteNonQuery();
                        new SqlCommand($"DELETE FROM Warehouse_Import WHERE Import_ID = {_selectedImportID}", conn).ExecuteNonQuery();
                    }
                    MessageBox.Show("Đã xóa item nhập kho!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    _selectedImportID = 0; LoadAll();
                }
                catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
        }

        private void BtnDeleteImport_Click(object sender, EventArgs e)
        {
            if (_selectedImportID == 0) { MessageBox.Show("Vui lòng chọn dòng trong lịch sử nhập!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            var imp = _imports.Find(i => i.Import_ID == _selectedImportID);
            string batchNo = imp?.Import_No ?? "";
            if (MessageBox.Show($"Xóa TOÀN BỘ phiếu nhập: {batchNo}?\n(Tất cả items cùng mã phiếu sẽ bị xóa)", "Xác nhận xóa cả phiếu", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                try
                {
                    using (var conn = DatabaseHelper.GetConnection())
                    {
                        conn.Open();
                        var cmd = new SqlCommand("SELECT Import_ID FROM Warehouse_Import WHERE Import_No = @no", conn);
                        cmd.Parameters.AddWithValue("@no", batchNo);
                        var ids = new List<int>();
                        using (var r = cmd.ExecuteReader()) while (r.Read()) ids.Add(Convert.ToInt32(r["Import_ID"]));
                        foreach (int id in ids)
                        {
                            new SqlCommand($"DELETE FROM Warehouse_Export WHERE Import_ID = {id}", conn).ExecuteNonQuery();
                            new SqlCommand($"DELETE FROM Warehouse_Import WHERE Import_ID = {id}", conn).ExecuteNonQuery();
                        }
                    }
                    MessageBox.Show($"Đã xóa toàn bộ phiếu: {batchNo}!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    _selectedImportID = 0; LoadAll();
                }
                catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
        }

        private void BtnSaveExport_Click(object sender, EventArgs e)
        {
            if (_selectedStockImportID == 0) { MessageBox.Show("Vui lòng chọn vật tư từ kho!", "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (nudQtyExport.Value <= 0) { MessageBox.Show("Vui lòng nhập Số lượng xuất!", "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (GetWarehouseExportID() == 0)
            {
                cboWarehouseExport.BackColor = Color.FromArgb(255, 230, 230);
                MessageBox.Show("Vui lòng chọn Kho xuất!", "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cboWarehouseExport.Focus(); return;
            }
            try
            {
                string warehouseName = "";
                int wId = GetWarehouseExportID();
                var warehouseDt = cboWarehouseExport.DataSource as System.Data.DataTable;
                if (warehouseDt != null && wId > 0)
                    foreach (System.Data.DataRow dr in warehouseDt.Rows)
                        if (Convert.ToInt32(dr["ID"]) == wId) { warehouseName = dr["Name"].ToString(); break; }

                var exp = new WarehouseExport
                {
                    Export_No = txtExportNo.Text.Trim(),
                    Export_Date = dtpExportDate.Value,
                    Import_ID = _selectedStockImportID,
                    Qty_Export = nudQtyExport.Value,
                    Weight_kg = nudWeightExport.Value,
                    Project_Code = txtExportProjectCode.Text.Trim(),
                    WorkorderNo = txtExportWorkorderNo.Text.Trim(),
                    Export_To = warehouseName,
                    Purpose = txtPurpose.Text.Trim(),
                    Notes = txtNotesExport.Text.Trim()
                };
                _service.InsertExport(exp, _currentUser);
                MessageBox.Show($"Xuất kho thành công!\nMã: {exp.Export_No}\nKho xuất: {warehouseName}", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadAll(); ClearExportForm();
            }
            catch (Exception ex) { MessageBox.Show("Lỗi xuất kho: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void BtnDeleteExport_Click(object sender, EventArgs e)
        {
            if (_selectedExportID == 0) { MessageBox.Show("Vui lòng chọn phiếu xuất cần xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (MessageBox.Show("Xóa phiếu xuất này?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try { _service.DeleteExport(_selectedExportID); MessageBox.Show("Xóa thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information); LoadAll(); ClearExportForm(); }
                catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
        }

        private void ClearImportItemForm()
        {
            txtItemName.Text = ""; txtMaterial.Text = ""; txtSize.Text = ""; txtUnit.Text = "";
            txtIDCode.Text = ""; txtMTRno.Text = ""; txtHeatno.Text = ""; txtNotesImport.Text = "";
            nudQtyImport.Value = 0; nudWeightImport.Value = 0;
        }

        private void ClearExportForm()
        {
            _selectedStockImportID = 0;
            _currentStockQty = 0;
            txtExportNo.Text = "";
            txtExportTo.Text = "";
            txtPurpose.Text = "";
            txtNotesExport.Text = "";
            txtExportProjectCode.Text = "";
            txtExportWorkorderNo.Text = "";
            nudQtyExport.Value = 0;
            nudWeightExport.Value = 0;
            dtpExportDate.Value = DateTime.Today;
            if (cboWarehouseExport != null) { cboWarehouseExport.SelectedIndex = 0; cboWarehouseExport.BackColor = Color.White; }
            if (lblStockInfo != null) lblStockInfo.Text = "";
            if (lblExportStatus != null) lblExportStatus.Text = "";
        }

        private void BuildStockTab_V2(TabPage parent)
        {
            // --- CẤU HÌNH GỐC: CHO PHÉP SCROLL TOÀN TRANG ---
            Panel mainScrollPanel = new Panel();
            mainScrollPanel.Dock = DockStyle.Fill;
            mainScrollPanel.AutoScroll = true; // Kích hoạt cuộn ngang/dọc khi thu nhỏ
            parent.Controls.Add(mainScrollPanel);

            // Dùng một container để giữ độ rộng cố định khi scroll (tránh các control bị bóp méo)
            Panel container = new Panel();
            container.Width = 1300; // Độ rộng tối thiểu để không bị nhảy layout
            container.Height = 1200; // Độ cao ước tính cho 4 phần
            container.Location = new Point(0, 0);
            mainScrollPanel.Controls.Add(container);

            panelStockSummary = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(1200, 60),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            container.Controls.Add(panelStockSummary);

            lblStockTotal = AddStatLbl(panelStockSummary, "Tổng mục:", "0 mục", Color.FromArgb(0, 120, 212), 10);
            lblStockQty = AddStatLbl(panelStockSummary, "Tổng SL tồn:", "0", Color.FromArgb(40, 167, 69), 250);
            lblStockWeight = AddStatLbl(panelStockSummary, "Tổng KG tồn:", "0 kg", Color.FromArgb(255, 140, 0), 490);

            int fy = 80;
            container.Controls.Add(new Label { Text = "Tìm kiếm:", Location = new Point(10, fy + 3), Size = new Size(70, 20), Font = new Font("Segoe UI", 9) });
            txtSearchStock = new TextBox { Location = new Point(83, fy), Size = new Size(200, 25), Font = new Font("Segoe UI", 9), PlaceholderText = "Tìm tên, ID Code, PO No..." };
            container.Controls.Add(txtSearchStock);
            txtSearchStock.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) LoadStock(); };

            container.Controls.Add(new Label { Text = "Dự án:", Location = new Point(295, fy + 3), Size = new Size(50, 20), Font = new Font("Segoe UI", 9) });
            cboProjectFilter = new ComboBox { Location = new Point(347, fy), Size = new Size(180, 25), Font = new Font("Segoe UI", 9), DropDownStyle = ComboBoxStyle.DropDownList };
            cboProjectFilter.Items.Add("Tất cả dự án");
            cboProjectFilter.SelectedIndex = 0;
            cboProjectFilter.SelectedIndexChanged += (s, e) => LoadStock();
            container.Controls.Add(cboProjectFilter);

            var b1 = CreateBtn("🔍 Tìm", Color.FromArgb(0, 120, 212), new Point(537, fy - 1), 80, 28);
            var b2 = CreateBtn("📦 Chỉ còn tồn", Color.FromArgb(40, 167, 69), new Point(627, fy - 1), 130, 28);
            var b3 = CreateBtn("🔄 Làm mới", Color.FromArgb(108, 117, 125), new Point(767, fy - 1), 100, 28);
            b1.Click += (s, e) => LoadStock();
            b2.Click += (s, e) => LoadStockOnly();
            b3.Click += (s, e) => LoadStock();
            container.Controls.Add(b1);
            container.Controls.Add(b2);
            container.Controls.Add(b3);

            dgvStock = new DataGridView
            {
                Location = new Point(10, 115),
                Size = new Size(1200, 400),
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
            dgvStock.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvStock.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvStock.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvStock.EnableHeadersVisualStyles = false;
            dgvStock.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvStock.CellFormatting += DgvStock_CellFormatting;
            container.Controls.Add(dgvStock);
        }
    }
}