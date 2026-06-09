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
using System.Runtime.InteropServices;

namespace MPR_Managerment.Forms
{
    public partial class frmPO : Form
    {
        private POService _service = new POService();
        private List<POHead> _poList = new List<POHead>();
        private List<PODetail> _details = new List<PODetail>();
        private readonly Dictionary<int, System.Windows.Forms.Timer> _emailPollers = new();
        private int _selectedPO_ID = 0;
        private string _currentUser = AppSession.CurrentUser.Full_Name ?? "Admin";

        private string _targetPoNo = "";
        private string _importMprNo = "";
        private int _importMprId = 0;
        private List<int> _filteredMprDetailIds = null;

        private DataGridView dgvPO;
        private DataGridView dgvDocPO; // Document: INV + Delivery theo PO đang chọn
        private TextBox txtSearch;
        private Button btnSearch, btnNewPO, btnDeletePO, btnClearHeader, btnExport, btnSavePO, btnUnlockSupplier;
        private Button btnSearchBySupp, btnPOStatus;
        private bool _supplierLocked = false;
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
        // Culture định dạng số theo hệ thống Windows (dấu thập phân và ngàn tùy cài đặt)
        private static readonly System.Globalization.CultureInfo _numCulture =
            System.Globalization.CultureInfo.CurrentCulture;
        private Panel panelTop, panelHeader, panelDetail;
        private Panel _mainScroll; // Panel bọc ngoài hỗ trợ scroll trên màn hình nhỏ
        private DataGridView dgvMPRFiles; // Bảng file MPR Link bên phải chi tiết
        private ComboBox cboSupplier;
        private System.Data.DataTable _supplierTable;
        // File preview popup
        private Form _filePreviewForm;
        private System.Windows.Forms.Timer _previewTimer;
        private string _previewPath = null;
        private DataGridViewCell _previewCell = null;
        private System.Windows.Forms.Timer _previewCloseTimer;
        // State for current preview
        private class PreviewState
        {
            public string Type; // "image", "text", "pdf", "other"
            public Image OriginalImage;
            public double Scale = 1.0;
            public string FullText;
            public string[] Lines;
            public int TextOffset = 0;
            public int TextPageLines = 200;
        }
        private bool _isSearching = false;
        private bool _isDomestic = true; // true = Mua trong nước (làm tròn số nguyên), false = Mua nước ngoài (2 chữ số thập phân)
        private bool _printDialogOpen = false; // giữ preview mở trong khi hộp thoại in đang hiển thị
        private bool _isLoadingPO = false; // suppress SelectionChanged trong lúc bind lần đầu

        private string _projectCodeImport = string.Empty;
        private Form TopOwner => (this.TopLevelControl as Form) ?? this;
        private Form GetActiveOwner() { var f = TopOwner; f.BringToFront(); f.Activate(); return f; }
        // Suppress DgvPO_SelectionChanged khi show dialog để tránh sync DB calls
        // bị trigger trong message pump của MessageBox (đặc biệt khi frmPO là popup)
        private void SafeWarn(string text, string title = "Thông báo")
        {
            dgvPO.SelectionChanged -= DgvPO_SelectionChanged;
            try { var f = TopOwner; f.BringToFront(); f.Activate(); MessageBox.Show(f, text, title, MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            finally { dgvPO.SelectionChanged += DgvPO_SelectionChanged; }
        }
        private void SafeInfo(string text, string title = "Thông báo")
        {
            dgvPO.SelectionChanged -= DgvPO_SelectionChanged;
            try { var f = TopOwner; f.BringToFront(); f.Activate(); MessageBox.Show(f, text, title, MessageBoxButtons.OK, MessageBoxIcon.Information); }
            finally { dgvPO.SelectionChanged += DgvPO_SelectionChanged; }
        }
        private void SafeErr(string text, string title = "Lỗi")
        {
            dgvPO.SelectionChanged -= DgvPO_SelectionChanged;
            try { var f = TopOwner; f.BringToFront(); f.Activate(); MessageBox.Show(f, text, title, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally { dgvPO.SelectionChanged += DgvPO_SelectionChanged; }
        }
        private bool SafeAsk(string text, string title = "Xác nhận")
        {
            dgvPO.SelectionChanged -= DgvPO_SelectionChanged;
            try { var f = TopOwner; f.BringToFront(); f.Activate(); return MessageBox.Show(f, text, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes; }
            finally { dgvPO.SelectionChanged += DgvPO_SelectionChanged; }
        }

        // =========================================================================
        // XÁC THỰC MẬT KHẨU ADMIN — dùng cho unlock Trạng thái PO
        // =========================================================================
        private bool VerifyAdminPasswordForStatus()
        {
            var dlg = new Form
            {
                Text = "🔐 Xác thực Admin",
                Size = new Size(360, 160),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false, MinimizeBox = false,
                BackColor = Color.FromArgb(245, 245, 245), KeyPreview = true
            };
            dlg.Controls.Add(new Label { Text = "Nhập mật khẩu Admin để đổi trạng thái PO:", Font = new Font("Segoe UI", 9), Location = new Point(14, 14), Size = new Size(320, 20) });
            var txtPwd = new TextBox { Location = new Point(14, 38), Size = new Size(320, 26), Font = new Font("Segoe UI", 10), PasswordChar = '●' };
            dlg.Controls.Add(txtPwd);
            var lblErr = new Label { ForeColor = Color.FromArgb(220, 53, 69), Font = new Font("Segoe UI", 9, FontStyle.Bold), Location = new Point(14, 68), Size = new Size(320, 18), Text = "" };
            dlg.Controls.Add(lblErr);
            bool verified = false;

            var btnOK = new Button { Text = "✔ Xác nhận", Location = new Point(140, 92), Size = new Size(100, 28), BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            btnOK.FlatAppearance.BorderSize = 0;
            var btnCancel = new Button { Text = "Hủy", Location = new Point(250, 92), Size = new Size(84, 28), BackColor = Color.FromArgb(108, 117, 125), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), DialogResult = DialogResult.Cancel };
            btnCancel.FlatAppearance.BorderSize = 0;
            dlg.Controls.AddRange(new Control[] { btnOK, btnCancel });
            dlg.CancelButton = btnCancel;
            dlg.AcceptButton = btnOK;

            btnOK.Click += (s, ev) =>
            {
                string pwd = txtPwd.Text;
                if (string.IsNullOrEmpty(pwd)) { lblErr.Text = "Vui lòng nhập mật khẩu!"; return; }
                try
                {
                    // Hash hardcoded admin password
                    string inputHash;
                    using (var sha = System.Security.Cryptography.SHA256.Create())
                        inputHash = BitConverter.ToString(sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(pwd))).Replace("-", "").ToLower();

                    const string ADMIN_HASH = "e86f78a8a3caf0b60d8e74e5942aa6d86dc150cd3c03338aef25b7d2d7e3acc7";
                    if (inputHash == ADMIN_HASH) { verified = true; dlg.DialogResult = DialogResult.OK; dlg.Close(); return; }

                    // Kiểm tra DB
                    using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                    conn.Open();
                    var cmd = new Microsoft.Data.SqlClient.SqlCommand(
                        @"SELECT COUNT(1) FROM Users WHERE LOWER(Username)='admin' AND (LOWER(Password)=@h OR Password=@hu)", conn);
                    cmd.Parameters.AddWithValue("@h", inputHash);
                    cmd.Parameters.AddWithValue("@hu", inputHash.ToUpper());
                    if (Convert.ToInt32(cmd.ExecuteScalar()) > 0) { verified = true; dlg.DialogResult = DialogResult.OK; dlg.Close(); }
                    else lblErr.Text = "Mật khẩu không đúng!";
                }
                catch { lblErr.Text = "Lỗi xác thực. Thử lại."; }
            };

            dlg.ShowDialog(TopOwner);
            return verified;
        }

        public frmPO(string poNo = "")
        {
            _targetPoNo = poNo;
            InitializeComponent();
            BuildUI();
            ApplyPermissions();
            this.Resize += FrmPO_Resize;
            frmAIChat.Attach(this);
            this.Shown += async (s, e) =>
            {
                var toast = ToastHelper.Attach(this);
                toast.Show("⏳ Đang tải dữ liệu, vui lòng chờ...");
                try
                {
                    await Task.WhenAll(LoadPOAsync(), LoadSupplierComboAsync());
                    LoadDeliveries();
                    if (!string.IsNullOrEmpty(_targetPoNo))
                        SelectPOByNo(_targetPoNo);
                }
                finally { toast.Hide(); }
            };
        }

        public frmPO(string mprNo, bool importMode)
        {
            _importMprNo = mprNo;
            InitializeComponent();
            BuildUI();
            ApplyPermissions();
            this.Resize += FrmPO_Resize;
            frmAIChat.Attach(this);
            this.Shown += async (s, e) =>
            {
                this.BringToFront();
                this.Activate();
                var toast = ToastHelper.Attach(this);
                toast.Show("⏳ Đang tải dữ liệu, vui lòng chờ...");
                try
                {
                    await Task.WhenAll(LoadPOAsync(), LoadSupplierComboAsync());
                    await ImportMPRByNoAsync(_importMprNo, _importMprId);
                }
                finally { toast.Hide(); }
            };
        }

        public void SetImportMprId(int id) { _importMprId = id; }
        public void SetFilteredMprDetailIds(List<int> ids) { _filteredMprDetailIds = ids; }

        private void SelectPOByNo(string poNo)
        {
            // If the form or dgvPO handle is not yet created (called from ctor), defer selection until shown
            if (!this.IsHandleCreated || dgvPO == null || !dgvPO.IsHandleCreated)
            {
                // attach one-time handler to perform selection after the form is shown
                void OnShownOnce(object s, EventArgs e)
                {
                    this.Shown -= OnShownOnce;
                    SelectPOByNo(poNo);
                }
                this.Shown += OnShownOnce;
                return;
            }

            var targetPO = _poList.Find(p => p.PONo == poNo);
            if (targetPO != null) { txtSearch.Text = targetPO.PONo; BtnSearch_Click(null, null); }
            foreach (DataGridViewRow row in dgvPO.Rows)
            {
                if (row.Cells["PO_No"].Value?.ToString() == poNo)
                {
                    try
                    {
                        dgvPO.ClearSelection();
                        row.Selected = true;
                        if (row.Index >= 0) dgvPO.FirstDisplayedScrollingRowIndex = row.Index;
                    }
                    catch
                    {
                        // ignore if UI not ready
                    }
                    break;
                }
            }
        }

        private void BuildUI()
        {
            this.Text = "Quản lý Đơn Đặt Hàng (PO)";
            this.Size = new Size(1300, 780);
            this.MinimumSize = new Size(600, 450);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(245, 245, 245);
            this.AutoScroll = false;

            // ===== MAIN SCROLL PANEL =====
            _mainScroll = new Panel
            {
                Dock = DockStyle.Fill,   // Top = 0, chiếm toàn bộ form
                AutoScroll = true,       // Tự thêm scrollbar dọc/ngang khi nội dung vượt kích thước
                BackColor = Color.FromArgb(245, 245, 245),
                Padding = new Padding(0) // Không có padding thừa
            };
            this.Controls.Add(_mainScroll);

            // ===== PANEL TOP =====
            panelTop = new Panel { Location = new Point(10, 10), Size = new Size(1260, 210), BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Anchor = AnchorStyles.Top | AnchorStyles.Left };
            _mainScroll.Controls.Add(panelTop);
            panelTop.Controls.Add(new Label { Text = "DANH SÁCH ĐƠN ĐẶT HÀNG (PO)", Font = new Font("Segoe UI", 11, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212), Location = new Point(10, 10), Size = new Size(450, 30) });

            // ── Toolbar panelTop: search + buttons trong FlowLayout để scale đồng đều ──
            var flowTop = new FlowLayoutPanel
            {
                Location = new Point(10, 42),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = Color.White,
                Padding = new Padding(0),
                Name = "flowTop"
            };
            panelTop.Controls.Add(flowTop);

            txtSearch = new TextBox { Width = 300, Height = 28, Font = new Font("Segoe UI", 10), PlaceholderText = "Tìm theo PO No, MPR No, tên dự án...", Margin = new Padding(0, 1, 6, 0) };
            panelTop.Controls.Add(txtSearch);
            txtSearch.Location = new Point(10, 45);
            txtSearch.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) BtnSearch_Click(null, null); };

            btnSearch = CreateButton("Tìm", Color.FromArgb(0, 120, 212), Point.Empty, 70, 30); btnSearch.Tag = "70,30";
            btnNewPO = CreateButton("+ Tạo PO", Color.FromArgb(40, 167, 69), Point.Empty, 100, 30); btnNewPO.Tag = "100,30";
            btnDeletePO = CreateButton("Xóa PO", Color.FromArgb(220, 53, 69), Point.Empty, 90, 30); btnDeletePO.Tag = "90,30";
            btnSearchBySupp = CreateButton("🔍 Tìm theo NCC", Color.FromArgb(102, 51, 153), Point.Empty, 130, 30); btnSearchBySupp.Tag = "130,30";
            btnPOStatus = CreateButton("📊 PO status", Color.FromArgb(255, 140, 0), Point.Empty, 110, 30); btnPOStatus.Tag = "110,30";
            var btnPaymentTop = CreateButton("💳 Payment", Color.FromArgb(0, 150, 100), Point.Empty, 110, 30); btnPaymentTop.Tag = "110,30";
            var btnPrintBudget = CreateButton("📑 Budget", Color.FromArgb(70, 130, 180), Point.Empty, 90, 30); btnPrintBudget.Tag = "90,30";

            foreach (var b in new[] { btnSearch, btnNewPO, btnDeletePO, btnSearchBySupp, btnPOStatus, btnPaymentTop, btnPrintBudget })
                b.Margin = new Padding(0, 0, 4, 0);

            btnSearch.Click += BtnSearch_Click;
            btnNewPO.Click += BtnNewPO_Click;
            btnDeletePO.Click += BtnDeletePO_Click;
            btnSearchBySupp.Click += BtnSearchBySupp_Click;
            btnPOStatus.Click += BtnPOStatus_Click;
            btnPaymentTop.Click += (s, ev) =>
            {
                string poNo = dgvPO.SelectedRows.Count > 0 ? dgvPO.SelectedRows[0].Cells["PO_No"].Value?.ToString() ?? "" : "";
                new frmPayment(_currentUser, poNo).Show();
            };
            btnPrintBudget.Click += async (s, ev) => await BtnPrintBudget_ClickAsync();

            flowTop.Controls.Add(btnSearch);
            flowTop.Controls.Add(btnNewPO);
            flowTop.Controls.Add(btnDeletePO);
            flowTop.Controls.Add(btnSearchBySupp);
            flowTop.Controls.Add(btnPOStatus);
            flowTop.Controls.Add(btnPaymentTop);
            flowTop.Controls.Add(btnPrintBudget);

            // Đặt flowTop ngay sau txtSearch
            panelTop.SizeChanged += (s, e) =>
            {
                flowTop.Location = new Point(txtSearch.Right + 6, 44);
            };

            lblStatus = new Label { Location = new Point(10, 78), Size = new Size(800, 22), Font = new Font("Segoe UI", 9), ForeColor = Color.Gray };
            panelTop.Controls.Add(lblStatus);

            dgvPO = new DataGridView
            {
                Location = new Point(10, 105),
                Size = new Size(1235, 115),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            dgvPO.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvPO.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPO.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvPO.EnableHeadersVisualStyles = false;
            dgvPO.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvPO.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
            dgvPO.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgvPO.SelectionChanged += DgvPO_SelectionChanged;
            // Cancelled PO: toàn bộ row màu xám + gạch ngang
            dgvPO.CellFormatting += (s, ev) =>
            {
                if (ev.RowIndex < 0) return;
                var statusCell = dgvPO.Rows[ev.RowIndex].Cells["Trang_Thai"];
                if (statusCell?.Value?.ToString() == "Cancelled")
                {
                    ev.CellStyle.ForeColor = Color.FromArgb(160, 160, 160);
                    ev.CellStyle.BackColor = Color.FromArgb(245, 245, 245);
                    ev.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Strikeout);
                }
            };
            dgvPO.DataBindingComplete += (s, ev) =>
            {
                dgvPO.ClearSelection();
                try { dgvPO.CurrentCell = null; } catch { }
            };

            // Click chuột trái vào bất kỳ ô nào → copy số PO vào clipboard
            dgvPO.CellClick += (s, ev) =>
            {
                if (ev.RowIndex < 0) return;
                if (ev.ColumnIndex >= 0 && dgvPO.Columns[ev.ColumnIndex].Name == "col_SendEmail") return;
                string poNo = dgvPO.Rows[ev.RowIndex].Cells["PO_No"].Value?.ToString() ?? "";
                //if (!string.IsNullOrEmpty(poNo))
                //{
                //    //Clipboard.SetText(poNo);
                //    lblStatus.Text = $"✔ Đã copy: {poNo}";
                //}
            };

            dgvPO.CellContentClick += DgvPO_CellContentClick;

            // Context menu cho bảng PO
            var menuPO = new ContextMenuStrip();
            var itemUnlockSupp = new ToolStripMenuItem("🔓 Mở khóa chỉnh sửa NCC", null, (s, ev) => BtnUnlockSupplier_Click(null, null));
            menuPO.Items.Add(itemUnlockSupp);
            dgvPO.ContextMenuStrip = menuPO;

            panelTop.Controls.Add(dgvPO);

            // ── Bảng Document (📎) góc phải panelTop — INV + Delivery ──
            const int docPOW = 300;
            panelTop.Controls.Add(new Label
            {
                Text = "📎 Document",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(panelTop.Width - docPOW, 8),
                Size = new Size(docPOW - 5, 20),
                Name = "lblDocPO",
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            });
            dgvDocPO = new DataGridView
            {
                Location = new Point(panelTop.Width - docPOW, 28),
                Size = new Size(docPOW - 5, 172), // full height panelTop
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 10),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            dgvDocPO.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvDocPO.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvDocPO.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            dgvDocPO.EnableHeadersVisualStyles = false;
            dgvDocPO.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvDocPO.Columns.Add(new DataGridViewTextBoxColumn { Name = "DocPO_Path", Visible = false });
            dgvDocPO.Columns.Add(new DataGridViewTextBoxColumn { Name = "DocPO_Name", HeaderText = "Tên file (INV / Delivery)", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            dgvDocPO.CellDoubleClick += (s, ev) =>
            {
                if (ev.RowIndex < 0) return;
                string path = dgvDocPO.Rows[ev.RowIndex].Cells["DocPO_Path"].Value?.ToString() ?? "";
                if (System.IO.File.Exists(path))
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
            };
            panelTop.Controls.Add(dgvDocPO);
            // Thu nhỏ dgvPO để nhường chỗ cho dgvDocPO
            dgvPO.Size = new Size(panelTop.Width - docPOW - 20, 115); // dgvPO nhường docPOW=300

            // btnPayment đã được tích hợp vào flowTop phía trên

            // ===== PANEL HEADER =====
            panelHeader = new Panel { Location = new Point(10, 230), Size = new Size(1260, 245), BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Anchor = AnchorStyles.Top | AnchorStyles.Left };
            _mainScroll.Controls.Add(panelHeader);
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
            dgvFiles.CellMouseEnter += DgvFiles_CellMouseEnter;
            dgvFiles.CellMouseLeave += DgvFiles_CellMouseLeave;
            panelHeader.Controls.Add(dgvFiles);
            // Initialize preview handlers for file list
            EnsureFilePreview();

            // BẢNG THEO DÕI GIAO HÀNG — bọc trong Panel con để tọa độ nội bộ luôn chính xác
            const int delivW = 600;
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

            var btnDelivSaveNote = new Button
            {
                Text = "📝 Cập nhật ghi chú",
                Location = new Point(419, 1),
                Size = new Size(125, 22),
                BackColor = Color.FromArgb(255, 140, 0),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnDelivSaveNote.FlatAppearance.BorderSize = 0;
            btnDelivSaveNote.Click += (s, e) => SaveDeliveryReceiverNote();
            panelDelivery.Controls.Add(btnDelivSaveNote);

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

            // ===== VÙNG NHẬP LIỆU: 3 rows nằm ngang, mỗi row là 1 FlowLayoutPanel ──
            // Mỗi row trải dài hết chiều rộng bên trái (cách bảng giao hàng 50px).
            // Chỉ wrap xuống dòng khi form bị kéo nhỏ hơn kích thước gốc.

            // Khoảng cách bên phải dành cho Delivery + Files (delivW=550 + gap=6 + files=200 + 10 + 50)
            const int rightReserved = 816;

            // Hàm tạo Label + Control dàn đều trong một panel nhỏ
            Panel MakeField(string labelText, Control ctrl, int labelW = 68, int gap = 4)
            {
                var lbl = new Label
                {
                    Text = labelText,
                    Font = new Font("Segoe UI", 9),
                    AutoSize = false,
                    Size = new Size(labelW, 25),
                    TextAlign = System.Drawing.ContentAlignment.MiddleLeft
                };
                ctrl.Size = new Size(ctrl.Width, 25);
                var pnl = new Panel
                {
                    Size = new Size(labelW + gap + ctrl.Width + 4, 30),
                    BackColor = Color.White,
                    Margin = new Padding(0, 2, 8, 2)
                };
                lbl.Location = new Point(0, 3);
                ctrl.Location = new Point(labelW + gap, 2);
                pnl.Controls.Add(lbl);
                pnl.Controls.Add(ctrl);
                return pnl;
            }

            // Hàm tạo một row FlowLayoutPanel — không wrap ở kích thước bình thường
            FlowLayoutPanel MakeRow(int top)
            {
                var row = new FlowLayoutPanel
                {
                    Location = new Point(10, top),
                    FlowDirection = FlowDirection.LeftToRight,
                    WrapContents = true,          // wrap xuống dòng khi form bị kéo nhỏ
                    AutoSize = true,              // tự tăng chiều cao khi wrap
                    AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    BackColor = Color.White,
                    Padding = new Padding(0),
                    Anchor = AnchorStyles.Top | AnchorStyles.Left
                };
                // Khi row tự thay đổi chiều cao (do wrap), đẩy các row bên dưới xuống theo
                panelHeader.Controls.Add(row);
                return row;
            }

            // ── Row 1 (y=35): PO No | Tên dự án | Workorder | MPR No | Nhà CC ──
            var flowRow1 = MakeRow(35);
            txtPONo = new TextBox { Width = 100, Font = new Font("Segoe UI", 9) };
            txtProjectName = new TextBox { Width = 160, Font = new Font("Segoe UI", 9) };
            txtWorkorderNo = new TextBox { Width = 120, Font = new Font("Segoe UI", 9) };
            txtMPRNo = new TextBox { Width = 220, Font = new Font("Segoe UI", 9) };
            txtMPRNo.Leave += async (s, e) => await LoadMPRFilesAsync();

            cboSupplier = new ComboBox { Width = 250, Font = new Font("Segoe UI", 9), DropDownStyle = ComboBoxStyle.DropDown, AutoCompleteMode = AutoCompleteMode.None, DropDownWidth = 600, MaxDropDownItems = 20 };
            cboSupplier.Validating += CboSupplier_Validating;
            cboSupplier.SelectedIndexChanged += CboSupplier_SelectedIndexChanged;
            cboSupplier.TextChanged += CboSupplier_TextChanged;
            cboSupplier.KeyDown += CboSupplier_KeyDown;

            // Nút 🔓 unlock Nhà cung cấp ngay cạnh ComboBox
            var btnUnlockSuppSmall = new Button
            {
                Text = "🔒",
                Size = new Size(26, 26),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 8),
                Cursor = Cursors.Hand,
                Margin = new Padding(2, 1, 0, 0)
            };
            btnUnlockSuppSmall.FlatAppearance.BorderSize = 0;
            btnUnlockSuppSmall.Click += (s, ev) => BtnUnlockSupplier_Click(null, null);

            // Cập nhật text nút dựa theo trạng thái cboSupplier
            cboSupplier.EnabledChanged += (s, ev) =>
            {
                btnUnlockSuppSmall.Text = cboSupplier.Enabled ? "🔓" : "🔒";
                btnUnlockSuppSmall.BackColor = cboSupplier.Enabled ? Color.FromArgb(220, 53, 69) : Color.FromArgb(108, 117, 125);
            };

            flowRow1.Controls.Add(MakeField("PO No (*):", txtPONo, 60));
            flowRow1.Controls.Add(MakeField("Tên dự án:", txtProjectName, 65));
            flowRow1.Controls.Add(MakeField("Workorder:", txtWorkorderNo, 65));
            flowRow1.Controls.Add(MakeField("MPR No:", txtMPRNo, 55));
            
            var suppPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, WrapContents = false, AutoSize = true, Margin = new Padding(0) };
            suppPanel.Controls.Add(MakeField("Nhà CC:", cboSupplier, 50));
            suppPanel.Controls.Add(btnUnlockSuppSmall);
            flowRow1.Controls.Add(suppPanel);

            // ── Row 2 (y=71): Ngày PO | Trạng thái | Revise ──
            var flowRow2 = MakeRow(71);
            // LoadSupplierCombo() được gọi async trong Shown event
            dtpPODate = new DateTimePicker { Width = 108, Font = new Font("Segoe UI", 9), Format = DateTimePickerFormat.Short };
            cboStatus = new ComboBox { Width = 108, Font = new Font("Segoe UI", 9), DropDownStyle = ComboBoxStyle.DropDownList, Enabled = false };
            cboStatus.Items.AddRange(new[] { "Draft", "Pending", "Approved", "In Progress", "Completed", "Cancelled" });
            cboStatus.SelectedIndex = 0; // Draft
            nudRevise = new NumericUpDown { Width = 52, Font = new Font("Segoe UI", 9), Minimum = 0, Maximum = 99 };
            flowRow2.Controls.Add(MakeField("Ngày PO:", dtpPODate, 60));

            // ── Trạng thái + nút 🔓 unlock bằng mật khẩu Admin ──
            var statusPanel = new Panel { Width = 185, Height = 28, BackColor = Color.Transparent, Margin = new Padding(0, 2, 6, 0) };
            var lblStatusLbl = new Label { Text = "Trạng thái:", Width = 68, Height = 28, Font = new Font("Segoe UI", 9), TextAlign = ContentAlignment.MiddleLeft };
            cboStatus.Location = new Point(72, 1);
            var btnUnlockStatus = new Button
            {
                Text = "🔒",
                Location = new Point(184, 1),
                Size = new Size(26, 26),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9),
                Cursor = Cursors.Hand,
                Tag = "locked"
            };
            btnUnlockStatus.FlatAppearance.BorderSize = 0;
            btnUnlockStatus.Click += (s, ev) =>
            {
                if (btnUnlockStatus.Tag?.ToString() == "locked")
                {
                    // Xác thực mật khẩu Admin để unlock
                    if (VerifyAdminPasswordForStatus())
                    {
                        cboStatus.Enabled = true;
                        btnUnlockStatus.Text = "🔓";
                        btnUnlockStatus.BackColor = Color.FromArgb(220, 53, 69);
                        btnUnlockStatus.Tag = "unlocked";
                    }
                }
                else
                {
                    // Khoá lại
                    cboStatus.Enabled = false;
                    btnUnlockStatus.Text = "🔒";
                    btnUnlockStatus.BackColor = Color.FromArgb(108, 117, 125);
                    btnUnlockStatus.Tag = "locked";
                }
            };
            // Tự động lock lại sau khi chọn xong
            cboStatus.SelectedIndexChanged += (s, ev) =>
            {
                if (cboStatus.Enabled && btnUnlockStatus.Tag?.ToString() == "unlocked")
                {
                    cboStatus.Enabled = false;
                    btnUnlockStatus.Text = "🔒";
                    btnUnlockStatus.BackColor = Color.FromArgb(108, 117, 125);
                    btnUnlockStatus.Tag = "locked";
                }
            };
            statusPanel.Controls.Add(lblStatusLbl);
            statusPanel.Controls.Add(cboStatus);
            statusPanel.Controls.Add(btnUnlockStatus);
            statusPanel.Width = btnUnlockStatus.Right + 2;
            flowRow2.Controls.Add(statusPanel);

            flowRow2.Controls.Add(MakeField("Revise:", nudRevise, 48));

            // Khởi tạo các TextBox ẩn để giữ tương thích — không hiển thị trên UI
            txtPrepared = new TextBox { Visible = false };
            txtReviewed = new TextBox { Visible = false };
            txtAgreement = new TextBox { Visible = false };
            txtApproved = new TextBox { Visible = false };

            // ── Row 3 (y=107): Ghi chú | Payment Term | Ngày giao hàng ──
            var flowRow3 = MakeRow(107);
            txtNotes = new TextBox { Width = 160, Font = new Font("Segoe UI", 9) };
            cboPaymentTerm = new ComboBox
            {
                Width = 210,
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
            dtpPOExpectDelivery = new DateTimePicker { Width = 108, Font = new Font("Segoe UI", 9), Format = DateTimePickerFormat.Short };
            flowRow3.Controls.Add(MakeField("Ghi chú:", txtNotes, 52));
            flowRow3.Controls.Add(MakeField("Payment Term:", cboPaymentTerm, 88));
            flowRow3.Controls.Add(MakeField("Ngày Giao hàng:", dtpPOExpectDelivery, 96));

            // Lưu tham chiếu flowHeader = flowRow1 để các handler cũ vẫn hoạt động
            var flowHeader = flowRow1;

            // ── Row 5: Buttons — dùng FlowLayoutPanel riêng để wrap ──
            var flowBtns = new FlowLayoutPanel
            {
                Location = new Point(10, 145), // ngay dưới row3 (107+34+4)
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowOnly,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                BackColor = Color.White,
                Padding = new Padding(0, 2, 0, 2),
                Margin = new Padding(0),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            panelHeader.Controls.Add(flowBtns);

            // flowRow3.SizeChanged → đã tích hợp vào relayoutHeader bên dưới

            btnSavePO = CreateButton("💾 Lưu Toàn Bộ PO", Color.FromArgb(0, 120, 212), Point.Empty, 150, 32);
            btnSavePO.Margin = new Padding(0, 0, 6, 0); btnSavePO.Tag = "150,32";
            btnSavePO.Click += BtnSavePO_Click;
            btnClearHeader = CreateButton("Làm mới", Color.FromArgb(108, 117, 125), Point.Empty, 100, 32);
            btnClearHeader.Margin = new Padding(0, 0, 6, 0); btnClearHeader.Tag = "100,32";
            btnClearHeader.Click += BtnClearHeader_Click;

            btnUnlockSupplier = CreateButton("🔓 Mở khóa NCC", Color.FromArgb(255, 140, 0), Point.Empty, 125, 32);
            btnUnlockSupplier.Margin = new Padding(0, 0, 6, 0); btnUnlockSupplier.Tag = "125,32";
            btnUnlockSupplier.Visible = false;
            btnUnlockSupplier.Click += BtnUnlockSupplier_Click;
            var btnImportMPR = CreateButton("Import MPR", Color.FromArgb(255, 140, 0), Point.Empty, 120, 32); btnImportMPR.Tag = "120,32";
            btnImportMPR.Margin = new Padding(0, 0, 6, 0);
            btnImportMPR.Click += BtnImportMPR_Click;
            var btnHistory = CreateButton("Revise History", Color.FromArgb(102, 51, 153), Point.Empty, 130, 32); btnHistory.Tag = "130,32";
            btnHistory.Margin = new Padding(0, 0, 6, 0);
            btnHistory.Click += (s, e) => { if (string.IsNullOrEmpty(txtPONo.Text)) { SafeWarn("Vui lòng chọn một PO trước!"); return; } new frmReviseHistory(txtPONo.Text).ShowDialog(); };
            flowBtns.Controls.Add(btnSavePO);
            flowBtns.Controls.Add(btnClearHeader);
            flowBtns.Controls.Add(btnUnlockSupplier);
            flowBtns.Controls.Add(btnImportMPR);
            flowBtns.Controls.Add(btnHistory);

            // Resize 3 rows và tính lại vị trí dọc khi form thay đổi kích thước
            Action relayoutHeader = () =>
            {
                // Chiều rộng field rows = panelHeader.Width - vùng phải (delivery+files+50px)
                int flowW = Math.Max(panelHeader.Width - rightReserved - 10, 280);
                // Các row field giới hạn theo flowW → wrap khi form nhỏ
                flowRow1.MaximumSize = new Size(flowW, 0);
                flowRow2.MaximumSize = new Size(flowW, 0);
                flowRow3.MaximumSize = new Size(flowW, 0);
                // flowBtns KHÔNG giới hạn theo rightReserved — luôn hiển thị 1 hàng
                // khi form đủ rộng; chỉ wrap khi form quá hẹp (< tổng width 4 buttons)
                flowBtns.Width = Math.Max(panelHeader.Width - 20, 100);
                flowBtns.PerformLayout();
                // Tính lại Top từng row dựa trên Bottom của row trước (xử lý wrap)
                flowRow1.Location = new Point(10, 35);
                flowRow2.Location = new Point(10, flowRow1.Bottom + 2);
                flowRow3.Location = new Point(10, flowRow2.Bottom + 2);
                flowBtns.Location = new Point(10, flowRow3.Bottom + 2);
                panelHeader.Height = flowBtns.Bottom + 4;

                // ── Cập nhật chiều cao panelDelivery và dgvFiles theo panelHeader.Height ──
                int newH = panelHeader.Height;
                if (panelDelivery != null)
                {
                    panelDelivery.Location = new Point(panelHeader.Width - rightReserved, 8);
                    panelDelivery.Height = newH - 16;
                    if (dgvDelivery != null)
                        dgvDelivery.Height = panelDelivery.Height - 26;
                }
                if (dgvFiles != null)
                {
                    dgvFiles.Height = newH - 20;
                    dgvFiles.Location = new Point(panelHeader.Width - dgvFiles.Width - 10, 10);
                }
            };
            panelHeader.SizeChanged += (s, e) => relayoutHeader();
            // Mỗi row tự wrap → Bottom thay đổi → đẩy các row bên dưới
            flowRow1.SizeChanged += (s, e) => relayoutHeader();
            flowRow2.SizeChanged += (s, e) => relayoutHeader();
            flowRow3.SizeChanged += (s, e) => relayoutHeader();
            flowBtns.SizeChanged += (s, e) => relayoutHeader();
            relayoutHeader(); // Tính chiều cao chính xác ngay khi khởi tạo

            // ===== PANEL DETAIL =====
            panelDetail = new Panel { Location = new Point(10, 500), Size = new Size(980, 285), BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Anchor = AnchorStyles.Top | AnchorStyles.Left };
            _mainScroll.Controls.Add(panelDetail);

            // ── Panel MPR Files (bên phải panelDetail) ──
            var panelMPRFiles = new Panel
            {
                Location = new Point(panelDetail.Right + 10, panelDetail.Top),
                Size = new Size(280, 285),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom
            };
            _mainScroll.Controls.Add(panelMPRFiles);

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
                    SafeWarn("Không tìm thấy file!");
            };
            dgvMPRFiles.CellMouseEnter += DgvFiles_CellMouseEnter;
            dgvMPRFiles.CellMouseLeave += DgvFiles_CellMouseLeave;
            panelMPRFiles.Controls.Add(dgvMPRFiles);
            panelDetail.Controls.Add(new Label { Text = "CHI TIẾT ĐƠN HÀNG", Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212), Location = new Point(10, 8), Size = new Size(160, 25) });

            // FlowLayoutPanel cho toolbar chi tiết — tự wrap khi form nhỏ
            var flowDetailBtns = new FlowLayoutPanel
            {
                Location = new Point(10, 36),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,   // wrap xuống dòng khi form bị kéo nhỏ
                BackColor = Color.White,
                Padding = new Padding(0),
                Name = "flowDetailBtns",
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            panelDetail.Controls.Add(flowDetailBtns);
            // Cập nhật MaximumSize khi panelDetail resize để tránh wrap vào vùng panelRight (590px)
            const int panelRightW = 600; // panelRight width + buffer
            panelDetail.Resize += (s, ev) =>
            {
                int maxFlowW = Math.Max(panelDetail.Width - panelRightW, 200);
                flowDetailBtns.MaximumSize = new Size(maxFlowW, 0);
            };

            btnAddDetail = CreateButton("+ Thêm", Color.FromArgb(40, 167, 69), Point.Empty, 80, 28);
            btnDeleteDetail = CreateButton("Xóa dòng", Color.FromArgb(220, 53, 69), Point.Empty, 85, 28);
            var btnSaveDetail = CreateButton("💾 Lưu chi tiết", Color.FromArgb(0, 120, 212), Point.Empty, 115, 28); btnSaveDetail.Tag = "115,28";
            btnExport = CreateButton("📄 Xuất Excel", Color.FromArgb(0, 150, 100), Point.Empty, 110, 28);
            var btnCheckBySize = CreateButton("🔍 Check size", Color.FromArgb(102, 51, 153), Point.Empty, 110, 28); btnCheckBySize.Tag = "110,28";

            btnAddDetail.Margin = new Padding(0, 0, 3, 0); btnAddDetail.Tag = "80,28";
            btnDeleteDetail.Margin = new Padding(0, 0, 3, 0); btnDeleteDetail.Tag = "85,28";
            btnSaveDetail.Margin = new Padding(0, 0, 4, 0);
            btnExport.Margin = new Padding(0, 0, 3, 0); btnExport.Tag = "110,28";
            btnCheckBySize.Margin = new Padding(0, 0, 4, 0);

            btnAddDetail.Click += BtnAddDetail_Click;
            btnDeleteDetail.Click += BtnDeleteDetail_Click;
            btnSaveDetail.Click += BtnSaveDetail_Click;
            btnExport.Click += BtnExport_Click;
            btnCheckBySize.Click += BtnCheckBySize_Click;

            flowDetailBtns.Controls.Add(btnAddDetail);
            flowDetailBtns.Controls.Add(btnDeleteDetail);
            flowDetailBtns.Controls.Add(btnSaveDetail);
            flowDetailBtns.Controls.Add(btnExport);        // Xuất Excel
            flowDetailBtns.Controls.Add(btnCheckBySize);   // Check by size — cạnh Xuất Excel

            var btnCalculartor = CreateButton("📠 Tính", Color.FromArgb(102, 51, 153), Point.Empty, 110, 28); /*btnCalculartor.Tag = "110,28";*/
            flowDetailBtns.Controls.Add(btnCalculartor);
            btnCalculartor.Margin = new Padding(0, 0, 4, 0);

            btnCalculartor.Click += (s, e) =>
            {
                if (dgvDetails.Rows.Count == 0 || (dgvDetails.Rows.Count == 1 && dgvDetails.Rows[0].IsNewRow))
                {
                    SafeInfo("Không có dòng nào để tính!");
                    return;
                }
                foreach (DataGridViewRow row in dgvDetails.Rows)
                {
                    if (row.IsNewRow || row.Tag?.ToString() == "TOTAL") continue;
                    RecalculateAmount(row.Index);
                }
                //UpdateTotal();
            };


            // ── Bộ lọc "Đã lên PO" — di chuyển lên cạnh tiêu đề "CHI TIẾT ĐƠN HÀNG" ──────────
            var lblFilterPO = new Label
            {
                Text = "Đã lên PO:",
                AutoSize = true,
                Location = new Point(180, 11),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.Black
            };
            panelDetail.Controls.Add(lblFilterPO);

            var cboFilterPO = new ComboBox
            {
                Size = new Size(110, 25),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Name = "cboFilterPO",
                Location = new Point(255, 6)
            };
            cboFilterPO.Items.AddRange(new object[] { "Tất cả", "Đã lên PO", "Chưa lên PO" });
            cboFilterPO.SelectedIndex = 0;
            cboFilterPO.SelectedIndexChanged += (s, ev) =>
            {
                string sel = cboFilterPO.SelectedItem?.ToString() ?? "Tất cả";
                foreach (DataGridViewRow row in dgvDetails.Rows)
                {
                    if (row.IsNewRow) continue;
                    string poVal = row.Cells["Ordered_PO"]?.Value?.ToString() ?? "";
                    row.Visible = sel switch
                    {
                        "Đã lên PO" => !string.IsNullOrWhiteSpace(poVal),
                        "Chưa lên PO" => string.IsNullOrWhiteSpace(poVal),
                        _ => true
                    };
                }
            };
            panelDetail.Controls.Add(cboFilterPO);

            // Đẩy dgvDetails xuống dưới toolbar
            flowDetailBtns.SizeChanged += (s, e) =>
            {
                int dgvTop = flowDetailBtns.Bottom + 4;
                if (dgvDetails != null) dgvDetails.Top = dgvTop;
            };


            // ── Panel phai: chua labels TruocVAT/SauVAT + apply VAT + apply Calc ──
            // Anchor Right de tu can le phai khi resize
            var panelRight = new Panel
            {
                Size = new Size(740, 68),
                BackColor = Color.Transparent,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            // Vi tri: sat le phai panelDetail, dong voi hang buttons
            panelRight.Location = new Point(panelDetail.Width - 740 - 5, 2);
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
                    {
                        row.Cells["VAT"].Value = vatVal;
                        RecalculateAmount(row.Index);
                    }
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
                    {
                        row.Cells["Calc_Method"].Value = calcVal;
                        RecalculateAmount(row.Index);
                    }
                UpdateTotal();
            };
            panelRight.Controls.Add(btnApplyCalc);

            // -- Cot 4: Loai mua (x=500) --
            int xMode = 500;
            panelRight.Controls.Add(new Label
            {
                Text = "Loại mua:",
                Location = new Point(xMode, 2),
                Size = new Size(65, 18),
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            });
            var rdoDomestic = new RadioButton
            {
                Text = "Mua trong nước",
                Location = new Point(xMode, 22),
                AutoSize = true,
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 60),
                Checked = true
            };
            var rdoForeign = new RadioButton
            {
                Text = "Mua nước ngoài",
                Location = new Point(xMode, 45),
                AutoSize = true,
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.FromArgb(180, 90, 0)
            };
            panelRight.Controls.Add(rdoDomestic);
            panelRight.Controls.Add(rdoForeign);
            rdoDomestic.CheckedChanged += (s, ev) =>
            {
                if (!rdoDomestic.Checked) return;
                _isDomestic = true;
                foreach (DataGridViewRow r in dgvDetails.Rows)
                    if (!r.IsNewRow && r.Tag?.ToString() != "TOTAL") RecalculateAmount(r.Index);
                UpdateTotal();
            };
            rdoForeign.CheckedChanged += (s, ev) =>
            {
                if (!rdoForeign.Checked) return;
                _isDomestic = false;
                foreach (DataGridViewRow r in dgvDetails.Rows)
                    if (!r.IsNewRow && r.Tag?.ToString() != "TOTAL") RecalculateAmount(r.Index);
                UpdateTotal();
            };

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
            dgvDetails.CellBeginEdit += (s, e) => { if (dgvDetails.Columns[e.ColumnIndex].Name == "NhapKho") e.Cancel = true; };
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
                    MessageBox.Show(GetActiveOwner(), "Không thể mở file: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else if (!string.IsNullOrEmpty(path))
            {
                SafeWarn("File không tồn tại hoặc đã bị xóa / di chuyển khỏi thư mục!");
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

                if (prj != null && !string.IsNullOrEmpty(prj.PO_Link))
                {
                    // If the project's PO_Link ends with a folder named "Excel", prefer the sibling "PDF" folder
                    string configured = prj.PO_Link.Trim();
                    string useDir = configured;
                    try
                    {
                        string trimmed = configured.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                        string last = Path.GetFileName(trimmed) ?? "";
                        if (last.Equals("Excel", StringComparison.OrdinalIgnoreCase))
                        {
                            string parent = Path.GetDirectoryName(trimmed) ?? trimmed;
                            string alt = Path.Combine(parent, "PDF");
                            if (Directory.Exists(alt)) useDir = alt;
                            // if alt doesn't exist, fall back to configured
                        }
                    }
                    catch { /* ignore path parse errors and use configured path */ }

                    if (Directory.Exists(useDir))
                    {
                        var files = Directory.GetFiles(useDir);
                        foreach (var f in files)
                        {
                            dgvFiles.Rows.Add(Path.GetFileName(f), f);
                        }
                    }
                    else
                    {
                        dgvFiles.Rows.Add("(Thư mục không tồn tại)", "");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Lỗi load files: " + ex.Message);
            }
        }

        // Initialize preview timer and handlers for file grid — call once when form init
        private void EnsureFilePreview()
        {
            if (_previewTimer != null) return;
            _previewTimer = new System.Windows.Forms.Timer { Interval = 300 };
            _previewTimer.Tick += (s, e) =>
            {
                _previewTimer.Stop();
                if (!string.IsNullOrEmpty(_previewPath) && _previewCell != null)
                {
                    try { ShowFilePreview(_previewPath, Cursor.Position); }
                    catch { }
                }
            };
        }

        private void DgvFiles_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
                var dgv = sender as DataGridView;
                if (dgv == null) return;
                var cell = dgv.Rows[e.RowIndex].Cells[e.ColumnIndex];
                var fullPathCell = dgv.Rows[e.RowIndex].Cells.Count > 1 ? dgv.Rows[e.RowIndex].Cells[1] : null;
                string path = fullPathCell?.Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(path)) { _previewPath = null; _previewCell = null; return; }
                _previewPath = path;
                _previewCell = cell;
                EnsureFilePreview();
                _previewTimer.Stop();
                _previewTimer.Start();
            }
            catch { }
        }

        private void DgvFiles_CellMouseLeave(object sender, DataGridViewCellEventArgs e)
        {
            // stop show timer and start close timer so user can move mouse into preview
            _previewTimer?.Stop();
            _previewPath = null; _previewCell = null;
            StartPreviewCloseTimer();
        }

        private void HideFilePreview()
        {
            try
            {
                if (_filePreviewForm != null)
                {
                    if (!_filePreviewForm.IsDisposed) _filePreviewForm.Hide();
                    _filePreviewForm.Dispose();
                    _filePreviewForm = null;
                }
            }
            catch { }
        }

        private void StartPreviewCloseTimer(int intervalMs = 800)
        {
            try
            {
                if (_printDialogOpen) return; // Không đóng khi hộp thoại in đang mở
                if (_previewCloseTimer == null)
                {
                    _previewCloseTimer = new System.Windows.Forms.Timer { Interval = intervalMs };
                    _previewCloseTimer.Tick += (s, e) =>
                    {
                        _previewCloseTimer.Stop();
                        HideFilePreview();
                    };
                }
                _previewCloseTimer.Interval = intervalMs;
                _previewCloseTimer.Stop();
                _previewCloseTimer.Start();
            }
            catch { }
        }

        private void StopPreviewCloseTimer()
        {
            try { _previewCloseTimer?.Stop(); }
            catch { }
        }

        private void ShowFilePreview(string filePath, Point screenPos)
        {
            HideFilePreview();
            if (!File.Exists(filePath)) return;
            Form _previewOwner = this.TopLevelControl as Form ?? this;

            var f = new Form
            {
                FormBorderStyle = FormBorderStyle.SizableToolWindow,
                Size            = new Size(1024, 768),
                StartPosition   = FormStartPosition.Manual,
                ShowInTaskbar   = false,
                TopMost         = true,
                BackColor       = Color.FromArgb(38, 38, 38),
                Text            = Path.GetFileName(filePath)
            };
            var scr = Screen.FromPoint(screenPos)?.WorkingArea ?? new Rectangle(0, 0, 1024, 768);
            f.Bounds = new Rectangle(
                scr.Left + (scr.Width  - f.Width)  / 2,
                scr.Top  + (scr.Height - f.Height) / 2,
                f.Width, f.Height);

            // ── Toolbar ──────────────────────────────────────────────────────
            var toolbar = new Panel { Dock = DockStyle.Top, Height = 38, BackColor = Color.FromArgb(45, 45, 48) };
            var clrDark = Color.FromArgb(62, 62, 66);
            var clrBlue = Color.FromArgb(0, 100, 180);

            Button MkBtn(string txt, int left, int w, Color bg)
            {
                var b = new Button
                {
                    Text = txt, Left = left, Top = 5, Width = w, Height = 28,
                    FlatStyle = FlatStyle.Flat, BackColor = bg,
                    ForeColor = Color.White,
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    Cursor = Cursors.Hand
                };
                b.FlatAppearance.BorderColor = Color.FromArgb(80, 80, 80);
                toolbar.Controls.Add(b);
                return b;
            }

            var btnZoomIn  = MkBtn("＋",    8, 36, clrDark);
            var btnZoomOut = MkBtn("－",   48, 36, clrDark);
            var btnFit     = MkBtn("⊡ Fit", 88, 60, clrDark);
            var btn100     = MkBtn("1:1",  152, 44, clrBlue);
            var lblZoom    = new Label { Left = 202, Top = 11, Width = 58, Text = "—", ForeColor = Color.FromArgb(220, 220, 220), Font = new Font("Segoe UI", 9, FontStyle.Bold), BackColor = Color.Transparent };
            var btnPrev    = MkBtn("◀",   268, 36, clrDark); btnPrev.Visible = false;
            var btnNext    = MkBtn("▶",   308, 36, clrDark); btnNext.Visible = false;
            var lblInfo    = new Label { Left = 352, Top = 11, Width = 360, Text = Path.GetFileName(filePath), ForeColor = Color.FromArgb(155, 155, 155), Font = new Font("Segoe UI", 8.5f), BackColor = Color.Transparent };
            // Add labels before right-side buttons so buttons have higher z-order
            toolbar.Controls.Add(lblZoom);
            toolbar.Controls.Add(lblInfo);
            var btnPrint  = MkBtn("⎙ In file",      0, 90, Color.FromArgb(80, 60, 140));
            var btnOpen   = MkBtn("↗ Mở file",     0, 84, Color.FromArgb(0, 120, 60));
            var btnFolder = MkBtn("📁 Mở thư mục", 0, 108, Color.FromArgb(160, 100, 0));
            var btnClose  = MkBtn("✖",              0, 36, Color.FromArgb(196, 43, 28));
            // Position right-side buttons dynamically — Anchor fails when toolbar width is 0 at button creation
            void PlaceRightBtns()
            {
                int r = toolbar.ClientSize.Width - 4;
                btnClose.Left  = r - btnClose.Width;  r = btnClose.Left  - 4;
                btnFolder.Left = r - btnFolder.Width; r = btnFolder.Left - 4;
                btnOpen.Left   = r - btnOpen.Width;   r = btnOpen.Left   - 4;
                btnPrint.Left  = r - btnPrint.Width;
            }
            toolbar.Resize += (s, e) =>
            {
                PlaceRightBtns();
                foreach (Control c in toolbar.Controls)
                    if (c is Button) c.BringToFront();
            };
            btnFolder.Click += (s, e) =>
            {
                try
                {
                    string dir = Path.GetDirectoryName(filePath);
                    if (Directory.Exists(dir))
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = "explorer.exe",
                            Arguments = $"/select,\"{filePath}\"",
                            UseShellExecute = true
                        });
                }
                catch { }
            };
            f.Controls.Add(toolbar);   // triggers first Resize → buttons placed correctly
            foreach (Control c in toolbar.Controls)
                if (c is Button) c.BringToFront();

            Action printAction = null;
            btnPrint.Click += (s, e) =>
            {
                try
                {
                    StopPreviewCloseTimer();
                    _printDialogOpen = true;
                    printAction?.Invoke();
                }
                catch { }
                finally
                {
                    _printDialogOpen = false;
                    // Sau khi hộp thoại in đóng: nếu chuột đã ra ngoài Preview thì đợi 1 giây rồi đóng
                    var pf = _filePreviewForm;
                    if (pf != null && !pf.IsDisposed)
                    {
                        var pt = pf.PointToClient(Cursor.Position);
                        if (!pf.ClientRectangle.Contains(pt))
                            StartPreviewCloseTimer(1000);
                    }
                }
            };
            btnOpen .Click += (s, e) => { try { Process.Start(new ProcessStartInfo { FileName = filePath, UseShellExecute = true }); } catch { } };
            btnClose.Click += (s, e) => f.Close();
            f.KeyPreview = true;
            f.KeyDown   += (s, e) => { if (e.KeyCode == Keys.Escape) f.Close(); };

            // ── Content ──────────────────────────────────────────────────────
            string ext      = Path.GetExtension(filePath).ToLowerInvariant();
            bool   isImgExt = ext == ".png" || ext == ".jpg" || ext == ".jpeg" || ext == ".bmp" || ext == ".gif";

            if (isImgExt || ext == ".pdf")
            {
                // Load source image (PDF via shell thumbnail, images directly)
                Bitmap origBmp = null;
                try
                {
                    Image src = null;
                    if (ext == ".pdf") src = GetShellThumbnail(filePath, 1800, 2400);
                    if (src == null && isImgExt) src = Image.FromFile(filePath);
                    if (src != null)
                    {
                        if (src is Bitmap bSrc) origBmp = bSrc;
                        else { origBmp = new Bitmap(src); src.Dispose(); }
                    }
                }
                catch { origBmp = null; }

                if (origBmp != null)
                {
                    // ── Canvas-based rendering: no new Bitmap on every zoom ──
                    double scale  = 1.0;
                    float  offX   = 0, offY = 0;
                    bool   fitted = false;
                    bool   dragging = false;
                    Point  dragStart = Point.Empty;
                    float  dragOX = 0, dragOY = 0;

                    var canvas = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(38, 38, 38) };
                    // Enable double-buffering to eliminate flicker during redraws
                    typeof(Panel).GetProperty("DoubleBuffered",
                        System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                        ?.SetValue(canvas, true);

                    void CalcFit()
                    {
                        if (canvas.ClientSize.Width <= 0 || canvas.ClientSize.Height <= 0) return;
                        double sx = (double)canvas.ClientSize.Width  / origBmp.Width;
                        double sy = (double)canvas.ClientSize.Height / origBmp.Height;
                        scale = Math.Min(sx, sy) * 0.97;
                        offX  = (canvas.ClientSize.Width  - (float)(origBmp.Width  * scale)) / 2f;
                        offY  = (canvas.ClientSize.Height - (float)(origBmp.Height * scale)) / 2f;
                    }

                    void ClampPan()
                    {
                        float dw = (float)(origBmp.Width  * scale);
                        float dh = (float)(origBmp.Height * scale);
                        if (dw <= canvas.ClientSize.Width)  offX = (canvas.ClientSize.Width  - dw) / 2f;
                        else offX = Math.Max(canvas.ClientSize.Width  - dw, Math.Min(0, offX));
                        if (dh <= canvas.ClientSize.Height) offY = (canvas.ClientSize.Height - dh) / 2f;
                        else offY = Math.Max(canvas.ClientSize.Height - dh, Math.Min(0, offY));
                    }

                    void UpdateZoomLabel() => lblZoom.Text = $"{(int)(scale * 100)}%";

                    // Zoom centered on a pixel coordinate in canvas space
                    void ZoomAt(int px, int py, double factor)
                    {
                        double ns = Math.Max(0.03, Math.Min(30.0, scale * factor));
                        float  r  = (float)(ns / scale);
                        offX  = px - r * (px - offX);
                        offY  = py - r * (py - offY);
                        scale = ns;
                        ClampPan();
                        UpdateZoomLabel();
                        canvas.Invalidate();
                    }

                    canvas.Paint += (s, e) =>
                    {
                        // First paint: fit image to canvas
                        if (!fitted && canvas.ClientSize.Width > 0 && canvas.ClientSize.Height > 0)
                        {
                            CalcFit(); fitted = true; UpdateZoomLabel();
                        }
                        var g = e.Graphics;
                        g.Clear(Color.FromArgb(38, 38, 38));
                        g.InterpolationMode  = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                        g.SmoothingMode      = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                        g.PixelOffsetMode    = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                        g.DrawImage(origBmp, offX, offY,
                            (float)(origBmp.Width * scale), (float)(origBmp.Height * scale));
                    };

                    // Mouse drag to pan
                    canvas.MouseDown  += (s, e) =>
                    {
                        if (e.Button != MouseButtons.Left) return;
                        dragging = true; dragStart = e.Location; dragOX = offX; dragOY = offY;
                        canvas.Cursor = Cursors.SizeAll;
                    };
                    canvas.MouseMove  += (s, e) =>
                    {
                        if (!dragging) return;
                        offX = dragOX + (e.X - dragStart.X);
                        offY = dragOY + (e.Y - dragStart.Y);
                        ClampPan();
                        canvas.Invalidate();
                    };
                    canvas.MouseUp    += (s, e) => { dragging = false; canvas.Cursor = Cursors.Default; };
                    canvas.MouseLeave += (s, e) => { dragging = false; canvas.Cursor = Cursors.Default; };

                    // Mouse wheel zoom — centered on cursor position
                    canvas.MouseWheel += (s, e) => ZoomAt(e.X, e.Y, e.Delta > 0 ? 1.15 : 1.0 / 1.15);
                    canvas.MouseEnter += (s, e) => canvas.Focus();

                    // Keyboard shortcuts
                    canvas.KeyDown += (s, e) =>
                    {
                        int cx = canvas.Width / 2, cy = canvas.Height / 2;
                        if      (e.KeyCode == Keys.Add      || e.KeyCode == Keys.Oemplus)   ZoomAt(cx, cy, 1.25);
                        else if (e.KeyCode == Keys.Subtract || e.KeyCode == Keys.OemMinus)  ZoomAt(cx, cy, 0.8);
                        else if (e.KeyCode == Keys.Home)  { CalcFit(); UpdateZoomLabel(); canvas.Invalidate(); }
                        else if (e.KeyCode == Keys.Escape) f.Close();
                    };

                    // Toolbar buttons
                    btnZoomIn .Click += (s, e) => ZoomAt(canvas.Width / 2, canvas.Height / 2, 1.25);
                    btnZoomOut.Click += (s, e) => ZoomAt(canvas.Width / 2, canvas.Height / 2, 0.8);
                    btnFit    .Click += (s, e) => { CalcFit(); UpdateZoomLabel(); canvas.Invalidate(); };
                    btn100    .Click += (s, e) =>
                    {
                        scale = 1.0;
                        offX  = (canvas.Width  - origBmp.Width)  / 2f;
                        offY  = (canvas.Height - origBmp.Height) / 2f;
                        ClampPan(); UpdateZoomLabel(); canvas.Invalidate();
                    };

                    f.Controls.Add(canvas);
                    f.FormClosed += (s, e) => { try { origBmp.Dispose(); } catch { } };

                    btnPrev.Visible = false; btnNext.Visible = false;
                    printAction = () =>
                    {
                        try
                        {
                            using var pd = new System.Drawing.Printing.PrintDocument();
                            pd.DefaultPageSettings.Margins =
                                new System.Drawing.Printing.Margins(10, 10, 10, 10);
                            pd.DefaultPageSettings.Landscape = true;
                            var dlg = new PrintDialog { Document = pd };
                            if (dlg.ShowDialog() != DialogResult.OK) return;
                            var bmp = origBmp;
                            pd.PrintPage += (ps, pe) =>
                            {
                                var g = pe.Graphics;
                                g.PageUnit = GraphicsUnit.Pixel;
                                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                float dpiX = g.DpiX, dpiY = g.DpiY;
                                var pb = pe.PageBounds;
                                float pageW = pb.Width  * dpiX / 100f;
                                float pageH = pb.Height * dpiY / 100f;
                                float pageL = pb.Left   * dpiX / 100f;
                                float pageT = pb.Top    * dpiY / 100f;
                                float sc = Math.Min(pageW / bmp.Width, pageH / bmp.Height);
                                float dw = bmp.Width * sc, dh = bmp.Height * sc;
                                g.DrawImage(bmp,
                                    pageL + (pageW - dw) / 2f,
                                    pageT + (pageH - dh) / 2f,
                                    dw, dh);
                                pe.HasMorePages = false;
                            };
                            pd.Print();
                        }
                        catch (Exception ex) { MessageBox.Show(_previewOwner, "Lỗi in: " + ex.Message, "In file"); }
                    };
                }
                else
                {
                    BuildPreviewInfoPanel(f, filePath);
                    btnZoomIn.Visible = btnZoomOut.Visible = btnFit.Visible = btn100.Visible = lblZoom.Visible = false;
                    printAction = () => { try { Process.Start(new ProcessStartInfo { FileName = filePath, Verb = "print", UseShellExecute = true }); } catch { MessageBox.Show(_previewOwner, "Không thể in file này.", "In file"); } };
                }
            }
            else if (ext == ".txt" || ext == ".csv" || ext == ".log" || ext == ".json" || ext == ".xml")
            {
                var tb = new TextBox
                {
                    Multiline = true, ReadOnly = true, Dock = DockStyle.Fill,
                    ScrollBars = ScrollBars.Both,
                    Font       = new Font("Consolas", 9.5f),
                    BackColor  = Color.FromArgb(30, 30, 30),
                    ForeColor  = Color.FromArgb(212, 212, 212),
                    BorderStyle = BorderStyle.None
                };
                var state = new PreviewState { Type = "text" };
                try
                {
                    var full = File.ReadAllText(filePath);
                    state.FullText = full.Replace("\r\n", "\n");
                    state.Lines    = state.FullText.Split('\n');
                    tb.Text = string.Join(Environment.NewLine, state.Lines.Take(state.TextPageLines));
                    lblInfo.Text = $"{Path.GetFileName(filePath)}  ({state.Lines.Length:N0} dòng)";
                }
                catch { tb.Text = "(Không thể đọc nội dung)"; }
                f.Controls.Add(tb);

                btnZoomIn.Visible = btnZoomOut.Visible = btnFit.Visible = btn100.Visible = lblZoom.Visible = false;
                btnPrev.Visible = btnNext.Visible = true;
                btnPrev.Enabled = false;
                btnNext.Enabled = state.Lines != null && state.Lines.Length > state.TextPageLines;

                void UpdatePageInfo() => lblInfo.Text =
                    $"{Path.GetFileName(filePath)}  — dòng {state.TextOffset + 1}–{Math.Min(state.Lines.Length, state.TextOffset + state.TextPageLines)} / {state.Lines.Length:N0}";

                btnPrev.Click += (s, e) =>
                {
                    state.TextOffset = Math.Max(0, state.TextOffset - state.TextPageLines);
                    tb.Text = string.Join(Environment.NewLine, state.Lines.Skip(state.TextOffset).Take(state.TextPageLines));
                    btnPrev.Enabled = state.TextOffset > 0;
                    btnNext.Enabled = state.TextOffset + state.TextPageLines < state.Lines.Length;
                    UpdatePageInfo();
                };
                btnNext.Click += (s, e) =>
                {
                    state.TextOffset = Math.Min(Math.Max(0, state.Lines.Length - state.TextPageLines), state.TextOffset + state.TextPageLines);
                    tb.Text = string.Join(Environment.NewLine, state.Lines.Skip(state.TextOffset).Take(state.TextPageLines));
                    btnPrev.Enabled = state.TextOffset > 0;
                    btnNext.Enabled = state.TextOffset + state.TextPageLines < state.Lines.Length;
                    UpdatePageInfo();
                };
                printAction = () =>
                {
                    try
                    {
                        using var pd = new System.Drawing.Printing.PrintDocument();
                        var dlg = new PrintDialog { Document = pd };
                        if (dlg.ShowDialog() != DialogResult.OK) return;
                        var lines = state.Lines ?? Array.Empty<string>();
                        int li = 0;
                        var pf = new Font("Consolas", 9);
                        pd.PrintPage += (ps, pe) =>
                        {
                            float y = pe.MarginBounds.Top, h = pf.GetHeight(pe.Graphics);
                            while (li < lines.Length)
                            {
                                if (y + h > pe.MarginBounds.Bottom) { pe.HasMorePages = true; return; }
                                pe.Graphics.DrawString(lines[li++], pf, Brushes.Black, pe.MarginBounds.Left, y);
                                y += h;
                            }
                            pe.HasMorePages = false;
                        };
                        pd.BeginPrint += (ps, pe) => li = 0;
                        pd.EndPrint   += (ps, pe) => pf.Dispose();
                        pd.Print();
                    }
                    catch (Exception ex) { MessageBox.Show(_previewOwner, "Lỗi in: " + ex.Message, "In file"); }
                };
            }
            else
            {
                BuildPreviewInfoPanel(f, filePath);
                btnZoomIn.Visible = btnZoomOut.Visible = btnFit.Visible = btn100.Visible = lblZoom.Visible = false;
                btnPrev.Visible = btnNext.Visible = false;
                printAction = () => { try { Process.Start(new ProcessStartInfo { FileName = filePath, Verb = "print", UseShellExecute = true }); } catch { MessageBox.Show(_previewOwner, "Không thể in file này.", "In file"); } };
            }

            try { StopPreviewCloseTimer(); AttachPreviewMouseHandlers(f); } catch { }
            _filePreviewForm = f;
            f.Show();
        }

        private void BuildPreviewInfoPanel(Form f, string filePath)
        {
            f.BackColor = Color.FromArgb(45, 45, 48);
            var pnl = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(45, 45, 48) };
            var pb  = new PictureBox { Left = 24, Top = 24, Width = 64, Height = 64, SizeMode = PictureBoxSizeMode.StretchImage };
            try { pb.Image = Icon.ExtractAssociatedIcon(filePath)?.ToBitmap(); } catch { }
            var lblName = new Label { Left = 24, Top = 96,  Width = 600, AutoSize = false, Height = 26, Text = Path.GetFileName(filePath), Font = new Font("Segoe UI", 11, FontStyle.Bold), ForeColor = Color.White };
            var lblSize = new Label { Left = 24, Top = 126, Width = 400, AutoSize = false, Height = 20, Text = $"Kích thước: {new FileInfo(filePath).Length:N0} bytes",       Font = new Font("Segoe UI", 9), ForeColor = Color.FromArgb(170, 170, 170) };
            var lblExt  = new Label { Left = 24, Top = 150, Width = 400, AutoSize = false, Height = 20, Text = $"Loại: {Path.GetExtension(filePath).ToUpperInvariant()}",     Font = new Font("Segoe UI", 9), ForeColor = Color.FromArgb(170, 170, 170) };
            var btnOp   = new Button { Left = 24, Top = 186, Width = 120, Height = 30, Text = "Mở file", FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White, Font = new Font("Segoe UI", 9, FontStyle.Bold), Cursor = Cursors.Hand };
            btnOp.FlatAppearance.BorderSize = 0;
            btnOp.Click += (s, e) => { try { Process.Start(new ProcessStartInfo { FileName = filePath, UseShellExecute = true }); } catch { } };
            pnl.Controls.AddRange(new Control[] { pb, lblName, lblSize, lblExt, btnOp });
            f.Controls.Add(pnl);
        }

        private void AttachPreviewMouseHandlers(Control ctl)
        {
            try
            {
                ctl.MouseEnter += (s, e) => StopPreviewCloseTimer();
                ctl.MouseLeave += (s, e) => StartPreviewCloseTimer();
                foreach (Control c in ctl.Controls)
                    AttachPreviewMouseHandlers(c);
            }
            catch { }
        }

        // Shell thumbnail extraction for files (PDF, Office, images)
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void SHCreateItemFromParsingName([MarshalAs(UnmanagedType.LPWStr)] string pszPath, IntPtr pbc, [In] ref Guid riid, out IntPtr ppv);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b")]
        private interface IShellItemImageFactory
        {
            void GetImage([In] SIZE size, [In] SIIGBF flags, out IntPtr phbm);
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SIZE { public int cx; public int cy; public SIZE(int x, int y) { cx = x; cy = y; } }

        [Flags]
        private enum SIIGBF : uint
        {
            SIIGBF_RESIZETOFIT = 0x00,
            SIIGBF_BIGGERSIZEOK = 0x01,
            SIIGBF_MEMORYONLY = 0x02,
            SIIGBF_ICONONLY = 0x04,
            SIIGBF_THUMBNAILONLY = 0x08,
            SIIGBF_INCACHEONLY = 0x10
        }

        private Image GetShellThumbnail(string path, int width, int height)
        {
            try
            {
                Guid guid = new Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b");
                SHCreateItemFromParsingName(path, IntPtr.Zero, ref guid, out IntPtr ppv);
                if (ppv == IntPtr.Zero) return null;
                var factory = (IShellItemImageFactory)Marshal.GetObjectForIUnknown(ppv);
                IntPtr hbm;
                factory.GetImage(new SIZE(Math.Max(1, width), Math.Max(1, height)), SIIGBF.SIIGBF_RESIZETOFIT, out hbm);
                Marshal.ReleaseComObject(factory);
                Marshal.Release(ppv);
                if (hbm == IntPtr.Zero) return null;
                var img = Image.FromHbitmap(hbm);
                DeleteObject(hbm);
                return img;
            }
            catch { return null; }
        }

        // ĐÃ KHẮC PHỤC LỖI CS7036: Chèn đúng danh sách _poList
        private void BtnSearchBySupp_Click(object sender, EventArgs e)
        {
            try
            {
                var suppliers = new SupplierService().GetAll();
                // Nạp ProjectCode (Mã dự án) từ ProjectService để hiển thị trong cột "Mã DA"
                var projects = new ProjectService().GetAll();
                var dt = new System.Data.DataTable();
                dt.Columns.Add("PO_ID", typeof(int));
                dt.Columns.Add("PO No", typeof(string));
                dt.Columns.Add("NCC", typeof(string));
                dt.Columns.Add("Mã DA", typeof(string));   // Mã dự án (ProjectCode)
                dt.Columns.Add("Dự án", typeof(string));   // Tên dự án (ẩn, dùng để filter nội bộ)
                dt.Columns.Add("MPR No", typeof(string));
                dt.Columns.Add("Workorder", typeof(string));
                dt.Columns.Add("Ngày PO", typeof(string));
                dt.Columns.Add("Trạng thái", typeof(string));
                dt.Columns.Add("Tổng tiền", typeof(string));
                foreach (var h in _poList)
                {
                    var supp = suppliers.Find(s => s.Supplier_ID == h.Supplier_ID);
                    // Tìm ProjectCode theo WorkorderNo hoặc ProjectName
                    var prj = projects.Find(p =>
                        (!string.IsNullOrEmpty(p.WorkorderNo) && p.WorkorderNo.Equals(h.WorkorderNo, StringComparison.OrdinalIgnoreCase)) ||
                        (!string.IsNullOrEmpty(p.ProjectName) && p.ProjectName.Equals(h.Project_Name, StringComparison.OrdinalIgnoreCase)));
                    string maDA = prj?.ProjectCode ?? "";
                    dt.Rows.Add(h.PO_ID, h.PONo, supp?.Company_Name ?? supp?.Short_Name ?? "",
                        maDA, h.Project_Name, h.MPR_No, h.WorkorderNo,
                        h.PO_Date.HasValue ? h.PO_Date.Value.ToString("dd/MM/yyyy") : "",
                        h.Status, h.Total_Amount.ToString("N2", _numCulture));
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
                pF.Controls.Add(new Label { Text = "Mã DA:", Location = new Point(255, 12), Size = new Size(45, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold) });
                var btnDropDA = new Button
                {
                    Location = new Point(302, 8),
                    Size = new Size(220, 26),
                    Text = "Tất cả mã DA  ▼",
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

                // ── Danh sách mã dự án unique (ProjectCode) ──
                var projectList = dtFull.AsEnumerable()
                    .Select(r => r["Mã DA"]?.ToString() ?? "")
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
                        ? "Tất cả mã DA  ▼"
                        : (sel.Count == 1 ? sel[0] + "  ▼" : $"{sel.Count} mã DA đã chọn  ▼");
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

                // ── Đóng dropdown khi click ra ngoài hoặc focus rời khỏi panel ──
                // 1. Leave của clbDA (khi focus chuyển sang control khác)
                clbDA.Leave += (s, ev) =>
                {
                    // Chỉ đóng nếu focus KHÔNG đến btnDropDA (để toggle vẫn hoạt động)
                    var focused = popup.ActiveControl;
                    if (focused != btnDropDA)
                        panelDropDA.Visible = false;
                };

                // 2. MouseDown trên popup (click vào vùng trống)
                popup.MouseDown += (s, ev) =>
                {
                    if (panelDropDA.Visible && !panelDropDA.Bounds.Contains(ev.Location))
                        panelDropDA.Visible = false;
                };

                // 3. Hook MouseDown vào tất cả control con để bắt click ở mọi vị trí
                Action<Control> hookClose = null!;
                hookClose = (ctrl) =>
                {
                    if (ctrl == panelDropDA || ctrl == btnDropDA) return;
                    ctrl.MouseDown += (s, ev) =>
                    {
                        if (panelDropDA.Visible)
                            panelDropDA.Visible = false;
                    };
                    foreach (Control child in ctrl.Controls)
                        hookClose(child);
                };
                // Gọi hook sau khi tất cả controls đã được thêm vào popup
                popup.Shown += (s, ev) =>
                {
                    foreach (Control ctrl in popup.Controls)
                        hookClose(ctrl);
                };

                var lblCount = new Label { Text = "", Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212), Location = new Point(10, 38 + PF_H + 4), Size = new Size(700, 20), Anchor = AnchorStyles.Top | AnchorStyles.Left };
                popup.Controls.Add(lblCount);

                int DGV_TOP = 38 + PF_H + 28;
                var dgv = new DataGridView { Location = new Point(10, DGV_TOP), Size = new Size(popup.ClientSize.Width - 20, popup.ClientSize.Height - DGV_TOP - 50), ReadOnly = false, AllowUserToAddRows = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, BackgroundColor = Color.White, BorderStyle = BorderStyle.FixedSingle, RowHeadersVisible = false, Font = new Font("Segoe UI", 9), AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom };
                dgv.CellBeginEdit += (s, ev) => { if (dgv.Columns[ev.ColumnIndex].DataPropertyName != "Chọn") ev.Cancel = true; };
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
                        if (checkedProjects.Count > 0 && !checkedProjects.Contains(r["Mã DA"].ToString())) return false;
                        return true;
                    });
                    dtCurrent = rows.Any() ? rows.CopyToDataTable() : dtFull.Clone();
                    if (!dtCurrent.Columns.Contains("Chọn"))
                        dtCurrent.Columns.Add("Chọn", typeof(bool));
                    foreach (System.Data.DataRow row in dtCurrent.Rows) row["Chọn"] = false;
                    dgv.DataSource = dtCurrent;
                    if (dgv.Columns.Contains("PO_ID")) dgv.Columns["PO_ID"].Visible = false;
                    if (dgv.Columns.Contains("Dự án")) dgv.Columns["Dự án"].Visible = false;
                    var chonCol = dgv.Columns.Cast<DataGridViewColumn>()
                        .FirstOrDefault(c => c.DataPropertyName == "Chọn");
                    if (chonCol != null)
                    {
                        chonCol.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                        chonCol.Width = 40;
                        chonCol.HeaderText = "✓";
                        chonCol.ReadOnly = false;
                        chonCol.DisplayIndex = 0; // đặt cuối cùng để tránh reorder làm null ref
                    }
                    foreach (DataGridViewColumn col in dgv.Columns)
                        if (col.DataPropertyName != "Chọn") col.ReadOnly = true;

                    int total = dtFull.Rows.Count, shown = dtCurrent.Rows.Count;
                    string daInfo = checkedProjects.Count == 0 ? "tất cả mã DA" : $"{checkedProjects.Count} mã DA";
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
                        // KHÔNG đóng dropdown ở đây — người dùng có thể tick nhiều mã
                    }));
                };

                // ── Đóng dropdown khi nhấn Enter trong clbDA ──
                clbDA.KeyDown += (s, ev) =>
                {
                    if (ev.KeyCode == Keys.Enter || ev.KeyCode == Keys.Escape)
                    {
                        panelDropDA.Visible = false;
                        ev.Handled = true;
                    }
                };

                // ── Lọc danh sách mã DA theo keyword NCC đang nhập ──
                Action updateProjectList = () =>
                {
                    string kNCC = txtNCC.Text.Trim().ToLower();
                    // Lấy các mã DA của PO có NCC khớp keyword (hoặc tất cả nếu NCC trống)
                    var filteredMaDA = dtFull.AsEnumerable()
                        .Where(r => string.IsNullOrEmpty(kNCC) ||
                                    r["NCC"].ToString().ToLower().Contains(kNCC))
                        .Select(r => r["Mã DA"]?.ToString() ?? "")
                        .Where(v => !string.IsNullOrWhiteSpace(v))
                        .Distinct()
                        .OrderBy(v => v)
                        .ToList();

                    // Ghi nhớ các mã đang được tick để khôi phục sau khi rebuild list
                    var checked_ = clbDA.CheckedItems.Cast<string>().ToList();
                    clbDA.Items.Clear();
                    foreach (var m in filteredMaDA)
                        clbDA.Items.Add(m, checked_.Contains(m));

                    // Resize panelDropDA theo số lượng mục mới
                    int h = Math.Min(Math.Max(filteredMaDA.Count, 1), 10) * 18 + 8;
                    panelDropDA.Height = h + 2;
                    updateDropBtn();
                };

                // Gọi updateProjectList mỗi khi txtNCC thay đổi
                txtNCC.TextChanged += (s, ev) =>
                {
                    updateProjectList();
                    applyFilter();
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
                btnSel.Click += (s, ev) => { if (dgv.SelectedRows.Count == 0) { SafeWarn("Vui lòng chọn một dòng!"); return; } selectedPONo = dgv.SelectedRows[0].Cells["PO No"].Value?.ToString(); popup.DialogResult = DialogResult.OK; popup.Close(); };
                popup.Controls.Add(btnSel);

                var btnSelAll = new Button { Text = "☑ Chọn tất cả", Location = new Point(150, btnY), Size = new Size(130, 32), BackColor = Color.FromArgb(255, 140, 0), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
                btnSelAll.FlatAppearance.BorderSize = 0;
                btnSelAll.Click += (s, ev) =>
                {
                    if (dtCurrent == null) return;
                    bool allChecked = dtCurrent.Rows.Cast<System.Data.DataRow>().All(r => r["Chọn"] is bool b && b);
                    foreach (System.Data.DataRow row in dtCurrent.Rows) row["Chọn"] = !allChecked;
                    btnSelAll.Text = allChecked ? "☑ Chọn tất cả" : "☐ Bỏ chọn tất cả";
                    dgv.Invalidate();
                };
                popup.Controls.Add(btnSelAll);

                var btnExp = new Button { Text = "📥 Xuất Excel", Location = new Point(290, btnY), Size = new Size(130, 32), BackColor = Color.FromArgb(0, 150, 100), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
                btnExp.FlatAppearance.BorderSize = 0;
                btnExp.Click += async (s, ev) =>
                {
                    var selectedRows = dtCurrent?.Rows.Cast<System.Data.DataRow>().Where(r => r["Chọn"] is bool b && b).ToList();
                    if (selectedRows == null || selectedRows.Count == 0) { SafeWarn("Vui lòng chọn ít nhất một PO để xuất!"); return; }
                    using var sfd = new SaveFileDialog { Title = "Xuất chi tiết PO theo NCC", Filter = "Excel Files|*.xlsx", FileName = $"PO_ChiTiet_{DateTime.Now:yyyyMMdd_HHmm}", InitialDirectory = Directory.Exists(@"D:\RÁC") ? @"D:\RÁC" : Environment.GetFolderPath(Environment.SpecialFolder.Desktop) };
                    if (sfd.ShowDialog() != DialogResult.OK) return;

                    // Snapshot dữ liệu cần trước khi vào Task.Run
                    string nccTitle = txtNCC.Text.Trim();
                    if (string.IsNullOrEmpty(nccTitle)) nccTitle = "Tất cả NCC";
                    string savePath = sfd.FileName;
                    bool isDomestic = _isDomestic;
                    int chiTietDec = isDomestic ? 0 : 2;
                    string chiTietFmt = isDomestic ? "#,##0" : "#,##0.00";
                    var rowSnapshots = selectedRows.Select(dr => new {
                        PoId = Convert.ToInt32(dr["PO_ID"]),
                        PoNo = dr["PO No"]?.ToString() ?? "",
                        Ncc  = dr["NCC"]?.ToString() ?? "",
                        DuAn = dr["Dự án"]?.ToString() ?? "",
                        MprNo = dr["MPR No"]?.ToString() ?? "",
                        NgayPO = dr["Ngày PO"]?.ToString() ?? "",
                        TrangThai = dr["Trạng thái"]?.ToString() ?? ""
                    }).ToList();

                    btnExp.Enabled = false;
                    popup.Cursor = Cursors.WaitCursor;
                    try
                    {
                        int totalRows = 0;
                        string errMsg = null;

                        await Task.Run(() =>
                        {
                            try
                            {
                                const int TOTAL_COLS = 12;
                                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                                using var pkg = new OfficeOpenXml.ExcelPackage();
                                var ws = pkg.Workbook.Worksheets.Add("Chi tiết PO theo NCC");

                                // Dòng 1: Tiêu đề
                                ws.Cells[1, 1].Value = $"CHI TIẾT ĐẶT HÀNG THEO NHÀ CUNG CẤP — {nccTitle}";
                                ws.Cells[1, 1, 1, TOTAL_COLS].Merge = true;
                                ws.Cells[1, 1].Style.Font.Bold = true; ws.Cells[1, 1].Style.Font.Size = 13;
                                ws.Cells[1, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(102, 51, 153));
                                ws.Cells[1, 1].Style.Font.Color.SetColor(Color.White);
                                ws.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                // Dòng 2: Thông tin xuất
                                ws.Cells[2, 1].Value = $"Tổng PO: {rowSnapshots.Count}   |   Xuất ngày: {DateTime.Now:dd/MM/yyyy HH:mm}";
                                ws.Cells[2, 1, 2, TOTAL_COLS].Merge = true;
                                ws.Cells[2, 1].Style.Font.Italic = true; ws.Cells[2, 1].Style.Font.Size = 9;
                                ws.Cells[2, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                ws.Cells[2, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(240, 230, 255));

                                // Dòng 3: Header cột
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
                                foreach (var r in rowSnapshots)
                                {
                                    // Dòng nhóm PO
                                    ws.Cells[curRow, 1].Value = $"  🛒  PO: {r.PoNo}   |   NCC: {r.Ncc}   |   Dự án: {r.DuAn}   |   MPR: {r.MprNo}   |   Ngày PO: {r.NgayPO}   |   Trạng thái: {r.TrangThai}";
                                    var poHdrRange = ws.Cells[curRow, 1, curRow, TOTAL_COLS];
                                    poHdrRange.Style.Font.Bold = true;
                                    poHdrRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    poHdrRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(102, 51, 153));
                                    poHdrRange.Style.Font.Color.SetColor(Color.White);
                                    poHdrRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
                                    ws.Row(curRow).Height = 18;
                                    curRow++;

                                    // Load details từ DB (đang ở background thread — an toàn)
                                    var dets = new POService().GetDetails(r.PoId);
                                    decimal sub = 0;
                                    for (int i = 0; i < dets.Count; i++)
                                    {
                                        var d = dets[i];
                                        decimal rp = d.Price, bv = d.Qty_Per_Sheet;
                                        if ((d.Remarks ?? "").Contains("[CALC:KG]") && d.Weight_kg > 0 && d.Qty_Per_Sheet > 0)
                                        { rp = Math.Round((d.Price * d.Qty_Per_Sheet) / d.Weight_kg, 2); bv = d.Weight_kg; }
                                        decimal amt = Math.Round(bv * rp, chiTietDec, MidpointRounding.AwayFromZero); sub += amt;

                                        ws.Cells[curRow, 1].Value = i + 1;
                                        ws.Cells[curRow, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                        ws.Cells[curRow, 2].Value = d.Item_Name ?? "";
                                        ws.Cells[curRow, 3].Value = d.Material ?? "";
                                        ws.Cells[curRow, 4].Value = d.Asize; ws.Cells[curRow, 4].Style.Numberformat.Format = "#,##0.##";
                                        ws.Cells[curRow, 5].Value = d.Bsize; ws.Cells[curRow, 5].Style.Numberformat.Format = "#,##0.##";
                                        ws.Cells[curRow, 6].Value = d.Csize; ws.Cells[curRow, 6].Style.Numberformat.Format = "#,##0.##";
                                        ws.Cells[curRow, 7].Value = d.Qty_Per_Sheet; ws.Cells[curRow, 7].Style.Numberformat.Format = "#,##0";
                                        ws.Cells[curRow, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                        ws.Cells[curRow, 8].Value = d.UNIT ?? "";
                                        ws.Cells[curRow, 8].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                        ws.Cells[curRow, 9].Value = d.Weight_kg; ws.Cells[curRow, 9].Style.Numberformat.Format = "#,##0.00";
                                        ws.Cells[curRow, 9].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                                        ws.Cells[curRow, 10].Value = rp; ws.Cells[curRow, 10].Style.Numberformat.Format = "#,##0.00";
                                        ws.Cells[curRow, 10].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                                        ws.Cells[curRow, 11].Value = d.VAT; ws.Cells[curRow, 11].Style.Numberformat.Format = "#,##0";
                                        ws.Cells[curRow, 11].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                        ws.Cells[curRow, 12].Value = amt; ws.Cells[curRow, 12].Style.Numberformat.Format = chiTietFmt;
                                        ws.Cells[curRow, 12].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                                        if (i % 2 == 1)
                                            for (int c = 1; c <= TOTAL_COLS; c++)
                                            { ws.Cells[curRow, c].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid; ws.Cells[curRow, c].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(240, 248, 255)); }

                                        ws.Cells[curRow, 1, curRow, TOTAL_COLS].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Hair);
                                        curRow++; totalRows++;
                                    }

                                    decimal vatAmt = isDomestic ? Math.Round(sub * 0.1m, 0, MidpointRounding.AwayFromZero) : 0m;
                                    ws.Cells[curRow, 11].Value = "Sub-Total:"; ws.Cells[curRow, 11].Style.Font.Bold = true;
                                    ws.Cells[curRow, 12].Value = sub; ws.Cells[curRow, 12].Style.Numberformat.Format = chiTietFmt; ws.Cells[curRow, 12].Style.Font.Bold = true;
                                    ws.Cells[curRow, 12].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right; curRow++;
                                    ws.Cells[curRow, 11].Value = "VAT (10%):";
                                    ws.Cells[curRow, 12].Value = vatAmt; ws.Cells[curRow, 12].Style.Numberformat.Format = chiTietFmt;
                                    ws.Cells[curRow, 12].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right; curRow++;
                                    ws.Cells[curRow, 10, curRow, 11].Merge = true; ws.Cells[curRow, 10].Value = "TOTAL (incl. VAT):"; ws.Cells[curRow, 10].Style.Font.Bold = true;
                                    ws.Cells[curRow, 12].Value = sub + vatAmt; ws.Cells[curRow, 12].Style.Numberformat.Format = chiTietFmt; ws.Cells[curRow, 12].Style.Font.Bold = true;
                                    ws.Cells[curRow, 12].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                                    curRow += 2;
                                }

                                ws.Cells[ws.Dimension.Address].AutoFitColumns();
                                ws.Column(2).Width = Math.Min(ws.Column(2).Width, 45);
                                pkg.SaveAs(new System.IO.FileInfo(savePath));
                            }
                            catch (Exception ex2) { errMsg = ex2.Message; }
                        });

                        if (errMsg != null)
                        {
                            MessageBox.Show(GetActiveOwner(), "Lỗi xuất Excel: " + errMsg, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else
                        {
                            if (lblStatus != null)
                                lblStatus.Text = $"✅ Đã xuất {rowSnapshots.Count} PO — {totalRows} dòng chi tiết";
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = savePath, UseShellExecute = true });
                        }
                    }
                    catch (Exception ex2) { MessageBox.Show(GetActiveOwner(), "Lỗi xuất Excel: " + ex2.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                    finally
                    {
                        btnExp.Enabled = true;
                        popup.Cursor = Cursors.Default;
                    }
                };
                popup.Controls.Add(btnExp);

                var btnClose = new Button { Text = "Đóng", Size = new Size(100, 32), BackColor = Color.FromArgb(108, 117, 125), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Anchor = AnchorStyles.Bottom | AnchorStyles.Right, DialogResult = DialogResult.Cancel };
                btnClose.FlatAppearance.BorderSize = 0; btnClose.Location = new Point(popup.ClientSize.Width - 115, btnY);
                popup.Controls.Add(btnClose); popup.CancelButton = btnClose;

                popup.Resize += (s, ev) => { pF.Width = popup.ClientSize.Width - 20; dgv.Size = new Size(popup.ClientSize.Width - 20, popup.ClientSize.Height - 165); btnSel.Location = new Point(10, popup.ClientSize.Height - 42); btnSelAll.Location = new Point(150, popup.ClientSize.Height - 42); btnExp.Location = new Point(290, popup.ClientSize.Height - 42); btnClose.Location = new Point(popup.ClientSize.Width - 115, popup.ClientSize.Height - 42); };
                if (popup.ShowDialog(this) == DialogResult.OK && !string.IsNullOrEmpty(selectedPONo)) SelectPOByNo(selectedPONo);
            }
            catch (Exception ex) { MessageBox.Show(GetActiveOwner(), "Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void AutoAdjustColumnWidths()
        {
            if (dgvDetails.Columns.Count == 0) return;
            // Dùng DisplayedCells thay vì AllCells — chỉ đo cells đang hiển thị, tránh block UI
            dgvDetails.SuspendLayout();
            try
            {
                foreach (DataGridViewColumn col in dgvDetails.Columns)
                {
                    if (!col.Visible) continue;
                    if (col.Name == "NhapKho") continue;
                    if (col.Name == "Asize" || col.Name == "Bsize" || col.Name == "Csize")
                    {
                        col.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
                        col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                        continue;
                    }
                    col.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCellsExceptHeader;
                    int measured = col.Width;
                    col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    col.Width = Math.Min(Math.Max(measured, 40), 220);
                }
            }
            finally { dgvDetails.ResumeLayout(); }
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
        {
            if (e.RowIndex < 0) return;
            string colName = dgvDetails.Columns[e.ColumnIndex].Name;
            // Tính lại khi thay đổi các cột ảnh hưởng đến thành tiền
            if (colName == "Price" || colName == "Qty" || colName == "Weight"
                || colName == "VAT" || colName == "Calc_Method")
                RecalculateAmount(e.RowIndex);
        }

        private void DgvDetails_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string colName = dgvDetails.Columns[e.ColumnIndex].Name;

            if (colName == "Nhập kho")
            {
                e.CellStyle.BackColor = Color.FromArgb(220, 252, 231);
                e.CellStyle.ForeColor = Color.FromArgb(22, 101, 52);
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                e.CellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                e.FormattingApplied = true;
            }

            // Cột Đơn giá, TT trước thuế, Thành tiền
            if (colName == "Price" || colName == "SubAmount")
            {
                if (e.Value != null && decimal.TryParse(e.Value.ToString(), out decimal num))
                {
                    string fmt = (colName == "SubAmount" && _isDomestic) ? "N0" : "N2";
                    e.Value = num.ToString(fmt, _numCulture);
                    e.FormattingApplied = true;
                }
                e.CellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
            if (colName == "Amount")
            {
                if (e.Value != null && decimal.TryParse(e.Value.ToString(), out decimal numAmt))
                {
                    string fmtAmt = _isDomestic ? "N0" : "N2";
                    e.Value = numAmt.ToString(fmtAmt, _numCulture);
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
        { try { _supplierTable = new SupplierService().GetForCombo(); BindSupplierCombo(_supplierTable); } catch (Exception ex) { MessageBox.Show(GetActiveOwner(), "Lỗi tải nhà cung cấp: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); } }

        private async Task LoadSupplierComboAsync()
        {
            try
            {
                var dt = await Task.Run(() => new SupplierService().GetForCombo());
                _supplierTable = dt;
                BindSupplierCombo(_supplierTable);
            }
            catch (Exception ex)
            {
                MessageBox.Show(GetActiveOwner(), "Lỗi tải nhà cung cấp: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

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
            else
            {
                BindSupplierCombo(filtered);
            }
            _isSearching = true; cboSupplier.Text = keyword;
            cboSupplier.SelectionStart = keyword.Length; 
            cboSupplier.DroppedDown = true; _isSearching = false;
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
        //    if (_selectedPO_ID == 0) { SafeWarn("Vui lòng chọn PO cần xuất!"); return; }
        //    try
        //    {
        //        var po = _poList.Find(p => p.PO_ID == _selectedPO_ID);
        //        var details = _service.GetDetails(_selectedPO_ID); if (po == null) return;
        //        var suppliers = new SupplierService().GetAll();
        //        var supplier = suppliers.Find(s => s.Supplier_ID == Convert.ToInt32(cboSupplier.SelectedValue?.ToString() ?? "0"));
        //        var projects = new ProjectService().GetAll();
        //        var project = projects.Find(p => p.WorkorderNo == po.WorkorderNo);
        //        string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "po_template.xlsx");
        //        if (!File.Exists(templatePath)) { MessageBox.Show(GetActiveOwner(), $"Lỗi: Không tìm thấy file template!\nĐường dẫn dự kiến: {templatePath}", "Lỗi hệ thống", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
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
        //        var result = MessageBox.Show(GetActiveOwner(), $"✅ Xuất Excel thành công!\nFile: {saveDialog.FileName}\n\nBạn có muốn mở file không?", "Thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
        //        if (result == DialogResult.Yes) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = saveDialog.FileName, UseShellExecute = true });
        //    }
        //    catch (Exception ex) { MessageBox.Show(GetActiveOwner(), "Lỗi xuất Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        //}

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("PO", "Xuất Excel", "Xuất Excel")) return;
            if (_selectedPO_ID == 0) { SafeWarn("Vui lòng chọn PO cần xuất!"); return; }
            try
            {
                var po = _poList.Find(p => p.PO_ID == _selectedPO_ID);
                var details = _service.GetDetails(_selectedPO_ID); if (po == null) return;
                var suppliers = new SupplierService().GetAll();
                var supplier = suppliers.Find(s => s.Supplier_ID == Convert.ToInt32(cboSupplier.SelectedValue?.ToString() ?? "0"));
                var projects = new ProjectService().GetAll();
                var project = projects.Find(p => p.WorkorderNo == po.WorkorderNo);

                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "po_template.xlsx");
                if (!File.Exists(templatePath)) { MessageBox.Show(GetActiveOwner(), $"Lỗi: Không tìm thấy template!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

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
                    string fullName = !string.IsNullOrWhiteSpace(AppSession.CurrentUser?.Full_Name)
                        ? AppSession.CurrentUser.Full_Name
                        : !string.IsNullOrWhiteSpace(_currentUser) ? _currentUser
                        : AppSession.CurrentUser?.Username ?? "Admin";
                    ReplaceCell(ws, "<<USER_NAME>>", fullName);
                    ReplaceCell(ws, "<<user name>>", fullName);
                    ReplaceCell(ws, "<<User Name>>", fullName);

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
                            ws.Cells[startRow, 1, startRow, 17].Copy(ws.Cells[startRow + i, 1]);
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
                        else if (rem.Contains("[CALC:SL]"))
                        {
                            rem = rem.Replace("[CALC:SL]", "").Trim();
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
                    ws.Cells[subTotalRow, 9].Formula = $"=SUM(I{startRow}:I{startRow + detailCount - 1})";
                    ws.Cells[subTotalRow, 9].Style.Numberformat.Format = "#,##0.00";
                    ws.Cells[subTotalRow, 14].Formula = $"=SUM(N{startRow}:N{startRow + detailCount - 1})";
                    ws.Cells[subTotalRow, 14].Style.Numberformat.Format = "#,##0.00";

                    // Gán nhãn cho dòng tổng thanh toán (đã bao gồm các mức thuế khác nhau)
                    ws.Cells[vatRow, 3].Value = "Final Price Requested (Included VAT)";

                    // 3. CẬP NHẬT: Gán giá trị tổng sau thuế đã tính toán vào ô kết quả
                    // Chúng ta gán giá trị số trực tiếp và định dạng hiển thị cho chuyên nghiệp
                    ws.Cells[vatRow, 14].Value = totalAfterVAT;
                    ws.Cells[vatRow, 14].Style.Numberformat.Format = "#,##0.00"; // Định dạng dấu phân cách hàng nghìn
                    ws.Cells[vatRow, 14].Style.Font.Bold = true;

                    package.Save();
                }

                if (MessageBox.Show(GetActiveOwner(), "✅ Xuất Excel thành công! Bạn có muốn mở file?", "Thành công", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = saveDialog.FileName, UseShellExecute = true });
            }
            catch (Exception ex) { MessageBox.Show(GetActiveOwner(), "Lỗi: " + ex.Message); }
        }

        //private void BtnExport_Click(object sender, EventArgs e)
        //{
        //    if (_selectedPO_ID == 0) { SafeWarn("Vui lòng chọn PO cần xuất!"); return; }
        //    try
        //    {
        //        var po = _poList.Find(p => p.PO_ID == _selectedPO_ID);
        //        var details = _service.GetDetails(_selectedPO_ID); if (po == null) return;
        //        var suppliers = new SupplierService().GetAll();
        //        var supplier = suppliers.Find(s => s.Supplier_ID == Convert.ToInt32(cboSupplier.SelectedValue?.ToString() ?? "0"));
        //        var projects = new ProjectService().GetAll();
        //        var project = projects.Find(p => p.WorkorderNo == po.WorkorderNo);
        //        string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "po_template.xlsx");

        //        if (!File.Exists(templatePath)) { MessageBox.Show(GetActiveOwner(), $"Lỗi: Không tìm thấy file template!\nĐường dẫn dự kiến: {templatePath}", "Lỗi hệ thống", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

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

        //        var result = MessageBox.Show(GetActiveOwner(), $"✅ Xuất Excel thành công!\nFile: {saveDialog.FileName}\n\nBạn có muốn mở file không?", "Thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
        //        if (result == DialogResult.Yes) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = saveDialog.FileName, UseShellExecute = true });
        //    }
        //    catch (Exception ex) { MessageBox.Show(GetActiveOwner(), "Lỗi xuất Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        //}

        private void ReplaceCell(OfficeOpenXml.ExcelWorksheet ws, string placeholder, string value)
        {
            if (ws.Dimension == null) return;
            foreach (var cell in ws.Cells[ws.Dimension.Address])
            {
                if (cell.IsRichText)
                {
                    // Rich-text cell: thay trong từng run
                    foreach (var run in cell.RichText)
                    {
                        if (run.Text.Contains(placeholder, StringComparison.OrdinalIgnoreCase))
                            run.Text = run.Text.Replace(placeholder, value, StringComparison.OrdinalIgnoreCase);
                    }
                }
                else
                {
                    string cellVal = cell.Value?.ToString() ?? "";
                    if (cellVal.Contains(placeholder, StringComparison.OrdinalIgnoreCase))
                        cell.Value = cellVal.Replace(placeholder, value, StringComparison.OrdinalIgnoreCase);
                }
            }
        }


        private void BuildDetailColumns()
        {
            dgvDetails.Columns.Clear();
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "DeliveryLocation", HeaderText = "Nơi giao" });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Item_No", HeaderText = "STT", ReadOnly = true });
            // Freeze 2 cột đầu: đặt Frozen=true trên cột cuối cùng trong nhóm freeze
            dgvDetails.Columns["DeliveryLocation"].Frozen = true;
            dgvDetails.Columns["Item_No"].Frozen = true;
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
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Received", HeaderText = "Đã nhận", Visible = false });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "MPSNo", HeaderText = "MPS No", Visible = false });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Remarks", HeaderText = "Ghi chú" });

            // Cách tính — mặc định Theo KG
            var colCalc = new DataGridViewComboBoxColumn { Name = "Calc_Method", HeaderText = "Cách tính" };
            colCalc.Items.AddRange("Theo KG", "Theo SL");
            dgvDetails.Columns.Add(colCalc);

            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ordered_PO", HeaderText = "Đã lên PO", ReadOnly = true });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "NhapKho", HeaderText = "Nhập kho", Visible = true, ReadOnly = true });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "PO_Detail_ID", HeaderText = "PO_ID", Visible = false });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "MPR_Detail_ID", HeaderText = "MPR_Detail_ID", Visible = false });

            // Header: wrap text, tự co giãn chiều cao
            dgvDetails.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvDetails.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;

            foreach (DataGridViewColumn col in dgvDetails.Columns) col.Width = 60;

            // A/B/C: tự co giãn theo nội dung (AutoAdjustColumnWidths sẽ bỏ qua 3 cột này)
            dgvDetails.Columns["Asize"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgvDetails.Columns["Bsize"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgvDetails.Columns["Csize"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

            // Nhập kho: căn giữa và hiển thị 2 chữ số thập phân
            if (dgvDetails.Columns.Contains("NhapKho"))
            {
                dgvDetails.Columns["NhapKho"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgvDetails.Columns["NhapKho"].DefaultCellStyle.Format = "N2";
                dgvDetails.Columns["NhapKho"].DefaultCellStyle.NullValue = "";
                dgvDetails.Columns["NhapKho"].Width = 75;
            }
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

        // =====================================================================
        // SCALE BUTTONS — thay đổi kích thước buttons theo tỉ lệ chiều rộng form
        // =====================================================================
        private const int _baseFormW = 1280; // Chiều rộng thiết kế gốc (không scale khi form rộng hơn)
        private void ScaleButtons()
        {
            if (_mainScroll == null) return;

            // Chỉ thu nhỏ khi form nhỏ hơn kích thước gốc.
            // Khi form >= 1280px → ratio = 1.0 → giữ nguyên kích thước gốc.
            double ratio = Math.Min(1.0, (double)this.ClientSize.Width / _baseFormW);
            // Clamp đáy 0.60 để button không bị quá nhỏ
            ratio = Math.Max(0.60, ratio);

            // Kích thước gốc của từng button (width, height)
            var btnSizes = new (Button btn, int baseW, int baseH)[]
            {
                (btnSearch,       70, 30),
                (btnNewPO,       100, 30),
                (btnDeletePO,     90, 30),
                (btnSearchBySupp,130, 30),
                (btnSavePO,      150, 32),
                (btnClearHeader, 100, 32),
                (btnAddDetail,   120, 30),
                (btnDeleteDetail,100, 30),
                (btnExport,      130, 30),
            };
            foreach (var (btn, bw, bh) in btnSizes)
            {
                if (btn == null) continue;
                btn.Width = (int)(bw * ratio);
                btn.Height = (int)(bh * ratio);
                btn.Font = new Font("Segoe UI", (float)Math.Max(7.0, 9.0 * ratio), FontStyle.Bold);
            }

            // Scale các buttons trong FlowLayoutPanel
            ScaleFlowPanelButtons(_mainScroll, ratio);
        }

        private void ScaleFlowPanelButtons(Control parent, double ratio)
        {
            foreach (Control c in parent.Controls)
            {
                if (c is FlowLayoutPanel flp)
                {
                    foreach (Control child in flp.Controls)
                    {
                        if (child is Button b)
                        {
                            // Lấy kích thước gốc từ Tag (lưu một lần khi tạo)
                            if (b.Tag == null)
                                b.Tag = $"{b.Width},{b.Height}";
                            var parts = b.Tag.ToString().Split(',');
                            if (parts.Length == 2 &&
                                int.TryParse(parts[0], out int ow) &&
                                int.TryParse(parts[1], out int oh))
                            {
                                b.Width = (int)(ow * ratio);
                                b.Height = (int)(oh * ratio);
                                b.Font = new Font("Segoe UI",
                                    (float)Math.Max(7.0, 9.0 * ratio), FontStyle.Bold);
                            }
                        }
                    }
                }
                else if (c.Controls.Count > 0)
                    ScaleFlowPanelButtons(c, ratio);
            }
        }
        private void FrmPO_Resize(object sender, EventArgs e)
        {
            if (_mainScroll == null) return;

            // ── Chiều rộng nội dung: tối thiểu 1280px ──
            // Nếu form hẹp hơn → _mainScroll tự thêm scrollbar ngang
            const int minContentW = 1280;
            int formW = this.ClientSize.Width - 2;
            int w = Math.Max(formW, minContentW);

            panelTop.Width = w;

            // panelHeader: FlowLayout tự wrap → Height tự điều chỉnh
            panelHeader.Width = w;
            panelHeader.Top = panelTop.Bottom + 10;
            panelHeader.PerformLayout();

            // panelMPRFiles: 280px cố định bên phải
            var panelMPRFiles = _mainScroll.Controls.OfType<Panel>()
                .FirstOrDefault(p => dgvMPRFiles != null &&
                    p.Controls.OfType<DataGridView>().Any(d => d == dgvMPRFiles));
            int mprW = 280;

            panelDetail.Top = panelHeader.Bottom + 10;
            panelDetail.Width = w - mprW - 10;

            // ── Đồng bộ chiều cao dgvFiles theo panelHeader ──
            if (dgvFiles != null)
            {
                dgvFiles.Height = panelHeader.Height - 20;
                dgvFiles.Location = new Point(panelHeader.Width - dgvFiles.Width - 10, 10);
            }

            // dgvDetails.Top: ngay dưới flowDetailBtns
            var flowDetailBtns = panelDetail.Controls.OfType<FlowLayoutPanel>().FirstOrDefault();
            int dgvTop = flowDetailBtns != null ? flowDetailBtns.Bottom + 4 : 75;
            dgvDetails.Top = dgvTop;
            dgvDetails.Width = panelDetail.Width - 20;

            // ── panelDetail.Height: cố định đủ để dgvDetails luôn hiển thị toàn bộ ──
            // Không co theo form → _mainScroll tự thêm scrollbar dọc kéo đến đáy dgvDetails
            int formH = this.ClientSize.Height;
            int availH = formH - panelDetail.Top - 10;  // khoảng còn lại nếu form đủ cao
            int minDgvH = 400;                           // dgvDetails tối thiểu 400px
            int fixedPanelH = dgvTop + minDgvH + 5;     // chiều cao panel đủ chứa dgvDetails
            // Lấy giá trị lớn hơn: form đủ cao thì lấp đầy form, form thấp thì giữ cố định
            panelDetail.Height = Math.Max(availH, fixedPanelH);

            // dgvDetails kéo dài đến đáy panelDetail
            dgvDetails.Height = panelDetail.Height - dgvTop - 5;

            if (panelMPRFiles != null)
            {
                panelMPRFiles.Location = new Point(panelDetail.Right + 10, panelDetail.Top);
                panelMPRFiles.Width = mprW;
                panelMPRFiles.Height = panelDetail.Height;
            }

            const int _docPOW = 300;
            dgvPO.Width = panelTop.Width - _docPOW - 20;
            // Cập nhật vị trí dgvDocPO và label
            if (dgvDocPO != null)
            {
                dgvDocPO.Left = panelTop.Width - _docPOW;
                dgvDocPO.Width = _docPOW - 5;
                // Chiều cao Document = chiều cao dgvPO (cùng top 105, đồng bộ height)
                dgvDocPO.Top = 28; // sát mép trên panelTop
                dgvDocPO.Height = panelTop.Height - 38; // full
                foreach (Control c in panelTop.Controls)
                    if (c is Label lb && lb.Name == "lblDocPO")
                    { lb.Left = panelTop.Width - _docPOW; lb.Top = 8; }
            }

            // ── Scale buttons: chỉ thu nhỏ khi form < 1280px ──
            ScaleButtons();
        }
        // Load danh sách file từ MPR_Link — tìm theo MPR No tại "Thông tin đơn đặt hàng"
        private async Task LoadMPRFilesAsync(string mprNoOverride = "")
        {
            if (dgvMPRFiles == null) return;
            dgvMPRFiles.Rows.Clear();
            try
            {
                string mprNo = !string.IsNullOrEmpty(mprNoOverride)
                    ? mprNoOverride
                    : txtMPRNo?.Text?.Trim() ?? "";
                if (string.IsNullOrEmpty(mprNo)) return;

                // Load DB trên background thread
                string projectName = "";
                string mprLink = "";
                List<string> files = new List<string>();

                await Task.Run(() =>
                {
                    using (var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection())
                    {
                        conn.Open();
                        var cmd = new Microsoft.Data.SqlClient.SqlCommand(
                            "SELECT TOP 1 Project_Name FROM MPR_Header WHERE MPR_No = @mprNo", conn);
                        cmd.Parameters.AddWithValue("@mprNo", mprNo);
                        projectName = cmd.ExecuteScalar()?.ToString() ?? "";
                    }
                    if (string.IsNullOrEmpty(projectName)) return;

                    var proj = new MPR_Managerment.Services.ProjectService().GetAll()
                        .Find(p => !string.IsNullOrEmpty(p.ProjectName) &&
                            (p.ProjectName.Equals(projectName, StringComparison.OrdinalIgnoreCase) ||
                             p.ProjectName.IndexOf(projectName, StringComparison.OrdinalIgnoreCase) >= 0));

                    mprLink = proj?.MPR_Link?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(mprLink) && System.IO.Directory.Exists(mprLink))
                        files = System.IO.Directory.GetFiles(mprLink).OrderBy(x => x).ToList();
                });

                // Cập nhật UI trên main thread
                foreach (var f in files)
                {
                    int idx = dgvMPRFiles.Rows.Add();
                    dgvMPRFiles.Rows[idx].Cells["FileName"].Value = System.IO.Path.GetFileName(f);
                    dgvMPRFiles.Rows[idx].Cells["FullPath"].Value = f;
                }
            }
            catch { }
        }

        // =========================================================================
        // PRINT BUDGET — xuất PaymentStatus_template.xlsx theo project của PO đang chọn
        // =========================================================================
        private async Task BtnPrintBudget_ClickAsync()
        {
            if (_selectedPO_ID == 0) { SafeWarn("Vui lòng chọn một PO trước!"); return; }

            var selPO = _poList.Find(p => p.PO_ID == _selectedPO_ID);
            if (selPO == null) return;

            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "PaymentStatus_template.xlsx");
            if (!File.Exists(templatePath)) { SafeErr($"Không tìm thấy template:\n{templatePath}"); return; }

            // Hỏi trước: mở trực tiếp hay lưu file
            var choice = MessageBox.Show(
                GetActiveOwner(),
                "Bạn muốn làm gì với file Budget?\n\n" +
                "  [Yes]    → Mở trực tiếp (không lưu)\n" +
                "  [No]     → Lưu file rồi mở\n" +
                "  [Cancel] → Huỷ",
                "Xuất Budget",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question);

            if (choice == DialogResult.Cancel) return;

            bool openOnly = (choice == DialogResult.Yes);
            string defaultName = $"PaymentStatus_{selPO.WorkorderNo}_{DateTime.Now:yyyyMMdd}.xlsx";

            string savePath;
            if (openOnly)
            {
                savePath = Path.Combine(Path.GetTempPath(), defaultName);
            }
            else
            {
                using var sfd = new SaveFileDialog
                {
                    Title = "Lưu Payment Status",
                    Filter = "Excel Files|*.xlsx",
                    FileName = Path.GetFileNameWithoutExtension(defaultName),
                    InitialDirectory = Directory.Exists(@"D:\RÁC") ? @"D:\RÁC" : Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                };
                if (sfd.ShowDialog() != DialogResult.OK) return;
                savePath = sfd.FileName;
            }
            string errMsg = null;
            this.Cursor = Cursors.WaitCursor;
            try
            {
                // Snapshot dữ liệu cần trước Task.Run
                string projName = selPO.Project_Name ?? "";
                string workorderNo = selPO.WorkorderNo ?? "";
                bool isDom = _isDomestic;

                await Task.Run(() =>
                {
                    try
                    {
                        // ── 1. Load project info ──
                        var allProjects = new ProjectService().GetAll();
                        var proj = allProjects.Find(p =>
                            (!string.IsNullOrEmpty(p.WorkorderNo) && p.WorkorderNo.Equals(workorderNo, StringComparison.OrdinalIgnoreCase)) ||
                            (!string.IsNullOrEmpty(p.ProjectName) && p.ProjectName.Equals(projName, StringComparison.OrdinalIgnoreCase)));

                        decimal budget = proj?.PJBudget ?? 0;
                        string projectCode = proj?.ProjectCode ?? workorderNo;

                        // ── 2. Tất cả PO cùng project (loại bỏ các PO đã huỷ - Cancelled/Cancel) ──
var sameProjPOs = _poList.Where(p =>
    ((!string.IsNullOrEmpty(p.WorkorderNo) && p.WorkorderNo.Equals(workorderNo, StringComparison.OrdinalIgnoreCase)) ||
         (!string.IsNullOrEmpty(p.Project_Name) && p.Project_Name.Equals(projName, StringComparison.OrdinalIgnoreCase)))
    && !string.Equals(p.Status?.Trim(), "Cancelled", StringComparison.OrdinalIgnoreCase)
    && !string.Equals(p.Status?.Trim(), "Cancel", StringComparison.OrdinalIgnoreCase))
    .OrderBy(p => p.PONo).ToList();
                        if (!sameProjPOs.Any() 
                            && !string.Equals(selPO.Status?.Trim(), "Cancelled", StringComparison.OrdinalIgnoreCase)
                            && !string.Equals(selPO.Status?.Trim(), "Cancel", StringComparison.OrdinalIgnoreCase)) 
                            sameProjPOs = new List<POHead> { selPO };

                        // ── 3. Load MPR Notes cho từng PO (PO Description = Notes của MPR tương ứng) ──
                        var mprNotes = new Dictionary<string, string>(); // MPR_No → Notes
                        var mprNosToLoad = sameProjPOs.Select(p => p.MPR_No).Where(s => !string.IsNullOrEmpty(s)).Distinct().ToList();
                        if (mprNosToLoad.Any())
                        {
                            using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                            conn.Open();
                            string inClause = string.Join(",", mprNosToLoad.Select((_, idx) => $"@p{idx}"));
                            var cmd = new Microsoft.Data.SqlClient.SqlCommand(
                                $"SELECT MPR_No, ISNULL(Notes,'') AS Notes FROM MPR_Header WHERE MPR_No IN ({inClause})", conn);
                            for (int idx = 0; idx < mprNosToLoad.Count; idx++)
                                cmd.Parameters.AddWithValue($"@p{idx}", mprNosToLoad[idx]);
                            using var rdr = cmd.ExecuteReader();
                            while (rdr.Read())
                                mprNotes[rdr["MPR_No"].ToString()] = rdr["Notes"].ToString();
                        }

                        // ── 4. Tính năm export ──
                        int exportYear = selPO.PO_Date?.Year ?? DateTime.Today.Year;

                        // ── 5. Tạo file Excel từ template ──
                        File.Copy(templatePath, savePath, true);
                        OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                        using var pkg = new OfficeOpenXml.ExcelPackage(new FileInfo(savePath));
                        var ws = pkg.Workbook.Worksheets[0];

                        // Row 4 — năm, Row 5 — tháng 1-12
                        ws.Cells[4, 7].Value = exportYear;
                        for (int m = 1; m <= 12; m++)
                            ws.Cells[5, 6 + m].Value = new DateTime(exportYear, m, 1);

                        // ── 6. Điều chỉnh số dòng data (template có sẵn 3 dòng 6-8) ──
                        int poCount = sameProjPOs.Count;
                        if (poCount > 3) ws.InsertRow(9, poCount - 3);
                        else if (poCount < 3) ws.DeleteRow(6 + poCount, 3 - poCount);
                        int sumRow = 6 + poCount;
                        int lastPoRow = sumRow - 1; // dòng PO cuối

                        // ── 7. Điền từng PO ──
                        decimal[] monthColSum = new decimal[13]; // index 1-12
                        decimal totalPoAmount = 0;

                        for (int i = 0; i < poCount; i++)
                        {
                            var po = sameProjPOs[i];
                            int r = 6 + i;

                            ws.Cells[r, 1].Value = i + 1;    // No.
                            ws.Cells[r, 4].Value = po.PONo;  // PO No.

                            // PO Description = MPR Notes
                            string mprNo = po.MPR_No ?? "";
                            ws.Cells[r, 5].Value = mprNotes.TryGetValue(mprNo, out var note) ? note : mprNo;

                            ws.Cells[r, 6].Value = po.Total_Amount; // PO Amount
                            totalPoAmount += po.Total_Amount;

                            // ── Monthly TT: PO_Date + 7 ngày → tháng thanh toán dự kiến ──
                            int payMonth = 0;
                            if (po.PO_Date.HasValue)
                            {
                                var payDate = po.PO_Date.Value.AddDays(7);
                                if (payDate.Year == exportYear) payMonth = payDate.Month;
                            }
                            for (int m = 1; m <= 12; m++)
                            {
                                decimal amt = (m == payMonth) ? po.Total_Amount : 0;
                                ws.Cells[r, 6 + m].Value = amt > 0 ? (object)amt : null;
                                monthColSum[m] += amt;
                            }

                            decimal rowAcc = payMonth > 0 ? po.Total_Amount : 0;
                            ws.Cells[r, 19].Value = rowAcc;  // Accumulate
                            ws.Cells[r, 20].Value = po.Total_Amount - rowAcc; // Balance
                        }

                        // ── 8. Project No + Budget: merge từ row 6 đến lastPoRow ──
                        if (poCount > 0)
                        {
                            // Xóa merge cũ nếu có
                            ws.Cells[6, 2, lastPoRow, 2].Merge = true;
                            ws.Cells[6, 2].Value = $"{projName}\n{projectCode}";
                            ws.Cells[6, 2].Style.WrapText = true;
                            ws.Cells[6, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                            ws.Cells[6, 3, lastPoRow, 3].Merge = true;
                            ws.Cells[6, 3].Value = budget;
                            ws.Cells[6, 3].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        }

                        // ── 9. Sum row ──
                        ws.Cells[sumRow, 2].Value = "S  u  m";
                        ws.Cells[sumRow, 3].Value = budget;
                        ws.Cells[sumRow, 6].Value = totalPoAmount;
                        for (int m = 1; m <= 12; m++)
                            ws.Cells[sumRow, 6 + m].Value = monthColSum[m] > 0 ? (object)monthColSum[m] : null;
                        decimal grandAcc = monthColSum.Skip(1).Sum();
                        ws.Cells[sumRow, 19].Value = grandAcc;
                        ws.Cells[sumRow, 20].Value = budget - grandAcc;

                        // ══════════════════════════════════════════
                        // ── 10. FORMAT toàn bộ sheet ──
                        // ══════════════════════════════════════════
                        const int COLS = 21; // A–U
                        var clrTitleBg   = System.Drawing.Color.FromArgb(68, 84, 106);
                        var clrHeaderBg  = System.Drawing.Color.FromArgb(31, 78, 121);
                        var clrMonthBg   = System.Drawing.Color.FromArgb(23, 54, 93);
                        var clrSumBg     = System.Drawing.Color.FromArgb(255, 230, 153);
                        var clrAltRow    = System.Drawing.Color.FromArgb(242, 242, 242);
                        var clrWhite     = System.Drawing.Color.White;
                        var clrBlack     = System.Drawing.Color.Black;
                        string numFmt    = "#,##0;[Red]-#,##0;-";

                        // Thin border helper
                        void SetBorder(OfficeOpenXml.ExcelRange rng)
                        {
                            rng.Style.Border.Top.Style    = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            rng.Style.Border.Left.Style   = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            rng.Style.Border.Right.Style  = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        }
                        void SetFill(OfficeOpenXml.ExcelRange rng, System.Drawing.Color c)
                        {
                            rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            rng.Style.Fill.BackgroundColor.SetColor(c);
                        }

                        // ── Row 1: Title ──
                        var title = ws.Cells[1, 1, 1, COLS];
                        SetFill(title, clrTitleBg);
                        title.Style.Font.Color.SetColor(clrWhite);
                        title.Style.Font.Bold = true;
                        title.Style.Font.Size = 14;
                        title.Style.Font.Name = "Times New Roman";
                        title.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        title.Style.VerticalAlignment   = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        ws.Row(1).Height = 28;

                        // ── Row 2: trống ──
                        ws.Row(2).Height = 6;

                        // ── Row 3: Unit VND ──
                        ws.Cells[3, COLS].Value = "Unit : VND";
                        ws.Cells[3, COLS].Style.Font.Italic = true;
                        ws.Cells[3, COLS].Style.Font.Size = 9;
                        ws.Cells[3, COLS].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        ws.Row(3).Height = 14;

                        // ── Row 4: Column headers ──
                        var hdr4 = ws.Cells[4, 1, 4, COLS];
                        SetFill(hdr4, clrHeaderBg);
                        hdr4.Style.Font.Color.SetColor(clrWhite);
                        hdr4.Style.Font.Bold = true;
                        hdr4.Style.Font.Name = "Arial";
                        hdr4.Style.Font.Size = 10;
                        hdr4.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        hdr4.Style.VerticalAlignment   = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        hdr4.Style.WrapText = true;
                        SetBorder(hdr4);
                        ws.Row(4).Height = 36;

                        // ── Row 5: Month sub-headers ──
                        var hdr5 = ws.Cells[5, 1, 5, COLS];
                        SetFill(hdr5, clrMonthBg);
                        hdr5.Style.Font.Color.SetColor(clrWhite);
                        hdr5.Style.Font.Bold = true;
                        hdr5.Style.Font.Name = "Arial";
                        hdr5.Style.Font.Size = 9;
                        hdr5.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        hdr5.Style.VerticalAlignment   = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        SetBorder(hdr5);
                        ws.Row(5).Height = 18;
                        // Format tháng: "Jan-26"
                        for (int m = 1; m <= 12; m++)
                            ws.Cells[5, 6 + m].Style.Numberformat.Format = "mmm-yy";

                        // ── Data rows (6 → sumRow-1) ──
                        for (int r = 6; r < sumRow; r++)
                        {
                            var rowRange = ws.Cells[r, 1, r, COLS];
                            if ((r - 6) % 2 == 1) SetFill(rowRange, clrAltRow);
                            rowRange.Style.Font.Name = "Arial";
                            rowRange.Style.Font.Size = 9;
                            rowRange.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            SetBorder(rowRange);
                            ws.Row(r).Height = 16;
                            // Cột số: right-align + number format
                            foreach (int nc in new[] { 3, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20 })
                            {
                                ws.Cells[r, nc].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                                ws.Cells[r, nc].Style.Numberformat.Format = numFmt;
                            }
                            // Cột text: center/left
                            ws.Cells[r, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // No.
                            ws.Cells[r, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // PO No.
                            ws.Cells[r, 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;   // Description
                            ws.Cells[r, 5].Style.WrapText = true;
                        }
                        // Project No + Budget: center + wrap
                        ws.Cells[6, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        ws.Cells[6, 2].Style.WrapText = true;
                        ws.Cells[6, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        ws.Cells[6, 3].Style.Numberformat.Format = numFmt;

                        // ── Sum row ──
                        var sumRng = ws.Cells[sumRow, 1, sumRow, COLS];
                        SetFill(sumRng, clrSumBg);
                        sumRng.Style.Font.Bold = true;
                        sumRng.Style.Font.Name = "Arial";
                        sumRng.Style.Font.Size = 10;
                        sumRng.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        // Thick outer border cho sum row
                        sumRng.Style.Border.Top.Style    = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        sumRng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        sumRng.Style.Border.Left.Style   = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        sumRng.Style.Border.Right.Style  = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        ws.Row(sumRow).Height = 20;
                        ws.Cells[sumRow, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        foreach (int nc in new[] { 3, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20 })
                        {
                            ws.Cells[sumRow, nc].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                            ws.Cells[sumRow, nc].Style.Numberformat.Format = numFmt;
                        }

                        // ── Column widths ──
                        ws.Column(1).Width  = 4.5;   // No.
                        ws.Column(2).Width  = 22;    // Project No.
                        ws.Column(3).Width  = 16;    // Budget
                        ws.Column(4).Width  = 18;    // PO No.
                        ws.Column(5).Width  = 26;    // PO Description
                        ws.Column(6).Width  = 14;    // PO Amount
                        for (int m = 1; m <= 12; m++)
                            ws.Column(6 + m).Width = 12; // Tháng 1–12
                        ws.Column(19).Width = 14;    // Accumulate
                        ws.Column(20).Width = 14;    // Balance
                        ws.Column(21).Width = 16;    // Remarks

                        // ── Freeze panes: đóng băng tại row 6, col 7 (sau Project/Budget/PO cols) ──
                        ws.View.FreezePanes(6, 7);

                        // ── Print setup ──
                        var ps = ws.PrinterSettings;
                        ps.PrintArea        = ws.Cells[1, 1, sumRow, COLS];   // vùng in: A1 → U{sumRow}
                        ps.RepeatRows       = ws.Cells["1:5"];                 // lặp header khi in nhiều trang
                        ps.Orientation      = OfficeOpenXml.eOrientation.Landscape;
                        ps.PaperSize        = OfficeOpenXml.ePaperSize.A4;
                        ps.FitToPage        = true;
                        ps.FitToWidth       = 1;   // vừa khít 1 trang ngang
                        ps.FitToHeight      = 0;   // tự động số trang dọc
                        ps.TopMargin        = 0.5m;
                        ps.BottomMargin     = 0.5m;
                        ps.LeftMargin       = 0.4m;
                        ps.RightMargin      = 0.4m;
                        ps.HeaderMargin     = 0.2m;
                        ps.FooterMargin     = 0.2m;
                        ps.ShowGridLines    = false;
                        ps.ShowHeaders      = false;
                        // Footer: tên file bên trái, số trang bên phải
                        ws.HeaderFooter.OddFooter.LeftAlignedText  = $"PaymentStatus_{workorderNo}";
                        ws.HeaderFooter.OddFooter.RightAlignedText = "Page &P of &N";

                        pkg.Save();
                    }
                    catch (Exception ex) { errMsg = ex.Message; }
                });

                if (errMsg != null)
                    SafeErr("Lỗi xuất Budget: " + errMsg);
                else
                {
                    if (lblStatus != null)
                        lblStatus.Text = openOnly
                            ? $"✅ Đã mở Budget (tạm thời)"
                            : $"✅ Đã lưu Payment Status: {Path.GetFileName(savePath)}";

                    try
                    {
                        Process.Start(new ProcessStartInfo { FileName = savePath, UseShellExecute = true });
                    }
                    catch (Exception openEx)
                    {
                        SafeErr($"Không thể mở file:\n{openEx.Message}");
                    }
                }
            }
            catch (Exception ex) { SafeErr("Lỗi: " + ex.Message); }
            finally { this.Cursor = Cursors.Default; }
        }

        private void BtnPOStatus_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgvPO == null || dgvPO.SelectedRows.Count == 0)
                {
                    SafeWarn("Vui lòng chọn 1 PO để xem trạng thái.");
                    return;
                }

                var row = dgvPO.SelectedRows[0];
                int poId = 0;
                string poNo = row.Cells["PO_No"]?.Value?.ToString() ?? "";
                string emailStatus = row.Cells["Email_Status"]?.Value?.ToString() ?? "";

                if (row.Cells["ID"]?.Value == null || !int.TryParse(row.Cells["ID"].Value.ToString(), out poId))
                {
                    SafeWarn("Không đọc được thông tin PO_ID.");
                    return;
                }

                var head = _poList?.FirstOrDefault(x => x.PO_ID == poId);
                var details = new POService().GetDetails(poId);

                decimal totalOrdered = details.Sum(d => d.Qty_Per_Sheet);
                decimal totalReceived = details.Sum(d => d.Received_Qty);
                decimal deliveryPercent = totalOrdered <= 0 ? 0 : Math.Round((totalReceived / totalOrdered) * 100m, 2);

                int rirCount = 0;
                using (var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var cmd = new Microsoft.Data.SqlClient.SqlCommand(
                        "SELECT COUNT(1) FROM RIR_head WHERE PONo = @poNo", conn);
                    cmd.Parameters.AddWithValue("@poNo", poNo);
                    var obj = cmd.ExecuteScalar();
                    rirCount = obj != null && obj != DBNull.Value ? Convert.ToInt32(obj) : 0;
                }

                string rirStatus = rirCount > 0 ? $"Đã tạo ({rirCount})" : "Chưa tạo";

                var frm = new Form
                {
                    Text = $"PO status - {poNo}",
                    Size = new Size(900, 600),
                    StartPosition = FormStartPosition.CenterParent,
                    BackColor = Color.White
                };

                var lblSummary = new Label
                {
                    Dock = DockStyle.Top,
                    Height = 95,
                    Font = new Font("Segoe UI", 10, FontStyle.Bold),
                    ForeColor = Color.FromArgb(40, 40, 40),
                    Padding = new Padding(10),
                    Text =
                        $"PO No: {poNo}\r\n" +
                        $"Email: {emailStatus}\r\n" +
                        $"% Giao hàng: {deliveryPercent:N2}% ({totalReceived:N2}/{totalOrdered:N2})\r\n" +
                        $"RIR status: {rirStatus}"
                };

                var tree = new TreeView
                {
                    Dock = DockStyle.Fill,
                    Font = new Font("Segoe UI", 9),
                    HideSelection = false
                };

                var root = new TreeNode($"PO {poNo} ({details.Count} dòng)");
                if (head != null)
                {
                    root.Nodes.Add($"Project: {head.Project_Name}");
                    root.Nodes.Add($"MPR No: {head.MPR_No}");
                    root.Nodes.Add($"Status: {head.Status}");
                }

                foreach (var d in details.OrderBy(x => x.Item_No))
                {
                    var n = new TreeNode($"Item {d.Item_No}: {d.Item_Name}");
                    n.Nodes.Add($"Material: {d.Material}");
                    n.Nodes.Add($"Size: {d.Asize} x {d.Bsize} x {d.Csize}");
                    n.Nodes.Add($"Ordered: {d.Qty_Per_Sheet:N2} {d.UNIT}");
                    n.Nodes.Add($"Received: {d.Received_Qty:N2} {d.UNIT}");
                    n.Nodes.Add($"Delivery %: {(d.Qty_Per_Sheet <= 0 ? 0 : (d.Received_Qty / d.Qty_Per_Sheet) * 100):N2}%");
                    root.Nodes.Add(n);
                }

                root.Expand();
                tree.Nodes.Add(root);

                frm.Controls.Add(tree);
                frm.Controls.Add(lblSummary);
                frm.ShowDialog(this);
            }
            catch (Exception ex)
            {
                SafeErr("Lỗi mở PO status: " + ex.Message);
            }
        }

        private void LoadPO()
        {
            try
            {
                _poList = _service.GetAll();
                _isLoadingPO = true;
                BindPOGrid(_poList);
                _isLoadingPO = false;
                lblStatus.Text = $"Tổng: {_poList.Count} đơn PO";
            }
            catch (Exception ex)
            {
                _isLoadingPO = false;
                MessageBox.Show(GetActiveOwner(), "Lỗi tải dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task LoadPOAsync()
        {
            try
            {
                List<POHead> poList = null;
                List<MPR_Managerment.Models.Supplier> suppliers = null;
                List<MPR_Managerment.Models.ProjectInfo> projects = null;

                await Task.Run(() =>
                {
                    poList = _service.GetAll();
                    suppliers = new SupplierService().GetAll();
                    try { projects = new ProjectService().GetAll(); } catch { projects = new List<MPR_Managerment.Models.ProjectInfo>(); }
                });

                _poList = poList;
                _isLoadingPO = true;
                BindPOGrid(_poList, suppliers, projects);
                _isLoadingPO = false;
                lblStatus.Text = $"Tổng: {_poList.Count} đơn PO";
            }
            catch (Exception ex)
            {
                _isLoadingPO = false;
                MessageBox.Show(GetActiveOwner(), "Lỗi tải dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =========================================================================
        // BIND PO GRID — sắp xếp theo Ngày tạo mới nhất
        // =========================================================================
        private void BindPOGrid(List<POHead> list,
            List<MPR_Managerment.Models.Supplier> suppliers = null,
            List<MPR_Managerment.Models.ProjectInfo> projects2 = null)
        {
            if (suppliers == null) suppliers = new SupplierService().GetAll();
            if (projects2 == null) { try { projects2 = new ProjectService().GetAll(); } catch { projects2 = new List<MPR_Managerment.Models.ProjectInfo>(); } }
            // Sắp xếp theo Created_Date giảm dần (mới nhất lên đầu)
            var sorted = list
                .OrderByDescending(h => h.Created_Date ?? DateTime.MinValue)
                .ThenBy(h => h.PONo, StringComparer.OrdinalIgnoreCase)
                .ToList();

            dgvPO.DataSource = sorted.ConvertAll(h =>
            {
                var supplier = suppliers.Find(s => s.Supplier_ID == h.Supplier_ID);
                // Tra mã dự án từ ProjectService
                var prj = projects2.Find(p =>
                    (!string.IsNullOrEmpty(p.ProjectName) && p.ProjectName.Equals(h.Project_Name, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(p.WorkorderNo) && p.WorkorderNo.Equals(h.WorkorderNo, StringComparison.OrdinalIgnoreCase)));
                string maDA = prj?.ProjectCode ?? h.Project_Name ?? "";
                return new
                {
                    ID = h.PO_ID,
                    PO_No = h.PONo,
                    NCC = supplier?.Short_Name ?? "",
                    Du_An = maDA,                  // Hiển thị mã dự án thay tên
                    MPR_No = h.MPR_No,
                    Ngay_PO = h.PO_Date.HasValue ? h.PO_Date.Value.ToString("dd/MM/yyyy") : "",
                    Trang_Thai = h.Status,
                    Tong_Tien = h.Total_Amount.ToString("N2", _numCulture),
                    Created_By = h.Created_By,
                    Revise = h.Revise,
                    Email_Status    = h.Email_Status,
                    Email_Sent_Info = h.Email_Sent_At.HasValue
                        ? $"{h.Email_Sent_By}\n{h.Email_Sent_At.Value:dd/MM/yy HH:mm}"
                        : ""
                };
            });
            if (dgvPO.Columns.Contains("ID")) dgvPO.Columns["ID"].Visible = false;
            // Auto-fit cột MPR_No, tối đa 200px, WrapText nếu quá dài
            if (dgvPO.Columns.Contains("MPR_No"))
            {
                var col = dgvPO.Columns["MPR_No"];
                col.HeaderText = "MPR No";
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                col.DefaultCellStyle.WrapMode = DataGridViewTriState.False;
            }
            // Cấu hình cột Email_Status
            if (dgvPO.Columns.Contains("Email_Status"))
            {
                var esc = dgvPO.Columns["Email_Status"];
                esc.HeaderText = "Email";
                esc.Width = 65;
                esc.ReadOnly = true;
                esc.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                esc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                esc.DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }

            // Thêm nút gửi email (chỉ thêm 1 lần)
            if (!dgvPO.Columns.Contains("col_SendEmail"))
            {
                var btnEmail = new DataGridViewButtonColumn
                {
                    Name = "col_SendEmail",
                    HeaderText = "Gửi Email",
                    Text = "📧 Gửi",
                    UseColumnTextForButtonValue = false,
                    Width = 80,
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                    FlatStyle = FlatStyle.Flat,
                    DefaultCellStyle = { BackColor = Color.FromArgb(40, 167, 69), ForeColor = Color.White, Font = new Font("Segoe UI", 8, FontStyle.Bold) },
                };
                dgvPO.Columns.Add(btnEmail);
            }
            // Cấu hình cột Created_By (user tạo)
            if (dgvPO.Columns.Contains("Created_By"))
            {
                var c = dgvPO.Columns["Created_By"];
                c.HeaderText = "User Tạo";
                c.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                c.ReadOnly = true;
                c.DefaultCellStyle.ForeColor = Color.FromArgb(80, 80, 80);
            }

            // Cấu hình cột Email_Sent_Info (người gửi + thời gian)
            if (dgvPO.Columns.Contains("Email_Sent_Info"))
            {
                var c = dgvPO.Columns["Email_Sent_Info"];
                c.AutoSizeMode  = DataGridViewAutoSizeColumnMode.None;
                c.Width         = 115;
                c.HeaderText    = "Người gửi / Thời gian";
                c.ReadOnly      = true;
                c.DefaultCellStyle.WrapMode  = DataGridViewTriState.True;
                c.DefaultCellStyle.Font      = new Font("Segoe UI", 8f);
                c.DefaultCellStyle.ForeColor = Color.FromArgb(80, 80, 80);
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            }

            // Tô màu cột Email_Status theo giá trị
            dgvPO.CellFormatting -= DgvPO_EmailCellFormatting;
            dgvPO.CellFormatting += DgvPO_EmailCellFormatting;
            dgvPO.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
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
                // Query Warehouse_Import to get sum of Qty_Import per PO_Detail and populate Nhập kho column
                using (var c = MPR_Managerment.Helpers.DatabaseHelper.GetConnection())
                {
                    c.Open();
                    // Use PO_Detail_ID to aggregate imported qty. Some imports may store value in Qty_Import
                    // or in Received_Qty — use COALESCE to prefer Qty_Import then Received_Qty.
                    var cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                        SELECT d.PO_Detail_ID,
                               ISNULL(SUM(ISNULL(wi.Qty_Import,0)), 0) AS Total
                        FROM PO_Detail d
                        LEFT JOIN Warehouse_Import wi ON wi.PO_Detail_ID = d.PO_Detail_ID
                        WHERE d.PO_ID = @pid
                        GROUP BY d.PO_Detail_ID", c);
                    cmd.Parameters.AddWithValue("@pid", poId);
                    var dt = new System.Data.DataTable();
                    new Microsoft.Data.SqlClient.SqlDataAdapter(cmd).Fill(dt);
                    var map = dt.Rows.Cast<System.Data.DataRow>()
                        .ToDictionary(r => Convert.ToInt32(r[0]), r => Convert.ToDecimal(r[1]));
                    foreach (DataGridViewRow row in dgvDetails.Rows)
                    {
                        if (row.IsNewRow || row.Tag?.ToString() == "TOTAL") continue;
                        var v = row.Cells["PO_Detail_ID"].Value;
                        if (v == null || v == DBNull.Value) continue;
                        if (!int.TryParse(v.ToString(), out int id)) continue;
                        if (map.TryGetValue(id, out var q))
                        {
                            // If zero, show empty; otherwise show number
                            row.Cells["NhapKho"].Value = (q == 0m) ? "" : (object)q;
                        }
                        else
                        {
                            row.Cells["NhapKho"].Value = "";
                        }
                    }
                    if (dgvDetails.Columns.Contains("NhapKho"))
                        dgvDetails.InvalidateColumn(dgvDetails.Columns["NhapKho"].Index);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(GetActiveOwner(), "Lỗi tải chi tiết PO: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateTotal()
        {
            decimal subTotal = 0;
            decimal total = 0;
            decimal totalQty = 0;
            decimal totalKg = 0;
            decimal totalSubAmt = 0;
            decimal totalAmount = 0;
            int dec = _isDomestic ? 0 : 2;

            foreach (DataGridViewRow row in dgvDetails.Rows)
            {
                if (row.IsNewRow || row.Tag?.ToString() == "TOTAL") continue;
                decimal qty = ParseDecimalRaw(row.Cells["Qty"].Value?.ToString() ?? "");
                decimal weight = ParseDecimalRaw(row.Cells["Weight"].Value?.ToString() ?? "");
                decimal price = ParseDecimalRaw(row.Cells["Price"].Value?.ToString() ?? "");
                decimal.TryParse(row.Cells["VAT"].Value?.ToString(), out decimal vat);
                string calcMethod = row.Cells["Calc_Method"].Value?.ToString() ?? "Theo KG";
                decimal baseValue = (calcMethod == "Theo KG") ? weight : qty;
                decimal rowSubAmt = baseValue * price;
                decimal rowAmount = _isDomestic
                    ? Math.Round(rowSubAmt * (1 + vat / 100), 0, MidpointRounding.AwayFromZero)
                    : Math.Round(rowSubAmt, 2, MidpointRounding.AwayFromZero);

                totalQty += qty;
                totalKg += weight;
                totalSubAmt += rowSubAmt;
                totalAmount += rowAmount;
            }
            subTotal = totalSubAmt;
            total = totalAmount;

            string lblFmt = _isDomestic ? "N0" : "N2";
            if (lblSubTotal != null && lblSubTotal.Visible)
                lblSubTotal.Text = "Trước VAT: " + subTotal.ToString(lblFmt, _numCulture) + " VND";
            if (lblTotal != null && lblTotal.Visible)
                lblTotal.Text = "Sau VAT: " + total.ToString(lblFmt, _numCulture) + " VND";

            // ── Cập nhật dòng Total ở cuối dgvDetails ──
            for (int i = dgvDetails.Rows.Count - 1; i >= 0; i--)
                if (dgvDetails.Rows[i].Tag?.ToString() == "TOTAL")
                { dgvDetails.Rows.RemoveAt(i); break; }

            if (dgvDetails.Rows.Count > 0)
            {
                int idx = dgvDetails.Rows.Add();
                var tr = dgvDetails.Rows[idx];
                tr.Tag = "TOTAL";
                tr.ReadOnly = true;

                if (dgvDetails.Columns.Contains("Item_Name")) tr.Cells["Item_Name"].Value = "TỔNG CỘNG";
                if (dgvDetails.Columns.Contains("Qty")) tr.Cells["Qty"].Value = totalQty > 0 ? (object)totalQty : "";
                if (dgvDetails.Columns.Contains("Weight")) tr.Cells["Weight"].Value = totalKg > 0 ? (object)Math.Round(totalKg, 2) : "";
                if (dgvDetails.Columns.Contains("SubAmount")) tr.Cells["SubAmount"].Value = totalSubAmt > 0 ? (object)Math.Round(totalSubAmt, dec, MidpointRounding.AwayFromZero) : "";
                if (dgvDetails.Columns.Contains("Amount")) tr.Cells["Amount"].Value = totalAmount > 0 ? (object)totalAmount : "";

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
                MessageBox.Show(GetActiveOwner(), "Loi khi dan du lieu: " + ex.Message,
                    "Loi Paste", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RecalculateAmount(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= dgvDetails.Rows.Count) return;
            var row = dgvDetails.Rows[rowIndex];
            if (row.IsNewRow || row.Tag?.ToString() == "TOTAL") return;

            decimal qty = ParseDecimalRaw(row.Cells["Qty"].Value?.ToString() ?? "");
            decimal weight = ParseDecimalRaw(row.Cells["Weight"].Value?.ToString() ?? "");
            decimal price = ParseDecimalRaw(row.Cells["Price"].Value?.ToString() ?? "");
            decimal.TryParse(row.Cells["VAT"].Value?.ToString(), out decimal vat);
            string calcMethod = row.Cells["Calc_Method"].Value?.ToString() ?? "Theo KG";

            decimal baseValue = (calcMethod == "Theo KG") ? weight : qty;

            // TT trước thuế = số lượng/KG × đơn giá
            decimal subAmt = baseValue * price;
            int dec = _isDomestic ? 0 : 2;
            decimal subAmtRounded = Math.Round(subAmt, dec, MidpointRounding.AwayFromZero);
            // Thành tiền: Trong nước = TT × (1+VAT%), Nước ngoài = TT (không VAT)
            decimal amount = _isDomestic
                ? Math.Round(subAmt * (1 + vat / 100), 0, MidpointRounding.AwayFromZero)
                : Math.Round(subAmt, 2, MidpointRounding.AwayFromZero);

            if (dgvDetails.Columns.Contains("SubAmount"))
                row.Cells["SubAmount"].Value = subAmtRounded;
            row.Cells["Amount"].Value = amount;

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
            //if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            //string colName = dgvDetails.Columns[e.ColumnIndex].Name;

            //// Chi xu ly cac cot so
            //if (colName != "Price" && colName != "Qty" && colName != "Weight"
            //    && colName != "Amount" && colName != "SubAmount" && colName != "Received") return;

            //string raw = e.Value?.ToString() ?? "";
            //if (string.IsNullOrWhiteSpace(raw)) return;

            //decimal parsed = ParseDecimalRaw(raw);
            //// Luu gia tri decimal thuc vao cell, tranh WinForms parse lai sai
            //e.Value = parsed;
            //e.ParsingApplied = true;
            //// Dong thoi ghi thang vao cell de dam bao
            //if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            //    dgvDetails.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = parsed;
        }

        // Parse số từ chuỗi bất kỳ — tự nhận dạng dấu thập phân và ngàn theo Windows system
        private decimal ParseDecimalRaw(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return 0;
            raw = raw.Trim();

            char sysDec = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0];
            char sysGrp = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator[0];

            // Có cả "." và ","
            if (raw.Contains(".") && raw.Contains(","))
            {
                if (raw.IndexOf('.') < raw.IndexOf(','))
                    // "1.234,56" → dấu "." là ngàn, "," là thập phân
                    raw = raw.Replace(".", "").Replace(',', sysDec);
                else
                    // "1,234.56" → dấu "," là ngàn, "." là thập phân
                    raw = raw.Replace(",", "").Replace('.', sysDec);
            }
            else if (raw.Contains(".") && !raw.Contains(","))
            {
                var parts = raw.Split('.');
                bool allThousand = parts.Length > 1 && parts.Skip(1).All(p => p.Length == 3);
                raw = allThousand
                    ? raw.Replace(".", "")           // "1.234.567" → ngàn separator
                    : raw.Replace('.', sysDec);      // "1.5" → thập phân
            }
            else if (raw.Contains(",") && !raw.Contains("."))
            {
                var parts = raw.Split(',');
                bool allThousand = parts.Length > 1 && parts.Skip(1).All(p => p.Length == 3);
                raw = allThousand
                    ? raw.Replace(",", "")           // "1,234,567" → ngàn separator
                    : raw.Replace(',', sysDec);      // "1,5" → thập phân
            }

            decimal.TryParse(raw,
                System.Globalization.NumberStyles.Number,
                System.Globalization.CultureInfo.CurrentCulture,
                out decimal result);
            return result;
        }

        private void DgvDetails_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0) RecalculateAmount(e.RowIndex);
            AutoAdjustColumnWidths();
        }

        private void BtnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtSearch.Text)) LoadPO();
                else
                {
                    var result = _service.Search(txtSearch.Text.Trim()); BindPOGrid(result);
                    if (result.Count == 0) { lblStatus.Text = "Không tìm thấy kết quả"; SafeInfo("Không tìm thấy kết quả phù hợp!", "Tìm kiếm"); }
                    else lblStatus.Text = $"Tìm thấy: {result.Count} đơn PO";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(GetActiveOwner(), "Lỗi tìm kiếm: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void DgvPO_SelectionChanged(object sender, EventArgs e)
        {
            if (_isLoadingPO || dgvPO.SelectedRows.Count == 0) return;
            var row = dgvPO.SelectedRows[0]; _selectedPO_ID = Convert.ToInt32(row.Cells["ID"].Value);
            var h = _poList.Find(x => x.PO_ID == _selectedPO_ID); if (h == null) return;
            txtPONo.Text = h.PONo; txtProjectName.Text = h.Project_Name; txtWorkorderNo.Text = h.WorkorderNo; txtMPRNo.Text = h.MPR_No;
            _ = LoadMPRFilesAsync();
            // Prepared được gán tự động khi lưu — không hiển thị trên UI
            txtNotes.Text = h.Notes; nudRevise.Value = h.Revise;
            if (h.PO_Date.HasValue) dtpPODate.Value = h.PO_Date.Value;
            var ptIdx = cboPaymentTerm.Items.IndexOf(h.Payment_Term ?? "");
            cboPaymentTerm.SelectedIndex = ptIdx > 0 ? ptIdx : 0;

            // Đọc trạng thái trực tiếp từ DB — đảm bảo luôn hiển thị đúng
            string currentStatus = h.Status ?? "";
            try
            {
                using var connSt = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                connSt.Open();
                var cmdSt = new Microsoft.Data.SqlClient.SqlCommand(
                    "SELECT Status FROM PO_head WHERE PO_ID = @id", connSt);
                cmdSt.Parameters.AddWithValue("@id", _selectedPO_ID);
                var dbStatus = cmdSt.ExecuteScalar()?.ToString();
                if (!string.IsNullOrEmpty(dbStatus)) currentStatus = dbStatus;
            }
            catch { }

            // Gán vào cboStatus — so sánh không phân biệt hoa thường
            cboStatus.SelectedIndex = 0;
            for (int i = 0; i < cboStatus.Items.Count; i++)
            {
                if (string.Equals(cboStatus.Items[i]?.ToString(), currentStatus,
                    StringComparison.OrdinalIgnoreCase))
                {
                    cboStatus.SelectedIndex = i;
                    break;
                }
            }
            // Đồng bộ lại _poList
            if (h.Status != currentStatus) h.Status = currentStatus;

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
            LoadDocumentsPO();

            // Khóa NCC đối với PO đã tồn tại
            if (h.Supplier_ID > 0)
            {
                cboSupplier.Enabled = false;
                _supplierLocked = true;
                btnUnlockSupplier.Visible = true;
            }

            // Kiểm tra tự động khi mở PO — không thông báo, chỉ cập nhật UI/DB lặng lẽ
            CheckAndAutoCompleteStatus(_selectedPO_ID, silent: true);
        }

        private void BtnUnlockSupplier_Click(object sender, EventArgs e)
        {
            // Nếu đang mở thì cho phép khóa lại ngay
            if (cboSupplier.Enabled)
            {
                // Kiểm tra xem NCC có thay đổi không
                if (_selectedPO_ID > 0)
                {
                    var currentPO = _poList.Find(x => x.PO_ID == _selectedPO_ID);
                    if (currentPO != null)
                    {
                        int newSupplierId = Convert.ToInt32(cboSupplier.SelectedValue ?? 0);
                        
                        // Nếu NCC đã thay đổi, cập nhật vào database và app
                        if (newSupplierId > 0 && newSupplierId != currentPO.Supplier_ID)
                        {
                            try
                            {
                                // Cập nhật Supplier_ID trong database
                                using (var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection())
                                {
                                    conn.Open();
                                    var cmd = new Microsoft.Data.SqlClient.SqlCommand(
                                        "UPDATE PO_head SET Supplier_ID = @SupplierID WHERE PO_ID = @PO_ID", conn);
                                    cmd.Parameters.AddWithValue("@SupplierID", newSupplierId);
                                    cmd.Parameters.AddWithValue("@PO_ID", _selectedPO_ID);
                                    cmd.ExecuteNonQuery();
                                }

                                // Cập nhật trong _poList (dữ liệu local)
                                currentPO.Supplier_ID = newSupplierId;

                                // Cập nhật UI - refresh grid để hiển thị tên NCC mới
                                _isLoadingPO = true;
                                BindPOGrid(_poList);
                                
                                // Chọn lại PO hiện tại trong grid
                                foreach (DataGridViewRow row in dgvPO.Rows)
                                {
                                    if (Convert.ToInt32(row.Cells["ID"].Value) == _selectedPO_ID)
                                    {
                                        row.Selected = true;
                                        dgvPO.CurrentCell = row.Cells[0];
                                        break;
                                    }
                                }
                                _isLoadingPO = false;

                                SafeInfo("Đã cập nhật thông tin Nhà cung cấp thành công!", "Thông báo");
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(GetActiveOwner(), 
                                    "Lỗi khi cập nhật Nhà cung cấp: " + ex.Message, 
                                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }
                    }
                }

                cboSupplier.Enabled = false;
                _supplierLocked = true;
                btnUnlockSupplier.Visible = true;
                return;
            }

            // Nếu đang khóa thì yêu cầu xác thực mật khẩu Admin để mở khóa
            if (VerifyAdminPassword())
            {
                cboSupplier.Enabled = true;
                _supplierLocked = false;
                btnUnlockSupplier.Visible = false;
                SafeInfo("Đã mở khóa lựa chọn Nhà cung cấp. Vui lòng cẩn trọng khi thay đổi NCC cho PO đã có dữ liệu!", "Mở khóa");
            }
        }

        private void BtnNewPO_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("PO", "Tạo PO", "Tạo PO")) return;
            ClearHeader(); _selectedPO_ID = 0; dgvDetails.Rows.Clear(); dgvFiles.Rows.Clear();
            dgvDelivery.Rows.Clear();
            UpdateTotal(); txtPONo.Focus(); lblStatus.Text = "Đang tạo đơn PO mới...";
            cboSupplier.Enabled = true; _supplierLocked = false; btnUnlockSupplier.Visible = false;
        }

        // =========================================================================
        // LƯU TOÀN BỘ PO (Header + Detail)
        // =========================================================================
        /// <summary>
        /// Tự động chuyển trạng thái PO thành "Complete" khi tiến độ nhập kho >= 100%.
        /// Tiến độ = SUM(Qty_Import) / SUM(Qty_Per_Sheet) từ Warehouse_Import và PO_Detail.
        /// </summary>
        private void CheckAndAutoCompleteStatus(int poId, bool silent = false)
        {
            if (poId <= 0) return;
            try
            {
                using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                conn.Open();

                var cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                    SELECT
                        ISNULL(SUM(d.Qty_Per_Sheet), 0)  AS TotalOrdered,
                        ISNULL(SUM(wi.Imported), 0)      AS TotalImported
                    FROM PO_Detail d
                    LEFT JOIN (
                        SELECT PO_Detail_ID, SUM(Qty_Import) AS Imported
                        FROM Warehouse_Import
                        GROUP BY PO_Detail_ID
                    ) wi ON wi.PO_Detail_ID = d.PO_Detail_ID
                    WHERE d.PO_ID = @poId", conn);
                cmd.Parameters.AddWithValue("@poId", poId);

                var dt = new System.Data.DataTable();
                new Microsoft.Data.SqlClient.SqlDataAdapter(cmd).Fill(dt);
                if (dt.Rows.Count == 0) return;

                decimal totalOrdered = Convert.ToDecimal(dt.Rows[0]["TotalOrdered"]);
                decimal totalImported = Convert.ToDecimal(dt.Rows[0]["TotalImported"]);
                if (totalOrdered <= 0) return;

                decimal progress = totalImported / totalOrdered * 100;
                if (progress < 100m) return;

                // Kiểm tra trạng thái hiện tại
                var cmdStatus = new Microsoft.Data.SqlClient.SqlCommand(
                    "SELECT Status FROM PO_head WHERE PO_ID = @poId", conn);
                cmdStatus.Parameters.AddWithValue("@poId", poId);
                string currentStatus = cmdStatus.ExecuteScalar()?.ToString() ?? "";
                if (currentStatus == "Completed" || currentStatus == "Cancelled") return;

                // Cập nhật DB
                var cmdUpdate = new Microsoft.Data.SqlClient.SqlCommand(
                    "UPDATE PO_head SET Status = 'Completed' WHERE PO_ID = @poId", conn);
                cmdUpdate.Parameters.AddWithValue("@poId", poId);
                cmdUpdate.ExecuteNonQuery();

                // Cập nhật UI
                var cboIdx = cboStatus.Items.IndexOf("Completed");
                if (cboIdx >= 0) cboStatus.SelectedIndex = cboIdx;

                // Cập nhật _poList
                var po = _poList.Find(p => p.PO_ID == poId);
                if (po != null) po.Status = "Completed";

                // Cập nhật dòng trong dgvPO
                if (dgvPO.Columns.Contains("Status"))
                    foreach (DataGridViewRow row in dgvPO.Rows)
                        if (Convert.ToInt32(row.Cells["ID"].Value ?? 0) == poId)
                        { row.Cells["Status"].Value = "Completed"; break; }

                if (!silent)
                    MessageBox.Show(GetActiveOwner(),
                        $"✅ Tiến độ nhập kho đạt {progress:F1}%\n" +
                        "Trạng thái PO đã tự động chuyển sang \"Completed\".",
                        "Hoàn tất giao hàng", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CheckAndAutoCompleteStatus: {ex.Message}");
            }
        }

        private async void BtnSavePO_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("PO", "Lưu PO", "Lưu PO")) return;
            dgvDetails.EndEdit();
            if (string.IsNullOrWhiteSpace(txtPONo.Text))
            {
                SafeWarn("Vui lòng nhập PO No!"); txtPONo.Focus(); return;
            }
            int _suppId = Convert.ToInt32(cboSupplier.SelectedValue ?? 0);
            if (_suppId <= 0 || string.IsNullOrWhiteSpace(cboSupplier.Text))
            {
                SafeWarn("Vui long chon Nha Cung Cap truoc khi luu PO!");
                cboSupplier.Focus(); return;
            }
            if (dgvDetails.Rows.Count == 0 && MessageBox.Show(GetActiveOwner(), "Don hang nay chua co chi tiet vat tu nao.\nBan co chac chan muon luu chi voi Header khong?", "Canh bao", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No) return;

            btnSavePO.Enabled = false;
            this.Cursor = Cursors.WaitCursor;
            try
            {
                string basePONo = txtPONo.Text.Trim();
                int revIdx = basePONo.LastIndexOf("_Rev"); if (revIdx > 0) basePONo = basePONo.Substring(0, revIdx);
                string finalPONo = basePONo;

                if (_selectedPO_ID == 0)
                {
                    bool isBaseDuplicate = _poList.Exists(p => p.PONo == basePONo && p.PO_ID != _selectedPO_ID);
                    if (isBaseDuplicate || nudRevise.Value > 0)
                    {
                        if (nudRevise.Value == 0)
                        {
                            SafeWarn("So PO nay da ton tai!\nVui long tang so Revise de tao ban sua doi.");
                            nudRevise.Focus(); return;
                        }
                        finalPONo = $"{basePONo}_Rev{nudRevise.Value}";
                        if (_poList.Exists(p => p.PONo == finalPONo && p.PO_ID != _selectedPO_ID))
                        {
                            MessageBox.Show(GetActiveOwner(), $"Ban '{finalPONo}' cung da ton tai!\nVui long tang Revise len cao hon.",
                                "Trung lap", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            nudRevise.Focus(); return;
                        }
                    }
                }
                else
                {
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
                    Prepared = _currentUser,
                    Reviewed = txtReviewed.Text.Trim(),
                    Agreement = txtAgreement.Text.Trim(),
                    Approved = txtApproved.Text.Trim(),
                    Notes = txtNotes.Text.Trim(),
                    Payment_Term = cboPaymentTerm.SelectedIndex > 0 ? cboPaymentTerm.SelectedItem.ToString() : "",
                    PO_Date = dtpPODate.Value,
                    Status = cboStatus.SelectedItem?.ToString() ?? "Draft",
                    Revise = (int)nudRevise.Value,
                    Supplier_ID = Convert.ToInt32(cboSupplier.SelectedValue ?? 0),
                    ProjectCode = _projectCodeImport,
                    Expected_Delivery = dtpPOExpectDelivery.Value
                };

                // Lưu DB trên background thread
                int savedId = _selectedPO_ID;
                await Task.Run(() =>
                {
                    if (_selectedPO_ID == 0) savedId = _service.InsertHead(h, _currentUser);
                    else _service.UpdateHead(h, _currentUser);
                });
                if (savedId > 0) _selectedPO_ID = savedId;
                h.PO_ID = _selectedPO_ID;
                txtPONo.Text = finalPONo;

                await Task.Run(() =>
                {
                    SaveDetailsToDb();
                    RefreshMPRDetailLinks(_selectedPO_ID);
                    CheckAndAutoCompleteStatus(_selectedPO_ID);
                });

                // Tự động xuất Excel (fire-and-forget, không block UI)
                _ = Task.Run(() => { try { AutoSavePOExcelToBackground(h, _selectedPO_ID); } catch { } });

                // Hiển thị kết quả lưu trên status label (không dùng MessageBox)
                if (lblStatus != null)
                    lblStatus.Text = $"✅ Đã lưu PO '{finalPONo}' ({dgvDetails.Rows.Count} dòng vật tư) lúc {DateTime.Now:HH:mm:ss}";

                // Trigger notification cho frmMain (nếu là PO mới)
                if (_selectedPO_ID == savedId && savedId > 0)
                {
                    frmMain.TriggerNewDataCheck();
                }

                // Reload danh sách PO async
                await LoadPOAsync();

                // Chọn lại dòng vừa lưu (suppress SelectionChanged để tránh cascade)
                _isLoadingPO = true;
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
                _isLoadingPO = false;
                _selectedPO_ID = savedId;

                // Reload details (sync nhưng chỉ 1 PO — nhanh)
                LoadDetails(savedId);
            }
            catch (Exception ex)
            {
                MessageBox.Show(GetActiveOwner(), "Lỗi khi lưu PO: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnSavePO.Enabled = true;
                this.Cursor = Cursors.Default;
            }
        }

        // =========================================================================
        // TỰ ĐỘNG XUẤT EXCEL — BACKGROUND VERSION (không blocking, không MessageBox)
        // Chạy trong Task.Run, silently skip nếu thiếu PO_Link hoặc template
        // =========================================================================
        private void AutoSavePOExcelToBackground(POHead poHead, int poId)
        {
            // Chạy trong Task.Run — không dùng MessageBox, không block UI
            try
            {
                var projects = new ProjectService().GetAll();
                var prj = projects.Find(p =>
                    (!string.IsNullOrEmpty(p.WorkorderNo) && p.WorkorderNo.Equals(poHead.WorkorderNo, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(p.ProjectName) && p.ProjectName.Equals(poHead.Project_Name, StringComparison.OrdinalIgnoreCase)));
                if (prj == null || string.IsNullOrEmpty(prj.PO_Link)) return;

                string poLinkDir = prj.PO_Link.Trim();
                if (!Directory.Exists(poLinkDir)) return;

                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "po_template.xlsx");
                if (!File.Exists(templatePath)) return;

                string safePoNo = poHead.PONo?.Replace("/", "-").Replace("\\", "-").Replace(":", "-") ?? "PO";
                string fileName = $"PO_{safePoNo}_{DateTime.Now:ddMMyyyy}.xlsx";
                string destPath = Path.Combine(poLinkDir, fileName);
                File.Copy(templatePath, destPath, true);

                var details = new POService().GetDetails(poId);
                var suppliers = new SupplierService().GetAll();
                var supplier = suppliers.Find(s => s.Supplier_ID == poHead.Supplier_ID);
                string fullName = AppSession.CurrentUser?.Full_Name ?? _currentUser ?? "Admin";
                bool isDomestic = _isDomestic;
                string amtFmt = isDomestic ? "#,##0" : "#,##0.00";

                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(destPath)))
                {
                    var ws = package.Workbook.Worksheets[0];
                    ReplaceCell(ws, "<<PROJECT_NAME>>", prj.ProjectName ?? poHead.Project_Name ?? "");
                    ReplaceCell(ws, "<<WO-NO>>", poHead.WorkorderNo ?? "");
                    ReplaceCell(ws, "<<REV.NUM>>", poHead.Revise.ToString());
                    ReplaceCell(ws, "<<DATE>>", poHead.PO_Date.HasValue ? poHead.PO_Date.Value.ToString("dd/MM/yyyy") : DateTime.Today.ToString("dd/MM/yyyy"));
                    ReplaceCell(ws, "<<MPR-NO>>", poHead.MPR_No ?? "");
                    ReplaceCell(ws, "<<PO-NO>>", poHead.PONo ?? "");
                    string supplierInfo = supplier != null ? $"{supplier.Company_Name}\nCert: {supplier.Cert ?? ""}\nEmail: {supplier.Email}" : "";
                    ReplaceCell(ws, "<<SUPPLIER-INFO>>", supplierInfo);
                    ReplaceCell(ws, "<<USER_NAME>>", fullName);
                    ReplaceCell(ws, "<<user name>>", fullName);
                    ReplaceCell(ws, "<<User Name>>", fullName);
                    ws.Cells[5, 15].Value = !string.IsNullOrEmpty(poHead.Payment_Term) ? poHead.Payment_Term : "Within 7 days after delivery";
                    ws.Cells[8, 11].Value = DateTime.Today.AddDays(7).ToString("dd/MM/yyyy");

                    int startRow = 8;
                    if (details.Count > 1) { ws.InsertRow(startRow + 1, details.Count - 1); for (int i = 1; i < details.Count; i++) ws.Cells[startRow, 1, startRow, 17].Copy(ws.Cells[startRow + i, 1]); }
                    decimal totalAfterVAT = 0;
                    for (int i = 0; i < details.Count; i++)
                    {
                        var d = details[i]; int row = startRow + i;
                        decimal q = d.Qty_Per_Sheet; decimal wk = d.Weight_kg; decimal realPrice = d.Price;
                        string rem = d.Remarks ?? "";
                        if (rem.Contains("[CALC:KG]")) 
                        { 
                            rem = rem.Replace("[CALC:KG]", "").Trim(); 
                            if (wk > 0 && q > 0) realPrice = (d.Price * q) / wk; 
                        }
                        else if (rem.Contains("[CALC:SL]"))
                        {
                            rem = rem.Replace("[CALC:SL]", "").Trim();
                        }
                        decimal dAmt = Math.Round(d.Amount, isDomestic ? 0 : 2, MidpointRounding.AwayFromZero);
                        ws.Cells[row, 1].Value = i + 1; ws.Cells[row, 2].Value = d.Item_Name ?? ""; ws.Cells[row, 3].Value = d.Material ?? "";
                        ws.Cells[row, 4].Value = d.Asize; ws.Cells[row, 5].Value = d.Bsize; ws.Cells[row, 6].Value = d.Csize;
                        ws.Cells[row, 7].Value = q; ws.Cells[row, 8].Value = d.UNIT ?? ""; ws.Cells[row, 9].Value = wk;
                        ws.Cells[row, 10].Value = d.MPSNo ?? ""; ws.Cells[row, 11].Value = poHead.Expected_Delivery; ws.Cells[row, 12].Value = "Kho DLHI";
                        ws.Cells[row, 13].Value = Math.Round(realPrice, 0); ws.Cells[row, 14].Value = dAmt; ws.Cells[row, 16].Value = rem;
                        totalAfterVAT += isDomestic ? Math.Round(dAmt * (1 + d.VAT / 100), 0, MidpointRounding.AwayFromZero) : dAmt;
                    }
                    int subTotalRow = startRow + details.Count;
                    ws.Cells[subTotalRow, 3].Value = "SUB-TOTAL";
                    ws.Cells[subTotalRow, 9].Formula = $"=SUM(I{startRow}:I{startRow + details.Count - 1})";
                    ws.Cells[subTotalRow, 9].Style.Numberformat.Format = "#,##0.00";
                    ws.Cells[subTotalRow, 14].Formula = $"=SUM(N{startRow}:N{startRow + details.Count - 1})";
                    ws.Cells[subTotalRow + 1, 3].Value = "Final Price Requested (Included VAT)";
                    ws.Cells[subTotalRow + 1, 14].Value = totalAfterVAT;
                    package.Save();
                }

                this.BeginInvoke((Action)(() =>
                {
                    if (lblStatus != null) lblStatus.Text = $"✅ Lưu PO '{poHead.PONo}' thành công | Excel: {fileName}";
                    try { Process.Start(new ProcessStartInfo { FileName = destPath, UseShellExecute = true }); }
                    catch { }
                }));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AutoSaveBg] {ex.Message}");
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
                    MessageBox.Show(GetActiveOwner(),
                        "⚠️ Không tìm thấy đường dẫn PO Link cho dự án này!\n\n" +
                        "Vui lòng cập nhật 'PO Link' trong mục Quản lý Dự án trước khi lưu file Excel tự động.",
                        "Thiếu PO Link", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string poLinkDir = prj.PO_Link.Trim();
                if (!Directory.Exists(poLinkDir))
                {
                    var createDir = MessageBox.Show(GetActiveOwner(),
                        $"⚠️ Thư mục PO Link không tồn tại:\n{poLinkDir}\n\nBạn có muốn tạo thư mục này không?",
                        "Thư mục không tồn tại", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (createDir == DialogResult.Yes)
                    {
                        try { Directory.CreateDirectory(poLinkDir); }
                        catch (Exception ex)
                        {
                            MessageBox.Show(GetActiveOwner(), $"❌ Không thể tạo thư mục:\n{poLinkDir}\n\nLỗi: {ex.Message}",
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
                    MessageBox.Show(GetActiveOwner(),
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

                var confirm = MessageBox.Show(GetActiveOwner(), confirmMsg,
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
                    string fullNameAuto = !string.IsNullOrWhiteSpace(AppSession.CurrentUser?.Full_Name)
                        ? AppSession.CurrentUser.Full_Name
                        : !string.IsNullOrWhiteSpace(_currentUser) ? _currentUser
                        : AppSession.CurrentUser?.Username ?? "Admin";
                    ReplaceCell(ws, "<<USER_NAME>>", fullNameAuto);
                    ReplaceCell(ws, "<<user name>>", fullNameAuto);
                    ReplaceCell(ws, "<<User Name>>", fullNameAuto);

                    ws.Cells[5, 15].Value = !string.IsNullOrEmpty(poHead.Payment_Term)
                        ? poHead.Payment_Term : "Within 7 days after delivery";
                    ws.Cells[8, 11].Value = DateTime.Today.AddDays(7).ToString("dd/MM/yyyy");

                    int startRow = 8;
                    int detailCount = details.Count;

                    if (detailCount > 1)
                    {
                        ws.InsertRow(startRow + 1, detailCount - 1);
                        for (int i = 1; i < detailCount; i++)
                            ws.Cells[startRow, 1, startRow, 17].Copy(ws.Cells[startRow + i, 1]);
                    }

                    decimal totalAfterVAT = 0;
                    string amtFmt = _isDomestic ? "#,##0" : "#,##0.00";
                    int amtDec = _isDomestic ? 0 : 2;
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
                        else if (rem.Contains("[CALC:SL]"))
                        {
                            rem = rem.Replace("[CALC:SL]", "").Trim();
                        }

                        decimal dAmt = Math.Round(d.Amount, amtDec, MidpointRounding.AwayFromZero);

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
                        ws.Cells[row, 11].Value = poHead.Expected_Delivery;
                        ws.Cells[row, 12].Value = "Kho DLHI";
                        ws.Cells[row, 13].Value = Math.Round(realPrice, 0);
                        ws.Cells[row, 14].Value = dAmt;
                        ws.Cells[row, 16].Value = rem;

                        // --- CỦNG CỐ ĐỊNH DẠNG MERGE CHO REMARKS ---
                        // Nếu trong template cột Remarks gộp từ cột 16 (P) và 17 (Q)
                        if (!ws.Cells[row, 16].Merge)
                        {
                            ws.Cells[row, 16, row, 17].Merge = true;
                        }

                        totalAfterVAT += _isDomestic
                            ? Math.Round(dAmt * (1 + d.VAT / 100), 0, MidpointRounding.AwayFromZero)
                            : dAmt;

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
                        ws.Cells[row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        ws.Cells[row, 16].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                        ws.Cells[row, 11].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        ws.Cells[row, 11].Style.Numberformat.Format = "dd/MM/yyyy";

                        ws.Cells[row, 13].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        ws.Cells[row, 13].Style.Numberformat.Format = "#,##0.00";

                        ws.Cells[row, 14].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        ws.Cells[row, 14].Style.Numberformat.Format = amtFmt;
                    }

                    int subTotalRow = startRow + detailCount;
                    int vatRow = subTotalRow + 1;
                    ws.Cells[subTotalRow, 3].Value = "SUB-TOTAL";
                    ws.Cells[subTotalRow, 9].Formula = $"=SUM(I{startRow}:I{startRow + detailCount - 1})";
                    ws.Cells[subTotalRow, 9].Style.Numberformat.Format = "#,##0.00";
                    ws.Cells[subTotalRow, 14].Formula = $"=SUM(N{startRow}:N{startRow + detailCount - 1})";
                    ws.Cells[subTotalRow, 14].Style.Numberformat.Format = amtFmt;
                    ws.Cells[vatRow, 3].Value = "Final Price Requested (Included VAT)";
                    ws.Cells[vatRow, 14].Value = totalAfterVAT;
                    ws.Cells[vatRow, 14].Style.Numberformat.Format = amtFmt;
                    ws.Cells[vatRow, 14].Style.Font.Bold = true;

                    package.Save();
                }

                // 6. Refresh danh sách file đính kèm
                LoadFiles(poHead.WorkorderNo, poHead.Project_Name);

                // 7. Thông báo thành công và tự động mở file
                MessageBox.Show(GetActiveOwner(),
                    $"✅ Đã lưu file Excel thành công!\n\n📄 File: {fileName}\n📂 Thư mục: {poLinkDir}",
                    "Lưu file thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);

                Process.Start(new ProcessStartInfo { FileName = destPath, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show(GetActiveOwner(), $"❌ Lỗi khi tự động lưu file Excel vào PO Link:\n\n{ex.Message}",
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
                SafeWarn("Vui lòng chọn hoặc lưu Header PO trước!"); return;
            }
            if (dgvDetails.Rows.Count == 0)
            {
                SafeWarn("Không có dòng nào để lưu!");
                return;
            }

            // ── Xác thực mật khẩu Admin trước khi lưu ──
            if (!VerifyAdminPassword()) return;

            try
            {
                SaveDetailsToDb();
                // Cập nhật lại link MPR_Detail_ID sau khi lưu
                RefreshMPRDetailLinks(_selectedPO_ID);
                MessageBox.Show(GetActiveOwner(), $"✅ Đã lưu {dgvDetails.Rows.Count} dòng chi tiết thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDetails(_selectedPO_ID);
            }
            catch (Exception ex)
            {
                MessageBox.Show(GetActiveOwner(), "Lỗi khi lưu chi tiết: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                //decimal q = decimal.TryParse(row.Cells["Qty"].Value?.ToString(), out decimal _q) ? _q : 0;
                //decimal wk = decimal.TryParse(row.Cells["Weight"].Value?.ToString(), out decimal _wk) ? _wk : 0;
                //decimal p = decimal.TryParse((row.Cells["Price"].Value?.ToString() ?? "0").Replace(",", ""), out decimal _p) ? _p : 0;
                //decimal vat = decimal.TryParse(row.Cells["VAT"].Value?.ToString(), out decimal _vt) ? _vt : 0;

                decimal q = Common.Common.ParseDecimalRaw(row.Cells["Qty"].Value?.ToString());
                decimal wk = Common.Common.ParseDecimalRaw(row.Cells["Weight"].Value?.ToString());
                decimal p = Common.Common.ParseDecimalRaw(row.Cells["Price"].Value?.ToString());
                decimal vat = Common.Common.ParseDecimalRaw(row.Cells["VAT"].Value?.ToString());

                string calcMethod = row.Cells["Calc_Method"].Value?.ToString() ?? "Theo KG";
                string remarks = row.Cells["Remarks"].Value?.ToString() ?? "";
                remarks = remarks.Replace("[CALC:KG]", "").Replace("[CALC:SL]", "").Trim();
                decimal dbPrice = p;
                if (calcMethod == "Theo KG")
                {
                    remarks += " [CALC:KG]";
                    if (q > 0 && wk > 0)
                        dbPrice = (wk * p) / q;
                }
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

        // =====================================================================
        //  Tự động cập nhật MPR_Detail_ID cho các dòng PO_Detail
        //  bằng cách match Item_Name + Material với MPR_Details cùng MPR
        //  Đảm bảo frmMPR theo dõi đúng số PO theo từng hạng mục
        // =====================================================================
        private void RefreshMPRDetailLinks(int poId)
        {
            if (poId <= 0) return;
            try
            {
                using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                conn.Open();

                // Lấy MPR_No của PO này
                var cmdMPR = new Microsoft.Data.SqlClient.SqlCommand(
                    "SELECT MPR_No FROM PO_head WHERE PO_ID = @poId AND ISNULL(Status,'') <> 'Cancelled'", conn);
                cmdMPR.Parameters.AddWithValue("@poId", poId);
                string mprNo = cmdMPR.ExecuteScalar()?.ToString() ?? "";
                if (string.IsNullOrEmpty(mprNo)) return;

                // Cập nhật MPR_Detail_ID cho các dòng PO_Detail chưa có link
                // hoặc link sai — match theo Item_Name + Material + kích thước
                var cmdUpdate = new Microsoft.Data.SqlClient.SqlCommand(@"
                    -- Bước 1: Update PO_Detail có MPR_Detail_ID = NULL hoặc không khớp
                    -- Match theo item_name + Material (case-insensitive)
                    UPDATE pod
                    SET    pod.MPR_Detail_ID = md.Detail_ID
                    FROM   PO_Detail pod
                    INNER JOIN MPR_Details md
                           ON LTRIM(RTRIM(LOWER(md.item_name)))  = LTRIM(RTRIM(LOWER(pod.item_name)))
                          AND LTRIM(RTRIM(LOWER(md.Material)))   = LTRIM(RTRIM(LOWER(pod.Material)))
                    INNER JOIN MPR_Header   mh ON mh.MPR_ID = md.MPR_ID
                    WHERE  pod.PO_ID = @poId
                      AND  mh.MPR_No = @mprNo
                      AND  (pod.MPR_Detail_ID IS NULL
                            OR pod.MPR_Detail_ID NOT IN (
                               SELECT Detail_ID FROM MPR_Details WHERE MPR_ID = mh.MPR_ID
                            ))", conn);
                cmdUpdate.Parameters.AddWithValue("@poId", poId);
                cmdUpdate.Parameters.AddWithValue("@mprNo", mprNo);
                int updated = cmdUpdate.ExecuteNonQuery();

                if (updated > 0)
                    System.Diagnostics.Debug.WriteLine(
                        $"[RefreshMPRDetailLinks] Updated {updated} PO_Detail rows for PO_ID={poId}, MPR={mprNo}");
            }
            catch (Exception ex)
            {
                // Không chặn luồng chính — chỉ log
                System.Diagnostics.Debug.WriteLine("[RefreshMPRDetailLinks] Error: " + ex.Message);
            }
        }

        // ── Load INV + Delivery vào bảng Document của dgvDocPO ─────────────
        private void LoadDocumentsPO()
        {
            if (dgvDocPO == null) return;
            dgvDocPO.Rows.Clear();
            if (_selectedPO_ID == 0) return;
            var h = _poList.Find(x => x.PO_ID == _selectedPO_ID);
            if (h == null) return;
            // Tìm project để lấy INV_Link và DeliveryNote_Link
            MPR_Managerment.Models.ProjectInfo proj = null;
            try
            {
                var all = new ProjectService().GetAll();
                proj = all.Find(p =>
                    (!string.IsNullOrEmpty(p.ProjectName) && p.ProjectName.Equals(h.Project_Name, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(p.WorkorderNo) && p.WorkorderNo.Equals(h.WorkorderNo, StringComparison.OrdinalIgnoreCase)));
            }
            catch { }
            string poNo = h.PONo ?? "";
            ScanFolderToDocPO(proj?.INV_Link ?? "", $"INV_{poNo}");
            ScanFolderToDocPO(proj?.DeliveryNote_Link ?? "", $"Delivery_{poNo}");
        }

        private void ScanFolderToDocPO(string folder, string prefix)
        {
            if (string.IsNullOrWhiteSpace(folder) || !System.IO.Directory.Exists(folder)) return;
            try
            {
                var files = System.IO.Directory.GetFiles(folder, $"{prefix}*", System.IO.SearchOption.TopDirectoryOnly);
                foreach (var f in files)
                {
                    int i = dgvDocPO.Rows.Add();
                    dgvDocPO.Rows[i].Cells["DocPO_Path"].Value = f;
                    dgvDocPO.Rows[i].Cells["DocPO_Name"].Value = System.IO.Path.GetFileName(f);
                    // Màu: xanh dương = INV, xanh lá = Delivery
                    dgvDocPO.Rows[i].DefaultCellStyle.ForeColor =
                        prefix.StartsWith("INV") ? Color.FromArgb(0, 120, 212) : Color.FromArgb(40, 167, 69);
                }
            }
            catch { }
        }

        private void BtnDeletePO_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("PO", "Xóa PO", "Xóa PO")) return;
            if (dgvPO.SelectedRows.Count == 0 || _selectedPO_ID == 0)
            {
                SafeWarn("Vui lòng chọn một đơn PO trong 'Danh sách đơn đặt hàng' để xóa!");
                return;
            }
            string poNo = dgvPO.SelectedRows[0].Cells["PO_No"].Value?.ToString() ?? "";
            if (MessageBox.Show(GetActiveOwner(), $"Bạn có chắc chắn muốn xóa đơn PO '{poNo}' và toàn bộ chi tiết vật tư bên trong không?\nHành động này không thể hoàn tác!", "Xác nhận xóa PO", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
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
                    MessageBox.Show(GetActiveOwner(), $"Đã xóa thành công đơn PO '{poNo}'!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information); ClearHeader(); dgvDetails.Rows.Clear(); dgvFiles.Rows.Clear(); UpdateTotal(); _selectedPO_ID = 0; LoadPO();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(GetActiveOwner(), "Lỗi khi xóa PO: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
        }

        private void BtnDeleteDetail_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("PO", "Xóa dòng", "Xóa dòng chi tiết")) return;
            if (dgvDetails.SelectedRows.Count == 0)
            {
                SafeWarn("Vui lòng chọn ít nhất một dòng trong danh sách vật tư để xóa!");
                return;
            }

            try
            {
                // Ngắt toàn bộ các sự kiện có thể gây loop hoặc block khi thay đổi cấu trúc/dữ liệu bảng
                dgvDetails.CellValueChanged -= DgvDetails_CellValueChanged;
                dgvDetails.CellEndEdit -= DgvDetails_CellEndEdit;
                dgvDetails.CellParsing -= DgvDetails_CellParsing;
                dgvDetails.CurrentCellDirtyStateChanged -= DgvDetails_CurrentCellDirtyStateChanged;
                dgvDetails.CellFormatting -= DgvDetails_CellFormatting;

                var rowsToDelete = new List<DataGridViewRow>();
                foreach (DataGridViewRow row in dgvDetails.SelectedRows)
                    if (!row.IsNewRow && row.Tag?.ToString() != "TOTAL") rowsToDelete.Add(row);
                foreach (var row in rowsToDelete) dgvDetails.Rows.Remove(row);
                
                int itemNo = 1;
                foreach (DataGridViewRow row in dgvDetails.Rows)
                    if (!row.IsNewRow && row.Tag?.ToString() != "TOTAL") row.Cells["Item_No"].Value = itemNo++;
                
                UpdateTotal();

                // Kết nối lại các sự kiện
                dgvDetails.CellValueChanged += DgvDetails_CellValueChanged;
                dgvDetails.CellEndEdit += DgvDetails_CellEndEdit;
                dgvDetails.CellParsing += DgvDetails_CellParsing;
                dgvDetails.CurrentCellDirtyStateChanged += DgvDetails_CurrentCellDirtyStateChanged;
                dgvDetails.CellFormatting += DgvDetails_CellFormatting;
            }
            catch (Exception ex)
            {
                SafeErr("Lỗi khi xóa: " + ex.Message);
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
                    // UNION: direct match (same revision) + cross-revision match by item_name
                    // Cross-revision part also checks base MPR_No (strips _Rev.X suffix) so we
                    // only match actual revisions of the same MPR, not unrelated MPRs in the project.
                    string sql = @"
SELECT md.Detail_ID AS CurDetailId, poh.PONo
FROM PO_Detail pod
INNER JOIN PO_head poh ON pod.PO_ID = poh.PO_ID
INNER JOIN MPR_Details md ON md.Detail_ID = pod.MPR_Detail_ID AND md.MPR_ID = @mprId
WHERE pod.MPR_Detail_ID IS NOT NULL AND ISNULL(poh.Status,'') <> 'Cancelled'

UNION

SELECT md_new.Detail_ID AS CurDetailId, poh.PONo
FROM PO_Detail pod
INNER JOIN PO_head poh ON pod.PO_ID = poh.PO_ID
INNER JOIN MPR_Details md_old ON md_old.Detail_ID = pod.MPR_Detail_ID AND md_old.MPR_ID <> @mprId
INNER JOIN MPR_Header mh_old ON mh_old.MPR_ID = md_old.MPR_ID
INNER JOIN MPR_Header mh_new ON mh_new.MPR_ID = @mprId
    AND mh_new.Project_Code = mh_old.Project_Code
    AND CASE WHEN CHARINDEX('_Rev.', mh_new.MPR_No) > 0
             THEN LEFT(mh_new.MPR_No, CHARINDEX('_Rev.', mh_new.MPR_No) - 1)
             ELSE mh_new.MPR_No END
      = CASE WHEN CHARINDEX('_Rev.', mh_old.MPR_No) > 0
             THEN LEFT(mh_old.MPR_No, CHARINDEX('_Rev.', mh_old.MPR_No) - 1)
             ELSE mh_old.MPR_No END
INNER JOIN MPR_Details md_new ON md_new.MPR_ID = @mprId
    -- Tránh khớp nhầm: Item_Name phải khác rỗng và Item_No phải giống nhau
    AND NULLIF(LTRIM(RTRIM(md_new.Item_Name)), '') IS NOT NULL
    AND NULLIF(LTRIM(RTRIM(md_old.Item_Name)), '') IS NOT NULL
    AND LTRIM(RTRIM(LOWER(md_new.Item_Name))) = LTRIM(RTRIM(LOWER(md_old.Item_Name)))
    AND LTRIM(RTRIM(ISNULL(md_new.Item_No, ''))) = LTRIM(RTRIM(ISNULL(md_old.Item_No, '')))
WHERE pod.MPR_Detail_ID IS NOT NULL AND ISNULL(poh.Status,'') <> 'Cancelled'";
                    using (var cmd = new Microsoft.Data.SqlClient.SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@mprId", mprId);
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                if (reader["CurDetailId"] != DBNull.Value)
                                {
                                    int detailId = Convert.ToInt32(reader["CurDetailId"]);
                                    string poNo = reader["PONo"]?.ToString() ?? "";
                                    if (dict.ContainsKey(detailId))
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

        private async void BtnImportMPR_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("PO", "Import MPR", "Import MPR")) return;
            using (var dlg = new frmSelectMPR())
            {
                if (dlg.ShowDialog() == DialogResult.OK && dlg.SelectedMPR != null)
                {
                    ClearHeader(); _selectedPO_ID = 0; _details.Clear(); dgvDetails.Rows.Clear(); dgvFiles.Rows.Clear();
                    var mpr = dlg.SelectedMPR;
                    var details = dlg.SelectedDetails;
                    txtProjectName.Text = mpr.Project_Name; txtMPRNo.Text = mpr.MPR_No;

                    // Load DB trên background thread
                    Dictionary<int, string> poMapping = null;
                    MPR_Managerment.Models.ProjectInfo project = null;
                    string autoPoNo = "", projectCode = "";
                    this.Cursor = Cursors.WaitCursor;
                    try
                    {
                        await Task.Run(() =>
                        {
                            poMapping = GetPoMappingForMpr(mpr.MPR_ID);
                            var projects = new ProjectService().GetAll();
                            project = projects.Find(p => !string.IsNullOrEmpty(p.ProjectName) && p.ProjectName.Equals(mpr.Project_Name, StringComparison.OrdinalIgnoreCase));
                            if (project == null) project = projects.Find(p => !string.IsNullOrEmpty(p.ProjectName) && (p.ProjectName.Contains(mpr.Project_Name, StringComparison.OrdinalIgnoreCase) || mpr.Project_Name.Contains(p.ProjectName, StringComparison.OrdinalIgnoreCase)));
                            if (project != null) { projectCode = project.ProjectCode; autoPoNo = GenerateAutoPoNo(project.POCode ?? project.ProjectCode ?? ""); }
                        });
                        if (project != null) { txtWorkorderNo.Text = project.WorkorderNo ?? ""; _projectCodeImport = projectCode; txtPONo.Text = autoPoNo; }
                        else MessageBox.Show(GetActiveOwner(), $"Không tìm thấy dự án khớp với tên \"{mpr.Project_Name}\".\nVui lòng kiểm tra lại thông tin Workorder và PO No.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(GetActiveOwner(), "Lỗi tìm project: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally { this.Cursor = Cursors.Default; }
                    _ = LoadMPRFilesAsync();
                    // Bỏ qua các dòng bị đánh dấu chỉ đọc/đã xóa khi revise
                    var activeDetails = details.Where(d => !d.Is_Deleted).ToList();
                    int skipped = details.Count - activeDetails.Count;
                    int itemNo = 1;
                    foreach (var d in activeDetails)
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
                    string skipNote = skipped > 0 ? $"\n(Đã bỏ qua {skipped} dòng chỉ đọc)" : "";
                    MessageBox.Show(GetActiveOwner(), $"✅ Đã import {activeDetails.Count} dòng từ MPR {mpr.MPR_No}!\nPO No dự kiến: {txtPONo.Text}\nWorkorder: {txtWorkorderNo.Text}{skipNote}", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        public async Task ImportMPRByNoAsync(string mprNo, int mprId = 0)
        {
            if (string.IsNullOrEmpty(mprNo)) return;
            this.Cursor = Cursors.WaitCursor;
            try
            {
                // Load tất cả dữ liệu từ DB trên background thread để không block UI
                MPR_Managerment.Models.MPRHeader mpr = null;
                List<MPR_Managerment.Models.MPRDetail> details = null;
                Dictionary<int, string> poMapping = null;
                MPR_Managerment.Models.ProjectInfo project = null;
                string autoPoNo = "";
                string projectCode = "";
                string projectErr = "";

                await Task.Run(() =>
                {
                    var mprService = new MPR_Managerment.Services.MPRService();
                    mpr = mprId > 0
                        ? mprService.GetById(mprId)
                        : mprService.GetByNo(mprNo);
                    if (mpr == null) return;
                    details = mprService.GetDetails(mpr.MPR_ID);
                    if (details == null || details.Count == 0) return;
                    poMapping = GetPoMappingForMpr(mpr.MPR_ID);
                    try
                    {
                        var projects = new MPR_Managerment.Services.ProjectService().GetAll();
                        project = projects.Find(p => !string.IsNullOrEmpty(p.ProjectName) && p.ProjectName.Equals(mpr.Project_Name, StringComparison.OrdinalIgnoreCase));
                        if (project == null) project = projects.Find(p => !string.IsNullOrEmpty(p.ProjectName) && (p.ProjectName.Contains(mpr.Project_Name, StringComparison.OrdinalIgnoreCase) || mpr.Project_Name.Contains(p.ProjectName, StringComparison.OrdinalIgnoreCase)));
                        if (project != null)
                        {
                            projectCode = project.ProjectCode;
                            autoPoNo = GenerateAutoPoNo(project.POCode ?? project.ProjectCode ?? "");
                        }
                    }
                    catch (Exception ex)
                    {
                        projectErr = ex.Message;
                    }
                });

                if (mpr == null) { SafeWarn($"Không tìm thấy MPR: {mprNo}"); return; }
                if (details == null || details.Count == 0) { SafeWarn($"MPR {mprNo} chưa có chi tiết vật tư!"); return; }
                if (!string.IsNullOrEmpty(projectErr))
                    MessageBox.Show(this, "Lỗi tìm project: " + projectErr, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                // Cập nhật UI trên main thread
                // Đưa việc khai báo biến ra ngoài khối try để chúng có phạm vi sử dụng ở phía sau
                List<MPR_Managerment.Models.MPRDetail> activeDetails = null;
                int skipped = 0;
                dgvDetails.SuspendLayout();
                try
                {
                    ClearHeader();
                    _selectedPO_ID = 0; _details.Clear(); dgvDetails.Rows.Clear(); dgvFiles.Rows.Clear();
                    txtProjectName.Text = mpr.Project_Name;
                    txtMPRNo.Text = mpr.MPR_No;
                    _ = LoadMPRFilesAsync();
                    if (project != null)
                    {
                        txtWorkorderNo.Text = project.WorkorderNo ?? "";
                        _projectCodeImport = projectCode;
                        txtPONo.Text = autoPoNo;
                    }

                    // Gán giá trị cho các biến đã khai báo trước
                    activeDetails = details.Where(d => !d.Is_Deleted).ToList();

                    // Nếu có danh sách Detail IDs được lọc từ frmMPR, chỉ lấy những dòng đó
                    if (_filteredMprDetailIds != null && _filteredMprDetailIds.Count > 0)
                    {
                        activeDetails = activeDetails.Where(d => _filteredMprDetailIds.Contains(d.Detail_ID)).ToList();
                    }

                    skipped = details.Count - activeDetails.Count;
                    int itemNo = 1;
                    foreach (var d in activeDetails)
                    {
                        string orderedPo = poMapping != null && poMapping.ContainsKey(d.Detail_ID) ? poMapping[d.Detail_ID] : "";
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
                }
                finally
                {
                    dgvDetails.ResumeLayout();
                }
                dgvPO.SelectionChanged -= DgvPO_SelectionChanged;
                dgvPO.ClearSelection();
                try { dgvPO.CurrentCell = null; } catch { }
                dgvPO.SelectionChanged += DgvPO_SelectionChanged;
                string skipNote = skipped > 0 ? $" | Bỏ qua {skipped} dòng chỉ đọc" : "";
                if (lblStatus != null)
                    lblStatus.Text = $"✅ Đã import {activeDetails.Count} dòng từ MPR {mpr.MPR_No} | PO: {txtPONo.Text}{skipNote}";
            }
            catch (Exception ex)
            {
                SafeErr("Lỗi import MPR: " + ex.Message);
            }
            finally
            {
                this.Cursor = Cursors.Default;
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
                       AND dt.PONo NOT IN (SELECT PONo FROM PO_head WHERE ISNULL(Status,'') = 'Cancelled')
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
                      AND ISNULL(Status,'Pending') NOT IN ('Done','Overdue','In Transit','Completed','Cancelled')";
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
                SafeWarn("Vui lòng chọn một dòng!");
                return;
            }
            int trackId = Convert.ToInt32(dgvDelivery.SelectedRows[0].Cells["TrackID"].Value ?? 0);
            if (trackId == 0) return;

            // Đọc ReceiverNote (kể cả khi ô đang ở chế độ edit) trước khi EndEdit
            string receiverNote = "";
            var currentCell = dgvDelivery.CurrentCell;
            if (currentCell != null && currentCell.RowIndex == dgvDelivery.SelectedRows[0].Index
                && currentCell.OwningColumn.Name == "ReceiverNote"
                && dgvDelivery.IsCurrentCellInEditMode)
            {
                receiverNote = (dgvDelivery.EditingControl as TextBox)?.Text ?? currentCell.EditedFormattedValue?.ToString() ?? "";
            }
            else
            {
                receiverNote = dgvDelivery.SelectedRows[0].Cells["ReceiverNote"].Value?.ToString() ?? "";
            }
            dgvDelivery.EndEdit();

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

                    // Check if all deliveries for this PO are done
                    var cmdCheck = new Microsoft.Data.SqlClient.SqlCommand(
                        "SELECT COUNT(*) FROM PO_DeliveryTracking WHERE PONo = (SELECT PONo FROM PO_DeliveryTracking WHERE TrackID = @id) AND Status != 'Done'", conn);
                    cmdCheck.Parameters.AddWithValue("@id", trackId);
                    int pendingCount = (int)cmdCheck.ExecuteScalar();

                    if (pendingCount == 0)
                    {
                        var cmdUpdateStatus = new Microsoft.Data.SqlClient.SqlCommand(
                            "UPDATE PO_head SET Status = 'Completed' WHERE PONo = (SELECT PONo FROM PO_DeliveryTracking WHERE TrackID = @id)", conn);
                        cmdUpdateStatus.Parameters.AddWithValue("@id", trackId);
                        cmdUpdateStatus.ExecuteNonQuery();
                        
                        // Update local list to reflect change in UI
                        var po = _poList.Find(p => p.PO_ID == _selectedPO_ID);
                        if (po != null) po.Status = "Completed";
                    }
                }
                LoadDeliveries();
                BindPOGrid(_poList); // Refresh the main grid to show "Completed"
            }
            catch (Exception ex)
            {
                MessageBox.Show(GetActiveOwner(), "Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SaveDeliveryReceiverNote()
        {
            if (dgvDelivery.SelectedRows.Count == 0)
            {
                SafeWarn("Vui lòng chọn một dòng!");
                return;
            }
            int trackId = Convert.ToInt32(dgvDelivery.SelectedRows[0].Cells["TrackID"].Value ?? 0);
            if (trackId == 0) return;

            // Đọc giá trị đang được edit (nếu có) TRƯỚC khi EndEdit làm mất nội dung
            string receiverNote = "";
            var currentCell = dgvDelivery.CurrentCell;
            if (currentCell != null && currentCell.RowIndex == dgvDelivery.SelectedRows[0].Index
                && currentCell.OwningColumn.Name == "ReceiverNote"
                && dgvDelivery.IsCurrentCellInEditMode)
            {
                receiverNote = (dgvDelivery.EditingControl as TextBox)?.Text ?? currentCell.EditedFormattedValue?.ToString() ?? "";
            }
            else
            {
                receiverNote = dgvDelivery.SelectedRows[0].Cells["ReceiverNote"].Value?.ToString() ?? "";
            }
            dgvDelivery.EndEdit();

            try
            {
                string sql = "UPDATE PO_DeliveryTracking SET ReceiverNote = @note WHERE TrackID = @id";
                using (var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var cmd = new Microsoft.Data.SqlClient.SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@note", receiverNote);
                    cmd.Parameters.AddWithValue("@id", trackId);
                    cmd.ExecuteNonQuery();
                }
                SafeInfo("✅ Đã lưu ghi chú thành công!");
                LoadDeliveries();
            }
            catch (Exception ex)
            {
                MessageBox.Show(GetActiveOwner(), "Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DeleteDeliveryRow()
        {
            if (dgvDelivery.SelectedRows.Count == 0)
            {
                SafeWarn("Vui lòng chọn một dòng!");
                return;
            }
            if (MessageBox.Show(GetActiveOwner(), "Xóa dòng theo dõi này?", "Xác nhận",
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
                    MessageBox.Show(GetActiveOwner(), "Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                if (po == null) { MessageBox.Show(GetActiveOwner(), $"Không tìm thấy PO: {poNo}", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

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
                popup.Owner = this; popup.Show();
            }
            catch (Exception ex) { MessageBox.Show(GetActiveOwner(), "Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
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
                btnOK.Click += async (s, ev) =>
                {
                    if (dgvDlg.SelectedRows.Count == 0)
                    {
                        SafeWarn("Vui lòng chọn một PO!");
                        return;
                    }
                    string selPONo = dgvDlg.SelectedRows[0].Cells["PONo"]?.Value?.ToString() ?? "";
                    string selDaAn = dgvDlg.SelectedRows[0].Cells["MaDuAn"]?.Value?.ToString() ?? "";
                    string selNCC  = dgvDlg.SelectedRows[0].Cells["NCC"]?.Value?.ToString() ?? "";

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

                    // Block UI trong suốt quá trình lưu + gửi thông báo
                    btnOK.Enabled     = false;
                    btnCancel.Enabled = false;
                    btnOK.Text        = "⏳ Đang gửi...";

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
                        // Gửi thông báo TRƯỚC KHI đóng dialog (tránh SynchronizationContext bị huỷ sau Close)
                        await SendDeliveryZaloNotificationAsync(selPONo, selDaAn, selNCC, expDate);
                        LoadDeliveries();
                        dlg.Close();
                    }
                    catch (Exception ex)
                    {
                        // Khôi phục nút nếu lỗi để user có thể thử lại
                        btnOK.Text        = "✔ OK";
                        btnOK.Enabled     = true;
                        btnCancel.Enabled = true;
                        MessageBox.Show(GetActiveOwner(), "Lỗi lưu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                dlg.Owner = this; dlg.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(GetActiveOwner(), "Lỗi mở popup: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Gửi thông báo lịch giao hàng lên nhóm Zalo của dự án
        private async System.Threading.Tasks.Task SendDeliveryZaloNotificationAsync(
            string poNo, string projectCode, string ncc, DateTime expDate)
        {
            try
            {
                var zSettings = ZaloHelper.LoadSettings();
                if (!zSettings.Enabled) return;

                // Lấy thông tin nhóm Zalo và chi tiết PO từ DB
                string zaloGroupName = "";
                string projectName   = "";
                System.Data.DataTable dtDetails;

                const string sqlInfo = @"
                    SELECT pi.ZaloGroupName, pi.ZaloGroupId, pi.ProjectName
                    FROM ProjectInfo pi
                    WHERE pi.ProjectCode = @pCode";

                const string sqlDetails = @"
                    SELECT pd.Item_No, pd.Item_Name, pd.Material,
                           pd.Asize, pd.Bsize, pd.Csize,
                           pd.Qty_Per_Sheet AS Qty, pd.UNIT
                    FROM PO_detail pd
                    INNER JOIN PO_head ph ON ph.PO_ID = pd.PO_ID
                    WHERE ph.PONo = @poNo
                    ORDER BY pd.Item_No";

                using (var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection())
                {
                    conn.Open();

                    // Lấy thông tin dự án / nhóm Zalo
                    if (!string.IsNullOrWhiteSpace(projectCode))
                    {
                        var cmdInfo = new Microsoft.Data.SqlClient.SqlCommand(sqlInfo, conn);
                        cmdInfo.Parameters.AddWithValue("@pCode", projectCode);
                        using var rdr = cmdInfo.ExecuteReader();
                        if (rdr.Read())
                        {
                            zaloGroupName = rdr["ZaloGroupName"]?.ToString() ?? "";
                            projectName   = rdr["ProjectName"]?.ToString() ?? "";
                            // Dùng ZaloGroupId nếu ZaloGroupName trống
                            if (string.IsNullOrWhiteSpace(zaloGroupName))
                                zaloGroupName = rdr["ZaloGroupId"]?.ToString() ?? "";
                        }
                    }

                    // Lấy chi tiết PO
                    var cmdDet = new Microsoft.Data.SqlClient.SqlCommand(sqlDetails, conn);
                    cmdDet.Parameters.AddWithValue("@poNo", poNo);
                    dtDetails = new System.Data.DataTable();
                    dtDetails.Load(cmdDet.ExecuteReader());
                }

                if (string.IsNullOrWhiteSpace(zaloGroupName)) return; // Không có nhóm Zalo

                // Xây nội dung tin nhắn
                var sb = new System.Text.StringBuilder();
                sb.AppendLine($"📦 LỊCH GIAO HÀNG DỰ KIẾN");
                sb.AppendLine($"━━━━━━━━━━━━━━━━━━━━━━━━");
                sb.AppendLine($"🔖 Số PO      : {poNo}");
                if (!string.IsNullOrWhiteSpace(projectName))
                    sb.AppendLine($"🏗️  Dự án     : {projectName}");
                if (!string.IsNullOrWhiteSpace(ncc))
                    sb.AppendLine($"🏭 Nhà CC     : {ncc}");
                sb.AppendLine($"📅 Exp.Deliv  : {expDate:dd/MM/yyyy}");

                if (dtDetails.Rows.Count > 0)
                {
                    sb.AppendLine($"━━━━━━━━━━━━━━━━━━━━━━━━");
                    sb.AppendLine("📋 Chi tiết vật tư:");
                    foreach (System.Data.DataRow row in dtDetails.Rows)
                    {
                        string itemNo   = row["Item_No"]?.ToString() ?? "";
                        string name     = row["Item_Name"]?.ToString() ?? "";
                        string material = row["Material"]?.ToString() ?? "";
                        string a        = row["Asize"]?.ToString() ?? "";
                        string b        = row["Bsize"]?.ToString() ?? "";
                        string c        = row["Csize"]?.ToString() ?? "";
                        string qty      = row["Qty"]?.ToString() ?? "";
                        string unit     = row["UNIT"]?.ToString() ?? "";

                        string size = string.Join(" x ", new[] { a, b, c }.Where(v => !string.IsNullOrWhiteSpace(v)));

                        sb.Append($"  {itemNo}. {name}");
                        if (!string.IsNullOrWhiteSpace(material)) sb.Append($" ({material})");
                        if (!string.IsNullOrWhiteSpace(size))     sb.Append($" — {size}");
                        if (!string.IsNullOrWhiteSpace(qty))      sb.Append($" — SL: {qty} {unit}");
                        sb.AppendLine();
                    }
                }

                string message = sb.ToString().TrimEnd();
                await ZaloHelper.SendToGroupAsync(zSettings, zaloGroupName, message);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Zalo] SendDeliveryZaloNotificationAsync error: {ex.Message}");
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
                        dt.TrackID                                              AS [TrackID],
                        dt.PONo                                                 AS [PO No],
                        ISNULL(s.Short_Name, ISNULL(s.Company_Name, N''))       AS [Nhà CC],
                        ISNULL(pi.ProjectCode, N'')                             AS [Mã dự án],
                        CONVERT(NVARCHAR(10), dt.ExpDelivery, 103)              AS [Exp. Delivery],
                        ISNULL(dt.GhiChu, N'')                                  AS [Ghi chú],
                        ISNULL(dt.Status, N'Pending')                           AS [Trạng thái],
                        ISNULL(dt.ReceiverNote, N'')                            AS [Receiver Note],
                        CONVERT(NVARCHAR(16), dt.Created_Date, 103)             AS [Ngày tạo]
                    FROM PO_DeliveryTracking dt
                    LEFT JOIN PO_head ph ON ph.PONo = dt.PONo
                    LEFT JOIN ProjectInfo pi ON pi.WorkorderNo = ph.WorkorderNo
                    LEFT JOIN Suppliers s ON s.Supplier_ID = ph.Supplier_ID
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

                // ── Thanh lọc ──────────────────────────────────────────────────
                // Layout: [🔍 PO:][txtPO 155px]  [Nhà CC:][txtNCC 155px]  [✖ Xóa]
                // Tổng ~590px — đủ rộng cho popup 900px
                var pFilter = new Panel
                {
                    Location = new Point(10, 62),
                    Size = new Size(popup.ClientSize.Width - 20, 32),
                    BackColor = Color.FromArgb(245, 245, 245),
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
                };
                popup.Controls.Add(pFilter);

                // x positions: label=0, txt=58, label2=225, txt2=283, btn=450
                pFilter.Controls.Add(new Label { Text = "🔍 PO:", Location = new Point(0, 7), Size = new Size(55, 18), Font = new Font("Segoe UI", 9) });
                var txtFilterPO = new TextBox
                {
                    Location = new Point(57, 4),
                    Size = new Size(160, 24),
                    Font = new Font("Segoe UI", 9),
                    PlaceholderText = "Nhập PO No..."
                };
                pFilter.Controls.Add(txtFilterPO);

                pFilter.Controls.Add(new Label { Text = "Nhà CC:", Location = new Point(225, 7), Size = new Size(52, 18), Font = new Font("Segoe UI", 9) });
                var txtFilterNCC = new TextBox
                {
                    Location = new Point(279, 4),
                    Size = new Size(160, 24),
                    Font = new Font("Segoe UI", 9),
                    PlaceholderText = "Tên viết tắt NCC..."
                };
                pFilter.Controls.Add(txtFilterNCC);

                var btnClearFilter = new Button
                {
                    Text = "✖ Xóa lọc",
                    Location = new Point(450, 3),
                    Size = new Size(80, 26),
                    BackColor = Color.FromArgb(108, 117, 125),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    Cursor = Cursors.Hand
                };
                btnClearFilter.FlatAppearance.BorderSize = 0;
                pFilter.Controls.Add(btnClearFilter);

                // DataGridView
                var dgv = new DataGridView
                {
                    Location = new Point(10, 98),
                    Size = new Size(popup.ClientSize.Width - 20, popup.ClientSize.Height - 144),
                    ReadOnly = false,
                    AllowUserToAddRows = false,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                    MultiSelect = true,
                    BackgroundColor = Color.White,
                    BorderStyle = BorderStyle.FixedSingle,
                    RowHeadersVisible = false,
                    Font = new Font("Segoe UI", 9),
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
                };

                dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 150, 100);
                dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                dgv.EnableHeadersVisualStyles = false;
                dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(235, 255, 245);
                dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
                dgv.DefaultCellStyle.SelectionForeColor = Color.Black;

                // ── Thêm cột unbound để tránh InvalidOperationException ──
                dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "TrackID", HeaderText = "TrackID", Visible = false });
                dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "PO No", HeaderText = "PO No", ReadOnly = true, FillWeight = 10 });
                dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Nhà CC", HeaderText = "Nhà CC", ReadOnly = true, FillWeight = 12 });
                dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Mã dự án", HeaderText = "Mã dự án", ReadOnly = true, FillWeight = 10 });
                dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Exp. Delivery", HeaderText = "Exp. Delivery", ReadOnly = true, FillWeight = 12 });
                dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ghi chú", HeaderText = "Ghi chú", ReadOnly = true, FillWeight = 16 });
                dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Trạng thái", HeaderText = "Trạng thái", ReadOnly = true, FillWeight = 10 });
                dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Receiver Note", HeaderText = "Receiver Note", ReadOnly = false, FillWeight = 20 });
                dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ngày tạo", HeaderText = "Ngày tạo", ReadOnly = true, FillWeight = 14 });

                foreach (System.Data.DataRow dr in dt.Rows)
                {
                    int ri = dgv.Rows.Add();
                    dgv.Rows[ri].Cells["TrackID"].Value = dr["TrackID"];
                    dgv.Rows[ri].Cells["PO No"].Value = dr["PO No"];
                    dgv.Rows[ri].Cells["Nhà CC"].Value = dr["Nhà CC"];
                    dgv.Rows[ri].Cells["Mã dự án"].Value = dr["Mã dự án"];
                    dgv.Rows[ri].Cells["Exp. Delivery"].Value = dr["Exp. Delivery"];
                    dgv.Rows[ri].Cells["Ghi chú"].Value = dr["Ghi chú"];
                    dgv.Rows[ri].Cells["Trạng thái"].Value = dr["Trạng thái"];
                    dgv.Rows[ri].Cells["Receiver Note"].Value = dr["Receiver Note"];
                    dgv.Rows[ri].Cells["Ngày tạo"].Value = dr["Ngày tạo"];
                }

                // ── Logic lọc theo PO No ──
                Action applyFilter = () =>
                {
                    string kwPO = txtFilterPO.Text.Trim().ToLower();
                    string kwNCC = txtFilterNCC.Text.Trim().ToLower();
                    foreach (DataGridViewRow r in dgv.Rows)
                    {
                        string poVal = r.Cells["PO No"].Value?.ToString()?.ToLower() ?? "";
                        string nccVal = r.Cells["Nhà CC"].Value?.ToString()?.ToLower() ?? "";
                        bool matchPO = string.IsNullOrEmpty(kwPO) || poVal.Contains(kwPO);
                        bool matchNCC = string.IsNullOrEmpty(kwNCC) || nccVal.Contains(kwNCC);
                        r.Visible = matchPO && matchNCC;
                    }
                };
                txtFilterPO.TextChanged += (s, ev) => applyFilter();
                txtFilterPO.KeyDown += (s, ev) => { if (ev.KeyCode == Keys.Enter) applyFilter(); };
                txtFilterNCC.TextChanged += (s, ev) => applyFilter();
                txtFilterNCC.KeyDown += (s, ev) => { if (ev.KeyCode == Keys.Enter) applyFilter(); };
                btnClearFilter.Click += (s, ev) => { txtFilterPO.Clear(); txtFilterNCC.Clear(); applyFilter(); };

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

                // Nút lưu ghi chú
                var btnSaveNote = new Button
                {
                    Text = "💾 Lưu ghi chú",
                    Size = new Size(120, 30),
                    BackColor = Color.FromArgb(255, 140, 0),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
                    Location = new Point(10, popup.ClientSize.Height - 40)
                };
                btnSaveNote.FlatAppearance.BorderSize = 0;
                btnSaveNote.Click += (s, ev) =>
                {
                    if (dgv.SelectedRows.Count == 0)
                    {
                        SafeWarn("Vui lòng chọn một dòng!");
                        return;
                    }
                    dgv.EndEdit();
                    var row = dgv.SelectedRows[0];
                    int trackId = Convert.ToInt32(row.Cells["TrackID"].Value ?? 0);
                    if (trackId == 0) return;
                    string note = row.Cells["Receiver Note"].Value?.ToString() ?? "";
                    try
                    {
                        string updateSql = "UPDATE PO_DeliveryTracking SET ReceiverNote = @note WHERE TrackID = @id";
                        using (var conn2 = MPR_Managerment.Helpers.DatabaseHelper.GetConnection())
                        {
                            conn2.Open();
                            var cmd2 = new Microsoft.Data.SqlClient.SqlCommand(updateSql, conn2);
                            cmd2.Parameters.AddWithValue("@note", note);
                            cmd2.Parameters.AddWithValue("@id", trackId);
                            cmd2.ExecuteNonQuery();
                        }
                        SafeInfo("✅ Đã lưu ghi chú thành công!");
                        LoadDeliveries();
                    }
                    catch (Exception ex2)
                    {
                        MessageBox.Show(GetActiveOwner(), "Lỗi lưu ghi chú: " + ex2.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                };
                popup.Controls.Add(btnSaveNote);

                // ── Nút Xóa dòng (yêu cầu mật khẩu Admin) ──
                var btnDel = new Button
                {
                    Text = "🗑 Xóa dòng",
                    Size = new Size(115, 30),
                    BackColor = Color.FromArgb(220, 53, 69),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
                    Cursor = Cursors.Hand,
                    Location = new Point(140, popup.ClientSize.Height - 40)
                };
                btnDel.FlatAppearance.BorderSize = 0;
                btnDel.Click += (s, ev) =>
                {
                    var selected = dgv.SelectedRows.Cast<DataGridViewRow>()
                        .Where(r => !r.IsNewRow)
                        .ToList();
                    if (selected.Count == 0)
                    {
                        SafeWarn("Vui lòng chọn ít nhất một dòng cần xóa!");
                        return;
                    }
                    string confirmMsg = selected.Count == 1
                        ? "Xóa dòng này khỏi Delivery Tracking History?"
                        : $"Xóa {selected.Count} dòng đã chọn khỏi Delivery Tracking History?";
                    if (MessageBox.Show(GetActiveOwner(), confirmMsg, "Xác nhận", MessageBoxButtons.YesNo,
                            MessageBoxIcon.Warning) != DialogResult.Yes) return;

                    // Yêu cầu mật khẩu Admin
                    if (!VerifyAdminPassword()) return;

                    try
                    {
                        using var conn3 = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                        conn3.Open();
                        int deleted = 0;
                        foreach (var row in selected)
                        {
                            int tid = Convert.ToInt32(row.Cells["TrackID"].Value ?? 0);
                            if (tid <= 0) continue;
                            var cmdDel = new Microsoft.Data.SqlClient.SqlCommand(
                                "DELETE FROM PO_DeliveryTracking WHERE TrackID = @id", conn3);
                            cmdDel.Parameters.AddWithValue("@id", tid);
                            deleted += cmdDel.ExecuteNonQuery();
                        }
                        // Xóa dòng khỏi grid
                        foreach (var row in selected)
                            if (dgv.Rows.Contains(row)) dgv.Rows.Remove(row);
                        LoadDeliveries();
                        MessageBox.Show(GetActiveOwner(), $"✅ Đã xóa {deleted} dòng thành công!", "Thành công",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex3)
                    {
                        MessageBox.Show(GetActiveOwner(), "Lỗi xóa: " + ex3.Message, "Lỗi",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                };
                popup.Controls.Add(btnDel);

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
                    pFilter.Width = popup.ClientSize.Width - 20;
                    dgv.Size = new Size(popup.ClientSize.Width - 20, popup.ClientSize.Height - 144);
                    btnClose.Location = new Point(popup.ClientSize.Width - 115, popup.ClientSize.Height - 40);
                    btnDel.Location = new Point(140, popup.ClientSize.Height - 40);
                    btnSaveNote.Location = new Point(10, popup.ClientSize.Height - 40);
                };

                popup.Owner = this; popup.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(GetActiveOwner(), "Lỗi tải lịch sử: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ClearHeader()
        {
            txtPONo.Text = ""; txtProjectName.Text = ""; txtWorkorderNo.Text = ""; txtMPRNo.Text = "";
            txtNotes.Text = "";
            nudRevise.Value = 0; dtpPODate.Value = DateTime.Today; cboStatus.SelectedIndex = 0; // Draft
            cboPaymentTerm.SelectedIndex = 0;
            // Reset Nha CC ve rong
            _isSearching = true; cboSupplier.Text = ""; cboSupplier.SelectedIndex = -1;
            BindSupplierCombo(_supplierTable); _isSearching = false;
            cboSupplier.BackColor = Color.White;
            cboSupplier.Enabled = true; _supplierLocked = false; btnUnlockSupplier.Visible = false;
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
                    WHERE ISNULL(ph.Status,'') <> 'Cancelled'
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
                    { SafeWarn("Không có dữ liệu để xuất!"); return; }

                    using var sfd = new SaveFileDialog
                    {
                        Title = "Xuất Check by Size",
                        Filter = "Excel Files|*.xlsx",
                        FileName = $"CheckBySize_{DateTime.Now:yyyyMMdd_HHmm}",
                        InitialDirectory = Directory.Exists(@"D:\RÁC") ? @"D:\RÁC" : Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
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
                        MessageBox.Show(GetActiveOwner(), $"✅ Đã xuất {dgv.Rows.Count} dòng thành công!", "Thành công",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        { FileName = sfd.FileName, UseShellExecute = true });
                    }
                    catch (Exception ex2)
                    {
                        MessageBox.Show(GetActiveOwner(), "Lỗi xuất Excel: " + ex2.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                popup.Owner = this; popup.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(GetActiveOwner(), "Lỗi khi tra cứu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

            // ── Ẩn/hiện cột giá & tổng tiền theo phân quyền ──────────────
            ApplyPriceVisibility();
        }

        /// <summary>
        /// Ẩn/hiện cột Đơn giá, TT trước thuế, Thành tiền và label tổng
        /// dựa theo quyền "Xem đơn giá", "Xem TT trước thuế", "Xem TT sau thuế"
        /// </summary>
        private void ApplyPriceVisibility()
        {
            bool canViewPrice = AppSession.HasPermission("PO", "Xem đơn giá");
            bool canViewSubTotal = AppSession.HasPermission("PO", "Xem TT trước thuế");
            bool canViewTotal = AppSession.HasPermission("PO", "Xem TT sau thuế");

            // ── Cột trong dgvDetails ──────────────────────────────────────
            if (dgvDetails != null)
            {
                if (dgvDetails.Columns.Contains("Price"))
                    dgvDetails.Columns["Price"].Visible = canViewPrice;
                if (dgvDetails.Columns.Contains("VAT"))
                    dgvDetails.Columns["VAT"].Visible = canViewPrice || canViewSubTotal || canViewTotal;
                if (dgvDetails.Columns.Contains("SubAmount"))
                    dgvDetails.Columns["SubAmount"].Visible = canViewSubTotal;
                if (dgvDetails.Columns.Contains("Amount"))
                    dgvDetails.Columns["Amount"].Visible = canViewTotal;
            }

            // ── Label tổng tiền phía dưới lưới ───────────────────────────
            if (lblSubTotal != null)
            {
                lblSubTotal.Visible = canViewSubTotal;
                if (!canViewSubTotal) lblSubTotal.Text = "";
            }
            if (lblTotal != null)
            {
                lblTotal.Visible = canViewTotal;
                if (!canViewTotal) lblTotal.Text = "";
            }
        }

        // =====================================================================
        //  EMAIL CELL FORMATTING
        // =====================================================================
        private void DgvPO_EmailCellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string colName = dgvPO.Columns[e.ColumnIndex].Name;

            // Đọc Email_Status của row này
            string status = "";
            if (dgvPO.Columns.Contains("Email_Status"))
                status = dgvPO.Rows[e.RowIndex].Cells["Email_Status"].Value?.ToString() ?? "";
            bool done = status.Equals("done", StringComparison.OrdinalIgnoreCase);

            if (colName == "col_SendEmail")
            {
                if (done)
                {
                    e.Value = "✔ Đã gửi";
                    e.CellStyle.BackColor = Color.FromArgb(200, 200, 200);
                    e.CellStyle.ForeColor = Color.Black;
                }
                else
                {
                    e.Value = "📧 Gửi";
                    e.CellStyle.BackColor = Color.FromArgb(40, 167, 69);
                    e.CellStyle.ForeColor = Color.White;
                }
                e.FormattingApplied = true;
            }
            else if (colName == "Email_Status")
            {
                if (done)
                {
                    e.Value = "✔ done";
                    e.CellStyle.ForeColor = Color.Black;
                    e.CellStyle.BackColor = Color.FromArgb(200, 200, 200);
                    e.FormattingApplied = true;
                }
            }
        }

        // =====================================================================
        //  EMAIL BUTTON CLICK
        // =====================================================================
        private void DgvPO_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (dgvPO.Columns[e.ColumnIndex].Name != "col_SendEmail") return;

            int poId = Convert.ToInt32(dgvPO.Rows[e.RowIndex].Cells["ID"].Value ?? 0);
            if (poId == 0) return;

            var po = _poList.Find(p => p.PO_ID == poId);
            if (po == null) return;

            // Nếu đã gửi email → hỏi người dùng muốn làm gì
            if (po.Email_Status?.Equals("done", StringComparison.OrdinalIgnoreCase) == true)
            {
                ShowEmailOrZaloChoicePO(po);
                return;
            }

            ShowEmailComposePopup(po);
        }

        // =====================================================================
        //  EMAIL ĐÃ GỬI — hỏi "Tiếp tục gửi Email" hay "Gửi tin nhắn thông báo"
        // =====================================================================
        private void ShowEmailOrZaloChoicePO(POHead po)
        {
            var dlg = new Form
            {
                Text = $"PO {po.PONo} — Đã gửi email",
                Size = new Size(410, 175),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false, MinimizeBox = false,
                BackColor = Color.White, Font = new Font("Segoe UI", 9.5f)
            };

            dlg.Controls.Add(new Label
            {
                Text = "✔ Email đã được gửi. Bạn muốn làm gì?",
                Location = new Point(16, 18), Size = new Size(370, 22),
                Font = new Font("Segoe UI", 9.5f)
            });

            var btnEmail = new Button
            {
                Text = "📧 Tiếp tục gửi Email",
                Location = new Point(16, 55), Size = new Size(175, 40),
                BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };
            btnEmail.FlatAppearance.BorderSize = 0;

            var btnZalo = new Button
            {
                Text = "🔔 Gửi tin nhắn thông báo",
                Location = new Point(208, 55), Size = new Size(175, 40),
                BackColor = Color.FromArgb(0, 150, 136), ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };
            btnZalo.FlatAppearance.BorderSize = 0;

            dlg.Controls.AddRange(new Control[] { btnEmail, btnZalo });

            btnEmail.Click += (s, ev) => { dlg.Close(); ShowEmailComposePopup(po); };
            btnZalo.Click  += (s, ev) => { dlg.Close(); ShowZaloGroupSelectAndSendPO(po); };

            dlg.ShowDialog(this);
        }

        // =====================================================================
        //  CHỌN NHÓM ZALO & GỬI THÔNG BÁO (PO)
        // =====================================================================
        private void ShowZaloGroupSelectAndSendPO(POHead po)
        {
            var cfg = Helpers.ZaloHelper.LoadSettings();
            if (!cfg.Enabled)
            {
                SafeWarn("Zalo chưa được bật.\nVào 🔔 Zalo trên thanh tiêu đề để bật.", "Zalo chưa bật");
                return;
            }

            var allGroups = new SupplierService().GetAll()
                .Where(s => !string.IsNullOrWhiteSpace(s.Zalo_Group_ID))
                .OrderBy(s => s.Company_Name)
                .ToList();

            if (allGroups.Count == 0)
            {
                SafeWarn("Không có nhà cung cấp nào có nhóm Zalo được cấu hình.\nVào 🏢 Nhà Cung Cấp để điền tên nhóm.", "Không có nhóm Zalo");
                return;
            }

            // checkedIds giữ Supplier_ID được tích — không bị mất khi lọc
            var checkedIds = new HashSet<int>();
            // filteredGroups là danh sách hiện tại đang hiển thị trong clb
            var filteredGroups = new List<MPR_Managerment.Models.Supplier>(allGroups);

            var dlg = new Form
            {
                Text = $"Chọn nhóm Zalo — PO {po.PONo}",
                Size = new Size(440, 570),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false, MinimizeBox = false,
                BackColor = Color.White, Font = new Font("Segoe UI", 9f)
            };

            dlg.Controls.Add(new Label
            {
                Text = "Chọn một hoặc nhiều nhóm Zalo để gửi thông báo:",
                Location = new Point(12, 12), Size = new Size(400, 20)
            });

            // ── Thanh tìm kiếm ───────────────────────────────────────────────
            var txtSearch = new TextBox
            {
                Location = new Point(12, 36), Size = new Size(404, 26),
                Font = new Font("Segoe UI", 9f),
                ForeColor = Color.Gray, Text = "🔍  Tìm kiếm nhóm...",
                BorderStyle = BorderStyle.FixedSingle
            };
            dlg.Controls.Add(txtSearch);

            var clb = new CheckedListBox
            {
                Location = new Point(12, 68),
                Size = new Size(404, 370),
                CheckOnClick = true,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9f)
            };
            foreach (var s in allGroups)
                clb.Items.Add($"{s.Company_Name}  [{s.Zalo_Group_ID}]", false);
            dlg.Controls.Add(clb);

            // Rebuild clb theo keyword, giữ checkmark qua checkedIds
            void Refilter(string keyword)
            {
                // Flush trạng thái check hiện tại vào checkedIds trước khi rebuild
                for (int i = 0; i < clb.Items.Count; i++)
                {
                    int sid = filteredGroups[i].Supplier_ID;
                    if (clb.GetItemChecked(i)) checkedIds.Add(sid);
                    else checkedIds.Remove(sid);
                }

                filteredGroups = string.IsNullOrWhiteSpace(keyword)
                    ? new List<MPR_Managerment.Models.Supplier>(allGroups)
                    : allGroups.Where(s =>
                        s.Company_Name.Contains(keyword, StringComparison.OrdinalIgnoreCase) ||
                        s.Zalo_Group_ID.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                      .ToList();

                clb.ItemCheck -= ClbItemCheck;
                clb.Items.Clear();
                foreach (var s in filteredGroups)
                    clb.Items.Add($"{s.Company_Name}  [{s.Zalo_Group_ID}]", checkedIds.Contains(s.Supplier_ID));
                clb.ItemCheck += ClbItemCheck;
            }

            void ClbItemCheck(object sender, ItemCheckEventArgs e)
            {
                int sid = filteredGroups[e.Index].Supplier_ID;
                if (e.NewValue == CheckState.Checked) checkedIds.Add(sid);
                else checkedIds.Remove(sid);
            }

            clb.ItemCheck += ClbItemCheck;

            // Placeholder logic
            txtSearch.GotFocus  += (s, ev) => { if (txtSearch.ForeColor == Color.Gray) { txtSearch.Text = ""; txtSearch.ForeColor = Color.Black; } };
            txtSearch.LostFocus += (s, ev) => { if (string.IsNullOrWhiteSpace(txtSearch.Text)) { txtSearch.ForeColor = Color.Gray; txtSearch.Text = "🔍  Tìm kiếm nhóm..."; } };
            txtSearch.TextChanged += (s, ev) =>
            {
                if (txtSearch.ForeColor == Color.Gray) return; // đang hiện placeholder
                Refilter(txtSearch.Text.Trim());
            };

            var btnSend = new Button
            {
                Text = "🔔 Gửi thông báo",
                Location = new Point(12, 455), Size = new Size(185, 38),
                BackColor = Color.FromArgb(0, 150, 136), ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };
            btnSend.FlatAppearance.BorderSize = 0;

            var btnAll = new Button
            {
                Text = "Chọn tất cả",
                Location = new Point(213, 455), Size = new Size(100, 38),
                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand
            };
            var btnCancel = new Button
            {
                Text = "Huỷ",
                Location = new Point(328, 455), Size = new Size(88, 38),
                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand
            };
            dlg.Controls.AddRange(new Control[] { btnSend, btnAll, btnCancel });

            btnAll.Click += (s, ev) =>
            {
                // Tích tất cả item đang hiện (filtered)
                for (int i = 0; i < clb.Items.Count; i++) clb.SetItemChecked(i, true);
            };
            btnCancel.Click += (s, ev) => dlg.Close();

            btnSend.Click += (s, ev) =>
            {
                // Flush lần cuối trước khi lấy kết quả
                for (int i = 0; i < clb.Items.Count; i++)
                {
                    int sid = filteredGroups[i].Supplier_ID;
                    if (clb.GetItemChecked(i)) checkedIds.Add(sid);
                    else checkedIds.Remove(sid);
                }

                var selected = allGroups.Where(s => checkedIds.Contains(s.Supplier_ID)).ToList();
                if (selected.Count == 0)
                {
                    MessageBox.Show(dlg, "Vui lòng chọn ít nhất một nhóm.",
                        "Chưa chọn nhóm", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                dlg.Close();

                string msg =
                    $"📋 DLHI Thông báo:\n" +
                    $"Chúng tôi đã gửi đơn đặt hàng (PO) cho:\n" +
                    $"🔹 Số PO     : {po.PONo}\n" +
                    $"🔹 Dự án    : {po.Project_Name}\n" +
                    $"🔹 Thời gian: {DateTime.Now:dd/MM/yyyy HH:mm}\n" +
                    $"Xin vui lòng kiểm tra Email, Rất mong sớm nhận được xác nhận từ Quý công ty!";

                _SendZaloToSelectedGroupsPO(selected, msg, po.PONo);
            };

            dlg.ShowDialog(this);
        }

        private void _SendZaloToSelectedGroupsPO(
            List<MPR_Managerment.Models.Supplier> selected, string msg, string poNo)
        {
            void SetSt(string text) { if (lblStatus != null && !lblStatus.IsDisposed) lblStatus.Text = text; }

            SetSt($"⏳ Đang đặt {selected.Count} tin nhắn Zalo vào hàng đợi...");

            foreach (var sup in selected)
            {
                // Lấy tên nhóm từ Zalo_Group_ID (giữ nguyên để dễ debug)
                string groupName = sup.Zalo_Group_ID ?? sup.Company_Name ?? "Unknown";
                ZaloNotificationService.Instance.Enqueue(groupName, msg, $"PO: {poNo}");
            }

            SetSt($"🔔 Đã đặt {selected.Count} tin nhắn Zalo vào hàng đợi (sẽ gửi theo thứ tự)");
        }

        // =====================================================================
        //  EMAIL SETTINGS DIALOG
        // =====================================================================
        private bool ShowEmailSettingsDialog()
        {
            var settings = Helpers.EmailHelper.LoadSettings();
            var dlg = new Form
            {
                Text = "Cài đặt Email (SMTP)",
                Size = new Size(440, 340),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false, MinimizeBox = false,
                BackColor = Color.White, Font = new Font("Segoe UI", 9)
            };

            int y = 15;
            Control MakeRow(string label, Control ctrl) {
                dlg.Controls.Add(new Label { Text = label, Location = new Point(15, y + 3), Size = new Size(120, 20), TextAlign = ContentAlignment.MiddleRight });
                ctrl.Location = new Point(140, y); ctrl.Width = 260; dlg.Controls.Add(ctrl); y += 35; return ctrl;
            }

            var txtHost = new TextBox { Text = settings.Host };
            var txtPort = new TextBox { Text = settings.Port.ToString() };
            var txtUser = new TextBox { Text = settings.Username };
            var txtPass = new TextBox { Text = settings.Password, UseSystemPasswordChar = true };
            var txtFrom = new TextBox { Text = settings.FromName };
            var chkSsl  = new CheckBox { Text = "Bật SSL/TLS", Checked = settings.EnableSsl, Location = new Point(140, y), AutoSize = true };

            MakeRow("SMTP Host:", txtHost);
            MakeRow("SMTP Port:", txtPort);
            MakeRow("Username (email):", txtUser);
            MakeRow("Password:", txtPass);
            MakeRow("Tên hiển thị:", txtFrom);
            dlg.Controls.Add(chkSsl); y += 35;

            var btnOk = new Button { Text = "Lưu", DialogResult = DialogResult.OK, Location = new Point(200, y), Size = new Size(90, 30), BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            var btnCancel = new Button { Text = "Huỷ", DialogResult = DialogResult.Cancel, Location = new Point(300, y), Size = new Size(90, 30), FlatStyle = FlatStyle.Flat };
            btnOk.FlatAppearance.BorderSize = 0;
            dlg.Controls.AddRange(new Control[] { btnOk, btnCancel });
            dlg.AcceptButton = btnOk; dlg.CancelButton = btnCancel;
            dlg.ClientSize = new Size(420, y + 50);

            if (dlg.ShowDialog(this) != DialogResult.OK) return false;

            settings.Host       = txtHost.Text.Trim();
            settings.Port       = int.TryParse(txtPort.Text, out int p) ? p : 587;
            settings.Username   = txtUser.Text.Trim();
            settings.Password   = txtPass.Text;
            settings.FromName   = txtFrom.Text.Trim();
            settings.EnableSsl  = chkSsl.Checked;
            Helpers.EmailHelper.SaveSettings(settings);
            return true;
        }

        // =====================================================================
        //  EMAIL COMPOSE + SEND VIA OUTLOOK (mailto:)
        // =====================================================================
        private void ShowEmailComposePopup(POHead po)
        {
            var suppliers = new SupplierService().GetAll();
            var supplier  = suppliers.Find(s => s.Supplier_ID == po.Supplier_ID);
            // Tách nhiều email phân biệt bởi ";" và loại bỏ khoảng trắng thừa
            string toEmail = string.Join("; ", (supplier?.Email ?? "")
                .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(e => e.Trim())
                .Where(e => !string.IsNullOrWhiteSpace(e)));

            if (string.IsNullOrWhiteSpace(toEmail))
            {
                SafeWarn($"Nhà cung cấp của PO \"{po.PONo}\" chưa có email trong hệ thống.\nVui lòng cập nhật thông tin nhà cung cấp trước.", "Thiếu email");
                return;
            }

            // Build body (không có tên user)
            var details = new POService().GetDetails(po.PO_ID);
            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"Kính gửi: {supplier?.Company_Name ?? "Quý Nhà Cung Cấp"},");
            sb.AppendLine();
            sb.AppendLine("Chúng tôi xin gửi Đơn Đặt Hàng (PO) với thông tin như sau:");
            sb.AppendLine($"  - Số PO   : {po.PONo}");
            sb.AppendLine($"  - Dự án   : {po.Project_Name}");
            sb.AppendLine($"  - Ngày PO : {po.PO_Date?.ToString("dd/MM/yyyy") ?? ""}");
            if (po.Revise > 0) sb.AppendLine($"  - Revise  : {po.Revise}");
            sb.AppendLine();
            sb.AppendLine("Chi tiết vật tư:");
            int stt = 1;
            foreach (var d in details)
                sb.AppendLine($"  {stt++}. {d.Item_Name}  |  SL: {d.Qty_Per_Sheet} {d.UNIT}  |  Đơn giá: {d.Price:N2}");
            sb.AppendLine();
            sb.AppendLine($"Tổng giá trị: {po.Total_Amount:N2}");
            sb.AppendLine();
            sb.AppendLine("Vui lòng xác nhận đơn hàng và liên hệ nếu có thắc mắc.");

            string subject = $"Đơn Đặt Hàng (PO) - {po.PONo}";
            string body    = sb.ToString();

            // Tìm PDF trước để hiện thị trong popup
            string pdfPath = FindPOPdfFile(po, supplier);

            // Popup xem trước
            var popup = new Form
            {
                Text = $"Gửi Email PO — {po.PONo}",
                Size = new Size(660, 560),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false, MinimizeBox = false,
                BackColor = Color.White, Font = new Font("Segoe UI", 9)
            };

            int py = 12;
            void AddRow(string lbl, Control ctrl, int h = 26)
            {
                popup.Controls.Add(new Label { Text = lbl, Location = new Point(10, py + 3), Size = new Size(65, 20), TextAlign = ContentAlignment.MiddleRight });
                ctrl.Location = new Point(80, py); ctrl.Size = new Size(560, h);
                popup.Controls.Add(ctrl); py += h + 6;
            }

            var txtTo      = new TextBox { Text = toEmail };
            var txtSubject = new TextBox { Text = subject };
            var txtBody    = new TextBox { Text = body, Multiline = true, ScrollBars = ScrollBars.Vertical };

            AddRow("Đến:",     txtTo);
            AddRow("Chủ đề:",  txtSubject);
            AddRow("Nội dung:", txtBody, 320);

            // Label + nút duyệt đính kèm
            bool hasPdf = !string.IsNullOrEmpty(pdfPath) && File.Exists(pdfPath);
            var lblAttach = new Label
            {
                Text      = hasPdf
                    ? $"📎 {Path.GetFileName(pdfPath)}"
                    : "⚠ Không tìm thấy file PDF trong thư mục PO Link",
                Location  = new Point(80, py), Size = new Size(470, 18),
                ForeColor = hasPdf ? Color.FromArgb(0, 130, 60) : Color.FromArgb(180, 80, 0),
                Font      = new Font("Segoe UI", 8.5f, FontStyle.Italic)
            };
            popup.Controls.Add(lblAttach);
            var btnBrowse = new Button
            {
                Text = "...", Location = new Point(555, py - 2), Size = new Size(85, 22),
                FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8)
            };
            popup.Controls.Add(btnBrowse);
            py += 28;

            // Buttons
            var btnSend  = new Button { Text = "📧 Mở trong Outlook", Location = new Point(10, py), Size = new Size(165, 34), BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Cursor = Cursors.Hand };
            var btnClose = new Button { Text = "Huỷ", Location = new Point(560, py), Size = new Size(80, 34), FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnSend.FlatAppearance.BorderSize = 0;
            popup.Controls.AddRange(new Control[] { btnSend, btnClose });
            popup.ClientSize = new Size(650, py + 54);

            string resolvedAttach = pdfPath;

            btnBrowse.Click += (s, ev) =>
            {
                using var ofd = new OpenFileDialog
                {
                    Title  = "Chọn file PO đính kèm",
                    Filter = "Excel / PDF|*.xlsx;*.xls;*.pdf|Tất cả|*.*",
                    InitialDirectory = string.IsNullOrEmpty(resolvedAttach) ? "" : Path.GetDirectoryName(resolvedAttach) ?? ""
                };
                if (ofd.ShowDialog() != DialogResult.OK) return;
                resolvedAttach      = ofd.FileName;
                lblAttach.Text      = $"📎 {Path.GetFileName(resolvedAttach)}";
                lblAttach.ForeColor = Color.FromArgb(0, 130, 60);
            };

            btnClose.Click += (s, ev) => popup.Close();

            btnSend.Click += (s, ev) =>
            {
                if (string.IsNullOrWhiteSpace(txtTo.Text)) { SafeWarn("Chưa có email nhận!"); return; }
                btnSend.Enabled = false;
                btnSend.Text    = "Đang mở Outlook...";
                try
                {
                    var outlookType = Type.GetTypeFromProgID("Outlook.Application");
                    if (outlookType == null) throw new Exception("Không tìm thấy Outlook trên máy tính này.");

                    dynamic outlook = Activator.CreateInstance(outlookType);
                    dynamic mail    = outlook.CreateItem(0); // olMailItem

                    mail.To      = txtTo.Text.Trim();
                    mail.Subject = txtSubject.Text.Trim();

                    // Chèn nội dung vào trước chữ ký Outlook
                    dynamic inspector  = mail.GetInspector;
                    string htmlWithSig = (string)(mail.HTMLBody ?? "");
                    string ourHtml     = BuildHtmlBody(txtBody.Text);
                    if (string.IsNullOrEmpty(htmlWithSig))
                    {
                        mail.HTMLBody = ourHtml;
                    }
                    else
                    {
                        int tagEnd = htmlWithSig.IndexOf('>', htmlWithSig.IndexOf("<body", StringComparison.OrdinalIgnoreCase));
                        mail.HTMLBody = tagEnd >= 0 ? htmlWithSig.Insert(tagEnd + 1, ourHtml) : ourHtml + htmlWithSig;
                    }

                    // Đính kèm file nếu tìm thấy
                    if (!string.IsNullOrEmpty(resolvedAttach) && File.Exists(resolvedAttach))
                        mail.Attachments.Add(resolvedAttach, 1, Type.Missing, Path.GetFileName(resolvedAttach));

                    // Mở cửa sổ soạn thảo Outlook (không gửi ngay)
                    mail.Display(false);

                    popup.Close();
                    _StartPOEmailSentPoller(outlook, mail, txtSubject.Text.Trim(), po.PO_ID, supplier);
                }
                catch (Exception ex)
                {
                    btnSend.Enabled = true;
                    btnSend.Text    = "📧 Mở trong Outlook";
                    SafeErr($"Lỗi mở Outlook:\n{ex.Message}", "Lỗi");
                }
            };

            popup.ShowDialog(this);
        }

        // =====================================================================
        //  TÌM FILE PDF CỦA PO (theo PO_Link dự án, thay Excel→PDF)
        // =====================================================================
        private string FindPOPdfFile(POHead po, Supplier supplier)
        {
            try
            {
                var projects = new ProjectService().GetAll();
                var prj = projects.Find(p =>
                    (!string.IsNullOrEmpty(p.WorkorderNo)  && p.WorkorderNo.Equals(po.WorkorderNo,   StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(p.ProjectName)  && p.ProjectName.Equals(po.Project_Name,  StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(p.ProjectCode)  && p.ProjectCode.Equals(po.ProjectCode,   StringComparison.OrdinalIgnoreCase)));

                if (prj == null || string.IsNullOrEmpty(prj.PO_Link)) return null;

                // Xác định thư mục PDF
                string configured = prj.PO_Link.Trim()
                    .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                string lastName = Path.GetFileName(configured) ?? "";
                string pdfDir   = lastName.Equals("Excel", StringComparison.OrdinalIgnoreCase)
                    ? Path.Combine(Path.GetDirectoryName(configured) ?? configured, "PDF")
                    : configured;

                if (!Directory.Exists(pdfDir)) return null;

                // Chuẩn hoá token tìm kiếm
                string poNo      = Normalize(po.PONo);

                string suppShort = Normalize(supplier?.Short_Name ?? "");

                // Lấy tất cả PDF trong thư mục
                var candidates = Directory.GetFiles(pdfDir, "*.pdf", SearchOption.TopDirectoryOnly);

                // Chấm điểm: ưu tiên file có cả số PO lẫn tên viết tắt NCC
                string bestPath  = null;
                int    bestScore = -1;

                foreach (var f in candidates)
                {
                    string nameNorm = Normalize(Path.GetFileNameWithoutExtension(f));
                    int score = 0;
                    if (!string.IsNullOrEmpty(poNo)      && nameNorm.Contains(poNo))      score += 2;
                    if (!string.IsNullOrEmpty(suppShort) && nameNorm.Contains(suppShort)) score += 1;
                    if (score > bestScore) { bestScore = score; bestPath = f; }
                }

                // Chỉ trả về nếu ít nhất khớp số PO (score >= 2)
                return bestScore >= 2 ? bestPath : null;
            }
            catch { return null; }
        }

        private static string Normalize(string s) =>
            (s ?? "").ToUpperInvariant()
                     .Replace(" ", "")
                     .Replace("-", "")
                     .Replace("_", "")
                     .Replace(".", "");

        private static string BuildHtmlBody(string plainText)
        {
            // Chuyển plain text → HTML, giữ xuống dòng và khoảng trắng
            string html = plainText
                .Replace("&",  "&amp;")
                .Replace("<",  "&lt;")
                .Replace(">",  "&gt;")
                .Replace("\r\n", "<br>")
                .Replace("\n",   "<br>")
                .Replace("  ",  "&nbsp;&nbsp;");
            return $"<div style='font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000000;'>{html}</div><br>";
        }

        // =====================================================================
        //  POLL mail.Sent — tự cập nhật status khi Outlook thực sự bấm Send
        //  Fallback: khi COM object bị giải phóng (sau khi gửi) → check Sent Items
        // =====================================================================
        private void _StartPOEmailSentPoller(dynamic outlook, dynamic mail, string subject, int poId,
            MPR_Managerment.Models.Supplier supplier = null)
        {
            if (_emailPollers.TryGetValue(poId, out var old))
            {
                old.Stop(); old.Dispose();
                _emailPollers.Remove(poId);
            }

            if (lblStatus != null)
                lblStatus.Text = $"⏳ Đang chờ Outlook gửi email cho PO #{poId}...";

            var startedAt = DateTime.Now;
            var timer     = new System.Windows.Forms.Timer { Interval = 2000 };
            var deadline  = DateTime.Now.AddMinutes(15);
            bool comDead  = false;
            _emailPollers[poId] = timer;

            timer.Tick += (_, __) =>
            {
                if (DateTime.Now > deadline)
                {
                    timer.Stop(); timer.Dispose();
                    _emailPollers.Remove(poId);
                    if (lblStatus != null) lblStatus.Text = "⚠ Hết thời gian chờ — email chưa được ghi nhận.";
                    return;
                }

                // Pha 1: COM còn sống → kiểm tra mail.Sent
                if (!comDead)
                {
                    bool sent = false;
                    try { sent = (bool)mail.Sent; }
                    catch { comDead = true; }

                    if (sent)
                    {
                        _MarkPOEmailDone(poId, supplier);
                        timer.Stop(); timer.Dispose();
                        _emailPollers.Remove(poId);
                        return;
                    }
                }

                // Pha 2: COM đã chết → tiếp tục poll Sent Items mỗi 2s
                if (comDead)
                {
                    int elapsed = (int)(DateTime.Now - startedAt).TotalSeconds;
                    if (lblStatus != null)
                        lblStatus.Text = $"⏳ Đang kiểm tra Sent Items... ({elapsed}s)";

                    if (_POCheckSentItemsFolder(subject))
                    {
                        _MarkPOEmailDone(poId, supplier);
                        timer.Stop(); timer.Dispose();
                        _emailPollers.Remove(poId);
                    }
                    // Chưa thấy → giữ timer tiếp tục, không dừng
                    return;
                }

                int sec = (int)(DateTime.Now - startedAt).TotalSeconds;
                if (lblStatus != null)
                    lblStatus.Text = $"⏳ Chờ Outlook gửi... ({sec}s)";
            };

            timer.Start();
        }

        private void _MarkPOEmailDone(int poId,
            MPR_Managerment.Models.Supplier supplier = null)
        {
            var sentAt = DateTime.Now;
            new POService().UpdateEmailStatus(poId, "done", _currentUser, sentAt);
            var h = _poList.Find(x => x.PO_ID == poId);
            if (h != null)
            {
                h.Email_Status  = "done";
                h.Email_Sent_By = _currentUser;
                h.Email_Sent_At = sentAt;
            }
            BindPOGrid(_poList);
            if (lblStatus != null)
                lblStatus.Text = "✅ Email PO đã được gửi thành công!";

            // Tự động chuyển trạng thái PO thành "Pending" sau khi gửi mail
            try
            {
                using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                conn.Open();
                var cmd = new Microsoft.Data.SqlClient.SqlCommand(
                    "UPDATE PO_head SET Status = 'Pending' WHERE PO_ID = @poId", conn);
                cmd.Parameters.AddWithValue("@poId", poId);
                cmd.ExecuteNonQuery();
            }
            catch { }
            
            // Cập nhật UI và list
            var poPending = _poList.Find(p => p.PO_ID == poId);
            if (poPending != null) poPending.Status = "Pending";
            BindPOGrid(_poList);

            // Gửi thông báo Zalo vào nhóm NCC
            if (supplier != null && h != null)
            {
                string zaloMsg =
                    $"📋 DLHI Thông báo:\n" +
                    $"Chúng tôi đã gửi Đơn Đặt Hàng (PO) đến Quý công ty:\n" +
                    $"🔹 Số PO    : {h.PONo}\n" +
                    $"🔹 Dự án   : {h.Project_Name}\n" +
                    $"🔹 Thời gian: {sentAt:dd/MM/yyyy HH:mm}\n" +
                    $"Xin vui lòng kiểm tra Email, Rất mong sớm nhận được xác nhận từ Quý công ty!";
                _SendZaloPONotify(supplier, zaloMsg);
            }
        }

        private void _SendZaloPONotify(
            MPR_Managerment.Models.Supplier supplier, string msg)
        {
            void SetSt(string text) { if (lblStatus != null) lblStatus.Text = text; }

            if (string.IsNullOrWhiteSpace(supplier?.Zalo_Group_ID))
            {
                SetSt("✅ Email PO đã gửi  |  ⚠ NCC chưa có tên nhóm Zalo — vào Nhà Cung Cấp để điền");
                return;
            }

            SetSt("✅ Email PO đã gửi  |  ⏳ Đang đưa thông báo Zalo vào hàng đợi...");
            
            string groupName = supplier.Zalo_Group_ID;
            // Lấy PO No để làm tag
            string poNo = _selectedPO_ID > 0 ? _poList.Find(p => p.PO_ID == _selectedPO_ID)?.PONo : "N/A";
            
            ZaloNotificationService.Instance.Enqueue(groupName, msg, $"PO: {poNo}");
            
            SetSt("✅ Email PO đã gửi  |  🔔 Thông báo Zalo đã được xếp hàng gửi");
        }

        private bool _POCheckSentItemsFolder(string subject)
        {
            try
            {
                // Tạo instance Outlook mới (singleton COM — trả về Outlook đang chạy)
                var outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null) return false;
                dynamic ol = Activator.CreateInstance(outlookType);
                dynamic ns = ol.GetNamespace("MAPI");
                try
                {
                    dynamic stores     = ns.Stores;
                    int     storeCount = Convert.ToInt32(stores.Count);
                    for (int s = 1; s <= storeCount; s++)
                    {
                        try
                        {
                            dynamic store = stores[s];
                            dynamic sentFolder;
                            try { sentFolder = store.GetDefaultFolder(5); } catch { continue; }
                            if (_POScanFolderForSubject(sentFolder, subject)) return true;
                        }
                        catch { }
                    }
                }
                catch { }
                try
                {
                    dynamic sentFolder = ns.GetDefaultFolder(5);
                    if (_POScanFolderForSubject(sentFolder, subject)) return true;
                }
                catch { }
                return false;
            }
            catch { return false; }
        }

        private bool _POScanFolderForSubject(dynamic folder, string subject)
        {
            try
            {
                dynamic items = folder.Items;
                items.Sort("[SentOn]", true);
                int count = Convert.ToInt32(items.Count);
                for (int i = 1; i <= Math.Min(100, count); i++)
                {
                    try
                    {
                        dynamic item = items[i];
                        string  itemSubj;
                        try { itemSubj = (string)item.Subject; } catch { continue; }
                        if (string.Equals(itemSubj, subject, StringComparison.OrdinalIgnoreCase)) return true;
                    }
                    catch { }
                }
                return false;
            }
            catch { return false; }
        }

    }
}
