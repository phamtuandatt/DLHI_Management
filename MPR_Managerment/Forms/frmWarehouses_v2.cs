using Microsoft.Data.SqlClient;
using Microsoft.IdentityModel.Tokens;
using MPR_Managerment.Forms.DeliveryGUI;
using MPR_Managerment.Forms.ExportGUI;
using MPR_Managerment.Forms.ImportWarehouseGUI;
using MPR_Managerment.Forms.ItemCodeGUI;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;
using MPR_Managerment.Services;
using OfficeOpenXml;
using OfficeOpenXml.Packaging.Ionic.Zlib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Design;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static OfficeOpenXml.ExcelErrorValue;

namespace MPR_Managerment.Forms
{
    public partial class frmWarehouses_v2 : Form
    {
        private TabControl mainTabControl;
        private TabPage pageImport, pageExport, pageWarehouse, pageFillInvoiceNo, pageFillInvoiceNo_v2;
        private TabPage pageSaveDelivertNote;
        private List<ProjectInfo> _dtProject = new List<ProjectInfo>();

        private Button btnSearch, btnCancelSearch, btnSearchHistory;

        private Button btnAdd, btnSave, btnCancel, btnDeleteRow;
        private ComboBox cboProject, cboPONo;
        private Button btnPrintPNK, btnDeleteFullBill;
        private ComboBox cboFilterProject, cboFilterPO;

        private DataGridView dgvImportQueue, dgvImport;

        private WarehouseService _service = new WarehouseService();
        private POService _poService = new POService();
        private WarehouseLocationService _warehouseService = new WarehouseLocationService();
        private string _currentUser = "Admin";
        private ProductServices _productServices = new ProductServices();
        private List<WarehouseImport> _imports = new List<WarehouseImport>();
        private List<WarehouseImport> _importQueue = new List<WarehouseImport>();
        private int _selectedImportID = 0;
        private int _pendingPO_ID = 0;
        private string _currentBatchNo = "";

        private Dictionary<string, string> _importList = new Dictionary<string, string>();
        private object oldValue = null;
        private bool _useItemCodeExisted = false;

        // ===== TỒN KHO =====
        private DataGridView dgvStock;
        private TextBox txtSearchStock;
        private ComboBox cboProjectFilter;
        private Label lblStockTotal, lblStockQty, lblStockWeight;
        private Panel panelStockSummary;

        private string _targetPONo = "";

        public frmWarehouses_v2(string targetPONo = "")
        {
            _targetPONo = targetPONo; // Tham số nhận từ màn hình Dashboard
            InitializeComponent();
            _dtProject = new ProjectService().GetAll();
            BuidUI();
            SetupImportLayout(pageImport);
            SetupExportLayout(pageExport);
            SetupFillInvoiceNotLayout(pageFillInvoiceNo);
            SetupFillInvoiceNoLayout_v2(pageFillInvoiceNo_v2);
            SetupSaveDeliveryNotetLayout(pageSaveDelivertNote);
            TrackButtonClick();
            LoadComboboxProject();
            HandleComboBoxIndexChange();
            BuildQueueColumns();
            BuildStockTab_V2(pageWarehouse);

            this.Load += FrmWarehouses_v2_Load;


            dgvImportQueue.Columns["ID_Code"].ReadOnly = false;
            //Button btnPaste = new Button()
            //{
            //    Text = " 📋 Dán từ Excel",
            //    Size = new Size(130, 35),
            //    BackColor = Color.FromArgb(46, 204, 113), // Màu xanh lá nhẹ (Emerald)
            //    ForeColor = Color.White,
            //    FlatStyle = FlatStyle.Flat,
            //    Font = new Font("Segoe UI", 9, FontStyle.Bold),
            //    Cursor = Cursors.Hand,
            //    //Location = new Point(btnDeleteRow.Location.X + btnDeleteRow.Width + 10, btnDeleteRow.Location.Y) // Đặt bên cạnh nút "Thêm vào phiếu"
            //    Location = new Point(150, 500)
            //};
            //btnPaste.BringToFront();
            //btnPaste.Click += (s, e) => PasteToEditableCells();
            //this.Controls.Add(btnPaste);
        }

        private void FrmWarehouses_v2_Load(object? sender, EventArgs e)
        {
            LoadAll();

            // Tự động nhảy đến PO được chọn từ Dashboard
            if (!string.IsNullOrEmpty(_targetPONo))
            {
                cboProject.SelectedIndex = 0;
                LoadPOFilterByProject("");

                for (int i = 0; i < cboPONo.Items.Count; i++)
                {
                    if (cboPONo.Items[i].ToString() == _targetPONo)
                    {
                        cboPONo.SelectedIndex = i;
                        break;
                    }
                }
                mainTabControl.SelectedTab = pageImport;
            }
        }

        public void BuidUI()
        {
            mainTabControl = new TabControl();
            mainTabControl.Dock = DockStyle.Fill;
            mainTabControl.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            pageImport = new TabPage();
            pageImport.Text = "  📥  Nhập kho  ";
            pageImport.BackColor = Color.White;

            pageExport = new TabPage();
            pageExport.Text = "  📤  Xuất kho  ";
            pageExport.BackColor = Color.White;

            pageWarehouse = new TabPage();
            pageWarehouse.Text = "  📦  Tồn kho  ";
            pageWarehouse.BackColor = Color.White;

            pageFillInvoiceNo = new TabPage();
            pageFillInvoiceNo.Text = "🧾 Xem Hóa đơn";
            pageFillInvoiceNo.BackColor = Color.White;

            pageFillInvoiceNo_v2 = new TabPage();
            pageFillInvoiceNo_v2.Text = "📝 Hóa đơn";
            pageFillInvoiceNo_v2.BackColor = Color.White;

            pageSaveDelivertNote = new TabPage();
            pageSaveDelivertNote.Text = "🗄️ Phiếu giao hàng";
            pageSaveDelivertNote.BackColor = Color.White;

            // 5. Thêm các Page vào TabControl
            mainTabControl.TabPages.Add(pageWarehouse);
            mainTabControl.TabPages.Add(pageImport);
            mainTabControl.TabPages.Add(pageExport);
            mainTabControl.TabPages.Add(pageFillInvoiceNo);
            mainTabControl.TabPages.Add(pageFillInvoiceNo_v2);
            mainTabControl.TabPages.Add(pageSaveDelivertNote);

            this.Controls.Add(mainTabControl);
        }

        public void SetupImportLayout(TabPage parent)
        {
            Panel mainScrollPanel = new Panel();
            mainScrollPanel.Dock = DockStyle.Fill;
            mainScrollPanel.AutoScroll = true;
            parent.Controls.Add(mainScrollPanel);

            Panel container = new Panel();
            container.Width = 1300;
            container.Height = 1200;
            container.Location = new Point(0, 0);
            mainScrollPanel.Controls.Add(container);

            GroupBox gbHeader = new GroupBox();
            gbHeader.Text = "Bộ lọc tìm kiếm";
            gbHeader.Size = new Size(1280, 70);
            gbHeader.Location = new Point(10, 10);
            container.Controls.Add(gbHeader);

            Label lblProject = new Label()
            {
                Text = "Dự án:",
                Location = new Point(15, 33),
                AutoSize = true
            };
            cboProject = new ComboBox()
            {
                Name = "cbProject",
                Location = new Point(70, 30),
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.ListItems
            };
            //cboProject.Validating += CboProject_Validating;

            Label lblPONo = new Label()
            {
                Text = "PO NO:",
                Location = new Point(290, 33),
                AutoSize = true
            };
            cboPONo = new ComboBox()
            {
                Name = "cbPONo",
                Location = new Point(350, 30),
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.ListItems
            };
            //cboPONo.Validating += CboPONo_Validating;

            btnSearch = new Button()
            {
                Name = "btnSearch",
                Text = "🔍 Tìm",
                Location = new Point(570, 28),
                Size = new Size(100, 30),
                Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(0, 120, 212),
                FlatStyle = FlatStyle.Flat,
            };
            btnCancelSearch = new Button()
            {
                Name = "btnCancelSearch",
                Text = "✖ Xóa lọc",
                Location = new Point(700, 28),
                Size = new Size(100, 30),
                Cursor = Cursors.Hand,
                BackColor = Color.FromArgb(108, 117, 125),
                FlatStyle = FlatStyle.Flat,
            };
            gbHeader.Controls.AddRange(new Control[] { lblProject, cboProject, lblPONo, cboPONo, btnSearch, btnCancelSearch });

            GroupBox gbActions = new GroupBox();
            gbActions.Text = "Thao tác nghiệp vụ";
            gbActions.Size = new Size(1280, 80);
            gbActions.Location = new Point(10, 90);
            container.Controls.Add(gbActions);

            btnAdd = new Button()
            {
                Text = "➕ Thêm vào phiếu",
                Location = new Point(15, 30),
                Size = new Size(120, 35),
                BackColor = Color.FromArgb(255, 140, 0),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnSave = new Button()
            {
                Text = "💾 Lưu phiếu nhập",
                Location = new Point(145, 30),
                Size = new Size(120, 35),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnCancel = new Button()
            {
                Text = "🆕 Phiếu mới",
                Location = new Point(275, 30),
                Size = new Size(120, 35),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnCancel.Click += BtnCancel_Click;
            gbActions.Controls.AddRange(new Control[] { btnAdd, btnSave, btnCancel });

            GroupBox gbDetails = new GroupBox();
            gbDetails.Text = "Chi tiết nhập kho";
            gbDetails.Size = new Size(1280, 400);
            gbDetails.Location = new Point(10, 180);
            container.Controls.Add(gbDetails);

            Label lblDetail = new Label()
            {
                Text = "Danh sách vật tư trong phiếu",
                Location = new Point(15, 25),
                AutoSize = true,
                ForeColor = Color.Orange,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };

            btnDeleteRow = new Button()
            {
                Text = "🗑 Xóa dòng chọn",
                Location = new Point(gbDetails.Width - 150, 20),
                Size = new Size(130, 35),
                BackColor = Color.FromArgb(220, 53, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
            };
            btnDeleteRow.Click += BtnDeleteRow_Click;

            Button btnPaste = new Button()
            {
                Text = " 📋 Dán từ Excel",
                Size = new Size(130, 35),
                Location = new Point(btnDeleteRow.Location.X - btnDeleteRow.Width + 20, btnDeleteRow.Location.Y), // Đặt bên cạnh nút "Thêm vào phiếu"
                BackColor = Color.FromArgb(46, 204, 113), // Màu xanh lá nhẹ (Emerald)
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand
                //Location = new Point(gbDetails.Width - 500, 20),
            };
            btnPaste.Click += (s, e) => PasteToEditableCells();
            btnPaste.Visible = false;

            dgvImportQueue = new DataGridView()
            {
                Location = new Point(15, 60),
                Size = new Size(gbDetails.Width - 40, 320),
                BackgroundColor = Color.White,
                AutoGenerateColumns = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            };
            dgvImportQueue.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 140, 0);
            dgvImportQueue.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvImportQueue.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvImportQueue.EnableHeadersVisualStyles = false;
            dgvImportQueue.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(255, 248, 235);

            // Xanh nhạt cho selection
            dgvImportQueue.DefaultCellStyle.SelectionBackColor = Color.FromArgb(204, 232, 255);
            dgvImportQueue.DefaultCellStyle.SelectionForeColor = Color.Black;

            dgvImportQueue.CellBeginEdit += DgvImportQueue_CellBeginEdit;
            dgvImportQueue.CellEndEdit += DgvImportQueue_CellEndEdit;
            dgvImportQueue.EditingControlShowing += DgvImportQueue_EditingControlShowing;
            dgvImportQueue.CellDoubleClick += DgvImportQueue_CellDoubleClick;
            dgvImportQueue.KeyDown += DgvImportQueue_KeyDown;

            gbDetails.Controls.AddRange(new Control[] { lblDetail, btnPaste, btnDeleteRow, dgvImportQueue });

            //GroupBox gbHistory = new GroupBox();
            //gbHistory.Text = "Truy xuất lịch sử";
            //gbHistory.Size = new Size(1280, 450);
            //gbHistory.Location = new Point(10, 590);
            //container.Controls.Add(gbHistory);

            //Label lblHistory = new Label()
            //{
            //    Text = "Lịch sử nhập kho",
            //    Location = new Point(15, 25),
            //    AutoSize = true,
            //    ForeColor = Color.Blue,
            //    Font = new Font("Segoe UI", 10, FontStyle.Bold)
            //};
            //btnPrintPNK = new Button()
            //{
            //    Text = "🖨 In phiếu nhập kho",
            //    Location = new Point(15, 55),
            //    Size = new Size(150, 35),
            //    BackColor = Color.FromArgb(33, 115, 70),
            //    ForeColor = Color.White,
            //    FlatStyle = FlatStyle.Flat
            //};

            //dgvImport = new DataGridView()
            //{
            //    Location = new Point(15, 100),
            //    Size = new Size(gbDetails.Width - 40, 320),
            //    BackgroundColor = Color.White,
            //    AutoGenerateColumns = true,
            //    ReadOnly = true,
            //    AllowUserToAddRows = false,
            //    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            //    BorderStyle = BorderStyle.None,
            //    RowHeadersVisible = false,
            //    Font = new Font("Segoe UI", 9),
            //    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            //    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
            //};
            //dgvImport.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            //dgvImport.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            //dgvImport.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            //dgvImport.EnableHeadersVisualStyles = false;
            //dgvImport.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);

            //// Xanh nhạt cho selection
            //dgvImport.DefaultCellStyle.SelectionBackColor = Color.FromArgb(204, 232, 255);
            //dgvImport.DefaultCellStyle.SelectionForeColor = Color.Black;

            //gbHistory.Controls.AddRange(new Control[] { lblHistory, btnPrintPNK, dgvImport });
            GroupBox gbHistory = new GroupBox();
            gbHistory.Text = "Truy xuất lịch sử";
            gbHistory.Size = new Size(1280, 450);
            gbHistory.Location = new Point(10, 590);
            container.Controls.Add(gbHistory);

            Label lblHistory = new Label()
            {
                Text = "Lịch sử nhập kho",
                Location = new Point(15, 25),
                AutoSize = true,
                ForeColor = Color.Blue,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };

            // 1. Nút In Phiếu Nhập Kho
            btnPrintPNK = new Button()
            {
                Text = "🖨 In phiếu nhập kho",
                Location = new Point(15, 55),
                Size = new Size(150, 35),
                BackColor = Color.FromArgb(33, 115, 70),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };

            // --- PHẦN UPDATE MỚI: 2 CẶP LABEL-COMBOBOX VÀ NÚT SEARCH ---

            // 2. Cặp 1: Dự án
            Label lblFilterProject = new Label()
            {
                Text = "Dự án:",
                Location = new Point(180, 62), // Căn chỉnh Y để text nằm giữa chiều cao ComboBox
                AutoSize = true,
                Font = new Font("Segoe UI", 9)
            };

            cboFilterProject = new ComboBox()
            {
                Name = "cboFilterProject",
                Location = new Point(230, 57),
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.ListItems
            };
            cboFilterProject.SelectedIndexChanged += CboFilterProject_SelectedIndexChanged; ; // Tải lại PO khi chọn dự án khác

            // 3. Cặp 2: Số PO
            Label lblFilterPO = new Label()
            {
                Text = "Số PO:",
                Location = new Point(450, 62),
                AutoSize = true,
                Font = new Font("Segoe UI", 9)
            };

            cboFilterPO = new ComboBox()
            {
                Name = "cboFilterPO",
                Location = new Point(500, 57),
                Width = 180,
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.ListItems
            };

            // 4. Nút Search (Nằm cuối hàng)
            btnSearchHistory = new Button()
            {
                Text = "🔍 Tìm kiếm",
                Location = new Point(700, 55),
                Size = new Size(120, 35),
                BackColor = Color.FromArgb(0, 120, 212), // Màu xanh dương đồng bộ với Header Grid
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnSearchHistory.Click += BtnSearchHistory_Click;

            // --- KẾT THÚC PHẦN UPDATE ---

            dgvImport = new DataGridView()
            {
                Location = new Point(15, 100),
                Size = new Size(gbHistory.Width - 40, 320), // Đã sửa gbDetails thành gbHistory để khớp code
                BackgroundColor = Color.White,
                AutoGenerateColumns = true,
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
            };

            // Định dạng Grid giữ nguyên như code cũ của bạn
            dgvImport.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvImport.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvImport.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvImport.EnableHeadersVisualStyles = false;
            dgvImport.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvImport.DefaultCellStyle.SelectionBackColor = Color.FromArgb(204, 232, 255);
            dgvImport.DefaultCellStyle.SelectionForeColor = Color.Black;

            // Thêm tất cả vào GroupBox
            gbHistory.Controls.AddRange(new Control[] {
                lblHistory,
                btnPrintPNK,
                lblFilterProject, cboFilterProject,
                lblFilterPO, cboFilterPO,
                btnSearchHistory,
                dgvImport
            });
        }

        private void CboPONo_Validating(object? sender, CancelEventArgs e)
        {
            Common.Common.AutoCompleteComboboxValidating(sender as ComboBox, e);
        }

        private void CboProject_Validating(object? sender, CancelEventArgs e)
        {
            Common.Common.AutoCompleteComboboxValidating(sender as ComboBox, e);
        }

        private void CboFilterProject_SelectedIndexChanged(object? sender, EventArgs e)
        {
            try
            {
                string project = (cboFilterProject != null && cboFilterProject.SelectedIndex > 0) ? cboFilterProject.SelectedItem.ToString() : "";
                LoadPOFilterByProject(project);
                //LoadImports(); // Không thực hiện lấy dữ liệu từ combobox trên nữa
            }
            catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void BtnSearchHistory_Click(object? sender, EventArgs e)
        {
            if (!Common.Common.IsComboBoxValid(cboFilterProject, "Dự án")
                || !Common.Common.IsComboBoxValid(cboFilterPO, "PO"))
                return;
            try
            {
                string poNo = (cboFilterPO != null && cboFilterPO.SelectedIndex > 0) ? cboFilterPO.SelectedItem.ToString() : "";
                string projectCode = (cboFilterProject != null && cboFilterProject.SelectedIndex > 0) ? cboFilterProject.SelectedItem.ToString() : "";
                LoadImports(poNo, projectCode);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Load dữ liệu thất bại: {ex.Message}", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void DgvImportQueue_KeyDown(object? sender, KeyEventArgs e)
        {
            //if (e.RowIndex < 0 || e.RowIndex >= _importQueue.Count) return;
            //string colName = dgvImportQueue.Columns[e.ColumnIndex].Name;
            //if (colName != "ID_Code") return;

            //if (e.Control && e.KeyCode == Keys.V)
            //{
            //    PasteToEditableCells();
            //    e.Handled = true;
            //}
        }

        public void SetupExportLayout(TabPage parent)
        {
            // --- CẤU HÌNH GỐC: CHO PHÉP SCROLL TOÀN TRANG ---
            Panel mainScrollPanel = new Panel();
            mainScrollPanel.Dock = DockStyle.Fill;
            mainScrollPanel.AutoScroll = true; // Kích hoạt cuộn ngang/dọc khi thu nhỏ
            parent.Controls.Add(mainScrollPanel);

            // Dùng một container để giữ độ rộng cố định khi scroll (tránh các control bị bóp méo)
            Panel container = new Panel();
            container.Width = 1300; // Độ rộng tối thiểu để không bị nhảy layout
            container.Height = 2000; // Độ cao ước tính cho 4 phần
            container.Location = new Point(0, 0);
            mainScrollPanel.Controls.Add(container);

            ucExportWarehouse ucExportWarehouse = new ucExportWarehouse();
            ucExportWarehouse.Dock = DockStyle.Fill;
            container.Controls.Add(ucExportWarehouse);
            ucExportWarehouse.BringToFront();
        }

        public void SetupFillInvoiceNoLayout_v2(TabPage parent)
        {
            // --- CẤU HÌNH GỐC: CHO PHÉP SCROLL TOÀN TRANG ---
            Panel mainScrollPanel = new Panel();
            mainScrollPanel.Dock = DockStyle.Fill;
            //mainScrollPanel.AutoScroll = true; // Kích hoạt cuộn ngang/dọc khi thu nhỏ
            parent.Controls.Add(mainScrollPanel);
            parent.AllowDrop = true;
            parent.Padding = new Padding(0);

            // Dùng một container để giữ độ rộng cố định khi scroll (tránh các control bị bóp méo)
            Panel container = new Panel();
            container.Width = 1300; // Độ rộng tối thiểu để không bị nhảy layout
            container.Height = 1200; // Độ cao ước tính cho 4 phần
            container.Location = new Point(0, 0);
            mainScrollPanel.Controls.Add(container);

            ucFillInvoiceNo ucFillInvoiceNo = new ucFillInvoiceNo();
            ucFillInvoiceNo.Dock = DockStyle.Fill;
            container.Controls.Add(ucFillInvoiceNo);
            ucFillInvoiceNo.BringToFront();
        }

        public void SetupSaveDeliveryNotetLayout(TabPage parent)
        {
            ucDelivery ucDelivery = new ucDelivery(_dtProject);
            ucDelivery.Dock = DockStyle.Fill;
            parent.Controls.Add(ucDelivery);
            ucDelivery.BringToFront();
        }

        public void SetupFillInvoiceNotLayout(TabPage parent)
        {
            parent.AllowDrop = true;
            parent.Padding = new Padding(0);

            // ── Header: hàng 1 = Dự án + PO No + btn Lưu | hàng 2 = INV Link ──
            var pHead = new Panel
            {
                Dock = DockStyle.Top,
                Height = 72,
                BackColor = Color.FromArgb(240, 240, 240),
            };
            parent.Controls.Add(pHead);

            // Hàng 1: Dự án
            pHead.Controls.Add(new Label { Text = "Dự án:", Location = new Point(8, 10), Size = new Size(50, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold) });
            var cboInvProject = new ComboBox
            {
                Location = new Point(60, 7),
                Size = new Size(160, 26),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            pHead.Controls.Add(cboInvProject);

            // Hàng 1: PO No — lọc theo dự án đang chọn
            pHead.Controls.Add(new Label { Text = "PO No:", Location = new Point(232, 10), Size = new Size(50, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold) });
            var cboPONoInv = new ComboBox
            {
                Location = new Point(284, 7),
                Size = new Size(200, 26),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            pHead.Controls.Add(cboPONoInv);

            // Hàng 1: nút Lưu hóa đơn (góc phải)
            var btnSaveInv = new Button
            {
                Text = "💾 Lưu hóa đơn",
                Size = new Size(145, 28),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Cursor = Cursors.Hand,
                Enabled = false
            };
            btnSaveInv.Location = new Point(pHead.Width - 155, 6);
            btnSaveInv.FlatAppearance.BorderSize = 0;
            pHead.Controls.Add(btnSaveInv);

            // Hàng 2: INV Link path
            var lblInvPath = new Label
            {
                Location = new Point(8, 40),
                Size = new Size(pHead.Width - 16, 18),
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.FromArgb(100, 100, 100),
                Text = "INV Link: —",
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            pHead.Controls.Add(lblInvPath);

            // Load PO theo dự án đang chọn vào cboPONoInv
            Action loadPOForProject = () =>
            {
                cboPONoInv.Items.Clear();
                cboPONoInv.Items.Add("-- Chọn PO No --");
                try
                {
                    string code = cboInvProject.SelectedItem?.ToString() ?? "";
                    var allPO = _poService.GetAll()
                        .Where(p => string.IsNullOrEmpty(code) ||
                                    p.ProjectCode == code ||
                                    p.Project_Name?.Contains(code) == true)
                        .OrderBy(p => p.PONo)
                        .ToList();
                    foreach (var po in allPO)
                        cboPONoInv.Items.Add(po.PONo);
                }
                catch { }
                cboPONoInv.SelectedIndex = 0;
                btnSaveInv.Enabled = false; // reset khi đổi dự án
            };

            // Khi chọn PO No → cập nhật trạng thái nút Lưu
            // (cboPONoInv.SelectedIndexChanged đăng ký sau khi khai báo _pendingDropPath)

            // Resize
            pHead.Resize += (s, e) =>
            {
                btnSaveInv.Location = new Point(pHead.Width - 155, 6);
                lblInvPath.Width = pHead.Width - 16;
            };

            // ── State — khai báo sớm để các lambda bên dưới dùng được ──
            string _invFolderPath = "";
            string _pendingDropPath = "";
            string _pendingDropName = "";

            // Đăng ký sau khi đã khai báo _pendingDropPath
            cboPONoInv.SelectedIndexChanged += (s, e) =>
            {
                string sel = cboPONoInv.SelectedItem?.ToString() ?? "";
                bool validPO = !string.IsNullOrEmpty(sel) && sel != "-- Chọn PO No --";
                btnSaveInv.Enabled = validPO && !string.IsNullOrEmpty(_pendingDropPath);
            };

            // ── Layout thủ công: pInvLeft (320px) | pInvRight (fill) ──
            const int LEFT_W = 320;

            var pInvRight = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(0)
            };
            parent.Controls.Add(pInvRight);

            var pInvLeft = new Panel
            {
                Dock = DockStyle.Left,
                Width = LEFT_W,
                BackColor = Color.White,
                Padding = new Padding(0)
            };
            parent.Controls.Add(pInvLeft);
            pInvLeft.BringToFront();

            // Alias để code bên dưới dùng splitMain.Panel1 / Panel2 vẫn đúng

            // ── Panel trái: INV list ──
            pInvLeft.Controls.Add(new Label
            {
                Text = "📄  INV List  —  kéo thả file PDF vào đây",
                Dock = DockStyle.Top,
                Height = 26,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(4, 0, 0, 0)
            });

            var dgvInv = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowDrop = true
            };
            dgvInv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvInv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvInv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvInv.EnableHeadersVisualStyles = false;
            dgvInv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvInv.Columns.Add(new DataGridViewTextBoxColumn { Name = "FileName", HeaderText = "Tên file", FillWeight = 70 });
            dgvInv.Columns.Add(new DataGridViewTextBoxColumn { Name = "FileSize", HeaderText = "Kích thước", FillWeight = 20 });
            dgvInv.Columns.Add(new DataGridViewTextBoxColumn { Name = "FullPath", HeaderText = "FullPath", Visible = false });
            dgvInv.Columns.Add(new DataGridViewTextBoxColumn { Name = "IsPending", HeaderText = "Trạng thái", FillWeight = 10 });
            pInvLeft.Controls.Add(dgvInv);

            // ── Panel phải: PDF Preview ──
            pInvRight.Controls.Add(new Label
            {
                Text = "🔍  Xem nhanh PDF",
                Dock = DockStyle.Top,
                Height = 26,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(102, 51, 153),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(4, 0, 0, 0)
            });

            var webView = new System.Windows.Forms.WebBrowser
            {
                Dock = DockStyle.Fill,
                ScrollBarsEnabled = true,
                IsWebBrowserContextMenuEnabled = false
            };
            pInvRight.Controls.Add(webView);

            // ── Load INV list từ thư mục ──
            Action loadInvList = () =>
            {
                dgvInv.Rows.Clear();
                _pendingDropPath = "";
                _pendingDropName = "";
                btnSaveInv.Enabled = false;

                if (string.IsNullOrEmpty(_invFolderPath) || !System.IO.Directory.Exists(_invFolderPath))
                {
                    lblInvPath.Text = "INV Link: (thư mục không tồn tại)";
                    lblInvPath.ForeColor = Color.FromArgb(200, 53, 69);
                    return;
                }
                lblInvPath.Text = $"INV Link: {_invFolderPath}";
                lblInvPath.ForeColor = Color.FromArgb(100, 100, 100);

                foreach (var f in System.IO.Directory.GetFiles(_invFolderPath, "*.pdf")
                                       .OrderBy(x => x))
                {
                    var fi = new System.IO.FileInfo(f);
                    int idx = dgvInv.Rows.Add();
                    dgvInv.Rows[idx].Cells["FileName"].Value = fi.Name;
                    dgvInv.Rows[idx].Cells["FileSize"].Value = $"{fi.Length / 1024.0:0.#} KB";
                    dgvInv.Rows[idx].Cells["FullPath"].Value = f;
                    dgvInv.Rows[idx].Cells["IsPending"].Value = "";
                }
            };

            // ── Chọn dự án → load INV path ──
            try
            {
                foreach (var p in _dtProject)    
                    cboInvProject.Items.Add(p.ProjectCode);
                if (cboInvProject.Items.Count > 0) cboInvProject.SelectedIndex = 0;
            }
            catch { }

            cboInvProject.SelectedIndexChanged += (s, e) =>
            {
                _invFolderPath = "";
                _pendingDropPath = "";
                webView.Navigate("about:blank");
                try
                {
                    string code = cboInvProject.SelectedItem?.ToString() ?? "";
                    var proj = _dtProject.Find(p => p.ProjectCode == code);
                    _invFolderPath = proj?.INV_Link?.Trim() ?? "";
                }
                catch { }
                loadPOForProject();
                loadInvList();
            };

            // ── Chọn dòng → preview PDF ──
            dgvInv.SelectionChanged += (s, e) =>
            {
                if (dgvInv.SelectedRows.Count == 0) return;
                string path = dgvInv.SelectedRows[0].Cells["FullPath"].Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(path) && System.IO.File.Exists(path))
                    webView.Navigate(path);
                else
                    webView.Navigate("about:blank");
            };

            // ── Double click → mở file PDF ──
            dgvInv.CellDoubleClick += (s, ev) =>
            {
                if (ev.RowIndex < 0) return;
                string path = dgvInv.Rows[ev.RowIndex].Cells["FullPath"].Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(path) && System.IO.File.Exists(path))
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    { FileName = path, UseShellExecute = true });
            };

            // ── Kéo thả file PDF — hỗ trợ cả File Explorer và Outlook ──

            // Hàm xử lý drop dùng chung cho dgvInv và pInvLeft
            Action<DragEventArgs> handleDrop = (e) =>
            {
                string pdfPath = null;
                string pdfName = null;
                byte[] pdfBytes = null;

                // Cách 1: Kéo từ File Explorer (DataFormats.FileDrop)
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                    var pdf = System.Array.Find(files, f =>
                        f.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase));
                    if (pdf != null) pdfPath = pdf;
                }

                // Cách 2: Kéo từ Outlook (FileGroupDescriptor + FileContents)
                if (pdfPath == null &&
                    e.Data.GetDataPresent("FileGroupDescriptorW") &&
                    e.Data.GetDataPresent("FileContents"))
                {
                    try
                    {
                        // Lấy tên file từ FileGroupDescriptorW
                        var fgd = e.Data.GetData("FileGroupDescriptorW") as System.IO.MemoryStream;
                        if (fgd != null)
                        {
                            fgd.Position = 0;
                            byte[] buf = fgd.ToArray();
                            // Tên file bắt đầu từ byte 76, encoding Unicode
                            string fname = System.Text.Encoding.Unicode
                                .GetString(buf, 76, buf.Length - 76)
                                .TrimEnd('\0').Trim();
                            if (fname.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                            {
                                pdfName = fname;
                                // Lấy nội dung file từ FileContents
                                var fc = e.Data.GetData("FileContents", true) as System.IO.MemoryStream;
                                if (fc != null)
                                {
                                    fc.Position = 0;
                                    pdfBytes = fc.ToArray();
                                }
                            }
                        }
                    }
                    catch { }
                }

                // Nếu không có gì hợp lệ
                if (pdfPath == null && pdfBytes == null)
                {
                    MessageBox.Show("Chỉ hỗ trợ file PDF!", "Thông báo",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
                }

                // Nếu từ Outlook → lưu tạm vào temp folder
                if (pdfPath == null && pdfBytes != null)
                {
                    string tmpDir = System.IO.Path.GetTempPath();
                    string tmpFile = System.IO.Path.Combine(tmpDir, pdfName ?? "invoice_temp.pdf");
                    System.IO.File.WriteAllBytes(tmpFile, pdfBytes);
                    pdfPath = tmpFile;
                }

                // Xóa dòng pending cũ
                for (int r = dgvInv.Rows.Count - 1; r >= 0; r--)
                    if (dgvInv.Rows[r].Cells["IsPending"].Value?.ToString() == "⏳ Chờ lưu")
                        dgvInv.Rows.RemoveAt(r);

                _pendingDropPath = pdfPath;
                _pendingDropName = System.IO.Path.GetFileName(pdfPath);
                var fi = new System.IO.FileInfo(pdfPath);

                int idx = dgvInv.Rows.Add();
                dgvInv.Rows[idx].Cells["FileName"].Value = _pendingDropName;
                dgvInv.Rows[idx].Cells["FileSize"].Value = $"{fi.Length / 1024.0:0.#} KB";
                dgvInv.Rows[idx].Cells["FullPath"].Value = pdfPath;
                dgvInv.Rows[idx].Cells["IsPending"].Value = "⏳ Chờ lưu";
                dgvInv.Rows[idx].DefaultCellStyle.ForeColor = Color.FromArgb(255, 140, 0);
                dgvInv.Rows[idx].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                dgvInv.ClearSelection();
                dgvInv.Rows[idx].Selected = true;
                webView.Navigate(pdfPath);
                btnSaveInv.Enabled = !string.IsNullOrEmpty(_invFolderPath)
                                  && cboPONoInv.SelectedItem?.ToString() != "-- Chọn PO No --"
                                  && cboPONoInv.SelectedIndex > 0;
            };

            Action<DragEventArgs> handleDragEnter = (e) =>
            {
                bool hasFileDrop = e.Data.GetDataPresent(DataFormats.FileDrop);
                bool hasOutlook = e.Data.GetDataPresent("FileGroupDescriptorW");
                e.Effect = (hasFileDrop || hasOutlook)
                    ? DragDropEffects.Copy
                    : DragDropEffects.None;
            };

            // Đăng ký trên dgvInv
            dgvInv.AllowDrop = true;
            dgvInv.DragEnter += (s, e) => handleDragEnter(e);
            dgvInv.DragOver += (s, e) => handleDragEnter(e);
            dgvInv.DragDrop += (s, e) => handleDrop(e);

            // Đăng ký trên pInvLeft (panel chứa) — bắt khi drop vào vùng trống
            pInvLeft.AllowDrop = true;
            pInvLeft.DragEnter += (s, e) => handleDragEnter(e);
            pInvLeft.DragOver += (s, e) => handleDragEnter(e);
            pInvLeft.DragDrop += (s, e) => handleDrop(e);

            // ── Lưu hóa đơn ──
            btnSaveInv.Click += (s, e) =>
            {
                if (string.IsNullOrEmpty(_pendingDropPath) || !System.IO.File.Exists(_pendingDropPath))
                { MessageBox.Show("Không có file nào chờ lưu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

                if (!System.IO.Directory.Exists(_invFolderPath))
                { MessageBox.Show("Thư mục INV Link không tồn tại!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                // Lấy PO No từ bộ lọc trong tab hóa đơn
                string poNo = cboPONoInv?.SelectedItem?.ToString()?.Trim() ?? "";
                if (string.IsNullOrEmpty(poNo) || poNo == "-- Chọn PO No --")
                { MessageBox.Show("Vui lòng chọn số PO trước khi lưu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

                // Tạo tên file: INV_PONo.pdf, nếu trùng → INV_PONo_2.pdf, _3.pdf...
                string baseName = $"INV_{poNo}";
                string destPath = System.IO.Path.Combine(_invFolderPath, baseName + ".pdf");
                int counter = 2;
                while (System.IO.File.Exists(destPath))
                {
                    destPath = System.IO.Path.Combine(_invFolderPath,
                        $"{baseName}_tờ{counter}.pdf");
                    counter++;
                }

                try
                {
                    System.IO.File.Copy(_pendingDropPath, destPath, false);
                    MessageBox.Show(
                        $"✅ Đã lưu hóa đơn thành công!\nFile: {System.IO.Path.GetFileName(destPath)}",
                        "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    _pendingDropPath = "";
                    _pendingDropName = "";
                    btnSaveInv.Enabled = false;
                    loadInvList();

                    // Preview file vừa lưu
                    webView.Navigate(destPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi lưu file: " + ex.Message, "Lỗi",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            // Load lần đầu
            if (cboInvProject.Items.Count > 0)
            {
                cboInvProject.SelectedIndex = 0;
                loadPOForProject();
            }
        }

        private void BtnDeleteRow_Click(object? sender, EventArgs e)
        {
            if (!PermissionHelper.Check("WAREHOUSE", "Lưu hóa đơn", "Xóa dòng nhập kho")) return;
            if (dgvImportQueue.SelectedRows.Count == 0) return;
            int idx = dgvImportQueue.SelectedRows[0].Index;
            if (idx >= 0 && idx < _importQueue.Count)
            {
                var key = $"{dgvImportQueue.SelectedRows[0].Cells[0].Value.ToString().Trim()}_{dgvImportQueue.SelectedRows[0].Cells[1].Value.ToString().Trim().ToLower()}";
                _importList.Remove(key);
                _importQueue.RemoveAt(idx);
                if (_importQueue.Count == 0) _currentBatchNo = "";
                RefreshQueueGrid();
            }
        }

        private void BtnCancel_Click(object? sender, EventArgs e)
        {
            if (_importQueue.Count > 0)
                if (MessageBox.Show($"Bạn có {_importQueue.Count} items chưa lưu. Tạo phiếu mới sẽ xóa danh sách. Tiếp tục?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            _importQueue.Clear(); _currentBatchNo = ""; _pendingPO_ID = 0;
            ClearImportItemForm(); RefreshQueueGrid();
        }

        private void ClearImportItemForm()
        {
            _importList.Clear();
        }

        private void DgvImportQueue_CellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _importQueue.Count) return;
            string colName = dgvImportQueue.Columns[e.ColumnIndex].Name;
            if (colName != "ID_Code") return;

            var item = _importQueue[e.RowIndex];
            frmCreateItemCode frmCreateItemCode = new frmCreateItemCode($"{item.Item_Name} - {item.Size} ");
            frmCreateItemCode.ShowDialog();
            if (string.IsNullOrEmpty(frmCreateItemCode.itemCode)) return;
            _useItemCodeExisted = frmCreateItemCode.isUseCodeAvailable;

            _importQueue[e.RowIndex].ID_Code = frmCreateItemCode.itemCode;
            if (!_useItemCodeExisted)
            {
                _importQueue[e.RowIndex].Material_Detail_ID = frmCreateItemCode.itemDetailId;
                _importQueue[e.RowIndex].Material_Detail_Number = frmCreateItemCode.itemDetailNumber;
            }

            dgvImportQueue.CurrentRow.Cells[colName].Value = frmCreateItemCode.itemCode;
            dgvImportQueue.CurrentRow.Cells["Material_Detail_Id"].Value = frmCreateItemCode.itemDetailId;
            dgvImportQueue.CurrentRow.Cells["Material_Detail_Number"].Value = frmCreateItemCode.itemDetailNumber;
        }

        private void BuildStockTab_V2(TabPage parent)
        {
            Panel mainScrollPanel = new Panel();
            mainScrollPanel.Dock = DockStyle.Fill;
            mainScrollPanel.AutoScroll = true;
            parent.Controls.Add(mainScrollPanel);

            Panel container = new Panel();
            container.Width = 1300;
            container.Height = 900;
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

            GroupBox gbHeader = new GroupBox();
            gbHeader.Text = "Lịch sử nhập hàng";
            gbHeader.Size = new Size(1280, 700);
            gbHeader.Location = new Point(10, 115);
            container.Controls.Add(gbHeader);

            dgvStock = new DataGridView
            {
                //Location = new Point(10, 115),
                Size = new Size(1200, 1200),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                Margin = new Padding(0, 100, 0, 0),
            };
            dgvStock.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvStock.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvStock.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvStock.EnableHeadersVisualStyles = false;
            dgvStock.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvStock.Dock = DockStyle.Fill;

            // Xanh nhạt cho selection
            dgvStock.DefaultCellStyle.SelectionBackColor = Color.FromArgb(204, 232, 255);
            dgvStock.DefaultCellStyle.SelectionForeColor = Color.Black;

            dgvStock.CellFormatting += DgvStock_CellFormatting;
            gbHeader.Controls.Add(dgvStock);
        }

        private void DgvStock_CellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string col = dgvStock.Columns[e.ColumnIndex].Name;
            if (col == "SL_Ton" || col == "KG_Ton" || col == "SL_Nhap")
            {
                decimal val = e.Value != null ? Convert.ToDecimal(e.Value) : 0;
                e.CellStyle.ForeColor = val > 0 ? Color.FromArgb(40, 167, 69) : Color.FromArgb(220, 53, 69);
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
        }

        private void DgvImportQueue_EditingControlShowing(object? sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress -= new KeyPressEventHandler(Column_KeyPress);
            if (dgvImportQueue.CurrentCell.ColumnIndex == dgvImportQueue.Columns["Qty_Import"].Index
                || dgvImportQueue.CurrentCell.ColumnIndex == dgvImportQueue.Columns["Weight_kg"].Index)
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress += new KeyPressEventHandler(Column_KeyPress);
                }
            }
        }

        private void Column_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void DgvImportQueue_CellEndEdit(object? sender, DataGridViewCellEventArgs e)
        {
            if (dgvImportQueue.Columns[e.ColumnIndex].Name == "Qty_Import")
            {
                var cell = dgvImportQueue.Rows[e.RowIndex].Cells[e.ColumnIndex];
                if (cell.Value != null)
                {
                    decimal newValue;
                    if (decimal.TryParse(cell.Value.ToString(), out newValue))
                    {
                        decimal originalLimit = Convert.ToDecimal(oldValue ?? 0);
                        if (newValue > originalLimit)
                        {
                            MessageBox.Show($"Số lượng không được vượt quá số lượng của PO !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            cell.Value = oldValue;
                        }
                    }
                    else
                    {
                        cell.Value = oldValue;
                    }
                }
            }

            if (dgvImportQueue.Columns[e.ColumnIndex].Name == "Weight_kg")
            {
                var cell = dgvImportQueue.Rows[e.RowIndex].Cells[e.ColumnIndex];
                if (cell.Value != null)
                {
                    decimal newValue;
                    if (decimal.TryParse(cell.Value.ToString(), out newValue))
                    {
                        decimal originalLimit = Convert.ToDecimal(oldValue ?? 0);
                        if (newValue > originalLimit + 10)
                        {
                            MessageBox.Show($"Khối lượng không được vượt quá khối lượng của PO !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            cell.Value = oldValue;
                        }
                    }
                    else
                    {
                        cell.Value = oldValue;
                    }
                }
            }
        }

        private void DgvImportQueue_CellBeginEdit(object? sender, DataGridViewCellCancelEventArgs e)
        {
            oldValue = dgvImportQueue.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
        }

        private Label AddStatLbl(Panel p, string title, string value, Color color, int x)
        {
            var card = new Panel { Location = new Point(x, 8), Size = new Size(220, 42), BackColor = color };
            p.Controls.Add(card);
            card.Controls.Add(new Label { Text = title, Font = new Font("Segoe UI", 8, FontStyle.Bold), ForeColor = Color.White, Location = new Point(6, 3), Size = new Size(208, 18) });
            var lbl = new Label { Text = value, Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.White, Location = new Point(6, 22), Size = new Size(208, 18) };
            card.Controls.Add(lbl);
            return lbl;
        }

        private Button CreateBtn(string text, Color color, Point loc, int w, int h)
        {
            var btn = new Button { Text = text, Location = loc, Size = new Size(w, h), BackColor = color, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Cursor = Cursors.Hand };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        private void LoadStock()
        {
            try
            {
                if (dgvStock == null) return;
                string kw = txtSearchStock?.Text.Trim() ?? "";
                string project = (cboProjectFilter != null && cboProjectFilter.SelectedIndex > 0) ? cboProjectFilter.SelectedItem.ToString() : "";
                if (!string.IsNullOrEmpty(project))
                {
                    BindStockGrid(_service.GetStock(project, kw));
                }
            }
            catch (Exception ex) { MessageBox.Show("Lỗi tải tồn kho: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void LoadStockOnly()
        {
            try { if (cboProject.Items.Count <= 0) return; if (dgvStock != null) BindStockGrid(_service.GetStockWithRemaining(cboProject.SelectedText ?? "")); }
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
            dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "Material_Detail_Id", HeaderText = "Material Detail Id", Width = 160, ReadOnly = true });
            dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "Material_Detail_Number", HeaderText = "Material Detail Number", Width = 160, ReadOnly = true });
        }

        public void TrackButtonClick()
        {
            btnSearch.Click += BtnSearch_Click;
            btnSave.Click += BtnSave_Click;
            btnPrintPNK.Click += BtnPrintPNK_Click;
        }

        private async void BtnPrintPNK_Click(object? sender, EventArgs e)
        {
            if (dgvImport.Rows.Count <= 0) return;
            int rsl = dgvImport.CurrentRow.Index;
            var billNo = dgvImport.Rows[rsl].Cells[1].Value.ToString();
            var poID = Convert.ToInt32(dgvImport.Rows[rsl].Cells[13].Value.ToString());

            var dtImports = await _service.GetImportRows(billNo, poID);
            PrintBill(dtImports, poID);
        }

        public void PrintBill(DataTable dtDetails, int poId)
        {
            try
            {
                if (dgvImport.CurrentRow == null)
                {
                    MessageBox.Show("Vui lòng chọn một phiếu nhập kho để in!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int rsl = dgvImport.CurrentRow.Index;
                var billNo = dgvImport.Rows[rsl].Cells[1].Value.ToString();

                var poModel = _poService.GetPOByPONo(poId);
                if (poModel == null) throw new Exception("Không tìm thấy thông tin PO tương ứng.");

                var supplier = new SupplierService().GetBySupId(poModel.Supplier_ID);
                var projects = new ProjectService().GetByProjectCode(poModel.ProjectCode ?? poModel.Notes);

                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "pnk_template_v2.xlsx");
                //string exportFolder = projects.PNK_Link;
                string exportFolder = @"D:\RAC\";
                if (!Directory.Exists(exportFolder)) Directory.CreateDirectory(exportFolder);

                string fileName = $"VMNP_PNK_{billNo}_{DateTime.Now:ddMMyyyy_HHmmss}.xlsx";
                string actualSavePath = Path.Combine(exportFolder, fileName);

                if (!File.Exists(templatePath))
                {
                    MessageBox.Show("Không tìm thấy file template tại: " + templatePath, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // --- BƯỚC FIX LỖI "Closed File": Copy trước khi mở ---
                File.Copy(templatePath, actualSavePath, true);

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(new FileInfo(actualSavePath)))
                {
                    ExcelWorksheet ws = package.Workbook.Worksheets[0];

                    // 1. Fill Header (A1:J13)
                    var headerRange = ws.Cells["A1:J14"];
                    foreach (var cell in headerRange)
                    {
                        if (cell.Value == null) continue;
                        string txt = cell.Value.ToString();
                        if (txt.Contains("<<BILL-NO>>")) cell.Value = txt.Replace("<<BILL-NO>>", "VMNP-" + billNo);
                        if (txt.Contains("<<DATE>>")) cell.Value = txt.Replace("<<DATE>>", poModel.Created_Date.HasValue ? poModel.Created_Date.Value.ToString("dd/MM/yyyy") : DateTime.Now.ToString("dd/MM/yyyy"));
                        if (txt.Contains("<<SUPPLIER_NAME>>")) cell.Value = txt.Replace("<<SUPPLIER_NAME>>", supplier?.Company_Name ?? "");
                    }

                    // 2. Xác định startRow (Tìm dòng chứa tiêu đề cột)
                    int startRow = 15;
                    for (int r = 1; r <= 25; r++)
                    {
                        if (ws.Cells[r, 1].Value?.ToString().Trim().ToUpper() == "NO.")
                        {
                            startRow = r + 1;
                            break;
                        }
                    }

                    int count = dtDetails.Rows.Count;
                    if (count > 1)
                    {
                        ws.InsertRow(startRow + 1, count - 1);
                        for (int i = 1; i < count; i++)
                        {
                            // Copy định dạng dòng gốc xuống các dòng mới
                            ws.Cells[startRow, 1, startRow, 9].Copy(ws.Cells[startRow + i, 1]);
                        }
                    }

                    // 3. Fill Data và Định dạng đồng nhất
                    for (int i = 0; i < count; i++)
                    {
                        DataRow dr = dtDetails.Rows[i];
                        int currentRow = startRow + i;
                        ws.Row(currentRow).Height = 25;

                        ws.Cells[currentRow, 1].Value = i + 1;
                        ws.Cells[currentRow, 2].Value = dr["ID_Code"];
                        ws.Cells[currentRow, 3].Value = dr["Item_Name"];
                        ws.Cells[currentRow, 4].Value = Convert.ToDecimal(dr["Qty_Import"] ?? 0);
                        ws.Cells[currentRow, 5].Value = dr["UNIT"];
                        ws.Cells[currentRow, 6].Value = dr["Weight_kg"];

                        // --- ÁP DỤNG ĐỊNH DẠNG ĐỒNG NHẤT CHO MỌI DÒNG ---
                        using (var range = ws.Cells[currentRow, 1, currentRow, 9])
                        {
                            range.Style.Font.Name = "Times New Roman";
                            range.Style.Font.Size = 16;
                            range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            range.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                            range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        }
                        // Cột Tên vật tư căn trái cho dễ nhìn
                        ws.Cells[currentRow, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    }

                    // 4. Cập nhật Footer Info (A18:J50 sau khi chèn dòng)
                    var footerRange = ws.Cells[startRow + count, 1, ws.Dimension.End.Row, ws.Dimension.End.Column];
                    foreach (var cell in footerRange)
                    {
                        if (cell.Value == null) continue;
                        string txt = cell.Value.ToString();
                        if (txt.Contains("<<PROJECT_NAME>>")) cell.Value = txt.Replace("<<PROJECT_NAME>>", projects?.ProjectName ?? "");
                        if (txt.Contains("<<PROJECT_CODE>>")) cell.Value = txt.Replace("<<PROJECT_CODE>>", projects?.ProjectCode ?? "");
                        if (txt.Contains("<<PROJECT_WO_NO>>")) cell.Value = txt.Replace("<<PROJECT_WO_NO>>", projects?.WorkorderNo ?? "");
                        if (txt.Contains("<<MPR_NO>>")) cell.Value = txt.Replace("<<MPR_NO>>", poModel.MPR_No ?? "");
                        if (txt.Contains("<<PO_NO>>")) cell.Value = txt.Replace("<<PO_NO>>", poModel.PONo ?? "");

                        // Xử lý hàm SUM() tại Excel
                        if (txt.Contains("<<SUM>>"))
                        {
                            cell.Value = ""; // Xóa text tag
                                             // Cột Qty là cột 4 (D). Hàm sum từ startRow đếncurrentRow cuối
                            cell.Formula = $"=SUM(D{startRow}:D{startRow + count - 1})";
                        }
                        if (txt.Contains("<<SUM-W>>"))
                        {
                            cell.Value = ""; // Xóa text tag
                                             // Cột Qty là cột 4 (D). Hàm sum từ startRow đếncurrentRow cuối
                            cell.Formula = $"=SUM(F{startRow}:F{startRow + count - 1})";
                        }
                    }

                    package.Save();
                }

                if (MessageBox.Show($"✅ Xuất phiếu PNK thành công!\nBạn có muốn mở file không?", "Thành công",
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

        //public void PrintBill(DataTable dtDetails, int poId)
        //{
        //    try
        //    {
        //        if (dgvImport.CurrentRow == null)
        //        {
        //            MessageBox.Show("Vui lòng chọn một phiếu nhập kho để in!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //            return;
        //        }

        //        int rsl = dgvImport.CurrentRow.Index;
        //        var billNo = dgvImport.Rows[rsl].Cells[1].Value.ToString();

        //        var poModel = _poService.GetPOByPONo(poId);
        //        if (poModel == null) throw new Exception("Không tìm thấy thông tin PO tương ứng.");

        //        var supplier = new SupplierService().GetBySupId(poModel.Supplier_ID);
        //        var projects = new ProjectService().GetByProjectCode(poModel.ProjectCode ?? poModel.Notes);

        //        string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "pnk_template.xlsx");
        //        string exportFolder = projects.PNK_Link;
        //        if (!Directory.Exists(exportFolder))
        //        {
        //            Directory.CreateDirectory(exportFolder);
        //        }

        //        string fileName = $"VMNP_PNK_{billNo}_{DateTime.Now:ddMMyyyy}.xlsx";
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
        //            var headerRange = ws.Cells["A1:J13"];
        //            foreach (var cell in headerRange)
        //            {
        //                if (cell.Value == null) continue;
        //                string txt = cell.Value.ToString();

        //                if (txt.Contains("<<BILL_NO>>")) cell.Value = txt.Replace("<<BILL_NO>>", "VMNP-" + billNo);
        //                if (txt.Contains("<<DATE>>")) cell.Value = txt.Replace("<<DATE>>", poModel.Created_Date.ToString());
        //                if (txt.Contains("<<SUPPLIER_NAME>>")) cell.Value = txt.Replace("<<SUPPLIER_NAME>>", supplier?.Company_Name ?? "");
        //            }

        //            var footerRange = ws.Cells["A18:J50"];
        //            foreach (var cell in footerRange)
        //            {
        //                if (cell.Value == null) continue;
        //                string txt = cell.Value.ToString();
        //                if (txt.Contains("<<PROJECT_NAME>>")) cell.Value = txt.Replace("<<PROJECT_NAME>>", projects?.ProjectName ?? "");
        //                if (txt.Contains("<<PROJECT_CODE>>")) cell.Value = txt.Replace("<<PROJECT_CODE>>", projects?.ProjectCode ?? "");
        //                if (txt.Contains("<<PROJECT_WO_NO>>")) cell.Value = txt.Replace("<<PROJECT_WO_NO>>", projects?.WorkorderNo ?? "");
        //                if (txt.Contains("<<MPR_NO>>")) cell.Value = txt.Replace("<<MPR_NO>>", poModel.MPR_No);
        //                if (txt.Contains("<<PO_NO>>")) cell.Value = txt.Replace("<<PO_NO>>", poModel.PONo);
        //            }

        //            int startRow = 15;
        //            for (int r = 1; r <= 25; r++)
        //            {
        //                if (ws.Cells[r, 1].Value?.ToString().Trim().ToUpper() == "NO.")
        //                {
        //                    startRow = r + 1;
        //                    break;
        //                }
        //            }

        //            int count = dtDetails.Rows.Count;
        //            if (count > 1)
        //            {
        //                ws.InsertRow(startRow + 1, count - 1);
        //                for (int i = 1; i < count; i++)
        //                {
        //                    ws.Cells[startRow, 1, startRow, 10].Copy(ws.Cells[startRow + i, 1]);
        //                }
        //            }

        //            decimal totalSum = 0;
        //            for (int i = 0; i < count; i++)
        //            {
        //                DataRow dr = dtDetails.Rows[i];
        //                int currentRow = startRow + i;
        //                ws.Row(currentRow).Height = 25;

        //                ws.Cells[currentRow, 1].Value = i + 1;
        //                ws.Cells[currentRow, 2].Value = dr["ID_Code"];
        //                ws.Cells[currentRow, 3].Value = dr["Item_Name"];
        //                ws.Cells[currentRow, 5].Value = dr["Qty_Import"];
        //                ws.Cells[currentRow, 6].Value = dr["UNIT"];
        //                ws.Cells[currentRow, 7].Value = dr["Weight_kg"];

        //                totalSum += Convert.ToDecimal(dr["Qty_Import"] ?? 0);
        //                    if (i > 0)
        //                    {
        //                        for (int col = 1; col <= 16; col++)
        //                        {
        //                            ws.Cells[currentRow, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        //                            ws.Cells[currentRow, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        //                            ws.Cells[currentRow, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        //                            ws.Cells[currentRow, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        //                            ws.Cells[currentRow, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        //                            ws.Cells[currentRow, col].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
        //                            ws.Cells[currentRow, col].Style.Font.Name = "Times New Roman";
        //                            ws.Cells[currentRow, col].Style.Font.Size = 9;
        //                        }
        //                    }
        //            }

        //            var sumCell = ws.Cells["A1:J60"].FirstOrDefault(c => c.Value?.ToString().Contains("<<SUM>>") == true);
        //            if (sumCell != null)
        //            {
        //                sumCell.Value = sumCell.Value.ToString().Replace("<<SUM>>", totalSum.ToString("N0"));
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

        private async void BtnSave_Click(object? sender, EventArgs e)
        {
            if (!PermissionHelper.Check("WAREHOUSE", "Lưu hóa đơn", "Lưu hóa đơn nhập kho")) return;
            if (!Common.Common.IsDataGridViewValid(dgvImportQueue, "Danh sách vật tư")) return;
            foreach (DataGridViewRow item in dgvImportQueue.Rows)
            {
                if (string.IsNullOrEmpty(item.Cells["ID_Code"].Value.ToString()))
                {
                    MessageBox.Show($"Hãy tạo code cho item: {item.Cells["Item_Name"].Value.ToString()}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            if (_importQueue.Count == 0) { MessageBox.Show("Danh sách phiếu đang trống!\nHãy thêm vật tư trước.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            try
            {
                int saved = 0;
                foreach (var imp in _importQueue)
                {
                    imp.Import_Date = DateTime.Now;
                    _service.InsertImport(imp, _currentUser);
                    saved++;
                }

                POHead h = new POHead()
                {
                    PONo = cboPONo.SelectedItem.ToString().Trim(),
                };
                _poService.MakeImported(h, DateTime.UtcNow.ToString());
                MessageBox.Show($"✅ Lưu phiếu nhập kho thành công!\nMã phiếu: {_currentBatchNo}\nSố vật tư: {saved} items", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                _importQueue.Clear(); _currentBatchNo = ""; _pendingPO_ID = 0;
                RefreshQueueGrid();
                LoadAll();
                _importList.Clear();
            }
            catch (Exception ex) { MessageBox.Show("Lỗi nhập kho: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void LoadAll()
        {
            LoadProjectImportFilter();
            LoadProjectFilter();
        }

        private void LoadProjectFilter()
        {
            try
            {
                cboProjectFilter.Items.Clear();
                cboProjectFilter.Items.Add("Tất cả dự án");
                foreach (var p in _dtProject) cboProjectFilter.Items.Add(p.ProjectCode);
                cboProjectFilter.SelectedIndex = 0;
            }
            catch { }
        }

        private void LoadProjectImportFilter()
        {
            try
            {
                cboProject.Items.Clear();
                //cboProject.Items.Add("Tất cả dự án");
                foreach (var p in _dtProject)
                    cboProject.Items.Add(p.ProjectCode);
                cboProject.SelectedIndex = 0;
            }
            catch { }
        }

        public void HandleComboBoxIndexChange()
        {
            cboProject.SelectedIndexChanged += CboProject_SelectedIndexChanged;
        }

        private void CboProject_SelectedIndexChanged(object? sender, EventArgs e)
        {
            try
            {
                string project = (cboProject != null && cboProject.SelectedIndex > 0) ? cboProject.SelectedItem.ToString() : "";
                LoadPOFilterByProject(project);
                //LoadImports(); // Không thực hiện lấy dữ liệu từ combobox trên nữa
            }
            catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void LoadImports(string poNo, string projectCode)
        {
            try
            {
                if (dgvImport == null) return;
                //string poNo = (cboPONo != null && cboPONo.SelectedIndex > 0) ? cboPONo.SelectedItem.ToString() : "";
                //string project = (cboProject != null && cboProject.SelectedIndex > 0) ? cboProject.SelectedItem.ToString() : "";

                var all = _service.GetAllImports();
                if (!string.IsNullOrEmpty(poNo))
                {
                    var po = _poService.GetAll().Find(p => p.PONo == poNo);
                    all = po != null ? all.FindAll(i => i.PO_ID == po.PO_ID) : new List<WarehouseImport>();
                }
                if (!string.IsNullOrEmpty(projectCode))
                    all = all.FindAll(i => i.Project_Code == projectCode);

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
                    Vi_Tri = i.Location,
                    PO_ID = i.PO_ID
                });
                if (dgvImport.Columns.Contains("ID")) dgvImport.Columns["ID"].Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải nhập kho: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadComboboxProject()
        {
            try
            {
                cboProject.Items.Clear();
                cboFilterProject.Items.Clear();
                //cboProject.Items.Add("Tất cả dự án");
                foreach (var p in _dtProject)
                {
                    cboProject.Items.Add(p.ProjectCode);
                    cboFilterProject.Items.Add(p.ProjectCode);
                }
                cboProject.SelectedIndex = 0;
                cboFilterProject.SelectedIndex = 0;
            }
            catch { }
        }

        private void LoadPOFilterByProject(string projectCode)
        {
            try
            {
                var allPO = _poService.GetAllPOForImport();
                if (string.IsNullOrEmpty(projectCode))
                {
                    cboPONo.Items.Clear();
                    cboFilterPO.Items.Clear();
                    cboPONo.Items.Add("-- Chọn PO --");
                    cboFilterPO.Items.Add("-- Chọn PO --");
                    foreach (var po in allPO)
                    {
                        cboPONo.Items.Add(po.PONo);
                        cboFilterPO.Items.Add(po.PONo);
                    }
                    cboPONo.SelectedIndex = 0;
                    cboFilterPO.SelectedIndex = 0;
                    return;
                }
                var projects = _dtProject;
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

                cboPONo.Items.Clear();
                cboFilterPO.Items.Clear();
                cboPONo.Items.Add("-- Chọn PO --");
                cboFilterPO.Items.Add("-- Chọn PO --");
                if (filtered.Count == 0)
                {
                    cboPONo.Items.Add("(Không có PO)");
                    cboFilterPO.Items.Add("(Không có PO)");
                    cboPONo.SelectedIndex = 0;
                    cboFilterPO.SelectedIndex = 0;
                    return;
                }
                foreach (var po in filtered)
                {
                    cboPONo.Items.Add(po.PONo);
                    cboFilterPO.Items.Add(po.PONo);
                }
                cboPONo.SelectedIndex = 0;
                cboFilterPO.SelectedIndex = 0;
            }
            catch (Exception ex) { MessageBox.Show("Lỗi lọc PO: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void BtnSearch_Click(object? sender, EventArgs e)
        {
            try
            {
                if (!Common.Common.IsComboBoxValid(cboProject, "Dự án")
                    || !Common.Common.IsComboBoxValid(cboPONo, "PO"))
                    return;


                string poNo = cboPONo.SelectedItem.ToString();
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
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "SL_NK", HeaderText = "SL nhập", Width = 75, ReadOnly = true });
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
                            if (cboProject != null && cboProject.SelectedIndex > 0)
                                projectCode = cboProject.SelectedItem.ToString();
                            else
                            {
                                try
                                {
                                    var pjs = _dtProject;
                                    projectCode = pjs.Find(p => p.WorkorderNo == po.WorkorderNo)?.ProjectCode ?? po.MPR_No ?? "";
                                }
                                catch { projectCode = po.MPR_No ?? ""; }
                            }

                            _importList.Add($"{row.Cells["STT"].Value.ToString()}_{detail.Item_Name.ToString().Trim().ToLower()}",
                                qty.ToString());

                            _importQueue.Add(new WarehouseImport
                            {
                                Import_No = _currentBatchNo,
                                Import_Date = DateTime.Now,
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

        private void RefreshQueueGrid()
        {
            dgvImportQueue.Rows.Clear();
            for (int i = 0; i < _importQueue.Count; i++)
            {
                var item = _importQueue[i];
                dgvImportQueue.Rows.Add(i + 1, item.Item_Name, item.Material, item.Size, item.UNIT, item.Qty_Import, item.Weight_kg, item.ID_Code, item.Import_No);
            }
        }

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

        private void PasteToEditableCells()
        {
            try
            {
                // 1. Lấy dữ liệu từ Clipboard
                string copiedData = Clipboard.GetText();
                if (string.IsNullOrEmpty(copiedData))
                {
                    MessageBox.Show("Bộ nhớ tạm (Clipboard) đang trống!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 2. Tách dữ liệu thành các dòng và các ô (tab-separated)
                string[] lines = copiedData.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);

                // 3. Xác định tọa độ bắt đầu
                int startRow = dgvImportQueue.CurrentCell?.RowIndex ?? 0;
                int startCol = dgvImportQueue.CurrentCell?.ColumnIndex ?? 0;

                DataTable dt = (DataTable)dgvImportQueue.DataSource;

                for (int i = 0; i < lines.Length; i++)
                {
                    int currentRow = startRow + i;

                    // Nếu dòng hiện tại vượt quá số dòng trong Grid, thêm dòng mới vào DataTable
                    if (currentRow >= dgvImportQueue.Rows.Count)
                    {
                        if (dt != null) dt.Rows.Add(dt.NewRow());
                        else dgvImportQueue.Rows.Add();
                    }

                    string[] cells = lines[i].Split('\t');
                    int currentGridCol = startCol;

                    for (int j = 0; j < cells.Length; j++)
                    {
                        // VÒNG LẶP TÌM CỘT ĐƯỢC PHÉP NHẬP (Skip ReadOnly/Hidden)
                        while (currentGridCol < dgvImportQueue.Columns.Count &&
                              (dgvImportQueue.Columns[currentGridCol].ReadOnly || !dgvImportQueue.Columns[currentGridCol].Visible))
                        {
                            currentGridCol++;
                        }

                        // Nếu vẫn nằm trong phạm vi cột của Grid thì mới dán
                        if (currentGridCol < dgvImportQueue.Columns.Count)
                        {
                            dgvImportQueue.Rows[currentRow].Cells[currentGridCol].Value = cells[j].Trim();
                            currentGridCol++; // Di chuyển sang cột tiếp theo cho ô Excel kế tiếp
                        }
                    }
                }

                MessageBox.Show(" ✅ Đã dán dữ liệu vào các ô cho phép!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(" ❌ Lỗi dán dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // =====================================================
        //  ÁP DỤNG PHÂN QUYỀN
        // =====================================================
        private void ApplyPermissions()
        {
            if (btnSave != null) PermissionHelper.Apply(btnSave, "WAREHOUSE", "Lưu hóa đơn");
            if (btnDeleteRow != null) PermissionHelper.Apply(btnDeleteRow, "WAREHOUSE", "Lưu hóa đơn");
        }

    }
}