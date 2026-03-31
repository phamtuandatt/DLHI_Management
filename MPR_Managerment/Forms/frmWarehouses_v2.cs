using Microsoft.Data.SqlClient;
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
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static OfficeOpenXml.ExcelErrorValue;

namespace MPR_Managerment.Forms
{
    public partial class frmWarehouses_v2 : Form
    {
        private TabControl mainTabControl;
        private TabPage pageImport, pageExport, pageWarehouse;

        private Button btnSearch, btnCancelSearch;

        private Button btnAdd, btnSave, btnCancel, btnDeleteRow;
        private ComboBox cboProject, cboPONo;

        private Button btnPrintPNK, btnDeleteFullBill;

        private DataGridView dgvImportQueue, dgvImport;

        private WarehouseService _service = new WarehouseService();
        private POService _poService = new POService();
        private WarehouseLocationService _warehouseService = new WarehouseLocationService();
        private string _currentUser = "Admin";

        private List<WarehouseImport> _imports = new List<WarehouseImport>();
        private List<WarehouseImport> _importQueue = new List<WarehouseImport>();
        private int _selectedImportID = 0;
        private int _pendingPO_ID = 0;
        private string _currentBatchNo = "";

        private Dictionary<string, string> _importList = new Dictionary<string, string>();
        private object oldValue = null;

        // ===== TỒN KHO =====
        private DataGridView dgvStock;
        private TextBox txtSearchStock;
        private ComboBox cboProjectFilter;
        private Label lblStockTotal, lblStockQty, lblStockWeight;
        private Panel panelStockSummary;

        private POService _poServices = new POService();

        public frmWarehouses_v2()
        {
            InitializeComponent();

            BuidUI();
            SetupImportLayout(pageImport);
            TrackButtonClick();
            LoadComboboxProject();
            HandleComboBoxIndexChange();
            BuildQueueColumns();
            BuildStockTab_V2(pageWarehouse);

            this.Load += FrmWarehouses_v2_Load;
        }

        private void FrmWarehouses_v2_Load(object? sender, EventArgs e)
        {
            LoadAll();
        }

        public void BuidUI()
        {
            // 1. Khởi tạo TabControl mặc định
            mainTabControl = new TabControl();
            mainTabControl.Dock = DockStyle.Fill;

            // Tùy chỉnh font chữ cho tiêu đề Tab cho chuyên nghiệp
            mainTabControl.Font = new Font("Segoe UI", 10, FontStyle.Regular);

            // 2. Tạo TabPage: Nhập kho (Import)
            pageImport = new TabPage();
            pageImport.Text = "  📥  Nhập kho  "; // Thêm khoảng trắng để tab rộng hơn
            pageImport.BackColor = Color.White;      // Đặt nền trắng cho sạch sẽ
                                                     // pageImport.Controls.Add(yourControl); 

            // 3. Tạo TabPage: Xuất kho (Export)
            pageExport = new TabPage();
            pageExport.Text = "  📤  Xuất kho  ";
            pageExport.BackColor = Color.White;

            // 4. Tạo TabPage: Kho bãi (Warehouse)
            pageWarehouse = new TabPage();
            pageWarehouse.Text = "  📦  Tồn kho  ";
            pageWarehouse.BackColor = Color.White;

            // 5. Thêm các Page vào TabControl
            mainTabControl.TabPages.Add(pageImport);
            mainTabControl.TabPages.Add(pageExport);
            mainTabControl.TabPages.Add(pageWarehouse);

            // 6. Đưa TabControl vào Panel chính

            this.Controls.Add(mainTabControl);
        }

        public void SetupImportLayout(TabPage parent)
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

            // =========================================================================
            // PHẦN 1: HEADER (Tìm kiếm)
            // =========================================================================
            GroupBox gbHeader = new GroupBox();
            gbHeader.Text = "Bộ lọc tìm kiếm";
            gbHeader.Size = new Size(1280, 70);
            gbHeader.Location = new Point(10, 10);
            container.Controls.Add(gbHeader);

            // Label và ComboBox cho Dự án
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
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            // Label và ComboBox cho PO NO
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
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            // Nút Tìm kiếm
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

            // Nút Tìm kiếm
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

            // Thêm tất cả vào GroupBox Header
            gbHeader.Controls.AddRange(new Control[] { lblProject, cboProject, lblPONo, cboPONo, btnSearch, btnCancelSearch });

            // =========================================================================
            // PHẦN 2: THAO TÁC PHIẾU
            // =========================================================================
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
            gbActions.Controls.AddRange(new Control[] { btnAdd, btnSave, btnCancel });

            // =========================================================================
            // PHẦN 3: CHI TIẾT VẬT TƯ
            // =========================================================================
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
                Location = new Point(gbDetails.Width - 125, 20),
                Size = new Size(120, 30),
                BackColor = Color.FromArgb(220, 53, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
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
            dgvImportQueue.CellBeginEdit += DgvImportQueue_CellBeginEdit;
            dgvImportQueue.CellEndEdit += DgvImportQueue_CellEndEdit;
            dgvImportQueue.EditingControlShowing += DgvImportQueue_EditingControlShowing;
            dgvImportQueue.CellDoubleClick += DgvImportQueue_CellDoubleClick;

            gbDetails.Controls.AddRange(new Control[] { lblDetail, btnDeleteRow, dgvImportQueue });

            // =========================================================================
            // PHẦN 4: LỊCH SỬ NHẬP KHO
            // =========================================================================
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
            btnPrintPNK = new Button()
            {
                Text = "🖨 In phiếu nhập kho",
                Location = new Point(15, 55),
                Size = new Size(150, 35),
                BackColor = Color.FromArgb(33, 115, 70),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            //btnDeleteFullBill = new Button()
            //{
            //    Text = "🗑 Xóa cả phiếu",
            //    Location = new Point(175, 55),
            //    Size = new Size(150, 35),
            //    BackColor = Color.FromArgb(180, 30, 30),
            //    ForeColor = Color.White,
            //    FlatStyle = FlatStyle.Flat
            //};
            dgvImport = new DataGridView()
            {
                Location = new Point(15, 100),
                Size = new Size(gbDetails.Width - 40, 320),
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
            dgvImport.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvImport.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvImport.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvImport.EnableHeadersVisualStyles = false;
            dgvImport.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);

            gbHistory.Controls.AddRange(new Control[] { lblHistory, btnPrintPNK,/* btnDeleteFullBill,*/ dgvImport });
        }

        private void DgvImportQueue_CellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _importQueue.Count) return;
            string colName = dgvImportQueue.Columns[e.ColumnIndex].Name;
            if (colName != "ID_Code") return;

            var item = _importQueue[e.RowIndex];
            frmCreateItemCode frmCreateItemCode = new frmCreateItemCode();
            frmCreateItemCode.ShowDialog();

            if (string.IsNullOrEmpty(frmCreateItemCode.itemCode)) return;

            _importQueue[e.RowIndex].ID_Code = frmCreateItemCode.itemCode;
            dgvImportQueue.CurrentRow.Cells[colName].Value = frmCreateItemCode.itemCode;
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
            container.Height = 900; // Độ cao ước tính cho 4 phần
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
                Size = new Size(1200, 1200),
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
            dgvStock.CellFormatting += DgvStock_CellFormatting; ;
            container.Controls.Add(dgvStock);
        }

        private void DgvStock_CellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
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

        private void DgvImportQueue_EditingControlShowing(object? sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress -= new KeyPressEventHandler(Column_KeyPress);
            if (dgvImportQueue.CurrentCell.ColumnIndex == dgvImportQueue.Columns["Qty_Import"].Index // Cột số lượng
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
            // Chỉ cho phép nhập số và dấu chấm thập phân
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void DgvImportQueue_CellEndEdit(object? sender, DataGridViewCellEventArgs e)
        {
            // Giả sử cột bạn muốn kiểm tra tên là "Qty" hoặc "Quantity"
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
                            MessageBox.Show($"Số lượng không được vượt quá số lượng của PO !",
                                            "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                            MessageBox.Show($"Khối lượng không được vượt quá khối lượng của PO !",
                                            "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                BindStockGrid(_service.GetStock(project, kw));
            }
            catch (Exception ex) { MessageBox.Show("Lỗi tải tồn kho: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void LoadStockOnly()
        {
            try { if (dgvStock != null) BindStockGrid(_service.GetStockWithRemaining()); }
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
                // 1. Lấy thông tin từ Grid đang chọn
                if (dgvImport.CurrentRow == null)
                {
                    MessageBox.Show("Vui lòng chọn một phiếu nhập kho để in!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int rsl = dgvImport.CurrentRow.Index;
                var billNo = dgvImport.Rows[rsl].Cells[1].Value.ToString();
                // Giả sử logic lấy PO No của bạn là xóa tiền tố "PNK-"
                //var poNO = billNo.Replace("PNK-", "");

                // Gọi Service lấy dữ liệu liên quan
                var poModel = _poService.GetPOByPONo(poId);
                if (poModel == null) throw new Exception("Không tìm thấy thông tin PO tương ứng.");

                var supplier = new SupplierService().GetBySupId(poModel.Supplier_ID);
                var projects = new ProjectService().GetByProjectCode(poModel.Notes);

                // 2. Thiết lập đường dẫn Template
                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "pnk_template.xlsx");

                if (!File.Exists(templatePath))
                {
                    MessageBox.Show($"Lỗi: Không tìm thấy file template!\nĐường dẫn dự kiến: {templatePath}",
                                    "Lỗi hệ thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 3. Hiển thị hộp thoại Lưu file
                var saveDialog = new SaveFileDialog
                {
                    Title = "Lưu phiếu nhập kho",
                    Filter = "Excel Files|*.xlsx",
                    FileName = $"PNK_{billNo}_{DateTime.Now:ddMMyyyy}",
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                };

                if (saveDialog.ShowDialog() != DialogResult.OK) return;

                // ĐƯỜNG DẪN LƯU THỰC TẾ
                string actualSavePath = saveDialog.FileName;

                // 4. Tiến hành xuất Excel bằng EPPlus
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                FileInfo templateFile = new FileInfo(templatePath);
                FileInfo newFile = new FileInfo(actualSavePath);

                // Nếu file đã tồn tại và đang mở, cần xử lý lỗi ghi đè
                using (ExcelPackage package = new ExcelPackage(newFile, templateFile))
                {
                    ExcelWorksheet ws = package.Workbook.Worksheets[0];

                    // --- ĐIỀN THÔNG TIN HEADER ---
                    var headerRange = ws.Cells["A1:J13"];
                    foreach (var cell in headerRange)
                    {
                        if (cell.Value == null) continue;
                        string txt = cell.Value.ToString();

                        if (txt.Contains("<<BILL_NO>>")) cell.Value = txt.Replace("<<BILL_NO>>", billNo);
                        if (txt.Contains("<<DATE>>")) cell.Value = txt.Replace("<<DATE>>", poModel.Created_Date.ToString());
                        if (txt.Contains("<<SUPPLIER_NAME>>")) cell.Value = txt.Replace("<<SUPPLIER_NAME>>", supplier?.Company_Name ?? "");
                    }

                    // --- ĐIỀN THÔNG TIN FOOTER ---
                    var footerRange = ws.Cells["A18:J50"];
                    foreach (var cell in footerRange)
                    {
                        if (cell.Value == null) continue;
                        string txt = cell.Value.ToString();
                        if (txt.Contains("<<PROJECT_NAME>>")) cell.Value = txt.Replace("<<PROJECT_NAME>>", projects?.ProjectName ?? "");
                        if (txt.Contains("<<PROJECT_CODE>>")) cell.Value = txt.Replace("<<PROJECT_CODE>>", projects?.ProjectCode ?? "");
                        if (txt.Contains("<<PROJECT_WO_NO>>")) cell.Value = txt.Replace("<<PROJECT_WO_NO>>", projects?.WorkorderNo ?? "");
                        if (txt.Contains("<<MPR_NO>>")) cell.Value = txt.Replace("<<MPR_NO>>", poModel.MPR_No); // Sửa lại Tag nếu viết nhầm
                        if (txt.Contains("<<PO_NO>>")) cell.Value = txt.Replace("<<PO_NO>>", poModel.PONo);
                    }

                    // --- ĐIỀN CHI TIẾT VẬT TƯ ---
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
                        // Chèn thêm dòng và giữ định dạng
                        ws.InsertRow(startRow + 1, count - 1);
                        for (int i = 1; i < count; i++)
                        {
                            ws.Cells[startRow, 1, startRow, 10].Copy(ws.Cells[startRow + i, 1]);
                        }
                    }

                    decimal totalSum = 0;
                    for (int i = 0; i < count; i++)
                    {
                        DataRow dr = dtDetails.Rows[i];
                        int currentRow = startRow + i;

                        ws.Cells[currentRow, 1].Value = i + 1;
                        ws.Cells[currentRow, 2].Value = "NOT";
                        //ws.Cells[currentRow, 2].Value = dr["Material_Code"];
                        ws.Cells[currentRow, 3].Value = dr["Item_Name"];
                        ws.Cells[currentRow, 5].Value = dr["Qty_Import"];
                        ws.Cells[currentRow, 6].Value = dr["UNIT"];
                        ws.Cells[currentRow, 7].Value = dr["Weight_kg"];
                        //ws.Cells[currentRow, 8].Value = dr["Price"];
                        //ws.Cells[currentRow, 9].Value = dr["Amount"];
                        //ws.Cells[currentRow, 10].Value = dr["Remarks"];

                        totalSum += Convert.ToDecimal(dr["Qty_Import"] ?? 0);
                    }

                    // --- ĐIỀN TỔNG CỘNG ---
                    var sumCell = ws.Cells["A1:J60"].FirstOrDefault(c => c.Value?.ToString().Contains("<<SUM>>") == true);
                    if (sumCell != null)
                    {
                        sumCell.Value = sumCell.Value.ToString().Replace("<<SUM>>", totalSum.ToString("N0"));
                    }

                    // Lưu file thực tế
                    package.Save();
                }

                // 5. Thông báo và mở file
                var result = MessageBox.Show(
                    $"✅ Xuất phiếu nhập kho thành công!\nFile: {actualSavePath}\n\nBạn có muốn mở file ngay không?",
                    "Thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                if (result == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = actualSavePath,
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi in phiếu: " + ex.Message, "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //public void PrintBill(DataTable dtDetails)
        //{
        //    //Lấy thông tin PO
        //    DataRowView row = (DataRowView)dgvImport.CurrentRow.DataBoundItem;
        //    var billNo = row["Import_No"].ToString();
        //    var poNO = billNo.Replace("PNK-", "");
        //    var poModel = _poService.GetPOByPONo(poNO);

        //    // Lấy thông tin supplier
        //    var supplier = new SupplierService().GetBySupId(poModel.Supplier_ID);


        //    // Lấy thông tin project
        //    var projects = new ProjectService().GetByProjectCode(poModel.Notes);

        //    // Đường dẫn template
        //    string templatePath = Path.Combine(Application.StartupPath, "Templates", "pnk_template.xlsx");
        //    if (!File.Exists(templatePath))
        //    {
        //        MessageBox.Show("Không tìm thấy file template!\nVui lòng đặt file template.xlsx vào thư mục Templates.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        return;
        //    }

        //    // Chọn nơi lưu file
        //    var saveDialog = new SaveFileDialog
        //    {
        //        Title = "Lưu file PO",
        //        Filter = "Excel Files|*.xlsx",
        //        FileName = $"{billNo}_{DateTime.Now:ddMMyyyy}",
        //        InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        //    };
        //    if (saveDialog.ShowDialog() != DialogResult.OK) return;

        //    // Thiết lập bản quyền EPPlus
        //    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

        //    FileInfo templateFile = new FileInfo(templatePath);
        //    FileInfo newFile = new FileInfo(savePath);

        //    using (ExcelPackage package = new ExcelPackage(newFile, templateFile))
        //    {
        //        // Lấy Worksheet đầu tiên
        //        ExcelWorksheet ws = package.Workbook.Worksheets[0];

        //        // 1. ĐIỀN THÔNG TIN HEADER (Dựa trên các tag << >> trong template)
        //        // Tìm và thay thế các biến trong vùng Header (Từ dòng 1 đến dòng 15)
        //        var headerRange = ws.Cells["A1:J15"];
        //        foreach (var cell in headerRange)
        //        {
        //            if (cell.Value == null) continue;
        //            string txt = cell.Value.ToString();

        //            if (txt.Contains("<<BILL_NO>>")) cell.Value = txt.Replace("<<BILL_NO>>", billNo);
        //            if (txt.Contains("<<DATE>>")) cell.Value = txt.Replace("<<DATE>>", poModel.Created_Date.ToString());
        //            if (txt.Contains("<<SUPPLIER_NAME>>")) cell.Value = txt.Replace("<<SUPPLIER_NAME>>", supplier.Company_Name);
        //        }

        //        // 2. ĐIỀN THÔNG TIN DỰ ÁN Ở PHẦN CUỐI (SHIPMENT DETAILS)
        //        // Thông thường tag <<PROJECT_NAME>> nằm ở phía dưới bảng vật tư
        //        var footerRange = ws.Cells["A20:J40"]; // Quét vùng rộng phía dưới
        //        foreach (var cell in footerRange)
        //        {
        //            if (cell.Value == null) continue;
        //            string txt = cell.Value.ToString();
        //            if (txt.Contains("<<PROJECT_NAME>>")) cell.Value = txt.Replace("<<PROJECT_NAME>>", projects.ProjectName);
        //            if (txt.Contains("<<PROJECT_CODE>>")) cell.Value = txt.Replace("<<PROJECT_CODE>>", projects.ProjectCode);
        //            if (txt.Contains("<<PROJECT_WO_NO>>")) cell.Value = txt.Replace("<<PROJECT_WO_NO>>", projects.WorkorderNo);
        //            if (txt.Contains("<<MPR_NO>>")) cell.Value = txt.Replace("<<MPR_NO>>", poModel.MPR_No);
        //            if (txt.Contains("<<PO_NO>>")) cell.Value = txt.Replace("<<PO_NO>>", poModel.PONo);
        //        }


        //        // 3. ĐIỀN DANH SÁCH VẬT TƯ (BẮT ĐẦU TỪ DÒNG CÓ TIÊU ĐỀ "No.")
        //        int startRow = 15; // Theo template, bảng bắt đầu khoảng dòng 15-16
        //        for (int r = 1; r <= 20; r++)
        //        {
        //            if (ws.Cells[r, 1].Value?.ToString().Trim().ToUpper() == "NO.")
        //            {
        //                startRow = r + 1;
        //                break;
        //            }
        //        }

        //        int count = dtDetails.Rows.Count;
        //        if (count > 1)
        //        {
        //            // Chèn thêm dòng nếu cần (giữ lại định dạng của dòng startRow)
        //            ws.InsertRow(startRow + 1, count - 1);
        //            for (int i = 1; i < count; i++)
        //            {
        //                ws.Cells[startRow, 1, startRow, 10].Copy(ws.Cells[startRow + i, 1]);
        //            }
        //        }

        //        decimal totalSum = 0;

        //        // Đổ dữ liệu từ DataTable/Grid vào bảng
        //        for (int i = 0; i < count; i++)
        //        {
        //            DataRow dr = dtDetails.Rows[i];
        //            int currentRow = startRow + i;

        //            ws.Cells[currentRow, 1].Value = i + 1; // NO.
        //            ws.Cells[currentRow, 2].Value = dr["Material_Code"];
        //            ws.Cells[currentRow, 3].Value = dr["Material_Size"];
        //            // Cột 4 thường để trống hoặc merge theo template
        //            ws.Cells[currentRow, 5].Value = dr["Qty"];
        //            ws.Cells[currentRow, 6].Value = dr["Unit"];
        //            ws.Cells[currentRow, 7].Value = dr["Weight"];
        //            ws.Cells[currentRow, 8].Value = dr["Price"];
        //            ws.Cells[currentRow, 9].Value = dr["Amount"];
        //            ws.Cells[currentRow, 10].Value = dr["Remarks"];

        //            totalSum += Convert.ToDecimal(dr["Qty"] ?? 0);
        //        }

        //        // 4. ĐIỀN TỔNG CỘNG (Tag <<SUM>>)
        //        // Tìm ô chứa <<SUM>> để điền tổng số lượng
        //        var sumCell = ws.Cells["A1:J50"].FirstOrDefault(c => c.Value?.ToString().Contains("<<SUM>>") == true);
        //        if (sumCell != null)
        //        {
        //            sumCell.Value = sumCell.Value.ToString().Replace("<<SUM>>", totalSum.ToString("N0"));
        //        }

        //        // Lưu file
        //        if (!Directory.Exists(Path.GetDirectoryName(savePath)))
        //        {
        //            Directory.CreateDirectory(Path.GetDirectoryName(savePath));
        //        }
        //        package.Save();
        //    }

        //    // Tự động mở file sau khi xuất
        //    //Process.Start(new ProcessStartInfo(savePath) { UseShellExecute = true });
        //    // Mở file sau khi xuất
        //    var result = MessageBox.Show(
        //        $"✅ Xuất Excel thành công!\nFile: {saveDialog.FileName}\n\nBạn có muốn mở file không?",
        //        "Thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
        //    if (result == DialogResult.Yes)
        //        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
        //        {
        //            FileName = saveDialog.FileName,
        //            UseShellExecute = true
        //        });
        //}

        private void BtnSave_Click(object? sender, EventArgs e)
        {
            frmCreateItemCode frmCreateItemCode = new frmCreateItemCode();
            frmCreateItemCode.ShowDialog();
            //if (_importQueue.Count == 0) { MessageBox.Show("Danh sách phiếu đang trống!\nHãy thêm vật tư trước.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            //try
            //{
            //    int saved = 0;
            //    foreach (var imp in _importQueue) 
            //    { 
            //        imp.Import_Date = DateTime.Now; 
            //        _service.InsertImport(imp, _currentUser);
            //        saved++; 
            //    }
            //    MessageBox.Show($"✅ Lưu phiếu nhập kho thành công!\nMã phiếu: {_currentBatchNo}\nSố vật tư: {saved} items", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    _importQueue.Clear(); _currentBatchNo = ""; _pendingPO_ID = 0;
            //    RefreshQueueGrid(); 
            //    LoadAll();
            //}
            //catch (Exception ex) { MessageBox.Show("Lỗi nhập kho: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }

        }



        private void LoadAll()
        {
            LoadProjectImportFilter();
            //LoadProjectExportFilter();
            LoadProjectFilter();
            //LoadPOFilterByProject("");
            //LoadPOExportFilterByProject("");
            //LoadWarehouseExportCombo();
            //LoadImports();
            //LoadExports();
            //LoadStock();
            //LoadStockForExport();

            // Đảm bảo dgvExport hiển thị đúng
        }

        private void LoadProjectFilter()
        {
            try { cboProjectFilter.Items.Clear(); cboProjectFilter.Items.Add("Tất cả dự án"); foreach (var p in new ProjectService().GetAll()) cboProjectFilter.Items.Add(p.ProjectCode); cboProjectFilter.SelectedIndex = 0; }
            catch { }
        }

        private void LoadProjectImportFilter()
        {
            try {
                cboProject.Items.Clear(); 
                cboProject.Items.Add("Tất cả dự án"); 
                foreach (var p in new ProjectService().GetAll()) 
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
                LoadImports();
            }
            catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void LoadImports()
        {
            try
            {
                if (dgvImport == null) return;
                string poNo = (cboPONo != null && cboPONo.SelectedIndex > 0) ? cboPONo.SelectedItem.ToString() : "";
                string project = (cboProject != null && cboProject.SelectedIndex > 0) ? cboProject.SelectedItem.ToString() : "";

                var all = _service.GetAllImports();

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
                    Vi_Tri = i.Location,

                    PO_ID = i.PO_ID
                });
                if (dgvImport.Columns.Contains("ID")) dgvImport.Columns["ID"].Visible = false;
                //if (lblImportStatus != null) lblImportStatus.Text = $"Tổng: {_imports.Count} bản ghi";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải nhập kho: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadComboboxProject()
        {
            try { 
                cboProject.Items.Clear();
                cboProject.Items.Add("Tất cả dự án"); 
                foreach (var p in new ProjectService().GetAll())
                {
                    cboProject.Items.Add(p.ProjectCode);
                }
                cboProject.SelectedIndex = 0; }
            catch { }
        }

        private void LoadPOFilterByProject(string projectCode)
        {
            try
            {
                var allPO = _poService.GetAll();
                if (string.IsNullOrEmpty(projectCode))
                {
                    cboPONo.Items.Clear();
                    cboPONo.Items.Add("-- Chọn PO --");
                    foreach (var po in allPO) cboPONo.Items.Add(po.PONo);
                    cboPONo.SelectedIndex = 0;
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

                cboPONo.Items.Clear();
                cboPONo.Items.Add("-- Chọn PO --");
                if (filtered.Count == 0) 
                { 
                    cboPONo.Items.Add("(Không có PO)");
                    cboPONo.SelectedIndex = 0; 
                    return; 
                }
                foreach (var po in filtered)
                {
                    cboPONo.Items.Add(po.PONo);
                } 
                cboPONo.SelectedIndex = 0;
            }
            catch (Exception ex) { MessageBox.Show("Lỗi lọc PO: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void BtnSearch_Click(object? sender, EventArgs e)
        {
            try
            {
                if (cboProject.Items.Count <= 0 || cboPONo.Items.Count <= 0 ||
                    cboProject.SelectedIndex == -1 || cboPONo.SelectedIndex == -1) return;
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
                                try { var pjs = new ProjectService().GetAll(); projectCode = pjs.Find(p => p.WorkorderNo == po.WorkorderNo)?.ProjectCode ?? po.MPR_No ?? ""; }
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
                                //Location = txtLocation.Text.Trim()
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
            //int count = _importQueue.Count;
            //if (lblQueueStatus != null) lblQueueStatus.Text = count > 0 ? $"📋 Phiếu: {_currentBatchNo}  |  {count} vật tư" : "";
            //if (lblCurrentBatch != null) lblCurrentBatch.Text = string.IsNullOrEmpty(_currentBatchNo)
            //    ? "Mã phiếu: (chưa có — chọn PO hoặc thêm thủ công)"
            //    : $"✅ Mã phiếu: {_currentBatchNo}  ({count} items)";
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

    }
}
