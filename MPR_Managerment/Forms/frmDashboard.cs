using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace MPR_Managerment.Forms
{
    public partial class frmDashboard : Form
    {
        private TabControl tabMain;
        private TabPage tabPO, tabMPR, tabRIR;

        // PO Tab
        private DataGridView dgvPO;
        private DataGridView dgvPOImports;
        private Label lblPOTotal, lblPOOverdue, lblPOCompleted, lblPOInProgress;
        private Panel panelPOSummary;
        private ComboBox cboFilterPO;
        private TextBox txtSearchPO;

        // MPR Tab
        private DataGridView dgvMPR;
        private DataGridView dgvMPRPO;      // Bảng PO của MPR đang chọn
        private Label lblMPRPOTitle; // Tiêu đề bảng PO
        private Label lblMPRTotal, lblMPRHasPO, lblMPRNoPO, lblMPRCompleted;
        private Panel panelMPRSummary;
        private ComboBox cboFilterMPR;
        private ComboBox cboFilterPOStatus;  // Lọc theo Tình trạng PO
        private TextBox txtSearchMPR;
        private Button btnExportMPR;         // Xuất Excel danh sách MPR

        // RIR Tab
        private DataGridView dgvRIR;
        private Label lblRIRTotal, lblRIRPending, lblRIRInspecting, lblRIRDone;
        private Panel panelRIRSummary;
        private ComboBox cboFilterRIR;
        private TextBox txtSearchRIR;
        private DataGridView dgvRIRDetail;

        public frmDashboard()
        {
            InitializeComponent();
            BuildUI();
            LoadData();

            // Ép form gọi sự kiện Resize lần đầu để chia tỷ lệ ngay khi mở
            this.OnResize(EventArgs.Empty);
        }

        private void BuildUI()
        {
            this.Text = "Dashboard - Theo dõi tiến độ";
            this.BackColor = Color.FromArgb(245, 245, 245);

            // Header
            var panelHeader = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(this.Width, 45),
                BackColor = Color.FromArgb(0, 120, 212),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            panelHeader.Controls.Add(new Label
            {
                Text = "📊 DASHBOARD THEO DÕI TIẾN ĐỘ",
                Font = new Font("Segoe UI", 13, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(15, 10),
                Size = new Size(500, 28)
            });
            var btnRefreshAll = new Button
            {
                Text = "🔄 Làm mới tất cả",
                Size = new Size(140, 28),
                BackColor = Color.FromArgb(0, 90, 170),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Cursor = Cursors.Hand
            };
            btnRefreshAll.FlatAppearance.BorderSize = 0;
            btnRefreshAll.Click += (s, e) => LoadData();
            panelHeader.Controls.Add(btnRefreshAll);

            this.Controls.Add(panelHeader);

            // Tab Control
            tabMain = new TabControl
            {
                Location = new Point(0, 45),
                Size = new Size(this.Width, this.Height - 45),
                Font = new Font("Segoe UI", 10),
                Padding = new Point(20, 5),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            tabPO = new TabPage("🛒  Tiến độ giao hàng PO");
            tabMPR = new TabPage("📋  Tiến độ đặt hàng MPR");
            tabRIR = new TabPage("📦  Tiến độ kiểm tra RIR theo PO");

            tabMain.TabPages.Add(tabPO);
            tabMain.TabPages.Add(tabMPR);
            tabMain.TabPages.Add(tabRIR);
            this.Controls.Add(tabMain);

            // SỰ KIỆN TỰ ĐỘNG CHIA TỶ LỆ SONG SONG 70/30
            this.Resize += (s, e) =>
            {
                if (panelHeader != null)
                {
                    panelHeader.Width = this.ClientSize.Width;
                    btnRefreshAll.Location = new Point(panelHeader.Width - 150, 8);
                }

                if (tabMain != null)
                {
                    tabMain.Size = new Size(this.ClientSize.Width, this.ClientSize.Height - 45);
                }

                if (dgvPO != null && dgvPOImports != null && tabPO != null)
                {
                    int totalW = tabPO.ClientSize.Width - 30;
                    int totalH = tabPO.ClientSize.Height - 175 - 10;

                    int poW = (int)(totalW * 0.65);
                    int impW = totalW - poW - 10;

                    dgvPO.Width = Math.Max(100, poW);
                    dgvPO.Height = Math.Max(80, totalH);

                    var lblImport = tabPO.Controls.Find("lblImportTitle", false).FirstOrDefault();
                    if (lblImport != null) { lblImport.Left = dgvPO.Right + 10; lblImport.Width = Math.Max(50, impW); }

                    dgvPOImports.Left = dgvPO.Right + 10;
                    dgvPOImports.Width = Math.Max(80, impW);
                    dgvPOImports.Height = Math.Max(80, totalH);
                }
            };

            BuildPOTab();
            BuildMPRTab();
            BuildRIRTab();
        }

        // ===== PO TAB =====
        private void BuildPOTab()
        {
            tabPO.BackColor = Color.FromArgb(245, 245, 245);
            panelPOSummary = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(this.ClientSize.Width - 20, 95),
                BackColor = Color.Transparent,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            tabPO.Controls.Add(panelPOSummary);
            lblPOTotal = AddSummaryCard(panelPOSummary, "Tổng PO", "0", Color.FromArgb(0, 120, 212), 0);
            lblPOInProgress = AddSummaryCard(panelPOSummary, "Đang giao", "0", Color.FromArgb(255, 140, 0), 1);
            lblPOOverdue = AddSummaryCard(panelPOSummary, "Quá hạn", "0", Color.FromArgb(220, 53, 69), 2);
            lblPOCompleted = AddSummaryCard(panelPOSummary, "Hoàn thành", "0", Color.FromArgb(40, 167, 69), 3);

            // Filter row
            int fy = 115;
            // Filter bar — dùng FlowLayoutPanel để tự wrap khi màn hình nhỏ
            var pFilterPO = new FlowLayoutPanel
            {
                Location = new Point(10, fy),
                Size = new Size(tabPO.ClientSize.Width - 20, 32),
                AutoSize = false,
                WrapContents = false,
                FlowDirection = FlowDirection.LeftToRight,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = Color.Transparent
            };
            tabPO.Controls.Add(pFilterPO);
            tabPO.ClientSizeChanged += (s, e) => pFilterPO.Width = tabPO.ClientSize.Width - 20;

            pFilterPO.Controls.Add(new Label { Text = "Tìm kiếm:", Size = new Size(75, 25), TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Segoe UI", 9) });
            txtSearchPO = new TextBox
            {
                Size = new Size(200, 25),
                Font = new Font("Segoe UI", 9),
                PlaceholderText = "PO No hoặc MPR No..."
            };
            txtSearchPO.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) LoadPOData(); };
            pFilterPO.Controls.Add(txtSearchPO);

            pFilterPO.Controls.Add(new Label { Text = "Trạng thái:", Size = new Size(75, 25), TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Segoe UI", 9) });
            cboFilterPO = new ComboBox
            {
                Size = new Size(150, 25),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboFilterPO.Items.AddRange(new[]
            {
    "Tất cả",
    // ── Trạng thái tính theo % giao hàng ──
    "New",
    "Completed",    // % = 100 → giao đủ
    "Pending",      // 0 < % < 100 → đang giao dở
    // ── Trạng thái gốc từ DB (khi % = 0) ──
    "Draft",
    "Approved",
    "In Progress",
    "Cancelled"
});
            cboFilterPO.SelectedIndex = 0;
            cboFilterPO.SelectedIndexChanged += (s, e) => LoadPOData();
            tabPO.Controls.Add(cboFilterPO);

            var btnSearch = CreateButton("🔍 Tìm", Color.FromArgb(0, 120, 212), Point.Empty, 90, 28);
            btnSearch.Click += (s, e) => LoadPOData();
            pFilterPO.Controls.Add(btnSearch);

            var btnClear = CreateButton("✖ Xóa lọc", Color.FromArgb(108, 117, 125), Point.Empty, 90, 28);
            btnClear.Click += (s, e) => { txtSearchPO.Text = ""; cboFilterPO.SelectedIndex = 0; LoadPOData(); };
            pFilterPO.Controls.Add(btnClear);

            // TIÊU ĐỀ BẢNG BÊN TRÁI
            tabPO.Controls.Add(new Label
            {
                Text = "📑 DANH SÁCH ĐƠN HÀNG (PO)",
                Location = new Point(10, 155),
                Size = new Size(300, 20),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212)
            });

            // BẢNG BÊN TRÁI
            dgvPO = new DataGridView
            {
                Location = new Point(10, 175),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom | AnchorStyles.Right
            };
            dgvPO.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvPO.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPO.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvPO.EnableHeadersVisualStyles = false;
            dgvPO.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvPO.DefaultCellStyle.SelectionBackColor = Color.FromArgb(204, 232, 255);
            dgvPO.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgvPO.RowPrePaint += DgvPO_RowPrePaint;
            dgvPO.SelectionChanged += DgvPO_SelectionChanged;
            tabPO.Controls.Add(dgvPO);

            // TIÊU ĐỀ BẢNG BÊN PHẢI
            tabPO.Controls.Add(new Label
            {
                Text = "📋 PHIẾU NHẬP KHO CỦA PO",
                Location = new Point(600, 155), // Sẽ tự cập nhật lại trong Form_Resize
                Size = new Size(250, 20),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(255, 140, 0),
                Name = "lblImportTitle"
            });

            // BẢNG BÊN PHẢI
            dgvPOImports = new DataGridView
            {
                Location = new Point(600, 175),
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
            dgvPOImports.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 140, 0);
            dgvPOImports.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPOImports.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvPOImports.EnableHeadersVisualStyles = false;
            dgvPOImports.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(255, 248, 235);
            dgvPOImports.DefaultCellStyle.SelectionBackColor = Color.FromArgb(204, 232, 255);
            dgvPOImports.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgvPOImports.CellDoubleClick += DgvPOImports_CellDoubleClick;
            tabPO.Controls.Add(dgvPOImports);
        }

        // =========================================================================
        // ĐỘ RỘNG CỘT BẢNG "DANH SÁCH ĐƠN HÀNG (PO)"
        // Chỉnh width tại đây để thay đổi độ rộng từng cột
        // =========================================================================
        private void AutoAdjustPOColumns()
        {
            if (dgvPO.Columns.Count == 0) return;
            dgvPO.SuspendLayout();
            dgvPO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

            // ── Cấu hình độ rộng từng cột — chỉnh số ở đây ──
            var colWidths = new Dictionary<string, int>
            {
                { "PO No",                 160 },
                { "Dự án",                  50 },
                { "MPR No",                130 },
                { "Ngày PO",                90 },
                { "Rev",                    40 },
                { "Tổng items",             70 },
                { "Tổng SL đặt",            80 },
                { "Tổng SL nhận",           80 },
                { "Ngày giao sớm nhất",    110 },
                { "Trạng thái",            100 },
                { "% Giao hàng",            85 },
                { "Cảnh báo",               90 },
            };

            foreach (DataGridViewColumn col in dgvPO.Columns)
            {
                if (!col.Visible) continue;
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;

                if (colWidths.TryGetValue(col.Name, out int w))
                    col.Width = w;
                else
                    col.Width = 80; // mặc định cho cột chưa khai báo
            }

            dgvPO.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            dgvPO.ResumeLayout();
        }

        // =========================================================================
        // ĐỘ RỘNG CỘT BẢNG "DANH SÁCH YÊU CẦU MUA HÀNG MPR" (bảng phải - tab MPR)
        // Chỉnh width tại đây để thay đổi độ rộng từng cột
        // =========================================================================
        private void AutoAdjustMPRColumns()
        {
            if (dgvMPR == null || dgvMPR.Columns.Count == 0) return;
            dgvMPR.SuspendLayout();
            dgvMPR.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

            var colWidths = new Dictionary<string, int>
            {
                { "MPR No",             180 },
                { "Dự án",               55 },
                { "Ngày cần",            90 },
                { "Trạng thái",          95 },
                { "Rev",                 40 },
                { "Tổng items",          75 },
                { "Tình trạng PO",      110 },
                { "% Item đặt hàng",     95 },
                { "Ngày đến PO",         85 },
                { "Ngày tạo",            90 },
            };

            foreach (DataGridViewColumn col in dgvMPR.Columns)
            {
                if (!col.Visible) continue;
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                if (colWidths.TryGetValue(col.Name, out int w))
                    col.Width = w;
                else
                    col.Width = 80;
            }

            dgvMPR.ResumeLayout();
        }

        // =========================================================================
        // ĐỘ RỘNG CỘT BẢNG "PO CỦA MPR ĐANG CHỌN" (bảng trái - tab MPR)
        // Chỉnh width tại đây để thay đổi độ rộng từng cột
        // =========================================================================
        private void AutoAdjustMPRPOColumns()
        {
            if (dgvMPRPO == null || dgvMPRPO.Columns.Count == 0) return;
            dgvMPRPO.SuspendLayout();
            dgvMPRPO.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

            var colWidths = new Dictionary<string, int>
            {
                { "PO No",               120 },
                { "Dự án",                55 },
                { "Ngày PO",              90 },
                { "Trạng thái",           95 },
                { "Tổng tiền",           110 },
                { "Số dòng vật tư",       90 },
                { "Số RIR",              130 },
            };

            foreach (DataGridViewColumn col in dgvMPRPO.Columns)
            {
                if (!col.Visible) continue;
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                if (colWidths.TryGetValue(col.Name, out int w))
                    col.Width = w;
                else
                    col.Width = 80;
            }

            dgvMPRPO.ResumeLayout();
        }
        private void DgvPO_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvPO.SelectedRows.Count == 0) return;

            int poId = Convert.ToInt32(dgvPO.SelectedRows[0].Cells["PO_ID"].Value);
            string poNo = dgvPO.SelectedRows[0].Cells["PO No"].Value.ToString().Replace("🔥 ", "").Replace(" (Mới)", "");

            Control lbl = tabPO.Controls.Find("lblImportTitle", false)[0];
            lbl.Text = $"📋 PHIẾU NHẬP KHO CỦA: {poNo}";

            // Query lấy danh sách các mã phiếu nhập của PO này
            string sql = $@"
                SELECT 
                    Import_No AS [Mã phiếu], 
                    MAX(Import_Date) AS [Ngày nhập],
                    SUM(Qty_Import) AS [SL Nhập]
                FROM Warehouse_Import 
                WHERE PO_ID = {poId} 
                GROUP BY Import_No 
                ORDER BY MAX(Import_Date) DESC";

            try
            {
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var dt = new DataTable();
                    dt.Load(new SqlCommand(sql, conn).ExecuteReader());
                    dgvPOImports.DataSource = dt;
                    if (dgvPOImports.Columns.Contains("Ngày nhập"))
                        dgvPOImports.Columns["Ngày nhập"].DefaultCellStyle.Format = "dd/MM/yyyy";
                }
            }
            catch { }
        }

        private void DgvPOImports_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            // Lấy PO No từ dgvPO (bảng bên trái)
            string poNo = dgvPO.SelectedRows[0].Cells["PO No"].Value.ToString().Replace("🔥 ", "").Replace(" (Mới)", "");

            // Khởi tạo frmWarehouses_v2 và truyền tham số poNo để nó tự auto search
            frmWarehouses_v2 frm = new frmWarehouses_v2(poNo);
            frm.Show();
        }

        // ===== MPR TAB =====
        private void BuildMPRTab()
        {
            tabMPR.BackColor = Color.FromArgb(245, 245, 245);

            // ===== SUMMARY CARDS =====
            panelMPRSummary = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(this.ClientSize.Width - 20, 95),
                BackColor = Color.Transparent,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            tabMPR.Controls.Add(panelMPRSummary);
            lblMPRTotal = AddSummaryCard(panelMPRSummary, "Tổng MPR", "0", Color.FromArgb(0, 120, 212), 0);
            lblMPRHasPO = AddSummaryCard(panelMPRSummary, "Đã có PO", "0", Color.FromArgb(40, 167, 69), 1);
            lblMPRNoPO = AddSummaryCard(panelMPRSummary, "Chưa có PO", "0", Color.FromArgb(220, 53, 69), 2);
            lblMPRCompleted = AddSummaryCard(panelMPRSummary, "Hoàn thành", "0", Color.FromArgb(102, 51, 153), 3);

            // ===== FILTER BAR =====
            int fy = 115;
            var pFilterMPR = new FlowLayoutPanel
            {
                Location = new Point(10, fy),
                Size = new Size(tabMPR.ClientSize.Width - 20, 32),
                WrapContents = false,
                FlowDirection = FlowDirection.LeftToRight,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = Color.Transparent
            };
            tabMPR.Controls.Add(pFilterMPR);
            tabMPR.ClientSizeChanged += (s, e) => pFilterMPR.Width = tabMPR.ClientSize.Width - 20;

            pFilterMPR.Controls.Add(new Label { Text = "Tìm kiếm:", Size = new Size(75, 25), TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Segoe UI", 9) });
            txtSearchMPR = new TextBox
            {
                Size = new Size(180, 25),
                Font = new Font("Segoe UI", 9),
                PlaceholderText = "MPR No hoặc tên dự án..."
            };
            txtSearchMPR.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) LoadMPRData(); };
            pFilterMPR.Controls.Add(txtSearchMPR);

            pFilterMPR.Controls.Add(new Label { Text = "Trạng thái:", Size = new Size(75, 25), TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Segoe UI", 9) });
            cboFilterMPR = new ComboBox
            {
                Size = new Size(140, 25),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboFilterMPR.Items.AddRange(new[] { "Tất cả", "Mới", "Đang xử lý", "Đã duyệt", "Hoàn thành", "Hủy" });
            cboFilterMPR.SelectedIndex = 0;
            cboFilterMPR.SelectedIndexChanged += (s, e) => LoadMPRData();
            tabMPR.Controls.Add(cboFilterMPR);

            pFilterMPR.Controls.Add(new Label { Text = "% Đặt hàng:", Size = new Size(80, 25), TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Segoe UI", 9) });
            cboFilterPOStatus = new ComboBox
            {
                Size = new Size(160, 25),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboFilterPOStatus.Items.AddRange(new[] { "Tất cả", "✅ Hoàn thành (≥100%)", "⏳ Chưa hoàn thành (<100%)" });
            cboFilterPOStatus.SelectedIndex = 0;
            cboFilterPOStatus.SelectedIndexChanged += (s, e) => FilterMPRByPOStatus();
            pFilterMPR.Controls.Add(cboFilterPOStatus);

            var btnSearch = CreateButton("🔍 Tìm", Color.FromArgb(0, 120, 212), Point.Empty, 80, 28);
            var btnClear = CreateButton("✖ Xóa lọc", Color.FromArgb(108, 117, 125), Point.Empty, 80, 28);
            btnExportMPR = CreateButton("📥 Excel", Color.FromArgb(0, 150, 100), Point.Empty, 80, 28);
            btnSearch.Click += (s, e) => LoadMPRData();
            btnClear.Click += (s, e) =>
            {
                txtSearchMPR.Text = "";
                cboFilterMPR.SelectedIndex = 0;
                cboFilterPOStatus.SelectedIndex = 0;
                LoadMPRData();
            };
            btnExportMPR.Click += BtnExportMPR_Click;
            pFilterMPR.Controls.Add(btnSearch);
            pFilterMPR.Controls.Add(btnClear);
            pFilterMPR.Controls.Add(btnExportMPR);

            // ===== LAYOUT: dgvMPRPO (trái) | dgvMPR (phải) =====
            // Dùng hằng số, KHÔNG dùng tabMPR.Width/Height vì lúc init = 0
            const int topGrid = 150;
            const int poW = 600;  // bảng PO bên trái cố định 520px
            const int gap = 6;
            const int poLeft = 10;
            const int mprLeft = poLeft + poW + gap;

            // ── Label + Bảng PO (TRÁI) ──
            lblMPRPOTitle = new Label
            {
                Text = "📋  PO của MPR đang chọn",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(255, 140, 0),
                Location = new Point(poLeft, topGrid),
                Size = new Size(poW, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            tabMPR.Controls.Add(lblMPRPOTitle);

            dgvMPRPO = new DataGridView
            {
                Location = new Point(poLeft, topGrid + 22),
                Size = new Size(poW, 400),   // chiều cao sẽ do Resize handler
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom
            };
            dgvMPRPO.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 140, 0);
            dgvMPRPO.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvMPRPO.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvMPRPO.EnableHeadersVisualStyles = false;
            dgvMPRPO.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(255, 248, 235);
            dgvMPRPO.DefaultCellStyle.SelectionBackColor = Color.FromArgb(204, 232, 255);
            dgvMPRPO.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgvMPRPO.CellDoubleClick += DgvMPRPO_CellDoubleClick;

            // Tô màu tím cho cột RIR No
            dgvMPRPO.CellFormatting += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                if (dgvMPRPO.Columns[e.ColumnIndex].Name == "RIR No")
                {
                    string val = e.Value?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(val))
                    {
                        e.CellStyle.ForeColor = Color.FromArgb(102, 51, 153);
                        e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    }
                }
            };

            tabMPR.Controls.Add(dgvMPRPO);

            // ── Label + Bảng MPR (PHẢI) ──
            var lblMPRListTitle = new Label
            {
                Text = "DANH SÁCH YÊU CẦU MUA HÀNG MPR",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(mprLeft, topGrid),
                Size = new Size(800, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            tabMPR.Controls.Add(lblMPRListTitle);

            dgvMPR = new DataGridView
            {
                Location = new Point(mprLeft, topGrid + 22),
                Size = new Size(800, 400),   // chiều cao/rộng do Resize handler
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
            dgvMPR.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvMPR.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvMPR.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvMPR.EnableHeadersVisualStyles = false;
            dgvMPR.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvMPR.DefaultCellStyle.SelectionBackColor = Color.FromArgb(204, 232, 255);
            dgvMPR.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgvMPR.RowPrePaint += DgvMPR_RowPrePaint;
            dgvMPR.CellDoubleClick += DgvMPR_CellDoubleClick;
            dgvMPR.SelectionChanged += DgvMPR_SelectionChanged;
            tabMPR.Controls.Add(dgvMPR);

            // Resize: chạy khi tab thay đổi kích thước — điều chỉnh chiều rộng/cao thực tế
            void ApplyMPRLayout()
            {
                if (tabMPR == null || dgvMPR == null || dgvMPRPO == null) return;
                int w = tabMPR.ClientSize.Width;
                int h = tabMPR.ClientSize.Height;
                if (w < 100 || h < 100) return;

                int gridHeight = h - topGrid - 32;
                int mprW = w - mprLeft - 10;

                lblMPRPOTitle.Size = new Size(poW, 20);
                dgvMPRPO.Size = new Size(poW, Math.Max(80, gridHeight));

                lblMPRListTitle.Left = mprLeft;
                lblMPRListTitle.Size = new Size(Math.Max(100, mprW), 20);
                dgvMPR.Left = mprLeft;
                dgvMPR.Size = new Size(Math.Max(100, mprW), Math.Max(80, gridHeight));
            }

            tabMPR.ClientSizeChanged += (s, e) => ApplyMPRLayout();
            // Gọi ngay trong Load của form để layout đúng khi mở
            this.Load += (s, e) => ApplyMPRLayout();
        }

        // Double click dgvMPR → mở frmMPR
        private void DgvMPR_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var row = dgvMPR.Rows[e.RowIndex];
            int mprId = Convert.ToInt32(row.Cells["MPR_ID"].Value);
            new frmMPR(mprId).Show();
        }

        // Chọn dòng MPR → load danh sách PO vào dgvMPRPO
        private void DgvMPR_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvMPR.SelectedRows.Count == 0) return;
            string mprNo = dgvMPR.SelectedRows[0].Cells["MPR No"].Value?.ToString() ?? "";
            if (lblMPRPOTitle != null)
                lblMPRPOTitle.Text = $"📋  PO của MPR: {mprNo}  —  double click để mở";
            LoadPOForMPR(mprNo);
        }

        // =====================================================================
        // LoadPOForMPR — dùng join qua MPR_Detail_ID (không phụ thuộc PO_head.MPR_No)
        // Hỗ trợ 1 MPR có nhiều PO: hiển thị từng PO riêng với đầy đủ thông tin
        // =====================================================================
        private void LoadPOForMPR(string mprNo)
        {
            if (dgvMPRPO == null || string.IsNullOrEmpty(mprNo)) return;
            try
            {
                string sql = @"
                    SELECT
                        po.PO_ID,
                        po.PONo                                             AS [PO No],
                        po.Project_Name                                     AS [Dự án],
                        CONVERT(NVARCHAR(10), po.PO_Date, 103)              AS [Ngày PO],
                        po.Status                                           AS [Trạng thái],
                        FORMAT(po.Total_Amount, 'N0')                       AS [Tổng tiền],
                        (SELECT COUNT(DISTINCT pod2.PO_Detail_ID)
                         FROM PO_Detail pod2 WHERE pod2.PO_ID = po.PO_ID)   AS [Số dòng vật tư],
                        ISNULL(
                            STUFF((
                                SELECT DISTINCT ', ' + r.RIR_No
                                FROM RIR_head r
                                WHERE r.PONo = po.PONo
                                FOR XML PATH(''), TYPE
                            ).value('.', 'NVARCHAR(MAX)'), 1, 2, ''),
                        'Chưa có RIR')                                      AS [Số RIR],
                        po.PO_Date                                          AS _SortDate
                    FROM PO_head po
                    INNER JOIN MPR_Header mh ON mh.MPR_No = @mprNo
                    WHERE
                        po.MPR_No = @mprNo
                        OR po.PO_ID IN (
                            SELECT DISTINCT pod.PO_ID
                            FROM PO_Detail pod
                            INNER JOIN MPR_Details md ON md.Detail_ID = pod.MPR_Detail_ID
                            WHERE md.MPR_ID = mh.MPR_ID
                        )
                    ORDER BY po.PO_Date DESC";

                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@mprNo", mprNo);
                    var dt = new DataTable();
                    dt.Load(cmd.ExecuteReader());
                    dgvMPRPO.DataSource = dt;
                    if (dgvMPRPO.Columns.Contains("PO_ID"))
                        dgvMPRPO.Columns["PO_ID"].Visible = false;
                    if (dgvMPRPO.Columns.Contains("_SortDate"))
                        dgvMPRPO.Columns["_SortDate"].Visible = false;

                    AutoAdjustMPRPOColumns();

                    // Format màu cột RIR
                    if (!dgvMPRPO.Columns.Contains("Số RIR")) return;
                    int rirColIdx = dgvMPRPO.Columns["Số RIR"].Index;
                    foreach (DataGridViewRow row in dgvMPRPO.Rows)
                    {
                        string rirVal = row.Cells["Số RIR"].Value?.ToString() ?? "";
                        if (rirVal == "Chưa có RIR")
                        {
                            row.Cells["Số RIR"].Style.ForeColor = Color.FromArgb(220, 53, 69);
                            row.Cells["Số RIR"].Style.Font = new Font("Segoe UI", 9, FontStyle.Italic);
                        }
                        else if (!string.IsNullOrEmpty(rirVal))
                        {
                            row.Cells["Số RIR"].Style.ForeColor = Color.FromArgb(40, 167, 69);
                            row.Cells["Số RIR"].Style.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                        }
                    }

                    // Cập nhật tiêu đề với số PO tìm được
                    if (lblMPRPOTitle != null)
                        lblMPRPOTitle.Text = $"📋  PO của MPR: {mprNo}  —  Tìm thấy {dt.Rows.Count} PO  —  double click để mở";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("LoadPOForMPR: " + ex.Message);
                MessageBox.Show("Lỗi tải danh sách PO:\n" + ex.Message, "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Double click dgvMPRPO → mở frmPO và tìm PO đó
        private void DgvMPRPO_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string poNo = dgvMPRPO.Rows[e.RowIndex].Cells["PO No"].Value?.ToString() ?? "";
            if (!string.IsNullOrEmpty(poNo))
                new frmPO(poNo).Show();
        }

        // ===== RIR TAB =====
        private void BuildRIRTab()
        {
            tabRIR.BackColor = Color.FromArgb(245, 245, 245);
            panelRIRSummary = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(this.ClientSize.Width - 20, 95),
                BackColor = Color.Transparent,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            tabRIR.Controls.Add(panelRIRSummary);
            lblRIRTotal = AddSummaryCard(panelRIRSummary, "Tổng RIR", "0", Color.FromArgb(0, 120, 212), 0);
            lblRIRPending = AddSummaryCard(panelRIRSummary, "Chờ kiểm tra", "0", Color.FromArgb(255, 140, 0), 1);
            lblRIRInspecting = AddSummaryCard(panelRIRSummary, "Đang kiểm tra", "0", Color.FromArgb(102, 51, 153), 2);
            lblRIRDone = AddSummaryCard(panelRIRSummary, "Hoàn thành", "0", Color.FromArgb(40, 167, 69), 3);

            int fy = 115;
            var pFilterRIR = new FlowLayoutPanel
            {
                Location = new Point(10, fy),
                Size = new Size(tabRIR.ClientSize.Width - 20, 32),
                WrapContents = false,
                FlowDirection = FlowDirection.LeftToRight,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = Color.Transparent
            };
            tabRIR.Controls.Add(pFilterRIR);
            tabRIR.ClientSizeChanged += (s, e) => pFilterRIR.Width = tabRIR.ClientSize.Width - 20;

            pFilterRIR.Controls.Add(new Label { Text = "Tìm kiếm:", Size = new Size(75, 25), TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Segoe UI", 9) });
            txtSearchRIR = new TextBox
            {
                Size = new Size(200, 25),
                Font = new Font("Segoe UI", 9),
                PlaceholderText = "RIR No hoặc PO No..."
            };
            txtSearchRIR.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) LoadRIRData(); };
            pFilterRIR.Controls.Add(txtSearchRIR);

            pFilterRIR.Controls.Add(new Label { Text = "Trạng thái:", Size = new Size(75, 25), TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Segoe UI", 9) });
            cboFilterRIR = new ComboBox
            {
                Size = new Size(150, 25),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboFilterRIR.Items.AddRange(new[] { "Tất cả", "Chờ kiểm tra", "Đang kiểm tra", "Hoàn thành" });
            cboFilterRIR.SelectedIndex = 0;
            cboFilterRIR.SelectedIndexChanged += (s, e) => LoadRIRData();
            pFilterRIR.Controls.Add(cboFilterRIR);

            var btnSearch = CreateButton("🔍 Tìm", Color.FromArgb(0, 120, 212), Point.Empty, 90, 28);
            btnSearch.Click += (s, e) => LoadRIRData();
            pFilterRIR.Controls.Add(btnSearch);

            var btnClear = CreateButton("✖ Xóa lọc", Color.FromArgb(108, 117, 125), Point.Empty, 90, 28);
            btnClear.Click += (s, e) => { txtSearchRIR.Text = ""; cboFilterRIR.SelectedIndex = 0; LoadRIRData(); };
            pFilterRIR.Controls.Add(btnClear);
            tabRIR.Controls.Add(new Label
            {
                Text = "DANH SÁCH PO & TIẾN ĐỘ RIR",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 152),
                Size = new Size(300, 20)
            });

            const int RIR_TOP = 173;
            const int RIR_LBL_H = 22;

            dgvRIR = BuildGrid(tabRIR, RIR_TOP);
            dgvRIR.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            dgvRIR.RowPrePaint += DgvRIR_RowPrePaint;
            dgvRIR.SelectionChanged += DgvRIR_SelectionChanged;

            var lblDetailTitle = new Label
            {
                Text = "CHI TIẾT RIR THEO PO (click vào PO ở trên để xem)",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Size = new Size(600, RIR_LBL_H),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            tabRIR.Controls.Add(lblDetailTitle);

            dgvRIRDetail = BuildGrid(tabRIR, RIR_TOP + 100 + RIR_LBL_H);
            dgvRIRDetail.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(102, 51, 153);
            dgvRIRDetail.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 240, 255);
            dgvRIRDetail.CellFormatting += DgvRIRDetail_CellFormatting;

            // Resize động: dgvRIR = 40%, dgvRIRDetail = 55% của vùng còn lại
            void ApplyRIRLayout()
            {
                if (tabRIR == null || dgvRIR == null || dgvRIRDetail == null) return;
                int w = tabRIR.ClientSize.Width - 20;
                int h = tabRIR.ClientSize.Height;
                if (w < 50 || h < 200) return;

                int available = h - RIR_TOP - 10;
                int topH = Math.Max(80, (int)(available * 0.40));
                int lblY = RIR_TOP + topH + 4;
                int bottomTop = lblY + RIR_LBL_H + 2;
                int bottomH = Math.Max(80, h - bottomTop - 10);

                dgvRIR.Location = new Point(10, RIR_TOP);
                dgvRIR.Size = new Size(w, topH);

                lblDetailTitle.Location = new Point(10, lblY);
                lblDetailTitle.Width = w;

                dgvRIRDetail.Location = new Point(10, bottomTop);
                dgvRIRDetail.Size = new Size(w, bottomH);
            }
            tabRIR.ClientSizeChanged += (s, e) => ApplyRIRLayout();
            this.Load += (s, e) => ApplyRIRLayout();
        }

        // ===== HELPERS =====
        private DataGridView BuildGrid(TabPage tab, int top)
        {
            var dgv = new DataGridView
            {
                Location = new Point(10, top),
                Size = new Size(tab.Width - 20, tab.Height - top - 10),
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
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.EnableHeadersVisualStyles = false;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);

            // MÀU CHỌN XANH NHẠT CHO TẤT CẢ CÁC GRID
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(204, 232, 255);
            dgv.DefaultCellStyle.SelectionForeColor = Color.Black;

            tab.Controls.Add(dgv);
            return dgv;
        }

        private Label AddSummaryCard(Panel parent, string title, string value, Color color, int slotIndex)
        {
            // Tự tính vị trí và kích thước theo tỷ lệ, 4 card đều nhau
            const int CARD_COUNT = 4;
            const int GAP = 8;
            // Card sẽ resize khi parent thay đổi — dùng Anchor + SizeChanged
            int cardW = Math.Max(100, (parent.Width - GAP * (CARD_COUNT + 1)) / CARD_COUNT);
            int cardX = GAP + slotIndex * (cardW + GAP);

            var card = new Panel
            {
                Location = new Point(cardX, 4),
                Size = new Size(cardW, 86),
                BackColor = color,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            parent.Controls.Add(card);

            card.Controls.Add(new Label
            {
                Text = title,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.White,
                Dock = DockStyle.None,
                Location = new Point(0, 8),
                Size = new Size(cardW, 22),
                TextAlign = ContentAlignment.MiddleCenter,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            });
            var lbl = new Label
            {
                Text = value,
                Font = new Font("Segoe UI", 22, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(0, 32),
                Size = new Size(cardW, 50),
                TextAlign = ContentAlignment.MiddleCenter,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            card.Controls.Add(lbl);

            // Resize card khi parent resize
            parent.SizeChanged += (s, e) =>
            {
                int newW = Math.Max(100, (parent.Width - GAP * (CARD_COUNT + 1)) / CARD_COUNT);
                int newX = GAP + slotIndex * (newW + GAP);
                card.Location = new Point(newX, 4);
                card.Width = newW;
                foreach (Control c in card.Controls) c.Width = newW;
            };

            return lbl;
        }

        private void AddLabel(TabPage tab, string text, int x, int y)
        {
            tab.Controls.Add(new Label { Text = text, Location = new Point(x, y + 3), Size = new Size(75, 20), Font = new Font("Segoe UI", 9) });
        }

        private Button CreateButton(string text, Color color, Point loc, int w, int h)
        {
            var btn = new Button { Text = text, Location = loc, Size = new Size(w, h), BackColor = color, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Cursor = Cursors.Hand };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        // ===== LOAD DATA =====
        private void LoadData()
        {
            LoadPOData();
            LoadMPRData();
            LoadRIRData();
        }

        private void LoadPOData()
        {
            try
            {
                string search = txtSearchPO.Text.Trim();
                string filter = cboFilterPO.SelectedItem?.ToString() ?? "Tất cả";

                string searchCondition = "";
                if (!string.IsNullOrEmpty(search))
                {
                    searchCondition = $" AND (h.PONo LIKE N'%{search}%' OR h.MPR_No LIKE N'%{search}%' OR h.Project_Name LIKE N'%{search}%')";
                }

                string filterCondition = "";
                if (filter != "Tất cả")
                {
                    filterCondition = $" WHERE [Trạng thái] = N'{filter}'";
                }

                string sql = $@"
                    WITH POStats AS (
                        SELECT
                            h.PO_ID,
                            h.PONo                             AS [PO No],
                            h.Project_Name                     AS [Dự án],
                            h.MPR_No                           AS [MPR No],
                            h.PO_Date                          AS [Ngày PO],
                            h.Revise                           AS [Rev],
                            COUNT(d.PO_Detail_ID)              AS [Tổng items],
                            ISNULL(SUM(d.Qty_Per_Sheet), 0)    AS [Tổng SL đặt],
                            ISNULL((SELECT SUM(Qty_Import) FROM Warehouse_Import wi WHERE wi.PO_ID = h.PO_ID), ISNULL(SUM(d.Received), 0)) AS [Tổng SL nhận],
                            MIN(d.RequestDay)                  AS [Ngày giao sớm nhất],
                            h.Status                           AS [TrangThaiDB]
                        FROM PO_head h
                        LEFT JOIN PO_Detail d ON h.PO_ID = d.PO_ID
                        WHERE 1=1 {searchCondition}
                        GROUP BY h.PO_ID, h.PONo, h.Project_Name, h.MPR_No, h.PO_Date, h.Status, h.Revise
                    ),
                    CalculatedPO AS (
                        SELECT
                            PO_ID,
                            [PO No],
                            [Dự án],
                            [MPR No],
                            [Ngày PO],
                            CASE
                                WHEN [Tổng SL đặt] > 0 AND CAST([Tổng SL nhận] * 100.0 / [Tổng SL đặt] AS DECIMAL(5,1)) >= 100 THEN N'Completed'
                                WHEN [Tổng SL đặt] > 0 AND CAST([Tổng SL nhận] * 100.0 / [Tổng SL đặt] AS DECIMAL(5,1)) > 0 THEN N'Pending'
                                ELSE [TrangThaiDB]
                            END AS [Trạng thái],
                            [Rev],
                            [Tổng items],
                            [Tổng SL đặt],
                            [Tổng SL nhận],
                            CASE
                                WHEN [Tổng SL đặt] = 0 THEN 0
                                ELSE CAST([Tổng SL nhận] * 100.0 / [Tổng SL đặt] AS DECIMAL(5,1))
                            END AS [% Giao hàng],
                            [Ngày giao sớm nhất],
                            CASE
                                WHEN [Ngày giao sớm nhất] < GETDATE() AND [Tổng SL nhận] < [Tổng SL đặt] THEN N'⚠ Quá hạn'
                                ELSE N'✅ Đúng hạn'
                            END AS [Cảnh báo]
                        FROM POStats
                    )
                    SELECT * FROM CalculatedPO
                    {filterCondition}
                    ORDER BY [Ngày PO] DESC";
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var dt = new DataTable();
                    dt.Load(new SqlCommand(sql, conn).ExecuteReader());
                    dgvPO.DataSource = dt;
                    if (dgvPO.Columns.Contains("PO_ID"))
                        dgvPO.Columns["PO_ID"].Visible = false;

                    dgvPO.CellFormatting -= DgvPO_CellFormatting;
                    dgvPO.CellFormatting += DgvPO_CellFormatting;

                    AutoAdjustPOColumns(); // Gọi hàm auto giãn cột

                    // Summary
                    int total = dt.Rows.Count, overdue = 0, completed = 0, inProgress = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        decimal pct = row["% Giao hàng"] != DBNull.Value ? Convert.ToDecimal(row["% Giao hàng"]) : 0;
                        string canh = row["Cảnh báo"]?.ToString() ?? "";

                        if (pct >= 100) completed++;
                        else if (canh.Contains("Quá")) overdue++;
                        else if (pct > 0) inProgress++;
                    }
                    lblPOTotal.Text = total.ToString();
                    lblPOInProgress.Text = inProgress.ToString();
                    lblPOOverdue.Text = overdue.ToString();
                    lblPOCompleted.Text = completed.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải PO: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DgvPO_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string col = dgvPO.Columns[e.ColumnIndex].Name;

            if (col == "% Giao hàng")
            {
                if (decimal.TryParse(e.Value?.ToString(), out decimal pct))
                {
                    e.CellStyle.ForeColor = pct >= 100 ? Color.FromArgb(40, 167, 69) : pct > 0 ? Color.FromArgb(255, 140, 0) : Color.FromArgb(220, 53, 69);
                    e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    e.Value = $"{pct}%";
                    e.FormattingApplied = true;
                }
            }
            else if (col == "Cảnh báo")
            {
                e.CellStyle.ForeColor = e.Value?.ToString().Contains("Quá") == true ? Color.Red : Color.FromArgb(40, 167, 69);
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
            else if (col == "Trạng thái")
            {
                string val = e.Value?.ToString() ?? "";
                if (val == "Completed") e.CellStyle.ForeColor = Color.FromArgb(40, 167, 69);
                else if (val == "Pending") e.CellStyle.ForeColor = Color.FromArgb(255, 140, 0);
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
        }

        private void DgvPO_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex < 0 || dgvPO.Rows[e.RowIndex].IsNewRow) return;
            var row = dgvPO.Rows[e.RowIndex];
            string canh = row.Cells["Cảnh báo"].Value?.ToString() ?? "";
            string status = row.Cells["Trạng thái"].Value?.ToString() ?? "";

            // Xử lý background color bình thường (màu chữ chọn đã cấu hình chung Xanh nhạt)
            if (canh.Contains("Quá")) row.DefaultCellStyle.BackColor = Color.FromArgb(255, 235, 235);
            else if (status == "Completed") row.DefaultCellStyle.BackColor = Color.FromArgb(235, 255, 235);
            else if (status == "In Progress" || status == "Approved" || status == "Pending")
                row.DefaultCellStyle.BackColor = Color.FromArgb(255, 248, 235);
        }

        private void LoadMPRData()
        {
            try
            {
                string search = txtSearchMPR.Text.Trim();
                string filter = cboFilterMPR.SelectedItem?.ToString() ?? "Tất cả";

                string where = "WHERE 1=1";
                if (!string.IsNullOrEmpty(search))
                    where += $" AND (h.MPR_No LIKE N'%{search}%' OR h.Project_Name LIKE N'%{search}%')";
                if (filter != "Tất cả")
                    where += $" AND h.Status = N'{filter}'";

                string sql = $@"
                    SELECT
                        h.MPR_ID,
                        h.MPR_No                           AS [MPR No],
                        h.Project_Name                     AS [Dự án],
                        h.Required_Date                    AS [Ngày cần],
                        h.Status                           AS [Trạng thái],
                        h.Rev                              AS [Rev],
                        COUNT(DISTINCT d.Detail_ID)        AS [Tổng items],
                        CASE
                            WHEN COUNT(DISTINCT po.PO_ID) > 0
                            THEN N'✅ ' + CAST(COUNT(DISTINCT po.PO_ID) AS NVARCHAR(10)) + N' PO'
                            ELSE N'❌ Chưa có PO'
                        END                                AS [Tình trạng PO],
                        CASE
                            WHEN COUNT(DISTINCT d.Detail_ID) = 0 THEN 0
                            ELSE CAST(
                                COUNT(DISTINCT pod.PO_Detail_ID) * 100.0
                                / COUNT(DISTINCT d.Detail_ID)
                                AS DECIMAL(5,1))
                        END                                AS [% Item đặt hàng],
                        DATEDIFF(DAY, h.Created_Date, MIN(po.Created_Date)) AS [Ngày đến PO],
                        h.Created_Date                     AS [Ngày tạo]
                    FROM MPR_Header h
                    LEFT JOIN MPR_Details d   ON d.MPR_ID = h.MPR_ID
                    LEFT JOIN PO_Detail   pod ON pod.MPR_Detail_ID = d.Detail_ID
                    -- Lấy PO qua 2 cách: qua MPR_Detail_ID HOẶC qua PO_head.MPR_No trực tiếp
                    LEFT JOIN PO_head     po  ON po.PO_ID = pod.PO_ID
                                              OR po.MPR_No = h.MPR_No
                    {where}
                    GROUP BY h.MPR_ID, h.MPR_No, h.Project_Name,
                             h.Required_Date, h.Status, h.Rev, h.Created_Date
                    ORDER BY h.Created_Date DESC";
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var dt = new DataTable();
                    dt.Load(new SqlCommand(sql, conn).ExecuteReader());
                    dgvMPR.DataSource = dt;

                    if (dgvMPR.Columns.Contains("MPR_ID"))
                        dgvMPR.Columns["MPR_ID"].Visible = false;

                    AutoAdjustMPRColumns();

                    dgvMPR.CellFormatting -= DgvMPR_CellFormatting;
                    dgvMPR.CellFormatting += DgvMPR_CellFormatting;

                    int total = dt.Rows.Count, hasPO = 0, noPO = 0, completed = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        string tinh = row["Tình trạng PO"]?.ToString() ?? "";
                        string status = row["Trạng thái"]?.ToString() ?? "";

                        if (!tinh.Contains("Chưa có")) hasPO++;
                        else noPO++;

                        if (status == "Hoàn thành") completed++;
                    }
                    lblMPRTotal.Text = total.ToString();
                    lblMPRHasPO.Text = hasPO.ToString();
                    lblMPRNoPO.Text = noPO.ToString();
                    lblMPRCompleted.Text = completed.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải MPR: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DgvMPR_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string col = dgvMPR.Columns[e.ColumnIndex].Name;

            if (col == "% Item đặt hàng")
            {
                if (decimal.TryParse(e.Value?.ToString(), out decimal pct))
                {
                    e.CellStyle.ForeColor = pct >= 100 ? Color.FromArgb(40, 167, 69) : pct >= 50 ? Color.FromArgb(255, 140, 0) : Color.FromArgb(220, 53, 69);
                    e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    e.Value = $"{pct}%";
                    e.FormattingApplied = true;
                }
            }
            if (col == "Tình trạng PO")
            {
                e.CellStyle.ForeColor = e.Value?.ToString().Contains("Chưa có") == true ? Color.FromArgb(220, 53, 69) : Color.FromArgb(40, 167, 69);
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
            if (col == "Ngày đến PO")
            {
                e.Value = e.Value != DBNull.Value && e.Value != null ? $"{e.Value} ngày" : "—";
                e.FormattingApplied = true;
            }
        }

        private void DgvMPR_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex < 0 || dgvMPR.Rows[e.RowIndex].IsNewRow) return;
            var row = dgvMPR.Rows[e.RowIndex];
            string tinh = row.Cells["Tình trạng PO"].Value?.ToString() ?? "";
            string status = row.Cells["Trạng thái"].Value?.ToString() ?? "";
            if (status == "Hoàn thành") row.DefaultCellStyle.BackColor = Color.FromArgb(235, 255, 235);
            else if (tinh.Contains("Chưa có")) row.DefaultCellStyle.BackColor = Color.FromArgb(255, 235, 235);
            else row.DefaultCellStyle.BackColor = Color.FromArgb(255, 248, 235);
        }

        // Auto giãn cột dgvMPR: min 30, max 150 — cột "Dự án" đã set 60 trước khi gọi hàm này
        // Lọc dgvMPR theo Tình trạng PO (client-side, không query lại DB)
        private void FilterMPRByPOStatus()
        {
            if (dgvMPR == null || dgvMPR.Rows.Count == 0) return;
            string sel = cboFilterPOStatus.SelectedItem?.ToString() ?? "Tất cả";

            foreach (DataGridViewRow row in dgvMPR.Rows)
            {
                if (row.IsNewRow) continue;
                if (sel == "Tất cả") { row.Visible = true; continue; }

                // Đọc % Item đặt hàng từ cột (có thể dạng "100.0%" hoặc số)
                string pctRaw = row.Cells["% Item đặt hàng"].Value?.ToString() ?? "0";
                pctRaw = pctRaw.Replace("%", "").Trim();
                decimal.TryParse(pctRaw, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out decimal pct);

                if (sel.Contains("Hoàn thành"))
                    row.Visible = pct >= 100;
                else if (sel.Contains("Chưa hoàn thành"))
                    row.Visible = pct < 100;
                else
                    row.Visible = true;
            }
        }

        // Xuất Excel tổng hợp MPR + PO
        private void BtnExportMPR_Click(object sender, EventArgs e)
        {
            if (dgvMPR == null || dgvMPR.Rows.Count == 0)
            { MessageBox.Show("Không có dữ liệu để xuất!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            using var sfd = new SaveFileDialog
            {
                Title = "Lưu báo cáo MPR",
                Filter = "Excel|*.xlsx",
                FileName = $"BaoCao_MPR_{DateTime.Now:yyyyMMdd_HHmm}"
            };
            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using var pkg = new ExcelPackage();

                var ws = pkg.Workbook.Worksheets.Add("Chi tiết MPR");

                // ── Tiêu đề file ──
                int TOTAL_COLS = 16; // sẽ cập nhật theo hdrs
                ws.Cells[1, 1].Value = "BÁO CÁO CHI TIẾT ĐẶT HÀNG MPR";
                ws.Cells[1, 1, 1, TOTAL_COLS].Merge = true;
                ws.Cells[1, 1].Style.Font.Size = 14;
                ws.Cells[1, 1].Style.Font.Bold = true;
                ws.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[2, 1].Value = $"Xuất ngày: {DateTime.Now:dd/MM/yyyy HH:mm}";
                ws.Cells[2, 1, 2, TOTAL_COLS].Merge = true;

                // ── Lấy danh sách MPR No đang HIỂN THỊ ──
                var mprNos = new System.Collections.Generic.List<string>();
                foreach (DataGridViewRow row in dgvMPR.Rows)
                {
                    if (row.IsNewRow || !row.Visible) continue;
                    string mno = row.Cells["MPR No"].Value?.ToString();
                    if (!string.IsNullOrEmpty(mno)) mprNos.Add(mno);
                }
                if (mprNos.Count == 0) { MessageBox.Show("Không có MPR nào!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

                string inClause = string.Join(",", mprNos.Select(m => $"N'{m.Replace("'", "''")}'"));

                // ── Query: mỗi hạng mục MPR = 1 dòng, đầy đủ tất cả cột MPR_Details ──
                string sql = @"
                    SELECT
                        h.MPR_No,
                        h.Project_Name,
                        h.Status                                    AS MPR_Status,
                        CONVERT(NVARCHAR(10), h.Required_Date, 103) AS Required_Date,
                        ISNULL(h.Notes, '')                         AS MPR_Notes,
                        d.Item_No,
                        ISNULL(d.item_name,     '')  AS Item_Name,
                        ISNULL(d.Description,   '')  AS Description,
                        ISNULL(d.Material,      '')  AS Material,
                        ISNULL(CAST(NULLIF(d.Thickness_mm,0) AS NVARCHAR),'') AS A_Day,
                        ISNULL(CAST(NULLIF(d.Depth_mm,    0) AS NVARCHAR),'') AS B_Sau,
                        ISNULL(CAST(NULLIF(d.C_Width_mm,  0) AS NVARCHAR),'') AS C_Rong,
                        ISNULL(CAST(NULLIF(d.D_Web_mm,    0) AS NVARCHAR),'') AS D_Bung,
                        ISNULL(CAST(NULLIF(d.E_Flange_mm, 0) AS NVARCHAR),'') AS E_Canh,
                        ISNULL(CAST(NULLIF(d.F_Length_mm, 0) AS NVARCHAR),'') AS F_Dai,
                        ISNULL(d.UNIT,          '')  AS UNIT,
                        ISNULL(d.Qty_Per_Sheet, 0)   AS SL,
                        ISNULL(d.Weight_kg,     0)   AS KG,
                        ISNULL(d.MPS_Info,     '')   AS MPS_Info,
                        ISNULL(d.Usage_Location,'')  AS Usage_Location,
                        ISNULL(d.REV,          '0')  AS REV,
                        ISNULL(d.Remarks,      '')   AS Detail_Remarks,
                        ISNULL(STUFF((
                            SELECT DISTINCT ', ' + pox.PONo
                            FROM PO_Detail podx
                            INNER JOIN PO_head pox ON pox.PO_ID = podx.PO_ID
                            WHERE podx.MPR_Detail_ID = d.Detail_ID
                            FOR XML PATH(''), TYPE
                        ).value('.','NVARCHAR(MAX)'), 1, 2, ''), '') AS PO_List,
                        ISNULL(STUFF((
                            SELECT DISTINCT ', ' + r.RIR_No
                            FROM RIR_head r
                            WHERE r.PONo IN (
                                SELECT pox2.PONo
                                FROM PO_Detail podx2
                                INNER JOIN PO_head pox2 ON pox2.PO_ID = podx2.PO_ID
                                WHERE podx2.MPR_Detail_ID = d.Detail_ID
                            )
                            FOR XML PATH(''), TYPE
                        ).value('.','NVARCHAR(MAX)'), 1, 2, ''), '') AS RIR_List
                    FROM MPR_Header  h
                    INNER JOIN MPR_Details d ON d.MPR_ID = h.MPR_ID
                    WHERE h.MPR_No IN (" + inClause + @")
                    ORDER BY h.MPR_No, d.Item_No";

                DataTable dt;
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    dt = new DataTable();
                    dt.Load(new SqlCommand(sql, conn).ExecuteReader());
                }

                // ── Header cột — khớp đúng với SQL ──
                // Cột 1-5  : thông tin MPR header
                // Cột 6-22 : chi tiết hạng mục MPR_Details đầy đủ
                // Cột 23-24: PO và RIR
                string[] hdrs = {
                    // MPR header (5 cột)
                    "MPR No", "Dự án", "TT MPR", "Ngày cần", "Ghi chú MPR",
                    // Chi tiết hạng mục MPR_Details (17 cột)
                    "STT", "Tên vật tư", "Mô tả", "Vật liệu",
                    "A-Dày(mm)", "B-Sâu(mm)", "C-Rộng(mm)", "D-Bụng(mm)", "E-Cánh(mm)", "F-Dài(mm)",
                    "ĐVT", "Số lượng", "KG",
                    "MPS Info", "Nơi dùng", "REV", "Ghi chú",
                    // PO và RIR (2 cột)
                    "Số PO", "Số RIR"
                };
                TOTAL_COLS = hdrs.Length; // = 24

                // Cập nhật merge tiêu đề
                ws.Cells[1, 1, 1, TOTAL_COLS].Merge = true;
                ws.Cells[2, 1, 2, TOTAL_COLS].Merge = true;

                // Ghi header (dòng 4)
                for (int c = 0; c < hdrs.Length; c++)
                {
                    var hCell = ws.Cells[4, c + 1];
                    hCell.Value = hdrs[c];
                    hCell.Style.Font.Bold = true;
                    hCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    hCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    hCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    hCell.Style.WrapText = true;
                    hCell.Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    // Màu header: MPR=xanh đậm (1-5), chi tiết=xanh dương (6-22), PO/RIR=tím (23-24)
                    Color hColor = c < 5 ? Color.FromArgb(0, 70, 127) :
                                   c < 22 ? Color.FromArgb(0, 120, 212) :
                                            Color.FromArgb(102, 51, 153);
                    hCell.Style.Fill.BackgroundColor.SetColor(hColor);
                    hCell.Style.Font.Color.SetColor(Color.White);
                }
                ws.Row(4).Height = 30;

                // ── Ghi dữ liệu ──
                int rowIdx = 5;
                string lastMprNo = "";
                int colorToggle = 0;

                foreach (DataRow dr in dt.Rows)
                {
                    string mprNo = dr["MPR_No"]?.ToString() ?? "";

                    // Dòng tiêu đề nhóm khi đổi MPR
                    if (mprNo != lastMprNo)
                    {
                        if (lastMprNo != "") rowIdx++; // dòng trống ngăn cách

                        ws.Cells[rowIdx, 1, rowIdx, TOTAL_COLS].Merge = true;
                        ws.Cells[rowIdx, 1].Value =
                            $"  📋  MPR: {mprNo}  |  Dự án: {dr["Project_Name"]}  " +
                            $"|  Ngày cần: {dr["Required_Date"]}  |  Trạng thái: {dr["MPR_Status"]}";
                        ws.Cells[rowIdx, 1].Style.Font.Bold = true;
                        ws.Cells[rowIdx, 1].Style.Font.Size = 10;
                        ws.Cells[rowIdx, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[rowIdx, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 140, 0));
                        ws.Cells[rowIdx, 1].Style.Font.Color.SetColor(Color.White);
                        ws.Row(rowIdx).Height = 20;
                        rowIdx++;
                        lastMprNo = mprNo;
                        colorToggle = 0;
                    }

                    // Màu nền xen kẽ
                    var bg = colorToggle % 2 == 0 ? Color.White : Color.FromArgb(240, 248, 255);

                    // ── Cột 1-5: thông tin MPR ──
                    ws.Cells[rowIdx, 1].Value = dr["MPR_No"]?.ToString();
                    ws.Cells[rowIdx, 2].Value = dr["Project_Name"]?.ToString();
                    ws.Cells[rowIdx, 3].Value = dr["MPR_Status"]?.ToString();
                    ws.Cells[rowIdx, 4].Value = dr["Required_Date"]?.ToString();
                    ws.Cells[rowIdx, 5].Value = dr["MPR_Notes"]?.ToString();

                    // ── Cột 6-22: chi tiết hạng mục MPR_Details ──
                    ws.Cells[rowIdx, 6].Value = dr["Item_No"] != DBNull.Value ? Convert.ToInt32(dr["Item_No"]) : (object)"";
                    ws.Cells[rowIdx, 7].Value = dr["Item_Name"]?.ToString();      // Tên vật tư
                    ws.Cells[rowIdx, 8].Value = dr["Description"]?.ToString();    // Mô tả
                    ws.Cells[rowIdx, 9].Value = dr["Material"]?.ToString();       // Vật liệu
                    ws.Cells[rowIdx, 10].Value = dr["A_Day"]?.ToString();          // A-Dày
                    ws.Cells[rowIdx, 11].Value = dr["B_Sau"]?.ToString();          // B-Sâu
                    ws.Cells[rowIdx, 12].Value = dr["C_Rong"]?.ToString();         // C-Rộng
                    ws.Cells[rowIdx, 13].Value = dr["D_Bung"]?.ToString();         // D-Bụng
                    ws.Cells[rowIdx, 14].Value = dr["E_Canh"]?.ToString();         // E-Cánh
                    ws.Cells[rowIdx, 15].Value = dr["F_Dai"]?.ToString();          // F-Dài
                    ws.Cells[rowIdx, 16].Value = dr["UNIT"]?.ToString();           // ĐVT
                    ws.Cells[rowIdx, 17].Value = dr["SL"] != DBNull.Value ? Convert.ToDecimal(dr["SL"]) : (object)"";  // SL
                    ws.Cells[rowIdx, 18].Value = dr["KG"] != DBNull.Value ? Convert.ToDecimal(dr["KG"]) : (object)"";  // KG
                    ws.Cells[rowIdx, 19].Value = dr["MPS_Info"]?.ToString();       // MPS Info
                    ws.Cells[rowIdx, 20].Value = dr["Usage_Location"]?.ToString(); // Nơi dùng
                    ws.Cells[rowIdx, 21].Value = dr["REV"]?.ToString();            // REV
                    ws.Cells[rowIdx, 22].Value = dr["Detail_Remarks"]?.ToString(); // Ghi chú

                    // ── Cột 23: Số PO ──
                    string poList = dr["PO_List"]?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(poList))
                    {
                        ws.Cells[rowIdx, 23].Value = poList;
                        ws.Cells[rowIdx, 23].Style.Font.Color.SetColor(Color.FromArgb(0, 120, 212));
                        ws.Cells[rowIdx, 23].Style.Font.Bold = poList.Contains(",");
                    }
                    else
                    {
                        ws.Cells[rowIdx, 23].Value = "Chưa có PO";
                        ws.Cells[rowIdx, 23].Style.Font.Color.SetColor(Color.FromArgb(220, 53, 69));
                        ws.Cells[rowIdx, 23].Style.Font.Italic = true;
                    }

                    // ── Cột 24: Số RIR ──
                    string rirList = dr["RIR_List"]?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(rirList))
                    {
                        ws.Cells[rowIdx, 24].Value = rirList;
                        ws.Cells[rowIdx, 24].Style.Font.Color.SetColor(Color.FromArgb(40, 167, 69));
                        ws.Cells[rowIdx, 24].Style.Font.Bold = rirList.Contains(",");
                    }
                    else
                    {
                        ws.Cells[rowIdx, 24].Value = "";
                    }

                    // Tô màu toàn dòng
                    ws.Cells[rowIdx, 1, rowIdx, TOTAL_COLS].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[rowIdx, 1, rowIdx, TOTAL_COLS].Style.Fill.BackgroundColor.SetColor(bg);

                    // Tô nền đỏ nhạt vùng PO/RIR nếu chưa có PO
                    if (string.IsNullOrEmpty(poList))
                    {
                        ws.Cells[rowIdx, 23, rowIdx, 24].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[rowIdx, 23, rowIdx, 24].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 235, 235));
                    }

                    // Border từng dòng
                    ws.Cells[rowIdx, 1, rowIdx, TOTAL_COLS].Style.Border.BorderAround(ExcelBorderStyle.Hair);

                    colorToggle++;
                    rowIdx++;
                }

                // Border toàn bộ vùng data
                if (dt.Rows.Count > 0)
                {
                    var dataRange = ws.Cells[4, 1, rowIdx - 1, TOTAL_COLS];
                    dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }

                // Căn chỉnh cột số (STT, kích thước, SL, KG)
                foreach (int c in new[] { 6, 10, 11, 12, 13, 14, 15, 17, 18 })
                    if (c <= TOTAL_COLS)
                        ws.Column(c).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                ws.Cells[ws.Dimension.Address].AutoFitColumns();

                // Giới hạn width cột PO và RIR
                ws.Column(23).Width = Math.Min(ws.Column(23).Width, 50);
                ws.Column(24).Width = Math.Min(ws.Column(24).Width, 50);

                ws.View.FreezePanes(5, 1);

                pkg.SaveAs(new FileInfo(sfd.FileName));
                MessageBox.Show(
                    $"✅ Xuất Excel thành công!\n" +
                    $"{mprNos.Count} MPR, {dt.Rows.Count} hạng mục.",
                    "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                { FileName = sfd.FileName, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xuất Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadRIRData()
        {
            try
            {
                string search = txtSearchRIR.Text.Trim();
                string filter = cboFilterRIR.SelectedItem?.ToString() ?? "Tất cả";

                string where = "WHERE 1=1";
                if (!string.IsNullOrEmpty(search))
                    where += $" AND (h.RIR_No LIKE N'%{search}%' OR h.PONo LIKE N'%{search}%' OR h.Project_Name LIKE N'%{search}%')";
                if (filter != "Tất cả")
                    where += $" AND h.Status = N'{filter}'";
                string sql = $@"
                    SELECT
                        h.RIR_ID,
                        h.RIR_No                             AS [RIR No],
                        h.PONo                                              AS [PO No],
                        h.MPR_No                                            AS [MPR No],
                        h.Project_Name                                      AS [Dự án],
                        h.Issue_Date                                        AS [Ngày phát hành],
                        h.Customer                                          AS [Khách hàng],
                        h.Status                                            AS [Trạng thái],
                        COUNT(d.RIR_Detail_ID)                              AS [Tổng items],
                        ISNULL(SUM(d.Qty_Required), 0)                      AS [Tổng SL YC],
                        ISNULL(SUM(d.Qty_Received), 0)                      AS [Tổng SL nhận],
                        COUNT(CASE WHEN d.Inspect_Result = 'Pass' THEN 1 END) AS [Pass],
                        COUNT(CASE WHEN d.Inspect_Result = 'Fail' THEN 1 END) AS [Fail],
                        COUNT(CASE WHEN d.Inspect_Result = 'Hold' THEN 1 END) AS [Hold],
                        CASE
                            WHEN COUNT(d.RIR_Detail_ID) = 0 THEN 0
                            ELSE CAST(COUNT(CASE WHEN d.Inspect_Result = 'Pass' THEN 1 END) * 100.0 / COUNT(d.RIR_Detail_ID) AS DECIMAL(5,1))
                        END                                                 AS [% Pass]
                    FROM RIR_head h
                    LEFT JOIN RIR_detail d ON h.RIR_ID = d.RIR_ID
                    {where}
                    GROUP BY h.RIR_ID, h.RIR_No, h.PONo, h.MPR_No, h.Project_Name,
                             h.Issue_Date, h.Customer, h.Status
                    ORDER BY h.Issue_Date DESC";

                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var dt = new DataTable();
                    dt.Load(new SqlCommand(sql, conn).ExecuteReader());
                    dgvRIR.DataSource = dt;
                    if (dgvRIR.Columns.Contains("RIR_ID"))
                        dgvRIR.Columns["RIR_ID"].Visible = false;
                    dgvRIR.CellFormatting -= DgvRIR_CellFormatting;
                    dgvRIR.CellFormatting += DgvRIR_CellFormatting;

                    int total = dt.Rows.Count, pending = 0, inspecting = 0, done = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        string status = row["Trạng thái"]?.ToString() ?? "";
                        if (status == "Chờ kiểm tra") pending++;
                        else if (status == "Đang kiểm tra") inspecting++;
                        else if (status == "Hoàn thành") done++;
                    }
                    lblRIRTotal.Text = total.ToString();
                    lblRIRPending.Text = pending.ToString();
                    lblRIRInspecting.Text = inspecting.ToString();
                    lblRIRDone.Text = done.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải RIR: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DgvRIR_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string col = dgvRIR.Columns[e.ColumnIndex].Name;
            if (col == "% Pass")
            {
                if (decimal.TryParse(e.Value?.ToString(), out decimal pct))
                {
                    e.CellStyle.ForeColor = pct >= 100 ? Color.FromArgb(40, 167, 69) : pct >= 50 ? Color.FromArgb(255, 140, 0) : Color.FromArgb(220, 53, 69);
                    e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    e.Value = $"{pct}%";
                    e.FormattingApplied = true;
                }
            }
            if (col == "Trạng thái")
            {
                e.CellStyle.ForeColor = e.Value?.ToString() == "Hoàn thành" ? Color.FromArgb(40, 167, 69) :
                                        e.Value?.ToString() == "Đang kiểm tra" ? Color.FromArgb(255, 140, 0) :
                                        Color.FromArgb(0, 120, 212);
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
        }

        private void DgvRIR_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex < 0 || dgvRIR.Rows[e.RowIndex].IsNewRow) return;
            var row = dgvRIR.Rows[e.RowIndex];
            string status = row.Cells["Trạng thái"].Value?.ToString() ?? "";
            if (status == "Hoàn thành") row.DefaultCellStyle.BackColor = Color.FromArgb(235, 255, 235);
            else if (status == "Đang kiểm tra") row.DefaultCellStyle.BackColor = Color.FromArgb(255, 248, 235);
            else row.DefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
        }

        private void DgvRIR_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvRIR.SelectedRows.Count == 0) return;
            var row = dgvRIR.SelectedRows[0];
            string poNo = row.Cells["PO No"].Value?.ToString() ?? "";
            if (string.IsNullOrEmpty(poNo)) return;
            LoadRIRDetailByPO(poNo);
        }

        private void LoadRIRDetailByPO(string poNo)
        {
            try
            {
                string sql = $@"
                    SELECT
                        h.RIR_No                                            AS [RIR No],
                        h.Issue_Date                                        AS [Ngày phát hành],
                        h.Status                                            AS [Trạng thái RIR],
                        h.Customer                                          AS [Khách hàng],
                        d.Item_No                                           AS [STT],
                        d.item_name                                         AS [Tên vật tư],
                        d.Material                                          AS [Vật liệu],
                        d.Size                                              AS [Kích thước],
                        d.UNIT                                              AS [ĐVT],
                        d.Qty_Required                                      AS [SL YC],
                        d.Qty_Received                                      AS [SL nhận],
                        d.MTRno                                             AS [MTR No],
                        d.Heatno                                            AS [Heat No],
                        d.ID_Code                                           AS [ID Code],
                        ISNULL(d.Inspect_Result, N'Chưa KT')                AS [Kết quả KT]
                    FROM RIR_head h
                    INNER JOIN RIR_detail d ON h.RIR_ID = d.RIR_ID
                    WHERE h.PONo = N'{poNo}'
                    ORDER BY h.RIR_No, d.Item_No";
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var dt = new DataTable();
                    dt.Load(new SqlCommand(sql, conn).ExecuteReader());
                    dgvRIRDetail.DataSource = dt;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải chi tiết RIR: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DgvRIRDetail_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string col = dgvRIRDetail.Columns[e.ColumnIndex].Name;
            if (col == "Kết quả KT")
            {
                string val = e.Value?.ToString() ?? "";
                e.CellStyle.ForeColor = val == "Pass" ? Color.FromArgb(40, 167, 69) :
                                        val == "Fail" ? Color.FromArgb(220, 53, 69) :
                                        val == "Hold" ? Color.FromArgb(255, 140, 0) :
                                        Color.Gray;
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
            if (col == "Trạng thái RIR")
            {
                e.CellStyle.ForeColor = e.Value?.ToString() == "Hoàn thành" ? Color.FromArgb(40, 167, 69) :
                                        e.Value?.ToString() == "Đang kiểm tra" ? Color.FromArgb(255, 140, 0) :
                                        Color.FromArgb(0, 120, 212);
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
        }
    }
}