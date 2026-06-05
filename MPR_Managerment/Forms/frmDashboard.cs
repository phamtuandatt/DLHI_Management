using System;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;
using MPR_Managerment.Services;
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
        private DataGridView dgvMPRDetail;  // Bảng chi tiết vật tư của PO đang chọn
        private Panel panelMPRDetail;        // Panel bao quanh bảng chi tiết
        private Label lblMPRDetailTitle;     // Tiêu đề bảng chi tiết
        private Label lblMPRPOTitle; // Tiêu đề bảng PO
        private Label lblMPRTotal, lblMPRHasPO, lblMPRNoPO, lblMPRCompleted;
        private Panel panelMPRSummary;
        private ComboBox cboFilterMPR;
        private ComboBox cboFilterPOStatus;  // Lọc theo Tình trạng PO
        private TextBox txtSearchMPR;
        private Button btnExportMPR;         // Xuất Excel danh sách MPR
        private Button btnSaveMPRNote;        // Lưu ghi chú MPR

        // RIR Tab
        private DataGridView dgvRIR;
        private Label lblRIRTotal, lblRIRPending, lblRIRInspecting, lblRIRDone;
        private Panel panelRIRSummary;
        private ComboBox cboFilterRIR;
        private TextBox txtSearchRIR;
        private DataGridView dgvRIRDetail;
        // NOTIFICATION SYSTEM
        private Panel panelNotify;
        private ListBox lstNotify;
        private Label lblNotifyTitle, lblNotifyCount;
        private System.Windows.Forms.Timer _notifyTimer;
        private DateTime _lastCheckTime = DateTime.MinValue;
        private Button btnNotifyToggle;
        private Panel _toastPanel;

        // Lưu trạng thái đã gửi Zalo để không bị mất khi làm mới (trong 1 session)
        private HashSet<int> _sentZaloPOs = new HashSet<int>();

        private void EnsureZaloNotificationTable()
        {
            try
            {
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    string checkSql = @"
                        IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='PO_ZaloNotification' AND xtype='U')
                        BEGIN
                            CREATE TABLE PO_ZaloNotification (
                                PO_ID INT PRIMARY KEY,
                                SentDate DATETIME DEFAULT GETDATE()
                            )
                        END";
                    using (var cmd = new SqlCommand(checkSql, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("EnsureZaloNotificationTable error: " + ex.Message);
            }
        }

        private void SaveZaloSentStatus(int poId)
        {
            try
            {
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    string sql = "IF NOT EXISTS (SELECT 1 FROM PO_ZaloNotification WHERE PO_ID = @poId) INSERT INTO PO_ZaloNotification (PO_ID) VALUES (@poId)";
                    using (var cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@poId", poId);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("SaveZaloSentStatus error: " + ex.Message);
            }
        }

        public frmDashboard()
        {
            InitializeComponent();
            BuildUI();
            BuildNotificationPanel();
            StartNotifyTimer();

            // Ép form gọi sự kiện Resize lần đầu để chia tỷ lệ ngay khi mở
            this.OnResize(EventArgs.Empty);
            frmAIChat.Attach(this);

            // *** PERFORMANCE FIX: Defer DB loading to after form is shown ***
            // Form hiển thị NGAY LẬP TỨC, sau đó mới load dữ liệu ở background
            this.Shown += FrmDashboard_Shown;
        }

        private async void FrmDashboard_Shown(object sender, EventArgs e)
        {
            this.Shown -= FrmDashboard_Shown;
            var toast = ToastHelper.Attach(this);
            toast.Show("⏳ Đang tải dữ liệu, vui lòng chờ...");
            this.Cursor = Cursors.WaitCursor;
            try
            {
                // EnsureZaloNotificationTable chạy trên background thread
                await Task.Run(() => EnsureZaloNotificationTable());
                await LoadDataAsync();
            }
            finally
            {
                toast.Hide();
                this.Cursor = Cursors.Default;
            }
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
            btnRefreshAll.Click += async (s, e) => await LoadDataAsync();
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
                if (panelNotify != null && panelNotify.Visible)
                    panelNotify.Location = new Point(
                        this.ClientSize.Width - panelNotify.Width - 10,
                        this.ClientSize.Height - panelNotify.Height - 10);

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
            pFilterPO.Controls.Add(cboFilterPO);

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
            dgvPO.CellClick += DgvPO_CellClick;
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
                { "PO No",                 130 },
                { "NCC",                   100 },
                { "Dự án",                  50 },
                { "MPR No",                120 },
                { "Ngày PO",                85 },
                { "Rev",                    35 },
                { "Tổng SL đặt",            75 },
                { "Tổng SL nhận",           75 },
                { "Ngày giao sớm nhất",    105 },
                { "Trạng thái",             85 },
                { "% Giao hàng",            80 },
                { "Cảnh báo",               80 },
                { "Gửi Zalo",              80 },
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

            // Ẩn các cột không cần hiển thị
            foreach (var hidden in new[] { "Trạng thái", "BaseNo", "MaxRev" })
                if (dgvMPR.Columns.Contains(hidden))
                    dgvMPR.Columns[hidden].Visible = false;

            // Ngày tạo chỉ hiển thị ngày/tháng/năm
            if (dgvMPR.Columns.Contains("Ngày tạo"))
                dgvMPR.Columns["Ngày tạo"].DefaultCellStyle.Format = "dd/MM/yyyy";

            var colWidths = new Dictionary<string, int>
            {
                { "MPR No",             200 },
                { "Dự án",               55 },
                { "Ngày cần",            90 },
                { "Rev",                 40 },

                { "Tình trạng PO",      110 },
                { "% Item đặt hàng",     95 },

                { "Ngày tạo",           100 },
                { "Ghi chú",            160 },
            };

            foreach (DataGridViewColumn col in dgvMPR.Columns)
            {
                if (!col.Visible) continue;
                if (col.Name == "Ghi chú" || col.Name == "Ghi chu")
                {
                    col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    continue;
                }
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
            txtSearchMPR.KeyDown += async (s, e) => { if (e.KeyCode == Keys.Enter) await LoadMPRDataAsync(); };
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
            cboFilterMPR.SelectedIndexChanged += async (s, e) => await LoadMPRDataAsync();
            pFilterMPR.Controls.Add(cboFilterMPR);

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
            btnSearch.Click += async (s, e) => await LoadMPRDataAsync();
            btnClear.Click += async (s, e) =>
            {
                txtSearchMPR.Text = "";
                cboFilterMPR.SelectedIndex = 0;
                cboFilterPOStatus.SelectedIndex = 0;
                await LoadMPRDataAsync();
            };
            btnExportMPR.Click += BtnExportMPR_Click;
            btnSaveMPRNote = CreateButton("💾 Lưu ghi chú", Color.FromArgb(0, 120, 212), Point.Empty, 120, 28);
            btnSaveMPRNote.Click += BtnSaveMPRNote_Click;
            pFilterMPR.Controls.Add(btnSearch);
            pFilterMPR.Controls.Add(btnClear);
            pFilterMPR.Controls.Add(btnExportMPR);
            pFilterMPR.Controls.Add(btnSaveMPRNote);

            // ===== LAYOUT: dgvMPRPO (trái) | dgvMPR (phải) =====
            // Dùng hằng số, KHÔNG dùng tabMPR.Width/Height vì lúc init = 0
            const int topGrid = 150;
            const int poW = 600;  // initial value — sẽ override trong ApplyMPRLayout
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
            dgvMPRPO.SelectionChanged += DgvMPRPO_SelectionChanged;

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

            // ── Panel chi tiết vật tư PO (bên dưới dgvMPRPO) ──
            lblMPRDetailTitle = new Label
            {
                Text = "📦  Chi tiết vật tư — click vào PO để xem",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(poLeft, topGrid + 22 + 200 + 5),
                Size = new Size(poW, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            tabMPR.Controls.Add(lblMPRDetailTitle);

            panelMPRDetail = new Panel
            {
                Location = new Point(poLeft, topGrid + 22 + 200 + 27),
                Size = new Size(poW, 180),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom
            };
            tabMPR.Controls.Add(panelMPRDetail);

            dgvMPRDetail = new DataGridView
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
                ColumnHeadersHeight = 28
            };
            dgvMPRDetail.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvMPRDetail.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvMPRDetail.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvMPRDetail.EnableHeadersVisualStyles = false;
            dgvMPRDetail.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(235, 245, 255);
            dgvMPRDetail.DefaultCellStyle.SelectionBackColor = Color.FromArgb(204, 232, 255);
            dgvMPRDetail.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgvMPRDetail.CellFormatting += (s, ev) =>
            {
                if (ev.RowIndex < 0) return;
                string col = dgvMPRDetail.Columns[ev.ColumnIndex].Name;
                if (col == "SL PO" || col == "SL MPR" || col == "Còn lại" || col == "Nhập kho")
                    ev.CellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                if (col == "SL PO" || col == "SL MPR")
                {
                    if (decimal.TryParse(ev.Value?.ToString(), out decimal sl))
                    {
                        ev.Value = sl % 1 == 0 ? ((long)sl).ToString() : sl.ToString("G29");
                        ev.FormattingApplied = true;
                    }
                }
                if (col == "Nhập kho")
                {
                    if (ev.Value == null || ev.Value == DBNull.Value ||
                        (decimal.TryParse(ev.Value?.ToString(), out decimal nk) && nk == 0))
                    {
                        ev.Value = "";
                        ev.FormattingApplied = true;
                    }
                    else if (decimal.TryParse(ev.Value?.ToString(), out decimal nkVal))
                    {
                        ev.Value = nkVal % 1 == 0 ? ((long)nkVal).ToString() : nkVal.ToString("G29");
                        ev.FormattingApplied = true;
                    }
                }
                if (col == "Còn lại")
                {
                    if (decimal.TryParse(ev.Value?.ToString(), out decimal rem))
                    {
                        ev.CellStyle.ForeColor = rem == 0
                            ? Color.FromArgb(40, 167, 69)      // = 0 → xanh lá
                            : rem < 0
                                ? Color.FromArgb(102, 0, 153)  // < 0 → tím đậm
                                : Color.FromArgb(180, 0, 0);   // > 0 → đỏ đậm
                        ev.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    }
                }
            };
            panelMPRDetail.Controls.Add(dgvMPRDetail);

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
                ReadOnly = false,
                EditMode = DataGridViewEditMode.EditOnKeystroke,
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
            // One-click vao cot Ghi chu -> bat dau edit ngay
            dgvMPR.CellClick += (s, ev) =>
            {
                if (ev.RowIndex < 0 || ev.ColumnIndex < 0) return;
                if (dgvMPR.Columns[ev.ColumnIndex].Name == "Ghi chu")
                    dgvMPR.BeginEdit(true);
            };
            tabMPR.Controls.Add(dgvMPR);

            // Resize: chạy khi tab thay đổi kích thước — điều chỉnh chiều rộng/cao thực tế
            void ApplyMPRLayout()
            {
                if (tabMPR == null || dgvMPR == null || dgvMPRPO == null) return;
                int w = tabMPR.ClientSize.Width;
                int h = tabMPR.ClientSize.Height;
                if (w < 100 || h < 100) return;

                int totalH = h - topGrid - 32;
                int halfW = (int)((w - 26) * 0.4);  // 40% bên trái
                int dynPoW = Math.Max(280, halfW);
                int dynMprLeft = poLeft + dynPoW + gap;
                int mprW = Math.Max(100, w - dynMprLeft - 10);
                // Chia chiều cao: dgvMPRPO 30%, panelMPRDetail 70%
                int poGridH = Math.Max(80, (int)((totalH - 27) * 0.30));
                int detailH = Math.Max(80, totalH - poGridH - 27);

                lblMPRPOTitle.Size = new Size(dynPoW, 20);
                dgvMPRPO.Size = new Size(dynPoW, poGridH);

                if (lblMPRDetailTitle != null)
                {
                    lblMPRDetailTitle.Location = new Point(poLeft, dgvMPRPO.Bottom + 5);
                    lblMPRDetailTitle.Size = new Size(dynPoW, 20);
                }
                if (panelMPRDetail != null)
                {
                    panelMPRDetail.Location = new Point(poLeft, dgvMPRPO.Bottom + 27);
                    panelMPRDetail.Size = new Size(dynPoW, detailH);
                }

                lblMPRListTitle.Left = dynMprLeft;
                lblMPRListTitle.Size = new Size(Math.Max(100, mprW), 20);
                dgvMPR.Left = dynMprLeft;
                dgvMPR.Size = new Size(Math.Max(100, mprW), Math.Max(80, totalH));
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
        // =====================================================================
        // LoadPOForMPR — load danh sách PO của MPR vào dgvMPRPO
        // =====================================================================
        private void LoadPOForMPR(string mprNo)
        {
            if (dgvMPRPO == null || string.IsNullOrEmpty(mprNo)) return;
            try
            {
                string sql = @"
                    SELECT DISTINCT
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
                    CROSS APPLY (
                        SELECT CASE WHEN CHARINDEX('_Rev.', @mprNo) > 0
                                    THEN LEFT(@mprNo, CHARINDEX('_Rev.', @mprNo) - 1)
                                    ELSE @mprNo END AS BaseNo
                    ) bn
                    WHERE
                        -- PO liên kết trực tiếp qua MPR_No (bất kỳ revision nào trong cùng series)
                        po.MPR_No = @mprNo
                        OR po.MPR_No = bn.BaseNo
                        OR po.MPR_No LIKE bn.BaseNo + '_Rev.%'
                        -- PO liên kết qua PO_Detail → MPR_Details → bất kỳ revision nào trong series
                        OR po.PO_ID IN (
                            SELECT DISTINCT pod.PO_ID
                            FROM PO_Detail pod
                            INNER JOIN MPR_Details md ON md.Detail_ID = pod.MPR_Detail_ID
                            INNER JOIN MPR_Header mh ON mh.MPR_ID = md.MPR_ID
                            WHERE mh.MPR_No = @mprNo
                               OR mh.MPR_No = bn.BaseNo
                               OR mh.MPR_No LIKE bn.BaseNo + '_Rev.%'
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

                    if (!dgvMPRPO.Columns.Contains("Số RIR")) return;
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

                    if (lblMPRPOTitle != null)
                        lblMPRPOTitle.Text = $"📋  PO của MPR: {mprNo}  —  Tìm thấy {dt.Rows.Count} PO  —  double click để mở";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("LoadPOForMPR: " + ex.Message);
                SafeMsg("Lỗi tải danh sách PO:\n" + ex.Message, "Lỗi");
            }
        }

        // Double click dgvMPRPO → mở frmPO
        private void DgvMPRPO_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string poNo = dgvMPRPO.Rows[e.RowIndex].Cells["PO No"]?.Value?.ToString() ?? "";
            if (!string.IsNullOrEmpty(poNo))
                new frmPO(poNo).Show();
        }

        // Click vào dòng PO trong dgvMPRPO → load chi tiết vật tư
        private void DgvMPRPO_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvMPRPO == null || dgvMPRPO.SelectedRows.Count == 0) return;
            string poNo = dgvMPRPO.SelectedRows[0].Cells["PO No"]?.Value?.ToString()?.Trim() ?? "";
            if (string.IsNullOrEmpty(poNo)) return;
            LoadPODetailForMPR(poNo);
        }

        // Load chi tiết vật tư của PO được chọn vào dgvMPRDetail
        private void LoadPODetailForMPR(string poNo)
        {
            if (dgvMPRDetail == null || string.IsNullOrEmpty(poNo)) return;
            if (lblMPRDetailTitle != null)
                lblMPRDetailTitle.Text = $"📦  Chi tiết vật tư — PO: {poNo} (đang tải...)";
            try
            {
                string sql = @"
                    WITH WI AS (
                        SELECT PO_Detail_ID, ISNULL(SUM(Qty_Import), 0) AS Total
                        FROM Warehouse_Import
                        GROUP BY PO_Detail_ID
                    )
                    SELECT
                        pod.Item_No                                                     AS [STT],
                        ISNULL(pod.item_name,  ISNULL(md.item_name,  ''))              AS [Tên hàng],
                        ISNULL(pod.Material,   ISNULL(md.Material,   ''))              AS [Vật liệu],
                        CASE
                            WHEN NULLIF(pod.Asize, '') IS NOT NULL
                              AND NULLIF(pod.Bsize, '') IS NOT NULL
                              AND NULLIF(pod.Csize, '') IS NOT NULL
                                THEN CAST(pod.Asize AS NVARCHAR(50))
                                   + ' x ' + CAST(pod.Bsize AS NVARCHAR(50))
                                   + ' x ' + CAST(pod.Csize AS NVARCHAR(50))
                            WHEN NULLIF(pod.Asize, '') IS NOT NULL
                              AND NULLIF(pod.Bsize, '') IS NOT NULL
                                THEN CAST(pod.Asize AS NVARCHAR(50))
                                   + ' x ' + CAST(pod.Bsize AS NVARCHAR(50))
                            WHEN NULLIF(pod.Asize, '') IS NOT NULL
                                THEN CAST(pod.Asize AS NVARCHAR(50))
                            ELSE ''
                        END                                                             AS [Size (mm)],
                        ISNULL(pod.Qty_Per_Sheet, 0)                                   AS [SL PO],
                        ISNULL(NULLIF(pod.Unit,''), ISNULL(md.UNIT, ''))               AS [ĐVT],
                        ISNULL(md.Qty_Per_Sheet,   0)                                  AS [SL MPR],
                        ISNULL(md.Qty_Per_Sheet, 0) - ISNULL(pod.Qty_Per_Sheet, 0)    AS [Còn lại],
                        NULLIF(ISNULL(wi.Total, 0), 0)                                 AS [Nhập kho]
                    FROM PO_head ph
                    INNER JOIN PO_Detail   pod ON pod.PO_ID    = ph.PO_ID
                    LEFT  JOIN MPR_Details md  ON md.Detail_ID = pod.MPR_Detail_ID
                    LEFT  JOIN WI          wi  ON wi.PO_Detail_ID = pod.PO_Detail_ID
                    WHERE ph.PONo = @poNo
                    ORDER BY pod.Item_No";

                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@poNo", poNo);
                    var dt = new DataTable();
                    dt.Load(cmd.ExecuteReader());
                    dgvMPRDetail.DataSource = dt;

                    dgvMPRDetail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                    var widths = new Dictionary<string, int>
                    {
                        { "STT", 35 }, { "Vật liệu", 130 },
                        { "Size (mm)", 148 }, { "SL PO", 52 }, { "ĐVT", 45 },
                        { "SL MPR", 55 }, { "Còn lại", 55 }, { "Nhập kho", 65 }
                    };
                    foreach (DataGridViewColumn col in dgvMPRDetail.Columns)
                    {
                        if (col.Name == "Tên hàng")
                        {
                            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                            continue;
                        }
                        col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                        col.Width = widths.TryGetValue(col.Name, out int w) ? w : 80;
                        if (col.Name == "Nhập kho")
                        {
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            col.DefaultCellStyle.NullValue  = "";
                        }
                    }

                    if (lblMPRDetailTitle != null)
                        lblMPRDetailTitle.Text = $"📦  Chi tiết vật tư — PO: {poNo}  ({dt.Rows.Count} hạng mục)  |  double-click PO để mở frmPO";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("LoadPODetailForMPR: " + ex.Message);
                if (lblMPRDetailTitle != null)
                    lblMPRDetailTitle.Text = $"📦  Chi tiết vật tư — PO: {poNo}  (lỗi: {ex.Message})";
            }
        }
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
            txtSearchRIR.KeyDown += async (s, e) => { if (e.KeyCode == Keys.Enter) await LoadRIRDataAsync(); };
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
            cboFilterRIR.SelectedIndexChanged += async (s, e) => await LoadRIRDataAsync();
            pFilterRIR.Controls.Add(cboFilterRIR);

            var btnSearch = CreateButton("🔍 Tìm", Color.FromArgb(0, 120, 212), Point.Empty, 90, 28);
            btnSearch.Click += async (s, e) => await LoadRIRDataAsync();
            pFilterRIR.Controls.Add(btnSearch);

            var btnClear = CreateButton("✖ Xóa lọc", Color.FromArgb(108, 117, 125), Point.Empty, 90, 28);
            btnClear.Click += async (s, e) => { txtSearchRIR.Text = ""; cboFilterRIR.SelectedIndex = 0; await LoadRIRDataAsync(); };
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
            LoadPODataAsync().ConfigureAwait(false);
        }

        private void LoadMPRData()
        {
            LoadMPRDataAsync().ConfigureAwait(false);
        }

        private void LoadRIRData()
        {
            LoadRIRDataAsync().ConfigureAwait(false);
        }

        private async Task LoadDataAsync()
        {
            await Task.WhenAll(LoadPODataAsync(), LoadMPRDataAsync(), LoadRIRDataAsync());
        }

        private async Task LoadPODataAsync()
        {
            try
            {
                string search = txtSearchPO.Text.Trim();
                string filter = cboFilterPO.SelectedItem?.ToString() ?? "Tất cả";

                string searchCondition = "";
                if (!string.IsNullOrEmpty(search))
                    searchCondition = $" AND (h.PONo LIKE N'%{search}%' OR h.MPR_No LIKE N'%{search}%' OR h.Project_Name LIKE N'%{search}%')";

                string filterCondition = "";
                if (filter != "Tất cả")
                    filterCondition = $" WHERE [Trạng thái] = N'{filter}'";

                string sql = $@"
                    WITH WI_Agg AS (
                        SELECT PO_ID, SUM(Qty_Import) AS TotalImport
                        FROM Warehouse_Import
                        GROUP BY PO_ID
                    ),
                    POStats AS (
                        SELECT
                            h.PO_ID,
                            h.PONo                             AS [PO No],
                            h.Project_Name                     AS [Dự án],
                            h.MPR_No                           AS [MPR No],
                            h.PO_Date                          AS [Ngày PO],
                            s.Short_Name                       AS [NCC],
                            h.Revise                           AS [Rev],
                            ISNULL(SUM(d.Qty_Per_Sheet), 0)    AS [Tổng SL đặt],
                            ISNULL(wi.TotalImport, ISNULL(SUM(d.Received), 0)) AS [Tổng SL nhận],
                            MIN(d.RequestDay)                  AS [Ngày giao sớm nhất],
                            h.Status                           AS [TrangThaiDB]
                        FROM PO_head h
                        LEFT JOIN PO_Detail d ON h.PO_ID = d.PO_ID
                        LEFT JOIN Suppliers s ON h.Supplier_ID = s.Supplier_ID
                        LEFT JOIN WI_Agg wi ON wi.PO_ID = h.PO_ID
                        WHERE 1=1 {searchCondition}
                        GROUP BY h.PO_ID, h.PONo, h.Project_Name, h.MPR_No, h.PO_Date, h.Status, h.Revise, s.Short_Name, wi.TotalImport
                    ),
                    CalculatedPO AS (
                        SELECT
                            PO_ID,
                            [PO No],
                            [NCC],
                            [Dự án],
                            [MPR No],
                            [Ngày PO],
                            CASE
                                WHEN [Tổng SL đặt] > 0 AND CAST([Tổng SL nhận] * 100.0 / [Tổng SL đặt] AS DECIMAL(5,1)) >= 100 THEN N'Completed'
                                WHEN [Tổng SL đặt] > 0 AND CAST([Tổng SL nhận] * 100.0 / [Tổng SL đặt] AS DECIMAL(5,1)) > 0 THEN N'Pending'
                                ELSE [TrangThaiDB]
                            END AS [Trạng thái],
                            [Rev],
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

                // Chạy toàn bộ DB I/O trên thread pool, không block UI thread
                var (dt, sentZalos) = await Task.Run(() =>
                {
                    using var conn = DatabaseHelper.GetConnection();
                    conn.Open();

                    var zalos = new HashSet<int>();
                    try
                    {
                        using var cmdZalo = new SqlCommand("SELECT PO_ID FROM PO_ZaloNotification", conn);
                        using var drZalo = cmdZalo.ExecuteReader();
                        while (drZalo.Read())
                            zalos.Add(Convert.ToInt32(drZalo["PO_ID"]));
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("Load sent Zalo error: " + ex.Message);
                    }

                    var table = new DataTable();
                    table.Load(new SqlCommand(sql, conn).ExecuteReader());
                    return (table, zalos);
                });

                // *** ISSUE #2 FIX: Suspend layout & binding during grid update ***
                dgvPO.SuspendLayout();
                try
                {
                    // Temporarily detach event handlers to prevent SelectionChanged from firing 100+ times
                    dgvPO.SelectionChanged -= DgvPO_SelectionChanged;
                    dgvPO.CellFormatting -= DgvPO_CellFormatting;

                    // Cập nhật UI (đang trên UI thread sau await)
                    _sentZaloPOs = sentZalos;
                    dgvPO.DataSource = dt;
                    if (dgvPO.Columns.Contains("PO_ID"))
                        dgvPO.Columns["PO_ID"].Visible = false;

                    // Re-attach event handlers and trigger CellFormatting
                    dgvPO.CellFormatting += DgvPO_CellFormatting;
                    dgvPO.SelectionChanged += DgvPO_SelectionChanged;
                    dgvPO.Invalidate(); // Force redraw
                }
                finally
                {
                    dgvPO.ResumeLayout(true); // Resume layout with force layout = true
                }

                if (!dgvPO.Columns.Contains("Gửi Zalo"))
                {
                    var colZalo = new DataGridViewButtonColumn
                    {
                        Name = "Gửi Zalo",
                        HeaderText = "Gửi Zalo",
                        Text = "📱 Zalo",
                        UseColumnTextForButtonValue = true,
                        Width = 80,
                        FlatStyle = FlatStyle.Flat,
                        DisplayIndex = dgvPO.Columns.Count
                    };
                    colZalo.DefaultCellStyle.BackColor = Color.FromArgb(40, 167, 69);
                    colZalo.DefaultCellStyle.ForeColor = Color.White;
                    colZalo.DefaultCellStyle.Font = new Font("Segoe UI", 8, FontStyle.Bold);
                    colZalo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dgvPO.Columns.Add(colZalo);
                }

                AutoAdjustPOColumns();

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
            catch (Exception ex)
            {
                SafeMsg("Lỗi tải PO: " + ex.Message, "Lỗi");
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
            else if (col == "Gửi Zalo")
            {
                // Lấy PO_ID từ dòng hiện tại
                int poId = Convert.ToInt32(dgvPO.Rows[e.RowIndex].Cells["PO_ID"].Value);

                // Giữ nút xanh lá mặc định, chỉ xám nếu ID nằm trong danh sách đã gửi
                if (_sentZaloPOs.Contains(poId))
                {
                    e.CellStyle.BackColor = Color.FromArgb(180, 180, 180);
                    e.CellStyle.SelectionBackColor = Color.FromArgb(180, 180, 180);
                }
                else
                {
                    e.CellStyle.BackColor = Color.FromArgb(40, 167, 69);
                    e.CellStyle.SelectionBackColor = Color.FromArgb(40, 167, 69);
                }
                e.CellStyle.ForeColor = Color.White;
                e.CellStyle.Font = new Font("Segoe UI", 8, FontStyle.Bold);
            }
        }

        // ── Xử lý click nút Gửi Zalo trong dgvPO ──
        private async void DgvPO_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (dgvPO.Columns[e.ColumnIndex].Name != "Gửi Zalo") return;

            var row = dgvPO.Rows[e.RowIndex];
            int poId = Convert.ToInt32(row.Cells["PO_ID"].Value);
            string poNo = row.Cells["PO No"].Value?.ToString()?.Replace("🔥 ", "").Replace(" (Mới)", "") ?? "";
            string duan = row.Cells["Dự án"].Value?.ToString() ?? "";
            string trangThai = row.Cells["Trạng thái"].Value?.ToString() ?? "";
            string pctGiao = row.Cells["% Giao hàng"].Value?.ToString() ?? "0";
            string canhBao = row.Cells["Cảnh báo"].Value?.ToString() ?? "";

            try
            {
                // Mở form cấu hình giao hàng Zalo
                using (var configForm = new frmZaloDeliveryConfig(poId))
                {
                    if (configForm.ShowDialog() != DialogResult.OK) return;

                    // 1. Lấy Tên nhóm Zalo từ bảng Project
                    string zaloGroupName = "";
                    using (var conn = DatabaseHelper.GetConnection())
                    {
                        conn.Open();
                        string sqlProj = "SELECT ZaloGroupName FROM ProjectInfo WHERE ProjectName = @proj";
                        using var cmdProj = new SqlCommand(sqlProj, conn);
                        cmdProj.Parameters.AddWithValue("@proj", duan);
                        zaloGroupName = cmdProj.ExecuteScalar()?.ToString() ?? "";
                    }

                    // 2. Định dạng danh sách chi tiết các hạng mục được chọn để giao
                    string detailsMsg = "";
                    foreach (var item in configForm.SelectedItems)
                    {
                        if (item.DeliveryQtyNow <= 0) continue; // Bỏ qua nếu số lượng giao lần này = 0

                        string techInfo = "";
                        if (!string.IsNullOrEmpty(item.Material)) techInfo += item.Material;

                        string sizeStr = "";
                        if (!string.IsNullOrEmpty(item.Asize)) sizeStr += item.Asize;
                        if (!string.IsNullOrEmpty(item.Bsize)) sizeStr += " x " + item.Bsize;
                        if (!string.IsNullOrEmpty(item.Csize)) sizeStr += " x " + item.Csize;

                        if (!string.IsNullOrEmpty(sizeStr))
                        {
                            if (!string.IsNullOrEmpty(techInfo)) techInfo += " | ";
                            techInfo += sizeStr;
                        }

                        string itemHeader = item.ItemName;
                        if (!string.IsNullOrEmpty(techInfo)) itemHeader += $" ({techInfo})";

                        detailsMsg += $"\n- {itemHeader}: {item.TotalQty:G29} {item.Unit}";
                    }

                    if (string.IsNullOrWhiteSpace(detailsMsg))
                    {
                        ShowToast("⚠ Vui lòng cấu hình số lượng giao hàng lớn hơn 0 cho ít nhất 1 hạng mục!", true);
                        return;
                    }

                    // 3. Tự động chuyển trạng thái thành "Giao hàng lần 2" nếu % giao hàng thuộc (0%, 100%)
                    decimal pctGiaoVal = 0;
                    string pctClean = pctGiao.Replace("%", "").Trim();
                    decimal.TryParse(pctClean, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out pctGiaoVal);

                    string dynamicTrangThai = trangThai;
                    if (pctGiaoVal > 0 && pctGiaoVal < 100)
                    {
                        dynamicTrangThai = "Giao hàng lần 2";
                    }
                    else if (pctGiaoVal == 0)
                    {
                        dynamicTrangThai = "Giao hàng lần 1";
                    }

                    string msg = $"📦 THÔNG BÁO GIAO HÀNG\n" +
                                 $"━━━━━━━━━━━━━━━━━━\n" +
                                 $"📋 PO: {poNo}\n" +
                                 $"🏗 Dự án: {duan}\n" +
                                 $"📊 Trạng thái: {dynamicTrangThai}\n" +
                                 $"📅 Ngày giao: {configForm.SelectedDate:dd/MM/yyyy}\n" +
                                 $"⚡ Cảnh báo: {canhBao}\n" +
                                 $"📦 Chi tiết:\n{detailsMsg}\n" +
                                 $"━━━━━━━━━━━━━━━━━━";

                    if (Helpers.ZaloHelper.IsConfigured())
                    {
                        var settings = Helpers.ZaloHelper.LoadSettings();
                        // Gửi đến nhóm Zalo của dự án (nếu không có thì gửi nhóm mặc định "Giao hàng")
                        string targetGroup = !string.IsNullOrWhiteSpace(zaloGroupName) ? zaloGroupName : "Giao hàng";

                        var result = await Helpers.ZaloHelper.SendToGroupAsync(settings, targetGroup, msg);
                        if (result.ok)
                        {
                            ShowToast($"✅ Đã gửi thông báo Zalo đến nhóm {targetGroup} cho PO: {poNo}");
                            _sentZaloPOs.Add(poId);
                            SaveZaloSentStatus(poId);

                            // Ghi log vào lịch sử
                            try
                            {
                                new NotificationLogService().AddLog(new NotificationLog
                                {
                                    Sent_At = DateTime.Now,
                                    Sent_By = AppSession.CurrentUser?.Full_Name ?? "System",
                                    Recipient = targetGroup,
                                    Type = "Zalo",
                                    Content = msg,
                                    Status = "Success",
                                    Project_Code = duan
                                });
                            }
                            catch { }
                        }
                        else
                        {
                            ShowToast($"⚠ Lỗi gửi Zalo: {result.error}", true);
                        }
                    }
                    else
                    {
                        Clipboard.SetText(msg);
                        ShowToast($"📋 Đã copy thông báo PO: {poNo} vào clipboard (Zalo chưa cấu hình)");
                        _sentZaloPOs.Add(poId);
                        SaveZaloSentStatus(poId);
                    }

                    dgvPO.InvalidateRow(e.RowIndex);
                }
            }
            catch (Exception ex)
            {
                try { Clipboard.SetText("Lỗi xử lý thông báo giao hàng"); } catch { }
                ShowToast($"⚠ Lỗi: {ex.Message}", true);
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

        // ── Toast: thông báo không modal, tự ẩn sau 3 giây ──────────────────────
        private async void ShowToast(string message, bool isError = false)
        {
            if (_toastPanel != null && !_toastPanel.IsDisposed)
            {
                this.Controls.Remove(_toastPanel);
                _toastPanel.Dispose();
            }
            _toastPanel = new Panel
            {
                BackColor = isError ? Color.FromArgb(200, 53, 69) : Color.FromArgb(40, 167, 69),
                Size      = new Size(this.ClientSize.Width, 36),
                Location  = new Point(0, 0)
            };
            var lbl = new Label
            {
                Text      = "  " + message,
                ForeColor = Color.White,
                Font      = new Font("Segoe UI", 10, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft,
                Dock      = DockStyle.Fill,
                AutoSize  = false
            };
            _toastPanel.Controls.Add(lbl);
            this.Controls.Add(_toastPanel);
            _toastPanel.BringToFront();
            await Task.Delay(3000);
            if (_toastPanel != null && !_toastPanel.IsDisposed)
            {
                this.Controls.Remove(_toastPanel);
                _toastPanel.Dispose();
                _toastPanel = null;
            }
        }

        private Form TopOwner => (this.TopLevelControl as Form) ?? this;
        private void SafeMsg(string text, string title, MessageBoxIcon icon = MessageBoxIcon.Error)
        { var f = TopOwner; f.BringToFront(); f.Activate(); MessageBox.Show(f, text, title, MessageBoxButtons.OK, icon); }
        private void SafeInfo(string text, string title = "Thông báo")
        { var f = TopOwner; f.BringToFront(); f.Activate(); MessageBox.Show(f, text, title, MessageBoxButtons.OK, MessageBoxIcon.Information); }
        private void SafeWarn(string text, string title = "Thông báo")
        { var f = TopOwner; f.BringToFront(); f.Activate(); MessageBox.Show(f, text, title, MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        private bool SafeAsk(string text, string title = "Xác nhận")
        { var f = TopOwner; f.BringToFront(); f.Activate(); return MessageBox.Show(f, text, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes; }


        private async Task LoadMPRDataAsync()
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

                // Adjusted to consider all revisions in the same MPR series (baseNo + _Rev.x)
                string sql = $@"
                    SELECT
                        h.MPR_ID,
                        h.MPR_No                           AS [MPR No],
                        h.Project_Name                     AS [Dự án],
                        h.Required_Date                    AS [Ngày cần],
                        h.Status                           AS [Trạng thái],
                        h.Rev                              AS [Rev],

                        CASE
                            WHEN COUNT(DISTINCT po.PO_ID) > 0
                            THEN N'✅ ' + CAST(COUNT(DISTINCT po.PO_ID) AS NVARCHAR(10)) + N' PO'
                            ELSE N'❌ Chưa có PO'
                        END                                AS [Tình trạng PO],

                        -- % Item đặt hàng: ordered items across the whole MPR series / total items in the series
                        CASE
                            WHEN COUNT(DISTINCT series_d.Detail_ID) = 0 THEN 0
                            ELSE CAST(
                                COUNT(DISTINCT pod.PO_Detail_ID) * 100.0
                                / COUNT(DISTINCT series_d.Detail_ID)
                                AS DECIMAL(5,1))
                        END                                AS [% Item đặt hàng],

                        h.Created_Date                     AS [Ngày tạo]
                    FROM MPR_Header h
                    -- compute baseNo: strip suffix _Rev.x if present
                    CROSS APPLY (SELECT CASE WHEN CHARINDEX('_Rev.', h.MPR_No) > 0 THEN LEFT(h.MPR_No, CHARINDEX('_Rev.', h.MPR_No)-1) ELSE h.MPR_No END AS BaseNo) bn
                    -- include all headers in same series (baseNo and its revisions)
                    LEFT JOIN MPR_Header series_h ON (series_h.MPR_No = h.MPR_No OR series_h.MPR_No = bn.BaseNo OR series_h.MPR_No LIKE bn.BaseNo + '_Rev.%')
                    LEFT JOIN MPR_Details series_d ON series_d.MPR_ID = series_h.MPR_ID
                    LEFT JOIN PO_Detail pod ON pod.MPR_Detail_ID = series_d.Detail_ID
                    -- consider PO linked either via PO_Detail or via PO_head.MPR_No referencing any name in series
                    LEFT JOIN PO_head po ON po.PO_ID = pod.PO_ID
                                          OR po.MPR_No = h.MPR_No
                                          OR po.MPR_No = bn.BaseNo
                                          OR po.MPR_No LIKE bn.BaseNo + '_Rev.%'
                    {where}
                    GROUP BY h.MPR_ID, h.MPR_No, h.Project_Name,
                             h.Required_Date, h.Status, h.Rev, h.Created_Date, h.Notes
                    ORDER BY h.Created_Date DESC";
                // Chạy toàn bộ DB I/O trên thread pool — SP + Notes trong 1 connection
                var (dt, noteMap) = await Task.Run(() =>
                {
                    using var conn = DatabaseHelper.GetConnection();
                    conn.Open();

                    using var cmd = new SqlCommand("sp_GetMPRDashboardSummary", conn)
                        { CommandType = System.Data.CommandType.StoredProcedure };
                    cmd.Parameters.AddWithValue("@search", string.IsNullOrEmpty(search) ? (object)DBNull.Value : search);
                    cmd.Parameters.AddWithValue("@status", (filter == "Tất cả") ? (object)DBNull.Value : filter);
                    var table = new DataTable();
                    table.Load(cmd.ExecuteReader());

                    // Lọc bỏ revision cũ (Rev < MaxRev)
                    if (table.Columns.Contains("MaxRev"))
                    {
                        var oldRevRows = table.AsEnumerable()
                            .Where(r =>
                            {
                                int.TryParse(r["Rev"]?.ToString() ?? "0", out int rev);
                                int.TryParse(r["MaxRev"]?.ToString() ?? "0", out int maxRev);
                                return maxRev > 0 && rev < maxRev;
                            })
                            .ToList();
                        foreach (var r in oldRevRows)
                            table.Rows.Remove(r);
                    }

                    // Load Notes trong cùng connection, không cần mở thêm
                    var notes = new System.Collections.Generic.Dictionary<int, string>();
                    try
                    {
                        using var cmdNote = new SqlCommand(
                            "SELECT MPR_ID, ISNULL(Notes,'') AS Notes FROM MPR_Header", conn);
                        using var rNote = cmdNote.ExecuteReader();
                        while (rNote.Read())
                            notes[Convert.ToInt32(rNote["MPR_ID"])] = rNote["Notes"].ToString();
                    }
                    catch { }

                    return (table, notes);
                });

                // *** ISSUE #2 FIX: Suspend layout & binding during grid update ***
                dgvMPR.SuspendLayout();
                try
                {
                    // Temporarily detach event handlers
                    dgvMPR.SelectionChanged -= DgvMPR_SelectionChanged;
                    dgvMPR.CellFormatting -= DgvMPR_CellFormatting;
                    dgvMPR.RowPrePaint -= DgvMPR_RowPrePaint;

                    // Cập nhật UI (đang trên UI thread sau await)
                    foreach (DataColumn col in dt.Columns) col.ReadOnly = false;
                    dgvMPR.DataSource = dt;

                if (dgvMPR.Columns.Contains("MPR_ID"))
                    dgvMPR.Columns["MPR_ID"].Visible = false;
                if (dgvMPR.Columns.Contains("Tổng items")) dgvMPR.Columns["Tổng items"].Visible = false;
                if (dgvMPR.Columns.Contains("Ngày đến PO")) dgvMPR.Columns["Ngày đến PO"].Visible = false;

                foreach (DataGridViewColumn col in dgvMPR.Columns)
                    col.ReadOnly = true;
                if (dgvMPR.Columns.Contains("Ghi chu")) dgvMPR.Columns["Ghi chu"].ReadOnly = false;

                // Style theo MaxRev / Hủy
                try
                {
                    var grayStyle = new DataGridViewCellStyle
                    {
                        ForeColor = Color.FromArgb(160, 160, 160),
                        BackColor = Color.FromArgb(245, 245, 245),
                        Font = new Font("Segoe UI", 9, FontStyle.Italic)
                    };
                    var normalStyle = new DataGridViewCellStyle
                    {
                        ForeColor = Color.Black,
                        BackColor = Color.White,
                        Font = new Font("Segoe UI", 9, FontStyle.Regular)
                    };
                    var cancelStyle = new DataGridViewCellStyle(grayStyle)
                    {
                        Font = new Font("Segoe UI", 9, FontStyle.Strikeout | FontStyle.Bold)
                    };

                    foreach (DataGridViewRow row in dgvMPR.Rows)
                    {
                        if (row.IsNewRow) continue;
                        int.TryParse(row.Cells["Rev"].Value?.ToString() ?? "0", out int rev);
                        int.TryParse(row.Cells["MaxRev"]?.Value?.ToString() ?? "0", out int maxRev);
                        string status = row.Cells["Trạng thái"]?.Value?.ToString() ?? "";
                        bool isCancelled = status == "Hủy";

                        if ((maxRev > 0 && rev < maxRev) || isCancelled)
                        {
                            row.ReadOnly = true;
                            row.DefaultCellStyle = isCancelled ? cancelStyle : grayStyle;
                        }
                        else
                        {
                            row.ReadOnly = false;
                            row.DefaultCellStyle = normalStyle;
                        }
                    }
                }
                catch { }

                // Cột Ghi chú
                if (!dgvMPR.Columns.Contains("Ghi chu"))
                {
                    var colNote = new DataGridViewTextBoxColumn
                    {
                        Name = "Ghi chu",
                        HeaderText = "Ghi chu",
                        Width = 160,
                        ReadOnly = false,
                        DisplayIndex = dgvMPR.Columns.Count
                    };
                    colNote.DefaultCellStyle.BackColor = Color.FromArgb(255, 255, 230);
                    colNote.DefaultCellStyle.SelectionBackColor = Color.FromArgb(255, 245, 180);
                    colNote.ToolTipText = "Click de nhap ghi chu, bam Luu ghi chu de luu";
                    dgvMPR.Columns.Add(colNote);
                }
                else
                {
                    dgvMPR.Columns["Ghi chu"].ReadOnly = false;
                }

                // Điền Notes vào unbound column
                foreach (DataGridViewRow row in dgvMPR.Rows)
                {
                    if (row.IsNewRow) continue;
                    object idObj = row.Cells["MPR_ID"]?.Value;
                    if (idObj == null || idObj == DBNull.Value) continue;
                    if (noteMap.TryGetValue(Convert.ToInt32(idObj), out string note))
                        row.Cells["Ghi chu"].Value = note;
                }

                    AutoAdjustMPRColumns();

                    // Re-attach event handlers
                    dgvMPR.CellFormatting += DgvMPR_CellFormatting;
                    dgvMPR.SelectionChanged += DgvMPR_SelectionChanged;
                    dgvMPR.RowPrePaint += DgvMPR_RowPrePaint;
                    dgvMPR.Invalidate(); // Force redraw
                }
                finally
                {
                    dgvMPR.ResumeLayout(true); // Resume layout with force layout = true
                }

                int total = dt.Rows.Count, hasPO = 0, noPO = 0, completed = 0;
                bool hasTinhCol = dt.Columns.Contains("Tình trạng PO");
                bool hasStatusCol = dt.Columns.Contains("Trạng thái");
                foreach (DataRow row in dt.Rows)
                {
                    string tinh = hasTinhCol ? row["Tình trạng PO"]?.ToString() ?? "" : "";
                    string status = hasStatusCol ? row["Trạng thái"]?.ToString() ?? "" : "";
                    if (!tinh.Contains("Chưa có")) hasPO++;
                    else noPO++;
                    if (status == "Hoàn thành") completed++;
                }
                lblMPRTotal.Text = total.ToString();
                lblMPRHasPO.Text = hasPO.ToString();
                lblMPRNoPO.Text = noPO.ToString();
                lblMPRCompleted.Text = completed.ToString();
            }
            catch (Exception ex)
            {
                SafeMsg("Lỗi tải MPR: " + ex.Message, "Lỗi");
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
            if (!dgvMPR.Columns.Contains("Tình trạng PO") || !dgvMPR.Columns.Contains("Trạng thái")) return;
            var row = dgvMPR.Rows[e.RowIndex];
            string tinh = row.Cells["Tình trạng PO"].Value?.ToString() ?? "";
            string status = row.Cells["Trạng thái"].Value?.ToString() ?? "";
            if (status == "Hoàn thành") row.DefaultCellStyle.BackColor = Color.FromArgb(235, 255, 235);
            else if (tinh.Contains("Chưa có")) row.DefaultCellStyle.BackColor = Color.FromArgb(255, 235, 235);
            else row.DefaultCellStyle.BackColor = Color.FromArgb(255, 248, 235);
        }

        // Auto giãn cột dgvMPR: min 30, max 150 — cột "Dự án" đã set 60 trước khi gọi hàm này
        // Lọc dgvMPR theo Tình trạng PO (client-side, không query lại DB)
        //private void FilterMPRByPOStatus()
        //{
        //    if (dgvMPR == null || dgvMPR.Rows.Count == 0) return;
        //    string sel = cboFilterPOStatus.SelectedItem?.ToString() ?? "Tất cả";

        //    foreach (DataGridViewRow row in dgvMPR.Rows)
        //    {
        //        if (row.IsNewRow) continue;
        //        if (sel == "Tất cả") { row.Visible = true; continue; }

        //        // Đọc % Item đặt hàng từ cột (có thể dạng "100.0%" hoặc số)
        //        string pctRaw = row.Cells["% Item đặt hàng"].Value?.ToString() ?? "0";
        //        pctRaw = pctRaw.Replace("%", "").Trim();
        //        decimal.TryParse(pctRaw, System.Globalization.NumberStyles.Any,
        //            System.Globalization.CultureInfo.InvariantCulture, out decimal pct);

        //        if (sel.Contains("Hoàn thành"))
        //            row.Visible = pct >= 100;
        //        else if (sel.Contains("Chưa hoàn thành"))
        //            row.Visible = pct < 100;
        //        else
        //            row.Visible = true;
        //    }
        //}
        private void FilterMPRByPOStatus()
        {
            // Kiểm tra điều kiện đầu vào
            if (dgvMPR == null || dgvMPR.Rows.Count == 0) return;

            // 1. QUAN TRỌNG: Hủy chọn dòng hiện tại để tránh lỗi InvalidOperationException
            dgvMPR.CurrentCell = null;

            string sel = cboFilterPOStatus.SelectedItem?.ToString() ?? "Tất cả";

            // Sử dụng CurrencyManager để tạm dừng quản lý vị trí dòng, giúp ẩn dòng mượt hơn
            CurrencyManager currencyManager = (CurrencyManager)BindingContext[dgvMPR.DataSource];
            currencyManager.SuspendBinding();

            try
            {
                foreach (DataGridViewRow row in dgvMPR.Rows)
                {
                    if (row.IsNewRow) continue;

                    if (sel == "Tất cả")
                    {
                        row.Visible = true;
                        continue;
                    }

                    // Đọc % Item đặt hàng (Xử lý an toàn với CultureInfo.InvariantCulture như chúng ta đã làm)
                    string pctRaw = row.Cells["% Item đặt hàng"].Value?.ToString() ?? "0";
                    pctRaw = pctRaw.Replace("%", "").Trim();

                    decimal.TryParse(pctRaw, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out decimal pct);

                    // Thực hiện ẩn/hiện dựa trên điều kiện
                    if (sel.Contains("Hoàn thành"))
                    {
                        row.Visible = (pct >= 100);
                    }
                    else if (sel.Contains("Chưa hoàn thành"))
                    {
                        row.Visible = (pct < 100);
                    }
                    else
                    {
                        row.Visible = true;
                    }
                }
            }
            catch (Exception ex)
            {
                // Debug nếu có lỗi phát sinh trong quá trình lọc
                Console.WriteLine("Lỗi lọc MPR: " + ex.Message);
            }
            finally
            {
                // 2. QUAN TRỌNG: Kích hoạt lại Binding sau khi lọc xong
                currencyManager.ResumeBinding();
            }
        }

        // Xuất Excel tổng hợp MPR + PO
        private async void BtnExportMPR_Click(object sender, EventArgs e)
        {
            if (dgvMPR == null || dgvMPR.Rows.Count == 0)
            { ShowToast("Không có dữ liệu để xuất!", isError: true); return; }

            // ── Lấy danh sách MPR No đang HIỂN THỊ (UI thread) ──
            var mprNos = new System.Collections.Generic.List<string>();
            foreach (DataGridViewRow row in dgvMPR.Rows)
            {
                if (row.IsNewRow || !row.Visible) continue;

                string status = row.Cells["Trạng thái"]?.Value?.ToString() ?? "";
                if (status == "Hủy") continue;

                string mno = row.Cells["MPR No"].Value?.ToString();
                if (!string.IsNullOrEmpty(mno)) mprNos.Add(mno);
            }
            if (mprNos.Count == 0) { ShowToast("Không có MPR nào!", isError: true); return; }

            using var sfd = new SaveFileDialog
            {
                Title        = "Lưu báo cáo MPR",
                Filter       = "Excel|*.xlsx",
                FileName     = $"BaoCao_MPR_{DateTime.Now:yyyyMMdd_HHmm}",
                InitialDirectory = System.IO.Directory.Exists(@"D:\RÁC") ? @"D:\RÁC" : Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };
            if (sfd.ShowDialog() != DialogResult.OK) return;

            string savePath = sfd.FileName;
            btnExportMPR.Enabled = false;

            try
            {
                string inClause = string.Join(",", mprNos.Select(m => $"N'{m.Replace("'", "''")}'"));

                // ── Query với CTE — cross-revision PO/RIR ──
                // Sib_Details tính 1 lần toàn bộ cặp (CurrDetailID, SibDetailID) cho tất cả revision
                // PO_Flat / RIR_Flat là DISTINCT nhỏ gọn — FOR XML PATH cuối chỉ duyệt set nhỏ
                string sql = @"
                    WITH
                    Qry_MPR AS (
                        SELECT MPR_ID, MPR_No,
                            CASE WHEN CHARINDEX('_Rev.',MPR_No)>0
                                 THEN LEFT(MPR_No,CHARINDEX('_Rev.',MPR_No)-1)
                                 ELSE MPR_No END AS BaseNo
                        FROM MPR_Header WHERE MPR_No IN (" + inClause + @")
                    ),
                    Sib_Details AS (
                        SELECT DISTINCT dC.Detail_ID AS CurrID, dS.Detail_ID AS SibID
                        FROM Qry_MPR q
                        INNER JOIN MPR_Details dC ON dC.MPR_ID = q.MPR_ID
                            AND TRY_CAST(TRY_CAST(dC.Item_No AS DECIMAL(10,2)) AS INT) > 0
                        INNER JOIN MPR_Header hS
                            ON (CASE WHEN CHARINDEX('_Rev.',hS.MPR_No)>0
                                     THEN LEFT(hS.MPR_No,CHARINDEX('_Rev.',hS.MPR_No)-1)
                                     ELSE hS.MPR_No END) = q.BaseNo
                        INNER JOIN MPR_Details dS ON dS.MPR_ID = hS.MPR_ID
                            AND TRY_CAST(TRY_CAST(dS.Item_No AS DECIMAL(10,2)) AS INT)
                              = TRY_CAST(TRY_CAST(dC.Item_No AS DECIMAL(10,2)) AS INT)
                    ),
                    PO_Flat AS (
                        SELECT DISTINCT sd.CurrID, pox.PONo
                        FROM Sib_Details sd
                        INNER JOIN PO_Detail podx ON podx.MPR_Detail_ID = sd.SibID
                        INNER JOIN PO_head   pox  ON pox.PO_ID = podx.PO_ID
                        WHERE ISNULL(pox.Status,'') <> 'Cancelled'
                    ),
                    RIR_Flat AS (
                        SELECT DISTINCT sd.CurrID, r.RIR_No
                        FROM Sib_Details sd
                        INNER JOIN PO_Detail podx ON podx.MPR_Detail_ID = sd.SibID
                        INNER JOIN PO_head   pox  ON pox.PO_ID = podx.PO_ID
                        INNER JOIN RIR_head  r    ON r.PONo = pox.PONo
                        WHERE ISNULL(pox.Status,'') <> 'Cancelled'
                    )
                    SELECT
                        h.MPR_No,
                        h.Project_Name,
                        h.Status                                    AS MPR_Status,
                        CONVERT(NVARCHAR(10), h.Required_Date, 103) AS Required_Date,
                        ISNULL(h.Notes, '')                         AS MPR_Notes,
                        ISNULL(TRY_CAST(TRY_CAST(d.Item_No AS DECIMAL(10,2)) AS INT), 0) AS Item_No,
                        ISNULL(d.item_name,     '')  AS Item_Name,
                        ISNULL(d.Description,   '')  AS Description,
                        ISNULL(d.Material,      '')  AS Material,
                        ISNULL(CAST(NULLIF(TRY_CAST(d.Thickness_mm AS DECIMAL(18,4)),0) AS NVARCHAR),'') AS A_Day,
                        ISNULL(CAST(NULLIF(TRY_CAST(d.Depth_mm     AS DECIMAL(18,4)),0) AS NVARCHAR),'') AS B_Sau,
                        ISNULL(CAST(NULLIF(TRY_CAST(d.C_Width_mm   AS DECIMAL(18,4)),0) AS NVARCHAR),'') AS C_Rong,
                        ISNULL(CAST(NULLIF(TRY_CAST(d.D_Web_mm     AS DECIMAL(18,4)),0) AS NVARCHAR),'') AS D_Bung,
                        ISNULL(CAST(NULLIF(TRY_CAST(d.E_Flange_mm  AS DECIMAL(18,4)),0) AS NVARCHAR),'') AS E_Canh,
                        ISNULL(CAST(NULLIF(TRY_CAST(d.F_Length_mm  AS DECIMAL(18,4)),0) AS NVARCHAR),'') AS F_Dai,
                        ISNULL(d.UNIT,          '')  AS UNIT,
                        ISNULL(d.Qty_Per_Sheet, 0)   AS SL,
                        ISNULL(d.Weight_kg,     0)   AS KG,
                        ISNULL(d.MPS_Info,     '')   AS MPS_Info,
                        ISNULL(d.Usage_Location,'')  AS Usage_Location,
                        ISNULL(d.REV,          '0')  AS REV,
                        ISNULL(d.Remarks,      '')   AS Detail_Remarks,
                        ISNULL(STUFF((
                            SELECT ', ' + pf.PONo FROM PO_Flat pf
                            WHERE pf.CurrID = d.Detail_ID
                            FOR XML PATH(''), TYPE
                        ).value('.','NVARCHAR(MAX)'), 1, 2, ''), '') AS PO_List,
                        ISNULL(STUFF((
                            SELECT ', ' + rf.RIR_No FROM RIR_Flat rf
                            WHERE rf.CurrID = d.Detail_ID
                            FOR XML PATH(''), TYPE
                        ).value('.','NVARCHAR(MAX)'), 1, 2, ''), '') AS RIR_List
                    FROM MPR_Header  h
                    INNER JOIN MPR_Details d ON d.MPR_ID = h.MPR_ID
                    WHERE h.MPR_No IN (" + inClause + @")
                      AND ISNULL(d.Is_Deleted, 0) = 0
                      AND ISNULL(TRY_CAST(TRY_CAST(h.Rev AS DECIMAL(10,2)) AS INT),0) = (
                          SELECT ISNULL(MAX(TRY_CAST(TRY_CAST(h2.Rev AS DECIMAL(10,2)) AS INT)),0)
                          FROM MPR_Header h2
                          WHERE h2.MPR_No IN (" + inClause + @")
                            AND (CASE WHEN CHARINDEX('_Rev.',h2.MPR_No)>0 THEN LEFT(h2.MPR_No,CHARINDEX('_Rev.',h2.MPR_No)-1) ELSE h2.MPR_No END)
                                = (CASE WHEN CHARINDEX('_Rev.',h.MPR_No)>0  THEN LEFT(h.MPR_No, CHARINDEX('_Rev.',h.MPR_No)-1)  ELSE h.MPR_No  END)
                      )
                    ORDER BY h.MPR_No,
                             ISNULL(TRY_CAST(TRY_CAST(d.Item_No AS DECIMAL(10,2)) AS INT), 0)";

                DataTable dt = await Task.Run(() =>
                {
                    using var conn = DatabaseHelper.GetConnection();
                    conn.Open();
                    var table = new DataTable();
                    table.Load(new SqlCommand(sql, conn) { CommandTimeout = 120 }.ExecuteReader());
                    return table;
                });

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using var pkg = new ExcelPackage();
                var ws = pkg.Workbook.Worksheets.Add("Chi tiết MPR");

                int TOTAL_COLS = 16;
                ws.Cells[1, 1].Value = "BÁO CÁO CHI TIẾT ĐẶT HÀNG MPR";
                ws.Cells[1, 1, 1, TOTAL_COLS].Merge = true;
                ws.Cells[1, 1].Style.Font.Size = 14;
                ws.Cells[1, 1].Style.Font.Bold = true;
                ws.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[2, 1].Value = $"Xuất ngày: {DateTime.Now:dd/MM/yyyy HH:mm}";
                ws.Cells[2, 1, 2, TOTAL_COLS].Merge = true;

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

                Func<object, object> dimVal = v => {
                    if (v == null || v == DBNull.Value) return (object)"";
                    var s = v.ToString(); if (string.IsNullOrEmpty(s)) return (object)"";
                    return decimal.TryParse(s, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out decimal d) && d != 0
                        ? (object)Math.Round(d, 2, MidpointRounding.AwayFromZero)
                            .ToString("#,##0.##", System.Globalization.CultureInfo.InvariantCulture)
                        : (object)"";
                };

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
                    ws.Cells[rowIdx, 10].Value = dimVal(dr["A_Day"]);
                    ws.Cells[rowIdx, 11].Value = dimVal(dr["B_Sau"]);
                    ws.Cells[rowIdx, 12].Value = dimVal(dr["C_Rong"]);
                    ws.Cells[rowIdx, 13].Value = dimVal(dr["D_Bung"]);
                    ws.Cells[rowIdx, 14].Value = dimVal(dr["E_Canh"]);
                    ws.Cells[rowIdx, 15].Value = dimVal(dr["F_Dai"]);
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

                // Căn chỉnh phải cột kích thước A-F (cell-level vì string values cần ghi đè column style)
                if (dt.Rows.Count > 0)
                    foreach (int c in new[] { 10, 11, 12, 13, 14, 15 })
                        if (c <= TOTAL_COLS)
                            ws.Cells[5, c, rowIdx - 1, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                // Căn chỉnh cột số (STT, kích thước, SL, KG)
                foreach (int c in new[] { 6, 10, 11, 12, 13, 14, 15, 17, 18 })
                    if (c <= TOTAL_COLS)
                        ws.Column(c).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                // Độ rộng cột cố định — thay AutoFitColumns() để tránh chậm với dữ liệu lớn
                double[] colWidths = {
                    18,  // 1  MPR No
                    28,  // 2  Dự án
                    14,  // 3  TT MPR
                    12,  // 4  Ngày cần
                    22,  // 5  Ghi chú MPR
                    6,   // 6  STT
                    32,  // 7  Tên vật tư
                    28,  // 8  Mô tả
                    14,  // 9  Vật liệu
                    10,  // 10 A-Dày
                    10,  // 11 B-Sâu
                    10,  // 12 C-Rộng
                    10,  // 13 D-Bụng
                    10,  // 14 E-Cánh
                    12,  // 15 F-Dài
                    8,   // 16 ĐVT
                    10,  // 17 Số lượng
                    10,  // 18 KG
                    16,  // 19 MPS Info
                    16,  // 20 Nơi dùng
                    6,   // 21 REV
                    22,  // 22 Ghi chú
                    35,  // 23 Số PO
                    22,  // 24 Số RIR
                };
                for (int ci = 0; ci < colWidths.Length && ci < TOTAL_COLS; ci++)
                    ws.Column(ci + 1).Width = colWidths[ci];

                ws.View.FreezePanes(5, 1);

                await Task.Run(() => pkg.SaveAs(new FileInfo(savePath)));

                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        { FileName = savePath, UseShellExecute = true });
                }
                catch { }
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                this.BeginInvoke(new Action(() => SafeMsg("Lỗi xuất Excel: " + msg, "Lỗi")));
            }
            finally
            {
                btnExportMPR.Enabled = true;
            }
        }

        private async Task LoadRIRDataAsync()
        {
            try
            {
                string search = txtSearchRIR.Text.Trim();
                string filter = cboFilterRIR.SelectedItem?.ToString() ?? "Tất cả";

                string where = "WHERE 1=1";
                if (!string.IsNullOrEmpty(search))
                    where += $" AND (h.RIR_No LIKE N'%{search}%' OR h.PONo LIKE N'%{search}%' OR h.Project_Name LIKE N'%{search}%' OR pi.ProjectCode LIKE N'%{search}%')";
                if (filter != "Tất cả")
                    where += $" AND h.Status = N'{filter}'";

                string sql = $@"
                    SELECT
                        h.RIR_ID,
                        h.RIR_No                                            AS [RIR No],
                        h.PONo                                              AS [PO No],
                        h.MPR_No                                            AS [MPR No],
                        ISNULL(pi.ProjectCode, h.Project_Name)             AS [Mã dự án],
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
                    LEFT JOIN PO_head po ON po.PONo = h.PONo
                    LEFT JOIN ProjectInfo pi ON pi.ProjectCode = po.ProjectCode
                    {where}
                    GROUP BY h.RIR_ID, h.RIR_No, h.PONo, h.MPR_No, ISNULL(pi.ProjectCode, h.Project_Name),
                             h.Issue_Date, h.Customer, h.Status
                    ORDER BY h.Issue_Date DESC";

                // Chạy DB I/O trên thread pool
                var dt = await Task.Run(() =>
                {
                    using var conn = DatabaseHelper.GetConnection();
                    conn.Open();
                    var table = new DataTable();
                    table.Load(new SqlCommand(sql, conn).ExecuteReader());
                    return table;
                });

                // *** ISSUE #2 FIX: Suspend layout & binding during grid update ***
                dgvRIR.SuspendLayout();
                try
                {
                    // Temporarily detach event handlers
                    dgvRIR.SelectionChanged -= DgvRIR_SelectionChanged;
                    dgvRIR.CellFormatting -= DgvRIR_CellFormatting;
                    dgvRIR.RowPrePaint -= DgvRIR_RowPrePaint;

                    // Cập nhật UI (đang trên UI thread sau await)
                    dgvRIR.DataSource = dt;
                    if (dgvRIR.Columns.Contains("RIR_ID"))
                        dgvRIR.Columns["RIR_ID"].Visible = false;

                    // Re-attach event handlers
                    dgvRIR.CellFormatting += DgvRIR_CellFormatting;
                    dgvRIR.SelectionChanged += DgvRIR_SelectionChanged;
                    dgvRIR.RowPrePaint += DgvRIR_RowPrePaint;
                    dgvRIR.Invalidate(); // Force redraw
                }
                finally
                {
                    dgvRIR.ResumeLayout(true); // Resume layout with force layout = true
                }

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
            catch (Exception ex)
            {
                SafeMsg("Lỗi tải RIR: " + ex.Message, "Lỗi");
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
                SafeMsg("Lỗi tải chi tiết RIR: " + ex.Message, "Lỗi");
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
        // =====================================================================
        //  NOTIFICATION SYSTEM
        // =====================================================================
        private void BuildNotificationPanel()
        {
            btnNotifyToggle = new Button
            {
                Text = "N",
                Size = new Size(36, 28),
                BackColor = Color.FromArgb(0, 90, 170),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            btnNotifyToggle.FlatAppearance.BorderSize = 0;
            btnNotifyToggle.Click += (s, e) =>
            {
                panelNotify.Visible = !panelNotify.Visible;
                if (panelNotify.Visible)
                {
                    panelNotify.BringToFront();
                    panelNotify.Location = new Point(
                        this.ClientSize.Width - panelNotify.Width - 10,
                        this.ClientSize.Height - panelNotify.Height - 10);
                    btnNotifyToggle.BackColor = Color.FromArgb(0, 90, 170);
                }
            };

            var panelHeader = this.Controls.OfType<Panel>()
                .FirstOrDefault(p => p.BackColor == Color.FromArgb(0, 120, 212));
            if (panelHeader != null)
            {
                btnNotifyToggle.Location = new Point(panelHeader.Width - 195, 8);
                panelHeader.Controls.Add(btnNotifyToggle);
            }

            panelNotify = new Panel
            {
                Size = new Size(340, 420),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Visible = false
            };
            panelNotify.Location = new Point(
                this.ClientSize.Width - 350, this.ClientSize.Height - 430);

            // Header chatbox
            var pHead = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(340, 40),
                BackColor = Color.FromArgb(0, 120, 212)
            };
            lblNotifyTitle = new Label
            {
                Text = "Thong bao he thong",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(10, 10),
                Size = new Size(240, 22)
            };
            var btnClose = new Button
            {
                Text = "X",
                Size = new Size(28, 28),
                Location = new Point(308, 6),
                BackColor = Color.FromArgb(0, 90, 170),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Click += (s, e) => panelNotify.Visible = false;
            pHead.Controls.Add(lblNotifyTitle);
            pHead.Controls.Add(btnClose);
            panelNotify.Controls.Add(pHead);

            lblNotifyCount = new Label
            {
                Text = "Chua co thong bao moi",
                Font = new Font("Segoe UI", 8, FontStyle.Italic),
                ForeColor = Color.Gray,
                Location = new Point(10, 46),
                Size = new Size(320, 18)
            };
            panelNotify.Controls.Add(lblNotifyCount);

            var btnRefreshNow = new Button
            {
                Text = "Lam moi ngay",
                Size = new Size(130, 26),
                Location = new Point(10, 68),
                BackColor = Color.FromArgb(0, 150, 100),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnRefreshNow.FlatAppearance.BorderSize = 0;
            btnRefreshNow.Click += (s, e) => CheckAndNotify(true);
            panelNotify.Controls.Add(btnRefreshNow);

            var btnClear = new Button
            {
                Text = "Xoa tat ca",
                Size = new Size(110, 26),
                Location = new Point(148, 68),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnClear.FlatAppearance.BorderSize = 0;
            btnClear.Click += (s, e) =>
            {
                lstNotify.Items.Clear();
                lblNotifyCount.Text = "Da xoa tat ca thong bao";
                btnNotifyToggle.BackColor = Color.FromArgb(0, 90, 170);
            };
            panelNotify.Controls.Add(btnClear);

            lstNotify = new ListBox
            {
                Location = new Point(8, 100),
                Size = new Size(322, 278),
                Font = new Font("Segoe UI", 9),
                BorderStyle = BorderStyle.None,
                BackColor = Color.FromArgb(248, 248, 252),
                ItemHeight = 44,
                DrawMode = DrawMode.OwnerDrawFixed,
                IntegralHeight = false
            };
            lstNotify.DrawItem += LstNotify_DrawItem;
            panelNotify.Controls.Add(lstNotify);

            var lblNext = new Label
            {
                Name = "lblNextRefresh",
                Text = "Tu dong cap nhat moi 5 phut",
                Font = new Font("Segoe UI", 7, FontStyle.Italic),
                ForeColor = Color.Silver,
                Location = new Point(10, 384),
                Size = new Size(320, 16)
            };
            panelNotify.Controls.Add(lblNext);

            this.Controls.Add(panelNotify);
            panelNotify.BringToFront();
        }

        private void LstNotify_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            string msg = lstNotify.Items[e.Index].ToString();
            bool isPO = msg.StartsWith("[PO]");
            bool isMPR = msg.StartsWith("[MPR]");

            Color bg = e.Index % 2 == 0 ? Color.White : Color.FromArgb(245, 245, 252);
            e.Graphics.FillRectangle(new SolidBrush(bg), e.Bounds);

            Color barColor = isPO ? Color.FromArgb(0, 120, 212) :
                             isMPR ? Color.FromArgb(40, 167, 69) :
                                     Color.FromArgb(200, 200, 200);
            e.Graphics.FillRectangle(new SolidBrush(barColor),
                new Rectangle(e.Bounds.X, e.Bounds.Y, 4, e.Bounds.Height));

            string[] parts = msg.Split('|');
            string line1 = parts.Length > 0 ? parts[0].Trim() : msg;
            string line2 = parts.Length > 1 ? parts[1].Trim() : "";

            e.Graphics.DrawString(line1,
                new Font("Segoe UI", 9, FontStyle.Bold),
                new SolidBrush(barColor),
                new RectangleF(e.Bounds.X + 10, e.Bounds.Y + 4, e.Bounds.Width - 14, 20));
            if (!string.IsNullOrEmpty(line2))
                e.Graphics.DrawString(line2,
                    new Font("Segoe UI", 8),
                    Brushes.DimGray,
                    new RectangleF(e.Bounds.X + 10, e.Bounds.Y + 24, e.Bounds.Width - 14, 18));

            e.Graphics.DrawLine(Pens.LightGray,
                e.Bounds.X, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
        }

        private void StartNotifyTimer()
        {
            _lastCheckTime = DateTime.Now;
            try
            {
                using var conn = DatabaseHelper.GetConnection();
                conn.Open();
                _lastCheckTime = DateTime.Now;
            }
            catch { }

            _notifyTimer = new System.Windows.Forms.Timer { Interval = 5 * 60 * 1000 };
            _notifyTimer.Tick += (s, e) => CheckAndNotify(false);
            _notifyTimer.Start();
        }

        private void CheckAndNotify(bool force)
        {
            try
            {
                int newPO = 0, newMPR = 0;
                var msgList = new System.Collections.Generic.List<string>();

                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    string sqlPO = "SELECT PONo, Project_Name, Created_Date FROM PO_head WHERE Created_Date > @since ORDER BY Created_Date DESC";
                    using var cmdPO = new SqlCommand(sqlPO, conn);
                    cmdPO.Parameters.AddWithValue("@since", _lastCheckTime);
                    using var rPO = cmdPO.ExecuteReader();
                    while (rPO.Read())
                    {
                        newPO++;
                        string poNo = rPO["PONo"]?.ToString() ?? "";
                        string proj = rPO["Project_Name"]?.ToString() ?? "";
                        string dt = rPO["Created_Date"] != DBNull.Value
                                      ? Convert.ToDateTime(rPO["Created_Date"]).ToString("dd/MM HH:mm") : "";
                        msgList.Add("[PO] PO moi: " + poNo + " | " + proj + "  " + dt);
                    }
                    rPO.Close();

                    string sqlMPR = "SELECT MPR_No, Project_Name, Modified_Date FROM MPR_Header WHERE Modified_Date > @since ORDER BY Modified_Date DESC";
                    using var cmdMPR = new SqlCommand(sqlMPR, conn);
                    cmdMPR.Parameters.AddWithValue("@since", _lastCheckTime);
                    using var rMPR = cmdMPR.ExecuteReader();
                    while (rMPR.Read())
                    {
                        newMPR++;
                        string mprNo = rMPR["MPR_No"]?.ToString() ?? "";
                        string proj = rMPR["Project_Name"]?.ToString() ?? "";
                        string dt = rMPR["Modified_Date"] != DBNull.Value
                                       ? Convert.ToDateTime(rMPR["Modified_Date"]).ToString("dd/MM HH:mm") : "";
                        msgList.Add("[MPR] MPR cap nhat: " + mprNo + " | " + proj + "  " + dt);
                    }
                }

                _lastCheckTime = DateTime.Now;
                if (newPO == 0 && newMPR == 0 && !force) return;

                if (this.InvokeRequired)
                    this.Invoke(new Action(() => UpdateNotifyUI(newPO, newMPR, msgList, force)));
                else
                    UpdateNotifyUI(newPO, newMPR, msgList, force);
            }
            catch { }
        }

        private void UpdateNotifyUI(int newPO, int newMPR,
            System.Collections.Generic.List<string> msgList, bool force)
        {
            string checkTime = DateTime.Now.ToString("HH:mm dd/MM");

            if (newPO > 0 || newMPR > 0)
            {
                foreach (var msg in msgList)
                    lstNotify.Items.Insert(0, msg);

                var parts = new System.Collections.Generic.List<string>();
                if (newPO > 0) parts.Add(newPO + " PO moi");
                if (newMPR > 0) parts.Add(newMPR + " MPR cap nhat");
                lblNotifyCount.Text = string.Join("  |  ", parts) + "  (" + checkTime + ")";
                lblNotifyCount.ForeColor = Color.FromArgb(220, 53, 69);

                int total = newPO + newMPR;
                btnNotifyToggle.Text = total.ToString();
                btnNotifyToggle.BackColor = Color.FromArgb(220, 53, 69);

                if (!panelNotify.Visible)
                {
                    panelNotify.Visible = true;
                    panelNotify.BringToFront();
                    panelNotify.Location = new Point(
                        this.ClientSize.Width - panelNotify.Width - 10,
                        this.ClientSize.Height - panelNotify.Height - 10);
                }
                _ = LoadDataAsync();
            }
            else if (force)
            {
                lblNotifyCount.Text = "Kiem tra luc " + checkTime + " - Khong co moi";
                lblNotifyCount.ForeColor = Color.Gray;
            }

            var lblNext = panelNotify.Controls.Find("lblNextRefresh", false).FirstOrDefault() as Label;
            if (lblNext != null)
                lblNext.Text = "Kiem tra tiep: " + DateTime.Now.AddMinutes(5).ToString("HH:mm") + "  (moi 5 phut)";
        }


        // =====================================================================
        //  LƯU GHI CHÚ MPR
        // =====================================================================
        private async void BtnSaveMPRNote_Click(object sender, EventArgs e)
        {
            if (dgvMPR == null || dgvMPR.Rows.Count == 0) return;

            // Commit ô đang edit trước khi lưu
            if (dgvMPR.IsCurrentCellInEditMode) dgvMPR.EndEdit();

            // Thu thập dữ liệu trên UI thread trước khi chuyển sang background
            var updates = new System.Collections.Generic.List<(int id, string note)>();
            foreach (DataGridViewRow row in dgvMPR.Rows)
            {
                if (row.IsNewRow) continue;
                object mprIdObj = row.Cells["MPR_ID"]?.Value;
                if (mprIdObj == null || mprIdObj == DBNull.Value) continue;
                updates.Add((Convert.ToInt32(mprIdObj), row.Cells["Ghi chu"]?.Value?.ToString() ?? ""));
            }

            btnSaveMPRNote.Enabled = false;
            try
            {
                int saved = await Task.Run(() =>
                {
                    using var conn = DatabaseHelper.GetConnection();
                    conn.Open();
                    using var tx = conn.BeginTransaction();
                    var cmd = new SqlCommand(
                        "UPDATE MPR_Header SET Notes = @note WHERE MPR_ID = @id", conn, tx);
                    cmd.Parameters.Add("@note", System.Data.SqlDbType.NVarChar);
                    cmd.Parameters.Add("@id",   System.Data.SqlDbType.Int);
                    int count = 0;
                    foreach (var (id, note) in updates)
                    {
                        cmd.Parameters["@note"].Value = note;
                        cmd.Parameters["@id"].Value   = id;
                        cmd.ExecuteNonQuery();
                        count++;
                    }
                    tx.Commit();
                    return count;
                });

                ShowToast($"Đã lưu ghi chú cho {saved} MPR.");
            }
            catch (Exception ex)
            {
                SafeMsg("Lỗi lưu ghi chú: " + ex.Message, "Lỗi");
            }
            finally
            {
                btnSaveMPRNote.Enabled = true;
            }
        }


    }
}