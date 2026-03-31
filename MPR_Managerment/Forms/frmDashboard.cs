using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;

namespace MPR_Managerment.Forms
{
    public partial class frmDashboard : Form
    {
        private TabControl tabMain;
        private TabPage tabPO, tabMPR, tabRIR;

        // PO Tab
        private DataGridView dgvPO;
        private Label lblPOTotal, lblPOOverdue, lblPOCompleted, lblPOInProgress;
        private Panel panelPOSummary;
        private ComboBox cboFilterPO;
        private TextBox txtSearchPO;

        // MPR Tab
        private DataGridView dgvMPR;
        private Label lblMPRTotal, lblMPRHasPO, lblMPRNoPO, lblMPRCompleted;
        private Panel panelMPRSummary;
        private ComboBox cboFilterMPR;
        private TextBox txtSearchMPR;

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
            panelHeader.Resize += (s, e) =>
            {
                panelHeader.Width = this.ClientSize.Width;
                btnRefreshAll.Location = new Point(panelHeader.Width - 150, 8);
            };
            this.Controls.Add(panelHeader);

            // Tab Control — đặt dưới header
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

            // Resize TabControl theo form
            this.Resize += (s, e) =>
            {
                panelHeader.Width = this.ClientSize.Width;
                tabMain.Size = new Size(this.ClientSize.Width, this.ClientSize.Height - 45);
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
                Size = new Size(900, 95),
                BackColor = Color.Transparent,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            tabPO.Controls.Add(panelPOSummary);

            lblPOTotal = AddSummaryCard(panelPOSummary, "Tổng PO", "0", Color.FromArgb(0, 120, 212), 0);
            lblPOInProgress = AddSummaryCard(panelPOSummary, "Đang giao", "0", Color.FromArgb(255, 140, 0), 220);
            lblPOOverdue = AddSummaryCard(panelPOSummary, "Quá hạn", "0", Color.FromArgb(220, 53, 69), 440);
            lblPOCompleted = AddSummaryCard(panelPOSummary, "Hoàn thành", "0", Color.FromArgb(40, 167, 69), 660);

            // Filter row
            int fy = 115;
            AddLabel(tabPO, "Tìm kiếm:", 10, fy);
            txtSearchPO = new TextBox
            {
                Location = new Point(85, fy),
                Size = new Size(200, 25),
                Font = new Font("Segoe UI", 9),
                PlaceholderText = "PO No hoặc MPR No..."
            };
            txtSearchPO.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) LoadPOData(); };
            tabPO.Controls.Add(txtSearchPO);

            AddLabel(tabPO, "Trạng thái:", 300, fy);
            cboFilterPO = new ComboBox
            {
                Location = new Point(380, fy),
                Size = new Size(160, 25),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboFilterPO.Items.AddRange(new[] { "Tất cả", "Draft", "Pending", "Approved", "In Progress", "Completed", "Cancelled" });
            cboFilterPO.SelectedIndex = 0;
            cboFilterPO.SelectedIndexChanged += (s, e) => LoadPOData();
            tabPO.Controls.Add(cboFilterPO);

            var btnSearch = CreateButton("🔍 Tìm", Color.FromArgb(0, 120, 212), new Point(550, fy - 1), 90, 28);
            btnSearch.Click += (s, e) => LoadPOData();
            tabPO.Controls.Add(btnSearch);

            var btnClear = CreateButton("✖ Xóa lọc", Color.FromArgb(108, 117, 125), new Point(650, fy - 1), 90, 28);
            btnClear.Click += (s, e) => { txtSearchPO.Text = ""; cboFilterPO.SelectedIndex = 0; LoadPOData(); };
            tabPO.Controls.Add(btnClear);

            dgvPO = BuildGrid(tabPO, 150);
            dgvPO.RowPrePaint += DgvPO_RowPrePaint;
        }

        // ===== MPR TAB =====
        private void BuildMPRTab()
        {
            tabMPR.BackColor = Color.FromArgb(245, 245, 245);
            panelMPRSummary = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(900, 95),
                BackColor = Color.Transparent,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            tabMPR.Controls.Add(panelMPRSummary);

            lblMPRTotal = AddSummaryCard(panelMPRSummary, "Tổng MPR", "0", Color.FromArgb(0, 120, 212), 0);
            lblMPRHasPO = AddSummaryCard(panelMPRSummary, "Đã có PO", "0", Color.FromArgb(40, 167, 69), 220);
            lblMPRNoPO = AddSummaryCard(panelMPRSummary, "Chưa có PO", "0", Color.FromArgb(220, 53, 69), 440);
            lblMPRCompleted = AddSummaryCard(panelMPRSummary, "Hoàn thành", "0", Color.FromArgb(102, 51, 153), 660);

            int fy = 115;
            AddLabel(tabMPR, "Tìm kiếm:", 10, fy);
            txtSearchMPR = new TextBox
            {
                Location = new Point(85, fy),
                Size = new Size(200, 25),
                Font = new Font("Segoe UI", 9),
                PlaceholderText = "MPR No hoặc tên dự án..."
            };
            txtSearchMPR.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) LoadMPRData(); };
            tabMPR.Controls.Add(txtSearchMPR);

            AddLabel(tabMPR, "Trạng thái:", 300, fy);
            cboFilterMPR = new ComboBox
            {
                Location = new Point(380, fy),
                Size = new Size(160, 25),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboFilterMPR.Items.AddRange(new[] { "Tất cả", "Mới", "Đang xử lý", "Đã duyệt", "Hoàn thành", "Hủy" });
            cboFilterMPR.SelectedIndex = 0;
            cboFilterMPR.SelectedIndexChanged += (s, e) => LoadMPRData();
            tabMPR.Controls.Add(cboFilterMPR);

            var btnSearch = CreateButton("🔍 Tìm", Color.FromArgb(0, 120, 212), new Point(550, fy - 1), 90, 28);
            btnSearch.Click += (s, e) => LoadMPRData();
            tabMPR.Controls.Add(btnSearch);

            var btnClear = CreateButton("✖ Xóa lọc", Color.FromArgb(108, 117, 125), new Point(650, fy - 1), 90, 28);
            btnClear.Click += (s, e) => { txtSearchMPR.Text = ""; cboFilterMPR.SelectedIndex = 0; LoadMPRData(); };
            tabMPR.Controls.Add(btnClear);

            dgvMPR = BuildGrid(tabMPR, 150);
            dgvMPR.RowPrePaint += DgvMPR_RowPrePaint;
            dgvMPR.CellDoubleClick += DgvMPR_CellDoubleClick; // MỞ FORM MPR KHI DOUBLE CLICK
        }

        private void DgvMPR_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            // Lấy MPR_ID từ dòng được click (cột này đã bị ẩn)
            var row = dgvMPR.Rows[e.RowIndex];
            int mprId = Convert.ToInt32(row.Cells["MPR_ID"].Value);

            // Mở form frmMPR và truyền MPR_ID sang
            var frm = new frmMPR(mprId);
            frm.Show();
        }

        // ===== RIR TAB =====
        private void BuildRIRTab()
        {
            tabRIR.BackColor = Color.FromArgb(245, 245, 245);
            panelRIRSummary = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(900, 95),
                BackColor = Color.Transparent,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            tabRIR.Controls.Add(panelRIRSummary);

            lblRIRTotal = AddSummaryCard(panelRIRSummary, "Tổng RIR", "0", Color.FromArgb(0, 120, 212), 0);
            lblRIRPending = AddSummaryCard(panelRIRSummary, "Chờ kiểm tra", "0", Color.FromArgb(255, 140, 0), 220);
            lblRIRInspecting = AddSummaryCard(panelRIRSummary, "Đang kiểm tra", "0", Color.FromArgb(102, 51, 153), 440);
            lblRIRDone = AddSummaryCard(panelRIRSummary, "Hoàn thành", "0", Color.FromArgb(40, 167, 69), 660);

            // Filter row
            int fy = 115;
            AddLabel(tabRIR, "Tìm kiếm:", 10, fy);
            txtSearchRIR = new TextBox
            {
                Location = new Point(85, fy),
                Size = new Size(200, 25),
                Font = new Font("Segoe UI", 9),
                PlaceholderText = "RIR No hoặc PO No..."
            };
            txtSearchRIR.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) LoadRIRData(); };
            tabRIR.Controls.Add(txtSearchRIR);

            AddLabel(tabRIR, "Trạng thái:", 300, fy);
            cboFilterRIR = new ComboBox
            {
                Location = new Point(380, fy),
                Size = new Size(160, 25),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboFilterRIR.Items.AddRange(new[] { "Tất cả", "Chờ kiểm tra", "Đang kiểm tra", "Hoàn thành" });
            cboFilterRIR.SelectedIndex = 0;
            cboFilterRIR.SelectedIndexChanged += (s, e) => LoadRIRData();
            tabRIR.Controls.Add(cboFilterRIR);

            var btnSearch = CreateButton("🔍 Tìm", Color.FromArgb(0, 120, 212), new Point(550, fy - 1), 90, 28);
            btnSearch.Click += (s, e) => LoadRIRData();
            tabRIR.Controls.Add(btnSearch);

            var btnClear = CreateButton("✖ Xóa lọc", Color.FromArgb(108, 117, 125), new Point(650, fy - 1), 90, 28);
            btnClear.Click += (s, e) => { txtSearchRIR.Text = ""; cboFilterRIR.SelectedIndex = 0; LoadRIRData(); };
            tabRIR.Controls.Add(btnClear);

            // Label tiêu đề grid trên
            tabRIR.Controls.Add(new Label
            {
                Text = "DANH SÁCH PO & TIẾN ĐỘ RIR",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 152),
                Size = new Size(300, 20)
            });

            // Grid trên — danh sách PO kèm tiến độ RIR
            dgvRIR = new DataGridView
            {
                Location = new Point(10, 173),
                Size = new Size(tabRIR.Width - 20, 200),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            dgvRIR.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvRIR.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvRIR.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvRIR.EnableHeadersVisualStyles = false;
            dgvRIR.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvRIR.RowPrePaint += DgvRIR_RowPrePaint;
            dgvRIR.SelectionChanged += DgvRIR_SelectionChanged;
            tabRIR.Controls.Add(dgvRIR);

            // Label tiêu đề grid dưới
            var lblDetailTitle = new Label
            {
                Text = "CHI TIẾT RIR THEO PO (click vào PO ở trên để xem)",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 383),
                Size = new Size(500, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            tabRIR.Controls.Add(lblDetailTitle);

            // Grid dưới — chi tiết RIR theo PO được chọn
            dgvRIRDetail = new DataGridView
            {
                Location = new Point(10, 404),
                Size = new Size(tabRIR.Width - 20, tabRIR.Height - 415),
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
            dgvRIRDetail.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(102, 51, 153);
            dgvRIRDetail.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvRIRDetail.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvRIRDetail.EnableHeadersVisualStyles = false;
            dgvRIRDetail.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 240, 255);
            dgvRIRDetail.CellFormatting += DgvRIRDetail_CellFormatting;
            tabRIR.Controls.Add(dgvRIRDetail);
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
            tab.Controls.Add(dgv);
            return dgv;
        }

        private Label AddSummaryCard(Panel parent, string title, string value, Color color, int x)
        {
            var card = new Panel { Location = new Point(x, 0), Size = new Size(200, 90), BackColor = color };
            parent.Controls.Add(card);
            card.Controls.Add(new Label
            {
                Text = title,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(0, 8),
                Size = new Size(200, 22),
                TextAlign = ContentAlignment.MiddleCenter
            });
            var lbl = new Label
            {
                Text = value,
                Font = new Font("Segoe UI", 26, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(0, 32),
                Size = new Size(200, 50),
                TextAlign = ContentAlignment.MiddleCenter
            };
            card.Controls.Add(lbl);
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

                string where = "WHERE 1=1";
                if (!string.IsNullOrEmpty(search))
                    where += $" AND (h.PONo LIKE N'%{search}%' OR h.MPR_No LIKE N'%{search}%' OR h.Project_Name LIKE N'%{search}%')";
                if (filter != "Tất cả")
                    where += $" AND h.Status = N'{filter}'";

                string sql = $@"
                    SELECT
                        h.PO_ID,
                        h.PONo                             AS [PO No],
                        h.Project_Name                     AS [Dự án],
                        h.MPR_No                           AS [MPR No],
                        h.PO_Date                          AS [Ngày PO],
                        h.Status                           AS [Trạng thái],
                        h.Revise                           AS [Rev],
                        COUNT(d.PO_Detail_ID)              AS [Tổng items],
                        ISNULL(SUM(d.Qty_Per_Sheet), 0)    AS [Tổng SL đặt],
                        ISNULL(SUM(d.Received), 0)         AS [Tổng SL nhận],
                        CASE
                            WHEN ISNULL(SUM(d.Qty_Per_Sheet), 0) = 0 THEN 0
                            ELSE CAST(SUM(d.Received) * 100.0 / SUM(d.Qty_Per_Sheet) AS DECIMAL(5,1))
                        END                                AS [% Giao hàng],
                        MIN(d.RequestDay)                  AS [Ngày giao sớm nhất],
                        CASE
                            WHEN MIN(d.RequestDay) < GETDATE()
                             AND ISNULL(SUM(d.Received), 0) < ISNULL(SUM(d.Qty_Per_Sheet), 0)
                            THEN N'⚠ Quá hạn'
                            ELSE N'✅ Đúng hạn'
                        END                                AS [Cảnh báo]
                    FROM PO_head h
                    LEFT JOIN PO_Detail d ON h.PO_ID = d.PO_ID
                    {where}
                    GROUP BY h.PO_ID, h.PONo, h.Project_Name, h.MPR_No, h.PO_Date, h.Status, h.Revise
                    ORDER BY h.PO_Date DESC";

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
                    e.CellStyle.ForeColor = pct >= 100 ? Color.FromArgb(40, 167, 69) : pct >= 50 ? Color.FromArgb(255, 140, 0) : Color.FromArgb(220, 53, 69);
                    e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    e.Value = $"{pct}%";
                    e.FormattingApplied = true;
                }
            }
            if (col == "Cảnh báo")
            {
                e.CellStyle.ForeColor = e.Value?.ToString().Contains("Quá") == true ? Color.Red : Color.FromArgb(40, 167, 69);
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
        }

        private void DgvPO_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex < 0 || dgvPO.Rows[e.RowIndex].IsNewRow) return;
            var row = dgvPO.Rows[e.RowIndex];
            string canh = row.Cells["Cảnh báo"].Value?.ToString() ?? "";
            string status = row.Cells["Trạng thái"].Value?.ToString() ?? "";

            if (canh.Contains("Quá")) row.DefaultCellStyle.BackColor = Color.FromArgb(255, 235, 235);
            else if (status == "Completed") row.DefaultCellStyle.BackColor = Color.FromArgb(235, 255, 235);
            else if (status == "In Progress" || status == "Approved")
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
                        h.Requestor                        AS [Người YC],
                        h.Required_Date                    AS [Ngày cần],
                        h.Status                           AS [Trạng thái],
                        h.Rev                              AS [Rev],
                        COUNT(DISTINCT d.Detail_ID)        AS [Tổng items],
                        COUNT(DISTINCT po.PO_ID)           AS [Số PO],
                        CASE
                            WHEN COUNT(DISTINCT po.PO_ID) > 0 THEN N'✅ Đã có PO'
                            ELSE N'❌ Chưa có PO'
                        END                                AS [Tình trạng PO],
                        CASE
                            WHEN COUNT(DISTINCT d.Detail_ID) = 0 THEN 0
                            ELSE CAST(COUNT(DISTINCT pod.PO_Detail_ID) * 100.0 / COUNT(DISTINCT d.Detail_ID) AS DECIMAL(5,1))
                        END                                AS [% Item đặt hàng],
                        DATEDIFF(DAY, h.Created_Date, MIN(po.Created_Date)) AS [Ngày đến PO],
                        h.Created_Date                     AS [Ngày tạo]
                    FROM MPR_Header h
                    LEFT JOIN MPR_Details d   ON h.MPR_ID = d.MPR_ID
                    LEFT JOIN PO_head po      ON po.MPR_No = h.MPR_No
                    LEFT JOIN PO_Detail pod   ON pod.PO_ID = po.PO_ID AND pod.MPR_Detail_ID = d.Detail_ID
                    {where}
                    GROUP BY h.MPR_ID, h.MPR_No, h.Project_Name, h.Requestor,
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

                    dgvMPR.CellFormatting -= DgvMPR_CellFormatting;
                    dgvMPR.CellFormatting += DgvMPR_CellFormatting;

                    // Summary
                    int total = dt.Rows.Count, hasPO = 0, noPO = 0, completed = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        string tinh = row["Tình trạng PO"]?.ToString() ?? "";
                        string status = row["Trạng thái"]?.ToString() ?? "";
                        if (tinh.Contains("Đã có")) hasPO++;
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
                e.CellStyle.ForeColor = e.Value?.ToString().Contains("Đã có") == true ? Color.FromArgb(40, 167, 69) : Color.FromArgb(220, 53, 69);
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
                        h.RIR_No                                            AS [RIR No],
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

                    // Summary
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