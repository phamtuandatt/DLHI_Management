using MPR_Managerment.Helpers;
using MPR_Managerment.Services;
using System.Threading.Tasks;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace MPR_Managerment.Forms
{
    public class frmZaloImport : Form
    {
        // Controls
        private Label lblFolder;
        private Button btnImportLatest, btnImportSelected, btnRefreshFiles, btnViewDB, btnMarkOldPaid;
        private TextBox txtSearchPO;
        private ListBox lstFiles;
        private DataGridView dgvPreview, dgvDB;
        private Panel panelTop, panelLeft, panelRight;
        private Label lblStatus;
        private ProgressBar pbImport;
        private TabControl tabs;
        private TabPage tabPreview, tabDB;
        private CheckBox chkAutoWatch;
        private Label lblWatch;

        // FileSystemWatcher
        private FileSystemWatcher _watcher;
        private System.Windows.Forms.Timer _debounceTimer;
        private string _pendingFile;


        public frmZaloImport()
        {
            BuildUI();
            RefreshFileList();
            ZaloImportService.EnsureTable();
            frmAIChat.Attach(this);
        }

        private void BuildUI()
        {
            Text = "📥  Zalo Import — 품의서";
            BackColor = Color.FromArgb(245, 245, 245);
            Size = new Size(1400, 750);
            MinimumSize = new Size(900, 550);

            SuspendLayout();

            // ── Top toolbar: 2 hàng ──
            panelTop = new Panel { Dock = DockStyle.Top, Height = 82, BackColor = Color.White };
            // Hàng 1: folder path
            lblFolder = new Label
            {
                Text = $"📁  {ZaloImportService.ZaloDownloadFolder}",
                Location = new Point(8, 8), Size = new Size(950, 18),
                Font = new Font("Segoe UI", 8.5f), ForeColor = Color.FromArgb(80, 80, 80)
            };
            panelTop.Controls.Add(lblFolder);
            // Hàng 2: buttons
            btnRefreshFiles  = MkBtn("🔄 Làm mới",            Color.FromArgb(108,117,125),  8, 33, 110, 28);
            btnImportLatest  = MkBtn("⚡ Import file mới nhất", Color.FromArgb(0,120,212),   126, 33, 175, 28);
            btnImportSelected= MkBtn("📥 Import file đã chọn", Color.FromArgb(40,167,69),   309, 33, 175, 28);
            btnViewDB        = MkBtn("🗄️ Xem DB",              Color.FromArgb(102,51,153),  492, 33, 100, 28);
            btnMarkOldPaid   = MkBtn("✅ Đánh dấu Đã TT (≤30/12/25)", Color.FromArgb(180, 90, 0), 600, 33, 210, 28);
            btnRefreshFiles.Click   += (s, e) => RefreshFileList();
            btnImportLatest.Click   += BtnImportLatest_Click;
            btnImportSelected.Click += BtnImportSelected_Click;
            btnViewDB.Click         += BtnViewDB_Click;
            btnMarkOldPaid.Click    += BtnMarkOldPaid_Click;
            chkAutoWatch = new CheckBox
            {
                Text = "Tự động import khi có file mới",
                Location = new Point(1115, 38), Size = new Size(235, 18),
                Font = new Font("Segoe UI", 9), Checked = true
            };
            chkAutoWatch.CheckedChanged += ChkAutoWatch_Changed;

            var lblSearch = new Label { Text = "🔍 PO No:", Location = new Point(820, 38), Size = new Size(65, 26), Font = new Font("Segoe UI", 9), TextAlign = ContentAlignment.MiddleLeft };
            txtSearchPO = new TextBox { Location = new Point(887, 36), Size = new Size(220, 26), Font = new Font("Segoe UI", 9), PlaceholderText = "Tìm theo PO No..." };
            txtSearchPO.TextChanged += (s, e) => FilterByPO();

            panelTop.Controls.AddRange(new Control[] { btnRefreshFiles, btnImportLatest, btnImportSelected, btnViewDB, btnMarkOldPaid, chkAutoWatch, lblSearch, txtSearchPO });

            // ── Watch label (Bottom) ──
            lblWatch = new Label
            {
                Dock = DockStyle.Bottom, Height = 22,
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = Color.FromArgb(0, 150, 100),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 0, 0),
                BackColor = Color.FromArgb(235, 255, 235),
                Text = "👀  Đang theo dõi folder Zalo..."
            };

            // ── Status bar (Bottom) ──
            var panelStatus = new Panel { Dock = DockStyle.Bottom, Height = 28, BackColor = Color.FromArgb(230, 230, 230) };
            lblStatus = new Label { Text = "Sẵn sàng.", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Segoe UI", 9), Padding = new Padding(8, 0, 0, 0) };
            pbImport = new ProgressBar { Dock = DockStyle.Right, Width = 160, Style = ProgressBarStyle.Marquee, MarqueeAnimationSpeed = 0 };
            panelStatus.Controls.Add(lblStatus);
            panelStatus.Controls.Add(pbImport);

            // ── SplitContainer (Fill) ──
            var splitter = new SplitContainer { Dock = DockStyle.Fill };

            // Thứ tự add vào form: Fill trước, Bottom sau, Top cuối
            // (WinForms xử lý dock theo thứ tự ngược: cuối cùng được xử lý trước)
            Controls.Add(splitter);     // Fill  → xử lý cuối (= lấy phần còn lại)
            Controls.Add(panelStatus);  // Bottom→ xử lý trước splitter
            Controls.Add(lblWatch);     // Bottom→ xử lý trước panelStatus
            Controls.Add(panelTop);     // Top   → xử lý đầu tiên

            // SplitterDistance phải set sau khi form hiển thị đầy đủ
            Shown += (s, e) =>
            {
                try { splitter.SplitterDistance = Math.Max(200, Math.Min(260, splitter.Width - 350)); }
                catch { }
            };

            // ── Left panel: danh sách file ──
            var lblFiles = new Label
            {
                Text = "📋  Danh sách file 품의서",
                Dock = DockStyle.Top, Height = 28,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 0, 0), BackColor = Color.White
            };
            lstFiles = new ListBox
            {
                Dock = DockStyle.Fill, Font = new Font("Segoe UI", 9),
                BorderStyle = BorderStyle.None, BackColor = Color.White
            };
            lstFiles.SelectedIndexChanged += LstFiles_SelectedIndexChanged;
            // Fill trước, Top sau
            splitter.Panel1.Controls.Add(lstFiles);
            splitter.Panel1.Controls.Add(lblFiles);

            // ── Right panel: tabs ──
            tabs = new TabControl { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            tabPreview = new TabPage("🔍  Xem trước file");
            tabDB      = new TabPage("🗄️  Dữ liệu đã import vào DB");
            tabs.TabPages.AddRange(new[] { tabPreview, tabDB });
            splitter.Panel2.Controls.Add(tabs);

            // Tab Preview: chỉ grid, Dock.Fill
            dgvPreview = BuildGrid(Color.FromArgb(0, 120, 212));
            BuildPreviewCols();
            tabPreview.Controls.Add(dgvPreview);

            // Tab DB: chỉ grid, Dock.Fill (nút đã đưa lên toolbar)
            dgvDB = BuildGrid(Color.FromArgb(102, 51, 153));
            BuildDBCols();
            tabDB.Controls.Add(dgvDB);

            ResumeLayout(true);

            // Khởi động watcher
            StartWatcher();
        }

        private void RefreshFileList()
        {
            lstFiles.Items.Clear();
            var cutoff = DateTime.Today.AddDays(-7);
            var files = ZaloImportService.FindImportFiles()
                .Where(f => f.Item2 >= cutoff)
                .ToList();
            foreach (var f in files)
                lstFiles.Items.Add(new FileItem(f.Item1, f.Item2));

            if (files.Count > 0)
            {
                lstFiles.SelectedIndex = 0;
                lblStatus.Text = $"Tìm thấy {files.Count} file 품의서 (7 ngày gần nhất). File mới nhất: {files[0].Item2:dd/MM/yyyy}";
            }
            else
            {
                lblStatus.Text = "Không tìm thấy file 품의서 nào trong 7 ngày gần nhất.";
            }
        }

        private void LstFiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstFiles.SelectedItem is FileItem item)
                LoadPreview(item.Path);
        }

        private void LoadPreview(string filePath)
        {
            dgvPreview.Rows.Clear();
            lblStatus.Text = "Đang đọc file...";
            try
            {
                var rows = ZaloImportService.ReadFile(filePath);
                foreach (var r in rows)
                {
                    int i = dgvPreview.Rows.Add();
                    var dr = dgvPreview.Rows[i];
                    dr.Cells["P_Row"].Value   = r.RowIndex;
                    dr.Cells["P_PO"].Value    = r.PONo;
                    dr.Cells["P_Ecount"].Value= r.EcountNo;
                    dr.Cells["P_Title"].Value = r.TitleEN;
                    dr.Cells["P_Amount"].Value= FormatAmt(r.Amount);
                    dr.Cells["P_VAT"].Value   = FormatAmt(r.VAT);
                    dr.Cells["P_Final"].Value = FormatAmt(r.FinalAmount);

                    decimal finalAmt = r.FinalAmount ?? 0;
                    decimal paidAmt = r.PaidDate.HasValue ? finalAmt : 0;
                    decimal remain = finalAmt - paidAmt;

                    dr.Cells["P_Paid"].Value  = paidAmt > 0 ? FormatAmt(paidAmt) : "";
                    dr.Cells["P_Remain"].Value = FormatAmt(remain);

                    dr.Cells["P_D1"].Value    = FormatAmt(r.Dot1);
                    dr.Cells["P_D2"].Value    = FormatAmt(r.Dot2);
                    dr.Cells["P_D3"].Value    = FormatAmt(r.Dot3);
                    dr.Cells["P_D4"].Value    = FormatAmt(r.Dot4);
                    dr.Cells["P_FTCash"].Value = r.FTCash;
                    dr.Cells["P_PayDate"].Value= r.PaidDate.HasValue ? r.PaidDate.Value.ToString("dd/MM/yyyy") : "";
                    dr.Cells["P_Status"].Value= r.ProgressStatus;
                }
                lblStatus.Text = $"Preview: {rows.Count} dòng dữ liệu từ {Path.GetFileName(filePath)}";
                FilterByPO();
            }
            catch (Exception ex)
            {
                lblStatus.Text = "Lỗi đọc file: " + ex.Message;
            }
        }

        private void BtnImportLatest_Click(object sender, EventArgs e)
        {
            var files = ZaloImportService.FindImportFiles();
            if (files.Count == 0) { Status("Không tìm thấy file nào!"); return; }
            DoImport(files[0].Item1, files[0].Item2, clearFirst: true);
        }

        private void BtnImportSelected_Click(object sender, EventArgs e)
        {
            if (lstFiles.SelectedItem is not FileItem item) { Status("Vui lòng chọn file trong danh sách!"); return; }
            DoImport(item.Path, item.Date, clearFirst: false);
        }

        private async void DoImport(string path, DateTime date, bool clearFirst = false)
        {
            SetBusy(true);
            try
            {
                if (clearFirst)
                {
                    Status("Đang xóa dữ liệu cũ cùng ngày...");
                    await Task.Run(() => ZaloImportService.ClearByDate(date));
                }

                Status($"Đang import {Path.GetFileName(path)}...");
                var result = await Task.Run(() =>
                    ZaloImportService.ImportToDB(path, date));

                if (result.Success)
                {
                    Status($"✅  Import thành công! {result.InsertedRows} dòng mới, {result.UpdatedRows} dòng cập nhật (tổng {result.TotalRows} dòng) — {date:dd/MM/yyyy}");
                    LoadPreview(path);
                    // Chuyển sang tab DB và load
                    tabs.SelectedTab = tabDB;
                    LoadDBGrid();

                    // -------------------------------------------------
                    // Update PO payment status based on imported data (Background)
                    // -------------------------------------------------
                    var rows = ZaloImportService.ReadFile(path);
                    var distinctPo = rows.Select(r => r.PONo).Where(p => !string.IsNullOrEmpty(p)).Distinct().ToList();
                    
                    Task.Run(() =>
                    {
                        foreach (var po in distinctPo)
                        {
                            UpdatePaymentStatusByPONo(po);
                        }
                        Invoke(new Action(() => Status("✅ Đã cập nhật xong trạng thái PO.")));
                    });
                }
                else
                {
                    Status($"❌  Lỗi import: {result.Error}");
                    MessageBox.Show(TopLevelControl, result.Error, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            finally { SetBusy(false); }
        }

        // -----------------------------------------------------------------
        // Replicated logic from frmPayment.cs to update PO payment status
        // -----------------------------------------------------------------
        private void UpdatePaymentStatusByPONo(string poNo)
        {
            if (string.IsNullOrEmpty(poNo)) return;

            using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
            conn.Open();

            // 1. Tổng đã TT từ PO_HistoryPaid theo PONo
            var cmdPaid = new Microsoft.Data.SqlClient.SqlCommand(
                "SELECT ISNULL(SUM(Amount_Total), 0) AS Total_Paid FROM PO_HistoryPaid WHERE PONo = @poNo",
                conn);
            cmdPaid.Parameters.AddWithValue("@poNo", poNo);
            decimal totalPaid = Convert.ToDecimal(cmdPaid.ExecuteScalar() ?? 0);

            // 2. Tổng kế hoạch từ PO_Payment_Schedule
            var cmdPlan = new Microsoft.Data.SqlClient.SqlCommand(@"
                SELECT ISNULL(SUM(ps.Amount_Plan), 0) AS Total_Plan
                FROM PO_Payment_Schedule ps
                INNER JOIN PO_head ph ON ph.PO_ID = ps.PO_ID
                WHERE ph.PONo = @poNo2", conn);
            cmdPlan.Parameters.AddWithValue("@poNo2", poNo);
            decimal totalPlan = Convert.ToDecimal(cmdPlan.ExecuteScalar() ?? 0);

            // 3. Xác định trạng thái tổng (Tính cả Zalo_PaymentImport)
            var cmdZalo = new Microsoft.Data.SqlClient.SqlCommand(
                "SELECT ISNULL(SUM(Final_Amount), 0) FROM Zalo_PaymentImport WHERE PO_No = @poNo", conn);
            cmdZalo.Parameters.AddWithValue("@poNo", poNo);
            decimal zaloPaid = Convert.ToDecimal(cmdZalo.ExecuteScalar() ?? 0);
            
            decimal effectivePaid = totalPaid + zaloPaid;
            
            string newPoStatus = effectivePaid <= 0 ? "Chưa TT"
                               : effectivePaid >= totalPlan ? "Đã TT đủ"
                               : "Một phần";

            // 4. Cập nhật từng đợt
            var cmdDots = new Microsoft.Data.SqlClient.SqlCommand(@"
                SELECT ps.Schedule_ID, ps.Amount_Plan, ps.Dot_TT
                FROM PO_Payment_Schedule ps
                INNER JOIN PO_head ph ON ph.PO_ID = ps.PO_ID
                WHERE ph.PONo = @poNo3", conn);
            cmdDots.Parameters.AddWithValue("@poNo3", poNo);

            var dotUpdates = new List<(int sid, string status)>();
            using (var rdr = cmdDots.ExecuteReader())
                while (rdr.Read())
                {
                    int sid = Convert.ToInt32(rdr["Schedule_ID"]);
                    int dotTT = Convert.ToInt32(rdr["Dot_TT"]);
                    decimal plan = rdr["Amount_Plan"] != DBNull.Value ? Convert.ToDecimal(rdr["Amount_Plan"]) : 0;
                    
                    // Lấy Paid từ cả 2 nguồn
                    var cmdDotPaid = new Microsoft.Data.SqlClient.SqlCommand(@"
                        SELECT (SELECT ISNULL(SUM(hp.Amount_Total), 0) 
                                FROM PO_PrintRequestHistory prh 
                                INNER JOIN PO_HistoryPaid hp ON hp.Print_ID = prh.Print_ID 
                                WHERE prh.PONo = @p AND prh.Dot_TT = @d)
                             + (SELECT ISNULL(SUM(Final_Amount), 0) 
                                FROM Zalo_PaymentImport 
                                WHERE PO_No = @p AND Progress_Status LIKE '%Đợt ' + CAST(@d AS NVARCHAR))", conn);
                    cmdDotPaid.Parameters.AddWithValue("@p", poNo);
                    cmdDotPaid.Parameters.AddWithValue("@d", dotTT);
                    decimal dotPaid = Convert.ToDecimal(cmdDotPaid.ExecuteScalar() ?? 0);
                    
                    string dotSt = dotPaid <= 0 ? "Chưa TT"
                                   : dotPaid >= plan ? "Đã TT đủ"
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

            // 5. Cập nhật PO_head.Status
            var cmdPO = new Microsoft.Data.SqlClient.SqlCommand(
                "UPDATE PO_head SET Status = @st WHERE PONo = @poNo5", conn);
            cmdPO.Parameters.AddWithValue("@st", newPoStatus);
            cmdPO.Parameters.AddWithValue("@poNo5", poNo);
            cmdPO.ExecuteNonQuery();

            // Debug
            System.Diagnostics.Debug.WriteLine(
                $"[UpdatePaymentStatus] {poNo} → {newPoStatus} (paid={totalPaid:N2} / plan={totalPlan:N2})");
        }

        private void BtnViewDB_Click(object sender, EventArgs e) => LoadDBGrid();

        private async void BtnMarkOldPaid_Click(object sender, EventArgs e)
        {
            var confirm = MessageBox.Show(
                "Thao tác này sẽ đánh dấu Đã thanh toán cho tất cả các PO có Ecount No từ 30/12/2025 trở về trước (chưa có ngày thanh toán).\n\nBạn có chắc chắn muốn tiếp tục?",
                "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (confirm != DialogResult.Yes) return;

            SetBusy(true);
            Status("Đang cập nhật trạng thái thanh toán...");
            try
            {
                var cutoff = new DateTime(2025, 12, 30);
                List<string> affectedPOs = new();

                await Task.Run(() =>
                {
                    using var conn = DatabaseHelper.GetConnection();
                    conn.Open();

                    // Lấy danh sách PO bị ảnh hưởng trước khi update
                    var cmdSelect = new Microsoft.Data.SqlClient.SqlCommand(@"
                        SELECT DISTINCT PO_No
                        FROM Zalo_PaymentImport
                        WHERE Paid_Date IS NULL
                          AND Ecount_No IS NOT NULL AND Ecount_No != ''
                          AND TRY_CAST(REPLACE(
                                CASE WHEN CHARINDEX('-', Ecount_No) > 0
                                     THEN SUBSTRING(Ecount_No, 1, CHARINDEX('-', Ecount_No) - 1)
                                     ELSE Ecount_No END,
                              '/', '-') AS DATE) <= @cutoff", conn);
                    cmdSelect.Parameters.AddWithValue("@cutoff", cutoff);
                    using (var rdr = cmdSelect.ExecuteReader())
                        while (rdr.Read())
                            if (!string.IsNullOrEmpty(rdr[0]?.ToString()))
                                affectedPOs.Add(rdr[0].ToString()!);

                    // Cập nhật Paid_Date = date parse từ Ecount_No
                    var cmdUpdate = new Microsoft.Data.SqlClient.SqlCommand(@"
                        UPDATE Zalo_PaymentImport
                        SET Paid_Date = TRY_CAST(REPLACE(
                                CASE WHEN CHARINDEX('-', Ecount_No) > 0
                                     THEN SUBSTRING(Ecount_No, 1, CHARINDEX('-', Ecount_No) - 1)
                                     ELSE Ecount_No END,
                              '/', '-') AS DATE)
                        WHERE Paid_Date IS NULL
                          AND Ecount_No IS NOT NULL AND Ecount_No != ''
                          AND TRY_CAST(REPLACE(
                                CASE WHEN CHARINDEX('-', Ecount_No) > 0
                                     THEN SUBSTRING(Ecount_No, 1, CHARINDEX('-', Ecount_No) - 1)
                                     ELSE Ecount_No END,
                              '/', '-') AS DATE) <= @cutoff2", conn);
                    cmdUpdate.Parameters.AddWithValue("@cutoff2", cutoff);
                    cmdUpdate.ExecuteNonQuery();

                    // Cập nhật trạng thái PO_head và PO_Payment_Schedule cho các PO bị ảnh hưởng
                    foreach (var po in affectedPOs)
                        UpdatePaymentStatusByPONo(po);
                });

                Status($"✅ Đã đánh dấu Đã TT cho {affectedPOs.Count} PO có Ecount No ≤ 30/12/2025.");
                tabs.SelectedTab = tabDB;
                LoadDBGrid();
            }
            catch (Exception ex)
            {
                Status("❌ Lỗi: " + ex.Message);
                MessageBox.Show(ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { SetBusy(false); }
        }

        private void LoadDBGrid()
        {
            dgvDB.Rows.Clear();
            try
            {
                using var conn = DatabaseHelper.GetConnection();
                conn.Open();
                var chk = new Microsoft.Data.SqlClient.SqlCommand(
                    "SELECT COUNT(1) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='Zalo_PaymentImport'", conn);
                if ((int)chk.ExecuteScalar() == 0) return;

                var cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                    SELECT TOP 2000 Import_ID, File_Date, Row_Index, PO_No, Ecount_No, GW_No,
                           Amount, VAT, Final_Amount, Dot1, Dot2, Dot3, Dot4,
                           FT_Cash, Paid_Date,
                           Progress_Status, Note, Imported_At
                    FROM Zalo_PaymentImport
                    ORDER BY File_Date DESC, Row_Index ASC", conn);
                using var rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    int i = dgvDB.Rows.Add();
                    var dr = dgvDB.Rows[i];
                    dr.Cells["DB_ID"].Value      = rdr["Import_ID"];
                    dr.Cells["DB_Date"].Value    = rdr["File_Date"] != DBNull.Value ? ((DateTime)rdr["File_Date"]).ToString("dd/MM/yyyy") : "";
                    dr.Cells["DB_Row"].Value     = rdr["Row_Index"];
                    dr.Cells["DB_PO"].Value      = rdr["PO_No"]?.ToString() ?? "";
                    dr.Cells["DB_Ecount"].Value  = rdr["Ecount_No"]?.ToString() ?? "";
                    dr.Cells["DB_Amount"].Value  = rdr["Amount"] != DBNull.Value ? FormatAmt(Convert.ToDecimal(rdr["Amount"])) : "";
                    dr.Cells["DB_VAT"].Value     = rdr["VAT"] != DBNull.Value ? FormatAmt(Convert.ToDecimal(rdr["VAT"])) : "";
                    decimal dbFinal = rdr["Final_Amount"] != DBNull.Value ? Convert.ToDecimal(rdr["Final_Amount"]) : 0;
                    bool dbHasPaid = rdr["Paid_Date"] != DBNull.Value;
                    decimal dbPaid = dbHasPaid ? dbFinal : 0;
                    decimal dbRemain = dbFinal - dbPaid;

                    dr.Cells["DB_Final"].Value   = dbFinal > 0 ? FormatAmt(dbFinal) : "";
                    dr.Cells["DB_Paid"].Value    = dbPaid > 0 ? FormatAmt(dbPaid) : "";
                    dr.Cells["DB_Remain"].Value  = FormatAmt(dbRemain);

                    dr.Cells["DB_D1"].Value      = rdr["Dot1"] != DBNull.Value ? FormatAmt(Convert.ToDecimal(rdr["Dot1"])) : "";
                    dr.Cells["DB_D2"].Value      = rdr["Dot2"] != DBNull.Value ? FormatAmt(Convert.ToDecimal(rdr["Dot2"])) : "";
                    dr.Cells["DB_D3"].Value      = rdr["Dot3"] != DBNull.Value ? FormatAmt(Convert.ToDecimal(rdr["Dot3"])) : "";
                    dr.Cells["DB_D4"].Value      = rdr["Dot4"] != DBNull.Value ? FormatAmt(Convert.ToDecimal(rdr["Dot4"])) : "";
                    dr.Cells["DB_FTCash"].Value  = rdr["FT_Cash"]?.ToString() ?? "";
                    dr.Cells["DB_PayDate"].Value = rdr["Paid_Date"] != DBNull.Value ? ((DateTime)rdr["Paid_Date"]).ToString("dd/MM/yyyy") : "";
                    dr.Cells["DB_Status"].Value  = rdr["Progress_Status"]?.ToString() ?? "";
                    dr.Cells["DB_Note"].Value    = rdr["Note"]?.ToString() ?? "";
                    dr.Cells["DB_ImportAt"].Value= rdr["Imported_At"] != DBNull.Value ? ((DateTime)rdr["Imported_At"]).ToString("dd/MM/yyyy HH:mm") : "";
                    dr.Cells["DB_ImportID"].Value= rdr["Import_ID"];
                }
                lblStatus.Text = $"Đang hiển thị {dgvDB.Rows.Count} bản ghi từ DB.";
                FilterByPO();
            }
            catch (Exception ex) { lblStatus.Text = "Lỗi load DB: " + ex.Message; }
        }

        // ── FileSystemWatcher ──
        private void StartWatcher()
        {
            string folder = ZaloImportService.ZaloDownloadFolder;
            if (!Directory.Exists(folder)) return;

            _debounceTimer = new System.Windows.Forms.Timer { Interval = 2000 };
            _debounceTimer.Tick += (s, e) =>
            {
                _debounceTimer.Stop();
                if (_pendingFile != null)
                {
                    string f = _pendingFile; _pendingFile = null;
                    string fname = Path.GetFileNameWithoutExtension(f);
                    string datePart = fname.Substring("0. 품의서 ".Length).Trim();
                    if (DateTime.TryParseExact(datePart, "yyyy-MM-dd",
                        System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None, out DateTime dt))
                    {
                        Invoke(() =>
                        {
                            lblWatch.Text = $"🔔  File mới phát hiện: {Path.GetFileName(f)}";
                            RefreshFileList();
                            if (chkAutoWatch.Checked) DoImport(f, dt);
                        });
                    }
                }
            };

            _watcher = new FileSystemWatcher(folder, "0. 품의서 *.xlsx")
            {
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite,
                EnableRaisingEvents = chkAutoWatch.Checked
            };
            _watcher.Created += (s, e) => { _pendingFile = e.FullPath; Invoke(() => { _debounceTimer.Stop(); _debounceTimer.Start(); }); };
            _watcher.Changed += (s, e) => { _pendingFile = e.FullPath; Invoke(() => { _debounceTimer.Stop(); _debounceTimer.Start(); }); };

            lblWatch.Text = chkAutoWatch.Checked ? "👀  Đang theo dõi folder Zalo..." : "⏸  Tắt theo dõi tự động.";
        }

        private void ChkAutoWatch_Changed(object sender, EventArgs e)
        {
            if (_watcher != null)
                _watcher.EnableRaisingEvents = chkAutoWatch.Checked;
            if (lblWatch != null)
                lblWatch.Text = chkAutoWatch.Checked ? "👀  Đang theo dõi folder Zalo..." : "⏸  Tắt theo dõi tự động.";
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _watcher?.Dispose();
            _debounceTimer?.Dispose();
            base.OnFormClosed(e);
        }

        // ── Helpers ──
        private void BuildPreviewCols()
        {
            dgvPreview.Columns.Clear();
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "P_Row",    HeaderText = "Row",          Width = 45,  ReadOnly = true });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "P_PO",     HeaderText = "PO No",        Width = 160, ReadOnly = true });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "P_Ecount", HeaderText = "Ecount No",    Width = 110, ReadOnly = true });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "P_Title",  HeaderText = "Title",        Width = 200, ReadOnly = true });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "P_Amount", HeaderText = "Amount",       Width = 110, ReadOnly = true });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "P_VAT",    HeaderText = "VAT",          Width = 90,  ReadOnly = true });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "P_Final",  HeaderText = "Final Amount", Width = 110, ReadOnly = true });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "P_Paid",   HeaderText = "Đã TT",        Width = 110, ReadOnly = true });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "P_Remain", HeaderText = "Còn lại",      Width = 110, ReadOnly = true });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "P_D1",     HeaderText = "Đợt 1",        Width = 100, ReadOnly = true });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "P_D2",     HeaderText = "Đợt 2",        Width = 100, ReadOnly = true });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "P_D3",     HeaderText = "Đợt 3",        Width = 100, ReadOnly = true });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "P_D4",     HeaderText = "Đợt 4",        Width = 100, ReadOnly = true });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "P_FTCash", HeaderText = "FT / Cash",    Width = 90,  ReadOnly = true });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "P_PayDate",HeaderText = "Paid Date",        Width = 90,  ReadOnly = true });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "P_Status", HeaderText = "Status",       Width = 90,  ReadOnly = true });
            StyleAmtCols(dgvPreview, "P_Amount", "P_VAT", "P_Final", "P_Paid", "P_Remain", "P_D1", "P_D2", "P_D3", "P_D4");
        }

        private void BuildDBCols()
        {
            dgvDB.Columns.Clear();
            dgvDB.Columns.Add(new DataGridViewTextBoxColumn { Name = "DB_ID",       HeaderText = "ID",           Width = 55,  ReadOnly = true });
            dgvDB.Columns.Add(new DataGridViewTextBoxColumn { Name = "DB_Date",     HeaderText = "Ngày file",    Width = 90,  ReadOnly = true });
            dgvDB.Columns.Add(new DataGridViewTextBoxColumn { Name = "DB_Row",      HeaderText = "Row",          Width = 45,  ReadOnly = true });
            dgvDB.Columns.Add(new DataGridViewTextBoxColumn { Name = "DB_PO",       HeaderText = "PO No",        Width = 160, ReadOnly = true });
            dgvDB.Columns.Add(new DataGridViewTextBoxColumn { Name = "DB_Ecount",   HeaderText = "Ecount No",    Width = 110, ReadOnly = true });
            dgvDB.Columns.Add(new DataGridViewTextBoxColumn { Name = "DB_Amount",   HeaderText = "Amount",       Width = 110, ReadOnly = true });
            dgvDB.Columns.Add(new DataGridViewTextBoxColumn { Name = "DB_VAT",      HeaderText = "VAT",          Width = 90,  ReadOnly = true });
            dgvDB.Columns.Add(new DataGridViewTextBoxColumn { Name = "DB_Final",    HeaderText = "Final Amount", Width = 110, ReadOnly = true });
            dgvDB.Columns.Add(new DataGridViewTextBoxColumn { Name = "DB_Paid",     HeaderText = "Đã TT",        Width = 110, ReadOnly = true });
            dgvDB.Columns.Add(new DataGridViewTextBoxColumn { Name = "DB_Remain",   HeaderText = "Còn lại",      Width = 110, ReadOnly = true });
            dgvDB.Columns.Add(new DataGridViewTextBoxColumn { Name = "DB_D1",       HeaderText = "Đợt 1",        Width = 100, ReadOnly = true });
            dgvDB.Columns.Add(new DataGridViewTextBoxColumn { Name = "DB_D2",       HeaderText = "Đợt 2",        Width = 100, ReadOnly = true });
            dgvDB.Columns.Add(new DataGridViewTextBoxColumn { Name = "DB_D3",       HeaderText = "Đợt 3",        Width = 100, ReadOnly = true });
            dgvDB.Columns.Add(new DataGridViewTextBoxColumn { Name = "DB_D4",       HeaderText = "Đợt 4",        Width = 100, ReadOnly = true });
            dgvDB.Columns.Add(new DataGridViewTextBoxColumn { Name = "DB_FTCash",   HeaderText = "FT / Cash",    Width = 90,  ReadOnly = true });
            dgvDB.Columns.Add(new DataGridViewTextBoxColumn { Name = "DB_PayDate",  HeaderText = "Paid Date",        Width = 90,  ReadOnly = true });
            dgvDB.Columns.Add(new DataGridViewTextBoxColumn { Name = "DB_Status",   HeaderText = "Status",       Width = 90,  ReadOnly = true });
            dgvDB.Columns.Add(new DataGridViewTextBoxColumn { Name = "DB_Note",     HeaderText = "Ghi chú",      Width = 200, ReadOnly = false });
            dgvDB.Columns.Add(new DataGridViewTextBoxColumn { Name = "DB_ImportAt", HeaderText = "Imported At",  Width = 130, ReadOnly = true });
            dgvDB.Columns.Add(new DataGridViewTextBoxColumn { Name = "DB_ImportID", Visible = false });
            StyleAmtCols(dgvDB, "DB_Amount", "DB_VAT", "DB_Final", "DB_Paid", "DB_Remain", "DB_D1", "DB_D2", "DB_D3", "DB_D4");
            dgvDB.CellEndEdit += DgvDB_NoteEndEdit;
        }

        private void StyleAmtCols(DataGridView dgv, params string[] cols)
        {
            dgv.CellFormatting += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                string cn = dgv.Columns[e.ColumnIndex].Name;
                if (Array.IndexOf(cols, cn) >= 0)
                {
                    e.CellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    // Color logic for Paid/Remain to match frmPayment
                    if (cn.EndsWith("_Paid"))
                    {
                        string val = e.Value?.ToString() ?? "";
                        if (!string.IsNullOrEmpty(val))
                        {
                            e.CellStyle.ForeColor = Color.FromArgb(40, 167, 69); // Green
                            e.CellStyle.Font = new Font(dgv.Font, FontStyle.Bold);
                        }
                    }
                    if (cn.EndsWith("_Remain"))
                    {
                        string val = e.Value?.ToString() ?? "";
                        if (val != "0" && !string.IsNullOrEmpty(val))
                        {
                            e.CellStyle.ForeColor = Color.FromArgb(220, 53, 69); // Red
                            e.CellStyle.Font = new Font(dgv.Font, FontStyle.Bold);
                        }
                    }
                }
            };
        }

        private DataGridView BuildGrid(Color headerColor)
        {
            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = headerColor;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.EnableHeadersVisualStyles = false;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 248, 255);
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(210, 230, 255);
            dgv.DefaultCellStyle.SelectionForeColor = Color.Black;
            return dgv;
        }

        private Button MkBtn(string text, Color color, int x, int y, int w, int h)
        {
            var b = new Button
            {
                Text = text, Location = new Point(x, y), Size = new Size(w, h),
                BackColor = color, ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            b.FlatAppearance.BorderSize = 0;
            return b;
        }

        private void SetBusy(bool busy)
        {
            pbImport.Style = busy ? ProgressBarStyle.Marquee : ProgressBarStyle.Continuous;
            pbImport.MarqueeAnimationSpeed = busy ? 30 : 0;
            btnImportLatest.Enabled = !busy;
            btnImportSelected.Enabled = !busy;
        }

        private void Status(string msg) => lblStatus.Text = msg;

        private static string FormatAmt(decimal? v) =>
            v.HasValue ? v.Value.ToString("#,##0.##") : "";

        private void FilterByPO()
        {
            string kw = (txtSearchPO?.Text ?? "").Trim().ToUpperInvariant();

            // Lọc tab Preview
            foreach (DataGridViewRow row in dgvPreview.Rows)
            {
                string po = (row.Cells["P_PO"].Value?.ToString() ?? "").ToUpperInvariant();
                row.Visible = string.IsNullOrEmpty(kw) || po.Contains(kw);
            }

            // Lọc tab DB
            foreach (DataGridViewRow row in dgvDB.Rows)
            {
                string po = (row.Cells["DB_PO"].Value?.ToString() ?? "").ToUpperInvariant();
                row.Visible = string.IsNullOrEmpty(kw) || po.Contains(kw);
            }
        }

        private void DgvDB_NoteEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (dgvDB.Columns[e.ColumnIndex].Name != "DB_Note") return;
            var row = dgvDB.Rows[e.RowIndex];
            if (!int.TryParse(row.Cells["DB_ImportID"].Value?.ToString(), out int id) || id <= 0) return;
            string note = row.Cells["DB_Note"].Value?.ToString() ?? "";
            try { ZaloImportService.SaveNote(id, note); }
            catch (Exception ex) { Status("Lỗi lưu ghi chú: " + ex.Message); }
        }

        // ── Inner class ──
        private class FileItem
        {
            public string Path { get; }
            public DateTime Date { get; }
            public FileItem(string path, DateTime date) { Path = path; Date = date; }
            public override string ToString() => $"{Date:yyyy-MM-dd}  —  {System.IO.Path.GetFileName(Path)}";
        }
    }
}
