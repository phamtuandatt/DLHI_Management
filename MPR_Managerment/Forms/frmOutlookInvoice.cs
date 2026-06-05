using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;
using MPR_Managerment.Services;

namespace MPR_Managerment.Forms
{
    public class frmOutlookInvoice : Form
    {
        private TextBox txtSaveDir = null!;
        private NumericUpDown numDaysBack = null!;
        private Button btnBrowse = null!;
        private Button btnDownload = null!;
        private Button btnOpenFolder = null!;
        private Button btnStartMonitor = null!;
        private Button btnStopMonitor = null!;
        private Button btnDownloadLinks = null!;
        private Label lblMonitorStatus = null!;
        private Label lblOutlookStatus = null!;
        private ListView lvFiles = null!;
        private Label lblStatus = null!;
        private ProgressBar progressBar = null!;
        private System.Windows.Forms.Timer timerMonitorCheck = null!;
        private System.Windows.Forms.Timer timerOutlookCheck = null!;
        private bool _outlookCheckBusy = false;
        private SplitContainer splitMain = null!;

        // Log panel
        private ListView lvLog = null!;
        private Label lblLogEmailCount = null!;
        private Label lblLogFileCount = null!;
        private Label lblLogLastTime = null!;
        private int _totalEmailsReceived = 0;
        private int _totalFilesDownloaded = 0;

        // Assign panel
        private Panel pnlAssign = null!;
        private ComboBox cboProject = null!;
        private ComboBox cboPO = null!;
        private Button btnAssign = null!;
        private Label lblAssignInfo = null!;

        private readonly ProjectService _projectSvc = new();
        private readonly POService _poSvc = new();
        private List<ProjectInfo> _projects = new();

        public frmOutlookInvoice()
        {
            InitializeComponent();
            OutlookMonitorService.OnNewEvent += OnMonitorEvent;
            UpdateMonitorStatus();
            _ = LoadProjectsAsync();
            _ = UpdateOutlookStatusAsync();
        }

        private void InitializeComponent()
        {
            Text = "Tải Hóa Đơn từ Outlook";
            Size = new Size(1000, 720);
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;

            // ── Header trạng thái theo dõi ──────────────────────────────────────
            var pnlHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 52,
                BackColor = Color.FromArgb(30, 30, 45)
            };

            var lblTitle = new Label
            {
                Text = "📧  Hóa Đơn Outlook",
                Font = new Font("Segoe UI", 13, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(14, 12),
                AutoSize = true
            };

            lblMonitorStatus = new Label
            {
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(280, 12)
            };

            // Dấu phân cách
            var lblSep = new Label
            {
                Text = "│",
                Font = new Font("Segoe UI", 14),
                ForeColor = Color.FromArgb(80, 80, 100),
                AutoSize = true,
                Location = new Point(580, 12)
            };

            lblOutlookStatus = new Label
            {
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(605, 15),
                Text = "Outlook: ..."
            };

            pnlHeader.Controls.AddRange(new Control[] { lblTitle, lblMonitorStatus, lblSep, lblOutlookStatus });

            // ── Panel cấu hình ──────────────────────────────────────────────────
            var pnlTop = new Panel { Dock = DockStyle.Top, Height = 130, Padding = new Padding(10) };

            var lblDir = new Label { Text = "Thư mục lưu:", Location = new Point(10, 15), AutoSize = true };
            txtSaveDir = new TextBox { Location = new Point(110, 12), Width = 550, Text = OutlookMonitorService.SaveDir };
            btnBrowse = new Button { Text = "...", Location = new Point(668, 11), Width = 40 };
            btnBrowse.Click += BtnBrowse_Click;

            var lblDays = new Label { Text = "Số ngày nhìn lại:", Location = new Point(10, 50), AutoSize = true };
            numDaysBack = new NumericUpDown { Location = new Point(130, 47), Width = 70, Minimum = 1, Maximum = 365, Value = 30 };

            btnDownload = new Button
            {
                Text = "Tải thủ công",
                Location = new Point(350, 44), Width = 130, Height = 30,
                BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
            };
            btnDownload.Click += BtnDownload_Click;

            btnOpenFolder = new Button { Text = "Mở thư mục", Location = new Point(490, 44), Width = 110, Height = 30 };
            btnOpenFolder.Click += (s, e) => OutlookInvoiceService.OpenSaveFolder(txtSaveDir.Text);

            var btnClassify = new Button
            {
                Text = "📁 Phân loại hóa đơn",
                Location = new Point(610, 44), Width = 160, Height = 30,
                BackColor = Color.FromArgb(142, 68, 173), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
            };
            btnClassify.Click += (s, e) =>
            {
                var selected = lvFiles.SelectedItems.Cast<ListViewItem>()
                    .Select(i => i.Tag?.ToString()).Where(f => f != null && File.Exists(f)).ToList();
                var frm = new frmInvoiceClassifier
                {
                    PreSelectedFiles = selected.Count > 0 ? selected! : new List<string>()
                };
                frm.Show(this);
            };

            var separator = new Label
            {
                Text = "── Tự động theo dõi email mới (chạy liên tục, kể cả khi tắt app) ──────",
                Location = new Point(10, 84), AutoSize = true, ForeColor = Color.Gray
            };

            btnStartMonitor = new Button
            {
                Text = "▶ Bật Automatic",
                Location = new Point(10, 104), Width = 155, Height = 28,
                BackColor = Color.FromArgb(39, 174, 96), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
            };
            btnStartMonitor.Click += BtnStartMonitor_Click;

            btnStopMonitor = new Button
            {
                Text = "■ Tắt Automatic",
                Location = new Point(175, 104), Width = 140, Height = 28,
                BackColor = Color.FromArgb(192, 57, 43), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
            };
            btnStopMonitor.Click += BtnStopMonitor_Click;

            btnDownloadLinks = new Button
            {
                Text = "🔗 Tải từ link email",
                Location = new Point(330, 104), Width = 155, Height = 28,
                BackColor = Color.FromArgb(41, 128, 185), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
            };
            btnDownloadLinks.Click += BtnDownloadLinks_Click;

            pnlTop.Controls.AddRange(new Control[] {
                lblDir, txtSaveDir, btnBrowse,
                lblDays, numDaysBack, btnDownload, btnOpenFolder, btnClassify,
                separator, btnStartMonitor, btnStopMonitor, btnDownloadLinks
            });

            // ── SplitContainer: trái = danh sách file, phải = log monitor ────────
            splitMain = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterWidth = 5,
                BackColor = Color.FromArgb(200, 210, 230)
            };

            // ── ListView (trái) ──────────────────────────────────────────────────
            lvFiles = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                MultiSelect = true
            };
            lvFiles.Columns.Add("Tên file", 220);
            lvFiles.Columns.Add("Tiêu đề email", 240);
            lvFiles.Columns.Add("Người gửi", 140);
            lvFiles.Columns.Add("Thời gian", 120);
            lvFiles.Columns.Add("Dự án", 100);
            lvFiles.Columns.Add("PO No", 110);

            lvFiles.DoubleClick += (s, e) =>
            {
                if (lvFiles.SelectedItems.Count > 0)
                {
                    string? path = lvFiles.SelectedItems[0].Tag?.ToString();
                    if (path != null && File.Exists(path))
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
                }
            };
            lvFiles.SelectedIndexChanged += LvFiles_SelectedIndexChanged;

            splitMain.Panel1.Controls.Add(lvFiles);

            // ── Panel log monitor (phải) ─────────────────────────────────────────
            var pnlLog = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(24, 26, 38) };

            var pnlLogHeader = new Panel
            {
                Dock = DockStyle.Top, Height = 40,
                BackColor = Color.FromArgb(35, 37, 55),
                Padding = new Padding(10, 0, 10, 0)
            };
            var lblLogTitle = new Label
            {
                Text = "📋  Nhật ký theo dõi tự động",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(180, 190, 220),
                Dock = DockStyle.Left,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoSize = false, Width = 250
            };
            var btnClearLog = new Button
            {
                Text = "Xóa log",
                Dock = DockStyle.Right,
                Width = 75,
                Height = 28,
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.FromArgb(150, 160, 190),
                BackColor = Color.FromArgb(50, 52, 75),
                Font = new Font("Segoe UI", 8)
            };
            btnClearLog.FlatAppearance.BorderColor = Color.FromArgb(70, 75, 110);
            btnClearLog.Click += (s, e) => { lvLog.Items.Clear(); _totalEmailsReceived = 0; _totalFilesDownloaded = 0; UpdateLogStats(); };
            pnlLogHeader.Controls.Add(lblLogTitle);
            pnlLogHeader.Controls.Add(btnClearLog);

            // Thống kê
            var pnlLogStats = new Panel
            {
                Dock = DockStyle.Top, Height = 52,
                BackColor = Color.FromArgb(30, 32, 48),
                Padding = new Padding(10, 6, 10, 6)
            };
            lblLogEmailCount = new Label
            {
                Text = "📧  Email nhận:  0",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(100, 200, 255),
                Location = new Point(10, 8), AutoSize = true
            };
            lblLogFileCount = new Label
            {
                Text = "📄  File tải:  0",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(100, 220, 140),
                Location = new Point(160, 8), AutoSize = true
            };
            lblLogLastTime = new Label
            {
                Text = "🕐  Lần cuối:  —",
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = Color.FromArgb(150, 155, 180),
                Location = new Point(300, 10), AutoSize = true
            };
            pnlLogStats.Controls.AddRange(new Control[] { lblLogEmailCount, lblLogFileCount, lblLogLastTime });

            lvLog = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = false,
                BackColor = Color.FromArgb(24, 26, 38),
                ForeColor = Color.FromArgb(210, 215, 235),
                Font = new Font("Segoe UI", 8.5f),
                BorderStyle = BorderStyle.None,
                HeaderStyle = ColumnHeaderStyle.Nonclickable
            };
            lvLog.Columns.Add("Thời gian", 80);
            lvLog.Columns.Add("Loại", 90);
            lvLog.Columns.Add("Tiêu đề / Người gửi", 160);
            lvLog.Columns.Add("File", 55);
            lvLog.Columns.Add("Chi tiết", 200);

            pnlLog.Controls.Add(lvLog);
            pnlLog.Controls.Add(pnlLogStats);
            pnlLog.Controls.Add(pnlLogHeader);

            splitMain.Panel2.Controls.Add(pnlLog);

            // ── Panel gán dự án / PO ────────────────────────────────────────────
            pnlAssign = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 48,
                BackColor = Color.FromArgb(240, 245, 255),
                Padding = new Padding(8, 6, 8, 6),
                Visible = false
            };
            pnlAssign.Paint += (s, e) =>
            {
                e.Graphics.DrawLine(Pens.LightSteelBlue, 0, 0, pnlAssign.Width, 0);
            };

            lblAssignInfo = new Label
            {
                Location = new Point(8, 14), AutoSize = true,
                ForeColor = Color.DimGray,
                Font = new Font("Segoe UI", 8.5f)
            };

            var lblProj = new Label { Text = "Dự án:", Location = new Point(160, 14), AutoSize = true };
            cboProject = new ComboBox
            {
                Location = new Point(210, 10), Width = 160, DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboProject.SelectedIndexChanged += CboProject_SelectedIndexChanged;

            var lblPO = new Label { Text = "PO No:", Location = new Point(380, 14), AutoSize = true };
            cboPO = new ComboBox
            {
                Location = new Point(425, 10), Width = 160, DropDownStyle = ComboBoxStyle.DropDownList
            };

            btnAssign = new Button
            {
                Text = "📌 Cập nhật",
                Location = new Point(600, 8), Width = 110, Height = 30,
                BackColor = Color.FromArgb(39, 174, 96), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
            };
            btnAssign.Click += BtnAssign_Click;

            pnlAssign.Controls.AddRange(new Control[] {
                lblAssignInfo, lblProj, cboProject, lblPO, cboPO, btnAssign
            });

            // ── Status bar ──────────────────────────────────────────────────────
            var pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 35 };
            progressBar = new ProgressBar { Location = new Point(10, 8), Width = 200, Height = 18, Style = ProgressBarStyle.Marquee, Visible = false };
            lblStatus = new Label { Location = new Point(220, 10), AutoSize = true, Text = "Sẵn sàng." };
            pnlBottom.Controls.AddRange(new Control[] { progressBar, lblStatus });

            Controls.Add(splitMain);
            Controls.Add(pnlAssign);
            Controls.Add(pnlTop);
            Controls.Add(pnlHeader);
            Controls.Add(pnlBottom);

            timerMonitorCheck = new System.Windows.Forms.Timer { Interval = 3000, Enabled = true };
            timerMonitorCheck.Tick += (s, e) => UpdateMonitorStatus();

            timerOutlookCheck = new System.Windows.Forms.Timer { Interval = 5000, Enabled = true };
            timerOutlookCheck.Tick += (s, e) => _ = UpdateOutlookStatusAsync();

            LoadExistingLog();
        }

        // ── Load dữ liệu ────────────────────────────────────────────────────────

        private async Task LoadProjectsAsync()
        {
            await Task.Run(() => { _projects = _projectSvc.GetAll(); });
            if (IsDisposed) return;
            Invoke(() =>
            {
                cboProject.Items.Clear();
                cboProject.Items.Add(new ProjectComboItem(null, "-- Chọn dự án --"));
                foreach (var p in _projects)
                    cboProject.Items.Add(new ProjectComboItem(p, p.ProjectCode));
                if (cboProject.Items.Count > 0) cboProject.SelectedIndex = 0;
            });
        }

        private async void CboProject_SelectedIndexChanged(object? sender, EventArgs e)
        {
            cboPO.Items.Clear();
            cboPO.Items.Add(new POComboItem(null, "-- Chọn PO --"));
            cboPO.SelectedIndex = 0;

            if (cboProject.SelectedItem is not ProjectComboItem { Project: not null } sel) return;

            var dt = await _poSvc.GetPOByProjectCode(sel.Project.ProjectCode);
            foreach (DataRow row in dt.Rows)
                cboPO.Items.Add(new POComboItem((int)row["PO_ID"], row["PONo"]?.ToString() ?? ""));
        }

        // ── ListView selection ───────────────────────────────────────────────────

        private void LvFiles_SelectedIndexChanged(object? sender, EventArgs e)
        {
            int count = lvFiles.SelectedItems.Count;
            pnlAssign.Visible = count > 0;
            if (count > 0)
                lblAssignInfo.Text = count == 1
                    ? $"1 file được chọn"
                    : $"{count} file được chọn";
        }

        // ── Gán dự án / PO và di chuyển file ───────────────────────────────────

        private async void BtnAssign_Click(object? sender, EventArgs e)
        {
            if (cboProject.SelectedItem is not ProjectComboItem { Project: not null } projSel)
            {
                MessageBox.Show("Vui lòng chọn dự án.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var project = projSel.Project;

            string? poNo = null;
            if (cboPO.SelectedItem is POComboItem { PONo: not null and not "" } poSel)
                poNo = poSel.PONo;

            var selectedItems = lvFiles.SelectedItems.Cast<ListViewItem>().ToList();
            if (selectedItems.Count == 0) return;

            // Xác định thư mục đích
            string destDir = string.IsNullOrWhiteSpace(project.INV_Link)
                ? txtSaveDir.Text
                : project.INV_Link;

            try { Directory.CreateDirectory(destDir); }
            catch (Exception ex)
            {
                MessageBox.Show($"Không tạo được thư mục:\n{destDir}\n{ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int success = 0, fail = 0;
            var errors = new List<string>();

            await Task.Run(() =>
            {
                foreach (var item in selectedItems)
                {
                    string? srcPath = item.Tag?.ToString();
                    if (srcPath == null || !File.Exists(srcPath))
                    {
                        fail++;
                        errors.Add($"{item.Text}: file không tồn tại");
                        continue;
                    }

                    try
                    {
                        string ext = Path.GetExtension(srcPath);
                        string origName = Path.GetFileNameWithoutExtension(srcPath);

                        // Tên mới: <PONo>_<ProjectCode>_<origName>.pdf (hoặc bỏ PONo nếu chưa chọn)
                        string prefix = poNo != null
                            ? $"{poNo}_{project.ProjectCode}_"
                            : $"{project.ProjectCode}_";

                        // Tránh prefix trùng nếu đã gán trước
                        string newName = origName.StartsWith(prefix)
                            ? origName + ext
                            : prefix + origName + ext;

                        string destPath = Path.Combine(destDir, newName);
                        // Tránh ghi đè
                        if (File.Exists(destPath))
                            destPath = Path.Combine(destDir, prefix + origName + $"_{DateTimeOffset.Now.ToUnixTimeSeconds()}" + ext);

                        File.Move(srcPath, destPath);

                        // Cập nhật ListView trên UI thread
                        Invoke(() =>
                        {
                            item.Text = Path.GetFileName(destPath);
                            item.Tag = destPath;
                            item.SubItems[4].Text = project.ProjectCode;
                            item.SubItems[5].Text = poNo ?? "";
                            item.BackColor = Color.FromArgb(224, 240, 255);
                        });
                        success++;
                    }
                    catch (Exception ex)
                    {
                        fail++;
                        errors.Add($"{item.Text}: {ex.Message}");
                    }
                }
            });

            string summary = $"Cập nhật xong: {success} file → {destDir}";
            if (fail > 0) summary += $"\n{fail} file thất bại.";
            lblStatus.Text = summary.Split('\n')[0];

            if (errors.Count > 0)
                MessageBox.Show(string.Join("\n", errors), "Một số file lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
                MessageBox.Show($"Đã đổi tên và di chuyển {success} file vào:\n{destDir}", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ── Helpers ─────────────────────────────────────────────────────────────

        private void LoadExistingLog()
        {
            var logs = OutlookMonitorService.ReadLog(100);
            foreach (var evt in logs)
            {
                if (evt.Event == "downloaded" && evt.Files != null)
                {
                    foreach (var file in evt.Files)
                        AddFileToList(file, evt.Subject ?? "", evt.Sender ?? "", evt.Time, fromLog: true);
                    _totalEmailsReceived++;
                    _totalFilesDownloaded += evt.Files.Count;
                    AddLogEntry(evt.Time, "✅ Tải được",
                        $"{evt.Subject ?? ""} / {evt.Sender ?? ""}",
                        evt.Files.Count.ToString(),
                        string.Join(", ", evt.Files.Select(Path.GetFileName)),
                        Color.FromArgb(40, 70, 45));
                }
                else if (evt.Event == "started")
                {
                    AddLogEntry(evt.Time, "▶ Khởi động", "Monitor bắt đầu lắng nghe", "", "", Color.FromArgb(30, 55, 85));
                }
            }
            UpdateLogStats();
        }

        private void AddFileToList(string filePath, string subject, string sender, string time, bool fromLog = false)
        {
            if (InvokeRequired) { Invoke(() => AddFileToList(filePath, subject, sender, time, fromLog)); return; }
            var item = new ListViewItem(Path.GetFileName(filePath));
            item.SubItems.Add(subject);
            item.SubItems.Add(sender);
            item.SubItems.Add(time);
            item.SubItems.Add(""); // Dự án
            item.SubItems.Add(""); // PO No
            item.Tag = filePath;
            if (!fromLog) item.BackColor = Color.FromArgb(232, 255, 232);
            lvFiles.Items.Insert(0, item);
        }

        private void OnMonitorEvent(MonitorEvent evt)
        {
            if (InvokeRequired) { Invoke(() => OnMonitorEvent(evt)); return; }
            if (evt.Event == "downloaded" && evt.Files != null)
            {
                foreach (var file in evt.Files)
                    AddFileToList(file, evt.Subject ?? "", evt.Sender ?? "", evt.Time);
                _totalEmailsReceived++;
                _totalFilesDownloaded += evt.Files.Count;
                UpdateLogStats();
                AddLogEntry(evt.Time, "✅ Tải được",
                    $"{evt.Subject ?? ""} / {evt.Sender ?? ""}",
                    evt.Files.Count.ToString(),
                    string.Join(", ", evt.Files.Select(Path.GetFileName)),
                    Color.FromArgb(60, 100, 60));
                lblStatus.Text = $"[{DateTime.Now:HH:mm:ss}] Tải {evt.Files.Count} file từ: {evt.Subject}";
            }
            else if (evt.Event == "started")
            {
                AddLogEntry(evt.Time, "▶ Khởi động", "Monitor bắt đầu lắng nghe", "", "", Color.FromArgb(40, 80, 120));
            }
            else if (evt.Event == "error")
            {
                AddLogEntry(evt.Time, "⚠ Lỗi", evt.Message ?? "", "", "", Color.FromArgb(100, 40, 40));
                lblStatus.Text = $"Lỗi: {evt.Message}";
            }
        }

        private void AddLogEntry(string time, string type, string subject, string fileCount, string detail, Color rowColor)
        {
            if (InvokeRequired) { Invoke(() => AddLogEntry(time, type, subject, fileCount, detail, rowColor)); return; }
            string displayTime = time;
            if (DateTime.TryParse(time, out var dt)) displayTime = dt.ToString("HH:mm:ss");

            var item = new ListViewItem(displayTime);
            item.SubItems.Add(type);
            item.SubItems.Add(subject.Length > 45 ? subject[..42] + "..." : subject);
            item.SubItems.Add(fileCount);
            item.SubItems.Add(detail.Length > 60 ? detail[..57] + "..." : detail);
            item.BackColor = rowColor;
            item.ForeColor = Color.FromArgb(210, 215, 235);
            lvLog.Items.Insert(0, item);

            // Giữ tối đa 200 dòng
            while (lvLog.Items.Count > 200)
                lvLog.Items.RemoveAt(lvLog.Items.Count - 1);
        }

        private void UpdateLogStats()
        {
            if (InvokeRequired) { Invoke(UpdateLogStats); return; }
            lblLogEmailCount.Text = $"📧  Email nhận:  {_totalEmailsReceived}";
            lblLogFileCount.Text  = $"📄  File tải:  {_totalFilesDownloaded}";
            lblLogLastTime.Text   = $"🕐  Lần cuối:  {DateTime.Now:HH:mm dd/MM}";
        }

        private bool _restarting = false;

        private async void UpdateMonitorStatus()
        {
            bool persistent = OutlookMonitorService.IsPersistentEnabled;
            bool running    = OutlookMonitorService.IsRunning;

            btnStartMonitor.Enabled = !persistent;
            btnStopMonitor.Enabled  = persistent;

            if (persistent && running)
            {
                lblMonitorStatus.ForeColor = Color.FromArgb(0, 230, 120);
                lblMonitorStatus.Text = "● Automatic";
                _restarting = false;
            }
            else if (persistent && !running)
            {
                // Tự động restart nếu bị crash
                if (!_restarting)
                {
                    _restarting = true;
                    lblMonitorStatus.ForeColor = Color.FromArgb(255, 200, 0);
                    lblMonitorStatus.Text = "⟳ Automatic (đang khởi động lại...)";
                    string result = await OutlookMonitorService.StartAsync(OutlookMonitorService.SaveDir);
                    lblStatus.Text = result;
                    // Nếu vẫn không chạy được, hiện lỗi
                    if (!OutlookMonitorService.IsRunning)
                    {
                        lblMonitorStatus.ForeColor = Color.FromArgb(255, 120, 0);
                        lblMonitorStatus.Text = "⚠ Automatic (lỗi script)";
                        _restarting = false;
                    }
                }
            }
            else
            {
                lblMonitorStatus.ForeColor = Color.FromArgb(255, 80, 80);
                lblMonitorStatus.Text = "● OFF";
                _restarting = false;
            }
        }

        // ── Button handlers ─────────────────────────────────────────────────────

        private async void BtnStartMonitor_Click(object? sender, EventArgs e)
        {
            btnStartMonitor.Enabled = false;
            lblStatus.Text = "Đang bật chế độ Automatic...";
            UpdateMonitorStatus();
            string msg = await OutlookMonitorService.EnablePersistentAsync(txtSaveDir.Text);
            lblStatus.Text = msg;
            UpdateMonitorStatus();
        }

        private void BtnStopMonitor_Click(object? sender, EventArgs e)
        {
            var confirm = MessageBox.Show(
                "Tắt chế độ Automatic sẽ dừng theo dõi email ngay cả khi khởi động lại máy.\nBạn có chắc không?",
                "Xác nhận tắt Automatic", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes) return;

            lblStatus.Text = OutlookMonitorService.DisablePersistent();
            UpdateMonitorStatus();
        }

        private void BtnBrowse_Click(object? sender, EventArgs e)
        {
            using var dlg = new FolderBrowserDialog { SelectedPath = txtSaveDir.Text, Description = "Chọn thư mục lưu file hóa đơn" };
            if (dlg.ShowDialog() == DialogResult.OK)
                txtSaveDir.Text = dlg.SelectedPath;
        }

        private async void BtnDownload_Click(object? sender, EventArgs e)
        {
            btnDownload.Enabled = false;
            btnDownloadLinks.Enabled = false;
            progressBar.Visible = true;
            var toast = ToastHelper.Attach(this);
            toast.Show("⏳ Đang tải hóa đơn đính kèm và link email...");

            string saveDir = txtSaveDir.Text;
            int daysBack = (int)numDaysBack.Value;

            var progressAttach = new Progress<string>(msg => { if (InvokeRequired) Invoke(() => lblStatus.Text = msg); else lblStatus.Text = msg; });
            var taskAttach = OutlookInvoiceService.DownloadInvoicesAsync(saveDir: saveDir, daysBack: daysBack, progress: progressAttach);
            var taskLink   = InvoiceLinkDownloaderService.DownloadFromLinksAsync(saveDir: saveDir, daysBack: daysBack);

            await Task.WhenAll(taskAttach, taskLink);

            var attachResult = taskAttach.Result;
            var linkResult   = taskLink.Result;

            toast.Hide();
            progressBar.Visible = false;
            btnDownload.Enabled = true;
            btnDownloadLinks.Enabled = true;

            if (attachResult.Success)
                foreach (var f in attachResult.Files)
                    AddFileToList(f.FilePath, f.Subject, f.Sender, f.Received);

            if (linkResult.Success)
                foreach (var r in linkResult.Results)
                    if (r.Status == "downloaded" && r.File != null)
                        AddFileToList(r.File, r.Subject ?? "", r.Sender ?? "", DateTime.Now.ToString());

            var pending = linkResult.Results?.FindAll(r => r.Status == "pending_manual") ?? new();
            if (pending.Count > 0)
            {
                string msg = $"⚠ {pending.Count} email cần xử lý thủ công:\n\n";
                foreach (var p in pending)
                    msg += $"• {p.Subject}\n  Mã: {p.LookupCode ?? "không trích xuất được"} ({p.Provider})\n";
                MessageBox.Show(msg, "Cần xử lý thủ công", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            int attachCount  = attachResult.FilesDownloaded;
            int linkCount    = linkResult.Summary?.Downloaded ?? 0;
            int pendingCount = linkResult.Summary?.PendingManual ?? 0;
            lblStatus.Text = $"Hoàn thành: {attachCount} file đính kèm + {linkCount} file từ link.{(pendingCount > 0 ? $" {pendingCount} chờ thủ công." : "")}";
        }

        private async void BtnDownloadLinks_Click(object? sender, EventArgs e)
        {
            btnDownloadLinks.Enabled = false;
            progressBar.Visible = true;
            lblStatus.Text = "Đang quét email chứa link hóa đơn...";

            var result = await InvoiceLinkDownloaderService.DownloadFromLinksAsync(
                saveDir: txtSaveDir.Text, daysBack: (int)numDaysBack.Value,
                progress: new Progress<string>(msg => lblStatus.Text = msg));

            progressBar.Visible = false;
            btnDownloadLinks.Enabled = true;

            if (!result.Success) { MessageBox.Show(result.Error, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

            foreach (var r in result.Results)
                if (r.Status == "downloaded" && r.File != null)
                    AddFileToList(r.File, r.Subject ?? "", r.Sender ?? "", DateTime.Now.ToString());

            var pending = result.Results.FindAll(r => r.Status == "pending_manual");
            if (pending.Count > 0)
            {
                string msg = $"⚠ {pending.Count} email cần xử lý thủ công:\n\n";
                foreach (var p in pending)
                    msg += $"• {p.Subject}\n  Mã: {p.LookupCode ?? "không trích xuất được"} ({p.Provider})\n";
                MessageBox.Show(msg, "Cần xử lý thủ công", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            lblStatus.Text = $"Hoàn thành: {result.Summary.Downloaded} tải được, {result.Summary.PendingManual} chờ thủ công.";
        }

        private async Task UpdateOutlookStatusAsync()
        {
            if (_outlookCheckBusy || IsDisposed) return;
            _outlookCheckBusy = true;
            try
            {
                var status = await Task.Run(() => OutlookStatusChecker.Check());
                if (IsDisposed) return;
                Invoke(() =>
                {
                    lblOutlookStatus.Text = status.Text;
                    lblOutlookStatus.ForeColor = status.Color;
                });
            }
            finally { _outlookCheckBusy = false; }
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            BeginInvoke(() =>
            {
                try
                {
                    splitMain.Panel1MinSize = 400;
                    splitMain.Panel2MinSize = 300;
                    splitMain.SplitterDistance = Math.Max(400, splitMain.Width - 380);
                }
                catch { }
            });
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            OutlookMonitorService.OnNewEvent -= OnMonitorEvent;
            timerMonitorCheck.Dispose();
            timerOutlookCheck.Dispose();
            base.OnFormClosed(e);
        }

        // ── Helper wrapper classes cho ComboBox ─────────────────────────────────

        private class ProjectComboItem(ProjectInfo? project, string display)
        {
            public ProjectInfo? Project { get; } = project;
            public override string ToString() => display;
        }

        private class POComboItem(int? poId, string poNo)
        {
            public int? POId { get; } = poId;
            public string PONo { get; } = poNo;
            public override string ToString() => poNo;
        }
    }
}
