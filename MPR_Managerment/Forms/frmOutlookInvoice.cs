using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
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
        private Label lblMonitorStatus = null!;
        private ListView lvFiles = null!;
        private Label lblStatus = null!;
        private ProgressBar progressBar = null!;
        private System.Windows.Forms.Timer timerMonitorCheck = null!;

        public frmOutlookInvoice()
        {
            InitializeComponent();
            OutlookMonitorService.OnNewEvent += OnMonitorEvent;
            UpdateMonitorStatus();
        }

        private void InitializeComponent()
        {
            Text = "Tải Hóa Đơn từ Outlook";
            Size = new Size(860, 600);
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;

            // --- Panel cấu hình ---
            var pnlTop = new Panel { Dock = DockStyle.Top, Height = 130, Padding = new Padding(10) };

            var lblDir = new Label { Text = "Thư mục lưu:", Location = new Point(10, 15), AutoSize = true };
            txtSaveDir = new TextBox { Location = new Point(110, 12), Width = 480, Text = OutlookMonitorService.SaveDir };
            btnBrowse = new Button { Text = "...", Location = new Point(598, 11), Width = 40 };
            btnBrowse.Click += BtnBrowse_Click;

            var lblDays = new Label { Text = "Số ngày nhìn lại:", Location = new Point(10, 50), AutoSize = true };
            numDaysBack = new NumericUpDown { Location = new Point(130, 47), Width = 70, Minimum = 1, Maximum = 365, Value = 30 };

            btnDownload = new Button
            {
                Text = "Tải thủ công",
                Location = new Point(350, 44),
                Width = 130,
                Height = 30,
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnDownload.Click += BtnDownload_Click;

            btnOpenFolder = new Button { Text = "Mở thư mục", Location = new Point(490, 44), Width = 110, Height = 30 };
            btnOpenFolder.Click += (s, e) => OutlookInvoiceService.OpenSaveFolder(txtSaveDir.Text);

            var btnClassify = new Button
            {
                Text = "📁 Phân loại hóa đơn",
                Location = new Point(610, 44),
                Width = 160, Height = 30,
                BackColor = Color.FromArgb(142, 68, 173),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnClassify.Click += (s, e) =>
            {
                var selectedFiles = lvFiles.SelectedItems
                    .Cast<ListViewItem>()
                    .Select(i => i.Tag?.ToString())
                    .Where(f => f != null && File.Exists(f))
                    .ToList();

                var frm = new frmInvoiceClassifier
                {
                    PreSelectedFiles = selectedFiles.Count > 0
                        ? selectedFiles!
                        : new System.Collections.Generic.List<string>()
                };
                frm.Show(this);
            };

            // --- Monitor controls ---
            var separator = new Label
            {
                Text = "── Tự động theo dõi email mới ──────────────────────────────",
                Location = new Point(10, 84),
                AutoSize = true,
                ForeColor = Color.Gray
            };

            btnStartMonitor = new Button
            {
                Text = "▶ Bắt đầu theo dõi",
                Location = new Point(10, 104),
                Width = 155,
                Height = 28,
                BackColor = Color.FromArgb(39, 174, 96),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnStartMonitor.Click += BtnStartMonitor_Click;

            btnStopMonitor = new Button
            {
                Text = "■ Dừng theo dõi",
                Location = new Point(175, 104),
                Width = 140,
                Height = 28,
                BackColor = Color.FromArgb(192, 57, 43),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnStopMonitor.Click += BtnStopMonitor_Click;

            lblMonitorStatus = new Label
            {
                Location = new Point(330, 109),
                AutoSize = true,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };

            pnlTop.Controls.AddRange(new Control[] {
                lblDir, txtSaveDir, btnBrowse,
                lblDays, numDaysBack, btnDownload, btnOpenFolder, btnClassify,
                separator, btnStartMonitor, btnStopMonitor, lblMonitorStatus
            });

            // --- ListView ---
            lvFiles = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true
            };
            lvFiles.Columns.Add("Tên file", 230);
            lvFiles.Columns.Add("Tiêu đề email", 280);
            lvFiles.Columns.Add("Người gửi", 150);
            lvFiles.Columns.Add("Thời gian", 130);

            lvFiles.DoubleClick += (s, e) =>
            {
                if (lvFiles.SelectedItems.Count > 0)
                {
                    string? path = lvFiles.SelectedItems[0].Tag?.ToString();
                    if (path != null && File.Exists(path))
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
                }
            };

            // --- Status bar ---
            var pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 35 };
            progressBar = new ProgressBar { Location = new Point(10, 8), Width = 200, Height = 18, Style = ProgressBarStyle.Marquee, Visible = false };
            lblStatus = new Label { Location = new Point(220, 10), AutoSize = true, Text = "Sẵn sàng." };
            pnlBottom.Controls.AddRange(new Control[] { progressBar, lblStatus });

            Controls.Add(lvFiles);
            Controls.Add(pnlTop);
            Controls.Add(pnlBottom);

            // Timer kiểm tra trạng thái monitor mỗi 3 giây
            timerMonitorCheck = new System.Windows.Forms.Timer { Interval = 3000, Enabled = true };
            timerMonitorCheck.Tick += (s, e) => UpdateMonitorStatus();

            // Load log cũ khi mở form
            LoadExistingLog();
        }

        private void LoadExistingLog()
        {
            var logs = OutlookMonitorService.ReadLog(100);
            foreach (var evt in logs)
            {
                if (evt.Event == "downloaded" && evt.Files != null)
                {
                    foreach (var file in evt.Files)
                        AddFileToList(file, evt.Subject ?? "", evt.Sender ?? "", evt.Time, fromLog: true);
                }
            }
        }

        private void AddFileToList(string filePath, string subject, string sender, string time, bool fromLog = false)
        {
            if (InvokeRequired)
            {
                Invoke(() => AddFileToList(filePath, subject, sender, time, fromLog));
                return;
            }
            var item = new ListViewItem(Path.GetFileName(filePath));
            item.SubItems.Add(subject);
            item.SubItems.Add(sender);
            item.SubItems.Add(time);
            item.Tag = filePath;
            if (!fromLog) item.BackColor = Color.FromArgb(232, 255, 232); // highlight xanh nếu mới
            lvFiles.Items.Insert(0, item); // thêm lên đầu
        }

        private void OnMonitorEvent(MonitorEvent evt)
        {
            if (InvokeRequired) { Invoke(() => OnMonitorEvent(evt)); return; }

            if (evt.Event == "downloaded" && evt.Files != null)
            {
                foreach (var file in evt.Files)
                    AddFileToList(file, evt.Subject ?? "", evt.Sender ?? "", evt.Time);

                lblStatus.Text = $"[{DateTime.Now:HH:mm:ss}] Tải {evt.Files.Count} file từ: {evt.Subject}";
            }
            else if (evt.Event == "error")
            {
                lblStatus.Text = $"Lỗi: {evt.Message}";
            }
        }

        private void UpdateMonitorStatus()
        {
            bool running = OutlookMonitorService.IsRunning;
            btnStartMonitor.Enabled = !running;
            btnStopMonitor.Enabled = running;
            if (running)
            {
                lblMonitorStatus.ForeColor = Color.FromArgb(39, 174, 96);
                lblMonitorStatus.Text = "● Đang theo dõi";
            }
            else
            {
                lblMonitorStatus.ForeColor = Color.Gray;
                lblMonitorStatus.Text = "○ Chưa theo dõi";
            }
        }

        private async void BtnStartMonitor_Click(object? sender, EventArgs e)
        {
            btnStartMonitor.Enabled = false;
            OutlookMonitorService.SaveDir = txtSaveDir.Text;
            string msg = await OutlookMonitorService.StartAsync(txtSaveDir.Text);
            lblStatus.Text = msg;
            UpdateMonitorStatus();
        }

        private void BtnStopMonitor_Click(object? sender, EventArgs e)
        {
            string msg = OutlookMonitorService.Stop();
            lblStatus.Text = msg;
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
            progressBar.Visible = true;
            var progress = new Progress<string>(msg => lblStatus.Text = msg);
            var result = await OutlookInvoiceService.DownloadInvoicesAsync(
                saveDir: txtSaveDir.Text, daysBack: (int)numDaysBack.Value, progress: progress);
            progressBar.Visible = false;
            btnDownload.Enabled = true;

            if (!result.Success) { MessageBox.Show(result.Error, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

            foreach (var f in result.Files)
                AddFileToList(f.FilePath, f.Subject, f.Sender, f.Received);

            lblStatus.Text = $"Hoàn thành: {result.FilesDownloaded} file từ {result.EmailsProcessed} email.";
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            OutlookMonitorService.OnNewEvent -= OnMonitorEvent;
            timerMonitorCheck.Dispose();
            base.OnFormClosed(e);
        }
    }
}
