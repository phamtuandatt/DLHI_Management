using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using MPR_Managerment.Services;

namespace MPR_Managerment.Forms
{
    public class frmInvoiceClassifier : Form
    {
        private TextBox txtScanDir = null!;
        private TextBox txtUnclassifiedBase = null!;
        private Button btnBrowseScan = null!;
        private Button btnBrowseUnclassified = null!;
        private Button btnScanFolder = null!;
        private Button btnClassifySelected = null!;
        private Button btnOpenDest = null!;
        private Button btnCutToFolder = null!;
        private ListView lvResults = null!;
        private Label lblStatus = null!;
        private ProgressBar progressBar = null!;
        private Label lblSummary = null!;

        // Danh sách file được chọn từ form Outlook (nếu mở từ đó)
        public List<string> PreSelectedFiles { get; set; } = new();

        public frmInvoiceClassifier()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            Text = "Phân loại Hóa Đơn → Thư mục Dự án";
            Size = new Size(1000, 640);
            StartPosition = FormStartPosition.CenterParent;

            // ── Panel trên ────────────────────────────────────────────────────
            var pnlTop = new Panel { Dock = DockStyle.Top, Height = 110, Padding = new Padding(10) };

            var lblScan = new Label { Text = "Thư mục PDF nguồn:", Location = new Point(10, 15), AutoSize = true };
            txtScanDir = new TextBox
            {
                Location = new Point(155, 12), Width = 580,
                Text = OutlookMonitorService.SaveDir
            };
            btnBrowseScan = new Button { Text = "...", Location = new Point(743, 11), Width = 35 };
            btnBrowseScan.Click += (s, e) => BrowseFolder(txtScanDir);

            var lblUnclassified = new Label { Text = "Thư mục 'Chưa phân loại':", Location = new Point(10, 50), AutoSize = true };
            txtUnclassifiedBase = new TextBox
            {
                Location = new Point(185, 47), Width = 550,
                Text = OutlookMonitorService.SaveDir
            };
            btnBrowseUnclassified = new Button { Text = "...", Location = new Point(743, 46), Width = 35 };
            btnBrowseUnclassified.Click += (s, e) => BrowseFolder(txtUnclassifiedBase);

            btnScanFolder = new Button
            {
                Text = "🔍 Quét & Phân loại toàn bộ thư mục",
                Location = new Point(10, 78),
                Width = 250, Height = 28,
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnScanFolder.Click += BtnScanFolder_Click;

            btnClassifySelected = new Button
            {
                Text = "📋 Phân loại file đã chọn",
                Location = new Point(270, 78),
                Width = 200, Height = 28,
                BackColor = Color.FromArgb(39, 174, 96),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Enabled = false
            };
            btnClassifySelected.Click += BtnClassifySelected_Click;

            btnOpenDest = new Button
            {
                Text = "Mở thư mục đích",
                Location = new Point(480, 78),
                Width = 150, Height = 28,
                Enabled = false
            };
            btnOpenDest.Click += BtnOpenDest_Click;

            btnCutToFolder = new Button
            {
                Text = "✂ Di chuyển vào...",
                Location = new Point(640, 78),
                Width = 155, Height = 28,
                BackColor = Color.FromArgb(211, 84, 0),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Enabled = false
            };
            btnCutToFolder.Click += BtnCutToFolder_Click;

            pnlTop.Controls.AddRange(new Control[] {
                lblScan, txtScanDir, btnBrowseScan,
                lblUnclassified, txtUnclassifiedBase, btnBrowseUnclassified,
                btnScanFolder, btnClassifySelected, btnOpenDest, btnCutToFolder
            });

            // ── ListView kết quả ───────────────────────────────────────────────
            lvResults = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true
            };
            lvResults.Columns.Add("Trạng thái", 110);
            lvResults.Columns.Add("Tên file đích", 200);
            lvResults.Columns.Add("PO No", 155);
            lvResults.Columns.Add("Dự án", 110);
            lvResults.Columns.Add("Trùng xóa", 70);
            lvResults.Columns.Add("Thư mục đích", 300);
            lvResults.SelectedIndexChanged += (s, e) =>
            {
                bool any = lvResults.SelectedItems.Count > 0;
                btnOpenDest.Enabled = any;
                btnCutToFolder.Enabled = any;
            };
            lvResults.DoubleClick += LvResults_DoubleClick;

            // ── Panel dưới ────────────────────────────────────────────────────
            var pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 40 };
            progressBar = new ProgressBar
            {
                Location = new Point(10, 10),
                Width = 200, Height = 18,
                Style = ProgressBarStyle.Marquee,
                Visible = false
            };
            lblStatus = new Label { Location = new Point(220, 13), AutoSize = true, Text = "Sẵn sàng." };
            lblSummary = new Label
            {
                Dock = DockStyle.Right,
                AutoSize = true,
                Padding = new Padding(0, 13, 15, 0),
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            pnlBottom.Controls.AddRange(new Control[] { progressBar, lblStatus, lblSummary });

            Controls.Add(lvResults);
            Controls.Add(pnlTop);
            Controls.Add(pnlBottom);
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            // Nếu được truyền file từ form Outlook
            if (PreSelectedFiles.Count > 0)
            {
                btnClassifySelected.Enabled = true;
                lblStatus.Text = $"{PreSelectedFiles.Count} file được chọn từ Outlook. Nhấn 'Phân loại file đã chọn'.";
            }
        }

        private async void BtnScanFolder_Click(object? sender, EventArgs e)
        {
            if (!Directory.Exists(txtScanDir.Text))
            {
                MessageBox.Show("Thư mục nguồn không tồn tại.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            await RunClassify(scanFolder: true);
        }

        private async void BtnClassifySelected_Click(object? sender, EventArgs e)
        {
            if (PreSelectedFiles.Count == 0) return;
            await RunClassify(scanFolder: false);
        }

        private async System.Threading.Tasks.Task RunClassify(bool scanFolder)
        {
            SetBusy(true);
            lvResults.Items.Clear();
            lblSummary.Text = "";

            var progress = new Progress<string>(msg => lblStatus.Text = msg);
            ClassifyResponse result;

            if (scanFolder)
            {
                result = await InvoiceClassifierService.ClassifyFolderAsync(
                    txtScanDir.Text, txtUnclassifiedBase.Text, progress);
            }
            else
            {
                result = await InvoiceClassifierService.ClassifyFilesAsync(
                    PreSelectedFiles, txtUnclassifiedBase.Text, progress);
            }

            SetBusy(false);

            if (!result.Success)
            {
                MessageBox.Show(result.Error, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "Lỗi!";
                return;
            }

            foreach (var r in result.Results)
                AddResultRow(r);

            lblSummary.Text =
                $"Tổng: {result.Summary.Total}  |  ✅ Phân loại được: {result.Summary.Classified}  |  ⚠ Chưa phân loại: {result.Summary.Unclassified}";
            lblStatus.Text = "Hoàn thành.";
        }

        private void AddResultRow(ClassifyResult r)
        {
            var item = new ListViewItem(GetStatusText(r.Status));
            item.SubItems.Add(Path.GetFileName(r.Dest ?? r.File));
            item.SubItems.Add(r.PoNo ?? "-");
            item.SubItems.Add(r.ProjectCode ?? "-");
            item.SubItems.Add(r.DuplicatesRemoved > 0 ? $"🗑 {r.DuplicatesRemoved}" : "-");
            item.SubItems.Add(r.Dest != null ? Path.GetDirectoryName(r.Dest) ?? "-" : (r.Reason ?? "-"));
            item.Tag = r.Dest;

            item.BackColor = r.Status switch
            {
                "classified" => Color.FromArgb(230, 255, 230),
                "no_inv_link" => Color.FromArgb(255, 245, 200),
                _ => Color.FromArgb(255, 230, 230)
            };

            lvResults.Items.Add(item);
        }

        private static string GetStatusText(string status) => status switch
        {
            "classified" => "✅ Đã phân loại",
            "no_inv_link" => "⚠ Thiếu INV_Link",
            _ => "❌ Chưa tìm thấy"
        };

        private void BtnCutToFolder_Click(object? sender, EventArgs e)
        {
            var selected = lvResults.SelectedItems.Cast<ListViewItem>()
                .Where(i => i.Tag?.ToString() is string f && File.Exists(f))
                .ToList();

            if (selected.Count == 0)
            {
                MessageBox.Show("Không có file hợp lệ nào được chọn (file có thể đã bị xóa hoặc chưa được phân loại).",
                    "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var dlg = new FolderBrowserDialog { Description = $"Chọn thư mục đích để di chuyển {selected.Count} file" };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            string targetDir = dlg.SelectedPath;
            int moved = 0, failed = 0;

            foreach (var item in selected)
            {
                string src = item.Tag!.ToString()!;
                string destPath = Path.Combine(targetDir, Path.GetFileName(src));

                // Tránh ghi đè: thêm số thứ tự nếu tồn tại
                if (File.Exists(destPath))
                {
                    string base_ = Path.GetFileNameWithoutExtension(src);
                    string ext = Path.GetExtension(src);
                    int idx = 1;
                    while (File.Exists(destPath))
                        destPath = Path.Combine(targetDir, $"{base_}_{idx++}{ext}");
                }

                try
                {
                    File.Move(src, destPath);
                    item.Tag = destPath;
                    // Cập nhật cột thư mục đích trên ListView
                    item.SubItems[5].Text = targetDir;
                    item.BackColor = Color.FromArgb(220, 235, 255); // xanh nhạt = đã move thủ công
                    moved++;
                }
                catch (Exception ex)
                {
                    failed++;
                    lblStatus.Text = $"Lỗi di chuyển {Path.GetFileName(src)}: {ex.Message}";
                }
            }

            lblStatus.Text = $"Đã di chuyển {moved} file vào: {targetDir}" +
                             (failed > 0 ? $" ({failed} lỗi)" : "");
        }

        private void BtnOpenDest_Click(object? sender, EventArgs e)
        {
            if (lvResults.SelectedItems.Count == 0) return;
            string? dest = lvResults.SelectedItems[0].Tag?.ToString();
            if (dest == null) return;
            string? dir = File.Exists(dest) ? Path.GetDirectoryName(dest) : dest;
            if (dir != null && Directory.Exists(dir))
                System.Diagnostics.Process.Start("explorer.exe", dir);
        }

        private void LvResults_DoubleClick(object? sender, EventArgs e)
        {
            if (lvResults.SelectedItems.Count == 0) return;
            string? dest = lvResults.SelectedItems[0].Tag?.ToString();
            if (dest != null && File.Exists(dest))
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(dest) { UseShellExecute = true });
        }

        private void BrowseFolder(TextBox target)
        {
            using var dlg = new FolderBrowserDialog { SelectedPath = target.Text };
            if (dlg.ShowDialog() == DialogResult.OK)
                target.Text = dlg.SelectedPath;
        }

        private void SetBusy(bool busy)
        {
            progressBar.Visible = busy;
            btnScanFolder.Enabled = !busy;
            btnClassifySelected.Enabled = !busy && PreSelectedFiles.Count > 0;
        }
    }
}
