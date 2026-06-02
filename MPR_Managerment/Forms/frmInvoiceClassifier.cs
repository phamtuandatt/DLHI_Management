using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using MPR_Managerment.Models;
using MPR_Managerment.Services;

namespace MPR_Managerment.Forms
{
    public class frmInvoiceClassifier : Form
    {
        private TextBox txtScanDir = null!;
        private TextBox txtUnclassifiedBase = null!;
        private Button btnBrowseScan = null!;
        private Button btnBrowseUnclassified = null!;
        private Button btnAddFromUnclassified = null!;
        private Button btnScanFolder = null!;
        private Button btnClassifySelected = null!;
        private Button btnOpenDest = null!;
        private Button btnCutToFolder = null!;
        private ListView lvResults = null!;
        private Label lblStatus = null!;
        private ProgressBar progressBar = null!;
        private Label lblSummary = null!;
        private SplitContainer splitMain = null!;
        private WebView2 webPreview = null!;
        private Label lblPreviewHint = null!;

        // Assign panel
        private Panel pnlAssign = null!;
        private ComboBox cboProject = null!;
        private ComboBox cboPO = null!;
        private Button btnAssign = null!;
        private Label lblAssignInfo = null!;

        private readonly ProjectService _projectSvc = new();
        private readonly POService _poSvc = new();
        private List<ProjectInfo> _projects = new();

        public List<string> PreSelectedFiles { get; set; } = new();

        public frmInvoiceClassifier()
        {
            InitializeComponent();
            _ = LoadProjectsAsync();
            _ = InitWebViewAsync();
        }

        private void InitializeComponent()
        {
            Text = "Phân loại Hóa Đơn → Thư mục Dự án";
            Size = new Size(1400, 720);
            StartPosition = FormStartPosition.CenterParent;

            // ── Panel trên ────────────────────────────────────────────────────
            var pnlTop = new Panel { Dock = DockStyle.Top, Height = 110, Padding = new Padding(10) };

            var lblScan = new Label { Text = "Thư mục PDF nguồn:", Location = new Point(10, 15), AutoSize = true };
            txtScanDir = new TextBox { Location = new Point(155, 12), Width = 580, Text = OutlookMonitorService.SaveDir };
            btnBrowseScan = new Button { Text = "...", Location = new Point(743, 11), Width = 35 };
            btnBrowseScan.Click += (s, e) => BrowseFolder(txtScanDir);

            var lblUnclassified = new Label { Text = "Thư mục 'Chưa phân loại':", Location = new Point(10, 50), AutoSize = true };
            txtUnclassifiedBase = new TextBox
            {
                Location = new Point(185, 47), Width = 550,
                Text = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "MPR_Invoices", "Chưa phân loại")
            };
            btnBrowseUnclassified = new Button { Text = "...", Location = new Point(743, 46), Width = 35 };
            btnBrowseUnclassified.Click += (s, e) => BrowseFolder(txtUnclassifiedBase);

            btnAddFromUnclassified = new Button
            {
                Text = "+ Thêm từ thư mục",
                Location = new Point(786, 44), Width = 145, Height = 28,
                BackColor = Color.FromArgb(52, 152, 219), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
            };
            btnAddFromUnclassified.Click += BtnAddFromUnclassified_Click;

            btnScanFolder = new Button
            {
                Text = "🔍 Quét & Phân loại toàn bộ thư mục",
                Location = new Point(10, 78), Width = 250, Height = 28,
                BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
            };
            btnScanFolder.Click += BtnScanFolder_Click;

            btnClassifySelected = new Button
            {
                Text = "📋 Phân loại file đã chọn",
                Location = new Point(270, 78), Width = 200, Height = 28,
                BackColor = Color.FromArgb(39, 174, 96), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Enabled = false
            };
            btnClassifySelected.Click += BtnClassifySelected_Click;

            btnOpenDest = new Button { Text = "Mở thư mục đích", Location = new Point(480, 78), Width = 150, Height = 28, Enabled = false };
            btnOpenDest.Click += BtnOpenDest_Click;

            btnCutToFolder = new Button
            {
                Text = "✂ Di chuyển vào...",
                Location = new Point(640, 78), Width = 155, Height = 28,
                BackColor = Color.FromArgb(211, 84, 0), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Enabled = false
            };
            btnCutToFolder.Click += BtnCutToFolder_Click;

            pnlTop.Controls.AddRange(new Control[] {
                lblScan, txtScanDir, btnBrowseScan,
                lblUnclassified, txtUnclassifiedBase, btnBrowseUnclassified, btnAddFromUnclassified,
                btnScanFolder, btnClassifySelected, btnOpenDest, btnCutToFolder
            });

            // ── SplitContainer (ListView | Preview) ───────────────────────────
            splitMain = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical
            };

            // Trái: ListView
            lvResults = new ListView
            {
                Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, GridLines = true, MultiSelect = true
            };
            lvResults.Columns.Add("Trạng thái", 110);
            lvResults.Columns.Add("Tên file đích", 200);
            lvResults.Columns.Add("PO No", 130);
            lvResults.Columns.Add("Dự án", 100);
            lvResults.Columns.Add("Trùng xóa", 65);
            lvResults.Columns.Add("Thư mục đích", 280);
            lvResults.SelectedIndexChanged += LvResults_SelectedIndexChanged;
            lvResults.DoubleClick += LvResults_DoubleClick;
            splitMain.Panel1.Controls.Add(lvResults);

            // Phải: Preview panel
            var pnlPreview = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(245, 245, 245) };

            lblPreviewHint = new Label
            {
                Text = "🔍 Chọn một file để xem trước",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 11, FontStyle.Regular),
                ForeColor = Color.Gray
            };

            webPreview = new WebView2 { Dock = DockStyle.Fill, Visible = false };

            pnlPreview.Controls.Add(webPreview);
            pnlPreview.Controls.Add(lblPreviewHint);
            splitMain.Panel2.Controls.Add(pnlPreview);

            // ── Panel gán Dự án / PO ──────────────────────────────────────────
            pnlAssign = new Panel
            {
                Dock = DockStyle.Bottom, Height = 48,
                BackColor = Color.FromArgb(240, 245, 255),
                Padding = new Padding(8, 6, 8, 6),
                Visible = false
            };
            pnlAssign.Paint += (s, e) => e.Graphics.DrawLine(Pens.LightSteelBlue, 0, 0, pnlAssign.Width, 0);

            lblAssignInfo = new Label { Location = new Point(8, 14), AutoSize = true, ForeColor = Color.DimGray, Font = new Font("Segoe UI", 8.5f) };
            var lblProj = new Label { Text = "Dự án:", Location = new Point(160, 14), AutoSize = true };
            cboProject = new ComboBox { Location = new Point(210, 10), Width = 160, DropDownStyle = ComboBoxStyle.DropDownList };
            cboProject.SelectedIndexChanged += CboProject_SelectedIndexChanged;
            var lblPO = new Label { Text = "PO No:", Location = new Point(382, 14), AutoSize = true };
            cboPO = new ComboBox { Location = new Point(425, 10), Width = 180, DropDownStyle = ComboBoxStyle.DropDownList };
            btnAssign = new Button
            {
                Text = "📌 Cập nhật",
                Location = new Point(620, 8), Width = 120, Height = 30,
                BackColor = Color.FromArgb(39, 174, 96), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
            };
            btnAssign.Click += BtnAssign_Click;
            pnlAssign.Controls.AddRange(new Control[] { lblAssignInfo, lblProj, cboProject, lblPO, cboPO, btnAssign });

            // ── Panel dưới ────────────────────────────────────────────────────
            var pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 40 };
            progressBar = new ProgressBar { Location = new Point(10, 10), Width = 200, Height = 18, Style = ProgressBarStyle.Marquee, Visible = false };
            lblStatus = new Label { Location = new Point(220, 13), AutoSize = true, Text = "Sẵn sàng." };
            lblSummary = new Label { Dock = DockStyle.Right, AutoSize = true, Padding = new Padding(0, 13, 15, 0), Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            pnlBottom.Controls.AddRange(new Control[] { progressBar, lblStatus, lblSummary });

            Controls.Add(splitMain);
            Controls.Add(pnlAssign);
            Controls.Add(pnlTop);
            Controls.Add(pnlBottom);
        }

        // ── WebView2 init ─────────────────────────────────────────────────────

        private async Task InitWebViewAsync()
        {
            try
            {
                var env = await CoreWebView2Environment.CreateAsync(null,
                    Path.Combine(Path.GetTempPath(), "MPR_WebView2Cache"));
                await webPreview.EnsureCoreWebView2Async(env);
                webPreview.CoreWebView2.Settings.IsStatusBarEnabled = false;
                webPreview.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
            }
            catch { /* WebView2 không khởi động được — preview không hiển thị */ }
        }

        private void ShowPreview(string? filePath)
        {
            if (webPreview.CoreWebView2 == null || string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                webPreview.Visible = false;
                lblPreviewHint.Visible = true;
                return;
            }

            try
            {
                webPreview.CoreWebView2.Navigate(new Uri(filePath).AbsoluteUri);
                webPreview.Visible = true;
                lblPreviewHint.Visible = false;
            }
            catch
            {
                webPreview.Visible = false;
                lblPreviewHint.Text = "⚠ Không thể xem trước file này";
                lblPreviewHint.Visible = true;
            }
        }

        // ── Load dự án / PO ──────────────────────────────────────────────────

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

        // ── ListView selection ────────────────────────────────────────────────

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            // Set splitter và min size sau khi layout hoàn tất
            BeginInvoke(() =>
            {
                splitMain.Panel1MinSize = 400;
                splitMain.Panel2MinSize = 280;
                splitMain.SplitterDistance = (int)(splitMain.Width * 0.60);
            });

            if (PreSelectedFiles.Count > 0)
            {
                btnClassifySelected.Enabled = true;
                lblStatus.Text = $"{PreSelectedFiles.Count} file được chọn từ Outlook. Nhấn 'Phân loại file đã chọn'.";
            }
        }

        private void LvResults_SelectedIndexChanged(object? sender, EventArgs e)
        {
            int count = lvResults.SelectedItems.Count;
            bool any = count > 0;
            btnOpenDest.Enabled = any;
            btnCutToFolder.Enabled = any;
            pnlAssign.Visible = any;

            if (any)
            {
                lblAssignInfo.Text = count == 1 ? "1 file được chọn" : $"{count} file được chọn";
                // Preview file đầu tiên được chọn
                string? path = lvResults.SelectedItems[0].Tag?.ToString();
                ShowPreview(path);
            }
            else
            {
                webPreview.Visible = false;
                lblPreviewHint.Text = "🔍 Chọn một file để xem trước";
                lblPreviewHint.Visible = true;
            }
        }

        // ── Gán dự án / PO ───────────────────────────────────────────────────

        private async void BtnAssign_Click(object? sender, EventArgs e)
        {
            if (cboProject.SelectedItem is not ProjectComboItem { Project: not null } projSel)
            {
                MessageBox.Show("Vui lòng chọn dự án.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var project = projSel.Project;
            string? poNo = cboPO.SelectedItem is POComboItem { PONo: not null and not "" } poSel ? poSel.PONo : null;

            var selectedItems = lvResults.SelectedItems.Cast<ListViewItem>().ToList();
            if (selectedItems.Count == 0) return;

            string destDir = string.IsNullOrWhiteSpace(project.INV_Link) ? txtUnclassifiedBase.Text : project.INV_Link;
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
                    if (srcPath == null || !File.Exists(srcPath)) { fail++; errors.Add($"{item.SubItems[1].Text}: file không tồn tại"); continue; }
                    try
                    {
                        string ext = Path.GetExtension(srcPath);
                        string origName = Path.GetFileNameWithoutExtension(srcPath);
                        string prefix = poNo != null ? $"{poNo}_{project.ProjectCode}_" : $"{project.ProjectCode}_";
                        string newName = origName.StartsWith(prefix) ? origName + ext : prefix + origName + ext;
                        string destPath = Path.Combine(destDir, newName);
                        if (File.Exists(destPath))
                            destPath = Path.Combine(destDir, prefix + origName + $"_{DateTimeOffset.Now.ToUnixTimeSeconds()}" + ext);

                        File.Move(srcPath, destPath);
                        Invoke(() =>
                        {
                            item.SubItems[0].Text = "✅ Đã phân loại";
                            item.SubItems[1].Text = Path.GetFileName(destPath);
                            item.SubItems[2].Text = poNo ?? "-";
                            item.SubItems[3].Text = project.ProjectCode;
                            item.SubItems[5].Text = destDir;
                            item.Tag = destPath;
                            item.BackColor = Color.FromArgb(224, 240, 255);
                        });
                        success++;
                    }
                    catch (Exception ex) { fail++; errors.Add($"{item.SubItems[1].Text}: {ex.Message}"); }
                }
            });

            lblStatus.Text = $"Cập nhật xong: {success} file → {destDir}{(fail > 0 ? $" ({fail} lỗi)" : "")}";
            if (errors.Count > 0)
                MessageBox.Show(string.Join("\n", errors), "Một số file lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
                MessageBox.Show($"Đã đổi tên và di chuyển {success} file vào:\n{destDir}", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ── Classify ──────────────────────────────────────────────────────────

        private async void BtnScanFolder_Click(object? sender, EventArgs e)
        {
            if (!Directory.Exists(txtScanDir.Text)) { MessageBox.Show("Thư mục nguồn không tồn tại.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            await RunClassify(scanFolder: true);
        }

        private async void BtnClassifySelected_Click(object? sender, EventArgs e)
        {
            if (PreSelectedFiles.Count == 0) return;
            await RunClassify(scanFolder: false);
        }

        private async Task RunClassify(bool scanFolder)
        {
            SetBusy(true);
            lvResults.Items.Clear();
            lblSummary.Text = "";
            webPreview.Visible = false;
            lblPreviewHint.Text = "🔍 Chọn một file để xem trước";
            lblPreviewHint.Visible = true;

            var progress = new Progress<string>(msg => lblStatus.Text = msg);
            ClassifyResponse result;
            if (scanFolder)
                result = await InvoiceClassifierService.ClassifyFolderAsync(txtScanDir.Text, txtUnclassifiedBase.Text, progress);
            else
                result = await InvoiceClassifierService.ClassifyFilesAsync(PreSelectedFiles, txtUnclassifiedBase.Text, progress);

            SetBusy(false);
            if (!result.Success) { MessageBox.Show(result.Error, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); lblStatus.Text = "Lỗi!"; return; }

            foreach (var r in result.Results) AddResultRow(r);

            lblSummary.Text = $"Tổng: {result.Summary.Total}  |  ✅ Phân loại được: {result.Summary.Classified}  |  ⚠ Chưa phân loại: {result.Summary.Unclassified}";
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
            item.Tag = r.Dest ?? r.File;
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

        // ── Thêm từ thư mục ───────────────────────────────────────────────────

        private void BtnAddFromUnclassified_Click(object? sender, EventArgs e)
        {
            string dir = txtUnclassifiedBase.Text;
            if (!Directory.Exists(dir)) { MessageBox.Show("Thư mục 'Chưa phân loại' không tồn tại.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            var pdfFiles = Directory.GetFiles(dir, "*.pdf", SearchOption.TopDirectoryOnly)
                .Where(f => !PreSelectedFiles.Contains(f, StringComparer.OrdinalIgnoreCase))
                .OrderBy(f => f).ToList();

            if (pdfFiles.Count == 0) { MessageBox.Show("Không có file PDF nào mới trong thư mục.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }

            using var dlg = new frmPickFiles(pdfFiles);
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            var chosen = dlg.SelectedFiles;
            if (chosen.Count == 0) return;

            PreSelectedFiles.AddRange(chosen);
            btnClassifySelected.Enabled = true;

            foreach (var f in chosen)
            {
                var item = new ListViewItem("⏳ Chờ phân loại");
                item.SubItems.Add(Path.GetFileName(f));
                item.SubItems.Add("-"); item.SubItems.Add("-"); item.SubItems.Add("-");
                item.SubItems.Add(Path.GetDirectoryName(f) ?? "");
                item.Tag = f;
                item.BackColor = Color.FromArgb(255, 250, 220);
                lvResults.Items.Insert(0, item);
            }
            lblStatus.Text = $"Đã thêm {chosen.Count} file. Chọn file để gán Dự án/PO, hoặc nhấn 'Phân loại file đã chọn'.";
        }

        // ── Di chuyển thủ công ───────────────────────────────────────────────

        private void BtnCutToFolder_Click(object? sender, EventArgs e)
        {
            var selected = lvResults.SelectedItems.Cast<ListViewItem>()
                .Where(i => i.Tag?.ToString() is string f && File.Exists(f)).ToList();
            if (selected.Count == 0) { MessageBox.Show("Không có file hợp lệ nào được chọn.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }

            using var dlg = new FolderBrowserDialog { Description = $"Chọn thư mục đích để di chuyển {selected.Count} file" };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            string targetDir = dlg.SelectedPath;
            int moved = 0, failed = 0;
            foreach (var item in selected)
            {
                string src = item.Tag!.ToString()!;
                string destPath = Path.Combine(targetDir, Path.GetFileName(src));
                if (File.Exists(destPath))
                {
                    string b = Path.GetFileNameWithoutExtension(src), x = Path.GetExtension(src);
                    int idx = 1;
                    while (File.Exists(destPath)) destPath = Path.Combine(targetDir, $"{b}_{idx++}{x}");
                }
                try { File.Move(src, destPath); item.Tag = destPath; item.SubItems[5].Text = targetDir; item.BackColor = Color.FromArgb(220, 235, 255); moved++; }
                catch (Exception ex) { failed++; lblStatus.Text = $"Lỗi: {ex.Message}"; }
            }
            lblStatus.Text = $"Đã di chuyển {moved} file vào: {targetDir}" + (failed > 0 ? $" ({failed} lỗi)" : "");
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
            if (dlg.ShowDialog() == DialogResult.OK) target.Text = dlg.SelectedPath;
        }

        private void SetBusy(bool busy)
        {
            progressBar.Visible = busy;
            btnScanFolder.Enabled = !busy;
            btnClassifySelected.Enabled = !busy && PreSelectedFiles.Count > 0;
        }

        // ── ComboBox wrappers ────────────────────────────────────────────────

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
