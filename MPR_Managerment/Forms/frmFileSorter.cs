using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MPR_Managerment.Forms
{
    /// <summary>
    /// Phân loại file scan PO/MPR từ thư mục nguồn vào đúng thư mục dự án.
    ///
    /// Pattern trích xuất WorkorderNo từ tên file:
    ///   PO  → DV-{WorkorderNo}-PC-{số}   → lấy phần giữa "DV-" và "-PC-{số}"
    ///         Ví dụ: DV-GHA-PC-067.pdf   → WorkorderNo = "GHA"
    ///
    ///   MPR → DV-{WorkorderNo}-MPR-{số}  → lấy phần giữa "DV-" và "-MPR-{số}"
    ///         Ví dụ: DV-FT-2502-MPR-001.pdf → WorkorderNo = "FT-2502"
    ///
    /// Đối chiếu với DB.ProjectInfo → cut vào PO_Link / MPR_Link.
    /// File trùng tên → tự động đổi tên thêm _1, _2...
    /// </summary>
    public class frmFileSorter : Form
    {
        // ── Controls ─────────────────────────────────────────────────────
        private TextBox _txtSource;
        private DataGridView _dgv;
        private Button _btnBrowse, _btnScan, _btnRun, _btnClose;
        private Label _lblStatus;
        private ProgressBar _progress;
        private CheckBox _chkPreview;

        // ── Data ─────────────────────────────────────────────────────────
        private List<SortItem> _items = new List<SortItem>();
        private List<ProjectRec> _projects = new List<ProjectRec>();

        // ── Regex ─────────────────────────────────────────────────────────
        // PO:  DV-GHA-PC-067.pdf      → group(1) = "GHA"
        private static readonly Regex RexPO =
            new Regex(@"DV-(.+?)-PC-\d+",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);
        // MPR: trích tất cả ký tự trước dấu "-" cuối cùng (bao gồm dấu "-" đó)
        // DV-SL-2511-ABT-MPR-001.pdf → "DV-SL-2511-ABT-MPR-"
        private static readonly Regex RexMPR =
            new Regex(@"^(.*-)\d+\.",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // ═════════════════════════════════════════════════════════════════
        // MODELS
        // ═════════════════════════════════════════════════════════════════
        private class SortItem
        {
            public string FileName { get; set; }
            public string SourcePath { get; set; }
            public string ExtractedCode { get; set; }
            public string FileType { get; set; } // "PO" | "MPR" | "-"
            public string MatchedProject { get; set; }
            public string TargetFolder { get; set; }
            public string FinalName { get; set; }
            public string Status { get; set; }
        }

        private class ProjectRec
        {
            public string WorkorderNo { get; set; }
            public string ProjectName { get; set; }
            public string POCode { get; set; } // khớp với file PO:  DV-{POCode}-PC-{số}
            public string MPRCode { get; set; } // khớp với file MPR: DV-{MPRCode}-MPR-{số}
            public string POLink { get; set; }
            public string MPRLink { get; set; }
        }

        // ═════════════════════════════════════════════════════════════════
        // CONSTRUCTOR
        // ═════════════════════════════════════════════════════════════════
        public frmFileSorter()
        {
            Text = "📂 Phân loại file Scan PO / MPR";
            Size = new Size(1150, 680);
            MinimumSize = new Size(920, 520);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.FromArgb(245, 247, 250);
            Font = new Font("Segoe UI", 9);
            BuildUI();
        }

        // ═════════════════════════════════════════════════════════════════
        // BUILD UI
        // ═════════════════════════════════════════════════════════════════
        private void BuildUI()
        {
            // ── Top panel ────────────────────────────────────────────────
            var pTop = new Panel
            {
                Dock = DockStyle.Top,
                Height = 108,
                BackColor = Color.White,
                Padding = new Padding(16, 10, 16, 8)
            };

            // Dòng 1: label
            pTop.Controls.Add(MkLabel(
                "📁 Thư mục nguồn (PO_TEST):",
                new Point(16, 12), bold: true));

            // Dòng 2: textbox + buttons
            _txtSource = new TextBox
            {
                Location = new Point(16, 34),
                Size = new Size(640, 26),
                Text = @"D:\PO_TEST",
                Font = new Font("Segoe UI", 9)
            };
            pTop.Controls.Add(_txtSource);

            _btnBrowse = MkBtn("…", new Point(664, 34), new Size(36, 26),
                Color.FromArgb(108, 117, 125));
            _btnBrowse.Click += (s, e) =>
            {
                using var d = new FolderBrowserDialog
                { SelectedPath = _txtSource.Text };
                if (d.ShowDialog() == DialogResult.OK)
                    _txtSource.Text = d.SelectedPath;
            };
            pTop.Controls.Add(_btnBrowse);

            _btnScan = MkBtn("🔍  Quét file", new Point(710, 34), new Size(120, 26),
                Color.FromArgb(0, 120, 212));
            _btnScan.Click += BtnScan_Click;
            pTop.Controls.Add(_btnScan);

            _btnRun = MkBtn("▶  Thực hiện Cut", new Point(840, 34), new Size(150, 26),
                Color.FromArgb(40, 167, 69));
            _btnRun.Enabled = false;
            _btnRun.Click += BtnRun_Click;
            pTop.Controls.Add(_btnRun);

            // Dòng 3: options + ghi chú
            _chkPreview = new CheckBox
            {
                Text = "Preview — chỉ xem, không di chuyển file thực sự",
                Location = new Point(16, 70),
                AutoSize = true,
                Checked = false
            };
            pTop.Controls.Add(_chkPreview);

            pTop.Controls.Add(MkLabel(
                "PO: DV-{POCode}-PC-{số} → lấy {POCode}   |   MPR: lấy tất cả trước dấu \"-\" cuối (gồm cả \"-\")   |   File trùng → _1, _2...",
                new Point(380, 74),
                color: Color.FromArgb(108, 117, 125)));

            // ── Grid ─────────────────────────────────────────────────────
            _dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 8.5f),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ColumnHeadersHeight = 28,
            };
            _dgv.ColumnHeadersDefaultCellStyle.Font =
                new Font("Segoe UI", 8.5f, FontStyle.Bold);
            _dgv.RowTemplate.Height = 22;

            var cols = new[]
            {
                ("FileName",       "Tên file gốc",    22),
                ("FileType",       "Loại",             5),
                ("ExtractedCode",  "Mã trích xuất",   11),
                ("MatchedProject", "Dự án khớp",      18),
                ("TargetFolder",   "Thư mục đích",    27),
                ("FinalName",      "Tên lưu",         13),
                ("Status",         "Trạng thái",      10),
            };
            foreach (var (n, h, w) in cols)
                _dgv.Columns.Add(new DataGridViewTextBoxColumn
                { Name = n, HeaderText = h, FillWeight = w });

            _dgv.CellFormatting += DgvFormat;

            // ── Bottom bar ───────────────────────────────────────────────
            var pBot = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 36,
                BackColor = Color.FromArgb(33, 37, 41),
                Padding = new Padding(12, 6, 12, 4)
            };
            _lblStatus = new Label
            {
                Text = "Sẵn sàng.  Nhấn 🔍 Quét file để bắt đầu.",
                ForeColor = Color.White,
                AutoSize = true,
                Dock = DockStyle.Left,
                Font = new Font("Segoe UI", 9)
            };
            _progress = new ProgressBar
            {
                Dock = DockStyle.Right,
                Width = 200,
                Visible = false
            };
            _btnClose = MkBtn("✕  Đóng", Point.Empty, new Size(90, 26),
                Color.FromArgb(200, 53, 69));
            _btnClose.Dock = DockStyle.Right;
            _btnClose.Click += (s, e) => Close();

            pBot.Controls.AddRange(new Control[]
                { _lblStatus, _progress, _btnClose });

            Controls.Add(_dgv);
            Controls.Add(pTop);
            Controls.Add(pBot);
        }

        // ═════════════════════════════════════════════════════════════════
        // SCAN
        // ═════════════════════════════════════════════════════════════════
        private void BtnScan_Click(object sender, EventArgs e)
        {
            string src = _txtSource.Text.Trim();
            if (!Directory.Exists(src))
            {
                MessageBox.Show($"Thư mục không tồn tại:\n{src}",
                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SetStatus("⏳ Đang tải dữ liệu dự án từ DB...");
            _items.Clear();
            _dgv.Rows.Clear();
            _btnRun.Enabled = false;

            // Load projects
            _projects = LoadProjects();
            if (_projects.Count == 0)
            {
                SetStatus("⚠️ Không tìm thấy dự án nào trong DB (bảng ProjectInfo).");
                return;
            }

            // Đọc file
            SetStatus("⏳ Đang quét file...");
            string[] files;
            try
            {
                files = Directory.GetFiles(src, "*.*", SearchOption.TopDirectoryOnly)
                    .Where(f => !Path.GetFileName(f).StartsWith("~$")
                             && !Path.GetFileName(f).EndsWith(".log"))
                    .ToArray();
            }
            catch (Exception ex)
            {
                SetStatus($"❌ Lỗi đọc thư mục: {ex.Message}");
                return;
            }

            if (files.Length == 0)
            {
                SetStatus("📂 Thư mục không có file nào.");
                return;
            }

            // Phân tích
            foreach (string fp in files)
                _items.Add(AnalyzeFile(Path.GetFileName(fp), fp));

            RefreshGrid();

            int ready = _items.Count(i => i.Status == "✅ Sẵn sàng");
            int renamed = _items.Count(i => i.Status == "⚠️ Đổi tên");
            int noMatch = _items.Count(i => i.Status == "❌ Không khớp");
            int err = _items.Count(i =>
                i.Status != "✅ Sẵn sàng" &&
                i.Status != "⚠️ Đổi tên" &&
                i.Status != "❌ Không khớp");

            SetStatus($"🔍 {files.Length} file  |  " +
                $"✅ {ready} sẵn sàng  |  " +
                $"⚠️ {renamed} cần đổi tên  |  " +
                $"❌ {noMatch} không khớp  |  " +
                $"⚠️ {err} lỗi link");

            _btnRun.Enabled = _items.Any(i =>
                i.Status == "✅ Sẵn sàng" || i.Status == "⚠️ Đổi tên");
        }

        // ─────────────────────────────────────────────────────────────────
        private SortItem AnalyzeFile(string fileName, string sourcePath)
        {
            var item = new SortItem
            {
                FileName = fileName,
                SourcePath = sourcePath,
                FileType = "-",
                Status = "❌ Không khớp"
            };

            string fileNoExt = Path.GetFileNameWithoutExtension(fileName);
            string extracted = null;
            bool isMPR = false;

            // ── Thử PO trước: DV-{POCode}-PC-{số} ────────────────────────
            var mPO = RexPO.Match(fileName);
            if (mPO.Success)
            {
                extracted = mPO.Groups[1].Value; // VD: "GHA"
                isMPR = false;
            }
            else
            {
                // ── Thử MPR: lấy tất cả trước dấu "-" cuối cùng (kể cả "-")
                // VD: "DV-SL-2511-ABT-MPR-001" → "DV-SL-2511-ABT-MPR-"
                var mMPR = RexMPR.Match(fileName);
                if (mMPR.Success)
                {
                    extracted = mMPR.Groups[1].Value; // VD: "DV-SL-2511-ABT-MPR-"
                    isMPR = true;
                }
            }

            if (extracted == null) return item;

            item.ExtractedCode = extracted;
            item.FileType = isMPR ? "MPR" : "PO";

            // ── Đối chiếu với DB ─────────────────────────────────────────
            ProjectRec proj = null;
            if (isMPR)
            {
                // MPRCode trong DB lưu dạng "DV-SL-2511-ABT-MPR-"
                proj = _projects.FirstOrDefault(p =>
                    string.Equals(p.MPRCode, extracted, StringComparison.OrdinalIgnoreCase))
                    ?? _projects.FirstOrDefault(p =>
                        !string.IsNullOrEmpty(p.MPRCode) &&
                        (p.MPRCode.IndexOf(extracted, StringComparison.OrdinalIgnoreCase) >= 0 ||
                         extracted.IndexOf(p.MPRCode, StringComparison.OrdinalIgnoreCase) >= 0));
            }
            else
            {
                // POCode trong DB lưu dạng "GHA"
                proj = _projects.FirstOrDefault(p =>
                    string.Equals(p.POCode, extracted, StringComparison.OrdinalIgnoreCase))
                    ?? _projects.FirstOrDefault(p =>
                        !string.IsNullOrEmpty(p.POCode) &&
                        (p.POCode.IndexOf(extracted, StringComparison.OrdinalIgnoreCase) >= 0 ||
                         extracted.IndexOf(p.POCode, StringComparison.OrdinalIgnoreCase) >= 0));
            }

            if (proj == null)
            {
                item.Status = "❌ Không khớp";
                return item;
            }

            string matchCode = isMPR ? proj.MPRCode : proj.POCode;
            item.MatchedProject = $"{matchCode}  —  {proj.ProjectName}";

            string targetFolder = isMPR ? proj.MPRLink : proj.POLink;

            if (string.IsNullOrWhiteSpace(targetFolder))
            {
                item.TargetFolder = isMPR
                    ? "(MPR_link chưa cấu hình trong frmProject)"
                    : "(PO_Link chưa cấu hình trong frmProject)";
                item.Status = "⚠️ Chưa có Link";
                return item;
            }

            if (!Directory.Exists(targetFolder))
            {
                item.TargetFolder = targetFolder;
                item.Status = "⚠️ Thư mục đích không tồn tại";
                return item;
            }

            item.TargetFolder = targetFolder;
            item.FinalName = ResolveFileName(targetFolder, fileName);
            item.Status = item.FinalName != fileName
                ? "⚠️ Đổi tên"
                : "✅ Sẵn sàng";
            return item;
        }

        // ═════════════════════════════════════════════════════════════════
        // CUT FILE
        // ═════════════════════════════════════════════════════════════════
        private void BtnRun_Click(object sender, EventArgs e)
        {
            bool preview = _chkPreview.Checked;
            var toMove = _items.Where(i =>
                i.Status == "✅ Sẵn sàng" ||
                i.Status == "⚠️ Đổi tên").ToList();

            if (toMove.Count == 0)
            {
                MessageBox.Show("Không có file nào sẵn sàng để di chuyển.",
                    "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (MessageBox.Show(
                $"Chế độ: {(preview ? "PREVIEW (không di chuyển thực sự)" : "DI CHUYỂN THỰC SỰ")}\n\n" +
                $"Sẽ xử lý {toMove.Count} file.\nBạn có chắc chắn?",
                "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                != DialogResult.Yes) return;

            _btnRun.Enabled = false;
            _btnScan.Enabled = false;
            _progress.Visible = true;
            _progress.Maximum = toMove.Count;
            _progress.Value = 0;

            int ok = 0, fail = 0;
            var log = new StringBuilder();
            log.AppendLine($"=== FileSorter Log — {DateTime.Now:dd/MM/yyyy HH:mm:ss} ===");
            log.AppendLine($"Chế độ : {(preview ? "PREVIEW" : "CUT")}");
            log.AppendLine($"Nguồn  : {_txtSource.Text}");
            log.AppendLine();

            foreach (var item in toMove)
            {
                try
                {
                    // Tính lại FinalName tại thời điểm thực hiện (tránh race condition)
                    string finalName = ResolveFileName(item.TargetFolder, item.FileName);
                    string dest = Path.Combine(item.TargetFolder, finalName);

                    if (!preview)
                        File.Move(item.SourcePath, dest);

                    item.FinalName = finalName;
                    item.Status = preview ? "🔍 Preview OK" : "✅ Đã di chuyển";
                    log.AppendLine($"OK  | {item.FileType,-3} | {item.FileName}");
                    log.AppendLine($"      → {dest}");
                    ok++;
                }
                catch (Exception ex)
                {
                    item.Status = $"❌ {ex.Message}";
                    log.AppendLine($"LỖI | {item.FileType,-3} | {item.FileName}");
                    log.AppendLine($"      → {ex.Message}");
                    fail++;
                }

                _progress.Value++;
                Application.DoEvents();
            }

            log.AppendLine();
            log.AppendLine($"Kết quả: {ok} thành công, {fail} lỗi.");

            // Lưu log
            try
            {
                string logPath = Path.Combine(_txtSource.Text,
                    $"FileSorter_{DateTime.Now:yyyyMMdd_HHmmss}.log");
                File.WriteAllText(logPath, log.ToString(), Encoding.UTF8);
            }
            catch { }

            RefreshGrid();
            _progress.Visible = false;
            _btnScan.Enabled = true;
            _btnRun.Enabled = _items.Any(i =>
                i.Status == "✅ Sẵn sàng" || i.Status == "⚠️ Đổi tên");

            SetStatus(preview
                ? $"🔍 Preview: {ok} file sẽ được di chuyển, {fail} lỗi."
                : $"✅ Hoàn thành: {ok} file thành công, {fail} lỗi. Log đã lưu.");

            MessageBox.Show(
                $"{(preview ? "🔍 Preview xong" : "✅ Hoàn thành")}\n\n" +
                $"Thành công : {ok} file\n" +
                $"Lỗi        : {fail} file",
                preview ? "Preview" : "Hoàn thành",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ═════════════════════════════════════════════════════════════════
        // DB
        // ═════════════════════════════════════════════════════════════════
        private List<ProjectRec> LoadProjects()
        {
            var list = new List<ProjectRec>();
            try
            {
                using var conn = DatabaseHelper.GetConnection();
                conn.Open();
                var dt = new DataTable();
                new SqlDataAdapter(new SqlCommand(@"
                    SELECT WorkorderNo,
                           ISNULL(ProjectName,'') AS ProjectName,
                           ISNULL(POCode,'')      AS POCode,
                           ISNULL(MPRCode,'')     AS MPRCode,
                           ISNULL(PO_Link,'')     AS PO_Link,
                           ISNULL(MPR_link,'')    AS MPR_Link
                    FROM   ProjectInfo
                    WHERE  (POCode IS NOT NULL AND POCode <> '')
                        OR (MPRCode IS NOT NULL AND MPRCode <> '')
                    ORDER  BY LEN(ISNULL(POCode,'')) DESC", conn)).Fill(dt);

                foreach (DataRow r in dt.Rows)
                    list.Add(new ProjectRec
                    {
                        WorkorderNo = r["WorkorderNo"].ToString(),
                        ProjectName = r["ProjectName"].ToString(),
                        POCode = r["POCode"].ToString(),
                        MPRCode = r["MPRCode"].ToString(),
                        POLink = r["PO_Link"].ToString(),
                        MPRLink = r["MPR_Link"].ToString()
                    });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi tải dữ liệu DB:\n{ex.Message}",
                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return list;
        }

        // ═════════════════════════════════════════════════════════════════
        // HELPERS
        // ═════════════════════════════════════════════════════════════════

        /// <summary>
        /// Nếu file đã tồn tại trong thư mục đích → thêm _1, _2... cho đến khi không trùng.
        /// </summary>
        private static string ResolveFileName(string folder, string fileName)
        {
            if (!File.Exists(Path.Combine(folder, fileName)))
                return fileName;

            string name = Path.GetFileNameWithoutExtension(fileName);
            string ext = Path.GetExtension(fileName);
            for (int i = 1; i < 9999; i++)
            {
                string candidate = $"{name}_{i}{ext}";
                if (!File.Exists(Path.Combine(folder, candidate)))
                    return candidate;
            }
            return $"{name}_{Guid.NewGuid():N}{ext}";
        }

        private void RefreshGrid()
        {
            _dgv.Rows.Clear();
            foreach (var item in _items)
                _dgv.Rows.Add(
                    item.FileName,
                    item.FileType,
                    item.ExtractedCode ?? "-",
                    item.MatchedProject ?? "-",
                    item.TargetFolder ?? "-",
                    item.FinalName ?? item.FileName,
                    item.Status);
        }

        private void DgvFormat(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string status = _dgv.Rows[e.RowIndex]
                .Cells["Status"].Value?.ToString() ?? "";

            // Màu nền theo trạng thái
            Color bg =
                status == "✅ Sẵn sàng" ? Color.FromArgb(236, 253, 245) :
                status == "✅ Đã di chuyển" ? Color.FromArgb(212, 237, 218) :
                status == "🔍 Preview OK" ? Color.FromArgb(219, 234, 254) :
                status == "⚠️ Đổi tên" ? Color.FromArgb(254, 249, 195) :
                status == "❌ Không khớp" ? Color.FromArgb(249, 250, 251) :
                                               Color.FromArgb(254, 226, 226);

            _dgv.Rows[e.RowIndex].DefaultCellStyle.BackColor = bg;

            // Cột FileType — màu riêng
            if (_dgv.Columns[e.ColumnIndex].Name == "FileType")
            {
                e.CellStyle.ForeColor = e.Value?.ToString() == "MPR"
                    ? Color.FromArgb(29, 78, 216)
                    : Color.FromArgb(154, 52, 18);
                e.CellStyle.Font = new Font("Segoe UI", 8.5f, FontStyle.Bold);
                e.FormattingApplied = true;
            }
        }

        private void SetStatus(string msg) => _lblStatus.Text = msg;

        private static Button MkBtn(string text, Point loc, Size sz, Color bg)
        {
            var b = new Button
            {
                Text = text,
                Location = loc,
                Size = sz,
                FlatStyle = FlatStyle.Flat,
                BackColor = bg,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                TabStop = false
            };
            b.FlatAppearance.BorderSize = 0;
            return b;
        }

        private static Label MkLabel(string text, Point loc,
            bool bold = false, Color? color = null)
            => new Label
            {
                Text = text,
                Location = loc,
                AutoSize = true,
                Font = new Font("Segoe UI", 8.5f,
                    bold ? FontStyle.Bold : FontStyle.Regular),
                ForeColor = color ?? Color.FromArgb(33, 37, 41)
            };
    }
}