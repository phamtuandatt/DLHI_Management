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
            public string ProjectCode { get; set; } // mã dự án dùng để tìm đường dẫn
            public string POCode { get; set; }
            public string MPRCode { get; set; }
            public string POLink { get; set; }
            public string MPRLink { get; set; }
        }

        private Form TopOwner => (this.TopLevelControl as Form) ?? this;

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
                ReadOnly = false,  // cho phép edit cột FinalName
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
                ("FileName",       "Tên file gốc (double-click để mở)", 22, true),
                ("FileType",       "Loại ▼",                             5, true),
                ("ExtractedCode",  "Mã trích xuất",                     11, true),
                ("TargetFolder",   "Thư mục đích",                      27, true),
                ("FinalName",      "Tên lưu (click để sửa)",            13, false), // cho phép edit
                ("Status",         "Trạng thái",                        10, true),
            };
            foreach (var (n, h, w, ro) in cols)
                _dgv.Columns.Add(new DataGridViewTextBoxColumn
                { Name = n, HeaderText = h, FillWeight = w, ReadOnly = ro });

            // Cột "Dự án khớp" — dùng TextBox + EditingControlShowing để inject ComboBox
            var colProject = new DataGridViewTextBoxColumn
            {
                Name = "MatchedProject",
                HeaderText = "Dự án khớp ▼",
                FillWeight = 18,
            };
            _dgv.Columns.Insert(3, colProject);

            // Inject ComboBox khi user click vào cột MatchedProject
            _dgv.EditingControlShowing += (s, e) =>
            {
                if (_dgv.CurrentCell?.OwningColumn.Name != "MatchedProject") return;
                if (e.Control is TextBox tb)
                {
                    tb.ReadOnly = true; // không gõ tay, chỉ chọn từ dropdown
                }
            };
            _dgv.CellClick += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                if (_dgv.Columns[e.ColumnIndex].Name == "MatchedProject")
                    ShowProjectDropdown(e.RowIndex);
                else if (_dgv.Columns[e.ColumnIndex].Name == "FileType")
                    ShowFileTypeDropdown(e.RowIndex);
            };

            // Double-click cột FileName → mở file gốc
            _dgv.CellDoubleClick += (s, e) =>
            {
                if (e.RowIndex < 0 || e.RowIndex >= _items.Count) return;
                if (_dgv.Columns[e.ColumnIndex].Name != "FileName") return;
                string path = _items[e.RowIndex].SourcePath;
                if (File.Exists(path))
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    { FileName = path, UseShellExecute = true });
                else
                    MessageBox.Show(TopOwner, $"File không còn tồn tại:\n{path}",
                        "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            };

            // Cho phép sửa tên trực tiếp tại cột FinalName
            _dgv.CellEndEdit += (s, e) =>
            {
                if (e.RowIndex < 0 || e.RowIndex >= _items.Count) return;
                if (_dgv.Columns[e.ColumnIndex].Name != "FinalName") return;

                string newName = _dgv.Rows[e.RowIndex]
                    .Cells["FinalName"].Value?.ToString()?.Trim() ?? "";
                var item = _items[e.RowIndex];

                if (string.IsNullOrEmpty(newName) || newName == item.FinalName) return;

                // Kiểm tra tên hợp lệ
                if (newName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                {
                    MessageBox.Show(TopOwner, "Tên file chứa ký tự không hợp lệ.",
                        "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    _dgv.Rows[e.RowIndex].Cells["FinalName"].Value = item.FinalName;
                    return;
                }

                // Đổi tên file gốc ngay trên disk
                if (!File.Exists(item.SourcePath))
                {
                    MessageBox.Show(TopOwner, $"File gốc không còn tồn tại:\n{item.SourcePath}",
                        "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    _dgv.Rows[e.RowIndex].Cells["FinalName"].Value = item.FinalName;
                    return;
                }

                string dir = Path.GetDirectoryName(item.SourcePath)!;
                string newPath = Path.Combine(dir, newName);

                if (File.Exists(newPath))
                {
                    MessageBox.Show(TopOwner, $"File \"{newName}\" đã tồn tại trong thư mục nguồn.",
                        "Trùng tên", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    _dgv.Rows[e.RowIndex].Cells["FinalName"].Value = item.FinalName;
                    return;
                }

                try
                {
                    File.Move(item.SourcePath, newPath);
                    item.SourcePath = newPath;
                    item.FileName = newName;
                    item.FinalName = newName;
                    _dgv.Rows[e.RowIndex].Cells["FileName"].Value = newName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(TopOwner, $"Không thể đổi tên:\n{ex.Message}",
                        "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    _dgv.Rows[e.RowIndex].Cells["FinalName"].Value = item.FinalName;
                }
            };

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
                MessageBox.Show(TopOwner, $"Thư mục không tồn tại:\n{src}",
                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _btnScan.Enabled = false;
            _btnRun.Enabled = false;
            SetStatus("⏳ Đang chạy OCR...");

            Task.Run(() =>
            {
                // ── OCR (chạy ngầm, không block UI) ─────────────────────
                try
                {
                    var psi = new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = "cmd.exe",
                        Arguments = "/c python test_po_ocr.py",
                        UseShellExecute = false,
                        RedirectStandardOutput = false, // không redirect stdout tránh deadlock
                        RedirectStandardError = true,
                        CreateNoWindow = true,
                        WorkingDirectory = @"C:\Users\PCPV"
                    };
                    using var proc = System.Diagnostics.Process.Start(psi);
                    string stderr = proc.StandardError.ReadToEnd();
                    proc.WaitForExit();

                    if (proc.ExitCode != 0)
                    {
                        DialogResult res = DialogResult.No;
                        this.Invoke(() => res = MessageBox.Show(TopOwner,
                            $"Script OCR kết thúc với lỗi:\n{stderr}\n\nBạn có muốn tiếp tục quét không?",
                            "Cảnh báo OCR", MessageBoxButtons.YesNo, MessageBoxIcon.Warning));
                        if (res != DialogResult.Yes)
                        {
                            this.Invoke(() => { SetStatus("Đã hủy."); _btnScan.Enabled = true; });
                            return;
                        }
                    }
                    else
                        this.Invoke(() => SetStatus("✅ OCR hoàn thành. Tiếp tục quét file..."));
                }
                catch (Exception ex)
                {
                    DialogResult res = DialogResult.No;
                    this.Invoke(() => res = MessageBox.Show(TopOwner,
                        $"Không thể chạy OCR:\n{ex.Message}\n\nBạn có muốn tiếp tục quét không?",
                        "Cảnh báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning));
                    if (res != DialogResult.Yes)
                    {
                        this.Invoke(() => { SetStatus("Đã hủy."); _btnScan.Enabled = true; });
                        return;
                    }
                }

                // ── Load projects từ DB ──────────────────────────────────
                this.Invoke(() => SetStatus("⏳ Đang tải dữ liệu dự án từ DB..."));
                var projects = LoadProjects();

                if (projects.Count == 0)
                {
                    this.Invoke(() =>
                    {
                        SetStatus("⚠️ Không tìm thấy dự án nào trong DB (bảng ProjectInfo).");
                        _btnScan.Enabled = true;
                    });
                    return;
                }

                // ── Đọc file ─────────────────────────────────────────────
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
                    this.Invoke(() => { SetStatus($"❌ Lỗi đọc thư mục: {ex.Message}"); _btnScan.Enabled = true; });
                    return;
                }

                if (files.Length == 0)
                {
                    this.Invoke(() => { SetStatus("📂 Thư mục không có file nào."); _btnScan.Enabled = true; });
                    return;
                }

                // ── Phân tích (có thể dùng projects mới load) ────────────
                this.Invoke(() => SetStatus("⏳ Đang quét file..."));
                _projects = projects; // cập nhật trước khi gọi AnalyzeFile
                var newItems = new List<SortItem>();
                foreach (string fp in files)
                    newItems.Add(AnalyzeFile(Path.GetFileName(fp), fp));

                // ── Cập nhật UI trên UI thread ───────────────────────────
                this.Invoke(() =>
                {
                    _items = newItems;
                    _dgv.Rows.Clear();
                    PopulateProjectCombo();
                    RefreshGrid();

                    int ready   = _items.Count(i => i.Status == "✅ Sẵn sàng");
                    int renamed = _items.Count(i => i.Status == "⚠️ Đổi tên");
                    int noMatch = _items.Count(i => i.Status == "❌ Không khớp");
                    int err     = _items.Count(i =>
                        i.Status != "✅ Sẵn sàng" &&
                        i.Status != "⚠️ Đổi tên" &&
                        i.Status != "❌ Không khớp");

                    SetStatus($"🔍 {files.Length} file  |  " +
                        $"✅ {ready} sẵn sàng  |  " +
                        $"⚠️ {renamed} cần đổi tên  |  " +
                        $"❌ {noMatch} không khớp  |  " +
                        $"⚠️ {err} lỗi link");

                    _btnScan.Enabled = true;
                    _btnRun.Enabled = _items.Any(i =>
                        i.Status == "✅ Sẵn sàng" || i.Status == "⚠️ Đổi tên");
                });
            });
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

            // ── Phân loại theo nội dung tên file ─────────────────────────
            bool hasMPR = fileName.IndexOf("-MPR-", StringComparison.OrdinalIgnoreCase) >= 0;
            bool hasPC = fileName.IndexOf("-PC-", StringComparison.OrdinalIgnoreCase) >= 0;

            if (hasPC && !hasMPR)
            {
                // ── PO: DV-{POCode}-PC-{số} → trích {POCode} ─────────────
                var mPO = RexPO.Match(fileName);
                if (mPO.Success)
                {
                    extracted = mPO.Groups[1].Value; // VD: "GHA"
                    isMPR = false;
                }
            }
            else if (hasMPR)
            {
                // ── MPR: lấy tất cả trước dấu "-" cuối cùng (kể cả "-") ──
                // VD: "DV-SL-2511-ABT-MPR-001.pdf" → "DV-SL-2511-ABT-MPR-"
                var mMPR = RexMPR.Match(fileName);
                if (mMPR.Success)
                {
                    extracted = mMPR.Groups[1].Value;
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

            string poCode2 = isMPR
                ? (!string.IsNullOrEmpty(proj.MPRCode) ? proj.MPRCode : proj.WorkorderNo)
                : (!string.IsNullOrEmpty(proj.POCode) ? proj.POCode : proj.WorkorderNo);
            item.MatchedProject = $"{poCode2} - {proj.ProjectCode}";

            string targetFolder = isMPR ? proj.MPRLink : proj.POLink;

            if (string.IsNullOrWhiteSpace(targetFolder))
            {
                item.TargetFolder = isMPR
                    ? "(MPR_link chưa cấu hình trong frmProject)"
                    : "(PO_Link chưa cấu hình trong frmProject)";
                item.Status = "⚠️ Chưa có Link";
                return item;
            }

            item.TargetFolder = targetFolder;

            // PO: đổi thư mục Excel → PDF (scan file lưu vào PDF, không ảnh hưởng DB)
            if (!isMPR)
                targetFolder = Regex.Replace(targetFolder,
                    @"\\Excel$", @"\PDF",
                    RegexOptions.IgnoreCase);

            if (!Directory.Exists(targetFolder))
            {
                item.TargetFolder = targetFolder;
                item.Status = "⚠️ Thư mục PDF không tồn tại";
                return item;
            }

            item.TargetFolder = targetFolder;

            // PO: preview tên mới = {FinalName_không_ext}_{ShortName}.ext
            string previewFileName = item.FileName;
            if (!isMPR)
            {
                // Lấy PONo từ tên file hiện tại (bỏ extension)
                string poNo = Path.GetFileNameWithoutExtension(item.FileName);
                string shortName = GetSupplierShortName(poNo);
                if (!string.IsNullOrEmpty(shortName))
                {
                    string ext = Path.GetExtension(item.FileName);
                    previewFileName = $"{poNo}_{shortName}{ext}";
                }
            }
            item.FinalName = ResolveFileName(targetFolder, previewFileName);
            item.Status = "✅ Sẵn sàng";
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
                MessageBox.Show(TopOwner, "Không có file nào sẵn sàng để di chuyển.",
                    "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (MessageBox.Show(TopOwner,
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
                    string destFileName = item.FinalName ?? item.FileName;

                    // ── File PO: lấy PONo từ FinalName (bỏ extension) → tìm NCC ──
                    if (item.FileType == "PO")
                    {
                        // PONo = FinalName không có extension
                        // VD: "DV-GHA-PC-067.pdf" → "DV-GHA-PC-067"
                        string poNo = Path.GetFileNameWithoutExtension(item.FinalName ?? item.FileName);
                        string shortName = GetSupplierShortName(poNo);

                        if (!string.IsNullOrEmpty(shortName))
                        {
                            string ext = Path.GetExtension(item.FileName);
                            destFileName = $"{poNo}_{shortName}{ext}";
                        }
                    }

                    destFileName = ResolveFileName(item.TargetFolder, destFileName);
                    string dest = Path.Combine(item.TargetFolder, destFileName);

                    if (!preview)
                        File.Move(item.SourcePath, dest);

                    item.FinalName = destFileName;
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

            MessageBox.Show(TopOwner,
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
                           ISNULL(ProjectName,'')  AS ProjectName,
                           ISNULL(ProjectCode,'')  AS ProjectCode,
                           ISNULL(POCode,'')       AS POCode,
                           ISNULL(MPRCode,'')      AS MPRCode,
                           ISNULL(PO_Link,'')      AS PO_Link,
                           ISNULL(MPR_link,'')     AS MPR_Link
                    FROM   ProjectInfo
                    WHERE  (POCode IS NOT NULL AND POCode <> '')
                        OR (MPRCode IS NOT NULL AND MPRCode <> '')
                    ORDER  BY LEN(ISNULL(POCode,'')) DESC", conn)).Fill(dt);

                foreach (DataRow r in dt.Rows)
                    list.Add(new ProjectRec
                    {
                        WorkorderNo = r["WorkorderNo"].ToString(),
                        ProjectName = r["ProjectName"].ToString(),
                        ProjectCode = r["ProjectCode"].ToString(),
                        POCode = r["POCode"].ToString(),
                        MPRCode = r["MPRCode"].ToString(),
                        POLink = r["PO_Link"].ToString(),
                        MPRLink = r["MPR_Link"].ToString()
                    });
            }
            catch (Exception ex)
            {
                MessageBox.Show(TopOwner, $"Lỗi tải dữ liệu DB:\n{ex.Message}",
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
        /// <summary>
        /// Trích PONo từ tên file PO.
        /// VD: "DV-GHA-PC-067_scan.pdf" → "DV-GHA-PC-067"
        /// </summary>
        private string ExtractPONo(string fileName)
        {
            // PONo dạng DV-{POCode}-PC-{số} — lấy phần này từ tên file
            var m = RexPO.Match(fileName);
            if (!m.Success) return null;

            // Reconstruct PONo = DV-{group1}-PC-{số}
            // Lấy toàn bộ match trước dấu _ hoặc khoảng trắng hoặc hết
            string raw = Path.GetFileNameWithoutExtension(fileName);
            // Tìm vị trí match và lấy đúng độ dài
            int start = m.Index;
            // Tìm ký tự kết thúc PONo (dấu _ hoặc khoảng trắng hoặc hết chuỗi)
            int end = m.Index + m.Length;
            // m.Length bao gồm cả số PO → đây chính là PONo đầy đủ
            return raw.Substring(start, Math.Min(end, raw.Length) - start);
        }

        /// <summary>
        /// Lấy Short_Name của NCC từ DB theo PONo.
        /// Query: PO_head JOIN Suppliers WHERE PONo = @poNo
        /// </summary>
        private string GetSupplierShortName(string poNo)
        {
            if (string.IsNullOrEmpty(poNo)) return null;
            try
            {
                using var conn = DatabaseHelper.GetConnection();
                conn.Open();
                var cmd = new SqlCommand(@"
                    SELECT TOP 1 ISNULL(s.Short_Name, s.Company_Name)
                    FROM   PO_head   h
                    JOIN   Suppliers s ON s.Supplier_ID = h.Supplier_ID
                    WHERE  h.PONo = @poNo
                    ORDER  BY h.Revise DESC", conn);
                cmd.Parameters.AddWithValue("@poNo", poNo);
                string result = cmd.ExecuteScalar()?.ToString();

                // Loại bỏ ký tự không hợp lệ trong tên file
                if (!string.IsNullOrEmpty(result))
                    result = string.Concat(result
                        .Split(Path.GetInvalidFileNameChars()))
                        .Trim();
                return result;
            }
            catch { return null; }
        }

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

        private void PopulateProjectCombo() { } // không còn dùng

        /// <summary>
        /// Hiện ContextMenuStrip danh sách dự án khi click vào cột "Dự án khớp".
        /// Dùng ContextMenuStrip thay vì ComboBox — đáng tin cậy hơn trong DataGridView.
        /// </summary>
        private void ShowProjectDropdown(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= _items.Count) return;
            var item = _items[rowIndex];

            var menu = new ContextMenuStrip();
            menu.Items.Add(new ToolStripMenuItem("— chưa chọn —")
            {
                Tag = (ProjectRec)null
            });
            menu.Items.Add(new ToolStripSeparator());

            foreach (var p in _projects)
            {
                // Hiển thị: "POCode - ProjectCode" (dùng ProjectCode để tìm đường dẫn)
                string poCode = !string.IsNullOrEmpty(p.POCode) ? p.POCode :
                                 !string.IsNullOrEmpty(p.MPRCode) ? p.MPRCode : p.WorkorderNo;
                string display = $"{poCode} - {p.ProjectCode}";
                var mi = new ToolStripMenuItem(display) { Tag = p };
                menu.Items.Add(mi);
            }

            menu.ItemClicked += (s, e) =>
            {
                var proj = e.ClickedItem.Tag as ProjectRec;
                ApplyProjectSelection(rowIndex, proj);
            };

            // Hiện menu tại vị trí cell
            var cellRect = _dgv.GetCellDisplayRectangle(
                _dgv.Columns["MatchedProject"].Index, rowIndex, false);
            menu.Show(_dgv, new Point(cellRect.X, cellRect.Bottom));
        }

        private void ShowFileTypeDropdown(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= _items.Count) return;
            var item = _items[rowIndex];

            var menu = new ContextMenuStrip();
            foreach (string type in new[] { "PO", "MPR", "-" })
            {
                var mi = new ToolStripMenuItem(type) { Tag = type };
                if (item.FileType == type) mi.Checked = true;
                menu.Items.Add(mi);
            }

            menu.ItemClicked += (s, e) =>
            {
                string newType = e.ClickedItem.Tag as string;
                if (newType == item.FileType) return;
                item.FileType = newType;
                _dgv.Rows[rowIndex].Cells["FileType"].Value = newType;

                // Re-resolve target folder using current matched project
                ProjectRec proj = null;
                string currentMatch = item.MatchedProject ?? "";
                if (!string.IsNullOrEmpty(currentMatch) && currentMatch != "-")
                    proj = _projects.FirstOrDefault(p =>
                    {
                        string code = !string.IsNullOrEmpty(p.POCode) ? p.POCode :
                                      !string.IsNullOrEmpty(p.MPRCode) ? p.MPRCode : p.WorkorderNo;
                        return $"{code} - {p.ProjectCode}" == currentMatch;
                    });

                ApplyProjectSelection(rowIndex, proj);
            };

            var cellRect = _dgv.GetCellDisplayRectangle(
                _dgv.Columns["FileType"].Index, rowIndex, false);
            menu.Show(_dgv, new Point(cellRect.X, cellRect.Bottom));
        }

        private void ApplyProjectSelection(int rowIndex, ProjectRec proj)
        {
            if (rowIndex < 0 || rowIndex >= _items.Count) return;
            var item = _items[rowIndex];

            if (proj == null)
            {
                item.MatchedProject = "";
                item.TargetFolder = "-";
                item.FinalName = item.FileName;
                item.Status = "❌ Không khớp";
                RefreshRow(rowIndex);
                return;
            }

            bool isMPR = item.FileType == "MPR";
            string poCode = !string.IsNullOrEmpty(proj.POCode) ? proj.POCode :
                               !string.IsNullOrEmpty(proj.MPRCode) ? proj.MPRCode : proj.WorkorderNo;
            item.MatchedProject = $"{poCode} - {proj.ProjectCode}";

            string targetFolder = isMPR ? proj.MPRLink : proj.POLink;

            // PO: đổi \Excel → \PDF
            if (!isMPR && !string.IsNullOrEmpty(targetFolder))
                targetFolder = Regex.Replace(targetFolder,
                    @"\\Excel$", @"\PDF", RegexOptions.IgnoreCase);

            if (string.IsNullOrWhiteSpace(targetFolder))
            {
                item.TargetFolder = "(Chưa có Link)";
                item.Status = "⚠️ Chưa có Link";
            }
            else if (!Directory.Exists(targetFolder))
            {
                item.TargetFolder = targetFolder;
                item.Status = "⚠️ Thư mục đích không tồn tại";
            }
            else
            {
                item.TargetFolder = targetFolder;
                item.FinalName = ResolveFileName(targetFolder, item.FileName);
                item.Status = item.FinalName != item.FileName
                    ? "⚠️ Đổi tên" : "✅ Sẵn sàng";
            }

            RefreshRow(rowIndex);
            _btnRun.Enabled = _items.Any(i =>
                i.Status == "✅ Sẵn sàng" || i.Status == "⚠️ Đổi tên");
        }

        private void RefreshRow(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= _dgv.Rows.Count) return;
            var item = _items[rowIndex];
            var row = _dgv.Rows[rowIndex];
            row.Cells["MatchedProject"].Value = item.MatchedProject ?? "-";
            row.Cells["TargetFolder"].Value = item.TargetFolder ?? "-";
            row.Cells["FinalName"].Value = item.FinalName ?? item.FileName;
            row.Cells["Status"].Value = item.Status;
            _dgv.InvalidateRow(rowIndex);
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