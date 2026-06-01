using Microsoft.Data.SqlClient;
using MPR_Managerment.Forms.MPRGUI;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;
using MPR_Managerment.Services;
using OfficeOpenXml;
using Syncfusion.XlsIO.Implementation.XmlSerialization;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Linq;

namespace MPR_Managerment.Forms
{
    public partial class frmMPR : Form
    {
        private MPRService _service = new MPRService();
        // Backup PO links khi Admin xóa và insert lại Details — dùng để re-link sau insert
        private Dictionary<int, List<int>> _adminPoLinkBackup = new Dictionary<int, List<int>>();

        // Đảm bảo cột Is_Deleted tồn tại trong DB (chạy 1 lần, idempotent)
        private static void EnsureIsDeletedColumn()
        {
            using var conn = Helpers.DatabaseHelper.GetConnection();
            try { conn.Open(); } catch { return; }

            try
            {
                new SqlCommand(@"
                    IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
                                  WHERE TABLE_NAME='MPR_Details' AND COLUMN_NAME='Is_Deleted')
                        ALTER TABLE MPR_Details ADD Is_Deleted BIT NOT NULL DEFAULT 0", conn)
                    .ExecuteNonQuery();
            }
            catch { }

            try
            {
                new SqlCommand(@"
                    UPDATE MPR_Header
                    SET Rev = CAST(FLOOR(TRY_CAST(TRY_CAST(Rev AS DECIMAL(10,2)) AS INT)) AS VARCHAR(10))
                    WHERE TRY_CAST(TRY_CAST(Rev AS DECIMAL(10,2)) AS INT) IS NULL
                       OR Rev LIKE '%.%'
                       OR (LEN(Rev) > 1 AND LEFT(Rev,1) = '0')", conn)
                    .ExecuteNonQuery();
            }
            catch { }

            try
            {
                new SqlCommand(@"
                    IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
                                  WHERE TABLE_NAME='MPR_Header' AND COLUMN_NAME='Is_Latest')
                    BEGIN
                        ALTER TABLE MPR_Header ADD Is_Latest BIT NOT NULL DEFAULT 1;
                        UPDATE MPR_Header SET Is_Latest = 0;
                        UPDATE h SET h.Is_Latest = 1
                        FROM MPR_Header h
                        INNER JOIN (
                            SELECT Project_Code, MAX(TRY_CAST(TRY_CAST(Rev AS DECIMAL(10,2)) AS INT)) AS MaxRev,
                                   CASE WHEN MPR_No LIKE '%_Rev.%'
                                        THEN SUBSTRING(MPR_No,1,CHARINDEX('_Rev.',MPR_No)-1)
                                        ELSE MPR_No END AS BaseNo
                            FROM MPR_Header
                            GROUP BY Project_Code,
                                CASE WHEN MPR_No LIKE '%_Rev.%'
                                     THEN SUBSTRING(MPR_No,1,CHARINDEX('_Rev.',MPR_No)-1)
                                     ELSE MPR_No END
                        ) mx ON mx.Project_Code = h.Project_Code
                               AND mx.MaxRev = TRY_CAST(TRY_CAST(h.Rev AS DECIMAL(10,2)) AS INT)
                               AND (
                                  CASE WHEN h.MPR_No LIKE '%_Rev.%'
                                       THEN SUBSTRING(h.MPR_No,1,CHARINDEX('_Rev.',h.MPR_No)-1)
                                       ELSE h.MPR_No END = mx.BaseNo
                               )
                    END", conn)
                    .ExecuteNonQuery();
            }
            catch { }

            try
            {
                new SqlCommand(@"
                    IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
                                  WHERE TABLE_NAME='MPR_Header' AND COLUMN_NAME='Email_Status')
                        ALTER TABLE MPR_Header ADD Email_Status NVARCHAR(20) NULL DEFAULT '';
                    IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
                                  WHERE TABLE_NAME='MPR_Header' AND COLUMN_NAME='Email_Sent_By')
                        ALTER TABLE MPR_Header ADD Email_Sent_By NVARCHAR(100) NULL;
                    IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
                                  WHERE TABLE_NAME='MPR_Header' AND COLUMN_NAME='Email_Sent_At')
                        ALTER TABLE MPR_Header ADD Email_Sent_At DATETIME NULL;", conn)
                    .ExecuteNonQuery();
            }
            catch { }
        }

        private List<MPRHeader> _mprList = new List<MPRHeader>();
        private List<MPRDetail> _details = new List<MPRDetail>();
        private int _selectedMPR_ID = 0;
        private string _currentUser = "Admin";
        private readonly Dictionary<int, System.Windows.Forms.Timer> _emailPollers = new();

        // Thêm biến để lưu ID truyền từ Dashboard sang
        private int _targetMprId = 0;

        private DataGridView dgvMPR;
        private TextBox txtSearch;
        private Button btnSearch, btnNewMPR, btnSaveHeader, btnDeleteMPR, btnClearHeader;
        private Label lblStatus;

        private TextBox txtMPRNo, txtProjectName, txtProjectCode, txtDepartment, txtRequestor, txtRev, txtNotes;
        private DateTimePicker dtpRequiredDate;
        private ComboBox cboStatus;

        // BẢNG MỚI: Danh sách file đính kèm
        private DataGridView dgvFiles;

        // File preview popup for MPR files
        private Form _filePreviewForm_MPR;
        private System.Windows.Forms.Timer _previewTimer_MPR;
        private string _previewPath_MPR = null;
        private DataGridViewCell _previewCell_MPR = null;
        private class PreviewStateMPR
        {
            public string Type;
            public Image OriginalImage;
            public double Scale = 1.0;
            public string FullText;
            public string[] Lines;
            public int TextOffset = 0;
            public int TextPageLines = 200;
        }

        private DataGridView dgvDetails;

        // Preview helpers for MPR files
        private void EnsureFilePreview_MPR()
        {
            if (_previewTimer_MPR != null) return;
            _previewTimer_MPR = new System.Windows.Forms.Timer { Interval = 300 };
            _previewTimer_MPR.Tick += (s, e) =>
            {
                _previewTimer_MPR.Stop();
                if (!string.IsNullOrEmpty(_previewPath_MPR) && _previewCell_MPR != null)
                {
                    try { ShowFilePreview_MPR(_previewPath_MPR, Cursor.Position); }
                    catch { }
                }
            };

            if (dgvFiles != null)
            {
                dgvFiles.CellMouseEnter += DgvFiles_CellMouseEnter_MPR;
                dgvFiles.CellMouseLeave += DgvFiles_CellMouseLeave_MPR;
                dgvFiles.MouseLeave += (s, e) => StartPreviewCloseTimer_MPR();
                dgvFiles.Scroll += (s, e) => StartPreviewCloseTimer_MPR();
            }
        }

        private void DgvFiles_CellMouseEnter_MPR(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
                var cell = dgvFiles.Rows[e.RowIndex].Cells[e.ColumnIndex];
                var fullPathCell = dgvFiles.Rows[e.RowIndex].Cells.Count > 1 ? dgvFiles.Rows[e.RowIndex].Cells[1] : null;
                string path = fullPathCell?.Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(path)) { _previewPath_MPR = null; _previewCell_MPR = null; return; }
                _previewPath_MPR = path;
                _previewCell_MPR = cell;
                _previewTimer_MPR?.Stop();
                _previewTimer_MPR?.Start();
            }
            catch { }
        }

        private void DgvFiles_CellMouseLeave_MPR(object sender, DataGridViewCellEventArgs e)
        {
            _previewTimer_MPR?.Stop();
            _previewPath_MPR = null; _previewCell_MPR = null;
            StartPreviewCloseTimer_MPR();
        }

        private void HideFilePreview_MPR()
        {
            try
            {
                if (_filePreviewForm_MPR != null)
                {
                    if (!_filePreviewForm_MPR.IsDisposed) _filePreviewForm_MPR.Hide();
                    _filePreviewForm_MPR.Dispose();
                    _filePreviewForm_MPR = null;
                }
            }
            catch { }
        }

        private System.Windows.Forms.Timer _previewCloseTimer_MPR;
        private bool _printDialogOpen_MPR = false; // giữ preview mở trong khi hộp thoại in đang hiển thị

        private void StartPreviewCloseTimer_MPR(int intervalMs = 800)
        {
            try
            {
                if (_printDialogOpen_MPR) return; // Không đóng khi hộp thoại in đang mở
                if (_previewCloseTimer_MPR == null)
                {
                    _previewCloseTimer_MPR = new System.Windows.Forms.Timer { Interval = intervalMs };
                    _previewCloseTimer_MPR.Tick += (s, e) =>
                    {
                        _previewCloseTimer_MPR.Stop();
                        HideFilePreview_MPR();
                    };
                }
                _previewCloseTimer_MPR.Interval = intervalMs;
                _previewCloseTimer_MPR.Stop();
                _previewCloseTimer_MPR.Start();
            }
            catch { }
        }

        private void StopPreviewCloseTimer_MPR()
        {
            try { _previewCloseTimer_MPR?.Stop(); }
            catch { }
        }

        private void ShowFilePreview_MPR(string filePath, Point screenPos)
        {
            HideFilePreview_MPR();
            if (!File.Exists(filePath)) return;
            Form _previewOwner = this.TopLevelControl as Form ?? this;

            var f = new Form
            {
                FormBorderStyle = FormBorderStyle.SizableToolWindow,
                Size            = new Size(1024, 768),
                StartPosition   = FormStartPosition.Manual,
                ShowInTaskbar   = false,
                BackColor       = Color.FromArgb(38, 38, 38),
                Text            = Path.GetFileName(filePath)
            };
            var scr = Screen.FromPoint(screenPos)?.WorkingArea ?? new Rectangle(0, 0, 1024, 768);
            f.Bounds = new Rectangle(
                scr.Left + (scr.Width  - f.Width)  / 2,
                scr.Top  + (scr.Height - f.Height) / 2,
                f.Width, f.Height);

            // ── Toolbar ──────────────────────────────────────────────────────
            var toolbar  = new Panel { Dock = DockStyle.Top, Height = 38, BackColor = Color.FromArgb(45, 45, 48) };
            var clrDark  = Color.FromArgb(62, 62, 66);
            var clrBlue  = Color.FromArgb(0, 100, 180);

            Button MkBtn(string txt, int left, int w, Color bg)
            {
                var b = new Button
                {
                    Text = txt, Left = left, Top = 5, Width = w, Height = 28,
                    FlatStyle = FlatStyle.Flat, BackColor = bg,
                    ForeColor = Color.White,
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    Cursor = Cursors.Hand
                };
                b.FlatAppearance.BorderColor = Color.FromArgb(80, 80, 80);
                toolbar.Controls.Add(b);
                return b;
            }

            var btnZoomIn  = MkBtn("＋",    8, 36, clrDark);
            var btnZoomOut = MkBtn("－",   48, 36, clrDark);
            var btnFit     = MkBtn("⊡ Fit", 88, 60, clrDark);
            var btn100     = MkBtn("1:1",  152, 44, clrBlue);
            var lblZoom    = new Label { Left = 202, Top = 11, Width = 58, Text = "—", ForeColor = Color.FromArgb(220, 220, 220), Font = new Font("Segoe UI", 9, FontStyle.Bold), BackColor = Color.Transparent };
            var btnPrev    = MkBtn("◀",   268, 36, clrDark); btnPrev.Visible = false;
            var btnNext    = MkBtn("▶",   308, 36, clrDark); btnNext.Visible = false;
            var lblInfo    = new Label { Left = 352, Top = 11, Width = 360, Text = Path.GetFileName(filePath), ForeColor = Color.FromArgb(155, 155, 155), Font = new Font("Segoe UI", 8.5f), BackColor = Color.Transparent };
            // Add labels before right-side buttons so buttons have higher z-order
            toolbar.Controls.Add(lblZoom);
            toolbar.Controls.Add(lblInfo);
            var btnPrint = MkBtn("⎙ In file", 0, 90, Color.FromArgb(80, 60, 140));
            var btnOpen  = MkBtn("↗ Mở file", 0, 84, Color.FromArgb(0, 120, 60));
            var btnClose = MkBtn("✖",          0, 36, Color.FromArgb(196, 43, 28));
            // Position right-side buttons dynamically — Anchor fails when toolbar width is 0 at button creation
            void PlaceRightBtns()
            {
                int r = toolbar.ClientSize.Width - 4;
                btnClose.Left = r - btnClose.Width; r = btnClose.Left - 4;
                btnOpen.Left  = r - btnOpen.Width;  r = btnOpen.Left  - 4;
                btnPrint.Left = r - btnPrint.Width;
            }
            toolbar.Resize += (s, e) => PlaceRightBtns();
            f.Controls.Add(toolbar);   // triggers first Resize → buttons placed correctly

            Action printAction = null;
            btnPrint.Click += (s, e) =>
            {
                try
                {
                    StopPreviewCloseTimer_MPR();
                    _printDialogOpen_MPR = true;
                    printAction?.Invoke();
                }
                catch { }
                finally
                {
                    _printDialogOpen_MPR = false;
                    // Sau khi hộp thoại in đóng: nếu chuột đã ra ngoài Preview thì đợi 1 giây rồi đóng
                    var pf = _filePreviewForm_MPR;
                    if (pf != null && !pf.IsDisposed)
                    {
                        var pt = pf.PointToClient(Cursor.Position);
                        if (!pf.ClientRectangle.Contains(pt))
                            StartPreviewCloseTimer_MPR(1000);
                    }
                }
            };
            btnOpen .Click += (s, e) => { try { Process.Start(new ProcessStartInfo { FileName = filePath, UseShellExecute = true }); } catch { } };
            btnClose.Click += (s, e) => f.Close();
            f.KeyPreview = true;
            f.KeyDown   += (s, e) => { if (e.KeyCode == Keys.Escape) f.Close(); };

            // ── Content ──────────────────────────────────────────────────────
            string ext      = Path.GetExtension(filePath).ToLowerInvariant();
            bool   isImgExt = ext == ".png" || ext == ".jpg" || ext == ".jpeg" || ext == ".bmp" || ext == ".gif";

            if (isImgExt || ext == ".pdf")
            {
                Bitmap origBmp = null;
                try
                {
                    Image src = null;
                    if (ext == ".pdf") src = GetShellThumbnail_MPR(filePath, 1800, 2400);
                    if (src == null && isImgExt) src = Image.FromFile(filePath);
                    if (src != null)
                    {
                        if (src is Bitmap bSrc) origBmp = bSrc;
                        else { origBmp = new Bitmap(src); src.Dispose(); }
                    }
                }
                catch { origBmp = null; }

                if (origBmp != null)
                {
                    double scale  = 1.0;
                    float  offX   = 0, offY = 0;
                    bool   fitted = false;
                    bool   dragging = false;
                    Point  dragStart = Point.Empty;
                    float  dragOX = 0, dragOY = 0;

                    var canvas = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(38, 38, 38) };
                    typeof(Panel).GetProperty("DoubleBuffered",
                        System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                        ?.SetValue(canvas, true);

                    void CalcFit()
                    {
                        if (canvas.ClientSize.Width <= 0 || canvas.ClientSize.Height <= 0) return;
                        double sx = (double)canvas.ClientSize.Width  / origBmp.Width;
                        double sy = (double)canvas.ClientSize.Height / origBmp.Height;
                        scale = Math.Min(sx, sy) * 0.97;
                        offX  = (canvas.ClientSize.Width  - (float)(origBmp.Width  * scale)) / 2f;
                        offY  = (canvas.ClientSize.Height - (float)(origBmp.Height * scale)) / 2f;
                    }

                    void ClampPan()
                    {
                        float dw = (float)(origBmp.Width  * scale);
                        float dh = (float)(origBmp.Height * scale);
                        if (dw <= canvas.ClientSize.Width)  offX = (canvas.ClientSize.Width  - dw) / 2f;
                        else offX = Math.Max(canvas.ClientSize.Width  - dw, Math.Min(0, offX));
                        if (dh <= canvas.ClientSize.Height) offY = (canvas.ClientSize.Height - dh) / 2f;
                        else offY = Math.Max(canvas.ClientSize.Height - dh, Math.Min(0, offY));
                    }

                    void UpdateZoomLabel() => lblZoom.Text = $"{(int)(scale * 100)}%";

                    void ZoomAt(int px, int py, double factor)
                    {
                        double ns = Math.Max(0.03, Math.Min(30.0, scale * factor));
                        float  r  = (float)(ns / scale);
                        offX  = px - r * (px - offX);
                        offY  = py - r * (py - offY);
                        scale = ns;
                        ClampPan();
                        UpdateZoomLabel();
                        canvas.Invalidate();
                    }

                    canvas.Paint += (s, e) =>
                    {
                        if (!fitted && canvas.ClientSize.Width > 0 && canvas.ClientSize.Height > 0)
                        {
                            CalcFit(); fitted = true; UpdateZoomLabel();
                        }
                        var g = e.Graphics;
                        g.Clear(Color.FromArgb(38, 38, 38));
                        g.InterpolationMode  = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                        g.SmoothingMode      = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                        g.PixelOffsetMode    = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                        g.DrawImage(origBmp, offX, offY,
                            (float)(origBmp.Width * scale), (float)(origBmp.Height * scale));
                    };

                    canvas.MouseDown  += (s, e) =>
                    {
                        if (e.Button != MouseButtons.Left) return;
                        dragging = true; dragStart = e.Location; dragOX = offX; dragOY = offY;
                        canvas.Cursor = Cursors.SizeAll;
                    };
                    canvas.MouseMove  += (s, e) =>
                    {
                        if (!dragging) return;
                        offX = dragOX + (e.X - dragStart.X);
                        offY = dragOY + (e.Y - dragStart.Y);
                        ClampPan();
                        canvas.Invalidate();
                    };
                    canvas.MouseUp    += (s, e) => { dragging = false; canvas.Cursor = Cursors.Default; };
                    canvas.MouseLeave += (s, e) => { dragging = false; canvas.Cursor = Cursors.Default; };

                    canvas.MouseWheel += (s, e) => ZoomAt(e.X, e.Y, e.Delta > 0 ? 1.15 : 1.0 / 1.15);
                    canvas.MouseEnter += (s, e) => canvas.Focus();

                    canvas.KeyDown += (s, e) =>
                    {
                        int cx = canvas.Width / 2, cy = canvas.Height / 2;
                        if      (e.KeyCode == Keys.Add      || e.KeyCode == Keys.Oemplus)  ZoomAt(cx, cy, 1.25);
                        else if (e.KeyCode == Keys.Subtract || e.KeyCode == Keys.OemMinus) ZoomAt(cx, cy, 0.8);
                        else if (e.KeyCode == Keys.Home)  { CalcFit(); UpdateZoomLabel(); canvas.Invalidate(); }
                        else if (e.KeyCode == Keys.Escape) f.Close();
                    };

                    btnZoomIn .Click += (s, e) => ZoomAt(canvas.Width / 2, canvas.Height / 2, 1.25);
                    btnZoomOut.Click += (s, e) => ZoomAt(canvas.Width / 2, canvas.Height / 2, 0.8);
                    btnFit    .Click += (s, e) => { CalcFit(); UpdateZoomLabel(); canvas.Invalidate(); };
                    btn100    .Click += (s, e) =>
                    {
                        scale = 1.0;
                        offX  = (canvas.Width  - origBmp.Width)  / 2f;
                        offY  = (canvas.Height - origBmp.Height) / 2f;
                        ClampPan(); UpdateZoomLabel(); canvas.Invalidate();
                    };

                    f.Controls.Add(canvas);
                    f.FormClosed += (s, e) => { try { origBmp.Dispose(); } catch { } };

                    btnPrev.Visible = false; btnNext.Visible = false;
                    printAction = () =>
                    {
                        try
                        {
                            using var pd = new System.Drawing.Printing.PrintDocument();
                            // Margin nhỏ ~2.5mm để ảnh phủ gần hết trang A4
                            pd.DefaultPageSettings.Margins =
                                new System.Drawing.Printing.Margins(10, 10, 10, 10);
                            var dlg = new PrintDialog { Document = pd };
                            if (dlg.ShowDialog() != DialogResult.OK) return;
                            var bmp = origBmp;
                            pd.PrintPage += (ps, pe) =>
                            {
                                var g = pe.Graphics;
                                // Dùng Pixel để tránh sai số unit ở các máy in DPI khác nhau
                                g.PageUnit = GraphicsUnit.Pixel;
                                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                float dpiX = g.DpiX, dpiY = g.DpiY;
                                // Chuyển PageBounds (1/100 inch) → pixel thực của máy in
                                var pb = pe.PageBounds;
                                float pageW = pb.Width  * dpiX / 100f;
                                float pageH = pb.Height * dpiY / 100f;
                                float pageL = pb.Left   * dpiX / 100f;
                                float pageT = pb.Top    * dpiY / 100f;
                                // Scale giữ tỷ lệ, phủ toàn trang
                                float sc = Math.Min(pageW / bmp.Width, pageH / bmp.Height);
                                float dw = bmp.Width * sc, dh = bmp.Height * sc;
                                g.DrawImage(bmp,
                                    pageL + (pageW - dw) / 2f,
                                    pageT + (pageH - dh) / 2f,
                                    dw, dh);
                                pe.HasMorePages = false;
                            };
                            pd.Print();
                        }
                        catch (Exception ex) { MessageBox.Show("Lỗi in: " + ex.Message, "In file"); }
                    };
                }
                else
                {
                    BuildPreviewInfoPanel_MPR(f, filePath);
                    btnZoomIn.Visible = btnZoomOut.Visible = btnFit.Visible = btn100.Visible = lblZoom.Visible = false;
                printAction = () => { try { Process.Start(new ProcessStartInfo { FileName = filePath, Verb = "print", UseShellExecute = true }); } catch { MessageBox.Show("Không thể in file này.", "In file"); } };
                }
            }
            else if (ext == ".txt" || ext == ".csv" || ext == ".log" || ext == ".json" || ext == ".xml")
            {
                var tb = new TextBox
                {
                    Multiline = true, ReadOnly = true, Dock = DockStyle.Fill,
                    ScrollBars = ScrollBars.Both,
                    Font       = new Font("Consolas", 9.5f),
                    BackColor  = Color.FromArgb(30, 30, 30),
                    ForeColor  = Color.FromArgb(212, 212, 212),
                    BorderStyle = BorderStyle.None
                };
                var state = new PreviewStateMPR { Type = "text" };
                try
                {
                    var full = File.ReadAllText(filePath);
                    state.FullText = full.Replace("\r\n", "\n");
                    state.Lines    = state.FullText.Split('\n');
                    tb.Text = string.Join(Environment.NewLine, state.Lines.Take(state.TextPageLines));
                    lblInfo.Text = $"{Path.GetFileName(filePath)}  ({state.Lines.Length:N0} dòng)";
                }
                catch { tb.Text = "(Không thể đọc nội dung)"; }
                f.Controls.Add(tb);

                btnZoomIn.Visible = btnZoomOut.Visible = btnFit.Visible = btn100.Visible = lblZoom.Visible = false;
                btnPrev.Visible = btnNext.Visible = true;
                btnPrev.Enabled = false;
                btnNext.Enabled = state.Lines != null && state.Lines.Length > state.TextPageLines;

                void UpdatePageInfo() => lblInfo.Text =
                    $"{Path.GetFileName(filePath)}  — dòng {state.TextOffset + 1}–{Math.Min(state.Lines.Length, state.TextOffset + state.TextPageLines)} / {state.Lines.Length:N0}";

                btnPrev.Click += (s, e) =>
                {
                    state.TextOffset = Math.Max(0, state.TextOffset - state.TextPageLines);
                    tb.Text = string.Join(Environment.NewLine, state.Lines.Skip(state.TextOffset).Take(state.TextPageLines));
                    btnPrev.Enabled = state.TextOffset > 0;
                    btnNext.Enabled = state.TextOffset + state.TextPageLines < state.Lines.Length;
                    UpdatePageInfo();
                };
                btnNext.Click += (s, e) =>
                {
                    state.TextOffset = Math.Min(Math.Max(0, state.Lines.Length - state.TextPageLines), state.TextOffset + state.TextPageLines);
                    tb.Text = string.Join(Environment.NewLine, state.Lines.Skip(state.TextOffset).Take(state.TextPageLines));
                    btnPrev.Enabled = state.TextOffset > 0;
                    btnNext.Enabled = state.TextOffset + state.TextPageLines < state.Lines.Length;
                    UpdatePageInfo();
                };
                printAction = () =>
                {
                    try
                    {
                        using var pd = new System.Drawing.Printing.PrintDocument();
                        var dlg = new PrintDialog { Document = pd };
                        if (dlg.ShowDialog() != DialogResult.OK) return;
                        var lines = state.Lines ?? Array.Empty<string>();
                        int li = 0;
                        var pf = new Font("Consolas", 9);
                        pd.PrintPage += (ps, pe) =>
                        {
                            float y = pe.MarginBounds.Top, h = pf.GetHeight(pe.Graphics);
                            while (li < lines.Length)
                            {
                                if (y + h > pe.MarginBounds.Bottom) { pe.HasMorePages = true; return; }
                                pe.Graphics.DrawString(lines[li++], pf, Brushes.Black, pe.MarginBounds.Left, y);
                                y += h;
                            }
                            pe.HasMorePages = false;
                        };
                        pd.BeginPrint += (ps, pe) => li = 0;
                        pd.EndPrint   += (ps, pe) => pf.Dispose();
                        pd.Print();
                    }
                    catch (Exception ex) { MessageBox.Show(_previewOwner, "Lỗi in: " + ex.Message, "In file"); }
                };
            }
            else
            {
                BuildPreviewInfoPanel_MPR(f, filePath);
                btnZoomIn.Visible = btnZoomOut.Visible = btnFit.Visible = btn100.Visible = lblZoom.Visible = false;
                btnPrev.Visible = btnNext.Visible = false;
                printAction = () => { try { Process.Start(new ProcessStartInfo { FileName = filePath, Verb = "print", UseShellExecute = true }); } catch { MessageBox.Show(_previewOwner, "Không thể in file này.", "In file"); } };
            }

            try { StopPreviewCloseTimer_MPR(); AttachPreviewMouseHandlers_MPR(f); } catch { }
            _filePreviewForm_MPR = f;
            f.Show();
        }

        private void AttachPreviewMouseHandlers_MPR(Control ctl)
        {
            try
            {
                ctl.MouseEnter += (s, e) => StopPreviewCloseTimer_MPR();
                ctl.MouseLeave += (s, e) => StartPreviewCloseTimer_MPR();
                foreach (Control c in ctl.Controls)
                    AttachPreviewMouseHandlers_MPR(c);
            }
            catch { }
        }

        private void BuildPreviewInfoPanel_MPR(Form f, string filePath)
        {
            f.BackColor = Color.FromArgb(45, 45, 48);
            var pnl = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(45, 45, 48) };
            var pb  = new PictureBox { Left = 24, Top = 24, Width = 64, Height = 64, SizeMode = PictureBoxSizeMode.StretchImage };
            try { pb.Image = Icon.ExtractAssociatedIcon(filePath)?.ToBitmap(); } catch { }
            var lblName = new Label { Left = 24, Top = 96,  Width = 600, AutoSize = false, Height = 26, Text = Path.GetFileName(filePath), Font = new Font("Segoe UI", 11, FontStyle.Bold), ForeColor = Color.White };
            var lblSize = new Label { Left = 24, Top = 126, Width = 400, AutoSize = false, Height = 20, Text = $"Kích thước: {new FileInfo(filePath).Length:N0} bytes",   Font = new Font("Segoe UI", 9), ForeColor = Color.FromArgb(170, 170, 170) };
            var lblExt  = new Label { Left = 24, Top = 150, Width = 400, AutoSize = false, Height = 20, Text = $"Loại: {Path.GetExtension(filePath).ToUpperInvariant()}", Font = new Font("Segoe UI", 9), ForeColor = Color.FromArgb(170, 170, 170) };
            var btnOp   = new Button { Left = 24, Top = 186, Width = 120, Height = 30, Text = "Mở file", FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White, Font = new Font("Segoe UI", 9, FontStyle.Bold), Cursor = Cursors.Hand };
            btnOp.FlatAppearance.BorderSize = 0;
            btnOp.Click += (s, e) => { try { Process.Start(new ProcessStartInfo { FileName = filePath, UseShellExecute = true }); } catch { } };
            pnl.Controls.AddRange(new Control[] { pb, lblName, lblSize, lblExt, btnOp });
            f.Controls.Add(pnl);
        }

        // Shell thumbnail extraction (PDF, Office, images)
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false, EntryPoint = "SHCreateItemFromParsingName")]
        private static extern void SHCreateItemFromParsingName_MPR([MarshalAs(UnmanagedType.LPWStr)] string pszPath, IntPtr pbc, [In] ref Guid riid, out IntPtr ppv);

        [DllImport("gdi32.dll", EntryPoint = "DeleteObject")]
        private static extern bool DeleteObject_MPR(IntPtr hObject);

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b")]
        private interface IShellItemImageFactory_MPR
        {
            void GetImage([In] SIZE_MPR size, [In] SIIGBF_MPR flags, out IntPtr phbm);
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SIZE_MPR { public int cx; public int cy; public SIZE_MPR(int x, int y) { cx = x; cy = y; } }

        [Flags]
        private enum SIIGBF_MPR : uint
        {
            SIIGBF_RESIZETOFIT  = 0x00,
            SIIGBF_BIGGERSIZEOK = 0x01,
            SIIGBF_MEMORYONLY   = 0x02,
            SIIGBF_ICONONLY     = 0x04,
            SIIGBF_THUMBNAILONLY = 0x08,
            SIIGBF_INCACHEONLY  = 0x10
        }

        private Image GetShellThumbnail_MPR(string path, int width, int height)
        {
            try
            {
                Guid guid = new Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b");
                SHCreateItemFromParsingName_MPR(path, IntPtr.Zero, ref guid, out IntPtr ppv);
                if (ppv == IntPtr.Zero) return null;
                var factory = (IShellItemImageFactory_MPR)Marshal.GetObjectForIUnknown(ppv);
                factory.GetImage(new SIZE_MPR(Math.Max(1, width), Math.Max(1, height)), SIIGBF_MPR.SIIGBF_RESIZETOFIT, out IntPtr hbm);
                Marshal.ReleaseComObject(factory);
                Marshal.Release(ppv);
                if (hbm == IntPtr.Zero) return null;
                var img = Image.FromHbitmap(hbm);
                DeleteObject_MPR(hbm);
                return img;
            }
            catch { return null; }
        }

        private Button btnAddDetail, btnDeleteDetail, btnSaveDetail;

        // BẢNG: Tiến độ PO
        private DataGridView dgvPOProgress;
        private Label lblPOProgressTitle;

        private Panel panelTop, panelHeader, panelDetail;
        private ComboBox _cboFilterPO;    // Loc theo Da len PO
        private Button _btnExportDetail; // Xuat Excel chi tiet
        private Button _btnMPRStatus;   // MPR Status popup

        // Khai báo phía trên cùng của Class Form
        private System.Diagnostics.Process _excelProcess = null;

        private List<ProductModel> _selectedItem = new List<ProductModel>();


        public frmMPR(int mprId = 0)
        {
            _targetMprId = mprId;
            InitializeComponent();
            EnsureIsDeletedColumn();
            BuildUI();
            ApplyPermissions();
            LoadMPR();
            this.Resize += FrmMPR_Resize;
            this.WindowState = FormWindowState.Maximized;

            if (_targetMprId > 0)
            {
                SelectMPRById(_targetMprId);
            }
            frmAIChat.Attach(this);
        }

        private void SelectMPRById(int id)
        {
            var targetMPR = _mprList.Find(m => m.MPR_ID == id);
            if (targetMPR != null)
            {
                txtSearch.Text = targetMPR.MPR_No;
                BtnSearch_Click(null, null);
            }

            foreach (DataGridViewRow row in dgvMPR.Rows)
            {
                if (Convert.ToInt32(row.Cells["ID"].Value) == id)
                {
                    dgvMPR.ClearSelection();
                    row.Selected = true;

                    if (row.Index >= 0)
                        dgvMPR.FirstDisplayedScrollingRowIndex = row.Index;
                    break;
                }
            }
        }

        private void BuildUI()
        {
            this.Text = "Quản lý Phiếu Yêu Cầu Mua Hàng (MPR)";
            this.BackColor = Color.FromArgb(245, 245, 245);

            // ===== PANEL TOP =====
            panelTop = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(1360, 220),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(panelTop);

            panelTop.Controls.Add(new Label
            {
                Text = "DANH SÁCH PHIẾU MPR",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 10),
                Size = new Size(400, 30)
            });

            txtSearch = new TextBox
            {
                Location = new Point(10, 48),
                Size = new Size(300, 28),
                Font = new Font("Segoe UI", 10),
                PlaceholderText = "Tìm theo MPR No hoặc tên dự án..."
            };
            panelTop.Controls.Add(txtSearch);
            txtSearch.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) BtnSearch_Click(null, null); };

            btnSearch = CreateButton("🔍 Tìm", Color.FromArgb(0, 120, 212), new Point(320, 47), 85, 30);
            btnSearch.Click += BtnSearch_Click;
            panelTop.Controls.Add(btnSearch);

            // ── "➕ Tạo MPR" mới: mở popup tạo MPR ──
            var btnCreateMPR = CreateButton("➕ Tạo MPR", Color.FromArgb(40, 167, 69), new Point(415, 47), 110, 30);
            btnCreateMPR.Click += BtnCreateMPR_Click;
            panelTop.Controls.Add(btnCreateMPR);

            // ── "Update from Excel": chức năng cũ của btnNewMPR ──
            btnNewMPR = CreateButton("📥 Update from Excel", Color.FromArgb(0, 140, 120), new Point(533, 47), 155, 30);
            btnNewMPR.Click += BtnNewMPR_Click;
            panelTop.Controls.Add(btnNewMPR);

            btnDeleteMPR = CreateButton("🗑 Xóa MPR", Color.FromArgb(220, 53, 69), new Point(696, 47), 110, 30);
            btnDeleteMPR.Click += BtnDeleteMPR_Click;
            panelTop.Controls.Add(btnDeleteMPR);

            var btnPrint = new Button
            {
                Text = "🖨 In MPR",
                Location = new Point(860, 47),
                Size = new Size(110, 30),
                BackColor = Color.FromArgb(33, 115, 70),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            btnPrint.FlatAppearance.BorderSize = 0;
            btnPrint.Click += BtnPrint_Click;
            panelTop.Controls.Add(btnPrint);

            lblStatus = new Label
            {
                Location = new Point(820, 52),
                Size = new Size(500, 25),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.Gray
            };
            panelTop.Controls.Add(lblStatus);

            dgvMPR = new DataGridView
            {
                Location = new Point(10, 85),
                Size = new Size(1335, 125),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            dgvMPR.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvMPR.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvMPR.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvMPR.EnableHeadersVisualStyles = false;
            dgvMPR.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvMPR.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
            dgvMPR.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgvMPR.SelectionChanged += DgvMPR_SelectionChanged;
            panelTop.Controls.Add(dgvMPR);

            // ===== PANEL HEADER =====
            panelHeader = new Panel
            {
                Location = new Point(10, 240),
                Size = new Size(1360, 160),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(panelHeader);

            panelHeader.Controls.Add(new Label
            {
                Text = "THÔNG TIN PHIẾU MPR",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 8),
                Size = new Size(300, 25)
            });

            // === BẢNG FILE ĐÍNH KÈM ===
            int gridFilesWidth = 450;
            int filesLeft = panelHeader.Width - gridFilesWidth - 10;
            dgvFiles = new DataGridView
            {
                Location = new Point(filesLeft, 10),
                Size = new Size(gridFilesWidth, panelHeader.Height - 20),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom
            };
            dgvFiles.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(108, 117, 125);
            dgvFiles.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvFiles.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvFiles.EnableHeadersVisualStyles = false;
            dgvFiles.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvFiles.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
            dgvFiles.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgvFiles.Columns.Add("FileName", "Tệp đính kèm (MPR Link)");
            dgvFiles.Columns.Add("FullPath", "FullPath");
            dgvFiles.Columns["FullPath"].Visible = false;
            dgvFiles.CellDoubleClick += DgvFiles_CellDoubleClick;
            panelHeader.Controls.Add(dgvFiles);
            // Initialize preview handlers for MPR file list
            EnsureFilePreview_MPR();

            // === CÁC CONTROL NHẬP LIỆU BÊN TRÁI ===
            int y = 38;

            // Hàng 1
            AddLabel(panelHeader, "MPR No (*):", 10, y);
            txtMPRNo = AddTextBox(panelHeader, 90, y, 140);

            AddLabel(panelHeader, "Tên dự án:", 240, y);
            txtProjectName = AddTextBox(panelHeader, 320, y, 200);

            AddLabel(panelHeader, "Mã dự án:", 530, y);
            txtProjectCode = AddTextBox(panelHeader, 610, y, 130);

            AddLabel(panelHeader, "Trạng thái:", 750, y);
            cboStatus = new ComboBox
            {
                Location = new Point(830, y),
                Size = new Size(60, 25),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboStatus.Items.AddRange(new[] { "Mới", "Đang xử lý", "Đã duyệt", "Hoàn thành", "Hủy" });
            cboStatus.SelectedIndex = 0;
            panelHeader.Controls.Add(cboStatus);

            // Hàng 2
            y += 38;
            AddLabel(panelHeader, "Phòng ban:", 10, y);
            txtDepartment = AddTextBox(panelHeader, 90, y, 140);

            AddLabel(panelHeader, "Người YC:", 240, y);
            txtRequestor = AddTextBox(panelHeader, 320, y, 200);

            AddLabel(panelHeader, "Ngày cần:", 530, y);
            dtpRequiredDate = new DateTimePicker
            {
                Location = new Point(610, y),
                Size = new Size(130, 25),
                Font = new Font("Segoe UI", 9),
                Format = DateTimePickerFormat.Short
            };
            panelHeader.Controls.Add(dtpRequiredDate);

            AddLabel(panelHeader, "Rev:", 750, y);
            txtRev = AddTextBox(panelHeader, 790, y, 100);
            txtRev.Text = "0";

            // Hàng 3 (Buttons & Notes)
            y += 38;
            btnSaveHeader = CreateButton("💾 Lưu Header", Color.FromArgb(0, 120, 212), new Point(10, y), 130, 32);
            btnSaveHeader.Click += BtnSaveHeader_Click;
            panelHeader.Controls.Add(btnSaveHeader);

            btnClearHeader = CreateButton("🔄 Làm mới", Color.FromArgb(108, 117, 125), new Point(150, y), 110, 32);
            btnClearHeader.Click += BtnClearHeader_Click;
            panelHeader.Controls.Add(btnClearHeader);

            AddLabel(panelHeader, "Ghi chú:", 270, y + 5);
            txtNotes = AddTextBox(panelHeader, 340, y + 2, filesLeft - 340 - 15);
            txtNotes.Anchor = AnchorStyles.Top | AnchorStyles.Left;

            // ===== PANEL DETAIL =====
            panelDetail = new Panel
            {
                Location = new Point(10, panelHeader.Bottom + 10),
                Size = new Size(1360, 345),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            this.Controls.Add(panelDetail);

            panelDetail.Controls.Add(new Label
            {
                Text = "CHI TIẾT VẬT TƯ",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 8),
                Size = new Size(300, 25)
            });

            btnAddDetail = CreateButton("➕ Thêm dòng", Color.FromArgb(40, 167, 69), new Point(10, 38), 120, 30);
            btnAddDetail.Click += BtnAddDetail_Click;
            panelDetail.Controls.Add(btnAddDetail);

            btnDeleteDetail = CreateButton("🗑 Xóa dòng", Color.FromArgb(220, 53, 69), new Point(140, 38), 110, 30);
            btnDeleteDetail.Click += BtnDeleteDetail_Click;
            panelDetail.Controls.Add(btnDeleteDetail);

            btnSaveDetail = CreateButton("💾 Lưu chi tiết", Color.FromArgb(0, 120, 212), new Point(260, 38), 130, 30);
            btnSaveDetail.Click += BtnSaveDetail_Click;
            panelDetail.Controls.Add(btnSaveDetail);

            var btnCreatePO = CreateButton("🛒 Tạo PO", Color.FromArgb(255, 140, 0), new Point(400, 38), 120, 30);
            btnCreatePO.Click += BtnCreatePO_Click;
            panelDetail.Controls.Add(btnCreatePO);

            var btnCheckAll = CreateButton("🔎 Check All Items", Color.FromArgb(102, 51, 153), new Point(530, 38), 150, 30);
            btnCheckAll.Click += BtnCheckAllItems_Click;
            panelDetail.Controls.Add(btnCheckAll);

            // ── Loc theo Da len PO ──
            panelDetail.Controls.Add(new Label
            {
                Text = "Da len PO:",
                Location = new Point(695, 45),
                Size = new Size(72, 22),
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            });
            _cboFilterPO = new ComboBox
            {
                Location = new Point(770, 44),
                Size = new Size(120, 24),
                Font = new Font("Segoe UI", 8),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _cboFilterPO.Items.Add("(Tat ca)"); // placeholder, se duoc load dong
            _cboFilterPO.SelectedIndex = 0;
            _cboFilterPO.SelectedIndexChanged += (s, ev) => FilterDetailByPO();
            panelDetail.Controls.Add(_cboFilterPO);

            // ── Xuat Excel ──
            _btnExportDetail = CreateButton("📥 Xuat Excel", Color.FromArgb(0, 150, 100), new Point(898, 38), 120, 30);
            _btnExportDetail.Click += BtnExportDetail_Click;
            panelDetail.Controls.Add(_btnExportDetail);

            _btnMPRStatus = CreateButton("📊 MPR Status", Color.FromArgb(75, 0, 130), new Point(1030, 38), 130, 30);
            _btnMPRStatus.Click += BtnMPRStatus_Click;
            panelDetail.Controls.Add(_btnMPRStatus);

            dgvDetails = new DataGridView
            {
                Location = new Point(10, 75),
                Size = new Size(900, 260),
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom
            };
            dgvDetails.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvDetails.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvDetails.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvDetails.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvDetails.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvDetails.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dgvDetails.ColumnHeadersHeight = 45;
            dgvDetails.EnableHeadersVisualStyles = false;
            dgvDetails.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvDetails.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
            dgvDetails.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgvDetails.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvDetails.CellFormatting += DgvDetails_CellFormatting;
            dgvDetails.KeyDown += DgvDetails_GridKeyDown;
            // Cho phep copy nhieu o
            dgvDetails.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;

            BuildDetailColumns();
            panelDetail.Controls.Add(dgvDetails);

            lblPOProgressTitle = new Label
            {
                Text = "TỔNG HỢP TIẾN ĐỘ PO ĐÃ ĐẶT",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(255, 140, 0),
                Location = new Point(930, 48),
                Size = new Size(300, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            panelDetail.Controls.Add(lblPOProgressTitle);

            dgvPOProgress = new DataGridView
            {
                Location = new Point(930, 75),
                Size = new Size(415, 260),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom
            };
            dgvPOProgress.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 140, 0);
            dgvPOProgress.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPOProgress.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvPOProgress.EnableHeadersVisualStyles = false;
            dgvPOProgress.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(255, 248, 235);
            dgvPOProgress.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
            dgvPOProgress.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgvPOProgress.CellFormatting += DgvPOProgress_CellFormatting;
            dgvPOProgress.CellDoubleClick += DgvPOProgress_CellDoubleClick;
            panelDetail.Controls.Add(dgvPOProgress);

            Common.Common.AutoBringToFontControl(new[] { panelTop, panelHeader, panelDetail });
        }

        private void BtnPrint_Click(object? sender, EventArgs e)
        {
            if (dgvMPR.Rows.Count <= 0) return;
            int rsl = dgvMPR.CurrentRow.Index;
            int mprId = Convert.ToInt32(dgvMPR.Rows[rsl].Cells["ID"].Value.ToString().Trim());

            var mpr = _mprList.Find(x => x.MPR_ID == mprId);
            if (mpr != null && mpr.Status == "Hủy")
            {
                SafeMsg("Phiếu MPR này đã bị Hủy, không thể in báo cáo!", "Thông báo", MessageBoxIcon.Warning);
                return;
            }

            var projectName = dgvMPR.CurrentRow.Cells["Ten_Du_An"].Value.ToString().Trim();
            var mprNo = dgvMPR.CurrentRow.Cells["MPR_No"].Value.ToString().Trim();
            var mpr_detail = _service.GetActiveDetails(mprId);

            var mpr_header = new MPRHeader() { MPR_ID = mprId, MPR_No = mprNo, Project_Name = projectName };

            ExportMPRToExcel(mpr_header, mpr_detail);
        }

        // =====================================================================
        // SỰ KIỆN DOUBLE CLICK VÀO FILE TRONG BẢNG ĐÍNH KÈM
        // =====================================================================
        private void DgvFiles_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string path = dgvFiles.Rows[e.RowIndex].Cells["FullPath"].Value?.ToString();

            if (!string.IsNullOrEmpty(path) && File.Exists(path))
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = path,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                MessageBox.Show("Không thể mở file: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else if (!string.IsNullOrEmpty(path))
            {
                MessageBox.Show("File không tồn tại hoặc đã bị xóa / di chuyển khỏi thư mục!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // =====================================================================
        // HÀM QUÉT THƯ MỤC LẤY DANH SÁCH FILE CỦA DỰ ÁN ĐANG CHỌN
        // =====================================================================
        private void LoadFiles(string projectName)
        {
            dgvFiles.Rows.Clear();
            if (string.IsNullOrEmpty(projectName)) return;

            try
            {
                var projects = new ProjectService().GetAll();
                var prj = projects.Find(p =>
                    !string.IsNullOrEmpty(p.ProjectName) &&
                    p.ProjectName.Equals(projectName, StringComparison.OrdinalIgnoreCase));

                if (prj == null)
                {
                    prj = projects.Find(p =>
                        !string.IsNullOrEmpty(p.ProjectName) &&
                        p.ProjectName.IndexOf(projectName, StringComparison.OrdinalIgnoreCase) >= 0);
                }

                if (prj != null && !string.IsNullOrEmpty(prj.MPR_Link) && Directory.Exists(prj.MPR_Link))
                {
                    var files = Directory.GetFiles(prj.MPR_Link);
                    foreach (var f in files)
                    {
                        dgvFiles.Rows.Add(Path.GetFileName(f), f);
                    }
                }
                else if (prj != null && !string.IsNullOrEmpty(prj.MPR_Link))
                {
                    dgvFiles.Rows.Add("(Thư mục không tồn tại)", "");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(TopOwner, "Lỗi load files: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DgvPOProgress_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string poNo = dgvPOProgress.Rows[e.RowIndex].Cells["PO No"].Value?.ToString() ?? "";

            if (!string.IsNullOrEmpty(poNo))
            {
                var frm = new frmPO(poNo);
                frm.Show();
            }
        }

        private void BuildDetailColumns()
        {
            dgvDetails.Columns.Clear();
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Detail_ID", HeaderText = "ID", Visible = false });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Item_No",
                HeaderText = "STT",
                Width = 35,
                ReadOnly = true,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Item_Name",
                HeaderText = "Tên vật tư",
                Width = 120,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleLeft }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Description",
                HeaderText = "Mô tả",
                Width = 110,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleLeft }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Material",
                HeaderText = "Vật liệu",
                Width = 130,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Thickness_mm",
                HeaderText = "Dày\n(mm)",
                Width = 38,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Depth_mm",
                HeaderText = "-Sâu\n(mm)",
                Width = 38,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "C_Width_mm",
                HeaderText = "Rộng\n(mm)",
                Width = 38,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "D_Web_mm",
                HeaderText = "Bụng\n(mm)",
                Width = 38,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "E_Flange_mm",
                HeaderText = "Cánh\n(mm)",
                Width = 38,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "F_Length_mm",
                HeaderText = "F-Dài\n(mm)",
                Width = 65,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "UNIT",
                HeaderText = "ĐVT",
                Width = 40,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Qty",
                HeaderText = "SL",
                Width = 35,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Weight",
                HeaderText = "KG",
                Width = 35,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "MPS_Info",
                HeaderText = "MPS Info",
                Width = 60,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Usage_Location",
                HeaderText = "Vị trí dùng",
                Width = 70,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleLeft }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "REV",
                HeaderText = "REV",
                Width = 35,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Remarks",
                HeaderText = "Ghi chú",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleLeft }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "PO_No",
                HeaderText = "Đã lên PO",
                Width = 100,
                ReadOnly = true,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
        }

        private void DgvDetails_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string colName = dgvDetails.Columns[e.ColumnIndex].Name;

            if (colName == "Thickness_mm" || colName == "Depth_mm" ||
                colName == "C_Width_mm" || colName == "D_Web_mm" ||
                colName == "E_Flange_mm" || colName == "F_Length_mm" ||
                colName == "Qty" || colName == "Weight")
            {
                if (e.Value != null && decimal.TryParse(e.Value.ToString(), out decimal num) && num == 0)
                {
                    e.Value = "";
                    e.FormattingApplied = true;
                }
                e.CellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            // Cột PO_No — màu xanh bold
            if (colName == "PO_No")
            {
                string val = e.Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(val))
                {
                    e.CellStyle.ForeColor = Color.FromArgb(40, 167, 69);
                    e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                }
            }
        }

        private void DgvPOProgress_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string colName = dgvPOProgress.Columns[e.ColumnIndex].Name;

            if (colName == "% Giao")
            {
                if (decimal.TryParse(e.Value?.ToString(), out decimal pct))
                {
                    e.CellStyle.ForeColor = pct >= 100 ? Color.FromArgb(40, 167, 69) : pct >= 50 ? Color.FromArgb(255, 140, 0) : Color.FromArgb(220, 53, 69);
                    e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    e.Value = $"{pct}%";
                    e.FormattingApplied = true;
                }
            }
            if (colName == "Ngày PO")
            {
                if (e.Value != null && e.Value != DBNull.Value)
                {
                    e.Value = Convert.ToDateTime(e.Value).ToString("dd/MM/yyyy");
                    e.FormattingApplied = true;
                }
            }
        }

        private void AddLabel(Panel panel, string text, int x, int y)
        {
            panel.Controls.Add(new Label
            {
                Text = text,
                Location = new Point(x, y + 3),
                Size = new Size(80, 20),
                Font = new Font("Segoe UI", 9),
                Margin = new Padding(0)
            });
        }

        private TextBox AddTextBox(Panel panel, int x, int y, int width)
        {
            var txt = new TextBox
            {
                Location = new Point(x, y),
                Size = new Size(width, 25),
                Font = new Font("Segoe UI", 9)
            };
            panel.Controls.Add(txt);
            return txt;
        }

        private Button CreateButton(string text, Color color, Point location, int w, int h)
        {
            var btn = new Button
            {
                Text = text,
                Location = location,
                Size = new Size(w, h),
                BackColor = color,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        // ===== RESIZE =====
        private void FrmMPR_Resize(object sender, EventArgs e)
        {
            try
            {
                int w = this.ClientSize.Width - 20;
                int h = this.ClientSize.Height;

                panelTop.Width = w;
                panelHeader.Width = w;
                panelDetail.Width = w;
                panelDetail.Height = h - panelDetail.Top - 10;
                dgvMPR.Width = panelTop.Width - 20;

                int rightWidth = 420;
                dgvPOProgress.Width = rightWidth;
                dgvPOProgress.Left = panelDetail.Width - rightWidth - 10;
                dgvPOProgress.Height = panelDetail.Height - 85;

                lblPOProgressTitle.Left = dgvPOProgress.Left;

                dgvDetails.Width = dgvPOProgress.Left - 20;
                dgvDetails.Height = panelDetail.Height - 85;
            }
            catch { }
        }

        // ===== LOAD MPR =====
        private void FrmMPR_Load_Extra() => EnsureIsDeletedColumn();

        private void LoadMPR()
        {
            try
            {
                // Load TẤT CẢ MPR — bao gồm cả các Rev cũ (hiển thị dạng chỉ đọc/xám)
                _mprList = _service.GetAll();
                BindMPRGrid(_mprList);
                lblStatus.Text = $"Tổng: {_mprList.Count} phiếu MPR";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải MPR: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BindMPRGrid(List<MPRHeader> list)
        {
            // Xác định MPR nào là "mới nhất" trong mỗi nhóm cùng MPR_No
            // Dùng decimal để sort Rev an toàn kể cả khi Rev lưu dạng "0.40", "01", "1"
            // Build mapping baseKey -> max Rev and corresponding MPR_ID
            var latestIds = new HashSet<int>();
            var groups = new Dictionary<string, (int MaxRev, int MprId)>();
            foreach (var m in list)
            {
                string baseNo = !string.IsNullOrEmpty(m.MPR_No) && m.MPR_No.Contains("_Rev.")
                    ? m.MPR_No.Substring(0, m.MPR_No.IndexOf("_Rev."))
                    : (m.MPR_No ?? "");
                string key = (m.Project_Code ?? "") + "||" + baseNo;
                int revVal = 0;
                try { revVal = Convert.ToInt32(m.Rev); } catch { revVal = 0; }
                if (!groups.TryGetValue(key, out var cur) || revVal > cur.MaxRev)
                {
                    groups[key] = (revVal, m.MPR_ID);
                }
            }
            // collect latest ids
            foreach (var kv in groups) latestIds.Add(kv.Value.MprId);

            dgvMPR.DataSource = list.ConvertAll(m => new
            {
                ID = m.MPR_ID,
                MPR_No = m.MPR_No,
                Ten_Du_An = m.Project_Name,
                Ma_Du_An = m.Project_Code,
                Phong_Ban = m.Department,
                Nguoi_YC = m.Requestor,
                Ngay_Can = m.Required_Date.HasValue ? m.Required_Date.Value.ToString("dd/MM/yyyy") : "",
                Rev = m.Rev,
                Trang_Thai = m.Status,
                Ngay_Tao = m.Created_Date.HasValue ? m.Created_Date.Value.ToString("dd/MM/yyyy") : "",
                Email_Status    = m.Email_Status,
                Email_Sent_Info = m.Email_Sent_At.HasValue
                    ? $"{m.Email_Sent_By}\n{m.Email_Sent_At.Value:dd/MM/yy HH:mm}"
                    : "",
                _IsLatest = latestIds.Contains(m.MPR_ID) ? 1 : 0  // cột ẩn để format dòng
            });
            if (dgvMPR.Columns.Contains("ID"))
                dgvMPR.Columns["ID"].Visible = false;
            if (dgvMPR.Columns.Contains("_IsLatest"))
                dgvMPR.Columns["_IsLatest"].Visible = false;

            // Cố định kích thước các cột nhỏ trước (AutoSizeMode.None TRƯỚC khi set Width)
            void FixCol(string name, int w)
            {
                if (!dgvMPR.Columns.Contains(name)) return;
                var c = dgvMPR.Columns[name];
                c.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                c.Width = w;
            }
            FixCol("Rev",        45);
            FixCol("Trang_Thai", 90);
            FixCol("Ngay_Can",   95);
            FixCol("Ngay_Tao",   95);

            // Cấu hình cột Email_Status (AutoSizeMode TRƯỚC Width)
            if (dgvMPR.Columns.Contains("Email_Status"))
            {
                var esc = dgvMPR.Columns["Email_Status"];
                esc.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                esc.Width = 65;
                esc.HeaderText = "Email";
                esc.ReadOnly = true;
                esc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                esc.DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }

            // Thêm nút gửi email MPR (chỉ thêm 1 lần)
            if (!dgvMPR.Columns.Contains("col_SendEmail_MPR"))
            {
                var btnEmail = new DataGridViewButtonColumn
                {
                    Name = "col_SendEmail_MPR",
                    HeaderText = "Gửi Email",
                    Text = "📧 Gửi",
                    UseColumnTextForButtonValue = false,
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                    Width = 85,
                    FlatStyle = FlatStyle.Flat,
                    DefaultCellStyle = { BackColor = Color.FromArgb(40, 167, 69), ForeColor = Color.White, Font = new Font("Segoe UI", 8, FontStyle.Bold) },
                };
                dgvMPR.Columns.Add(btnEmail);
            }

            // Cấu hình cột Email_Sent_Info (người gửi + thời gian)
            if (dgvMPR.Columns.Contains("Email_Sent_Info"))
            {
                var c = dgvMPR.Columns["Email_Sent_Info"];
                c.AutoSizeMode  = DataGridViewAutoSizeColumnMode.None;
                c.Width         = 115;
                c.HeaderText    = "Người gửi / Thời gian";
                c.ReadOnly      = true;
                c.DefaultCellStyle.WrapMode   = DataGridViewTriState.True;
                c.DefaultCellStyle.Font       = new Font("Segoe UI", 8f);
                c.DefaultCellStyle.ForeColor  = Color.FromArgb(80, 80, 80);
                c.DefaultCellStyle.Alignment  = DataGridViewContentAlignment.MiddleLeft;
            }

            // Tô màu cột Email_Status và nút gửi
            dgvMPR.CellFormatting -= DgvMPR_EmailCellFormatting;
            dgvMPR.CellFormatting += DgvMPR_EmailCellFormatting;
            dgvMPR.CellContentClick -= DgvMPR_EmailCellContentClick;
            dgvMPR.CellContentClick += DgvMPR_EmailCellContentClick;

            // Tô màu xám (chỉ đọc) cho các bản nằm giữa (Rev > 1 && Rev < MaxRev) — vẫn hiển thị trong danh sách
            foreach (DataGridViewRow row in dgvMPR.Rows)
            {
                if (row.IsNewRow) continue;
                bool isLatest = Convert.ToInt32(row.Cells["_IsLatest"].Value) == 1;
                // Lấy Rev và base key để tra MaxRev
                string mprNo = row.Cells["MPR_No"].Value?.ToString() ?? "";
                string projCode = row.Cells["Ma_Du_An"].Value?.ToString() ?? "";
                int revVal = 0; int.TryParse(row.Cells["Rev"].Value?.ToString(), out revVal);
                string baseNo = !string.IsNullOrEmpty(mprNo) && mprNo.Contains("_Rev.") ? mprNo.Substring(0, mprNo.IndexOf("_Rev.")) : mprNo;
                string key = projCode + "||" + baseNo;
                int maxRev = groups.TryGetValue(key, out var gp) ? gp.MaxRev : revVal;

                // Lấy trạng thái để xem có bị Hủy không
                string status = row.Cells["Trang_Thai"].Value?.ToString() ?? "";
                bool isCancelled = status == "Hủy";

                bool shouldGray = (!isLatest && revVal > 1 && revVal < maxRev) || isCancelled;
                if (shouldGray)
                {
                    row.ReadOnly = true;
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        cell.Style.ForeColor = Color.FromArgb(160, 160, 160);
                        cell.Style.BackColor = Color.FromArgb(245, 245, 245);
                        cell.Style.Font = new Font("Segoe UI", 9, isCancelled ? FontStyle.Bold : FontStyle.Italic);
                    }
                }
                else
                {
                    // ensure default style for rows that should be editable
                    row.ReadOnly = false;
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        cell.Style.ForeColor = Color.Black;
                        cell.Style.BackColor = Color.White;
                        cell.Style.Font = new Font("Segoe UI", 9, FontStyle.Regular);
                    }
                }
            }
        }

        // ===== LOAD TỔNG HỢP TIẾN ĐỘ PO =====
        private void LoadPOProgress(string mprNo)
        {
            if (string.IsNullOrEmpty(mprNo))
            {
                dgvPOProgress.DataSource = null;
                return;
            }

            try
            {
                // Tính baseMprNo để match PO của tất cả Rev cùng chuỗi
                string baseMprNo = mprNo.Contains("_Rev.")
                    ? mprNo.Substring(0, mprNo.IndexOf("_Rev."))
                    : mprNo;

                string sql = @"
                    SELECT
                        h.PONo AS [PO No],
                        h.PO_Date AS [Ngày PO],
                        h.Status AS [Trạng thái],
                        CASE
                            WHEN ISNULL(SUM(d.Qty_Per_Sheet), 0) = 0 THEN 0
                            ELSE CAST(
                                ISNULL(
                                    (SELECT SUM(Qty_Import) FROM Warehouse_Import wi WHERE wi.PO_ID = h.PO_ID), 
                                    ISNULL(SUM(d.Received), 0)
                                ) * 100.0 / SUM(d.Qty_Per_Sheet) 
                            AS DECIMAL(5,1))
                        END AS [% Giao]
                    FROM PO_head h
                    LEFT JOIN PO_Detail d ON h.PO_ID = d.PO_ID
                    WHERE h.MPR_No = @baseNo OR h.MPR_No LIKE @baseNo + '_Rev.%'
                    GROUP BY h.PO_ID, h.PONo, h.PO_Date, h.Status
                    ORDER BY h.PO_Date DESC";
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@baseNo", baseMprNo);
                    var dt = new DataTable();
                    dt.Load(cmd.ExecuteReader());
                    dgvPOProgress.DataSource = dt;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(TopOwner, "Lỗi tải PO Progress: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =====================================================================
        // LẤY PO MAPPING CHO MPR — HỖ TRỢ CẢ MPR REVISE
        // Logic:
        //   Bước 1: Tìm PO liên kết trực tiếp qua MPR_Detail_ID (Detail_ID hiện tại)
        //   Bước 2: Với các dòng chưa có PO, tìm sang các phiên bản MPR khác cùng MPR_No
        //           (revise) khớp theo Item_Name + Material để lấy PO đã đặt
        //           từ phiên bản cũ — rồi điền vào dòng tương ứng của phiên bản mới
        // =====================================================================
        // Trả về dict: Detail_ID → danh sách PONo đã đặt cho từng vật tư
        // Chỉ lấy PO qua PO_Detail.MPR_Detail_ID = MPR_Details.Detail_ID
        // (đúng 1 vật tư MPR → 1 hoặc nhiều dòng PO_Detail → 1 hoặc nhiều PO)
        private Dictionary<int, string> GetPoMappingForMpr(int mprId)
        {
            var dict = new Dictionary<int, string>();
            if (mprId <= 0) return dict;
            try
            {
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();

                    // Lấy MPR_No để tính baseMprNo
                    string curMprNo = "";
                    using (var cmdNo = new SqlCommand(
                        "SELECT MPR_No FROM MPR_Header WHERE MPR_ID = @mprId", conn))
                    {
                        cmdNo.Parameters.AddWithValue("@mprId", mprId);
                        curMprNo = cmdNo.ExecuteScalar()?.ToString() ?? "";
                    }
                    string baseMprNo = curMprNo.Contains("_Rev.")
                        ? curMprNo.Substring(0, curMprNo.IndexOf("_Rev."))
                        : curMprNo;

                    // ── Query hợp nhất: lấy PO từ TẤT CẢ Rev trong cùng chuỗi MPR_No ──
                    // Chiến lược: với mỗi Detail_ID của MPR hiện tại,
                    //   1. Tìm trực tiếp qua PO_Detail.MPR_Detail_ID = Detail_ID hiện tại
                    //   2. Tìm qua các Rev khác: match Item_No (số thứ tự vật tư) trong cùng chuỗi
                    //      vì khi Revise, Item_No được giữ nguyên theo thứ tự gốc
                    string sql = @"
                        -- Bước 1: PO liên kết trực tiếp với Detail_ID của MPR hiện tại
                        SELECT cur.Detail_ID AS CurDetailId, poh.PONo
                        FROM   MPR_Details cur
                        INNER JOIN PO_Detail pod ON pod.MPR_Detail_ID = cur.Detail_ID
                        INNER JOIN PO_head   poh ON poh.PO_ID = pod.PO_ID
                        WHERE  cur.MPR_ID = @mprId
                          AND  ISNULL(poh.Status, '') <> 'Cancelled'

                        UNION

                        -- Bước 2: PO từ Rev khác — match theo Item_No (số thứ tự)
                        -- Item_No được giữ nguyên khi tạo revision, là ID chính xác nhất
                        SELECT cur.Detail_ID AS CurDetailId, poh.PONo
                        FROM   MPR_Details cur
                        INNER JOIN MPR_Details old
                               ON TRY_CAST(TRY_CAST(old.Item_No AS DECIMAL(10,2)) AS INT)
                                = TRY_CAST(TRY_CAST(cur.Item_No AS DECIMAL(10,2)) AS INT)
                              AND TRY_CAST(TRY_CAST(cur.Item_No AS DECIMAL(10,2)) AS INT) > 0
                              AND old.MPR_ID  != cur.MPR_ID
                              AND ISNULL(old.Is_Deleted, 0) = 0
                        INNER JOIN MPR_Header oldH ON oldH.MPR_ID = old.MPR_ID
                        INNER JOIN PO_Detail pod   ON pod.MPR_Detail_ID = old.Detail_ID
                        INNER JOIN PO_head   poh   ON poh.PO_ID = pod.PO_ID
                        WHERE  cur.MPR_ID = @mprId
                          AND  ISNULL(cur.Is_Deleted, 0) = 0
                          AND  (oldH.MPR_No = @baseNo OR oldH.MPR_No LIKE @baseNo + N'_Rev.%')
                          AND  ISNULL(poh.Status, '') <> 'Cancelled'

                        ORDER BY CurDetailId, PONo";

                    using (var cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@mprId", mprId);
                        cmd.Parameters.AddWithValue("@baseNo", baseMprNo);
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                if (reader["CurDetailId"] == DBNull.Value) continue;
                                int detailId = Convert.ToInt32(reader["CurDetailId"]);
                                string poNo = reader["PONo"]?.ToString()?.Trim() ?? "";
                                if (string.IsNullOrEmpty(poNo)) continue;

                                if (dict.ContainsKey(detailId))
                                {
                                    var existing = dict[detailId].Split(
                                        new[] { ", " }, StringSplitOptions.RemoveEmptyEntries);
                                    if (!Array.Exists(existing, p => p == poNo))
                                        dict[detailId] += ", " + poNo;
                                }
                                else
                                {
                                    dict[detailId] = poNo;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(TopOwner, "Lỗi lấy PO Mapping: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return dict;
        }

        private void LoadDetails(int mprId)
        {
            try
            {
                _details = _service.GetDetails(mprId);
                dgvDetails.Rows.Clear();

                var poMapping = GetPoMappingForMpr(mprId);

                // Load Is_Deleted từ DB
                var deletedIds = new System.Collections.Generic.HashSet<int>();
                try
                {
                    using var connD = Helpers.DatabaseHelper.GetConnection();
                    connD.Open();
                    var cmdD = new SqlCommand(
                        "SELECT Detail_ID FROM MPR_Details WHERE MPR_ID=@id AND Is_Deleted=1", connD);
                    cmdD.Parameters.AddWithValue("@id", mprId);
                    using var rdrD = cmdD.ExecuteReader();
                    while (rdrD.Read()) deletedIds.Add(Convert.ToInt32(rdrD["Detail_ID"]));
                }
                catch { }

                foreach (var d in _details)
                {
                    int idx = dgvDetails.Rows.Add();
                    var row = dgvDetails.Rows[idx];
                    bool isDeleted = deletedIds.Contains(d.Detail_ID);
                    row.Tag = isDeleted ? "deleted" : null; // dùng để filter

                    row.Cells["Detail_ID"].Value = d.Detail_ID;
                    row.Cells["Item_No"].Value = d.Item_No;
                    row.Cells["Item_Name"].Value = d.Item_Name;
                    row.Cells["Description"].Value = d.Description;
                    row.Cells["Material"].Value = d.Material;
                    row.Cells["Thickness_mm"].Value = d.Thickness_mm;
                    row.Cells["Depth_mm"].Value = d.Depth_mm;
                    row.Cells["C_Width_mm"].Value = d.C_Width_mm;
                    row.Cells["D_Web_mm"].Value = d.D_Web_mm;
                    row.Cells["E_Flange_mm"].Value = d.E_Flange_mm;
                    row.Cells["F_Length_mm"].Value = d.F_Length_mm;
                    row.Cells["UNIT"].Value = d.UNIT;
                    row.Cells["Qty"].Value = d.Qty_Per_Sheet;
                    row.Cells["Weight"].Value = d.Weight_kg;
                    row.Cells["MPS_Info"].Value = d.MPS_Info;
                    row.Cells["Usage_Location"].Value = d.Usage_Location;
                    // Dòng Is_Deleted: chỉ đọc, xám gạch ngang
                    if (isDeleted)
                    {
                        row.ReadOnly = true;
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            cell.Style.ForeColor = Color.FromArgb(160, 160, 160);
                            cell.Style.BackColor = Color.FromArgb(245, 245, 245);
                            cell.Style.Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Strikeout);
                        }
                    }
                    row.Cells["REV"].Value = d.REV;
                    row.Cells["Remarks"].Value = d.Remarks;
                    row.Cells["PO_No"].Value = poMapping.ContainsKey(d.Detail_ID) ? poMapping[d.Detail_ID] : "";
                }
                // Populate combobox filter voi cac gia tri thuc te tu cot PO_No
                RefreshPOFilterCombo();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải chi tiết: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ===== SỰ KIỆN =====
        private void DgvMPR_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvMPR.SelectedRows.Count == 0) return;
            var row = dgvMPR.SelectedRows[0];
            _selectedMPR_ID = Convert.ToInt32(row.Cells["ID"].Value);

            var m = _mprList.Find(x => x.MPR_ID == _selectedMPR_ID);
            if (m == null) return;
            txtMPRNo.Text = m.MPR_No;
            txtProjectName.Text = m.Project_Name;
            txtProjectCode.Text = m.Project_Code;
            txtDepartment.Text = m.Department;
            txtRequestor.Text = m.Requestor;
            txtRev.Text = m.Rev.ToString();
            txtNotes.Text = m.Notes;
            dtpRequiredDate.Value = m.Required_Date ?? DateTime.Today;

            int idx = cboStatus.Items.IndexOf(m.Status);
            cboStatus.SelectedIndex = idx >= 0 ? idx : 0;

            // Xử lý Read-only nếu trạng thái là "Hủy"
            bool isCancelled = m.Status == "Hủy";
            txtMPRNo.ReadOnly = isCancelled;
            txtProjectName.ReadOnly = isCancelled;
            txtProjectCode.ReadOnly = isCancelled;
            txtDepartment.ReadOnly = isCancelled;
            txtRequestor.ReadOnly = isCancelled;
            txtRev.ReadOnly = isCancelled;
            txtNotes.ReadOnly = isCancelled;
            cboStatus.Enabled = !isCancelled;
            dtpRequiredDate.Enabled = !isCancelled;
            dgvDetails.ReadOnly = isCancelled;

            btnSaveHeader.Enabled = !isCancelled && PermissionHelper.Check("MPR", "Lưu Header", "Lưu Header");
            btnAddDetail.Enabled = !isCancelled && PermissionHelper.Check("MPR", "Thêm dòng", "Thêm dòng");
            btnDeleteDetail.Enabled = !isCancelled && PermissionHelper.Check("MPR", "Xóa dòng", "Xóa dòng");
            btnSaveDetail.Enabled = !isCancelled && PermissionHelper.Check("MPR", "Lưu chi tiết", "Lưu chi tiết");

            LoadDetails(_selectedMPR_ID);
            LoadPOProgress(m.MPR_No);
            LoadFiles(m.Project_Name);
        }

        private void BtnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                string kw = txtSearch.Text.Trim();
                var allList = _service.GetAll();
                _mprList = string.IsNullOrEmpty(kw)
                    ? allList
                    : allList.FindAll(m =>
                        (m.MPR_No ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                        (m.Project_Name ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                        (m.Project_Code ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase));

                BindMPRGrid(_mprList);
                if (!string.IsNullOrEmpty(kw) && _mprList.Count == 0)
                { lblStatus.Text = "Không tìm thấy kết quả"; SafeMsg("Không tìm thấy kết quả phù hợp!", "Tìm kiếm", MessageBoxIcon.Information); }
                else
                    lblStatus.Text = $"Tìm thấy: {_mprList.Count} phiếu MPR";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tìm kiếm: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnNewMPR_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("MPR", "Tạo MPR", "Tạo MPR")) return;

            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "SQLTesting-Template.xlsm");
            if (!File.Exists(templatePath)) { MessageBox.Show("Không tìm thấy file!"); return; }

            Form mainForm = this.ParentForm ?? this.FindForm();

            try
            {
                if (mainForm != null) mainForm.Hide();

                // 1. Khởi chạy file
                ProcessStartInfo startInfo = new ProcessStartInfo(templatePath) { UseShellExecute = true };
                Process p = Process.Start(startInfo);

                // 2. Chờ một chút để Excel kịp load file
                System.Threading.Thread.Sleep(2000);

                // 3. Tìm tiến trình thực sự đang giữ file Excel đó
                // (Vì Excel thường gom các file vào 1 tiến trình duy nhất "EXCEL")
                Process actualProcess = null;
                string fileName = Path.GetFileNameWithoutExtension(templatePath);

                // Lặp lại việc tìm kiếm cho đến khi Excel thực sự đóng
                bool isExcelRunning = true;
                while (isExcelRunning)
                {
                    // Kiểm tra xem có tiến trình Excel nào đang mở file của mình không
                    // Lưu ý: MainWindowTitle của Excel thường có dạng "Tên_File - Excel"
                    var processes = Process.GetProcessesByName("EXCEL");
                    isExcelRunning = false;

                    foreach (var proc in processes)
                    {
                        if (proc.MainWindowTitle.Contains(fileName))
                        {
                            isExcelRunning = true;
                            break;
                        }
                    }

                    if (isExcelRunning)
                    {
                        System.Threading.Thread.Sleep(1000); // Đợi 1 giây rồi kiểm tra lại
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
            finally
            {
                // 4. Luôn đảm bảo hiện lại Form khi kết thúc vòng lặp
                if (mainForm != null)
                {
                    mainForm.Show();
                    mainForm.BringToFront();
                }

                _selectedMPR_ID = 0;
                ClearHeader();
                txtMPRNo.Focus();
            }
        }

        // ── Popup Tạo MPR mới ─────────────────────────────────────────────────
        private void BtnCreateMPR_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("MPR", "Tạo MPR", "Tạo MPR")) return;
            ShowCreateMPRPopup();
        }

        private void ShowCreateMPRPopup()
        {
            bool isUpdatingTotal = false; // Biến cờ ngăn chặn vòng lặp vô tận

            var dlg = new Form
            {
                Text = "➕ Tạo MPR mới",
                Size = new Size(1100, 680),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.Sizable,
                MinimizeBox = false,
                BackColor = Color.FromArgb(245, 245, 245)
            };

            // ── Bảng tìm dự án (trái) ──────────────────────────────────────────
            dlg.Controls.Add(new Label { Text = "🔍 Tìm dự án:", Location = new Point(10, 10), Size = new Size(90, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold) });
            var txtSearch = new TextBox { Location = new Point(102, 8), Size = new Size(180, 24), Font = new Font("Segoe UI", 9), PlaceholderText = "Mã/tên dự án..." };
            dlg.Controls.Add(txtSearch);

            var dgvProj = new DataGridView
            {
                Location = new Point(10, 38),
                Size = new Size(272, 200),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            dgvProj.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvProj.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvProj.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvProj.EnableHeadersVisualStyles = false;
            dgvProj.Columns.Add(new DataGridViewTextBoxColumn { Name = "ProjCode", HeaderText = "Mã DA", Width = 100 });
            dgvProj.Columns.Add(new DataGridViewTextBoxColumn { Name = "ProjName", HeaderText = "Tên DA", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            dlg.Controls.Add(dgvProj);

            // Load và lọc projects
            var allProjects = new List<MPR_Managerment.Models.ProjectInfo>();
            try { allProjects = new ProjectService().GetAll(); } catch { }
            var projDict = new Dictionary<int, MPR_Managerment.Models.ProjectInfo>();

            bool _suppressSelection = false;
            void FilterProjects()
            {
                _suppressSelection = true; // tắt SelectionChanged khi rebuild
                string kw = txtSearch.Text.Trim().ToLower();
                dgvProj.Rows.Clear();
                projDict.Clear();
                foreach (var p in allProjects)
                {
                    if (string.IsNullOrEmpty(kw)
                        || (p.ProjectCode ?? "").ToLower().Contains(kw)
                        || (p.ProjectName ?? "").ToLower().Contains(kw))
                    {
                        int r = dgvProj.Rows.Add();
                        dgvProj.Rows[r].Cells["ProjCode"].Value = p.ProjectCode;
                        dgvProj.Rows[r].Cells["ProjName"].Value = p.ProjectName;
                        projDict[r] = p;
                    }
                }
                dgvProj.ClearSelection();
                dgvProj.CurrentCell = null;
                _suppressSelection = false;
            }
            FilterProjects();
            // Chỉ lọc khi nhấn Enter
            // ApplyProject sẽ được gọi qua SelectionChanged (định nghĩa sau)
            txtSearch.KeyDown += (s, ev) =>
            {
                if (ev.KeyCode != Keys.Enter) return;
                ev.SuppressKeyPress = true;
                FilterProjects();
                if (dgvProj.Rows.Count == 0)
                    SafeMsg("Không tìm thấy kết quả phù hợp!", "Tìm kiếm", MessageBoxIcon.Information);
                // Nếu đúng 1 kết quả → chọn ngay, SelectionChanged tự gọi ApplyProject
                else if (dgvProj.Rows.Count == 1)
                {
                    dgvProj.ClearSelection();
                    dgvProj.Rows[0].Selected = true;
                    dgvProj.CurrentCell = dgvProj.Rows[0].Cells[0];
                }
                // Nếu nhiều kết quả → chọn row đầu tiên để user thấy ngay
                else
                {
                    dgvProj.ClearSelection();
                    dgvProj.Rows[0].Selected = true;
                    dgvProj.CurrentCell = dgvProj.Rows[0].Cells[0];
                }
            };

            // ── Thông tin dự án (chỉ đọc) ──────────────────────────────────────
            int xInfo = 292;
            dlg.Controls.Add(new Label { Text = "THÔNG TIN DỰ ÁN", Location = new Point(xInfo, 10), Size = new Size(300, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212) });

            Label LblF(string t2, int y) => new Label { Text = t2, Location = new Point(xInfo, y), Size = new Size(105, 20), Font = new Font("Segoe UI", 9) };
            TextBox TxtR(int y) => new TextBox { Location = new Point(xInfo + 108, y), Size = new Size(280, 24), Font = new Font("Segoe UI", 9), ReadOnly = true, BackColor = Color.FromArgb(240, 240, 240) };

            var tProjCode = TxtR(34); var tProjName = TxtR(62); var tDept = TxtR(90); var tReq = TxtR(118);
            dlg.Controls.Add(LblF("Mã dự án:", 34)); dlg.Controls.Add(tProjCode);
            dlg.Controls.Add(LblF("Tên dự án:", 62)); dlg.Controls.Add(tProjName);
            dlg.Controls.Add(LblF("Department:", 90)); dlg.Controls.Add(tDept);
            dlg.Controls.Add(LblF("Requestor:", 118)); dlg.Controls.Add(tReq);

            // MPR No (có thể chỉnh sửa)
            dlg.Controls.Add(new Label { Text = "MPR No (*):", Location = new Point(xInfo, 148), Size = new Size(105, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold) });
            var tMPRNo = new TextBox { Location = new Point(xInfo + 108, 146), Size = new Size(280, 24), Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            dlg.Controls.Add(tMPRNo);

            // dlg.Shown SAU khi tất cả textbox đã khai báo → tránh CS0841
            dlg.Shown += (s, ev) =>
            {
                _suppressSelection = true;
                dgvProj.ClearSelection();
                dgvProj.CurrentCell = null;
                _suppressSelection = false;
                tProjCode.Text = ""; tProjName.Text = "";
                tDept.Text = ""; tReq.Text = ""; tMPRNo.Text = "";
            };

            dlg.Controls.Add(new Label { Text = "Required Date:", Location = new Point(xInfo, 176), Size = new Size(105, 20), Font = new Font("Segoe UI", 9) });
            var dtp = new DateTimePicker { Location = new Point(xInfo + 108, 174), Size = new Size(180, 24), Font = new Font("Segoe UI", 9), Value = DateTime.Today.AddDays(30) };
            dlg.Controls.Add(dtp);

            dlg.Controls.Add(new Label { Text = "Notes:", Location = new Point(xInfo, 204), Size = new Size(105, 20), Font = new Font("Segoe UI", 9) });
            var tNotes = new TextBox { Location = new Point(xInfo + 108, 202), Size = new Size(280, 24), Font = new Font("Segoe UI", 9) };
            dlg.Controls.Add(tNotes);

            // Hàm điền thông tin dự án được chọn
            void ApplyProject(MPR_Managerment.Models.ProjectInfo prj)
            {
                if (prj == null) return;
                tProjCode.Text = prj.ProjectCode ?? "";
                tProjName.Text = prj.ProjectName ?? "";
                tDept.Text = "";
                tReq.Text = "";
                // Tính MPR No tiếp theo: MPRCode-001, MPRCode-002, ...
                try
                {
                    string mprPrefix = (prj.MPRCode ?? prj.ProjectCode ?? "").TrimEnd('-');
                    int nextSeq = 1;
                    using var conn = DatabaseHelper.GetConnection();
                    conn.Open();
                    var cmd = new SqlCommand(
                        "SELECT MAX(TRY_CAST(SUBSTRING(MPR_No, LEN(@prefix)+2, 10) AS INT)) FROM MPR_Header WHERE MPR_No LIKE @like",
                        conn);
                    cmd.Parameters.AddWithValue("@prefix", mprPrefix);
                    cmd.Parameters.AddWithValue("@like", mprPrefix + "-%");
                    var res = cmd.ExecuteScalar();
                    if (res != DBNull.Value && res != null) nextSeq = Convert.ToInt32(res) + 1;
                    tMPRNo.Text = $"{mprPrefix}-{nextSeq:D3}";
                }
                catch { tMPRNo.Text = (prj.MPRCode ?? prj.ProjectCode ?? "").TrimEnd('-') + "-001"; }
            }
            // Click chuột chọn row
            dgvProj.CellClick += (s, ev) =>
            {
                if (ev.RowIndex < 0) return;
                if (projDict.TryGetValue(ev.RowIndex, out var prj2)) ApplyProject(prj2);
            };
            // Dùng bàn phím
            dgvProj.SelectionChanged += (s, ev) =>
            {
                if (_suppressSelection) return; // đang rebuild → bỏ qua
                if (dgvProj.SelectedRows.Count == 0) return;
                int ri = dgvProj.SelectedRows[0].Index;
                if (projDict.TryGetValue(ri, out var prj3)) ApplyProject(prj3);
            };

            // ── Bảng MPR Details ────────────────────────────────────────────────
            dlg.Controls.Add(new Label { Text = "CHI TIẾT VẬT TƯ (MPR_Details)", Location = new Point(10, 248), Size = new Size(400, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212) });

            var dgvDet = new DataGridView
            {
                Location = new Point(10, 270),
                Size = new Size(dlg.ClientSize.Width - 20, 300),
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 8),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText
            };
            dgvDet.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvDet.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvDet.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 8, FontStyle.Bold);
            dgvDet.EnableHeadersVisualStyles = false;

            dgvDet.CellDoubleClick += (e, s) =>
            {
                frmSelectItem frmSelectItem = new frmSelectItem();
                frmSelectItem.ShowDialog();

                if (frmSelectItem.selectedItems.Count <= 0) return;
                int startRow = dgvDet.CurrentCell?.RowIndex ?? 0;
                foreach (var item in frmSelectItem.selectedItems)
                {
                    dgvDet.Rows.Add();
                    dgvDet.Rows[startRow].Cells["Item_No"].Value = startRow + 1;
                    dgvDet.Rows[startRow].Cells["item_name"].Value = item.Name;
                    dgvDet.Rows[startRow].Cells["Description"].Value = item.Des2;
                    dgvDet.Rows[startRow].Cells["Material"].Value = item.ProdMaterialCode;
                    dgvDet.Rows[startRow].Cells["Thickness_mm"].Value = item.A_Thickness;
                    dgvDet.Rows[startRow].Cells["Depth_mm"].Value = item.B_Depth;
                    dgvDet.Rows[startRow].Cells["C_Width_mm"].Value = item.C_Width;
                    dgvDet.Rows[startRow].Cells["D_Web_mm"].Value = item.D_Web;
                    dgvDet.Rows[startRow].Cells["E_Flange_mm"].Value = item.E_Flag;
                    dgvDet.Rows[startRow].Cells["F_Length_mm"].Value = item.F_Length;
                    dgvDet.Rows[startRow].Cells["UNIT"].Value = "";
                    dgvDet.Rows[startRow].Cells["Weight_kg"].Value = item.G_Weight;
                    dgvDet.Rows[startRow].Cells["Id"].Value = item.Id;

                    startRow++;
                }
            };

            dgvDet.CellFormatting += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                Common.Common.ColorRowsByIdGroups(dgvDet, "Id");
            };

            dgvDet.CellEndEdit += (s, e) =>
            {
                // Chỉ kiểm tra nếu cột đang sửa là "SL_Xuat"
                if (dgvDet.Columns[e.ColumnIndex].Name == "Qty_Per_Sheet")
                {
                    var row = dgvDet.Rows[e.RowIndex];

                    // Lấy giá trị nhập vào và giá trị tồn
                    decimal slNhap = 0;

                    // Ép kiểu an toàn (sử dụng decimal.TryParse để tránh lỗi nhập chữ)
                    decimal.TryParse(row.Cells["Qty_Per_Sheet"].Value?.ToString() ?? "0", out slNhap);

                    if (slNhap == 0)
                    {
                        // Gán lại giá trị Xuất bằng giá trị Tồn
                        row.Cells["Qty_Per_Sheet"].Value = 0;
                    }
                }
                // Chỉ kiểm tra nếu cột đang sửa là "SL_Xuat"
                if (dgvDet.Columns[e.ColumnIndex].Name == "Weight_kg")
                {
                    var row = dgvDet.Rows[e.RowIndex];

                    // Lấy giá trị nhập vào và giá trị tồn
                    decimal slNhap = 0;

                    // Ép kiểu an toàn (sử dụng decimal.TryParse để tránh lỗi nhập chữ)
                    decimal.TryParse(row.Cells["Weight_kg"].Value?.ToString() ?? "0", out slNhap);

                    if (slNhap == 0)
                    {
                        // Gán lại giá trị Xuất bằng giá trị Tồn
                        row.Cells["Weight_kg"].Value = 0;
                    }
                }
            };

            // Các cột theo MPR_Details
            void AddCol(string name, string hdr, int w, bool ro = false)
            {
                dgvDet.Columns.Add(new DataGridViewTextBoxColumn
                {
                    Name = name,
                    HeaderText = hdr,
                    Width = w,
                    ReadOnly = ro,
                    DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleLeft }
                });
            }
            AddCol("Item_No", "STT", 40);
            AddCol("item_name", "Tên hàng", 140);
            AddCol("Description", "Mô tả", 120);
            AddCol("Material", "Vật liệu", 80);
            AddCol("Thickness_mm", "T(mm)", 55);
            AddCol("Depth_mm", "D(mm)", 55);
            AddCol("C_Width_mm", "W(mm)", 55);
            AddCol("D_Web_mm", "Web(mm)", 60);
            AddCol("E_Flange_mm", "Flange(mm)", 70);
            AddCol("F_Length_mm", "L(mm)", 60);
            AddCol("Usage_Location", "Vị trí", 90);
            AddCol("MPS_Info", "MPS", 70);
            AddCol("REV", "REV", 45);
            AddCol("DWG_BOQ_Receive_Date", "DWG Date", 80);
            AddCol("Issue_Date", "Issue Date", 80);
            AddCol("UNIT", "Đơn vị", 55);
            AddCol("Qty_Per_Sheet", "SL", 45);
            AddCol("Weight_kg", "KG", 55);
            AddCol("Remarks", "Ghi chú", 100);
            AddCol("Id", "Item Id", 50);
            dgvDet.Columns["Id"].Visible = false;
            dlg.Controls.Add(dgvDet);

            // --- THÊM DÒNG TỔNG BAN ĐẦU ---
            void AddTotalRow()
            {
                int idx = dgvDet.Rows.Add();
                DataGridViewRow totalRow = dgvDet.Rows[idx];
                totalRow.Tag = "TOTAL"; // Đánh dấu đây là dòng tổng
                totalRow.ReadOnly = true;
                totalRow.DefaultCellStyle.BackColor = Color.LightGray;
                totalRow.DefaultCellStyle.Font = new Font(dgvDet.Font, FontStyle.Bold);
                totalRow.Cells["item_name"].Value = "🔥 TỔNG CỘNG:";
                UpdateGridTotal();
            }

            // --- HÀM TÍNH TỔNG ---
            void UpdateGridTotal()
            {
                decimal totalQty = 0;
                decimal totalWeight = 0;

                foreach (DataGridViewRow row in dgvDet.Rows)
                {
                    // Chỉ tính những dòng thường, bỏ qua dòng TOTAL và NewRow
                    if (row.Tag?.ToString() != "TOTAL" && !row.IsNewRow)
                    {
                        if (decimal.TryParse(row.Cells["Qty_Per_Sheet"].Value?.ToString(), out decimal q)) totalQty += q;
                        if (decimal.TryParse(row.Cells["Weight_kg"].Value?.ToString(), out decimal w)) totalWeight += w;
                    }
                }

                // Hiển thị lên dòng TOTAL
                foreach (DataGridViewRow row in dgvDet.Rows)
                {
                    if (row.Tag?.ToString() == "TOTAL")
                    {
                        row.Cells["Qty_Per_Sheet"].Value = totalQty.ToString("N0");
                        row.Cells["Weight_kg"].Value = totalWeight.ToString("N2");
                        break;
                    }
                }
            }

            // Sự kiện này kích hoạt khi thêm dòng mới hoặc xóa dòng
            dgvDet.RowsAdded += (s, e) => RebuildTotalRow();
            dgvDet.RowsRemoved += (s, e) => RebuildTotalRow();

            // Gọi lần đầu để hiển thị dòng Total ngay khi mở popup
            RebuildTotalRow();

            // Đăng ký sự kiện ngay sau khi khởi tạo dgvDet
            dgvDet.CellValueChanged += (s, e) => {
                if (e.RowIndex >= 0 && (dgvDet.Columns[e.ColumnIndex].Name == "Qty_Per_Sheet" ||
                                       dgvDet.Columns[e.ColumnIndex].Name == "Weight_kg"))
                {
                    RebuildTotalRow();
                }
            };

            void RebuildTotalRow()
            {
                if (isUpdatingTotal) return;
                isUpdatingTotal = true;

                decimal totalQty = 0;
                decimal totalWeight = 0;

                dgvDet.SuspendLayout();
                try
                {
                    // 1. Tìm và xóa dòng TOTAL cũ, đồng thời tính tổng
                    for (int i = dgvDet.Rows.Count - 1; i >= 0; i--)
                    {
                        var row = dgvDet.Rows[i];
                        if (row.IsNewRow) continue; // Bỏ qua dòng trống mặc định

                        if (row.Tag?.ToString() == "TOTAL")
                        {
                            dgvDet.Rows.RemoveAt(i);
                        }
                        else
                        {
                            if (decimal.TryParse(row.Cells["Qty_Per_Sheet"].Value?.ToString(), out decimal q)) totalQty += q;
                            if (decimal.TryParse(row.Cells["Weight_kg"].Value?.ToString(), out decimal w)) totalWeight += w;
                        }
                    }

                    // 2. Tạm thời tắt AllowUserToAddRows để dòng TOTAL là dòng cuối cùng thực sự
                    dgvDet.AllowUserToAddRows = false;

                    // 3. Thêm dòng TOTAL vào cuối cùng
                    int totalIdx = dgvDet.Rows.Add();
                    DataGridViewRow totalRow = dgvDet.Rows[totalIdx];

                    totalRow.Tag = "TOTAL";
                    totalRow.ReadOnly = true;
                    totalRow.DefaultCellStyle.BackColor = Color.FromArgb(230, 230, 230);
                    totalRow.DefaultCellStyle.Font = new Font(dgvDet.Font, FontStyle.Bold);
                    totalRow.DefaultCellStyle.ForeColor = Color.FromArgb(0, 120, 212);

                    totalRow.Cells["item_name"].Value = "🔥 TỔNG CỘNG:";
                    totalRow.Cells["Qty_Per_Sheet"].Value = totalQty.ToString("N0");
                    totalRow.Cells["Weight_kg"].Value = totalWeight.ToString("N2");
                }
                finally
                {
                    isUpdatingTotal = false;
                    dgvDet.ResumeLayout();
                }
            }

            // Gọi hàm thêm dòng tổng ngay sau khi khởi tạo
            AddTotalRow();

            // Paste từ Excel — 1 lần, xử lý đúng new row placeholder và dòng TOTAL
            dgvDet.KeyDown += (s, ev) =>
            {
                if (ev.Control && ev.KeyCode == Keys.V)
                {
                    ev.Handled = true;
                    ev.SuppressKeyPress = true;
                    string clip = Clipboard.GetText();
                    if (string.IsNullOrEmpty(clip)) return;
                    dgvDet.EndEdit();

                    // Dùng None để không bỏ dòng cuối có ô trống.
                    // Chỉ loại bỏ đúng 1 phần tử rỗng ở CUỐI (do Excel tự thêm \r\n).
                    string[] rows2 = clip.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                    if (rows2.Length > 0 && string.IsNullOrEmpty(rows2[rows2.Length - 1]))
                        rows2 = rows2.Take(rows2.Length - 1).ToArray();
                    if (rows2.Length == 0) return;

                    // Tìm index của dòng TOTAL (luôn là dòng cuối cùng thực sự)
                    int totalRowIndex = -1;
                    for (int i = dgvDet.Rows.Count - 1; i >= 0; i--)
                    {
                        if (!dgvDet.Rows[i].IsNewRow && dgvDet.Rows[i].Tag?.ToString() == "TOTAL")
                        {
                            totalRowIndex = i;
                            break;
                        }
                    }

                    // Số dòng data thực = tất cả rows trừ TOTAL và new-row placeholder
                    int dataRowCount = dgvDet.Rows.Count
                        - (dgvDet.AllowUserToAddRows ? 1 : 0)
                        - (totalRowIndex >= 0 ? 1 : 0);

                    // Xác định ô bắt đầu
                    int startCol = dgvDet.CurrentCell?.ColumnIndex ?? 0;
                    int curRow = dgvDet.CurrentCell?.RowIndex ?? 0;

                    // Nếu con trỏ đang ở dòng TOTAL hoặc vượt quá data → paste từ cuối data
                    int startRow = (curRow >= dataRowCount || (totalRowIndex >= 0 && curRow == totalRowIndex))
                        ? dataRowCount
                        : curRow;

                    // Số dòng cần thêm mới (insert TRƯỚC dòng TOTAL)
                    int rowsNeeded = (startRow + rows2.Length) - dataRowCount;
                    dgvDet.SuspendLayout();
                    for (int ri = 0; ri < rowsNeeded; ri++)
                    {
                        // Insert ngay trước dòng TOTAL (hoặc ở cuối nếu không có TOTAL)
                        int insertAt = (totalRowIndex >= 0) ? totalRowIndex : dgvDet.Rows.Count;
                        dgvDet.Rows.Insert(insertAt, 1);
                        // Sau mỗi lần insert, totalRowIndex bị đẩy xuống 1
                        if (totalRowIndex >= 0) totalRowIndex++;
                    }

                    // Ghi data — bỏ qua nếu targetRow là dòng TOTAL
                    for (int ri2 = 0; ri2 < rows2.Length; ri2++)
                    {
                        string[] cells = rows2[ri2].Split('\t');
                        int targetRow = startRow + ri2;
                        if (targetRow >= dgvDet.Rows.Count) break;
                        if (dgvDet.Rows[targetRow].Tag?.ToString() == "TOTAL") break; // không ghi đè TOTAL
                        for (int c = 0; c < cells.Length && startCol + c < dgvDet.Columns.Count; c++)
                            if (!dgvDet.Columns[startCol + c].ReadOnly)
                                dgvDet.Rows[targetRow].Cells[startCol + c].Value = cells[c].Trim();
                    }
                    dgvDet.ResumeLayout();
                    dgvDet.Refresh();
                    RebuildTotalRow(); // Cập nhật lại tổng sau khi paste
                }
            };

            // ── Buttons ─────────────────────────────────────────────────────────
            var lblErr2 = new Label { Location = new Point(10, dlg.ClientSize.Height - 78), Size = new Size(500, 20), ForeColor = Color.Red, Font = new Font("Segoe UI", 8), Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            dlg.Controls.Add(lblErr2);

            var btnAddRow = new Button
            {
                Text = "+ Thêm dòng",
                Location = new Point(10, dlg.ClientSize.Height - 50),
                Size = new Size(110, 32),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            btnAddRow.FlatAppearance.BorderSize = 0;

            var btnDeleteRow = new Button
            {
                Text = "🗑 Xóa dòng",
                Location = new Point(130, dlg.ClientSize.Height - 50),
                Size = new Size(110, 34),
                BackColor = Color.FromArgb(255, 67, 60),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            btnDeleteDetail.FlatAppearance.BorderSize = 0;

            var btnCreate = new Button
            {
                Text = "✅ Tạo MPR",
                Location = new Point(dlg.ClientSize.Width - 460, dlg.ClientSize.Height - 50),
                Size = new Size(110, 34),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            btnCreate.FlatAppearance.BorderSize = 0;

            var btnRevise = new Button
            {
                Text = "🔄 Revise MPR",
                Location = new Point(dlg.ClientSize.Width - 342, dlg.ClientSize.Height - 50),
                Size = new Size(125, 34),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            btnRevise.FlatAppearance.BorderSize = 0;

            var btnAdmin = new Button
            {
                Text = "🔑 For Admin",
                Location = new Point(dlg.ClientSize.Width - 210, dlg.ClientSize.Height - 50),
                Size = new Size(110, 34),
                BackColor = Color.FromArgb(128, 0, 128),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                Visible = AppSession.IsAdmin
            };
            btnAdmin.FlatAppearance.BorderSize = 0;

            var btnClose2 = new Button
            {
                Text = "Đóng",
                Location = new Point(dlg.ClientSize.Width - 92, dlg.ClientSize.Height - 50),
                Size = new Size(80, 34),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                DialogResult = DialogResult.Cancel
            };
            btnClose2.FlatAppearance.BorderSize = 0;
            dlg.Controls.AddRange(new Control[] { btnAddRow, btnDeleteRow, btnCreate, btnRevise, btnAdmin, btnClose2 });
            dlg.CancelButton = btnClose2;

            // Resize handler
            dlg.Resize += (s, ev) =>
            {
                dgvDet.Width = dlg.ClientSize.Width - 20;
                dgvDet.Height = dlg.ClientSize.Height - 270 - 62;
                lblErr2.Top = dlg.ClientSize.Height - 78;
            };

            // -- THÊM DÒNG
            btnAddRow.Click += (s, e) =>
            {
                int rowIndex = dgvDet.Rows.Add();
                dgvDet.CurrentCell = dgvDet.Rows[rowIndex].Cells[0];
                dgvDet.BeginEdit(true);
            };

            // -- XÓA DÒNG
            btnDeleteRow.Click += (s, e) =>
            {
                if (dgvDet.SelectedRows.Count == 0)
                {
                    SafeMsg("Vui lòng chọn ít nhất một dòng trong danh sách vật tư để xóa!", "Thông báo", MessageBoxIcon.Warning);
                    return;
                }
                string msg = dgvDet.SelectedRows.Count == 1 ?
                "Bạn có chắc chắn muốn xóa dòng này?" : $"Bạn có chắc chắn muốn xóa {dgvDet.SelectedRows.Count} dòng đã chọn?";
                if (SafeConfirm(msg, "Xác nhận xóa") == DialogResult.Yes)
                {
                    try
                    {
                        var rowsToDelete = new List<DataGridViewRow>();
                        foreach (DataGridViewRow row in dgvDet.SelectedRows)
                        {
                            // KHÔNG cho phép xóa dòng TOTAL
                            if (!row.IsNewRow && row.Tag?.ToString() != "TOTAL")
                                rowsToDelete.Add(row);
                        }

                        foreach (var row in rowsToDelete) dgvDet.Rows.Remove(row);
                        int itemNo = 1;
                        foreach (DataGridViewRow row in dgvDet.Rows)
                            if (!row.IsNewRow && row.Tag?.ToString() != "TOTAL") row.Cells["Item_No"].Value = itemNo++;
                        Common.Common.AutoAdjustColumnWidths(dgvDet);
                    }
                    catch (Exception ex)
                    {
                        dlg.BringToFront(); dlg.Activate();
                        MessageBox.Show("Lỗi khi xóa: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            };

             // ── Tạo MPR ─────────────────────────────────────────────────────────
             btnCreate.Click += (s, ev) =>
             {
                 if (string.IsNullOrWhiteSpace(tMPRNo.Text)) { lblErr2.Text = "⚠ Nhập MPR No!"; return; }
                 if (string.IsNullOrWhiteSpace(tProjCode.Text)) { lblErr2.Text = "⚠ Chọn dự án!"; return; }

                 // --- BẮT ĐẦU KIỂM TRA DỮ LIỆU GRID ---
                 dgvDet.EndEdit(); // Kết thúc biên tập để đảm bảo dữ liệu được ghi nhận

                 // Kiểm tra xem bảng Chi tiết vật tư có dữ liệu không
                 bool hasValidDetail = false;
                 foreach (DataGridViewRow row in dgvDet.Rows)
                 {
                     // Bỏ qua dòng placeholder mới và dòng TOTAL
                     if (row.IsNewRow || row.Tag?.ToString() == "TOTAL") continue;
                     if (!IsRowEmpty(row))
                     {
                         hasValidDetail = true;
                         break;
                     }
                 }

                 if (!hasValidDetail)
                 {
                     dlg.BringToFront(); dlg.Activate();
                     MessageBox.Show("⚠ Vui lòng nhập chi tiết vật tư!\n\nBảng 'Chi tiết vật tư' không được để trống. Hãy thêm ít nhất một dòng vật tư trước khi tạo MPR.", "Yêu cầu nhập liệu", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                     return;
                 }

                foreach (DataGridViewRow row in dgvDet.Rows)
                {
                    // Bỏ qua dòng placeholder mới và dòng TOTAL
                    if (row.IsNewRow || row.Tag?.ToString() == "TOTAL") continue;
                    if (IsRowEmpty(row)) continue; // Bỏ qua dòng trống
                    // Lấy giá trị từ 2 cột Qty và Weight
                    object valQty = row.Cells["Qty_Per_Sheet"].Value;
                    string strSoLuong = valQty != null ? valQty.ToString() : "";

                    var valWeight = row.Cells["Weight_kg"].Value;

                    if (string.IsNullOrWhiteSpace(strSoLuong))
                    {
                        MessageBox.Show("Cảnh báo: Ô 'Số lượng' đang để trống hoặc chứa khoảng trắng!\nVui lòng nhập giá trị số vào ô này.",
                                        "Yêu cầu nhập liệu",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Warning);

                        // Đưa con trỏ chuột và kích hoạt chế độ chỉnh sửa ngay tại ô Số lượng đó cho người dùng tiện nhập liệu
                        dgvDet.CurrentCell = row.Cells["Qty_Per_Sheet"]; // Focus vào ô lỗi
                        dgvDet.BeginEdit(true); // Mở chế độ nhập liệu ngay
                        return;
                    }
                    // Trường hợp 2: Có dữ liệu, tiến hành dùng decimal.TryParse để kiểm tra xem có parse sang số được không
                    decimal parsedQty;
                    bool isNumeric = decimal.TryParse(strSoLuong.Trim(), out parsedQty);

                    if (isNumeric)
                    {
                        // GIAI ĐOẠN PARSE THÀNH CÔNG: Lúc này biến 'parsedQty' đã mang giá trị kiểu decimal
                        // Bạn có thể đem biến 'parsedQty' này đi tính toán hoặc truyền vào Stored Procedure
                        System.Diagnostics.Debug.WriteLine($"Parse thành công số lượng: {parsedQty}");

                        // Ví dụ một thông báo nhỏ dưới thanh trạng thái hoặc thông báo kiểm tra (tùy nhu cầu UX của bạn)
                        // MessageBox.Show($"Số lượng hợp lệ: {parsedQty}", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        // Trường hợp ô nhập chữ hoặc ký tự đặc biệt không thể chuyển thành số (Ví dụ: "ABC", "12a3")
                        MessageBox.Show($"Giá trị '{strSoLuong}' tại ô 'Số lượng' không phải là số hợp lệ!\nVui lòng nhập lại.",
                                        "Sai định dạng số",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Error);

                        dgvDet.CurrentCell = row.Cells["SoLuong"];
                        dgvDet.BeginEdit(true);
                    }
                }

                // Kiểm tra MPR_No đã tồn tại chưa
                try
                {
                    using var conn = DatabaseHelper.GetConnection();
                    conn.Open();
                    var chk = new SqlCommand("SELECT COUNT(1) FROM MPR_Header WHERE MPR_No=@no", conn);
                    chk.Parameters.AddWithValue("@no", tMPRNo.Text.Trim());
                    if (Convert.ToInt32(chk.ExecuteScalar()) > 0)
                    { lblErr2.Text = "❌ MPR No đã tồn tại! Vui lòng đổi số MPR."; return; }
                }
                catch (Exception ex) { lblErr2.Text = "❌ " + ex.Message; return; }

                try
                {
                    var header = new MPRHeader
                    {
                        MPR_ID = 0,
                        MPR_No = tMPRNo.Text.Trim(),
                        Project_Name = tProjName.Text.Trim(),
                        Project_Code = tProjCode.Text.Trim(),
                        Department = AppSession.CurrentUser.Department.Split("||")[0] ?? txtDepartment.Text.Trim(),
                        Requestor = AppSession.CurrentUser.Full_Name ?? txtRequestor.Text.Trim(),
                        Required_Date = dtp.Value,
                        Rev = 0,
                        Status = "Mới",
                        Notes = tNotes.Text.Trim(),
                        Created_By = AppSession.CurrentUser.Full_Name
                    };
                    int newId = _service.InsertHeader(header, _currentUser);

                    // Lưu details
                    // Insert details dùng MPRService.InsertDetail
                    int stt = 1;
                    foreach (DataGridViewRow row2 in dgvDet.Rows)
                    {
                        if (row2.IsNewRow || row2.Tag?.ToString() == "TOTAL") continue; // THÊM ĐIỀU KIỆN NÀY

                        string nm = row2.Cells["item_name"].Value?.ToString() ?? "";
                        if (string.IsNullOrWhiteSpace(nm)) continue;
                        _service.InsertDetail(new MPRDetail
                        {
                            MPR_ID = newId,
                            Item_No = stt++,
                            Item_Name = nm,
                            Description = row2.Cells["Description"].Value?.ToString() ?? "",
                            Material = row2.Cells["Material"].Value?.ToString() ?? "",
                            Thickness_mm = decimal.TryParse(row2.Cells["Thickness_mm"].Value?.ToString(), out decimal th) ? th : 0,
                            Depth_mm = decimal.TryParse(row2.Cells["Depth_mm"].Value?.ToString(), out decimal dp) ? dp : 0,
                            C_Width_mm = decimal.TryParse(row2.Cells["C_Width_mm"].Value?.ToString(), out decimal cw) ? cw : 0,
                            D_Web_mm = decimal.TryParse(row2.Cells["D_Web_mm"].Value?.ToString(), out decimal dw) ? dw : 0,
                            E_Flange_mm = decimal.TryParse(row2.Cells["E_Flange_mm"].Value?.ToString(), out decimal ef) ? ef : 0,
                            F_Length_mm = decimal.TryParse(row2.Cells["F_Length_mm"].Value?.ToString(), out decimal fl) ? fl : 0,
                            UNIT = row2.Cells["UNIT"].Value?.ToString() ?? "",
                            Qty_Per_Sheet = decimal.TryParse(row2.Cells["Qty_Per_Sheet"].Value?.ToString(), out decimal qs) ? qs : 0,
                            Weight_kg = decimal.TryParse(row2.Cells["Weight_kg"].Value?.ToString(), out decimal wk) ? wk : 0,
                            MPS_Info = row2.Cells["MPS_Info"].Value?.ToString() ?? "",
                            Usage_Location = row2.Cells["Usage_Location"].Value?.ToString() ?? "",
                            REV = "0",
                            Remarks = row2.Cells["Remarks"].Value?.ToString() ?? ""
                        }, _currentUser);
                    }

                    MessageBox.Show(TopOwner, $"✅ Tạo MPR '{header.MPR_No}' thành công ({stt - 1} vật tư)!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    dlg.DialogResult = DialogResult.OK;
                    dlg.Close();
                    LoadMPR();
                    _selectedMPR_ID = newId;
                    foreach (DataGridViewRow row2 in dgvMPR.Rows)
                        if (Convert.ToInt32(row2.Cells["MPR_ID"]?.Value ?? 0) == newId)
                        { row2.Selected = true; dgvMPR.CurrentCell = row2.Cells[1]; break; }
                }
                catch (Exception ex) { lblErr2.Text = "❌ " + ex.Message; }
            };

            // ── Revise MPR ───────────────────────────────────────────────────────
            btnRevise.Click += (s, ev) =>
            {
                if (string.IsNullOrWhiteSpace(tProjCode.Text)) { lblErr2.Text = "⚠ Chọn dự án trước!"; return; }
                ShowReviseMPRPopup(tProjCode.Text, tProjName.Text, false);
            };

            // ── For Admin ────────────────────────────────────────────────────────
            btnAdmin.Click += (s, ev) =>
            {
                if (string.IsNullOrWhiteSpace(tProjCode.Text)) { lblErr2.Text = "⚠ Chọn dự án trước!"; return; }
                ShowReviseMPRPopup(tProjCode.Text, tProjName.Text, true);
            };

            dlg.Owner = TopOwner;
            dlg.ShowDialog();
        }

        private bool IsRowEmpty(DataGridViewRow row)
        {
            // Ignore the new row placeholder
            if (row.IsNewRow) return true;

            // Check if all cells are null or empty
            return row.Cells.Cast<DataGridViewCell>()
                .All(cell => cell.Value == null || string.IsNullOrWhiteSpace(cell.Value.ToString()));
        }

        public void ExportMPRToExcel(MPRHeader header, List<MPRDetail> details)
        {
            // 1. Kiểm tra template
            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "mpr_template.xlsx");
            if (!File.Exists(templatePath))
            {
                MessageBox.Show(TopOwner, "Không tìm thấy file template mpr_template.xlsx!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 2. Chọn nơi lưu file
            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"MPR_{header.MPR_No}_{DateTime.Now:ddMMyyyy_HHmm}.xlsx",
                Title = "Lưu file MPR"
            };

            if (saveDialog.ShowDialog(TopOwner) != DialogResult.OK) return;

            try
            {
                // Copy từ template
                File.Copy(templatePath, saveDialog.FileName, true);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(saveDialog.FileName)))
                {
                    var ws = package.Workbook.Worksheets[0]; // Giả định dữ liệu ở Sheet đầu tiên

                    // --- PHẦN 1: THAY THẾ HEADER <<...>> ---
                    // Dựa vào snippet của bạn, các giá trị này nằm ở các dòng đầu (dòng 1 và 2)
                    ReplaceText(ws, "<<MPR-NO>>", header.MPR_No);
                    ReplaceText(ws, "<<PROJECT-NAME>>", header.Project_Name);
                    ReplaceText(ws, "<<WO-NO>>", ""); // Hoặc WO_No của bạn
                    ReplaceText(ws, "<<REV>>", header.Rev.ToString() ?? "0");
                    ReplaceText(ws, "<<DATE>>", DateTime.Now.ToString("dd/MM/yyyy"));

                    // --- PHẦN 2: ĐIỀN DỮ LIỆU CHI TIẾT ---
                    int startRow = 4; // Dòng bắt đầu điền item đầu tiên (Dưới tiêu đề cột)
                    int detailCount = details.Count;

                    // Nếu có nhiều hơn 1 dòng, chèn thêm dòng để giữ định dạng và không đè phần Footer
                    if (detailCount > 1)
                    {
                        // Chèn thêm (n-1) dòng, copy format từ dòng startRow
                        ws.InsertRow(startRow + 1, detailCount - 1, startRow);
                    }

                    decimal totalQty = 0;
                    decimal totalWeight = 0;

                    for (int i = 0; i < detailCount; i++)
                    {
                        var d = details[i];
                        int currentRow = startRow + i;

                        // Cột dựa theo cấu trúc: NO(A), DESCRIPTION(B,C), MATERIAL(D), A(E), B(F)...
                        ws.Cells[currentRow, 1].Value = i + 1; // NO
                        ws.Cells[currentRow, 2].Value = d.Item_Name; // Cột B
                        ws.Cells[currentRow, 3].Value = d.Description; // Cột C (nếu có)
                        ws.Cells[currentRow, 4].Value = d.Material; // MATERIAL
                        ws.Cells[currentRow, 5].Value = d.Thickness_mm; // A
                        ws.Cells[currentRow, 6].Value = d.Depth_mm; // B
                        ws.Cells[currentRow, 7].Value = d.C_Width_mm; // C
                        ws.Cells[currentRow, 8].Value = d.D_Web_mm; // D
                        ws.Cells[currentRow, 9].Value = d.E_Flange_mm; // E
                        ws.Cells[currentRow, 10].Value = d.F_Length_mm; // F

                        ws.Cells[currentRow, 11].Value = d.Usage_Location; // F
                        ws.Cells[currentRow, 12].Value = d.MPS_Info; // F
                        ws.Cells[currentRow, 13].Value = d.REV; // F
                        ws.Cells[currentRow, 14].Value = d.DWG_BOQ_Receive_Date; // F
                        ws.Cells[currentRow, 15].Value = d.Issue_Date; // F

                        // Các cột khác...
                        ws.Cells[currentRow, 16].Value = d.UNIT; // UNIT (Cột P)
                        ws.Cells[currentRow, 17].Value = d.Qty_Per_Sheet; // Q'ty / Sh't (Cột Q)
                        ws.Cells[currentRow, 18].Value = d.Weight_kg; // Weight(kg) (Cột R)

                        totalQty += d.Qty_Per_Sheet > 0 ? d.Qty_Per_Sheet : 0;
                        totalWeight += d.Weight_kg > 0 ? d.Weight_kg : 0;
                    }

                    // --- PHẦN 3: TÍNH TỔNG (GRAND-TOTAL) ---
                    // Tìm dòng có chữ "GRAND-TOTAL" để điền tổng
                    int grandTotalRow = -1;
                    // Tìm trong phạm vi 10 dòng sau khi kết thúc dữ liệu
                    for (int r = startRow + detailCount; r < startRow + detailCount + 10; r++)
                    {
                        if (ws.Cells[r, 2].Text.Contains("GRAND-TOTAL"))
                        {
                            grandTotalRow = r;
                            break;
                        }
                    }

                    if (grandTotalRow != -1)
                    {
                        ws.Cells[grandTotalRow, 17].Value = totalQty; // Cột Q
                        ws.Cells[grandTotalRow, 18].Value = totalWeight; // Cột R
                        ws.Cells[grandTotalRow, 17, grandTotalRow, 18].Style.Font.Bold = true;
                    }

                    package.Save();
                }

                if (MessageBox.Show(TopOwner, $"✅ Xuất phiếu MPR thành công!\nBạn có muốn mở file không?", "Thành công",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(saveDialog.FileName) { UseShellExecute = true });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(TopOwner, "Lỗi xuất Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Hàm bổ trợ thay thế text trong Header
        private void ReplaceText(ExcelWorksheet ws, string findText, string replaceValue)
        {
            // Quét vùng Header (ví dụ dòng 1 đến 3)
            for (int r = 1; r <= 5; r++)
            {
                for (int c = 1; c <= 20; c++)
                {
                    if (ws.Cells[r, c].Text.Contains(findText))
                    {
                        ws.Cells[r, c].Value = ws.Cells[r, c].Text.Replace(findText, replaceValue ?? "");
                    }
                }
            }
        }

        // ── Revise MPR Popup ─────────────────────────────────────────────────────
        private void ShowReviseMPRPopup(string projCode, string projName, bool isAdmin)
        {
            var dlg = new Form
            {
                Text = isAdmin ? "🔑 Admin Edit MPR" : "🔄 Revise MPR",
                Size = new Size(1100, 700),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.Sizable,
                MinimizeBox = false,
                BackColor = Color.FromArgb(245, 245, 245)
            };

            // Danh sách MPR của dự án
            dlg.Controls.Add(new Label { Text = $"MPR của dự án: {projCode}", Location = new Point(10, 8), Size = new Size(400, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212) });
            var dgvMPRList = new DataGridView
            {
                Location = new Point(10, 32),
                Size = new Size(260, 580),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom
            };
            dgvMPRList.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvMPRList.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvMPRList.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvMPRList.EnableHeadersVisualStyles = false;
            dgvMPRList.Columns.Add(new DataGridViewTextBoxColumn { Name = "RMprId", Visible = false });
            dgvMPRList.Columns.Add(new DataGridViewTextBoxColumn { Name = "RMprNo", HeaderText = "MPR No", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            dgvMPRList.Columns.Add(new DataGridViewTextBoxColumn { Name = "RMprRev", HeaderText = "Rev", Width = 45 });
            dgvMPRList.Columns.Add(new DataGridViewTextBoxColumn { Name = "RIsLatest", Visible = false }); // "1"=bản mới nhất
            dlg.Controls.Add(dgvMPRList);

            // Load MPR theo project
            try
            {
                using var conn = DatabaseHelper.GetConnection();
                conn.Open();
                var cmd = new SqlCommand(
                    "SELECT MPR_ID, MPR_No, " +
                    "ISNULL(TRY_CAST(TRY_CAST(Rev AS DECIMAL(10,2)) AS INT), 0) AS Rev, " +
                    "ISNULL(Is_Latest,1) AS Is_Latest " +
                    "FROM MPR_Header WHERE Project_Code=@code " +
                    "ORDER BY ISNULL(TRY_CAST(TRY_CAST(Rev AS DECIMAL(10,2)) AS INT), 0) DESC, MPR_No", conn);
                cmd.Parameters.AddWithValue("@code", projCode);
                using var rdr = cmd.ExecuteReader();
                var temp = new List<(int Id, string No, int Rev)>();
                while (rdr.Read())
                {
                    temp.Add((Convert.ToInt32(rdr["MPR_ID"]), rdr["MPR_No"]?.ToString() ?? "", Convert.ToInt32(rdr["Rev"])));
                }
                rdr.Close();

                // Determine latest per (Project_Code, baseNo)
                var groups = new Dictionary<string, (int MaxRev, int MprId)>();
                foreach (var t in temp)
                {
                    string baseNo = !string.IsNullOrEmpty(t.No) && t.No.Contains("_Rev.")
                        ? t.No.Substring(0, t.No.IndexOf("_Rev."))
                        : t.No;
                    string key = (projCode ?? "") + "||" + baseNo;
                    if (!groups.TryGetValue(key, out var cur) || t.Rev > cur.MaxRev)
                        groups[key] = (t.Rev, t.Id);
                }
                var latestIds = new HashSet<int>(groups.Values.Select(v => v.MprId));

                int latestRowIdx = -1;
                for (int i = 0; i < temp.Count; i++)
                {
                    var t = temp[i];
                    int r = dgvMPRList.Rows.Add();
                    dgvMPRList.Rows[r].Cells["RMprId"].Value = t.Id;
                    dgvMPRList.Rows[r].Cells["RMprNo"].Value = t.No;
                    dgvMPRList.Rows[r].Cells["RMprRev"].Value = t.Rev;
                    bool isLatest = latestIds.Contains(t.Id);
                    dgvMPRList.Rows[r].Cells["RIsLatest"].Value = isLatest ? "1" : "0";
                    // Determine whether this row should be shown as gray: Rev > 1 and Rev < MaxRev within the same series
                    string baseNo = !string.IsNullOrEmpty(t.No) && t.No.Contains("_Rev.") ? t.No.Substring(0, t.No.IndexOf("_Rev.")) : t.No;
                    string key = (projCode ?? "") + "||" + baseNo;
                    int maxRev = groups.TryGetValue(key, out var g) ? g.MaxRev : t.Rev;
                    bool shouldGray = !isLatest && t.Rev > 1 && t.Rev < maxRev;
                    if (isLatest && latestRowIdx < 0)
                    {
                        dgvMPRList.Rows[r].DefaultCellStyle.ForeColor = Color.FromArgb(0, 120, 212);
                        dgvMPRList.Rows[r].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                        latestRowIdx = r;
                    }
                    else if (shouldGray)
                    {
                        dgvMPRList.Rows[r].DefaultCellStyle.ForeColor = Color.FromArgb(160, 160, 160);
                    }
                    else
                    {
                        // Normal display for Rev = 0, Rev = 1, or Rev == maxRev
                        dgvMPRList.Rows[r].DefaultCellStyle.ForeColor = Color.Black;
                    }
                }

                // Tự động chọn MPR mới nhất → load details ngay
                if (latestRowIdx >= 0)
                {
                    dgvMPRList.Rows[latestRowIdx].Selected = true;
                    dgvMPRList.CurrentCell = dgvMPRList.Rows[latestRowIdx].Cells["RMprNo"];
                }
            }
            catch { }

            // Bảng chi tiết
            dlg.Controls.Add(new Label { Text = "CHI TIẾT VẬT TƯ", Location = new Point(278, 8), Size = new Size(400, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212) });
            var dgvRevDet = new DataGridView
            {
                Location = new Point(278, 32),
                Size = new Size(dlg.ClientSize.Width - 288, 530),
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 8),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2
            };
            dgvRevDet.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvRevDet.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvRevDet.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 8, FontStyle.Bold);
            dgvRevDet.EnableHeadersVisualStyles = false;
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RDetId", Visible = false });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RDeleted", Visible = false }); // "1"=xóa mềm
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "ROrigHash", Visible = false }); // snapshot khi load
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RIsNew", Visible = false }); // "1"=dòng mới thêm
            // Cột đồng bộ với dgvDet (Tạo MPR): STT, Tên hàng, Mô tả, Vật liệu, T,D,W,Web,Flange,L, Vị trí, MPS, REV, DWG, Issue, ĐVT, SL, KG, Ghi chú
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RItem_No", HeaderText = "STT", Width = 40, ReadOnly = true });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ritem_name", HeaderText = "Tên hàng", Width = 140 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RDesc", HeaderText = "Mô tả", Width = 120 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RMaterial", HeaderText = "Vật liệu", Width = 80 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RT_mm", HeaderText = "T(mm)", Width = 55 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RD_mm", HeaderText = "D(mm)", Width = 55 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RW_mm", HeaderText = "W(mm)", Width = 55 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RWeb_mm", HeaderText = "Web(mm)", Width = 60 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RFlange_mm", HeaderText = "Flange(mm)", Width = 70 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RL_mm", HeaderText = "L(mm)", Width = 60 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RUsage", HeaderText = "Vị trí", Width = 90 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RMPS", HeaderText = "MPS", Width = 70 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RREV", HeaderText = "REV", Width = 45, ReadOnly = true });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RDWG", HeaderText = "DWG Date", Width = 80 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RIssue", HeaderText = "Issue Date", Width = 80 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RUNIT", HeaderText = "Đơn vị", Width = 55 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RQty", HeaderText = "SL", Width = 45 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RKG", HeaderText = "KG", Width = 55 });
            dgvRevDet.Columns.Add(new DataGridViewTextBoxColumn { Name = "RRemarks", HeaderText = "Ghi chú", Width = 100 });
            dlg.Controls.Add(dgvRevDet);

            // ── Paste từ Excel vào dgvRevDet (Ctrl+V) ────────────────────
            dgvRevDet.KeyDown += (s2, ev2) =>
            {
                if (!(ev2.Control && ev2.KeyCode == Keys.V)) return;
                ev2.Handled = true;
                ev2.SuppressKeyPress = true;

                string clip = Clipboard.GetText();
                if (string.IsNullOrEmpty(clip)) return;

                // If grid was readonly (e.g., non-latest), temporarily allow edits for paste operation
                bool prevReadOnly = dgvRevDet.ReadOnly;
                if (prevReadOnly) dgvRevDet.ReadOnly = false;
                dgvRevDet.EndEdit();
                // Dùng None thay vì RemoveEmptyEntries để không bỏ dòng cuối có ô trống.
                // Chỉ loại bỏ phần tử rỗng hoàn toàn ở CUỐI cùng (do Excel tự thêm \r\n).
                string[] pasteRows = clip.Split(
                    new[] { "\r\n", "\n" }, StringSplitOptions.None);
                if (pasteRows.Length > 0 && string.IsNullOrEmpty(pasteRows[pasteRows.Length - 1]))
                    pasteRows = pasteRows.Take(pasteRows.Length - 1).ToArray();
                if (pasteRows.Length == 0) return;

                // Cột hiển thị có thể paste (bỏ ẩn và ReadOnly)
                // Thứ tự visible: STT(ro), Tên hàng, Mô tả, Vật liệu, T,D,W,Web,Flange,L, ĐVT,SL,KG, Ghi chú, REV(ro)
                var pasteCols = new[] {
                    "Ritem_name", "RDesc", "RMaterial",
                    "RT_mm", "RD_mm", "RW_mm", "RWeb_mm", "RFlange_mm", "RL_mm",
                    "RUsage", "RMPS", "RDWG", "RIssue",
                    "RUNIT", "RQty", "RKG", "RRemarks"
                };

                // Tìm startCol trong pasteCols từ CurrentCell
                int startPasteCol = 0;
                if (dgvRevDet.CurrentCell != null)
                {
                    string curName = dgvRevDet.Columns[dgvRevDet.CurrentCell.ColumnIndex].Name;
                    int ci = System.Array.IndexOf(pasteCols, curName);
                    if (ci >= 0) startPasteCol = ci;
                }

                // startRow: ưu tiên CurrentCell, nếu là dòng ẩn/deleted thì xuống
                int curRowIdx = dgvRevDet.CurrentCell?.RowIndex ?? 0;
                // Đếm data rows (không tính ẩn)
                int dataRows = dgvRevDet.Rows.Count - (dgvRevDet.AllowUserToAddRows ? 1 : 0);
                int startRow = (curRowIdx >= dataRows) ? dataRows : curRowIdx;

                // Bước 1: Pre-add tất cả rows cần thiết
                int rowsNeeded = startRow + pasteRows.Length - dataRows;
                dgvRevDet.SuspendLayout();
                for (int ri = 0; ri < rowsNeeded; ri++)
                {
                    dgvRevDet.AllowUserToAddRows = true;
                    int ni = dgvRevDet.Rows.Add();
                    dgvRevDet.AllowUserToAddRows = false;
                    // Đánh dấu dòng mới thêm
                    dgvRevDet.Rows[ni].Cells["RIsNew"].Value = "1";
                    dgvRevDet.Rows[ni].Cells["ROrigHash"].Value = "__NEW__";
                }

                // Bước 2: Ghi data
                for (int ri2 = 0; ri2 < pasteRows.Length; ri2++)
                {
                    string[] cells = pasteRows[ri2].Split('\t');
                    int targetRow = startRow + ri2;
                    if (targetRow >= dgvRevDet.Rows.Count) break;
                    // Bỏ qua dòng bị xóa mềm
                    if (dgvRevDet.Rows[targetRow].Cells["RDeleted"].Value?.ToString() == "1") continue;
                    // Đánh dấu dòng cũ được paste vào là có thay đổi
                    if (dgvRevDet.Rows[targetRow].Cells["RIsNew"].Value?.ToString() != "1")
                        dgvRevDet.Rows[targetRow].Cells["ROrigHash"].Value = "__CHANGED__";
                    for (int c = 0; c < cells.Length && startPasteCol + c < pasteCols.Length; c++)
                    {
                        string colName = pasteCols[startPasteCol + c];
                        var col = dgvRevDet.Columns[colName];
                        if (col != null && !col.ReadOnly)
                            dgvRevDet.Rows[targetRow].Cells[colName].Value = cells[c].Trim();
                    }
                }
                dgvRevDet.ResumeLayout();
                dgvRevDet.Refresh();
                // Restore readonly state
                if (prevReadOnly) dgvRevDet.ReadOnly = true;
            };

            // Hiệu ứng: dòng xóa mềm = xám gạch ngang
            dgvRevDet.CellFormatting += (s2, ev2) =>
            {
                if (ev2.RowIndex < 0) return;
                string del = dgvRevDet.Rows[ev2.RowIndex].Cells["RDeleted"].Value?.ToString() ?? "";
                if (del == "1") { ev2.CellStyle.ForeColor = Color.FromArgb(180, 180, 180); ev2.CellStyle.Font = new Font("Segoe UI", 8, FontStyle.Strikeout); }
            };

            int selMprId = 0;

            // Buttons Thêm / Xóa vật tư
            var btnAddRow = new Button
            {
                Text = "+ Thêm vật tư",
                Location = new Point(278, dlg.ClientSize.Height - 58),
                Size = new Size(120, 30),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            btnAddRow.FlatAppearance.BorderSize = 0;
            btnAddRow.Click += (s2, ev2) =>
            {
                dgvRevDet.AllowUserToAddRows = true;
                int ni = dgvRevDet.Rows.Add();
                dgvRevDet.AllowUserToAddRows = false;
                dgvRevDet.Rows[ni].Cells["RIsNew"].Value = "1";       // dòng mới
                dgvRevDet.Rows[ni].Cells["ROrigHash"].Value = "__NEW__"; // luôn nhảy REV
            };

            var btnDelRow = new Button
            {
                Text = "🗑 Xóa vật tư",
                Location = new Point(406, dlg.ClientSize.Height - 58),
                Size = new Size(120, 30),
                BackColor = Color.FromArgb(220, 53, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            btnDelRow.FlatAppearance.BorderSize = 0;
            // Xóa mềm: đánh dấu "1" vào cột RDeleted
            btnDelRow.Click += (s2, ev2) =>
            {
                foreach (DataGridViewRow row2 in dgvRevDet.SelectedRows)
                {
                    if (row2.IsNewRow) continue;
                    string cur = row2.Cells["RDeleted"].Value?.ToString() ?? "";
                    row2.Cells["RDeleted"].Value = cur == "1" ? "" : "1"; // toggle
                }
                dgvRevDet.Refresh();
            };

            var lblSave = new Label { Location = new Point(10, dlg.ClientSize.Height - 30), Size = new Size(500, 22), ForeColor = Color.Red, Font = new Font("Segoe UI", 8), Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            dlg.Controls.Add(lblSave);

            var btnSaveMPR = new Button
            {
                Text = "💾 Lưu MPR",
                Location = new Point(dlg.ClientSize.Width - 312, dlg.ClientSize.Height - 58),
                Size = new Size(120, 30),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            btnSaveMPR.FlatAppearance.BorderSize = 0;

            // Khi chọn MPR → load details (cho phép chọn bất kỳ MPR; chỉ Revise khi là bản mới nhất)
            dgvMPRList.SelectionChanged += (s2, ev2) =>
            {
                if (dgvMPRList.SelectedRows.Count == 0) return;
                var selectedRow = dgvMPRList.SelectedRows[0];

                bool isLatestRow = selectedRow.Cells["RIsLatest"].Value?.ToString() == "1";
                // Nếu người dùng là Admin thì luôn được phép lưu
                btnSaveMPR.Enabled = isLatestRow || isAdmin;
                dgvRevDet.ReadOnly = !(isLatestRow || isAdmin);

                selMprId = Convert.ToInt32(selectedRow.Cells["RMprId"].Value ?? 0);
                dgvRevDet.Rows.Clear();
                try
                {
                    using var conn = DatabaseHelper.GetConnection();
                    conn.Open();
                    var cmd = new SqlCommand("SELECT * FROM MPR_Details WHERE MPR_ID=@id ORDER BY Item_No", conn);
                    cmd.Parameters.AddWithValue("@id", selMprId);
                    using var rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        int r = dgvRevDet.Rows.Add();
                        dgvRevDet.Rows[r].Cells["RDetId"].Value = rdr["Detail_ID"];
                        dgvRevDet.Rows[r].Cells["RDeleted"].Value = "";
                        dgvRevDet.Rows[r].Cells["RItem_No"].Value = rdr["Item_No"];
                        dgvRevDet.Rows[r].Cells["Ritem_name"].Value = rdr["item_name"];
                        dgvRevDet.Rows[r].Cells["RDesc"].Value = rdr["Description"];
                        dgvRevDet.Rows[r].Cells["RMaterial"].Value = rdr["Material"];
                        dgvRevDet.Rows[r].Cells["RT_mm"].Value = rdr["Thickness_mm"];
                        dgvRevDet.Rows[r].Cells["RD_mm"].Value = rdr["Depth_mm"];
                        dgvRevDet.Rows[r].Cells["RW_mm"].Value = rdr["C_Width_mm"];
                        dgvRevDet.Rows[r].Cells["RWeb_mm"].Value = rdr["D_Web_mm"];
                        dgvRevDet.Rows[r].Cells["RFlange_mm"].Value = rdr["E_Flange_mm"];
                        dgvRevDet.Rows[r].Cells["RL_mm"].Value = rdr["F_Length_mm"];
                        dgvRevDet.Rows[r].Cells["RUsage"].Value = rdr["Usage_Location"];
                        dgvRevDet.Rows[r].Cells["RMPS"].Value = rdr["MPS_Info"];
                        dgvRevDet.Rows[r].Cells["RDWG"].Value = rdr["DWG_BOQ_Receive_Date"] != DBNull.Value
                            ? Convert.ToDateTime(rdr["DWG_BOQ_Receive_Date"]).ToString("dd/MM/yyyy") : "";
                        dgvRevDet.Rows[r].Cells["RIssue"].Value = rdr["Issue_Date"] != DBNull.Value
                            ? Convert.ToDateTime(rdr["Issue_Date"]).ToString("dd/MM/yyyy") : "";
                        dgvRevDet.Rows[r].Cells["RUNIT"].Value = rdr["UNIT"];
                        dgvRevDet.Rows[r].Cells["RQty"].Value = rdr["Qty_Per_Sheet"];
                        dgvRevDet.Rows[r].Cells["RKG"].Value = rdr["Weight_kg"];
                        dgvRevDet.Rows[r].Cells["RRemarks"].Value = rdr["Remarks"];
                        dgvRevDet.Rows[r].Cells["RREV"].Value = rdr["REV"];
                        // Snapshot hash để detect thay đổi khi Lưu
                        dgvRevDet.Rows[r].Cells["RIsNew"].Value = ""; // dòng từ DB
                        string Norm(object v) => (v == DBNull.Value || v == null) ? "" : v.ToString()!.Trim();
                        dgvRevDet.Rows[r].Cells["ROrigHash"].Value = string.Join("|",
                            Norm(rdr["item_name"]), Norm(rdr["Description"]), Norm(rdr["Material"]),
                            Norm(rdr["Thickness_mm"]), Norm(rdr["Depth_mm"]), Norm(rdr["C_Width_mm"]),
                            Norm(rdr["D_Web_mm"]), Norm(rdr["E_Flange_mm"]), Norm(rdr["F_Length_mm"]),
                            Norm(rdr["Usage_Location"]), Norm(rdr["MPS_Info"]),
                            Norm(rdr["UNIT"]), Norm(rdr["Qty_Per_Sheet"]), Norm(rdr["Weight_kg"]),
                            Norm(rdr["Remarks"]));
                    }
                }
                catch { }
            };

            // Nếu đã có row được chọn trước khi handler đăng ký (auto-select latest),
            // thì khởi tạo trạng thái ban đầu tương tự để btnSaveMPR và selMprId có giá trị.
            if (dgvMPRList.SelectedRows.Count > 0)
            {
                var selectedRow = dgvMPRList.SelectedRows[0];
                bool isLatestRow = selectedRow.Cells["RIsLatest"].Value?.ToString() == "1";
                btnSaveMPR.Enabled = isLatestRow || isAdmin;
                dgvRevDet.ReadOnly = !(isLatestRow || isAdmin);
                selMprId = Convert.ToInt32(selectedRow.Cells["RMprId"].Value ?? 0);
                // Load details now (handler would have done this)
                try
                {
                    using var conn = DatabaseHelper.GetConnection();
                    conn.Open();
                    var cmd = new SqlCommand("SELECT * FROM MPR_Details WHERE MPR_ID=@id ORDER BY Item_No", conn);
                    cmd.Parameters.AddWithValue("@id", selMprId);
                    using var rdr = cmd.ExecuteReader();
                    dgvRevDet.Rows.Clear();
                    while (rdr.Read())
                    {
                        int r = dgvRevDet.Rows.Add();
                        dgvRevDet.Rows[r].Cells["RDetId"].Value = rdr["Detail_ID"];
                        dgvRevDet.Rows[r].Cells["RDeleted"].Value = "";
                        dgvRevDet.Rows[r].Cells["RItem_No"].Value = rdr["Item_No"];
                        dgvRevDet.Rows[r].Cells["Ritem_name"].Value = rdr["item_name"];
                        dgvRevDet.Rows[r].Cells["RDesc"].Value = rdr["Description"];
                        dgvRevDet.Rows[r].Cells["RMaterial"].Value = rdr["Material"];
                        dgvRevDet.Rows[r].Cells["RT_mm"].Value = rdr["Thickness_mm"];
                        dgvRevDet.Rows[r].Cells["RD_mm"].Value = rdr["Depth_mm"];
                        dgvRevDet.Rows[r].Cells["RW_mm"].Value = rdr["C_Width_mm"];
                        dgvRevDet.Rows[r].Cells["RWeb_mm"].Value = rdr["D_Web_mm"];
                        dgvRevDet.Rows[r].Cells["RFlange_mm"].Value = rdr["E_Flange_mm"];
                        dgvRevDet.Rows[r].Cells["RL_mm"].Value = rdr["F_Length_mm"];
                        dgvRevDet.Rows[r].Cells["RUsage"].Value = rdr["Usage_Location"];
                        dgvRevDet.Rows[r].Cells["RMPS"].Value = rdr["MPS_Info"];
                        dgvRevDet.Rows[r].Cells["RDWG"].Value = rdr["DWG_BOQ_Receive_Date"] != DBNull.Value
                            ? Convert.ToDateTime(rdr["DWG_BOQ_Receive_Date"]).ToString("dd/MM/yyyy") : "";
                        dgvRevDet.Rows[r].Cells["RIssue"].Value = rdr["Issue_Date"] != DBNull.Value
                            ? Convert.ToDateTime(rdr["Issue_Date"]).ToString("dd/MM/yyyy") : "";
                        dgvRevDet.Rows[r].Cells["RUNIT"].Value = rdr["UNIT"];
                        dgvRevDet.Rows[r].Cells["RQty"].Value = rdr["Qty_Per_Sheet"];
                        dgvRevDet.Rows[r].Cells["RKG"].Value = rdr["Weight_kg"];
                        dgvRevDet.Rows[r].Cells["RRemarks"].Value = rdr["Remarks"];
                        dgvRevDet.Rows[r].Cells["RREV"].Value = rdr["REV"];
                        dgvRevDet.Rows[r].Cells["RIsNew"].Value = "";
                        string Norm(object v) => (v == DBNull.Value || v == null) ? "" : v.ToString()!.Trim();
                        dgvRevDet.Rows[r].Cells["ROrigHash"].Value = string.Join("|",
                            Norm(rdr["item_name"]), Norm(rdr["Description"]), Norm(rdr["Material"]),
                            Norm(rdr["Thickness_mm"]), Norm(rdr["Depth_mm"]), Norm(rdr["C_Width_mm"]),
                            Norm(rdr["D_Web_mm"]), Norm(rdr["E_Flange_mm"]), Norm(rdr["F_Length_mm"]),
                            Norm(rdr["Usage_Location"]), Norm(rdr["MPS_Info"]),
                            Norm(rdr["UNIT"]), Norm(rdr["Qty_Per_Sheet"]), Norm(rdr["Weight_kg"]),
                            Norm(rdr["Remarks"]));
                    }
                }
                catch { }
            }

            var btnCloseRev = new Button
            {
                Text = "Đóng",
                Location = new Point(dlg.ClientSize.Width - 184, dlg.ClientSize.Height - 58),
                Size = new Size(80, 30),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                DialogResult = DialogResult.Cancel
            };
            btnCloseRev.FlatAppearance.BorderSize = 0;
            dlg.Controls.AddRange(new Control[] { btnAddRow, btnDelRow, btnSaveMPR, btnCloseRev });
            dlg.CancelButton = btnCloseRev;

            dlg.Resize += (s2, ev2) =>
            {
                int btnY = dlg.ClientSize.Height - 50;
                int gridH = Math.Max(80, btnY - 36 - 8);
                dgvMPRList.Height = gridH;
                dgvRevDet.Width = dlg.ClientSize.Width - 288;
                dgvRevDet.Height = gridH;
                btnAddRow.Top = btnDelRow.Top = btnSaveMPR.Top = btnCloseRev.Top = btnY;
                btnSaveMPR.Left = dlg.ClientSize.Width - 312;
                btnCloseRev.Left = dlg.ClientSize.Width - 184;
                lblSave.Top = dlg.ClientSize.Height - 26;
            };

            // Lưu MPR Revise
            btnSaveMPR.Click += (s2, ev2) =>
            {
                // Basic validation
                if (dgvMPRList.SelectedRows.Count == 0) { lblSave.Text = "⚠ Chọn MPR cần Revise!"; return; }
                var selRow = dgvMPRList.SelectedRows[0];
                bool isLatestRowNow = selRow.Cells["RIsLatest"].Value?.ToString() == "1";
                if (selMprId == 0) { lblSave.Text = "⚠ Chọn MPR cần Revise!"; return; }
                // Non-admin users may only revise from the latest revision
                if (!isAdmin && !isLatestRowNow) { lblSave.Text = "⚠ Chỉ được Revise từ bản mới nhất!"; return; }
                try
                {
                    using var conn = DatabaseHelper.GetConnection();
                    conn.Open();

                    // ── Lấy baseMprNo trước để query maxRev đúng từ MPR_Header ──
                    string oldMprNo = dgvMPRList.SelectedRows[0].Cells["RMprNo"].Value?.ToString() ?? "";
                    string baseMprNo = oldMprNo.Contains("_Rev.")
                        ? oldMprNo.Substring(0, oldMprNo.IndexOf("_Rev."))
                        : oldMprNo;

                    // ── Tính maxRev đúng: đọc từ MPR_Header.Rev của cùng chuỗi MPR_No ──
                    var cmdMaxRev = new SqlCommand(
                        @"SELECT ISNULL(MAX(TRY_CAST(TRY_CAST(Rev AS DECIMAL(10,2)) AS INT)), 0)
                          FROM MPR_Header
                          WHERE MPR_No = @baseNo OR MPR_No LIKE @baseNo + '_Rev.%'", conn);
                    cmdMaxRev.Parameters.AddWithValue("@baseNo", baseMprNo);
                    // Rev đã được chuẩn hóa về INT trong DB
                    int maxRev = Convert.ToInt32(cmdMaxRev.ExecuteScalar());
                    int nextRev = isAdmin ? maxRev : maxRev + 1;
                    string newMprNo = isAdmin ? oldMprNo : $"{baseMprNo}_Rev.{nextRev}";

                    // ── Kiểm tra trùng MPR_No trước khi INSERT ──
                    if (!isAdmin)
                    {
                        var cmdChk = new SqlCommand(
                            "SELECT COUNT(1) FROM MPR_Header WHERE MPR_No = @no", conn);
                        cmdChk.Parameters.AddWithValue("@no", newMprNo);
                        int exists = Convert.ToInt32(cmdChk.ExecuteScalar());
                        if (exists > 0)
                        {
                            // Auto-tăng nextRev cho đến khi tìm được MPR_No chưa tồn tại
                            while (exists > 0)
                            {
                                nextRev++;
                                newMprNo = $"{baseMprNo}_Rev.{nextRev}";
                                cmdChk.Parameters["@no"].Value = newMprNo;
                                exists = Convert.ToInt32(cmdChk.ExecuteScalar());
                            }
                        }
                    }

                    // Nếu không phải Admin → tạo bản Revise mới
                    int targetId = selMprId;
                    if (!isAdmin)
                    {
                        // Lấy header cũ
                        var cmdH = new SqlCommand("SELECT * FROM MPR_Header WHERE MPR_ID=@id", conn);
                        cmdH.Parameters.AddWithValue("@id", selMprId);
                        using var rdrH = cmdH.ExecuteReader();
                        rdrH.Read();
                        var newHeader = new MPRHeader
                        {
                            MPR_ID = 0,
                            MPR_No = newMprNo,
                            Project_Name = rdrH["Project_Name"]?.ToString() ?? "",
                            Project_Code = rdrH["Project_Code"]?.ToString() ?? "",
                            Department = rdrH["Department"]?.ToString() ?? "",
                            Requestor = rdrH["Requestor"]?.ToString() ?? "",
                            Required_Date = rdrH["Required_Date"] is DateTime dt ? dt : DateTime.Today,
                            Rev = nextRev,
                            Status = rdrH["Status"]?.ToString() ?? "Mới",
                            Notes = rdrH["Notes"]?.ToString() ?? ""
                        };
                        rdrH.Close();
                        targetId = _service.InsertHeader(newHeader, _currentUser);
                        // Đánh dấu MPR cũ Is_Latest=0, MPR mới Is_Latest=1
                        var cmdLatest = new SqlCommand(@"
                            UPDATE MPR_Header SET Is_Latest = 0
                            WHERE Project_Code = (SELECT Project_Code FROM MPR_Header WHERE MPR_ID=@oldId)
                              AND (MPR_No = @baseNo OR MPR_No LIKE @baseNo + '_Rev.%');
                            UPDATE MPR_Header SET Is_Latest = 1 WHERE MPR_ID = @newId;", conn);
                        cmdLatest.Parameters.AddWithValue("@oldId", selMprId);
                        cmdLatest.Parameters.AddWithValue("@baseNo", baseMprNo);
                        cmdLatest.Parameters.AddWithValue("@newId", targetId);
                        cmdLatest.ExecuteNonQuery();
                    }
                    else
                    {
                        // ── Admin: chỉnh sửa TẠI CHỖ — không tạo Rev mới, không xóa Details ──
                        // UPDATE dòng đã có Detail_ID, INSERT dòng mới, soft-delete dòng bị xóa
                        // → PO_Detail.MPR_Detail_ID không bị đụng đến → PO link giữ nguyên
                        int stt = 1;
                        foreach (DataGridViewRow row2 in dgvRevDet.Rows)
                        {
                            if (row2.IsNewRow) continue;
                            string nm = row2.Cells["Ritem_name"].Value?.ToString() ?? "";
                            if (string.IsNullOrWhiteSpace(nm)) continue;

                            bool isSoftDeleted = row2.Cells["RDeleted"].Value?.ToString() == "1";
                            int oldDetId = int.TryParse(row2.Cells["RDetId"].Value?.ToString(), out int odid) ? odid : 0;
                            int detRev = int.TryParse(row2.Cells["RREV"].Value?.ToString(), out int rvA) ? rvA : 0;

                            var det = new MPRDetail
                            {
                                MPR_ID = targetId,
                                Item_No = isSoftDeleted ? 0 : stt++,
                                Item_Name = nm,
                                Description = row2.Cells["RDesc"].Value?.ToString() ?? "",
                                Material = row2.Cells["RMaterial"].Value?.ToString() ?? "",
                                Thickness_mm = decimal.TryParse(row2.Cells["RT_mm"].Value?.ToString(), out decimal rTh) ? rTh : 0,
                                Depth_mm = decimal.TryParse(row2.Cells["RD_mm"].Value?.ToString(), out decimal rDp) ? rDp : 0,
                                C_Width_mm = decimal.TryParse(row2.Cells["RW_mm"].Value?.ToString(), out decimal rCw) ? rCw : 0,
                                D_Web_mm = decimal.TryParse(row2.Cells["RWeb_mm"].Value?.ToString(), out decimal rDw) ? rDw : 0,
                                E_Flange_mm = decimal.TryParse(row2.Cells["RFlange_mm"].Value?.ToString(), out decimal rEf) ? rEf : 0,
                                F_Length_mm = decimal.TryParse(row2.Cells["RL_mm"].Value?.ToString(), out decimal rFl) ? rFl : 0,
                                Usage_Location = row2.Cells["RUsage"].Value?.ToString() ?? "",
                                MPS_Info = row2.Cells["RMPS"].Value?.ToString() ?? "",
                                DWG_BOQ_Receive_Date = DateTime.TryParse(row2.Cells["RDWG"].Value?.ToString(), out DateTime rDwg) ? rDwg : (DateTime?)null,
                                Issue_Date = DateTime.TryParse(row2.Cells["RIssue"].Value?.ToString(), out DateTime rIss) ? rIss : (DateTime?)null,
                                UNIT = row2.Cells["RUNIT"].Value?.ToString() ?? "",
                                Qty_Per_Sheet = isSoftDeleted ? 0 : (decimal.TryParse(row2.Cells["RQty"].Value?.ToString(), out decimal rQs) ? rQs : 0),
                                Weight_kg = isSoftDeleted ? 0 : (decimal.TryParse(row2.Cells["RKG"].Value?.ToString(), out decimal rWk) ? rWk : 0),
                                Remarks = row2.Cells["RRemarks"].Value?.ToString() ?? "",
                                REV = detRev.ToString(),
                                Is_Deleted = isSoftDeleted
                            };

                            if (oldDetId > 0)
                                // Dòng đã tồn tại → UPDATE tại chỗ (PO link không bị ảnh hưởng)
                                _service.UpdateDetail(det, _currentUser, oldDetId);
                            else
                                // Dòng mới thêm → INSERT
                                _service.InsertDetail(det, _currentUser);
                        }

                        // Xóa cứng các dòng trong DB mà user đã xóa hẳn (không phải soft-delete)
                        // Đây là các Detail_ID cũ không còn xuất hiện trong grid
                        var currentIds = new HashSet<int>(
                            dgvRevDet.Rows.Cast<DataGridViewRow>()
                                .Where(r2 => !r2.IsNewRow)
                                .Select(r2 => int.TryParse(r2.Cells["RDetId"].Value?.ToString(), out int rid) ? rid : 0)
                                .Where(id => id > 0));

                        using (var cmdGetOld = new SqlCommand(
                            "SELECT Detail_ID FROM MPR_Details WHERE MPR_ID=@id", conn))
                        {
                            cmdGetOld.Parameters.AddWithValue("@id", targetId);
                            var oldIds = new List<int>();
                            using var r3 = cmdGetOld.ExecuteReader();
                            while (r3.Read()) oldIds.Add(Convert.ToInt32(r3["Detail_ID"]));
                            r3.Close();

                            foreach (int orphanId in oldIds.Where(id => !currentIds.Contains(id)))
                            {
                                // Trước khi xóa: NULL hóa PO link nếu có (bảo vệ PO data)
                                new SqlCommand(
                                    $"UPDATE PO_Detail SET MPR_Detail_ID=NULL WHERE MPR_Detail_ID={orphanId}", conn)
                                    .ExecuteNonQuery();
                                new SqlCommand(
                                    $"DELETE FROM MPR_Details WHERE Detail_ID={orphanId}", conn)
                                    .ExecuteNonQuery();
                            }
                        }
                    }

                    // ── Vòng lặp INSERT Details cho non-Admin (Revise tạo bản mới) ──────
                    if (!isAdmin)
                    {
                        int stt2 = 1;
                        foreach (DataGridViewRow row2 in dgvRevDet.Rows)
                        {
                            if (row2.IsNewRow) continue;
                            bool isSoftDeleted = row2.Cells["RDeleted"].Value?.ToString() == "1";
                            string nm = row2.Cells["Ritem_name"].Value?.ToString() ?? "";
                            if (string.IsNullOrWhiteSpace(nm)) continue;

                            bool isNewRow = row2.Cells["RIsNew"].Value?.ToString() == "1";
                            string origHash = row2.Cells["ROrigHash"].Value?.ToString() ?? "";
                            string C(DataGridViewCell c) => c.Value?.ToString()?.Trim() ?? "";
                            string curHash = string.Join("|",
                                C(row2.Cells["Ritem_name"]), C(row2.Cells["RDesc"]),
                                C(row2.Cells["RMaterial"]), C(row2.Cells["RT_mm"]),
                                C(row2.Cells["RD_mm"]), C(row2.Cells["RW_mm"]),
                                C(row2.Cells["RWeb_mm"]), C(row2.Cells["RFlange_mm"]),
                                C(row2.Cells["RL_mm"]), C(row2.Cells["RUsage"]),
                                C(row2.Cells["RMPS"]), C(row2.Cells["RUNIT"]),
                                C(row2.Cells["RQty"]), C(row2.Cells["RKG"]),
                                C(row2.Cells["RRemarks"]));
                            bool changed = isNewRow || origHash == "__NEW__" || curHash != origHash;
                            int detRev = changed
                                ? nextRev
                                : (int.TryParse(row2.Cells["RREV"].Value?.ToString(), out int rvO) ? rvO : 0);

                            _service.InsertDetail(new MPRDetail
                            {
                                MPR_ID = targetId,
                                Item_No = isSoftDeleted ? 0 : stt2++,
                                Item_Name = nm,
                                Description = row2.Cells["RDesc"].Value?.ToString() ?? "",
                                Material = row2.Cells["RMaterial"].Value?.ToString() ?? "",
                                Thickness_mm = decimal.TryParse(row2.Cells["RT_mm"].Value?.ToString(), out decimal rTh) ? rTh : 0,
                                Depth_mm = decimal.TryParse(row2.Cells["RD_mm"].Value?.ToString(), out decimal rDp) ? rDp : 0,
                                C_Width_mm = decimal.TryParse(row2.Cells["RW_mm"].Value?.ToString(), out decimal rCw) ? rCw : 0,
                                D_Web_mm = decimal.TryParse(row2.Cells["RWeb_mm"].Value?.ToString(), out decimal rDw) ? rDw : 0,
                                E_Flange_mm = decimal.TryParse(row2.Cells["RFlange_mm"].Value?.ToString(), out decimal rEf) ? rEf : 0,
                                F_Length_mm = decimal.TryParse(row2.Cells["RL_mm"].Value?.ToString(), out decimal rFl) ? rFl : 0,
                                Usage_Location = row2.Cells["RUsage"].Value?.ToString() ?? "",
                                MPS_Info = row2.Cells["RMPS"].Value?.ToString() ?? "",
                                DWG_BOQ_Receive_Date = DateTime.TryParse(row2.Cells["RDWG"].Value?.ToString(), out DateTime rDwg) ? rDwg : (DateTime?)null,
                                Issue_Date = DateTime.TryParse(row2.Cells["RIssue"].Value?.ToString(), out DateTime rIss) ? rIss : (DateTime?)null,
                                UNIT = row2.Cells["RUNIT"].Value?.ToString() ?? "",
                                Qty_Per_Sheet = isSoftDeleted ? 0 : (decimal.TryParse(row2.Cells["RQty"].Value?.ToString(), out decimal rQs) ? rQs : 0),
                                Weight_kg = isSoftDeleted ? 0 : (decimal.TryParse(row2.Cells["RKG"].Value?.ToString(), out decimal rWk) ? rWk : 0),
                                Remarks = row2.Cells["RRemarks"].Value?.ToString() ?? "",
                                REV = isSoftDeleted ? "0" : detRev.ToString(),
                                Is_Deleted = isSoftDeleted
                            }, _currentUser);
                        }

                        // Đánh dấu Is_Deleted=1 cho các dòng xóa mềm trong DB
                        var cmdMarkDel = new SqlCommand(
                            "UPDATE MPR_Details SET Is_Deleted=1 WHERE MPR_ID=@mid AND Item_No=0", conn);
                        cmdMarkDel.Parameters.AddWithValue("@mid", targetId);
                        cmdMarkDel.ExecuteNonQuery();
                    }

                    string msg = isAdmin
                        ? $"✅ Admin đã cập nhật MPR '{oldMprNo}' (giữ nguyên REV)!"
                        : $"✅ Đã tạo Revise '{newMprNo}' (REV={nextRev})!";
                    MessageBox.Show(TopOwner, msg, "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadMPR();
                    dlg.Close();
                }
                catch (Exception ex) { lblSave.Text = "❌ " + ex.Message; }
            };

            dlg.Owner = TopOwner;
            dlg.ShowDialog();
        }

        // =====================================================================
        // TẠO PO TỪ MPR
        // =====================================================================

        private void BtnSaveHeader_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("MPR", "Lưu Header", "Lưu Header")) return;
            if (string.IsNullOrWhiteSpace(txtMPRNo.Text))
            {
                MessageBox.Show(TopOwner, "Vui lòng nhập MPR No!", "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtMPRNo.Focus();
                return;
            }

            try
            {
                var m = new MPRHeader
                {
                    MPR_ID = _selectedMPR_ID,
                    MPR_No = txtMPRNo.Text.Trim(),
                    Project_Name = txtProjectName.Text.Trim(),
                    Project_Code = txtProjectCode.Text.Trim(),
                    Department = txtDepartment.Text.Trim(),
                    Requestor = txtRequestor.Text.Trim(),
                    Required_Date = dtpRequiredDate.Value,
                    Rev = int.TryParse(txtRev.Text, out int rev) ? rev : 0,
                    Status = cboStatus.SelectedItem?.ToString() ?? "Mới",
                    Notes = txtNotes.Text.Trim()
                };
                if (_selectedMPR_ID == 0)
                {
                    _selectedMPR_ID = _service.InsertHeader(m, _currentUser);
                    MessageBox.Show(TopOwner, "Tạo phiếu MPR thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    _service.UpdateHeader(m, _currentUser);
                    MessageBox.Show(TopOwner, "Cập nhật phiếu MPR thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                LoadMPR();
            }
            catch (Exception ex)
            {
                MessageBox.Show(TopOwner, "Lỗi lưu header: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =====================================================================
        // XÓA MPR — CASCADE DELETE ĐÚNG THỨ TỰ TRONG TRANSACTION
        // Thứ tự: PO_Detail (FK) → MPR_Details → MPR_Header
        // =====================================================================
        private void BtnDeleteMPR_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("MPR", "Xóa MPR", "Xóa MPR")) return;
            if (_selectedMPR_ID == 0)
            {
                MessageBox.Show(TopOwner, "Vui lòng chọn phiếu MPR cần xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var mprToDelete = _mprList.Find(m => m.MPR_ID == _selectedMPR_ID);
            string mprNoDisplay = mprToDelete?.MPR_No ?? _selectedMPR_ID.ToString();

            string confirmMsg =
                $"Bạn có chắc chắn muốn xóa phiếu MPR: [{mprNoDisplay}] ?\n\n" +
                $"⚠ Thao tác này sẽ xóa:\n" +
                $"   • Chi tiết vật tư của phiếu MPR này\n" +
                $"   • Liên kết PO sẽ được chuyển về phiên bản Rev trước (nếu có)\n\n" +
                $"Dữ liệu PO sẽ KHÔNG bị xóa!";

            if (MessageBox.Show(TopOwner, confirmMsg, "⚠ Xác nhận xóa MPR",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2) != DialogResult.Yes)
                return;

            try
            {
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    using (var tran = conn.BeginTransaction())
                    {
                        try
                        {
                            // Tính baseMprNo để xác định chuỗi Rev
                            string baseMprNo = mprNoDisplay.Contains("_Rev.")
                                ? mprNoDisplay.Substring(0, mprNoDisplay.IndexOf("_Rev."))
                                : mprNoDisplay;

                            // ── Bước 1: Tìm Rev ngay trước để chuyển PO_Detail sang ──
                            // Lấy MPR_ID của Rev cao nhất trong chuỗi (không tính MPR đang xóa)
                            var cmdPrevRev = new SqlCommand(@"
                                SELECT TOP 1 MPR_ID
                                FROM   MPR_Header
                                WHERE  MPR_ID <> @mprId
                                  AND  (MPR_No = @baseNo OR MPR_No LIKE @baseNo + N'_Rev.%')
                                ORDER BY TRY_CAST(TRY_CAST(Rev AS DECIMAL(10,2)) AS INT) DESC", conn, tran);
                            cmdPrevRev.Parameters.AddWithValue("@mprId", _selectedMPR_ID);
                            cmdPrevRev.Parameters.AddWithValue("@baseNo", baseMprNo);
                            object prevRevObj = cmdPrevRev.ExecuteScalar();
                            int prevRevMprId = prevRevObj != null && prevRevObj != DBNull.Value
                                ? Convert.ToInt32(prevRevObj) : 0;

                            int poDetailMoved = 0;
                            if (prevRevMprId > 0)
                            {
                                // ── Bước 2: Chuyển PO_Detail.MPR_Detail_ID sang Detail_ID
                                //            tương ứng của Rev trước (match Item_No + item_name) ──
                                // TUYỆT ĐỐI không xóa PO_Detail, chỉ cập nhật FK
                                var cmdMovePO = new SqlCommand(@"
                                    UPDATE pod
                                    SET    pod.MPR_Detail_ID = prev.Detail_ID
                                    FROM   PO_Detail pod
                                    INNER JOIN MPR_Details cur
                                           ON cur.Detail_ID = pod.MPR_Detail_ID
                                          AND cur.MPR_ID    = @mprId
                                    INNER JOIN MPR_Details prev
                                           ON prev.MPR_ID    = @prevMprId
                                          AND prev.Item_No   = cur.Item_No
                                          AND prev.item_name = cur.item_name", conn, tran);
                                cmdMovePO.Parameters.AddWithValue("@mprId", _selectedMPR_ID);
                                cmdMovePO.Parameters.AddWithValue("@prevMprId", prevRevMprId);
                                poDetailMoved = cmdMovePO.ExecuteNonQuery();
                            }
                            // Nếu không có Rev trước (MPR đơn lẻ): PO_Detail.MPR_Detail_ID
                            // sẽ bị orphan sau khi xóa Detail → set về NULL để giữ PO nguyên vẹn
                            else
                            {
                                var cmdNullPO = new SqlCommand(@"
                                    UPDATE pod
                                    SET    pod.MPR_Detail_ID = NULL
                                    FROM   PO_Detail pod
                                    INNER JOIN MPR_Details cur
                                           ON cur.Detail_ID = pod.MPR_Detail_ID
                                          AND cur.MPR_ID    = @mprId", conn, tran);
                                cmdNullPO.Parameters.AddWithValue("@mprId", _selectedMPR_ID);
                                cmdNullPO.ExecuteNonQuery();
                            }

                            // ── Bước 3: Xóa MPR_Details của MPR đang xóa ──
                            var cmdDelDet = new SqlCommand(
                                "DELETE FROM dbo.MPR_Details WHERE MPR_ID = @mprId", conn, tran);
                            cmdDelDet.Parameters.AddWithValue("@mprId", _selectedMPR_ID);
                            int detailDeleted = cmdDelDet.ExecuteNonQuery();

                            // ── Bước 4: Cập nhật Is_Latest đúng trong chuỗi MPR_No ──
                            var cmdLatest = new SqlCommand(@"
                                -- Reset toàn bộ chuỗi về 0
                                UPDATE MPR_Header
                                SET    Is_Latest = 0
                                WHERE  MPR_No = @baseNo
                                  OR   MPR_No LIKE @baseNo + N'_Rev.%';

                                -- Đặt Is_Latest=1 cho Rev cao nhất còn lại (không tính MPR đang xóa)
                                UPDATE MPR_Header
                                SET    Is_Latest = 1
                                WHERE  MPR_ID = (
                                    SELECT TOP 1 MPR_ID
                                    FROM   MPR_Header
                                    WHERE  MPR_ID <> @mprId
                                      AND  (MPR_No = @baseNo OR MPR_No LIKE @baseNo + N'_Rev.%')
                                    ORDER BY TRY_CAST(TRY_CAST(Rev AS DECIMAL(10,2)) AS INT) DESC
                                );", conn, tran);
                            cmdLatest.Parameters.AddWithValue("@mprId", _selectedMPR_ID);
                            cmdLatest.Parameters.AddWithValue("@baseNo", baseMprNo);
                            cmdLatest.ExecuteNonQuery();

                            // ── Bước 5: Xóa MPR Header ──
                            var cmdDelHead = new SqlCommand(
                                "DELETE FROM dbo.MPR_Header WHERE MPR_ID = @mprId", conn, tran);
                            cmdDelHead.Parameters.AddWithValue("@mprId", _selectedMPR_ID);
                            cmdDelHead.ExecuteNonQuery();

                            tran.Commit();

                            string resultMsg =
                                $"✅ Xóa phiếu MPR [{mprNoDisplay}] thành công!\n\n" +
                                $"   • {detailDeleted} dòng vật tư đã xóa\n" +
                                (poDetailMoved > 0
                                    ? $"   • {poDetailMoved} liên kết PO đã chuyển về Rev trước ✅\n"
                                    : $"   • Không có liên kết PO nào cần chuyển\n") +
                                $"   • Dữ liệu PO được giữ nguyên hoàn toàn ✅";
                            MessageBox.Show(TopOwner, resultMsg, "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch
                        {
                            tran.Rollback();
                            throw;
                        }
                    }
                }

                _selectedMPR_ID = 0;
                ClearHeader();
                dgvDetails.Rows.Clear();
                dgvPOProgress.DataSource = null;
                dgvFiles.Rows.Clear();
                LoadMPR();
            }
            catch (Exception ex)
            {
                MessageBox.Show(TopOwner, "Lỗi khi xóa MPR: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnClearHeader_Click(object sender, EventArgs e)
        {
            _selectedMPR_ID = 0;
            ClearHeader();
            dgvDetails.Rows.Clear();
            dgvPOProgress.DataSource = null;
            dgvFiles.Rows.Clear();
            _details.Clear();
        }

        private async void BtnCreatePO_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("MPR", "Tạo PO", "Tạo PO từ MPR")) return;
            if (_selectedMPR_ID == 0)
            {
                MessageBox.Show(TopOwner, "Vui lòng chọn một MPR trước!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var mpr = _mprList.Find(m => m.MPR_ID == _selectedMPR_ID);
            if (mpr == null) return;

            if (mpr.Status == "Hủy")
            {
                SafeMsg("Phiếu MPR này đã bị Hủy, không thể tạo PO!", "Thông báo", MessageBoxIcon.Warning);
                return;
            }

            string mprNo = mpr.MPR_No;

            this.Cursor = Cursors.WaitCursor;
            try
            {
                // Nhường 1 nhịp cho UI render cursor trước khi dựng form nặng.
                await Task.Yield();

                var frm = new frmPO(mprNo, true);
                frm.SetImportMprId(mpr.MPR_ID);
                frm.Show();

                // Đưa focus sang form PO ngay sau khi mở.
                frm.BringToFront();
                frm.Activate();
            }
            catch (Exception ex)
            {
                SafeMsg("Không thể mở form PO: " + ex.Message, "Lỗi");
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void BtnAddDetail_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("MPR", "Thêm dòng", "Thêm dòng")) return;
            if (_selectedMPR_ID == 0)
            {
                MessageBox.Show(TopOwner, "Vui lòng chọn hoặc lưu phiếu MPR trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int nextNo = dgvDetails.Rows.Count + 1;
            int newIdx = dgvDetails.Rows.Add();
            var newRow = dgvDetails.Rows[newIdx];

            newRow.Cells["Detail_ID"].Value = 0;
            newRow.Cells["Item_No"].Value = nextNo;
            newRow.Cells["Item_Name"].Value = "";
            newRow.Cells["Description"].Value = "";
            newRow.Cells["Material"].Value = "";
            newRow.Cells["Thickness_mm"].Value = 0;
            newRow.Cells["Depth_mm"].Value = 0;
            newRow.Cells["C_Width_mm"].Value = 0;
            newRow.Cells["D_Web_mm"].Value = 0;
            newRow.Cells["E_Flange_mm"].Value = 0;
            newRow.Cells["F_Length_mm"].Value = 0;
            newRow.Cells["UNIT"].Value = "cái";
            newRow.Cells["Qty"].Value = 0;
            newRow.Cells["Weight"].Value = 0;
            newRow.Cells["MPS_Info"].Value = "";
            newRow.Cells["Usage_Location"].Value = "";
            newRow.Cells["REV"].Value = "0";
            newRow.Cells["Remarks"].Value = "";
            newRow.Cells["PO_No"].Value = "";

            dgvDetails.CurrentCell = dgvDetails.Rows[newIdx].Cells["Item_Name"];
            dgvDetails.BeginEdit(true);
        }

        private Form TopOwner => this;
        private void SafeMsg(string text, string caption, MessageBoxIcon icon = MessageBoxIcon.Error)
        {
            var owner = TopOwner;
            if (owner != null && !owner.IsDisposed)
            {
                owner.BringToFront();
                owner.Activate();
                MessageBox.Show(owner, text, caption, MessageBoxButtons.OK, icon);
            }
            else
            {
                MessageBox.Show(text, caption, MessageBoxButtons.OK, icon);
            }
        }
        private DialogResult SafeConfirm(string text, string caption)
        {
            var owner = TopOwner;
            if (owner != null && !owner.IsDisposed)
            {
                owner.BringToFront();
                owner.Activate();
                return MessageBox.Show(owner, text, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            }
            else
            {
                return MessageBox.Show(text, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            }
        }

        private void BtnDeleteDetail_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("MPR", "Xóa dòng", "Xóa dòng")) return;
            if (dgvDetails.SelectedRows.Count == 0)
            {
                SafeMsg("Vui lòng chọn dòng cần xóa!", "Thông báo", MessageBoxIcon.Warning);
                return;
            }

            string msg = dgvDetails.SelectedRows.Count == 1
                ? "Bạn có chắc chắn muốn xóa dòng này?"
                : $"Bạn có chắc chắn muốn xóa {dgvDetails.SelectedRows.Count} dòng đã chọn?";
            if (SafeConfirm(msg, "Xác nhận xóa") == DialogResult.Yes)
            {
                try
                {
                    var rowsToDelete = new List<DataGridViewRow>();
                    foreach (DataGridViewRow row in dgvDetails.SelectedRows)
                    {
                        if (!row.IsNewRow)
                        {
                            rowsToDelete.Add(row);
                        }
                    }

                    foreach (var row in rowsToDelete)
                    {
                        dgvDetails.Rows.Remove(row);
                    }

                    int itemNo = 1;
                    foreach (DataGridViewRow row in dgvDetails.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            row.Cells["Item_No"].Value = itemNo++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    SafeMsg("Lỗi khi xóa: " + ex.Message, "Lỗi");
                }
            }
        }

        private void BtnSaveDetail_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("MPR", "Lưu chi tiết", "Lưu chi tiết")) return;
            if (_selectedMPR_ID == 0)
            {
                MessageBox.Show(TopOwner, "Vui lòng lưu header MPR trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (dgvDetails.Rows.Count == 0)
            {
                MessageBox.Show(TopOwner, "Không có dòng nào để lưu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ── Xác thực mật khẩu Admin trước khi lưu ──
            if (!VerifyAdminPassword()) return;

            try
            {
                int saved = 0;
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    foreach (DataGridViewRow row in dgvDetails.Rows)
                    {
                        if (row.IsNewRow) continue;
                        string itemName = row.Cells["Item_Name"].Value?.ToString() ?? "";
                        if (string.IsNullOrWhiteSpace(itemName)) continue;

                        int detailId = Convert.ToInt32(row.Cells["Detail_ID"].Value ?? 0);
                        int itemNo = Convert.ToInt32(row.Cells["Item_No"].Value ?? 0);
                        string desc = row.Cells["Description"].Value?.ToString() ?? "";
                        string material = row.Cells["Material"].Value?.ToString() ?? "";
                        decimal thickMm = DecimalVal(row.Cells["Thickness_mm"].Value);
                        decimal depthMm = DecimalVal(row.Cells["Depth_mm"].Value);
                        decimal cWidthMm = DecimalVal(row.Cells["C_Width_mm"].Value);
                        decimal dWebMm = DecimalVal(row.Cells["D_Web_mm"].Value);
                        decimal eFlangeMm = DecimalVal(row.Cells["E_Flange_mm"].Value);
                        decimal fLengthMm = DecimalVal(row.Cells["F_Length_mm"].Value);
                        string unit = row.Cells["UNIT"].Value?.ToString() ?? "";
                        int qty = (int)DecimalVal(row.Cells["Qty"].Value);
                        decimal weight = DecimalVal(row.Cells["Weight"].Value);
                        string mpsInfo = row.Cells["MPS_Info"].Value?.ToString() ?? "";
                        string usageLoc = row.Cells["Usage_Location"].Value?.ToString() ?? "";
                        string rev = row.Cells["REV"].Value?.ToString() ?? "0";
                        string remarks = row.Cells["Remarks"].Value?.ToString() ?? "";
                        string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                        Microsoft.Data.SqlClient.SqlCommand cmd;

                        if (detailId == 0)
                        {
                            cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                                INSERT INTO MPR_Details
                                    (MPR_ID, Item_No, item_name, Description, Material,
                                     Thickness_mm, Depth_mm, C_Width_mm, D_Web_mm, E_Flange_mm, F_Length_mm,
                                     UNIT, Qty_Per_Sheet, Weight_kg, MPS_Info, Usage_Location, REV, Remarks,
                                     Created_Date, Created_By, Modified_Date, Modified_By)
                                VALUES
                                    (@mprId, @itemNo, @itemName, @desc, @material,
                                     @thick, @depth, @cWidth, @dWeb, @eFlange, @fLen,
                                     @unit, @qty, @weight, @mps, @usage, @rev, @remarks,
                                     @now, @user, @now, @user);
                                SELECT SCOPE_IDENTITY();", conn);
                        }
                        else
                        {
                            cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                                UPDATE MPR_Details SET
                                    Item_No       = @itemNo,
                                    item_name     = @itemName,
                                    Description   = @desc,
                                    Material      = @material,
                                    Thickness_mm  = @thick,
                                    Depth_mm      = @depth,
                                    C_Width_mm    = @cWidth,
                                    D_Web_mm      = @dWeb,
                                    E_Flange_mm   = @eFlange,
                                    F_Length_mm   = @fLen,
                                    UNIT          = @unit,
                                    Qty_Per_Sheet = @qty,
                                    Weight_kg     = @weight,
                                    MPS_Info      = @mps,
                                    Usage_Location= @usage,
                                    REV           = @rev,
                                    Remarks       = @remarks,
                                    Modified_Date = @now,
                                    Modified_By   = @user
                                WHERE Detail_ID = @detailId", conn);
                            cmd.Parameters.AddWithValue("@detailId", detailId);
                        }

                        cmd.Parameters.AddWithValue("@mprId", _selectedMPR_ID);
                        cmd.Parameters.AddWithValue("@itemNo", itemNo);
                        cmd.Parameters.AddWithValue("@itemName", itemName);
                        cmd.Parameters.AddWithValue("@desc", desc);
                        cmd.Parameters.AddWithValue("@material", material);
                        cmd.Parameters.AddWithValue("@thick", thickMm);
                        cmd.Parameters.AddWithValue("@depth", depthMm);
                        cmd.Parameters.AddWithValue("@cWidth", cWidthMm);
                        cmd.Parameters.AddWithValue("@dWeb", dWebMm);
                        cmd.Parameters.AddWithValue("@eFlange", eFlangeMm);
                        cmd.Parameters.AddWithValue("@fLen", fLengthMm);
                        cmd.Parameters.AddWithValue("@unit", unit);
                        cmd.Parameters.AddWithValue("@qty", qty);
                        cmd.Parameters.AddWithValue("@weight", weight);
                        cmd.Parameters.AddWithValue("@mps", mpsInfo);
                        cmd.Parameters.AddWithValue("@usage", usageLoc);
                        cmd.Parameters.AddWithValue("@rev", rev);
                        cmd.Parameters.AddWithValue("@remarks", remarks);
                        cmd.Parameters.AddWithValue("@now", now);
                        cmd.Parameters.AddWithValue("@user", _currentUser ?? "Admin");

                        if (detailId == 0)
                        {
                            var newId = cmd.ExecuteScalar();
                            if (newId != null && newId != DBNull.Value)
                                row.Cells["Detail_ID"].Value = Convert.ToInt32(newId);
                        }
                        else
                        {
                            cmd.ExecuteNonQuery();
                        }

                        saved++;
                    }
                }

                MessageBox.Show(TopOwner, $"✅ Đã lưu {saved} dòng chi tiết thành công!", "Thành công",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDetails(_selectedMPR_ID);
            }
            catch (Exception ex)
            {
                MessageBox.Show(TopOwner, "Lỗi lưu chi tiết: " + ex.Message, "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =========================================================================
        // XÁC THỰC MẬT KHẨU ADMIN
        // =========================================================================
        private bool VerifyAdminPassword()
        {
            var dlg = new Form
            {
                Text = "🔐 Xác thực Admin",
                Size = new Size(380, 170),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(245, 245, 245),
                KeyPreview = true
            };

            dlg.Controls.Add(new Label
            {
                Text = "Nhập mật khẩu tài khoản Admin để xác nhận lưu:",
                Font = new Font("Segoe UI", 9),
                Location = new Point(15, 15),
                Size = new Size(340, 20)
            });

            var txtPwd = new TextBox
            {
                Location = new Point(15, 40),
                Size = new Size(340, 26),
                Font = new Font("Segoe UI", 10),
                PasswordChar = '●',
                UseSystemPasswordChar = false
            };
            dlg.Controls.Add(txtPwd);

            var lblErr = new Label
            {
                Text = "",
                ForeColor = Color.FromArgb(220, 53, 69),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Location = new Point(15, 72),
                Size = new Size(340, 20)
            };
            dlg.Controls.Add(lblErr);

            var btnOK = new Button
            {
                Text = "✔ Xác nhận",
                Location = new Point(155, 98),
                Size = new Size(100, 30),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            btnOK.FlatAppearance.BorderSize = 0;
            dlg.Controls.Add(btnOK);

            var btnCancel = new Button
            {
                Text = "Hủy",
                Location = new Point(265, 98),
                Size = new Size(90, 30),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                DialogResult = DialogResult.Cancel
            };
            btnCancel.FlatAppearance.BorderSize = 0;
            dlg.Controls.Add(btnCancel);
            dlg.CancelButton = btnCancel;

            bool verified = false;

            btnOK.Click += (s, ev) =>
            {
                string pwd = txtPwd.Text;
                if (string.IsNullOrEmpty(pwd))
                { lblErr.Text = "Vui lòng nhập mật khẩu!"; return; }

                try
                {
                    string inputHash;
                    using (var sha256 = System.Security.Cryptography.SHA256.Create())
                    {
                        byte[] bytes = sha256.ComputeHash(System.Text.Encoding.UTF8.GetBytes(pwd));
                        inputHash = BitConverter.ToString(bytes).Replace("-", "").ToLower();
                    }

                    const string ADMIN_HASH = "e86f78a8a3caf0b60d8e74e5942aa6d86dc150cd3c03338aef25b7d2d7e3acc7";

                    bool match = false;

                    if (inputHash == ADMIN_HASH)
                    {
                        match = true;
                    }
                    else
                    {
                        using var conn = DatabaseHelper.GetConnection();
                        conn.Open();

                        var cmd1 = new Microsoft.Data.SqlClient.SqlCommand(
                            @"SELECT COUNT(1) FROM Users
                              WHERE LOWER(Username) = 'admin'
                                AND (LOWER(Password) = @hash
                                  OR Password = @hashUpper)", conn);
                        cmd1.Parameters.AddWithValue("@hash", inputHash);
                        cmd1.Parameters.AddWithValue("@hashUpper", inputHash.ToUpper());
                        if (Convert.ToInt32(cmd1.ExecuteScalar()) > 0)
                            match = true;

                        if (!match)
                        {
                            var cmd2 = new Microsoft.Data.SqlClient.SqlCommand(
                                @"SELECT COUNT(1) FROM Users
                                  WHERE LOWER(Username) = 'admin'
                                    AND Password = @pwd", conn);
                            cmd2.Parameters.AddWithValue("@pwd", pwd);
                            if (Convert.ToInt32(cmd2.ExecuteScalar()) > 0)
                                match = true;
                        }
                    }

                    if (match)
                    {
                        verified = true;
                        dlg.DialogResult = DialogResult.OK;
                        dlg.Close();
                    }
                    else
                    {
                        lblErr.Text = "❌ Mật khẩu không đúng!";
                        txtPwd.Clear();
                        txtPwd.Focus();
                    }
                }
                catch (Exception ex)
                {
                    lblErr.Text = "Lỗi xác thực: " + ex.Message;
                }
            };

            dlg.KeyDown += (s, ev) =>
            {
                if (ev.KeyCode == Keys.Enter) { btnOK.PerformClick(); ev.SuppressKeyPress = true; }
            };

            txtPwd.Focus();
            dlg.ShowDialog(this);
            return verified;
        }

        // =========================================================================
        // CHECK ALL ITEMS — Popup tổng hợp toàn bộ MPR Detail + RIR kết quả KT
        // =========================================================================
        private void BtnCheckAllItems_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("MPR", "Check All Items", "Check All Items")) return;
            ShowCheckAllItemsPopup();
        }

        private void ShowCheckAllItemsPopup()
        {
            try
            {
                // SQL tối ưu: CTE pre-aggregate, tương thích SQL Server 2012+
                // Bước 1: Materialized Sib_Details vào temp table — tính 1 lần, tránh re-evaluate 4 lần
                const string SQL_SIB = @"
                    IF OBJECT_ID('tempdb..#SibDet') IS NOT NULL DROP TABLE #SibDet;
                    SELECT DISTINCT dC.Detail_ID AS CurrID, dS.Detail_ID AS SibID
                    INTO #SibDet
                    FROM (
                        SELECT MPR_ID, MPR_No,
                               ROW_NUMBER() OVER (
                                   PARTITION BY Project_Code,
                                       SUBSTRING(MPR_No, 1, CHARINDEX('_Rev.', MPR_No + '_Rev.') - 1)
                                   ORDER BY TRY_CAST(TRY_CAST(Rev AS DECIMAL(10,2)) AS INT) DESC, MPR_ID DESC
                               ) AS rn
                        FROM MPR_Header
                    ) q
                    INNER JOIN MPR_Details dC ON dC.MPR_ID = q.MPR_ID
                        AND TRY_CAST(TRY_CAST(dC.Item_No AS DECIMAL(10,2)) AS INT) > 0
                    INNER JOIN MPR_Header hS
                        ON SUBSTRING(hS.MPR_No, 1, CHARINDEX('_Rev.', hS.MPR_No + '_Rev.') - 1)
                         = SUBSTRING(q.MPR_No,  1, CHARINDEX('_Rev.', q.MPR_No  + '_Rev.') - 1)
                    INNER JOIN MPR_Details dS ON dS.MPR_ID = hS.MPR_ID
                        AND TRY_CAST(TRY_CAST(dS.Item_No AS DECIMAL(10,2)) AS INT)
                          = TRY_CAST(TRY_CAST(dC.Item_No AS DECIMAL(10,2)) AS INT)
                    WHERE q.rn = 1";

                // Bước 2: Main query dùng #SibDet (temp table) — mỗi flat CTE scan 1 lần duy nhất
                const string SQL = @"
                    WITH
                    FilteredMPR AS (
                        SELECT *,
                               ROW_NUMBER() OVER (
                                   PARTITION BY Project_Code,
                                       SUBSTRING(MPR_No, 1, CHARINDEX('_Rev.', MPR_No + '_Rev.') - 1)
                                   ORDER BY TRY_CAST(TRY_CAST(Rev AS DECIMAL(10,2)) AS INT) DESC, MPR_ID DESC
                               ) AS rn
                        FROM MPR_Header
                        WHERE ISNULL(Status, '') <> N'Hủy'
                    ),
                    PO_Flat AS (
                        SELECT DISTINCT sd.CurrID, ph.PONo
                        FROM #SibDet sd
                        INNER JOIN PO_Detail pod ON pod.MPR_Detail_ID = sd.SibID
                        INNER JOIN PO_head   ph  ON ph.PO_ID = pod.PO_ID
                        WHERE ISNULL(ph.Status,'') <> 'Cancelled'
                    ),
                    Heat_Flat AS (
                        SELECT DISTINCT sd.CurrID, rd.Heatno
                        FROM #SibDet sd
                        INNER JOIN PO_Detail  pod ON pod.MPR_Detail_ID = sd.SibID
                        INNER JOIN PO_head    ph  ON ph.PO_ID  = pod.PO_ID
                        INNER JOIN RIR_head   rh  ON rh.PONo   = ph.PONo
                        INNER JOIN RIR_detail rd  ON rd.RIR_ID = rh.RIR_ID
                        WHERE ISNULL(rd.Heatno,N'') != N''
                    ),
                    KT_Flat AS (
                        SELECT sd.CurrID,
                               MIN(CASE rd.Inspect_Result
                                   WHEN N'Fail' THEN 1 WHEN N'Hold' THEN 2
                                   WHEN N'Pass' THEN 3 ELSE 4 END) AS KT_Rank
                        FROM #SibDet sd
                        INNER JOIN PO_Detail  pod ON pod.MPR_Detail_ID = sd.SibID
                        INNER JOIN PO_head    ph  ON ph.PO_ID  = pod.PO_ID
                        INNER JOIN RIR_head   rh  ON rh.PONo   = ph.PONo
                        INNER JOIN RIR_detail rd  ON rd.RIR_ID = rh.RIR_ID
                        GROUP BY sd.CurrID
                    ),
                    RIR_Flat AS (
                        SELECT DISTINCT sd.CurrID, rh.RIR_No,
                                        rh.Status AS RIR_Status, rh.Issue_Date
                        FROM #SibDet sd
                        INNER JOIN PO_Detail pod ON pod.MPR_Detail_ID = sd.SibID
                        INNER JOIN PO_head   ph  ON ph.PO_ID = pod.PO_ID
                        INNER JOIN RIR_head  rh  ON rh.PONo  = ph.PONo
                        WHERE ISNULL(rh.RIR_No,N'') != N''
                    ),
                    cte_PO AS (
                        SELECT pf.CurrID AS CurrDetailID,
                               ISNULL(STUFF((
                                   SELECT DISTINCT N', ' + pf2.PONo
                                   FROM PO_Flat pf2
                                   WHERE pf2.CurrID = pf.CurrID
                                   FOR XML PATH(''), TYPE).value('.','NVARCHAR(MAX)'),1,2,N''),N'') AS POList
                        FROM PO_Flat pf
                        GROUP BY pf.CurrID
                    ),
                    cte_Heat AS (
                        SELECT hf.CurrID AS CurrDetailID,
                               ISNULL(STUFF((
                                   SELECT DISTINCT N', ' + hf2.Heatno
                                   FROM Heat_Flat hf2
                                   WHERE hf2.CurrID = hf.CurrID
                                   FOR XML PATH(''), TYPE).value('.','NVARCHAR(MAX)'),1,2,N''),N'') AS HeatList
                        FROM Heat_Flat hf
                        GROUP BY hf.CurrID
                    ),
                    cte_KT AS (
                        SELECT CurrID AS CurrDetailID, KT_Rank
                        FROM KT_Flat
                    ),
                    cte_RIR AS (
                        SELECT rf.CurrID AS CurrDetailID,
                               ISNULL(STUFF((
                                   SELECT DISTINCT N', ' + rf2.RIR_No
                                   FROM RIR_Flat rf2
                                   WHERE rf2.CurrID = rf.CurrID
                                   FOR XML PATH(''), TYPE).value('.','NVARCHAR(MAX)'),1,2,N''),N'') AS RIRList,
                               (SELECT TOP 1 rf3.RIR_Status
                                FROM RIR_Flat rf3
                                WHERE rf3.CurrID = rf.CurrID
                                  AND ISNULL(rf3.RIR_Status,N'') != N''
                                ORDER BY rf3.Issue_Date DESC) AS RIR_Status
                        FROM RIR_Flat rf
                        GROUP BY rf.CurrID
                    )
                    SELECT
                        ISNULL(pi.ProjectCode, N'')                         AS [Mã dự án],
                        h.MPR_No                                            AS [MPR No],
                        h.Rev                                               AS [Rev],
                        d.Item_No                                           AS [Item No],
                        d.item_name                                         AS [Tên vật tư],
                        ISNULL(d.Description, N'')                          AS [Mô tả],
                        d.Material                                          AS [Vật liệu],
                        ISNULL(CAST(NULLIF(d.Thickness_mm,'') AS NVARCHAR),N'') AS [A-Dày(mm)],
                        ISNULL(CAST(NULLIF(d.Depth_mm,    '') AS NVARCHAR),N'') AS [B-Sâu(mm)],
                        ISNULL(CAST(NULLIF(d.C_Width_mm,  '') AS NVARCHAR),N'') AS [C-Rộng(mm)],
                        ISNULL(CAST(NULLIF(d.D_Web_mm,    '') AS NVARCHAR),N'') AS [D-Bụng(mm)],
                        ISNULL(CAST(NULLIF(d.E_Flange_mm, '') AS NVARCHAR),N'') AS [E-Cánh(mm)],
                        ISNULL(CAST(NULLIF(d.F_Length_mm, '') AS NVARCHAR),N'') AS [F-Dài(mm)],
                        ISNULL(d.UNIT,       N'')                           AS [ĐVT],
                        d.Qty_Per_Sheet                                     AS [SL],
                        ISNULL(d.Weight_kg,  0)                             AS [KG],
                        ISNULL(cp.POList,    N'')                           AS [PO No],
                        ISNULL(ch.HeatList,  N'')                           AS [Heat No],
                        ISNULL(CASE ck.KT_Rank
                            WHEN 1 THEN N'Fail' WHEN 2 THEN N'Hold'
                            WHEN 3 THEN N'Pass' ELSE N'Chưa KT' END, N'Chưa KT') AS [Kết quả KT],
                        ISNULL(cr.RIRList,   N'')                           AS [RIR No],
                        ISNULL(cr.RIR_Status,N'')                           AS [Trạng thái RIR]
                    FROM FilteredMPR h
                    INNER JOIN MPR_Details d  ON d.MPR_ID       = h.MPR_ID
                    LEFT  JOIN ProjectInfo pi ON pi.ProjectCode = h.Project_Code
                    LEFT  JOIN cte_PO   cp ON cp.CurrDetailID   = d.Detail_ID
                    LEFT  JOIN cte_Heat ch ON ch.CurrDetailID   = d.Detail_ID
                    LEFT  JOIN cte_KT   ck ON ck.CurrDetailID   = d.Detail_ID
                    LEFT  JOIN cte_RIR  cr ON cr.CurrDetailID   = d.Detail_ID
                    WHERE h.rn = 1
                    ORDER BY pi.ProjectCode, h.MPR_No, d.Item_No";


                DataTable dtFull;
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    // Bước 1: tạo temp table #SibDet (tính 1 lần)
                    new SqlCommand(SQL_SIB, conn) { CommandTimeout = 120 }.ExecuteNonQuery();
                    // Bước 2: main query dùng #SibDet
                    dtFull = new DataTable();
                    dtFull.Load(new SqlCommand(SQL, conn) { CommandTimeout = 120 }.ExecuteReader());
                }

                // ── Popup ──
                var popup = new Form
                {
                    Text = "🔎 Check All Items — Tổng hợp toàn bộ vật tư MPR",
                    Size = new Size(1400, 700),
                    StartPosition = FormStartPosition.CenterParent,
                    BackColor = Color.FromArgb(245, 245, 245),
                    MinimumSize = new Size(1100, 500),
                    KeyPreview = true
                };

                popup.Controls.Add(new Label
                {
                    Text = "🔎  CHECK ALL ITEMS — Tổng hợp vật tư tất cả MPR & kết quả kiểm tra RIR  |  💡 Double click → mở MPR",
                    Font = new Font("Segoe UI", 10, FontStyle.Bold),
                    ForeColor = Color.FromArgb(102, 51, 153),
                    Location = new Point(10, 8),
                    Size = new Size(900, 24)
                });

                int total = dtFull.Rows.Count;
                int pass = 0, fail = 0, hold = 0, notYet = 0;
                foreach (DataRow r in dtFull.Rows)
                {
                    string kt = r["Kết quả KT"]?.ToString() ?? "";
                    if (kt == "Pass") pass++;
                    else if (kt == "Fail") fail++;
                    else if (kt == "Hold") hold++;
                    else notYet++;
                }
                var lblStat = new Label
                {
                    Text = $"Tổng: {total}  |  ✅ Pass: {pass}  |  ❌ Fail: {fail}  |  ⏸ Hold: {hold}  |  ⏳ Chưa KT: {notYet}",
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    ForeColor = Color.FromArgb(0, 120, 212),
                    Location = new Point(10, 36),
                    Size = new Size(1360, 20),
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
                };
                popup.Controls.Add(lblStat);

                var pFilter = new Panel
                {
                    Location = new Point(10, 62),
                    Size = new Size(popup.ClientSize.Width - 20, 72),
                    BackColor = Color.White,
                    BorderStyle = BorderStyle.FixedSingle,
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
                };
                popup.Controls.Add(pFilter);

                Action<string, int, int, int> addFL = (txt, x, y2, w) =>
                    pFilter.Controls.Add(new Label
                    {
                        Text = txt,
                        Location = new Point(x, y2 + 3),
                        Size = new Size(w, 18),
                        Font = new Font("Segoe UI", 8, FontStyle.Bold),
                        ForeColor = Color.FromArgb(60, 60, 60)
                    });

                int row1Y = 6;
                int x1 = 6;

                addFL("Mã DA:", x1, row1Y, 48);

                // ── Checked-combo dropdown cho Mã DA (chọn nhiều) ──
                var daProjectList = dtFull.AsEnumerable()
                    .Select(r => r["Mã dự án"].ToString())
                    .Where(v => !string.IsNullOrEmpty(v))
                    .Distinct().OrderBy(v => v).ToList();

                var clbDA = new CheckedListBox
                {
                    Font = new Font("Segoe UI", 9),
                    CheckOnClick = true,
                    BorderStyle = BorderStyle.None,
                    BackColor = Color.White,
                    IntegralHeight = false,
                    Width = 220
                };
                foreach (var p in daProjectList) clbDA.Items.Add(p, false);
                clbDA.Height = Math.Min(clbDA.Items.Count, 10) * 17 + 4;

                var panelDropDA = new Panel
                {
                    Size = new Size(224, clbDA.Height + 2),
                    BackColor = Color.White,
                    BorderStyle = BorderStyle.FixedSingle,
                    Visible = false
                };
                popup.Controls.Add(panelDropDA);
                panelDropDA.Controls.Add(clbDA);
                panelDropDA.BringToFront();

                var btnDropDA = new Button
                {
                    Location = new Point(x1 + 50, row1Y),
                    Size = new Size(120, 22),
                    Text = "(Tất cả)  ▼",
                    TextAlign = ContentAlignment.MiddleLeft,
                    Font = new Font("Segoe UI", 8),
                    BackColor = Color.White,
                    ForeColor = Color.FromArgb(30, 30, 30),
                    FlatStyle = FlatStyle.Flat,
                    Cursor = Cursors.Hand,
                    Padding = new Padding(3, 0, 0, 0)
                };
                btnDropDA.FlatAppearance.BorderColor = Color.FromArgb(171, 173, 179);
                btnDropDA.FlatAppearance.BorderSize = 1;
                btnDropDA.Paint += (s, ev) =>
                {
                    int ax = btnDropDA.Width - 14, ay = btnDropDA.Height / 2;
                    ev.Graphics.FillPolygon(Brushes.DimGray, new[]
                    {
                        new Point(ax, ay - 3), new Point(ax + 7, ay - 3), new Point(ax + 3, ay + 3)
                    });
                };
                pFilter.Controls.Add(btnDropDA);

                Action updateBtnDA = () =>
                {
                    var sel = clbDA.CheckedItems.Cast<string>().ToList();
                    btnDropDA.Text = sel.Count == 0 ? "(Tất cả)  ▼" :
                                     sel.Count == 1 ? sel[0] + "  ▼" :
                                     $"({sel.Count} DA)  ▼";
                    btnDropDA.ForeColor = sel.Count > 0 ? Color.FromArgb(102, 51, 153) : Color.FromArgb(30, 30, 30);
                    btnDropDA.Font = new Font("Segoe UI", 8, sel.Count > 0 ? FontStyle.Bold : FontStyle.Regular);
                };

                btnDropDA.Click += (s, ev) =>
                {
                    if (panelDropDA.Visible) { panelDropDA.Visible = false; return; }
                    var pt = popup.PointToClient(btnDropDA.Parent.PointToScreen(
                        new Point(btnDropDA.Left, btnDropDA.Bottom + 2)));
                    panelDropDA.Location = pt;
                    panelDropDA.BringToFront();
                    panelDropDA.Visible = true;
                    clbDA.Focus();
                };

                popup.MouseDown += (s, ev) =>
                {
                    if (panelDropDA.Visible && !panelDropDA.Bounds.Contains(ev.Location))
                        panelDropDA.Visible = false;
                };

                x1 += 180;

                addFL("Tên VT:", x1, row1Y, 50);
                var txtFName = new TextBox { Location = new Point(x1 + 52, row1Y), Size = new Size(140, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "Tên vật tư..." };
                pFilter.Controls.Add(txtFName);
                x1 += 202;

                addFL("A(mm):", x1, row1Y, 48);
                var txtFA = new TextBox { Location = new Point(x1 + 50, row1Y), Size = new Size(44, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "A" };
                pFilter.Controls.Add(txtFA); x1 += 102;

                addFL("B(mm):", x1, row1Y, 48);
                var txtFB = new TextBox { Location = new Point(x1 + 50, row1Y), Size = new Size(44, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "B" };
                pFilter.Controls.Add(txtFB); x1 += 102;

                addFL("C(mm):", x1, row1Y, 48);
                var txtFC = new TextBox { Location = new Point(x1 + 50, row1Y), Size = new Size(44, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "C" };
                pFilter.Controls.Add(txtFC); x1 += 102;

                addFL("D(mm):", x1, row1Y, 48);
                var txtFD = new TextBox { Location = new Point(x1 + 50, row1Y), Size = new Size(44, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "D" };
                pFilter.Controls.Add(txtFD); x1 += 102;

                addFL("E(mm):", x1, row1Y, 48);
                var txtFE = new TextBox { Location = new Point(x1 + 50, row1Y), Size = new Size(44, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "E" };
                pFilter.Controls.Add(txtFE); x1 += 102;

                addFL("F(mm):", x1, row1Y, 48);
                var txtFF = new TextBox { Location = new Point(x1 + 50, row1Y), Size = new Size(44, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "F" };
                pFilter.Controls.Add(txtFF);

                int row2Y = 38;
                int x2 = 6;

                addFL("Heat No:", x2, row2Y, 55);
                var txtFHeat = new TextBox { Location = new Point(x2 + 57, row2Y), Size = new Size(100, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "Heat No..." };
                pFilter.Controls.Add(txtFHeat);
                x2 += 167;

                addFL("KQ KT:", x2, row2Y, 48);
                var cboFKQ = new ComboBox
                {
                    Location = new Point(x2 + 50, row2Y),
                    Size = new Size(100, 22),
                    Font = new Font("Segoe UI", 9),
                    DropDownStyle = ComboBoxStyle.DropDownList
                };
                cboFKQ.Items.AddRange(new[] { "Tất cả", "Pass", "Fail", "Hold", "Chưa KT" });
                cboFKQ.SelectedIndex = 0;
                pFilter.Controls.Add(cboFKQ);
                x2 += 160;

                addFL("Vật liệu:", x2, row2Y, 58);
                var txtFMat = new TextBox
                {
                    Location = new Point(x2 + 60, row2Y),
                    Size = new Size(110, 22),
                    Font = new Font("Segoe UI", 9),
                    PlaceholderText = "Vật liệu..."
                };
                pFilter.Controls.Add(txtFMat);
                x2 += 178;

                var btnFSearch = new Button
                {
                    Text = "🔍 Lọc",
                    Location = new Point(x2, row2Y - 2),
                    Size = new Size(75, 26),
                    BackColor = Color.FromArgb(0, 120, 212),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 8, FontStyle.Bold)
                };
                btnFSearch.FlatAppearance.BorderSize = 0;
                pFilter.Controls.Add(btnFSearch);

                var btnFClear = new Button
                {
                    Text = "✖ Xóa lọc",
                    Location = new Point(x2 + 79, row2Y - 2),
                    Size = new Size(85, 26),
                    BackColor = Color.FromArgb(108, 117, 125),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 8, FontStyle.Bold)
                };
                btnFClear.FlatAppearance.BorderSize = 0;
                pFilter.Controls.Add(btnFClear);

                foreach (Control c in pFilter.Controls)
                    if (c is TextBox || c is ComboBox || c is Button) c.BringToFront();
                // Dropdown panel phải luôn nằm trên cùng khi hiện
                panelDropDA.BringToFront();

                var dgv = new DataGridView
                {
                    Location = new Point(10, 142),
                    Size = new Size(popup.ClientSize.Width - 20, popup.ClientSize.Height - 190),
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
                dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(102, 51, 153);
                dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                dgv.EnableHeadersVisualStyles = false;
                dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 240, 255);
                dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
                dgv.DefaultCellStyle.SelectionForeColor = Color.Black;
                popup.Controls.Add(dgv);

                dgv.CellFormatting += (s, ev) =>
                {
                    if (ev.RowIndex < 0) return;
                    string col = dgv.Columns[ev.ColumnIndex].Name;
                    if (col == "Kết quả KT")
                    {
                        string v = ev.Value?.ToString() ?? "";
                        ev.CellStyle.ForeColor =
                            v == "Pass" ? Color.FromArgb(40, 167, 69) :
                            v == "Fail" ? Color.FromArgb(220, 53, 69) :
                            v == "Hold" ? Color.FromArgb(255, 140, 0) :
                            Color.FromArgb(108, 117, 125);
                        ev.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    }
                    if (col == "Trạng thái RIR")
                    {
                        string v = ev.Value?.ToString() ?? "";
                        ev.CellStyle.ForeColor =
                            v == "Hoàn thành" ? Color.FromArgb(40, 167, 69) :
                            v == "Đang kiểm tra" ? Color.FromArgb(255, 140, 0) :
                            string.IsNullOrEmpty(v) ? Color.FromArgb(180, 180, 180) :
                            Color.FromArgb(0, 120, 212);
                        ev.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    }
                };

                dgv.RowPrePaint += (s, ev) =>
                {
                    if (ev.RowIndex < 0 || dgv.Rows[ev.RowIndex].IsNewRow) return;
                    string kt = dgv.Rows[ev.RowIndex].Cells["Kết quả KT"].Value?.ToString() ?? "";
                    dgv.Rows[ev.RowIndex].DefaultCellStyle.BackColor =
                        kt == "Pass" ? Color.FromArgb(235, 255, 235) :
                        kt == "Fail" ? Color.FromArgb(255, 235, 235) :
                        kt == "Hold" ? Color.FromArgb(255, 248, 230) :
                        Color.White;
                };

                Action applyFilter = () =>
                {
                    var selDA = clbDA.CheckedItems.Cast<string>().ToList();
                    string kName = txtFName.Text.Trim().ToLower();
                    string kA = txtFA.Text.Trim();
                    string kB = txtFB.Text.Trim();
                    string kC = txtFC.Text.Trim();
                    string kD = txtFD.Text.Trim();
                    string kE = txtFE.Text.Trim();
                    string kF = txtFF.Text.Trim();
                    string kHeat = txtFHeat.Text.Trim().ToLower();
                    string kKQ = cboFKQ.SelectedItem?.ToString() ?? "Tất cả";
                    string kMat = txtFMat.Text.Trim().ToLower();

                    var rows = dtFull.AsEnumerable().Where(r =>
                    {
                        if (selDA.Count > 0 && !selDA.Contains(r["Mã dự án"].ToString())) return false;
                        if (!string.IsNullOrEmpty(kName) && !r["Tên vật tư"].ToString().ToLower().Contains(kName)) return false;
                        if (!string.IsNullOrEmpty(kMat) && !r["Vật liệu"].ToString().ToLower().Contains(kMat)) return false;
                        if (!string.IsNullOrEmpty(kA) && !r["A-Dày(mm)"].ToString().Contains(kA)) return false;
                        if (!string.IsNullOrEmpty(kB) && !r["B-Sâu(mm)"].ToString().Contains(kB)) return false;
                        if (!string.IsNullOrEmpty(kC) && !r["C-Rộng(mm)"].ToString().Contains(kC)) return false;
                        if (!string.IsNullOrEmpty(kD) && !r["D-Bụng(mm)"].ToString().Contains(kD)) return false;
                        if (!string.IsNullOrEmpty(kE) && !r["E-Cánh(mm)"].ToString().Contains(kE)) return false;
                        if (!string.IsNullOrEmpty(kF) && !r["F-Dài(mm)"].ToString().Contains(kF)) return false;
                        if (!string.IsNullOrEmpty(kHeat) && !r["Heat No"].ToString().ToLower().Contains(kHeat)) return false;
                        if (kKQ != "Tất cả" && r["Kết quả KT"].ToString() != kKQ) return false;
                        return true;
                    });

                    DataTable dtView = rows.Any() ? rows.CopyToDataTable() : dtFull.Clone();
                    dgv.DataSource = dtView;

                    foreach (string colName in new[] { "MPR No", "PO No" })
                    {
                        if (!dgv.Columns.Contains(colName)) continue;
                        var col = dgv.Columns[colName];
                        col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                        int measuredW = col.GetPreferredWidth(DataGridViewAutoSizeColumnMode.AllCells, true);
                        col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                        col.Width = Math.Min(measuredW, 400);
                    }

                    int t = dtView.Rows.Count, p = 0, f = 0, h2 = 0, n = 0;
                    foreach (DataRow r in dtView.Rows)
                    {
                        string kt = r["Kết quả KT"]?.ToString() ?? "";
                        if (kt == "Pass") p++;
                        else if (kt == "Fail") f++;
                        else if (kt == "Hold") h2++;
                        else n++;
                    }
                    lblStat.Text = $"Hiển thị: {t}  |  ✅ Pass: {p}  |  ❌ Fail: {f}  |  ⏸ Hold: {h2}  |  ⏳ Chưa KT: {n}";
                };

                dgv.DataSource = null;
                lblStat.Text = "Nhấn [🔍 Lọc] để tìm kiếm dữ liệu.";

                dgv.DataBindingComplete += (s, ev) =>
                {
                    foreach (string colName in new[] { "MPR No", "PO No" })
                    {
                        if (!dgv.Columns.Contains(colName)) continue;
                        var col = dgv.Columns[colName];
                        col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                        int measuredW = col.GetPreferredWidth(DataGridViewAutoSizeColumnMode.AllCells, true);
                        col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                        col.Width = Math.Min(measuredW, 400);
                    }
                };

                // ItemCheck: update button text + refilter (after applyFilter to avoid CS0841)
                clbDA.ItemCheck += (s, ev) =>
                {
                    clbDA.BeginInvoke(new Action(() =>
                    {
                        updateBtnDA();
                        applyFilter();
                    }));
                };

                btnFSearch.Click += (s, ev) => applyFilter();
                btnFClear.Click += (s, ev) =>
                {
                    for (int i = 0; i < clbDA.Items.Count; i++) clbDA.SetItemChecked(i, false);
                    updateBtnDA();
                    panelDropDA.Visible = false;
                    txtFName.Text = "";
                    txtFMat.Text = "";
                    txtFA.Text = ""; txtFB.Text = ""; txtFC.Text = "";
                    txtFD.Text = ""; txtFE.Text = ""; txtFF.Text = "";
                    txtFHeat.Text = ""; cboFKQ.SelectedIndex = 0;
                    dgv.DataSource = null;
                    lblStat.Text = "Nhấn [🔍 Lọc] để tìm kiếm dữ liệu.";
                };

                dgv.CellDoubleClick += (s, ev) =>
                {
                    if (ev.RowIndex < 0) return;
                    string mprNo = dgv.Rows[ev.RowIndex].Cells["MPR No"].Value?.ToString() ?? "";
                    if (string.IsNullOrEmpty(mprNo)) return;

                    var target = _mprList.Find(m => m.MPR_No == mprNo);
                    if (target == null)
                    {
                        MessageBox.Show(TopOwner, $"Không tìm thấy MPR: {mprNo}", "Thông báo",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    popup.Close();
                    SelectMPRById(target.MPR_ID);
                };

                var ttip = new ToolTip();
                ttip.SetToolTip(dgv, "Double click vào dòng để mở MPR tương ứng");

                popup.KeyDown += (s, ev) =>
                {
                    if (ev.KeyCode == Keys.Enter)
                    {
                        applyFilter();
                        ev.Handled = true;
                        ev.SuppressKeyPress = true;
                    }
                    if (ev.KeyCode == Keys.Escape) popup.Close();
                };

                var btnExport = new Button
                {
                    Text = "📥 Xuất Excel",
                    Size = new Size(120, 30),
                    BackColor = Color.FromArgb(0, 150, 100),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Right
                };
                btnExport.FlatAppearance.BorderSize = 0;
                btnExport.Location = new Point(popup.ClientSize.Width - 245, popup.ClientSize.Height - 40);
                popup.Controls.Add(btnExport);

                btnExport.Click += (s, ev) =>
                {
                    var dtExport = dgv.DataSource as DataTable;
                    if (dtExport == null || dtExport.Rows.Count == 0)
                    {
                        MessageBox.Show(TopOwner, "Không có dữ liệu để xuất!", "Thông báo",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    var sfd = new SaveFileDialog
                    {
                        Title = "Lưu file Excel",
                        Filter = "Excel Files|*.xlsx",
                        FileName = $"CheckAllItems_{DateTime.Now:yyyyMMdd_HHmm}",
                        InitialDirectory = Directory.Exists(@"D:\RÁC") ? @"D:\RÁC" : Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                    };
                    if (sfd.ShowDialog() != DialogResult.OK) return;

                    try
                    {
                        OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                        using (var pkg = new OfficeOpenXml.ExcelPackage())
                        {
                            var ws = pkg.Workbook.Worksheets.Add("Check All Items");

                            ws.Cells[1, 1].Value = "CHECK ALL ITEMS — Tổng hợp vật tư MPR & kết quả kiểm tra RIR";
                            ws.Cells[1, 1, 1, dtExport.Columns.Count].Merge = true;
                            ws.Cells[1, 1].Style.Font.Size = 13;
                            ws.Cells[1, 1].Style.Font.Bold = true;
                            ws.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            ws.Cells[1, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(102, 51, 153));
                            ws.Cells[1, 1].Style.Font.Color.SetColor(Color.White);

                            ws.Cells[2, 1].Value = $"Xuất ngày: {DateTime.Now:dd/MM/yyyy HH:mm}  |  Tổng: {dtExport.Rows.Count} dòng";
                            ws.Cells[2, 1, 2, dtExport.Columns.Count].Merge = true;
                            ws.Cells[2, 1].Style.Font.Italic = true;
                            ws.Cells[2, 1].Style.Font.Size = 9;

                            for (int c = 0; c < dtExport.Columns.Count; c++)
                            {
                                var cell = ws.Cells[3, c + 1];
                                cell.Value = dtExport.Columns[c].ColumnName;
                                cell.Style.Font.Bold = true;
                                cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(102, 51, 153));
                                cell.Style.Font.Color.SetColor(Color.White);
                                cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            }

                            for (int row = 0; row < dtExport.Rows.Count; row++)
                            {
                                var dr = dtExport.Rows[row];
                                bool isAlt = row % 2 == 1;

                                for (int c = 0; c < dtExport.Columns.Count; c++)
                                {
                                    var cell = ws.Cells[row + 4, c + 1];
                                    string colName = dtExport.Columns[c].ColumnName;

                                    if (colName == "KG")
                                    {
                                        if (dr[c] != DBNull.Value && decimal.TryParse(dr[c]?.ToString(), out decimal kg))
                                        {
                                            cell.Value = Math.Round(kg, 2);
                                            cell.Style.Numberformat.Format = "#,##0.00";
                                            cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                                        }
                                        else
                                        {
                                            cell.Value = 0.00m;
                                            cell.Style.Numberformat.Format = "#,##0.00";
                                        }
                                    }
                                    else
                                    {
                                        cell.Value = dr[c]?.ToString() ?? "";
                                    }

                                    cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                    cell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                                    if (isAlt)
                                    {
                                        cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                        cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(248, 240, 255));
                                    }

                                    if (colName == "Kết quả KT")
                                    {
                                        string kt = dr[c]?.ToString() ?? "";
                                        cell.Style.Font.Bold = true;
                                        cell.Style.Font.Color.SetColor(
                                            kt == "Pass" ? Color.FromArgb(40, 167, 69) :
                                            kt == "Fail" ? Color.FromArgb(220, 53, 69) :
                                            kt == "Hold" ? Color.FromArgb(255, 140, 0) :
                                            Color.FromArgb(108, 117, 125));
                                    }
                                }
                            }

                            ws.Cells[ws.Dimension.Address].AutoFitColumns();
                            ws.View.FreezePanes(4, 1);

                            pkg.SaveAs(new System.IO.FileInfo(sfd.FileName));
                        }

                        var res = MessageBox.Show(TopOwner,
                            $"✅ Xuất Excel thành công!\nFile: {sfd.FileName}\n\nBạn có muốn mở file không?",
                            "Thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                        if (res == DialogResult.Yes)
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                            { FileName = sfd.FileName, UseShellExecute = true });
                    }
                    catch (Exception exExport)
                    {
                        MessageBox.Show(TopOwner, "Lỗi xuất Excel: " + exExport.Message, "Lỗi",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                };

                var btnClose = new Button
                {
                    Text = "Đóng",
                    Size = new Size(100, 30),
                    BackColor = Color.FromArgb(108, 117, 125),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                    DialogResult = DialogResult.OK
                };
                btnClose.FlatAppearance.BorderSize = 0;
                btnClose.Location = new Point(popup.ClientSize.Width - 115, popup.ClientSize.Height - 40);
                popup.Controls.Add(btnClose);
                popup.AcceptButton = btnFSearch;
                popup.CancelButton = btnClose;

                popup.Resize += (s, ev) =>
                {
                    pFilter.Width = popup.ClientSize.Width - 20;
                    btnExport.Location = new Point(popup.ClientSize.Width - 245, popup.ClientSize.Height - 40);
                    btnClose.Location = new Point(popup.ClientSize.Width - 115, popup.ClientSize.Height - 40);
                };

                popup.Owner = this;
                popup.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(TopOwner, "Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ===== HELPERS =====
        private decimal DecimalVal(object val)
        {
            if (val == null || val == DBNull.Value) return 0;
            return decimal.TryParse(val.ToString(), out decimal d) ? d : 0;
        }

        private void ClearHeader()
        {
            txtMPRNo.Text = "";
            txtProjectName.Text = "";
            txtProjectCode.Text = "";
            txtDepartment.Text = "";
            txtRequestor.Text = "";
            txtRev.Text = "0";
            txtNotes.Text = "";
            dtpRequiredDate.Value = DateTime.Today;
            cboStatus.SelectedIndex = 0;
        }

        // =====================================================
        //  ÁP DỤNG PHÂN QUYỀN
        // =====================================================
        private void ApplyPermissions()
        {
            if (btnNewMPR != null) PermissionHelper.Apply(btnNewMPR, "MPR", "Tạo MPR");
            // btnCreateMPR dùng cùng quyền "Tạo MPR"
            if (btnDeleteMPR != null) PermissionHelper.Apply(btnDeleteMPR, "MPR", "Xóa MPR");
            if (btnSaveHeader != null) PermissionHelper.Apply(btnSaveHeader, "MPR", "Lưu Header");
            if (btnAddDetail != null) PermissionHelper.Apply(btnAddDetail, "MPR", "Thêm dòng");
            if (btnDeleteDetail != null) PermissionHelper.Apply(btnDeleteDetail, "MPR", "Xóa dòng");
            if (btnSaveDetail != null) PermissionHelper.Apply(btnSaveDetail, "MPR", "Lưu chi tiết");
            foreach (var c in this.Controls.Find("btnCreatePO", true))
                PermissionHelper.Apply(c, "MPR", "Tạo PO");
            foreach (var c in this.Controls.Find("btnCheckAll", true))
                PermissionHelper.Apply(c, "MPR", "Check All Items");
        }

        // =====================================================================
        //  LOC DETAIL THEO DA LEN PO
        // =====================================================================
        // Load dong cac gia tri PO_No vao combobox filter
        private void RefreshPOFilterCombo()
        {
            if (_cboFilterPO == null) return;
            _cboFilterPO.SelectedIndexChanged -= (s, ev) => FilterDetailByPO();
            _cboFilterPO.Items.Clear();
            _cboFilterPO.Items.Add("(Tat ca)"); // mac dinh

            // Lay cac gia tri duy nhat tu cot PO_No
            var seen = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
            bool hasEmpty = false;
            foreach (DataGridViewRow row in dgvDetails.Rows)
            {
                if (row.IsNewRow) continue;
                string val = row.Cells["PO_No"]?.Value?.ToString() ?? "";
                if (string.IsNullOrWhiteSpace(val)) hasEmpty = true;
                else if (seen.Add(val)) _cboFilterPO.Items.Add(val);
            }
            // Them muc loc rong neu co dong chua len PO
            if (hasEmpty) _cboFilterPO.Items.Add("(Chua len PO)");
            _cboFilterPO.SelectedIndex = 0; // chon "(Tat ca)"
            _cboFilterPO.SelectedIndexChanged += (s, ev) => FilterDetailByPO();
        }

        private void FilterDetailByPO()
        {
            if (_cboFilterPO == null || dgvDetails == null) return;
            string sel = _cboFilterPO.SelectedItem?.ToString() ?? "(Tat ca)";

            int visibleCount = 0;
            foreach (DataGridViewRow row in dgvDetails.Rows)
            {
                if (row.IsNewRow) continue;
                string poVal = row.Cells["PO_No"]?.Value?.ToString() ?? "";
                bool show;
                if (sel == "(Tat ca)")
                    show = true;
                else if (sel == "(Chua len PO)")
                    show = string.IsNullOrWhiteSpace(poVal);
                else
                    show = string.Equals(poVal, sel, StringComparison.OrdinalIgnoreCase);
                row.Visible = show;
                if (show) visibleCount++;
            }

            if (visibleCount == 0 && sel != "(Tat ca)")
                SafeMsg("Không tìm thấy kết quả phù hợp!", "Tìm kiếm", MessageBoxIcon.Information);
        }

        // =====================================================================
        //  MPR STATUS POPUP
        // =====================================================================
        private void BtnMPRStatus_Click(object sender, EventArgs e)
        {
            ShowMPRStatusPopup();
        }

        private void ShowMPRStatusPopup()
        {
            var popup = new Form
            {
                Text = "Theo dõi tiến độ xử lý MPR",
                Size = new Size(1100, 680),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.White,
                Font = new Font("Segoe UI", 9)
            };

            // ── Toolbar ──────────────────────────────────────────────────────
            var pnlTop = new Panel { Dock = DockStyle.Top, Height = 50, BackColor = Color.FromArgb(45, 45, 48), Padding = new Padding(8, 8, 8, 4) };
            popup.Controls.Add(pnlTop);

            // Project filter
            var lblProj = new Label { Text = "Dự án:", ForeColor = Color.White, AutoSize = true, Location = new Point(8, 16) };
            var cboProject = new ComboBox { Location = new Point(55, 12), Size = new Size(200, 24), DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 9) };

            // MPR No filter
            var lblMPR = new Label { Text = "Số MPR:", ForeColor = Color.White, AutoSize = true, Location = new Point(270, 16) };
            var txtMPRNo = new TextBox { Location = new Point(322, 12), Size = new Size(180, 24), Font = new Font("Segoe UI", 9) };

            // Search button
            var btnSearch = new Button
            {
                Text = "🔍 Tìm kiếm",
                Location = new Point(516, 10),
                Size = new Size(100, 28),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnSearch.FlatAppearance.BorderSize = 0;

            pnlTop.Controls.AddRange(new Control[] { lblProj, cboProject, lblMPR, txtMPRNo, btnSearch });

            // ── Status filter buttons ────────────────────────────────────────
            var pnlFilter = new Panel { Dock = DockStyle.Top, Height = 44, BackColor = Color.FromArgb(240, 240, 245), Padding = new Padding(8, 6, 8, 4) };
            popup.Controls.Add(pnlFilter);

            string[] filterLabels = { "Tất cả", "Chưa gửi báo giá", "Đang chờ báo giá", "Đã gửi PO" };
            Color[] filterColors = {
                Color.FromArgb(100, 100, 100),
                Color.FromArgb(180, 50, 50),
                Color.FromArgb(200, 140, 0),
                Color.FromArgb(0, 130, 60)
            };
            var filterBtns = new Button[4];
            int fx = 8;
            for (int i = 0; i < filterLabels.Length; i++)
            {
                var b = new Button
                {
                    Text = filterLabels[i],
                    Location = new Point(fx, 7),
                    Size = new Size(150, 30),
                    BackColor = filterColors[i],
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    Tag = filterLabels[i],
                    Cursor = Cursors.Hand
                };
                b.FlatAppearance.BorderSize = 0;
                pnlFilter.Controls.Add(b);
                filterBtns[i] = b;
                fx += 158;
            }

            // Summary label
            var lblSummary = new Label
            {
                Text = "",
                Location = new Point(fx + 10, 12),
                Size = new Size(400, 20),
                ForeColor = Color.FromArgb(60, 60, 60),
                Font = new Font("Segoe UI", 9)
            };
            pnlFilter.Controls.Add(lblSummary);

            // ── DataGridView ─────────────────────────────────────────────────
            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgv.EnableHeadersVisualStyles = false;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
            dgv.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

            var pnlGrid = new Panel { Dock = DockStyle.Fill };
            pnlGrid.Controls.Add(dgv);
            popup.Controls.Add(pnlGrid);

            // Raise panels (added top-to-bottom, dock=Top means last-added is topmost)
            pnlFilter.BringToFront();
            pnlTop.BringToFront();

            // ── State ────────────────────────────────────────────────────────
            DataTable dtAll = new DataTable();
            string currentFilter = "Tất cả";

            // ── Columns ──────────────────────────────────────────────────────
            dgv.Columns.Clear();
            var cols = new[]
            {
                ("MPR_No",      "Số MPR",         90),
                ("Project_Name","Tên dự án",       200),
                ("Project_Code","Mã dự án",        110),
                ("Required_Date","Ngày yêu cầu",   105),
                ("Total_Items", "Tổng SL hạng mục", 55),
                ("PO_Count",    "Số lượng PO",     65),
                ("Status_Cat",  "Trạng thái",      140),
                ("PO_List",     "Số PO",           220),
            };
            foreach (var (name, hdr, w) in cols)
            {
                var col = new DataGridViewTextBoxColumn { Name = name, HeaderText = hdr, DataPropertyName = name, FillWeight = w };
                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                dgv.Columns.Add(col);
            }
            dgv.Columns["Total_Items"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv.Columns["PO_Count"].DefaultCellStyle.Alignment    = DataGridViewContentAlignment.MiddleCenter;
            dgv.Columns["Status_Cat"].DefaultCellStyle.Alignment  = DataGridViewContentAlignment.MiddleCenter;

            // ── Cell coloring by status ───────────────────────────────────
            dgv.CellFormatting += (s, ev) =>
            {
                if (ev.RowIndex < 0) return;
                var row = dgv.Rows[ev.RowIndex];
                var cat = row.Cells["Status_Cat"].Value?.ToString() ?? "";
                Color bg = cat switch
                {
                    "Đã gửi PO"           => Color.FromArgb(210, 240, 220),
                    "Đang chờ báo giá"     => Color.FromArgb(255, 243, 205),
                    "Chưa gửi báo giá"     => Color.FromArgb(255, 220, 220),
                    _                      => Color.White
                };
                if (ev.RowIndex % 2 == 0)
                    row.DefaultCellStyle.BackColor = bg;
                else
                    row.DefaultCellStyle.BackColor = ControlPaint.Light(bg, 0.3f);
            };

            // ── Apply filter helper ───────────────────────────────────────
            void ApplyFilter(string cat)
            {
                currentFilter = cat;
                // Highlight active button
                for (int i = 0; i < filterBtns.Length; i++)
                {
                    filterBtns[i].Font = new Font("Segoe UI", 9,
                        filterBtns[i].Tag?.ToString() == cat ? FontStyle.Bold : FontStyle.Regular);
                    filterBtns[i].FlatAppearance.BorderSize = filterBtns[i].Tag?.ToString() == cat ? 2 : 0;
                }

                var view = cat == "Tất cả"
                    ? dtAll
                    : new DataView(dtAll, $"Status_Cat = '{cat}'", null, DataViewRowState.CurrentRows).ToTable();

                dgv.DataSource = null;
                dgv.DataSource = view;

                // Summary counts
                int total    = dtAll.Rows.Count;
                int chua     = dtAll.Select("Status_Cat = 'Chưa gửi báo giá'").Length;
                int cho      = dtAll.Select("Status_Cat = 'Đang chờ báo giá'").Length;
                int daPO     = dtAll.Select("Status_Cat = 'Đã gửi PO'").Length;
                lblSummary.Text = $"Tổng: {total}  |  Chưa BG: {chua}  |  Chờ BG: {cho}  |  Đã PO: {daPO}";
            }

            // ── Load data ────────────────────────────────────────────────────
            void LoadData()
            {
                string projFilter  = cboProject.SelectedValue?.ToString() ?? "";
                string mprFilter   = txtMPRNo.Text.Trim();

                const string SQL_SIB = @"
IF OBJECT_ID('tempdb..#SibDetStatus') IS NOT NULL DROP TABLE #SibDetStatus;
SELECT DISTINCT dC.Detail_ID AS CurrID, dS.Detail_ID AS SibID
INTO #SibDetStatus
FROM (
    SELECT h.MPR_ID, h.MPR_No, h.Project_Code,
           ROW_NUMBER() OVER (
               PARTITION BY h.Project_Code,
                   SUBSTRING(h.MPR_No, 1, CHARINDEX('_Rev.', h.MPR_No + '_Rev.') - 1)
               ORDER BY TRY_CAST(TRY_CAST(h.Rev AS DECIMAL(10,2)) AS INT) DESC, h.MPR_ID DESC
           ) AS rn
    FROM MPR_Header h
    WHERE ISNULL(h.Is_Deleted,0) = 0
) q
INNER JOIN MPR_Details dC ON dC.MPR_ID = q.MPR_ID
    AND TRY_CAST(TRY_CAST(dC.Item_No AS DECIMAL(10,2)) AS INT) > 0
    AND ISNULL(dC.Is_Deleted,0) = 0
INNER JOIN MPR_Header hS
    ON SUBSTRING(hS.MPR_No, 1, CHARINDEX('_Rev.', hS.MPR_No + '_Rev.') - 1)
     = SUBSTRING(q.MPR_No,  1, CHARINDEX('_Rev.', q.MPR_No  + '_Rev.') - 1)
    AND ISNULL(hS.Is_Deleted,0) = 0
INNER JOIN MPR_Details dS ON dS.MPR_ID = hS.MPR_ID
    AND TRY_CAST(TRY_CAST(dS.Item_No AS DECIMAL(10,2)) AS INT)
      = TRY_CAST(TRY_CAST(dC.Item_No AS DECIMAL(10,2)) AS INT)
    AND ISNULL(dS.Is_Deleted,0) = 0
WHERE q.rn = 1;";

                string SQL = @"
WITH FilteredMPR AS (
    SELECT h.MPR_ID, h.MPR_No, h.Rev, h.Status,
           h.Required_Date, h.Project_Code,
           ROW_NUMBER() OVER (
               PARTITION BY h.Project_Code,
                   SUBSTRING(h.MPR_No, 1, CHARINDEX('_Rev.', h.MPR_No + '_Rev.') - 1)
               ORDER BY TRY_CAST(TRY_CAST(h.Rev AS DECIMAL(10,2)) AS INT) DESC, h.MPR_ID DESC
           ) AS rn
    FROM MPR_Header h
    WHERE ISNULL(h.Is_Deleted,0) = 0 AND ISNULL(h.Status, '') <> N'Hủy'
),
PO_Flat AS (
    SELECT DISTINCT sd.CurrID, ph.PONo
    FROM #SibDetStatus sd
    INNER JOIN PO_Detail pod ON pod.MPR_Detail_ID = sd.SibID
    INNER JOIN PO_head   ph  ON ph.PO_ID = pod.PO_ID
    WHERE ISNULL(ph.Status,'') <> 'Cancelled'
),
cte_PO AS (
    SELECT pf.CurrID,
           COUNT(DISTINCT pf.PONo) AS PO_Count,
           ISNULL(STUFF((
               SELECT DISTINCT N', ' + pf2.PONo
               FROM PO_Flat pf2
               WHERE pf2.CurrID = pf.CurrID
               FOR XML PATH(''), TYPE).value('.','NVARCHAR(MAX)'),1,2,N''),N'') AS PO_List
    FROM PO_Flat pf
    GROUP BY pf.CurrID
),
MPR_PO AS (
    SELECT d.MPR_ID,
           COUNT(d.Detail_ID) AS Total_Items,
           SUM(CASE WHEN cp.PO_Count > 0 THEN 1 ELSE 0 END) AS Items_With_PO,
           SUM(ISNULL(cp.PO_Count,0)) AS Total_PO_Links,
           ISNULL(STUFF((
               SELECT DISTINCT N', ' + cp2.PO_List
               FROM MPR_Details d2
               INNER JOIN cte_PO cp2 ON cp2.CurrID = d2.Detail_ID
               WHERE d2.MPR_ID = d.MPR_ID
                 AND ISNULL(d2.Is_Deleted,0) = 0
                 AND LEN(ISNULL(cp2.PO_List,'')) > 0
               FOR XML PATH(''), TYPE).value('.','NVARCHAR(MAX)'),1,2,N''),N'') AS PO_List
    FROM MPR_Details d
    LEFT JOIN cte_PO cp ON cp.CurrID = d.Detail_ID
    WHERE ISNULL(d.Is_Deleted,0) = 0
    GROUP BY d.MPR_ID
)
SELECT
    f.MPR_No,
    ISNULL(pi.ProjectName, f.Project_Code) AS Project_Name,
    f.Project_Code,
    CONVERT(NVARCHAR(10), f.Required_Date, 103) AS Required_Date,
    ISNULL(mp.Total_Items,0)    AS Total_Items,
    ISNULL(mp.Total_PO_Links,0) AS PO_Count,
    CASE
        WHEN ISNULL(mp.Total_PO_Links,0) > 0        THEN N'Đã gửi PO'
        WHEN ISNULL(f.Status,'') <> N'Mới'          THEN N'Đang chờ báo giá'
        ELSE                                              N'Chưa gửi báo giá'
    END AS Status_Cat,
    ISNULL(mp.PO_List, N'')     AS PO_List
FROM FilteredMPR f
LEFT JOIN MPR_PO mp ON mp.MPR_ID = f.MPR_ID
LEFT JOIN Project_Info pi ON pi.ProjectCode = f.Project_Code
WHERE f.rn = 1";

                // Dynamic filters
                var whereParts = new System.Collections.Generic.List<string>();
                if (!string.IsNullOrEmpty(projFilter))
                    whereParts.Add($"  AND f.Project_Code = @projCode");
                if (!string.IsNullOrEmpty(mprFilter))
                    whereParts.Add($"  AND f.MPR_No LIKE @mprNo");
                if (whereParts.Count > 0)
                    SQL += "\n" + string.Join("\n", whereParts);

                SQL += "\nORDER BY pi.ProjectName, f.MPR_No";

                try
                {
                    using (var conn = DatabaseHelper.GetConnection())
                    {
                        conn.Open();
                        var cmdSib = new SqlCommand(SQL_SIB, conn) { CommandTimeout = 120 };
                        cmdSib.ExecuteNonQuery();

                        var cmd = new SqlCommand(SQL, conn) { CommandTimeout = 120 };
                        if (!string.IsNullOrEmpty(projFilter))
                            cmd.Parameters.AddWithValue("@projCode", projFilter);
                        if (!string.IsNullOrEmpty(mprFilter))
                            cmd.Parameters.AddWithValue("@mprNo", "%" + mprFilter + "%");

                        dtAll = new DataTable();
                        dtAll.Load(cmd.ExecuteReader());
                    }
                    ApplyFilter(currentFilter);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(TopOwner, "Lỗi tải dữ liệu MPR Status:\n" + ex.Message, "Lỗi",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // ── Load project list ─────────────────────────────────────────────
            void LoadProjects()
            {
                try
                {
                    using (var conn = DatabaseHelper.GetConnection())
                    {
                        conn.Open();
                        string sqlProj = @"
SELECT DISTINCT ISNULL(pi.ProjectName, h.Project_Code) AS DisplayName, h.Project_Code
FROM MPR_Header h
LEFT JOIN Project_Info pi ON pi.ProjectCode = h.Project_Code
WHERE ISNULL(h.Is_Deleted,0) = 0
ORDER BY DisplayName";
                        var dt = new DataTable();
                        dt.Columns.Add("DisplayName");
                        dt.Columns.Add("Project_Code");
                        dt.Rows.Add("(Tất cả)", "");
                        var dtProj = new DataTable();
                        dtProj.Load(new SqlCommand(sqlProj, conn).ExecuteReader());
                        foreach (DataRow row in dtProj.Rows)
                            dt.Rows.Add(row["DisplayName"].ToString(), row["Project_Code"].ToString());

                        cboProject.DataSource    = dt;
                        cboProject.DisplayMember = "DisplayName";
                        cboProject.ValueMember   = "Project_Code";
                    }
                }
                catch { /* ignore — combobox just shows (Tất cả) */ }
                cboProject.SelectedIndex = 0;
            }

            // ── Wire events ──────────────────────────────────────────────────
            btnSearch.Click += (s, ev) => LoadData();
            txtMPRNo.KeyDown += (s, ev) => { if (ev.KeyCode == Keys.Enter) LoadData(); };
            foreach (var btn in filterBtns)
            {
                var cap = btn.Tag?.ToString() ?? "";
                btn.Click += (s, ev) => ApplyFilter(cap);
            }

            // ── Initialize ───────────────────────────────────────────────────
            popup.Load += (s, ev) =>
            {
                LoadProjects();
                LoadData();
            };

            popup.ShowDialog(this);
        }

        // =====================================================================
        //  XUAT EXCEL CHI TIET VAT TU
        // =====================================================================
        private void BtnExportDetail_Click(object sender, EventArgs e)
        {
            var m = _mprList.Find(x => x.MPR_ID == _selectedMPR_ID);
            if (m != null && m.Status == "Hủy")
            {
                SafeMsg("Phiếu MPR này đã bị Hủy, không thể xuất chi tiết!", "Thông báo", MessageBoxIcon.Warning);
                return;
            }

            if (dgvDetails.Rows.Count == 0)
            {
                MessageBox.Show(TopOwner, "Khong co du lieu de xuat!", "Thong bao",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using var sfd = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = "MPR_ChiTiet_" + (txtMPRNo?.Text.Trim() ?? "export")
                           + "_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".xlsx",
                Title = "Luu file Excel",
                InitialDirectory = Directory.Exists(@"D:\RÁC") ? @"D:\RÁC" : Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };
            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using var pkg = new OfficeOpenXml.ExcelPackage();
                var ws = pkg.Workbook.Worksheets.Add("Chi tiet vat tu");

                // Header row
                var visCols = new System.Collections.Generic.List<DataGridViewColumn>();
                foreach (DataGridViewColumn col in dgvDetails.Columns)
                    if (col.Visible && col.Name != "Detail_ID")
                        visCols.Add(col);

                // Style header
                for (int ci = 0; ci < visCols.Count; ci++)
                {
                    var cell = ws.Cells[1, ci + 1];
                    cell.Value = visCols[ci].HeaderText;
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 120, 212));
                    cell.Style.Font.Color.SetColor(System.Drawing.Color.White);
                    cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                }

                // Lay gia tri truc tiep tu _details (so nguyen/decimal thuc)
                // de tranh bi format string boi CellFormatting
                var poMap2 = GetPoMappingForMpr(_selectedMPR_ID);
                // Build map Detail_ID -> row.Visible de biet dong nao dang hien
                var visibleIds = new System.Collections.Generic.HashSet<int>();
                foreach (DataGridViewRow dgvRow in dgvDetails.Rows)
                {
                    if (dgvRow.IsNewRow || !dgvRow.Visible) continue;

                    // --- BỎ QUA CÁC DÒNG ĐÃ BỊ XÓA (Gạch ngang / STT = 0) ---
                    if (dgvRow.Tag?.ToString() == "deleted" || dgvRow.Cells["Item_No"]?.Value?.ToString() == "0") continue;

                    if (dgvRow.Cells["Detail_ID"]?.Value is int rid) visibleIds.Add(rid);
                    else if (int.TryParse(dgvRow.Cells["Detail_ID"]?.Value?.ToString(), out int rid2))
                        visibleIds.Add(rid2);
                }

                int excelRow = 2;
                foreach (var d in _details)
                {
                    if (!visibleIds.Contains(d.Detail_ID)) continue;

                    string poNo = poMap2.ContainsKey(d.Detail_ID) ? poMap2[d.Detail_ID] : "";

                    // Map gia tri theo ten cot
                    var rowData = new System.Collections.Generic.Dictionary<string, object>
                    {
                        ["Item_No"] = d.Item_No,
                        ["Item_Name"] = d.Item_Name ?? "",
                        ["Description"] = d.Description ?? "",
                        ["Material"] = d.Material ?? "",
                        ["Thickness_mm"] = d.Thickness_mm,
                        ["Depth_mm"] = d.Depth_mm,
                        ["C_Width_mm"] = d.C_Width_mm,
                        ["D_Web_mm"] = d.D_Web_mm,
                        ["E_Flange_mm"] = d.E_Flange_mm,
                        ["F_Length_mm"] = d.F_Length_mm,
                        ["UNIT"] = d.UNIT ?? "",
                        ["Qty"] = d.Qty_Per_Sheet,
                        ["Weight"] = d.Weight_kg,
                        ["MPS_Info"] = d.MPS_Info ?? "",
                        ["Usage_Location"] = d.Usage_Location ?? "",
                        ["REV"] = d.REV,
                        ["Remarks"] = d.Remarks ?? "",
                        ["PO_No"] = poNo
                    };

                    var numericSet = new System.Collections.Generic.HashSet<string>
                        { "Thickness_mm","Depth_mm","C_Width_mm","D_Web_mm",
                          "E_Flange_mm","F_Length_mm","Qty","Weight" };

                    for (int ci = 0; ci < visCols.Count; ci++)
                    {
                        string colN = visCols[ci].Name;
                        var cell = ws.Cells[excelRow, ci + 1];

                        if (!rowData.ContainsKey(colN)) { cell.Value = ""; continue; }

                        object rawVal = rowData[colN];

                        if (numericSet.Contains(colN))
                        {
                            // Set so thuc truc tiep - KHONG qua string
                            double dbl = Convert.ToDouble(rawVal);
                            cell.Value = dbl;
                            // Format: hien so thap phan neu co, khong co separator
                            cell.Style.Numberformat.Format = "0.##";
                        }
                        else
                            cell.Value = rawVal;

                        // Mau dong xen ke
                        if (excelRow % 2 == 0)
                        {
                            cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            cell.Style.Fill.BackgroundColor.SetColor(
                                System.Drawing.Color.FromArgb(240, 248, 255));
                        }

                        // Mau cot Da len PO
                        if (colN == "PO_No" && !string.IsNullOrWhiteSpace(poNo))
                        {
                            cell.Style.Font.Bold = true;
                            cell.Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(40, 167, 69));
                        }

                        cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Hair);
                    }
                    excelRow++;
                }

                // Title MPR info
                ws.InsertRow(1, 2);
                ws.Cells[1, 1].Value = "BANG CHI TIET VAT TU MPR";
                ws.Cells[1, 1].Style.Font.Bold = true;
                ws.Cells[1, 1].Style.Font.Size = 13;
                ws.Cells[1, 1, 1, visCols.Count].Merge = true;

                ws.Cells[2, 1].Value = "MPR No: " + (txtMPRNo?.Text ?? "") +
                                       "   Du an: " + (txtProjectName?.Text ?? "") +
                                       "   Ngay: " + DateTime.Now.ToString("dd/MM/yyyy");
                ws.Cells[2, 1, 2, visCols.Count].Merge = true;

                // Filter info
                string filterInfo = _cboFilterPO?.SelectedItem?.ToString() ?? "Tat ca";
                ws.Cells[2, 1].Value += "   Loc: " + filterInfo;

                // Auto fit
                ws.Cells[ws.Dimension.Address].AutoFitColumns(8, 50);

                pkg.SaveAs(new System.IO.FileInfo(sfd.FileName));

                if (MessageBox.Show(TopOwner, "Xuat Excel thanh cong!\nMo file ngay?", "Thanh cong",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    System.Diagnostics.Process.Start(
                        new System.Diagnostics.ProcessStartInfo
                        { FileName = sfd.FileName, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show(TopOwner, "Loi xuat Excel: " + ex.Message, "Loi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // =====================================================================
        //  COPY BANG SANG CLIPBOARD (Ctrl+C / Ctrl+Shift+C)
        // =====================================================================
        private void DgvDetails_GridKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C)
            {
                CopyGridToClipboard(e.Shift); // Shift = kem header
                e.Handled = true;
            }

            if (e.Control && e.KeyCode == Keys.V)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                PasteFromExcelToDetails();
            }
        }

        private void PasteFromExcelToDetails()
        {
            if (_selectedMPR_ID == 0)
            {
                MessageBox.Show(TopOwner, "Vui lòng chọn phiếu MPR trước khi dán dữ liệu!",
                    "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string clip = Clipboard.GetText();
            if (string.IsNullOrEmpty(clip)) return;

            // Tách dòng — bỏ dòng rỗng cuối (Excel luôn thêm \r\n)
            string[] rows = clip.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            if (rows.Length > 0 && string.IsNullOrEmpty(rows[rows.Length - 1]))
                rows = rows.Take(rows.Length - 1).ToArray();
            if (rows.Length == 0) return;

            // Các cột có thể paste (bỏ ẩn, bỏ ReadOnly)
            var pastableCols = dgvDetails.Columns
                .Cast<DataGridViewColumn>()
                .Where(c => c.Visible
                         && c.Name != "Detail_ID"
                         && c.Name != "Item_No"
                         && c.Name != "PO_No")
                .OrderBy(c => c.DisplayIndex)
                .ToList();

            // Cột bắt đầu paste = cột hiện tại (hoặc cột đầu tiên có thể paste)
            int startColDisplayIdx = dgvDetails.CurrentCell?.ColumnIndex >= 0
                ? dgvDetails.CurrentCell.ColumnIndex : 0;

            // Lọc cột paste từ vị trí bắt đầu trở đi
            var targetCols = pastableCols
                .Where(c => c.Index >= startColDisplayIdx)
                .ToList();

            if (targetCols.Count == 0) targetCols = pastableCols;

            // Dòng bắt đầu paste
            int startRowIdx = dgvDetails.CurrentCell?.RowIndex >= 0
                ? dgvDetails.CurrentCell.RowIndex : 0;

            // Cần thêm dòng mới nếu không đủ
            int rowsNeeded = (startRowIdx + rows.Length) - dgvDetails.Rows.Count;
            for (int i = 0; i < rowsNeeded; i++)
            {
                int newIdx = dgvDetails.Rows.Add();
                var nr = dgvDetails.Rows[newIdx];
                nr.Cells["Detail_ID"].Value = 0;
                nr.Cells["Item_No"].Value = newIdx + 1;
                nr.Cells["Thickness_mm"].Value = 0;
                nr.Cells["Depth_mm"].Value = 0;
                nr.Cells["C_Width_mm"].Value = 0;
                nr.Cells["D_Web_mm"].Value = 0;
                nr.Cells["E_Flange_mm"].Value = 0;
                nr.Cells["F_Length_mm"].Value = 0;
                nr.Cells["Qty"].Value = 0;
                nr.Cells["Weight"].Value = 0;
                nr.Cells["REV"].Value = "0";
                nr.Cells["UNIT"].Value = "cái";
            }

            // Ghi dữ liệu
            dgvDetails.SuspendLayout();
            for (int ri = 0; ri < rows.Length; ri++)
            {
                int targetRowIdx = startRowIdx + ri;
                if (targetRowIdx >= dgvDetails.Rows.Count) break;
                var row = dgvDetails.Rows[targetRowIdx];
                if (row.IsNewRow) break;

                string[] cells = rows[ri].Split('\t');
                for (int ci = 0; ci < cells.Length && ci < targetCols.Count; ci++)
                {
                    string colName = targetCols[ci].Name;
                    string val = cells[ci].Trim();

                    // Cột số: parse linh hoạt (bỏ dấu ngàn, đổi dấu thập phân)
                    bool isNumCol = colName is "Thickness_mm" or "Depth_mm" or "C_Width_mm"
                                             or "D_Web_mm" or "E_Flange_mm" or "F_Length_mm"
                                             or "Qty" or "Weight";
                    if (isNumCol)
                    {
                        string num = val.Replace(",", ".").Replace(" ", "");
                        // Xử lý dấu ngàn: "1.234.567" → "1234567"
                        int dotCount = num.Count(c => c == '.');
                        if (dotCount > 1) num = num.Replace(".", "");
                        row.Cells[colName].Value = decimal.TryParse(num,
                            System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture,
                            out decimal d) ? d : 0m;
                    }
                    else
                    {
                        row.Cells[colName].Value = val;
                    }
                }

                // Cập nhật lại STT
                row.Cells["Item_No"].Value = targetRowIdx + 1;
            }
            dgvDetails.ResumeLayout();
            dgvDetails.Refresh();

            MessageBox.Show(TopOwner,
                $"✅ Đã dán {rows.Length} dòng thành công!\n\nNhấn '💾 Lưu chi tiết' để lưu vào database.",
                "Dán dữ liệu", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void CopyGridToClipboard(bool withHeader)
        {
            try
            {
                // Chi copy cac dong DANG DUOC CHON (selected rows)
                var selectedRows = dgvDetails.SelectedRows
                    .Cast<DataGridViewRow>()
                    .Where(r => !r.IsNewRow && r.Tag?.ToString() != "deleted" && r.Cells["Item_No"]?.Value?.ToString() != "0")
                    .OrderBy(r => r.Index)
                    .ToList();

                if (selectedRows.Count == 0)
                {
                    MessageBox.Show(TopOwner, "Vui long chon it nhat mot dong de copy!",
                        "Thong bao", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Lay cot hien thi
                var visCols = dgvDetails.Columns
                    .Cast<DataGridViewColumn>()
                    .Where(c => c.Visible && c.Name != "Detail_ID")
                    .OrderBy(c => c.DisplayIndex)
                    .ToList();

                var numSet = new System.Collections.Generic.HashSet<string>
                    { "Qty","Weight","Thickness_mm","Depth_mm",
                      "C_Width_mm","D_Web_mm","E_Flange_mm","F_Length_mm" };

                var sb = new System.Text.StringBuilder();

                // Header neu can
                if (withHeader)
                    sb.AppendLine(string.Join("\t", visCols.Select(c => c.HeaderText)));

                // Doc gia tri tu _details theo Detail_ID de dam bao la so thuc
                var poMap = GetPoMappingForMpr(_selectedMPR_ID);

                foreach (var row in selectedRows)
                {
                    int detId = 0;
                    int.TryParse(row.Cells["Detail_ID"]?.Value?.ToString(), out detId);
                    var d = _details.Find(x => x.Detail_ID == detId);

                    var parts = new System.Collections.Generic.List<string>();
                    foreach (var col in visCols)
                    {
                        string colN = col.Name;
                        string cellVal;

                        if (d != null && numSet.Contains(colN))
                        {
                            // Doc so thuc tu _details, xuat dang InvariantCulture
                            double dbl = colN switch
                            {
                                "Qty" => (double)d.Qty_Per_Sheet,
                                "Weight" => (double)d.Weight_kg,
                                "Thickness_mm" => (double)d.Thickness_mm,
                                "Depth_mm" => (double)d.Depth_mm,
                                "C_Width_mm" => (double)d.C_Width_mm,
                                "D_Web_mm" => (double)d.D_Web_mm,
                                "E_Flange_mm" => (double)d.E_Flange_mm,
                                "F_Length_mm" => (double)d.F_Length_mm,
                                _ => 0
                            };
                            // Xuat: so nguyen thi khong co thap phan
                            cellVal = (dbl == Math.Floor(dbl))
                                ? ((long)dbl).ToString()
                                : dbl.ToString(System.Globalization.CultureInfo.InvariantCulture);
                        }
                        else if (colN == "PO_No")
                            cellVal = (d != null && poMap.ContainsKey(d.Detail_ID))
                                ? poMap[d.Detail_ID] : "";
                        else
                            cellVal = row.Cells[colN]?.Value?.ToString() ?? "";

                        parts.Add(cellVal);
                    }
                    sb.AppendLine(string.Join("\t", parts));
                }

                Clipboard.SetText(sb.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(TopOwner, "Loi copy: " + ex.Message, "Loi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =====================================================================
        //  EMAIL CELL FORMATTING
        // =====================================================================
        private void DgvMPR_EmailCellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            if (e.ColumnIndex >= dgvMPR.Columns.Count) return;
            var col = dgvMPR.Columns[e.ColumnIndex];
            if (col == null) return;
            string colName = col.Name;

            string status = "";
            if (dgvMPR.Columns.Contains("Email_Status") && e.RowIndex < dgvMPR.Rows.Count)
            {
                var cell = dgvMPR.Rows[e.RowIndex].Cells["Email_Status"];
                status = cell?.Value?.ToString() ?? "";
            }
            bool done = status.Equals("done", StringComparison.OrdinalIgnoreCase);

            if (colName == "col_SendEmail_MPR")
            {
                if (done)
                {
                    e.Value = "✔ Đã gửi";
                    e.CellStyle.BackColor = Color.FromArgb(200, 200, 200);
                    e.CellStyle.ForeColor = Color.Black;
                }
                else
                {
                    e.Value = "📧 Gửi";
                    e.CellStyle.BackColor = Color.FromArgb(40, 167, 69);
                    e.CellStyle.ForeColor = Color.White;
                }
                e.FormattingApplied = true;
            }
            else if (colName == "Email_Status")
            {
                if (done)
                {
                    e.Value = "✔ done";
                    e.CellStyle.ForeColor = Color.Black;
                    e.CellStyle.BackColor = Color.FromArgb(200, 200, 200);
                    e.FormattingApplied = true;
                }
            }
        }

        // =====================================================================
        //  EMAIL BUTTON CLICK
        // =====================================================================
        private void DgvMPR_EmailCellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            if (e.ColumnIndex >= dgvMPR.Columns.Count) return;
            if (dgvMPR.Columns[e.ColumnIndex]?.Name != "col_SendEmail_MPR") return;

            var idCell = dgvMPR.Rows[e.RowIndex].Cells["ID"];
            int mprId = Convert.ToInt32(idCell?.Value ?? 0);
            if (mprId == 0) return;

            var mpr = _mprList.Find(m => m.MPR_ID == mprId);
            if (mpr == null) return;

            // Nếu đã gửi email → hỏi người dùng muốn làm gì
            if (mpr.Email_Status?.Equals("done", StringComparison.OrdinalIgnoreCase) == true)
            {
                ShowEmailOrZaloChoiceMPR(mpr);
                return;
            }

            ShowEmailComposePopupMPR(mpr);
        }

        // =====================================================================
        //  EMAIL ĐÃ GỬI — hỏi "Tiếp tục gửi Email" hay "Gửi tin nhắn thông báo"
        // =====================================================================
        private void ShowEmailOrZaloChoiceMPR(MPRHeader mpr)
        {
            var dlg = new Form
            {
                Text = $"MPR {mpr.MPR_No} — Đã gửi email",
                Size = new Size(410, 175),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false, MinimizeBox = false,
                BackColor = Color.White, Font = new Font("Segoe UI", 9.5f)
            };

            dlg.Controls.Add(new Label
            {
                Text = "✔ Email đã được gửi. Bạn muốn làm gì?",
                Location = new Point(16, 18), Size = new Size(370, 22),
                Font = new Font("Segoe UI", 9.5f)
            });

            var btnEmail = new Button
            {
                Text = "📧 Tiếp tục gửi Email",
                Location = new Point(16, 55), Size = new Size(175, 40),
                BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };
            btnEmail.FlatAppearance.BorderSize = 0;

            var btnZalo = new Button
            {
                Text = "🔔 Gửi tin nhắn thông báo",
                Location = new Point(208, 55), Size = new Size(175, 40),
                BackColor = Color.FromArgb(0, 150, 136), ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };
            btnZalo.FlatAppearance.BorderSize = 0;

            dlg.Controls.AddRange(new Control[] { btnEmail, btnZalo });

            btnEmail.Click += (s, ev) => { dlg.Close(); ShowEmailComposePopupMPR(mpr); };
            btnZalo.Click  += (s, ev) => { dlg.Close(); ShowZaloGroupSelectAndSendMPR(mpr); };

            dlg.ShowDialog(this);
        }

        // =====================================================================
        //  CHỌN NHÓM ZALO & GỬI THÔNG BÁO (MPR)
        // =====================================================================
        private void ShowZaloGroupSelectAndSendMPR(MPRHeader mpr)
        {
            var cfg = Helpers.ZaloHelper.LoadSettings();
            if (!cfg.Enabled)
            {
                MessageBox.Show(TopOwner,
                    "Zalo chưa được bật.\nVào 🔔 Zalo trên thanh tiêu đề để bật.",
                    "Zalo chưa bật", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var allGroups = new Services.SupplierService().GetAll()
                .Where(s => !string.IsNullOrWhiteSpace(s.Zalo_Group_ID))
                .OrderBy(s => s.Company_Name)
                .ToList();

            if (allGroups.Count == 0)
            {
                MessageBox.Show(TopOwner,
                    "Không có nhà cung cấp nào có nhóm Zalo được cấu hình.\nVào 🏢 Nhà Cung Cấp để điền tên nhóm.",
                    "Không có nhóm Zalo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // checkedIds giữ Supplier_ID được tích — không bị mất khi lọc
            var checkedIds = new HashSet<int>();
            // filteredGroups là danh sách hiện tại đang hiển thị trong clb
            var filteredGroups = new List<MPR_Managerment.Models.Supplier>(allGroups);

            var dlg = new Form
            {
                Text = $"Chọn nhóm Zalo — MPR {mpr.MPR_No}",
                Size = new Size(440, 570),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false, MinimizeBox = false,
                BackColor = Color.White, Font = new Font("Segoe UI", 9f)
            };

            dlg.Controls.Add(new Label
            {
                Text = "Chọn một hoặc nhiều nhóm Zalo để gửi thông báo:",
                Location = new Point(12, 12), Size = new Size(400, 20)
            });

            // ── Thanh tìm kiếm ───────────────────────────────────────────────
            var txtSearch = new TextBox
            {
                Location = new Point(12, 36), Size = new Size(404, 26),
                Font = new Font("Segoe UI", 9f),
                ForeColor = Color.Gray, Text = "🔍  Tìm kiếm nhóm...",
                BorderStyle = BorderStyle.FixedSingle
            };
            dlg.Controls.Add(txtSearch);

            var clb = new CheckedListBox
            {
                Location = new Point(12, 68),
                Size = new Size(404, 370),
                CheckOnClick = true,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9f)
            };
            foreach (var s in allGroups)
                clb.Items.Add($"{s.Company_Name}  [{s.Zalo_Group_ID}]", false);
            dlg.Controls.Add(clb);

            // Rebuild clb theo keyword, giữ checkmark qua checkedIds
            void Refilter(string keyword)
            {
                // Flush trạng thái check hiện tại vào checkedIds trước khi rebuild
                for (int i = 0; i < clb.Items.Count; i++)
                {
                    int sid = filteredGroups[i].Supplier_ID;
                    if (clb.GetItemChecked(i)) checkedIds.Add(sid);
                    else checkedIds.Remove(sid);
                }

                filteredGroups = string.IsNullOrWhiteSpace(keyword)
                    ? new List<MPR_Managerment.Models.Supplier>(allGroups)
                    : allGroups.Where(s =>
                        s.Company_Name.Contains(keyword, StringComparison.OrdinalIgnoreCase) ||
                        s.Zalo_Group_ID.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                      .ToList();

                clb.ItemCheck -= ClbItemCheck;
                clb.Items.Clear();
                foreach (var s in filteredGroups)
                    clb.Items.Add($"{s.Company_Name}  [{s.Zalo_Group_ID}]", checkedIds.Contains(s.Supplier_ID));
                clb.ItemCheck += ClbItemCheck;
            }

            void ClbItemCheck(object sender, ItemCheckEventArgs e)
            {
                int sid = filteredGroups[e.Index].Supplier_ID;
                if (e.NewValue == CheckState.Checked) checkedIds.Add(sid);
                else checkedIds.Remove(sid);
            }

            clb.ItemCheck += ClbItemCheck;

            // Placeholder logic
            txtSearch.GotFocus  += (s, ev) => { if (txtSearch.ForeColor == Color.Gray) { txtSearch.Text = ""; txtSearch.ForeColor = Color.Black; } };
            txtSearch.LostFocus += (s, ev) => { if (string.IsNullOrWhiteSpace(txtSearch.Text)) { txtSearch.ForeColor = Color.Gray; txtSearch.Text = "🔍  Tìm kiếm nhóm..."; } };
            txtSearch.TextChanged += (s, ev) =>
            {
                if (txtSearch.ForeColor == Color.Gray) return; // đang hiện placeholder
                Refilter(txtSearch.Text.Trim());
            };

            var btnSend = new Button
            {
                Text = "🔔 Gửi thông báo",
                Location = new Point(12, 455), Size = new Size(185, 38),
                BackColor = Color.FromArgb(0, 150, 136), ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };
            btnSend.FlatAppearance.BorderSize = 0;

            var btnAll = new Button
            {
                Text = "Chọn tất cả",
                Location = new Point(213, 455), Size = new Size(100, 38),
                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand
            };
            var btnCancel = new Button
            {
                Text = "Huỷ",
                Location = new Point(328, 455), Size = new Size(88, 38),
                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand
            };
            dlg.Controls.AddRange(new Control[] { btnSend, btnAll, btnCancel });

            btnAll.Click += (s, ev) =>
            {
                // Tích tất cả item đang hiện (filtered)
                for (int i = 0; i < clb.Items.Count; i++) clb.SetItemChecked(i, true);
            };
            btnCancel.Click += (s, ev) => dlg.Close();

            btnSend.Click += (s, ev) =>
            {
                // Flush lần cuối trước khi lấy kết quả
                for (int i = 0; i < clb.Items.Count; i++)
                {
                    int sid = filteredGroups[i].Supplier_ID;
                    if (clb.GetItemChecked(i)) checkedIds.Add(sid);
                    else checkedIds.Remove(sid);
                }

                var selected = allGroups.Where(s => checkedIds.Contains(s.Supplier_ID)).ToList();
                if (selected.Count == 0)
                {
                    MessageBox.Show(dlg, "Vui lòng chọn ít nhất một nhóm.",
                        "Chưa chọn nhóm", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                dlg.Close();

                string msg =
                    $"📋 DLHI Thông báo:\n" +
                    $"Chúng tôi đã gửi yêu cầu báo giá cho:\n" +
                    $"🔹 Số MPR   : {mpr.MPR_No}\n" +
                    $"🔹 Dự án   : {mpr.Project_Name}\n" +
                    $"🔹 Thời gian: {DateTime.Now:dd/MM/yyyy HH:mm}\n" +
                    $"Xin vui lòng kiểm tra Email, Rất mong sớm nhận được báo giá từ Quý công ty!";

                _SendZaloToSelectedGroupsMPR(selected, msg, mpr.MPR_No);
            };

            dlg.ShowDialog(this);
        }

        private void _SendZaloToSelectedGroupsMPR(
            List<MPR_Managerment.Models.Supplier> selected, string msg, string mprNo)
        {
            void SetSt(string text) { if (lblStatus != null && !lblStatus.IsDisposed) lblStatus.Text = text; }

            SetSt($"⏳ Đang đặt {selected.Count} tin nhắn Zalo vào hàng đợi...");

            foreach (var sup in selected)
            {
                // Lấy tên nhóm từ Zalo_Group_ID (giữ nguyên để dễ debug)
                string groupName = sup.Zalo_Group_ID ?? sup.Company_Name ?? "Unknown";
                ZaloNotificationService.Instance.Enqueue(groupName, msg, $"MPR: {mprNo}");
            }

            SetSt($"🔔 Đã đặt {selected.Count} tin nhắn Zalo vào hàng đợi (sẽ gửi theo thứ tự)");
        }

        // =====================================================================
        //  EMAIL COMPOSE — mở Outlook trực tiếp, BCC tất cả nhà cung cấp được chọn
        // =====================================================================
        private void ShowEmailComposePopupMPR(MPRHeader mpr)
        {
            // --- Bước 1: Chọn nhà cung cấp ---
            var allSuppliers = new SupplierService().GetAll()
                .Where(s => !string.IsNullOrWhiteSpace(s.Email))
                .OrderBy(s => s.Company_Name)
                .ToList();

            if (allSuppliers.Count == 0)
            {
                MessageBox.Show(TopOwner, "Không có nhà cung cấp nào có email trong hệ thống.\nVui lòng cập nhật email nhà cung cấp trước.", "Thiếu email", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var selDlg = new Form
            {
                Text = $"Chọn nhà cung cấp — MPR: {mpr.MPR_No}",
                Size = new Size(520, 480),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false, MinimizeBox = false,
                BackColor = Color.White, Font = new Font("Segoe UI", 9)
            };

            selDlg.Controls.Add(new Label
            {
                Text = "Chọn nhà cung cấp nhận email (BCC). Có thể chọn nhiều:",
                Location = new Point(10, 10), Size = new Size(490, 20),
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            });

            var txtFilter = new TextBox { Location = new Point(10, 36), Size = new Size(490, 24), PlaceholderText = "Lọc nhanh theo tên / email..." };
            var clb = new CheckedListBox { Location = new Point(10, 66), Size = new Size(490, 310), CheckOnClick = true, BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 9) };
            var checkedEmails = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var s in allSuppliers)
                clb.Items.Add($"{s.Company_Name}  <{s.Email}>", false);

            clb.ItemCheck += (sender, e) =>
            {
                string itemText = clb.Items[e.Index].ToString() ?? "";
                int lt = itemText.IndexOf('<'), gt = itemText.IndexOf('>');
                if (lt >= 0 && gt > lt)
                {
                    string email = itemText.Substring(lt + 1, gt - lt - 1).Trim();
                    if (e.NewValue == CheckState.Checked)
                        checkedEmails.Add(email);
                    else
                        checkedEmails.Remove(email);
                }
            };

            txtFilter.TextChanged += (s, e) =>
            {
                string kw = txtFilter.Text.Trim().ToLower();
                clb.Items.Clear();
                foreach (var sup in allSuppliers)
                {
                    if (string.IsNullOrEmpty(kw) || sup.Company_Name.ToLower().Contains(kw) || sup.Email.ToLower().Contains(kw))
                    {
                        int idx = clb.Items.Add($"{sup.Company_Name}  <{sup.Email}>");
                        if (!string.IsNullOrEmpty(sup.Email) && checkedEmails.Contains(sup.Email))
                            clb.SetItemChecked(idx, true);
                    }
                }
            };

            var btnAll      = new Button { Text = "Chọn tất cả",   Location = new Point(10,  382), Size = new Size(110, 28), FlatStyle = FlatStyle.Flat };
            var btnNone     = new Button { Text = "Bỏ chọn tất cả", Location = new Point(126, 382), Size = new Size(120, 28), FlatStyle = FlatStyle.Flat };
            var btnNext     = new Button { Text = "Tiếp theo ▶", Location = new Point(360, 382), Size = new Size(140, 28), BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, DialogResult = DialogResult.OK };
            var btnCancelSel = new Button { Text = "Huỷ", Location = new Point(260, 382), Size = new Size(90, 28), FlatStyle = FlatStyle.Flat, DialogResult = DialogResult.Cancel };
            btnNext.FlatAppearance.BorderSize = 0;
            btnAll.Click  += (s, e) => { for (int i = 0; i < clb.Items.Count; i++) clb.SetItemChecked(i, true); };
            btnNone.Click += (s, e) => { for (int i = 0; i < clb.Items.Count; i++) clb.SetItemChecked(i, false); };

            selDlg.Controls.AddRange(new Control[] { txtFilter, clb, btnAll, btnNone, btnNext, btnCancelSel });
            selDlg.AcceptButton = btnNext; selDlg.CancelButton = btnCancelSel;
            selDlg.ClientSize = new Size(510, 420);

            if (selDlg.ShowDialog(TopOwner) != DialogResult.OK) return;

            var selectedEmails    = new List<string>();
            var selectedNames     = new List<string>();
            var selectedSuppliers = new List<MPR_Managerment.Models.Supplier>();
            for (int i = 0; i < clb.Items.Count; i++)
            {
                if (!clb.GetItemChecked(i)) continue;
                string itemText = clb.Items[i].ToString() ?? "";
                int lt = itemText.IndexOf('<'), gt = itemText.IndexOf('>');
                if (lt < 0 || gt < 0) continue;
                string email = itemText.Substring(lt + 1, gt - lt - 1).Trim();
                var sup = allSuppliers.Find(x => x.Email.Equals(email, StringComparison.OrdinalIgnoreCase));
                if (sup != null)
                {
                    selectedEmails.Add(sup.Email);
                    selectedNames.Add(sup.Company_Name);
                    selectedSuppliers.Add(sup);
                }
            }

            if (selectedEmails.Count == 0)
            {
                MessageBox.Show(TopOwner, "Chưa chọn nhà cung cấp nào!", "Thiếu lựa chọn", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // --- Bước 2: Tìm file MPR đính kèm ---
            string attachPath = FindMPRFile(mpr);

            // --- Bước 3: Compose popup ---
            var details = _service.GetDetails(mpr.MPR_ID);
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("Kính gửi Quý Nhà Cung Cấp,");
            sb.AppendLine();
            sb.AppendLine("Chúng tôi gửi đến Quý vị Phiếu Yêu cầu Mua Vật tư (MPR) với thông tin như sau:");
            sb.AppendLine($"  - Số MPR     : {mpr.MPR_No}");
            sb.AppendLine($"  - Dự án      : {mpr.Project_Name}");
            sb.AppendLine($"  - Mã dự án   : {mpr.Project_Code}");
            sb.AppendLine($"  - Ngày cần   : {mpr.Required_Date?.ToString("dd/MM/yyyy") ?? ""}");
            if (mpr.Rev > 0) sb.AppendLine($"  - Revise     : {mpr.Rev}");
            sb.AppendLine();
            sb.AppendLine("Chi tiết vật tư:");
            int stt = 1;
            foreach (var d in details)
                sb.AppendLine($"  {stt++}. {d.Item_Name}  |  SL: {d.Qty_Per_Sheet} {d.UNIT}");
            sb.AppendLine();
            sb.AppendLine("Vui lòng xem xét và báo giá trong thời gian sớm nhất. Cảm ơn Quý vị!");

            string subject = $"Phiếu Yêu cầu Mua Vật tư (MPR) - {mpr.MPR_No}";
            string bccList = string.Join(";", selectedEmails);

            var popup = new Form
            {
                Text = $"Gửi Email MPR qua Outlook — {mpr.MPR_No}",
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false, MinimizeBox = false,
                BackColor = Color.White, Font = new Font("Segoe UI", 9)
            };

            int py = 12;
            void AddRow(string lbl, Control ctrl, int h = 26)
            {
                popup.Controls.Add(new Label { Text = lbl, Location = new Point(10, py + 3), Size = new Size(70, 20), TextAlign = ContentAlignment.MiddleRight });
                ctrl.Location = new Point(85, py); ctrl.Size = new Size(575, h);
                popup.Controls.Add(ctrl); py += h + 6;
            }

            var txtBcc     = new TextBox { Text = bccList };
            var txtSubject = new TextBox { Text = subject };
            var txtBody    = new TextBox { Text = sb.ToString(), Multiline = true, ScrollBars = ScrollBars.Vertical };

            AddRow("BCC:", txtBcc);
            AddRow("Chủ đề:", txtSubject);
            AddRow("Nội dung:", txtBody, 280);

            // Dòng đính kèm
            bool hasFile = !string.IsNullOrEmpty(attachPath) && File.Exists(attachPath);
            var lblAttach = new Label
            {
                Text      = hasFile ? $"📎  {Path.GetFileName(attachPath)}" : "⚠  Không tìm thấy file MPR trong thư mục MPR Link",
                Location  = new Point(85, py), Size = new Size(470, 18),
                ForeColor = hasFile ? Color.FromArgb(0, 120, 60) : Color.FromArgb(160, 80, 0),
                Font      = new Font("Segoe UI", 8.5f, FontStyle.Italic)
            };
            popup.Controls.Add(lblAttach);

            // Nút duyệt file thủ công
            var btnBrowse = new Button
            {
                Text = "...", Location = new Point(560, py - 2), Size = new Size(100, 22),
                FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8)
            };
            popup.Controls.Add(btnBrowse);
            py += 28;

            // BCC summary
            var lblBccNote = new Label
            {
                Text = $"BCC tới {selectedEmails.Count} nhà cung cấp: {string.Join(", ", selectedNames.Take(4))}{(selectedNames.Count > 4 ? $" ... (+{selectedNames.Count - 4})" : "")}",
                Location = new Point(10, py), Size = new Size(650, 18),
                ForeColor = Color.FromArgb(0, 100, 180),
                Font = new Font("Segoe UI", 8.5f, FontStyle.Italic)
            };
            popup.Controls.Add(lblBccNote); py += 26;

            var btnOpen  = new Button { Text = "📧 Mở trong Outlook", Location = new Point(10, py), Size = new Size(170, 34), BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Cursor = Cursors.Hand };
            var btnClose = new Button { Text = "Huỷ", Location = new Point(580, py), Size = new Size(80, 34), FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnOpen.FlatAppearance.BorderSize = 0;
            popup.Controls.AddRange(new Control[] { btnOpen, btnClose });
            popup.ClientSize = new Size(670, py + 54);

            // Cho phép duyệt file thủ công
            string resolvedAttach = attachPath;
            btnBrowse.Click += (s, ev) =>
            {
                using var ofd = new OpenFileDialog
                {
                    Title  = "Chọn file MPR đính kèm",
                    Filter = "Excel / PDF|*.xlsx;*.xls;*.pdf|Tất cả|*.*",
                    InitialDirectory = string.IsNullOrEmpty(resolvedAttach) ? "" : Path.GetDirectoryName(resolvedAttach) ?? ""
                };
                if (ofd.ShowDialog() != DialogResult.OK) return;
                resolvedAttach     = ofd.FileName;
                lblAttach.Text     = $"📎  {Path.GetFileName(resolvedAttach)}";
                lblAttach.ForeColor = Color.FromArgb(0, 120, 60);
            };

            btnClose.Click += (s, ev) => popup.Close();

            btnOpen.Click += (s, ev) =>
            {
                string bcc = txtBcc.Text.Trim();
                if (string.IsNullOrWhiteSpace(bcc))
                {
                    MessageBox.Show(TopOwner, "Chưa có địa chỉ BCC!", "Thiếu email", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                btnOpen.Enabled = false;
                btnOpen.Text    = "Đang mở Outlook...";
                try
                {
                    var outlookType = Type.GetTypeFromProgID("Outlook.Application");
                    if (outlookType == null) throw new Exception("Không tìm thấy Outlook trên máy tính này.");

                    dynamic outlook = Activator.CreateInstance(outlookType);
                    dynamic mail    = outlook.CreateItem(0); // olMailItem

                    mail.BCC     = bcc;
                    mail.Subject = txtSubject.Text.Trim();

                    // Chèn nội dung vào trước chữ ký Outlook
                    dynamic inspector  = mail.GetInspector;
                    string htmlWithSig = (string)(mail.HTMLBody ?? "");
                    string ourHtml     = BuildHtmlBodyMPR(txtBody.Text);
                    if (string.IsNullOrEmpty(htmlWithSig))
                    {
                        mail.HTMLBody = ourHtml;
                    }
                    else
                    {
                        int tagEnd = htmlWithSig.IndexOf('>', htmlWithSig.IndexOf("<body", StringComparison.OrdinalIgnoreCase));
                        mail.HTMLBody = tagEnd >= 0 ? htmlWithSig.Insert(tagEnd + 1, ourHtml) : ourHtml + htmlWithSig;
                    }

                    // Đính kèm file MPR nếu tìm thấy
                    if (!string.IsNullOrEmpty(resolvedAttach) && File.Exists(resolvedAttach))
                        mail.Attachments.Add(resolvedAttach, 1, Type.Missing, Path.GetFileName(resolvedAttach));

                    // Mở cửa sổ soạn thảo Outlook (không gửi ngay)
                    mail.Display(false);

                    // Đóng popup, rồi poll mail.Sent để tự nhận khi Outlook thực sự gửi
                    popup.Close();
                    _StartEmailSentPoller(outlook, mail, txtSubject.Text.Trim(), mpr.MPR_ID, selectedSuppliers);
                }
                catch (Exception ex)
                {
                    btnOpen.Enabled = true;
                    btnOpen.Text    = "📧 Mở trong Outlook";
                    MessageBox.Show(TopOwner, $"Lỗi mở Outlook:\n{ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            popup.ShowDialog(this);
        }

        // =====================================================================
        //  POLL mail.Sent — tự cập nhật status khi Outlook thực sự bấm Send
        //  Fallback: khi COM object bị giải phóng (sau khi gửi) → check Sent Items
        // =====================================================================
        private void _StartEmailSentPoller(dynamic outlook, dynamic mail, string subject, int mprId,
            List<MPR_Managerment.Models.Supplier> suppliers = null)
        {
            if (_emailPollers.TryGetValue(mprId, out var old))
            {
                old.Stop(); old.Dispose();
                _emailPollers.Remove(mprId);
            }

            if (lblStatus != null)
                lblStatus.Text = $"⏳ Đang chờ Outlook gửi email cho MPR #{mprId}...";

            var startedAt = DateTime.Now;
            var timer     = new System.Windows.Forms.Timer { Interval = 2000 };
            var deadline  = DateTime.Now.AddMinutes(15);
            bool comDead  = false;   // chuyển sang chế độ scan Sent Items
            _emailPollers[mprId] = timer;

            timer.Tick += (_, __) =>
            {
                if (DateTime.Now > deadline)
                {
                    timer.Stop(); timer.Dispose();
                    _emailPollers.Remove(mprId);
                    if (lblStatus != null) lblStatus.Text = "⚠ Hết thời gian chờ — email chưa được ghi nhận.";
                    return;
                }

                // Pha 1: COM còn sống → kiểm tra mail.Sent
                if (!comDead)
                {
                    bool sent = false;
                    try { sent = (bool)mail.Sent; }
                    catch { comDead = true; }

                    if (sent)
                    {
                        _MarkEmailDone(mprId, suppliers);
                        timer.Stop(); timer.Dispose();
                        _emailPollers.Remove(mprId);
                        return;
                    }
                }

                // Pha 2: COM đã chết → tiếp tục poll Sent Items mỗi 2s
                // (email có thể còn đang trong Outbox chưa kịp gửi)
                if (comDead)
                {
                    int elapsed = (int)(DateTime.Now - startedAt).TotalSeconds;
                    if (lblStatus != null)
                        lblStatus.Text = $"⏳ Đang kiểm tra Sent Items... ({elapsed}s)";

                    if (_CheckSentItemsFolder(subject))
                    {
                        _MarkEmailDone(mprId, suppliers);
                        timer.Stop(); timer.Dispose();
                        _emailPollers.Remove(mprId);
                    }
                    // Chưa thấy → giữ timer tiếp tục, không dừng
                    return;
                }

                int sec = (int)(DateTime.Now - startedAt).TotalSeconds;
                if (lblStatus != null)
                    lblStatus.Text = $"⏳ Chờ Outlook gửi... ({sec}s)";
            };

            timer.Start();
        }

        private void _MarkEmailDone(int mprId,
            List<MPR_Managerment.Models.Supplier> suppliers = null)
        {
            var sentAt = DateTime.Now;
            _service.UpdateEmailStatus(mprId, "done", _currentUser, sentAt);

            var h = _mprList.Find(x => x.MPR_ID == mprId);
            if (h != null)
            {
                h.Email_Status  = "done";
                h.Email_Sent_By = _currentUser;
                h.Email_Sent_At = sentAt;
            }

            // Dùng Dispatcher hoặc BeginInvoke để tránh lặp vô tận trong sự kiện UI
            if (this.IsHandleCreated)
            {
                this.BeginInvoke(new Action(() =>
                {
                    BindMPRGrid(_mprList);
                    if (lblStatus != null && !lblStatus.IsDisposed)
                        lblStatus.Text = "✅ Email MPR đã được gửi thành công!";
                }));
            }

            // Gửi thông báo Zalo vào nhóm của từng NCC
            if (suppliers != null && suppliers.Count > 0 && h != null)
            {
                string zaloMsg =
                    $"📋 DLHI Thông báo:\n" +
                    $"Chúng tôi đã gửi yêu cầu báo giá cho:\n" +
                    $"🔹 Số MPR   : {h.MPR_No}\n" +
                    $"🔹 Dự án   : {h.Project_Name}\n" +
                    $"🔹 Thời gian: {sentAt:dd/MM/yyyy HH:mm}\n" +
                    $"Xin vui lòng kiểm tra Email, Rất mong sớm nhận được báo giá từ Quý công ty!";
                _SendZaloMPRNotify(suppliers, zaloMsg, h.MPR_No);
            }
        }

        private void _SendZaloMPRNotify(
            List<MPR_Managerment.Models.Supplier> suppliers, string msg, string mprNo = "")
        {
            void SetSt(string text) { if (lblStatus != null) lblStatus.Text = text; }

            var cfg = Helpers.ZaloHelper.LoadSettings();
            if (!cfg.Enabled)
            {
                SetSt("✅ Email MPR đã gửi  |  ⚠ Zalo chưa bật — vào 🔔 Zalo trên tiêu đề để bật");
                return;
            }

            var targets = suppliers.FindAll(s => !string.IsNullOrWhiteSpace(s.Zalo_Group_ID));
            if (targets.Count == 0)
            {
                SetSt("✅ Email MPR đã gửi  |  ⚠ Không có NCC nào có tên nhóm Zalo");
                return;
            }

            SetSt($"✅ Email MPR đã gửi  |  ⏳ Đang đặt {targets.Count} tin nhắn Zalo vào hàng đợi...");

            foreach (var sup in targets)
            {
                // Lấy tên nhóm từ Zalo_Group_ID (giữ nguyên để dễ debug)
                string groupName = sup.Zalo_Group_ID ?? sup.Company_Name ?? "Unknown";
                ZaloNotificationService.Instance.Enqueue(groupName, msg, $"MPR: {mprNo}");
            }

            SetSt($"✅ Email MPR đã gửi  |  🔔 Đã đặt {targets.Count} tin nhắn Zalo vào hàng đợi (sẽ gửi theo thứ tự)");
        }

        private bool _CheckSentItemsFolder(string subject)
        {
            try
            {
                // Tạo instance Outlook mới (singleton COM — trả về Outlook đang chạy)
                var outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null) return false;
                dynamic ol = Activator.CreateInstance(outlookType);
                dynamic ns = ol.GetNamespace("MAPI");

                // Duyệt tất cả store (account) — tránh bỏ sót khi Outlook có nhiều tài khoản
                try
                {
                    dynamic stores     = ns.Stores;
                    int     storeCount = Convert.ToInt32(stores.Count);
                    for (int s = 1; s <= storeCount; s++)
                    {
                        try
                        {
                            dynamic store = stores[s];
                            dynamic sentFolder;
                            try { sentFolder = store.GetDefaultFolder(5); } catch { continue; }
                            if (_ScanFolderForSubject(sentFolder, subject)) return true;
                        }
                        catch { }
                    }
                }
                catch { }

                // Fallback: thư mục Sent Items mặc định của namespace
                try
                {
                    dynamic sentFolder = ns.GetDefaultFolder(5);
                    if (_ScanFolderForSubject(sentFolder, subject)) return true;
                }
                catch { }

                return false;
            }
            catch { return false; }
        }

        private bool _ScanFolderForSubject(dynamic folder, string subject)
        {
            try
            {
                dynamic items = folder.Items;
                items.Sort("[SentOn]", true); // mới nhất trước
                int count = Convert.ToInt32(items.Count);
                for (int i = 1; i <= Math.Min(100, count); i++)
                {
                    try
                    {
                        dynamic item = items[i];
                        string  itemSubj;
                        try { itemSubj = (string)item.Subject; } catch { continue; }
                        if (string.Equals(itemSubj, subject, StringComparison.OrdinalIgnoreCase)) return true;
                    }
                    catch { }
                }
                return false;
            }
            catch { return false; }
        }

        // =====================================================================
        //  TÌM FILE MPR TRONG THƯ MỤC MPR_Link (khớp MPR_No, ưu tiên PDF)
        // =====================================================================
        private string FindMPRFile(MPRHeader mpr)
        {
            try
            {
                var projects = new ProjectService().GetAll();
                var prj = projects.Find(p =>
                    (!string.IsNullOrEmpty(p.ProjectName) && p.ProjectName.Equals(mpr.Project_Name, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(p.ProjectCode) && p.ProjectCode.Equals(mpr.Project_Code, StringComparison.OrdinalIgnoreCase)));

                if (prj == null || string.IsNullOrEmpty(prj.MPR_Link)) return null;
                string folder = prj.MPR_Link.Trim();
                if (!Directory.Exists(folder)) return null;

                string mprNorm = NormalizeName(mpr.MPR_No);
                var candidates = Directory.GetFiles(folder, "*.*", SearchOption.TopDirectoryOnly)
                    .Where(f => { string ext = Path.GetExtension(f).ToLower(); return ext == ".xlsx" || ext == ".xls" || ext == ".pdf"; })
                    .ToList();

                string bestPath  = null;
                int    bestScore = -1;
                foreach (var f in candidates)
                {
                    string nameNorm = NormalizeName(Path.GetFileNameWithoutExtension(f));
                    int score = 0;
                    if (!string.IsNullOrEmpty(mprNorm) && nameNorm.Contains(mprNorm)) score += 4;
                    if (Path.GetExtension(f).ToLower() == ".pdf")  score += 2;
                    if (Path.GetExtension(f).ToLower() == ".xlsx") score += 1;
                    if (score > bestScore) { bestScore = score; bestPath = f; }
                }

                return bestScore >= 4 ? bestPath : null;
            }
            catch { return null; }
        }

        private static string NormalizeName(string s) =>
            (s ?? "").ToUpperInvariant().Replace(" ", "").Replace("-", "").Replace("_", "").Replace(".", "");

        private static string BuildHtmlBodyMPR(string plainText)
        {
            string html = plainText
                .Replace("&",  "&amp;")
                .Replace("<",  "&lt;")
                .Replace(">",  "&gt;")
                .Replace("\r\n", "<br>")
                .Replace("\n",   "<br>")
                .Replace("  ",  "&nbsp;&nbsp;");
            return $"<div style='font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000000;'>{html}</div><br>";
        }
    }
}