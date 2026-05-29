using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using MPR_Managerment.Services;

namespace MPR_Managerment.Forms
{
    // ════════════════════════════════════════════════════════════════════════
    // frmAIToggleBtn — Form nổi chứa icon robot 3D AI — có thể kéo tùy ý
    // Hiển thị full icon 128x128 PNG trên app
    // ════════════════════════════════════════════════════════════════════════
    public class frmAIToggleBtn : Form
    {
        private const int ICON_SIZE = 72;   // Kích thước hiển thị icon trên form
        private const int MARGIN = 18;

        private PictureBox _pic;
        private Label _lblClose;            // Dấu X khi chat mở
        private Action _onToggle;
        private Image _robotIcon;
        private bool _chatOpen = false;

        public frmAIToggleBtn(Action onToggle)
        {
            _onToggle = onToggle;

            // ── Load icon robot 3D từ Resources ──
            _robotIcon = LoadRobotIcon();

            // ── Form trong suốt, không viền ──
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            Size = new Size(ICON_SIZE, ICON_SIZE);
            BackColor = Color.Magenta;
            TransparencyKey = Color.Magenta;

            // ── PictureBox hiển thị icon robot 3D ──
            _pic = new PictureBox
            {
                Size = new Size(ICON_SIZE, ICON_SIZE),
                Location = new Point(0, 0),
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.Transparent,
                Cursor = Cursors.Hand,
                Image = _robotIcon
            };
            _pic.Click += (s, e) => _onToggle?.Invoke();

            // ── Label "✕" hiển thị khi chat đang mở ──
            _lblClose = new Label
            {
                Size = new Size(22, 22),
                Location = new Point(ICON_SIZE - 22, 0),
                BackColor = Color.FromArgb(220, 50, 50),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Text = "✕",
                TextAlign = ContentAlignment.MiddleCenter,
                Visible = false,
                Cursor = Cursors.Hand
            };
            _lblClose.Click += (s, e) => _onToggle?.Invoke();

            // ── Kéo icon để đổi vị trí (drag) ──
            bool dragging = false;
            Point dragStart = Point.Empty;
            _pic.MouseDown += (s, e) =>
            {
                if (e.Button == MouseButtons.Left) { dragging = true; dragStart = e.Location; }
            };
            _pic.MouseMove += (s, e) =>
            {
                if (dragging)
                    Location = new Point(Location.X + e.X - dragStart.X,
                                         Location.Y + e.Y - dragStart.Y);
            };
            _pic.MouseUp += (s, e) => dragging = false;

            // ── Hiệu ứng hover: phóng to nhẹ ──
            _pic.MouseEnter += (s, e) =>
            {
                _pic.Size = new Size(ICON_SIZE + 6, ICON_SIZE + 6);
                _pic.Location = new Point(-3, -3);
            };
            _pic.MouseLeave += (s, e) =>
            {
                _pic.Size = new Size(ICON_SIZE, ICON_SIZE);
                _pic.Location = new Point(0, 0);
            };

            Controls.Add(_lblClose);
            Controls.Add(_pic);
            _lblClose.BringToFront();
        }

        /// <summary>Đổi trạng thái icon: chat mở → hiện X, chat đóng → hiện robot</summary>
        public void SetIcon(bool chatOpen)
        {
            _chatOpen = chatOpen;
            _lblClose.Visible = chatOpen;
        }

        public void SnapToCorner()
        {
            var scr = Screen.PrimaryScreen.WorkingArea;
            Location = new Point(scr.Right - ICON_SIZE - MARGIN,
                                 scr.Bottom - ICON_SIZE - MARGIN);
        }

        /// <summary>Load icon robot 3D từ file Resources/ai_robot_icon.png</summary>
        private Image LoadRobotIcon()
        {
            try
            {
                // Tìm file icon trong thư mục Resources cạnh exe
                string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string iconPath = Path.Combine(exeDir, "Resources", "ai_robot_icon.png");

                // Nếu không có cạnh exe, thử tìm từ thư mục gốc project
                if (!File.Exists(iconPath))
                {
                    string projectDir = exeDir;
                    // Lùi lên tìm thư mục Resources (bin/Debug/net... → project root)
                    for (int i = 0; i < 5; i++)
                    {
                        projectDir = Path.GetDirectoryName(projectDir);
                        if (projectDir == null) break;
                        string tryPath = Path.Combine(projectDir, "Resources", "ai_robot_icon.png");
                        if (File.Exists(tryPath)) { iconPath = tryPath; break; }
                    }
                }

                if (File.Exists(iconPath))
                    return Image.FromFile(iconPath);
            }
            catch { /* fallback bên dưới */ }

            // Fallback: vẽ icon đơn giản nếu không tìm thấy file
            return CreateFallbackIcon();
        }

        /// <summary>Tạo icon fallback nếu không load được file PNG</summary>
        private Image CreateFallbackIcon()
        {
            var bmp = new Bitmap(ICON_SIZE, ICON_SIZE);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                // Vẽ hình tròn gradient nền
                using (var brush = new LinearGradientBrush(
                    new Rectangle(0, 0, ICON_SIZE, ICON_SIZE),
                    Color.FromArgb(99, 102, 241), Color.FromArgb(60, 63, 180),
                    LinearGradientMode.ForwardDiagonal))
                {
                    g.FillEllipse(brush, 4, 4, ICON_SIZE - 8, ICON_SIZE - 8);
                }

                // Vẽ đầu robot đơn giản
                int cx = ICON_SIZE / 2, cy = ICON_SIZE / 2;
                // Đầu
                using (var headBrush = new SolidBrush(Color.FromArgb(140, 145, 255)))
                    g.FillRoundedRectangle(headBrush, cx - 18, cy - 16, 36, 28, 8);
                // Mắt trái
                g.FillEllipse(Brushes.White, cx - 12, cy - 8, 10, 10);
                g.FillEllipse(Brushes.Black, cx - 9, cy - 5, 4, 4);
                // Mắt phải
                g.FillEllipse(Brushes.White, cx + 2, cy - 8, 10, 10);
                g.FillEllipse(Brushes.Black, cx + 5, cy - 5, 4, 4);
                // Miệng
                using (var pen = new Pen(Color.FromArgb(0, 220, 255), 2))
                    g.DrawLine(pen, cx - 8, cy + 8, cx + 8, cy + 8);
                // Antenna
                using (var pen = new Pen(Color.FromArgb(180, 185, 210), 2))
                    g.DrawLine(pen, cx, cy - 16, cx, cy - 24);
                g.FillEllipse(Brushes.Orange, cx - 4, cy - 28, 8, 8);
            }
            return bmp;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) _robotIcon?.Dispose();
            base.Dispose(disposing);
        }
    }

    // Extension method để vẽ rounded rectangle (cho fallback icon)
    internal static class GraphicsExtensions
    {
        public static void FillRoundedRectangle(this Graphics g, Brush brush,
            float x, float y, float w, float h, float r)
        {
            using (var path = new GraphicsPath())
            {
                path.AddArc(x, y, r * 2, r * 2, 180, 90);
                path.AddArc(x + w - r * 2, y, r * 2, r * 2, 270, 90);
                path.AddArc(x + w - r * 2, y + h - r * 2, r * 2, r * 2, 0, 90);
                path.AddArc(x, y + h - r * 2, r * 2, r * 2, 90, 90);
                path.CloseFigure();
                g.FillPath(brush, path);
            }
        }
    }


    // ════════════════════════════════════════════════════════════════════════
    // frmAIChat — Panel chat AI chính
    // Form bình thường (không WS_EX_TRANSPARENT) → input hoạt động hoàn toàn
    // ════════════════════════════════════════════════════════════════════════
    public class frmAIChat : Form
    {
        // ── Singleton ──────────────────────────────────────────────────────
        private static frmAIChat _chatInst;
        private static frmAIToggleBtn _btnInst;

        /// <summary>Gắn AI Chat vào form — gọi trong constructor của mọi Form.</summary>
        public static void Attach(Form parentForm)
        {
            parentForm.Shown += (s, e) =>
            {
                // Nút toggle
                if (_btnInst == null || _btnInst.IsDisposed)
                {
                    _btnInst = new frmAIToggleBtn(ToggleChat);
                    _btnInst.Show(parentForm);
                    _btnInst.SnapToCorner();
                }
                // Panel chat
                if (_chatInst == null || _chatInst.IsDisposed)
                {
                    _chatInst = new frmAIChat();
                    _chatInst.Show(parentForm);
                    _chatInst.Visible = false;
                }
            };
        }

        private static void ToggleChat()
        {
            if (_chatInst == null || _chatInst.IsDisposed) return;

            if (_chatInst.Visible)
            {
                _chatInst.Hide();
                _btnInst?.SetIcon(false);
            }
            else
            {
                _chatInst.PositionNearButton();
                _chatInst.Show();
                _chatInst.BringToFront();
                _chatInst._txtInput?.Focus();
                _btnInst?.SetIcon(true);
            }
        }

        // ── Cấu hình ─────────────────────────────────────────────────────
        // Ollama chạy local — không cần API key
        // Chỉnh model/URL trong AISearchService.cs: OLLAMA_MODEL, OLLAMA_BASE_URL

        private const int PANEL_W = 420;
        private const int PANEL_H = 580;
        private const int BTN_SIZE = 52;
        private const int MARGIN = 20;

        // ── Màu ──────────────────────────────────────────────────────────
        private static readonly Color C_BG = Color.FromArgb(18, 20, 28);
        private static readonly Color C_PANEL = Color.FromArgb(26, 30, 42);
        private static readonly Color C_HEADER = Color.FromArgb(35, 40, 58);
        private static readonly Color C_INPUT_BG = Color.FromArgb(44, 50, 72);
        private static readonly Color C_SEND = Color.FromArgb(99, 102, 241);
        private static readonly Color C_SEND_HOV = Color.FromArgb(79, 82, 221);
        private static readonly Color C_TEXT = Color.FromArgb(226, 232, 240);
        private static readonly Color C_SUBTEXT = Color.FromArgb(148, 163, 184);
        private static readonly Color C_USER_MSG = Color.FromArgb(129, 140, 248);
        private static readonly Color C_AI_MSG = Color.FromArgb(52, 211, 153);

        // ── Controls ─────────────────────────────────────────────────────
        private RichTextBox _rtbChat;
        private TextBox _txtInput;
        private Button _btnSend;
        private Button _btnExport;   // Tính năng 3: Xuất Excel
        private Label _lblStatus;

        // ── Service ──────────────────────────────────────────────────────
        internal readonly AISearchService _ai = new AISearchService();
        private bool _isThinking = false;
        private bool _isDbMode = true;
        private string _lastQuestion = "";
        private Microsoft.Web.WebView2.WinForms.WebView2 _webView;
        private Panel _inputPanel; // field để toggle mode có thể truy cập

        // ── Tính năng 2: Bookmark query ───────────────────────────────────
        private bool _summaryShown = false;
        private Panel _skillStrip;
        private ListBox _slashPopup;
        private Label _slashHeader;        // tiêu đề popup level 2
        private string _slashLevel2Parent;    // lệnh cha đang mở level 2 (null = level 1)
        private bool   _slashLevel3Active;    // đang ở level 3 (chọn sub-skill sau khi chọn dự án)
        private string _slashSelectedProject; // mã dự án đã chọn ở level 2 (dùng cho display text)
        private string _pendingDisplayText;   // text ngắn gọn hiện trong [Bạn] thay vì full template

        // ─────────────────────────────────────────────────────────────────
        public frmAIChat()
        {
            FormBorderStyle = FormBorderStyle.None;
            BackColor = C_PANEL;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            Size = new Size(PANEL_W, PANEL_H);

            // Viền bo tròn nhẹ
            this.Paint += (s, e) =>
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using var pen = new Pen(Color.FromArgb(80, 99, 102, 241), 1.5f);
                using var path = RoundedPath(new Rectangle(0, 0, Width - 1, Height - 1), 12);
                e.Graphics.DrawPath(pen, path);
            };

            BuildUI();
            RegisterHotkey();
            AISearchService.EnsureSkillTableAndSeed();
            RefreshSkillStrip();
            ShowGreeting();
        }

        // ════════════════════════════════════════════════════════════════════
        // BUILD UI
        // ════════════════════════════════════════════════════════════════════

        private void BuildUI()
        {
            // ── Header ────────────────────────────────────────────────────
            var header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 54,
                BackColor = C_HEADER
            };

            header.Controls.Add(new Label
            {
                Text = "✦",
                Font = new Font("Segoe UI Emoji", 15),
                ForeColor = C_SEND,
                Location = new Point(14, 13),
                Size = new Size(28, 28),
                AutoSize = false
            });
            header.Controls.Add(new Label
            {
                Text = "AI Trợ lý",
                Font = new Font("Segoe UI Semibold", 11),
                ForeColor = C_TEXT,
                Location = new Point(46, 8),
                Size = new Size(180, 20),
                AutoSize = false
            });
            _lblStatus = new Label
            {
                Text = "Sẵn sàng  •  Hỏi về MPR, PO, NCC, thanh toán...",
                Font = new Font("Segoe UI", 7.5f),
                ForeColor = C_SUBTEXT,
                Location = new Point(46, 30),
                Size = new Size(262, 16),
                AutoSize = false
            };
            header.Controls.Add(_lblStatus);

            // Nút xóa lịch sử
            var btnClear = HeaderBtn("🗑", PANEL_W - 80);
            new ToolTip().SetToolTip(btnClear, "Xóa lịch sử hội thoại");
            btnClear.Click += (s, e) =>
            {
                _rtbChat.Clear();
                _ai.ClearHistory();
                _summaryShown = false;
                ShowGreeting();
            };

            // Nút Bookmark
            var btnBookmark = HeaderBtn("⭐", PANEL_W - 116);
            new ToolTip().SetToolTip(btnBookmark, "Bookmark & tra cứu nhanh");
            btnBookmark.Click += (s, e) => ShowBookmarkMenu();

            // Nút Skill
            var btnSkill = HeaderBtn("⚡", PANEL_W - 152);
            new ToolTip().SetToolTip(btnSkill, "Quản lý AI Skill");
            btnSkill.Click += (s, e) => ShowSkillMenu();

            // Nút bộ nhớ AI
            var btnMemory = HeaderBtn("🧠", PANEL_W - 188);
            new ToolTip().SetToolTip(btnMemory, "Xem / quản lý bộ nhớ AI");
            btnMemory.Click += (s, e) => ShowMemoryMenu();

            // Nút đóng
            var btnClose = HeaderBtn("×", PANEL_W - 44);
            btnClose.Font = new Font("Segoe UI", 14);
            new ToolTip().SetToolTip(btnClose, "Đóng  (Ctrl+Space)");
            btnClose.Click += (s, e) => ToggleChat();

            header.Controls.AddRange(new Control[] { btnClear, btnSkill, btnMemory, btnBookmark, btnClose });

            // Kéo panel qua header
            bool drag = false; Point dragPt = default;
            header.MouseDown += (s, e) => { drag = true; dragPt = e.Location; };
            header.MouseMove += (s, e) =>
            {
                if (drag) Location = new Point(Location.X + e.X - dragPt.X,
                                                Location.Y + e.Y - dragPt.Y);
            };
            header.MouseUp += (s, e) => drag = false;

            // ── Mode selector — Tab to rõ ràng ngay dưới header ─────────
            var modePanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 38,
                BackColor = Color.FromArgb(22, 27, 42),
                Padding = new Padding(8, 5, 8, 5)
            };

            var btnDbMode = new Button
            {
                Location = new Point(8, 5),
                Size = new Size(188, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                Font = new Font("Segoe UI Semibold", 9.5f),
                Text = "🗄  Truy xuất dữ liệu",
                Cursor = Cursors.Hand,
                TabStop = false,
                Name = "btnDbMode"
            };
            btnDbMode.FlatAppearance.BorderSize = 0;
            btnDbMode.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 100, 190);

            var btnChatMode = new Button
            {
                Location = new Point(202, 5),
                Size = new Size(188, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(40, 46, 66),
                ForeColor = Color.FromArgb(148, 163, 184),
                Font = new Font("Segoe UI Semibold", 9.5f),
                Text = "💬  Chat tự do",
                Cursor = Cursors.Hand,
                TabStop = false,
                Name = "btnChatMode"
            };
            btnChatMode.FlatAppearance.BorderSize = 0;
            btnChatMode.FlatAppearance.MouseOverBackColor = Color.FromArgb(55, 60, 82);

            // Toggle logic
            btnDbMode.Click += (s, e) =>
            {
                if (_isDbMode) return;
                _isDbMode = true;
                btnDbMode.BackColor = Color.FromArgb(0, 120, 212);
                btnDbMode.ForeColor = Color.White;
                btnChatMode.BackColor = Color.FromArgb(40, 46, 66);
                btnChatMode.ForeColor = Color.FromArgb(148, 163, 184);
                _lblStatus.Text = "DB Mode  •  Truy xuất MPR, PO, NCC, thanh toán...";
                _txtInput.PlaceholderText = "Hỏi về MPR, PO, NCC... (Enter gửi | Shift+Enter xuống dòng)";
                _webView.Visible = false;
                _rtbChat.Visible = true;
                if (_inputPanel != null) _inputPanel.Visible = true;
                AppendMsg("Hệ thống",
                    "🗄 Đã chuyển sang **DB Mode** — tôi truy vấn dữ liệu thực từ hệ thống.",
                    Color.FromArgb(148, 163, 184));
                _ai.ClearHistory();
                _txtInput.Focus();
            };

            btnChatMode.Click += (s, e) =>
            {
                if (!_isDbMode) return;
                _isDbMode = false;
                btnChatMode.BackColor = Color.FromArgb(16, 163, 127);
                btnChatMode.ForeColor = Color.White;
                btnDbMode.BackColor = Color.FromArgb(40, 46, 66);
                btnDbMode.ForeColor = Color.FromArgb(148, 163, 184);
                _lblStatus.Text = "Chat Mode  •  ChatGPT (chat.openai.com)";
                _rtbChat.Visible = false;
                if (_inputPanel != null) _inputPanel.Visible = false;
                _webView.Visible = true;
                _webView.BringToFront();
            };

            modePanel.Controls.AddRange(new Control[] { btnDbMode, btnChatMode });

            // ── Chat area ────────────────────────────────────────────────
            _rtbChat = new RichTextBox
            {
                Dock = DockStyle.Fill,
                BackColor = C_BG,
                ForeColor = C_TEXT,
                Font = new Font("Segoe UI", 10),
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                Padding = new Padding(14, 10, 14, 10)
            };

            var chatWrap = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = C_BG,
                Padding = new Padding(10, 6, 10, 6)
            };
            chatWrap.Controls.Add(_rtbChat);

            // ── Input area ───────────────────────────────────────────────
            _inputPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 104,
                BackColor = C_HEADER,
                Padding = new Padding(10, 8, 10, 10)
            };

            _txtInput = new TextBox
            {
                Multiline = true,
                Location = new Point(10, 10),
                Size = new Size(PANEL_W - 72, 60),
                BackColor = C_INPUT_BG,
                ForeColor = C_TEXT,
                Font = new Font("Segoe UI", 10),
                BorderStyle = BorderStyle.FixedSingle,
                PlaceholderText = "Hỏi về MPR, PO, NCC... (Enter gửi | Shift+Enter xuống dòng)",
                TabIndex = 0
            };
            _txtInput.KeyDown += (s, e) =>
            {
                if (_slashPopup.Visible)
                {
                    if (e.KeyCode == Keys.Down)
                    {
                        e.SuppressKeyPress = true;
                        if (_slashPopup.SelectedIndex < _slashPopup.Items.Count - 1)
                            _slashPopup.SelectedIndex++;
                        return;
                    }
                    if (e.KeyCode == Keys.Up)
                    {
                        e.SuppressKeyPress = true;
                        if (_slashPopup.SelectedIndex > 0)
                            _slashPopup.SelectedIndex--;
                        return;
                    }
                    if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
                    {
                        e.SuppressKeyPress = true;
                        ApplySlashPopupSelection();
                        return;
                    }
                    if (e.KeyCode == Keys.Escape)
                    {
                        e.SuppressKeyPress = true;
                        HideSlashPopup();
                        return;
                    }
                }

                if (e.KeyCode == Keys.Enter && !e.Shift)
                {
                    e.SuppressKeyPress = true;
                    SendMessage();
                }
            };
            _txtInput.TextChanged += (s, e) => UpdateSlashPopup();

            // ── Slash popup — ListBox nổi hiển thị khi gõ "/" ────────────
            _slashPopup = new ListBox
            {
                Visible = false,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.FromArgb(35, 40, 58),
                ForeColor = Color.FromArgb(226, 232, 240),
                Font = new Font("Segoe UI", 9.5f),
                SelectionMode = SelectionMode.One,
                ItemHeight = 26,
                DrawMode = DrawMode.OwnerDrawFixed,
                IntegralHeight = false
            };
            _slashPopup.DrawItem += (s, e) =>
            {
                if (e.Index < 0) return;
                bool sel = (e.State & DrawItemState.Selected) != 0;
                e.Graphics.FillRectangle(
                    sel ? new SolidBrush(Color.FromArgb(99, 102, 241))
                        : new SolidBrush(Color.FromArgb(35, 40, 58)),
                    e.Bounds);
                e.Graphics.DrawString(
                    _slashPopup.Items[e.Index].ToString(),
                    e.Font, new SolidBrush(Color.FromArgb(226, 232, 240)),
                    e.Bounds.Left + 6, e.Bounds.Top + 4);
            };
            _slashPopup.MouseClick += (s, e) => ApplySlashPopupSelection();
            _slashPopup.MouseMove += (s, e) =>
            {
                int idx = _slashPopup.IndexFromPoint(e.Location);
                if (idx >= 0) _slashPopup.SelectedIndex = idx;
            };

            _btnSend = new Button
            {
                Location = new Point(PANEL_W - 58, 12),
                Size = new Size(46, 56),
                FlatStyle = FlatStyle.Flat,
                BackColor = C_SEND,
                ForeColor = Color.White,
                Text = "➤",
                Font = new Font("Segoe UI", 14),
                Cursor = Cursors.Hand,
                TabStop = false
            };
            _btnSend.FlatAppearance.BorderSize = 0;
            _btnSend.FlatAppearance.MouseOverBackColor = C_SEND_HOV;
            _btnSend.Click += (s, e) => SendMessage();

            // Nút xuất Excel — chỉ hiện khi có dữ liệu
            _btnExport = new Button
            {
                Location = new Point(10, 68),
                Size = new Size(PANEL_W - 20, 24),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(34, 197, 94),
                ForeColor = Color.White,
                Text = "📥 Xuất Excel kết quả vừa tra cứu",
                Font = new Font("Segoe UI", 8.5f),
                Cursor = Cursors.Hand,
                TabStop = false,
                Visible = false
            };
            _btnExport.FlatAppearance.BorderSize = 0;
            _btnExport.FlatAppearance.MouseOverBackColor = Color.FromArgb(22, 163, 74);
            _btnExport.Click += (s, e) => ExportToExcel();

            _inputPanel.Controls.AddRange(new Control[] { _txtInput, _btnSend, _btnExport });

            // Header label hiển thị trên popup khi ở Level 2
            _slashHeader = new Label
            {
                Visible = false,
                BackColor = Color.FromArgb(60, 65, 95),
                ForeColor = Color.FromArgb(199, 210, 254),
                Font = new Font("Segoe UI Semibold", 8.5f),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 0, 0),
                Height = 22
            };

            // Popup nổi — thêm vào form chính để có thể vẽ đè lên inputPanel
            Controls.Add(_slashHeader);
            Controls.Add(_slashPopup);

            // ── Skill Strip — hàng chip nút nhanh ────────────────────────
            _skillStrip = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 38,
                BackColor = Color.FromArgb(22, 27, 42),
                Padding = new Padding(6, 4, 6, 4),
                Visible = false
            };

            // ── Ghép lại ────────────────────────────────────────────────
            Controls.Add(chatWrap);
            Controls.Add(_skillStrip);
            Controls.Add(_inputPanel);
            Controls.Add(modePanel);
            Controls.Add(header);

            // ── WebView2 — hiển thị ChatGPT khi ở Chat Mode ──────────────
            _webView = new Microsoft.Web.WebView2.WinForms.WebView2
            {
                Dock = DockStyle.Fill,
                Visible = false
            };
            // Chèn WebView vào chatWrap (cùng vị trí với _rtbChat)
            chatWrap.Controls.Add(_webView);
            _webView.BringToFront();

            // Khởi tạo WebView2 async — không block UI
            _ = InitWebViewAsync();
        }

        private Button HeaderBtn(string text, int x)
        {
            var b = new Button
            {
                Text = text,
                Font = new Font("Segoe UI Emoji", 10),
                ForeColor = C_SUBTEXT,
                FlatStyle = FlatStyle.Flat,
                Location = new Point(x, 12),
                Size = new Size(32, 30),
                BackColor = Color.Transparent,
                Cursor = Cursors.Hand,
                TabStop = false
            };
            b.FlatAppearance.BorderSize = 0;
            b.FlatAppearance.MouseOverBackColor = Color.FromArgb(50, 250, 80, 80);
            return b;
        }

        // ════════════════════════════════════════════════════════════════════
        // SEND / RECEIVE (Streaming)
        // ════════════════════════════════════════════════════════════════════

        private async void SendMessage()
        {
            if (_isThinking) return;
            string q = _txtInput.Text.Trim();
            if (string.IsNullOrEmpty(q)) return;

            _lastQuestion = q;
            _txtInput.Clear();
            _txtInput.Focus();
            _btnExport.Visible = false;

            // Dùng display text ngắn gọn (tên skill) nếu có, ngược lại dùng q gốc
            string chatDisplay   = _pendingDisplayText ?? q;
            _pendingDisplayText  = null; // clear sau khi dùng
            AppendMsg("Bạn", chatDisplay, C_USER_MSG);

            // ── Xử lý lệnh ghi nhớ: "nhớ: ..." ──────────────────────────
            string qLower = q.ToLower();
            string[] memPrefixes = { "nhớ:", "nhớ ", "ghi nhớ:", "remember:" };
            foreach (string prefix in memPrefixes)
            {
                if (qLower.StartsWith(prefix))
                {
                    string rule = q.Substring(prefix.Length).Trim();
                    if (!string.IsNullOrEmpty(rule))
                    {
                        string result = AISearchService.AddMemory(rule);
                        AISearchService.GetMemories(forceRefresh: true); // refresh cache ngay
                        AppendMsg("🧠 Bộ nhớ",
                            result + "\n\nQuy tắc này sẽ được áp dụng vào tất cả câu hỏi tiếp theo.",
                            Color.FromArgb(168, 85, 247));
                    }
                    return;
                }
            }

            // ── Lệnh xem bộ nhớ ──────────────────────────────────────────
            if (qLower == "bộ nhớ" || qLower == "xem bộ nhớ" || qLower == "memory")
            {
                var mems = AISearchService.GetMemories(forceRefresh: true);
                string memList = mems.Count == 0
                    ? "Chưa có quy tắc nào.\n💡 Nói \"nhớ: [quy tắc]\" để thêm."
                    : $"📋 Có {mems.Count} quy tắc đang hoạt động:\n" +
                      string.Join("\n", mems.Select(m => $"  #{m.Id}. {m.Rule}"));
                AppendMsg("🧠 Bộ nhớ", memList, Color.FromArgb(168, 85, 247));
                return;
            }

            // ── Detect /slash_command ─────────────────────────────────────
            if (q.StartsWith("/"))
            {
                string slashCmd = q.Split(' ')[0];
                string param = q.Length > slashCmd.Length ? q.Substring(slashCmd.Length).Trim() : "";
                var skill = AISearchService.GetBySlashCommand(slashCmd);

                // ── /po và /mpr: dùng BuildProjectTemplate với param là project code ──
                if (skill != null && PROJECT_SLASH_COMMANDS.Contains(slashCmd))
                {
                    q = BuildProjectTemplate(slashCmd, param); // param có thể rỗng = All
                    _isDbMode = true;
                }
                else if (skill != null)
                {
                    q = skill.Template.Replace("{param}", param);
                    _isDbMode = true;
                }
            }

            // ── Resolve {project_code} — fallback cho slash/template gõ tay ──
            if (q.Contains("{project_code}"))
            {
                string code = PromptProjectCode();
                if (code == null)
                {
                    AppendMsg("AI", "❌ Đã hủy chọn dự án.", Color.FromArgb(156, 163, 175));
                    return;
                }
                q = ApplyProjectCodeToTemplate(q, code);
            }

            SetThinking(true);
            int thinkStart = _rtbChat.TextLength;
            AppendThinking();

            try
            {
                bool firstChunk = true;
                bool isDb = _isDbMode; // snapshot mode tại thời điểm gửi

                await Task.Run(async () =>
                {
                    // Gọi đúng method theo chế độ hiện tại
                    if (isDb)
                        await _ai.AskAsync(q, chunk =>
                        {
                            if (IsDisposed) return;
                            Invoke((Action)(() =>
                            {
                                if (firstChunk)
                                {
                                    _rtbChat.Select(thinkStart, _rtbChat.TextLength - thinkStart);
                                    _rtbChat.SelectedText = "";
                                    AppendPrefix("AI", C_AI_MSG);
                                    firstChunk = false;
                                }
                                _rtbChat.Select(_rtbChat.TextLength, 0);
                                _rtbChat.SelectionColor = C_TEXT;
                                _rtbChat.SelectionFont = new Font("Segoe UI", 10);
                                _rtbChat.AppendText(chunk);
                                _rtbChat.ScrollToCaret();
                            }));
                        });
                    else
                        await _ai.AskFreeAsync(q, chunk =>
                        {
                            if (IsDisposed) return;
                            Invoke((Action)(() =>
                            {
                                if (firstChunk)
                                {
                                    _rtbChat.Select(thinkStart, _rtbChat.TextLength - thinkStart);
                                    _rtbChat.SelectedText = "";
                                    AppendPrefix("AI", Color.FromArgb(167, 139, 250)); // tím nhạt cho chat mode
                                    firstChunk = false;
                                }
                                _rtbChat.Select(_rtbChat.TextLength, 0);
                                _rtbChat.SelectionColor = C_TEXT;
                                _rtbChat.SelectionFont = new Font("Segoe UI", 10);
                                _rtbChat.AppendText(chunk);
                                _rtbChat.ScrollToCaret();
                            }));
                        });
                });

                if (!IsDisposed)
                    Invoke((Action)(() =>
                    {
                        _rtbChat.AppendText("\n\n");
                        _rtbChat.ScrollToCaret();
                        // Nút Export chỉ hiện ở DB Mode
                        _btnExport.Visible = _isDbMode && _ai.HasExportableData;

                        // Tự động mở file Excel nếu AI vừa tạo tự động
                        if (_ai._pendingExcelPath != null)
                        {
                            try
                            {
                                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                                {
                                    FileName = _ai._pendingExcelPath,
                                    UseShellExecute = true
                                });
                            }
                            catch { }
                            _ai._pendingExcelPath = null;
                            _btnExport.Visible = false; // đã xuất tự động rồi
                        }
                    }));
            }
            catch (Exception ex)
            {
                if (!IsDisposed)
                    Invoke((Action)(() =>
                    {
                        _rtbChat.Select(thinkStart, _rtbChat.TextLength - thinkStart);
                        _rtbChat.SelectedText = "";
                        AppendMsg("AI", $"⚠ Lỗi: {ex.Message}",
                            Color.FromArgb(248, 113, 113));
                    }));
            }
            finally
            {
                if (!IsDisposed)
                    Invoke((Action)(() => SetThinking(false)));
            }
        }

        // ════════════════════════════════════════════════════════════════════
        // RENDERING HELPERS
        // ════════════════════════════════════════════════════════════════════

        private void AppendMsg(string sender, string text, Color color)
        {
            AppendPrefix(sender, color);
            _rtbChat.SelectionColor = C_TEXT;
            _rtbChat.SelectionFont = new Font("Segoe UI", 10);
            _rtbChat.AppendText(text + "\n\n");
            _rtbChat.ScrollToCaret();
        }

        private void AppendPrefix(string sender, Color color)
        {
            _rtbChat.Select(_rtbChat.TextLength, 0);
            _rtbChat.SelectionColor = color;
            _rtbChat.SelectionFont = new Font("Segoe UI Semibold", 9);
            _rtbChat.AppendText($"[{sender}]  ");
        }

        private void AppendThinking()
        {
            _rtbChat.Select(_rtbChat.TextLength, 0);
            _rtbChat.SelectionColor = C_SUBTEXT;
            _rtbChat.SelectionFont = new Font("Segoe UI", 9, FontStyle.Italic);
            _rtbChat.AppendText(_isDbMode ? "AI đang truy vấn dữ liệu...\n" : "AI đang suy nghĩ...\n");
            _rtbChat.SelectionFont = new Font("Segoe UI", 10);
            _rtbChat.ScrollToCaret();
        }

        private async void ShowGreeting()
        {
            AppendMsg("AI",
                "Xin chào! Tôi là trợ lý AI của MPR Management.\n\n" +
                "Bạn có thể hỏi tôi về:\n" +
                "• 📋 Tình trạng MPR / phiếu yêu cầu vật tư\n" +
                "• 🛒 Đơn đặt hàng PO và tiến độ giao hàng\n" +
                "• 🏢 Thông tin nhà cung cấp (NCC)\n" +
                "• 📦 Biên bản nghiệm thu (RIR)\n" +
                "• 💰 Lịch sử thanh toán & Payment Request\n\n" +
                "Ví dụ:\n" +
                "  \"MPR nào chưa có PO?\"\n" +
                "  \"Tổng tiền PO của NCC ABC tháng này?\"\n" +
                "  \"PO nào quá hạn giao hàng?\"",
                C_AI_MSG);

            // Tải báo cáo tóm tắt tự động (1 lần/session, không tốn API quota)
            if (!_summaryShown)
            {
                _summaryShown = true;
                _lblStatus.Text = "⏳ Đang tải báo cáo tóm tắt...";
                string summary = await _ai.GetDailySummaryAsync();
                if (!IsDisposed)
                {
                    AppendMsg("📊 Tóm tắt hôm nay", summary, Color.FromArgb(251, 191, 36));
                    _lblStatus.Text = "Sẵn sàng  •  Hỏi về MPR, PO, NCC, thanh toán...";
                }
            }
        }

        // ── Tính năng 2: Bookmark ────────────────────────────────────────
        private async System.Threading.Tasks.Task InitWebViewAsync()
        {
            try
            {
                await _webView.EnsureCoreWebView2Async();
                _webView.Source = new Uri("https://chat.openai.com");
            }
            catch { /* WebView2 runtime chưa cài — bỏ qua */ }
        }

        private void ShowMemoryMenu()
        {
            var menu = new ContextMenuStrip();
            var mems = AISearchService.GetMemories();

            if (mems.Count == 0)
            {
                menu.Items.Add(new ToolStripMenuItem("(Chưa có quy tắc nào)") { Enabled = false });
            }
            else
            {
                menu.Items.Add(new ToolStripMenuItem($"📋 {mems.Count} quy tắc đang hoạt động:") { Enabled = false });
                menu.Items.Add(new ToolStripSeparator());
                foreach (var mem in mems)
                {
                    int memId = mem.Id;
                    string rule = mem.Rule;
                    string label = rule.Length > 50 ? rule.Substring(0, 47) + "..." : rule;
                    var item = new ToolStripMenuItem($"  #{memId}. {label}");
                    var delItem = new ToolStripMenuItem("🗑 Xóa quy tắc này");
                    delItem.Click += (s, e) =>
                    {
                        string result = AISearchService.RemoveMemory(memId);
                        AppendMsg("🧠 Bộ nhớ", result, Color.FromArgb(168, 85, 247));
                    };
                    item.DropDownItems.Add(delItem);
                    menu.Items.Add(item);
                }
                menu.Items.Add(new ToolStripSeparator());
                var clearAll = new ToolStripMenuItem("🗑 Xóa tất cả quy tắc");
                clearAll.Click += (s, e) =>
                {
                    string result = AISearchService.ClearMemories();
                    AppendMsg("🧠 Bộ nhớ", result, Color.FromArgb(168, 85, 247));
                };
                menu.Items.Add(clearAll);
            }

            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(new ToolStripMenuItem("💡 Nói \"nhớ: [quy tắc]\" để thêm") { Enabled = false });
            menu.Show(this, new Point(PANEL_W - 160, 54));
        }

        private void ShowBookmarkMenu()
        {
            var menu = new ContextMenuStrip();
            menu.BackColor = Color.FromArgb(35, 40, 58);
            menu.ForeColor = Color.FromArgb(226, 232, 240);
            menu.RenderMode = ToolStripRenderMode.System;

            // ── Lưu câu hỏi hiện tại ──
            if (!string.IsNullOrEmpty(_lastQuestion))
            {
                string q = _lastQuestion;
                string label = q.Length > 45 ? q.Substring(0, 42) + "..." : q;
                var save = new ToolStripMenuItem($"⭐ Lưu: \"{label}\"");
                save.Click += (s, e) =>
                {
                    string result = AISearchService.AddBookmark(q);
                    AppendMsg("⭐ Bookmark", result, Color.FromArgb(234, 179, 8));
                };
                menu.Items.Add(save);
                menu.Items.Add(new ToolStripSeparator());
            }

            // ── Danh sách bookmark từ DB ──
            var bookmarks = AISearchService.GetBookmarks();
            if (bookmarks.Count == 0)
            {
                menu.Items.Add(new ToolStripMenuItem("(Chưa có bookmark nào)") { Enabled = false });
            }
            else
            {
                menu.Items.Add(new ToolStripMenuItem($"📌 {bookmarks.Count} câu hỏi đã lưu:") { Enabled = false });
                menu.Items.Add(new ToolStripSeparator());
                foreach (var (id, question) in bookmarks)
                {
                    int bmId = id;
                    string bm = question;
                    string lbl = bm.Length > 45 ? bm.Substring(0, 42) + "..." : bm;
                    var item = new ToolStripMenuItem($"  {lbl}");

                    // Click → điền vào input
                    item.Click += (s, e) =>
                    {
                        _txtInput.Text = bm;
                        _txtInput.Focus();
                        _txtInput.SelectAll();
                    };

                    // Sub-menu xóa từng bookmark
                    var del = new ToolStripMenuItem("🗑 Xóa bookmark này");
                    del.Click += (s, e) =>
                    {
                        string result = AISearchService.RemoveBookmark(bmId);
                        AppendMsg("⭐ Bookmark", result, Color.FromArgb(234, 179, 8));
                    };
                    item.DropDownItems.Add(del);
                    menu.Items.Add(item);
                }
                menu.Items.Add(new ToolStripSeparator());
                var clearAll = new ToolStripMenuItem("🗑 Xóa tất cả bookmark");
                clearAll.Click += (s, e) =>
                {
                    string result = AISearchService.ClearBookmarks();
                    AppendMsg("⭐ Bookmark", result, Color.FromArgb(234, 179, 8));
                };
                menu.Items.Add(clearAll);
            }

            menu.Show(this, new Point(PANEL_W - 120, 54));
        }

        // ── AI Skill ─────────────────────────────────────────────────────
        private void ShowSkillMenu()
        {
            var menu = new ContextMenuStrip();
            menu.BackColor = Color.FromArgb(35, 40, 58);
            menu.ForeColor = Color.FromArgb(226, 232, 240);
            menu.RenderMode = ToolStripRenderMode.System;

            var skills = AISearchService.GetSkills(forceRefresh: true);
            var typeLabels = new[] {
                ("quick",    "⚡ Câu hỏi nhanh"),
                ("slash",    "/ Lệnh gõ tắt"),
                ("workflow", "🔬 Workflow"),
            };

            bool hasAny = false;
            foreach (var (type, sectionLabel) in typeLabels)
            {
                var group = skills.FindAll(s => s.Type == type);
                if (group.Count == 0) continue;
                if (hasAny) menu.Items.Add(new ToolStripSeparator());
                menu.Items.Add(new ToolStripMenuItem(sectionLabel) { Enabled = false });
                foreach (var sk in group)
                {
                    int skId = sk.Id;
                    string tmpl = sk.Template;
                    string slashHint = !string.IsNullOrEmpty(sk.Slash) ? $"  [{sk.Slash}]" : "";
                    string lbl = (sk.Name.Length > 40 ? sk.Name.Substring(0, 37) + "..." : sk.Name) + slashHint;
                    var item = new ToolStripMenuItem($"  {lbl}");
                    if (!string.IsNullOrEmpty(sk.Description))
                        item.ToolTipText = sk.Description;
                    item.Click += (s, e) =>
                    {
                        _txtInput.Text = tmpl.Replace("{param}", "");
                        _txtInput.Focus();
                        _txtInput.SelectAll();
                    };
                    var del = new ToolStripMenuItem("🗑 Xóa skill này");
                    del.Click += (s2, e2) =>
                    {
                        string res = AISearchService.RemoveSkill(skId);
                        AppendMsg("⚡ Skill", res, Color.FromArgb(251, 191, 36));
                        RefreshSkillStrip();
                    };
                    item.DropDownItems.Add(del);
                    menu.Items.Add(item);
                }
                hasAny = true;
            }

            if (!hasAny)
                menu.Items.Add(new ToolStripMenuItem("(Chưa có skill nào)") { Enabled = false });

            menu.Items.Add(new ToolStripSeparator());
            var addNew = new ToolStripMenuItem("➕ Thêm skill mới...");
            addNew.Click += (s, e) =>
            {
                var dlg = new Forms.frmAddSkill();
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    AppendMsg("⚡ Skill", $"✅ Đã thêm skill: \"{dlg.SkillName}\"", Color.FromArgb(251, 191, 36));
                    RefreshSkillStrip();
                }
            };
            menu.Items.Add(addNew);
            menu.Show(this, new Point(PANEL_W - 152, 54));
        }

        private void RefreshSkillStrip()
        {
            if (_skillStrip == null) return;
            _skillStrip.Controls.Clear();
            var quickSkills = AISearchService.GetSkills("quick", forceRefresh: true);
            if (quickSkills.Count == 0) { _skillStrip.Visible = false; return; }

            int x = 6;
            foreach (var sk in quickSkills)
            {
                string tmpl = sk.Template;
                string label = sk.Icon + " " + (sk.Name.Length > 18 ? sk.Name.Substring(0, 15) + "…" : sk.Name);
                var chip = new Button
                {
                    Text = label,
                    Font = new Font("Segoe UI", 8f),
                    ForeColor = Color.FromArgb(226, 232, 240),
                    BackColor = Color.FromArgb(44, 50, 72),
                    FlatStyle = FlatStyle.Flat,
                    Location = new Point(x, 5),
                    Height = 26,
                    AutoSize = true,
                    Cursor = Cursors.Hand,
                    TabStop = false
                };
                chip.FlatAppearance.BorderColor = Color.FromArgb(80, 99, 102, 241);
                chip.FlatAppearance.BorderSize = 1;
                chip.FlatAppearance.MouseOverBackColor = Color.FromArgb(60, 65, 95);
                if (!string.IsNullOrEmpty(sk.Description))
                    new ToolTip().SetToolTip(chip, sk.Description);
                string chipName = sk.Icon + " " + sk.Name; // capture trước khi vào lambda
                chip.Click += (s, e) =>
                {
                    string resolved = tmpl;
                    string displayName = chipName;
                    if (resolved.Contains("{project_code}"))
                    {
                        string code = PromptProjectCode();
                        if (code == null) return; // user nhấn Hủy
                        resolved    = ApplyProjectCodeToTemplate(resolved, code);
                        displayName = chipName + (string.IsNullOrEmpty(code) ? "" : $" — {code}");
                    }
                    _pendingDisplayText = displayName;
                    _txtInput.Text      = resolved;
                    SendMessage();
                };
                _skillStrip.Controls.Add(chip);
                chip.PerformLayout();
                x += chip.Width + 6;
            }
            _skillStrip.Visible = true;
        }

        // ── Slash popup ───────────────────────────────────────────────────
        private record SlashEntry(string Display, string Template, string Type,
                                  string Slash = "", string ParentSlash = "", string Tag = "");

        /// <summary>
        /// Lệnh slash có Level 2 là danh sách dự án (thay vì sub-skills).
        /// Khi gõ "/po " hoặc "/mpr ", popup hiển thị project codes.
        /// </summary>
        private static readonly HashSet<string> PROJECT_SLASH_COMMANDS =
            new(StringComparer.OrdinalIgnoreCase) { "/po", "/mpr" };

        /// <summary>
        /// Thay thế {project_code} trong bất kỳ template nào.
        /// Xóa phần filter cũ mơ hồ ("Nếu X trống..."), thêm filter prefix BẮT BUỘC ở đầu prompt.
        /// Dùng cho quick chip, ApplySlashPopupSelection, và SendMessage fallback.
        /// </summary>
        private static string ApplyProjectCodeToTemplate(string template, string projectCode)
        {
            bool isAll = string.IsNullOrWhiteSpace(projectCode)
                      || projectCode.Equals("All", StringComparison.OrdinalIgnoreCase);

            // Xóa phần filter cũ mơ hồ trong template DB
            string cleaned = template
                .Replace(" Lọc dự án: {project_code}. Nếu {project_code} trống thì lấy tất cả dự án.", "")
                .Replace("{project_code}", isAll ? "" : projectCode);

            if (isAll) return cleaned;

            // Prefix BẮT BUỘC ở đầu — AI đọc trước câu hỏi, không bỏ qua được
            string prefix = $"⚠️ BẮT BUỘC: Chỉ lấy dữ liệu dự án '{projectCode}'. " +
                            $"SQL PHẢI có điều kiện: ProjectCode = '{projectCode}' (PO_head) " +
                            $"hoặc Project_Code = '{projectCode}' (MPR_Header). " +
                            $"KHÔNG được trả về dữ liệu dự án khác.\n";
            return prefix + cleaned;
        }

        /// <summary>
        /// Tạo prompt AI dành riêng cho truy vấn danh sách PO/MPR theo dự án.
        /// projectCode rỗng = tất cả dự án (All).
        /// Filter đặt ĐẦU TIÊN + ngôn ngữ bắt buộc để AI không bỏ qua.
        /// </summary>
        private static string BuildProjectTemplate(string slash, string projectCode)
        {
            bool isAll = string.IsNullOrWhiteSpace(projectCode)
                      || projectCode.Equals("All", StringComparison.OrdinalIgnoreCase);

            // Filter prefix: đặt trước câu hỏi chính, dùng ngôn ngữ BẮT BUỘC + tên cột chính xác
            string filterPrefix = isAll
                ? ""
                : $"⚠️ BẮT BUỘC: Chỉ lấy dữ liệu dự án '{projectCode}'. " +
                  $"SQL PHẢI có điều kiện: ph.ProjectCode = '{projectCode}' (bảng PO_head) " +
                  $"hoặc mh.Project_Code = '{projectCode}' (bảng MPR_Header). " +
                  $"KHÔNG được trả về dữ liệu của dự án khác.\n";

            string desc = isAll ? "tất cả dự án" : $"dự án {projectCode}";

            return slash.Equals("/po", StringComparison.OrdinalIgnoreCase)
                ? $"{filterPrefix}Danh sách tất cả PO của {desc}, kèm PONo, tên nhà cung cấp (Company_Name), " +
                  $"Total_Amount, Status, Expected_Delivery. Sắp xếp theo PO_Date giảm dần."
                : $"{filterPrefix}Danh sách tất cả MPR (Is_Latest=1) của {desc}, kèm MPR_No, Status, " +
                  $"Required_Date, số lượng item trong MPR_Details (Is_Deleted=0). " +
                  $"Sắp xếp theo Created_Date giảm dần.";
        }

        private void UpdateSlashPopup()
        {
            _slashLevel3Active = false; // user typed → Level 3 flow bị ngắt
            string text = _txtInput.Text;
            if (!text.StartsWith("/")) { HideSlashPopup(); return; }

            var skills = AISearchService.GetSkills();

            // ── Phát hiện "đã gõ đúng lệnh cha + space" → hiện Level 2 ──
            string[] parts = text.Split(' ');
            string firstWord = parts[0].ToLower();  // e.g. "/po"
            bool hasSpace = text.Contains(' ');

            if (hasSpace)
            {
                // Tìm skill cha khớp chính xác
                var parent = skills.Find(s =>
                    s.Type == "slash" && !string.IsNullOrEmpty(s.Slash) &&
                    s.Slash.Equals(firstWord, StringComparison.OrdinalIgnoreCase));

                if (parent != null)
                {
                    string filter = text.Substring(firstWord.Length).Trim().ToLower();

                    // ── /po và /mpr: Level 2 hiện danh sách mã dự án ────────
                    if (PROJECT_SLASH_COMMANDS.Contains(firstWord))
                    {
                        var allCodes = AISearchService.GetProjectCodes();
                        var entries  = new System.Collections.Generic.List<SlashEntry>();

                        // Mục "Tất cả dự án" luôn đứng đầu
                        bool showAll = string.IsNullOrEmpty(filter)
                            || "tất cả".Contains(filter) || "all".Contains(filter);
                        if (showAll)
                            entries.Add(new SlashEntry(
                                "📂  Tất cả dự án",
                                BuildProjectTemplate(firstWord, ""),
                                "slash", Tag: ""));

                        // Project codes khớp filter
                        foreach (var code in allCodes)
                            if (string.IsNullOrEmpty(filter) ||
                                code.ToLower().Contains(filter, StringComparison.OrdinalIgnoreCase))
                                entries.Add(new SlashEntry(
                                    $"📁  {code}",
                                    BuildProjectTemplate(firstWord, code),
                                    "slash", Tag: code));

                        if (entries.Count > 0)
                        {
                            string icon = parent.Icon;
                            string name = firstWord == "/po" ? "PO" : "MPR";
                            ShowPopup(entries, $"  ↳ {icon} Chọn dự án để xem danh sách {name}:");
                            _slashLevel2Parent = parent.Slash;
                            return;
                        }
                        HideSlashPopup();
                        return;
                    }

                    // ── Các lệnh khác: Level 2 hiện sub-skills như cũ ────────
                    var subs = AISearchService.GetSubSkills(parent.Slash);
                    var filtered = string.IsNullOrEmpty(filter)
                        ? subs
                        : subs.FindAll(s => s.Name.Contains(filter, StringComparison.OrdinalIgnoreCase)
                                         || s.Description.Contains(filter, StringComparison.OrdinalIgnoreCase));
                    if (filtered.Count > 0)
                    {
                        ShowPopup(filtered.ConvertAll(s =>
                            new SlashEntry($"{s.Icon}  {s.Name}", s.Template, s.Type, s.Slash, s.ParentSlash)),
                            $"  ↳ {parent.Icon} {parent.Name}  —  chọn tùy chọn:");
                        _slashLevel2Parent = parent.Slash;
                        return;
                    }
                }
            }

            // ── Level 1: lọc theo những gì đã gõ ──────────────────────────
            _slashLevel2Parent = null;
            string typed = firstWord;
            var matches = new System.Collections.Generic.List<SlashEntry>();

            foreach (var sk in skills)
            {
                // Chỉ hiện skill cha (không có Parent) ở level 1
                if (!string.IsNullOrEmpty(sk.ParentSlash)) continue;

                bool isSlash = sk.Type == "slash" && !string.IsNullOrEmpty(sk.Slash);
                bool isQuick = sk.Type == "quick";
                bool isWorkflow = sk.Type == "workflow";
                if (!isSlash && !isQuick && !isWorkflow) continue;

                string cmd = isSlash ? sk.Slash : "";
                string typeTag = isSlash ? "lệnh" : (isWorkflow ? "workflow" : "nhanh");
                string slashPart = isSlash ? $"{sk.Slash}  " : "";
                string display = $"{sk.Icon}  {slashPart}{sk.Name}   [{typeTag}]";

                bool show = typed == "/" ||
                            (!string.IsNullOrEmpty(cmd) && cmd.StartsWith(typed, StringComparison.OrdinalIgnoreCase)) ||
                            sk.Name.Contains(text.Substring(1), StringComparison.OrdinalIgnoreCase);
                if (show) matches.Add(new SlashEntry(display, sk.Template, sk.Type, sk.Slash, sk.ParentSlash));
            }

            if (matches.Count == 0) { HideSlashPopup(); return; }
            ShowPopup(matches, header: null);
        }

        private void ShowPopup(System.Collections.Generic.List<SlashEntry> entries, string header)
        {
            _slashPopup.Tag = entries;
            _slashPopup.Items.Clear();
            foreach (var m in entries) _slashPopup.Items.Add(m.Display);
            if (_slashPopup.Items.Count > 0) _slashPopup.SelectedIndex = 0;

            bool hasHeader = !string.IsNullOrEmpty(header);
            int headerH = hasHeader ? 22 : 0;
            int popupH = Math.Min(entries.Count, 7) * 26 + 2;
            int totalH = headerH + popupH;
            int baseY = _inputPanel.Top - totalH - 2;

            if (hasHeader)
            {
                _slashHeader.Text = header;
                _slashHeader.SetBounds(0, baseY, PANEL_W, headerH);
                _slashHeader.BringToFront();
                _slashHeader.Visible = true;
            }
            else
                _slashHeader.Visible = false;

            _slashPopup.SetBounds(0, baseY + headerH, PANEL_W, popupH);
            _slashPopup.BringToFront();
            _slashPopup.Visible = true;
        }

        private void HideSlashPopup()
        {
            _slashPopup?.Hide();
            _slashHeader?.Hide();
            _slashLevel2Parent    = null;
            _slashLevel3Active    = false;
            _slashSelectedProject = null;
            // _pendingDisplayText KHÔNG clear ở đây — cần sống đến khi SendMessage() đọc
        }

        private void ApplySlashPopupSelection()
        {
            if (_slashPopup.SelectedIndex < 0) return;
            if (_slashPopup.Tag is not System.Collections.Generic.List<SlashEntry> entries) return;
            var entry = entries[_slashPopup.SelectedIndex];

            // ── Level 3 đang hiển thị: template đã embedded project code → gửi ngay ──
            if (_slashLevel3Active)
            {
                // Hiện tên ngắn gọn trong chat thay vì full template
                string proj = string.IsNullOrEmpty(_slashSelectedProject) ? "Tất cả dự án"
                                                                           : _slashSelectedProject;
                _pendingDisplayText = $"{entry.Display.Trim()} — {proj}";

                HideSlashPopup();
                _txtInput.Text = entry.Template;
                _txtInput.SelectionStart = _txtInput.Text.Length;
                _txtInput.Focus();
                SendMessage();
                return;
            }

            // ── Level 1: slash có sub-skill → gõ "/lệnh " để hiện Level 2 ──
            if (string.IsNullOrEmpty(entry.ParentSlash) && entry.Type == "slash"
                && !string.IsNullOrEmpty(entry.Slash)
                && AISearchService.GetSubSkills(entry.Slash).Count > 0)
            {
                _txtInput.Text = entry.Slash + " ";
                _txtInput.SelectionStart = _txtInput.Text.Length;
                _txtInput.Focus();
                // TextChanged tự gọi UpdateSlashPopup() → hiện Level 2 (project codes hoặc sub-skills)
                return;
            }

            // ── Level 2 /po hoặc /mpr: dự án được chọn → hiện Level 3 sub-skills ──
            if (PROJECT_SLASH_COMMANDS.Contains(_slashLevel2Parent ?? ""))
            {
                ShowLevel3ProjectSubSkills(_slashLevel2Parent, entry.Tag);
                return;
            }

            // ── Các trường hợp còn lại (Level 2 sub-skill thường, hoặc Level 1 quick) ──
            HideSlashPopup();
            string resolvedEntry = entry.Template.Replace("{param}", "");

            // Skill có {project_code}: hiện dialog chọn dự án
            if (resolvedEntry.Contains("{project_code}"))
            {
                string code = PromptProjectCode();
                if (code == null) { _txtInput.Focus(); return; }
                resolvedEntry = ApplyProjectCodeToTemplate(resolvedEntry, code);
                _pendingDisplayText = entry.Display.Trim() +
                    (string.IsNullOrEmpty(code) ? "" : $" — {code}");
                _txtInput.Text = resolvedEntry;
                _txtInput.SelectionStart = _txtInput.Text.Length;
                _txtInput.Focus();
                SendMessage();
                return;
            }

            // Skill không có project code — hiện tên skill ngắn gọn
            _pendingDisplayText = entry.Display.Trim();
            _txtInput.Text = resolvedEntry;
            _txtInput.SelectionStart = _txtInput.Text.Length;
            _txtInput.Focus();
        }

        /// <summary>
        /// Hiện Level 3 popup: danh sách các loại kiểm tra cho /po hoặc /mpr
        /// sau khi đã chọn mã dự án ở Level 2.
        /// </summary>
        private void ShowLevel3ProjectSubSkills(string slash, string projectCode)
        {
            _slashLevel3Active    = true;
            _slashSelectedProject = projectCode; // lưu để dùng trong display text

            bool isAll = string.IsNullOrEmpty(projectCode)
                      || projectCode.Equals("All", StringComparison.OrdinalIgnoreCase);
            string listName    = slash.Equals("/po", StringComparison.OrdinalIgnoreCase) ? "PO" : "MPR";
            string projectDesc = isAll ? "Tất cả dự án" : $"Dự án {projectCode}";

            var entries = new System.Collections.Generic.List<SlashEntry>();

            // ── Mục đầu: "Tất cả [PO/MPR]" — lấy danh sách đầy đủ của dự án ──
            entries.Add(new SlashEntry(
                $"📋  Tất cả {listName}",
                BuildProjectTemplate(slash, projectCode),
                "slash"));

            // Filter prefix mạnh cho sub-skills — đặt ĐẦU prompt, tên cột chính xác
            // isAll = rỗng hoặc "All" → không lọc; ngược lại → bắt buộc lọc
            string subFilterPrefix = isAll
                ? ""
                : $"⚠️ BẮT BUỘC: Chỉ lấy dữ liệu dự án '{projectCode}'. " +
                  $"SQL PHẢI có điều kiện: ph.ProjectCode = '{projectCode}' (PO_head) " +
                  $"hoặc mh.Project_Code = '{projectCode}' (MPR_Header). " +
                  $"KHÔNG được trả về dữ liệu dự án khác.\n";

            // ── Sub-skills từ DB (Parent_Slash = slash) ──────────────────────────
            var subs = AISearchService.GetSubSkills(slash);
            foreach (var sub in subs)
            {
                // Xóa phần filter cũ mơ hồ trong template DB (tránh "Nếu PROJ-001 trống...")
                string tmpl = sub.Template
                    .Replace(" Lọc dự án: {project_code}. Nếu {project_code} trống thì lấy tất cả dự án.", "")
                    .Replace("{project_code}", projectCode); // fallback cho template chưa chuẩn hóa

                // Ghép filter prefix vào ĐẦU prompt
                entries.Add(new SlashEntry($"{sub.Icon}  {sub.Name}", subFilterPrefix + tmpl, sub.Type));
            }

            ShowPopup(entries, $"  ↳ 📁 {projectDesc}  —  chọn loại kiểm tra:");
        }

        // ── Chọn mã dự án — dialog tối màu, dropdown từ DB ───────────────
        /// <summary>
        /// Hiển thị dialog chọn mã dự án. Trả về mã đã chọn (chuỗi rỗng = tất cả dự án),
        /// hoặc null nếu người dùng nhấn Hủy.
        /// </summary>
        private string PromptProjectCode()
        {
            var codes = AISearchService.GetProjectCodes();

            using var dlg = new Form
            {
                Text            = "🔍 Chọn dự án",
                Size            = new Size(340, 150),
                MinimumSize     = new Size(340, 150),
                MaximumSize     = new Size(600, 150),
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition   = FormStartPosition.CenterScreen,
                MaximizeBox     = false,
                MinimizeBox     = false,
                TopMost         = true,
                BackColor       = Color.FromArgb(22, 27, 42)
            };

            var lbl = new Label
            {
                Text      = "Mã dự án (để trống = tất cả dự án):",
                ForeColor = Color.FromArgb(148, 163, 184),
                AutoSize  = true,
                Font      = new Font("Segoe UI", 9f),
                Location  = new Point(14, 16)
            };

            var cb = new ComboBox
            {
                Location           = new Point(14, 40),
                Width              = 298,
                BackColor          = Color.FromArgb(35, 40, 58),
                ForeColor          = Color.FromArgb(226, 232, 240),
                FlatStyle          = FlatStyle.Flat,
                Font               = new Font("Segoe UI", 10f),
                AutoCompleteMode   = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.ListItems
            };
            cb.Items.Add(""); // để trống = tất cả
            foreach (var c in codes) cb.Items.Add(c);
            cb.SelectedIndex = 0;

            var btnOk = new Button
            {
                Text         = "✔ Xác nhận",
                DialogResult = DialogResult.OK,
                Location     = new Point(140, 82),
                Width        = 90,
                BackColor    = Color.FromArgb(79, 70, 229),
                ForeColor    = Color.White,
                FlatStyle    = FlatStyle.Flat,
                Cursor       = Cursors.Hand,
                Font         = new Font("Segoe UI", 9f)
            };
            btnOk.FlatAppearance.BorderSize = 0;

            var btnCancel = new Button
            {
                Text         = "Hủy",
                DialogResult = DialogResult.Cancel,
                Location     = new Point(238, 82),
                Width        = 74,
                BackColor    = Color.FromArgb(50, 55, 80),
                ForeColor    = Color.FromArgb(148, 163, 184),
                FlatStyle    = FlatStyle.Flat,
                Cursor       = Cursors.Hand,
                Font         = new Font("Segoe UI", 9f)
            };
            btnCancel.FlatAppearance.BorderSize = 0;

            dlg.Controls.AddRange(new Control[] { lbl, cb, btnOk, btnCancel });
            dlg.AcceptButton = btnOk;
            dlg.CancelButton = btnCancel;

            return dlg.ShowDialog(this) == DialogResult.OK
                ? (cb.Text?.Trim() ?? "")
                : null;
        }

        // ── Tính năng 3: Xuất Excel ──────────────────────────────────────
        private void ExportToExcel()
        {
            string path = _ai.ExportLastResultToExcel(_lastQuestion);
            if (path != null)
            {
                AppendMsg("AI",
                    $"✅ Đã xuất Excel thành công!\nFile: {System.IO.Path.GetFileName(path)}",
                    Color.FromArgb(52, 211, 153));
                System.Diagnostics.Process.Start(
                    new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
                _btnExport.Visible = false;
            }
            else
            {
                AppendMsg("AI", "⚠️ Không có dữ liệu để xuất. Hãy hỏi câu hỏi trả về bảng dữ liệu trước.", C_SUBTEXT);
            }
        }

        private void SetThinking(bool thinking)
        {
            _isThinking = thinking;
            _btnSend.Enabled = !thinking;
            _txtInput.Enabled = !thinking;
            _lblStatus.Text = thinking
                ? "⏳ Đang phân tích và truy vấn dữ liệu..."
                : "Sẵn sàng  •  Hỏi về MPR, PO, NCC, thanh toán...";
            if (!thinking) _txtInput.Focus();
        }

        // ════════════════════════════════════════════════════════════════════
        // POSITIONING & HOTKEY
        // ════════════════════════════════════════════════════════════════════

        public void PositionNearButton()
        {
            var scr = Screen.PrimaryScreen.WorkingArea;
            int x, y;

            if (_btnInst != null && !_btnInst.IsDisposed)
            {
                x = Math.Max(scr.Left + 8,
                    Math.Min(_btnInst.Left + BTN_SIZE - PANEL_W,
                             scr.Right - PANEL_W - 8));
                y = Math.Max(scr.Top + 8, _btnInst.Top - PANEL_H - 8);
            }
            else
            {
                x = scr.Right - PANEL_W - MARGIN;
                y = scr.Bottom - PANEL_H - BTN_SIZE - MARGIN - 8;
            }
            Location = new Point(x, y);
        }

        private void RegisterHotkey()
        {
            Application.AddMessageFilter(new HotkeyFilter(Keys.Space, () =>
            {
                if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
                    ToggleChat();
            }));
        }

        private GraphicsPath RoundedPath(Rectangle r, int radius)
        {
            int d = radius * 2;
            var p = new GraphicsPath();
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        private class HotkeyFilter : IMessageFilter
        {
            private readonly Keys _key;
            private readonly Action _cb;
            private const int WM_KEYDOWN = 0x0100;
            public HotkeyFilter(Keys key, Action cb) { _key = key; _cb = cb; }
            public bool PreFilterMessage(ref Message m)
            {
                if (m.Msg == WM_KEYDOWN && (Keys)m.WParam == _key)
                    _cb?.Invoke();
                return false;
            }
        }
    }
}