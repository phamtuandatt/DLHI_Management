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



            // Nút bộ nhớ AI
            var btnMemory = HeaderBtn("🧠", PANEL_W - 152);
            new ToolTip().SetToolTip(btnMemory, "Xem / quản lý bộ nhớ AI");
            btnMemory.Click += (s, e) => ShowMemoryMenu();

            // Nút đóng
            var btnClose = HeaderBtn("×", PANEL_W - 44);
            btnClose.Font = new Font("Segoe UI", 14);
            new ToolTip().SetToolTip(btnClose, "Đóng  (Ctrl+Space)");
            btnClose.Click += (s, e) => ToggleChat();

            header.Controls.AddRange(new Control[] { btnClear, btnMemory, btnBookmark, btnClose });

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
                if (e.KeyCode == Keys.Enter && !e.Shift)
                {
                    e.SuppressKeyPress = true;
                    SendMessage();
                }
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

            // ── Ghép lại ────────────────────────────────────────────────
            Controls.Add(chatWrap);
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
            AppendMsg("Bạn", q, C_USER_MSG);

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