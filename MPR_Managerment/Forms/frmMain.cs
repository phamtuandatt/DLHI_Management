using Microsoft.Data.SqlClient;
using MPR_Managerment.Forms.RIRGUI;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;
using MPR_Managerment.Services;
using System;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MPR_Managerment.Forms
{
    // Interface để các form con implement refresh thủ công
    public interface IRefreshable
    {
        void RefreshData();
    }

    public partial class frmMain : Form
    {
        // ── Sự kiện tĩnh cho phép form con báo hiệu có dữ liệu mới ──────────
        public static event EventHandler NewDataCreated;
        public static void RaiseNewDataCreated()
        {
            NewDataCreated?.Invoke(null, EventArgs.Empty);
        }

        private Panel panelMenu;
        private Panel panelContent;
        private Panel panelHeader;
        private Label lblUser;
        private Form _activeForm = null;
        // Cache form — Hide thay vì Close khi chuyển, tránh reload data
        private readonly Dictionary<string, Form> _formCache = new Dictionary<string, Form>();
        private string _activeFormKey = "";
        private Button _btnRefresh;
        // Notification
        private Panel _panelNotify;
        private ListBox _lstNotify;
        private Label _lblNotifyCount;
        private System.Windows.Forms.Timer _notifyTimer;
        private DateTime _lastCheckTime = DateTime.MinValue;
        private Button _btnNotify;
        private int _notifyBadge = 0;          // so thong bao chua doc
        private int _unreadCount = 0;           // dem tu lan mo truoc
        private System.Collections.Generic.HashSet<string> _seenMsgs
            = new System.Collections.Generic.HashSet<string>(); // da hien
        private System.Collections.Generic.HashSet<string> _readMsgs
            = new System.Collections.Generic.HashSet<string>(); // da doc
        private System.Collections.Generic.List<string> _starredMsgs
            = new System.Collections.Generic.List<string>(); // quan trong
        private Button _btnTabStar;  // de cap nhat so luong Q Trong
        private Button _btnZalo;
        // Thông báo nội bộ (Admin → User)
        private readonly InternalNotificationService _internalNotifSvc = new();
        private readonly System.Collections.Generic.HashSet<int> _shownNotifIds = new();
        private System.Windows.Forms.Timer _internalNotifTimer;
        private ToolTip _zaloTip = new ToolTip();
        // ── Menu cố định 190px ──
        private const int MENU_FULL = 190;
        // ── File lưu thời gian check cuối (persist giữa các lần mở app) ──
        private static readonly string _lastCheckFile =
            Path.Combine(Application.StartupPath, "last_notify_check.txt");

        public frmMain()
        {
            InitializeComponent();
            // Gắn AI Trợ lý vào form chính
            frmAIChat.Attach(this);
            
            this.Load += FrmMain_Load;
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {
            this.Text = "MPR Management System";
            this.Size = new Size(1400, 820);
            this.MinimumSize = new Size(1100, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.WindowState = FormWindowState.Maximized;
            this.BackColor = Color.FromArgb(245, 245, 245);

            // Khởi tạo ZaloNotificationService trên UI thread
            // để đảm bảo Timer và WebView2 chạy đúng trên UI thread
            _ = ZaloNotificationService.Instance;

            BuildHeader();
            BuildMenu();
            BuildContent();
            this.Resize += (s, ev) => UpdateContentBounds();
            UpdateContentBounds();
            BuildNotifyPanel();
            StartNotifyTimer();
            ShowDashboard();

            // Đăng ký lắng nghe trạng thái Zalo một lần duy nhất
            ZaloSession.StatusChanged += status =>
            {
                if (_btnZalo == null || _btnZalo.IsDisposed) return;
                if (_btnZalo.InvokeRequired)
                    _btnZalo.Invoke(() => _UpdateZaloButton(status));
                else
                    _UpdateZaloButton(status);
            };

            // Khởi động session Zalo nền ngay khi mở app (nếu đã bật)
            _ = _WarmUpZaloSessionAsync();
        }

        private async Task _WarmUpZaloSessionAsync()
        {
            var cfg = ZaloHelper.LoadSettings();
            if (!cfg.Enabled)
            {
                _UpdateZaloButton(ZaloStatus.Disabled);
                return;
            }
            try
            {
                await ZaloSession.EnsureAsync(cfg);
            }
            catch
            {
                _UpdateZaloButton(ZaloStatus.Error);
            }
        }

        private void _UpdateZaloButton(ZaloStatus status)
        {
            switch (status)
            {
                case ZaloStatus.Ready:
                    _btnZalo.BackColor = Color.FromArgb(0, 140, 70);
                    _btnZalo.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 100, 50);
                    _zaloTip.SetToolTip(_btnZalo, "✅ Zalo đang kết nối — nhấn để cài đặt");
                    break;
                case ZaloStatus.Connecting:
                    _btnZalo.BackColor = Color.FromArgb(200, 130, 0);
                    _btnZalo.FlatAppearance.MouseOverBackColor = Color.FromArgb(160, 100, 0);
                    _zaloTip.SetToolTip(_btnZalo, "⏳ Đang kết nối Zalo...");
                    break;
                case ZaloStatus.Error:
                    _btnZalo.BackColor = Color.FromArgb(180, 30, 30);
                    _btnZalo.FlatAppearance.MouseOverBackColor = Color.FromArgb(140, 20, 20);
                    _zaloTip.SetToolTip(_btnZalo, "❌ Zalo mất kết nối hoặc chưa đăng nhập — nhấn để kiểm tra");
                    break;
                default: // Disabled
                    _btnZalo.BackColor = Color.FromArgb(120, 120, 120);
                    _btnZalo.FlatAppearance.MouseOverBackColor = Color.FromArgb(90, 90, 90);
                    _zaloTip.SetToolTip(_btnZalo, "⚠ Zalo chưa bật — nhấn để bật");
                    break;
            }
        }

        // =====================================================================
        //  HEADER
        // =====================================================================
        private void BuildHeader()
        {
            panelHeader = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(this.Width, 55),
                BackColor = Color.FromArgb(0, 120, 212),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            panelHeader.Controls.Add(new Label
            {
                Text = "⚙ DLHI ERP",
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(15, 12),
                Size = new Size(500, 32)
            });

            // Hiển thị user đang đăng nhập + Role
            string userText = AppSession.CurrentUser != null
                ? $"👤 {AppSession.CurrentUser.Full_Name}  [{AppSession.CurrentUser.Role_Name}]"
                : "👤 Admin";

            lblUser = new Label
            {
                Text = userText,
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.White,
                Size = new Size(280, 25),
                TextAlign = ContentAlignment.MiddleRight,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            lblUser.Location = new Point(panelHeader.Width - 300, 15);
            panelHeader.Controls.Add(lblUser);

            // Nút Làm mới — nằm bên trái lblUser
            _btnRefresh = new Button
            {
                Text = "🔄 Làm mới",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(0, 90, 180),
                FlatStyle = FlatStyle.Flat,
                Size = new Size(100, 30),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            _btnRefresh.FlatAppearance.BorderSize = 0;
            _btnRefresh.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 60, 140);
            new ToolTip().SetToolTip(_btnRefresh, "Làm mới dữ liệu form hiện tại");
            _btnRefresh.Location = new Point(panelHeader.Width - 415, 12);
            _btnRefresh.Click += (s, e) => RefreshActiveForm();
            panelHeader.Controls.Add(_btnRefresh);

            // Nút cài đặt Zalo
            _btnZalo = new Button
            {
                Text = "🔔 Zalo",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(120, 120, 120), // xám mặc định cho đến khi biết trạng thái
                FlatStyle = FlatStyle.Flat,
                Size = new Size(90, 30),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            _btnZalo.FlatAppearance.BorderSize = 0;
            _btnZalo.FlatAppearance.MouseOverBackColor = Color.FromArgb(90, 90, 90);
            _zaloTip.SetToolTip(_btnZalo, "Cài đặt thông báo Zalo");
            _btnZalo.Location = new Point(panelHeader.Width - 515, 12);
            _btnZalo.Click += (s, e) => ShowZaloSettingsDialog();
            panelHeader.Controls.Add(_btnZalo);

            panelHeader.Resize += (s, ev) =>
            {
                panelHeader.Width = this.ClientSize.Width;
                lblUser.Left = panelHeader.Width - 300;
                _btnRefresh.Left = panelHeader.Width - 415;
                _btnZalo.Left = panelHeader.Width - 515;
            };

            this.Controls.Add(panelHeader);
        }

        // =====================================================================
        //  MENU — ẩn/hiện theo quyền AppSession
        // =====================================================================
        private void BuildMenu()
        {
            panelMenu = new Panel
            {
                Location = new Point(0, 55),
                Size = new Size(190, this.Height - 55),
                BackColor = Color.FromArgb(30, 30, 45),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom
            };

            panelMenu.Controls.Add(new Label
            {
                Text = "MENU CHÍNH",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(150, 150, 180),
                Location = new Point(0, 10),
                Size = new Size(190, 22),
                TextAlign = ContentAlignment.MiddleCenter
            });

            int y = 55;

            // Tổng quan — luôn hiện
            AddMenuBtn("🏠  Tổng quan", Color.FromArgb(0, 120, 212), y); y += 42;
            AddMenuBtn("📊  Dashboard", Color.FromArgb(30, 30, 45), y); y += 42;

            // Các module — kiểm tra quyền
            if (AppSession.CanView("PROJECT"))
            {
                AddMenuBtn("🗂  Dự án", Color.FromArgb(30, 30, 45), y); y += 42;
            }
            if (AppSession.CanView("SUPPLIER"))
            {
                AddMenuBtn("🏢  Nhà cung cấp", Color.FromArgb(30, 30, 45), y); y += 42;
            }
            if (AppSession.CanView("MPR"))
            {
                AddMenuBtn("📋  MPR", Color.FromArgb(30, 30, 45), y); y += 42;
            }
            if (AppSession.CanView("PO"))
            {
                AddMenuBtn("🛒  Đơn đặt hàng (PO)", Color.FromArgb(30, 30, 45), y); y += 42;
            }
            if (AppSession.CanView("RIR"))
            {
                AddMenuBtn("📦  Kiểm tra (RIR)", Color.FromArgb(30, 30, 45), y); y += 42;
            }
            if (AppSession.CanView("Material Inspector Request"))
            {
                AddMenuBtn("📦  Material Inspector Request", Color.FromArgb(30, 30, 45), y); y += 42;
            }
            if (AppSession.CanView("WAREHOUSE"))
            {
                AddMenuBtn("🏭  Kho vật tư", Color.FromArgb(30, 30, 45), y); y += 42;
                AddMenuBtn("📂  Phân loại file Scan", Color.FromArgb(30, 30, 45), y); y += 42;
                AddMenuBtn("📧  Hóa đơn Outlook", Color.FromArgb(30, 30, 45), y); y += 42;

                AddMenuBtn("💳  Thanh toán Debit", Color.FromArgb(30, 30, 45), y); y += 42;
                AddMenuBtn("📥  Zalo Import", Color.FromArgb(30, 30, 45), y); y += 42;


            }

            // Divider
            panelMenu.Controls.Add(new Label
            {
                Location = new Point(10, y + 5),
                Size = new Size(200, 1),
                BackColor = Color.FromArgb(60, 60, 80)
            });
            y += 18;

            // Quản lý User — chỉ Admin
            if (AppSession.CanView("USER_MGT"))
            {
                AddMenuBtn("👤  Quản lý User", Color.FromArgb(63, 81, 181), y); y += 42;
                AddMenuBtn("📜  Lịch sử thông báo", Color.FromArgb(30, 30, 45), y); y += 42;
            }

            // Gửi thông báo nội bộ — chỉ Admin
            if (AppSession.IsAdmin)
            {
                AddMenuBtn("📢  Gửi thông báo", Color.FromArgb(198, 40, 40), y); y += 42;
            }

            // Đổi mật khẩu — luôn hiện
            AddMenuBtn("🔑  Đổi mật khẩu", Color.FromArgb(30, 30, 45), y); y += 42;
            AddMenuBtn("❌  Thoát", Color.FromArgb(30, 30, 45), y);

            this.Controls.Add(panelMenu);
        }

        private void AddMenuBtn(string text, Color backColor, int y)
        {
            var btn = new Button
            {
                Text = text,
                Location = new Point(0, y),
                Size = new Size(190, 38),
                BackColor = backColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(12, 0, 0, 0),
                Cursor = Cursors.Hand,
                Tag = text
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 100, 180);
            btn.Click += MenuBtn_Click;
            // Lưu text đầy đủ để dùng khi mở rộng, Tag dùng để nhận diện
            btn.Tag = text;
            panelMenu.Controls.Add(btn);
        }

        // =====================================================================
        //  MENU CLICK
        // =====================================================================
        private void MenuBtn_Click(object sender, EventArgs e)
        {
            var btn = sender as Button;
            if (btn == null) return;

            foreach (Control c in panelMenu.Controls)
                if (c is Button b) b.BackColor = Color.FromArgb(30, 30, 45);

            btn.BackColor = Color.FromArgb(0, 120, 212);
            string tag = btn.Tag?.ToString() ?? "";

            // Tag = text button (có emoji + spaces) → dùng Contains với keyword đủ unique
            if (tag.Contains("Tổng quan")) ShowDashboard();
            else if (tag.Contains("Dashboard")) OpenForm(new frmDashboard());
            else if (tag.Contains("Dự án")) OpenForm(new frmProject());
            else if (tag.Contains("Nhà cung cấp")) OpenForm(new frmSupplier());
            else if (tag.Contains("Material Inspector")) OpenForm(new frmRIRForQC()); // trước RIR
            else if (tag.Contains("Kiểm tra (RIR)")) OpenForm(new frmRIR());
            else if (tag.Contains("MPR")) OpenForm(new frmMPR());
            else if (tag.Contains("Đơn đặt hàng")) OpenForm(new frmPO());
            else if (tag.Contains("Thanh toán Debit")) OpenForm(new frmPayment());
            else if (tag.Contains("Zalo Import")) OpenForm(new frmZaloImport());
            else if (tag.Contains("Kho vật tư")) OpenForm(new frmWarehouses_v2());
            else if (tag.Contains("Phân loại file")) OpenForm(new frmFileSorter());
            else if (tag.Contains("Hóa đơn Outlook")) OpenForm(new frmOutlookInvoice());
            else if (tag.Contains("Quản lý User")) OpenForm(new frmUserManagement());
            else if (tag.Contains("Lịch sử thông báo")) OpenForm(new frmNotificationLog());
            else if (tag.Contains("Gửi thông báo"))
            {
                var frm = new frmSendNotification();
                frm.ShowDialog(this);
            }
            else if (tag.Contains("Đổi mật khẩu"))
            {
                if (AppSession.CurrentUser != null)
                {
                    var f = new frmChangePassword(AppSession.CurrentUser.User_ID);
                    f.ShowDialog(this);
                }
            }
            else if (tag.Contains("Thoát"))
            {
                this.BringToFront(); this.Activate();
                if (MessageBox.Show(this, "Bạn có chắc muốn thoát?", "Xác nhận",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    ZaloNotificationService.Instance.Dispose();
                    _notifyTimer?.Stop(); _notifyTimer?.Dispose();
                    _internalNotifTimer?.Stop(); _internalNotifTimer?.Dispose();
                    AppSession.Clear();
                    Application.Exit();
                }
            }
        }

        // =====================================================================
        //  CONTENT
        // =====================================================================
        private void BuildContent()
        {
            panelContent = new Panel
            {
                BackColor = Color.FromArgb(245, 245, 245),
                AutoScroll = true
            };
            this.Controls.Add(panelContent);
        }

        private void UpdateContentBounds()
        {
            int headerH = 55;
            panelHeader.Width = this.ClientSize.Width;
            panelMenu.Width = MENU_FULL;
            panelMenu.Height = this.ClientSize.Height - headerH;
            panelContent.Location = new Point(MENU_FULL, headerH);
            panelContent.Size = new Size(this.ClientSize.Width - MENU_FULL,
                                             this.ClientSize.Height - headerH);
            panelMenu.BringToFront();
            panelHeader.BringToFront();
        }


        private void OpenForm(Form form)
        {
            // Lấy key từ tên type của form
            string key = form.GetType().Name;

            // Nếu đang hiển thị đúng form này rồi → không làm gì
            if (_activeFormKey == key && _activeForm != null && !_activeForm.IsDisposed)
            {
                form.Dispose(); // dispose instance mới vì đã có cache
                return;
            }

            // Ẩn form hiện tại (không Close → giữ nguyên data)
            if (_activeForm != null && !_activeForm.IsDisposed)
                _activeForm.Hide();

            // Dùng cache nếu đã có, không tạo mới
            if (_formCache.TryGetValue(key, out Form cached) && !cached.IsDisposed)
            {
                form.Dispose(); // dispose instance mới, dùng cached
                _activeForm = cached;
            }
            else
            {
                _formCache[key] = form;
                _activeForm = form;
                _activeForm.TopLevel = false;
                _activeForm.FormBorderStyle = FormBorderStyle.None;
                _activeForm.Dock = DockStyle.Fill;
                panelContent.Controls.Add(_activeForm);
            }

            _activeFormKey = key;
            // Đảm bảo form vẫn là child của panelContent (có thể bị xóa khi ShowDashboard)
            if (_activeForm.Parent != panelContent)
            {
                panelContent.Controls.Clear(); // xóa dashboard controls
                _activeForm.TopLevel = false;
                _activeForm.FormBorderStyle = FormBorderStyle.None;
                _activeForm.Dock = DockStyle.Fill;
                panelContent.Controls.Add(_activeForm);
            }
            _activeForm.Show();
            _activeForm.BringToFront();
        }

        // Làm mới dữ liệu form đang hiển thị
        private void RefreshActiveForm()
        {
            if (_activeForm == null || _activeForm.IsDisposed) return;
            // Gọi interface IRefreshable nếu form hỗ trợ
            if (_activeForm is IRefreshable r)
            {
                r.RefreshData();
                return;
            }
            // Fallback: tạo lại form (xóa khỏi cache)
            string key = _activeFormKey;
            _activeForm.Close();
            _activeForm = null;
            if (_formCache.ContainsKey(key))
                _formCache.Remove(key);
            // Mở lại form mới qua menu handler
            foreach (Control c in panelMenu.Controls)
                if (c is Button btn && btn.Tag is string tag && tag.Contains(key.Replace("frm", "")))
                { btn.PerformClick(); return; }
        }

        // =====================================================================
        //  DASHBOARD
        // =====================================================================
        private void ShowDashboard()
        {
            // Ẩn và dispose tất cả cached forms trước khi clear panelContent
            foreach (var kv in _formCache)
            {
                try { kv.Value.Hide(); } catch { }
            }
            _formCache.Clear();
            if (_activeForm != null && !_activeForm.IsDisposed)
                _activeForm.Hide();
            _activeForm = null;
            _activeFormKey = "";

            panelContent.Controls.Clear();

            panelContent.Controls.Add(new Label
            {
                Text = "TỔNG QUAN HỆ THỐNG",
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(30, 20),
                Size = new Size(600, 40)
            });

            panelContent.Controls.Add(new Label
            {
                Text = $"📅 {DateTime.Now:dddd, dd/MM/yyyy  HH:mm}",
                Font = new Font("Segoe UI", 10),
                ForeColor = Color.Gray,
                Location = new Point(30, 65),
                Size = new Size(600, 25)
            });

            try
            {
                int supplierCount = new SupplierService().GetAll().Count;
                int mprCount = new MPRService().GetAll().Count;
                int poCount = new POService().GetAll().Count;
                int rirCount = new RIRService().GetAll().Count;
                int projectCount = new ProjectService().GetAll().Count;

                // Hàng 1
                AddCard("🏢 Nhà cung cấp", supplierCount.ToString(), Color.FromArgb(0, 120, 212), 30, 110);
                AddCard("📋 Phiếu MPR", mprCount.ToString(), Color.FromArgb(40, 167, 69), 260, 110);
                AddCard("🛒 Đơn PO", poCount.ToString(), Color.FromArgb(255, 140, 0), 490, 110);
                AddCard("📦 Phiếu RIR", rirCount.ToString(), Color.FromArgb(102, 51, 153), 720, 110);

                // Hàng 2
                AddCard("🗂 Dự án", projectCount.ToString(), Color.FromArgb(0, 150, 136), 30, 310);
                AddCard("🏭 Kho vật tư", "📦 Xem", Color.FromArgb(63, 81, 181), 260, 310);
                AddCard("📊 Dashboard", "📈 Xem", Color.FromArgb(233, 30, 99), 490, 310);


            }
            catch (Exception ex)
            {
                panelContent.Controls.Add(new Label
                {
                    Text = "⚠ Không thể tải thống kê: " + ex.Message,
                    Font = new Font("Segoe UI", 9),
                    ForeColor = Color.Red,
                    Location = new Point(30, 110),
                    Size = new Size(800, 25)
                });
            }

            // ── Panel tiến độ dự án (phía phải) ──────────────────────────
            BuildProjectProgressPanel();

            // Panel hướng dẫn
            var panelGuide = new Panel
            {
                Location = new Point(30, 510),
                Size = new Size(950, 200),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
            panelContent.Controls.Add(panelGuide);

            panelGuide.Controls.Add(new Label
            {
                Text = "📖  HƯỚNG DẪN SỬ DỤNG",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(15, 12),
                Size = new Size(500, 28)
            });

            string[] guides = {
                "1️⃣   Nhà cung cấp  →  Quản lý danh sách NCC, thông tin liên hệ, tài khoản ngân hàng",
                "2️⃣   MPR            →  Tạo phiếu yêu cầu mua vật tư, quản lý chi tiết vật tư từng dự án",
                "3️⃣   Đơn PO         →  Tạo đơn đặt hàng, import từ MPR, xuất Excel theo template",
                "4️⃣   Kiểm tra RIR   →  Tạo phiếu kiểm tra hàng nhập, ghi nhận MTR No / Heat No / ID Code",
                "5️⃣   Kho vật tư     →  Nhập kho từ PO, xuất kho cho dự án, theo dõi tồn kho"
            };

            int gy = 50;
            foreach (var g in guides)
            {
                panelGuide.Controls.Add(new Label
                {
                    Text = g,
                    Font = new Font("Segoe UI", 9),
                    ForeColor = Color.FromArgb(50, 50, 50),
                    Location = new Point(15, gy),
                    Size = new Size(920, 25)
                });
                gy += 30;
            }

            // Version bar
            var panelVersion = new Panel
            {
                Location = new Point(30, 730),
                Size = new Size(950, 50),
                BackColor = Color.FromArgb(30, 30, 45)
            };
            panelContent.Controls.Add(panelVersion);
            panelVersion.Controls.Add(new Label
            {
                Text = "MPR Management System  v2.0  —  C# Windows Forms + SQL Server Azure",
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.FromArgb(180, 180, 200),
                Location = new Point(15, 15),
                Size = new Size(920, 22)
            });
        }

        // ── Tiến độ mua hàng từng dự án active ─────────────────────────────
        private void BuildProjectProgressPanel()
        {
            // SQL dùng subquery độc lập — tránh double-count khi JOIN nhiều bảng
            const string SQL = @"
                SELECT
                    p.ProjectCode,
                    p.ProjectName,
                    ISNULL(p.PJBudget, 0)  AS Budget,
                    ISNULL(p.PJWeight, 0)  AS PJWeight_kg,
                    -- MPR: đếm từ MPR_Details trực tiếp
                    ISNULL((
                        SELECT SUM(md.Qty_Per_Sheet)
                        FROM MPR_Header mh
                        JOIN MPR_Details md ON md.MPR_ID = mh.MPR_ID
                        WHERE mh.Project_Code = p.ProjectCode
                    ), 0) AS MPR_Qty,
                    ISNULL((
                        SELECT SUM(md.Weight_kg * md.Qty_Per_Sheet)
                        FROM MPR_Header mh
                        JOIN MPR_Details md ON md.MPR_ID = mh.MPR_ID
                        WHERE mh.Project_Code = p.ProjectCode
                    ), 0) AS MPR_KG,
                    -- PO Qty/KG: từ PO_Detail, không qua MPR
                    ISNULL((
                        SELECT SUM(pod.Qty_Per_Sheet)
                        FROM PO_head ph
                        JOIN PO_Detail pod ON pod.PO_ID = ph.PO_ID
                        WHERE ph.WorkorderNo = p.WorkorderNo
                          AND ISNULL(ph.Status,'') <> 'Cancelled'
                    ), 0) AS PO_Qty,
                    -- PO_KG: SUM Weight_kg từ PO_Detail (Weight_kg = tổng KG dòng đó)
                    ISNULL((
                        SELECT SUM(pod.Weight_kg)
                        FROM PO_head ph
                        JOIN PO_Detail pod ON pod.PO_ID = ph.PO_ID
                        WHERE ph.WorkorderNo = p.WorkorderNo
                          AND ISNULL(ph.Status,'') <> 'Cancelled'
                    ), 0) AS PO_KG,
                    -- Budget: SUM từ PO_head — KHÔNG qua PO_Detail để tránh nhân bội
                    ISNULL((
                        SELECT SUM(ph.Total_Amount)
                        FROM PO_head ph
                        WHERE ph.WorkorderNo = p.WorkorderNo
                          AND ISNULL(ph.Status,'') <> 'Cancelled'
                    ), 0) AS PO_Amount
                FROM ProjectInfo p
                ORDER BY p.CreatedDate DESC, p.ProjectCode DESC";

            // Container chính
            // ── Header cố định (NGOÀI pRight, không bị scroll) ──────────────
            var pHeader = new Panel
            {
                Location = new Point(950, 80),
                Size = new Size(panelContent.Width - 970, 70),
                BackColor = Color.FromArgb(245, 248, 255),
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Left
            };
            panelContent.Controls.Add(pHeader);

            // ── pRight: scroll area cho cards, bên dưới header ──────────────
            var pRight = new Panel
            {
                Location = new Point(950, 152),
                Size = new Size(panelContent.Width - 970, panelContent.Height - 172),
                AutoScroll = true,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom | AnchorStyles.Left
            };
            panelContent.Controls.Add(pRight);

            pHeader.Controls.Add(new Label
            {
                Text = "📊  Project Progress",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 8),
                Size = new Size(200, 22)
            });

            // Hàng 2: txtSearch + cboFilter
            var txtSearch = new TextBox
            {
                Location = new Point(10, 36),
                Size = new Size(180, 24),
                Font = new Font("Segoe UI", 8),
                PlaceholderText = "🔍 Tìm dự án..."
            };
            pHeader.Controls.Add(txtSearch);

            pHeader.Controls.Add(new Label
            {
                Text = "Tiến độ:",
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.Gray,
                Location = new Point(198, 39),
                Size = new Size(50, 18)
            });
            var cboFilter = new ComboBox
            {
                Location = new Point(250, 36),
                Size = new Size(140, 24),
                Font = new Font("Segoe UI", 8),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboFilter.Items.AddRange(new object[]
            {
                "Tất cả",
                "Chưa đặt (0%)",
                "Đang đặt (1-99%)",
                "Hoàn thành (100%)"
            });
            cboFilter.SelectedIndex = 0;
            pHeader.Controls.Add(cboFilter);

            // Panel chứa các cards (có thể re-render)
            // pCards: không giới hạn chiều cao — tự mở rộng theo số dự án
            var pCards = new Panel
            {
                Location = new Point(0, 4),
                Size = new Size(pRight.Width - 18, 10), // chiều cao sẽ được set sau khi render
                AutoScroll = false,
                BackColor = Color.White
            };
            pRight.Controls.Add(pCards);

            // Load data một lần
            var rows = new System.Collections.Generic.List<(
                string code, string name,
                double budget, double pjKG,
                double mprQty, double mprKG,
                double poQty, double poKG, double poAmt,
                double pctQty, double pctBudg, double pctKG)>();

            try
            {
                using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                conn.Open();
                var dt = new System.Data.DataTable();
                dt.Load(new SqlCommand(SQL, conn).ExecuteReader());

                foreach (System.Data.DataRow r in dt.Rows)
                {
                    double budget = Convert.ToDouble(r["Budget"]);
                    double pjKG = Convert.ToDouble(r["PJWeight_kg"]);
                    double mprQty = Convert.ToDouble(r["MPR_Qty"]);
                    double mprKG = Convert.ToDouble(r["MPR_KG"]);
                    double poQty = Convert.ToDouble(r["PO_Qty"]);
                    double poKG = Convert.ToDouble(r["PO_KG"]);
                    double poAmt = Convert.ToDouble(r["PO_Amount"]);
                    // Hiển thị % thực — không cap 100 để phát hiện over-order/over-budget
                    double pctQty = mprQty > 0 ? Math.Round(poQty / mprQty * 100, 1) : 0;
                    double pctBudg = budget > 0 ? Math.Round(poAmt / budget * 100, 1) : 0;
                    double pctKG = pjKG > 0 ? Math.Round(poKG / pjKG * 100, 1) : 0;
                    rows.Add((
                        r["ProjectCode"]?.ToString() ?? "",
                        r["ProjectName"]?.ToString() ?? "",
                        budget, pjKG, mprQty, mprKG,
                        poQty, poKG, poAmt,
                        pctQty, pctBudg, pctKG));
                }
            }
            catch (Exception ex)
            {
                pCards.Controls.Add(new Label
                {
                    Text = "⚠ Lỗi tải tiến độ: " + ex.Message,
                    Font = new Font("Segoe UI", 9),
                    ForeColor = Color.Red,
                    Location = new Point(10, 10),
                    Size = new Size(pRight.Width - 20, 40)
                });
                return;
            }

            // ── Render cards theo thứ tự sort ───────────────────────────────
            void RenderCards(int _unused)
            {
                // Lấy filter hiện tại
                string kw = txtSearch.Text.Trim().ToLower();
                int cat = cboFilter.SelectedIndex; // 0=Tất cả,1=0%,2=1-99%,3=100%

                // Lọc theo search + tiến độ (dùng pctQty làm đại diện tiến độ đặt hàng)
                var sorted = rows
                    .FindAll(d =>
                    {
                        bool matchKW = string.IsNullOrEmpty(kw)
                            || d.code.ToLower().Contains(kw)
                            || d.name.ToLower().Contains(kw);
                        bool matchCat = cat switch
                        {
                            1 => d.pctQty == 0,
                            2 => d.pctQty > 0 && d.pctQty < 100,
                            3 => d.pctQty >= 100,
                            _ => true
                        };
                        return matchKW && matchCat;
                    })
                    .OrderByDescending(d => d.pctBudg)
                    .ToList();

                pRight.SuspendLayout();
                pCards.SuspendLayout();
                pCards.Controls.Clear();

                int yp = 6;
                int pw = pRight.Width - 24;

                if (sorted.Count == 0)
                {
                    pCards.Controls.Add(new Label
                    {
                        Text = "Không có dự án nào.",
                        Font = new Font("Segoe UI", 9),
                        ForeColor = Color.Gray,
                        Location = new Point(10, yp),
                        Size = new Size(pw, 24)
                    });
                    pCards.ResumeLayout();
                    return;
                }

                foreach (var d in sorted)
                {
                    var card = new Panel
                    {
                        Location = new Point(10, yp),
                        Size = new Size(pw, 106),
                        BackColor = Color.FromArgb(245, 248, 255),
                        BorderStyle = BorderStyle.FixedSingle
                    };
                    pCards.Controls.Add(card);

                    card.Controls.Add(new Label
                    {
                        Text = $"🏗 {d.code}  —  {d.name}",
                        Font = new Font("Segoe UI", 9, FontStyle.Bold),
                        ForeColor = Color.FromArgb(0, 80, 160),
                        Location = new Point(8, 6),
                        Size = new Size(pw - 16, 20)
                    });

                    void AddBar(string lbl, double pct, Color barColor, int barY)
                    {
                        int barW = pw - 190;
                        card.Controls.Add(new Label
                        {
                            Text = lbl,
                            Font = new Font("Segoe UI", 8),
                            ForeColor = Color.FromArgb(60, 60, 60),
                            Location = new Point(8, barY),
                            Size = new Size(108, 18)
                        });
                        var bg = new Panel
                        {
                            Location = new Point(120, barY + 2),
                            Size = new Size(barW, 14),
                            BackColor = Color.FromArgb(220, 225, 235)
                        };
                        card.Controls.Add(bg);
                        // fillW tối đa = barW (không tràn), nhưng % label hiển thị số thực
                        int fillW = (int)(barW * Math.Min(pct, 100) / 100);
                        if (fillW > 0)
                            bg.Controls.Add(new Panel
                            {
                                Location = new Point(0, 0),
                                Size = new Size(fillW, 14),
                                BackColor = pct > 100 ? Color.FromArgb(220, 53, 69)   // đỏ = vượt
                                           : pct >= 100 ? Color.FromArgb(40, 167, 69)  // xanh = đủ
                                           : pct >= 60 ? barColor                     // màu gốc
                                           : Color.FromArgb(255, 140, 0)               // cam = thấp
                            });
                        card.Controls.Add(new Label
                        {
                            Text = $"{pct:F1}%",
                            Font = new Font("Segoe UI", 8, FontStyle.Bold),
                            ForeColor = pct > 100 ? Color.FromArgb(220, 53, 69)  // đỏ = vượt
                                      : pct >= 100 ? Color.FromArgb(40, 167, 69)  // xanh = đủ
                                      : barColor,
                            Location = new Point(120 + barW + 6, barY),
                            Size = new Size(52, 18),
                            TextAlign = ContentAlignment.MiddleLeft
                        });
                    }

                    AddBar($"💰 Budget ({d.poAmt / 1e6:F1}M/{d.budget / 1e6:F1}M)",
                           d.pctBudg, Color.FromArgb(233, 30, 99), 30);
                    AddBar($"⚖ KG ({d.poKG:N0}/{d.pjKG:N0} kg)",
                           d.pctKG, Color.FromArgb(102, 51, 153), 52);

                    yp += 114;
                }

                // Set chiều cao pCards = nội dung thực, pRight tự scroll
                pCards.Height = yp + 10;
                pCards.Width = pRight.Width - 18;
                pRight.AutoScrollMinSize = new Size(0, yp + 10);
                pCards.ResumeLayout();
                pRight.ResumeLayout();
            }

            // Render lần đầu
            RenderCards(0);

            // Filter events
            txtSearch.TextChanged += (s, ev) => RenderCards(0);
            cboFilter.SelectedIndexChanged += (s, ev) => RenderCards(0);
        }

        private void AddCard(string title, string value, Color color, int x, int y)
        {
            var card = new Panel
            {
                Location = new Point(x, y),
                Size = new Size(210, 175),
                BackColor = color,
                Cursor = Cursors.Hand
            };

            var lblTitle = new Label
            {
                Text = title,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(0, 15),
                Size = new Size(210, 28),
                TextAlign = ContentAlignment.MiddleCenter
            };

            var lblValue = new Label
            {
                Text = value,
                Font = new Font("Segoe UI", 36, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(0, 48),
                Size = new Size(210, 80),
                TextAlign = ContentAlignment.MiddleCenter
            };

            var lblSub = new Label
            {
                Text = (value.Contains("Xem")) ? "nhấn để mở" : "bản ghi",
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.FromArgb(220, 255, 255, 255),
                Location = new Point(0, 135),
                Size = new Size(210, 25),
                TextAlign = ContentAlignment.MiddleCenter
            };

            card.Controls.Add(lblTitle);
            card.Controls.Add(lblValue);
            card.Controls.Add(lblSub);
            panelContent.Controls.Add(card);

            EventHandler clickHandler = (s, e) =>
            {
                foreach (Control c in panelMenu.Controls)
                    if (c is Button b) b.BackColor = Color.FromArgb(30, 30, 45);

                if (title.Contains("Nhà cung cấp"))
                { HighlightMenu("Nhà cung cấp"); OpenForm(new frmSupplier()); }
                else if (title.Contains("MPR"))
                { HighlightMenu("MPR"); OpenForm(new frmMPR()); }
                else if (title.Contains("PO"))
                { HighlightMenu("PO"); OpenForm(new frmPO()); }
                else if (title.Contains("RIR"))
                { HighlightMenu("RIR"); OpenForm(new frmRIR()); }
                else if (title.Contains("Dự án"))
                { HighlightMenu("Dự án"); OpenForm(new frmProject()); }
                else if (title.Contains("Kho"))
                { HighlightMenu("Kho vật tư"); OpenForm(new frmWarehouses_v2()); }
                else if (title.Contains("Dashboard"))
                { HighlightMenu("Dashboard"); OpenForm(new frmDashboard()); }
            };

            card.Click += clickHandler;
            lblTitle.Click += clickHandler;
            lblValue.Click += clickHandler;
            lblSub.Click += clickHandler;

            EventHandler enterH = (s, e) => card.BackColor = ControlPaint.Dark(color, 0.1f);
            EventHandler leaveH = (s, e) => card.BackColor = color;
            card.MouseEnter += enterH; card.MouseLeave += leaveH;
            lblTitle.MouseEnter += enterH; lblTitle.MouseLeave += leaveH;
            lblValue.MouseEnter += enterH; lblValue.MouseLeave += leaveH;
            lblSub.MouseEnter += enterH; lblSub.MouseLeave += leaveH;
        }

        private void HighlightMenu(string keyword)
        {
            foreach (Control c in panelMenu.Controls)
                if (c is Button b && b.Tag?.ToString().Contains(keyword) == true)
                    b.BackColor = Color.FromArgb(0, 120, 212);
        }
        // =====================================================================
        //  NOTIFICATION CHATBOX
        // =====================================================================
        private void BuildNotifyPanel()
        {
            // ── NUT THONG BAO 50x50 voi icon dep ──────────────────────────
            _btnNotify = new Button
            {
                Size = new Size(50, 50),
                BackColor = Color.FromArgb(25, 118, 210),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
                TabStop = false
            };
            _btnNotify.FlatAppearance.BorderSize = 0;
            _btnNotify.FlatAppearance.MouseOverBackColor = Color.FromArgb(21, 101, 192);

            // Ve chu Notice dep voi gradient
            _btnNotify.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                var btn = _btnNotify;
                var rc = btn.ClientRectangle;

                // Gradient background
                Color c1 = _notifyBadge > 0 ? Color.FromArgb(198, 40, 40) : Color.FromArgb(30, 136, 229);
                Color c2 = _notifyBadge > 0 ? Color.FromArgb(244, 67, 54) : Color.FromArgb(21, 101, 192);
                using var bgBrush = new System.Drawing.Drawing2D.LinearGradientBrush(
                    rc, c1, c2, System.Drawing.Drawing2D.LinearGradientMode.Vertical);
                g.FillRectangle(bgBrush, rc);

                // Duong ke trang mo phia tren
                using var linePen = new Pen(Color.FromArgb(60, 255, 255, 255), 1);
                g.DrawLine(linePen, 4, 1, rc.Width - 4, 1);

                // Icon chong nho phia tren
                int bx = rc.Width / 2;
                using var wBrush = new SolidBrush(Color.FromArgb(220, 255, 255, 255));
                var bp = new System.Drawing.Drawing2D.GraphicsPath();
                bp.AddArc(bx - 6, 6, 12, 8, 180, 180);
                bp.AddLine(bx + 6, 10, bx + 7, 16);
                bp.AddLine(bx + 7, 16, bx - 7, 16);
                bp.CloseFigure();
                g.FillPath(wBrush, bp);
                g.FillEllipse(wBrush, bx - 3, 16, 6, 3);

                // Badge so
                if (_notifyBadge > 0)
                {
                    var bRect = new Rectangle(bx + 2, 4, 16, 12);
                    g.FillEllipse(Brushes.White, bRect);
                    g.DrawString(_notifyBadge > 99 ? "99+" : _notifyBadge.ToString(),
                        new Font("Segoe UI", 6, FontStyle.Bold),
                        new SolidBrush(Color.FromArgb(198, 40, 40)), bRect,
                        new StringFormat
                        {
                            Alignment = StringAlignment.Center,
                            LineAlignment = StringAlignment.Center
                        });
                }

                // Chu "Notice"
                var sf = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Far
                };
                var textRect = new RectangleF(0, rc.Height - 22, rc.Width, 20);
                g.DrawString("Notice",
                    new Font("Segoe UI", 8, FontStyle.Bold),
                    Brushes.White, textRect, sf);

                // Duong ke trang mo phia duoi
                g.DrawLine(linePen, 4, rc.Height - 2, rc.Width - 4, rc.Height - 2);
            };

            _btnNotify.Location = new Point((190 - 50) / 2, panelMenu.Height - 60);
            panelMenu.Controls.Add(_btnNotify);
            _btnNotify.BringToFront();

            panelMenu.Resize += (s, e) =>
            {
                if (_btnNotify != null)
                    _btnNotify.Location = new Point((190 - 50) / 2, panelMenu.Height - 60);
                if (_panelNotify != null && _panelNotify.Visible)
                    PositionNotifyPanel();
            };

            // ── PANEL CHATBOX voi resize/drag ─────────────────────────────
            _panelNotify = new Panel
            {
                Size = new Size(340, 440),
                BackColor = Color.FromArgb(250, 250, 252),
                BorderStyle = BorderStyle.None,
                Visible = false
            };

            // Bo rang dep bang Paint
            _panelNotify.Paint += (s, e) =>
            {
                var g = e.Graphics;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                using var pen = new Pen(Color.FromArgb(200, 200, 220), 1.5f);
                g.DrawRectangle(pen, 0, 0, _panelNotify.Width - 1, _panelNotify.Height - 1);
                // Bo do phai: resize handle
                g.FillPolygon(new SolidBrush(Color.FromArgb(180, 180, 200)), new Point[]
                {
                    new Point(_panelNotify.Width - 1, _panelNotify.Height - 16),
                    new Point(_panelNotify.Width - 1, _panelNotify.Height - 1),
                    new Point(_panelNotify.Width - 16, _panelNotify.Height - 1)
                });
            };

            // Header gradient
            var pHead = new Panel
            {
                Location = new Point(1, 1),
                Size = new Size(_panelNotify.Width - 2, 48),
                BackColor = Color.FromArgb(25, 118, 210)
            };
            pHead.Paint += (s, e) =>
            {
                using var brush = new System.Drawing.Drawing2D.LinearGradientBrush(
                    pHead.ClientRectangle,
                    Color.FromArgb(30, 136, 229),
                    Color.FromArgb(21, 101, 192),
                    System.Drawing.Drawing2D.LinearGradientMode.Vertical);
                e.Graphics.FillRectangle(brush, pHead.ClientRectangle);
            };
            _panelNotify.Controls.Add(pHead);
            _panelNotify.Resize += (s, e) =>
                pHead.Size = new Size(_panelNotify.Width - 2, 48);

            // Icon chong trong header
            var lblIcon = new Label
            {
                Text = "Bell",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(180, 220, 255),
                Location = new Point(12, 14),
                Size = new Size(30, 20),
                AutoSize = false
            };
            // Ve icon nho trong header
            lblIcon.Paint += (s, e) =>
            {
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                var g = e.Graphics;
                g.FillEllipse(new SolidBrush(Color.FromArgb(200, 255, 255, 255)),
                    2, 0, 18, 18);
                g.DrawArc(new Pen(Color.FromArgb(25, 118, 210), 3), 5, 2, 12, 12, 180, 180);
                g.FillRectangle(new SolidBrush(Color.FromArgb(25, 118, 210)), 7, 8, 8, 6);
                g.FillEllipse(new SolidBrush(Color.FromArgb(25, 118, 210)), 9, 13, 4, 3);
            };
            pHead.Controls.Add(lblIcon);

            _lblNotifyCount = new Label
            {
                Text = "Thong bao he thong",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(38, 14),
                Size = new Size(220, 20),
                AutoEllipsis = true
            };
            pHead.Controls.Add(_lblNotifyCount);
            pHead.Resize += (s, e) =>
                _lblNotifyCount.Size = new Size(pHead.Width - 80, 20);

            // Nut dong
            var btnX = new Button
            {
                Text = "x",
                Size = new Size(30, 30),
                BackColor = Color.Transparent,
                ForeColor = Color.FromArgb(200, 230, 255),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 11),
                Cursor = Cursors.Hand,
                TabStop = false
            };
            btnX.FlatAppearance.BorderSize = 0;
            btnX.FlatAppearance.MouseOverBackColor = Color.FromArgb(40, 255, 255, 255);
            btnX.Click += (s, e) => _panelNotify.Visible = false;
            pHead.Controls.Add(btnX);
            pHead.Resize += (s, e2) => btnX.Location = new Point(pHead.Width - 34, 9);
            btnX.Location = new Point(pHead.Width - 34, 9);

            // ── Drag panel bang header ──────────────────────────────────
            Point _dragStart = Point.Empty;
            bool _dragging = false;
            pHead.MouseDown += (s, e) =>
            {
                if (e.Button == MouseButtons.Left)
                { _dragging = true; _dragStart = e.Location; pHead.Cursor = Cursors.SizeAll; }
            };
            pHead.MouseMove += (s, e) =>
            {
                if (_dragging)
                {
                    int dx = e.X - _dragStart.X;
                    int dy = e.Y - _dragStart.Y;
                    _panelNotify.Location = new Point(
                        Math.Max(0, _panelNotify.Left + dx),
                        Math.Max(55, _panelNotify.Top + dy));
                }
            };
            pHead.MouseUp += (s, e) =>
            { _dragging = false; pHead.Cursor = Cursors.Default; };
            _lblNotifyCount.MouseDown += (s, e) =>
            { if (e.Button == MouseButtons.Left) { _dragging = true; _dragStart = e.Location; pHead.Cursor = Cursors.SizeAll; } };
            _lblNotifyCount.MouseMove += (s, e) =>
            { if (_dragging) { _panelNotify.Location = new Point(Math.Max(0, _panelNotify.Left + e.X - _dragStart.X), Math.Max(55, _panelNotify.Top + e.Y - _dragStart.Y)); } };
            _lblNotifyCount.MouseUp += (s, e) =>
            { _dragging = false; pHead.Cursor = Cursors.Default; };

            // ── Toolbar ─────────────────────────────────────────────────
            // Tab bar: Tat ca | Q Trong
            var pTab = new Panel
            {
                Location = new Point(1, 49),
                Size = new Size(_panelNotify.Width - 2, 30),
                BackColor = Color.FromArgb(240, 242, 248)
            };
            _panelNotify.Controls.Add(pTab);
            _panelNotify.Resize += (s, e) => pTab.Size = new Size(_panelNotify.Width - 2, 30);

            var btnTabAll = new Button
            {
                Text = "Tat ca",
                Size = new Size(80, 28),
                Location = new Point(2, 1),
                BackColor = Color.FromArgb(25, 118, 210),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                Cursor = Cursors.Hand,
                TabStop = false,
                Tag = "ALL"
            };
            btnTabAll.FlatAppearance.BorderSize = 0;
            pTab.Controls.Add(btnTabAll);

            var btnTabStar = new Button
            {
                Text = "Q Trong (0)",
                Size = new Size(80, 28),
                Location = new Point(84, 1),
                BackColor = Color.FromArgb(240, 242, 248),
                ForeColor = Color.FromArgb(80, 80, 80),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                Cursor = Cursors.Hand,
                TabStop = false,
                Tag = "STAR"
            };
            btnTabStar.FlatAppearance.BorderSize = 0;
            _btnTabStar = btnTabStar; // luu ref de cap nhat so luong
            pTab.Controls.Add(btnTabStar);

            // Switch tab logic
            string _activeTab = "ALL";

            // Helper cap nhat so luong Q Trong tren button tab
            Action updateStarCount = () =>
            {
                if (_btnTabStar != null)
                {
                    _btnTabStar.Text = "Q Trong (" + _starredMsgs.Count + ")";
                    _btnTabStar.ForeColor = _starredMsgs.Count > 0
                        ? Color.FromArgb(198, 40, 40) : Color.FromArgb(80, 80, 80);
                }
            };

            Action refreshList = () =>
            {
                _lstNotify.BeginUpdate();
                _lstNotify.Items.Clear();
                var src = _activeTab == "STAR"
                    ? new System.Collections.Generic.List<string>(_starredMsgs)
                    : new System.Collections.Generic.List<string>(_seenMsgs);

                // Sap xep theo thoi gian trong chuoi msg (format: "| dd/MM HH:mm - ")
                src.Sort((a, b) =>
                {
                    DateTime dtA = ParseMsgTime(a);
                    DateTime dtB = ParseMsgTime(b);
                    return dtB.CompareTo(dtA); // moi nhat len dau
                });

                foreach (var m in src) _lstNotify.Items.Add(m);
                _lstNotify.EndUpdate();
                updateStarCount();
            };

            btnTabAll.Click += (s, e) =>
            {
                _activeTab = "ALL";
                btnTabAll.BackColor = Color.FromArgb(25, 118, 210); btnTabAll.ForeColor = Color.White;
                btnTabStar.BackColor = Color.FromArgb(240, 242, 248); btnTabStar.ForeColor = Color.FromArgb(80, 80, 80);
                refreshList();
            };
            btnTabStar.Click += (s, e) =>
            {
                _activeTab = "STAR";
                btnTabStar.BackColor = Color.FromArgb(255, 193, 7); btnTabStar.ForeColor = Color.FromArgb(30, 30, 30);
                btnTabAll.BackColor = Color.FromArgb(240, 242, 248); btnTabAll.ForeColor = Color.FromArgb(80, 80, 80);
                refreshList();
            };

            var pToolbar = new Panel
            {
                Location = new Point(1, 79),
                Size = new Size(_panelNotify.Width - 2, 34),
                BackColor = Color.FromArgb(245, 247, 252)
            };
            _panelNotify.Controls.Add(pToolbar);
            _panelNotify.Resize += (s, e) =>
                pToolbar.Size = new Size(_panelNotify.Width - 2, 34);

            var btnRef = new Button
            {
                Text = "Lam moi",
                Size = new Size(90, 26),
                Location = new Point(8, 6),
                BackColor = Color.FromArgb(46, 125, 50),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                Cursor = Cursors.Hand,
                TabStop = false
            };
            btnRef.FlatAppearance.BorderSize = 0;
            btnRef.FlatAppearance.MouseOverBackColor = Color.FromArgb(27, 94, 32);
            btnRef.Click += (s, e) => CheckAndNotify(true);
            pToolbar.Controls.Add(btnRef);

            var btnClr = new Button
            {
                Text = "Xoa tat ca",
                Size = new Size(90, 26),
                Location = new Point(104, 6),
                BackColor = Color.FromArgb(84, 84, 84),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                Cursor = Cursors.Hand,
                TabStop = false
            };
            btnClr.FlatAppearance.BorderSize = 0;
            btnClr.FlatAppearance.MouseOverBackColor = Color.FromArgb(60, 60, 60);
            btnClr.Click += (s, e) =>
            {
                _lstNotify.Items.Clear();
                _notifyBadge = 0;
                _lblNotifyCount.Text = "Thong bao he thong";
                _btnNotify.BackColor = Color.FromArgb(25, 118, 210);
                _btnNotify.Invalidate();
            };
            pToolbar.Controls.Add(btnClr);

            // Nút "Gửi TB" — chỉ hiển thị cho Admin
            if (AppSession.IsAdmin)
            {
                var btnSendNotif = new Button
                {
                    Text = "📢 Gửi TB",
                    Size = new Size(85, 26),
                    Location = new Point(200, 6),
                    BackColor = Color.FromArgb(198, 40, 40),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 8, FontStyle.Bold),
                    Cursor = Cursors.Hand,
                    TabStop = false
                };
                btnSendNotif.FlatAppearance.BorderSize = 0;
                btnSendNotif.FlatAppearance.MouseOverBackColor = Color.FromArgb(160, 20, 20);
                new ToolTip().SetToolTip(btnSendNotif, "Gửi thông báo nội bộ đến người dùng khác");
                btnSendNotif.Click += (s, e) =>
                {
                    _panelNotify.Visible = false;
                    var frm = new frmSendNotification();
                    frm.ShowDialog(this);
                };
                pToolbar.Controls.Add(btnSendNotif);
            }

            // ── Danh sach thong bao ─────────────────────────────────────
            _lstNotify = new ListBox
            {
                Location = new Point(1, 117),
                Size = new Size(_panelNotify.Width - 2, _panelNotify.Height - 140),
                Font = new Font("Segoe UI", 9),
                BorderStyle = BorderStyle.None,
                BackColor = Color.White,
                ItemHeight = 50,
                DrawMode = DrawMode.OwnerDrawFixed,
                IntegralHeight = false
            };
            _lstNotify.DrawItem += NotifyDrawItem;

            // Click vao ngoi sao (goc phai) de star/unstar
            _lstNotify.MouseClick += (s, e) =>
            {
                int idx = _lstNotify.IndexFromPoint(e.Location);
                if (idx < 0) return;
                var itemRect = _lstNotify.GetItemRectangle(idx);
                string msg = _lstNotify.Items[idx].ToString();

                // Vung ngoi sao: goc phai
                var starZone = new Rectangle(itemRect.Right - 30, itemRect.Y, 30, itemRect.Height);
                if (starZone.Contains(e.Location))
                {
                    // Toggle star
                    if (_starredMsgs.Contains(msg))
                        _starredMsgs.Remove(msg);
                    else
                        _starredMsgs.Insert(0, msg);
                    if (_btnTabStar != null)
                    {
                        _btnTabStar.Text = "Q Trong (" + _starredMsgs.Count + ")";
                        _btnTabStar.ForeColor = _starredMsgs.Count > 0
                            ? Color.FromArgb(198, 40, 40) : Color.FromArgb(80, 80, 80);
                    }
                }
                else
                {
                    // Click vao than dong → danh dau da doc
                    if (!_readMsgs.Contains(msg))
                    {
                        _readMsgs.Add(msg);
                        // Nếu là thông báo nội bộ → đánh dấu đã đọc trên DB
                        if (msg.StartsWith("[NOTIFY]"))
                        {
                            int notifId = ExtractNotifId(msg);
                            if (notifId > 0)
                                Task.Run(() => { try { _internalNotifSvc.MarkRead(notifId); } catch { } });
                        }
                        UpdateBadge();
                    }
                    // Double-click sẽ show detail — đây single click chỉ đánh dấu đọc
                    if (msg.StartsWith("[NOTIFY]"))
                        ShowInternalNotifyDetail(msg);
                }
                _lstNotify.Invalidate(_lstNotify.GetItemRectangle(idx));
            };

            _lstNotify.DoubleClick += (s, e) =>
            {
                if (_lstNotify.SelectedIndex < 0) return;
                bool isAdmin = AppSession.CurrentUser?.Role_ID == 1
                               || AppSession.CurrentUser?.Role_Name?.Equals("Admin",
                                  StringComparison.OrdinalIgnoreCase) == true;
                if (!isAdmin)
                {
                    this.BringToFront(); this.Activate();
                    MessageBox.Show(this, "Chi tai khoan Admin moi xem duoc chi tiet thong bao!",
                        "Khong co quyen", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                string msg = _lstNotify.Items[_lstNotify.SelectedIndex].ToString();
                ShowNotifyDetail(msg);
            };
            _panelNotify.Controls.Add(_lstNotify);
            _panelNotify.Resize += (s, e) =>
            {
                _lstNotify.Size = new Size(_panelNotify.Width - 2, _panelNotify.Height - 140);
                _panelNotify.Invalidate();
            };

            // Footer
            var lblFoot = new Label
            {
                Name = "lblFoot",
                Text = "Tu dong cap nhat moi 5 phut",
                Font = new Font("Segoe UI", 7, FontStyle.Italic),
                ForeColor = Color.FromArgb(160, 160, 180),
                BackColor = Color.FromArgb(245, 247, 252),
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Bottom,
                Height = 20,
                Padding = new Padding(8, 0, 0, 0)
            };
            _panelNotify.Controls.Add(lblFoot);

            // ── Resize handle o goc duoi phai ───────────────────────────
            const int GRIP = 16;
            bool _resizing = false;
            Point _resizeStart = Point.Empty;
            Size _resizeStartSize = Size.Empty;

            _panelNotify.MouseDown += (s, e) =>
            {
                if (e.Button == MouseButtons.Left)
                {
                    bool inGrip = e.X >= _panelNotify.Width - GRIP && e.Y >= _panelNotify.Height - GRIP;
                    if (inGrip)
                    {
                        _resizing = true;
                        _resizeStart = this.PointToScreen(e.Location) +
                            new Size(_panelNotify.Location);
                        _resizeStartSize = _panelNotify.Size;
                        _panelNotify.Cursor = Cursors.SizeNWSE;
                    }
                }
            };
            _panelNotify.MouseMove += (s, e) =>
            {
                if (_resizing)
                {
                    Point cur = this.PointToScreen(e.Location) + new Size(_panelNotify.Location);
                    int newW = Math.Max(280, _resizeStartSize.Width + (cur.X - _resizeStart.X));
                    int newH = Math.Max(300, _resizeStartSize.Height + (cur.Y - _resizeStart.Y));
                    _panelNotify.Size = new Size(newW, newH);
                }
                else
                {
                    bool inGrip = e.X >= _panelNotify.Width - GRIP && e.Y >= _panelNotify.Height - GRIP;
                    _panelNotify.Cursor = inGrip ? Cursors.SizeNWSE : Cursors.Default;
                }
            };
            _panelNotify.MouseUp += (s, e) =>
            { _resizing = false; _panelNotify.Cursor = Cursors.Default; };

            this.Controls.Add(_panelNotify);
            _panelNotify.BringToFront();

            // Toggle
            _btnNotify.Click += (s, e) =>
            {
                _panelNotify.Visible = !_panelNotify.Visible;
                if (_panelNotify.Visible)
                {
                    PositionNotifyPanel();
                    _panelNotify.BringToFront();
                }
            };
        }


        private void UpdateBadge()
        {
            // Badge = so msg trong _seenMsgs chua co trong _readMsgs
            int unread = 0;
            foreach (var m in _seenMsgs)
                if (!_readMsgs.Contains(m)) unread++;
            _notifyBadge = unread;
            _btnNotify.Invalidate();
        }

        private void PositionNotifyPanel()
        {
            int menuW = 190;
            int x = menuW + 5;
            int y = this.ClientSize.Height - _panelNotify.Height - 10;
            _panelNotify.Location = new Point(x, Math.Max(55, y));
        }

        private string GetMsgKey(string msg)
        {
            // Key ngan de nhan dien thong bao
            var p = msg.Split('|');
            return p.Length > 0 ? p[0].Trim() : msg.Substring(0, Math.Min(40, msg.Length));
        }

        private void NotifyDrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            string msg = _lstNotify.Items[e.Index].ToString();
            bool isPO = msg.StartsWith("[PO]");
            bool isMPR = msg.StartsWith("[MPR]");
            bool isRIR = msg.StartsWith("[RIR]");
            bool isNotify = msg.StartsWith("[NOTIFY]");

            bool isRead = _readMsgs.Contains(msg);
            Color bg = isRead
                ? Color.FromArgb(242, 242, 242)  // da doc: xam nhat
                : (e.Index % 2 == 0 ? Color.White : Color.FromArgb(245, 247, 252));
            e.Graphics.FillRectangle(new SolidBrush(bg), e.Bounds);

            Color bar = isPO ? Color.FromArgb(25, 118, 210) :
                        isMPR ? Color.FromArgb(46, 125, 50) :
                        isRIR ? Color.FromArgb(102, 51, 153) :
                        isNotify ? Color.FromArgb(198, 40, 40) :
                                Color.FromArgb(180, 180, 200);
            // Thanh mau: nhat hon neu da doc
            Color barDraw = isRead ? Color.FromArgb(180, bar.R, bar.G, bar.B) : bar;
            e.Graphics.FillRectangle(new SolidBrush(barDraw),
                new Rectangle(e.Bounds.X, e.Bounds.Y + 2, 4, e.Bounds.Height - 4));

            // Parse: [TYPE] Ten | Du an | Time - User
            // Bo @@timestamp truoc khi hien thi
            string cleanMsg = msg.Contains("@@") ? msg.Substring(0, msg.LastIndexOf("@@")) : msg;
            string[] parts = cleanMsg.Split('|');
            string l1 = parts.Length > 0 ? parts[0].Trim().Replace("[PO] ", "").Replace("[MPR] ", "").Replace("[RIR] ", "") : cleanMsg;
            string l2 = parts.Length > 1 ? parts[1].Trim() : "";
            string l3 = parts.Length > 2 ? parts[2].Trim() : "";

            Color textColor = isRead ? Color.FromArgb(140, 140, 140) : bar;
            FontStyle fs = isRead ? FontStyle.Regular : FontStyle.Bold;
            e.Graphics.DrawString(l1, new Font("Segoe UI", 9, fs),
                new SolidBrush(textColor),
                new RectangleF(e.Bounds.X + 10, e.Bounds.Y + 4, e.Bounds.Width - 14, 18));
            if (!string.IsNullOrEmpty(l2))
                e.Graphics.DrawString(l2, new Font("Segoe UI", 8), Brushes.DimGray,
                    new RectangleF(e.Bounds.X + 10, e.Bounds.Y + 22, e.Bounds.Width - 14, 15));
            if (!string.IsNullOrEmpty(l3))
                e.Graphics.DrawString(l3, new Font("Segoe UI", 7, FontStyle.Italic),
                    new SolidBrush(Color.FromArgb(150, 150, 170)),
                    new RectangleF(e.Bounds.X + 10, e.Bounds.Y + 36, e.Bounds.Width - 14, 12));

            // Ngoi sao o goc phai
            bool isStarred = _starredMsgs.Contains(msg);
            var starRect = new Rectangle(e.Bounds.Right - 26, e.Bounds.Y + (e.Bounds.Height - 18) / 2, 18, 18);
            DrawStar(e.Graphics, starRect, isStarred);

            e.Graphics.DrawLine(new Pen(Color.FromArgb(230, 230, 240)),
                e.Bounds.X + 6, e.Bounds.Bottom - 1, e.Bounds.Right - 6, e.Bounds.Bottom - 1);
        }

        private void DrawStar(System.Drawing.Graphics g, Rectangle r, bool filled)
        {
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            int cx = r.X + r.Width / 2, cy = r.Y + r.Height / 2;
            int outer = r.Width / 2, inner = outer / 2;
            var pts = new System.Drawing.Point[10];
            for (int i = 0; i < 10; i++)
            {
                double angle = Math.PI / 5 * i - Math.PI / 2;
                int radius = (i % 2 == 0) ? outer : inner;
                pts[i] = new System.Drawing.Point(
                    cx + (int)(radius * Math.Cos(angle)),
                    cy + (int)(radius * Math.Sin(angle)));
            }
            if (filled)
                g.FillPolygon(new SolidBrush(Color.FromArgb(255, 193, 7)), pts);
            else
                g.DrawPolygon(new Pen(Color.FromArgb(180, 180, 200), 1.2f), pts);
        }

        // ── Persist last check time giữa các lần mở app ────────────────
        private DateTime LoadLastCheckTime()
        {
            try
            {
                if (File.Exists(_lastCheckFile))
                {
                    string text = File.ReadAllText(_lastCheckFile).Trim();
                    if (DateTime.TryParse(text, out DateTime dt))
                        return dt;
                }
            }
            catch { }
            // Mặc định: lùi 7 ngày để hiển thị các thông báo 7 ngày qua
            return DateTime.Now.AddDays(-7);
        }

        private void SaveLastCheckTime()
        {
            try
            {
                File.WriteAllText(_lastCheckFile, DateTime.Now.ToString("o"));
            }
            catch { }
        }

        // ── Phương thức tĩnh cho form con gọi để trigger check ngay ────
        public static void TriggerNewDataCheck()
        {
            RaiseNewDataCreated();
        }

        private void StartNotifyTimer()
        {
            // Load thời gian check cuối từ file (persist giữa các lần mở app)
            _lastCheckTime = LoadLastCheckTime();

            // Đăng ký lắng nghe sự kiện từ form con (PO, MPR vừa được tạo)
            NewDataCreated += (s, e) =>
            {
                try
                {
                    if (!this.IsDisposed)
                        this.BeginInvoke(new Action(() => CheckAndNotify(true)));
                }
                catch { }
            };

            _notifyTimer = new System.Windows.Forms.Timer { Interval = 30 * 1000 }; // 30 giây
            _notifyTimer.Tick += (s, e) =>
            {
                try { CheckAndNotify(false); }
                catch { }
            };
            _notifyTimer.Start();

            // Timer riêng cho thông báo nội bộ — check mỗi 10 giây, độc lập với PO/MPR
            _internalNotifTimer = new System.Windows.Forms.Timer { Interval = 10 * 1000 };
            _internalNotifTimer.Tick += (s, e) =>
            {
                try { CheckInternalNotifications(); }
                catch { }
            };
            _internalNotifTimer.Start();

            // Check ngay khi app vừa mở
            this.BeginInvoke(new Action(() =>
            {
                try
                {
                    CheckAndNotify(true);
                    CheckInternalNotifications();
                }
                catch { }
            }));
        }

        private void CheckInternalNotifications()
        {
            if (AppSession.CurrentUser == null) return;
            try
            {
                var unread = _internalNotifSvc.GetUnread(AppSession.CurrentUser.Username);
                // Lọc những thông báo chưa từng hiển thị trong session này
                var fresh = unread.FindAll(n => !_shownNotifIds.Contains(n.Notif_ID));
                if (fresh.Count == 0) return;

                foreach (var n in fresh)
                    _shownNotifIds.Add(n.Notif_ID);

                // Chuyển sang UI thread
                if (this.InvokeRequired)
                    this.Invoke(new Action(() => ApplyInternalNotifications(fresh)));
                else
                    ApplyInternalNotifications(fresh);
            }
            catch { }
        }

        private void ApplyInternalNotifications(System.Collections.Generic.List<InternalNotification> items)
        {
            if (items == null || items.Count == 0) return;

            // Thêm vào listbox notification panel
            foreach (var n in items)
            {
                string isoTs = n.Sent_At.ToString("o");
                string dt = n.Sent_At.ToString("dd/MM HH:mm");
                string msg = $"[NOTIFY] {n.Title} | Từ: {n.Sender_FullName} | {dt}@@{isoTs}|notif_id={n.Notif_ID}";
                _seenMsgs.Add(msg);
            }

            // Rebuild listbox
            _lstNotify.BeginUpdate();
            _lstNotify.Items.Clear();
            var allMsgs = new System.Collections.Generic.List<string>(_seenMsgs);
            allMsgs.Sort((a, b) => ParseMsgTime(b).CompareTo(ParseMsgTime(a)));
            foreach (var m in allMsgs) _lstNotify.Items.Add(m);
            _lstNotify.EndUpdate();

            _unreadCount += items.Count;
            UpdateBadge();

            try { System.Media.SystemSounds.Exclamation.Play(); } catch { }

            // Hiển thị toast cảnh báo nổi cho từng thông báo
            frmNotificationToast.ShowToasts(items, this);
        }

        private void CheckAndNotify(bool force)
        {
            try
            {
                var msgs = new System.Collections.Generic.List<string>();
                // Khi force: lay 5 ban ghi moi nhat bat ke thoi gian
                // Khi tu dong: lay ban ghi moi hon _lastCheckTime
                DateTime since = force ? DateTime.Now.AddDays(-7) : _lastCheckTime;

                using (var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection())
                {
                    conn.Open();

                    // ── PO moi ──────────────────────────────────────────
                    string sqlPO = "SELECT TOP " + (force ? "5" : "20") + " PONo, Project_Name, Created_Date, Created_By " +
                                   "FROM PO_head WHERE Created_Date > @s ORDER BY Created_Date DESC";
                    using var cPO = new Microsoft.Data.SqlClient.SqlCommand(sqlPO, conn);
                    cPO.Parameters.AddWithValue("@s", since);
                    using var rPO = cPO.ExecuteReader();
                    while (rPO.Read())
                    {
                        DateTime rawDtPO = rPO["Created_Date"] != DBNull.Value
                            ? Convert.ToDateTime(rPO["Created_Date"]) : DateTime.MinValue;
                        // Chuyển đổi sang múi giờ hệ thống địa phương
                        if (rawDtPO != DateTime.MinValue && rawDtPO.Kind == DateTimeKind.Unspecified)
                            rawDtPO = DateTime.SpecifyKind(rawDtPO, DateTimeKind.Utc).ToLocalTime();
                        string dt = rawDtPO != DateTime.MinValue ? rawDtPO.ToString("dd/MM HH:mm") : "";
                        string usr = rPO["Created_By"]?.ToString() ?? "";
                        string isoTs = rawDtPO.ToString("o"); // ISO timestamp de sort
                        msgs.Add("[PO] PO moi: " + rPO["PONo"] + " | " + rPO["Project_Name"] + " | " + dt + " - " + usr + "@@" + isoTs);
                    }
                    rPO.Close();

                    // ── MPR moi hoac cap nhat ────────────────────────────
                    string sqlMPR = "SELECT TOP " + (force ? "5" : "20") + " MPR_No, Project_Name, Modified_Date, Modified_By " +
                                    "FROM MPR_Header WHERE Modified_Date > @s ORDER BY Modified_Date DESC";
                    using var cMPR = new Microsoft.Data.SqlClient.SqlCommand(sqlMPR, conn);
                    cMPR.Parameters.AddWithValue("@s", since);
                    using var rMPR = cMPR.ExecuteReader();
                    while (rMPR.Read())
                    {
                        DateTime rawDtMPR = rMPR["Modified_Date"] != DBNull.Value
                            ? Convert.ToDateTime(rMPR["Modified_Date"]) : DateTime.MinValue;
                        // Chuyển đổi sang múi giờ hệ thống địa phương
                        if (rawDtMPR != DateTime.MinValue && rawDtMPR.Kind == DateTimeKind.Unspecified)
                            rawDtMPR = DateTime.SpecifyKind(rawDtMPR, DateTimeKind.Utc).ToLocalTime();
                        string dt = rawDtMPR != DateTime.MinValue ? rawDtMPR.ToString("dd/MM HH:mm") : "";
                        string usr = rMPR["Modified_By"]?.ToString() ?? "";
                        string isoTs = rawDtMPR.ToString("o");
                        msgs.Add("[MPR] MPR cap nhat: " + rMPR["MPR_No"] + " | " + rMPR["Project_Name"] + " | " + dt + " - " + usr + "@@" + isoTs);
                    }
                    rMPR.Close();

                    // ── MPR moi (Created_Date) neu chua co ──────────────
                    string sqlMPRNew = "SELECT TOP " + (force ? "5" : "20") + " MPR_No, Project_Name, Created_Date, Created_By " +
                                       "FROM MPR_Header WHERE Created_Date > @s ORDER BY Created_Date DESC";
                    using var cMN = new Microsoft.Data.SqlClient.SqlCommand(sqlMPRNew, conn);
                    cMN.Parameters.AddWithValue("@s", since);
                    using var rMN = cMN.ExecuteReader();
                    while (rMN.Read())
                    {
                        string mNo = rMN["MPR_No"]?.ToString() ?? "";
                        if (!msgs.Exists(m => m.Contains(mNo)))
                        {
                            DateTime rawDtMN = rMN["Created_Date"] != DBNull.Value
                                ? Convert.ToDateTime(rMN["Created_Date"]) : DateTime.MinValue;
                            // Chuyển đổi sang múi giờ hệ thống địa phương
                            if (rawDtMN != DateTime.MinValue && rawDtMN.Kind == DateTimeKind.Unspecified)
                                rawDtMN = DateTime.SpecifyKind(rawDtMN, DateTimeKind.Utc).ToLocalTime();
                            string dt = rawDtMN != DateTime.MinValue ? rawDtMN.ToString("dd/MM HH:mm") : "";
                            string usr = rMN["Created_By"]?.ToString() ?? "";
                            string isoTs = rawDtMN.ToString("o");
                            msgs.Add("[MPR] MPR moi: " + mNo + " | " + rMN["Project_Name"] + " | " + dt + " - " + usr + "@@" + isoTs);
                        }
                    }
                    rMN.Close();

                    // ── RIR moi ──────────────────────────────────────────
                    string sqlRIR = "SELECT TOP " + (force ? "5" : "20") + " RIR_No, PONo, Project_Name, Issue_Date, Created_By " +
                                    "FROM RIR_head WHERE Issue_Date > @s ORDER BY Issue_Date DESC";
                    using var cRIR = new Microsoft.Data.SqlClient.SqlCommand(sqlRIR, conn);
                    cRIR.Parameters.AddWithValue("@s", since);
                    using var rRIR = cRIR.ExecuteReader();
                    while (rRIR.Read())
                    {
                        DateTime rawDtRIR = rRIR["Issue_Date"] != DBNull.Value
                            ? Convert.ToDateTime(rRIR["Issue_Date"]) : DateTime.MinValue;
                        // Chuyển đổi sang múi giờ hệ thống địa phương
                        if (rawDtRIR != DateTime.MinValue && rawDtRIR.Kind == DateTimeKind.Unspecified)
                            rawDtRIR = DateTime.SpecifyKind(rawDtRIR, DateTimeKind.Utc).ToLocalTime();
                        string dt = rawDtRIR != DateTime.MinValue ? rawDtRIR.ToString("dd/MM HH:mm") : "";
                        string usr = rRIR["Created_By"]?.ToString() ?? "";
                        string isoTs = rawDtRIR.ToString("o");
                        msgs.Add("[RIR] RIR moi: " + rRIR["RIR_No"] + " | PO: " + rRIR["PONo"] + " | " + dt + " - " + usr + "@@" + isoTs);
                    }
                    rRIR.Close();
                }

                // Cap nhat moc thoi gian SAU khi da lay het data
                DateTime checkTime = DateTime.Now;

                int newPO = msgs.FindAll(m => m.StartsWith("[PO]")).Count;
                int newMPR = msgs.FindAll(m => m.StartsWith("[MPR]")).Count;
                int newRIR = msgs.FindAll(m => m.StartsWith("[RIR]")).Count;

                if (newPO == 0 && newMPR == 0 && newRIR == 0 && !force) return;

                _lastCheckTime = checkTime; // cập nhật SAU khi query xong
                SaveLastCheckTime();        // persist để lần sau mở app check tiếp
                if (this.InvokeRequired)
                    this.Invoke(new Action(() => ApplyNotify(newPO, newMPR, newRIR, msgs, force)));
                else
                    ApplyNotify(newPO, newMPR, newRIR, msgs, force);
            }
            catch { }
        }

        private void ApplyNotify(int newPO, int newMPR, int newRIR,
            System.Collections.Generic.List<string> msgs, bool force)
        {
            // Lấy thời gian hệ thống hiện tại
            string t = DateTime.Now.ToString("HH:mm:ss");

            // Them cac thong bao chua tung hien vao _seenMsgs va listbox
            int added = 0;
            foreach (var m in msgs)
            {
                if (!_seenMsgs.Contains(m))
                {
                    _seenMsgs.Add(m);
                    added++;
                }
            }

            if (added > 0)
            {
                // Rebuild listbox theo thu tu thoi gian
                _lstNotify.BeginUpdate();
                _lstNotify.Items.Clear();
                var allMsgs = new System.Collections.Generic.List<string>(_seenMsgs);
                allMsgs.Sort((a, b) => ParseMsgTime(b).CompareTo(ParseMsgTime(a)));
                foreach (var m in allMsgs) _lstNotify.Items.Add(m);
                _lstNotify.EndUpdate();

                _unreadCount += added;
                UpdateBadge(); // tinh lai badge = so chua doc

                var parts = new System.Collections.Generic.List<string>();
                if (newPO > 0) parts.Add(newPO + " PO moi");
                if (newMPR > 0) parts.Add(newMPR + " MPR cap nhat");
                if (newRIR > 0) parts.Add(newRIR + " RIR moi");
                _lblNotifyCount.Text = string.Join(" | ", parts) + "  (" + t + ")";
                _lblNotifyCount.ForeColor = Color.FromArgb(255, 200, 200); // Do nhat

                // Luôn hiện và BringToFront khi có thông báo mới
                _panelNotify.Visible = true;
                PositionNotifyPanel();
                _panelNotify.BringToFront();
                
                // Phát âm thanh cảnh báo
                try
                {
                    System.Media.SystemSounds.Beep.Play();
                }
                catch { }
                
                // Flash badge để thu hút chú ý
                _btnNotify.BackColor = Color.FromArgb(220, 53, 69);
                var flashTimer = new System.Windows.Forms.Timer { Interval = 500 };
                int flashCount = 0;
                flashTimer.Tick += (ft, fe) =>
                {
                    flashCount++;
                    _btnNotify.BackColor = flashCount % 2 == 0
                        ? Color.FromArgb(220, 53, 69)
                        : Color.FromArgb(255, 140, 0);
                    if (flashCount >= 6) { ((System.Windows.Forms.Timer)ft).Stop(); }
                };
                flashTimer.Start();
            }
            else if (force)
            {
                _lblNotifyCount.Text = "Kiem tra luc " + t + " - Khong co moi";
            }

            var f = _panelNotify.Controls.Find("lblFoot", false);
            if (f.Length > 0 && f[0] is Label lf)
                lf.Text = "Tu dong cap nhat moi 1 phut | Tiep theo: " + DateTime.Now.AddMinutes(1).ToString("HH:mm:ss");
        }


        private void ShowNotifyDetail(string msg)
        {
            bool isPO = msg.StartsWith("[PO]");
            bool isMPR = msg.StartsWith("[MPR]");
            bool isRIR = msg.StartsWith("[RIR]");

            string[] parts = msg.Split('|');
            string title = parts.Length > 0 ? parts[0].Trim() : msg;
            string proj = parts.Length > 1 ? parts[1].Trim() : "";
            string detail = parts.Length > 2 ? parts[2].Trim() : "";

            // Tim ID tu title
            string id = "";
            if (isPO && title.Contains("PO moi: "))
                id = title.Replace("[PO] PO moi: ", "").Trim();
            else if (isMPR)
                id = title.Replace("[MPR] MPR cap nhat: ", "").Trim();
            else if (isRIR)
                id = title.Replace("[RIR] RIR moi: ", "").Replace("[RIR] RIR cap nhat: ", "").Trim();

            Color barColor = isPO ? Color.FromArgb(25, 118, 210) :
                             isMPR ? Color.FromArgb(46, 125, 50) :
                                     Color.FromArgb(102, 51, 153);
            string typeName = isPO ? "PO" : isMPR ? "MPR" : "RIR";

            // Tao popup chi tiet
            var popup = new Form
            {
                Text = "Chi tiet thong bao — " + typeName,
                Size = new Size(480, 380),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(250, 250, 252)
            };

            // Header mau
            var pTop = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(480, 56),
                BackColor = barColor
            };
            pTop.Paint += (s, e) =>
            {
                using var brush = new System.Drawing.Drawing2D.LinearGradientBrush(
                    pTop.ClientRectangle, barColor,
                    ControlPaint.Dark(barColor, 0.15f),
                    System.Drawing.Drawing2D.LinearGradientMode.Vertical);
                e.Graphics.FillRectangle(brush, pTop.ClientRectangle);
            };
            pTop.Controls.Add(new Label
            {
                Text = typeName + " — Chi tiet thong bao",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(16, 16),
                Size = new Size(440, 26),
                AutoEllipsis = true
            });
            popup.Controls.Add(pTop);

            // Query DB lay chi tiet
            var pInfo = new Panel
            {
                Location = new Point(12, 68),
                Size = new Size(450, 240),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
            popup.Controls.Add(pInfo);

            int ry = 12;
            Action<string, string, Color> addRow = (lbl, val, vc) =>
            {
                pInfo.Controls.Add(new Label
                {
                    Text = lbl,
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    ForeColor = Color.FromArgb(90, 90, 110),
                    Location = new Point(12, ry),
                    Size = new Size(110, 22)
                });
                pInfo.Controls.Add(new Label
                {
                    Text = val,
                    Font = new Font("Segoe UI", 9),
                    ForeColor = vc,
                    Location = new Point(128, ry),
                    Size = new Size(310, 22),
                    AutoEllipsis = true
                });
                ry += 28;
            };

            addRow("Loai:", typeName, barColor);
            addRow("Ma so:", id, Color.FromArgb(30, 30, 30));
            if (!string.IsNullOrEmpty(proj.Replace("PO:", "").Trim()))
                addRow("Du an / PO:", proj, Color.FromArgb(30, 30, 30));
            addRow("Thoi gian:", detail.Split('-')[0].Trim(), Color.FromArgb(100, 100, 120));
            if (detail.Contains("-"))
                addRow("Nguoi thuc hien:", detail.Split('-')[1].Trim(), barColor);

            // Lay them tu DB
            try
            {
                using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                conn.Open();
                if (isPO && !string.IsNullOrEmpty(id))
                {
                    var cmd = new Microsoft.Data.SqlClient.SqlCommand(
                        "SELECT Status, Notes, Supplier_ID, Total_Amount FROM PO_head WHERE PONo = @id", conn);
                    cmd.Parameters.AddWithValue("@id", id);
                    using var r = cmd.ExecuteReader();
                    if (r.Read())
                    {
                        addRow("Trang thai:", r["Status"]?.ToString() ?? "", Color.FromArgb(40, 120, 40));
                        addRow("Ghi chu:", r["Notes"]?.ToString() ?? "", Color.DimGray);
                        decimal amt = r["Total_Amount"] != DBNull.Value ? Convert.ToDecimal(r["Total_Amount"]) : 0;
                        if (amt > 0) addRow("Tong tien:", amt.ToString("N0") + " VND", Color.FromArgb(150, 40, 40));
                    }
                }
                else if (isMPR && !string.IsNullOrEmpty(id))
                {
                    var cmd = new Microsoft.Data.SqlClient.SqlCommand(
                        "SELECT Status, Notes, Required_Date FROM MPR_Header WHERE MPR_No = @id", conn);
                    cmd.Parameters.AddWithValue("@id", id);
                    using var r = cmd.ExecuteReader();
                    if (r.Read())
                    {
                        addRow("Trang thai:", r["Status"]?.ToString() ?? "", Color.FromArgb(40, 120, 40));
                        addRow("Ghi chu:", r["Notes"]?.ToString() ?? "", Color.DimGray);
                        string req = r["Required_Date"] != DBNull.Value
                            ? Convert.ToDateTime(r["Required_Date"]).ToString("dd/MM/yyyy") : "";
                        if (!string.IsNullOrEmpty(req)) addRow("Ngay can:", req, Color.FromArgb(150, 100, 0));
                    }
                }
                else if (isRIR && !string.IsNullOrEmpty(id))
                {
                    var cmd = new Microsoft.Data.SqlClient.SqlCommand(
                        "SELECT Status, PONo, Project_Name FROM RIR_head WHERE RIR_No = @id", conn);
                    cmd.Parameters.AddWithValue("@id", id);
                    using var r = cmd.ExecuteReader();
                    if (r.Read())
                    {
                        addRow("Trang thai:", r["Status"]?.ToString() ?? "", Color.FromArgb(40, 120, 40));
                        addRow("PO No:", r["PONo"]?.ToString() ?? "", Color.FromArgb(25, 118, 210));
                        addRow("Du an:", r["Project_Name"]?.ToString() ?? "", Color.DimGray);
                    }
                }
            }
            catch { }

            // Nut mo form tuong ung
            var btnOpen = new Button
            {
                Text = "Mo " + typeName,
                Size = new Size(120, 34),
                Location = new Point(12, 320),
                BackColor = barColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnOpen.FlatAppearance.BorderSize = 0;
            btnOpen.Click += (s, e) =>
            {
                popup.Close();
                if (isPO) { HighlightMenu("PO"); OpenForm(new frmPO(id)); }
                else if (isMPR) { HighlightMenu("MPR"); OpenForm(new frmMPR()); }
                else if (isRIR) { HighlightMenu("RIR"); OpenForm(new frmRIR()); }
            };
            popup.Controls.Add(btnOpen);

            var btnClose = new Button
            {
                Text = "Dong",
                Size = new Size(80, 34),
                Location = new Point(144, 320),
                BackColor = Color.FromArgb(84, 84, 84),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9),
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.Cancel
            };
            btnClose.FlatAppearance.BorderSize = 0;
            popup.Controls.Add(btnClose);
            popup.CancelButton = btnClose;
            popup.ShowDialog(this);
        }


        // Parse thoi gian tu ISO timestamp cuoi chuoi msg ("@@2024-01-15T14:30:00")
        private DateTime ParseMsgTime(string msg)
        {
            try
            {
                int idx = msg.LastIndexOf("@@");
                if (idx >= 0)
                {
                    string isoStr = msg.Substring(idx + 2);
                    if (DateTime.TryParse(isoStr, out DateTime dt))
                        return dt;
                }
            }
            catch { }
            return DateTime.MinValue;
        }

        // =====================================================================
        //  ZALO SETTINGS DIALOG
        // =====================================================================
        private void ShowZaloSettingsDialog()
        {
            var settings = ZaloHelper.LoadSettings();

            var dlg = new Form
            {
                Text = "Cài đặt thông báo Zalo",
                Size = new Size(560, 400),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false, MinimizeBox = false,
                BackColor = Color.White,
                Font = new Font("Segoe UI", 9)
            };

            int y = 16;

            // Bật/tắt
            var chkEnabled = new CheckBox
            {
                Text = "Bật gửi thông báo Zalo tự động sau khi gửi email",
                Location = new Point(16, y), Size = new Size(480, 24),
                Checked = settings.Enabled, Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            dlg.Controls.Add(chkEnabled); y += 38;

            // Thư mục session
            dlg.Controls.Add(new Label { Text = "Session path:", Location = new Point(16, y + 3), Size = new Size(110, 20) });
            var txtDataDir = new TextBox
            {
                Location = new Point(130, y), Size = new Size(330, 24),
                Text = string.IsNullOrWhiteSpace(settings.UserDataDir)
                    ? ZaloHelper.DefaultUserDataDir : settings.UserDataDir,
                Font = new Font("Segoe UI", 8)
            };
            dlg.Controls.Add(txtDataDir);
            var btnBrowse = new Button { Text = "...", Location = new Point(466, y - 1), Size = new Size(50, 26), FlatStyle = FlatStyle.Flat };
            btnBrowse.Click += (s, e) =>
            {
                using var fbd = new FolderBrowserDialog { Description = "Chọn thư mục lưu session Zalo", SelectedPath = txtDataDir.Text };
                if (fbd.ShowDialog() == DialogResult.OK) txtDataDir.Text = fbd.SelectedPath;
            };
            dlg.Controls.Add(btnBrowse); y += 36;

            // Hướng dẫn
            dlg.Controls.Add(new Label
            {
                Text = "── Thiết lập lần đầu ────────────────────────────────────────────",
                Location = new Point(16, y), Size = new Size(500, 18),
                ForeColor = Color.FromArgb(0, 100, 180), Font = new Font("Segoe UI", 8.5f)
            }); y += 22;

            dlg.Controls.Add(new Label
            {
                Text = "1. Bấm 'Lưu'  →  2. Bấm 'Mở & đăng nhập'  →  3. Đăng nhập Zalo Web  →  4. Đóng cửa sổ đó lại.",
                Location = new Point(16, y), Size = new Size(500, 18),
                Font = new Font("Segoe UI", 8.5f, FontStyle.Italic), ForeColor = Color.FromArgb(80, 80, 80)
            }); y += 28;

            var btnLogin = new Button
            {
                Text = "🌐 Mở & đăng nhập Zalo",
                Location = new Point(16, y), Size = new Size(200, 32),
                BackColor = Color.FromArgb(0, 140, 70), ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand
            };
            btnLogin.FlatAppearance.BorderSize = 0; y += 46;

            // Test
            dlg.Controls.Add(new Label { Text = "Tên nhóm test:", Location = new Point(16, y + 3), Size = new Size(110, 20) });
            var txtTestGroup = new TextBox
            {
                Location = new Point(130, y), Size = new Size(250, 24),
                PlaceholderText = "Nhập đúng tên nhóm Zalo"
            };
            dlg.Controls.Add(txtTestGroup);
            var btnTest = new Button
            {
                Text = "📨 Gửi thử",
                Location = new Point(386, y - 2), Size = new Size(100, 28),
                BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand
            };
            btnTest.FlatAppearance.BorderSize = 0;
            dlg.Controls.AddRange(new Control[] { btnLogin, btnTest }); y += 36;

            var lblResult = new Label
            {
                Location = new Point(16, y), Size = new Size(500, 22),
                ForeColor = Color.Gray, Font = new Font("Segoe UI", 8.5f)
            };
            dlg.Controls.Add(lblResult); y += 34;

            var btnSave  = new Button { Text = "💾 Lưu", Location = new Point(16, y),  Size = new Size(100, 32), BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            var btnClose = new Button { Text = "Đóng",   Location = new Point(430, y), Size = new Size(80, 32),  FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnSave.FlatAppearance.BorderSize = 0;
            dlg.Controls.AddRange(new Control[] { btnSave, btnClose });
            dlg.ClientSize = new Size(530, y + 52);

            ZaloSettings CurrentSettings() => new ZaloSettings
            {
                Enabled     = chkEnabled.Checked,
                UserDataDir = txtDataDir.Text.Trim()
            };

            btnSave.Click += (s, e) =>
            {
                var cfg = CurrentSettings();
                ZaloHelper.SaveSettings(cfg);
                lblResult.ForeColor = Color.Green;
                lblResult.Text = "✅ Đã lưu cài đặt.";
                if (!cfg.Enabled)
                {
                    ZaloSession.Shutdown();
                    _UpdateZaloButton(ZaloStatus.Disabled);
                }
            };

            btnLogin.Click += (s, e) =>
            {
                try
                {
                    ZaloHelper.SaveSettings(CurrentSettings());
                    ZaloHelper.OpenBrowserForLogin(CurrentSettings());
                    lblResult.ForeColor = Color.FromArgb(0, 100, 180);
                    lblResult.Text = "🌐 Đã mở cửa sổ Zalo — đăng nhập xong thì đóng lại.";
                }
                catch (Exception ex)
                {
                    lblResult.ForeColor = Color.OrangeRed;
                    lblResult.Text = $"❌ {ex.Message}";
                }
            };

            btnTest.Click += async (s, e) =>
            {
                string grp = txtTestGroup.Text.Trim();
                if (string.IsNullOrWhiteSpace(grp)) { lblResult.Text = "⚠ Nhập tên nhóm để gửi thử."; lblResult.ForeColor = Color.OrangeRed; return; }
                btnTest.Enabled = false;
                lblResult.ForeColor = Color.Gray;
                lblResult.Text = "⏳ Đang mở Zalo Web và tìm nhóm...";

                var (ok, err) = await ZaloHelper.SendToGroupAsync(CurrentSettings(), grp,
                    "🔔 Test thông báo từ hệ thống ERP — Zalo tự động đã hoạt động!");
                btnTest.Enabled = true;
                if (ok) { lblResult.Text = "✅ Gửi thành công!"; lblResult.ForeColor = Color.Green; }
                else    { lblResult.Text = $"❌ {err}";           lblResult.ForeColor = Color.OrangeRed; }
            };

            btnClose.Click += (s, e) => dlg.Close();
            dlg.ShowDialog(this);

            // Sau khi dialog đóng: restart session nếu đã bật nhưng session bị tắt (vd sau đăng nhập lại)
            var saved = ZaloHelper.LoadSettings();
            if (saved.Enabled && !ZaloSession.IsReady)
                _ = _WarmUpZaloSessionAsync();
            else if (!saved.Enabled)
                _UpdateZaloButton(ZaloStatus.Disabled);
        }

        // ─────────────────────────────────────────────────────────────────────
        //  HELPERS — Thông báo nội bộ
        // ─────────────────────────────────────────────────────────────────────
        private static int ExtractNotifId(string msg)
        {
            // Định dạng: "...@@{isoTs}|notif_id=123"
            int idx = msg.IndexOf("|notif_id=", StringComparison.Ordinal);
            if (idx < 0) return 0;
            string tail = msg.Substring(idx + "|notif_id=".Length);
            return int.TryParse(tail, out int id) ? id : 0;
        }

        private void ShowInternalNotifyDetail(string msg)
        {
            // Parse: [NOTIFY] {title} | Từ: {sender} | {dt}@@{iso}|notif_id={id}
            string clean = msg.Contains("@@") ? msg.Substring(0, msg.IndexOf("@@")) : msg;
            string[] parts = clean.Split('|');
            string titleLine = parts.Length > 0 ? parts[0].Replace("[NOTIFY]", "").Trim() : "";
            string senderLine = parts.Length > 1 ? parts[1].Trim() : "";
            string timeLine = parts.Length > 2 ? parts[2].Trim() : "";

            // Tìm nội dung đầy đủ từ DB (nếu cần)
            int notifId = ExtractNotifId(msg);

            var popup = new Form
            {
                Text = "📢 Thông báo nội bộ",
                Size = new Size(500, 380),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(250, 250, 252)
            };

            var pTop = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(500, 56),
                BackColor = Color.FromArgb(198, 40, 40)
            };
            pTop.Controls.Add(new Label
            {
                Text = "📢  Thông báo nội bộ",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(16, 16),
                Size = new Size(460, 26)
            });
            popup.Controls.Add(pTop);

            int y = 72;
            void AddRow(string label, string value, bool multiline = false)
            {
                popup.Controls.Add(new Label
                {
                    Text = label,
                    Font = new Font("Segoe UI", 8, FontStyle.Bold),
                    ForeColor = Color.Gray,
                    Location = new Point(16, y),
                    Size = new Size(70, 18)
                });
                var txt = new TextBox
                {
                    Text = value,
                    Font = new Font("Segoe UI", 9),
                    ReadOnly = true,
                    BorderStyle = BorderStyle.None,
                    BackColor = Color.FromArgb(250, 250, 252),
                    Location = new Point(90, y),
                    Size = new Size(390, multiline ? 100 : 20),
                    Multiline = multiline,
                    WordWrap = multiline,
                    ScrollBars = multiline ? ScrollBars.Vertical : ScrollBars.None
                };
                popup.Controls.Add(txt);
                y += multiline ? 116 : 28;
            }

            AddRow("Tiêu đề:", titleLine);
            AddRow("Người gửi:", senderLine);
            AddRow("Thời gian:", timeLine);

            // Tải nội dung đầy đủ từ DB
            if (notifId > 0)
            {
                try
                {
                    using var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection();
                    conn.Open();
                    using var cmd = new SqlCommand(
                        "SELECT Content FROM Internal_Notifications WHERE Notif_ID=@id", conn);
                    cmd.Parameters.AddWithValue("@id", notifId);
                    string fullContent = cmd.ExecuteScalar()?.ToString() ?? "";
                    AddRow("Nội dung:", fullContent, multiline: true);
                }
                catch
                {
                    AddRow("Nội dung:", "(Không thể tải nội dung)");
                }
            }

            var btnClose = new Button
            {
                Text = "✔ Đóng",
                Location = new Point(390, popup.ClientSize.Height - 46),
                Size = new Size(90, 32),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                DialogResult = DialogResult.OK
            };
            btnClose.FlatAppearance.BorderSize = 0;
            popup.Controls.Add(btnClose);
            popup.AcceptButton = btnClose;

            popup.ShowDialog(this);
        }

    }
}