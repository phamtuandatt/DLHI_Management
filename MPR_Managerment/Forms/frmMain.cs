using System;
using System.Drawing;
using System.Windows.Forms;
using MPR_Managerment.Models;
using MPR_Managerment.Services;

namespace MPR_Managerment.Forms
{
    public partial class frmMain : Form
    {
        private Panel panelMenu;
        private Panel panelContent;
        private Panel panelHeader;
        private Label lblUser;
        private Form _activeForm = null;

        public frmMain()
        {
            InitializeComponent();
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
            

            BuildHeader();
            BuildMenu();
            BuildContent();
            this.Resize += (s, ev) => UpdateContentBounds();
            UpdateContentBounds();
            ShowDashboard();
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

            panelHeader.Resize += (s, ev) =>
            {
                panelHeader.Width = this.ClientSize.Width;
                lblUser.Left = panelHeader.Width - 300;
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
                Size = new Size(220, this.Height - 55),
                BackColor = Color.FromArgb(30, 30, 45),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom
            };

            panelMenu.Controls.Add(new Label
            {
                Text = "MENU CHÍNH",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(150, 150, 180),
                Location = new Point(0, 15),
                Size = new Size(220, 25),
                TextAlign = ContentAlignment.MiddleCenter
            });

            int y = 55;

            // Tổng quan — luôn hiện
            AddMenuBtn("🏠  Tổng quan", Color.FromArgb(0, 120, 212), y); y += 52;
            AddMenuBtn("📊  Dashboard", Color.FromArgb(30, 30, 45), y); y += 52;

            // Các module — kiểm tra quyền
            if (AppSession.CanView("PROJECT"))
            {
                AddMenuBtn("🗂  Dự án", Color.FromArgb(30, 30, 45), y); y += 52;
            }
            if (AppSession.CanView("SUPPLIER"))
            {
                AddMenuBtn("🏢  Nhà cung cấp", Color.FromArgb(30, 30, 45), y); y += 52;
            }
            if (AppSession.CanView("MPR"))
            {
                AddMenuBtn("📋  MPR", Color.FromArgb(30, 30, 45), y); y += 52;
            }
            if (AppSession.CanView("PO"))
            {
                AddMenuBtn("🛒  Đơn đặt hàng (PO)", Color.FromArgb(30, 30, 45), y); y += 52;
            }
            if (AppSession.CanView("RIR"))
            {
                AddMenuBtn("📦  Kiểm tra (RIR)", Color.FromArgb(30, 30, 45), y); y += 52;
            }
            if (AppSession.CanView("WAREHOUSE"))
            {
                AddMenuBtn("🏭  Kho vật tư", Color.FromArgb(30, 30, 45), y); y += 52;

            AddMenuBtn("💳  Thanh toán Debit", Color.FromArgb(30, 30, 45), y); y += 52;


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
                AddMenuBtn("👤  Quản lý User", Color.FromArgb(63, 81, 181), y); y += 52;
            }

            // Đổi mật khẩu — luôn hiện
            AddMenuBtn("🔑  Đổi mật khẩu", Color.FromArgb(30, 30, 45), y); y += 52;
            AddMenuBtn("❌  Thoát", Color.FromArgb(30, 30, 45), y);

            this.Controls.Add(panelMenu);
        }

        private void AddMenuBtn(string text, Color backColor, int y)
        {
            var btn = new Button
            {
                Text = text,
                Location = new Point(0, y),
                Size = new Size(220, 48),
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

            if (tag.Contains("Tổng quan")) ShowDashboard();
            else if (tag.Contains("Dashboard")) OpenForm(new frmDashboard());
            else if (tag.Contains("Dự án")) OpenForm(new frmProject());
            else if (tag.Contains("Nhà cung cấp")) OpenForm(new frmSupplier());
            else if (tag.Contains("MPR")) OpenForm(new frmMPR());
            else if (tag.Contains("PO")) OpenForm(new frmPO());
            else if (tag.Contains("RIR")) OpenForm(new frmRIR());
            else if (tag.Contains("Thanh toán Debit")) OpenForm(new frmPayment());
            //else if (tag.Contains("Kho vật tư")) OpenForm(new frmWarehouse());
            else if (tag.Contains("Kho vật tư")) OpenForm(new frmWarehouses_v2());
            else if (tag.Contains("Quản lý User")) OpenForm(new frmUserManagement());
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
                if (MessageBox.Show("Bạn có chắc muốn thoát?", "Xác nhận",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
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
            int menuW = 220;
            panelHeader.Width = this.ClientSize.Width;
            panelMenu.Height = this.ClientSize.Height - headerH;
            panelContent.Location = new Point(menuW, headerH);
            panelContent.Size = new Size(this.ClientSize.Width - menuW, this.ClientSize.Height - headerH);
            panelMenu.BringToFront();
            panelHeader.BringToFront();
        }

        private void OpenForm(Form form)
        {
            if (_activeForm != null)
            {
                _activeForm.Close();
                _activeForm = null;
            }
            _activeForm = form;
            _activeForm.TopLevel = false;
            _activeForm.FormBorderStyle = FormBorderStyle.None;
            _activeForm.Dock = DockStyle.Fill;
            panelContent.Controls.Clear();
            panelContent.Controls.Add(_activeForm);
            _activeForm.Show();
        }

        // =====================================================================
        //  DASHBOARD
        // =====================================================================
        private void ShowDashboard()
        {
            if (_activeForm != null)
            {
                _activeForm.Close();
                _activeForm = null;
            }
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
                Text = "MPR Management System  v1.0  —  C# Windows Forms + SQL Server Azure",
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.FromArgb(180, 180, 200),
                Location = new Point(15, 15),
                Size = new Size(920, 22)
            });
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
                { HighlightMenu("Kho vật tư"); OpenForm(new frmWarehouse()); }
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
    }
}