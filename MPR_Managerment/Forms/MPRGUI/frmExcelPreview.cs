using Microsoft.Web.WebView2.WinForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MPR_Managerment.Forms.MPRGUI
{
    public partial class frmExcelPreview : Form
    {
        private string _filePath;
        // Khai báo WebBrowser bằng code để tránh lỗi "does not exist"
        private WebView2 webBrowser1;
        private Panel panToolbar;
        private Button btnClose;

        public frmExcelPreview(string filePath, string title = "Xem trước tài liệu")
        {
            InitializeComponentInternal(); // Tự khởi tạo giao diện bằng code
            this._filePath = filePath;
            this.Text = title;

            this.StartPosition = FormStartPosition.CenterScreen;
            this.WindowState = FormWindowState.Maximized;
            this.ShowInTaskbar = false;
        }

        /// <summary>
        /// Hàm này thay thế cho file Designer.cs để khởi tạo các thành phần giao diện
        /// </summary>
        private void InitializeComponentInternal()
        {
            this.webBrowser1 = new WebView2();
            this.panToolbar = new Panel();
            this.btnClose = new Button();

            // 1. Cấu hình Thanh công cụ (Toolbar)
            this.panToolbar.Dock = DockStyle.Top;
            this.panToolbar.Height = 50;
            this.panToolbar.BackColor = Color.FromArgb(240, 240, 240);

            // 2. Cấu hình Nút đóng
            this.btnClose.Text = "❌ Đóng (Esc)";
            this.btnClose.Size = new Size(120, 35);
            this.btnClose.Location = new Point(10, 7);
            this.btnClose.Cursor = Cursors.Hand;
            this.btnClose.FlatStyle = FlatStyle.Flat;
            this.btnClose.Click += (s, e) => this.Close();

            // 3. Cấu hình WebBrowser (Vùng hiển thị Excel)
            this.webBrowser1.Dock = DockStyle.Fill;
            // Ngăn chặn mở cửa sổ mới khi click vào link trong excel
            this.webBrowser1.AllowExternalDrop = false;

            // Thêm các control vào Form
            this.panToolbar.Controls.Add(this.btnClose);
            this.Controls.Add(this.webBrowser1);
            this.Controls.Add(this.panToolbar); // Add toolbar sau để nó nằm trên cùng

            this.Load += new EventHandler(frmExcelPreview_Load);
        }

        private void frmExcelPreview_Load(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists(_filePath))
                {
                    // Lấy đường dẫn tuyệt đối
                    string absolutePath = Path.GetFullPath(_filePath);

                    // Chuyển đổi sang URI chuẩn (Fix lỗi đường dẫn có khoảng trắng/tiếng Việt)
                    Uri fileUri = new Uri(absolutePath, UriKind.Absolute);

                    // Ví dụ copy ra thư mục tạm để xem
                    string tempFile = Path.Combine(Path.GetTempPath(), Path.GetFileName(_filePath));
                    File.Copy(_filePath, tempFile, true);
                    webBrowser1.Source = new Uri(_filePath);

                    // Navigate bằng URI
                    //webBrowser1.Navigate(fileUri);
                }
                else
                {
                    MessageBox.Show("Không tìm thấy file template!");
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi hiển thị: " + ex.Message);
            }
        }

        // Hỗ trợ phím tắt ESC để đóng nhanh
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                this.Close();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (webBrowser1 != null)
            {
                webBrowser1.Dispose();
            }
            base.OnFormClosing(e);
        }
    }
}
