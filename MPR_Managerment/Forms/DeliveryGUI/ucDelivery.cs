using Microsoft.Web.WebView2.Core;
using MPR_Managerment.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MPR_Managerment.Forms.DeliveryGUI
{
    public partial class ucDelivery : UserControl
    {
        private List<ProjectInfo> _dtProject = new List<ProjectInfo>();
        public ucDelivery(List<ProjectInfo> projectInfos)
        {
            InitializeComponent();
            this._dtProject = projectInfos;
            SetupSaveDeliveryNotetLayout();
        }

        public void Cleanup(Microsoft.Web.WebView2.WinForms.WebView2 webView)
        {
            if (webView != null && webView.CoreWebView2 != null)
            {
                // Chuyển về trang trống để giải phóng trang hiện tại
                webView.CoreWebView2.Navigate("about:blank");
                webView.Dispose();
            }
        }

        public void SetupSaveDeliveryNotetLayout()
        {
            // Hàng 1: Dự án
            pHead.Controls.Add(new Label { Text = "Dự án:", Location = new Point(8, 10), Size = new Size(50, 20), Font = new Font("Segoe UI", 9, FontStyle.Bold) });
            var cboDeliveryProject = new ComboBox
            {
                Location = new Point(60, 7),
                Size = new Size(160, 26),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            pHead.Controls.Add(cboDeliveryProject);

            var txtPONo = new TextBox
            {
                Location = new Point(284, 7),
                Size = new Size(200, 26),
                Font = new Font("Segoe UI", 9),
                //PlaceholderText = ""
            };
            pHead.Controls.Add(txtPONo);

            // Hàng 1: nút Lưu hóa đơn (góc phải)
            var btnSaveDelivery = new Button
            {
                Text = "💾 Lưu phiếu giao hàng",
                Size = new Size(145, 28),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Cursor = Cursors.Hand
            };
            btnSaveDelivery.Location = new Point(pHead.Width - 155, 6);
            btnSaveDelivery.FlatAppearance.BorderSize = 0;
            pHead.Controls.Add(btnSaveDelivery);

            // Hàng 2: Delivery Link path
            var lblDelivertPath = new Label
            {
                Location = new Point(8, 40),
                Size = new Size(pHead.Width - 16, 18),
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.FromArgb(100, 100, 100),
                Text = "Delivery Note Link: —",
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            pHead.Controls.Add(lblDelivertPath);

            // Resize
            pHead.Resize += (s, e) =>
            {
                btnSaveDelivery.Location = new Point(pHead.Width - 155, 6);
                lblDelivertPath.Width = pHead.Width - 16;
            };

            // ── State — khai báo sớm để các lambda bên dưới dùng được ──
            string _invFolderPath = "";
            string _pendingDropPath = "";
            string _pendingDropName = "";

            // Alias để code bên dưới dùng splitMain.Panel1 / Panel2 vẫn đúng

            // ── Panel trái: INV list ──

            pTutorial.Controls.Add(new Label
            {
                Text = "📄  Delivery note List  —  kéo thả file PDF vào đây",
                Dock = DockStyle.Fill,
                //Height = 26,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(4, 0, 0, 0)
            });

            var dgvDelivery = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowDrop = true
            };
            dgvDelivery.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvDelivery.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvDelivery.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvDelivery.EnableHeadersVisualStyles = false;
            dgvDelivery.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvDelivery.Columns.Add(new DataGridViewTextBoxColumn { Name = "FileName", HeaderText = "Tên file", FillWeight = 70 });
            dgvDelivery.Columns.Add(new DataGridViewTextBoxColumn { Name = "FileSize", HeaderText = "Kích thước", FillWeight = 20 });
            dgvDelivery.Columns.Add(new DataGridViewTextBoxColumn { Name = "FullPath", HeaderText = "FullPath", Visible = false });
            dgvDelivery.Columns.Add(new DataGridViewTextBoxColumn { Name = "IsPending", HeaderText = "Trạng thái", FillWeight = 10 });
            pGrid.Controls.Add(dgvDelivery);

            // WIN 11
            //var webView = new System.Windows.Forms.WebBrowser
            //{
            //    Dock = DockStyle.Fill,
            //    ScrollBarsEnabled = true,
            //    IsWebBrowserContextMenuEnabled = false
            //};
            //pDeliveryRight.Controls.Add(webView);

            // WIN 10
            var webView = new Microsoft.Web.WebView2.WinForms.WebView2
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White
            };
            pDeliveryRight.Controls.Add(webView);

            // Khởi tạo WebView2 (Bắt buộc vì WebView2 cần khởi tạo môi trường)
            async void InitializeWebView()
            {
                //try
                //{
                //    // Đặt màu nền của chính Control WebView2 thành trắng ngay từ đầu
                //    webView.BackColor = Color.White;

                //    // Khởi tạo môi trường WebView2
                //    await webView.EnsureCoreWebView2Async();

                //    // THIẾT LẬP QUAN TRỌNG: Chỉnh màu nền mặc định của nhân Chromium
                //    // System.Drawing.Color.White sẽ giúp loại bỏ vùng xám đen khi chưa load file
                //    webView.CoreWebView2.Environment.UserDataFolder.ToString(); // (Tùy chọn check path)
                //    webView.DefaultBackgroundColor = Color.White;

                //    // Tắt các tính năng không cần thiết để giao diện sạch hơn
                //    webView.CoreWebView2.Settings.IsZoomControlEnabled = true;

                //    // Điều hướng đến trang trắng để đảm bảo không có màu lạ
                //    webView.CoreWebView2.Navigate("about:blank");
                //}
                //catch (Exception ex)
                //{
                //    MessageBox.Show("Lỗi khởi tạo WebView2: " + ex.Message);
                //}
                try
                {
                    if (webView != null && webView.CoreWebView2 == null)
                    {
                        // Thiết lập thư mục UserData riêng biệt trong LocalAppData để tránh tranh chấp file
                        string userDataFolder = Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                            "MPR_Managerment",
                            "WebView2_Delivery_Cache"
                        );

                        // Tạo môi trường khởi tạo
                        var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);

                        // Khởi tạo CoreWebView2
                        await webView.EnsureCoreWebView2Async(env);

                        // Cấu hình thêm sau khi đã khởi tạo xong
                        webView.CoreWebView2.Settings.IsPasswordAutosaveEnabled = false;
                        webView.CoreWebView2.Settings.IsGeneralAutofillEnabled = false;
                    }
                }
                catch (Exception ex)
                {
                    // Ghi log hoặc hiển thị lỗi nếu cần
                    Debug.WriteLine("WebView2 Init Error: " + ex.Message);
                }
            }
            InitializeWebView();

            // ── Load DELIVERY list từ thư mục ──
            Action loadDeliveryList = () =>
            {
                dgvDelivery.Rows.Clear();
                _pendingDropPath = "";
                _pendingDropName = "";

                if (string.IsNullOrEmpty(_invFolderPath) || !System.IO.Directory.Exists(_invFolderPath))
                {
                    lblDelivertPath.Text = "Delivery note Link: (thư mục không tồn tại)";
                    lblDelivertPath.ForeColor = Color.FromArgb(200, 53, 69);
                    return;
                }
                lblDelivertPath.Text = $"Delivery note Link: {_invFolderPath}";
                lblDelivertPath.ForeColor = Color.FromArgb(100, 100, 100);

                foreach (var f in System.IO.Directory.GetFiles(_invFolderPath, "*.pdf")
                                       .OrderBy(x => x))
                {
                    var fi = new System.IO.FileInfo(f);
                    int idx = dgvDelivery.Rows.Add();
                    dgvDelivery.Rows[idx].Cells["FileName"].Value = fi.Name;
                    dgvDelivery.Rows[idx].Cells["FileSize"].Value = $"{fi.Length / 1024.0:0.#} KB";
                    dgvDelivery.Rows[idx].Cells["FullPath"].Value = f;
                    dgvDelivery.Rows[idx].Cells["IsPending"].Value = "";
                }
            };

            // ── Chọn dự án → load INV path ──
            try
            {
                foreach (var p in _dtProject)
                    cboDeliveryProject.Items.Add(p.ProjectCode);
                if (cboDeliveryProject.Items.Count > 0) cboDeliveryProject.SelectedIndex = 0;
            }
            catch { }

            cboDeliveryProject.SelectedIndexChanged += (s, e) =>
            {
                _invFolderPath = "";
                _pendingDropPath = "";
                //webView.Navigate("about:blank");
                if (webView != null && webView.CoreWebView2 != null)
                    webView.CoreWebView2.Navigate("about:blank");
                try
                {
                    string code = cboDeliveryProject.SelectedItem?.ToString() ?? "";
                    var proj = _dtProject.Find(p => p.ProjectCode == code);
                    _invFolderPath = proj?.DeliveryNote_Link?.Trim() ?? "";
                }
                catch { }
                //loadPOForProject();
                loadDeliveryList();
            };

            // ── Chọn dòng → preview PDF ──
            dgvDelivery.SelectionChanged += (s, e) =>
            {
                if (dgvDelivery.SelectedRows.Count == 0) return;
                string path = dgvDelivery.SelectedRows[0].Cells["FullPath"].Value?.ToString() ?? "";

                // WIN 11
                //if (!string.IsNullOrEmpty(path) && System.IO.File.Exists(path))
                //    webView.Navigate(path);
                //else
                //    webView.Navigate("about:blank");

                // WIN 10
                if (!string.IsNullOrEmpty(path) && System.IO.File.Exists(path))
                {
                    if (webView != null && webView.CoreWebView2 != null)
                        webView.CoreWebView2.Navigate(path);
                }
                else
                {
                    if (webView != null && webView.CoreWebView2 != null)
                        webView.CoreWebView2.Navigate("about:blank");
                }
            };

            // ── Double click → mở file PDF ──
            dgvDelivery.CellDoubleClick += (s, ev) =>
            {
                if (ev.RowIndex < 0) return;
                string path = dgvDelivery.Rows[ev.RowIndex].Cells["FullPath"].Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(path) && System.IO.File.Exists(path))
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    { FileName = path, UseShellExecute = true });
            };

            // ── Kéo thả file PDF — hỗ trợ cả File Explorer và Outlook ──

            // Hàm xử lý drop dùng chung cho dgvInv và pInvLeft
            Action<DragEventArgs> handleDrop = (e) =>
            {
                string pdfPath = null;
                string pdfName = null;
                byte[] pdfBytes = null;

                // Cách 1: Kéo từ File Explorer (DataFormats.FileDrop)
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                    var pdf = System.Array.Find(files, f =>
                        f.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase));
                    if (pdf != null) pdfPath = pdf;
                }

                // Cách 2: Kéo từ Outlook (FileGroupDescriptor + FileContents)
                if (pdfPath == null &&
                    e.Data.GetDataPresent("FileGroupDescriptorW") &&
                    e.Data.GetDataPresent("FileContents"))
                {
                    try
                    {
                        // Lấy tên file từ FileGroupDescriptorW
                        var fgd = e.Data.GetData("FileGroupDescriptorW") as System.IO.MemoryStream;
                        if (fgd != null)
                        {
                            fgd.Position = 0;
                            byte[] buf = fgd.ToArray();
                            // Tên file bắt đầu từ byte 76, encoding Unicode
                            string fname = System.Text.Encoding.Unicode
                                .GetString(buf, 76, buf.Length - 76)
                                .TrimEnd('\0').Trim();
                            if (fname.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                            {
                                pdfName = fname;
                                // Lấy nội dung file từ FileContents
                                var fc = e.Data.GetData("FileContents", true) as System.IO.MemoryStream;
                                if (fc != null)
                                {
                                    fc.Position = 0;
                                    pdfBytes = fc.ToArray();
                                }
                            }
                        }
                    }
                    catch { }
                }

                // Nếu không có gì hợp lệ
                if (pdfPath == null && pdfBytes == null)
                {
                    MessageBox.Show("Chỉ hỗ trợ file PDF!", "Thông báo",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
                }

                // Nếu từ Outlook → lưu tạm vào temp folder
                if (pdfPath == null && pdfBytes != null)
                {
                    string tmpDir = System.IO.Path.GetTempPath();
                    string tmpFile = System.IO.Path.Combine(tmpDir, pdfName ?? "invoice_temp.pdf");
                    System.IO.File.WriteAllBytes(tmpFile, pdfBytes);
                    pdfPath = tmpFile;
                }

                // Xóa dòng pending cũ
                for (int r = dgvDelivery.Rows.Count - 1; r >= 0; r--)
                    if (dgvDelivery.Rows[r].Cells["IsPending"].Value?.ToString() == "⏳ Chờ lưu")
                        dgvDelivery.Rows.RemoveAt(r);

                _pendingDropPath = pdfPath;
                _pendingDropName = System.IO.Path.GetFileName(pdfPath);
                var fi = new System.IO.FileInfo(pdfPath);

                int idx = dgvDelivery.Rows.Add();
                dgvDelivery.Rows[idx].Cells["FileName"].Value = _pendingDropName;
                dgvDelivery.Rows[idx].Cells["FileSize"].Value = $"{fi.Length / 1024.0:0.#} KB";
                dgvDelivery.Rows[idx].Cells["FullPath"].Value = pdfPath;
                dgvDelivery.Rows[idx].Cells["IsPending"].Value = "⏳ Chờ lưu";
                dgvDelivery.Rows[idx].DefaultCellStyle.ForeColor = Color.FromArgb(255, 140, 0);
                dgvDelivery.Rows[idx].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                dgvDelivery.ClearSelection();
                dgvDelivery.Rows[idx].Selected = true;

                // WIN 11
                //webView.Navigate(pdfPath);

                // WIN 10
                if (webView != null && webView.CoreWebView2 != null)
                    webView.CoreWebView2.Navigate(pdfPath);
            };

            Action<DragEventArgs> handleDragEnter = (e) =>
            {
                bool hasFileDrop = e.Data.GetDataPresent(DataFormats.FileDrop);
                bool hasOutlook = e.Data.GetDataPresent("FileGroupDescriptorW");
                e.Effect = (hasFileDrop || hasOutlook)
                    ? DragDropEffects.Copy
                    : DragDropEffects.None;
            };

            // Đăng ký trên dgvInv
            dgvDelivery.AllowDrop = true;
            dgvDelivery.DragEnter += (s, e) => handleDragEnter(e);
            dgvDelivery.DragOver += (s, e) => handleDragEnter(e);
            dgvDelivery.DragDrop += (s, e) => handleDrop(e);

            // Đăng ký trên pInvLeft (panel chứa) — bắt khi drop vào vùng trống
            pDeliveryLeft.AllowDrop = true;
            pDeliveryLeft.DragEnter += (s, e) => handleDragEnter(e);
            pDeliveryLeft.DragOver += (s, e) => handleDragEnter(e);
            pDeliveryLeft.DragDrop += (s, e) => handleDrop(e);

            // ── Lưu hóa đơn ──
            btnSaveDelivery.Click += (s, e) =>
            {
                if (string.IsNullOrEmpty(_pendingDropPath) || !System.IO.File.Exists(_pendingDropPath))
                { MessageBox.Show("Không có file nào chờ lưu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

                if (!System.IO.Directory.Exists(_invFolderPath))
                { MessageBox.Show("Thư mục INV Link không tồn tại!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                // Lấy PO No từ bộ lọc trong tab hóa đơn
                string poNo = txtPONo.Text;
                if (string.IsNullOrEmpty(poNo) || poNo == "-- Chọn PO No --")
                { MessageBox.Show("Vui lòng chọn số PO trước khi lưu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

                // Tạo tên file: INV_PONo.pdf, nếu trùng → INV_PONo_2.pdf, _3.pdf...
                string baseName = $"Delivery_{poNo}";
                string destPath = System.IO.Path.Combine(_invFolderPath, baseName + ".pdf");
                int counter = 2;
                while (System.IO.File.Exists(destPath))
                {
                    destPath = System.IO.Path.Combine(_invFolderPath,
                        $"{baseName}_tờ{counter}.pdf");
                    counter++;
                }

                try
                {
                    System.IO.File.Copy(_pendingDropPath, destPath, false);
                    MessageBox.Show(
                        $"✅ Đã lưu hóa đơn thành công!\nFile: {System.IO.Path.GetFileName(destPath)}",
                        "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    _pendingDropPath = "";
                    _pendingDropName = "";
                    //btnSaveDelivery.Enabled = false;
                    loadDeliveryList();

                    // WIN 11
                    // Preview file vừa lưu
                    //webView.Navigate(destPath);

                    // WIN 10
                    if (webView != null && webView.CoreWebView2 != null)
                        webView.CoreWebView2.Navigate(destPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi lưu file: " + ex.Message, "Lỗi",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            // Load lần đầu
            if (cboDeliveryProject.Items.Count > 0)
            {
                cboDeliveryProject.SelectedIndex = 0;
                //loadPOForProject();
            }
        }
    }
}
