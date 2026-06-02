using MPR_Managerment.Common;
using MPR_Managerment.Forms.ImportWarehouseGUI;
using MPR_Managerment.Forms.ItemCodeGUI;
using MPR_Managerment.Models;
using MPR_Managerment.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace MPR_Managerment.Forms.WarehouseGUI
{
    public partial class ucWarehouse : UserControl
    {
        private DataGridView dgvStock;
        private DataGridView dgvHisTranfer;
        private TextBox txtSearchStock;
        private ComboBox cboProjectFilter;
        private Label lblStockTotal, lblStockQty, lblStockWeight;
        private Panel panelStockSummary;

        private string _targetPONo = "";

        private List<POHead> _poList = new List<POHead>();
        private Label lblStatus;
        private List<object> _makeRequestExport = new List<object>();
        private List<SelectedItemModel> _selectedItem = new List<SelectedItemModel>();

        private List<string> originalPOList = new List<string>();

        private List<ProjectInfo> _dtProject = new List<ProjectInfo>();

        private Button btnSearch, btnCancelSearch, btnSearchHistory;
        private WarehouseService _service = new WarehouseService();
        private ComboBox cboProject;

        public ucWarehouse()
        {
            InitializeComponent();
            //frmAIChat.Attach(this); // Bỏ dòng này vì đây là UserControl, không phải Form chính

            BuidUI();
        }


        private void ucWarehouse_Load(object sender, EventArgs e)
        {
            _dtProject = new ProjectService().GetAll();
            LoadAll();
            //// Tự động nhảy đến PO được chọn từ Dashboard
            //if (!string.IsNullOrEmpty(_targetPONo))
            //{
            //    cboProject.SelectedIndex = 0;
            //    //LoadPOFilterByProject("");

            //    for (int i = 0; i < cboPONo.Items.Count; i++)
            //    {
            //        if (cboPONo.Items[i].ToString() == _targetPONo)
            //        {
            //            cboPONo.SelectedIndex = i;
            //            break;
            //        }
            //    }
            //    mainTabControl.SelectedTab = pageImport;
            //}
        }
        private void LoadAll()
        {
            LoadProjectImportFilter();
            //LoadProjectFilter();
            PopulateTreeView(_service.GetTranformHistory());
        }

        private void LoadProjectImportFilter()
        {
            try
            {
                cboProjectFilter.Items.Clear();
                cboProjectFilter.Items.Add("Tất cả dự án");
                foreach (var p in _dtProject)
                    cboProjectFilter.Items.Add(p.ProjectCode);
                cboProjectFilter.SelectedIndex = 0;
            }
            catch { }
        }


        private void BuidUI()
        {
            GroupBox gbHisTransfer = new GroupBox();
            gbHisTransfer.Text = "Vật tư chuyển - mượn";
            //gbHisTransferer.Size = new Size(1280, 700);
            gbHisTransfer.Dock = DockStyle.Fill;
            //gbHisTransfer.Height = 700;
            gbHisTransfer.Location = new Point(10, 115);
            pnHisTransfer.Controls.Add(gbHisTransfer);

            dgvHisTranfer = new DataGridView
            {
                //Size = new Size(1200, 1200),
                ReadOnly = false, // Phải để false để có thể click vào CheckBox
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                Margin = new Padding(0, 100, 0, 0),
            };

            dgvHisTranfer.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvHisTranfer.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvHisTranfer.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvHisTranfer.EnableHeadersVisualStyles = false;
            dgvHisTranfer.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvHisTranfer.Dock = DockStyle.Fill;
            dgvHisTranfer.DefaultCellStyle.SelectionBackColor = Color.FromArgb(204, 232, 255);
            dgvHisTranfer.DefaultCellStyle.SelectionForeColor = Color.Black;


            GroupBox gbHeader = new GroupBox();
            gbHeader.Text = "Lịch sử nhập hàng";
            //gbHeader.Size = new Size(1280, 700);
            gbHeader.Dock = DockStyle.Top;
            gbHeader.Height = 700;
            gbHeader.Location = new Point(10, 115);
            pnContent.Controls.Add(gbHeader);

            dgvStock = new DataGridView
            {
                //Size = new Size(1200, 1200),
                ReadOnly = false, // Phải để false để có thể click vào CheckBox
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                Margin = new Padding(0, 100, 0, 0),
            };

            dgvStock.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvStock.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvStock.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvStock.EnableHeadersVisualStyles = false;
            dgvStock.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvStock.Dock = DockStyle.Fill;
            dgvStock.DefaultCellStyle.SelectionBackColor = Color.FromArgb(204, 232, 255);
            dgvStock.DefaultCellStyle.SelectionForeColor = Color.Black;

            dgvStock.CellFormatting += DgvStock_CellFormatting; ;
            dgvStock.CellContentDoubleClick += DgvStock_CellContentDoubleClick; ;
            dgvStock.SelectionChanged += (s, e) => Common.Common.UpdateSelectionSum(dgvStock, lblStatus);
            dgvStock.CellValueChanged += (s, e) =>
            {
                if (dgvStock.Columns[e.ColumnIndex].Name == "Chon" && e.RowIndex >= 0)
                {
                    bool isChecked = Convert.ToBoolean(dgvStock.Rows[e.RowIndex].Cells["Chon"].Value);
                    int importId = Convert.ToInt32(dgvStock.Rows[e.RowIndex].Cells["Import_ID"]?.Value?.ToString()?.Trim() ?? "0");

                    if (isChecked)
                    {
                        if (_selectedItem.Any(i => i.Import_ID == importId)) return;
                        _selectedItem.Add(new SelectedItemModel
                        {
                            Import_ID = importId,
                            Ma_Phieu = dgvStock.Rows[e.RowIndex].Cells["Ma_Phieu"].Value?.ToString() ?? "",
                            Ten_Vat_Tu = dgvStock.Rows[e.RowIndex].Cells["Ten_Vat_Tu"].Value?.ToString() ?? "",
                            Vat_Lieu = dgvStock.Rows[e.RowIndex].Cells["Vat_Lieu"].Value?.ToString() ?? "",
                            Kich_Thuoc = dgvStock.Rows[e.RowIndex].Cells["Kich_Thuoc"].Value?.ToString() ?? "",
                            DVT = dgvStock.Rows[e.RowIndex].Cells["DVT"].Value?.ToString() ?? "",
                            Item_Code = dgvStock.Rows[e.RowIndex].Cells["Item_Code"].Value?.ToString() ?? "",
                            SL_Ton = Convert.ToDecimal(dgvStock.Rows[e.RowIndex].Cells["SL_Ton"].Value),
                            SL_Xuat = 0,
                            ID_Code = dgvStock.Rows[e.RowIndex].Cells["QC_Code"].Value?.ToString() ?? "",
                        });
                    }
                    else
                    {
                        _selectedItem.RemoveAll(impId => impId.Import_ID == importId);
                    }
                }
            };
            gbHeader.Controls.Add(dgvStock);

            // 1. Khởi tạo ContextMenuStrip
            ContextMenuStrip menuStock = new ContextMenuStrip();

            // 2. Thêm các mục (Items) vào menu
            ToolStripMenuItem itemXemChiTiet = new ToolStripMenuItem("📄 Chuyển vật tư");
            ToolStripMenuItem itemCapnhatIDCode = new ToolStripMenuItem("📋 Cập nhật ID Code");
            //ToolStripMenuItem itemXuatKho = new ToolStripMenuItem("📤 Xuất kho");

            menuStock.Items.AddRange(new ToolStripItem[] { itemXemChiTiet, itemCapnhatIDCode/*, itemSaoChep, new ToolStripSeparator(), itemXuatKho*/ });

            // 3. Gắn menu vào DataGridView
            if (AppSession.CurrentUser.Role_ID == 1)
            {
                dgvStock.ContextMenuStrip = menuStock;
            }

            // 4. Sự kiện khi click vào một mục trong menu
            itemXemChiTiet.Click += (s, e) =>
            {
                if (dgvStock.CurrentRow != null)
                {
                    // Lấy dữ liệu từ dòng đang chọn
                    var row = dgvStock.CurrentRow;
                    string id = row.Cells["Import_ID"].Value?.ToString();
                    var importId = Convert.ToInt32(row.Cells["Import_ID"].Value.ToString());
                    var maxQty = Common.Common.ParseDecimalRaw(row.Cells["SL_Nhap"].Value.ToString());
                    frmProjectMaterialTransform frmProjectMaterialTransform = new frmProjectMaterialTransform(_dtProject, importId, maxQty);
                    frmProjectMaterialTransform.ShowDialog();
                    //btnSearch.PerformClick();
                }
            };

            itemCapnhatIDCode.Click += (s, e) =>
            {
                if (dgvStock.CurrentRow != null)
                {
                    // Lấy dữ liệu từ dòng đang chọn
                    var row = dgvStock.CurrentRow;
                    string id = row.Cells["Import_ID"].Value?.ToString();
                    var importId = Convert.ToInt32(row.Cells["Import_ID"].Value.ToString());
                    var warehouseImport = new WarehouseImport() { Import_ID = importId };

                    frmUpdateIDCode frmUpdateIDCode = new frmUpdateIDCode(warehouseImport);
                    frmUpdateIDCode.ShowDialog();
                    LoadAll();
                }
            };

            // 5. QUAN TRỌNG: Xử lý để chuột phải vào dòng nào thì chọn dòng đó (thay vì chỉ hiện menu)
            dgvStock.MouseDown += (s, e) =>
            {
                if (e.Button == MouseButtons.Right)
                {
                    var hit = dgvStock.HitTest(e.X, e.Y);
                    if (hit.RowIndex >= 0)
                    {
                        // Xóa các lựa chọn cũ và chọn dòng vừa click chuột phải
                        dgvStock.ClearSelection();
                        dgvStock.Rows[hit.RowIndex].Selected = true;
                        dgvStock.CurrentCell = dgvStock.Rows[hit.RowIndex].Cells[hit.ColumnIndex];
                    }
                }
            };

            GroupBox gbAction = new GroupBox();
            gbAction.Text = "";
            gbAction.Dock = DockStyle.Top;
            gbAction.Height = 85;
            gbAction.Location = new Point(10, 115);
            pnAction.Controls.Add(gbAction);

            int fy = 20;
            gbAction.Controls.Add(new Label { Text = "Tìm kiếm:", Location = new Point(10, fy + 3), Size = new Size(70, 20), Font = new Font("Segoe UI", 9) });
            txtSearchStock = new TextBox { Location = new Point(83, fy), Size = new Size(200, 25), Font = new Font("Segoe UI", 9), PlaceholderText = "Tìm tên, ID Code, PO No..." };
            gbAction.Controls.Add(txtSearchStock);
            txtSearchStock.KeyDown += async (s, e) => { if (e.KeyCode == Keys.Enter) await LoadStock(); };
            gbAction.Controls.Add(new Label { Text = "Dự án:", Location = new Point(295, fy + 3), Size = new Size(50, 20), Font = new Font("Segoe UI", 9) });
            cboProjectFilter = new ComboBox { Location = new Point(347, fy), Size = new Size(180, 25), Font = new Font("Segoe UI", 9), DropDownStyle = ComboBoxStyle.DropDownList };
            cboProjectFilter.Items.Add("Tất cả dự án");
            cboProjectFilter.SelectedIndex = 0;
            cboProjectFilter.SelectedIndexChanged += async (s, e) => await LoadStock();
            gbAction.Controls.Add(cboProjectFilter);

            fy += 35;
            var b1 = CreateBtn("🔍 Tìm", Color.FromArgb(0, 120, 212), new Point(10, fy - 3), 80, 28);
            var b2 = CreateBtn("📦 Chỉ còn tồn", Color.FromArgb(40, 167, 69), new Point(100, fy - 3), 130, 28);
            var b3 = CreateBtn("🔄 Làm mới", Color.FromArgb(108, 117, 125), new Point(240, fy - 3), 100, 28);
            var b4 = CreateBtn("🔄 Yêu cầu xuất kho", Color.FromArgb(86, 56, 103), new Point(350, fy - 3), 200, 28);
            var b5 = CreateBtn("📝 Cập nhật chi tiết vật tư", Color.FromArgb(0, 176, 80), new Point(560, fy - 3), 195, 28);
            b1.Click += async (s, e) => await LoadStock();
            b2.Click += (s, e) => LoadStockOnly();
            b3.Click += async (s, e) => await LoadStock(true);
            b4.Click += (s, e) => btnLayDuLieu_Click();
            gbAction.Controls.Add(b1);
            gbAction.Controls.Add(b2);
            gbAction.Controls.Add(b3);
            gbAction.Controls.Add(b4);

            if (AppSession.CurrentUser.Role_ID == 1)
            {
                b5.Click += (s, ev) =>
                {
                    frmCreateItemCode frmCreateItemCode = new frmCreateItemCode("Cập nhật");
                    frmCreateItemCode.ShowDialog();
                };
                gbAction.Controls.Add(b5);
            }

            panelStockSummary = new Panel
            {
                Location = new Point(10, 10),
                //Size = new Size(1200, 60),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            panelStockSummary.Dock = DockStyle.Top;
            panelStockSummary.Height = 60;
            panelStockSummary.BringToFront();
            pnHeader.Controls.Add(panelStockSummary);

            lblStockTotal = AddStatLbl(panelStockSummary, "Tổng mục:", "0 mục", Color.FromArgb(0, 120, 212), 10);
            lblStockQty = AddStatLbl(panelStockSummary, "Tổng SL tồn:", "0", Color.FromArgb(40, 167, 69), 195);
            lblStockWeight = AddStatLbl(panelStockSummary, "Tổng KG tồn:", "0 kg", Color.FromArgb(255, 140, 0), 380);
            lblStatus = AddStatLbl(panelStockSummary, "Số lượng:", "0 kg", Color.FromArgb(254, 0, 51), 565);
        }

        private void btnLayDuLieu_Click()
        {
            // 1. Kết thúc việc biên tập ô hiện tại trên lưới gốc
            dgvStock.EndEdit();

            if (_selectedItem.Count > 0)
            {
                using (Form dlg = new Form())
                {
                    dlg.Text = "Chi Tiết Phiếu Xuất Kho";
                    dlg.StartPosition = FormStartPosition.CenterParent;
                    dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
                    dlg.Size = new Size(1100, 510);
                    dlg.BackColor = Color.White;

                    Panel pnlHeader = new Panel { Dock = DockStyle.Top, Height = 50, BackColor = Color.FromArgb(240, 240, 240) };
                    Button btnSave = new Button { Text = "💾 Lưu phiếu", Font = new Font("Segoe UI", 9, FontStyle.Bold), Width = 100, Height = 35, Location = new Point(10, 7), BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
                    Button btnDelete = new Button { Text = "🗑 Xóa dòng", Font = new Font("Segoe UI", 9, FontStyle.Bold), Width = 100, Height = 35, Location = new Point(120, 7), BackColor = Color.FromArgb(232, 17, 35), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
                    Button btnClose = new Button { Text = "✖ Thoát", Font = new Font("Segoe UI", 9, FontStyle.Bold), Width = 100, Height = 35, Location = new Point(230, 7), BackColor = Color.Gray, ForeColor = Color.White, FlatStyle = FlatStyle.Flat };

                    pnlHeader.Controls.AddRange(new Control[] { btnSave, btnDelete, btnClose });

                    DataGridView dgvSelected = new DataGridView
                    {
                        Dock = DockStyle.Fill,
                        AllowUserToAddRows = false,
                        AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                        BackgroundColor = Color.White,
                        RowHeadersVisible = true, // Bật lên để người dùng dễ click chọn dòng để xóa
                        SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                        Font = new Font("Segoe UI", 9),
                        ReadOnly = false
                    };
                    dgvSelected.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
                    dgvSelected.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                    dgvSelected.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    dgvSelected.EnableHeadersVisualStyles = false;
                    dgvSelected.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);

                    // Định nghĩa cột
                    dgvSelected.Columns.Add(new DataGridViewTextBoxColumn { Name = "Import_ID", HeaderText = "Import_ID", Visible = false });
                    dgvSelected.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ma_Phieu", HeaderText = "Mã phiếu", ReadOnly = true });
                    dgvSelected.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ten_Vat_Tu", HeaderText = "Tên vật tư", ReadOnly = true });
                    dgvSelected.Columns.Add(new DataGridViewTextBoxColumn { Name = "Vat_Lieu", HeaderText = "Vật liệu", ReadOnly = true });
                    dgvSelected.Columns.Add(new DataGridViewTextBoxColumn { Name = "Kich_Thuoc", HeaderText = "Kích thước", ReadOnly = true });
                    dgvSelected.Columns.Add(new DataGridViewTextBoxColumn { Name = "DVT", HeaderText = "ĐVT", ReadOnly = true });
                    dgvSelected.Columns.Add(new DataGridViewTextBoxColumn { Name = "Item_Code", HeaderText = "Item Code", ReadOnly = true });
                    dgvSelected.Columns.Add(new DataGridViewTextBoxColumn { Name = "SL_Ton", HeaderText = "Số lượng tồn", ReadOnly = true });
                    dgvSelected.Columns.Add(new DataGridViewTextBoxColumn { Name = "SL_Xuat", HeaderText = "Số Lượng Xuất (*)", ReadOnly = false });

                    dgvSelected.Columns.Add(new DataGridViewTextBoxColumn { Name = "ID_Code", HeaderText = "ID_Code", ReadOnly = false });
                    //dgvSelected.EditingControlShowing += DgvSelectedMakeExport_EditingControlShowing;

                    dgvSelected.CellEndEdit += (s, e) =>
                    {
                        // Chỉ kiểm tra nếu cột đang sửa là "SL_Xuat"
                        if (dgvSelected.Columns[e.ColumnIndex].Name == "SL_Xuat")
                        {
                            var row = dgvSelected.Rows[e.RowIndex];

                            // Lấy giá trị nhập vào và giá trị tồn
                            decimal slXuat = 0;
                            decimal slTon = 0;

                            // Ép kiểu an toàn (sử dụng decimal.TryParse để tránh lỗi nhập chữ)
                            decimal.TryParse(row.Cells["SL_Xuat"].Value?.ToString(), out slXuat);
                            decimal.TryParse(row.Cells["SL_Ton"].Value?.ToString(), out slTon);

                            if (slXuat > slTon)
                            {
                                MessageBox.Show($"Số lượng xuất ({slXuat}) không được lớn hơn số lượng tồn ({slTon})!",
                                                "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                                // Gán lại giá trị Xuất bằng giá trị Tồn
                                row.Cells["SL_Xuat"].Value = slTon;
                            }
                            else if (slXuat < 0)
                            {
                                // Tiện thể kiểm tra luôn trường hợp nhập số âm
                                MessageBox.Show("Số lượng xuất không được nhỏ hơn 0!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                row.Cells["SL_Xuat"].Value = 0;
                            }
                        }
                    };

                    // Nạp dữ liệu vào Grid
                    foreach (var d in _selectedItem)
                        dgvSelected.Rows.Add(d.Import_ID, d.Ma_Phieu, d.Ten_Vat_Tu, d.Vat_Lieu, d.Kich_Thuoc, d.DVT, d.Item_Code, d.SL_Ton, d.SL_Xuat, d.ID_Code);

                    // Định dạng cột SL_Xuat
                    if (dgvSelected.Columns.Contains("SL_Xuat"))
                    {
                        dgvSelected.Columns["SL_Xuat"].DefaultCellStyle.BackColor = Color.LightYellow;
                        dgvSelected.Columns["SL_Xuat"].DefaultCellStyle.ForeColor = Color.Blue;
                    }

                    // --- CẬP NHẬT: XỬ LÝ SỰ KIỆN XÓA ---
                    btnDelete.Click += (s, ev) =>
                    {
                        // Kiểm tra xem có dòng nào đang được chọn không
                        if (dgvSelected.CurrentRow != null && dgvSelected.CurrentRow.Index >= 0)
                        {
                            int rowIndex = dgvSelected.CurrentRow.Index;
                            string itemName = dgvSelected.CurrentRow.Cells["Ten_Vat_Tu"].Value?.ToString();

                            if (MessageBox.Show($"Bạn có chắc chắn muốn xóa dòng: {itemName}?", "Xác nhận xóa",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            {
                                // Xóa dòng trực tiếp trên giao diện DataGridView
                                dgvSelected.Rows.RemoveAt(rowIndex);
                            }
                        }
                        else
                        {
                            // Trường hợp không có dòng nào được chọn
                            MessageBox.Show("Vui lòng chọn một dòng để xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    };

                    // Xử lý sự kiện Lưu
                    btnSave.Click += (s, ev) =>
                    {
                        dgvSelected.EndEdit();

                        if (dgvSelected.Rows.Count == 0)
                        {
                            MessageBox.Show("Không có dữ liệu để xuất!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        try
                        {
                            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "pxk_template.xlsx");
                            if (!File.Exists(templatePath))
                            {
                                MessageBox.Show("Không tìm thấy file template!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }

                            var saveDialog = new SaveFileDialog
                            {
                                Title = "Lưu Phiếu Xuất Kho",
                                Filter = "Excel Files|*.xlsx",
                                FileName = $"PXK_{DateTime.Now:ddMMyyyy_HHmm}",
                                InitialDirectory = Directory.Exists(@"D:\RÁC") ? @"D:\RÁC" : Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                            };

                            if (saveDialog.ShowDialog() != DialogResult.OK) return;

                            File.Copy(templatePath, saveDialog.FileName, true);
                            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(saveDialog.FileName)))
                            {
                                var ws = package.Workbook.Worksheets[0]; // Lấy sheet "PXK"

                                // 1. Thay thế <<DATE>> (Giả định nằm ở ô H8 dựa trên cấu trúc template của bạn)
                                // Tìm kiếm text <<DATE>> trong vùng Header để thay thế
                                for (int r = 1; r <= 10; r++)
                                {
                                    for (int c = 1; c <= 10; c++)
                                    {
                                        if (ws.Cells[r, c].Text.Contains("<<DATE>>"))
                                        {
                                            ws.Cells[r, c].Value = ws.Cells[r, c].Text.Replace("<<DATE>>", DateTime.Now.ToString("dd/MM/yyyy"));
                                        }
                                    }
                                }

                                ReplaceCell(ws, "<<PROJECT-NAME>>", cboProjectFilter.Text ?? "");
                                ReplaceCell(ws, "<<USER>>", AppSession.CurrentUser.Username ?? "");

                                int startRow = 11; // Dòng bắt đầu điền dữ liệu (Dòng có STT 1)
                                int detailCount = dgvSelected.Rows.Count;
                                decimal totalQty = 0;

                                // 2. Chèn dòng nếu nhiều hơn 1 item để không đè lên phần chữ ký
                                if (detailCount > 1)
                                {
                                    ws.InsertRow(startRow + 1, detailCount - 1, startRow);
                                }

                                // 3. Vòng lặp điền dữ liệu
                                for (int i = 0; i < detailCount; i++)
                                {
                                    var row = dgvSelected.Rows[i];
                                    int currentRow = startRow + i;

                                    decimal slXuat = Convert.ToDecimal(row.Cells["SL_Xuat"].Value ?? 0);
                                    totalQty += slXuat;

                                    ws.Cells[currentRow, 1].Value = i + 1; // Cột No (A)
                                    ws.Cells[currentRow, 2].Value = row.Cells["Item_Code"].Value; // Cột Code (B)
                                    ws.Cells[currentRow, 3].Value = row.Cells["Ten_Vat_Tu"].Value; // Cột Name (C)
                                    ws.Cells[currentRow, 4].Value = /*row.Cells["Ma_Phieu"].Value*/ ""; // Cột DWG No (D)
                                    ws.Cells[currentRow, 5].Value = row.Cells["Kich_Thuoc"].Value; // Cột Size (E)
                                    ws.Cells[currentRow, 6].Value = row.Cells["Vat_Lieu"].Value; // Cột Grade (F)
                                    ws.Cells[currentRow, 7].Value = slXuat; // Cột Q'ty (G)
                                    ws.Cells[currentRow, 8].Value = row.Cells["DVT"].Value; // Cột Unit (H)
                                    ws.Cells[currentRow, 9].Value = row.Cells["ID_Code"].Value; // Cột ID_Code (I)
                                }

                                // 4. Tìm và thay thế <<SUM>> bằng tổng thực tế
                                // Duyệt tìm ô chứa <<SUM>> bên dưới vùng dữ liệu vừa điền
                                int searchEndRow = startRow + detailCount + 5;
                                for (int r = startRow + detailCount; r <= searchEndRow; r++)
                                {
                                    for (int c = 1; c <= 10; c++)
                                    {
                                        if (ws.Cells[r, c].Text.Contains("<<SUM>>"))
                                        {
                                            ws.Cells[r, c].Value = totalQty;
                                            ws.Cells[r, c].Style.Font.Bold = true;
                                            ws.Cells[r, c].Style.Numberformat.Format = "#,##0";
                                        }
                                    }
                                }

                                package.Save();
                            }

                            //MessageBox.Show("Xuất phiếu xuất kho thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            //dlg.DialogResult = DialogResult.OK;
                            var result = MessageBox.Show(
                            $"✅ Xuất phiếu nhập kho thành công!\nFile: {saveDialog.FileName}\n\nBạn có muốn mở file ngay không?",
                            "Thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                            if (result == DialogResult.Yes)
                            {
                                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                                {
                                    FileName = saveDialog.FileName,
                                    UseShellExecute = true
                                });
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Lỗi: " + ex.Message, "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    };

                    btnClose.Click += (s, ev) => { dlg.Close(); };

                    dlg.Controls.Add(dgvSelected);
                    dlg.Controls.Add(pnlHeader);
                    dlg.ShowDialog();
                }
            }
            else
            {
                MessageBox.Show("Vui lòng chọn ít nhất một dòng!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ReplaceCell(OfficeOpenXml.ExcelWorksheet ws, string placeholder, string value)
        { for (int r = 1; r <= ws.Dimension.End.Row; r++) for (int c = 1; c <= ws.Dimension.End.Column; c++) if (ws.Cells[r, c].Value?.ToString() == placeholder) ws.Cells[r, c].Value = value; }


        private Button CreateBtn(string text, Color color, Point loc, int w, int h)
        {
            var btn = new Button { Text = text, Location = loc, Size = new Size(w, h), BackColor = color, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Cursor = Cursors.Hand };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        private Label AddStatLbl(Panel p, string title, string value, Color color, int x)
        {
            var card = new Panel { Location = new Point(x, 8), Size = new Size(180, 42), BackColor = color };
            p.Controls.Add(card);
            card.Controls.Add(new Label { Text = title, Font = new Font("Segoe UI", 8, FontStyle.Bold), ForeColor = Color.White, Location = new Point(6, 3), Size = new Size(208, 18) });
            var lbl = new Label { Text = value, Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.White, Location = new Point(6, 22), Size = new Size(208, 18) };
            card.Controls.Add(lbl);
            return lbl;
        }

        private void LoadStockOnly()
        {
            try { if (cboProjectFilter.Items.Count <= 0) return; if (dgvStock != null) BindStockGrid(_service.GetStockWithRemaining(cboProjectFilter.SelectedText ?? "")); }
            catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private async Task LoadStock(bool isRefesh = false)
        {
            try
            {
                if (dgvStock == null) return;
                if ((isRefesh))
                {
                    txtSearchStock.Text = "";
                    _selectedItem.Clear();
                }
                string kw = txtSearchStock?.Text.Trim() ?? "";
                string project = (cboProjectFilter != null && cboProjectFilter.SelectedIndex > 0) ? cboProjectFilter.SelectedItem.ToString() : "";
                BindStockGrid(_service.GetStock(project, kw));
            }
            catch (Exception ex) { MessageBox.Show("Lỗi tải tồn kho: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }
        private void BindStockGrid(List<WarehouseStock> stocks)
        {
            // Xóa cột cũ và dữ liệu cũ trước khi nạp lại (tránh trùng lặp cột khi gọi hàm nhiều lần)
            dgvStock.Columns.Clear();
            dgvStock.DataSource = null;

            // Thêm cột Checkbox trước
            DataGridViewCheckBoxColumn checkColumn = new DataGridViewCheckBoxColumn();
            checkColumn.Name = "Chon";
            checkColumn.HeaderText = "Chọn";
            checkColumn.Width = 50;
            checkColumn.ReadOnly = false; // Cho phép tương tác
            dgvStock.Columns.Add(checkColumn);

            // Gán DataSource
            dgvStock.DataSource = stocks.Select(s => new
            {
                Import_ID = s.Import_ID,
                Ma_Phieu = s.Import_No,
                Ngay_Nhap = s.Import_Date.HasValue ? s.Import_Date.Value.ToString("dd/MM/yyyy") : "",
                Ten_Vat_Tu = s.Item_Name,
                Vat_Lieu = s.Material,
                Kich_Thuoc = s.Size,
                DVT = s.UNIT,
                Item_Code = s.ID_Code,
                //PO_No = s.PONo,
                //Ma_DA = s.Project_Code,
                Vi_Tri = s.Notes,
                SL_Nhap = s.Qty_Import,
                SL_Xuat = s.Qty_Exported,
                SL_Ton = s.Qty_Stock,
                QC_Code = s.QC_Code,
                QC_Status = s.QC_Status,
                QC_Remark = s.Remarks,
            }).ToList();

            // Thiết lập ReadOnly cho tất cả các cột ngoại trừ cột "Chon"
            foreach (DataGridViewColumn col in dgvStock.Columns)
            {
                if (col.Name != "Chon")
                {
                    col.ReadOnly = true;
                }
            }

            // Ẩn cột ID
            if (dgvStock.Columns.Contains("Import_ID")) dgvStock.Columns["Import_ID"].Visible = false;

            // Tính toán tổng số lượng
            decimal tQ = 0, tW = 0;
            if (stocks != null)
            {
                foreach (var s in stocks)
                {
                    tQ += s.Qty_Stock;
                    tW += s.Weight_Stock;
                }
                if (lblStockTotal != null) lblStockTotal.Text = $"{stocks.Count} mục";
                if (lblStockQty != null) lblStockQty.Text = tQ.ToString("N2");
                if (lblStockWeight != null) lblStockWeight.Text = tW.ToString("N2") + " kg";
            }

            // Thủ thuật nhỏ: Để CheckBox phản hồi click chuột ngay lập tức (không cần click 2 lần)
            dgvStock.EditMode = DataGridViewEditMode.EditOnEnter;
            SyncDataGridViewWithList(dgvStock, _selectedItem);
        }

        private void SyncDataGridViewWithList(DataGridView dgv, List<SelectedItemModel> selectedItem)
        {
            // 1. Kiểm tra điều kiện đầu vào
            if (dgv.Rows.Count == 0 || selectedItem == null) return;

            // 2. Tối ưu hiệu suất: Chuyển List ID sang HashSet để tìm kiếm nhanh O(1)
            // Thay vì duyệt List nhiều lần, HashSet giúp kiểm tra sự tồn tại tức thì.
            var selectedIds = new HashSet<int>(selectedItem.Select(p => p.Import_ID));

            // 3. Tạm dừng vẽ giao diện để tăng tốc độ xử lý nếu dữ liệu cực lớn (tùy chọn)
            // dgv.SuspendLayout(); 

            try
            {
                // Kết thúc biên tập ô để đảm bảo dữ liệu đồng bộ
                dgv.EndEdit();

                foreach (DataGridViewRow row in dgv.Rows)
                {
                    // Bỏ qua dòng trống mới (nếu có)
                    if (row.IsNewRow) continue;

                    // Lấy giá trị ID từ cột "ID" của dòng hiện tại
                    if (row.Cells["Import_ID"].Value != null && int.TryParse(row.Cells["Import_ID"].Value.ToString(), out int rowId))
                    {
                        // Nếu ID của dòng nằm trong danh sách chọn, tích checkbox "Chon"
                        if (selectedIds.Contains(rowId))
                        {
                            row.Cells["Chon"].Value = true;
                        }
                        else
                        {
                            row.Cells["Chon"].Value = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra: " + ex.Message);
            }
            finally
            {
                // dgv.ResumeLayout();
            }
        }

        private void DgvStock_CellContentDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            int importId = (int)dgvStock.Rows[e.RowIndex].Cells["Import_ID"].Value;
            OpenFormModifyItem(importId);
        }

        private void DgvStock_CellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string col = dgvStock.Columns[e.ColumnIndex].Name;
            if (col == "SL_Ton" || col == "KG_Ton" || col == "SL_Nhap")
            {
                decimal val = e.Value != null ? Convert.ToDecimal(e.Value) : 0;
                e.CellStyle.ForeColor = val > 0 ? Color.FromArgb(40, 167, 69) : Color.FromArgb(220, 53, 69);
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }

            // 2. Định dạng cho cột "Trạng thái" (String)
            var statusRules = new List<StringRule>
            {
                new StringRule { Value = "Pass", CellColor = Color.SeaGreen },
                new StringRule { Value = "Fail", CellColor = Color.Red },
                new StringRule { Value = "Hold", CellColor = Color.Orange },
                new StringRule { Value = "Pending", CellColor = Color.DimGray }
            };
            Common.Common.ApplyCustomFormatting(e, dgvStock, "QC_Status", statusRules, null);
        }

        private void OpenFormModifyItem(int importId)
        {
            if (AppSession.CurrentUser.Role_ID == 1)
            {
                // Sử dụng khối lệnh using để khởi tạo và tự động giải phóng tài nguyên Form
                using (Form frm = new Form())
                {
                    // --- 1. Cấu hình giao diện Form ---
                    frm.Text = "Cập nhật dữ liệu - Product Entry V2";
                    frm.Size = new Size(400, 380);
                    frm.StartPosition = FormStartPosition.CenterScreen;
                    frm.FormBorderStyle = FormBorderStyle.FixedDialog;
                    frm.MaximizeBox = false;
                    frm.Font = new Font("Segoe UI", 10);

                    // Biến hỗ trợ vị trí hiển thị
                    int startY = 25;
                    int spacing = 45;

                    // --- 2. Khởi tạo 4 cặp Label và TextBox ---
                    // Item_Code
                    Label lblItemCode = new Label() { Text = "Item_Code:", Location = new Point(20, startY), AutoSize = true };
                    TextBox txtItemCode = new TextBox() { Location = new Point(150, startY - 3), Size = new Size(200, 25) };

                    // Qty_Import
                    Label lblQty = new Label() { Text = "Qty_Import:", Location = new Point(20, startY + spacing), AutoSize = true };
                    TextBox txtQty = new TextBox() { Location = new Point(150, startY + spacing - 3), Size = new Size(200, 25) };

                    // Weight_kg
                    Label lblWeight = new Label() { Text = "Weight_kg:", Location = new Point(20, startY + (spacing * 2)), AutoSize = true };
                    TextBox txtWeight = new TextBox() { Location = new Point(150, startY + (spacing * 2) - 3), Size = new Size(200, 25) };

                    // Size
                    Label lblSize = new Label() { Text = "Size:", Location = new Point(20, startY + (spacing * 3)), AutoSize = true };
                    TextBox txtSize = new TextBox() { Location = new Point(150, startY + (spacing * 3) - 3), Size = new Size(200, 25) };

                    // Name
                    Label lblName = new Label() { Text = "Name:", Location = new Point(20, startY + (spacing * 4)), AutoSize = true };
                    TextBox txtName = new TextBox() { Location = new Point(150, startY + (spacing * 4) - 3), Size = new Size(200, 25) };

                    // QCCode
                    Label lblQCCode = new Label() { Text = "ID Code:", Location = new Point(20, startY + (spacing * 5)), AutoSize = true };
                    TextBox txtQCCode = new TextBox() { Location = new Point(150, startY + (spacing * 5) - 3), Size = new Size(200, 25) };

                    // --- 3. Cấu hình các Button ---
                    // Button Cancel (Nền xám, chữ trắng)
                    Button btnCancel = new Button()
                    {
                        Text = "Cancel",
                        Location = new Point(150, 275),
                        Size = new Size(90, 35),
                        BackColor = Color.Gray,
                        ForeColor = Color.White,
                        FlatStyle = FlatStyle.Flat
                    };
                    btnCancel.Click += (s, e) => { frm.Close(); };

                    // Button Save (Nền xanh, chữ trắng)
                    Button btnSave = new Button()
                    {
                        Text = "Save",
                        Location = new Point(260, 275),
                        Size = new Size(90, 35),
                        BackColor = Color.DodgerBlue,
                        ForeColor = Color.White,
                        FlatStyle = FlatStyle.Flat
                    };

                    // --- 4. Logic lấy dữ liệu khi Click Save ---
                    btnSave.Click += (s, e) =>
                    {
                        //[cite_start]// Truy xuất giá trị từ các TextBox [cite: 108, 111]
                        string itemCode = txtItemCode.Text;
                        string qty = txtQty.Text;
                        string weight = txtWeight.Text;
                        string sizeValue = txtSize.Text;
                        string name = txtName.Text;
                        string qcCode = txtQCCode.Text.Trim().ToUpper();

                        var w = new WarehouseImport()
                        {
                            Import_ID = importId,
                            ID_Code = itemCode,
                            Qty_Import = !string.IsNullOrEmpty(qty) ? Convert.ToDecimal(qty.Trim()) : 0,
                            Size = sizeValue,
                            Item_Name = name,
                            QC_Code = qcCode
                        };

                        if (!string.IsNullOrEmpty(txtItemCode.Text))
                        {
                            _service.ModifyIDCodeOfWarehouse(w);
                        }
                        if (!string.IsNullOrEmpty(txtQty.Text))
                        {
                            _service.ModifyQtyImportOfWarehouseImport(w);
                        }
                        if (!string.IsNullOrEmpty(txtWeight.Text))
                        {
                            _service.ModifyWeightOfWarehouseImport(w);
                        }
                        if (!string.IsNullOrEmpty(txtSize.Text))
                        {
                            _service.ModifySizeOfWarehouseImport(w);
                        }
                        if (!string.IsNullOrEmpty(txtName.Text))
                        {
                            _service.ModifyNameOfWarehouseImport(w);
                        }
                        if (!string.IsNullOrEmpty(txtQCCode.Text))
                        {
                            _service.UpdateIDCode(w);
                        }
                        // Hiển thị kết quả lấy được để kiểm tra
                        string info = $"Dữ liệu đã thu thập:\n" +
                                      $"- Item Code: {itemCode}\n" +
                                      $"- Qty: {qty}\n" +
                                      $"- Weight: {weight}\n" +
                                      $"- Size: {sizeValue}";

                        MessageBox.Show(info, "Kết quả lưu dữ liệu", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // Sau khi xử lý xong có thể đóng form hoặc giữ lại tùy ý
                        frm.DialogResult = DialogResult.OK;
                    };

                    // --- 5. Thêm Controls vào Form và hiển thị ---
                    frm.Controls.AddRange(new Control[] {
                        lblItemCode, txtItemCode,
                        lblQty, txtQty,
                        lblWeight, txtWeight,
                        lblSize, txtSize,
                        lblName, txtName,
                        lblQCCode, txtQCCode,
                        btnCancel, btnSave
                    });

                    frm.AcceptButton = btnSave; // Nhấn Enter để Save [cite: 115]
                    frm.CancelButton = btnCancel; // Nhấn Esc để Cancel [cite: 115]

                    frm.ShowDialog();
                }
            }
        }

        private void PopulateTreeView(List<TransferLog> logList)
        {
            // 1. Validation kiểm tra đầu vào
            if (logList == null || logList.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu để hiển thị.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 2. Tối ưu UI: Dừng vẽ giao diện trong lúc thêm hàng loạt Node
            tvLocations.BeginUpdate();

            try
            {
                tvLocations.Nodes.Clear();

                // Sử dụng Dictionary để kiểm tra và quản lý các Root Node (New_Value_Location) không bị trùng lặp
                Dictionary<string, TreeNode> rootNodesDict = new Dictionary<string, TreeNode>();

                foreach (var item in logList)
                {
                    string rootKey = item.NewValueLocation;

                    // Nếu Root Node (New_Value_Location) chưa tồn tại trong Dictionary, tiến hành tạo mới
                    if (!rootNodesDict.ContainsKey(rootKey))
                    {
                        TreeNode newRoot = new TreeNode(rootKey);
                        newRoot.Tag = "ROOT"; // Đánh dấu đây là Node cha

                        // Thêm vào TreeView và lưu lại vào Dictionary để tái sử dụng ở các dòng sau
                        tvLocations.Nodes.Add(newRoot);
                        rootNodesDict.Add(rootKey, newRoot);
                    }

                    // Lấy ra Node cha tương ứng (dù là mới tạo hay đã tồn tại trước đó)
                    TreeNode targetRootNode = rootNodesDict[rootKey];

                    // Kết hợp tất cả thông tin các cột còn lại thành 1 chuỗi làm nội dung cho Child Node
                    // Định dạng mẫu: [Tên mặt hàng] - Size: [Kích thước] - SL: [Số lượng] (Từ vị trí: [Vị trí cũ])
                    string childText = $"{item.ItemName} | Size: {item.Size} | SL: {item.NumberTransform} | Mượn: {item.OldValueLocation}";

                    TreeNode childNode = new TreeNode(childText);
                    childNode.Tag = item; // Gán toàn bộ Object vào Tag để khi cần click chọn có thể lấy lại toàn bộ data gốc

                    // Thêm Child Node vào Root Node tương ứng
                    targetRootNode.Nodes.Add(childNode);
                }

                // Tùy chọn: Tự động mở rộng (bung) tất cả các nhánh sau khi nạp xong dữ liệu
                tvLocations.ExpandAll();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Đã xảy ra lỗi khi dựng cây dữ liệu: {ex.Message}", "Lỗi hệ thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Bắt buộc phải gọi EndUpdate để giao diện TreeView cập nhật lại bình thường
                tvLocations.EndUpdate();
            }
        }

        private void tvLocations_DrawNode(object sender, DrawTreeNodeEventArgs e)
        {
            // 1. Kiểm tra Validation, nếu node không hiển thị thì bỏ qua để tối ưu hiệu năng
            if (e.Node == null || !e.Bounds.IntersectsWith(tvLocations.ClientRectangle))
                return;

            // 2. Xác định trạng thái Node (Có đang được chọn hay không)
            bool isSelected = (e.State & TreeNodeStates.Selected) != 0;

            // 3. XỬ LÝ NỀN: Luôn vẽ nền trắng (hoặc màu nền mặc định của TreeView) 
            // Bỏ hoàn toàn phần tô nền màu xanh dương cũ để đáp ứng yêu cầu của bạn
            e.Graphics.FillRectangle(SystemBrushes.Window, e.Bounds);

            // Chuẩn bị tọa độ vẽ chữ để căn chỉnh lề cho đẹp
            float currentX = e.Bounds.X + 2;
            float currentY = e.Bounds.Y + 4;

            // 4. TIẾN HÀNH VẼ TEXT THEO TỪNG LOẠI NODE
            if (e.Node.Tag != null && e.Node.Tag.ToString() == "ROOT")
            {
                // --- VẼ NODE CHA (NEW_VALUE_LOCATION) ---
                // Xác định kiểu chữ cho Node Cha: Mặc định là In đậm (Bold)
                FontStyle style = FontStyle.Bold;

                // Nếu Node Cha này đang được chọn, kết hợp thêm Gạch chân (Underline)
                if (isSelected)
                {
                    style |= FontStyle.Underline;
                }

                using (Font rootFont = new Font(e.Node.TreeView.Font, style))
                {
                    // Sử dụng màu xanh đen công nghiệp cho Node cha
                    Brush rootColor = Brushes.DarkSlateBlue;
                    e.Graphics.DrawString(e.Node.Text, rootFont, rootColor, currentX, currentY);
                }
            }
            else if (e.Node.Tag is TransferLog item)
            {
                // --- VẼ NODE CON (CÓ PHÂN TÁCH MÀU SẮC) ---
                Font defaultFont = e.Node.TreeView.Font;

                // Định nghĩa các kiểu chữ tùy biến
                Font nameFont = isSelected ? new Font(defaultFont, FontStyle.Underline) : defaultFont;
                Font sizeFont = isSelected ? new Font(defaultFont, FontStyle.Underline) : defaultFont;
                Font qtyFont = isSelected ? new Font(defaultFont, FontStyle.Underline) : defaultFont;
                Font italicFont = isSelected ? new Font(defaultFont, FontStyle.Italic | FontStyle.Underline) : new Font(defaultFont, FontStyle.Italic);

                // Đoạn 1: Tên mặt hàng (Màu đen)
                string part1 = $"{item.ItemName}  ";
                e.Graphics.DrawString(part1, nameFont, Brushes.Black, currentX, currentY);
                currentX += e.Graphics.MeasureString(part1, nameFont).Width;

                // Đoạn 2: Size (Màu xanh xám)
                string part2 = $"[Size: {item.Size}]  ";
                e.Graphics.DrawString(part2, sizeFont, Brushes.DimGray, currentX, currentY);
                currentX += e.Graphics.MeasureString(part2, sizeFont).Width;

                // Đoạn 3: Số lượng (Màu xanh lá nổi bật)
                string part3 = $"|  SL: {item.NumberTransform}  ";
                e.Graphics.DrawString(part3, qtyFont, Brushes.ForestGreen, currentX, currentY);
                currentX += e.Graphics.MeasureString(part3, qtyFont).Width;

                // Đoạn 4: Vị trí gốc (Chữ nghiêng màu xám)
                string part4 = $"|  Gốc: {item.OldValueLocation}";
                e.Graphics.DrawString(part4, italicFont, Brushes.Gray, currentX, currentY);

                // Giải phóng bộ nhớ cho các Font tự tạo để tránh rò rỉ bộ nhớ (Memory Leak)
                if (isSelected)
                {
                    nameFont.Dispose();
                    sizeFont.Dispose();
                    qtyFont.Dispose();
                    italicFont.Dispose();
                }
                else
                {
                    italicFont.Dispose(); // Font này luôn được tạo mới ở nhánh else nên phải giải phóng
                }
            }
        }
    }
}
