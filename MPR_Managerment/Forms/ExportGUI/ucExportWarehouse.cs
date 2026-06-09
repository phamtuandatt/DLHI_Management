using MPR_Managerment.Common;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;
using MPR_Managerment.Services;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;

namespace MPR_Managerment.Forms.ExportGUI
{
    public partial class ucExportWarehouse : UserControl
    {
        private DataTable _dtProject = new DataTable();
        private DataTable _dtStock = new DataTable();
        private DataTable dtSelected = new DataTable();

        private ProjectService _projectServices = new ProjectService();
        private WarehouseService _warehouseServies = new WarehouseService();

        private Dictionary<string, string> _exportQue = new Dictionary<string, string>();
        private BindingSource _bindingSource = new BindingSource();
        private Helpers.DataGridViewFilterHelper? _filterHelper;
        private CheckBox chkFilter;

        //private DataGridView dgvKho;

        private bool _isLoaded = false;

        public ucExportWarehouse()
        {
            InitializeComponent();
            //frmAIChat.Attach(this); // Bỏ dòng này vì đây là UserControl, không phải Form chính

            LoadProject();
            InitGridSelected();


            FormartGrid(dgvKho, Color.FromArgb(0, 120, 212), true);
            FormartGrid(dgvExportQue, Color.FromArgb(255, 140, 0), false);
            FormartGrid(dgvHis, Color.FromArgb(0, 120, 212), true);
        }

        private void FormartGrid(DataGridView dataGridView, Color colorHeader, bool isReadOnly)
        {
            dataGridView.ReadOnly = isReadOnly;
            dataGridView.AllowUserToAddRows = false;
            //dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView.BackgroundColor = Color.White;
            dataGridView.BorderStyle = BorderStyle.FixedSingle;
            dataGridView.RowHeadersVisible = false;
            dataGridView.Font = new Font("Segoe UI", 9);
            dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            dataGridView.ColumnHeadersDefaultCellStyle.BackColor = colorHeader;
            dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dataGridView.EnableHeadersVisualStyles = false;
            dataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
        }

        private async void LoadProject()
        {
            _dtProject = await _projectServices.GetProjects();

            cboProject.DisplayMember = "ProjectCode";
            cboProject.ValueMember = "ProjectCode";

            cboProjectCheck.DisplayMember = "ProjectCode";
            cboProjectCheck.ValueMember = "ProjectCode";

            cboProject.DataSource = _dtProject;
            cboProjectCheck.DataSource = _dtProject.Copy();

            _isLoaded = true;
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private async void btnSearch_Click(object sender, EventArgs e)
        {
            if (_isLoaded)
            {
                _dtStock = await _warehouseServies.GetImportForExport(cboProject.SelectedValue.ToString());
                _bindingSource.DataSource = _dtStock;
                dgvKho.DataSource = _bindingSource;
                _filterHelper = new Helpers.DataGridViewFilterHelper(dgvKho, _dtStock, _bindingSource);
                _filterHelper.EnableFilter(true);

                dgvKho.Columns["Import_ID"].Visible = false;
                dgvKho.Columns["PO_ID"].Visible = false;
                dgvKho.Columns["PO_Detail_ID"].Visible = false;
                //dgvKho.Columns["RIR_ID"].Visible = false;
                //dgvKho.Columns["Item_ID"].Visible = false;
                //dgvKho.Columns["Warehouse_ID"].Visible = false;
                //dgvKho.Columns["Item_Code"].Visible = false;
                dgvKho.Columns["ID_Code"].Visible = false;

                dgvKho.Columns["Location"].Visible = false;
                //dgvKho.Columns["Notes"].Visible = false;
                dgvKho.Columns["MTRno"].Visible = false;
                dgvKho.Columns["Heatno"].Visible = false;
                //dgvKho.Columns["InvoiceNo"].Visible = false;
                //dgvKho.Columns["InvoiceDate"].Visible = false;

                dgvKho.Columns["Project_Code"].Visible = false;
                dgvKho.Columns["PONo"].Visible = false;
                dgvKho.Columns["MPR_No"].Visible = false;
                dgvKho.Columns["PO_Date"].Visible = false;
                dgvKho.Columns["WorkorderNo"].Visible = false;
            }
        }


        private void txtSearchItem_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (_dtStock.Rows.Count <= 0 || string.IsNullOrEmpty(txtSearchItem.Text)) return;
                var lstProperty = new List<string>()
                {
                    "Item_Name", "Material", "Size", "ID_Code" // Tên cột phải khớp chính xác với DataTable
                };
                DataView dv = Common.Common.Search(txtSearchItem.Text.Trim(), _dtStock, lstProperty);
                dgvKho.DataSource = dv;
            }
        }

        private void dgvKho_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (cboProject.Items.Count <= 0 || cboProject.SelectedIndex == -1) return;

            if (e.RowIndex >= 0)
            {
                string key = string.Empty;
                string value = string.Empty;
                try
                {
                    // Lấy dữ liệu dòng từ Grid 1 (bảng Warehouse_Import)
                    DataGridViewRow currentRow = dgvKho.Rows[e.RowIndex];
                    DataRowView drv = (DataRowView)dgvKho.Rows[e.RowIndex].DataBoundItem;
                    DataRow sourceRow = drv.Row;

                    // Lấy số lượng tối đa hiện có trong kho để cảnh báo
                    decimal maxQtyInStock = Convert.ToDecimal(sourceRow["Qty_Stock"]);

                    key = $"{sourceRow["Import_ID"].ToString()}";
                    value = sourceRow["ID_Code"].ToString();
                    _exportQue.Add(key, value);

                    // Hiển thị Popup nhập số lượng
                    OpenInputQuantityForm(maxQtyInStock, (qtyToExport) =>
                    {
                        // Sau khi nhập hợp lệ, tiến hành add vào Grid 2
                        AddRowToExportGrid(sourceRow, qtyToExport);
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this.FindForm(), "Lỗi xuất kho: " + ex.Message + " - " + value, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
        }

        private void AddRowToExportGrid(DataRow sourceRow, decimal qty)
        {
            // Kiểm tra xem Item này đã được add vào danh sách xuất chưa (dựa trên Import_ID)
            string filter = $"Import_ID = {sourceRow["Import_ID"]}";
            DataRow[] existingRows = dtSelected.Select(filter);

            if (existingRows.Length > 0)
            {
                // Nếu đã chọn rồi thì cộng dồn số lượng xuất
                decimal currentQty = Convert.ToDecimal(existingRows[0]["Qty_Export"]);
                existingRows[0]["Qty_Export"] = currentQty + qty;
            }
            else
            {
                // Tạo dòng mới cho Grid 2
                DataRow newRow = dtSelected.NewRow();

                // Map dữ liệu từ Grid 1 sang Grid 2
                newRow["Export_No"] = "XK-" + sourceRow["Import_No"];
                newRow["Import_ID"] = sourceRow["Import_ID"];
                newRow["Item_Name"] = sourceRow["Item_Name"];
                newRow["Material"] = sourceRow["Material"];
                newRow["Size"] = sourceRow["Size"];
                newRow["UNIT"] = sourceRow["UNIT"];
                newRow["ID_Code"] = sourceRow["ID_Code"];
                newRow["Weight_kg"] = sourceRow["Weight_Import"];
                newRow["Project_Code"] = sourceRow["Project_Code"];
                newRow["WorkorderNo"] = sourceRow["WorkorderNo"] ?? null;
                //newRow["Warehouse_ID"] = sourceRow["Warehouse_ID"] ?? null;

                // Gán số lượng xuất từ Popup
                newRow["Qty_Export"] = qty;

                // Các thông tin mặc định khác (nếu có)
                newRow["Export_Date"] = DateTime.Now;
                newRow["Created_Date"] = DateTime.Now;
                newRow["QC_Code"] = sourceRow["QC_Code"] ?? "";

                dtSelected.Rows.Add(newRow);
            }

            dgvExportQue.Refresh();
        }

        private void InitGridSelected()
        {
            // Định nghĩa các cột cho Grid 2 theo danh sách thuộc tính của bạn
            dtSelected.Columns.Add("Export_ID", typeof(int));
            dtSelected.Columns.Add("Export_No", typeof(string));
            dtSelected.Columns.Add("Export_Date", typeof(DateTime));
            dtSelected.Columns.Add("Import_ID", typeof(int));
            dtSelected.Columns.Add("Item_Name", typeof(string));
            dtSelected.Columns.Add("Material", typeof(string));
            dtSelected.Columns.Add("Size", typeof(string));
            dtSelected.Columns.Add("UNIT", typeof(string));
            dtSelected.Columns.Add("Qty_Export", typeof(decimal)); // Cột số lượng xuất
            dtSelected.Columns.Add("Weight_kg", typeof(decimal));
            dtSelected.Columns.Add("ID_Code", typeof(string));
            dtSelected.Columns.Add("Project_Code", typeof(string));
            dtSelected.Columns.Add("WorkorderNo", typeof(string));
            dtSelected.Columns.Add("Export_To", typeof(string));
            dtSelected.Columns.Add("Purpose", typeof(string));
            dtSelected.Columns.Add("Notes", typeof(string));
            dtSelected.Columns.Add("Created_By", typeof(string));
            dtSelected.Columns.Add("Created_Date", typeof(DateTime));
            dtSelected.Columns.Add("Warehouse_ID", typeof(int));
            dtSelected.Columns.Add("QC_Code", typeof(string));

            dgvExportQue.DataSource = dtSelected;

            // --- PHẦN CẬP NHẬT ĐỂ NHẬP ĐƯỢC Ô NOTES ---

            // 2. Mở khóa tổng thể Grid (Quan trọng nhất)
            dgvExportQue.ReadOnly = false;

            // 3. Khóa tất cả các cột bằng vòng lặp
            foreach (DataGridViewColumn col in dgvExportQue.Columns)
            {
                col.ReadOnly = true;
                // Đổi màu nền nhẹ để báo hiệu các cột này không được sửa
                col.DefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
            }

            // 4. Mở khóa DUY NHẤT cột Notes
            if (dgvExportQue.Columns.Contains("Notes"))
            {
                dgvExportQue.Columns["Notes"].ReadOnly = false;

                // Làm nổi bật ô Notes để người dùng biết là nhập được
                dgvExportQue.Columns["Notes"].HeaderText = "✏️ Ghi chú";
                dgvExportQue.Columns["Notes"].DefaultCellStyle.BackColor = Color.White;
                dgvExportQue.Columns["Notes"].DefaultCellStyle.ForeColor = Color.Blue;
            }

            // --- KẾT THÚC PHẦN CẬP NHẬT ---


            // Cấu hình tiêu đề cột cho dễ nhìn
            dgvExportQue.Columns["Item_Name"].HeaderText = "Tên vật tư";
            dgvExportQue.Columns["Qty_Export"].HeaderText = "SL Xuất";
            dgvExportQue.Columns["ID_Code"].HeaderText = "Mã ITEM";

            dgvExportQue.Columns["Export_ID"].Visible = false;
            dgvExportQue.Columns["Export_No"].Visible = false;
            dgvExportQue.Columns["Export_To"].Visible = false;
            dgvExportQue.Columns["Purpose"].Visible = false;
            dgvExportQue.Columns["Warehouse_ID"].Visible = false;
            dgvExportQue.Columns["Import_ID"].Visible = false;
        }

        private void OpenInputQuantityForm(decimal maxQty, Action<decimal> onSuccess)
        {
            using (Form frm = new Form())
            {
                frm.Text = "Xác nhận số lượng xuất";
                frm.Size = new Size(350, 220);
                frm.StartPosition = FormStartPosition.CenterParent;
                frm.FormBorderStyle = FormBorderStyle.FixedDialog;

                Label lblWarn = new Label
                {
                    Text = $"Tồn kho dòng này: {maxQty}",
                    ForeColor = Color.Red,
                    Location = new Point(20, 15),
                    AutoSize = true,
                    Font = new Font("Segoe UI", 9, FontStyle.Bold)
                };

                Label lblText = new Label { Text = "Số lượng xuất:", Location = new Point(20, 60), AutoSize = true };
                NumericUpDown num = new NumericUpDown
                {
                    Location = new Point(120, 58),
                    Width = 150,
                    Maximum = 1000000,
                    DecimalPlaces = 2,
                    Value = maxQty
                };

                Button btnSave = new Button { Text = "Add to List", Location = new Point(180, 110), Size = new Size(100, 35), BackColor = Color.ForestGreen, ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
                Button btnExit = new Button { Text = "Hủy", Location = new Point(70, 110), Size = new Size(100, 35), BackColor = Color.Gray, ForeColor = Color.White, FlatStyle = FlatStyle.Flat };

                btnSave.Click += (s, e) =>
                {
                    if (num.Value > maxQty)
                    {
                        MessageBox.Show(frm, "Số lượng xuất không được lớn hơn số lượng tồn của dòng này!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    if (num.Value <= 0) return;

                    frm.DialogResult = DialogResult.OK;
                    frm.Close();
                };
                btnExit.Click += (s, e) => frm.Close();

                frm.Controls.AddRange(new Control[] { lblWarn, lblText, num, btnSave, btnExit });
                frm.AcceptButton = btnSave;

                if (frm.ShowDialog() == DialogResult.OK)
                {
                    onSuccess(num.Value);
                }
            }
        }

        private void btnXoaRow_Click(object sender, EventArgs e)
        {
            // 1. Kiểm tra xem có dòng nào đang được chọn ở Grid 2 không
            if (dgvExportQue.CurrentRow == null || dgvExportQue.CurrentRow.Index < 0)
            {
                MessageBox.Show(this.FindForm(), "Vui lòng chọn một dòng trong danh sách xuất để xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 2. Xác nhận trước khi xóa
            var result = MessageBox.Show(this.FindForm(), "Bạn có chắc chắn muốn xóa dòng này khỏi danh sách xuất không?",
                                         "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                try
                {
                    // 3. Lấy DataRow đang chọn từ Grid 2
                    DataRowView drv = (DataRowView)dgvExportQue.CurrentRow.DataBoundItem;
                    DataRow rowToDelete = drv.Row;

                    // Lấy Import_ID để xóa trong Dictionary (đây là khóa để dgvStock đổi màu)
                    string importId = rowToDelete["Import_ID"].ToString();

                    // 4. XÓA KHỎI DICTIONARY QUẢN LÝ (Quan trọng để hoàn nguyên màu)
                    if (_exportQue.ContainsKey(importId))
                    {
                        _exportQue.Remove(importId);
                    }

                    // 5. XÓA KHỎI DATATABLE CỦA GRID 2
                    dtSelected.Rows.Remove(rowToDelete);

                    // 6. CẬP NHẬT GIAO DIỆN
                    dgvExportQue.Refresh();

                    // 7. ÉP GRID 1 VẼ LẠI: Khi gọi Invalidate, sự kiện CellFormatting sẽ chạy lại.
                    // Vì ID không còn trong _exportQue, dòng ở Grid 1 sẽ tự động quay về màu trắng.
                    dgvKho.Invalidate();

                    MessageBox.Show(this.FindForm(), "Đã xóa dòng khỏi danh sách chờ xuất.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this.FindForm(), "Lỗi khi xóa dòng: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void dgvKho_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string col = dgvKho.Columns[e.ColumnIndex].Name;
            if (col == "Qty_Stock")
            {
                decimal val = e.Value != null ? Convert.ToDecimal(e.Value) : 0;
                e.CellStyle.ForeColor = val > 0 ? Color.FromArgb(40, 167, 69) : Color.FromArgb(220, 53, 69);
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
            // Lấy Import_ID của dòng đang hiển thị ở Grid 1
            var importId = dgvKho.Rows[e.RowIndex].Cells["Import_ID"].Value;

            if (importId != null && dtSelected != null)
            {
                // Kiểm tra xem Import_ID này đã có trong danh sách Grid 2 (dtSelected) chưa
                bool isExisted = dtSelected.AsEnumerable().Any(r => r.Field<int>("Import_ID") == Convert.ToInt32(importId));

                if (isExisted)
                {
                    dgvKho.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(126, 205, 50);
                    dgvKho.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Black;
                }
                else
                {
                    // Trả về màu mặc định nếu không có trong Grid 2
                    dgvKho.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.White;
                    dgvKho.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Black;
                }
            }

            //var inspector_Result = dgvKho.Rows[e.RowIndex].Cells["QC_Status"].Value;

            //if (dgvKho.Columns[e.ColumnIndex].Name == "QC_Status")
            //{
            //    string val = e.Value?.ToString() ?? "";
            //    e.CellStyle.ForeColor =
            //        val == "Pass" ? Color.FromArgb(40, 167, 69) :
            //        val == "Fail" ? Color.FromArgb(220, 53, 69) :
            //        val == "Hold" ? Color.FromArgb(255, 140, 0) :
            //                        Color.Black;
            //    e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            //}

            // 1. Định dạng cho cột "Số lượng" (Numeric)
            var qtyRules = new List<NumericRule>
            {
                new NumericRule { MinValue = 0.01m, MaxValue = decimal.MaxValue, CellColor = Color.ForestGreen }, // > 0
                new NumericRule { MinValue = decimal.MinValue, MaxValue = -0.01m, CellColor = Color.Red },       // < 0
                new NumericRule { MinValue = 0, MaxValue = 0, CellColor = Color.Red }                          // = 0
            };
            Common.Common.ApplyCustomFormatting(e, dgvKho, "Qty_Stock", null, qtyRules);

            // 2. Định dạng cho cột "Trạng thái" (String)
            var statusRules = new List<StringRule>
            {
                new StringRule { Value = "Pass", CellColor = Color.SeaGreen },
                new StringRule { Value = "Fail", CellColor = Color.Red },
                new StringRule { Value = "Hold", CellColor = Color.Orange },
                new StringRule { Value = "Pending", CellColor = Color.DimGray }
            };
            Common.Common.ApplyCustomFormatting(e, dgvKho, "QC_Status", statusRules, null);
        }

        private async void btnSearchHis_Click(object sender, EventArgs e)
        {
            if (cboProjectCheck.Items.Count == 0) return;
            var dtW = await _warehouseServies.GetHistoryExportByProject(cboProjectCheck.SelectedValue.ToString());

            dgvHis.DataSource = Common.Common.SearchDate(dtpStart.Value, dtpTo.Value, dtW, new List<string> { "Export_Date" });

            dgvHis.Columns["Export_ID"].Visible = false;
            dgvHis.Columns["Import_ID"].Visible = false;
            dgvHis.Columns["Warehouse_ID"].Visible = false;
            dgvHis.Columns["Export_To"].Visible = false;
            dgvHis.Columns["Purpose"].Visible = false;
            //dgvHis.Columns["Notes"].Visible = false;

            decimal totalQty = 0;
            decimal totalKg = 0;
            foreach (DataRow dataRow in dtW.Rows)
            {
                totalQty += Convert.ToDecimal(dataRow["Qty_Export"].ToString());
                totalKg += Convert.ToDecimal(dataRow["Weight_kg"].ToString());
            }

            lblInfoXK.Text = $"📋 Tổng: {dtW.Rows.Count} phiếu  |  SL xuất: {totalQty:N2}  |  KG xuất: {totalKg:N2}";
        }

        private void dgvKho_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow currentRow = dgvKho.Rows[e.RowIndex];
                lblStatus.Text = $"✅ {currentRow.Cells["Item_Name"].Value}  |  ID: {currentRow.Cells["ID_Code"].Value}  |  Tồn:  {currentRow.Cells["Qty_Import"].Value}   |  DA:  {cboProject.Text} ";
            }
        }

        private async void btnXuatKHO_Click(object sender, EventArgs e)
        {
            if (dgvExportQue.Rows.Count <= 0) return;
            try
            {
                var rs = await _warehouseServies.SaveExportList(dtSelected, "", AppSession.CurrentUser?.Full_Name ?? "");
                if (rs)
                {
                    ExportExcel();
                    MessageBox.Show(this.FindForm(), "Đã xuất vật tư !.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadProject();

                    _exportQue.Clear();

                    dtSelected.Rows.Clear();
                    dgvExportQue.Refresh();

                    _dtStock.Rows.Clear();
                    dgvKho.Refresh();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this.FindForm(), $"Lỗi xuất vật tư: {ex.Message}", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void ExportExcel()
        {
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

                    ReplaceCell(ws, "<<PROJECT-NAME>>", cboProject.Text ?? "");
                    ReplaceCell(ws, "<<USER>>", AppSession.CurrentUser.Full_Name ?? "");

                    int startRow = 11; // Dòng bắt đầu điền dữ liệu (Dòng có STT 1)
                    int detailCount = dgvExportQue.Rows.Count;
                    decimal totalQty = 0;

                    // 2. Chèn dòng nếu nhiều hơn 1 item để không đè lên phần chữ ký
                    if (detailCount > 1)
                    {
                        ws.InsertRow(startRow + 1, detailCount - 1, startRow);
                    }

                    // 3. Vòng lặp điền dữ liệu
                    for (int i = 0; i < detailCount; i++)
                    {
                        var row = dgvExportQue.Rows[i];
                        int currentRow = startRow + i;

                        decimal slXuat = Convert.ToDecimal(row.Cells["Qty_Export"].Value ?? 0);
                        totalQty += slXuat;

                        ws.Cells[currentRow, 1].Value = i + 1; // Cột No (A)
                        ws.Cells[currentRow, 2].Value = row.Cells["ID_Code"].Value; // Cột Code (B)
                        ws.Cells[currentRow, 3].Value = row.Cells["Item_Name"].Value; // Cột Name (C)
                        ws.Cells[currentRow, 4].Value = /*row.Cells["Ma_Phieu"].Value*/ ""; // Cột DWG No (D)
                        ws.Cells[currentRow, 5].Value = row.Cells["Size"].Value; // Cột Size (E)
                        ws.Cells[currentRow, 6].Value = row.Cells["Material"].Value; // Cột Grade (F)
                        ws.Cells[currentRow, 7].Value = slXuat; // Cột Q'ty (G)
                        ws.Cells[currentRow, 8].Value = row.Cells["UNIT"].Value; // Cột Unit (H)
                        ws.Cells[currentRow, 9].Value = row.Cells["QC_Code"].Value; // Cột ID_Code (I)
                        ws.Cells[currentRow, 10].Value = row.Cells["Notes"].Value; // Cột ID_Code (I)
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
        }

        private void ReplaceCell(OfficeOpenXml.ExcelWorksheet ws, string placeholder, string value)
        { for (int r = 1; r <= ws.Dimension.End.Row; r++) for (int c = 1; c <= ws.Dimension.End.Column; c++) if (ws.Cells[r, c].Value?.ToString() == placeholder) ws.Cells[r, c].Value = value; }


        private void btnClear_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show(
                $"❓Bạn đang tạo phiếu xuất kho cho dự án {cboProject.Text.Trim().ToUpper()}\n\nBạn có muốn hủy bỏ phiếu hiện tại không?",
                "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                _exportQue.Clear();
                dtSelected.Rows.Clear();
                dgvExportQue.Refresh();
                _dtStock.Rows.Clear();
                dgvKho.Refresh();
            }
        }


        private void dgvHis_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //if (e.RowIndex < 0) return;

            //DataGridViewRow row = dgvHis.Rows[e.RowIndex];

            //// 1. LẤY GIÁ TRỊ SỐ LƯỢNG (Giả sử cột tên là Qty_Import)
            //decimal qty = 0;
            //if (row.Cells["Qty_Export"].Value != null)
            //{
            //    decimal.TryParse(row.Cells["Qty_Export"].Value.ToString(), out qty);
            //}

            //// 3. THIẾT LẬP MÀU CẢNH BÁO THEO SỐ LƯỢNG TỒN
            //if (qty == 0)
            //{
            //    // TRẠNG THÁI HẾT HÀNG: Nền hồng nhạt, chữ đỏ đậm
            //    row.DefaultCellStyle.BackColor = Color.FromArgb(255, 235, 235);
            //    row.DefaultCellStyle.ForeColor = Color.FromArgb(192, 57, 43); // Đỏ Alizarin

            //    // Thêm font in nghiêng cho hàng hết
            //    row.DefaultCellStyle.Font = new Font(dgvHis.Font, FontStyle.Italic);
            //}
            //else if (qty > 0 && qty < 5)
            //{
            //    // TRẠNG THÁI SẮP HẾT (Dưới 5): Nền vàng nhạt, chữ nâu/cam đậm
            //    row.DefaultCellStyle.BackColor = Color.FromArgb(255, 250, 205); // Lemon Chiffon
            //    row.DefaultCellStyle.ForeColor = Color.FromArgb(211, 84, 0);   // Cam đậm (Pumpkin)

            //    row.DefaultCellStyle.Font = new Font(dgvHis.Font, FontStyle.Regular);
            //}
            //else
            //{
            //    // TRẠNG THÁI BÌNH THƯỜNG
            //    row.DefaultCellStyle.BackColor = Color.White;
            //    row.DefaultCellStyle.ForeColor = Color.Black;
            //    row.DefaultCellStyle.Font = new Font(dgvHis.Font, FontStyle.Regular);
            //}
        }

        private void cboProject_Validating(object sender, CancelEventArgs e)
        {
            Common.Common.AutoCompleteComboboxValidating(sender as ComboBox, e);
        }

        private void cboProjectCheck_Validating(object sender, CancelEventArgs e)
        {
            Common.Common.AutoCompleteComboboxValidating(sender as ComboBox, e);
        }

        private void dgvExportQue_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string col = dgvExportQue.Columns[e.ColumnIndex].Name;
            if (col == "Qty_Export")
            {
                decimal val = e.Value != null ? Convert.ToDecimal(e.Value) : 0;
                e.CellStyle.ForeColor = val > 0 ? Color.FromArgb(40, 167, 69) : Color.FromArgb(220, 53, 69);
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
            //// Lấy Import_ID của dòng đang hiển thị ở Grid 1
            //var importId = dgvKho.Rows[e.RowIndex].Cells["Import_ID"].Value;

            //if (importId != null && dtSelected != null)
            //{
            //    // Kiểm tra xem Import_ID này đã có trong danh sách Grid 2 (dtSelected) chưa
            //    bool isExisted = dtSelected.AsEnumerable().Any(r => r.Field<int>("Import_ID") == Convert.ToInt32(importId));

            //    if (isExisted)
            //    {
            //        dgvKho.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(126, 205, 50);
            //        dgvKho.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Black;
            //    }
            //    else
            //    {
            //        // Trả về màu mặc định nếu không có trong Grid 2
            //        dgvKho.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.White;
            //        dgvKho.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Black;
            //    }
            //}
        }

        private void btnCancelSer_Click(object sender, EventArgs e)
        {
            _filterHelper.ClearAllFilters();
        }
    }
}
