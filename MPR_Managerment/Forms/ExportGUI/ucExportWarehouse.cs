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

        //private DataGridView dgvKho;

        private bool _isLoaded = false;

        public ucExportWarehouse()
        {
            InitializeComponent();
            LoadProject();
            InitGridSelected();


            FormartGrid(dgvKho, Color.FromArgb(0, 120, 212));
            FormartGrid(dgvExportQue, Color.FromArgb(255, 140, 0));
            FormartGrid(dgvHis, Color.FromArgb(0, 120, 212));
        }

        private void FormartGrid(DataGridView dataGridView, Color colorHeader)
        {
            dataGridView.ReadOnly = true;
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
                dgvKho.DataSource = _dtStock;
                dgvKho.Columns["Import_ID"].Visible = false;
                dgvKho.Columns["PO_ID"].Visible = false;
                dgvKho.Columns["PO_Detail_ID"].Visible = false;
                dgvKho.Columns["RIR_ID"].Visible = false;
                dgvKho.Columns["Item_ID"].Visible = false;
                dgvKho.Columns["Warehouse_ID"].Visible = false;
                dgvKho.Columns["Item_Code"].Visible = false;

                dgvKho.Columns["Location"].Visible = false;
                dgvKho.Columns["Notes"].Visible = false;
                dgvKho.Columns["MTRno"].Visible = false;
                dgvKho.Columns["Heatno"].Visible = false;
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
                    decimal maxQtyInStock = Convert.ToDecimal(sourceRow["Qty_Import"]);

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
                    MessageBox.Show("Lỗi xuất kho: " + ex.Message + " - " + value, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                newRow["Weight_kg"] = sourceRow["Weight_kg"];
                newRow["Project_Code"] = sourceRow["Project_Code"];
                newRow["WorkorderNo"] = sourceRow["WorkorderNo"] ?? null;
                newRow["Warehouse_ID"] = sourceRow["Warehouse_ID"] ?? null;

                // Gán số lượng xuất từ Popup
                newRow["Qty_Export"] = qty;

                // Các thông tin mặc định khác (nếu có)
                newRow["Export_Date"] = DateTime.Now;
                newRow["Created_Date"] = DateTime.Now;

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

            dgvExportQue.DataSource = dtSelected;

            // Cấu hình tiêu đề cột cho dễ nhìn
            dgvExportQue.Columns["Item_Name"].HeaderText = "Tên vật tư";
            dgvExportQue.Columns["Qty_Export"].HeaderText = "SL Xuất";
            dgvExportQue.Columns["ID_Code"].HeaderText = "Mã ITEM";

            dgvExportQue.Columns["Export_ID"].Visible = false;
            dgvExportQue.Columns["Export_No"].Visible = false;
            dgvExportQue.Columns["Export_To"].Visible = false;
            dgvExportQue.Columns["Purpose"].Visible = false;
            dgvExportQue.Columns["Notes"].Visible = false;
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
                        MessageBox.Show("Số lượng xuất không được lớn hơn số lượng tồn của dòng này!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                MessageBox.Show("Vui lòng chọn một dòng trong danh sách xuất để xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 2. Xác nhận trước khi xóa
            var result = MessageBox.Show("Bạn có chắc chắn muốn xóa dòng này khỏi danh sách xuất không?",
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

                    MessageBox.Show("Đã xóa dòng khỏi danh sách chờ xuất.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi xóa dòng: " + ex.Message);
                }
            }
        }

        private void dgvKho_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string col = dgvKho.Columns[e.ColumnIndex].Name;
            if (col == "Qty_Import")
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
        }

        private async void btnSearchHis_Click(object sender, EventArgs e)
        {
            if (cboProjectCheck.Items.Count == 0) return;
            var dtW = await _warehouseServies.GetHistoryExportByProject(cboProjectCheck.SelectedValue.ToString());
            dgvHis.DataSource = dtW;
            dgvHis.Columns["Export_ID"].Visible = false;
            dgvHis.Columns["Import_ID"].Visible = false;
            dgvHis.Columns["Warehouse_ID"].Visible = false;
            dgvHis.Columns["Export_To"].Visible = false;
            dgvHis.Columns["Purpose"].Visible = false;
            dgvHis.Columns["Notes"].Visible = false;

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
                var rs = await _warehouseServies.SaveExportList(dtSelected, "", "Admin");
                if (rs)
                {
                    MessageBox.Show("Đã xuất vật tư !.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi xuất vật tư: {ex.Message}", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void btnClear_Click(object sender, EventArgs e)
        {

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
    }
}
