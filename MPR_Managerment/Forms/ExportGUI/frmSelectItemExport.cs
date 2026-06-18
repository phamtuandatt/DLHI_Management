using MPR_Managerment.Models;
using MPR_Managerment.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MPR_Managerment.Forms.ExportGUI
{
    public partial class frmSelectItemExport : Form
    {
        private DataTable _dtItems = new DataTable();
        private WarehouseService _warehouseService = new WarehouseService();
        private bool _isAdd = false;

        public List<WarehouseImport> selectedList { get; set; } = new List<WarehouseImport>();
        public int CheckedQuanty = 0;
        public bool isCancel = false;


        public frmSelectItemExport()
        {
            InitializeComponent();

            Common.Common.CreateButtonAdd(btnPreImage, "📸 Hình ảnh");
            Common.Common.CreateButtonSave(btnSelect, null);
            Common.Common.CreateButtonCancel(btnCancels, "");
            Common.Common.CreateButtonSearch(btnSearch, "🔍 Tìm");
            Common.Common.CreateButtonDelete(btnDelete, "🗑 Bỏ chọn");
            txtSearch.PlaceholderText = "Mã/tên vật tư...";

        }

        public frmSelectItemExport(bool isAdd)
        {
            InitializeComponent();
            _isAdd = isAdd;
            Common.Common.CreateButtonAdd(btnPreImage, "📸 Hình ảnh");
            Common.Common.CreateButtonSave(btnSelect, null);
            Common.Common.CreateButtonCancel(btnCancels, "");
            Common.Common.CreateButtonSearch(btnSearch, "🔍 Tìm");
            Common.Common.CreateButtonDelete(btnDelete, "🗑 Bỏ chọn");
            txtSearch.PlaceholderText = "Mã/tên vật tư...";
        }

        private async void frmSelectItemExport_Load(object sender, EventArgs e)
        {
            await LoadItems();
            InitGridItems();
        }

        private async Task LoadItems()
        {
            //_dtItems = await _productServices.GetProducts();
            //BindStockGrid(_dtItems);
            _dtItems = await _warehouseService.GetWarehouse_ForExport_V2();

            // Thêm cột logic 'Chon' vào DataTable nếu chưa có
            if (!_dtItems.Columns.Contains("Chon"))
            {
                DataColumn col = new DataColumn("Chon", typeof(bool));
                col.DefaultValue = false;
                _dtItems.Columns.Add(col);
            }

            BindStockGrid(_dtItems);
        }

        private void BindStockGrid(DataTable models)
        {
            // KHÔNG xóa toàn bộ cột nếu đã có cột "Chon"
            if (dgvItems.Columns.Count == 0)
            {
                DataGridViewCheckBoxColumn checkColumn = new DataGridViewCheckBoxColumn();
                checkColumn.Name = "Chon";
                checkColumn.HeaderText = "Chọn";
                checkColumn.DataPropertyName = "Chon"; // Map trực tiếp với cột trong DataTable
                dgvItems.Columns.Add(checkColumn);
            }

            // Gán DataSource là DataTable/DataView chứa cột 'Chon'
            dgvItems.DataSource = models;

            SyncDataGridViewWithList(dgvItems, selectedList);

            // Ẩn các cột ID và cấu hình hiển thị
            foreach (DataGridViewColumn column in dgvItems.Columns)
            {
                if (column.Name.IndexOf("id", StringComparison.OrdinalIgnoreCase) >= 0)
                    column.Visible = false;

                // Cấu hình các cột hiển thị đẹp hơn (tùy chọn)
                if (column.Name == "Chon") column.DisplayIndex = 0;
            }

            dgvItems.EditMode = DataGridViewEditMode.EditOnEnter;
            lblStatus.Text = $"Đã chọn: {selectedList.Count} vật tư";
        }

        private void InitGridItems()
        {
            Common.Common.CreateDataGridView_Hide_RowHeader(dgvItems);

            dgvItems.CellContentClick += DgvItems_CellContentClick; ;

            dgvItems.CellFormatting += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                //string col = dgvItems.Columns[e.ColumnIndex].Name.ToLower();
                //if (col.Contains("Id".ToLower()))
                //{
                //    //decimal val = e.Value != null ? Convert.ToDecimal(e.Value) : 0;
                //    //e.CellStyle.ForeColor = Color.FromArgb(40, 167, 69);
                //    //e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                //    dgvItems.Columns[col].Visible = false;
                //}

            };

            dgvItems.EditingControlShowing += (s, e) =>
            {
                //e.Control.KeyPress -= new KeyPressEventHandler(Common.Common.Column_KeyPress_Digital);
                //if (dgvItems.CurrentCell.ColumnIndex == dgvItems.Columns["Qty"].Index)
                //{
                //    TextBox tb = e.Control as TextBox;
                //    if (tb != null)
                //    {
                //        tb.KeyPress += new KeyPressEventHandler(Common.Common.Column_KeyPress_Digital);
                //    }
                //}
            };

            dgvItems.CellEndEdit += (s, e) =>
            {
                //// Chỉ kiểm tra nếu cột đang sửa là "SL_Xuat"
                //if (dgvItems.Columns[e.ColumnIndex].Name == "Qty")
                //{
                //    var row = dgvItems.Rows[e.RowIndex];

                //    // Lấy giá trị nhập vào và giá trị tồn
                //    decimal slNhap = 0;

                //    // Ép kiểu an toàn (sử dụng decimal.TryParse để tránh lỗi nhập chữ)
                //    decimal.TryParse(row.Cells["Qty"].Value?.ToString() ?? "0", out slNhap);

                //    if (slNhap == 0)
                //    {
                //        // Gán lại giá trị Xuất bằng giá trị Tồn
                //        row.Cells["Qty"].Value = 1;
                //    }
                //}
            };

            dgvItems.CellValueChanged += (s, e) =>
            {
                if (dgvItems.Columns[e.ColumnIndex].Name == "Chon" && e.RowIndex >= 0)
                {
                    bool isChecked = Convert.ToBoolean(dgvItems.Rows[e.RowIndex].Cells["Chon"].Value);
                    int Id = Convert.ToInt32(dgvItems.Rows[e.RowIndex].Cells["Import_ID"]?.Value?.ToString()?.Trim() ?? "0");
                    if (isChecked)
                    {
                        CheckedQuanty++; // Người dùng check -> +1
                        if (selectedList.Any(i => i.Import_ID == Id)) return;
                        selectedList.Add(new WarehouseImport
                        {
                            Import_ID = Id,
                            ID_Code = dgvItems.Rows[e.RowIndex].Cells["ID_Code"].Value?.ToString() ?? "",
                            Item_Name = dgvItems.Rows[e.RowIndex].Cells["Item_Name"].Value?.ToString() ?? "",
                            Material = dgvItems.Rows[e.RowIndex].Cells["Material"].Value?.ToString() ?? "",
                            Size = dgvItems.Rows[e.RowIndex].Cells["Size"].Value?.ToString() ?? "",
                            UNIT = dgvItems.Rows[e.RowIndex].Cells["UNIT"].Value?.ToString() ?? "",

                        });
                    }
                    else
                    {
                        // Chỉ trừ nếu biến đang lớn hơn 0 để tránh số âm ngoài ý muốn
                        if (CheckedQuanty > 0) CheckedQuanty--;
                        int id = Convert.ToInt32(dgvItems.Rows[e.RowIndex].Cells["Import_ID"]?.Value?.ToString()?.Trim() ?? "0");
                        selectedList.RemoveAll(p => p.Import_ID == id);
                    }

                    // Hiển thị lên giao diện
                    lblStatus.Text = $"Đã chọn: {selectedList.Count} vật tư";
                }

            };

            dgvItems.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (dgvItems.IsCurrentCellDirty)
                {
                    // This commits the edit immediately instead of waiting for focus to change
                    dgvItems.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            };
        }

        private void DgvItems_CellContentClick(object? sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            Search(txtSearch.Text.Trim());
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {

        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            dgvItems.EndEdit();

            if (!_isAdd && selectedList.Count > 1)
            {
                MessageBox.Show("Hãy chọn vật tư muốn cập nhật", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                isCancel = true;
                this.Close();
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            // Ensure you end any active edits to commit pending changes
            dgvItems.EndEdit();

            foreach (DataGridViewRow row in dgvItems.Rows)
            {
                // Set the cell value to false (unchecked)
                row.Cells["Chon"].Value = false;
            }
            selectedList.Clear();
            SyncDataGridViewWithList(dgvItems, selectedList);
            lblStatus.Text = $"Đã chọn: {selectedList.Count} vật tư";
        }

        private void SyncDataGridViewWithList(DataGridView dgv, List<WarehouseImport> selectedProducts)
        {
            // 1. Kiểm tra điều kiện đầu vào
            if (dgv.Rows.Count == 0 || selectedProducts == null) return;

            // 2. Tối ưu hiệu suất: Chuyển List ID sang HashSet để tìm kiếm nhanh O(1)
            // Thay vì duyệt List nhiều lần, HashSet giúp kiểm tra sự tồn tại tức thì.
            var selectedIds = new HashSet<int>(selectedProducts.Select(p => p.Import_ID));

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

        private void btnPreImage_Click(object sender, EventArgs e)
        {
            frmShowImage frm = new frmShowImage(@"\\Dlhivina\SHARE\Old\Stationery");
            frm.ShowDialog();
        }

        private void btnCancels_Click(object sender, EventArgs e)
        {
            dgvItems.EndEdit();
            isCancel = false;
            this.Close();
        }

        private void txtSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search(txtSearch.Text.Trim());
            }
        }

        private void Search(string text)
        {
            var lstProperty = new List<string>()
            {
                "ID_Code", "Item_Name", "Size", "UNIT"
            };

            DataView dv = Common.Common.Search(text, _dtItems, lstProperty);

            // Thay vì tạo bảng mới, chỉ cần cập nhật DataSource bằng View đã lọc
            dgvItems.DataSource = dv;

            SyncDataGridViewWithList(dgvItems, selectedList);
            lblStatus.Text = $"Đã chọn: {selectedList.Count} vật tư";
        }
    }
}
