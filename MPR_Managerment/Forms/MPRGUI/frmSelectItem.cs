using MPR_Managerment.Common;
using MPR_Managerment.Models;
using MPR_Managerment.Services;
using Syncfusion.XlsIO.Parser.Biff_Records;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.Contracts;
using System.Drawing;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MPR_Managerment.Forms.MPRGUI
{
    public partial class frmSelectItem : Form
    {
        private const string SearchPlaceholder = "Mã/tên vật tư...";
        private readonly WarehouseService _warehouseService = new WarehouseService();
        private readonly ProductServices _productServices = new ProductServices();
        private DataTable _dtItems = new DataTable();
        public List<ProductModel> selectedItems { get; set; } = new List<ProductModel>();

        public List<ProductModel> selectedProducts { get; set; } = new List<ProductModel>();

        public ProductModel ProductModel { get; set; }
        public int CheckedQuanty = 0;

        public frmSelectItem()
        {
            InitializeComponent();
            // Gắn AI Trợ lý — hiển thị nút floating + chat panel
            frmAIChat.Attach(this);
            SetSearchPlaceholder();
            Common.Common.CreateButtonRefresh(btnRefresh);
            Common.Common.CreateButtonSave(btnSelect);
            Common.Common.CreateButtonCancel(btnCancels, "");
            Common.Common.CreateButtonSearch(btnSearch, "");
            Common.Common.CreateButtonDelete(btnDelete, "🗑 Bỏ chọn");
            txtSearch.PlaceholderText = "Mã/tên vật tư...";
        }


        private async void frmSelectItem_Load_1(object sender, EventArgs e)
        {
            await LoadItems();
            InitGridItems();
        }

        private void InitGridItems()
        {
            Common.Common.CreateDataGridView_Hide_RowHeader(dgvItems);

            dgvItems.CellContentClick += DgvItems_CellContentClick;

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
                    int Id = Convert.ToInt32(dgvItems.Rows[e.RowIndex].Cells["Id"]?.Value?.ToString()?.Trim() ?? "0");
                    if (isChecked)
                    {
                        CheckedQuanty++; // Người dùng check -> +1
                        if (selectedProducts.Any(i => i.Id == Id)) return;
                        selectedProducts.Add(new ProductModel
                        {
                            Id = Id,
                            Name = dgvItems.Rows[e.RowIndex].Cells["name"].Value?.ToString() ?? "",
                            Des2 = dgvItems.Rows[e.RowIndex].Cells["des_2"].Value?.ToString() ?? "",
                            Code = dgvItems.Rows[e.RowIndex].Cells["code"].Value?.ToString() ?? "",
                            ProdMaterialCode = dgvItems.Rows[e.RowIndex].Cells["prod_material_code"].Value?.ToString() ?? "",
                            A_Thickness = dgvItems.Rows[e.RowIndex].Cells["a_thinkness"].Value?.ToString() ?? "",
                            B_Depth = dgvItems.Rows[e.RowIndex].Cells["b_depth"].Value?.ToString() ?? "",
                            C_Width = dgvItems.Rows[e.RowIndex].Cells["c_witdth"].Value?.ToString() ?? "",
                            D_Web = dgvItems.Rows[e.RowIndex].Cells["d_web"].Value?.ToString() ?? "",
                            E_Flag = dgvItems.Rows[e.RowIndex].Cells["e_flag"].Value?.ToString() ?? "",
                            F_Length = dgvItems.Rows[e.RowIndex].Cells["f_length"].Value?.ToString() ?? "",
                            G_Weight = dgvItems.Rows[e.RowIndex].Cells["g_weight"].Value?.ToString() ?? "",
                        });
                    }
                    else
                    {
                        // Chỉ trừ nếu biến đang lớn hơn 0 để tránh số âm ngoài ý muốn
                        if (CheckedQuanty > 0) CheckedQuanty--;
                        int id = Convert.ToInt32(dgvItems.Rows[e.RowIndex].Cells["Id"]?.Value?.ToString()?.Trim() ?? "0");
                        selectedProducts.RemoveAll(p => p.Id == id);
                    }

                    // Hiển thị lên giao diện
                    lblStatus.Text = $"Đã chọn: {selectedProducts.Count} vật tư";
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

        private void SyncDataGridViewWithList(DataGridView dgv, List<ProductModel> selectedProducts)
        {
            // 1. Kiểm tra điều kiện đầu vào
            if (dgv.Rows.Count == 0 || selectedProducts == null) return;

            // 2. Tối ưu hiệu suất: Chuyển List ID sang HashSet để tìm kiếm nhanh O(1)
            // Thay vì duyệt List nhiều lần, HashSet giúp kiểm tra sự tồn tại tức thì.
            var selectedIds = new HashSet<int>(selectedProducts.Select(p => p.Id));

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
                    if (row.Cells["Id"].Value != null && int.TryParse(row.Cells["Id"].Value.ToString(), out int rowId))
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

        private async Task LoadItems()
        {
            //_dtItems = await _productServices.GetProducts();
            //BindStockGrid(_dtItems);
            _dtItems = await _productServices.GetProducts();

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
            //// Xóa cột cũ và dữ liệu cũ trước khi nạp lại (tránh trùng lặp cột khi gọi hàm nhiều lần)
            //dgvItems.Columns.Clear();
            //dgvItems.DataSource = null;

            //// Thêm cột Checkbox trước
            //DataGridViewCheckBoxColumn checkColumn = new DataGridViewCheckBoxColumn();
            //checkColumn.Name = "Chon";
            //checkColumn.HeaderText = "Chọn";
            //checkColumn.ReadOnly = false; // Cho phép tương tác
            //dgvItems.Columns.Add(checkColumn);
            ////dgvItems.Columns["Chon"].Width = 300;

            //// Thêm cột Checkbox trước
            //DataGridViewImageColumn imgCol = new DataGridViewImageColumn();
            //imgCol.HeaderText = "Hình ảnh";
            //imgCol.Name = "imgColumn";
            //imgCol.ReadOnly = false;
            //dgvItems.Columns.Add(imgCol);

            ////dgvItems.Columns["imgColumn"].Width = 35;

            //// Gán DataSource
            //dgvItems.DataSource = models.AsEnumerable().Select(row => new
            //{
            //    Id = row.Field<int>("Id"),
            //    Name = row.Field<string>("name"),
            //    Desciption = row.Field<string>("des_2"),
            //    Code = row.Field<string>("code"),
            //    MaterialCode = row.Field<string>("prod_material_code"),
            //    Thinkness = row.Field<string>("a_thinkness"),
            //    Depth = row.Field<string>("b_depth"),
            //    Width = row.Field<string>("c_witdth"),
            //    Web = row.Field<string>("d_web"),
            //    Flag = row.Field<string>("e_flag"),
            //    Lenght = row.Field<string>("f_length"),
            //    Weight = row.Field<string>("g_weight"),
            //}).ToList();

            //foreach (DataGridViewColumn column in dgvItems.Columns)
            //{
            //    // Check if the column name or header contains "id" (case-insensitive)
            //    if (column.Name.IndexOf("id", StringComparison.OrdinalIgnoreCase) >= 0)
            //    {
            //        column.Visible = false;
            //    }
            //}

            ////// Tính toán tổng số lượng
            ////decimal tQ = 0, tW = 0;
            ////if (stocks != null)
            ////{
            ////    foreach (var s in stocks)
            ////    {
            ////        tQ += s.Qty_Stock;
            ////        tW += s.Weight_Stock;
            ////    }
            ////    if (lblStockTotal != null) lblStockTotal.Text = $"{stocks.Count} mục";
            ////    if (lblStockQty != null) lblStockQty.Text = tQ.ToString("N2");
            ////    if (lblStockWeight != null) lblStockWeight.Text = tW.ToString("N2") + " kg";
            ////}

            //// Thủ thuật nhỏ: Để CheckBox phản hồi click chuột ngay lập tức (không cần click 2 lần)
            //dgvItems.EditMode = DataGridViewEditMode.EditOnEnter;

            // KHÔNG xóa toàn bộ cột nếu đã có cột "Chon"
            if (dgvItems.Columns.Count == 0)
            {
                DataGridViewCheckBoxColumn checkColumn = new DataGridViewCheckBoxColumn();
                checkColumn.Name = "Chon";
                checkColumn.HeaderText = "Chọn";
                checkColumn.DataPropertyName = "Chon"; // Map trực tiếp với cột trong DataTable
                dgvItems.Columns.Add(checkColumn);

                //DataGridViewImageColumn imgCol = new DataGridViewImageColumn();
                //imgCol.HeaderText = "Hình ảnh";
                //imgCol.Name = "imgColumn";
                //dgvItems.Columns.Add(imgCol);
            }

            // Gán DataSource là DataTable/DataView chứa cột 'Chon'
            dgvItems.DataSource = models;

            SyncDataGridViewWithList(dgvItems, selectedProducts);

            // Ẩn các cột ID và cấu hình hiển thị
            foreach (DataGridViewColumn column in dgvItems.Columns)
            {
                if (column.Name.IndexOf("id", StringComparison.OrdinalIgnoreCase) >= 0)
                    column.Visible = false;

                // Cấu hình các cột hiển thị đẹp hơn (tùy chọn)
                if (column.Name == "Chon") column.DisplayIndex = 0;
            }
            dgvItems.Columns["picture_link"].Visible = false;
            dgvItems.Columns["picture"].Visible = false;

            dgvItems.EditMode = DataGridViewEditMode.EditOnEnter;
            lblStatus.Text = $"Đã chọn: {selectedProducts.Count} vật tư";
        }

        private void DgvItems_CellContentClick(object? sender, DataGridViewCellEventArgs e)
        {
            //if (!Common.Common.IsDataGridViewValid(dgvItems)) return;
            if (dgvItems.Rows.Count <= 0) return;

            int rsl = dgvItems.CurrentRow.Index;

            try
            {
                ProductModel = new ProductModel()
                {
                    Id = Convert.ToInt32(dgvItems.CurrentRow.Cells["ID"].Value.ToString()),
                    Name = !string.IsNullOrEmpty(dgvItems.CurrentRow.Cells["Name"].Value.ToString().Trim()) ? dgvItems.CurrentRow.Cells["Name"].Value.ToString().Trim() : "",
                    Des2 = !string.IsNullOrEmpty(dgvItems.CurrentRow.Cells["Desciption"].Value.ToString().Trim()) ? dgvItems.CurrentRow.Cells["Desciption"].Value.ToString().Trim() : "",
                    ProdMaterialCode = !string.IsNullOrEmpty(dgvItems.CurrentRow.Cells["MaterialCode"].Value.ToString().Trim()) ? dgvItems.CurrentRow.Cells["MaterialCode"].Value.ToString().Trim() : "",
                    Code = !string.IsNullOrEmpty(dgvItems.CurrentRow.Cells["Code"].Value.ToString().Trim()) ? dgvItems.CurrentRow.Cells["Code"].Value.ToString().Trim() : "",
                };
            }
            catch (Exception)
            {
                Debug.WriteLine("Lỗi khi lấy dữ liệu từ DataGridView. Vui lòng kiểm tra lại cấu trúc cột và dữ liệu.");
            }
        }

        private void SetSearchPlaceholder()
        {
            if (string.IsNullOrWhiteSpace(txtSearch.Text))
            {
                txtSearch.ForeColor = SystemColors.GrayText;
                txtSearch.PlaceholderText = SearchPlaceholder;
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            Search(txtSearch.Text.Trim());
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
            //if (txtSearch.Text.Length == 0)
            //{
            //    //dgvProds.DataSource = this.dtProds;
            //    dgvItems.Refresh();
            //}

            //var lstProperty = new List<string>()
            //{
            //    "Name", "des_2", "code", "prod_material_code", "a_thinkness", "b_depth", "c_witdth", "d_web", "e_flag", "f_length", "g_weight"
            //};

            //DataView dv = Common.Common.Search(txtSearch.Text, _dtItems, lstProperty);

            //BindStockGrid(dv.ToTable());
            // Sử dụng RowFilter của DefaultView để ẩn/hiện dòng mà không làm mất dữ liệu
            var lstProperty = new List<string>()
            {
                "name", "des_2", "code", "prod_material_code" // Tên cột phải khớp chính xác với DataTable
            };

            DataView dv = Common.Common.Search(text, _dtItems, lstProperty);

            // Thay vì tạo bảng mới, chỉ cần cập nhật DataSource bằng View đã lọc
            dgvItems.DataSource = dv;

            // Cập nhật lại số lượng đã chọn sau khi search
            //CalculateCheckedData();
            SyncDataGridViewWithList(dgvItems, selectedProducts);
            lblStatus.Text = $"Đã chọn: {selectedProducts.Count} vật tư";
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            dgvItems.EndEdit();
            selectedItems = new List<ProductModel>();
            //foreach (DataGridViewRow row in dgvItems.Rows)
            //{
            //    if (row.IsNewRow) continue;
            //    bool isChecked = Convert.ToBoolean(row.Cells["Chon"].Value);
            //    if (isChecked)
            //    {
            //        selectedItems.Add(new ProductModel
            //        {
            //            Id = Convert.ToInt32(row.Cells["Id"].Value.ToString().Trim()),
            //            Name = row.Cells["name"].Value?.ToString(),
            //            Des2 = row.Cells["des_2"].Value?.ToString(),
            //            Code = row.Cells["code"].Value?.ToString(),
            //            ProdMaterialCode = row.Cells["prod_material_code"].Value?.ToString(),
            //            A_Thickness = row.Cells["a_thinkness"].Value?.ToString(),
            //            B_Depth = row.Cells["b_depth"].Value?.ToString(),
            //            C_Width = row.Cells["c_witdth"].Value?.ToString(),
            //            D_Web = row.Cells["d_web"].Value?.ToString(),
            //            E_Flag = row.Cells["e_flag"].Value?.ToString(),
            //            F_Length = row.Cells["f_length"].Value?.ToString(),
            //            G_Weight = row.Cells["g_weight"].Value?.ToString(),
            //        });
            //    }
            //}
            selectedItems = selectedProducts;
            this.Close();
        }

        private void btnCancels_Click(object sender, EventArgs e)
        {
            this.Close();
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
            selectedProducts.Clear();
            SyncDataGridViewWithList(dgvItems, selectedProducts);
            lblStatus.Text = $"Đã chọn: {selectedProducts.Count} vật tư";
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            frmShowImage frm = new frmShowImage(@"D:\RAC\Image");
            frm.ShowDialog();
        }
    }
}
