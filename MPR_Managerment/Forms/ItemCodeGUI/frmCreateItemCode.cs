using Microsoft.Data.SqlClient;
using MPR_Managerment.Common;
using MPR_Managerment.Models;
using MPR_Managerment.Services;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MPR_Managerment.Forms.ItemCodeGUI
{
    public partial class frmCreateItemCode : Form
    {
        private ProductServices _productServices = new ProductServices();
        private WarehouseService _warehouseServices = new WarehouseService();
        private bool _isLoadedMaterialCate = false;
        private string itemNumberOfMaterial = "";
        private bool _isStandardLoaded = false;
        private bool _isClickGrid = false;
        private DataTable dtOrgins = new DataTable();
        private DataTable dtItemCode = new DataTable();

        public string itemCode { get; set; } = string.Empty;
        public string itemDetailId { get; set; } = string.Empty;
        public string itemDetailNumber { get; set; } = string.Empty;
        public bool isUseCodeAvailable { get; set; } = false;

        public frmCreateItemCode(string title)
        {
            InitializeComponent();
            // Đăng ký sự kiện Shown thay vì gọi trực tiếp ở đây
            //this.Shown += (s, e) => {
            //    SetTabOrder();
            //};
            this.Text = $"Tên vật tư: {title}";
            //this.Size = new Size(450, 650); // Tăng nhẹ chiều cao để cân đối khoảng cách
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            txtSearch.Font = new Font("Segoe UI", 9);
            txtSearch.PlaceholderText = "Nhập tên vật tư...";
        }

        private async void frmCreateItemCode_Load(object sender, EventArgs e)
        {
            await LoadMaterialCate();
            if (_isLoadedMaterialCate)
            {
                await LoadMaterialByCate(Convert.ToInt32(cboMaterialCate.SelectedValue.ToString()));
            }
            await LoadOriginals();
            await LoadStandards();
        }

        private void SetTabOrder()
        {
            // 1. Thiết lập thứ tự từ trên xuống dưới theo yêu cầu của bạn
            cboMaterialCate.TabIndex = 0;
            cboMaterial.TabIndex = 1;
            cboOriginal.TabIndex = 2;
            cboStandard.TabIndex = 3;
            txtCode.TabIndex = 4;
            btnGenerate.TabIndex = 5;

            // 2. Các nút phụ ở Footer (nếu có) nên để số lớn hơn
            btnSave.TabIndex = 6;
            btnCancel.TabIndex = 7;

            // 3. Đảm bảo thuộc tính TabStop là true (mặc định là true)
            // Nếu TabStop = false, phím Tab sẽ nhảy qua control đó.
            cboMaterialCate.TabStop = true;
            txtCode.TabStop = true;

            // 4. Focus vào control đầu tiên khi mở Form
            this.ActiveControl = cboMaterialCate;
        }

        private async Task LoadMaterialCate()
        {
            DataTable dtCates = await _productServices.GetMaterialCates();
            cboMaterialCate.DisplayMember = "cat_name";
            cboMaterialCate.ValueMember = "cat_id";
            cboMaterialCate.DataSource = dtCates;
            if (cboMaterialCate.Items.Count > 0)
            {
                _isLoadedMaterialCate = true;
                //cboMaterial.SelectedIndex = 0;
            }
        }

        private async Task LoadOriginals()
        {
            // 1. Lấy dữ liệu từ Service
            DataTable dtMaterials = await _productServices.GetOriginals();

            // 2. Khởi tạo DataTable mới cho ComboBox và định nghĩa cấu trúc ngay từ đầu
            DataTable dtCbo = new DataTable();
            dtCbo.Columns.Add("ID", typeof(int));
            dtCbo.Columns.Add("NAME", typeof(string));

            // 3. Duyệt và đổ dữ liệu
            foreach (DataRow dr in dtMaterials.Rows)
            {
                DataRow r = dtCbo.NewRow();

                // Nên dùng tên cột từ Database thay vì chỉ số 0, 1 để tránh nhầm lẫn
                // Giả sử dr[0] là ID, dr[1] là Code, dr[2] là Name
                r["ID"] = dr[0];
                r["NAME"] = $"{dr[1]}-{dr[2]}"; // Ghép mã và tên hiển thị

                dtCbo.Rows.Add(r);
            }

            // 4. Gán nguồn dữ liệu cho ComboBox
            // Lưu ý: Phải gán DisplayMember và ValueMember TRƯỚC khi gán DataSource
            cboOriginal.DataSource = null; // Clear dữ liệu cũ nếu có
            cboOriginal.DisplayMember = "NAME";
            cboOriginal.ValueMember = "ID";
            cboOriginal.DataSource = dtCbo;
            dtOrgins = dtCbo.Copy();
        }

        private async Task LoadStandards()
        {
            // 1. Lấy dữ liệu từ Service
            DataTable dtMaterials = await _productServices.GetStandards();

            // 2. Khởi tạo DataTable mới cho ComboBox và định nghĩa cấu trúc ngay từ đầu
            DataTable dtCbo = new DataTable();
            dtCbo.Columns.Add("ID", typeof(int));
            dtCbo.Columns.Add("NAME", typeof(string));

            // 3. Duyệt và đổ dữ liệu
            foreach (DataRow dr in dtMaterials.Rows)
            {
                DataRow r = dtCbo.NewRow();

                // Nên dùng tên cột từ Database thay vì chỉ số 0, 1 để tránh nhầm lẫn
                // Giả sử dr[0] là ID, dr[1] là Code, dr[2] là Name
                r["ID"] = dr[0];
                r["NAME"] = $"{dr[2]}|{dr[1]}"; // Ghép mã và tên hiển thị

                dtCbo.Rows.Add(r);
            }

            // 4. Gán nguồn dữ liệu cho ComboBox
            // Lưu ý: Phải gán DisplayMember và ValueMember TRƯỚC khi gán DataSource
            cboStandard.DataSource = null; // Clear dữ liệu cũ nếu có
            //cboStandard.DisplayMember = "NAME";
            //cboStandard.ValueMember = "ID";
            //cboStandard.DataSource = dtCbo;
            _isStandardLoaded = true;

            Common.Common.SetupComboBoxSearchUsingDataTable(cboStandard, dtCbo, "NAME", "ID");
            cboStandard.TextUpdate += CboStandard_TextUpdate;
        }

        private void CboStandard_TextUpdate(object? sender, EventArgs e)
        {
            Common.Common.ComboBoxTextUpdateForDataTable(sender, e);
        }

        private async Task LoadMaterialByCate(int cateId)
        {
            // 1. Lấy dữ liệu từ Service
            DataTable dtMaterials = await _productServices.GetMaterials(cateId);

            // 2. Khởi tạo DataTable mới cho ComboBox và định nghĩa cấu trúc ngay từ đầu
            DataTable dtCbo = new DataTable();
            dtCbo.Columns.Add("ID", typeof(int));
            dtCbo.Columns.Add("NAME", typeof(string));

            // 3. Duyệt và đổ dữ liệu
            foreach (DataRow dr in dtMaterials.Rows)
            {
                DataRow r = dtCbo.NewRow();

                // Nên dùng tên cột từ Database thay vì chỉ số 0, 1 để tránh nhầm lẫn
                // Giả sử dr[0] là ID, dr[1] là Code, dr[2] là Name
                r["ID"] = dr[0];
                r["NAME"] = $"{dr[1]}-{dr[2]}"; // Ghép mã và tên hiển thị

                dtCbo.Rows.Add(r);
            }

            // 4. Gán nguồn dữ liệu cho ComboBox
            // Lưu ý: Phải gán DisplayMember và ValueMember TRƯỚC khi gán DataSource
            cboMaterial.DataSource = null; // Clear dữ liệu cũ nếu có
            //cboMaterial.DisplayMember = "NAME";
            //cboMaterial.ValueMember = "ID";
            //cboMaterial.DataSource = dtCbo;
            Common.Common.SetupComboBoxSearchUsingDataTable(cboMaterial, dtCbo, "NAME", "ID");
            cboMaterial.TextUpdate += CboMaterial_TextUpdate; ;
        }

        private void CboMaterial_TextUpdate(object? sender, EventArgs e)
        {
            Common.Common.ComboBoxTextUpdateForDataTable(sender, e);
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private async void btnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtCode.Text) || txtCode.Text.Trim().Length < 12)
            {
                MessageBox.Show($"Item code chưa đúng định dạng !\nHãy chọn lại Standard !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;

            }
            itemCode = txtCode.Text.Trim();

            this.Close();
            //var material_Detail = new Material_Detail()
            //{
            //    Material_Detail_Id = Convert.ToInt32(cboMaterial.SelectedValue.ToString()),
            //    Material_Detail_Number = itemNumberOfMaterial,
            //};

            //var product = new ProductAddModel()
            //{
            //    Code = txtCode.Text.Trim()
            //};

            //if (await _productServices.InsertMaterialTypeDetailItem(material_Detail) > 0
            //    && await _productServices.InsertProduct(product) > 0)
            //{
            //    MessageBox.Show($"OK !",
            //    "OK", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //}
            //else
            //{
            //    bool rs = await _productServices.InsertProduct(product) > 0;
            //    MessageBox.Show($"No00000 !",
            //    "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //}
        }

        private async void cboMaterialCate_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_isLoadedMaterialCate)
            {
                await LoadMaterialByCate(Convert.ToInt32(cboMaterialCate.SelectedValue.ToString()));
            }
        }

        private async void btnGenerate_Click(object sender, EventArgs e)
        {
            //itemNumberOfMaterial = await _productServices.GetItemNumberOfMaterialType(Convert.ToInt32(cboMaterial.SelectedValue.ToString()));
            //var orgiCode = cboOriginal.Text.ToString().Trim().Split('-')[0];
            //var stanCode = cboStandard.Text.ToString().Trim().Split('|')[1];
            //var materialCode = cboMaterial.Text.ToString().Trim().Split('-')[0];
            //itemDetailNumber = itemNumberOfMaterial;
            //itemDetailId = cboMaterial.SelectedValue.ToString().Trim();

            //var itemCOde = orgiCode + materialCode + itemNumberOfMaterial + stanCode;
            //txtCode.Text = itemCOde;
            //isUseCodeAvailable = true;

            //frmOptions frmOptions = new frmOptions(Convert.ToInt32(cboMaterial.SelectedValue.ToString()));
            //frmOptions.ShowDialog();

            //// option 1: use code exist
            //if (!string.IsNullOrEmpty(frmOptions.ItemCode))
            //{
            //    txtCode.Text = frmOptions.ItemCode;
            //    isUseCodeAvailable = true;
            //}
            //else
            //{
            //    // option 2: create code
            //    itemNumberOfMaterial = await _productServices.GetItemNumberOfMaterialType(Convert.ToInt32(cboMaterial.SelectedValue.ToString()));
            //    var orgiCode = cboOriginal.Text.ToString().Trim().Split('-')[0];
            //    var stanCode = cboStandard.Text.ToString().Trim().Split('|')[1];
            //    var materialCode = cboMaterial.Text.ToString().Trim().Split('-')[0];
            //    itemDetailNumber = itemNumberOfMaterial;
            //    itemDetailId = cboMaterial.SelectedValue.ToString().Trim();

            //    var itemCOde = orgiCode + materialCode + itemNumberOfMaterial + stanCode;
            //    txtCode.Text = itemCOde;
            //}
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            itemCode = string.Empty;
            itemDetailId = string.Empty;
            itemDetailNumber = string.Empty;
            this.Close();
        }

        private void cboMaterialCate_Validating(object sender, CancelEventArgs e)
        {
            Common.Common.AutoCompleteComboboxValidating(sender as ComboBox, e);
        }

        private void cboMaterial_Validating(object sender, CancelEventArgs e)
        {
            if (cboMaterial.Items.Count <= 0) return;
            Common.Common.AutoCompleteComboboxValidating(sender as ComboBox, e);
        }

        private void cboOriginal_Validating(object sender, CancelEventArgs e)
        {
            Common.Common.AutoCompleteComboboxValidating(sender as ComboBox, e);
        }

        private void cboStandard_Validating(object sender, CancelEventArgs e)
        {
            Common.Common.AutoCompleteComboboxValidating(sender as ComboBox, e);
        }

        private async void cboMaterial_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (!_isLoaded || cboMaterial.Items.Count <= 0) return;
            //var dtItemExistedList = await _productServices.GetitemExistedList(Convert.ToInt32(cboMaterial.SelectedValue.ToString()));
            //dgvItemExist.DataSource = dtItemExistedList;
        }

        private void dgvItemExist_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {

        }

        private void dgvItemExist_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            Common.Common.RenderNumbering(sender, e);
        }

        private void dgvItemExist_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            DataGridViewRow row = dgvItemExist.Rows[e.RowIndex];
            txtCode.Text = row.Cells[4].Value.ToString();
            _isClickGrid = true;
        }

        private void cboStandard_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!_isStandardLoaded) return;
            txtCode.Text = !string.IsNullOrEmpty(txtCode.Text) ? txtCode.Text.Substring(0, 9) + cboStandard.Text.ToString().Trim().Split('|')[1] : "";
        }

        private async void btnShowExisted_Click(object sender, EventArgs e)
        {
            var dtItemExistedList = await _productServices.GetitemExistedList(Convert.ToInt32(cboMaterial.SelectedValue.ToString()));
            dgvItemExist.DataSource = dtItemExistedList;
            dtItemCode = dtItemExistedList.Copy();

            // 1. Khởi tạo ContextMenuStrip
            ContextMenuStrip menuStock = new ContextMenuStrip();

            // 2. Thêm các mục (Items) vào menu
            ToolStripMenuItem itemDeleteItem = new ToolStripMenuItem("📄 Xóa vật tư");

            menuStock.Items.AddRange(new ToolStripItem[] { itemDeleteItem });

            // 3. Gắn menu vào DataGridView
            if (AppSession.CurrentUser.Role_ID == 1)
            {
                dgvItemExist.ContextMenuStrip = menuStock;
            }

            // 4. Sự kiện khi click vào một mục trong menu
            itemDeleteItem.Click += (s, e) =>
            {
                if (dgvItemExist.CurrentRow != null)
                {
                    try
                    {
                        // Lấy dữ liệu từ dòng đang chọn
                        var row = dgvItemExist.CurrentRow;
                        int material_detail_id = int.Parse(row.Cells["material_detail_id"].Value?.ToString());

                        _warehouseServices.DeleteMaterialDetail(material_detail_id);
                        MessageBox.Show($"Xóa vật tư thành công !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (SqlException ex)
                    {
                        MessageBox.Show($"Xóa vật tư không thành công\nLỗi: {ex.Message}", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            };
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void btnAddMAterial_Click(object sender, EventArgs e)
        {
            var itemNumber = string.Empty;
            if (dgvItemExist.Rows.Count <= 0)
            {
                itemNumber = "001";
            }
            MaterialAddViewModel materialAddViewModel = new MaterialAddViewModel()
            {
                MaterialDetailNumber = string.IsNullOrEmpty(itemNumber) ? dgvItemExist.Rows[0].Cells[1].Value.ToString().Trim() : itemNumber,
                MaterialCode = cboMaterial.Text.ToString().Trim().Split('-')[0],
                MaterialID = cboMaterial.SelectedValue.ToString(),
            };

            frmAddMaterialDetail frm = new frmAddMaterialDetail(dtOrgins, materialAddViewModel);
            frm.ShowDialog();
            btnShowExisted.PerformClick();
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            if (txtSearch.Text.Length == 0)
            {
                //dgvProds.DataSource = this.dtProds;
                dgvItemExist.Refresh();
            }
            DataView dv = Common.Common.Search(txtSearch.Text, dtItemCode, ["material_detail_name"]);

            dgvItemExist.DataSource = dv;

        }

        private void dgvItemExist_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
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
                    Label lblItemCode = new Label() { Text = "Item Code:", Location = new Point(20, startY), AutoSize = true };
                    TextBox txtItemCode = new TextBox() { Location = new Point(150, startY - 3), Size = new Size(200, 25) };

                    // Qty_Import
                    Label lblItemNumber = new Label() { Text = "Item number:", Location = new Point(20, startY + spacing), AutoSize = true };
                    TextBox txtItemNumber = new TextBox() { Location = new Point(150, startY + spacing - 3), Size = new Size(200, 25) };

                    //// Weight_kg
                    //Label lblWeight = new Label() { Text = "Weight_kg:", Location = new Point(20, startY + (spacing * 2)), AutoSize = true };
                    //TextBox txtWeight = new TextBox() { Location = new Point(150, startY + (spacing * 2) - 3), Size = new Size(200, 25) };

                    //// Size
                    //Label lblSize = new Label() { Text = "Size:", Location = new Point(20, startY + (spacing * 3)), AutoSize = true };
                    //TextBox txtSize = new TextBox() { Location = new Point(150, startY + (spacing * 3) - 3), Size = new Size(200, 25) };

                    //// Name
                    //Label lblName = new Label() { Text = "Name:", Location = new Point(20, startY + (spacing * 4)), AutoSize = true };
                    //TextBox txtName = new TextBox() { Location = new Point(150, startY + (spacing * 4) - 3), Size = new Size(200, 25) };

                    // --- 3. Cấu hình các Button ---
                    // Button Cancel (Nền xám, chữ trắng)
                    Button btnCancel = new Button()
                    {
                        Text = "Cancel",
                        Location = new Point(150, startY + spacing - 3),
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
                        Location = new Point(260, startY + spacing - 3),
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
                        string itemNumber = txtItemNumber.Text;
                        //string weight = txtWeight.Text;
                        //string sizeValue = txtSize.Text;
                        //string name = txtName.Text;

                        var material_detail_id = dgvItemExist.CurrentRow.Cells[0].Value;
                        var item_code = txtItemCode.Text.Trim().Substring(0, 9);

                        _warehouseServices.UpdateMaterialDetail(Convert.ToInt32(material_detail_id), item_code, itemNumber);

                        //if (!string.IsNullOrEmpty(txtQty.Text))
                        //{
                        //    _service.ModifyQtyImportOfWarehouseImport(w);
                        //}
                        //if (!string.IsNullOrEmpty(txtWeight.Text))
                        //{
                        //    _service.ModifyWeightOfWarehouseImport(w);
                        //}
                        //if (!string.IsNullOrEmpty(txtSize.Text))
                        //{
                        //    _service.ModifySizeOfWarehouseImport(w);
                        //}
                        //if (!string.IsNullOrEmpty(txtName.Text))
                        //{
                        //    _service.ModifyNameOfWarehouseImport(w);
                        //}
                        // Hiển thị kết quả lấy được để kiểm tra
                        string info = $"Dữ liệu đã thu thập:\n" +
                                      $"- Item Code: {itemCode}\n";
                                      //$"- Qty: {qty}\n" +
                                      //$"- Weight: {weight}\n" +
                                      //$"- Size: {sizeValue}";

                        MessageBox.Show(info, "Kết quả lưu dữ liệu", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // Sau khi xử lý xong có thể đóng form hoặc giữ lại tùy ý
                        frm.DialogResult = DialogResult.OK;
                    };

                    // --- 5. Thêm Controls vào Form và hiển thị ---
                    frm.Controls.AddRange(new Control[] {
                        lblItemCode, txtItemCode,
                        lblItemNumber, txtItemNumber,
                        //lblWeight, txtWeight,
                        //lblSize, txtSize,
                        //lblName, txtName,
                        btnCancel, btnSave
                    });

                    frm.AcceptButton = btnSave; // Nhấn Enter để Save [cite: 115]
                    frm.CancelButton = btnCancel; // Nhấn Esc để Cancel [cite: 115]

                    frm.ShowDialog();
                }
            }
        }
    }
}
