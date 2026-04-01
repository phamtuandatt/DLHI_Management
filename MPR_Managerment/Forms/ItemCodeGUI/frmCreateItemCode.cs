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
        private bool _isLoaded = false;
        private string itemNumberOfMaterial = "";


        public string itemCode { get; set; } = string.Empty;
        public string itemDetailId { get; set; } = string.Empty;
        public string itemDetailNumber { get; set; } = string.Empty;
        public bool isUseCodeAvailable { get; set; } = false;

        public frmCreateItemCode()
        {
            InitializeComponent();
            // Đăng ký sự kiện Shown thay vì gọi trực tiếp ở đây
            this.Shown += (s, e) => {
                SetTabOrder();
            };
            this.Text = "Cấu hình Item Code";
            //this.Size = new Size(450, 650); // Tăng nhẹ chiều cao để cân đối khoảng cách
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
        }

        private async void frmCreateItemCode_Load(object sender, EventArgs e)
        {
            await LoadMaterialCate();
            if (_isLoaded)
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
                _isLoaded = true;
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
            cboStandard.DisplayMember = "NAME";
            cboStandard.ValueMember = "ID";
            cboStandard.DataSource = dtCbo;
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
            cboMaterial.DisplayMember = "NAME";
            cboMaterial.ValueMember = "ID";
            cboMaterial.DataSource = dtCbo;
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private async void btnSave_Click(object sender, EventArgs e)
        {
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
            if (_isLoaded)
            {
                await LoadMaterialByCate(Convert.ToInt32(cboMaterialCate.SelectedValue.ToString()));
            }
        }

        private async void btnGenerate_Click(object sender, EventArgs e)
        {
            frmOptions frmOptions = new frmOptions(Convert.ToInt32(cboMaterial.SelectedValue.ToString()));
            frmOptions.ShowDialog();
            // option 1: use code exist
            if (!string.IsNullOrEmpty(frmOptions.ItemCode))
            {
                txtCode.Text = frmOptions.ItemCode;
                isUseCodeAvailable = true;
            }
            else
            {
                // option 2: create code
                itemNumberOfMaterial = await _productServices.GetItemNumberOfMaterialType(Convert.ToInt32(cboMaterial.SelectedValue.ToString()));
                var orgiCode = cboOriginal.Text.ToString().Trim().Split('-')[0];
                var stanCode = cboStandard.Text.ToString().Trim().Split('|')[1];
                var materialCode = cboMaterial.Text.ToString().Trim().Split('-')[0];
                itemDetailNumber = itemNumberOfMaterial;
                itemDetailId = cboMaterial.SelectedValue.ToString().Trim();

                var itemCOde = orgiCode + materialCode + itemNumberOfMaterial + stanCode;
                txtCode.Text = itemCOde;
            }
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
            if (!_isLoaded || cboMaterial.Items.Count <= 0 || cboMaterial.SelectedValue.ToString() is null) return;
            var dtItemExistedList = await _productServices.GetitemExistedList(Convert.ToInt32(cboMaterial.SelectedValue.ToString()));
            dgvItemExist.DataSource = dtItemExistedList;
        }
    }
}
