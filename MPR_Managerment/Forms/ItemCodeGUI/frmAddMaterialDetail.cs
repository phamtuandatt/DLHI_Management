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

namespace MPR_Managerment.Forms.ItemCodeGUI
{
    public partial class frmAddMaterialDetail : Form
    {
        private DataTable dtOrigin = new DataTable();   
        private MaterialAddViewModel MaterialDetailAdd = new MaterialAddViewModel();
        private ProductServices _productServices = new ProductServices();

        public frmAddMaterialDetail(DataTable dtOrigin, MaterialAddViewModel materialDetailAdd)
        {
            InitializeComponent();
            this.dtOrigin = dtOrigin;
            this.MaterialDetailAdd = materialDetailAdd;
            lblName.Text = $"Thêm vật tư cho: {materialDetailAdd.MaterialCode}";

            cboOriginal.DisplayMember = "NAME";
            cboOriginal.ValueMember = "ID";
            cboOriginal.DataSource = dtOrigin;
        }

        private async void btnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                var number = IncrementString(MaterialDetailAdd.MaterialDetailNumber);
                var item_code = ($"{cboOriginal.Text.ToString().Split('-')[0]}{MaterialDetailAdd.MaterialCode}{number}");
                var material_Detail = new Models.Material_Detail()
                {
                    Material_Detail_Number = number ?? "001",
                    Material_Detail_Code = MaterialDetailAdd.MaterialCode,
                    Item_Code_Existed = item_code,
                    Material_Detail_Name = txtCode.Text,
                    MaterialID = MaterialDetailAdd.MaterialID,
                };

                var pModel = new ProductModel()
                {
                    Name = txtCode.Text,
                    Des2 = MaterialDetailAdd.MaterialSize,
                    Code = item_code,
                    //ProdMaterialCode = imp.Material
                };

                var rs_t = await _productServices.InsertMaterialTypeDetailItem(material_Detail);
                var rs = await _productServices.SaveProduct_Async(pModel, false);

                MessageBox.Show($"Thêm vật tư thành công !","Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"No00000 !",
                "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public static string IncrementString(string input)
        {
            // 1. Kiểm tra nếu chuỗi rỗng hoặc không phải là số
            if (string.IsNullOrWhiteSpace(input) || !int.TryParse(input, out int number))
            {
                return "001"; // Hoặc xử lý lỗi tùy nhu cầu của bạn
            }

            // 2. Tăng giá trị lên 1
            int incrementedNumber = number + 1;

            // 3. Trả về chuỗi mới với cùng độ dài bằng cách thêm các số 0 phía trước (Padding)
            // D2 là 2 chữ số, D3 là 3 chữ số. Ở đây ta dùng độ dài của chuỗi đầu vào.
            return incrementedNumber.ToString().PadLeft(input.Length, '0');
        }
    }

    public class MaterialAddViewModel
    {
        public string MaterialDetailNumber { get; set; }
        public string MaterialCode { get; set; }
        public string MaterialID { get; set; }
        public string MaterialSize { get; set; }
        //public string MaterialCode { get; set; }
    }
}
