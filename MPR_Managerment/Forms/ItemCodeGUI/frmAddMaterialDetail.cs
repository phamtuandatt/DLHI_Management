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
                var item_code = $"{cboOriginal.Text.ToString().Split('-')[0]}{MaterialDetailAdd.MaterialCode}{MaterialDetailAdd.MaterialDetailNumber}";
                var material_Detail = new Models.Material_Detail()
                {
                    Material_Detail_Number = !string.IsNullOrEmpty(MaterialDetailAdd.MaterialDetailNumber) ? MaterialDetailAdd.MaterialDetailNumber : "001",
                    Material_Detail_Code = MaterialDetailAdd.MaterialCode,
                    Item_Code_Existed = item_code,
                    Material_Detail_Name = txtCode.Text,
                    MaterialID = MaterialDetailAdd.MaterialID,
                };

                var rs = await _productServices.InsertMaterialTypeDetailItem(material_Detail);
                MessageBox.Show($"Thêm vật tư thành công !","Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"No00000 !",
                "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }

    public class MaterialAddViewModel
    {
        public string MaterialDetailNumber { get; set; }
        public string MaterialCode { get; set; }
        public string MaterialID { get; set; }
        //public string MaterialCode { get; set; }
    }
}
