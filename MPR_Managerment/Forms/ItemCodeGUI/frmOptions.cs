using Accessibility;
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
    public partial class frmOptions : Form
    {
        public string ItemCode { get; set; } = string.Empty;
        public int materialCode;

        private ProductServices _productServices = new ProductServices();
        public frmOptions(int materialCode)
        {
            InitializeComponent();
            this.materialCode = materialCode;
        }

        private async void btnGetCode_Click(object sender, EventArgs e)
        {
            ItemCode = await _productServices.GetCodeExistedByMaterilDetail(materialCode);
            this.Close();
        }

        private void btnCreateCode_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
