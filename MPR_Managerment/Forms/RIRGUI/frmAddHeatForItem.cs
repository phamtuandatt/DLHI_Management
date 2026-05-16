using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MPR_Managerment.Forms.RIRGUI
{
    public partial class frmAddHeatForItem : Form
    {
        public decimal Quantity { get; set; }
        public string MTRNo { get; set; }
        public string HeatNo { get; set; }
        public string ID_Code { get; set; }

        public frmAddHeatForItem(string MRTNo)
        {
            InitializeComponent();
            txtMTR.Text = MRTNo;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Quantity = Convert.ToInt32(txtQty.Text.Trim());
            MTRNo = txtMTR.Text;
            HeatNo = txtHeat.Text;
            ID_Code = txtIDCode.Text;
            this.Close();
        }

        private void txtQty_KeyPress(object sender, KeyPressEventArgs e)
        {
            Common.Common.Column_KeyPress_Digital(sender, e);
        }
    }
}
