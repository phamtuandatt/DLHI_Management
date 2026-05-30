using Microsoft.IdentityModel.Tokens;
using MPR_Managerment.Common;
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

        public bool IsClose { get; set; } = false;

        public frmAddHeatForItem(string MRTNo)
        {
            InitializeComponent();
            // Gắn AI Trợ lý — hiển thị nút floating + chat panel
            frmAIChat.Attach(this);
            txtMTR.Text = MRTNo;

            Common.Common.CreateButtonSave(btnSave);
            Common.Common.CreateButtonCancel(btnCancel, "Hủy");

        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Quantity = Convert.ToInt32(txtQty.Text.Trim() ?? "0");
            MTRNo = txtMTR.Text;
            HeatNo = txtHeat.Text;
            ID_Code = txtIDCode.Text;
            IsClose = true;
            this.Close();
        }

        private void txtQty_KeyPress(object sender, KeyPressEventArgs e)
        {
            Common.Common.Column_KeyPress_Digital(sender, e);
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            IsClose = false;
        }
    }
}
