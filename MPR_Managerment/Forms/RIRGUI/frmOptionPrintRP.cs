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
    public partial class frmOptionPrintRP : Form
    {
        public bool IsProjectSelected { get; set; } = false;
        public bool IsRIRNoSelected { get; set; } = false;
        public frmOptionPrintRP(string btnProjectText, string btnRIRNoText)
        {
            InitializeComponent();
            btnProject.Text = btnProjectText;
            btnRIRNo.Text = btnRIRNoText;
        }

        private void btnProject_Click(object sender, EventArgs e)
        {
            IsProjectSelected = true;
            this.Close();
        }

        private void btnRIRNo_Click(object sender, EventArgs e)
        {
            IsRIRNoSelected = true;
            this.Close();
        }
    }
}
