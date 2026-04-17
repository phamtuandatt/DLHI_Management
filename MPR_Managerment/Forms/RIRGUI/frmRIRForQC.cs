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
    public partial class frmRIRForQC : Form
    {
        public frmRIRForQC()
        {
            InitializeComponent();
            ucRIRForQC ucRIRForQC = new ucRIRForQC();
            ucRIRForQC.Dock = DockStyle.Fill;
            this.Controls.Add(ucRIRForQC);
            ucRIRForQC.BringToFront();
        }
    }
}
