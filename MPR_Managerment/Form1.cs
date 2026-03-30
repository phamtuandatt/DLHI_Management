using System;
using System.Windows.Forms;
using MPR_Managerment.Forms;
using MPR_Managerment.Helpers;

namespace MPR_Managerment
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.Load += Form1_Load;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (!DatabaseHelper.TestConnection())
            {
                MessageBox.Show("Không thể kết nối Database!",
                    "Lỗi kết nối", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Mở form MPR (đổi thành frmSupplier nếu muốn test Supplier)
            var frm = new frmMain();
            frm.Show();
        }
    }
}