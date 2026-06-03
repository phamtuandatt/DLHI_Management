using Microsoft.Data.SqlClient;
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

namespace MPR_Managerment.Forms.ImportWarehouseGUI
{
    public partial class frmUpdateIDCode : Form
    {
        private WarehouseImport _warehouseImport;
        private WarehouseService _warehouseService = new WarehouseService();

        public frmUpdateIDCode(WarehouseImport warehouseImport)
        {
            InitializeComponent();
            Common.Common.CreateButtonSave(btnSave, "");
            Common.Common.CreateButtonCancel(btnCancel, "Hủy");
            _warehouseImport = warehouseImport;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (_warehouseService != null)
            {
                try
                {
                    _warehouseImport.QC_Code = txtIDCode.Text.Trim().ToUpper();
                    _warehouseService.UpdateIDCode(_warehouseImport);
                    this.Close();
                }
                catch (SqlException ex)
                {
                    MessageBox.Show("Cảnh báo", $"Lỗi: {ex.Message} khi cập nhật ID cho {_warehouseImport.Item_Name}", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
