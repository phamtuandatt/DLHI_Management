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
    public partial class ucFillInvoiceNo : UserControl
    {
        private ProjectService _projectService = new ProjectService();
        private POService _poServices = new POService();
        private WarehouseService _warehouseServices = new WarehouseService();
        private bool _isLoaded = false;
        private bool _isPOLoaded = false;

        public ucFillInvoiceNo()
        {
            InitializeComponent();
            LoadProjects();
        }

        private async void LoadProjects()
        {
            var dt = await _projectService.GetProjects();
            cboProject.DisplayMember = "ProjectCode";
            cboProject.ValueMember = "ProjectCode";
            cboProject.DataSource = dt;
            _isLoaded = true;
        }

        private async void LoadPOByProjectCode(string projectCode)
        {
            var dt = await _poServices.GetPOByProjectCode(projectCode);
            cboPO.DisplayMember = "PONo";
            cboPO.ValueMember = "PO_ID";
            cboPO.DataSource = dt;
            _isPOLoaded = true;
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            if (!_isPOLoaded || !_isLoaded) return;
            var dt = _warehouseServices.GetWarehouseImportByPOId(Convert.ToInt32(cboPO.SelectedValue.ToString()));
            dgvList.DataSource = dt;
        }

        private void cboProject_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!_isLoaded) return;
            LoadPOByProjectCode(cboProject.SelectedValue.ToString());
        }
    }
}
