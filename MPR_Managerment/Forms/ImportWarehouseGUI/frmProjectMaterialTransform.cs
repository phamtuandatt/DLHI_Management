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
    public partial class frmProjectMaterialTransform : Form
    {
        private readonly WarehouseService _warehouseServices = new WarehouseService();
        private int importId = 0;
        public frmProjectMaterialTransform(List<ProjectInfo> dtProject, int importId, decimal maxQty)
        {
            InitializeComponent();
            // Gắn AI Trợ lý — hiển thị nút floating + chat panel
            frmAIChat.Attach(this);

            Common.Common.CreateButtonSave(btnSave);
            Common.Common.CreateButtonCancel(btnCancel, "🔄 Hủy");

            cboProject.Items.Clear();
            foreach (var p in dtProject)
            {
                cboProject.Items.Add(p.ProjectCode);
            }
            cboProject.SelectedIndex = 0;
            this.importId = importId;
            txtQty.Maximum = maxQty;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                var wh = _warehouseServices.GetImportByImportID(this.importId);
                wh.Notes = $"{txtNotes.Text.Trim()} - Chuyển {txtQty.Value} {wh.UNIT} của dự án {wh.Project_Code.ToUpper()} sang cho dự án {cboProject.SelectedItem.ToString().ToUpper()}";
                string notes_2 = $"Mượn {txtQty.Value} {wh.UNIT} của dự án {wh.Project_Code.ToUpper()}";
                var rs = _warehouseServices.ProjectMaterialTransform(wh, wh.Project_Code, cboProject.SelectedItem.ToString() ?? wh.Project_Code, txtQty.Value, notes_2);
                MessageBox.Show($"✅ Chuyển vật tư thành công!\nTên vật tư: {wh.Item_Name}\nKích thước vật tư: {wh.Size} items", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi chuyển vật tư: {ex.Message}", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
