using System;
using System.Data;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;

namespace MPR_Managerment.Forms.ExportGUI
{
    public partial class frmPreviewExportWarehouse : Form
    {
        private DataRow _headerRow;
        private DataTable _dtDetails;

        public frmPreviewExportWarehouse(DataRow headerRow)
        {
            InitializeComponent();
            _headerRow = headerRow;
            LoadData();
        }

        private void LoadData()
        {
            txtExportNo.Text = _headerRow["Export_No"].ToString();
            txtFromProject.Text = _headerRow["From_Project_Name"].ToString();
            txtToProject.Text = _headerRow["To_Project_Name"].ToString();
            txtCreateBy.Text = _headerRow["Create_By"].ToString();

            txtExportNo.ReadOnly = true;
            
            LoadDetails();
        }

        private void LoadDetails()
        {
            string sql = @"SELECT we.[Export_Detail_Id],we.[Export_ID],we.[Import_ID],wi.Item_Code,wi.Item_Name,wi.Material,wi.Size,we.[Qty_Export],wi.UNIT,we.[Notes]
                           FROM [dbo].[ExportWarehouseDetail] we 
                           INNER JOIN Warehouse_Import wi ON we.Import_ID = wi.Import_ID 
                           WHERE we.Export_ID = @ExportID";

            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@ExportID", _headerRow["Export_ID"]);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    _dtDetails = new DataTable();
                    da.Fill(_dtDetails);
                }
            }

            dgvDetails.DataSource = _dtDetails;
            
            // Add CheckBox column
            DataGridViewCheckBoxColumn chk = new DataGridViewCheckBoxColumn();
            chk.Name = "chkSelect";
            chk.HeaderText = "";
            dgvDetails.Columns.Insert(0, chk);
            dgvDetails.CellContentClick += DgvDetails_CellContentClick;
        }

        private void DgvDetails_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0 && e.RowIndex >= 0)
            {
                bool isChecked = (bool)dgvDetails.Rows[e.RowIndex].Cells[0].EditedFormattedValue;
                btnDelete.Visible = dgvDetails.Rows.Cast<DataGridViewRow>().Any(r => (bool)(r.Cells[0].EditedFormattedValue ?? false));
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            for (int i = dgvDetails.Rows.Count - 1; i >= 0; i--)
            {
                if ((bool)(dgvDetails.Rows[i].Cells[0].EditedFormattedValue ?? false))
                {
                    _dtDetails.Rows.RemoveAt(i);
                }
            }
            btnDelete.Visible = false;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            // Logic to save header and details to database
            MessageBox.Show("Save functionality to be implemented.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }
    }
}