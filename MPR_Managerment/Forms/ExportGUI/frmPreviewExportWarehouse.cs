using System;
using System.Data;
using System.Diagnostics.CodeAnalysis;
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
            string sql = @"SELECT we.[Export_Detail_Id],we.[Export_ID],we.[Import_ID],wi.ID_Code,wi.Item_Name,wi.Material,wi.Size,we.[Qty_Export],wi.UNIT,we.[Notes]
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
            ConfigureHisExportGrid();
        }

        private void ConfigureHisExportGrid()
        {
            dgvDetails.ReadOnly = false;
            dgvDetails.AllowUserToAddRows = false;
            dgvDetails.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dgvDetails.BackgroundColor = Color.White;
            dgvDetails.BorderStyle = BorderStyle.FixedSingle;
            dgvDetails.RowHeadersVisible = false;
            dgvDetails.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvDetails.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvDetails.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvDetails.EnableHeadersVisualStyles = false;
            dgvDetails.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);

            if (!dgvDetails.Columns.Contains("Select"))
            {
                var chkCol = new DataGridViewCheckBoxColumn { Name = "chkSelect", HeaderText = "", Width = 30, ReadOnly = false };
                dgvDetails.Columns.Insert(0, chkCol);
            }

            foreach (DataGridViewColumn col in dgvDetails.Columns)
            {
                col.ReadOnly = true;
                if (col.Name.Contains("ID", StringComparison.OrdinalIgnoreCase)) col.Visible = false;
                if (col.Name == "chkSelect") col.ReadOnly = false;
                if (col.Name == "Qty_Export") col.ReadOnly = false;
                if (col.Name == "Notes") col.ReadOnly = false;
            }

            // Set specific column widths
            if (dgvDetails.Columns.Contains("chkSelect")) { dgvDetails.Columns["chkSelect"].Width = 30; dgvDetails.Columns["chkSelect"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvDetails.Columns.Contains("ID_Code")) { dgvDetails.Columns["ID_Code"].Width = 150; dgvDetails.Columns["ID_Code"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvDetails.Columns.Contains("Item_Name")) { dgvDetails.Columns["Item_Name"].Width = 200; dgvDetails.Columns["Item_Name"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvDetails.Columns.Contains("Material")) { dgvDetails.Columns["Material"].Width = 150; dgvDetails.Columns["Material"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvDetails.Columns.Contains("Size")) { dgvDetails.Columns["Size"].Width = 120; dgvDetails.Columns["Size"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvDetails.Columns.Contains("Qty_Export")) { dgvDetails.Columns["Qty_Export"].Width = 100; dgvDetails.Columns["Qty_Export"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvDetails.Columns.Contains("UNIT")) { dgvDetails.Columns["UNIT"].Width = 80; dgvDetails.Columns["UNIT"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvDetails.Columns.Contains("Notes")) { dgvDetails.Columns["Notes"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; }
            dgvDetails.CellContentClick += DgvDetails_CellContentClick;
            dgvDetails.EditingControlShowing += DgvDetails_EditingControlShowing;
        }

        private void DgvDetails_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress -= NumericOnly_KeyPress;
            if (dgvDetails.CurrentCell != null && dgvDetails.Columns[dgvDetails.CurrentCell.ColumnIndex].Name == "Qty_Export")
            {
                e.Control.KeyPress += NumericOnly_KeyPress;
            }
        }

        private void NumericOnly_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '.' && e.KeyChar != ',')
            {
                e.Handled = true;
            }

            // Only allow one decimal separator
            if ((e.KeyChar == '.' || e.KeyChar == ',') && sender is TextBox tb && (tb.Text.Contains('.') || tb.Text.Contains(',')))
            {
                e.Handled = true;
            }
        }

        private void DgvDetails_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0 && e.RowIndex >= 0)
            {
                bool isChecked = (bool)dgvDetails.Rows[e.RowIndex].Cells[0].EditedFormattedValue;
                btnDelete.Enabled = dgvDetails.Rows.Cast<DataGridViewRow>().Any(r => (bool)(r.Cells[0].EditedFormattedValue ?? false));
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
            btnDelete.Enabled = false;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            // Logic to save header and details to database
            MessageBox.Show("Save functionality to be implemented.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }

        private void frmPreviewExportWarehouse_Load(object sender, EventArgs e)
        {

        }

        private void dgvDetails_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            int export_detail_id = Convert.ToInt32(dgvDetails.Rows[e.RowIndex].Cells[1].Value.ToString() ?? "0");

            frmSelectItemExport frmSelectItemExport = new frmSelectItemExport();
            frmSelectItemExport.ShowDialog();

            if (!frmSelectItemExport.isCancel || frmSelectItemExport.selectedList.Count <= 0) return;

            var model = frmSelectItemExport.selectedList[0];

            dgvDetails.Rows[e.RowIndex].Cells[4].Value = model.ID_Code;
            dgvDetails.Rows[e.RowIndex].Cells[5].Value = model.Item_Name;
            dgvDetails.Rows[e.RowIndex].Cells[6].Value = model.Size;
            dgvDetails.Rows[e.RowIndex].Cells[9].Value = model.UNIT;
        }
    }
}