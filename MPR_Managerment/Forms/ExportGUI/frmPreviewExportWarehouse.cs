using System;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;
using MPR_Managerment.Services;

namespace MPR_Managerment.Forms.ExportGUI
{
    public partial class frmPreviewExportWarehouse : Form
    {
        private DataRow _headerRow;
        private DataTable _dtDetails;
        private bool _isCreate = false;
        private string _number = "000";
        private ProjectInfo _projectFrom = new ProjectInfo();
        private ProjectInfo _projectTo = new ProjectInfo();
        private WarehouseService _warehouseServices = new WarehouseService();

        public bool IsAction { get; set; } = false;

        public frmPreviewExportWarehouse(DataRow headerRow)
        {
            InitializeComponent();
            _headerRow = headerRow;
            LoadData();
        }

        public frmPreviewExportWarehouse(bool isCreate)
        {
            InitializeComponent();
            _isCreate = isCreate;
            GetExportNumber();
            txtCreatedBy.Text = AppSession.CurrentUser.Full_Name.ToUpper();
            cboStatus.SelectedIndex = 0;
            cboStatus.SelectedItem = "Chưa xác nhận";
        }

        private void LoadData()
        {
            txtExportNo.Text = _headerRow["Export_No"].ToString();
            txtToProject.Text = _headerRow["From_Project_Name"].ToString();
            txtFromProject.Text = _headerRow["To_Project_Name"].ToString();
            txtCreatedBy.Text = _headerRow["Create_By"].ToString();
            cboStatus.SelectedItem = _headerRow["Status"].ToString();

            txtExportNo.ReadOnly = true;

            LoadDetails();
        }

        private void GetExportNumber()
        {
            string sql = @"SELECT RIGHT('000' + CAST(MAX(CAST(Export_Number AS INT)) + 1 AS VARCHAR), 3) AS Next_Export_Number FROM ExportWarehouseHeader;";

            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    try
                    {
                        conn.Open();
                        // 4. Execute and safely convert the result to a string
                        object result = cmd.ExecuteScalar();
                        string myString = result != DBNull.Value && result != null ? result.ToString() : "Default Value";

                        // 5. Display the string in a WinForms control
                        _number = myString;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                }
            }
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
            var total = 0;
            var model = new ExportWarehouseHeaderModel()
            {
                ExportNo = txtExportNo.Text.Trim(),
                ExportNumber = _number,
                FromProjectID = _projectFrom.Id,
                FromProjectCode = _projectFrom.ProjectCode,
                FromProjectName = _projectFrom.ProjectName,
                ToProjectID = _projectTo.Id,
                ToProjectCode = _projectTo.ProjectCode,
                ToProjectName = _projectTo.ProjectName,
                ExportTotals = total,
                Status = cboStatus.Text.Trim(),
                Notes = "",
                CreateBy = txtCreatedBy.Text.Trim()
            };

            // Gọi tầng xử lý dữ liệu
            bool isSaved = _warehouseServices.SaveFullInvoice(model, _dtDetails); // _dtDetails là DataTable chi tiết hàng hóa

            if (isSaved)
            {
                MessageBox.Show("Đã lưu toàn bộ Phiếu Xuất và Chi Tiết thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                IsAction = true;
                this.Close();
            }
            else
            {
                MessageBox.Show(" Lưu toàn bộ Phiếu Xuất và Chi Tiết không thành công!", "Thất bại", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                IsAction = false;
            }
        }

        private void InitDataTable()
        {
            _dtDetails = new DataTable();
            _dtDetails.Columns.Add("Export_Detail_Id", typeof(int));
            _dtDetails.Columns.Add("Export_ID", typeof(int));
            _dtDetails.Columns.Add("Import_ID", typeof(int));
            _dtDetails.Columns.Add("ID_Code", typeof(string));
            _dtDetails.Columns.Add("Item_Name", typeof(string));
            _dtDetails.Columns.Add("Material", typeof(string));
            _dtDetails.Columns.Add("Size", typeof(string));
            _dtDetails.Columns.Add("Qty_Export", typeof(decimal));
            _dtDetails.Columns.Add("UNIT", typeof(string));
            _dtDetails.Columns.Add("Notes", typeof(string));

            dgvDetails.DataSource = _dtDetails;
            ConfigureHisExportGrid();
        }

        private void frmPreviewExportWarehouse_Load(object sender, EventArgs e)
        {
            if (_dtDetails == null)
            {
                InitDataTable();
            }
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

        private void btnAddRow_Click(object sender, EventArgs e)
        {
            frmSelectItemExport frmSelectItemExport = new frmSelectItemExport(_isCreate);
            frmSelectItemExport.ShowDialog();

            if (!frmSelectItemExport.isCancel || frmSelectItemExport.selectedList.Count <= 0) return;

            var list = frmSelectItemExport.selectedList;

            if (_dtDetails == null)
            {
                InitDataTable();
            }

            foreach (var item in list)
            {
                bool exists = false;
                foreach (DataRow r in _dtDetails.Rows)
                {
                    if (r["Import_ID"] != DBNull.Value && Convert.ToInt32(r["Import_ID"]) == item.Import_ID)
                    {
                        exists = true;
                        break;
                    }
                }

                if (!exists)
                {
                    DataRow row = _dtDetails.NewRow();
                    row["Export_Detail_Id"] = 0;
                    row["Export_ID"] = 0;
                    row["Import_ID"] = item.Import_ID;
                    row["ID_Code"] = item.ID_Code;
                    row["Item_Name"] = item.Item_Name;
                    row["Material"] = item.Material;
                    row["Size"] = item.Size;
                    row["Qty_Export"] = 1;
                    row["UNIT"] = item.UNIT;
                    row["Notes"] = item.Notes;
                    _dtDetails.Rows.Add(row);
                }
            }
        }

        private void tableLayoutPanel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnFromWarehouse_Click(object sender, EventArgs e)
        {
            frmSelectWarehouse frmSelectWarehouse = new frmSelectWarehouse();
            frmSelectWarehouse.ShowDialog();

            if (frmSelectWarehouse.ProjectInfo is not null
                && frmSelectWarehouse.ProjectInfo.Id != 0)
            {
                txtFromProject.Text = $"{frmSelectWarehouse.ProjectInfo.ProjectCode} - {frmSelectWarehouse.ProjectInfo.ProjectName}";
                _projectFrom = frmSelectWarehouse.ProjectInfo;
                txtExportNo.Text = $"XK-PNK-DV-{_projectFrom.ProjectCode}-PC-{_number}";
                txtExportNo.ReadOnly = true;
            }
        }

        private void btnToWarehosue_Click(object sender, EventArgs e)
        {
            frmSelectWarehouse frmSelectWarehouse = new frmSelectWarehouse();
            frmSelectWarehouse.ShowDialog();

            if (frmSelectWarehouse.ProjectInfo is not null
                && frmSelectWarehouse.ProjectInfo.Id != 0)
            {
                txtToProject.Text = $"{frmSelectWarehouse.ProjectInfo.ProjectCode} - {frmSelectWarehouse.ProjectInfo.ProjectName}";
                _projectTo = frmSelectWarehouse.ProjectInfo;
            }
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void txtFromProject_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            IsAction = false;
            this.Close();
        }
    }
}