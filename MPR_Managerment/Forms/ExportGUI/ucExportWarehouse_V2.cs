using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using MPR_Managerment.Services;

namespace MPR_Managerment.Forms.ExportGUI
{
    public partial class ucExportWarehouse_V2 : UserControl
    {
        private DataTable _dtHisExport = new DataTable();
        private ProjectService _projectService = new ProjectService();

        public ucExportWarehouse_V2()
        {
            InitializeComponent();
            this.Load += async (s, e) => await InitializeFormAsync();
        }

        private async Task InitializeFormAsync()
        {
            await LoadProjectsAsync();
            LoadStatuses();
            await LoadHisExportAsync();

            Common.Common.CreateButtonSearch(btnSearch, "🔍 Tìm kiếm");
            Common.Common.CreateButtonRefresh(btnRefresh);
            Common.Common.CreateButtonAdd(btnAddXK, "✅ Thêm phiếu mới");
            Common.Common.CreateButtonPrint(btnInXK, "🖨 In");
            Common.Common.CreateButtonSave(btnUpdateStatus, "Cập nhật trạng thái        ⏷");

            // Initialize status dropdown menu
            btnUpdateStatus.Click += (s, e) => {
                _statusMenu.Show(btnUpdateStatus, new System.Drawing.Point(0, btnUpdateStatus.Height));
            };

            btnSearch.Click += BtnSearch_Click;
            btnRefresh.Click += BtnRefresh_Click;

            dtpFromDate.Value = DateTime.Today.AddDays(-30);
        }

        private async Task LoadProjectsAsync()
        {
            DataTable dtProjects = await _projectService.GetProjects();
            cboProject.DataSource = dtProjects;
            cboProject.DisplayMember = "ProjectCode";
            cboProject.ValueMember = "ProjectCode";
            cboProject.SelectedIndex = -1;
        }

        private void LoadStatuses()
        {
            cboStatus.Items.AddRange(new string[] { "Pending", "Completed", "Cancelled" });
        }

        private async Task LoadHisExportAsync()
        {
            try
            {
                string sql = @"SELECT TOP (1000) [Export_ID],[Export_No],[From_Project_Name],[To_Project_Name],[Export_Totals],[Status],[Notes],[Create_By],[Create_Date],[Update_By],[Update_Date] FROM [dbo].[ExportWarehouseHeader]";

                using (SqlConnection conn = DatabaseHelper.GetConnection())
                {
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        await conn.OpenAsync();
                        using (SqlDataReader reader = await cmd.ExecuteReaderAsync())
                        {
                            _dtHisExport.Load(reader);
                        }
                    }
                }

                dgvHisExport.DataSource = _dtHisExport;
                ConfigureHisExportGrid();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi tải dữ liệu: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ConfigureHisExportGrid()
        {
            dgvHisExport.ReadOnly = false;
            dgvHisExport.AllowUserToAddRows = false;
            dgvHisExport.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dgvHisExport.BackgroundColor = Color.White;
            dgvHisExport.BorderStyle = BorderStyle.FixedSingle;
            dgvHisExport.RowHeadersVisible = false;
            dgvHisExport.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvHisExport.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvHisExport.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvHisExport.EnableHeadersVisualStyles = false;
            dgvHisExport.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);

            if (!dgvHisExport.Columns.Contains("Select"))
            {
                var chkCol = new DataGridViewCheckBoxColumn { Name = "Select", HeaderText = "", Width = 30, ReadOnly = false };
                dgvHisExport.Columns.Insert(0, chkCol);
            }

            if (!dgvHisExport.Columns.Contains("Print"))
            {
                var btnCol = new DataGridViewButtonColumn { Name = "Print", HeaderText = "In", Text = "In", UseColumnTextForButtonValue = true, Width = 60, ReadOnly = true };
                dgvHisExport.Columns.Add(btnCol);
            }

            foreach (DataGridViewColumn col in dgvHisExport.Columns)
            {
                if (col.Name != "Select") col.ReadOnly = true;
                if (col.Name.Contains("ID", StringComparison.OrdinalIgnoreCase)) col.Visible = false;
            }

            // Set specific column widths
            if (dgvHisExport.Columns.Contains("Select")) { dgvHisExport.Columns["Select"].Width = 30; dgvHisExport.Columns["Select"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvHisExport.Columns.Contains("Export_No")) { dgvHisExport.Columns["Export_No"].Width = 150; dgvHisExport.Columns["Export_No"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvHisExport.Columns.Contains("From_Project_Name")) { dgvHisExport.Columns["From_Project_Name"].Width = 180; dgvHisExport.Columns["From_Project_Name"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvHisExport.Columns.Contains("To_Project_Name")) { dgvHisExport.Columns["To_Project_Name"].Width = 180; dgvHisExport.Columns["To_Project_Name"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvHisExport.Columns.Contains("Export_Totals")) { dgvHisExport.Columns["Export_Totals"].Width = 120; dgvHisExport.Columns["Export_Totals"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvHisExport.Columns.Contains("Status")) { dgvHisExport.Columns["Status"].Width = 100; dgvHisExport.Columns["Status"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvHisExport.Columns.Contains("Notes")) { dgvHisExport.Columns["Notes"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; }
            if (dgvHisExport.Columns.Contains("Create_By")) { dgvHisExport.Columns["Create_By"].Width = 120; dgvHisExport.Columns["Create_By"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvHisExport.Columns.Contains("Create_Date")) { dgvHisExport.Columns["Create_Date"].Width = 150; dgvHisExport.Columns["Create_Date"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvHisExport.Columns.Contains("Update_By")) { dgvHisExport.Columns["Update_By"].Width = 120; dgvHisExport.Columns["Update_By"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvHisExport.Columns.Contains("Update_Date")) { dgvHisExport.Columns["Update_Date"].Width = 150; dgvHisExport.Columns["Update_Date"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvHisExport.Columns.Contains("Print")) { dgvHisExport.Columns["Print"].Width = 60; dgvHisExport.Columns["Print"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }

            dgvHisExport.CellContentClick -= DgvHisExport_CellContentClick;
            dgvHisExport.CellContentClick += DgvHisExport_CellContentClick;
            dgvHisExport.CellDoubleClick -= dgvHisExport_CellDoubleClick;
            dgvHisExport.CellDoubleClick += dgvHisExport_CellDoubleClick;
        }

        private void dgvHisExport_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataRow row = ((DataRowView)dgvHisExport.Rows[e.RowIndex].DataBoundItem).Row;
                frmPreviewExportWarehouse frm = new frmPreviewExportWarehouse(row);
                frm.ShowDialog();
            }
        }

        private void BtnSearch_Click(object? sender, EventArgs e)
        {
            string filter = "1=1";
            if (cboProject.SelectedValue != null) filter += $" AND From_Project_Name = '{cboProject.Text}'";
            if (!string.IsNullOrEmpty(txtSearch.Text)) filter += $" AND (Create_By LIKE '%{txtSearch.Text}%' OR From_Project_Name LIKE '%{txtSearch.Text}%')";
            if (cboStatus.SelectedItem != null) filter += $" AND Status = '{cboStatus.SelectedItem}'";
            
            // Date filter (assuming Create_Date is the column)
            filter += $" AND Create_Date  >= '{dtpFromDate.Value:yyyy-MM-dd}' AND Create_Date  <= '{dtpToDate.Value:yyyy-MM-dd}'";

            DataView dv = _dtHisExport.DefaultView;
            dv.RowFilter = filter;
            dgvHisExport.DataSource = dv;
        }

        private void BtnRefresh_Click(object? sender, EventArgs e)
        {
            cboProject.SelectedIndex = -1;
            txtSearch.Clear();
            cboStatus.SelectedIndex = -1;
            dtpFromDate.Value = DateTime.Now;
            _dtHisExport.DefaultView.RowFilter = "";
            dgvHisExport.DataSource = _dtHisExport;
        }

        private void DgvHisExport_CellContentClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var grid = sender as DataGridView;
            if (grid != null && grid.Columns[e.ColumnIndex].Name == "Print")
            {
                string exportNo = grid.Rows[e.RowIndex].Cells["Export_No"].Value?.ToString() ?? "N/A";
                MessageBox.Show($"Chức năng In phiếu xuất kho {exportNo} sẽ được cập nhật sau.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // Handle status menu item click
        private async void StatusMenu_ItemClicked(object? sender, ToolStripItemClickedEventArgs e)
        {
            string newStatus = e.ClickedItem.Text;
            await UpdateSelectedRowsStatusAsync(newStatus);
        }

        // Update status of selected rows in the grid and database
        private async Task UpdateSelectedRowsStatusAsync(string newStatus)
        {
            var rowsToUpdate = new List<int>();
            foreach (DataGridViewRow row in dgvHisExport.Rows)
            {
                if (row.Cells["Select"].Value is bool isSelected && isSelected)
                {
                    if (row.Cells["Export_ID"].Value != null)
                    {
                        rowsToUpdate.Add(Convert.ToInt32(row.Cells["Export_ID"].Value));
                        row.Cells["Status"].Value = newStatus;
                    }
                }
            }

            if (rowsToUpdate.Count == 0) return;

            string ids = string.Join(",", rowsToUpdate);
            string sql = $"UPDATE [dbo].[ExportWarehouseHeader] SET Status = @status, Update_By = @user, Update_Date = GETDATE() WHERE Export_ID IN ({ids})";

            try
            {
                using (SqlConnection conn = DatabaseHelper.GetConnection())
                {
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@status", newStatus);
                        cmd.Parameters.AddWithValue("@user", Environment.UserName);
                        await conn.OpenAsync();
                        await cmd.ExecuteNonQueryAsync();
                    }
                }
                MessageBox.Show($"Cập nhật trạng thái thành công cho {rowsToUpdate.Count} dòng.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi cập nhật trạng thái: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}