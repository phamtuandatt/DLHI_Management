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
            
            btnSearch.Click += BtnSearch_Click;
            btnRefresh.Click += BtnRefresh_Click;
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
            dgvHisExport.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
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
                var chkCol = new DataGridViewCheckBoxColumn { Name = "Select", HeaderText = "", Width = 40, ReadOnly = false };
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
            filter += $" AND CONVERT(DATE, Create_Date) = '{dtpDate.Value:yyyy-MM-dd}'";

            DataView dv = _dtHisExport.DefaultView;
            dv.RowFilter = filter;
            dgvHisExport.DataSource = dv;
        }

        private void BtnRefresh_Click(object? sender, EventArgs e)
        {
            cboProject.SelectedIndex = -1;
            txtSearch.Clear();
            cboStatus.SelectedIndex = -1;
            dtpDate.Value = DateTime.Now;
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
    }
}