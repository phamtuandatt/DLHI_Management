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
using System.IO;
using OfficeOpenXml;
using MPR_Managerment.Models;
using MPR_Managerment.Common;
using System.DirectoryServices.ActiveDirectory;
using OfficeOpenXml.Style;

namespace MPR_Managerment.Forms.ExportGUI
{
    public partial class ucExportWarehouse_V2 : UserControl
    {
        private DataTable _dtHisExport = new DataTable();
        private ProjectService _projectService = new ProjectService();
        private string[] _statusList = new string[] { "Xác nhận", "Chưa xác nhận", "Hủy" };
        private WarehouseService _warehouseServices = new WarehouseService();

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
            Common.Common.CreateButtonPrint(btnExportExcel, "⌛Lịch sử xuất kho");
            Common.Common.CreateButtonSave_V2(btnSaveServer, "🛢 Cập nhật Server");

            // Initialize status dropdown menu
            btnUpdateStatus.Click += (s, e) =>
            {
                _statusMenu.Show(btnUpdateStatus, new System.Drawing.Point(0, btnUpdateStatus.Height));
            };

            btnSearch.Click += BtnSearch_Click;
            btnInXK.Click += btnInXK_Click;

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
            cboStatus.Items.AddRange(_statusList);
            //_statusMenu.Items.AddRange(_statusList.Select(text => new ToolStripMenuItem(text)).ToArray());
            cboStatus.SelectedIndex = -1;
        }

        private async Task LoadHisExportAsync(string project_code = "")
        {
            try
            {
                string sql = @"EXEC [dbo].[sp_GetHisExportV2]";
                if (!string.IsNullOrEmpty(project_code))
                    sql += $" @From_Project_Code = N'{project_code}'";

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
            if (dgvHisExport.Columns.Contains("Export_No")) { dgvHisExport.Columns["Export_No"].Width = 190; dgvHisExport.Columns["Export_No"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvHisExport.Columns.Contains("From_Project_Name")) { dgvHisExport.Columns["From_Project_Name"].Width = 180; dgvHisExport.Columns["From_Project_Name"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvHisExport.Columns.Contains("To_Project_Name")) { dgvHisExport.Columns["To_Project_Name"].Width = 180; dgvHisExport.Columns["To_Project_Name"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvHisExport.Columns.Contains("Export_Totals")) { dgvHisExport.Columns["Export_Totals"].Width = 120; dgvHisExport.Columns["Export_Totals"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvHisExport.Columns.Contains("Status")) { dgvHisExport.Columns["Status"].Width = 120; dgvHisExport.Columns["Status"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvHisExport.Columns.Contains("Notes")) { dgvHisExport.Columns["Notes"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; }
            if (dgvHisExport.Columns.Contains("Create_By")) { dgvHisExport.Columns["Create_By"].Width = 120; dgvHisExport.Columns["Create_By"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvHisExport.Columns.Contains("Create_Date")) { dgvHisExport.Columns["Create_Date"].Width = 150; dgvHisExport.Columns["Create_Date"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvHisExport.Columns.Contains("Update_By")) { dgvHisExport.Columns["Update_By"].Width = 120; dgvHisExport.Columns["Update_By"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvHisExport.Columns.Contains("Update_Date")) { dgvHisExport.Columns["Update_Date"].Width = 150; dgvHisExport.Columns["Update_Date"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvHisExport.Columns.Contains("IS_UPDATE")) { dgvHisExport.Columns["IS_UPDATE"].Width = 130; dgvHisExport.Columns["IS_UPDATE"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }
            if (dgvHisExport.Columns.Contains("Print")) { dgvHisExport.Columns["Print"].Width = 60; dgvHisExport.Columns["Print"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; }

            dgvHisExport.CellContentClick -= DgvHisExport_CellContentClick;
            dgvHisExport.CellContentClick += DgvHisExport_CellContentClick;
            dgvHisExport.CellDoubleClick -= dgvHisExport_CellDoubleClick;
            dgvHisExport.CellDoubleClick += dgvHisExport_CellDoubleClick;
            dgvHisExport.CellFormatting += (s, e) =>
            {
                var statusRules = new List<StringRule>
                {
                    new StringRule { Value = "Đã cập nhật", CellColor = Color.SeaGreen },
                    new StringRule { Value = "Chưa cập nhật", CellColor = Color.Red },
                };
                Common.Common.ApplyCustomFormatting(e, dgvHisExport, "IS_UPDATE", statusRules, null);

                var statusExRules = new List<StringRule>
                {
                    new StringRule { Value = "Xác nhận", CellColor = Color.SeaGreen },
                    new StringRule { Value = "Chưa xác nhận", CellColor = Color.Red },
                };
                Common.Common.ApplyCustomFormatting(e, dgvHisExport, "Status", statusExRules, null);
            };
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

        private async void BtnRefresh_Click(object? sender, EventArgs e)
        {
            try
            {
                // Clear current selections and filters
                cboProject.SelectedIndex = -1;
                txtSearch.Text = string.Empty;
                cboStatus.SelectedIndex = -1;
                dtpFromDate.Value = DateTime.Today.AddDays(-30);
                dtpToDate.Value = DateTime.Today;

                // Clear the DataTable and reload all data
                _dtHisExport.Clear();
                await LoadHisExportAsync();

                // Reset the data grid to show all data
                dgvHisExport.DataSource = _dtHisExport;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi làm mới dữ liệu: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void DgvHisExport_CellContentClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var grid = sender as DataGridView;
            if (grid != null && grid.Columns[e.ColumnIndex].Name == "Print")
            {
                int exportId = Convert.ToInt32(grid.Rows[e.RowIndex].Cells["Export_ID"].Value);
                DataRow headerRow = ((DataRowView)grid.Rows[e.RowIndex].DataBoundItem).Row;
                await ExportExcelAsync(exportId, headerRow);
            }
        }

        private async Task ExportExcelAsync(int exportId, DataRow headerRow)
        {
            try
            {
                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "pxk_template.xlsx");
                if (!File.Exists(templatePath))
                {
                    MessageBox.Show("Không tìm thấy file template!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string exportNo = headerRow["Export_No"].ToString() ?? "";
                string fromProject = headerRow["From_Project_Name"].ToString() ?? "";
                string createBy = headerRow["Create_By"].ToString() ?? "";
                DateTime createDate = headerRow["Create_Date"] != DBNull.Value ? Convert.ToDateTime(headerRow["Create_Date"]) : DateTime.Now;

                // Load details
                DataTable dtDetails = new DataTable();
                string sqlDetails = @"SELECT we.[Export_Detail_Id],we.[Export_ID],we.[Import_ID],wi.ID_Code,wi.Item_Name,wi.Material,wi.Size,we.[Qty_Export],wi.UNIT,we.[Notes],wi.QC_Code
                                      FROM [dbo].[ExportWarehouseDetail] we 
                                      INNER JOIN Warehouse_Import wi ON we.Import_ID = wi.Import_ID 
                                      WHERE we.Export_ID = @ExportID";

                using (SqlConnection conn = DatabaseHelper.GetConnection())
                {
                    using (SqlCommand cmd = new SqlCommand(sqlDetails, conn))
                    {
                        cmd.Parameters.AddWithValue("@ExportID", exportId);
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dtDetails);
                    }
                }

                if (dtDetails.Rows.Count == 0)
                {
                    MessageBox.Show("Phiếu xuất kho này không có chi tiết vật tư!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var saveDialog = new SaveFileDialog
                {
                    Title = "Lưu Phiếu Xuất Kho",
                    Filter = "Excel Files|*.xlsx",
                    FileName = $"PXK_{exportNo}_{DateTime.Now:ddMMyyyy_HHmm}",
                    InitialDirectory = Directory.Exists(@"D:\RÁC") ? @"D:\RÁC" : Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                };

                if (saveDialog.ShowDialog() != DialogResult.OK) return;

                File.Copy(templatePath, saveDialog.FileName, true);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(saveDialog.FileName)))
                {
                    var ws = package.Workbook.Worksheets[0]; // Lấy sheet "PXK"

                    for (int r = 1; r <= 10; r++)
                    {
                        for (int c = 1; c <= 10; c++)
                        {
                            if (ws.Cells[r, c].Text.Contains("<<DATE>>"))
                            {
                                ws.Cells[r, c].Value = ws.Cells[r, c].Text.Replace("<<DATE>>", createDate.ToString("dd/MM/yyyy"));
                            }
                        }
                    }

                    ReplaceCell(ws, "<<PROJECT-NAME>>", fromProject);
                    ReplaceCell(ws, "<<USER>>", createBy);

                    int startRow = 11;
                    int detailCount = dtDetails.Rows.Count;
                    decimal totalQty = 0;

                    if (detailCount > 1)
                    {
                        ws.InsertRow(startRow + 1, detailCount - 1, startRow);
                    }

                    for (int i = 0; i < detailCount; i++)
                    {
                        DataRow row = dtDetails.Rows[i];
                        int currentRow = startRow + i;

                        decimal slXuat = Convert.ToDecimal(row["Qty_Export"] != DBNull.Value ? row["Qty_Export"] : 0);
                        totalQty += slXuat;

                        ws.Cells[currentRow, 1].Value = i + 1; // Cột No (A)
                        ws.Cells[currentRow, 2].Value = row["ID_Code"]; // Cột Code (B)
                        ws.Cells[currentRow, 3].Value = row["Item_Name"]; // Cột Name (C)
                        ws.Cells[currentRow, 4].Value = ""; // Cột DWG No (D)
                        ws.Cells[currentRow, 5].Value = row["Size"]; // Cột Size (E)
                        ws.Cells[currentRow, 6].Value = row["Material"]; // Cột Grade (F)
                        ws.Cells[currentRow, 7].Value = slXuat; // Cột Q'ty (G)
                        ws.Cells[currentRow, 8].Value = row["UNIT"]; // Cột Unit (H)
                        ws.Cells[currentRow, 9].Value = row["QC_Code"]; // Cột QC_Code (I)
                        ws.Cells[currentRow, 10].Value = row["Notes"]; // Cột Notes (J)
                    }

                    int searchEndRow = startRow + detailCount + 5;
                    for (int r = startRow + detailCount; r <= searchEndRow; r++)
                    {
                        for (int c = 1; c <= 10; c++)
                        {
                            if (ws.Cells[r, c].Text.Contains("<<SUM>>"))
                            {
                                ws.Cells[r, c].Value = totalQty;
                                ws.Cells[r, c].Style.Font.Bold = true;
                                ws.Cells[r, c].Style.Numberformat.Format = "#,##0";
                            }
                        }
                    }

                    package.Save();
                }

                var result = MessageBox.Show(
                $"✅ Xuất phiếu xuất kho thành công!\nFile: {saveDialog.FileName}\n\nBạn có muốn mở file ngay không?",
                "Thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (result == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = saveDialog.FileName,
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ReplaceCell(ExcelWorksheet ws, string placeholder, string value)
        {
            for (int r = 1; r <= ws.Dimension.End.Row; r++)
                for (int c = 1; c <= ws.Dimension.End.Column; c++)
                    if (ws.Cells[r, c].Value?.ToString() == placeholder)
                        ws.Cells[r, c].Value = value;
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
                        cmd.Parameters.AddWithValue("@user", AppSession.CurrentUser?.Full_Name ?? Environment.UserName);
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

        private async void btnInXK_Click(object sender, EventArgs e)
        {
            try
            {
                // Collect selected rows
                var selectedExports = new List<(int ExportId, string ExportNo, DataRow HeaderRow)>();
                foreach (DataGridViewRow row in dgvHisExport.Rows)
                {
                    if (row.Cells["Select"].Value is bool isSelected && isSelected)
                    {
                        int exportId = Convert.ToInt32(row.Cells["Export_ID"].Value);
                        string exportNo = row.Cells["Export_No"].Value?.ToString() ?? "";
                        DataRow headerRow = ((DataRowView)row.DataBoundItem).Row;
                        selectedExports.Add((exportId, exportNo, headerRow));
                    }
                }

                if (selectedExports.Count == 0)
                {
                    MessageBox.Show("Vui lòng chọn ít nhất một phiếu xuất kho để in!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Show confirmation dialog
                string exportList = string.Join("\n• ", selectedExports.Select(x => x.ExportNo));
                var result = MessageBox.Show(
                    $"Bạn có muốn in {selectedExports.Count} phiếu xuất kho sau?\n\n• {exportList}",
                    "Xác nhận in phiếu",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result != DialogResult.Yes) return;

                // Export all selected rows to one file with multiple sheets
                await ExportMultipleSheetExcelAsync(selectedExports);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task ExportMultipleSheetExcelAsync(List<(int ExportId, string ExportNo, DataRow HeaderRow)> selectedExports)
        {
            try
            {
                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "pxk_ad_template.xlsx");
                if (!File.Exists(templatePath))
                {
                    MessageBox.Show("Không tìm thấy file template!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var saveDialog = new SaveFileDialog
                {
                    Title = "Lưu Phiếu Xuất Kho",
                    Filter = "Excel Files|*.xlsx",
                    FileName = $"PXK_Batch_{DateTime.Now:ddMMyyyy_HHmm}",
                    InitialDirectory = Directory.Exists(@"D:\RÁC") ? @"D:\RÁC" : Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                };

                if (saveDialog.ShowDialog() != DialogResult.OK) return;

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (var package = new ExcelPackage())
                {
                    int sheetIndex = 0;
                    foreach (var (exportId, exportNo, headerRow) in selectedExports)
                    {
                        sheetIndex++;

                        // Load template for this sheet
                        var templatePackage = new ExcelPackage(new FileInfo(templatePath));
                        var templateWs = templatePackage.Workbook.Worksheets[0];

                        // Create new sheet with safe name
                        string sheetName = exportNo.Length > 31 ? exportNo.Substring(0, 31) : exportNo;
                        sheetName = System.Text.RegularExpressions.Regex.Replace(sheetName, @"[\\\/\?\*\[\]]", "_");
                        var ws = package.Workbook.Worksheets.Add(sheetName, templateWs);
                        templatePackage.Dispose();

                        // Get export data
                        string exportNoVal = headerRow["Export_No"].ToString() ?? "";
                        string fromProject = headerRow["From_Project_Name"].ToString() ?? "";
                        string createBy = headerRow["Create_By"].ToString() ?? "";
                        DateTime createDate = headerRow["Create_Date"] != DBNull.Value ? Convert.ToDateTime(headerRow["Create_Date"]) : DateTime.Now;

                        // Load details
                        DataTable dtDetails = new DataTable();
                        string sqlDetails = @"SELECT we.[Export_Detail_Id],we.[Export_ID],we.[Import_ID],wi.ID_Code,wi.Item_Name,wi.Material,wi.Size,we.[Qty_Export],wi.UNIT,we.[Notes],wi.QC_Code
                                              FROM [dbo].[ExportWarehouseDetail] we 
                                              INNER JOIN Warehouse_Import wi ON we.Import_ID = wi.Import_ID 
                                              WHERE we.Export_ID = @ExportID";

                        using (SqlConnection conn = DatabaseHelper.GetConnection())
                        {
                            using (SqlCommand cmd = new SqlCommand(sqlDetails, conn))
                            {
                                cmd.Parameters.AddWithValue("@ExportID", exportId);
                                SqlDataAdapter da = new SqlDataAdapter(cmd);
                                da.Fill(dtDetails);
                            }
                        }

                        if (dtDetails.Rows.Count > 0)
                        {
                            //// Replace placeholders
                            //for (int r = 1; r <= 10; r++)
                            //{
                            //    for (int c = 1; c <= 10; c++)
                            //    {
                            //        if (ws.Cells[r, c].Text.Contains("<<DATE>>"))
                            //        {
                            //            ws.Cells[r, c].Value = ws.Cells[r, c].Text.Replace("<<DATE>>", createDate.ToString("dd/MM/yyyy"));
                            //        }
                            //    }
                            //}

                            ReplaceCell(ws, "<<EXPORT_NO>>", exportNo);
                            ReplaceCell(ws, "<<DATE>>", createDate.ToString("dd/MM/yyyy"));
                            ReplaceCell(ws, "<<USER>>", createBy);

                            int startRow = 8;
                            int detailCount = dtDetails.Rows.Count;
                            decimal totalQty = 0;

                            if (detailCount > 1)
                            {
                                ws.InsertRow(startRow + 1, detailCount - 1, startRow);
                            }

                            for (int i = 0; i < detailCount; i++)
                            {
                                DataRow row = dtDetails.Rows[i];
                                int currentRow = startRow + i;

                                decimal slXuat = Convert.ToDecimal(row["Qty_Export"] != DBNull.Value ? row["Qty_Export"] : 0);
                                totalQty += slXuat;

                                ws.Cells[currentRow, 1].Value = i + 1;

                                ws.Cells[currentRow, 2, currentRow, 4].Merge = true;
                                ws.Cells[currentRow, 2].Value = row["ID_Code"];

                                ws.Cells[currentRow, 5].Value = row["Item_Name"];
                                ws.Cells[currentRow, 6].Value = $"{row["Size"]} - {row["Material"]}";
                                ws.Cells[currentRow, 7].Value = slXuat;
                                ws.Cells[currentRow, 8].Value = row["UNIT"];
                                ws.Cells[currentRow, 9].Value = row["Notes"];
                            }

                            int searchEndRow = startRow + detailCount + 5;
                            for (int r = startRow + detailCount; r <= searchEndRow; r++)
                            {
                                for (int c = 1; c <= 10; c++)
                                {
                                    if (ws.Cells[r, c].Text.Contains("<<SUM>>"))
                                    {
                                        ws.Cells[r, c].Value = totalQty;
                                        ws.Cells[r, c].Style.Font.Bold = true;
                                        ws.Cells[r, c].Style.Numberformat.Format = "#,##0";
                                    }
                                }
                            }
                        }
                    }

                    package.SaveAs(new FileInfo(saveDialog.FileName));
                }

                var openResult = MessageBox.Show(
                $"✅ Xuất {selectedExports.Count} phiếu xuất kho thành công!\nFile: {saveDialog.FileName}\n\nBạn có muốn mở file ngay không?",
                "Thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (openResult == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = saveDialog.FileName,
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAddXK_Click(object sender, EventArgs e)
        {
            frmPreviewExportWarehouse frm = new frmPreviewExportWarehouse(true);
            frm.ShowDialog();
            btnRefresh.PerformClick();
        }

        private void btnUpdateStatus_Click(object sender, EventArgs e)
        {

        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            DataTable dt = _warehouseServices.GetExportWarehouseForExportExcel(fromProjectCode: string.Empty, fromDate: dtpFromDate.Value, toDate: dtpToDate.Value);
            List<ExportDetailModel> exportDetailsList = _warehouseServices.ConvertDataTableToList(dt);

            ExportExportDetailsToExcel(exportDetailsList);
        }

        private async Task ExportToOvppTemplateAsync()
        {
            try
            {
                // Get the template file path
                string templatePath = @"D:\ERP_Final_Management\MPR_Managerment\Templates\ovpp-template.xlsx";
                if (!File.Exists(templatePath))
                {
                    MessageBox.Show("Không tìm thấy file template (ovpp-template.xlsx)!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Get data from service
                DataTable dt = _warehouseServices.GetExportWarehouseForExportExcel(
                    fromProjectCode: string.Empty,
                    fromDate: dtpFromDate.Value,
                    toDate: dtpToDate.Value);

                if (dt == null || dt.Rows.Count == 0)
                {
                    MessageBox.Show("Không có dữ liệu để xuất!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Show Save File Dialog
                var saveDialog = new SaveFileDialog
                {
                    Title = "Lưu Phiếu Xuất Kho Theo Template",
                    Filter = "Excel Files|*.xlsx",
                    FileName = $"XuatKho_OVPP_{DateTime.Now:yyyyMMdd_HHmmss}",
                    InitialDirectory = Directory.Exists(@"D:\RÁC") ? @"D:\RÁC" : Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                };

                if (saveDialog.ShowDialog() != DialogResult.OK) return;

                // Set EPPlus license
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                // Copy template to destination
                File.Copy(templatePath, saveDialog.FileName, true);

                // Open and modify the Excel file
                using (var package = new ExcelPackage(new FileInfo(saveDialog.FileName)))
                {
                    var ws = package.Workbook.Worksheets["Sheet1"];

                    int startRow = 8;
                    int rowIndex = 0;

                    // Populate data starting from row 8
                    foreach (DataRow row in dt.Rows)
                    {
                        int currentRow = startRow + rowIndex;

                        // Get values from DataTable
                        string idCode = row["ID_Code"]?.ToString() ?? "";
                        string itemName = row["Item_Name"]?.ToString() ?? "";
                        string description = row["Notes"]?.ToString() ?? "";
                        decimal qtyExport = Convert.ToDecimal(row["Qty_Export"] ?? 0);
                        string unit = row["UNIT"]?.ToString() ?? "";
                        string remark = row["Notes"]?.ToString() ?? "";

                        // Write to cells in Sheet1
                        // Assuming columns: A=ID_Code, B=Item_Name, C=Description, D=Qty_Export, E=UNIT, F=Notes/Remark
                        ws.Cells[currentRow, 1].Value = idCode;        // Column A
                        ws.Cells[currentRow, 2].Value = itemName;      // Column B
                        ws.Cells[currentRow, 3].Value = description;   // Column C
                        ws.Cells[currentRow, 4].Value = qtyExport;     // Column D
                        ws.Cells[currentRow, 5].Value = unit;          // Column E
                        ws.Cells[currentRow, 6].Value = remark;        // Column F

                        // Format quantity column
                        ws.Cells[currentRow, 4].Style.Numberformat.Format = "#,##0.00";

                        rowIndex++;
                    }

                    // Save the package
                    package.Save();
                }

                var result = MessageBox.Show(
                    $"✅ Xuất dữ liệu thành công!\nFile: {saveDialog.FileName}\n\nBạn có muốn mở file ngay không?",
                    "Thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                if (result == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = saveDialog.FileName,
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi xuất file Excel: {ex.Message}\n\n{ex.StackTrace}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void ExportExportDetailsToExcel(List<ExportDetailModel> dataList)
        {
            if (dataList == null || dataList.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu để xuất Excel!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Thiết lập License cho EPPlus (Bắt buộc từ bản 5.x trở lên)
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                sfd.FileName = $"BaoCao_XuatKho_ChuaCapNhat_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                sfd.Title = "Chọn nơi lưu file Excel báo cáo";
                sfd.InitialDirectory = Directory.Exists(@"D:\RÁC") ? @"D:\RÁC" : Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        FileInfo fileInfo = new FileInfo(sfd.FileName);
                        using (ExcelPackage package = new ExcelPackage(fileInfo))
                        {
                            // Tạo một Worksheet mới
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Chi Tiết Xuất Kho");

                            // 1. Tạo Tiêu đề Báo cáo (Header Title)
                            worksheet.Cells["A1:P1"].Merge = true;
                            worksheet.Cells["A1"].Value = "DANH SÁCH CHI TIẾT XUẤT KHO CHƯA CẬP NHẬT";
                            worksheet.Cells["A1"].Style.Font.Size = 16;
                            worksheet.Cells["A1"].Style.Font.Bold = true;
                            worksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Row(1).Height = 35;

                            // 2. Định nghĩa các Tiêu đề cột (Grid Headers)
                            string[] headers = new string[] {
                                "STT", "Mã Chi Tiết", "ID Xuất", "Số Phiếu Xuất", "Mã Chứng Từ",
                                "Mã Dự Án Nguồn", "Tên Dự Án Nguồn", "Tên Dự Án Đích", "ID Nhập",
                                "Mã Vật Tư", "Tên Vật Tư", "Chất Liệu", "Kích Thước",
                                "Số Lượng Xuất", "Đơn Vị Tính", "Ghi Chú", "Ngày Tạo"
                            };

                            for (int i = 0; i < headers.Length; i++)
                            {
                                var cell = worksheet.Cells[3, i + 1];
                                cell.Value = headers[i];
                                cell.Style.Font.Bold = true;
                                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(23, 162, 184)); // Màu Teal/Xanh dịu mắt
                                cell.Style.Font.Color.SetColor(Color.White);
                                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            }
                            worksheet.Row(3).Height = 26;

                            // 3. Đổ dữ liệu từ List<Model> vào các dòng
                            int startRow = 4;
                            int stt = 1;

                            foreach (var item in dataList)
                            {
                                worksheet.Cells[startRow, 1].Value = stt++;
                                worksheet.Cells[startRow, 2].Value = item.Export_Detail_Id;
                                worksheet.Cells[startRow, 3].Value = item.Export_ID;
                                worksheet.Cells[startRow, 4].Value = item.Export_No;
                                worksheet.Cells[startRow, 5].Value = item.Export_Number;
                                worksheet.Cells[startRow, 6].Value = item.From_Project_Code;
                                worksheet.Cells[startRow, 7].Value = item.From_Project_Name;
                                worksheet.Cells[startRow, 8].Value = item.To_Project_Name;
                                worksheet.Cells[startRow, 9].Value = item.Import_ID;
                                worksheet.Cells[startRow, 10].Value = item.ID_Code;
                                worksheet.Cells[startRow, 11].Value = item.Item_Name;
                                worksheet.Cells[startRow, 12].Value = item.Material;
                                worksheet.Cells[startRow, 13].Value = item.Size;

                                // Định dạng hiển thị số lượng
                                var qtyCell = worksheet.Cells[startRow, 14];
                                qtyCell.Value = item.Qty_Export;
                                qtyCell.Style.Numberformat.Format = "#,##0.00";

                                worksheet.Cells[startRow, 15].Value = item.UNIT;
                                worksheet.Cells[startRow, 16].Value = item.Notes;

                                // Định dạng hiển thị ngày tháng năm
                                var dateCell = worksheet.Cells[startRow, 17];
                                if (item.Create_Date.HasValue)
                                {
                                    dateCell.Value = item.Create_Date.Value;
                                    dateCell.Style.Numberformat.Format = "dd/MM/yyyy HH:mm";
                                }

                                // Căn chỉnh lề cho các trường số/ID/Mã code
                                worksheet.Cells[startRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                worksheet.Cells[startRow, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                worksheet.Cells[startRow, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                worksheet.Cells[startRow, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                worksheet.Cells[startRow, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                worksheet.Cells[startRow, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                worksheet.Cells[startRow, 15].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                worksheet.Cells[startRow, 17].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                startRow++;
                            }

                            // 4. Định dạng đường viền (Borders) cho toàn bộ vùng dữ liệu
                            int endRow = startRow - 1;
                            using (var range = worksheet.Cells[3, 1, endRow, headers.Length])
                            {
                                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                range.Style.Border.Top.Color.SetColor(Color.LightGray);
                                range.Style.Border.Bottom.Color.SetColor(Color.LightGray);
                                range.Style.Border.Left.Color.SetColor(Color.LightGray);
                                range.Style.Border.Right.Color.SetColor(Color.LightGray);
                            }

                            // 5. Tự động căn chỉnh độ rộng cột theo nội dung (Auto-fit columns)
                            worksheet.Cells[3, 1, endRow, headers.Length].AutoFitColumns();

                            // Lưu file lại
                            package.Save();
                        }

                        var result = MessageBox.Show(
                                $"✅ Xuất phiếu nhập kho thành công!\nFile: {sfd.FileName}\n\nBạn có muốn mở file ngay không?",
                                "Thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                        if (result == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                            {
                                FileName = sfd.FileName,
                                UseShellExecute = true
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Lỗi khi xuất file Excel: {ex.Message}", "Lỗi hệ thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnSaveServer_Click(object sender, EventArgs e)
        {
            try
            {
                // Collect selected rows
                var selectedExports = new List<(int ExportId, string ExportNo, DataRow HeaderRow)>();
                foreach (DataGridViewRow row in dgvHisExport.Rows)
                {
                    if (row.Cells["Select"].Value is bool isSelected && isSelected)
                    {
                        int exportId = Convert.ToInt32(row.Cells["Export_ID"].Value);
                        string exportNo = row.Cells["Export_No"].Value?.ToString() ?? "";
                        DataRow headerRow = ((DataRowView)row.DataBoundItem).Row;
                        selectedExports.Add((exportId, exportNo, headerRow));
                    }
                }

                if (selectedExports.Count == 0)
                {
                    MessageBox.Show("Vui lòng chọn ít nhất một phiếu xuất kho để cập nhật!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var rs = 0;
                foreach (var item_id in selectedExports)
                {
                    var _dtDetails = LoadDetails(item_id.ExportId);
                    foreach (DataRow item in _dtDetails.Rows)
                    {
                        var model = new ExportWarehouseDetailModel()
                        {
                            Import_Id = Convert.ToInt32(item["Import_ID"].ToString() ?? "0"),
                            Qty_Export = Convert.ToInt32(item["Qty_Export"].ToString() ?? "0"),
                            Notes = item["Notes"]?.ToString()?.Trim() ?? "",
                        };

                        var isSave = _warehouseServices.InsertExportWarehouseHeader(model, item_id.ExportNo);
                        rs++;
                    }
                    if (rs > 0)
                    {
                        _warehouseServices.UpdateStatusExportWarehouseHeader("Xác nhận", item_id.ExportId);
                        MessageBox.Show("Đã cập nhật toàn bộ Phiếu Xuất và Chi Tiết lên Server thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show(" Cập nhật toàn bộ Phiếu Xuất và Chi Tiết lên Server không thành công!", "Thất bại", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ExportIdCodeListFromDatabase(DataTable dtDetail)
        {
            try
            {
                if (dtDetail == null || dtDetail.Rows.Count == 0)
                {
                    MessageBox.Show("Không có dữ liệu để xuất!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string templatePath = @"D:\ERP_Final_Management\MPR_Managerment\Templates\pxk_ad_template.xlsx";
                if (!File.Exists(templatePath))
                {
                    MessageBox.Show("Không tìm thấy file template (pxk_ad_template.xlsx)!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var saveDialog = new SaveFileDialog
                {
                    Title = "Lưu Phiếu Xuất Kho",
                    Filter = "Excel Files|*.xlsx",
                    FileName = $"PXK_AD_{DateTime.Now:ddMMyyyy_HHmm}",
                    InitialDirectory = Directory.Exists(@"D:\RÁC") ? @"D:\RÁC" : Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                };

                if (saveDialog.ShowDialog() != DialogResult.OK) return;

                File.Copy(templatePath, saveDialog.FileName, true);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(saveDialog.FileName)))
                {
                    var ws = package.Workbook.Worksheets[0];

                    // Replace header placeholders
                    for (int r = 1; r <= 10; r++)
                    {
                        for (int c = 1; c <= 10; c++)
                        {
                            if (ws.Cells[r, c].Text.Contains("<<DATE>>"))
                            {
                                ws.Cells[r, c].Value = ws.Cells[r, c].Text.Replace("<<DATE>>", DateTime.Now.ToString("dd/MM/yyyy"));
                            }
                        }
                    }

                    int startRow = 11;
                    int detailCount = dtDetail.Rows.Count;
                    decimal totalQty = 0;

                    if (detailCount > 1)
                    {
                        ws.InsertRow(startRow + 1, detailCount - 1, startRow);
                    }

                    for (int i = 0; i < detailCount; i++)
                    {
                        DataRow row = dtDetail.Rows[i];
                        int currentRow = startRow + i;

                        decimal slXuat = 0;
                        if (dtDetail.Columns.Contains("Qty_Export") && row["Qty_Export"] != DBNull.Value)
                            slXuat = Convert.ToDecimal(row["Qty_Export"]);
                        totalQty += slXuat;

                        ws.Cells[currentRow, 1].Value = i + 1; // No
                        ws.Cells[currentRow, 2].Value = dtDetail.Columns.Contains("ID_Code") ? row["ID_Code"] : ""; // ID Code
                        ws.Cells[currentRow, 3].Value = dtDetail.Columns.Contains("Item_Name") ? row["Item_Name"] : ""; // Item Name
                        ws.Cells[currentRow, 4].Value = ""; // DWG No
                        ws.Cells[currentRow, 5].Value = dtDetail.Columns.Contains("Size") ? row["Size"] : ""; // Size
                        ws.Cells[currentRow, 6].Value = dtDetail.Columns.Contains("Material") ? row["Material"] : ""; // Material/Grade
                        ws.Cells[currentRow, 7].Value = slXuat; // Qty
                        ws.Cells[currentRow, 8].Value = dtDetail.Columns.Contains("UNIT") ? row["UNIT"] : ""; // Unit
                        ws.Cells[currentRow, 9].Value = dtDetail.Columns.Contains("QC_Code") ? row["QC_Code"] : ""; // QC Code
                        ws.Cells[currentRow, 10].Value = dtDetail.Columns.Contains("Notes") ? row["Notes"] : ""; // Notes
                    }

                    // Replace <<SUM>> placeholder
                    int searchEndRow = startRow + detailCount + 5;
                    for (int r = startRow + detailCount; r <= searchEndRow; r++)
                    {
                        for (int c = 1; c <= 10; c++)
                        {
                            if (ws.Cells[r, c].Text.Contains("<<SUM>>"))
                            {
                                ws.Cells[r, c].Value = totalQty;
                                ws.Cells[r, c].Style.Font.Bold = true;
                                ws.Cells[r, c].Style.Numberformat.Format = "#,##0";
                            }
                        }
                    }

                    package.Save();
                }

                var result = MessageBox.Show(
                    $"✅ Xuất dữ liệu thành công!\nFile: {saveDialog.FileName}\n\nBạn có muốn mở file ngay không?",
                    "Thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (result == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = saveDialog.FileName,
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi xuất file Excel: {ex.Message}", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private DataTable LoadDetails(int export_Id)
        {
            var _dtDetails = new DataTable();
            string sql = @"SELECT 
                            we.[Export_Detail_Id],
                            we.[Export_ID],
                            we.[Import_ID],
                            wi.ID_Code,
                            wi.Item_Name,
                            wi.Material,
                            wi.Size,
                            we.[Qty_Export],
                            wi.UNIT,
                            we.[Notes],
                            -- Lấy cột số lượng tồn kho hiện tại, dùng ISNULL để tránh trả về NULL nếu chưa có dữ liệu tồn
                            ISNULL(stk.Qty_Stock, 0) AS [Qty_Stock]
                        FROM [dbo].[ExportWarehouseDetail] we 
                        INNER JOIN dbo.Warehouse_Import wi ON we.Import_ID = wi.Import_ID 
                        -- Join với View tính toán tồn kho động của hệ thống
                        LEFT OUTER JOIN dbo.vw_Warehouse_Stock stk ON we.Import_ID = stk.Import_ID
                        WHERE we.Export_ID = @ExportID;";

            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@ExportID", export_Id);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(_dtDetails);
                    return _dtDetails;
                }
            }
        }
    }
}
