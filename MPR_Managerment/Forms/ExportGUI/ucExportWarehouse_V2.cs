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
            btnUpdateStatus.Click += (s, e) =>
            {
                _statusMenu.Show(btnUpdateStatus, new System.Drawing.Point(0, btnUpdateStatus.Height));
            };

            btnSearch.Click += BtnSearch_Click;
            btnRefresh.Click += BtnRefresh_Click;
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
                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "pxk_template.xlsx");
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
                        var ws = package.Workbook.Worksheets.Add(sheetName);

                        // Copy template structure to new sheet
                        for (int r = 1; r <= templateWs.Dimension?.Rows; r++)
                        {
                            for (int c = 1; c <= templateWs.Dimension?.Columns; c++)
                            {
                                var sourceCell = templateWs.Cells[r, c];
                                var targetCell = ws.Cells[r, c];
                                targetCell.Value = sourceCell.Value;
                                
                                // Copy style properties
                                try
                                {
                                    targetCell.Style.Font.Name = sourceCell.Style.Font.Name;
                                    targetCell.Style.Font.Size = sourceCell.Style.Font.Size;
                                    targetCell.Style.Font.Bold = sourceCell.Style.Font.Bold;
                                    targetCell.Style.Font.Italic = sourceCell.Style.Font.Italic;
                                    targetCell.Style.Fill.PatternType = sourceCell.Style.Fill.PatternType;
                                    targetCell.Style.Border.Left.Style = sourceCell.Style.Border.Left.Style;
                                    targetCell.Style.Border.Right.Style = sourceCell.Style.Border.Right.Style;
                                    targetCell.Style.Border.Top.Style = sourceCell.Style.Border.Top.Style;
                                    targetCell.Style.Border.Bottom.Style = sourceCell.Style.Border.Bottom.Style;
                                    targetCell.Style.HorizontalAlignment = sourceCell.Style.HorizontalAlignment;
                                    targetCell.Style.VerticalAlignment = sourceCell.Style.VerticalAlignment;
                                    targetCell.Style.WrapText = sourceCell.Style.WrapText;
                                    targetCell.Style.Numberformat.Format = sourceCell.Style.Numberformat.Format;
                                }
                                catch { /* Skip style copying on error */ }
                            }
                        }
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
                            // Replace placeholders
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

                                ws.Cells[currentRow, 1].Value = i + 1;
                                ws.Cells[currentRow, 2].Value = row["ID_Code"];
                                ws.Cells[currentRow, 3].Value = row["Item_Name"];
                                ws.Cells[currentRow, 4].Value = "";
                                ws.Cells[currentRow, 5].Value = row["Size"];
                                ws.Cells[currentRow, 6].Value = row["Material"];
                                ws.Cells[currentRow, 7].Value = slXuat;
                                ws.Cells[currentRow, 8].Value = row["UNIT"];
                                ws.Cells[currentRow, 9].Value = row["QC_Code"];
                                ws.Cells[currentRow, 10].Value = row["Notes"];
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
        }
    }
}
