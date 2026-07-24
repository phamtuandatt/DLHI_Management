using MPR_Managerment.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using Microsoft.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MPR_Managerment.Forms.WarehouseGUI
{
    public partial class ucTrackDelivery : UserControl
    {
        private readonly CommonServices _dbService;
        private readonly POService _poService = new POService();

        // Track modified rows (PO_Detail_ID -> new Received_Qty)
        private Dictionary<int, decimal> _modifiedRows = new Dictionary<int, decimal>();
        private object _oldCellValue = null;

        public ucTrackDelivery()
        {
            InitializeComponent();
            _dbService = new CommonServices();
            this.Load += ucTrackDelivery_Load;
        }

        private void ucTrackDelivery_Load(object sender, EventArgs e)
        {
            LoadSuppliers();
            LoadProjectCodes();
            SetDefaultDates();
            LoadTrackDeliveryData();
        }

        /// <summary>
        /// Load supplier names into ComboBox for filtering
        /// </summary>
        private void LoadSuppliers()
        {
            try
            {
                string query = "SELECT Supplier_ID, Company_Name FROM Suppliers ORDER BY Company_Name";
                DataTable dt = _dbService.ExecuteQuery(query);

                // Insert "All" option at position 0
                DataRow allRow = dt.NewRow();
                allRow["Supplier_ID"] = 0;
                allRow["Company_Name"] = "-- All Suppliers --";
                dt.Rows.InsertAt(allRow, 0);

                cboSupplier.DataSource = dt;
                cboSupplier.DisplayMember = "Company_Name";
                cboSupplier.ValueMember = "Supplier_ID";
                cboSupplier.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading suppliers: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Load project code list into ComboBox for filtering
        /// </summary>
        private void LoadProjectCodes()
        {
            try
            {
                string query = "SELECT DISTINCT Project_Code FROM Warehouse_Import WHERE Project_Code IS NOT NULL AND Project_Code <> '' ORDER BY Project_Code";
                DataTable dt = _dbService.ExecuteQuery(query);

                cboProjectCode.Items.Clear();
                cboProjectCode.Items.Add("-- All Projects --");

                foreach (DataRow row in dt.Rows)
                {
                    cboProjectCode.Items.Add(row["Project_Code"].ToString());
                }

                cboProjectCode.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading project codes: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Set default date range (FromDate = first day of current month, ToDate = today)
        /// </summary>
        private void SetDefaultDates()
        {
            dtpFromDate.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            dtpToDate.Value = DateTime.Now;
        }

        /// <summary>
        /// Load delivery tracking data by calling stored procedure sp_GetPODeliveryStatusReport
        /// </summary>
        private void LoadTrackDeliveryData(int? supplierId = null, DateTime? fromDate = null, DateTime? toDate = null, string projectCode = null)
        {
            try
            {
                DataTable dt = _dbService.ExecuteStoredProcedure(
                    "sp_GetPODeliveryStatusReport",
                    new SqlParameter("@SupplierInput", supplierId.HasValue && supplierId.Value > 0 ? (object)supplierId.Value : DBNull.Value),
                    new SqlParameter("@FromDate", fromDate.HasValue ? (object)fromDate.Value.Date : DBNull.Value),
                    new SqlParameter("@ToDate", toDate.HasValue ? (object)toDate.Value.Date : DBNull.Value),
                    new SqlParameter("@ProjectCode", (object)projectCode ?? DBNull.Value)
                );

                // Clear modified rows tracking when reloading
                _modifiedRows.Clear();

                dgvTrackDelivery.DataSource = dt;
                FormatGrid();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading delivery data: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Format DataGridView columns with proper headers and visibility
        /// </summary>
        private void FormatGrid()
        {
            if (dgvTrackDelivery.Columns.Count == 0) return;

            // Make grid editable only for Received_Qty column
            dgvTrackDelivery.ReadOnly = false;

            // Set font size 12 for the entire grid
            dgvTrackDelivery.DefaultCellStyle.Font = new Font("Segoe UI", 9);
            dgvTrackDelivery.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);

            // Hide columns containing ID and set ReadOnly for all columns except Received_Qty
            foreach (DataGridViewColumn col in dgvTrackDelivery.Columns)
            {
                string colName = col.Name.ToUpper();
                if (colName.Contains("ID") || colName == "ID")
                {
                    col.Visible = false;
                }

                // Make only Received_Qty editable
                col.ReadOnly = (col.Name != "Received_Qty");
            }

            // Highlight the editable column
            if (dgvTrackDelivery.Columns.Contains("Received_Qty"))
            {
                dgvTrackDelivery.Columns["Received_Qty"].DefaultCellStyle.BackColor = Color.FromArgb(255, 255, 220);
                dgvTrackDelivery.Columns["Received_Qty"].ReadOnly = false;
            }

            // Set column headers
            SetColumnHeader("Short_Name", "Supplier Name");
            SetColumnHeader("PONo", "PO No");
            SetColumnHeader("ProjectCode", "Project Code");
            SetColumnHeader("Item_Name", "Item Name");
            SetColumnHeader("Material", "Material");
            SetColumnHeader("Size", "Size");
            SetColumnHeader("UNIT", "Unit");
            SetColumnHeader("Qty_Per_Sheet", "Qty PO");
            SetColumnHeader("Received_Qty", "Qty Received ✏️");
            SetColumnHeader("Qty_Remaining", "Qty Remaining");
            SetColumnHeader("Status_Delivery", "Delivery Status");
            SetColumnHeader("Import_No", "Import No");
            SetColumnHeader("Import_Date", "Import Date");
            SetColumnHeader("Weight_kg", "Weight (kg)");
            SetColumnHeader("ID_Code", "ID Code");
            SetColumnHeader("WorkorderNo", "Workorder No");

            // Format date columns
            if (dgvTrackDelivery.Columns.Contains("Import_Date"))
                dgvTrackDelivery.Columns["Import_Date"].DefaultCellStyle.Format = "dd/MM/yyyy";

            // Format numeric columns
            if (dgvTrackDelivery.Columns.Contains("Qty_Per_Sheet"))
                dgvTrackDelivery.Columns["Qty_Per_Sheet"].DefaultCellStyle.Format = "N2";
            if (dgvTrackDelivery.Columns.Contains("Received_Qty"))
                dgvTrackDelivery.Columns["Received_Qty"].DefaultCellStyle.Format = "N2";
            if (dgvTrackDelivery.Columns.Contains("Qty_Remaining"))
                dgvTrackDelivery.Columns["Qty_Remaining"].DefaultCellStyle.Format = "N2";
            if (dgvTrackDelivery.Columns.Contains("Weight_kg"))
                dgvTrackDelivery.Columns["Weight_kg"].DefaultCellStyle.Format = "N2";

            // Style the grid
            dgvTrackDelivery.EnableHeadersVisualStyles = false;
            dgvTrackDelivery.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 122, 204);
            dgvTrackDelivery.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvTrackDelivery.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvTrackDelivery.DefaultCellStyle.SelectionBackColor = Color.FromArgb(204, 232, 255);
            dgvTrackDelivery.DefaultCellStyle.SelectionForeColor = Color.Black;

            // Wire up cell edit events (remove first to avoid duplicates)
            dgvTrackDelivery.CellBeginEdit -= DgvTrackDelivery_CellBeginEdit;
            dgvTrackDelivery.CellBeginEdit += DgvTrackDelivery_CellBeginEdit;
            dgvTrackDelivery.CellEndEdit -= DgvTrackDelivery_CellEndEdit;
            dgvTrackDelivery.CellEndEdit += DgvTrackDelivery_CellEndEdit;
            dgvTrackDelivery.CellFormatting -= DgvTrackDelivery_CellFormatting;
            dgvTrackDelivery.CellFormatting += DgvTrackDelivery_CellFormatting;
        }

        /// <summary>
        /// Store old value before editing begins
        /// </summary>
        private void DgvTrackDelivery_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            _oldCellValue = dgvTrackDelivery.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
        }

        /// <summary>
        /// Handle cell edit completion - validate and track changes
        /// </summary>
        private void DgvTrackDelivery_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (dgvTrackDelivery.Columns[e.ColumnIndex].Name != "Received_Qty") return;

            var row = dgvTrackDelivery.Rows[e.RowIndex];
            var cell = row.Cells[e.ColumnIndex];

            if (cell.Value == null || !decimal.TryParse(cell.Value.ToString(), out decimal newQty))
            {
                cell.Value = _oldCellValue;
                return;
            }

            // Validate: Received_Qty cannot exceed Qty_Per_Sheet
            decimal qtyPO = 0;
            if (row.Cells["Qty_Per_Sheet"].Value != null)
                decimal.TryParse(row.Cells["Qty_Per_Sheet"].Value.ToString(), out qtyPO);

            if (newQty < 0)
            {
                MessageBox.Show("Received quantity cannot be negative!", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cell.Value = _oldCellValue;
                return;
            }

            if (newQty > qtyPO)
            {
                MessageBox.Show($"Received quantity ({newQty:N2}) cannot exceed PO quantity ({qtyPO:N2})!",
                    "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cell.Value = _oldCellValue;
                return;
            }

            //// Update Qty_Remaining and DeliveryStatus in the grid
            //decimal remaining = qtyPO - newQty;
            //row.Cells["Qty_Remaining"].Value = remaining;

            // Determine delivery status
            string status;
            if (newQty <= 0)
                status = "Chưa giao hàng";
            else if (newQty >= qtyPO)
                status = "Đã giao hàng đủ";
            else
                status = "Đã Giao một phần";
            row.Cells["Status_Delivery"].Value = status;

            // Track modified row by PO_Detail_ID
            if (dgvTrackDelivery.Columns.Contains("PO_Detail_ID"))
            {
                var poDetailIdValue = row.Cells["PO_Detail_ID"].Value;
                if (poDetailIdValue != null && poDetailIdValue != DBNull.Value)
                {
                    int poDetailId = Convert.ToInt32(poDetailIdValue);
                    _modifiedRows[poDetailId] = newQty;

                    // Highlight modified row
                    row.DefaultCellStyle.BackColor = Color.FromArgb(255, 255, 200);
                }
            }
        }

        /// <summary>
        /// Color-code the DeliveryStatus column
        /// </summary>
        private void DgvTrackDelivery_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (!dgvTrackDelivery.Columns.Contains("Status_Delivery")) return;
            if (e.ColumnIndex != dgvTrackDelivery.Columns["Status_Delivery"].Index) return;

            string status = e.Value?.ToString() ?? "";
            switch (status)
            {
                case "Đã giao hàng đủ":
                    e.CellStyle.ForeColor = Color.White;
                    e.CellStyle.BackColor = Color.SeaGreen;
                    break;
                case "Đã giao một phần":
                    e.CellStyle.ForeColor = Color.Black;
                    e.CellStyle.BackColor = Color.Orange;
                    break;
                case "Chưa giao hàng":
                    e.CellStyle.ForeColor = Color.White;
                    e.CellStyle.BackColor = Color.Crimson;
                    break;
            }
        }

        /// <summary>
        /// Helper method to set column header text safely
        /// </summary>
        private void SetColumnHeader(string columnName, string headerText)
        {
            if (dgvTrackDelivery.Columns.Contains(columnName))
                dgvTrackDelivery.Columns[columnName].HeaderText = headerText;
        }

        /// <summary>
        /// Search button click - apply filters and reload data
        /// </summary>
        private void btnSearch_Click(object sender, EventArgs e)
        {
            int? supplierId = null;
            if (cboSupplier.SelectedValue != null && Convert.ToInt32(cboSupplier.SelectedValue) > 0)
            {
                supplierId = Convert.ToInt32(cboSupplier.SelectedValue);
            }

            string projectCode = null;
            if (cboProjectCode.SelectedIndex > 0)
            {
                projectCode = cboProjectCode.SelectedItem.ToString();
            }

            DateTime? fromDate = dtpFromDate.Value;
            DateTime? toDate = dtpToDate.Value;

            // Validate date range
            if (fromDate > toDate)
            {
                MessageBox.Show("'From Date' cannot be later than 'To Date'. Please adjust the date range.",
                    "Invalid Date Range", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            LoadTrackDeliveryData(supplierId, fromDate, toDate, projectCode);
        }

        /// <summary>
        /// Clear button click - reset all filters and reload all data
        /// </summary>
        private void btnClear_Click(object sender, EventArgs e)
        {
            cboSupplier.SelectedIndex = 0;
            cboProjectCode.SelectedIndex = 0;
            SetDefaultDates();
            LoadTrackDeliveryData();
        }

        /// <summary>
        /// Save button click - save all modified Received_Qty values to database
        /// </summary>
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (_modifiedRows.Count == 0)
            {
                MessageBox.Show("No changes to save.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (MessageBox.Show($"Save {_modifiedRows.Count} modified row(s)?", "Confirm Save",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            try
            {
                int savedCount = 0;
                foreach (var kvp in _modifiedRows)
                {
                    int poDetailId = kvp.Key;
                    decimal receivedQty = kvp.Value;

                    // Determine delivery status
                    var statusDelivery = false;
                    decimal qtyPO = GetQtyPOByDetailId(poDetailId);
                    if (receivedQty >= qtyPO && qtyPO > 0)
                        statusDelivery = true;
                    //else if (receivedQty > 0)
                    //    statusDelivery = "Partial";

                    // Call existing stored procedure to update
                    _poService.UpdateReceiveQtyPODetail(
                        new Models.PODetail
                        {
                            PO_Detail_ID = poDetailId,
                            Received_Qty = receivedQty,
                            Status_Delivery = statusDelivery
                        },
                        DateTime.Now.ToString("yyyy-MM-dd")
                    );

                    savedCount++;
                }

                _modifiedRows.Clear();
                MessageBox.Show($"✅ Successfully saved {savedCount} row(s)!", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Reload data to reflect changes
                btnSearch_Click(sender, e);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving data: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Get Qty_Per_Sheet for a PO_Detail_ID from the current grid data
        /// </summary>
        private decimal GetQtyPOByDetailId(int poDetailId)
        {
            if (!dgvTrackDelivery.Columns.Contains("PO_Detail_ID") || !dgvTrackDelivery.Columns.Contains("Qty_Per_Sheet"))
                return 0;

            foreach (DataGridViewRow row in dgvTrackDelivery.Rows)
            {
                if (row.IsNewRow) continue;
                var idVal = row.Cells["PO_Detail_ID"].Value;
                if (idVal != null && idVal != DBNull.Value && Convert.ToInt32(idVal) == poDetailId)
                {
                    var qtyVal = row.Cells["Qty_Per_Sheet"].Value;
                    if (qtyVal != null && qtyVal != DBNull.Value)
                        return Convert.ToDecimal(qtyVal);
                }
            }
            return 0;
        }
    }
}