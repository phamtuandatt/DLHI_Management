using Microsoft.Data.SqlClient;
using MPR_Managerment.Forms.ItemCodeGUI;
using MPR_Managerment.Helpers;
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
using System.Xml.Linq;

namespace MPR_Managerment.Forms.WarehouseGUI
{
    public partial class ucImportWarehouse : UserControl
    {
        private List<ProjectInfo> _dtProject = new List<ProjectInfo>();
        private POService _poService = new POService();
        private WarehouseService _warehouseServices = new WarehouseService();

        private List<POHead> _poList = new List<POHead>();
        private List<string> originalPOList = new List<string>();
        private string _currentBatchNo = "";
        private List<WarehouseImport> _importQueue = new List<WarehouseImport>();
        private Dictionary<string, string> _importList = new Dictionary<string, string>();
        private object oldValue = null;

        //private Form TopOwner { get { var f = (this.TopLevelControl as Form) ?? this; if (!f.IsDisposed) { f.BringToFront(); f.Activate(); } return f; } }

        public ucImportWarehouse()
        {
            InitializeComponent();
            BuildButton();
            BuildGrid();
            _dtProject = new ProjectService().GetAll();

        }

        private void ucImportWarehouse_Load(object sender, EventArgs e)
        {
            LoadForImport();
        }

        private void LoadForImport()
        {
            try
            {
                cboProjectForImport.Items.Clear();
                cboProjectForImport.Items.Add("Tất cả dự án");
                foreach (var p in _dtProject) cboProjectForImport.Items.Add(p.ProjectCode);
                cboProjectForImport.SelectedIndex = 0;
            }
            catch { }
        }

        private void BuildButton()
        {
            Common.Common.CreateButtonSearch(btnSearchItemPO, "🔍 Tìm kiếm");
            Common.Common.CreateButtonRefresh(btnRefresh);
            Common.Common.CreateButtonDelete(btnDeleteRow);
            Common.Common.CreateButtonSave(btnSaveImport);
        }

        private void BuildGrid()
        {
            dgvImportQueue.BackgroundColor = Color.White;
            dgvImportQueue.AutoGenerateColumns = true;
            dgvImportQueue.AllowUserToAddRows = false;
            dgvImportQueue.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvImportQueue.BorderStyle = BorderStyle.None;
            dgvImportQueue.RowHeadersVisible = false;
            dgvImportQueue.Font = new Font("Segoe UI", 9);
            dgvImportQueue.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvImportQueue.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            dgvImportQueue.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 140, 0);
            dgvImportQueue.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvImportQueue.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvImportQueue.EnableHeadersVisualStyles = false;
            dgvImportQueue.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(255, 248, 235);

            // Xanh nhạt cho selection
            dgvImportQueue.DefaultCellStyle.SelectionBackColor = Color.FromArgb(204, 232, 255);
            dgvImportQueue.DefaultCellStyle.SelectionForeColor = Color.Black;

            dgvImportQueue.CellBeginEdit += DgvImportQueue_CellBeginEdit; ;
            dgvImportQueue.CellEndEdit += DgvImportQueue_CellEndEdit; ;
            dgvImportQueue.EditingControlShowing += DgvImportQueue_EditingControlShowing; ;
            dgvImportQueue.CellDoubleClick += DgvImportQueue_CellDoubleClick; ;
            dgvImportQueue.KeyDown += DgvImportQueue_KeyDown;

            dgvImportQueue.Columns.Clear();
            dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "QIdx", HeaderText = "#", Width = 35, ReadOnly = true }); // 0
            dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "Item_Name", HeaderText = "Tên vật tư", Width = 220, ReadOnly = true }); // 1
            dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "Material", HeaderText = "Vật liệu", Width = 90, ReadOnly = true }); // 2
            dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "Size", HeaderText = "Kích thước", Width = 110, ReadOnly = true }); // 3
            dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "UNIT", HeaderText = "ĐVT", Width = 55, ReadOnly = true }); // 4
            dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "Qty_Import", HeaderText = "SL nhập", Width = 80 }); // 5
            dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "Weight_kg", HeaderText = "KG", Width = 75 }); // 6
            dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "ID_Code", HeaderText = "ID Code", Width = 100, ReadOnly = true }); // 7
            //dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "Recevied_Qty", HeaderText = "Số lượng đã nhận", Width = 160, ReadOnly = true }); /// New
            dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ma_Phieu", HeaderText = "Mã phiếu", Width = 160, ReadOnly = true }); // 8
            //dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "Material_Detail_Id", HeaderText = "Material Detail Id", Width = 160, ReadOnly = true, Visible = false }); 
            //dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "Material_Detail_Number", HeaderText = "Material Detail Number", Width = 160, ReadOnly = true, Visible = false });
            dgvImportQueue.Columns.Add(new DataGridViewTextBoxColumn { Name = "PO_Detail_ID", HeaderText = "PO_Detail_ID", Width = 160, ReadOnly = true, Visible = false }); // 9
        }

        private void RefreshQueueGrid()
        {
            dgvImportQueue.Rows.Clear();
            for (int i = 0; i < _importQueue.Count; i++)
            {
                var item = _importQueue[i];
                dgvImportQueue.Rows.Add(
                    i + 1,
                    item.Item_Name,
                    item.Material,
                    item.Size,
                    item.UNIT,
                    item.Qty_Import,
                    item.Weight_kg,
                    item.ID_Code,
                    item.Import_No,
                    item.PO_Detail_ID);
            }
        }

        private void DgvImportQueue_KeyDown(object? sender, KeyEventArgs e)
        {
            //throw new NotImplementedException();
        }

        private void DgvImportQueue_CellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _importQueue.Count) return;
            string colName = dgvImportQueue.Columns[e.ColumnIndex].Name;
            if (colName != "ID_Code") return;

            var item = _importQueue[e.RowIndex];
            frmCreateItemCode frmCreateItemCode = new frmCreateItemCode($"{item.Item_Name} - {item.Size} ");
            frmCreateItemCode.ShowDialog();
            if (string.IsNullOrEmpty(frmCreateItemCode.itemCode)) return;
            _importQueue[e.RowIndex].ID_Code = frmCreateItemCode.itemCode;
            dgvImportQueue.CurrentRow.Cells[colName].Value = frmCreateItemCode.itemCode;
        }

        private void DgvImportQueue_EditingControlShowing(object? sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress -= new KeyPressEventHandler(Column_KeyPress);
            if (dgvImportQueue.CurrentCell.ColumnIndex == dgvImportQueue.Columns["Qty_Import"].Index
                || dgvImportQueue.CurrentCell.ColumnIndex == dgvImportQueue.Columns["Weight_kg"].Index)
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress += new KeyPressEventHandler(Column_KeyPress);
                }
            }
        }

        private void Column_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void DgvImportQueue_CellEndEdit(object? sender, DataGridViewCellEventArgs e)
        {
            if (dgvImportQueue.Columns[e.ColumnIndex].Name == "Qty_Import")
            {
                var cell = dgvImportQueue.Rows[e.RowIndex].Cells[e.ColumnIndex];
                var poDetailIdCell = dgvImportQueue.Rows[e.RowIndex].Cells[9].Value;
                if (cell.Value != null)
                {
                    decimal newValue;
                    if (decimal.TryParse(cell.Value.ToString(), out newValue))
                    {
                        decimal originalLimit = Convert.ToDecimal(oldValue ?? 0);
                        if (newValue > originalLimit || newValue <= 0)    
                        {
                            MessageBox.Show($"Số lượng không được vượt quá số lượng của PO !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            cell.Value = oldValue;
                        }
                        //_importQueueActual[(int)poDetailIdCell] = newValue;
                        _importQueue[e.RowIndex].Qty_Import = newValue;
                    }
                    else
                    {
                        cell.Value = oldValue;
                    }
                }
            }

            if (dgvImportQueue.Columns[e.ColumnIndex].Name == "Weight_kg")
            {
                var cell = dgvImportQueue.Rows[e.RowIndex].Cells[e.ColumnIndex];
                if (cell.Value != null)
                {
                    decimal newValue;
                    if (decimal.TryParse(cell.Value.ToString(), out newValue))
                    {
                        decimal originalLimit = Convert.ToDecimal(oldValue ?? 0);
                        if (newValue > originalLimit + 10)
                        {
                            MessageBox.Show($"Khối lượng không được vượt quá khối lượng của PO !", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            cell.Value = oldValue;
                        }
                    }
                    else
                    {
                        cell.Value = oldValue;
                    }
                }
            }
        }

        private void DgvImportQueue_CellBeginEdit(object? sender, DataGridViewCellCancelEventArgs e)
        {
            oldValue = dgvImportQueue.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
        }

        private void btnSearchItemPO_Click(object sender, EventArgs e)
        {
            try
            {
                if (!Common.Common.IsComboBoxValid(cboProjectForImport, "Dự án")
                    || !Common.Common.IsComboBoxValid(cboPONoForImport, "PO"))
                    return;


                string poNo = cboPONoForImport.SelectedItem.ToString();
                _poList = _poService.GetAll();
                _poList = _poList.FindAll(p => !string.Equals(p.Status, "Cancelled", StringComparison.OrdinalIgnoreCase));
                var po = _poList.Find(p => p.PONo == poNo);
                if (po == null) return;
                var details = _poService.GetDetails(po.PO_ID);
                if (details.Count == 0) { MessageBox.Show("PO này chưa có chi tiết vật tư!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }

                using (var dlg = new Form())
                {
                    dlg.Text = $"Chọn vật tư nhập kho từ PO: {poNo}";
                    dlg.Size = new Size(1100, 510);
                    dlg.StartPosition = FormStartPosition.CenterParent;
                    dlg.BackColor = Color.White;
                    dlg.Controls.Add(new Label { Text = $"PO: {poNo}  —  {po.Project_Name}  —  Tick chọn vật tư, sửa SL nếu cần:", Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212), Location = new Point(10, 10), Size = new Size(900, 25) });
                    var dgv = new DataGridView { Location = new Point(10, 45), Size = new Size(1060, 350), AllowUserToAddRows = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, BackgroundColor = Color.White, BorderStyle = BorderStyle.FixedSingle, RowHeadersVisible = false, Font = new Font("Segoe UI", 9), AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill };
                    dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
                    dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                    dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    dgv.EnableHeadersVisualStyles = false;
                    dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
                    dlg.Controls.Add(dgv);

                    // Add DataGridView Columns
                    dgv.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Chon", HeaderText = "Chọn", Width = 50 });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "PO_Detail_ID", HeaderText = "ID", Visible = false });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "STT", HeaderText = "STT", Width = 40, ReadOnly = true });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ten_Hang", HeaderText = "Tên hàng", Width = 210, ReadOnly = true });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Vat_Lieu", HeaderText = "Vật liệu", Width = 80, ReadOnly = true });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "A_mm", HeaderText = "A(mm)", Width = 60, ReadOnly = true });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "B_mm", HeaderText = "B(mm)", Width = 60, ReadOnly = true });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "C_mm", HeaderText = "C(mm)", Width = 60, ReadOnly = true });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "DVT", HeaderText = "ĐVT", Width = 50, ReadOnly = true });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "SL_NK", HeaderText = "SL nhập", Width = 75, ReadOnly = true });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "KG", HeaderText = "KG", Width = 65, ReadOnly = true });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "MPS_No", HeaderText = "MPS No", Width = 90, ReadOnly = true });
                    dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Received_Qty", HeaderText = "SL đã nhập", Width = 90, ReadOnly = true }); /// NEW

                    foreach (var d in details)
                    {
                        dgv.Rows.Add(false, d.PO_Detail_ID, d.Item_No, d.Item_Name, d.Material, d.Asize, d.Bsize, d.Csize, d.UNIT, d.Qty_Per_Sheet, d.Weight_kg, d.MPSNo, d.Received_Qty);
                    }

                    var btnAll = new Button { Text = "☑ Chọn tất cả", Location = new Point(10, 405), Size = new Size(120, 32), BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
                    btnAll.FlatAppearance.BorderSize = 0;
                    btnAll.Click += (s2, e2) => { foreach (DataGridViewRow r in dgv.Rows) r.Cells["Chon"].Value = true; };
                    dlg.Controls.Add(btnAll);

                    var btnAdd = new Button { Text = "✔ Thêm vào phiếu", Location = new Point(140, 405), Size = new Size(160, 32), BackColor = Color.FromArgb(40, 167, 69), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), DialogResult = DialogResult.OK };
                    btnAdd.FlatAppearance.BorderSize = 0;
                    dlg.Controls.Add(btnAdd);

                    var btnCan = new Button { Text = "Hủy", Location = new Point(310, 405), Size = new Size(80, 32), BackColor = Color.FromArgb(108, 117, 125), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), DialogResult = DialogResult.Cancel };
                    btnCan.FlatAppearance.BorderSize = 0;
                    dlg.Controls.Add(btnCan);
                    dlg.AcceptButton = btnAdd;
                    dlg.CancelButton = btnCan;

                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        if (string.IsNullOrEmpty(_currentBatchNo)) _currentBatchNo = GenerateImportNo(poNo);
                        int addedCount = 0;
                        foreach (DataGridViewRow row in dgv.Rows)
                        {
                            bool ticked = row.Cells["Chon"].Value != null && Convert.ToBoolean(row.Cells["Chon"].Value);
                            if (!ticked) continue;
                            int pdId = Convert.ToInt32(row.Cells["PO_Detail_ID"].Value);
                            var detail = details.Find(d => d.PO_Detail_ID == pdId);
                            if (detail == null) continue;
                            decimal qty = decimal.TryParse(row.Cells["SL_NK"].Value?.ToString(), out decimal q) ? q : detail.Qty_Per_Sheet;

                            string projectCode = "";
                            if (cboProjectForImport != null && cboProjectForImport.SelectedIndex > 0)
                                projectCode = cboProjectForImport.SelectedItem?.ToString() ?? "25G0";
                            else
                            {
                                try
                                {
                                    var pjs = _dtProject;
                                    projectCode = pjs.Find(p => p.WorkorderNo == po.WorkorderNo)?.ProjectCode ?? po.MPR_No ?? "";
                                }
                                catch { projectCode = po.MPR_No ?? ""; }
                            }

                            _importList.Add($"{row.Cells["STT"].Value.ToString()}_{detail.Item_Name.ToString().Trim().ToLower()}",
                                qty.ToString());

                            _importQueue.Add(new WarehouseImport
                            {
                                Import_No = _currentBatchNo,
                                Import_Date = DateTime.Now,
                                PO_ID = po.PO_ID,
                                PO_Detail_ID = detail.PO_Detail_ID,
                                Item_Name = detail.Item_Name ?? "",
                                Material = detail.Material ?? "",
                                Size = $"{detail.Asize}x{detail.Bsize}x{detail.Csize}",
                                UNIT = detail.UNIT ?? "",
                                Qty_Import = qty,
                                Weight_kg = detail.Weight_kg,
                                Project_Code = projectCode,
                                WorkorderNo = po.WorkorderNo ?? "",

                                Received_Qty = Convert.ToDecimal(row.Cells["Received_Qty"].Value.ToString().Trim() ?? "0")
                            });

                            addedCount++;
                        }
                        RefreshQueueGrid();
                        if (addedCount > 0)
                            MessageBox.Show($"✅ Đã thêm {addedCount} vật tư vào phiếu: {_currentBatchNo}\nTổng: {_importQueue.Count} items — Nhấn 'Lưu phiếu nhập' để hoàn tất.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }


        private string GenerateImportNo(string poNo)
        {
            try
            {
                string baseNo = $"PNK-{poNo}";
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var cmd = new SqlCommand(
                        "SELECT COUNT(DISTINCT Import_No) FROM Warehouse_Import WHERE Import_No LIKE @base", conn);
                    cmd.Parameters.AddWithValue("@base", baseNo + "%");
                    int count = Convert.ToInt32(cmd.ExecuteScalar());
                    var uniqueQ = new HashSet<string>();
                    foreach (var q in _importQueue)
                        if (q.Import_No.StartsWith(baseNo)) uniqueQ.Add(q.Import_No);
                    int total = count + uniqueQ.Count;
                    return total == 0 ? baseNo : $"{baseNo}_{total + 1}";
                }
            }
            catch { return $"PNK-{poNo}-{DateTime.Now:ddMMHHmm}"; }
        }


        private void cboProjectForImport_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (dgvImportQueue.Rows.Count > 0)
                {
                    if (MessageBox.Show($"Bạn có {_importQueue.Count} items chưa lưu. Tạo phiếu mới sẽ xóa danh sách. Tiếp tục?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
                    _importQueue.Clear();
                    _importList.Clear();
                    dgvImportQueue.Refresh();
                    dgvImportQueue.Rows.Clear();
                }
                string project = (cboProjectForImport != null && cboProjectForImport.SelectedIndex > 0) ? cboProjectForImport.SelectedItem.ToString() : "";
                LoadPOFilterByProject(project);
            }
            catch (Exception ex) { MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }
        private void LoadPOFilterByProject(string projectCode)
        {
            try
            {
                var allPO = _poService.GetAllPOForImport();
                // Loại bỏ PO bị Cancelled
                allPO = allPO.FindAll(p => !string.Equals(p.Status, "Cancelled", StringComparison.OrdinalIgnoreCase));
                if (string.IsNullOrEmpty(projectCode))
                {
                    cboPONoForImport.Items.Clear();
                    cboPONoForImport.Items.Add("-- Chọn PO --");
                    foreach (var po in allPO)
                    {
                        cboPONoForImport.Items.Add(po.PONo);
                    }
                    cboPONoForImport.SelectedIndex = 0;
                    return;
                }
                var projects = _dtProject;
                var proj = projects.Find(p => p.ProjectCode == projectCode);
                List<POHead> filtered;

                if (proj != null)
                    filtered = allPO.FindAll(p =>
                        (!string.IsNullOrEmpty(proj.WorkorderNo) && (p.WorkorderNo ?? "").Equals(proj.WorkorderNo, StringComparison.OrdinalIgnoreCase)) ||
                        (!string.IsNullOrEmpty(proj.MPRCode) && (p.MPR_No ?? "").Contains(proj.MPRCode, StringComparison.OrdinalIgnoreCase)) ||
                        (!string.IsNullOrEmpty(proj.ProjectCode) && (p.WorkorderNo ?? "").Contains(proj.ProjectCode, StringComparison.OrdinalIgnoreCase)));
                else
                    filtered = allPO.FindAll(p =>
                        (p.WorkorderNo ?? "").Contains(projectCode, StringComparison.OrdinalIgnoreCase) ||
                        (p.MPR_No ?? "").Contains(projectCode, StringComparison.OrdinalIgnoreCase));

                cboPONoForImport.Items.Clear();
                cboPONoForImport.Items.Add("-- Chọn PO --");

                if (filtered.Count == 0)
                {
                    cboPONoForImport.Items.Add("(Không có PO)");
                    cboPONoForImport.SelectedIndex = 0;
                    return;
                }
                originalPOList.Clear();
                foreach (var po in filtered)
                {
                    cboPONoForImport.Items.Add(po.PONo);
                    originalPOList.Add(po.PONo);
                }
                cboPONoForImport.SelectedIndex = 0;
            }
            catch (Exception ex) { MessageBox.Show("Lỗi lọc PO: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void cboPONoForImport_TextUpdate(object sender, EventArgs e)
        {
            Common.Common.ComboBoxTextUpdateForListItem(sender, e, originalPOList);
        }

        private void btnDeleteRow_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("WAREHOUSE", "Lưu hóa đơn", "Xóa dòng nhập kho")) return;
            if (dgvImportQueue.SelectedRows.Count == 0) return;
            int idx = dgvImportQueue.SelectedRows[0].Index;
            if (idx >= 0 && idx < _importQueue.Count)
            {
                var key = $"{dgvImportQueue.SelectedRows[0].Cells[0].Value.ToString().Trim()}_{dgvImportQueue.SelectedRows[0].Cells[1].Value.ToString().Trim().ToLower()}";
                _importList.Remove(key);
                _importQueue.RemoveAt(idx);

                if (_importQueue.Count == 0) _currentBatchNo = "";
                RefreshQueueGrid();
            }
        }

        private void btnSaveImport_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("WAREHOUSE", "Lưu hóa đơn", "Lưu hóa đơn nhập kho")) return;
            if (!Common.Common.IsDataGridViewValid(dgvImportQueue, "Danh sách vật tư")) return;
            foreach (DataGridViewRow item in dgvImportQueue.Rows)
            {
                if (string.IsNullOrEmpty(item.Cells["ID_Code"].Value.ToString()))
                {
                    MessageBox.Show($"Hãy tạo code cho item: {item.Cells["Item_Name"].Value.ToString()}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            if (_importQueue.Count == 0) { MessageBox.Show("Danh sách phiếu đang trống!\nHãy thêm vật tư trước.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            try
            {
                int saved = 0;
                foreach (var imp in _importQueue)
                {
                    imp.Import_Date = DateTime.Now;
                    _warehouseServices.InsertImport(imp, AppSession.CurrentUser.Full_Name);
                    saved++;
                }

                MessageBox.Show($"✅ Lưu phiếu nhập kho thành công!\nMã phiếu: {_currentBatchNo}\nSố vật tư: {saved} items", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                _currentBatchNo = ""; 
                _importQueue.Clear();
                _importList.Clear();
                RefreshQueueGrid();
                LoadPOFilterByProject(cboProjectForImport.Text.Trim());
            }
            catch (Exception ex) { MessageBox.Show("Lỗi nhập kho: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }
    }
}
