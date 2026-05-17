using MPR_Managerment.Helpers;
using MPR_Managerment.Models;
using MPR_Managerment.Services;
using Syncfusion.XlsIO.Parser.Biff_Records.ObjRecords;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;

namespace MPR_Managerment.Forms.RIRGUI
{
    public partial class ucRIRForQC : UserControl
    {
        private DataTable _dtProject = new DataTable();
        private DataTable _dtRIRs = new DataTable();

        private WarehouseService _warehouseServies = new WarehouseService();
        private RIRService _service = new RIRService();

        private List<RIRDetail> _details = new List<RIRDetail>();

        private bool _isSearching = false;
        private int _selectedRIR_ID = 0;

        private List<WarehouseImport> lstItemAdd = new List<WarehouseImport>();
        private List<int> lstRootItem = new List<int>();

        public ucRIRForQC()
        {
            InitializeComponent();
            BuildDetailColumns();
            ApplyPermissions();

            dgvRIR.BackgroundColor = Color.White;
            dgvRIR.BorderStyle = BorderStyle.None;
            dgvRIR.RowHeadersVisible = false;
            dgvRIR.Font = new Font("Segoe UI", 9);
            dgvRIR.AllowUserToAddRows = false;
            dgvRIR.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvRIR.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            dgvRIR.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvRIR.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvRIR.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvRIR.EnableHeadersVisualStyles = false;
            dgvRIR.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);

            txtSearch.KeyDown += (s, ev) => { if (ev.KeyCode == Keys.Enter) { btnSearch.PerformClick(); ev.SuppressKeyPress = true; } };

            CreateContextMenuStripForGrid();

        }

        private void CreateContextMenuStripForGrid()
        {
            // 1. Khởi tạo ContextMenuStrip
            ContextMenuStrip menuStock = new ContextMenuStrip();

            // 2. Thêm các mục (Items) vào menu
            ToolStripMenuItem itemXemChiTiet = new ToolStripMenuItem("📄 Thêm dòng cho vật tư");
            //ToolStripMenuItem itemSaoChep = new ToolStripMenuItem("📋 Sao chép mã");
            //ToolStripMenuItem itemXuatKho = new ToolStripMenuItem("📤 Xuất kho");

            menuStock.Items.AddRange(new ToolStripItem[] { itemXemChiTiet/*, itemSaoChep, new ToolStripSeparator(), itemXuatKho*/ });

            // 3. Gắn menu vào DataGridView
            if (AppSession.CurrentUser.Role_ID == 1)
            {
                dgvRIR.ContextMenuStrip = menuStock;
            }

            // 4. Sự kiện khi click vào một mục trong menu
            itemXemChiTiet.Click += (s, e) =>
            {
                if (dgvRIR.CurrentRow != null)
                {
                    // Lấy dữ liệu từ dòng đang chọn
                    var currentR = dgvRIR.CurrentRow;
                    var po_detail_id = currentR.Cells["PO_Detail_ID"].Value?.ToString();

                    // Nếu số lượng = 1 thì không cho tách
                    if (Convert.ToDecimal(currentR.Cells["Qty_Required"].Value.ToString()) == 1)
                    {
                        return;
                    }

                    // Tạo form cho nhập heat / MTR / SỐ lượng / QC_Code
                    frmAddHeatForItem frmAddHeatForItem = new frmAddHeatForItem(currentR.Cells["MTRno"].Value.ToString() ?? "");
                    frmAddHeatForItem.ShowDialog();
                    if (!frmAddHeatForItem.IsClose) return;
                    var qty = frmAddHeatForItem.Quantity;
                    var mtr = frmAddHeatForItem.MTRNo;
                    var heat = frmAddHeatForItem.HeatNo;
                    var qc_code = frmAddHeatForItem.ID_Code;

                    // Lưu thông tin dòng đã được tách
                    var wI = new WarehouseImport
                    {
                        PO_Detail_ID = Convert.ToInt32(po_detail_id),
                        Qty_Import = qty,
                        MTRno = mtr,
                        Heatno = heat,
                        QC_Code = qc_code, 
                    };
                    lstItemAdd.Add(wI);

                    // Cập nhật số lượng sau khi tách cho dòng cũ
                    currentR.Cells["Qty_Required"].Value = Convert.ToDecimal(currentR.Cells["Qty_Required"].Value) - qty;

                    // Ghi nhận dữ liệu mới -> Tạo dòng mới
                    int idx = dgvRIR.Rows.Add();
                    var row = dgvRIR.Rows[idx];
                    

                    row.Cells["RIR_Detail_ID"].Value = currentR.Cells["RIR_Detail_ID"].Value;
                    row.Cells["Item_No"].Value = currentR.Cells["Item_No"].Value;
                    row.Cells["Item_Name"].Value = currentR.Cells["Item_Name"].Value;
                    row.Cells["Material"].Value = currentR.Cells["Material"].Value;
                    row.Cells["Size"].Value = currentR.Cells["Size"].Value;
                    row.Cells["UNIT"].Value = currentR.Cells["UNIT"].Value;
                    row.Cells["Qty_Required"].Value = qty;
                    row.Cells["Qty_Received"].Value = 0;
                    row.Cells["MTRno"].Value = mtr;
                    row.Cells["Heatno"].Value = heat;
                    row.Cells["ID_Code"].Value = qc_code;
                    row.Cells["Inspect_Result"].Value = "";
                    row.Cells["Remarks"].Value = currentR.Cells["Remarks"].Value ?? "";

                    row.Cells["PO_Detail_ID"].Value = currentR.Cells["PO_Detail_ID"].Value;

                    row.Cells["IsAdded"].Value = "true";
                }
            };

            // 5. QUAN TRỌNG: Xử lý để chuột phải vào dòng nào thì chọn dòng đó (thay vì chỉ hiện menu)
            dgvRIR.MouseDown += (s, e) =>
            {
                if (e.Button == MouseButtons.Right)
                {
                    var hit = dgvRIR.HitTest(e.X, e.Y);
                    if (hit.RowIndex >= 0)
                    {
                        // Xóa các lựa chọn cũ và chọn dòng vừa click chuột phải
                        dgvRIR.ClearSelection();
                        dgvRIR.Rows[hit.RowIndex].Selected = true;
                        dgvRIR.CurrentCell = dgvRIR.Rows[hit.RowIndex].Cells[hit.ColumnIndex];
                    }
                }
            };
        }

        private void BuildDetailColumns()
        {
            dgvRIR.Columns.Clear();
            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "RIR_Detail_ID", HeaderText = "ID", Visible = false, ReadOnly = true });
            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "Item_No", HeaderText = "STT", Width = 45, ReadOnly = true });
            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "Item_Name", HeaderText = "Tên vật tư", Width = 200, ReadOnly = true, });
            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "Material", HeaderText = "Vật liệu", Width = 90, ReadOnly = true, });
            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "Size", HeaderText = "Kích thước", Width = 110, ReadOnly = true, });
            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "UNIT", HeaderText = "ĐVT", Width = 55, ReadOnly = true, });
            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "Qty_Required", HeaderText = "SL Yêu cầu", Width = 80, ReadOnly = true, });
            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "Qty_Received", HeaderText = "SL Thực nhận", Width = 85 });
            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "MTRno", HeaderText = "MTR No", Width = 100, ReadOnly = true, });
            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "Heatno", HeaderText = "Heat No", Width = 90, ReadOnly = true, });
            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "ID_Code", HeaderText = "ID Code", Width = 100 });

            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "PO_Detail_ID", HeaderText = "PO Detail No", Width = 100, ReadOnly = true, Visible = false }); // Add column PO_Detail_ID

            var cboResult = new DataGridViewComboBoxColumn
            {
                Name = "Inspect_Result",
                HeaderText = "Kết quả KT",
                Width = 100,
                FlatStyle = FlatStyle.Flat
            };
            cboResult.Items.AddRange(new[] { "", "Pass", "Fail", "Hold" });
            dgvRIR.Columns.Add(cboResult);

            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "Remarks", HeaderText = "Ghi chú", FillWeight = 100 });

            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "IsAdded", HeaderText = "Dòng mới", FillWeight = 100, ReadOnly = true });
        }


        private void dgvRIR_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress -= new KeyPressEventHandler(Column_KeyPress);
            if (dgvRIR.CurrentCell.ColumnIndex == dgvRIR.Columns["Qty_Required"].Index
                || dgvRIR.CurrentCell.ColumnIndex == dgvRIR.Columns["Qty_Received"].Index)
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


        private async void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                string kw = txtSearch.Text.Trim();
                _dtRIRs = await _warehouseServies.GetRIROfProject(kw);

                cboRIRs.DisplayMember = "RIR_No";
                cboRIRs.ValueMember = "RIR_ID";
                cboRIRs.DataSource = _dtRIRs;

                lblCountRIR.Text = $"Tìm thấy: {_dtRIRs.Rows.Count} phiếu";
                _isSearching = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void btnSave_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("Material Inspector Request", "Lưu chi tiết", "Lưu chi tiết")) return;
            if (!Common.Common.IsDataGridViewValid(dgvRIR)) return;

            try
            {
                int saved = 0;
                foreach (DataGridViewRow row in dgvRIR.Rows)
                {
                    string itemName = row.Cells["Item_Name"].Value?.ToString() ?? "";
                    if (string.IsNullOrWhiteSpace(itemName)) continue;

                    var d = new RIRDetail
                    {
                        RIR_Detail_ID = Convert.ToInt32(row.Cells["RIR_Detail_ID"].Value ?? 0),
                        RIR_ID = _selectedRIR_ID,
                        Item_No = Convert.ToInt32(row.Cells["Item_No"].Value ?? 0),
                        Item_Name = itemName,
                        Material = row.Cells["Material"].Value?.ToString() ?? "",
                        Size = row.Cells["Size"].Value?.ToString() ?? "",
                        UNIT = row.Cells["UNIT"].Value?.ToString() ?? "",
                        Qty_Required = (int)Math.Round(Convert.ToDecimal(row.Cells["Qty_Required"].Value ?? 0)),
                        Qty_Received = (int)Math.Round(Convert.ToDecimal(row.Cells["Qty_Received"].Value ?? 0)),
                        Qty_Per_Sheet = (int)Math.Round(Convert.ToDecimal(row.Cells["Qty_Required"].Value ?? 0)) > 0 ? (int)Math.Round(Convert.ToDecimal(row.Cells["Qty_Required"].Value ?? 0)) : (int)Math.Round(Convert.ToDecimal(row.Cells["Qty_Received"].Value ?? 0)),
                        MTRno = row.Cells["MTRno"].Value?.ToString() ?? "",
                        Heatno = row.Cells["Heatno"].Value?.ToString() ?? "",
                        ID_Code = row.Cells["ID_Code"].Value?.ToString() ?? "",
                        Inspect_Result = row.Cells["Inspect_Result"].Value?.ToString() ?? "",
                        Remarks = row.Cells["Remarks"].Value?.ToString() ?? "",
                        PO_Detail_ID = Convert.ToInt32(row.Cells["PO_Detail_ID"].Value?.ToString() ?? ""),

                        IsNewRow = row.Cells["IsAdded"].Value?.ToString() ?? ""
                    };

                    await _service.UpdateDetailForQC(d);

                    saved++;
                }

                foreach (DataGridViewRow row in dgvRIR.Rows)
                {
                    int po_d_id = Convert.ToInt32(row.Cells["PO_Detail_ID"].Value?.ToString());
                    var Qty_Required = (int)Math.Round(Convert.ToDecimal(row.Cells["Qty_Required"].Value ?? 0));
                    var MTRno = row.Cells["MTRno"].Value?.ToString() ?? "";
                    var Heatno = row.Cells["Heatno"].Value?.ToString() ?? "";
                    var ID_Code = row.Cells["ID_Code"].Value?.ToString() ?? "";
                    var Inspect_Result = row.Cells["Inspect_Result"].Value?.ToString() ?? "";
                    var isNewRow = row.Cells["IsAdded"].Value?.ToString() ?? "";

                    if (lstRootItem.Contains(po_d_id) && string.IsNullOrEmpty(isNewRow))
                    {
                        var lstAdd = lstItemAdd.Where(i => i.PO_Detail_ID == po_d_id).ToList();
                        bool rs = await _warehouseServies.SaveQCCodeForItemOfWarehouseImportTable(po_d_id, Qty_Required, 0, MTRno, Heatno, ID_Code, Inspect_Result, lstAdd);
                    }
                }

                // Kiểm tra nội dung khi truyền vào Procedure SQL
                // Tính số Weight sau khi tách

                MessageBox.Show($"Đã lưu {saved} dòng thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDetails(_selectedRIR_ID);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi lưu chi tiết: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cboRIRs_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_isSearching && Common.Common.IsComboBoxValid(cboRIRs))
            {
                int rirId = (int)cboRIRs.SelectedValue;
                LoadDetails(rirId);
                _selectedRIR_ID = rirId;
                lblStatus.Text = $"Phiếu gồm {dgvRIR.Rows.Count} dòng";
            }
        }

        private void LoadDetails(int rirId)
        {
            try
            {
                _details = _service.GetDetails(rirId);
                dgvRIR.Rows.Clear();

                foreach (var d in _details)
                {
                    int idx = dgvRIR.Rows.Add();
                    var row = dgvRIR.Rows[idx];

                    row.Cells["RIR_Detail_ID"].Value = d.RIR_Detail_ID;
                    row.Cells["Item_No"].Value = d.Item_No;
                    row.Cells["Item_Name"].Value = d.Item_Name;
                    row.Cells["Material"].Value = d.Material;
                    row.Cells["Size"].Value = d.Size;
                    row.Cells["UNIT"].Value = d.UNIT;
                    row.Cells["Qty_Required"].Value = d.Qty_Required;
                    row.Cells["Qty_Received"].Value = d.Qty_Received;
                    row.Cells["MTRno"].Value = d.MTRno;
                    row.Cells["Heatno"].Value = d.Heatno;
                    row.Cells["ID_Code"].Value = d.ID_Code;
                    row.Cells["Inspect_Result"].Value = d.Inspect_Result;
                    row.Cells["Remarks"].Value = d.Remarks ?? "";

                    row.Cells["PO_Detail_ID"].Value = d.PO_Detail_ID;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải chi tiết: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dgvRIR_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (dgvRIR.Columns[e.ColumnIndex].Name == "Inspect_Result")
            {
                string val = e.Value?.ToString() ?? "";
                e.CellStyle.ForeColor =
                    val == "Pass" ? Color.FromArgb(40, 167, 69) :
                    val == "Fail" ? Color.FromArgb(220, 53, 69) :
                    val == "Hold" ? Color.FromArgb(255, 140, 0) :
                                    Color.Black;
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }

            if ((dgvRIR.Columns[e.ColumnIndex].Name == "Qty_Required" && e.Value != null)
                && (dgvRIR.Columns[e.ColumnIndex].Name == "Qty_Received" && e.Value != null))
            {
                if (decimal.TryParse(e.Value.ToString(), out decimal qty))
                {
                    e.Value = qty.ToString("N0");
                    e.FormattingApplied = true;
                }
            }
        }

        private void dgvRIR_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            var qtyRequireCell = Convert.ToInt32(dgvRIR.CurrentRow.Cells["Qty_Required"].Value);
            var qtyRecivedCell = Convert.ToInt32(dgvRIR.CurrentRow.Cells["Qty_Received"].Value);

            if (qtyRecivedCell > qtyRequireCell)
            {
                MessageBox.Show("SL Thực nhận không được lớn hơn SL Yêu cầu!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                dgvRIR.CurrentCell.Value = qtyRequireCell;
            }
        }

        // =====================================================================
        //  APPLY PERMISSIONS
        // =====================================================================
        private void ApplyPermissions()
        {
            // Nút Lưu chi tiết — quyền "Lưu chi tiết" trong module MIR
            foreach (var c in this.Controls.Find("btnSave", true))
                PermissionHelper.Apply(c, "Material Inspector Request", "Lưu chi tiết");
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtSearch.Clear();
            txtSearch.Focus();
            cboRIRs.DataSource = null;
            _details.Clear();
            dgvRIR.Refresh();
        }
    }
}