using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;
using MPR_Managerment.Services;
using OfficeOpenXml;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Parser.Biff_Records.ObjRecords;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.WebSockets;
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
        private ProjectService _projectServices = new ProjectService();

        private List<RIRDetail> _details = new List<RIRDetail>();

        private bool _isSearching = false;
        private int _selectedRIR_ID = 0;

        private List<WarehouseImport> lstItemAdd = new List<WarehouseImport>();
        private List<int> lstRootItem = new List<int>();
        private int _rirId = 0;
        private List<int> lstRemoveItem = new List<int>();

        private DataTable _dtProjectMaterial = new DataTable();
        private DataTable _dtProjectPaint = new DataTable();
        private DataTable _dtProjectWelding = new DataTable();
        private bool _isLoaded = false;

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
            //dgvRIR.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            dgvRIR.Dock = DockStyle.Fill;
            dgvRIR.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvRIR.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvRIR.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvRIR.EnableHeadersVisualStyles = false;
            dgvRIR.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);

            //txtSearch.KeyDown += (s, ev) => { if (ev.KeyCode == Keys.Enter) { btnSearch.PerformClick(); ev.SuppressKeyPress = true; } };

            CreateContextMenuStripForGrid();
        }

        private async void ucRIRForQC_Load(object sender, EventArgs e)
        {
            var dt = await _projectServices.GetProjects();
            cboProjectMaterial.DisplayMember = "ProjectCode";
            cboProjectMaterial.ValueMember = "Id";
            cboProjectMaterial.DataSource = dt.Copy();
            _isLoaded = true;
        }

        private void CreateContextMenuStripForGrid()
        {
            // 1. Khởi tạo ContextMenuStrip
            ContextMenuStrip menuStock = new ContextMenuStrip();

            // 2. Thêm các mục (Items) vào menu
            ToolStripMenuItem itemXemChiTiet = new ToolStripMenuItem("📄 Thêm dòng cho vật tư");
            ToolStripMenuItem itemXoaVatTu = new ToolStripMenuItem("❌ Xóa vật tư"); // Thêm mục Xóa vật tư mới

            // Thêm các mục vào menu chung (bao gồm cả nút xóa)
            menuStock.Items.AddRange(new ToolStripItem[] { itemXemChiTiet, itemXoaVatTu });

            // 3. Gắn menu vào DataGridView
            if (AppSession.CurrentUser.Role_ID == 1)
            {
                dgvRIR.ContextMenuStrip = menuStock;
            }

            // 4. Sự kiện khi click vào một mục trong menu
            itemXemChiTiet.Click += (s, e) =>
            {
                if (dgvRIR.CurrentRow != null && string.IsNullOrEmpty(dgvRIR.CurrentRow.Cells["IsAdded"].Value?.ToString() ?? ""))
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

                    row.Cells["IsAdded"].Value = "New";
                }
            };

            // 4b. HÀNH ĐỘNG MỚI: Xử lý sự kiện khi click vào mục "❌ Xóa vật tư"
            itemXoaVatTu.Click += (s, e) =>
            {
                try
                {
                    // Kiểm tra chắc chắn đang chọn một dòng hợp lệ và dòng đó không phải dòng trống cuối cùng phục vụ nhập liệu (NewRow)
                    if (dgvRIR.CurrentRow != null && !dgvRIR.CurrentRow.IsNewRow)
                    {
                        // Hiển thị hộp thoại xác nhận trước khi xóa để tránh người dùng click nhầm
                        DialogResult confirmResult = MessageBox.Show(
                            "Bạn có chắc chắn muốn xóa vật tư thuộc dòng đang chọn này không?",
                            "Xác nhận xóa",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question
                        );

                        if (confirmResult == DialogResult.Yes)
                        {
                            if (dgvRIR.CurrentRow.Cells["IsAdded"].Value == "New")
                            {
                                dgvRIR.Rows.Remove(dgvRIR.CurrentRow);
                            }
                            else
                            {
                                try
                                {
                                    // Thực hiện xóa dòng trong database
                                    int rir_detail_id = Convert.ToInt32(dgvRIR.CurrentRow.Cells["RIR_Detail_ID"].Value.ToString().Trim());
                                    _service.DeleteDetail(rir_detail_id);
                                    // Thực hiện xóa dòng hiện tại ra khỏi DataGridView
                                    dgvRIR.Rows.Remove(dgvRIR.CurrentRow);
                                }
                                catch (SqlException ex)
                                {
                                    MessageBox.Show($"Không thể xóa dòng: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Không thể xóa dòng: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "Qty_Required", HeaderText = "SL Yêu cầu", Width = 80, ReadOnly = false });
            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "Qty_Received", HeaderText = "SL Thực nhận", Width = 85, Visible = false });
            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "MTRno", HeaderText = "MTR No", Width = 100, ReadOnly = false, });
            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "Heatno", HeaderText = "Heat No", Width = 90, ReadOnly = false, });
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
                string kw = cboProjectMaterial.Text.Trim();
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

                        IsNewRow = string.IsNullOrEmpty(row.Cells["IsAdded"].Value?.ToString()) ? "false" : "true",
                    };

                    await _service.UpdateDetailForQC(d);

                    saved++;
                }

                //foreach (DataGridViewRow row in dgvRIR.Rows)
                //{
                //    int po_d_id = Convert.ToInt32(row.Cells["PO_Detail_ID"].Value?.ToString());
                //    var Qty_Required = (int)Math.Round(Convert.ToDecimal(row.Cells["Qty_Required"].Value ?? 0));
                //    var MTRno = row.Cells["MTRno"].Value?.ToString() ?? "";
                //    var Heatno = row.Cells["Heatno"].Value?.ToString() ?? "";
                //    var ID_Code = row.Cells["ID_Code"].Value?.ToString() ?? "";
                //    var Inspect_Result = row.Cells["Inspect_Result"].Value?.ToString() ?? "";
                //    var isNewRow = row.Cells["IsAdded"].Value?.ToString() ?? "";

                //    if (lstRootItem.Contains(po_d_id) && string.IsNullOrEmpty(isNewRow))
                //    {
                //        var lstAdd = lstItemAdd.Where(i => i.PO_Detail_ID == po_d_id).ToList();
                //        bool rs = await _warehouseServies.SaveQCCodeForItemOfWarehouseImportTable(_rirId, po_d_id, Qty_Required, 0, MTRno, Heatno, ID_Code, Inspect_Result, lstAdd);
                //    }
                //}

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

                    lstRootItem.Add(Convert.ToInt32(d.PO_Detail_ID));
                    _rirId = d.RIR_ID;
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

            if (dgvRIR.Columns[e.ColumnIndex].Name == "IsAdded")
            {
                string val = e.Value?.ToString() ?? "";
                e.CellStyle.ForeColor =
                    val == "New" ? Color.FromArgb(40, 167, 69) : Color.Black;
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
            cboRIRs.DataSource = null;
            _details.Clear();
            dgvRIR.Refresh();
        }

        private void dgvRIR_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        public async void ExportIdCodeListFromDatabase(DataTable dtDetails)
        {
            // 1. Kiểm tra file Template
            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "raw_material_id_code_template.xlsx");
            if (!File.Exists(templatePath))
            {
                MessageBox.Show("Không tìm thấy file template [2. Raw Material ID Code List.xlsx!]", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 2. Lấy dữ liệu từ SQL Server
            // Lấy thông tin Header ra biến để sử dụng
            string projectName = cboProjectMaterial.Text.Trim().ToUpper();

            if (dtDetails.Rows.Count == 0)
            {
                MessageBox.Show("Không tìm thấy thông tin dự án tương ứng!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }


            // 3. Mở hộp thoại lưu file Excel mới
            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"ID_Code_List_{projectName}_{DateTime.Now:ddMMyyyy_HHmm}.xlsx",
                Title = "Lưu file ID Code List"
            };

            if (saveDialog.ShowDialog() != DialogResult.OK) return;

            try
            {
                // Sao chép từ template sang vị trí đích mới
                File.Copy(templatePath, saveDialog.FileName, true);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(saveDialog.FileName)))
                {
                    var ws = package.Workbook.Worksheets[0]; // Lấy sheet đầu tiên

                    // --- PHẦN 2: ĐỔ DỮ LIỆU CHI TIẾT VÀO BẢNG ---
                    int startRow = 6; // Dòng bắt đầu điền item dữ liệu đầu tiên (Dưới hàng tiêu đề)
                    int detailCount = dtDetails.Rows.Count;

                    // Nếu số dòng dữ liệu nhiều hơn 1 dòng mẫu thiết kế sẵn, tiến hành chèn dòng hàng loạt
                    if (detailCount > 1)
                    {
                        // Chèn thêm (detailCount - 1) dòng bên dưới dòng mẫu, kế thừa định dạng (style) của startRow
                        ws.InsertRow(startRow + 1, detailCount - 1, startRow);
                    }

                    // Duyệt danh sách điền dữ liệu vào từng ô tương ứng theo cấu trúc cột của template
                    for (int i = 0; i < detailCount; i++)
                    {
                        DataRow row = dtDetails.Rows[i];
                        int currentRow = startRow + i;
                        ws.Cells[currentRow, 2].Value = "";
                        ws.Cells[currentRow, 3].Value = row["Qty_Per_Sheet"]?.ToString();
                        ws.Cells[currentRow, 4].Value = row["Size"]?.ToString();
                        ws.Cells[currentRow, 5].Value = "";
                        ws.Cells[currentRow, 6].Value = row["Material"]?.ToString();
                        ws.Cells[currentRow, 7].Value = row["Heatno"]?.ToString();
                        ws.Cells[currentRow, 8].Value = row["MTRno"]?.ToString();
                        ws.Cells[currentRow, 9].Value = row["ID_Code"]?.ToString();
                    }
                    ws.Cells[startRow, 1, detailCount + startRow - 1, 1].Merge = true;
                    ws.Cells[startRow, 1].Value = projectName;

                    // Lưu dữ liệu lại vào file
                    package.Save();
                }
                if (MessageBox.Show($"✅ Xuất báo cáo dữ liệu thành công!\nTổng số dòng vật tư: {dtDetails.Rows.Count}\nXuất Excel thành công! Bạn có muốn mở file?", "Thành công", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = saveDialog.FileName, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi trong quá trình xử lý hoặc ghi file Excel: " + ex.Message, "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void btnExport_Click(object sender, EventArgs e)
        {
            if (!Common.Common.IsComboBoxValid(cboProjectMaterial)) return;
            var dt = await _service.GetMaterialIDCodeListOfRIRForQC(Convert.ToInt32(cboRIRs.SelectedValue?.ToString()));
            ExportIdCodeListFromDatabase(dt);
        }

        private void cboProjectMaterial_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!_isLoaded) return;
            btnSearch.PerformClick();
        }

        private async void btnPrintReportMaterial_Click(object sender, EventArgs e)
        {
            var dt = await _service.GetMaterialIDCodeListOfProjectForQC(cboProjectMaterial.Text);
            ExportIdCodeListFromDatabase(dt);
        }
    }
}