using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;
using MPR_Managerment.Services;
using OfficeOpenXml;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Parser.Biff_Records;
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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;

namespace MPR_Managerment.Forms.RIRGUI
{
    public partial class ucRIRForQC : UserControl
    {
        private DataTable _dtProject = new DataTable();
        private DataTable _dtRIRs = new DataTable();

        private WarehouseService _warehouseServies = new WarehouseService();
        private RIRService _rirServices = new RIRService();
        private ProjectService _projectServices = new ProjectService();

        private List<RIRDetail> _details = new List<RIRDetail>();
        private DataTable _dtRIRDetail = new DataTable();

        private bool _isSearching = false;
        private int _selectedRIR_ID = 0;

        private List<WarehouseImport> lstItemAdd = new List<WarehouseImport>();
        private List<int> lstRootItem = new List<int>();
        private int _rirId = 0;
        private List<int> lstRemoveItem = new List<int>();

        private DataTable _dtProjectMaterial = new DataTable();
        private DataTable _dtProjectPaint = new DataTable();
        private DataTable _dtProjectWelding = new DataTable();
        private DataTable _dtRIRNoPaint = new DataTable();
        private DataTable _dtRIRNOWelding = new DataTable();

        private bool _isLoaded = false;
        private bool _dtPaintIsLoaded = false;
        private bool _dtWeldingIsLoaded = false;
        private bool _dtRIRNoPaintIsLoaded = false;
        private bool _dtRIRNoWeldingIsLoaded = false;

        // Track modified rows
        private HashSet<DataGridViewRow> _modifiedRows = new HashSet<DataGridViewRow>();
        private bool _isLoadingData = false;

        public ucRIRForQC()
        {
            InitializeComponent();
            //frmAIChat.Attach(this); // Bỏ dòng này vì đây là UserControl, không phải Form chính

            BuildDetailColumns();
            ApplyPermissions();

            //if (AppSession.CurrentUser.Role_ID != 1)
            //{
            //    btnSave.Visible = false;
            //}

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

            dgvRIR.CellValueChanged += dgvRIR_CellValueChanged;

            CreateContextMenuStripForGrid();
        }

        private async void ucRIRForQC_Load(object sender, EventArgs e)
        {
            var dt = await _projectServices.GetProjects();
            cboProjectMaterial.DisplayMember = "ProjectCode";
            cboProjectMaterial.ValueMember = "Id";
            cboProjectMaterial.DataSource = dt.Copy();

            cboProjectForPain.DisplayMember = "ProjectCode";
            cboProjectForPain.ValueMember = "Id";
            cboProjectForPain.DataSource = dt.Copy();

            cboProjectForWelding.DisplayMember = "ProjectCode";
            cboProjectForWelding.ValueMember = "Id";
            cboProjectForWelding.DataSource = dt.Copy();

            _isLoaded = true;

            LoadDataPaint();
            LoadDataWelding();
            await LoadRIRNoPaint();
            await LoadRIRNoWelding();
        }

        private void btnShowAllPaint_Click(object sender, EventArgs e)
        {
            _dtPaintIsLoaded = false;
            LoadDataPaint();
        }

        private void btnShowAllWelding_Click(object sender, EventArgs e)
        {
            _dtWeldingIsLoaded = false;
            LoadDataWelding();
        }

        private async Task LoadRIRNoPaint()
        {
            try
            {
                DataTable dt = await _rirServices.GetRIRNoPaint();

                cboRIRNoPaint.DisplayMember = "RIR_No";
                cboRIRNoPaint.ValueMember = "RIR_ID";
                cboRIRNoPaint.DataSource = dt;
                //cboRIRNoPaint.SelectedIndex = -1;

                _dtRIRNoPaint = await _rirServices.GetRIRHeaderForPaintOrWelding(true, false);
                if (_dtRIRNoPaint.Rows.Count > 0) _dtRIRNoPaintIsLoaded = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi load danh sách RIR No Paint: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task LoadRIRNoWelding()
        {
            try
            {
                DataTable dt = await _rirServices.GetRIRNoWelding();

                cboRIRNoWelding.DisplayMember = "RIR_No";
                cboRIRNoWelding.ValueMember = "RIR_ID";
                cboRIRNoWelding.DataSource = dt;
                //cboRIRNoWelding.SelectedIndex = -1;

                _dtRIRNOWelding = await _rirServices.GetRIRHeaderForPaintOrWelding(false, true);
                if (_dtRIRNOWelding.Rows.Count > 0) _dtRIRNoWeldingIsLoaded = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi load danh sách RIR No Welding: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void LoadDataPaint(string keyword = "")
        {
            if (string.IsNullOrEmpty(keyword) && !_dtPaintIsLoaded)
            {
                _dtProjectPaint = await _rirServices.GetPaintDetails();
                dgvPaint.DataSource = _dtProjectPaint.Copy();
                _dtPaintIsLoaded = true;
            }
            //else
            //{
            //    _dtProjectPaint = await _rirServices.GetPaintDetails(keyword);
            //    dgvPaint.DataSource = _dtProjectPaint;
            //}

            dgvPaint.Columns["RIR_Detail_ID"].Visible = false;
            dgvPaint.Columns["RIR_ID"].Visible = false;
            dgvPaint.Columns["PO_Detail_ID"].Visible = false;
            dgvPaint.Columns["Item_No"].Visible = false;
            dgvPaint.Columns["RIR_No"].Visible = false;
            dgvPaint.Columns["Created_Date"].Visible = false;
            dgvPaint.RowPostPaint += (s, e) =>
            {
                Common.Common.RenderNumbering(s, e);
            };
        }

        private async void LoadDataWelding(string keyword = "")
        {
            if (string.IsNullOrEmpty(keyword) && !_dtWeldingIsLoaded)
            {
                _dtProjectWelding = await _rirServices.GetWeldingDetails();
                dgvWelding.DataSource = _dtProjectWelding;
                _dtWeldingIsLoaded = true;
            }
            //else
            //{
            //    _dtProjectWelding = await _rirServices.GetWeldingDetails(keyword);
            //    dgvWelding.DataSource = _dtProjectWelding;
            //}

            dgvWelding.Columns["RIR_Detail_ID"].Visible = false;
            dgvWelding.Columns["RIR_ID"].Visible = false;
            dgvWelding.Columns["PO_Detail_ID"].Visible = false;
            dgvWelding.Columns["Item_No"].Visible = false;
            dgvWelding.Columns["RIR_No"].Visible = false;
            dgvWelding.Columns["Created_Date"].Visible = false;
            dgvWelding.RowPostPaint += (s, e) =>
            {
                Common.Common.RenderNumbering(s, e);
            };
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

            //// 3. Gắn menu vào DataGridView
            //if (AppSession.CurrentUser.Role_ID == 1)
            //{
            //    dgvRIR.ContextMenuStrip = menuStock;
            //}
            dgvRIR.ContextMenuStrip = menuStock;

            // 4. Sự kiện khi click vào một mục trong menu
            itemXemChiTiet.Click += (s, e) =>
            {
                if (dgvRIR.CurrentRow != null && string.IsNullOrEmpty(dgvRIR.CurrentRow.Cells["IsAdded"].Value?.ToString() ?? ""))
                {
                    // Lấy dữ liệu từ dòng đang chọn
                    var currentR = dgvRIR.CurrentRow;
                    var po_detail_id = currentR.Cells["PO_Detail_ID"].Value?.ToString();
                    var rir_no = currentR.Cells["RIR_No"].Value?.ToString();

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

                    // Ghi nhận dữ liệu mới -> Chèn dòng mới ngay bên dưới dòng hiện tại
                    int insertIndex = currentR.Index + 1;
                    dgvRIR.Rows.Insert(insertIndex);
                    var row = dgvRIR.Rows[insertIndex];

                    // Highlight new row with light green background
                    row.DefaultCellStyle.BackColor = Color.FromArgb(198, 239, 206);
                    row.DefaultCellStyle.SelectionBackColor = Color.FromArgb(150, 210, 160);


                    //row.Cells["RIR_Detail_ID"].Value = currentR.Cells["RIR_Detail_ID"].Value;
                    row.Cells["RIR_ID"].Value = currentR.Cells["RIR_ID"].Value;
                    row.Cells["RIR_No"].Value = rir_no;
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

                    // Add new row to modified rows tracking
                    _modifiedRows.Add(row);
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
                            if (dgvRIR.CurrentRow.Cells["IsAdded"].Value?.ToString() == "New")
                            {
                                var currentRow = dgvRIR.CurrentRow;
                                var rirDetailId = currentRow.Cells["RIR_Detail_ID"].Value?.ToString();
                                var qtyToDelete = Convert.ToDecimal(currentRow.Cells["Qty_Required"].Value ?? 0);

                                // Tìm dòng cha (dòng gốc có cùng RIR_Detail_ID và không phải là dòng mới thêm)
                                DataGridViewRow parentRow = null;
                                foreach (DataGridViewRow r in dgvRIR.Rows)
                                {
                                    if (r.Index != currentRow.Index &&
                                        r.Cells["RIR_Detail_ID"].Value?.ToString() == rirDetailId &&
                                        string.IsNullOrEmpty(r.Cells["IsAdded"].Value?.ToString()))
                                    {
                                        parentRow = r;
                                        break;
                                    }
                                }

                                if (parentRow != null)
                                {
                                    // Tạm thời tắt cờ bắt sự kiện CellValueChanged để không bị highlight lại
                                    _isLoadingData = true;

                                    // Cộng trả lại số lượng cho dòng cha
                                    parentRow.Cells["Qty_Required"].Value = Convert.ToDecimal(parentRow.Cells["Qty_Required"].Value ?? 0) + qtyToDelete;

                                    // Xóa dòng cha khỏi danh sách modified và reset lại màu mặc định
                                    _modifiedRows.Remove(parentRow);
                                    parentRow.DefaultCellStyle.BackColor = Color.Empty;
                                    parentRow.DefaultCellStyle.SelectionBackColor = Color.Empty;

                                    _isLoadingData = false;
                                }

                                dgvRIR.Rows.Remove(currentRow);
                            }
                            else
                            {
                                //try
                                //{
                                //    // Thực hiện xóa dòng trong database
                                //    int rir_detail_id = Convert.ToInt32(dgvRIR.CurrentRow.Cells["RIR_Detail_ID"].Value.ToString().Trim());
                                //    _service.DeleteDetail(rir_detail_id);
                                //    // Thực hiện xóa dòng hiện tại ra khỏi DataGridView
                                //    dgvRIR.Rows.Remove(dgvRIR.CurrentRow);
                                //}
                                //catch (SqlException ex)
                                //{
                                //    MessageBox.Show($"Không thể xóa dòng: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                //}
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
            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "RIR_ID", HeaderText = "RIR ID", Visible = false, ReadOnly = true });
            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "RIR_No", HeaderText = "RIR No", Visible = true, ReadOnly = true });
            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "Item_No", HeaderText = "STT", Width = 45, ReadOnly = true });
            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "Item_Name", HeaderText = "Tên vật tư", Width = 200, ReadOnly = true, });
            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "Material", HeaderText = "Vật liệu", Width = 90, ReadOnly = true, });
            dgvRIR.Columns.Add(new DataGridViewTextBoxColumn { Name = "Size", HeaderText = "Kích thước", Width = 110, ReadOnly = false, });
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
                System.Windows.Forms.TextBox tb = e.Control as System.Windows.Forms.TextBox;
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
                await LoadRIRDetail();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task LoadRIRDetail()
        {
            try
            {
                // Đổi con trỏ chuột sang trạng thái chờ (WaitCursor) nhằm tăng trải nghiệm người dùng
                this.Cursor = Cursors.WaitCursor;

                // Gọi hàm và nhận về một đối tượng DataTable
                _dtRIRDetail = await _rirServices.GetRIRMaterialDetail(
                    projectCode: cboProjectMaterial.Text.Trim(),
                    rirNo: "",
                    projectName: "",
                    workorderNo: "",
                    poNo: ""
                );

                // Xóa dữ liệu cũ trên lưới và reset tracking
                _modifiedRows.Clear();
                _isLoadingData = true;
                dgvRIR.Rows.Clear();

                // Đổ dữ liệu từ _dtRIRDetail vào các cột đã định nghĩa sẵn trong dgvRIR
                foreach (DataRow dr in _dtRIRDetail.Rows)
                {
                    int idx = dgvRIR.Rows.Add();
                    var row = dgvRIR.Rows[idx];

                    row.Cells["RIR_Detail_ID"].Value = dr.Table.Columns.Contains("RIR_Detail_ID") ? dr["RIR_Detail_ID"] : DBNull.Value;
                    row.Cells["RIR_No"].Value = dr.Table.Columns.Contains("RIR_No") ? dr["RIR_No"] : "";
                    row.Cells["RIR_ID"].Value = dr.Table.Columns.Contains("RIR_ID") ? dr["RIR_ID"] : "";
                    row.Cells["Item_No"].Value = dr.Table.Columns.Contains("Item_No") ? dr["Item_No"] : DBNull.Value;
                    row.Cells["Item_Name"].Value = dr.Table.Columns.Contains("Item_Name")
                        ? dr["Item_Name"]
                        : (dr.Table.Columns.Contains("item_name") ? dr["item_name"] : "");
                    row.Cells["Material"].Value = dr.Table.Columns.Contains("Material") ? dr["Material"] : "";
                    row.Cells["Size"].Value = dr.Table.Columns.Contains("Size") ? dr["Size"] : "";
                    row.Cells["UNIT"].Value = dr.Table.Columns.Contains("UNIT") ? dr["UNIT"] : "";
                    row.Cells["Qty_Required"].Value = dr.Table.Columns.Contains("Qty_Required") ? dr["Qty_Required"] : 0;
                    row.Cells["Qty_Received"].Value = dr.Table.Columns.Contains("Qty_Received") ? dr["Qty_Received"] : 0;
                    row.Cells["MTRno"].Value = dr.Table.Columns.Contains("MTRno") ? dr["MTRno"] : "";
                    row.Cells["Heatno"].Value = dr.Table.Columns.Contains("Heatno") ? dr["Heatno"] : "";
                    row.Cells["ID_Code"].Value = dr.Table.Columns.Contains("ID_Code") ? dr["ID_Code"] : "";
                    row.Cells["PO_Detail_ID"].Value = dr.Table.Columns.Contains("PO_Detail_ID") ? dr["PO_Detail_ID"] : DBNull.Value;
                    row.Cells["Inspect_Result"].Value = dr.Table.Columns.Contains("Inspect_Result") ? (dr["Inspect_Result"]?.ToString() ?? "") : "";
                    row.Cells["Remarks"].Value = dr.Table.Columns.Contains("Remarks") ? (dr["Remarks"]?.ToString() ?? "") : "";
                    row.Cells["IsAdded"].Value = "";

                    StoreOriginalValues(row);
                }

                lblStatus.Text = $"Phiếu gồm {dgvRIR.Rows.Count} dòng";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi kết nối hoặc thực thi cơ sở dữ liệu:\n{ex.Message}", "Lỗi hệ thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _isLoadingData = false;
                // Trả lại con trỏ chuột bình thường
                this.Cursor = Cursors.Default;
            }
        }

        private void StoreOriginalValues(DataGridViewRow row)
        {
            var originalValues = new Dictionary<string, string>();
            foreach (string colName in new[] { "Qty_Required", "MTRno", "Heatno", "ID_Code", "Inspect_Result", "Remarks" })
            {
                if (row.Cells[colName] != null)
                {
                    originalValues[colName] = row.Cells[colName].Value?.ToString() ?? "";
                }
            }
            row.Tag = originalValues;
        }

        private bool IsRowActuallyModified(DataGridViewRow row)
        {
            // New rows are always considered modified
            string isAdded = row.Cells["IsAdded"].Value?.ToString() ?? "";
            if (isAdded == "New") return true;

            // If there are no stored original values, assume modified
            if (!(row.Tag is Dictionary<string, string> originalValues)) return true;

            foreach (var kvp in originalValues)
            {
                string colName = kvp.Key;
                string originalVal = kvp.Value;
                string newVal = row.Cells[colName]?.Value?.ToString() ?? "";

                if (colName == "Qty_Required")
                {
                    decimal origDec = 0, newDec = 0;
                    bool parsedOrig = decimal.TryParse(originalVal, out origDec);
                    bool parsedNew = decimal.TryParse(newVal, out newDec);

                    if (parsedOrig && parsedNew)
                    {
                        if (origDec != newDec) return true;
                        continue;
                    }
                }

                if (originalVal != newVal)
                {
                    return true;
                }
            }

            return false;
        }

        private void dgvRIR_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (_isLoadingData || e.RowIndex < 0) return;

            var row = dgvRIR.Rows[e.RowIndex];
            string isAdded = row.Cells["IsAdded"].Value?.ToString() ?? "";

            if (IsRowActuallyModified(row))
            {
                _modifiedRows.Add(row);
                if (string.IsNullOrEmpty(isAdded))
                {
                    // Highlight modified existing rows in yellow
                    row.DefaultCellStyle.BackColor = Color.FromArgb(72, 190, 165);
                    row.DefaultCellStyle.SelectionBackColor = Color.FromArgb(72, 157, 92);
                }
            }
            else
            {
                _modifiedRows.Remove(row);
                if (string.IsNullOrEmpty(isAdded))
                {
                    // Revert highlight
                    row.DefaultCellStyle.BackColor = Color.Empty;
                    row.DefaultCellStyle.SelectionBackColor = Color.Empty;
                }
            }
        }

        ///// <summary>
        ///// Checks whether any rows in the grid have been modified (edited or newly added) since the last load/save.
        ///// Cleans up stale references (rows that were removed from the grid) before checking.
        ///// Returns true if there are pending changes that need to be saved.
        ///// </summary>
        //public bool HasModifiedRows()
        //{
        //    // Remove any rows that may have been deleted from the grid but still exist in the tracking set
        //    _modifiedRows.RemoveWhere(r => r.DataGridView == null || r.Index < 0);
        //    return _modifiedRows.Count > 0;
        //}

        ///// <summary>
        ///// Collects all modified/new rows from the grid, converts them to RIRDetail objects,
        ///// and persists each one to the database via RIRService.UpdateDetailForQC.
        ///// Also updates ID codes for affected PO details and resets row highlights.
        ///// Returns the number of rows successfully saved.
        ///// Throws InvalidOperationException if validation fails (e.g. missing Item_Name).
        ///// </summary>
        //public async Task<int> SaveModifiedRowsToDatabase()
        //{
        //    // Clean up stale references
        //    _modifiedRows.RemoveWhere(r => r.DataGridView == null || r.Index < 0);

        //    if (_modifiedRows.Count == 0) return 0;

        //    int savedCount = 0;

        //    // 1. Validate and save each modified row
        //    foreach (DataGridViewRow gvRow in _modifiedRows)
        //    {
        //        // Skip the placeholder new-row at the bottom of the grid
        //        if (gvRow.IsNewRow) continue;

        //        // Validate: Item_Name must not be empty
        //        if (gvRow.Cells["Item_Name"].Value == null || string.IsNullOrWhiteSpace(gvRow.Cells["Item_Name"].Value.ToString()))
        //        {
        //            throw new InvalidOperationException(
        //                $"Dòng thứ {gvRow.Index + 1} chưa nhập Tên vật tư. Vui lòng kiểm tra lại!");
        //        }

        //        // Build RIRDetail object from grid row values
        //        var detail = new RIRDetail
        //        {
        //            RIR_Detail_ID = Convert.ToInt32(gvRow.Cells["RIR_Detail_ID"].Value ?? 0),
        //            RIR_ID = Convert.ToInt32(gvRow.Cells["RIR_ID"].Value ?? 0),
        //            Item_No = Convert.ToInt32(gvRow.Cells["Item_No"].Value ?? 0),
        //            Item_Name = gvRow.Cells["Item_Name"].Value?.ToString() ?? "",
        //            Material = gvRow.Cells["Material"].Value?.ToString() ?? "",
        //            Size = gvRow.Cells["Size"].Value?.ToString() ?? "",
        //            UNIT = gvRow.Cells["UNIT"].Value?.ToString() ?? "",
        //            Qty_Required = Convert.ToDecimal(gvRow.Cells["Qty_Required"].Value ?? 0),
        //            Qty_Per_Sheet = Convert.ToDecimal(gvRow.Cells["Qty_Required"].Value ?? 0),
        //            MTRno = gvRow.Cells["MTRno"].Value?.ToString() ?? "",
        //            Heatno = gvRow.Cells["Heatno"].Value?.ToString() ?? "",
        //            ID_Code = gvRow.Cells["ID_Code"].Value?.ToString() ?? "",
        //            Inspect_Result = gvRow.Cells["Inspect_Result"].Value?.ToString() ?? "",
        //            Remarks = gvRow.Cells["Remarks"].Value?.ToString() ?? "",
        //            PO_Detail_ID = Convert.ToInt32(gvRow.Cells["PO_Detail_ID"].Value ?? 0),
        //            IsNewRow = string.IsNullOrEmpty(gvRow.Cells["IsAdded"].Value?.ToString()) ? "false" : "true",
        //        };

        //        // Persist to database (Insert or Update depending on RIR_Detail_ID)
        //        await _rirServices.UpdateDetailForQC(detail);
        //        savedCount++;
        //    }

        //    // 2. Update ID codes for affected PO details
        //    var lstPODetailId = new List<int>();
        //    foreach (DataGridViewRow item in _modifiedRows)
        //    {
        //        if (item.IsNewRow) continue;
        //        int id = Convert.ToInt32(item.Cells["PO_Detail_ID"].Value ?? 0);
        //        if (!lstPODetailId.Contains(id))
        //        {
        //            lstPODetailId.Add(id);
        //        }
        //    }
        //    _rirServices.UpdateIDFormaterial(lstPODetailId);

        //    // 3. Accept changes on the underlying DataTable if bound
        //    if (dgvRIR.DataSource is DataTable dt)
        //    {
        //        dt.AcceptChanges();
        //    }

        //    // 4. Reset row highlights back to default
        //    foreach (DataGridViewRow gvRow in _modifiedRows)
        //    {
        //        if (gvRow.DataGridView != null && gvRow.Index >= 0)
        //        {
        //            gvRow.DefaultCellStyle.BackColor = Color.Empty;
        //            gvRow.DefaultCellStyle.SelectionBackColor = Color.Empty;
        //        }
        //    }

        //    // 5. Clear tracking set after successful save
        //    _modifiedRows.Clear();

        //    return savedCount;
        //}

        private async void btnSave_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("Material Inspector Request", "Lưu chi tiết", "Lưu chi tiết")) return;
            if (!Common.Common.IsDataGridViewValid(dgvRIR)) return;

            // Clean up deleted or stale row references
            _modifiedRows.RemoveWhere(r => r.DataGridView == null || r.Index < 0);

            // 1. Kiểm tra xem có dòng nào thay đổi để lưu hay không
            if (_modifiedRows == null || _modifiedRows.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu nào thay đổi để cập nhật!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                this.Cursor = Cursors.WaitCursor;

                // Giả định bạn đã có mã ID của chứng từ cha RIR_ID (ví dụ lấy từ TextBox hoặc biến toàn cục)
                //int currentRIRHeadID = Convert.ToInt32(txtRIR_ID.Text);

                // 2. Duyệt qua tập hợp HashSet chứa các dòng đã sửa đổi/thêm mới
                foreach (DataGridViewRow gvRow in _modifiedRows)
                {
                    // Bỏ qua dòng trống mới ở dưới cùng lưới dữ liệu (nếu AllowUserToAddRows = true)
                    if (gvRow.IsNewRow) continue;

                    // Kiểm tra tính hợp lệ sơ bộ (Ví dụ: tên mặt hàng không được trống)
                    if (gvRow.Cells["Item_Name"].Value == null || string.IsNullOrWhiteSpace(gvRow.Cells["Item_Name"].Value.ToString()))
                    {
                        MessageBox.Show($"Dòng thứ {gvRow.Index + 1} chưa nhập Tên vật tư. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // Tạo đối tượng RIRDetail từ các giá trị trong DataGridViewRow
                    var d = new RIRDetail
                    {
                        RIR_Detail_ID = Convert.ToInt32(gvRow.Cells["RIR_Detail_ID"].Value ?? 0),
                        RIR_ID = Convert.ToInt32(gvRow.Cells["RIR_ID"].Value ?? 0),
                        Item_No = Convert.ToInt32(gvRow.Cells["Item_No"].Value ?? 0),
                        Item_Name = gvRow.Cells["Item_Name"].Value?.ToString() ?? "",
                        Material = gvRow.Cells["Material"].Value?.ToString() ?? "",
                        Size = gvRow.Cells["Size"].Value?.ToString() ?? "",
                        UNIT = gvRow.Cells["UNIT"].Value?.ToString() ?? "",
                        Qty_Required = Convert.ToDecimal(gvRow.Cells["Qty_Required"].Value ?? 0),
                        //Qty_Received = Convert.ToDecimal(gvRow.Cells["Qty_Received"].Value ?? 0),
                        Qty_Per_Sheet = Convert.ToDecimal(gvRow.Cells["Qty_Required"].Value ?? 0), // Same as Qty_Required for QC
                        MTRno = gvRow.Cells["MTRno"].Value?.ToString() ?? "",
                        Heatno = gvRow.Cells["Heatno"].Value?.ToString() ?? "",
                        ID_Code = gvRow.Cells["ID_Code"].Value?.ToString() ?? "",
                        Inspect_Result = gvRow.Cells["Inspect_Result"].Value?.ToString() ?? "",
                        Remarks = gvRow.Cells["Remarks"].Value?.ToString() ?? "",
                        PO_Detail_ID = Convert.ToInt32(gvRow.Cells["PO_Detail_ID"].Value ?? 0),
                        IsNewRow = string.IsNullOrEmpty(gvRow.Cells["IsAdded"].Value?.ToString()) ? "false" : "true",
                    };

                    // Gọi hàm lưu động (Thêm mới/Cập nhật tùy thuộc vào RIR_Detail_ID)
                    await _rirServices.UpdateDetailForQC(d);
                }

                var lstPODetailId = new List<int>();
                foreach (DataGridViewRow item in _modifiedRows)
                {
                    int id = Convert.ToInt32(item.Cells["PO_Detail_ID"].Value ?? 0);
                    if (!lstPODetailId.Contains(id))
                    {
                        lstPODetailId.Add(id);
                    }
                }

                _rirServices.UpdateIDFormaterial(lstPODetailId);

                // 5. Đồng bộ hóa lại trạng thái dữ liệu trên DataTable gán với Grid
                if (dgvRIR.DataSource is DataTable dt)
                {
                    dt.AcceptChanges();
                }

                // Update original values and reset IsAdded status for newly saved rows
                _isLoadingData = true;
                foreach (DataGridViewRow gvRow in _modifiedRows)
                {
                    if (gvRow.IsNewRow) continue;
                    if (gvRow.Cells["IsAdded"].Value?.ToString() == "New")
                    {
                        gvRow.Cells["IsAdded"].Value = "";
                    }
                    StoreOriginalValues(gvRow);
                }
                _isLoadingData = false;

                // 6. Xóa sạch danh sách vết trong HashSet sau khi lưu thành công hoàn toàn
                _modifiedRows.Clear();

                MessageBox.Show("Đã lưu toàn bộ các dòng thay đổi thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Xảy ra lỗi trong quá trình lưu hàng loạt:\n{ex.Message}", "Lỗi hệ thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }

            //try
            //{
            //    int saved = 0;
            //    foreach (DataGridViewRow row in dgvRIR.Rows)
            //    {
            //        string itemName = row.Cells["Item_Name"].Value?.ToString() ?? "";
            //        if (string.IsNullOrWhiteSpace(itemName)) continue;

            //        var d = new RIRDetail
            //        {
            //            RIR_Detail_ID = Convert.ToInt32(row.Cells["RIR_Detail_ID"].Value ?? 0),
            //            RIR_ID = _selectedRIR_ID,
            //            Item_No = Convert.ToInt32(row.Cells["Item_No"].Value ?? 0),
            //            Item_Name = itemName,
            //            Material = row.Cells["Material"].Value?.ToString() ?? "",
            //            Size = row.Cells["Size"].Value?.ToString() ?? "",
            //            UNIT = row.Cells["UNIT"].Value?.ToString() ?? "",
            //            Qty_Required = (int)Math.Round(Convert.ToDecimal(row.Cells["Qty_Required"].Value ?? 0)),
            //            Qty_Received = (int)Math.Round(Convert.ToDecimal(row.Cells["Qty_Received"].Value ?? 0)),
            //            Qty_Per_Sheet = (int)Math.Round(Convert.ToDecimal(row.Cells["Qty_Required"].Value ?? 0)) > 0 ? (int)Math.Round(Convert.ToDecimal(row.Cells["Qty_Required"].Value ?? 0)) : (int)Math.Round(Convert.ToDecimal(row.Cells["Qty_Received"].Value ?? 0)),
            //            MTRno = row.Cells["MTRno"].Value?.ToString() ?? "",
            //            Heatno = row.Cells["Heatno"].Value?.ToString() ?? "",
            //            ID_Code = row.Cells["ID_Code"].Value?.ToString() ?? "",
            //            Inspect_Result = row.Cells["Inspect_Result"].Value?.ToString() ?? "",
            //            Remarks = row.Cells["Remarks"].Value?.ToString() ?? "",
            //            PO_Detail_ID = Convert.ToInt32(row.Cells["PO_Detail_ID"].Value?.ToString() ?? ""),

            //            IsNewRow = string.IsNullOrEmpty(row.Cells["IsAdded"].Value?.ToString()) ? "false" : "true",
            //        };

            //        await _rirServices.UpdateDetailForQC(d);

            //        saved++;
            //    }

            //    //foreach (DataGridViewRow row in dgvRIR.Rows)
            //    //{
            //    //    int po_d_id = Convert.ToInt32(row.Cells["PO_Detail_ID"].Value?.ToString());
            //    //    var Qty_Required = (int)Math.Round(Convert.ToDecimal(row.Cells["Qty_Required"].Value ?? 0));
            //    //    var MTRno = row.Cells["MTRno"].Value?.ToString() ?? "";
            //    //    var Heatno = row.Cells["Heatno"].Value?.ToString() ?? "";
            //    //    var ID_Code = row.Cells["ID_Code"].Value?.ToString() ?? "";
            //    //    var Inspect_Result = row.Cells["Inspect_Result"].Value?.ToString() ?? "";
            //    //    var isNewRow = row.Cells["IsAdded"].Value?.ToString() ?? "";

            //    //    if (lstRootItem.Contains(po_d_id) && string.IsNullOrEmpty(isNewRow))
            //    //    {
            //    //        var lstAdd = lstItemAdd.Where(i => i.PO_Detail_ID == po_d_id).ToList();
            //    //        bool rs = await _warehouseServies.SaveQCCodeForItemOfWarehouseImportTable(_rirId, po_d_id, Qty_Required, 0, MTRno, Heatno, ID_Code, Inspect_Result, lstAdd);
            //    //    }
            //    //}

            //    // Kiểm tra nội dung khi truyền vào Procedure SQL
            //    // Tính số Weight sau khi tách

            //    MessageBox.Show($"Đã lưu {saved} dòng thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    LoadDetails(_selectedRIR_ID);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Lỗi lưu chi tiết: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
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
                _modifiedRows.Clear();
                _isLoadingData = true;

                _details = _rirServices.GetDetails(rirId);
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

                    StoreOriginalValues(row);

                    lstRootItem.Add(Convert.ToInt32(d.PO_Detail_ID));
                    _rirId = d.RIR_ID;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải chi tiết: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _isLoadingData = false;
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
            //var qtyRequireCell = Convert.ToInt32(dgvRIR.CurrentRow.Cells["Qty_Required"].Value);
            //var qtyRecivedCell = Convert.ToInt32(dgvRIR.CurrentRow.Cells["Qty_Received"].Value);

            //if (qtyRecivedCell > qtyRequireCell)
            //{
            //    MessageBox.Show("SL Thực nhận không được lớn hơn SL Yêu cầu!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    dgvRIR.CurrentCell.Value = qtyRequireCell;
            //}
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
                Title = "Lưu file ID Code List",
                InitialDirectory = Directory.Exists(@"D:\RÁC") ? @"D:\RÁC" : Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
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
            var dt = await _rirServices.GetMaterialIDCodeListOfRIRForQC(Convert.ToInt32(cboRIRs.SelectedValue?.ToString()));
            ExportIdCodeListFromDatabase(dt);
        }

        private void cboProjectMaterial_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!_isLoaded) return;
            btnSearch.PerformClick();
        }

        private async void btnPrintReportMaterial_Click(object sender, EventArgs e)
        {
            var dt = await _rirServices.GetMaterialIDCodeListOfProjectForQC(cboProjectMaterial.Text);
            ExportIdCodeListFromDatabase(dt);
        }

        private async void btnExportListItem_Click(object sender, EventArgs e)
        {
            //var rirId = Convert.ToInt32(cboRIRs.SelectedValue ?? 1);
            var rirId = Convert.ToInt32(dgvRIR.Rows[0].Cells["RIR_ID"].Value ?? 0);

            var dtHeader = await _rirServices.GetRIRHeaderByRIRId(rirId);

            //var dtDetail = _rirServices.GetDetails(rirId);
            var dtDetail = _rirServices.GetDetailsForReportEX(cboProjectMaterial.Text.Trim());

            ExportMaterialInspectionReport(dtHeader, dtDetail);
        }

        public void ExportMaterialInspectionReport(RIRHead dtHeader, List<RIRDetail> dtDetails)
        {
            // 1. Kiểm tra file Template
            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "material_receiving_inspection_report__material_template.xlsx");
            if (!File.Exists(templatePath))
            {
                MessageBox.Show("Không tìm thấy file template báo cáo nghiệm thu!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 2. Mở hộp thoại lưu file Excel mới
            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"MRI_Report_{DateTime.Now:ddMMyyyy_HHmm}.xlsx",
                Title = "Lưu báo cáo nghiệm thu vật liệu",
                InitialDirectory = Directory.Exists(@"D:\RÁC") ? @"D:\RÁC" : Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };

            if (saveDialog.ShowDialog() != DialogResult.OK) return;

            try
            {
                // Sao chép từ template sang vị trí đích mới
                File.Copy(templatePath, saveDialog.FileName, true);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(saveDialog.FileName)))
                {
                    var ws = package.Workbook.Worksheets[0]; // Lấy sheet báo cáo đầu tiên

                    // --- PHẦN 1: THAY THẾ CÁC THẺ HEADER <<...>> ---
                    // Quét vùng Header từ dòng 1 đến dòng 8 để tìm và thay thế chuỗi cấu hình mẫu
                    if (dtHeader != null)
                    {
                        ReplaceCell(ws, "<<WO_NO>>", dtHeader.WorkorderNo ?? "");
                        ReplaceCell(ws, "<<MPR_NO>>", dtHeader.MPR_No ?? "");
                        ReplaceCell(ws, "<<PROJECT_NAME>>", dtHeader.Project_Name ?? "");
                        ReplaceCell(ws, "<<CUSTOMER>>", dtHeader.Customer ?? "");
                        ReplaceCell(ws, "<<PO_NO>>", dtHeader.PONo ?? "");
                    }

                    // --- PHẦN 2: ĐỔ DỮ LIỆU CHI TIẾT BẮT ĐẦU TỪ DÒNG 9 & ĐẨY FOOTER ---
                    int startRow = 9; // Dòng bắt đầu điền dữ liệu (Dòng mẫu có STT 1)
                    int detailCount = dtDetails != null ? dtDetails.Count : 0;

                    if (detailCount > 0)
                    {
                        // Nếu số dòng dữ liệu nhiều hơn 1 dòng mẫu thiết kế sẵn, tiến hành chèn dòng hàng loạt.
                        // Lệnh này tự động đẩy phần Note và Footer chữ ký xuống phía dưới mà không làm vỡ layout.
                        if (detailCount > 1)
                        {
                            ws.InsertRow(startRow + 1, detailCount - 1, startRow);
                        }

                        // Duyệt danh sách điền dữ liệu dựa theo cấu trúc cột biểu mẫu nghiệm thu vật liệu hàn
                        for (int i = 0; i < detailCount; i++)
                        {
                            var row = dtDetails[i];
                            int currentRow = startRow + i;

                            // Gán dữ liệu (Cột 1->A, 2->B, 3->C, ...)
                            ws.Cells[currentRow, 1, currentRow, 2].Merge = true;
                            ws.Cells[currentRow, 1].Value = i + 1;                                     // Cột A -> B: No.

                            ws.Cells[currentRow, 3, currentRow, 7].Merge = true;
                            ws.Cells[currentRow, 3].Value = row.Item_Name;  // Cột C -> G : Name

                            ws.Cells[currentRow, 8, currentRow, 12].Merge = true;
                            ws.Cells[currentRow, 8].Value = row.Material;                  // Cột H -> L: spec

                            ws.Cells[currentRow, 13, currentRow, 19].Merge = true;
                            ws.Cells[currentRow, 13].Value = row.Size;                  // Cột M -> S: Size

                            ws.Cells[currentRow, 20, currentRow, 22].Merge = true;
                            ws.Cells[currentRow, 20].Value = row.UNIT;              // Cột T -> V: UNIT

                            ws.Cells[currentRow, 23, currentRow, 24].Merge = true;
                            ws.Cells[currentRow, 23].Value = row.Qty_Required;              // Cột W -> X: Qty

                            ws.Cells[currentRow, 25, currentRow, 30].Merge = true;
                            ws.Cells[currentRow, 25].Value = row.MTRno;               // Cột Y -> AD: MTC No.

                            ws.Cells[currentRow, 31, currentRow, 37].Merge = true;
                            ws.Cells[currentRow, 31].Value = row.Heatno;               // Cột AE -> AK: Heat No.

                            if (row.Inspect_Result.Equals("Pass"))
                            {
                                ws.Cells[currentRow, 38, currentRow, 44].Merge = true;
                                ws.Cells[currentRow, 38].Value = "Accept";               // Cột AL -> AR: Result.
                            }
                            else
                            {
                                ws.Cells[currentRow, 38, currentRow, 44].Merge = true;
                                ws.Cells[currentRow, 38].Value = "Reject";               // Cột AL -> AR: Result.
                            }

                            ws.Cells[currentRow, 45, currentRow, 46].Merge = true;
                            ws.Cells[currentRow, 45].Value = row.ID_Code;               // Cột Y -> AD: MTC No.

                            ws.Cells[currentRow, 47].Value = "";               // Cột Y -> AD: MTC No.
                            ws.Cells[currentRow, 48].Value = "";               // Cột Y -> AD: MTC No.

                            ws.Row(currentRow).Height = 30;
                        }
                    }

                    // Lưu dữ liệu lại vào file
                    package.Save();
                }

                MessageBox.Show($"Xuất báo cáo nghiệm thu thành công!\nSố mục vật tư: {dtDetails?.Count ?? 0}", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (MessageBox.Show($"✅ Xuất báo cáo nghiệm thu thành công!\nSố mục vật tư: {dtDetails?.Count ?? 0}\nXuất Excel thành công! Bạn có muốn mở file?", "Thành công", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = saveDialog.FileName, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi trong quá trình ghi file Excel: " + ex.Message, "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void ReplaceCell(OfficeOpenXml.ExcelWorksheet ws, string placeholder, string value)
        {
            for (int r = 1; r <= ws.Dimension.End.Row; r++)
                for (int c = 1; c <= ws.Dimension.End.Column; c++)
                    if (ws.Cells[r, c].Value?.ToString() == placeholder)
                        ws.Cells[r, c].Value = value;
        }

        private void cboProjectForPain_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!Common.Common.IsComboBoxValid(cboProjectForPain) || !_dtPaintIsLoaded) return;

            DataView dv = Common.Common.Search(cboProjectForPain.Text.Trim(), _dtProjectPaint, ["ProjectCode"]);
            dgvPaint.DataSource = dv;
        }

        private void cboProjectForWelding_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!Common.Common.IsComboBoxValid(cboProjectForWelding) || !_dtWeldingIsLoaded) return;

            DataView dv = Common.Common.Search(cboProjectForWelding.Text.Trim(), _dtProjectWelding, ["ProjectCode"]);
            dgvWelding.DataSource = dv;
        }

        private void cboRIRNoPaint_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!Common.Common.IsComboBoxValid(cboRIRNoPaint) || !_dtRIRNoPaintIsLoaded) return;
            DataView dv = Common.Common.Search(cboRIRNoPaint.Text.Trim(), _dtProjectPaint, ["RIR_No"]);
            dgvPaint.DataSource = dv;
        }

        private void cboRIRNoWelding_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!Common.Common.IsComboBoxValid(cboRIRNoWelding) || !_dtRIRNoWeldingIsLoaded) return;
            dgvWelding.DataSource = Common.Common.Search(cboRIRNoWelding.Text.Trim(), _dtProjectWelding, ["RIR_No"]);
        }

        private void btnPrintRPPaint_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show($"Bạn muốn xuất danh sách vật tư Sơn của dự án {cboProjectForPain.Text.Trim().ToUpper()} ?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                var dt = Common.Common.Search(cboProjectForPain.Text.Trim(), _dtProjectPaint, ["ProjectCode"]);
                if (dt != null)
                {
                    PrintIDCodePaintList(dt.ToTable());
                }
            }
        }

        private async void btnReportPaintOfProject_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show($"Bạn muốn xuất danh sách vật tư Sơn của dự án {cboProjectForPain.Text.Trim().ToUpper()} ?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                frmOptionPrintRP frmOptionPrintRP = new frmOptionPrintRP(cboProjectForPain.Text.Trim(), cboRIRNoPaint.Text.Trim());
                frmOptionPrintRP.ShowDialog();

                DataTable dtDetail = new DataTable();
                ProjectInfo dtHeaderByProject = new ProjectInfo();
                RIRHead dtHeaderByRIRNo = new RIRHead();

                // Get Data Following ProjectCode
                if (frmOptionPrintRP.IsProjectSelected)
                {
                    dtDetail = Common.Common.Search(cboProjectForPain.Text.Trim(), _dtProjectPaint, ["ProjectCode"]).ToTable();
                    dtHeaderByProject = _projectServices.GetByProjectCode(cboProjectForPain.Text.Trim());
                }
                // Get Data Following RIR No
                else
                {
                    dtDetail = Common.Common.Search(cboRIRNoPaint.Text.Trim(), _dtProjectPaint, ["RIR_No"]).ToTable();
                    dtHeaderByRIRNo = await _rirServices.GetRIRHeaderByRIRId(Convert.ToInt32(cboRIRNoPaint.SelectedValue?.ToString() ?? "0"));
                }

                if (dtDetail != null)
                {
                    DataTable dataTable = dtDetail.Copy();
                    List<RIRDetail> lstD = new List<RIRDetail>();
                    foreach (DataRow r in dataTable.Rows)
                    {
                        RIRDetail detail = new RIRDetail();
                        if (dataTable.Columns.Contains("RIR_Detail_ID") && r["RIR_Detail_ID"] != DBNull.Value) detail.RIR_Detail_ID = Convert.ToInt32(r["RIR_Detail_ID"]);
                        if (dataTable.Columns.Contains("RIR_ID") && r["RIR_ID"] != DBNull.Value) detail.RIR_ID = Convert.ToInt32(r["RIR_ID"]);
                        if (dataTable.Columns.Contains("PO_Detail_ID") && r["PO_Detail_ID"] != DBNull.Value) detail.PO_Detail_ID = Convert.ToInt32(r["PO_Detail_ID"]);
                        if (dataTable.Columns.Contains("Item_No") && r["Item_No"] != DBNull.Value) detail.Item_No = Convert.ToInt32(r["Item_No"]);
                        if (dataTable.Columns.Contains("Item_Name") && r["Item_Name"] != DBNull.Value) detail.Item_Name = r["Item_Name"].ToString();
                        if (dataTable.Columns.Contains("Material") && r["Material"] != DBNull.Value) detail.Material = r["Material"].ToString();
                        if (dataTable.Columns.Contains("Size") && r["Size"] != DBNull.Value) detail.Size = r["Size"].ToString();
                        if (dataTable.Columns.Contains("UNIT") && r["UNIT"] != DBNull.Value) detail.UNIT = r["UNIT"].ToString();
                        if (dataTable.Columns.Contains("Qty_Required") && r["Qty_Required"] != DBNull.Value) detail.Qty_Required = Convert.ToDecimal(r["Qty_Required"]);
                        if (dataTable.Columns.Contains("Qty_Received") && r["Qty_Received"] != DBNull.Value) detail.Qty_Received = Convert.ToDecimal(r["Qty_Received"]);
                        if (dataTable.Columns.Contains("MTRno") && r["MTRno"] != DBNull.Value) detail.MTRno = r["MTRno"].ToString();
                        if (dataTable.Columns.Contains("Heatno") && r["Heatno"] != DBNull.Value) detail.Heatno = r["Heatno"].ToString();
                        if (dataTable.Columns.Contains("ID_Code") && r["ID_Code"] != DBNull.Value) detail.ID_Code = r["ID_Code"].ToString();
                        if (dataTable.Columns.Contains("Inspect_Result") && r["Inspect_Result"] != DBNull.Value) detail.Inspect_Result = r["Inspect_Result"].ToString();
                        if (dataTable.Columns.Contains("Remarks") && r["Remarks"] != DBNull.Value) detail.Remarks = r["Remarks"].ToString();
                        if (dataTable.Columns.Contains("Created_Date") && r["Created_Date"] != DBNull.Value) detail.Created_Date = Convert.ToDateTime(r["Created_Date"]);
                        if (dataTable.Columns.Contains("Qty_Per_Sheet") && r["Qty_Per_Sheet"] != DBNull.Value) detail.Qty_Per_Sheet = Convert.ToDecimal(r["Qty_Per_Sheet"]);
                        if (dataTable.Columns.Contains("Updated_Date") && r["Updated_Date"] != DBNull.Value) detail.Updated_Date = Convert.ToDateTime(r["Updated_Date"]);
                        if (dataTable.Columns.Contains("IsNewRow") && r["IsNewRow"] != DBNull.Value) detail.IsNewRow = r["IsNewRow"].ToString();

                        lstD.Add(detail);
                    }

                    if (frmOptionPrintRP.IsProjectSelected)
                    {
                        var dtH = new RIRHead
                        {
                            WorkorderNo = dtHeaderByProject.WorkorderNo,
                            MPR_No = "",
                            Project_Name = dtHeaderByProject.ProjectName,
                            Customer = dtHeaderByProject.Customer,
                            PONo = "",
                        };
                        PrintRPPaint(dtH, lstD);
                    }
                    else
                    {
                        PrintRPPaint(dtHeaderByRIRNo, lstD);
                    }
                    //foreach (DataRow dr in dtHeader.ToTable().Rows)
                    //{
                    //    if (dataTable.Columns.Contains("WorkorderNo") && dr["WorkorderNo"] != DBNull.Value) header.WorkorderNo = dr["WorkorderNo"].ToString() ?? "";
                    //    if (dataTable.Columns.Contains("MPR_No") && dr["MPR_No"] != DBNull.Value) header.MPR_No = dr["MPR_No"].ToString() ?? "";
                    //    if (dataTable.Columns.Contains("Project_Name") && dr["Project_Name"] != DBNull.Value) header.Project_Name = dr["Project_Name"].ToString() ?? "";
                    //    if (dataTable.Columns.Contains("Customer") && dr["Customer"] != DBNull.Value) header.Customer = dr["Customer"].ToString() ?? "";
                    //    if (dataTable.Columns.Contains("PONo") && dr["PONo"] != DBNull.Value) header.PONo = dr["PONo"].ToString() ?? "";
                    //}

                }
            }
        }

        public async void PrintIDCodePaintList(DataTable dtDetails)
        {
            // 1. Kiểm tra file Template
            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "painting_consumable_list_template.xlsx");
            if (!File.Exists(templatePath))
            {
                MessageBox.Show("Không tìm thấy file template [2. Paint ID Code List.xlsx!]", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 2. Lấy dữ liệu từ SQL Server
            // Lấy thông tin Header ra biến để sử dụng
            string projectName = cboProjectForPain.Text.Trim().ToUpper();

            if (dtDetails.Rows.Count == 0)
            {
                MessageBox.Show("Không tìm thấy thông tin dự án tương ứng!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }


            // 3. Mở hộp thoại lưu file Excel mới
            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"Paint_ID_Code_List_{projectName}_{DateTime.Now:ddMMyyyy_HHmm}.xlsx",
                Title = "Lưu file ID Code List",
                InitialDirectory = Directory.Exists(@"D:\RÁC") ? @"D:\RÁC" : Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
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
                        //ws.Cells[currentRow, 10].Value = row["ID_Code"]?.ToString();
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

        public void PrintRPPaint(RIRHead dtHeader, List<RIRDetail> dtDetails)
        {
            // 1. Kiểm tra file Template
            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "material_receiving_inspection_report__material_paint_consumable_template.xlsx");
            if (!File.Exists(templatePath))
            {
                MessageBox.Show("Không tìm thấy file template báo cáo nghiệm thu!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 2. Mở hộp thoại lưu file Excel mới
            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"MRI_Report_{DateTime.Now:ddMMyyyy_HHmm}.xlsx",
                Title = "Lưu báo cáo nghiệm thu vật liệu",
                InitialDirectory = Directory.Exists(@"D:\RÁC") ? @"D:\RÁC" : Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };

            if (saveDialog.ShowDialog() != DialogResult.OK) return;

            try
            {
                // Sao chép từ template sang vị trí đích mới
                File.Copy(templatePath, saveDialog.FileName, true);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(saveDialog.FileName)))
                {
                    var ws = package.Workbook.Worksheets[0]; // Lấy sheet báo cáo đầu tiên

                    // --- PHẦN 1: THAY THẾ CÁC THẺ HEADER <<...>> ---
                    // Quét vùng Header từ dòng 1 đến dòng 8 để tìm và thay thế chuỗi cấu hình mẫu
                    if (dtHeader != null)
                    {
                        ReplaceCell(ws, "<<WO_NO>>", dtHeader.WorkorderNo ?? "");
                        ReplaceCell(ws, "<<MPR_NO>>", dtHeader.MPR_No ?? "");
                        ReplaceCell(ws, "<<PROJECT_CODE>>", dtHeader.Project_Name ?? "");
                        ReplaceCell(ws, "<<CUSTOMER>>", dtHeader.Customer ?? "");
                        ReplaceCell(ws, "<<PO_NO>>", dtHeader.PONo ?? "");
                    }

                    // --- PHẦN 2: ĐỔ DỮ LIỆU CHI TIẾT BẮT ĐẦU TỪ DÒNG 9 & ĐẨY FOOTER ---
                    int startRow = 9; // Dòng bắt đầu điền dữ liệu (Dòng mẫu có STT 1)
                    int detailCount = dtDetails != null ? dtDetails.Count : 0;

                    if (detailCount > 0)
                    {
                        // Nếu số dòng dữ liệu nhiều hơn 1 dòng mẫu thiết kế sẵn, tiến hành chèn dòng hàng loạt.
                        // Lệnh này tự động đẩy phần Note và Footer chữ ký xuống phía dưới mà không làm vỡ layout.
                        if (detailCount > 1)
                        {
                            ws.InsertRow(startRow + 1, detailCount - 1, startRow);
                        }

                        // Duyệt danh sách điền dữ liệu dựa theo cấu trúc cột biểu mẫu nghiệm thu vật liệu hàn
                        for (int i = 0; i < detailCount; i++)
                        {
                            var row = dtDetails[i];
                            int currentRow = startRow + i;

                            // Gán dữ liệu (Cột 1->A, 2->B, 3->C, ...)
                            ws.Cells[currentRow, 1, currentRow, 2].Merge = true;
                            ws.Cells[currentRow, 1].Value = i + 1;                                     // Cột A -> B: No.

                            ws.Cells[currentRow, 3, currentRow, 9].Merge = true;
                            ws.Cells[currentRow, 3].Value = row.Item_Name;  // Cột C -> G : Name

                            ws.Cells[currentRow, 10, currentRow, 17].Merge = true;
                            ws.Cells[currentRow, 10].Value = row.Material;                  // Cột H -> L: spec

                            ws.Cells[currentRow, 18, currentRow, 22].Merge = true;
                            ws.Cells[currentRow, 18].Value = row.Size;                  // Cột M -> S: Size

                            ws.Cells[currentRow, 23, currentRow, 24].Merge = true;
                            ws.Cells[currentRow, 23].Value = row.UNIT;              // Cột T -> V: UNIT

                            ws.Cells[currentRow, 25, currentRow, 26].Merge = true;
                            ws.Cells[currentRow, 25].Value = row.Qty_Required;              // Cột W -> X: Qty

                            ws.Cells[currentRow, 27, currentRow, 32].Merge = true;
                            ws.Cells[currentRow, 27].Value = row.MTRno;               // Cột Y -> AD: MTC No.

                            ws.Cells[currentRow, 33, currentRow, 38].Merge = true;
                            ws.Cells[currentRow, 33].Value = row.Heatno;               // Cột AE -> AK: Heat No.

                            ws.Cells[currentRow, 39, currentRow, 43].Merge = true;
                            ws.Cells[currentRow, 39].Value = "";

                            if (row.Inspect_Result.Equals("Pass"))
                            {
                                ws.Cells[currentRow, 44, currentRow, 48].Merge = true;
                                ws.Cells[currentRow, 44].Value = "Accept";               // Cột AL -> AR: Result.
                            }
                            else
                            {
                                ws.Cells[currentRow, 44, currentRow, 48].Merge = true;
                                ws.Cells[currentRow, 44].Value = "Reject";               // Cột AL -> AR: Result.
                            }

                            ws.Cells[currentRow, 49, currentRow, 52].Merge = true;
                            ws.Cells[currentRow, 49].Value = row.ID_Code;

                            ws.Row(currentRow).Height = 30;
                        }
                    }

                    // Lưu dữ liệu lại vào file
                    package.Save();
                }

                MessageBox.Show($"Xuất báo cáo nghiệm thu thành công!\nSố mục vật tư: {dtDetails?.Count ?? 0}", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (MessageBox.Show($"✅ Xuất báo cáo nghiệm thu thành công!\nSố mục vật tư: {dtDetails?.Count ?? 0}\nXuất Excel thành công! Bạn có muốn mở file?", "Thành công", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = saveDialog.FileName, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi trong quá trình ghi file Excel: " + ex.Message, "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void btnPrinRPWelding_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show($"Bạn muốn xuất danh sách vật tư Hàn của dự án {cboProjectForWelding.Text.Trim().ToUpper()} ?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                var dt = Common.Common.Search(cboProjectForWelding.Text.Trim(), _dtProjectWelding, ["ProjectCode"]);
                if (dt != null)
                {
                    PrintIDCodeWeldingList(dt.ToTable());
                }
            }
        }

        public async void PrintIDCodeWeldingList(DataTable dtDetails)
        {
            // 1. Kiểm tra file Template
            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "welding_consumable_list_template.xlsx");
            if (!File.Exists(templatePath))
            {
                MessageBox.Show("Không tìm thấy file template [Welding ID Code List.xlsx!]", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 2. Lấy dữ liệu từ SQL Server
            // Lấy thông tin Header ra biến để sử dụng
            string projectName = cboProjectForWelding.Text.Trim().ToUpper();

            if (dtDetails.Rows.Count == 0)
            {
                MessageBox.Show("Không tìm thấy thông tin dự án tương ứng!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }


            // 3. Mở hộp thoại lưu file Excel mới
            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"Paint_ID_Code_List_{projectName}_{DateTime.Now:ddMMyyyy_HHmm}.xlsx",
                Title = "Lưu file ID Code List",
                InitialDirectory = Directory.Exists(@"D:\RÁC") ? @"D:\RÁC" : Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
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
                        //ws.Cells[currentRow, 10].Value = row["ID_Code"]?.ToString();
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

        private async void btnReportWeldingOfProject_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show($"Bạn muốn xuất danh sách vật tư Hàn của dự án {cboProjectForWelding.Text.Trim().ToUpper()} ?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                frmOptionPrintRP frmOptionPrintRP = new frmOptionPrintRP(cboProjectForWelding.Text.Trim(), cboRIRNoWelding.Text.Trim());
                frmOptionPrintRP.ShowDialog();

                DataTable dtDetail = new DataTable();
                ProjectInfo dtHeaderByProject = new ProjectInfo();
                RIRHead dtHeaderByRIRNo = new RIRHead();

                // Get Data Following ProjectCode
                if (frmOptionPrintRP.IsProjectSelected)
                {
                    dtDetail = Common.Common.Search(cboProjectForWelding.Text.Trim(), _dtProjectWelding, ["ProjectCode"]).ToTable();
                    dtHeaderByProject = _projectServices.GetByProjectCode(cboProjectForWelding.Text.Trim());
                }
                // Get Data Following RIR No
                else
                {
                    dtDetail = Common.Common.Search(cboRIRNoWelding.Text.Trim(), _dtProjectWelding, ["RIR_No"]).ToTable();
                    dtHeaderByRIRNo = await _rirServices.GetRIRHeaderByRIRId(Convert.ToInt32(cboRIRNoWelding.SelectedValue?.ToString() ?? "0"));
                }

                if (dtDetail != null)
                {
                    DataTable dataTable = dtDetail.Copy();
                    List<RIRDetail> lstD = new List<RIRDetail>();
                    foreach (DataRow r in dataTable.Rows)
                    {
                        RIRDetail detail = new RIRDetail();
                        if (dataTable.Columns.Contains("RIR_Detail_ID") && r["RIR_Detail_ID"] != DBNull.Value) detail.RIR_Detail_ID = Convert.ToInt32(r["RIR_Detail_ID"]);
                        if (dataTable.Columns.Contains("RIR_ID") && r["RIR_ID"] != DBNull.Value) detail.RIR_ID = Convert.ToInt32(r["RIR_ID"]);
                        if (dataTable.Columns.Contains("PO_Detail_ID") && r["PO_Detail_ID"] != DBNull.Value) detail.PO_Detail_ID = Convert.ToInt32(r["PO_Detail_ID"]);
                        if (dataTable.Columns.Contains("Item_No") && r["Item_No"] != DBNull.Value) detail.Item_No = Convert.ToInt32(r["Item_No"]);
                        if (dataTable.Columns.Contains("Item_Name") && r["Item_Name"] != DBNull.Value) detail.Item_Name = r["Item_Name"].ToString();
                        if (dataTable.Columns.Contains("Material") && r["Material"] != DBNull.Value) detail.Material = r["Material"].ToString();
                        if (dataTable.Columns.Contains("Size") && r["Size"] != DBNull.Value) detail.Size = r["Size"].ToString();
                        if (dataTable.Columns.Contains("UNIT") && r["UNIT"] != DBNull.Value) detail.UNIT = r["UNIT"].ToString();
                        if (dataTable.Columns.Contains("Qty_Required") && r["Qty_Required"] != DBNull.Value) detail.Qty_Required = Convert.ToDecimal(r["Qty_Required"]);
                        if (dataTable.Columns.Contains("Qty_Received") && r["Qty_Received"] != DBNull.Value) detail.Qty_Received = Convert.ToDecimal(r["Qty_Received"]);
                        if (dataTable.Columns.Contains("MTRno") && r["MTRno"] != DBNull.Value) detail.MTRno = r["MTRno"].ToString();
                        if (dataTable.Columns.Contains("Heatno") && r["Heatno"] != DBNull.Value) detail.Heatno = r["Heatno"].ToString();
                        if (dataTable.Columns.Contains("ID_Code") && r["ID_Code"] != DBNull.Value) detail.ID_Code = r["ID_Code"].ToString();
                        if (dataTable.Columns.Contains("Inspect_Result") && r["Inspect_Result"] != DBNull.Value) detail.Inspect_Result = r["Inspect_Result"].ToString();
                        if (dataTable.Columns.Contains("Remarks") && r["Remarks"] != DBNull.Value) detail.Remarks = r["Remarks"].ToString();
                        if (dataTable.Columns.Contains("Created_Date") && r["Created_Date"] != DBNull.Value) detail.Created_Date = Convert.ToDateTime(r["Created_Date"]);
                        if (dataTable.Columns.Contains("Qty_Per_Sheet") && r["Qty_Per_Sheet"] != DBNull.Value) detail.Qty_Per_Sheet = Convert.ToDecimal(r["Qty_Per_Sheet"]);
                        if (dataTable.Columns.Contains("Updated_Date") && r["Updated_Date"] != DBNull.Value) detail.Updated_Date = Convert.ToDateTime(r["Updated_Date"]);
                        if (dataTable.Columns.Contains("IsNewRow") && r["IsNewRow"] != DBNull.Value) detail.IsNewRow = r["IsNewRow"].ToString();

                        lstD.Add(detail);
                    }

                    if (frmOptionPrintRP.IsProjectSelected)
                    {
                        var dtH = new RIRHead
                        {
                            WorkorderNo = dtHeaderByProject.WorkorderNo,
                            MPR_No = "",
                            Project_Name = dtHeaderByProject.ProjectName,
                            Customer = dtHeaderByProject.Customer,
                            PONo = "",
                        };
                        PrintRPWelding(dtH, lstD);
                    }
                    else
                    {
                        PrintRPWelding(dtHeaderByRIRNo, lstD);
                    }
                    //foreach (DataRow dr in dtHeader.ToTable().Rows)
                    //{
                    //    if (dataTable.Columns.Contains("WorkorderNo") && dr["WorkorderNo"] != DBNull.Value) header.WorkorderNo = dr["WorkorderNo"].ToString() ?? "";
                    //    if (dataTable.Columns.Contains("MPR_No") && dr["MPR_No"] != DBNull.Value) header.MPR_No = dr["MPR_No"].ToString() ?? "";
                    //    if (dataTable.Columns.Contains("Project_Name") && dr["Project_Name"] != DBNull.Value) header.Project_Name = dr["Project_Name"].ToString() ?? "";
                    //    if (dataTable.Columns.Contains("Customer") && dr["Customer"] != DBNull.Value) header.Customer = dr["Customer"].ToString() ?? "";
                    //    if (dataTable.Columns.Contains("PONo") && dr["PONo"] != DBNull.Value) header.PONo = dr["PONo"].ToString() ?? "";
                    //}

                }
            }
        }

        public void PrintRPWelding(RIRHead dtHeader, List<RIRDetail> dtDetails)
        {
            // 1. Kiểm tra file Template
            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "material_receiving_inspection_report__material_welding_consumable_template.xlsx");
            if (!File.Exists(templatePath))
            {
                MessageBox.Show("Không tìm thấy file template báo cáo nghiệm thu!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 2. Mở hộp thoại lưu file Excel mới
            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"MRI_Report_{DateTime.Now:ddMMyyyy_HHmm}.xlsx",
                Title = "Lưu báo cáo nghiệm thu vật liệu",
                InitialDirectory = Directory.Exists(@"D:\RÁC") ? @"D:\RÁC" : Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };

            if (saveDialog.ShowDialog() != DialogResult.OK) return;

            try
            {
                // Sao chép từ template sang vị trí đích mới
                File.Copy(templatePath, saveDialog.FileName, true);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(saveDialog.FileName)))
                {
                    var ws = package.Workbook.Worksheets[0]; // Lấy sheet báo cáo đầu tiên

                    // --- PHẦN 1: THAY THẾ CÁC THẺ HEADER <<...>> ---
                    // Quét vùng Header từ dòng 1 đến dòng 8 để tìm và thay thế chuỗi cấu hình mẫu
                    if (dtHeader != null)
                    {
                        ReplaceCell(ws, "<<WO_NO>>", dtHeader.WorkorderNo ?? "");
                        ReplaceCell(ws, "<<MPR_NO>>", dtHeader.MPR_No ?? "");
                        ReplaceCell(ws, "<<PROJECT_CODE>>", dtHeader.Project_Name ?? "");
                        ReplaceCell(ws, "<<CUSTOMER>>", dtHeader.Customer ?? "");
                        ReplaceCell(ws, "<<PO_NO>>", dtHeader.PONo ?? "");
                    }

                    // --- PHẦN 2: ĐỔ DỮ LIỆU CHI TIẾT BẮT ĐẦU TỪ DÒNG 9 & ĐẨY FOOTER ---
                    int startRow = 9; // Dòng bắt đầu điền dữ liệu (Dòng mẫu có STT 1)
                    int detailCount = dtDetails != null ? dtDetails.Count : 0;

                    if (detailCount > 0)
                    {
                        // Nếu số dòng dữ liệu nhiều hơn 1 dòng mẫu thiết kế sẵn, tiến hành chèn dòng hàng loạt.
                        // Lệnh này tự động đẩy phần Note và Footer chữ ký xuống phía dưới mà không làm vỡ layout.
                        if (detailCount > 1)
                        {
                            ws.InsertRow(startRow + 1, detailCount - 1, startRow);
                        }

                        // Duyệt danh sách điền dữ liệu dựa theo cấu trúc cột biểu mẫu nghiệm thu vật liệu hàn
                        for (int i = 0; i < detailCount; i++)
                        {
                            var row = dtDetails[i];
                            int currentRow = startRow + i;

                            // Gán dữ liệu (Cột 1->A, 2->B, 3->C, ...)
                            ws.Cells[currentRow, 1, currentRow, 2].Merge = true;
                            ws.Cells[currentRow, 1].Value = i + 1;                                     // Cột A -> B: No.

                            ws.Cells[currentRow, 3, currentRow, 9].Merge = true;
                            ws.Cells[currentRow, 3].Value = row.Item_Name;  // Cột C -> G : Name

                            ws.Cells[currentRow, 10, currentRow, 15].Merge = true;
                            ws.Cells[currentRow, 10].Value = row.Material;                  // Cột H -> L: spec

                            ws.Cells[currentRow, 16, currentRow, 19].Merge = true;
                            ws.Cells[currentRow, 16].Value = row.Size;                  // Cột M -> S: Size

                            ws.Cells[currentRow, 20, currentRow, 24].Merge = true;
                            ws.Cells[currentRow, 20].Value = "";              // Cột T -> X: Manufacturer

                            ws.Cells[currentRow, 25, currentRow, 26].Merge = true;
                            ws.Cells[currentRow, 25].Value = row.UNIT;              // Cột W -> X: Qty

                            ws.Cells[currentRow, 27, currentRow, 28].Merge = true;
                            ws.Cells[currentRow, 27].Value = row.Qty_Per_Sheet;               // Cột Y -> AD: MTC No.

                            ws.Cells[currentRow, 29, currentRow, 34].Merge = true;
                            ws.Cells[currentRow, 29].Value = row.MTRno;               // Cột AE -> AK: Heat No.

                            ws.Cells[currentRow, 35, currentRow, 40].Merge = true;
                            ws.Cells[currentRow, 35].Value = row.Heatno;

                            if (row.Inspect_Result.Equals("Pass"))
                            {
                                ws.Cells[currentRow, 41, currentRow, 45].Merge = true;
                                ws.Cells[currentRow, 41].Value = "Accept";               // Cột AL -> AR: Result.
                            }
                            else
                            {
                                ws.Cells[currentRow, 41, currentRow, 45].Merge = true;
                                ws.Cells[currentRow, 41].Value = "Reject";               // Cột AL -> AR: Result.
                            }

                            ws.Cells[currentRow, 46, currentRow, 49].Merge = true;
                            ws.Cells[currentRow, 46].Value = row.ID_Code;

                            ws.Row(currentRow).Height = 30;
                        }
                    }

                    // Lưu dữ liệu lại vào file
                    package.Save();
                }

                MessageBox.Show($"Xuất báo cáo nghiệm thu thành công!\nSố mục vật tư: {dtDetails?.Count ?? 0}", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (MessageBox.Show($"✅ Xuất báo cáo nghiệm thu thành công!\nSố mục vật tư: {dtDetails?.Count ?? 0}\nXuất Excel thành công! Bạn có muốn mở file?", "Thành công", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = saveDialog.FileName, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi trong quá trình ghi file Excel: " + ex.Message, "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void btnUpdateIDCode_QC_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("Material Inspector Request", "Lưu chi tiết", "Lưu chi tiết")) return;
            if (!Common.Common.IsDataGridViewValid(dgvRIR)) return;

            // Clean up deleted or stale row references
            _modifiedRows.RemoveWhere(r => r.DataGridView == null || r.Index < 0);

            // 1. Kiểm tra xem có dòng nào thay đổi để lưu hay không
            if (_modifiedRows == null || _modifiedRows.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu nào thay đổi để cập nhật!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                this.Cursor = Cursors.WaitCursor;

                // Giả định bạn đã có mã ID của chứng từ cha RIR_ID (ví dụ lấy từ TextBox hoặc biến toàn cục)
                //int currentRIRHeadID = Convert.ToInt32(txtRIR_ID.Text);

                // 2. Duyệt qua tập hợp HashSet chứa các dòng đã sửa đổi/thêm mới
                foreach (DataGridViewRow gvRow in _modifiedRows)
                {
                    // Bỏ qua dòng trống mới ở dưới cùng lưới dữ liệu (nếu AllowUserToAddRows = true)
                    if (gvRow.IsNewRow) continue;

                    // Kiểm tra tính hợp lệ sơ bộ (Ví dụ: tên mặt hàng không được trống)
                    if (gvRow.Cells["Item_Name"].Value == null || string.IsNullOrWhiteSpace(gvRow.Cells["Item_Name"].Value.ToString()))
                    {
                        MessageBox.Show($"Dòng thứ {gvRow.Index + 1} chưa nhập Tên vật tư. Vui lòng kiểm tra lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // Tạo đối tượng RIRDetail từ các giá trị trong DataGridViewRow
                    var d = new RIRDetail
                    {
                        RIR_Detail_ID = Convert.ToInt32(gvRow.Cells["RIR_Detail_ID"].Value ?? 0),
                        RIR_ID = Convert.ToInt32(gvRow.Cells["RIR_ID"].Value ?? 0),
                        Item_No = Convert.ToInt32(gvRow.Cells["Item_No"].Value ?? 0),
                        Item_Name = gvRow.Cells["Item_Name"].Value?.ToString() ?? "",
                        Material = gvRow.Cells["Material"].Value?.ToString() ?? "",
                        Size = gvRow.Cells["Size"].Value?.ToString() ?? "",
                        UNIT = gvRow.Cells["UNIT"].Value?.ToString() ?? "",
                        Qty_Required = Convert.ToDecimal(gvRow.Cells["Qty_Required"].Value ?? 0),
                        //Qty_Received = Convert.ToDecimal(gvRow.Cells["Qty_Received"].Value ?? 0),
                        Qty_Per_Sheet = Convert.ToDecimal(gvRow.Cells["Qty_Required"].Value ?? 0), // Same as Qty_Required for QC
                        MTRno = gvRow.Cells["MTRno"].Value?.ToString() ?? "",
                        Heatno = gvRow.Cells["Heatno"].Value?.ToString() ?? "",
                        ID_Code = gvRow.Cells["ID_Code"].Value?.ToString() ?? "",
                        Inspect_Result = gvRow.Cells["Inspect_Result"].Value?.ToString() ?? "",
                        Remarks = gvRow.Cells["Remarks"].Value?.ToString() ?? "",
                        PO_Detail_ID = Convert.ToInt32(gvRow.Cells["PO_Detail_ID"].Value ?? 0),
                        IsNewRow = string.IsNullOrEmpty(gvRow.Cells["IsAdded"].Value?.ToString()) ? "false" : "true",
                    };

                    // Gọi hàm lưu động (Thêm mới/Cập nhật tùy thuộc vào RIR_Detail_ID)
                    await _rirServices.UpdateDetailForQC(d);
                }

                var lstPODetailId = new List<int>();
                foreach (DataGridViewRow item in _modifiedRows)
                {
                    int id = Convert.ToInt32(item.Cells["PO_Detail_ID"].Value ?? 0);
                    _rirServices.UpdateIDFormaterialForQC(new WarehouseImport()
                    {
                        PO_Detail_ID = id,
                        QC_Code = item.Cells["ID_Code"].Value?.ToString() ?? "",
                        QC_Status = item.Cells["Inspect_Result"].Value?.ToString() ?? "",
                        Notes = item.Cells["Remarks"].Value?.ToString() ?? "",
                    });
                }

                // 5. Đồng bộ hóa lại trạng thái dữ liệu trên DataTable gán với Grid
                if (dgvRIR.DataSource is DataTable dt)
                {
                    dt.AcceptChanges();
                }

                // Update original values and reset IsAdded status for newly saved rows
                _isLoadingData = true;
                foreach (DataGridViewRow gvRow in _modifiedRows)
                {
                    if (gvRow.IsNewRow) continue;
                    if (gvRow.Cells["IsAdded"].Value?.ToString() == "New")
                    {
                        gvRow.Cells["IsAdded"].Value = "";
                    }
                    StoreOriginalValues(gvRow);
                }
                _isLoadingData = false;

                // 6. Xóa sạch danh sách vết trong HashSet sau khi lưu thành công hoàn toàn
                _modifiedRows.Clear();
                btnSearch.PerformClick();

                MessageBox.Show("Đã lưu toàn bộ các dòng thay đổi thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Xảy ra lỗi trong quá trình lưu hàng loạt:\n{ex.Message}", "Lỗi hệ thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }
    }
}