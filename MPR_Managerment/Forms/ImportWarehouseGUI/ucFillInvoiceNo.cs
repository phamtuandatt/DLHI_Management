using MPR_Managerment.Common;
using MPR_Managerment.Models;
using MPR_Managerment.Services;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MPR_Managerment.Forms.ImportWarehouseGUI
{
    public partial class ucFillInvoiceNo : UserControl
    {
        private ProjectService _projectService = new ProjectService();
        private POService _poServices = new POService();
        private WarehouseService _warehouseServices = new WarehouseService();
        private bool _isLoaded = false;
        private bool _isPOLoaded = false;
        DateTimePicker dtp = new DateTimePicker();
        private DataTable dtSelected = new DataTable();

        public ucFillInvoiceNo()
        {
            InitializeComponent();
            LoadProjects();

            InitGridSelected();
            FormartGrid(dgvList, Color.FromArgb(0, 120, 212));
        }

        private async void LoadProjects()
        {
            lblStatus.Visible = true;
            var dt = await _projectService.GetProjects();
            cboProject.DisplayMember = "ProjectCode";
            cboProject.ValueMember = "ProjectCode";
            cboProject.DataSource = dt;
            _isLoaded = true;
        }

        private void FormatLableStatus()
        {
            int count = dgvList.Rows.Count;
            lblStatus.Text = $" ✅ PO có: {count} Item(s) chưa có hóa đơn";
            lblStatus.ForeColor = Color.FromArgb(120, 158, 41);    // Xanh đậm (Prussian Blue)
            //lblStatus.BackColor = Color.FromArgb(120, 158, 41); // Nền xanh nhạt (Light Sky Blue)
            lblStatus.Font = new Font("Segoe UI", 9, FontStyle.Bold);
        }

        private void FormartGrid(DataGridView dataGridView, Color colorHeader)
        {
            dataGridView.ReadOnly = true;
            dataGridView.AllowUserToAddRows = false;
            //dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView.BackgroundColor = Color.White;
            dataGridView.BorderStyle = BorderStyle.FixedSingle;
            dataGridView.RowHeadersVisible = false;
            dataGridView.Font = new Font("Segoe UI", 9);
            dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            dataGridView.ColumnHeadersDefaultCellStyle.BackColor = colorHeader;
            dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dataGridView.EnableHeadersVisualStyles = false;
            dataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
        }

        private void InitGridSelected()
        {
            //var dt = new DataTable();
            //// ... các cột khác ...
            ////dt.Columns.Add("Export_Date", typeof(DateTime));

            //dgvList.DataSource = dt;

            //// --- THAY THẾ CỘT THƯỜNG THÀNH CỘT LỊCH ---
            //if (dgvList.Columns.Contains("InvoiceDate"))
            //{
            //    int columnIndex = dgvList.Columns["InvoiceDate"].Index;
            //    dgvList.Columns.RemoveAt(columnIndex);

            //    DataGridViewCalendarColumn calCol = new DataGridViewCalendarColumn();
            //    calCol.Name = "InvoiceDate";
            //    calCol.HeaderText = "📅 Ngày Xuất";
            //    calCol.DataPropertyName = "InvoiceDate"; // Map với DataTable
            //    calCol.Width = 120;

            //    dgvList.Columns.Insert(columnIndex, calCol);
            //}
        }


        private async void LoadPOByProjectCode(string projectCode)
        {
            var dt = await _poServices.GetPOByProjectCode(projectCode);
            cboPO.DisplayMember = "PONo";
            cboPO.ValueMember = "PO_ID";
            cboPO.DataSource = dt;
            _isPOLoaded = true;
        }

        private async void btnSearch_Click(object sender, EventArgs e)
        {
            if (!_isPOLoaded || !_isLoaded) return;
            dtSelected = await _warehouseServices.GetWarehouseImportByPOId(Convert.ToInt32(cboPO.SelectedValue.ToString()));
            dgvList.DataSource = dtSelected;
            if (dgvList.Rows.Count > 0 && dtSelected.Rows.Count > 0)
            {
                SetupGridEditable();
                //FormatLableStatus();
            }
        }

        private void cboProject_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!_isLoaded) return;
            LoadPOByProjectCode(cboProject.SelectedValue.ToString());
        }

        private void ucFillInvoiceNo_Load(object sender, EventArgs e)
        {
            //// Thêm dtp vào Grid và ẩn đi
            //dgvList.Controls.Add(dtp);
            //dtp.Visible = false;
            //dtp.Format = DateTimePickerFormat.Custom;
            //dtp.CustomFormat = "dd-MM-yyyy";

            //// Sự kiện khi chọn ngày xong
            //dtp.ValueChanged += (s, ev) =>
            //{
            //    dgvList.CurrentCell.Value = dtp.Value.ToString("dd-MM-yyyy");
            //};
        }

        private void dgvList_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //// Kiểm tra nếu click vào cột Ngày (Giả sử cột 2 là cột ngày)
            //if (e.RowIndex >= 0 && dgvList.Columns[e.ColumnIndex].Name == "InvoiceDate")
            //{
            //    // Lấy tọa độ ô đang chọn
            //    Rectangle rect = dgvList.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);

            //    dtp.Size = new Size(rect.Width, rect.Height);
            //    dtp.Location = new Point(rect.X, rect.Y);
            //    dtp.Visible = true;
            //}
            //else
            //{
            //    dtp.Visible = false;
            //}
        }

        private void SetupGridEditable()
        {
            // 1. Cho phép Grid có thể chỉnh sửa (Tổng thể)
            dgvList.ReadOnly = false;

            // 2. Khóa tất cả các cột trước
            foreach (DataGridViewColumn col in dgvList.Columns)
            {
                col.ReadOnly = true;

                // Đổi màu nền các cột bị khóa để người dùng biết (Tùy chọn)
                col.DefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);
            }

            // 3. Mở khóa duy nhất cột invoice no
            if (dgvList.Columns.Contains("InvoiceNo"))
            {
                dgvList.Columns["InvoiceNo"].ReadOnly = false;

                // Đổi màu nền cột được sửa thành màu trắng hoặc vàng nhạt để nổi bật
                dgvList.Columns["InvoiceNo"].DefaultCellStyle.BackColor = Color.White;

                // Đổi màu chữ để báo hiệu ô này "sống"
                dgvList.Columns["InvoiceNo"].DefaultCellStyle.ForeColor = Color.Blue;
            }
            if (dgvList.Columns.Contains("InvoiceDate"))
            {
                dgvList.Columns["InvoiceDate"].ReadOnly = false;

                // Đổi màu nền cột được sửa thành màu trắng hoặc vàng nhạt để nổi bật
                dgvList.Columns["InvoiceDate"].DefaultCellStyle.BackColor = Color.White;

                // Đổi màu chữ để báo hiệu ô này "sống"
                dgvList.Columns["InvoiceDate"].DefaultCellStyle.ForeColor = Color.Blue;
            }
        }

        private void dgvList_Scroll(object sender, ScrollEventArgs e)
        {
            dtp.Visible = false;
        }

        private void dgvList_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnCancelSer_Click(object sender, EventArgs e)
        {

        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            if (dgvList.Rows.Count <= 0) return;
            var result = MessageBox.Show("Bạn có muốn xóa toàn bộ số lượng đã nhập và reset ngày tháng không?",
                                         "Xác nhận làm mới", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                // 2. Duyệt qua từng dòng trong DataTable gắn với Grid
                foreach (DataRow row in dtSelected.Rows)
                {
                    // Reset cột Số lượng (cột bạn cho phép nhập)
                    if (dtSelected.Columns.Contains("InvoiceNo"))
                    {
                        row["InvoiceNo"] = 0;
                    }

                    // Reset cột Ngày tháng (cột bạn cho phép chọn ngày)
                    if (dtSelected.Columns.Contains("InvoiceDate"))
                    {
                        row["InvoiceDate"] = DateTime.Now;
                    }

                    //// Nếu bạn có cột Ghi chú (Notes) cũng cho phép nhập:
                    //if (dtSelected.Columns.Contains("Notes"))
                    //{
                    //    row["Notes"] = DBNull.Value;
                    //}
                }

                // 3. Cập nhật lại giao diện
                dgvList.Refresh();

                MessageBox.Show("Đã làm mới các ô nhập liệu!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dgvList_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            // Kiểm tra xem có đúng là cột Số lượng đang được sửa không
            if (dgvList.CurrentCell.OwningColumn.Name == "InvoiceNo")
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    // Quan trọng: Gỡ bỏ sự kiện cũ trước khi gán mới để tránh bị lặp lại nhiều lần
                    tb.KeyPress -= OnlyNumber_KeyPress;
                    tb.KeyPress += OnlyNumber_KeyPress;
                }
            }
        }

        private void OnlyNumber_KeyPress(object? sender, KeyPressEventArgs e)
        {
            // Chỉ cho phép: Số (0-9), Phím xóa (BackSpace), và dấu chấm thập phân (.)
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true; // Từ chối ký tự này
            }

            // Kiểm tra nếu đã có dấu chấm rồi thì không cho nhập thêm dấu chấm thứ 2
            TextBox txt = sender as TextBox;
            if ((e.KeyChar == '.') && (txt.Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void btnSaveInvoice_Click(object sender, EventArgs e)
        {
            if (dgvList.Rows.Count <= 0) return;

            try
            {
                var count = 0;
                foreach (DataGridViewRow item in dgvList.Rows)
                {
                    var invoiceNo = item.Cells[0].Value.ToString();
                    var invoidDate = item.Cells[1].Value.ToString();
                    if (!string.IsNullOrEmpty(invoiceNo) && !string.IsNullOrEmpty(invoidDate))
                    {
                        var poId = item.Cells["PO_ID"].Value.ToString();
                        var importID = item.Cells["Import_ID"].Value.ToString();

                        var wM = new WarehouseImport()
                        {
                            InvoiceNo = invoiceNo,
                            InvoiceDate = invoidDate,
                            PO_ID = Convert.ToInt32(poId),
                            Import_ID = Convert.ToInt32(importID)
                        };
                        _warehouseServices.Update(wM);
                        count++;
                    }
                }

                MessageBox.Show($"Đã làm cập nhật hóa đơn cho {count} vật tư thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                dtSelected.Rows.Clear();
                dgvList.Refresh();
                lblStatus.Text = "";
                dtp.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Cập nhật hóa đơn thất bại!\nLỗi: {ex.Message}", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnFill_Click(object sender, EventArgs e)
        {
            FillBulkData("InvoiceNo", txtInvoiceNo.Text.Trim());
            FillBulkData("InvoiceDate", dtDate.Value.ToString("dd/MM/yyyy"));
        }

        private void FillBulkData(string columnName, string value)
        {
            if (dgvList.Rows.Count == 0) return;

            foreach (DataGridViewRow row in dgvList.Rows)
            {
                //if (row.IsNewRow) continue;
                if (row.Visible) // Chỉ điền cho các dòng đang hiển thị (đã lọc)
                {
                    row.Cells[columnName].Value = value;
                }
            }
        }
    }
}
