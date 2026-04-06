using MPR_Managerment.Common;
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

        public ucFillInvoiceNo()
        {
            InitializeComponent();
            LoadProjects();

            FormartGrid(dgvList, Color.FromArgb(0, 120, 212));
        }

        private async void LoadProjects()
        {
            var dt = await _projectService.GetProjects();
            cboProject.DisplayMember = "ProjectCode";
            cboProject.ValueMember = "ProjectCode";
            cboProject.DataSource = dt;
            _isLoaded = true;
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
            var dt = new DataTable();
            // ... các cột khác ...
            dt.Columns.Add("Export_Date", typeof(DateTime));

            dgvList.DataSource = dt;

            // --- THAY THẾ CỘT THƯỜNG THÀNH CỘT LỊCH ---
            if (dgvList.Columns.Contains("InvoiceDate"))
            {
                int columnIndex = dgvList.Columns["InvoiceDate"].Index;
                dgvList.Columns.RemoveAt(columnIndex);

                DataGridViewCalendarColumn calCol = new DataGridViewCalendarColumn();
                calCol.Name = "InvoiceDate";
                calCol.HeaderText = "📅 Ngày Xuất";
                calCol.DataPropertyName = "InvoiceDate"; // Map với DataTable
                calCol.Width = 120;

                dgvList.Columns.Insert(columnIndex, calCol);
            }
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
            var dt = await _warehouseServices.GetWarehouseImportByPOId(Convert.ToInt32(cboPO.SelectedValue.ToString()));
            dgvList.DataSource = dt;
        }

        private void cboProject_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!_isLoaded) return;
            LoadPOByProjectCode(cboProject.SelectedValue.ToString());
        }

        private void ucFillInvoiceNo_Load(object sender, EventArgs e)
        {
            // Thêm dtp vào Grid và ẩn đi
            dgvList.Controls.Add(dtp);
            dtp.Visible = false;
            dtp.Format = DateTimePickerFormat.Custom;
            dtp.CustomFormat = "dd-MM-yyyy";

            // Sự kiện khi chọn ngày xong
            dtp.ValueChanged += (s, ev) =>
            {
                dgvList.CurrentCell.Value = dtp.Value.ToString("dd-MM-yyyy");
            };
        }

        private void dgvList_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // Kiểm tra nếu click vào cột Ngày (Giả sử cột 2 là cột ngày)
            if (e.RowIndex >= 0 && dgvList.Columns[e.ColumnIndex].Name == "InvoiceDate")
            {
                // Lấy tọa độ ô đang chọn
                Rectangle rect = dgvList.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);

                dtp.Size = new Size(rect.Width, rect.Height);
                dtp.Location = new Point(rect.X, rect.Y);
                dtp.Visible = true;
            }
            else
            {
                dtp.Visible = false;
            }
        }

        private void dgvList_Scroll(object sender, ScrollEventArgs e)
        {
            dtp.Visible = false;
        }
    }
}
