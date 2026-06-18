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

namespace MPR_Managerment.Forms.ExportGUI
{
    public partial class frmSelectWarehouse : Form
    {
        private DataTable _dtItems = new DataTable();
        private ProjectService _projectServices = new ProjectService();

        public ProjectInfo ProjectInfo { get; set; }

        public frmSelectWarehouse()
        {
            InitializeComponent();
            Common.Common.CreateDataGridView_Hide_RowHeader(dgvItems);
        }


        private async void frmSelectWarehouse_Load(object sender, EventArgs e)
        {
            await LoadItems();
        }

        private async Task LoadItems()
        {
            //_dtItems = await _productServices.GetProducts();
            //BindStockGrid(_dtItems);
            _dtItems = await _projectServices.GetProjectForCreateExport();

            //// Thêm cột logic 'Chon' vào DataTable nếu chưa có
            //if (!_dtItems.Columns.Contains("Chon"))
            //{
            //    DataColumn col = new DataColumn("Chon", typeof(bool));
            //    col.DefaultValue = false;
            //    _dtItems.Columns.Add(col);
            //}

            BindStockGrid(_dtItems);
        }

        private void BindStockGrid(DataTable models)
        {
            // KHÔNG xóa toàn bộ cột nếu đã có cột "Chon"
            if (dgvItems.Columns.Count == 0)
            {
                //DataGridViewCheckBoxColumn checkColumn = new DataGridViewCheckBoxColumn();
                //checkColumn.Name = "Chon";
                //checkColumn.HeaderText = "Chọn";
                //checkColumn.DataPropertyName = "Chon"; // Map trực tiếp với cột trong DataTable
                //dgvItems.Columns.Add(checkColumn);
            }

            // Gán DataSource là DataTable/DataView chứa cột 'Chon'
            dgvItems.DataSource = models;

            // Ẩn các cột ID và cấu hình hiển thị
            foreach (DataGridViewColumn column in dgvItems.Columns)
            {
                if (column.Name.IndexOf("id", StringComparison.OrdinalIgnoreCase) >= 0)
                    column.Visible = false;

                //// Cấu hình các cột hiển thị đẹp hơn (tùy chọn)
                //if (column.Name == "Chon") column.DisplayIndex = 0;
            }

            dgvItems.EditMode = DataGridViewEditMode.EditOnEnter;
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            Search(txtSearch.Text.Trim());
        }

        private void Search(string text)
        {
            var lstProperty = new List<string>()
            {
                "ProjectCode", "ProjectName"
            };

            DataView dv = Common.Common.Search(text, _dtItems, lstProperty);

            // Thay vì tạo bảng mới, chỉ cần cập nhật DataSource bằng View đã lọc
            dgvItems.DataSource = dv;
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dgvItems_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            ProjectInfo = new ProjectInfo
            {
                Id = Convert.ToInt32(dgvItems.Rows[e.RowIndex].Cells["Id"].Value.ToString()),
                ProjectCode = dgvItems.Rows[e.RowIndex].Cells["ProjectCode"].Value.ToString() ?? "",
                ProjectName = dgvItems.Rows[e.RowIndex].Cells["ProjectName"].Value.ToString() ?? "",
            };
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void txtSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Search(txtSearch.Text.Trim());
            }
        }
    }
}
