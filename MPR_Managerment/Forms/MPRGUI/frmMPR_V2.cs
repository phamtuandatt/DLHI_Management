using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.DirectoryServices.ActiveDirectory;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MPR_Managerment.Forms.MPRGUI
{
    public partial class frmMPR_V2 : Form
    {
        private DataTable _dtItems = new DataTable();

        public frmMPR_V2()
        {
            this.Text = "Yêu cầu mua hàng mới";
            this.Size = new Size(1100, 800);
            this.MinimumSize = new Size(1000, 700);
            this.BackColor = Color.White;
            this.Font = new Font("Segoe UI", 9.5f);

            InitializeComponent();

            Common.Common.CreateButtonPrint(btnPrint);
            Common.Common.CreateButtonSave(btnSave);
            Common.Common.CreateButtonCancel(btnCancel, "");
            Common.Common.CreateButtonAdd(btnAddRow);
            Common.Common.CreateButtonDelete(btnDeleteRow, "");
            BuidDataSourceGrid();
 
            lblTotal.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            lblTotal.BackColor = Color.FromArgb(254, 0, 51);
            lblTotal.ForeColor = Color.White;
        }

        private void BuidDataSourceGrid()
        {
            Common.Common.CreateDataGridView(dgvItems);

            dgvItems.Columns.Clear();
            dgvItems.Columns.Add(new DataGridViewTextBoxColumn { Name = "Item_Name", HeaderText = "Tên vật tư", Width = 220, ReadOnly = true });
            dgvItems.Columns.Add(new DataGridViewTextBoxColumn { Name = "Description", HeaderText = "Mô tả", Width = 220 });
            dgvItems.Columns.Add(new DataGridViewTextBoxColumn { Name = "Material", HeaderText = "Tiêu chuẩn", Width = 110, ReadOnly = true });
            dgvItems.Columns.Add(new DataGridViewTextBoxColumn { Name = "A_Thinkness", HeaderText = "A_Thinkness", Width = 70 });
            dgvItems.Columns.Add(new DataGridViewTextBoxColumn { Name = "B_Depth", HeaderText = "B_Depth", Width = 70 });
            dgvItems.Columns.Add(new DataGridViewTextBoxColumn { Name = "C_Width", HeaderText = "C_Width", Width = 70 });
            dgvItems.Columns.Add(new DataGridViewTextBoxColumn { Name = "D_Web", HeaderText = "D_Web", Width = 70,  });
            dgvItems.Columns.Add(new DataGridViewTextBoxColumn { Name = "E_Flag", HeaderText = "E_Flag", Width = 70 });
            dgvItems.Columns.Add(new DataGridViewTextBoxColumn { Name = "F_Length", HeaderText = "F_Length", Width = 100 });
            dgvItems.Columns.Add(new DataGridViewTextBoxColumn { Name = "Usage", HeaderText = "Usage", Width = 160 });
            dgvItems.Columns.Add(new DataGridViewTextBoxColumn { Name = "Unit", HeaderText = "Unit", Width = 150, ReadOnly = true });
            dgvItems.Columns.Add(new DataGridViewTextBoxColumn { Name = "Qty", HeaderText = "Qty", Width = 150 });
            dgvItems.Columns.Add(new DataGridViewTextBoxColumn { Name = "Weights", HeaderText = "Weights (kg)", Width = 170 });
            dgvItems.Columns.Add(new DataGridViewTextBoxColumn { Name = "Remarks", HeaderText = "Remarks", Width = 200 });


            dgvItems.CellFormatting += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                string col = dgvItems.Columns[e.ColumnIndex].Name;
                if (col == "Qty")
                {
                    decimal val = e.Value != null ? Convert.ToDecimal(e.Value) : 0;
                    e.CellStyle.ForeColor = Color.FromArgb(40, 167, 69);
                    e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                }
            };

            dgvItems.EditingControlShowing += (s, e) =>
            {
                e.Control.KeyPress -= new KeyPressEventHandler(Common.Common.Column_KeyPress_Digital);
                if (dgvItems.CurrentCell.ColumnIndex == dgvItems.Columns["Qty"].Index)
                {
                    TextBox tb = e.Control as TextBox;
                    if (tb != null)
                    {
                        tb.KeyPress += new KeyPressEventHandler(Common.Common.Column_KeyPress_Digital);
                    }
                }
            };

            dgvItems.CellEndEdit += (s, e) =>
            {
                // Chỉ kiểm tra nếu cột đang sửa là "SL_Xuat"
                if (dgvItems.Columns[e.ColumnIndex].Name == "Qty")
                {
                    var row = dgvItems.Rows[e.RowIndex];

                    // Lấy giá trị nhập vào và giá trị tồn
                    decimal slNhap = 0;

                    // Ép kiểu an toàn (sử dụng decimal.TryParse để tránh lỗi nhập chữ)
                    decimal.TryParse(row.Cells["Qty"].Value?.ToString() ?? "0", out slNhap);

                    if (slNhap == 0)
                    {
                        // Gán lại giá trị Xuất bằng giá trị Tồn
                        row.Cells["Qty"].Value = 1;
                    }
                }
            };

        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            dgvItems.Rows.Clear();
            dgvItems.Refresh();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (!Common.Common.IsDataGridViewValid(dgvItems)) return;
            if (MessageBox.Show($"Bạn có muốn hủy thao tác hiện tại không?", "Thông báo",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                dgvItems.Rows.Clear();
                dgvItems.Refresh();
            }
        }

        private void btnDeleteRow_Click(object sender, EventArgs e)
        {

        }

        private void btnAddRow_Click(object sender, EventArgs e)
        {
            int rowIndex = dgvItems.Rows.Add();

            // 2. Tùy chọn: Focus vào ô đầu tiên của dòng mới để người dùng nhập liệu ngay
            dgvItems.CurrentCell = dgvItems.Rows[rowIndex].Cells[0];
            dgvItems.BeginEdit(true);
        }
    }
}
