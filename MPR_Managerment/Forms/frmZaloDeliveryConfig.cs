using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;

namespace MPR_Managerment.Forms
{
    public class ZaloDeliveryItem
    {
        public int DetailId { get; set; }
        public string ItemName { get; set; }
        public string Material { get; set; }
        public string Asize { get; set; }
        public string Bsize { get; set; }
        public string Csize { get; set; }
        public string Unit { get; set; }
        public decimal TotalQty { get; set; }
        public decimal DeliveredQty { get; set; }
        public decimal RemainingQty { get; set; }
        public decimal DeliveryQtyNow { get; set; }
    }

    public partial class frmZaloDeliveryConfig : Form
    {
        private int _poId;
        private DateTimePicker dtDeliveryDate;
        private DataGridView dgvItems;
        private Button btnOk, btnCancel;
        private Label lblTitle;

        public DateTime SelectedDate { get; private set; }
        public List<ZaloDeliveryItem> SelectedItems { get; private set; } = new List<ZaloDeliveryItem>();

        public frmZaloDeliveryConfig(int poId)
        {
            // Gắn AI Trợ lý — hiển thị nút floating + chat panel
            frmAIChat.Attach(this);
            _poId = poId;
            InitializeComponent();
            LoadData();
        }

        private void InitializeComponent()
        {
            this.Text = "Cấu hình giao hàng Zalo";
            this.Size = new Size(700, 500);
            this.StartPosition = FormStartPosition.CenterParent;
            this.BackColor = Color.FromArgb(245, 245, 245);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            lblTitle = new Label
            {
                Text = "Xác nhận thông tin giao hàng",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                Location = new Point(20, 20),
                Size = new Size(400, 25),
                ForeColor = Color.FromArgb(0, 120, 212)
            };

            var lblDate = new Label
            {
                Text = "Ngày giao hàng:",
                Location = new Point(20, 60),
                Size = new Size(120, 20),
                Font = new Font("Segoe UI", 9)
            };

            dtDeliveryDate = new DateTimePicker
            {
                Location = new Point(140, 57),
                Size = new Size(200, 25),
                Format = DateTimePickerFormat.Short
            };

            dgvItems = new DataGridView
            {
                Location = new Point(20, 100),
                Size = new Size(640, 300),
                ReadOnly = false,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            dgvItems.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvItems.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvItems.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvItems.EnableHeadersVisualStyles = false;

            btnOk = new Button
            {
                Text = "Xác nhận & Gửi",
                Location = new Point(460, 415),
                Size = new Size(110, 35),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnOk.FlatAppearance.BorderSize = 0;
            btnOk.Click += BtnOk_Click;

            btnCancel = new Button
            {
                Text = "Hủy",
                Location = new Point(580, 415),
                Size = new Size(80, 35),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnCancel.FlatAppearance.BorderSize = 0;
            btnCancel.Click += (s, e) => this.Close();

            this.Controls.Add(lblTitle);
            this.Controls.Add(lblDate);
            this.Controls.Add(dtDeliveryDate);
            this.Controls.Add(dgvItems);
            this.Controls.Add(btnOk);
            this.Controls.Add(btnCancel);
        }

        private void LoadData()
        {
            try
            {
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    string sql = @"
                        SELECT 
                            pod.PO_Detail_ID, 
                            pod.item_name, 
                            pod.Material,
                            pod.Asize, pod.Bsize, pod.Csize,
                            pod.Unit,
                            pod.Qty_Per_Sheet AS TotalQty,
                            ISNULL((SELECT SUM(Qty_Import) FROM Warehouse_Import WHERE PO_Detail_ID = pod.PO_Detail_ID), 0) AS DeliveredQty
                        FROM PO_Detail pod
                        WHERE pod.PO_ID = @poId";

                    using var cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@poId", _poId);
                    using var dr = cmd.ExecuteReader();
                    
                    DataTable dt = new DataTable();
                    dt.Columns.Add("ID", typeof(int));
                    dt.Columns.Add("Tên hàng", typeof(string));
                    dt.Columns.Add("Vật liệu", typeof(string));
                    dt.Columns.Add("A", typeof(string));
                    dt.Columns.Add("B", typeof(string));
                    dt.Columns.Add("C", typeof(string));
                    dt.Columns.Add("ĐVT", typeof(string));
                    dt.Columns.Add("Tổng SL", typeof(decimal));
                    dt.Columns.Add("Đã giao", typeof(decimal));
                    dt.Columns.Add("Còn lại", typeof(decimal));
                    dt.Columns.Add("Giao lần này", typeof(decimal));

                    while (dr.Read())
                    {
                        int id = Convert.ToInt32(dr["PO_Detail_ID"]);
                        string name = dr["item_name"]?.ToString() ?? "N/A";
                        string mat = dr["Material"]?.ToString() ?? "";
                        string a = dr["Asize"]?.ToString() ?? "";
                        string b = dr["Bsize"]?.ToString() ?? "";
                        string c = dr["Csize"]?.ToString() ?? "";
                        string unit = dr["Unit"]?.ToString() ?? "";
                        decimal total = Convert.ToDecimal(dr["TotalQty"]);
                        decimal delivered = Convert.ToDecimal(dr["DeliveredQty"]);
                        decimal remaining = total - delivered;

                        dt.Rows.Add(id, name, mat, a, b, c, unit, total, delivered, remaining, remaining);
                    }
                    dgvItems.DataSource = dt;
                    dgvItems.Columns["ID"].Visible = false;
                    
                    // Lock read-only columns
                    dgvItems.Columns["Tên hàng"].ReadOnly = true;
                    dgvItems.Columns["Vật liệu"].ReadOnly = true;
                    dgvItems.Columns["A"].ReadOnly = true;
                    dgvItems.Columns["B"].ReadOnly = true;
                    dgvItems.Columns["C"].ReadOnly = true;
                    dgvItems.Columns["ĐVT"].ReadOnly = true;
                    dgvItems.Columns["Tổng SL"].ReadOnly = true;
                    dgvItems.Columns["Đã giao"].ReadOnly = true;
                    dgvItems.Columns["Còn lại"].ReadOnly = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            SelectedDate = dtDeliveryDate.Value;
            SelectedItems.Clear();

            foreach (DataGridViewRow row in dgvItems.Rows)
            {
                if (row.IsNewRow) continue;
                
                SelectedItems.Add(new ZaloDeliveryItem
                {
                    DetailId = Convert.ToInt32(row.Cells["ID"].Value),
                    ItemName = row.Cells["Tên hàng"].Value?.ToString() ?? "",
                    Material = row.Cells["Vật liệu"].Value?.ToString() ?? "",
                    Asize = row.Cells["A"].Value?.ToString() ?? "",
                    Bsize = row.Cells["B"].Value?.ToString() ?? "",
                    Csize = row.Cells["C"].Value?.ToString() ?? "",
                    Unit = row.Cells["ĐVT"].Value?.ToString() ?? "",
                    TotalQty = Convert.ToDecimal(row.Cells["Tổng SL"].Value),
                    DeliveredQty = Convert.ToDecimal(row.Cells["Đã giao"].Value),
                    RemainingQty = Convert.ToDecimal(row.Cells["Còn lại"].Value),
                    DeliveryQtyNow = Convert.ToDecimal(row.Cells["Giao lần này"].Value)
                });
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}