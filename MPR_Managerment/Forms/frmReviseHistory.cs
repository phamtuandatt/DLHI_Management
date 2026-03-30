using System;
using System.Drawing;
using System.Windows.Forms;
using MPR_Managerment.Helpers;
using Microsoft.Data.SqlClient;

namespace MPR_Managerment.Forms
{
    public partial class frmReviseHistory : Form
    {
        private string _poNo;
        private DataGridView dgvHistory;

        public frmReviseHistory() : this("") { }

        public frmReviseHistory(string poNo)
        {
            _poNo = poNo;
            this.Text = $"Lịch sử Revise — {_poNo}";
            this.Size = new Size(900, 500);
            this.StartPosition = FormStartPosition.CenterParent;
            this.BackColor = Color.White;
            BuildUI();
            LoadHistory();
        }

        private void BuildUI()
        {
            this.Controls.Add(new Label
            {
                Text = $"LỊCH SỬ THAY ĐỔI — {_poNo}",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 10),
                Size = new Size(600, 30)
            });

            dgvHistory = new DataGridView
            {
                Location = new Point(10, 50),
                Size = new Size(860, 390),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            dgvHistory.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvHistory.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvHistory.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvHistory.EnableHeadersVisualStyles = false;
            dgvHistory.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            this.Controls.Add(dgvHistory);
        }

        private void LoadHistory()
        {
            try
            {
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var cmd = new SqlCommand(@"
                        SELECT 
                            Transaction_ID  AS [ID],
                            PONo            AS [PO No],
                            Old_Revise      AS [Revise cũ],
                            New_Revise      AS [Revise mới],
                            Changed_By      AS [Người thay đổi],
                            Changed_Date    AS [Ngày thay đổi],
                            Reason          AS [Lý do]
                        FROM PO_Revise_Transactions
                        WHERE PONo = @poNo
                        ORDER BY Changed_Date DESC", conn);
                    cmd.Parameters.AddWithValue("@poNo", _poNo);

                    var dt = new System.Data.DataTable();
                    dt.Load(cmd.ExecuteReader());

                    if (dt.Rows.Count == 0)
                    {
                        this.Controls.Add(new Label
                        {
                            Text = "Chưa có lịch sử thay đổi nào cho PO này.",
                            Location = new Point(10, 250),
                            Size = new Size(500, 30),
                            Font = new Font("Segoe UI", 10),
                            ForeColor = Color.Gray
                        });
                    }
                    else
                    {
                        dgvHistory.DataSource = dt;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải lịch sử: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}