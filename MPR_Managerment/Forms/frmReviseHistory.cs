using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;

namespace MPR_Managerment
{
    public class frmReviseHistory : Form
    {
        private readonly string _poNo;
        private DataGridView dgv;
        private Label lblTitle, lblCount;
        private TextBox txtSearch;
        private DataTable _dtFull;

        public frmReviseHistory(string poNo)
        {
            _poNo = poNo;
            BuildUI();
            LoadData();
        }

        private void BuildUI()
        {
            this.Text = $"📋 Revise History — {_poNo}";
            this.Size = new Size(1000, 600);
            this.MinimumSize = new Size(750, 400);
            this.StartPosition = FormStartPosition.CenterParent;
            this.BackColor = Color.FromArgb(245, 245, 245);
            this.KeyPreview = true;

            // ── Tiêu đề ──
            lblTitle = new Label
            {
                Text = $"📋  LỊCH SỬ THAY ĐỔI PO:  {_poNo}",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = Color.FromArgb(102, 51, 153),
                Location = new Point(10, 10),
                Size = new Size(700, 26),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            this.Controls.Add(lblTitle);

            // ── Filter bar ──
            var pFilter = new Panel
            {
                Location = new Point(10, 42),
                Size = new Size(this.ClientSize.Width - 20, 34),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(pFilter);

            pFilter.Controls.Add(new Label
            {
                Text = "Tìm kiếm:",
                Location = new Point(6, 8),
                Size = new Size(70, 20),
                Font = new Font("Segoe UI", 9),
                TextAlign = ContentAlignment.MiddleLeft
            });
            txtSearch = new TextBox
            {
                Location = new Point(80, 6),
                Size = new Size(220, 24),
                Font = new Font("Segoe UI", 9),
                PlaceholderText = "Cột thay đổi, giá trị..."
            };
            pFilter.Controls.Add(txtSearch);

            var btnSearch = new Button
            {
                Text = "🔍 Tìm",
                Location = new Point(310, 4),
                Size = new Size(80, 26),
                BackColor = Color.FromArgb(102, 51, 153),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnSearch.FlatAppearance.BorderSize = 0;
            btnSearch.Click += (s, e) => ApplyFilter();
            pFilter.Controls.Add(btnSearch);

            var btnClear = new Button
            {
                Text = "✖ Xóa lọc",
                Location = new Point(400, 4),
                Size = new Size(80, 26),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnClear.FlatAppearance.BorderSize = 0;
            btnClear.Click += (s, e) => { txtSearch.Text = ""; ApplyFilter(); };
            pFilter.Controls.Add(btnClear);

            // ── Label thống kê ──
            lblCount = new Label
            {
                Text = "",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 82),
                Size = new Size(600, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            this.Controls.Add(lblCount);

            // ── DataGridView ──
            dgv = new DataGridView
            {
                Location = new Point(10, 106),
                Size = new Size(this.ClientSize.Width - 20, this.ClientSize.Height - 160),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
                                    | AnchorStyles.Right | AnchorStyles.Bottom
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(102, 51, 153);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.EnableHeadersVisualStyles = false;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 240, 255);
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(204, 232, 255);
            dgv.DefaultCellStyle.SelectionForeColor = Color.Black;

            // Tô màu cột old/new value
            dgv.CellFormatting += (s, ev) =>
            {
                if (ev.RowIndex < 0) return;
                string col = dgv.Columns[ev.ColumnIndex].Name;
                if (col == "old_value")
                {
                    ev.CellStyle.ForeColor = Color.FromArgb(180, 60, 60);
                    ev.CellStyle.Font = new Font("Segoe UI", 9);
                }
                else if (col == "new_value")
                {
                    ev.CellStyle.ForeColor = Color.FromArgb(40, 130, 40);
                    ev.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                }
            };

            this.Controls.Add(dgv);

            // ── Nút đóng ──
            var btnClose = new Button
            {
                Text = "Đóng",
                Size = new Size(100, 32),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                DialogResult = DialogResult.OK,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Location = new Point(this.ClientSize.Width - 115, this.ClientSize.Height - 42);
            this.Controls.Add(btnClose);
            this.AcceptButton = btnClose;

            // ── Resize ──
            this.Resize += (s, e) =>
            {
                pFilter.Width = this.ClientSize.Width - 20;
                dgv.Size = new Size(this.ClientSize.Width - 20, this.ClientSize.Height - 160);
                btnClose.Location = new Point(this.ClientSize.Width - 115, this.ClientSize.Height - 42);
            };

            // Enter → tìm kiếm
            txtSearch.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter) { ApplyFilter(); e.SuppressKeyPress = true; }
            };
        }

        private void LoadData()
        {
            try
            {
                string sql = @"
                    SELECT
                        t.po_trans_id                               AS [ID],
                        h.PONo                                      AS [PO No],
                        t.item_no                                   AS [Item No],
                        t.column_name_change                        AS [Cột thay đổi],
                        ISNULL(t.old_value, '')                     AS [Giá trị cũ],
                        ISNULL(t.new_value, '')                     AS [Giá trị mới],
                        CONVERT(NVARCHAR(16), t.trans_date, 120)    AS [Ngày thay đổi],
                        ISNULL(t.old_json_value, '')                AS [JSON cũ]
                    FROM PO_Revise_Transactions t
                    INNER JOIN PO_head h ON h.PO_ID = t.po_id
                    WHERE h.PONo = @poNo
                    ORDER BY t.trans_date DESC, t.po_trans_id DESC";

                using var conn = DatabaseHelper.GetConnection();
                conn.Open();
                _dtFull = new DataTable();
                _dtFull.Load(new SqlCommand(sql, conn) { Parameters = { new SqlParameter("@poNo", _poNo) } }.ExecuteReader());

                dgv.DataSource = _dtFull;

                // Ẩn cột JSON cũ mặc định (ít dùng)
                if (dgv.Columns.Contains("JSON cũ"))
                    dgv.Columns["JSON cũ"].Visible = false;

                // Width cột
                SetColumnWidths();

                lblCount.Text = $"Tổng: {_dtFull.Rows.Count} bản ghi thay đổi  —  PO: {_poNo}";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải lịch sử: " + ex.Message, "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApplyFilter()
        {
            if (_dtFull == null) return;
            string kw = txtSearch.Text.Trim().ToLower();
            if (string.IsNullOrEmpty(kw))
            {
                dgv.DataSource = _dtFull;
                lblCount.Text = $"Tổng: {_dtFull.Rows.Count} bản ghi thay đổi  —  PO: {_poNo}";
                return;
            }

            var view = _dtFull.AsEnumerable().Where(r =>
                r["Cột thay đổi"].ToString().ToLower().Contains(kw) ||
                r["Giá trị cũ"].ToString().ToLower().Contains(kw) ||
                r["Giá trị mới"].ToString().ToLower().Contains(kw) ||
                r["Ngày thay đổi"].ToString().ToLower().Contains(kw)
            );

            var dt = view.Any() ? view.CopyToDataTable() : _dtFull.Clone();
            dgv.DataSource = dt;
            if (dgv.Columns.Contains("JSON cũ")) dgv.Columns["JSON cũ"].Visible = false;
            SetColumnWidths();
            lblCount.Text = $"Hiển thị: {dt.Rows.Count} / {_dtFull.Rows.Count} bản ghi  —  PO: {_poNo}";
        }

        private void SetColumnWidths()
        {
            var widths = new System.Collections.Generic.Dictionary<string, int>
            {
                { "ID",             50  },
                { "PO No",         120  },
                { "Item No",        60  },
                { "Cột thay đổi",  160  },
                { "Giá trị cũ",    200  },
                { "Giá trị mới",   200  },
                { "Ngày thay đổi", 130  },
            };
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            foreach (DataGridViewColumn col in dgv.Columns)
            {
                if (!col.Visible) continue;
                col.Width = widths.TryGetValue(col.Name, out int w) ? w : 100;
            }
        }
    }
}