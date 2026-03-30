using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using MPR_Managerment.Models;
using MPR_Managerment.Services;

namespace MPR_Managerment.Forms
{
    public partial class frmSelectMPR : Form
    {
        private MPRService _service = new MPRService();
        private List<MPRHeader> _mprList = new List<MPRHeader>();

        public MPRHeader SelectedMPR { get; private set; }
        public List<MPRDetail> SelectedDetails { get; private set; }

        private DataGridView dgvMPR;
        private TextBox txtSearch;

        //private Panel panelHead, panelTop, panelDetail;

        public frmSelectMPR()
        {
            InitializeComponent();
            BuildUI();
            LoadMPR();
        }

        private void BuildUI()
        {
            this.Text = "Chọn phiếu MPR";
            this.Size = new Size(900, 520);
            this.StartPosition = FormStartPosition.CenterParent;
            this.BackColor = Color.White;

            this.Controls.Add(new Label
            {
                Text = "CHỌN PHIẾU MPR ĐỂ IMPORT VÀO PO",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 10),
                Size = new Size(500, 30)
            });

            txtSearch = new TextBox
            {
                Location = new Point(10, 50),
                Size = new Size(350, 28),
                Font = new Font("Segoe UI", 10),
                PlaceholderText = "Tìm theo MPR No hoặc tên dự án..."
            };
            this.Controls.Add(txtSearch);
            txtSearch.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) LoadMPR(txtSearch.Text); };

            var btnSearch = new Button
            {
                Text = "Tìm",
                Location = new Point(370, 49),
                Size = new Size(70, 30),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            btnSearch.Click += (s, e) => LoadMPR(txtSearch.Text);
            this.Controls.Add(btnSearch);

            dgvMPR = new DataGridView
            {
                Location = new Point(10, 90),
                Size = new Size(860, 360),
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
            dgvMPR.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvMPR.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvMPR.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvMPR.EnableHeadersVisualStyles = false;
            dgvMPR.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvMPR.DoubleClick += DgvMPR_DoubleClick;
            this.Controls.Add(dgvMPR);

            var btnSelect = new Button
            {
                Text = "Chọn MPR này",
                Location = new Point(10, 458),
                Size = new Size(140, 32),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            btnSelect.Click += BtnSelect_Click;
            this.Controls.Add(btnSelect);

            var btnCancel = new Button
            {
                Text = "Hủy",
                Location = new Point(160, 458),
                Size = new Size(80, 32),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };
            //this.Controls.Add(btnCancel);
            //// Đưa tất cả TextBox và ComboBox lên trên Label
            //foreach (Panel panel in new[] { panelHead, panelTop, panelDetail })
            //{
            //    foreach (Control c in panel.Controls)
            //    {
            //        if (c is TextBox || c is ComboBox || c is DateTimePicker || c is NumericUpDown)
            //            c.BringToFront();
            //    }
            //}
        }

        private void LoadMPR(string keyword = "")
        {
            try
            {
                _mprList = string.IsNullOrWhiteSpace(keyword)
                    ? _service.GetAll()
                    : _service.Search(keyword);

                dgvMPR.DataSource = _mprList.ConvertAll(h => new
                {
                    ID = h.MPR_ID,
                    MPR_No = h.MPR_No,
                    Du_An = h.Project_Name,
                    Ma_DA = h.Project_Code,
                    Nguoi_YC = h.Requestor,
                    Ngay_Can = h.Required_Date.HasValue ? h.Required_Date.Value.ToString("dd/MM/yyyy") : "",
                    Trang_Thai = h.Status,
                    Rev = h.Rev
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SelectMPR()
        {
            if (dgvMPR.SelectedRows.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn một phiếu MPR!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            int mprId = Convert.ToInt32(dgvMPR.SelectedRows[0].Cells["ID"].Value);
            SelectedMPR = _mprList.Find(x => x.MPR_ID == mprId);
            SelectedDetails = _service.GetDetails(mprId);
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void BtnSelect_Click(object sender, EventArgs e) => SelectMPR();
        private void DgvMPR_DoubleClick(object sender, EventArgs e) => SelectMPR();
    }
}