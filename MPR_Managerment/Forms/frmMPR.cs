using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Models;
using MPR_Managerment.Services;
using MPR_Managerment.Helpers;

namespace MPR_Managerment.Forms
{
    public partial class frmMPR : Form
    {
        private MPRService _service = new MPRService();
        private List<MPRHeader> _mprList = new List<MPRHeader>();
        private List<MPRDetail> _details = new List<MPRDetail>();
        private int _selectedMPR_ID = 0;
        private string _currentUser = "Admin";

        // Thêm biến để lưu ID truyền từ Dashboard sang
        private int _targetMprId = 0;

        private DataGridView dgvMPR;
        private TextBox txtSearch;
        private Button btnSearch, btnNewMPR, btnSaveHeader, btnDeleteMPR, btnClearHeader;
        private Label lblStatus;

        private TextBox txtMPRNo, txtProjectName, txtProjectCode, txtDepartment, txtRequestor, txtRev, txtNotes;
        private DateTimePicker dtpRequiredDate;
        private ComboBox cboStatus;

        // BẢNG MỚI: Danh sách file đính kèm
        private DataGridView dgvFiles;

        private DataGridView dgvDetails;
        private Button btnAddDetail, btnDeleteDetail, btnSaveDetail;

        // BẢNG: Tiến độ PO
        private DataGridView dgvPOProgress;
        private Label lblPOProgressTitle;

        private Panel panelTop, panelHeader, panelDetail;

        public frmMPR(int mprId = 0)
        {
            _targetMprId = mprId;
            InitializeComponent();
            BuildUI();
            LoadMPR();
            this.Resize += FrmMPR_Resize;
            this.WindowState = FormWindowState.Maximized;

            if (_targetMprId > 0)
            {
                SelectMPRById(_targetMprId);
            }
        }

        private void SelectMPRById(int id)
        {
            var targetMPR = _mprList.Find(m => m.MPR_ID == id);
            if (targetMPR != null)
            {
                txtSearch.Text = targetMPR.MPR_No;
                BtnSearch_Click(null, null);
            }

            foreach (DataGridViewRow row in dgvMPR.Rows)
            {
                if (Convert.ToInt32(row.Cells["ID"].Value) == id)
                {
                    dgvMPR.ClearSelection();
                    row.Selected = true;

                    if (row.Index >= 0)
                        dgvMPR.FirstDisplayedScrollingRowIndex = row.Index;
                    break;
                }
            }
        }

        private void BuildUI()
        {
            this.Text = "Quản lý Phiếu Yêu Cầu Mua Hàng (MPR)";
            this.BackColor = Color.FromArgb(245, 245, 245);

            // ===== PANEL TOP =====
            panelTop = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(1360, 220),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(panelTop);

            panelTop.Controls.Add(new Label
            {
                Text = "DANH SÁCH PHIẾU MPR",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 10),
                Size = new Size(400, 30)
            });

            txtSearch = new TextBox
            {
                Location = new Point(10, 48),
                Size = new Size(300, 28),
                Font = new Font("Segoe UI", 10),
                PlaceholderText = "Tìm theo MPR No hoặc tên dự án..."
            };
            panelTop.Controls.Add(txtSearch);
            txtSearch.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) BtnSearch_Click(null, null); };

            btnSearch = CreateButton("🔍 Tìm", Color.FromArgb(0, 120, 212), new Point(320, 47), 85, 30);
            btnSearch.Click += BtnSearch_Click;
            panelTop.Controls.Add(btnSearch);

            btnNewMPR = CreateButton("➕ Tạo MPR", Color.FromArgb(40, 167, 69), new Point(415, 47), 110, 30);
            btnNewMPR.Click += BtnNewMPR_Click;
            panelTop.Controls.Add(btnNewMPR);

            btnDeleteMPR = CreateButton("🗑 Xóa MPR", Color.FromArgb(220, 53, 69), new Point(535, 47), 110, 30);
            btnDeleteMPR.Click += BtnDeleteMPR_Click;
            panelTop.Controls.Add(btnDeleteMPR);

            lblStatus = new Label
            {
                Location = new Point(660, 52),
                Size = new Size(500, 25),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.Gray
            };
            panelTop.Controls.Add(lblStatus);

            dgvMPR = new DataGridView
            {
                Location = new Point(10, 85),
                Size = new Size(1335, 125),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            dgvMPR.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvMPR.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvMPR.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvMPR.EnableHeadersVisualStyles = false;
            dgvMPR.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvMPR.SelectionChanged += DgvMPR_SelectionChanged;
            panelTop.Controls.Add(dgvMPR);

            // ===== PANEL HEADER =====
            panelHeader = new Panel
            {
                Location = new Point(10, 240),
                Size = new Size(1360, 160), // Set chiều cao tối ưu để vừa 3 hàng control
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(panelHeader);

            panelHeader.Controls.Add(new Label
            {
                Text = "THÔNG TIN PHIẾU MPR",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 8),
                Size = new Size(300, 25)
            });

            // === BẢNG FILE ĐÍNH KÈM (CĂN BẰNG CHIỀU CAO & RỘNG THÊM 100) ===
            int gridFilesWidth = 450; // Tăng thêm 100px so với 350 cũ
            dgvFiles = new DataGridView
            {
                Location = new Point(panelHeader.Width - gridFilesWidth - 10, 10), // Đẩy y=10 để cao ngang hàng Tiêu đề
                Size = new Size(gridFilesWidth, panelHeader.Height - 20), // Chiều cao full panel (trừ margin)
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom
            };
            dgvFiles.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(108, 117, 125);
            dgvFiles.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvFiles.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvFiles.EnableHeadersVisualStyles = false;
            dgvFiles.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvFiles.Columns.Add("FileName", "Tệp đính kèm (MPR Link)");
            dgvFiles.Columns.Add("FullPath", "FullPath");
            dgvFiles.Columns["FullPath"].Visible = false;
            dgvFiles.CellDoubleClick += DgvFiles_CellDoubleClick;
            panelHeader.Controls.Add(dgvFiles);

            // === CÁC CONTROL NHẬP LIỆU BÊN TRÁI ===
            int y = 38;

            // Hàng 1
            AddLabel(panelHeader, "MPR No (*):", 10, y);
            txtMPRNo = AddTextBox(panelHeader, 90, y, 140);

            AddLabel(panelHeader, "Tên dự án:", 240, y);
            txtProjectName = AddTextBox(panelHeader, 320, y, 200);

            AddLabel(panelHeader, "Mã dự án:", 530, y);
            txtProjectCode = AddTextBox(panelHeader, 610, y, 130);

            AddLabel(panelHeader, "Trạng thái:", 750, y);
            cboStatus = new ComboBox
            {
                Location = new Point(830, y),
                Size = new Size(60, 25),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboStatus.Items.AddRange(new[] { "Mới", "Đang xử lý", "Đã duyệt", "Hoàn thành", "Hủy" });
            cboStatus.SelectedIndex = 0;
            panelHeader.Controls.Add(cboStatus);

            // Hàng 2
            y += 38;
            AddLabel(panelHeader, "Phòng ban:", 10, y);
            txtDepartment = AddTextBox(panelHeader, 90, y, 140);

            AddLabel(panelHeader, "Người YC:", 240, y);
            txtRequestor = AddTextBox(panelHeader, 320, y, 200);

            AddLabel(panelHeader, "Ngày cần:", 530, y);
            dtpRequiredDate = new DateTimePicker
            {
                Location = new Point(610, y),
                Size = new Size(130, 25),
                Font = new Font("Segoe UI", 9),
                Format = DateTimePickerFormat.Short
            };
            panelHeader.Controls.Add(dtpRequiredDate);

            AddLabel(panelHeader, "Rev:", 750, y);
            txtRev = AddTextBox(panelHeader, 790, y, 100);
            txtRev.Text = "0";

            // Hàng 3 (Buttons & Notes)
            y += 38;
            btnSaveHeader = CreateButton("💾 Lưu Header", Color.FromArgb(0, 120, 212), new Point(10, y), 130, 32);
            btnSaveHeader.Click += BtnSaveHeader_Click;
            panelHeader.Controls.Add(btnSaveHeader);

            btnClearHeader = CreateButton("🔄 Làm mới", Color.FromArgb(108, 117, 125), new Point(150, y), 110, 32);
            btnClearHeader.Click += BtnClearHeader_Click;
            panelHeader.Controls.Add(btnClearHeader);

            AddLabel(panelHeader, "Ghi chú:", 270, y + 5);
            // txtNotes tự động giãn ra lấp kín phần trống ở giữa
            txtNotes = AddTextBox(panelHeader, 340, y + 2, dgvFiles.Left - 340 - 15);
            txtNotes.Anchor = AnchorStyles.Top | AnchorStyles.Left;

            // ===== PANEL DETAIL =====
            panelDetail = new Panel
            {
                Location = new Point(10, panelHeader.Bottom + 10),
                Size = new Size(1360, 345),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            this.Controls.Add(panelDetail);

            panelDetail.Controls.Add(new Label
            {
                Text = "CHI TIẾT VẬT TƯ",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 8),
                Size = new Size(300, 25)
            });

            btnAddDetail = CreateButton("➕ Thêm dòng", Color.FromArgb(40, 167, 69), new Point(10, 38), 120, 30);
            btnAddDetail.Click += BtnAddDetail_Click;
            panelDetail.Controls.Add(btnAddDetail);

            btnDeleteDetail = CreateButton("🗑 Xóa dòng", Color.FromArgb(220, 53, 69), new Point(140, 38), 110, 30);
            btnDeleteDetail.Click += BtnDeleteDetail_Click;
            panelDetail.Controls.Add(btnDeleteDetail);

            btnSaveDetail = CreateButton("💾 Lưu chi tiết", Color.FromArgb(0, 120, 212), new Point(260, 38), 130, 30);
            btnSaveDetail.Click += BtnSaveDetail_Click;
            panelDetail.Controls.Add(btnSaveDetail);

            var btnCreatePO = CreateButton("🛒 Tạo PO", Color.FromArgb(255, 140, 0), new Point(400, 38), 120, 30);
            btnCreatePO.Click += BtnCreatePO_Click;
            panelDetail.Controls.Add(btnCreatePO);

            dgvDetails = new DataGridView
            {
                Location = new Point(10, 75),
                Size = new Size(900, 260),
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom
            };
            dgvDetails.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvDetails.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvDetails.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvDetails.EnableHeadersVisualStyles = false;
            dgvDetails.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvDetails.CellFormatting += DgvDetails_CellFormatting;

            BuildDetailColumns();
            panelDetail.Controls.Add(dgvDetails);

            lblPOProgressTitle = new Label
            {
                Text = "TỔNG HỢP TIẾN ĐỘ PO ĐÃ ĐẶT",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(255, 140, 0),
                Location = new Point(930, 48),
                Size = new Size(300, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            panelDetail.Controls.Add(lblPOProgressTitle);

            dgvPOProgress = new DataGridView
            {
                Location = new Point(930, 75),
                Size = new Size(415, 260),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom
            };
            dgvPOProgress.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 140, 0);
            dgvPOProgress.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPOProgress.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvPOProgress.EnableHeadersVisualStyles = false;
            dgvPOProgress.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(255, 248, 235);
            dgvPOProgress.CellFormatting += DgvPOProgress_CellFormatting;
            dgvPOProgress.CellDoubleClick += DgvPOProgress_CellDoubleClick;
            panelDetail.Controls.Add(dgvPOProgress);

            Common.Common.AutoBringToFontControl(new[] { panelTop, panelHeader, panelDetail });
        }

        // =====================================================================
        // SỰ KIỆN DOUBLE CLICK VÀO FILE TRONG BẢNG ĐÍNH KÈM
        // =====================================================================
        private void DgvFiles_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string path = dgvFiles.Rows[e.RowIndex].Cells["FullPath"].Value?.ToString();

            if (!string.IsNullOrEmpty(path) && File.Exists(path))
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = path,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Không thể mở file: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else if (!string.IsNullOrEmpty(path))
            {
                MessageBox.Show("File không tồn tại hoặc đã bị xóa / di chuyển khỏi thư mục!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // =====================================================================
        // HÀM QUÉT THƯ MỤC LẤY DANH SÁCH FILE CỦA DỰ ÁN ĐANG CHỌN
        // =====================================================================
        private void LoadFiles(string projectName)
        {
            dgvFiles.Rows.Clear();
            if (string.IsNullOrEmpty(projectName)) return;

            try
            {
                var projects = new ProjectService().GetAll();
                var prj = projects.Find(p =>
                    !string.IsNullOrEmpty(p.ProjectName) &&
                    p.ProjectName.Equals(projectName, StringComparison.OrdinalIgnoreCase));

                if (prj == null)
                {
                    prj = projects.Find(p =>
                        !string.IsNullOrEmpty(p.ProjectName) &&
                        p.ProjectName.IndexOf(projectName, StringComparison.OrdinalIgnoreCase) >= 0);
                }

                if (prj != null && !string.IsNullOrEmpty(prj.MPR_Link) && Directory.Exists(prj.MPR_Link))
                {
                    var files = Directory.GetFiles(prj.MPR_Link);
                    foreach (var f in files)
                    {
                        dgvFiles.Rows.Add(Path.GetFileName(f), f);
                    }
                }
                else if (prj != null && !string.IsNullOrEmpty(prj.MPR_Link))
                {
                    dgvFiles.Rows.Add("(Thư mục không tồn tại)", "");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Lỗi load files: " + ex.Message);
            }
        }

        private void DgvPOProgress_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string poNo = dgvPOProgress.Rows[e.RowIndex].Cells["PO No"].Value?.ToString() ?? "";

            if (!string.IsNullOrEmpty(poNo))
            {
                var frm = new frmPO(poNo);
                frm.Show();
            }
        }

        private void BuildDetailColumns()
        {
            dgvDetails.Columns.Clear();
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Detail_ID", HeaderText = "ID", Visible = false });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Item_No", HeaderText = "STT", Width = 45, ReadOnly = true });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Item_Name", HeaderText = "Tên vật tư", Width = 180 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Description", HeaderText = "Mô tả", Width = 100 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Material", HeaderText = "Vật liệu", Width = 85 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Thickness_mm", HeaderText = "A-Dày(mm)", Width = 75 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Depth_mm", HeaderText = "B-Sâu(mm)", Width = 75 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "C_Width_mm", HeaderText = "C-Rộng(mm)", Width = 80 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "D_Web_mm", HeaderText = "D-Bụng(mm)", Width = 80 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "E_Flange_mm", HeaderText = "E-Cánh(mm)", Width = 80 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "F_Length_mm", HeaderText = "F-Dài(mm)", Width = 75 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "UNIT", HeaderText = "ĐVT", Width = 50 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Qty", HeaderText = "SL", Width = 50 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Weight", HeaderText = "KG", Width = 55 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "MPS_Info", HeaderText = "MPS Info", Width = 100 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Usage_Location", HeaderText = "Vị trí dùng", Width = 110 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "REV", HeaderText = "REV", Width = 45 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Remarks", HeaderText = "Ghi chú", FillWeight = 100 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "PO_No", HeaderText = "Đã lên PO", Width = 120, ReadOnly = true });
        }

        private void DgvDetails_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (dgvDetails.Columns[e.ColumnIndex].Name == "PO_No")
            {
                string val = e.Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(val))
                {
                    e.CellStyle.ForeColor = Color.FromArgb(40, 167, 69);
                    e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                }
            }
        }

        private void DgvPOProgress_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string colName = dgvPOProgress.Columns[e.ColumnIndex].Name;

            if (colName == "% Giao")
            {
                if (decimal.TryParse(e.Value?.ToString(), out decimal pct))
                {
                    e.CellStyle.ForeColor = pct >= 100 ? Color.FromArgb(40, 167, 69) : pct >= 50 ? Color.FromArgb(255, 140, 0) : Color.FromArgb(220, 53, 69);
                    e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    e.Value = $"{pct}%";
                    e.FormattingApplied = true;
                }
            }
            if (colName == "Ngày PO")
            {
                if (e.Value != null && e.Value != DBNull.Value)
                {
                    e.Value = Convert.ToDateTime(e.Value).ToString("dd/MM/yyyy");
                    e.FormattingApplied = true;
                }
            }
        }

        private void AddLabel(Panel panel, string text, int x, int y)
        {
            panel.Controls.Add(new Label
            {
                Text = text,
                Location = new Point(x, y + 3),
                Size = new Size(80, 20),
                Font = new Font("Segoe UI", 9),
                Margin = new Padding(0)
            });
        }

        private TextBox AddTextBox(Panel panel, int x, int y, int width)
        {
            var txt = new TextBox
            {
                Location = new Point(x, y),
                Size = new Size(width, 25),
                Font = new Font("Segoe UI", 9)
            };
            panel.Controls.Add(txt);
            return txt;
        }

        private Button CreateButton(string text, Color color, Point location, int w, int h)
        {
            var btn = new Button
            {
                Text = text,
                Location = location,
                Size = new Size(w, h),
                BackColor = color,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        // ===== RESIZE =====
        private void FrmMPR_Resize(object sender, EventArgs e)
        {
            try
            {
                int w = this.ClientSize.Width - 20;
                int h = this.ClientSize.Height;

                panelTop.Width = w;
                panelHeader.Width = w;
                panelDetail.Width = w;
                panelDetail.Height = h - panelDetail.Top - 10;
                dgvMPR.Width = panelTop.Width - 20;

                int rightWidth = 420;
                dgvPOProgress.Width = rightWidth;
                dgvPOProgress.Left = panelDetail.Width - rightWidth - 10;
                dgvPOProgress.Height = panelDetail.Height - 85;

                lblPOProgressTitle.Left = dgvPOProgress.Left;

                dgvDetails.Width = dgvPOProgress.Left - 20;
                dgvDetails.Height = panelDetail.Height - 85;

                // Tự động kéo dài ô Ghi chú để lấp đầy khoảng trống tới Grid Files
                if (txtNotes != null && panelHeader != null && dgvFiles != null)
                {
                    int noteW = dgvFiles.Left - txtNotes.Left - 15;
                    if (noteW > 50) txtNotes.Width = noteW;
                }
            }
            catch { }
        }

        // ===== LOAD MPR =====
        private void LoadMPR()
        {
            try
            {
                _mprList = _service.GetAll();
                BindMPRGrid(_mprList);
                lblStatus.Text = $"Tổng: {_mprList.Count} phiếu MPR";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải MPR: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BindMPRGrid(List<MPRHeader> list)
        {
            dgvMPR.DataSource = list.ConvertAll(m => new
            {
                ID = m.MPR_ID,
                MPR_No = m.MPR_No,
                Ten_Du_An = m.Project_Name,
                Ma_Du_An = m.Project_Code,
                Phong_Ban = m.Department,
                Nguoi_YC = m.Requestor,
                Ngay_Can = m.Required_Date.HasValue ? m.Required_Date.Value.ToString("dd/MM/yyyy") : "",
                Rev = m.Rev,
                Trang_Thai = m.Status,
                Ngay_Tao = m.Created_Date.HasValue ? m.Created_Date.Value.ToString("dd/MM/yyyy") : ""
            });
            if (dgvMPR.Columns.Contains("ID"))
                dgvMPR.Columns["ID"].Visible = false;
        }

        // ===== LOAD TỔNG HỢP TIẾN ĐỘ PO =====
        private void LoadPOProgress(string mprNo)
        {
            if (string.IsNullOrEmpty(mprNo))
            {
                dgvPOProgress.DataSource = null;
                return;
            }

            try
            {
                string sql = @"
                    SELECT
                        h.PONo AS [PO No],
                        h.PO_Date AS [Ngày PO],
                        h.Status AS [Trạng thái],
                        CASE
                            WHEN ISNULL(SUM(d.Qty_Per_Sheet), 0) = 0 THEN 0
                            ELSE CAST(
                                ISNULL(
                                    (SELECT SUM(Qty_Import) FROM Warehouse_Import wi WHERE wi.PO_ID = h.PO_ID), 
                                    ISNULL(SUM(d.Received), 0)
                                ) * 100.0 / SUM(d.Qty_Per_Sheet) 
                            AS DECIMAL(5,1))
                        END AS [% Giao]
                    FROM PO_head h
                    LEFT JOIN PO_Detail d ON h.PO_ID = d.PO_ID
                    WHERE h.MPR_No = @mprNo
                    GROUP BY h.PO_ID, h.PONo, h.PO_Date, h.Status
                    ORDER BY h.PO_Date DESC";
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var cmd = new SqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@mprNo", mprNo);
                    var dt = new DataTable();
                    dt.Load(cmd.ExecuteReader());
                    dgvPOProgress.DataSource = dt;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Lỗi tải PO Progress: " + ex.Message);
            }
        }

        private Dictionary<int, string> GetPoMappingForMpr(int mprId)
        {
            var dict = new Dictionary<int, string>();
            if (mprId <= 0) return dict;

            try
            {
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    string sql = @"
                        SELECT pod.MPR_Detail_ID, poh.PONo
                        FROM PO_Detail pod
                        INNER JOIN PO_head poh ON pod.PO_ID = poh.PO_ID
                        WHERE pod.MPR_Detail_ID IS NOT NULL
                          AND pod.MPR_Detail_ID IN (
                              SELECT Detail_ID FROM MPR_Details WHERE MPR_ID = @mprId
                          )";

                    using (var cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@mprId", mprId);
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                if (reader["MPR_Detail_ID"] != DBNull.Value)
                                {
                                    int detailId = Convert.ToInt32(reader["MPR_Detail_ID"]);
                                    string poNo = reader["PONo"]?.ToString() ?? "";

                                    if (dict.ContainsKey(detailId))
                                    {
                                        if (!dict[detailId].Contains(poNo))
                                            dict[detailId] += ", " + poNo;
                                    }
                                    else
                                    {
                                        dict[detailId] = poNo;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Lỗi lấy PO Mapping: " + ex.Message);
            }
            return dict;
        }

        private void LoadDetails(int mprId)
        {
            try
            {
                _details = _service.GetDetails(mprId);
                dgvDetails.Rows.Clear();

                var poMapping = GetPoMappingForMpr(mprId);

                foreach (var d in _details)
                {
                    int idx = dgvDetails.Rows.Add();
                    var row = dgvDetails.Rows[idx];

                    row.Cells["Detail_ID"].Value = d.Detail_ID;
                    row.Cells["Item_No"].Value = d.Item_No;
                    row.Cells["Item_Name"].Value = d.Item_Name;
                    row.Cells["Description"].Value = d.Description;
                    row.Cells["Material"].Value = d.Material;
                    row.Cells["Thickness_mm"].Value = d.Thickness_mm;
                    row.Cells["Depth_mm"].Value = d.Depth_mm;
                    row.Cells["C_Width_mm"].Value = d.C_Width_mm;
                    row.Cells["D_Web_mm"].Value = d.D_Web_mm;
                    row.Cells["E_Flange_mm"].Value = d.E_Flange_mm;
                    row.Cells["F_Length_mm"].Value = d.F_Length_mm;
                    row.Cells["UNIT"].Value = d.UNIT;
                    row.Cells["Qty"].Value = d.Qty_Per_Sheet;
                    row.Cells["Weight"].Value = d.Weight_kg;
                    row.Cells["MPS_Info"].Value = d.MPS_Info;
                    row.Cells["Usage_Location"].Value = d.Usage_Location;
                    row.Cells["REV"].Value = d.REV;
                    row.Cells["Remarks"].Value = d.Remarks;
                    row.Cells["PO_No"].Value = poMapping.ContainsKey(d.Detail_ID) ? poMapping[d.Detail_ID] : "";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải chi tiết: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ===== SỰ KIỆN =====
        private void DgvMPR_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvMPR.SelectedRows.Count == 0) return;
            var row = dgvMPR.SelectedRows[0];
            _selectedMPR_ID = Convert.ToInt32(row.Cells["ID"].Value);

            var m = _mprList.Find(x => x.MPR_ID == _selectedMPR_ID);
            if (m == null) return;
            txtMPRNo.Text = m.MPR_No;
            txtProjectName.Text = m.Project_Name;
            txtProjectCode.Text = m.Project_Code;
            txtDepartment.Text = m.Department;
            txtRequestor.Text = m.Requestor;
            txtRev.Text = m.Rev.ToString();
            txtNotes.Text = m.Notes;
            dtpRequiredDate.Value = m.Required_Date ?? DateTime.Today;

            int idx = cboStatus.Items.IndexOf(m.Status);
            cboStatus.SelectedIndex = idx >= 0 ? idx : 0;

            LoadDetails(_selectedMPR_ID);
            LoadPOProgress(m.MPR_No);
            LoadFiles(m.Project_Name); // Tự động lấy file thư mục khi chọn MPR
        }

        private void BtnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                string kw = txtSearch.Text.Trim();
                _mprList = string.IsNullOrEmpty(kw)
                    ? _service.GetAll()
                    : _service.GetAll().FindAll(m =>
                        (m.MPR_No ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                        (m.Project_Name ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                        (m.Project_Code ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase));

                BindMPRGrid(_mprList);
                lblStatus.Text = $"Tìm thấy: {_mprList.Count} phiếu";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tìm kiếm: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnNewMPR_Click(object sender, EventArgs e)
        {
            _selectedMPR_ID = 0;
            ClearHeader();
            dgvDetails.Rows.Clear();
            dgvPOProgress.DataSource = null;
            dgvFiles.Rows.Clear();
            _details.Clear();
            txtMPRNo.Focus();
        }

        private void BtnSaveHeader_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtMPRNo.Text))
            {
                MessageBox.Show("Vui lòng nhập MPR No!", "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtMPRNo.Focus();
                return;
            }

            try
            {
                var m = new MPRHeader
                {
                    MPR_ID = _selectedMPR_ID,
                    MPR_No = txtMPRNo.Text.Trim(),
                    Project_Name = txtProjectName.Text.Trim(),
                    Project_Code = txtProjectCode.Text.Trim(),
                    Department = txtDepartment.Text.Trim(),
                    Requestor = txtRequestor.Text.Trim(),
                    Required_Date = dtpRequiredDate.Value,
                    Rev = int.TryParse(txtRev.Text, out int rev) ? rev : 0,
                    Status = cboStatus.SelectedItem?.ToString() ?? "Mới",
                    Notes = txtNotes.Text.Trim()
                };
                if (_selectedMPR_ID == 0)
                {
                    _selectedMPR_ID = _service.InsertHeader(m, _currentUser);
                    MessageBox.Show("Tạo phiếu MPR thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    _service.UpdateHeader(m, _currentUser);
                    MessageBox.Show("Cập nhật phiếu MPR thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                LoadMPR();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi lưu header: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnDeleteMPR_Click(object sender, EventArgs e)
        {
            if (_selectedMPR_ID == 0)
            {
                MessageBox.Show("Vui lòng chọn phiếu MPR cần xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show("Xóa phiếu MPR này và toàn bộ chi tiết?", "Xác nhận",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    _service.DeleteMPR(_selectedMPR_ID);
                    MessageBox.Show("Xóa thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    _selectedMPR_ID = 0;
                    ClearHeader();
                    dgvDetails.Rows.Clear();
                    dgvPOProgress.DataSource = null;
                    dgvFiles.Rows.Clear();
                    LoadMPR();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnClearHeader_Click(object sender, EventArgs e)
        {
            _selectedMPR_ID = 0;
            ClearHeader();
            dgvDetails.Rows.Clear();
            dgvPOProgress.DataSource = null;
            dgvFiles.Rows.Clear();
            _details.Clear();
        }

        private void BtnCreatePO_Click(object sender, EventArgs e)
        {
            if (_selectedMPR_ID == 0)
            {
                MessageBox.Show("Vui lòng chọn một MPR trước!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var mpr = _mprList.Find(m => m.MPR_ID == _selectedMPR_ID);
            if (mpr == null) return;

            string mprNo = mpr.MPR_No;
            var frm = new frmPO(mprNo, true);
            frm.Show();
        }

        private void BtnAddDetail_Click(object sender, EventArgs e)
        {
            if (_selectedMPR_ID == 0)
            {
                MessageBox.Show("Vui lòng chọn hoặc lưu phiếu MPR trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int nextNo = dgvDetails.Rows.Count + 1;
            int newIdx = dgvDetails.Rows.Add();
            var newRow = dgvDetails.Rows[newIdx];

            newRow.Cells["Detail_ID"].Value = 0;
            newRow.Cells["Item_No"].Value = nextNo;
            newRow.Cells["Item_Name"].Value = "";
            newRow.Cells["Description"].Value = "";
            newRow.Cells["Material"].Value = "";
            newRow.Cells["Thickness_mm"].Value = 0;
            newRow.Cells["Depth_mm"].Value = 0;
            newRow.Cells["C_Width_mm"].Value = 0;
            newRow.Cells["D_Web_mm"].Value = 0;
            newRow.Cells["E_Flange_mm"].Value = 0;
            newRow.Cells["F_Length_mm"].Value = 0;
            newRow.Cells["UNIT"].Value = "cái";
            newRow.Cells["Qty"].Value = 0;
            newRow.Cells["Weight"].Value = 0;
            newRow.Cells["MPS_Info"].Value = "";
            newRow.Cells["Usage_Location"].Value = "";
            newRow.Cells["REV"].Value = "0";
            newRow.Cells["Remarks"].Value = "";
            newRow.Cells["PO_No"].Value = "";

            dgvDetails.CurrentCell = dgvDetails.Rows[newIdx].Cells["Item_Name"];
            dgvDetails.BeginEdit(true);
        }

        private void BtnDeleteDetail_Click(object sender, EventArgs e)
        {
            if (dgvDetails.SelectedRows.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn dòng cần xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string msg = dgvDetails.SelectedRows.Count == 1
                ? "Bạn có chắc chắn muốn xóa dòng này?"
                : $"Bạn có chắc chắn muốn xóa {dgvDetails.SelectedRows.Count} dòng đã chọn?";
            if (MessageBox.Show(msg, "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    var rowsToDelete = new List<DataGridViewRow>();
                    foreach (DataGridViewRow row in dgvDetails.SelectedRows)
                    {
                        if (!row.IsNewRow)
                        {
                            rowsToDelete.Add(row);
                        }
                    }

                    foreach (var row in rowsToDelete)
                    {
                        dgvDetails.Rows.Remove(row);
                    }

                    int itemNo = 1;
                    foreach (DataGridViewRow row in dgvDetails.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            row.Cells["Item_No"].Value = itemNo++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi xóa: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnSaveDetail_Click(object sender, EventArgs e)
        {
            if (_selectedMPR_ID == 0)
            {
                MessageBox.Show("Vui lòng lưu header MPR trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (dgvDetails.Rows.Count == 0)
            {
                MessageBox.Show("Không có dòng nào để lưu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                int saved = 0;
                foreach (DataGridViewRow row in dgvDetails.Rows)
                {
                    string itemName = row.Cells["Item_Name"].Value?.ToString() ?? "";
                    if (string.IsNullOrWhiteSpace(itemName)) continue;

                    var d = new MPRDetail
                    {
                        Detail_ID = Convert.ToInt32(row.Cells["Detail_ID"].Value ?? 0),
                        MPR_ID = _selectedMPR_ID,
                        Item_No = Convert.ToInt32(row.Cells["Item_No"].Value ?? 0),
                        Item_Name = itemName,
                        Description = row.Cells["Description"].Value?.ToString() ?? "",
                        Material = row.Cells["Material"].Value?.ToString() ?? "",
                        Thickness_mm = DecimalVal(row.Cells["Thickness_mm"].Value),
                        Depth_mm = DecimalVal(row.Cells["Depth_mm"].Value),
                        C_Width_mm = DecimalVal(row.Cells["C_Width_mm"].Value),
                        D_Web_mm = DecimalVal(row.Cells["D_Web_mm"].Value),
                        E_Flange_mm = DecimalVal(row.Cells["E_Flange_mm"].Value),
                        F_Length_mm = DecimalVal(row.Cells["F_Length_mm"].Value),
                        UNIT = row.Cells["UNIT"].Value?.ToString() ?? "",
                        Qty_Per_Sheet = (int)DecimalVal(row.Cells["Qty"].Value),
                        Weight_kg = DecimalVal(row.Cells["Weight"].Value),
                        MPS_Info = row.Cells["MPS_Info"].Value?.ToString() ?? "",
                        Usage_Location = row.Cells["Usage_Location"].Value?.ToString() ?? "",
                        REV = row.Cells["REV"].Value?.ToString() ?? "0",
                        Remarks = row.Cells["Remarks"].Value?.ToString() ?? ""
                    };
                    if (d.Detail_ID == 0)
                        _service.InsertDetail(d, _currentUser);
                    else
                        _service.UpdateDetail(d, _currentUser);
                    saved++;
                }

                MessageBox.Show($"Đã lưu {saved} dòng chi tiết thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDetails(_selectedMPR_ID);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi lưu chi tiết: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ===== HELPERS =====
        private decimal DecimalVal(object val)
        {
            if (val == null || val == DBNull.Value) return 0;
            return decimal.TryParse(val.ToString(), out decimal d) ? d : 0;
        }

        private void ClearHeader()
        {
            txtMPRNo.Text = "";
            txtProjectName.Text = "";
            txtProjectCode.Text = "";
            txtDepartment.Text = "";
            txtRequestor.Text = "";
            txtRev.Text = "0";
            txtNotes.Text = "";
            dtpRequiredDate.Value = DateTime.Today;
            cboStatus.SelectedIndex = 0;
        }
    }
}