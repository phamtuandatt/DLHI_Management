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
            ApplyPermissions();
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
            dgvMPR.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
            dgvMPR.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgvMPR.SelectionChanged += DgvMPR_SelectionChanged;
            panelTop.Controls.Add(dgvMPR);

            // ===== PANEL HEADER =====
            panelHeader = new Panel
            {
                Location = new Point(10, 240),
                Size = new Size(1360, 160),
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

            // === BẢNG FILE ĐÍNH KÈM ===
            int gridFilesWidth = 450;
            int filesLeft = panelHeader.Width - gridFilesWidth - 10;
            dgvFiles = new DataGridView
            {
                Location = new Point(filesLeft, 10),
                Size = new Size(gridFilesWidth, panelHeader.Height - 20),
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
            dgvFiles.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
            dgvFiles.DefaultCellStyle.SelectionForeColor = Color.Black;
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
            txtNotes = AddTextBox(panelHeader, 340, y + 2, filesLeft - 340 - 15);
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

            var btnCheckAll = CreateButton("🔎 Check All Items", Color.FromArgb(102, 51, 153), new Point(530, 38), 150, 30);
            btnCheckAll.Click += BtnCheckAllItems_Click;
            panelDetail.Controls.Add(btnCheckAll);

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
            dgvDetails.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvDetails.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvDetails.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvDetails.EnableHeadersVisualStyles = false;
            dgvDetails.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvDetails.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
            dgvDetails.DefaultCellStyle.SelectionForeColor = Color.Black;
            dgvDetails.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
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
            dgvPOProgress.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
            dgvPOProgress.DefaultCellStyle.SelectionForeColor = Color.Black;
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
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Item_No",
                HeaderText = "STT",
                Width = 45,
                ReadOnly = true,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Item_Name",
                HeaderText = "Tên vật tư",
                Width = 180,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleLeft }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Description",
                HeaderText = "Mô tả",
                Width = 100,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleLeft }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Material",
                HeaderText = "Vật liệu",
                Width = 85,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Thickness_mm",
                HeaderText = "A-Dày(mm)",
                Width = 45,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Depth_mm",
                HeaderText = "B-Sâu(mm)",
                Width = 45,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "C_Width_mm",
                HeaderText = "C-Rộng(mm)",
                Width = 45,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "D_Web_mm",
                HeaderText = "D-Bụng(mm)",
                Width = 45,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "E_Flange_mm",
                HeaderText = "E-Cánh(mm)",
                Width = 45,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "F_Length_mm",
                HeaderText = "F-Dài(mm)",
                Width = 75,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "UNIT",
                HeaderText = "ĐVT",
                Width = 50,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Qty",
                HeaderText = "SL",
                Width = 50,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Weight",
                HeaderText = "KG",
                Width = 55,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "MPS_Info",
                HeaderText = "MPS Info",
                Width = 100,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Usage_Location",
                HeaderText = "Vị trí dùng",
                Width = 110,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleLeft }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "REV",
                HeaderText = "REV",
                Width = 45,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Remarks",
                HeaderText = "Ghi chú",
                FillWeight = 100,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleLeft }
            });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "PO_No",
                HeaderText = "Đã lên PO",
                Width = 120,
                ReadOnly = true,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });
        }

        private void DgvDetails_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            string colName = dgvDetails.Columns[e.ColumnIndex].Name;

            if (colName == "Thickness_mm" || colName == "Depth_mm" ||
                colName == "C_Width_mm" || colName == "D_Web_mm" ||
                colName == "E_Flange_mm" || colName == "F_Length_mm" ||
                colName == "Qty" || colName == "Weight")
            {
                if (e.Value != null && decimal.TryParse(e.Value.ToString(), out decimal num) && num == 0)
                {
                    e.Value = "";
                    e.FormattingApplied = true;
                }
                e.CellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            // Cột PO_No — màu xanh bold
            if (colName == "PO_No")
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

        // =====================================================================
        // LẤY PO MAPPING CHO MPR — HỖ TRỢ CẢ MPR REVISE
        // Logic:
        //   Bước 1: Tìm PO liên kết trực tiếp qua MPR_Detail_ID (Detail_ID hiện tại)
        //   Bước 2: Với các dòng chưa có PO, tìm sang các phiên bản MPR khác cùng MPR_No
        //           (revise) khớp theo Item_No + Item_Name + Material để lấy PO đã đặt
        //           từ phiên bản cũ — rồi điền vào dòng tương ứng của phiên bản mới
        // =====================================================================
        private Dictionary<int, string> GetPoMappingForMpr(int mprId)
        {
            var dict = new Dictionary<int, string>();
            if (mprId <= 0) return dict;

            try
            {
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();

                    // ── Bước 1: PO liên kết trực tiếp qua Detail_ID ──
                    string sqlDirect = @"
                        SELECT pod.MPR_Detail_ID, poh.PONo
                        FROM PO_Detail pod
                        INNER JOIN PO_head poh ON pod.PO_ID = poh.PO_ID
                        WHERE pod.MPR_Detail_ID IS NOT NULL
                          AND pod.MPR_Detail_ID IN (
                              SELECT Detail_ID FROM MPR_Details WHERE MPR_ID = @mprId
                          )";

                    using (var cmd = new SqlCommand(sqlDirect, conn))
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

                    // ── Bước 2: Tìm PO từ các phiên bản revise khác cùng MPR_No ──
                    // Lấy thông tin các dòng của MPR hiện tại chưa có PO
                    string sqlCurrentDetails = @"
                        SELECT d.Detail_ID, d.Item_No, d.item_name, d.Material,
                               d.Thickness_mm, d.Depth_mm, d.C_Width_mm,
                               d.D_Web_mm, d.E_Flange_mm, d.F_Length_mm,
                               h.MPR_No
                        FROM MPR_Details d
                        INNER JOIN MPR_Header h ON h.MPR_ID = d.MPR_ID
                        WHERE d.MPR_ID = @mprId";

                    var currentDetails = new List<(int DetailId, int ItemNo, string ItemName, string Material,
                        decimal A, decimal B, decimal C, decimal D, decimal E, decimal F, string MprNo)>();

                    using (var cmd = new SqlCommand(sqlCurrentDetails, conn))
                    {
                        cmd.Parameters.AddWithValue("@mprId", mprId);
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                int detId = Convert.ToInt32(reader["Detail_ID"]);
                                // Chỉ xử lý các dòng chưa có PO từ bước 1
                                if (!dict.ContainsKey(detId))
                                {
                                    currentDetails.Add((
                                        detId,
                                        Convert.ToInt32(reader["Item_No"]),
                                        reader["item_name"]?.ToString() ?? "",
                                        reader["Material"]?.ToString() ?? "",
                                        reader["Thickness_mm"] != DBNull.Value ? Convert.ToDecimal(reader["Thickness_mm"]) : 0,
                                        reader["Depth_mm"] != DBNull.Value ? Convert.ToDecimal(reader["Depth_mm"]) : 0,
                                        reader["C_Width_mm"] != DBNull.Value ? Convert.ToDecimal(reader["C_Width_mm"]) : 0,
                                        reader["D_Web_mm"] != DBNull.Value ? Convert.ToDecimal(reader["D_Web_mm"]) : 0,
                                        reader["E_Flange_mm"] != DBNull.Value ? Convert.ToDecimal(reader["E_Flange_mm"]) : 0,
                                        reader["F_Length_mm"] != DBNull.Value ? Convert.ToDecimal(reader["F_Length_mm"]) : 0,
                                        reader["MPR_No"]?.ToString() ?? ""
                                    ));
                                }
                            }
                        }
                    }

                    if (currentDetails.Count > 0)
                    {
                        // Lấy MPR_No gốc (không có hậu tố Rev) để tìm các phiên bản liên quan
                        // Quy ước: MPR-001, MPR-001-R1, MPR-001-R2 → cùng gốc MPR-001
                        string mprNoBase = currentDetails[0].MprNo;

                        // Tìm PO từ tất cả các phiên bản MPR khác có MPR_No bắt đầu giống gốc
                        // Khớp theo Item_No + Item_Name + Material (+ kích thước nếu có)
                        string sqlRevise = @"
                            SELECT
                                d.Item_No,
                                d.item_name,
                                d.Material,
                                d.Thickness_mm, d.Depth_mm, d.C_Width_mm,
                                d.D_Web_mm, d.E_Flange_mm, d.F_Length_mm,
                                poh.PONo
                            FROM MPR_Details d
                            INNER JOIN MPR_Header h ON h.MPR_ID = d.MPR_ID
                            INNER JOIN PO_Detail pod ON pod.MPR_Detail_ID = d.Detail_ID
                            INNER JOIN PO_head poh ON poh.PO_ID = pod.PO_ID
                            WHERE h.MPR_ID <> @mprId
                              AND (
                                  h.MPR_No = @mprNoBase
                                  OR h.MPR_No LIKE @mprNoPattern
                              )";

                        using (var cmd = new SqlCommand(sqlRevise, conn))
                        {
                            cmd.Parameters.AddWithValue("@mprId", mprId);
                            cmd.Parameters.AddWithValue("@mprNoBase", mprNoBase);
                            cmd.Parameters.AddWithValue("@mprNoPattern", mprNoBase + "%");

                            // Build lookup từ các phiên bản revise: key = (ItemNo, ItemName, Material)
                            var revisePOLookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    int itemNo = Convert.ToInt32(reader["Item_No"]);
                                    string itemName = reader["item_name"]?.ToString() ?? "";
                                    string material = reader["Material"]?.ToString() ?? "";
                                    string poNo = reader["PONo"]?.ToString() ?? "";

                                    // Key khớp chặt: ItemNo + ItemName + Material
                                    string key = $"{itemNo}|{itemName.Trim().ToLower()}|{material.Trim().ToLower()}";

                                    if (revisePOLookup.ContainsKey(key))
                                    {
                                        if (!revisePOLookup[key].Contains(poNo))
                                            revisePOLookup[key] += ", " + poNo;
                                    }
                                    else
                                    {
                                        revisePOLookup[key] = poNo;
                                    }
                                }
                            }

                            // Gán PO vào các dòng của MPR hiện tại theo key khớp
                            foreach (var detail in currentDetails)
                            {
                                string key = $"{detail.ItemNo}|{detail.ItemName.Trim().ToLower()}|{detail.Material.Trim().ToLower()}";
                                if (revisePOLookup.TryGetValue(key, out string poNos))
                                {
                                    dict[detail.DetailId] = poNos;
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
            LoadFiles(m.Project_Name);
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
            if (!PermissionHelper.Check("MPR", "Tạo MPR", "Tạo MPR")) return;
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
            if (!PermissionHelper.Check("MPR", "Lưu Header", "Lưu Header")) return;
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

        // =====================================================================
        // XÓA MPR — CASCADE DELETE ĐÚNG THỨ TỰ TRONG TRANSACTION
        // Thứ tự: PO_Detail (FK) → MPR_Details → MPR_Header
        // =====================================================================
        private void BtnDeleteMPR_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("MPR", "Xóa MPR", "Xóa MPR")) return;
            if (_selectedMPR_ID == 0)
            {
                MessageBox.Show("Vui lòng chọn phiếu MPR cần xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Lấy tên MPR để hiển thị xác nhận
            var mprToDelete = _mprList.Find(m => m.MPR_ID == _selectedMPR_ID);
            string mprNoDisplay = mprToDelete?.MPR_No ?? _selectedMPR_ID.ToString();

            string confirmMsg =
                $"Bạn có chắc chắn muốn xóa phiếu MPR: [{mprNoDisplay}] ?\n\n" +
                $"⚠ Thao tác này sẽ xóa toàn bộ:\n" +
                $"   • Chi tiết vật tư của phiếu MPR này\n" +
                $"   • Liên kết PO_Detail đến các dòng vật tư\n\n" +
                $"Dữ liệu sẽ KHÔNG thể khôi phục!";

            if (MessageBox.Show(confirmMsg, "⚠ Xác nhận xóa MPR",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2) != DialogResult.Yes)
                return;

            try
            {
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    using (var tran = conn.BeginTransaction())
                    {
                        try
                        {
                            // Bước 1: Xóa PO_Detail liên kết đến các dòng MPR_Details của MPR này
                            // (giải phóng FK_PO_Detail_MPR_Details trước khi xóa MPR_Details)
                            var cmd1 = new SqlCommand(@"
                                DELETE pod
                                FROM dbo.PO_Detail pod
                                INNER JOIN dbo.MPR_Details md ON pod.MPR_Detail_ID = md.Detail_ID
                                WHERE md.MPR_ID = @mprId", conn, tran);
                            cmd1.Parameters.AddWithValue("@mprId", _selectedMPR_ID);
                            int poDetailDeleted = cmd1.ExecuteNonQuery();

                            // Bước 2: Xóa toàn bộ MPR_Details của MPR này
                            var cmd2 = new SqlCommand(
                                "DELETE FROM dbo.MPR_Details WHERE MPR_ID = @mprId", conn, tran);
                            cmd2.Parameters.AddWithValue("@mprId", _selectedMPR_ID);
                            int detailDeleted = cmd2.ExecuteNonQuery();

                            // Bước 3: Xóa MPR Header
                            var cmd3 = new SqlCommand(
                                "DELETE FROM dbo.MPR_Header WHERE MPR_ID = @mprId", conn, tran);
                            cmd3.Parameters.AddWithValue("@mprId", _selectedMPR_ID);
                            cmd3.ExecuteNonQuery();

                            tran.Commit();

                            string resultMsg = $"✅ Xóa phiếu MPR [{mprNoDisplay}] thành công!\n\n" +
                                               $"   • {detailDeleted} dòng vật tư đã xóa\n" +
                                               $"   • {poDetailDeleted} liên kết PO_Detail đã xóa";
                            MessageBox.Show(resultMsg, "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch
                        {
                            tran.Rollback();
                            throw;
                        }
                    }
                }

                _selectedMPR_ID = 0;
                ClearHeader();
                dgvDetails.Rows.Clear();
                dgvPOProgress.DataSource = null;
                dgvFiles.Rows.Clear();
                LoadMPR();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi xóa MPR: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            if (!PermissionHelper.Check("MPR", "Tạo PO", "Tạo PO từ MPR")) return;
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
            if (!PermissionHelper.Check("MPR", "Thêm dòng", "Thêm dòng")) return;
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
            if (!PermissionHelper.Check("MPR", "Xóa dòng", "Xóa dòng")) return;
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
            if (!PermissionHelper.Check("MPR", "Lưu chi tiết", "Lưu chi tiết")) return;
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

            // ── Xác thực mật khẩu Admin trước khi lưu ──
            if (!VerifyAdminPassword()) return;

            try
            {
                int saved = 0;
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    foreach (DataGridViewRow row in dgvDetails.Rows)
                    {
                        if (row.IsNewRow) continue;
                        string itemName = row.Cells["Item_Name"].Value?.ToString() ?? "";
                        if (string.IsNullOrWhiteSpace(itemName)) continue;

                        int detailId = Convert.ToInt32(row.Cells["Detail_ID"].Value ?? 0);
                        int itemNo = Convert.ToInt32(row.Cells["Item_No"].Value ?? 0);
                        string desc = row.Cells["Description"].Value?.ToString() ?? "";
                        string material = row.Cells["Material"].Value?.ToString() ?? "";
                        decimal thickMm = DecimalVal(row.Cells["Thickness_mm"].Value);
                        decimal depthMm = DecimalVal(row.Cells["Depth_mm"].Value);
                        decimal cWidthMm = DecimalVal(row.Cells["C_Width_mm"].Value);
                        decimal dWebMm = DecimalVal(row.Cells["D_Web_mm"].Value);
                        decimal eFlangeMm = DecimalVal(row.Cells["E_Flange_mm"].Value);
                        decimal fLengthMm = DecimalVal(row.Cells["F_Length_mm"].Value);
                        string unit = row.Cells["UNIT"].Value?.ToString() ?? "";
                        int qty = (int)DecimalVal(row.Cells["Qty"].Value);
                        decimal weight = DecimalVal(row.Cells["Weight"].Value);
                        string mpsInfo = row.Cells["MPS_Info"].Value?.ToString() ?? "";
                        string usageLoc = row.Cells["Usage_Location"].Value?.ToString() ?? "";
                        string rev = row.Cells["REV"].Value?.ToString() ?? "0";
                        string remarks = row.Cells["Remarks"].Value?.ToString() ?? "";
                        string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                        Microsoft.Data.SqlClient.SqlCommand cmd;

                        if (detailId == 0)
                        {
                            cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                                INSERT INTO MPR_Details
                                    (MPR_ID, Item_No, item_name, Description, Material,
                                     Thickness_mm, Depth_mm, C_Width_mm, D_Web_mm, E_Flange_mm, F_Length_mm,
                                     UNIT, Qty_Per_Sheet, Weight_kg, MPS_Info, Usage_Location, REV, Remarks,
                                     Created_Date, Created_By, Modified_Date, Modified_By)
                                VALUES
                                    (@mprId, @itemNo, @itemName, @desc, @material,
                                     @thick, @depth, @cWidth, @dWeb, @eFlange, @fLen,
                                     @unit, @qty, @weight, @mps, @usage, @rev, @remarks,
                                     @now, @user, @now, @user);
                                SELECT SCOPE_IDENTITY();", conn);
                        }
                        else
                        {
                            cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                                UPDATE MPR_Details SET
                                    Item_No       = @itemNo,
                                    item_name     = @itemName,
                                    Description   = @desc,
                                    Material      = @material,
                                    Thickness_mm  = @thick,
                                    Depth_mm      = @depth,
                                    C_Width_mm    = @cWidth,
                                    D_Web_mm      = @dWeb,
                                    E_Flange_mm   = @eFlange,
                                    F_Length_mm   = @fLen,
                                    UNIT          = @unit,
                                    Qty_Per_Sheet = @qty,
                                    Weight_kg     = @weight,
                                    MPS_Info      = @mps,
                                    Usage_Location= @usage,
                                    REV           = @rev,
                                    Remarks       = @remarks,
                                    Modified_Date = @now,
                                    Modified_By   = @user
                                WHERE Detail_ID = @detailId", conn);
                            cmd.Parameters.AddWithValue("@detailId", detailId);
                        }

                        cmd.Parameters.AddWithValue("@mprId", _selectedMPR_ID);
                        cmd.Parameters.AddWithValue("@itemNo", itemNo);
                        cmd.Parameters.AddWithValue("@itemName", itemName);
                        cmd.Parameters.AddWithValue("@desc", desc);
                        cmd.Parameters.AddWithValue("@material", material);
                        cmd.Parameters.AddWithValue("@thick", thickMm);
                        cmd.Parameters.AddWithValue("@depth", depthMm);
                        cmd.Parameters.AddWithValue("@cWidth", cWidthMm);
                        cmd.Parameters.AddWithValue("@dWeb", dWebMm);
                        cmd.Parameters.AddWithValue("@eFlange", eFlangeMm);
                        cmd.Parameters.AddWithValue("@fLen", fLengthMm);
                        cmd.Parameters.AddWithValue("@unit", unit);
                        cmd.Parameters.AddWithValue("@qty", qty);
                        cmd.Parameters.AddWithValue("@weight", weight);
                        cmd.Parameters.AddWithValue("@mps", mpsInfo);
                        cmd.Parameters.AddWithValue("@usage", usageLoc);
                        cmd.Parameters.AddWithValue("@rev", rev);
                        cmd.Parameters.AddWithValue("@remarks", remarks);
                        cmd.Parameters.AddWithValue("@now", now);
                        cmd.Parameters.AddWithValue("@user", _currentUser ?? "Admin");

                        if (detailId == 0)
                        {
                            var newId = cmd.ExecuteScalar();
                            if (newId != null && newId != DBNull.Value)
                                row.Cells["Detail_ID"].Value = Convert.ToInt32(newId);
                        }
                        else
                        {
                            cmd.ExecuteNonQuery();
                        }

                        saved++;
                    }
                }

                MessageBox.Show($"✅ Đã lưu {saved} dòng chi tiết thành công!", "Thành công",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDetails(_selectedMPR_ID);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi lưu chi tiết: " + ex.Message, "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =========================================================================
        // XÁC THỰC MẬT KHẨU ADMIN
        // =========================================================================
        private bool VerifyAdminPassword()
        {
            var dlg = new Form
            {
                Text = "🔐 Xác thực Admin",
                Size = new Size(380, 170),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(245, 245, 245),
                KeyPreview = true
            };

            dlg.Controls.Add(new Label
            {
                Text = "Nhập mật khẩu tài khoản Admin để xác nhận lưu:",
                Font = new Font("Segoe UI", 9),
                Location = new Point(15, 15),
                Size = new Size(340, 20)
            });

            var txtPwd = new TextBox
            {
                Location = new Point(15, 40),
                Size = new Size(340, 26),
                Font = new Font("Segoe UI", 10),
                PasswordChar = '●',
                UseSystemPasswordChar = false
            };
            dlg.Controls.Add(txtPwd);

            var lblErr = new Label
            {
                Text = "",
                ForeColor = Color.FromArgb(220, 53, 69),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Location = new Point(15, 72),
                Size = new Size(340, 20)
            };
            dlg.Controls.Add(lblErr);

            var btnOK = new Button
            {
                Text = "✔ Xác nhận",
                Location = new Point(155, 98),
                Size = new Size(100, 30),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            btnOK.FlatAppearance.BorderSize = 0;
            dlg.Controls.Add(btnOK);

            var btnCancel = new Button
            {
                Text = "Hủy",
                Location = new Point(265, 98),
                Size = new Size(90, 30),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                DialogResult = DialogResult.Cancel
            };
            btnCancel.FlatAppearance.BorderSize = 0;
            dlg.Controls.Add(btnCancel);
            dlg.CancelButton = btnCancel;

            bool verified = false;

            btnOK.Click += (s, ev) =>
            {
                string pwd = txtPwd.Text;
                if (string.IsNullOrEmpty(pwd))
                { lblErr.Text = "Vui lòng nhập mật khẩu!"; return; }

                try
                {
                    string inputHash;
                    using (var sha256 = System.Security.Cryptography.SHA256.Create())
                    {
                        byte[] bytes = sha256.ComputeHash(System.Text.Encoding.UTF8.GetBytes(pwd));
                        inputHash = BitConverter.ToString(bytes).Replace("-", "").ToLower();
                    }

                    const string ADMIN_HASH = "e86f78a8a3caf0b60d8e74e5942aa6d86dc150cd3c03338aef25b7d2d7e3acc7";

                    bool match = false;

                    if (inputHash == ADMIN_HASH)
                    {
                        match = true;
                    }
                    else
                    {
                        using var conn = DatabaseHelper.GetConnection();
                        conn.Open();

                        var cmd1 = new Microsoft.Data.SqlClient.SqlCommand(
                            @"SELECT COUNT(1) FROM Users
                              WHERE LOWER(Username) = 'admin'
                                AND (LOWER(Password) = @hash
                                  OR Password = @hashUpper)", conn);
                        cmd1.Parameters.AddWithValue("@hash", inputHash);
                        cmd1.Parameters.AddWithValue("@hashUpper", inputHash.ToUpper());
                        if (Convert.ToInt32(cmd1.ExecuteScalar()) > 0)
                            match = true;

                        if (!match)
                        {
                            var cmd2 = new Microsoft.Data.SqlClient.SqlCommand(
                                @"SELECT COUNT(1) FROM Users
                                  WHERE LOWER(Username) = 'admin'
                                    AND Password = @pwd", conn);
                            cmd2.Parameters.AddWithValue("@pwd", pwd);
                            if (Convert.ToInt32(cmd2.ExecuteScalar()) > 0)
                                match = true;
                        }
                    }

                    if (match)
                    {
                        verified = true;
                        dlg.DialogResult = DialogResult.OK;
                        dlg.Close();
                    }
                    else
                    {
                        lblErr.Text = "❌ Mật khẩu không đúng!";
                        txtPwd.Clear();
                        txtPwd.Focus();
                    }
                }
                catch (Exception ex)
                {
                    lblErr.Text = "Lỗi xác thực: " + ex.Message;
                }
            };

            dlg.KeyDown += (s, ev) =>
            {
                if (ev.KeyCode == Keys.Enter) { btnOK.PerformClick(); ev.SuppressKeyPress = true; }
            };

            txtPwd.Focus();
            dlg.ShowDialog(this);
            return verified;
        }

        // =========================================================================
        // CHECK ALL ITEMS — Popup tổng hợp toàn bộ MPR Detail + RIR kết quả KT
        // =========================================================================
        private void BtnCheckAllItems_Click(object sender, EventArgs e)
        {
            if (!PermissionHelper.Check("MPR", "Check All Items", "Check All Items")) return;
            ShowCheckAllItemsPopup();
        }

        private void ShowCheckAllItemsPopup()
        {
            try
            {
                const string SQL = @"
                    SELECT
                        ISNULL(pi.ProjectCode,  N'')                        AS [Mã dự án],
                        h.MPR_No                                            AS [MPR No],
                        h.Rev                                               AS [Rev],
                        d.Item_No                                           AS [Item No],
                        d.item_name                                         AS [Tên vật tư],
                        d.Material                                          AS [Vật liệu],
                        ISNULL(CAST(NULLIF(d.Thickness_mm,0) AS NVARCHAR), N'') AS [A-Dày(mm)],
                        ISNULL(CAST(NULLIF(d.Depth_mm,    0) AS NVARCHAR), N'') AS [B-Sâu(mm)],
                        ISNULL(CAST(NULLIF(d.C_Width_mm,  0) AS NVARCHAR), N'') AS [C-Rộng(mm)],
                        ISNULL(CAST(NULLIF(d.D_Web_mm,    0) AS NVARCHAR), N'') AS [D-Bụng(mm)],
                        ISNULL(CAST(NULLIF(d.E_Flange_mm, 0) AS NVARCHAR), N'') AS [E-Cánh(mm)],
                        ISNULL(CAST(NULLIF(d.F_Length_mm, 0) AS NVARCHAR), N'') AS [F-Dài(mm)],
                        ISNULL(d.UNIT,          N'')                        AS [ĐVT],
                        d.Qty_Per_Sheet                                     AS [SL],
                        ISNULL(d.Weight_kg,     0)                          AS [KG],
                        ISNULL(STUFF((
                            SELECT DISTINCT N', ' + ph2.PONo
                            FROM PO_Detail pod2
                            INNER JOIN PO_head ph2 ON ph2.PO_ID = pod2.PO_ID
                            WHERE pod2.MPR_Detail_ID = d.Detail_ID
                            FOR XML PATH(''), TYPE
                        ).value('.','NVARCHAR(MAX)'), 1, 2, N''), N'')      AS [PO No],
                        ISNULL(STUFF((
                            SELECT DISTINCT N', ' + rd2.Heatno
                            FROM PO_Detail pod3
                            INNER JOIN PO_head   ph3 ON ph3.PO_ID  = pod3.PO_ID
                            INNER JOIN RIR_head  rh2 ON rh2.PONo   = ph3.PONo
                            INNER JOIN RIR_detail rd2 ON rd2.RIR_ID = rh2.RIR_ID
                            WHERE pod3.MPR_Detail_ID = d.Detail_ID
                              AND ISNULL(rd2.Heatno, N'') != N''
                            FOR XML PATH(''), TYPE
                        ).value('.','NVARCHAR(MAX)'), 1, 2, N''), N'')      AS [Heat No],
                        ISNULL((
                            SELECT TOP 1
                                CASE rd3.Inspect_Result
                                    WHEN N'Fail' THEN N'Fail'
                                    WHEN N'Hold' THEN N'Hold'
                                    WHEN N'Pass' THEN N'Pass'
                                    ELSE N'Chưa KT'
                                END
                            FROM PO_Detail pod4
                            INNER JOIN PO_head    ph4 ON ph4.PO_ID  = pod4.PO_ID
                            INNER JOIN RIR_head   rh3 ON rh3.PONo   = ph4.PONo
                            INNER JOIN RIR_detail rd3 ON rd3.RIR_ID = rh3.RIR_ID
                            WHERE pod4.MPR_Detail_ID = d.Detail_ID
                            ORDER BY
                                CASE rd3.Inspect_Result
                                    WHEN N'Fail' THEN 1
                                    WHEN N'Hold' THEN 2
                                    WHEN N'Pass' THEN 3
                                    ELSE 4
                                END
                        ), N'Chưa KT')                                      AS [Kết quả KT],
                        ISNULL(STUFF((
                            SELECT DISTINCT N', ' + rh4.RIR_No
                            FROM PO_Detail pod5
                            INNER JOIN PO_head  ph5 ON ph5.PO_ID = pod5.PO_ID
                            INNER JOIN RIR_head rh4 ON rh4.PONo  = ph5.PONo
                            WHERE pod5.MPR_Detail_ID = d.Detail_ID
                              AND ISNULL(rh4.RIR_No, N'') != N''
                            FOR XML PATH(''), TYPE
                        ).value('.','NVARCHAR(MAX)'), 1, 2, N''), N'')      AS [RIR No],
                        ISNULL((
                            SELECT TOP 1 rh5.Status
                            FROM PO_Detail pod6
                            INNER JOIN PO_head  ph6 ON ph6.PO_ID = pod6.PO_ID
                            INNER JOIN RIR_head rh5 ON rh5.PONo  = ph6.PONo
                            WHERE pod6.MPR_Detail_ID = d.Detail_ID
                              AND ISNULL(rh5.Status, N'') != N''
                            ORDER BY rh5.Issue_Date DESC
                        ), N'')                                             AS [Trạng thái RIR]
                    FROM MPR_Header  h
                    INNER JOIN MPR_Details d  ON d.MPR_ID = h.MPR_ID
                    LEFT  JOIN ProjectInfo pi ON pi.ProjectCode = h.Project_Code
                    ORDER BY pi.ProjectCode, h.MPR_No, d.Item_No";

                DataTable dtFull;
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    dtFull = new DataTable();
                    dtFull.Load(new SqlCommand(SQL, conn).ExecuteReader());
                }

                // ── Popup ──
                var popup = new Form
                {
                    Text = "🔎 Check All Items — Tổng hợp toàn bộ vật tư MPR",
                    Size = new Size(1400, 700),
                    StartPosition = FormStartPosition.CenterParent,
                    BackColor = Color.FromArgb(245, 245, 245),
                    MinimumSize = new Size(1100, 500),
                    KeyPreview = true
                };

                popup.Controls.Add(new Label
                {
                    Text = "🔎  CHECK ALL ITEMS — Tổng hợp vật tư tất cả MPR & kết quả kiểm tra RIR  |  💡 Double click → mở MPR",
                    Font = new Font("Segoe UI", 10, FontStyle.Bold),
                    ForeColor = Color.FromArgb(102, 51, 153),
                    Location = new Point(10, 8),
                    Size = new Size(900, 24)
                });

                int total = dtFull.Rows.Count;
                int pass = 0, fail = 0, hold = 0, notYet = 0;
                foreach (DataRow r in dtFull.Rows)
                {
                    string kt = r["Kết quả KT"]?.ToString() ?? "";
                    if (kt == "Pass") pass++;
                    else if (kt == "Fail") fail++;
                    else if (kt == "Hold") hold++;
                    else notYet++;
                }
                var lblStat = new Label
                {
                    Text = $"Tổng: {total}  |  ✅ Pass: {pass}  |  ❌ Fail: {fail}  |  ⏸ Hold: {hold}  |  ⏳ Chưa KT: {notYet}",
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    ForeColor = Color.FromArgb(0, 120, 212),
                    Location = new Point(10, 36),
                    Size = new Size(1360, 20),
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
                };
                popup.Controls.Add(lblStat);

                var pFilter = new Panel
                {
                    Location = new Point(10, 62),
                    Size = new Size(popup.ClientSize.Width - 20, 72),
                    BackColor = Color.White,
                    BorderStyle = BorderStyle.FixedSingle,
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
                };
                popup.Controls.Add(pFilter);

                Action<string, int, int, int> addFL = (txt, x, y2, w) =>
                    pFilter.Controls.Add(new Label
                    {
                        Text = txt,
                        Location = new Point(x, y2 + 3),
                        Size = new Size(w, 18),
                        Font = new Font("Segoe UI", 8, FontStyle.Bold),
                        ForeColor = Color.FromArgb(60, 60, 60)
                    });

                int row1Y = 6;
                int x1 = 6;

                addFL("Mã DA:", x1, row1Y, 48);
                var cboDuAn = new ComboBox
                {
                    Location = new Point(x1 + 50, row1Y),
                    Size = new Size(115, 22),
                    Font = new Font("Segoe UI", 9),
                    DropDownStyle = ComboBoxStyle.DropDownList
                };
                cboDuAn.Items.Add("Tất cả");
                dtFull.AsEnumerable()
                    .Select(r => r["Mã dự án"].ToString()).Where(v => !string.IsNullOrEmpty(v))
                    .Distinct().OrderBy(v => v).ToList().ForEach(v => cboDuAn.Items.Add(v));
                cboDuAn.SelectedIndex = 0;
                pFilter.Controls.Add(cboDuAn);
                x1 += 175;

                addFL("Tên VT:", x1, row1Y, 50);
                var txtFName = new TextBox { Location = new Point(x1 + 52, row1Y), Size = new Size(140, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "Tên vật tư..." };
                pFilter.Controls.Add(txtFName);
                x1 += 202;

                addFL("A(mm):", x1, row1Y, 48);
                var txtFA = new TextBox { Location = new Point(x1 + 50, row1Y), Size = new Size(44, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "A" };
                pFilter.Controls.Add(txtFA); x1 += 102;

                addFL("B(mm):", x1, row1Y, 48);
                var txtFB = new TextBox { Location = new Point(x1 + 50, row1Y), Size = new Size(44, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "B" };
                pFilter.Controls.Add(txtFB); x1 += 102;

                addFL("C(mm):", x1, row1Y, 48);
                var txtFC = new TextBox { Location = new Point(x1 + 50, row1Y), Size = new Size(44, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "C" };
                pFilter.Controls.Add(txtFC); x1 += 102;

                addFL("D(mm):", x1, row1Y, 48);
                var txtFD = new TextBox { Location = new Point(x1 + 50, row1Y), Size = new Size(44, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "D" };
                pFilter.Controls.Add(txtFD); x1 += 102;

                addFL("E(mm):", x1, row1Y, 48);
                var txtFE = new TextBox { Location = new Point(x1 + 50, row1Y), Size = new Size(44, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "E" };
                pFilter.Controls.Add(txtFE); x1 += 102;

                addFL("F(mm):", x1, row1Y, 48);
                var txtFF = new TextBox { Location = new Point(x1 + 50, row1Y), Size = new Size(44, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "F" };
                pFilter.Controls.Add(txtFF);

                int row2Y = 38;
                int x2 = 6;

                addFL("Heat No:", x2, row2Y, 55);
                var txtFHeat = new TextBox { Location = new Point(x2 + 57, row2Y), Size = new Size(100, 22), Font = new Font("Segoe UI", 9), PlaceholderText = "Heat No..." };
                pFilter.Controls.Add(txtFHeat);
                x2 += 167;

                addFL("KQ KT:", x2, row2Y, 48);
                var cboFKQ = new ComboBox
                {
                    Location = new Point(x2 + 50, row2Y),
                    Size = new Size(100, 22),
                    Font = new Font("Segoe UI", 9),
                    DropDownStyle = ComboBoxStyle.DropDownList
                };
                cboFKQ.Items.AddRange(new[] { "Tất cả", "Pass", "Fail", "Hold", "Chưa KT" });
                cboFKQ.SelectedIndex = 0;
                pFilter.Controls.Add(cboFKQ);
                x2 += 160;

                addFL("Vật liệu:", x2, row2Y, 58);
                var txtFMat = new TextBox
                {
                    Location = new Point(x2 + 60, row2Y),
                    Size = new Size(110, 22),
                    Font = new Font("Segoe UI", 9),
                    PlaceholderText = "Vật liệu..."
                };
                pFilter.Controls.Add(txtFMat);
                x2 += 178;

                var btnFSearch = new Button
                {
                    Text = "🔍 Lọc",
                    Location = new Point(x2, row2Y - 2),
                    Size = new Size(75, 26),
                    BackColor = Color.FromArgb(0, 120, 212),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 8, FontStyle.Bold)
                };
                btnFSearch.FlatAppearance.BorderSize = 0;
                pFilter.Controls.Add(btnFSearch);

                var btnFClear = new Button
                {
                    Text = "✖ Xóa lọc",
                    Location = new Point(x2 + 79, row2Y - 2),
                    Size = new Size(85, 26),
                    BackColor = Color.FromArgb(108, 117, 125),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 8, FontStyle.Bold)
                };
                btnFClear.FlatAppearance.BorderSize = 0;
                pFilter.Controls.Add(btnFClear);

                foreach (Control c in pFilter.Controls)
                    if (c is TextBox || c is ComboBox) c.BringToFront();

                var dgv = new DataGridView
                {
                    Location = new Point(10, 142),
                    Size = new Size(popup.ClientSize.Width - 20, popup.ClientSize.Height - 190),
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
                dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(102, 51, 153);
                dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                dgv.EnableHeadersVisualStyles = false;
                dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 240, 255);
                dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 210, 255);
                dgv.DefaultCellStyle.SelectionForeColor = Color.Black;
                popup.Controls.Add(dgv);

                dgv.CellFormatting += (s, ev) =>
                {
                    if (ev.RowIndex < 0) return;
                    string col = dgv.Columns[ev.ColumnIndex].Name;
                    if (col == "Kết quả KT")
                    {
                        string v = ev.Value?.ToString() ?? "";
                        ev.CellStyle.ForeColor =
                            v == "Pass" ? Color.FromArgb(40, 167, 69) :
                            v == "Fail" ? Color.FromArgb(220, 53, 69) :
                            v == "Hold" ? Color.FromArgb(255, 140, 0) :
                            Color.FromArgb(108, 117, 125);
                        ev.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    }
                    if (col == "Trạng thái RIR")
                    {
                        string v = ev.Value?.ToString() ?? "";
                        ev.CellStyle.ForeColor =
                            v == "Hoàn thành" ? Color.FromArgb(40, 167, 69) :
                            v == "Đang kiểm tra" ? Color.FromArgb(255, 140, 0) :
                            string.IsNullOrEmpty(v) ? Color.FromArgb(180, 180, 180) :
                            Color.FromArgb(0, 120, 212);
                        ev.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    }
                };

                dgv.RowPrePaint += (s, ev) =>
                {
                    if (ev.RowIndex < 0 || dgv.Rows[ev.RowIndex].IsNewRow) return;
                    string kt = dgv.Rows[ev.RowIndex].Cells["Kết quả KT"].Value?.ToString() ?? "";
                    dgv.Rows[ev.RowIndex].DefaultCellStyle.BackColor =
                        kt == "Pass" ? Color.FromArgb(235, 255, 235) :
                        kt == "Fail" ? Color.FromArgb(255, 235, 235) :
                        kt == "Hold" ? Color.FromArgb(255, 248, 230) :
                        Color.White;
                };

                Action applyFilter = () =>
                {
                    string selDA = cboDuAn.SelectedItem?.ToString() ?? "Tất cả";
                    string kName = txtFName.Text.Trim().ToLower();
                    string kA = txtFA.Text.Trim();
                    string kB = txtFB.Text.Trim();
                    string kC = txtFC.Text.Trim();
                    string kD = txtFD.Text.Trim();
                    string kE = txtFE.Text.Trim();
                    string kF = txtFF.Text.Trim();
                    string kHeat = txtFHeat.Text.Trim().ToLower();
                    string kKQ = cboFKQ.SelectedItem?.ToString() ?? "Tất cả";
                    string kMat = txtFMat.Text.Trim().ToLower();

                    var rows = dtFull.AsEnumerable().Where(r =>
                    {
                        if (selDA != "Tất cả" && r["Mã dự án"].ToString() != selDA) return false;
                        if (!string.IsNullOrEmpty(kName) && !r["Tên vật tư"].ToString().ToLower().Contains(kName)) return false;
                        if (!string.IsNullOrEmpty(kMat) && !r["Vật liệu"].ToString().ToLower().Contains(kMat)) return false;
                        if (!string.IsNullOrEmpty(kA) && !r["A-Dày(mm)"].ToString().Contains(kA)) return false;
                        if (!string.IsNullOrEmpty(kB) && !r["B-Sâu(mm)"].ToString().Contains(kB)) return false;
                        if (!string.IsNullOrEmpty(kC) && !r["C-Rộng(mm)"].ToString().Contains(kC)) return false;
                        if (!string.IsNullOrEmpty(kD) && !r["D-Bụng(mm)"].ToString().Contains(kD)) return false;
                        if (!string.IsNullOrEmpty(kE) && !r["E-Cánh(mm)"].ToString().Contains(kE)) return false;
                        if (!string.IsNullOrEmpty(kF) && !r["F-Dài(mm)"].ToString().Contains(kF)) return false;
                        if (!string.IsNullOrEmpty(kHeat) && !r["Heat No"].ToString().ToLower().Contains(kHeat)) return false;
                        if (kKQ != "Tất cả" && r["Kết quả KT"].ToString() != kKQ) return false;
                        return true;
                    });

                    DataTable dtView = rows.Any() ? rows.CopyToDataTable() : dtFull.Clone();
                    dgv.DataSource = dtView;

                    foreach (string colName in new[] { "MPR No", "PO No" })
                    {
                        if (!dgv.Columns.Contains(colName)) continue;
                        var col = dgv.Columns[colName];
                        col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                        int measuredW = col.GetPreferredWidth(DataGridViewAutoSizeColumnMode.AllCells, true);
                        col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                        col.Width = Math.Min(measuredW, 400);
                    }

                    int t = dtView.Rows.Count, p = 0, f = 0, h2 = 0, n = 0;
                    foreach (DataRow r in dtView.Rows)
                    {
                        string kt = r["Kết quả KT"]?.ToString() ?? "";
                        if (kt == "Pass") p++;
                        else if (kt == "Fail") f++;
                        else if (kt == "Hold") h2++;
                        else n++;
                    }
                    lblStat.Text = $"Hiển thị: {t}  |  ✅ Pass: {p}  |  ❌ Fail: {f}  |  ⏸ Hold: {h2}  |  ⏳ Chưa KT: {n}";
                };

                dgv.DataSource = null;
                lblStat.Text = "Nhấn [🔍 Lọc] để tìm kiếm dữ liệu.";

                dgv.DataBindingComplete += (s, ev) =>
                {
                    foreach (string colName in new[] { "MPR No", "PO No" })
                    {
                        if (!dgv.Columns.Contains(colName)) continue;
                        var col = dgv.Columns[colName];
                        col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                        int measuredW = col.GetPreferredWidth(DataGridViewAutoSizeColumnMode.AllCells, true);
                        col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                        col.Width = Math.Min(measuredW, 400);
                    }
                };

                btnFSearch.Click += (s, ev) => applyFilter();
                btnFClear.Click += (s, ev) =>
                {
                    cboDuAn.SelectedIndex = 0;
                    txtFName.Text = "";
                    txtFMat.Text = "";
                    txtFA.Text = ""; txtFB.Text = ""; txtFC.Text = "";
                    txtFD.Text = ""; txtFE.Text = ""; txtFF.Text = "";
                    txtFHeat.Text = ""; cboFKQ.SelectedIndex = 0;
                    dgv.DataSource = null;
                    lblStat.Text = "Nhấn [🔍 Lọc] để tìm kiếm dữ liệu.";
                };

                dgv.CellDoubleClick += (s, ev) =>
                {
                    if (ev.RowIndex < 0) return;
                    string mprNo = dgv.Rows[ev.RowIndex].Cells["MPR No"].Value?.ToString() ?? "";
                    if (string.IsNullOrEmpty(mprNo)) return;

                    var target = _mprList.Find(m => m.MPR_No == mprNo);
                    if (target == null)
                    {
                        MessageBox.Show($"Không tìm thấy MPR: {mprNo}", "Thông báo",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    popup.Close();
                    SelectMPRById(target.MPR_ID);
                };

                var ttip = new ToolTip();
                ttip.SetToolTip(dgv, "Double click vào dòng để mở MPR tương ứng");

                popup.KeyDown += (s, ev) =>
                {
                    if (ev.KeyCode == Keys.Enter)
                    {
                        applyFilter();
                        ev.Handled = true;
                        ev.SuppressKeyPress = true;
                    }
                    if (ev.KeyCode == Keys.Escape) popup.Close();
                };

                var btnExport = new Button
                {
                    Text = "📥 Xuất Excel",
                    Size = new Size(120, 30),
                    BackColor = Color.FromArgb(0, 150, 100),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Right
                };
                btnExport.FlatAppearance.BorderSize = 0;
                btnExport.Location = new Point(popup.ClientSize.Width - 245, popup.ClientSize.Height - 40);
                popup.Controls.Add(btnExport);

                btnExport.Click += (s, ev) =>
                {
                    var dtExport = dgv.DataSource as DataTable;
                    if (dtExport == null || dtExport.Rows.Count == 0)
                    {
                        MessageBox.Show("Không có dữ liệu để xuất!", "Thông báo",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    var sfd = new SaveFileDialog
                    {
                        Title = "Lưu file Excel",
                        Filter = "Excel Files|*.xlsx",
                        FileName = $"CheckAllItems_{DateTime.Now:yyyyMMdd_HHmm}",
                        InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                    };
                    if (sfd.ShowDialog() != DialogResult.OK) return;

                    try
                    {
                        OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                        using (var pkg = new OfficeOpenXml.ExcelPackage())
                        {
                            var ws = pkg.Workbook.Worksheets.Add("Check All Items");

                            ws.Cells[1, 1].Value = "CHECK ALL ITEMS — Tổng hợp vật tư MPR & kết quả kiểm tra RIR";
                            ws.Cells[1, 1, 1, dtExport.Columns.Count].Merge = true;
                            ws.Cells[1, 1].Style.Font.Size = 13;
                            ws.Cells[1, 1].Style.Font.Bold = true;
                            ws.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            ws.Cells[1, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(102, 51, 153));
                            ws.Cells[1, 1].Style.Font.Color.SetColor(Color.White);

                            ws.Cells[2, 1].Value = $"Xuất ngày: {DateTime.Now:dd/MM/yyyy HH:mm}  |  Tổng: {dtExport.Rows.Count} dòng";
                            ws.Cells[2, 1, 2, dtExport.Columns.Count].Merge = true;
                            ws.Cells[2, 1].Style.Font.Italic = true;
                            ws.Cells[2, 1].Style.Font.Size = 9;

                            for (int c = 0; c < dtExport.Columns.Count; c++)
                            {
                                var cell = ws.Cells[3, c + 1];
                                cell.Value = dtExport.Columns[c].ColumnName;
                                cell.Style.Font.Bold = true;
                                cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(102, 51, 153));
                                cell.Style.Font.Color.SetColor(Color.White);
                                cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            }

                            for (int row = 0; row < dtExport.Rows.Count; row++)
                            {
                                var dr = dtExport.Rows[row];
                                bool isAlt = row % 2 == 1;

                                for (int c = 0; c < dtExport.Columns.Count; c++)
                                {
                                    var cell = ws.Cells[row + 4, c + 1];
                                    string colName = dtExport.Columns[c].ColumnName;

                                    if (colName == "KG")
                                    {
                                        if (dr[c] != DBNull.Value && decimal.TryParse(dr[c]?.ToString(), out decimal kg))
                                        {
                                            cell.Value = Math.Round(kg, 2);
                                            cell.Style.Numberformat.Format = "#,##0.00";
                                            cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                                        }
                                        else
                                        {
                                            cell.Value = 0.00m;
                                            cell.Style.Numberformat.Format = "#,##0.00";
                                        }
                                    }
                                    else
                                    {
                                        cell.Value = dr[c]?.ToString() ?? "";
                                    }

                                    cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                    cell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                                    if (isAlt)
                                    {
                                        cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                        cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(248, 240, 255));
                                    }

                                    if (colName == "Kết quả KT")
                                    {
                                        string kt = dr[c]?.ToString() ?? "";
                                        cell.Style.Font.Bold = true;
                                        cell.Style.Font.Color.SetColor(
                                            kt == "Pass" ? Color.FromArgb(40, 167, 69) :
                                            kt == "Fail" ? Color.FromArgb(220, 53, 69) :
                                            kt == "Hold" ? Color.FromArgb(255, 140, 0) :
                                            Color.FromArgb(108, 117, 125));
                                    }
                                }
                            }

                            ws.Cells[ws.Dimension.Address].AutoFitColumns();
                            ws.View.FreezePanes(4, 1);

                            pkg.SaveAs(new System.IO.FileInfo(sfd.FileName));
                        }

                        var res = MessageBox.Show(
                            $"✅ Xuất Excel thành công!\nFile: {sfd.FileName}\n\nBạn có muốn mở file không?",
                            "Thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                        if (res == DialogResult.Yes)
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                            { FileName = sfd.FileName, UseShellExecute = true });
                    }
                    catch (Exception exExport)
                    {
                        MessageBox.Show("Lỗi xuất Excel: " + exExport.Message, "Lỗi",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                };

                var btnClose = new Button
                {
                    Text = "Đóng",
                    Size = new Size(100, 30),
                    BackColor = Color.FromArgb(108, 117, 125),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                    DialogResult = DialogResult.OK
                };
                btnClose.FlatAppearance.BorderSize = 0;
                btnClose.Location = new Point(popup.ClientSize.Width - 115, popup.ClientSize.Height - 40);
                popup.Controls.Add(btnClose);
                popup.AcceptButton = btnFSearch;
                popup.CancelButton = btnClose;

                popup.Resize += (s, ev) =>
                {
                    pFilter.Width = popup.ClientSize.Width - 20;
                    btnExport.Location = new Point(popup.ClientSize.Width - 245, popup.ClientSize.Height - 40);
                    btnClose.Location = new Point(popup.ClientSize.Width - 115, popup.ClientSize.Height - 40);
                };

                popup.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        // =====================================================
        //  ÁP DỤNG PHÂN QUYỀN
        // =====================================================
        private void ApplyPermissions()
        {
            if (btnNewMPR != null) PermissionHelper.Apply(btnNewMPR, "MPR", "Tạo MPR");
            if (btnDeleteMPR != null) PermissionHelper.Apply(btnDeleteMPR, "MPR", "Xóa MPR");
            if (btnSaveHeader != null) PermissionHelper.Apply(btnSaveHeader, "MPR", "Lưu Header");
            if (btnAddDetail != null) PermissionHelper.Apply(btnAddDetail, "MPR", "Thêm dòng");
            if (btnDeleteDetail != null) PermissionHelper.Apply(btnDeleteDetail, "MPR", "Xóa dòng");
            if (btnSaveDetail != null) PermissionHelper.Apply(btnSaveDetail, "MPR", "Lưu chi tiết");
            foreach (var c in this.Controls.Find("btnCreatePO", true))
                PermissionHelper.Apply(c, "MPR", "Tạo PO");
            foreach (var c in this.Controls.Find("btnCheckAll", true))
                PermissionHelper.Apply(c, "MPR", "Check All Items");
        }

    }
}