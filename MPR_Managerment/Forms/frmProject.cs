using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;
using MPR_Managerment.Services;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace MPR_Managerment.Forms
{
    public partial class frmProject : Form
    {
        private ProjectService _service = new ProjectService();
        private List<ProjectInfo> _projects = new List<ProjectInfo>();
        private int _selectedId = 0;
        private string _currentUser = "Admin";

        // Controls - List
        private DataGridView dgvProjects;
        private TextBox txtSearch;
        private Button btnSearch, btnNew, btnSave, btnDelete, btnClear;
        private Label lblStatus;
        private ComboBox cboFilterStatus;

        // Controls - Form
        private TextBox txtProjectName, txtProjectCode, txtWorkorderNo, txtCustomer;
        private TextBox txtPOCode, txtMPRCode, txtNotes;
        private TextBox txtPOLink, txtRIRLink, txtMPRLink, txtINVLink, txtDeliveryNoteLink;
        private NumericUpDown nudWeight, nudBudget;
        private ComboBox cboStatus;

        // Panels
        private Panel panelTop, panelForm, panelStats;

        // Stats labels — hệ thống
        private Label lblMPRCount, lblPOCount, lblRIRCount, lblWeightTotal, lblBudgetTotal;

        // Stats labels — dự án đang chọn
        private Label lblMPRProject, lblPOProject, lblRIRProject, lblWeightProject, lblBudgetProject;

        public frmProject()
        {
            InitializeComponent();
            BuildUI();
            LoadProjects();
            this.Resize += FrmProject_Resize;
        }

        private void BuildUI()
        {
            this.Text = "Quản lý Dự án";
            this.BackColor = Color.FromArgb(245, 245, 245);

            // ===== PANEL STATS =====
            panelStats = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(1260, 150),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(panelStats);

            // Hàng 1 — toàn hệ thống
            panelStats.Controls.Add(new Label
            {
                Text = "🌐 Toàn hệ thống:",
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                ForeColor = Color.Gray,
                Location = new Point(10, 8),
                Size = new Size(150, 18)
            });

            lblMPRCount = AddStatCard(panelStats, "📋 Tổng MPR", "0", Color.FromArgb(0, 120, 212), 10, 28);
            lblPOCount = AddStatCard(panelStats, "🛒 Tổng PO", "0", Color.FromArgb(40, 167, 69), 220, 28);
            lblRIRCount = AddStatCard(panelStats, "📦 Tổng RIR", "0", Color.FromArgb(102, 51, 153), 430, 28);
            lblWeightTotal = AddStatCard(panelStats, "⚖ Tổng KG", "0", Color.FromArgb(255, 140, 0), 640, 28);
            lblBudgetTotal = AddStatCard(panelStats, "💰 Tổng Budget", "0", Color.FromArgb(220, 53, 69), 850, 28);

            // Hàng 2 — dự án đang chọn
            panelStats.Controls.Add(new Label
            {
                Text = "📁 Dự án đang chọn:",
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                ForeColor = Color.Gray,
                Location = new Point(10, 78),
                Size = new Size(150, 18)
            });

            lblMPRProject = AddStatCard(panelStats, "📋 MPR dự án", "0", Color.FromArgb(0, 120, 212), 10, 98);
            lblPOProject = AddStatCard(panelStats, "🛒 PO dự án", "0", Color.FromArgb(40, 167, 69), 220, 98);
            lblRIRProject = AddStatCard(panelStats, "📦 RIR dự án", "0", Color.FromArgb(102, 51, 153), 430, 98);
            lblWeightProject = AddStatCard(panelStats, "⚖ KG đặt / KG DA", "0", Color.FromArgb(255, 140, 0), 640, 98);
            lblBudgetProject = AddStatCard(panelStats, "💰 % Budget đã dùng", "0%", Color.FromArgb(220, 53, 69), 850, 98);

            // ===== PANEL TOP - Danh sách =====
            panelTop = new Panel
            {
                Location = new Point(10, 150),
                Size = new Size(1260, 280),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(panelTop);

            panelTop.Controls.Add(new Label
            {
                Text = "DANH SÁCH DỰ ÁN",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 10),
                Size = new Size(300, 28)
            });

            txtSearch = new TextBox
            {
                Location = new Point(10, 46),
                Size = new Size(250, 28),
                Font = new Font("Segoe UI", 10),
                PlaceholderText = "Tìm tên, mã dự án, khách hàng..."
            };
            panelTop.Controls.Add(txtSearch);

            txtSearch.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) BtnSearch_Click(null, null); };

            btnSearch = CreateButton("🔍 Tìm", Color.FromArgb(0, 120, 212), new Point(270, 45), 80, 30);
            btnSearch.Click += BtnSearch_Click;
            panelTop.Controls.Add(btnSearch);

            panelTop.Controls.Add(new Label
            {
                Text = "Trạng thái:",
                Location = new Point(365, 50),
                Size = new Size(75, 20),
                Font = new Font("Segoe UI", 9)
            });

            cboFilterStatus = new ComboBox
            {
                Location = new Point(445, 47),
                Size = new Size(150, 25),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            cboFilterStatus.Items.AddRange(new[] { "Tất cả", "Đang thực hiện", "Hoàn thành", "Tạm dừng", "Hủy" });
            cboFilterStatus.SelectedIndex = 0;

            cboFilterStatus.SelectedIndexChanged += (s, e) => LoadProjects();
            panelTop.Controls.Add(cboFilterStatus);

            btnNew = CreateButton("➕ Thêm mới", Color.FromArgb(40, 167, 69), new Point(615, 45), 110, 30);
            btnDelete = CreateButton("🗑 Xóa", Color.FromArgb(220, 53, 69), new Point(735, 45), 80, 30);
            btnClear = CreateButton("🔄 Làm mới", Color.FromArgb(108, 117, 125), new Point(825, 45), 100, 30);

            btnNew.Click += BtnNew_Click;
            btnDelete.Click += BtnDelete_Click;
            btnClear.Click += BtnClear_Click;

            panelTop.Controls.Add(btnNew);
            panelTop.Controls.Add(btnDelete);
            panelTop.Controls.Add(btnClear);

            lblStatus = new Label
            {
                Location = new Point(940, 50),
                Size = new Size(300, 22),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.Gray
            };
            panelTop.Controls.Add(lblStatus);

            dgvProjects = new DataGridView
            {
                Location = new Point(10, 85),
                Size = new Size(1235, 185),
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
            dgvProjects.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvProjects.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvProjects.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvProjects.EnableHeadersVisualStyles = false;
            dgvProjects.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvProjects.SelectionChanged += DgvProjects_SelectionChanged;
            dgvProjects.CellFormatting += DgvProjects_CellFormatting;
            panelTop.Controls.Add(dgvProjects);

            // ===== PANEL FORM =====
            panelForm = new Panel
            {
                Location = new Point(10, 440),
                Size = new Size(1260, 330),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            this.Controls.Add(panelForm);

            panelForm.Controls.Add(new Label
            {
                Text = "THÔNG TIN DỰ ÁN",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 8),
                Size = new Size(300, 25)
            });

            // Row 1
            int y = 38;
            AddLabel(panelForm, "Tên dự án (*):", 10, y);
            txtProjectName = AddTxt(panelForm, 120, y, 280);

            AddLabel(panelForm, "Mã dự án (*):", 415, y);
            txtProjectCode = AddTxt(panelForm, 515, y, 140);

            AddLabel(panelForm, "Workorder.No:", 670, y);
            txtWorkorderNo = AddTxt(panelForm, 780, y, 200);

            AddLabel(panelForm, "Trạng thái:", 1000, y);
            cboStatus = new ComboBox
            {
                Location = new Point(1120, y),
                Size = new Size(160, 25),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboStatus.Items.AddRange(new[] { "Active", "Completed", "Pending", "Cancel" });
            cboStatus.SelectedIndex = 0;
            panelForm.Controls.Add(cboStatus);

            // Row 2
            y += 38;
            AddLabel(panelForm, "Khách hàng:", 10, y);
            txtCustomer = AddTxt(panelForm, 120, y, 220);

            AddLabel(panelForm, "Mã PO:", 355, y);
            txtPOCode = AddTxt(panelForm, 410, y, 140);

            AddLabel(panelForm, "Mã MPR:", 565, y);
            txtMPRCode = AddTxt(panelForm, 635, y, 200);

            AddLabel(panelForm, "Tổng KG:", 840, y);
            nudWeight = new NumericUpDown
            {
                Location = new Point(935, y),
                Size = new Size(130, 25),
                Font = new Font("Segoe UI", 9),
                Maximum = 9999999,
                DecimalPlaces = 2,
                ThousandsSeparator = true
            };
            panelForm.Controls.Add(nudWeight);

            AddLabel(panelForm, "Budget:", 1070, y);
            nudBudget = new NumericUpDown
            {
                Location = new Point(1150, y),
                Size = new Size(140, 25),
                Font = new Font("Segoe UI", 9),
                Maximum = 999999999,
                DecimalPlaces = 0,
                ThousandsSeparator = true
            };
            panelForm.Controls.Add(nudBudget);

            // Row 3
            y += 38;
            AddLabel(panelForm, "PO Link:", 10, y);
            txtPOLink = AddTxt(panelForm, 120, y, 330);

            AddLabel(panelForm, "RIR Link:", 465, y);
            txtRIRLink = AddTxt(panelForm, 540, y, 300);

            AddLabel(panelForm, "MPR Link:", 855, y);
            txtMPRLink = AddTxt(panelForm, 930, y, 250);
            txtMPRLink.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            // Row 4 (Thêm INV Link và Delivery Link)
            y += 38;
            AddLabel(panelForm, "INV Link:", 10, y);
            txtINVLink = AddTxt(panelForm, 120, y, 330);

            AddLabel(panelForm, "Delivery Link:", 465, y);
            txtDeliveryNoteLink = AddTxt(panelForm, 560, y, 300);
            txtDeliveryNoteLink.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            // Row 5
            y += 38;
            AddLabel(panelForm, "Ghi chú:", 10, y);
            txtNotes = new TextBox
            {
                Location = new Point(120, y),
                Size = new Size(1110, 60),
                Font = new Font("Segoe UI", 9),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            panelForm.Controls.Add(txtNotes);

            // Buttons
            y += 80;
            btnSave = CreateButton("💾 Lưu", Color.FromArgb(0, 120, 212), new Point(10, y), 110, 32);
            btnSave.Click += BtnSave_Click;
            panelForm.Controls.Add(btnSave);

            var btnClearForm = CreateButton("🔄 Xóa form", Color.FromArgb(108, 117, 125), new Point(130, y), 110, 32);
            btnClearForm.Click += (s, e) => ClearForm();
            panelForm.Controls.Add(btnClearForm);

            Common.Common.AutoBringToFontControl(new[] { panelTop, panelForm, panelStats });
        }

        private Label AddStatCard(Panel parent, string title, string value, Color color, int x, int y)
        {
            var card = new Panel { Location = new Point(x, y), Size = new Size(195, 42), BackColor = color };
            parent.Controls.Add(card);
            card.Controls.Add(new Label
            {
                Text = title,
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(6, 3),
                Size = new Size(183, 18)
            });

            var lbl = new Label
            {
                Text = value,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(6, 22),
                Size = new Size(183, 18)
            };

            card.Controls.Add(lbl);
            return lbl;
        }

        private void AddLabel(Panel p, string text, int x, int y)
        {
            p.Controls.Add(new Label { Text = text, Location = new Point(x, y + 3), Size = new Size(105, 20), Font = new Font("Segoe UI", 9) });
        }

        private TextBox AddTxt(Panel p, int x, int y, int width)
        {
            var txt = new TextBox { Location = new Point(x, y), Size = new Size(width, 25), Font = new Font("Segoe UI", 9) };
            p.Controls.Add(txt);
            return txt;
        }

        private Button CreateButton(string text, Color color, Point loc, int w, int h)
        {
            var btn = new Button { Text = text, Location = loc, Size = new Size(w, h), BackColor = color, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Cursor = Cursors.Hand };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        // ===== RESIZE =====
        private void FrmProject_Resize(object sender, EventArgs e)
        {
            int w = this.ClientSize.Width - 20;
            int h = this.ClientSize.Height;

            panelStats.Width = w;
            panelTop.Width = w;
            panelForm.Width = w;
            panelForm.Height = h - panelForm.Top - 10;

            dgvProjects.Width = panelTop.Width - 20;
            txtNotes.Width = panelForm.Width - txtNotes.Left - 20;
            txtMPRLink.Width = panelForm.Width - txtMPRLink.Left - 20;
            txtDeliveryNoteLink.Width = panelForm.Width - txtDeliveryNoteLink.Left - 20;
        }

        // ===== LOAD DỮ LIỆU =====
        private void LoadProjects()
        {
            try
            {
                string filter = cboFilterStatus.SelectedItem?.ToString() ?? "Tất cả";
                var all = _service.GetAll();
                if (filter != "Tất cả")
                    all = all.FindAll(p => p.Status == filter);

                _projects = all;
                BindGrid(_projects);
                UpdateSystemStats();
                lblStatus.Text = $"Tổng: {_projects.Count} dự án";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BindGrid(List<ProjectInfo> list)
        {
            dgvProjects.DataSource = list.ConvertAll(p => new
            {
                ID = p.Id,
                Ten_Du_An = p.ProjectName,
                Ma_Du_An = p.ProjectCode,
                Workorder = p.WorkorderNo,
                Khach_Hang = p.Customer,
                Ma_PO = p.POCode,
                Ma_MPR = p.MPRCode,
                Tong_KG = p.PJWeight.ToString("N2"),
                Budget = p.PJBudget.ToString("N0"),
                Trang_Thai = p.Status,
                INV_Link = p.INV_Link,
                Delivery_Link = p.DeliveryNote_Link,
                Ngay_Tao = p.CreatedDate.HasValue ? p.CreatedDate.Value.ToString("dd/MM/yyyy") : "",
                Ngay_SuaDoi = p.ModifiedDate.HasValue ? p.ModifiedDate.Value.ToString("dd/MM/yyyy") : ""

            });
        }

        private void UpdateSystemStats()
        {
            try
            {
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var cmd = new SqlCommand(@"
                        SELECT
                            (SELECT COUNT(*) FROM MPR_Header)                         AS MPR_Count,
                            (SELECT COUNT(*) FROM PO_head)                            AS PO_Count,
                            (SELECT COUNT(*) FROM RIR_head)                           AS RIR_Count,
                            (SELECT ISNULL(SUM(PJWeight), 0) FROM ProjectInfo)        AS TotalWeight,
                            (SELECT ISNULL(SUM(PJBudget), 0) FROM ProjectInfo)        AS TotalBudget", conn);

                    var r = cmd.ExecuteReader();
                    if (r.Read())
                    {
                        lblMPRCount.Text = Convert.ToInt32(r["MPR_Count"]).ToString("N0");
                        lblPOCount.Text = Convert.ToInt32(r["PO_Count"]).ToString("N0");
                        lblRIRCount.Text = Convert.ToInt32(r["RIR_Count"]).ToString("N0");
                        lblWeightTotal.Text = Convert.ToDecimal(r["TotalWeight"]).ToString("N0") + " kg";
                        lblBudgetTotal.Text = Convert.ToDecimal(r["TotalBudget"]).ToString("N0");
                    }
                }
            }
            catch { }
        }

        private void UpdateProjectStats(ProjectInfo p)
        {
            try
            {
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var cmd = new SqlCommand(@"
                        SELECT
                            (SELECT COUNT(*) FROM MPR_Header WHERE Project_Code = @code)    AS MPR_Count,
                            (SELECT COUNT(*) FROM PO_head WHERE WorkorderNo  = @wono)       AS PO_Count,
                            (SELECT COUNT(*) FROM RIR_head WHERE WorkorderNo  = @wono)      AS RIR_Count,
                            (SELECT ISNULL(SUM(d.Weight_kg * d.Qty_Per_Sheet), 0)
                             FROM PO_Detail d
                             INNER JOIN PO_head h ON h.PO_ID = d.PO_ID
                             WHERE h.WorkorderNo = @wono)                                   AS KG_Ordered,
                            (SELECT ISNULL(SUM(Total_Amount), 0)
                             FROM PO_head WHERE WorkorderNo = @wono)                        AS Budget_Used", conn);

                    cmd.Parameters.AddWithValue("@code", p.ProjectCode ?? "");
                    cmd.Parameters.AddWithValue("@wono", p.WorkorderNo ?? "");

                    var r = cmd.ExecuteReader();
                    if (r.Read())
                    {
                        decimal kgOrdered = Convert.ToDecimal(r["KG_Ordered"]);
                        decimal budgetUsed = Convert.ToDecimal(r["Budget_Used"]);
                        decimal pctBudget = p.PJBudget > 0 ? Math.Round(budgetUsed * 100 / p.PJBudget, 1) : 0;

                        lblMPRProject.Text = Convert.ToInt32(r["MPR_Count"]).ToString();
                        lblPOProject.Text = Convert.ToInt32(r["PO_Count"]).ToString();
                        lblRIRProject.Text = Convert.ToInt32(r["RIR_Count"]).ToString();
                        lblWeightProject.Text = $"{kgOrdered:N0} / {p.PJWeight:N0} kg";
                        lblBudgetProject.Text = $"{pctBudget}%";
                    }
                }
            }
            catch { }
        }

        // ===== SỰ KIỆN =====
        private void BtnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                string kw = txtSearch.Text.Trim();
                string filter = cboFilterStatus.SelectedItem?.ToString() ?? "Tất cả";
                var result = string.IsNullOrEmpty(kw) ? _service.GetAll() : _service.Search(kw);

                if (filter != "Tất cả")
                    result = result.FindAll(p => p.Status == filter);

                _projects = result;
                BindGrid(_projects);
                lblStatus.Text = $"Tìm thấy: {result.Count} dự án";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tìm kiếm: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DgvProjects_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvProjects.SelectedRows.Count == 0) return;

            var row = dgvProjects.SelectedRows[0];
            _selectedId = Convert.ToInt32(row.Cells["ID"].Value);
            var p = _projects.Find(x => x.Id == _selectedId);
            if (p == null) return;

            txtProjectName.Text = p.ProjectName;
            txtProjectCode.Text = p.ProjectCode;
            txtWorkorderNo.Text = p.WorkorderNo;
            txtCustomer.Text = p.Customer;
            txtPOCode.Text = p.POCode;
            txtMPRCode.Text = p.MPRCode;
            txtNotes.Text = p.Notes;
            txtPOLink.Text = p.PO_Link;
            txtRIRLink.Text = p.RIR_Link;
            txtMPRLink.Text = p.MPR_Link;
            txtINVLink.Text = p.INV_Link;
            txtDeliveryNoteLink.Text = p.DeliveryNote_Link;
            nudWeight.Value = p.PJWeight > nudWeight.Maximum ? nudWeight.Maximum : p.PJWeight;
            nudBudget.Value = p.PJBudget > nudBudget.Maximum ? nudBudget.Maximum : p.PJBudget;

            var idx = cboStatus.Items.IndexOf(p.Status);
            cboStatus.SelectedIndex = idx >= 0 ? idx : 0;

            // Cập nhật stats cho dự án đang chọn
            UpdateProjectStats(p);
        }

        private void DgvProjects_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;

            if (dgvProjects.Columns[e.ColumnIndex].Name == "Trang_Thai")
            {
                string val = e.Value?.ToString() ?? "";
                e.CellStyle.ForeColor =
                    val == "Complete" ? Color.FromArgb(40, 167, 69) :
                    val == "Active" ? Color.FromArgb(0, 120, 212) :
                    val == "Pending" ? Color.FromArgb(255, 140, 0) :
                                              Color.FromArgb(220, 53, 69);

                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
        }

        private void BtnNew_Click(object sender, EventArgs e)
        {
            ClearForm();
            _selectedId = 0;
            txtProjectName.Focus();
            lblStatus.Text = "Đang thêm dự án mới...";
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtProjectName.Text))
            {
                MessageBox.Show("Vui lòng nhập Tên dự án!", "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtProjectName.Focus();
                return;
            }
            if (string.IsNullOrWhiteSpace(txtProjectCode.Text))
            {
                MessageBox.Show("Vui lòng nhập Mã dự án!", "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtProjectCode.Focus();
                return;
            }

            try
            {
                var p = new ProjectInfo
                {
                    Id = _selectedId,
                    ProjectName = txtProjectName.Text.Trim(),
                    ProjectCode = txtProjectCode.Text.Trim(),
                    WorkorderNo = txtWorkorderNo.Text.Trim(),
                    Customer = txtCustomer.Text.Trim(),
                    POCode = txtPOCode.Text.Trim(),
                    MPRCode = txtMPRCode.Text.Trim(),
                    PJWeight = nudWeight.Value,
                    PJBudget = nudBudget.Value,
                    Status = cboStatus.SelectedItem?.ToString() ?? "Active",
                    Notes = txtNotes.Text.Trim(),
                    PO_Link = txtPOLink.Text.Trim(),
                    RIR_Link = txtRIRLink.Text.Trim(),
                    MPR_Link = txtMPRLink.Text.Trim(),
                    INV_Link = txtINVLink.Text.Trim(),
                    DeliveryNote_Link = txtDeliveryNoteLink.Text.Trim()
                };

                if (_selectedId == 0)
                {
                    _service.Insert(p, _currentUser);
                    MessageBox.Show("Thêm dự án thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    _service.Update(p);
                    MessageBox.Show("Cập nhật thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                LoadProjects();
                ClearForm();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi lưu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (_selectedId == 0)
            {
                MessageBox.Show("Vui lòng chọn dự án cần xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBox.Show("Bạn có chắc muốn xóa dự án này?", "Xác nhận",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    _service.Delete(_selectedId);
                    MessageBox.Show("Xóa thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadProjects();
                    ClearForm();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi xóa: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            txtSearch.Text = "";
            cboFilterStatus.SelectedIndex = 0;
            ClearForm();
            LoadProjects();
        }

        private void ClearForm()
        {
            _selectedId = 0;
            txtProjectName.Text = "";
            txtProjectCode.Text = "";
            txtWorkorderNo.Text = "";
            txtCustomer.Text = "";
            txtPOCode.Text = "";
            txtMPRCode.Text = "";
            txtNotes.Text = "";
            txtPOLink.Text = "";
            txtRIRLink.Text = "";
            txtMPRLink.Text = "";
            txtINVLink.Text = "";
            txtDeliveryNoteLink.Text = "";
            nudWeight.Value = 0;
            nudBudget.Value = 0;
            cboStatus.SelectedIndex = 0;
            lblStatus.Text = "";

            // Reset project stats
            lblMPRProject.Text = "0";
            lblPOProject.Text = "0";
            lblRIRProject.Text = "0";
            lblWeightProject.Text = "0";
            lblBudgetProject.Text = "0%";
        }
    }
}