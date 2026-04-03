using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using MPR_Managerment.Models;
using MPR_Managerment.Services;

namespace MPR_Managerment.Forms
{
    public partial class frmSupplier : Form
    {
        // ── Service (đã sửa Search dùng SQL thẳng, không dùng SP lỗi) ──
        private readonly SupplierService _service = new SupplierService();

        private List<Supplier> _suppliers = new List<Supplier>();
        private int _selectedSupplierID = 0;
        private string _currentUser = "Admin";

        // ── Controls ──────────────────────────────────────────────────
        private DataGridView dgvSuppliers;
        private TextBox txtSearch;
        private TextBox txtCompanyName, txtShortName, txtSupplierType;
        private TextBox txtTaxCode, txtContactPerson, txtContactPhone;
        private TextBox txtEmail, txtAddress;
        private TextBox txtBankAccount, txtBankName;
        private TextBox txtWebsite, txtCert, txtNotes;
        private CheckBox chkIsActive;
        private Button btnSearch, btnNew, btnSave, btnDelete, btnClear;
        private Label lblStatus;
        private Panel panelLeft, panelRight;

        public frmSupplier()
        {
            InitializeComponent();
            BuildUI();
            LoadSuppliers();
        }

        // =================================================================
        // BUILD UI
        // =================================================================
        private void BuildUI()
        {
            this.Text = "Quản lý Nhà Cung Cấp";
            this.Size = new Size(1200, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(245, 245, 245);
            this.MinimumSize = new Size(1000, 600);

            // ── PANEL LEFT ────────────────────────────────────────────
            panelLeft = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(560, 640),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom
            };
            this.Controls.Add(panelLeft);

            panelLeft.Controls.Add(new Label
            {
                Text = "DANH SÁCH NHÀ CUNG CẤP",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 10),
                Size = new Size(400, 28)
            });

            // Ô tìm kiếm
            txtSearch = new TextBox
            {
                Location = new Point(10, 48),
                Size = new Size(390, 28),
                Font = new Font("Segoe UI", 10),
                PlaceholderText = "Gõ để tìm theo tên, SĐT, email... (real-time)"
            };
            panelLeft.Controls.Add(txtSearch);

            // Real-time: gõ là lọc ngay — KHÔNG gọi service.Search (SP lỗi)
            txtSearch.TextChanged += (s, e) => FilterSuppliers();
            txtSearch.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) FilterSuppliers(); };

            btnSearch = MkBtn("🔍 Tìm", Color.FromArgb(0, 120, 212), new Point(410, 47), 120, 30);
            btnSearch.Click += (s, e) => FilterSuppliers();
            panelLeft.Controls.Add(btnSearch);

            dgvSuppliers = new DataGridView
            {
                Location = new Point(10, 88),
                Size = new Size(535, 510),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
                                    | AnchorStyles.Right | AnchorStyles.Bottom
            };
            dgvSuppliers.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvSuppliers.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvSuppliers.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvSuppliers.EnableHeadersVisualStyles = false;
            dgvSuppliers.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvSuppliers.SelectionChanged += DgvSuppliers_SelectionChanged;
            panelLeft.Controls.Add(dgvSuppliers);

            // ── PANEL RIGHT ───────────────────────────────────────────
            panelRight = new Panel
            {
                Location = new Point(580, 10),
                Size = new Size(600, 640),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
                            | AnchorStyles.Right | AnchorStyles.Bottom
            };
            this.Controls.Add(panelRight);

            panelRight.Controls.Add(new Label
            {
                Text = "THÔNG TIN NHÀ CUNG CẤP",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 10),
                Size = new Size(400, 28)
            });

            // Các trường nhập — khớp với AddParams() trong SupplierService
            int y = 48;
            txtCompanyName = AddField(panelRight, "Tên công ty (*)", ref y);
            txtShortName = AddField(panelRight, "Tên viết tắt", ref y);
            txtSupplierType = AddField(panelRight, "Loại NCC", ref y);
            txtTaxCode = AddField(panelRight, "Mã số thuế", ref y);
            txtContactPerson = AddField(panelRight, "Người liên hệ", ref y);
            txtContactPhone = AddField(panelRight, "Số điện thoại", ref y);
            txtEmail = AddField(panelRight, "Email", ref y);
            txtAddress = AddField(panelRight, "Địa chỉ", ref y);
            txtBankAccount = AddField(panelRight, "Số tài khoản", ref y);
            txtBankName = AddField(panelRight, "Tên ngân hàng", ref y);
            txtWebsite = AddField(panelRight, "Website", ref y);
            txtCert = AddField(panelRight, "Chứng chỉ", ref y);
            txtNotes = AddField(panelRight, "Ghi chú", ref y);

            chkIsActive = new CheckBox
            {
                Text = "Đang hoạt động",
                Location = new Point(150, y),
                Size = new Size(200, 25),
                Font = new Font("Segoe UI", 9),
                Checked = true
            };
            panelRight.Controls.Add(chkIsActive);
            y += 38;

            btnNew = MkBtn("+ Thêm mới", Color.FromArgb(40, 167, 69), new Point(10, y), 120, 34);
            btnSave = MkBtn("💾 Lưu", Color.FromArgb(0, 120, 212), new Point(140, y), 110, 34);
            btnDelete = MkBtn("🗑 Xóa", Color.FromArgb(220, 53, 69), new Point(260, y), 100, 34);
            btnClear = MkBtn("🔄 Làm mới", Color.FromArgb(108, 117, 125), new Point(370, y), 110, 34);

            btnNew.Click += BtnNew_Click;
            btnSave.Click += BtnSave_Click;
            btnDelete.Click += BtnDelete_Click;
            btnClear.Click += BtnClear_Click;

            panelRight.Controls.Add(btnNew);
            panelRight.Controls.Add(btnSave);
            panelRight.Controls.Add(btnDelete);
            panelRight.Controls.Add(btnClear);

            lblStatus = new Label
            {
                Location = new Point(10, y + 44),
                Size = new Size(570, 25),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.Gray
            };
            panelRight.Controls.Add(lblStatus);
        }

        // ── Helper tạo field ──────────────────────────────────────────
        private TextBox AddField(Panel panel, string label, ref int y)
        {
            panel.Controls.Add(new Label
            {
                Text = label,
                Location = new Point(10, y + 3),
                Size = new Size(135, 20),
                Font = new Font("Segoe UI", 9)
            });
            var txt = new TextBox
            {
                Location = new Point(150, y),
                Size = new Size(420, 25),
                Font = new Font("Segoe UI", 9)
            };
            panel.Controls.Add(txt);
            y += 35;
            return txt;
        }

        private Button MkBtn(string text, Color color, Point loc, int w, int h)
        {
            var btn = new Button
            {
                Text = text,
                Location = loc,
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

        // =================================================================
        // LOAD — dùng _service.GetAll() (SQL thẳng, không SP)
        // =================================================================
        private void LoadSuppliers()
        {
            try
            {
                _suppliers = _service.GetAll();
                BindGrid(_suppliers);
                lblStatus.Text = $"Tổng: {_suppliers.Count} nhà cung cấp";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải dữ liệu: " + ex.Message, "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =================================================================
        // BIND GRID
        // =================================================================
        private void BindGrid(List<Supplier> list)
        {
            dgvSuppliers.DataSource = list.ConvertAll(s => new
            {
                ID = s.Supplier_ID,
                Ten_Cong_Ty = s.Company_Name,
                Viet_Tat = s.Short_Name,
                Loai_NCC = s.Supplier_Type,
                Lien_He = s.Contact_Person,
                SDT = s.Contact_Phone,
                Email = s.Email,
                Trang_Thai = s.IsActive ? "✅ Hoạt động" : "⛔ Ngừng"
            });

            if (dgvSuppliers.Columns.Contains("ID"))
                dgvSuppliers.Columns["ID"].Visible = false;
        }

        // =================================================================
        // LỌC REAL-TIME TRÊN MEMORY
        // Không gọi _service.Search() để tránh lỗi SP
        // Tìm theo: Company_Name, Short_Name, Contact_Person,
        //           Contact_Phone, Email, Supplier_Type, Tax_Code
        // =================================================================
        private void FilterSuppliers()
        {
            string kw = txtSearch.Text.Trim().ToLower();

            if (string.IsNullOrEmpty(kw))
            {
                BindGrid(_suppliers);
                lblStatus.Text = $"Tổng: {_suppliers.Count} nhà cung cấp";
                return;
            }

            var result = _suppliers.FindAll(s =>
                (s.Company_Name ?? "").ToLower().Contains(kw) ||
                (s.Short_Name ?? "").ToLower().Contains(kw) ||
                (s.Contact_Person ?? "").ToLower().Contains(kw) ||
                (s.Contact_Phone ?? "").ToLower().Contains(kw) ||
                (s.Email ?? "").ToLower().Contains(kw) ||
                (s.Supplier_Type ?? "").ToLower().Contains(kw) ||
                (s.Tax_Code ?? "").ToLower().Contains(kw)
            );

            BindGrid(result);
            lblStatus.Text = result.Count > 0
                ? $"Tìm thấy: {result.Count} nhà cung cấp"
                : "Không tìm thấy kết quả phù hợp";
        }

        // =================================================================
        // CHỌN DÒNG → điền form
        // Mapping theo MapSupplier() trong SupplierService
        // =================================================================
        private void DgvSuppliers_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvSuppliers.SelectedRows.Count == 0) return;
            if (!dgvSuppliers.Columns.Contains("ID")) return;

            _selectedSupplierID = Convert.ToInt32(
                dgvSuppliers.SelectedRows[0].Cells["ID"].Value);

            // Tìm trong _suppliers (danh sách đầy đủ, không bị filter cắt mất)
            var s = _suppliers.Find(x => x.Supplier_ID == _selectedSupplierID);
            if (s == null) return;

            // Khớp đúng tên field với MapSupplier() và AddParams()
            txtCompanyName.Text = s.Company_Name ?? "";
            txtShortName.Text = s.Short_Name ?? "";
            txtSupplierType.Text = s.Supplier_Type ?? "";
            txtTaxCode.Text = s.Tax_Code ?? "";
            txtContactPerson.Text = s.Contact_Person ?? "";
            txtContactPhone.Text = s.Contact_Phone ?? "";
            txtEmail.Text = s.Email ?? "";
            txtAddress.Text = s.Company_Address ?? "";
            txtBankAccount.Text = s.Bank_Account ?? "";
            txtBankName.Text = s.Bank_Name ?? "";
            txtWebsite.Text = s.Website ?? "";
            txtCert.Text = s.Cert ?? "";
            txtNotes.Text = s.Notes ?? "";
            chkIsActive.Checked = s.IsActive;

            lblStatus.Text = $"Đang xem: {s.Company_Name}";
        }

        // =================================================================
        // THÊM MỚI
        // =================================================================
        private void BtnNew_Click(object sender, EventArgs e)
        {
            ClearForm();
            _selectedSupplierID = 0;
            txtCompanyName.Focus();
            lblStatus.Text = "Đang thêm nhà cung cấp mới...";
        }

        // =================================================================
        // LƯU — dùng _service.Insert / _service.Update
        // Mapping đúng theo AddParams() trong SupplierService
        // =================================================================
        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtCompanyName.Text))
            {
                MessageBox.Show("Vui lòng nhập Tên công ty!", "Thiếu thông tin",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtCompanyName.Focus();
                return;
            }

            try
            {
                // Tạo object Supplier khớp đúng với AddParams() trong service
                var s = new Supplier
                {
                    Supplier_ID = _selectedSupplierID,
                    Company_Name = txtCompanyName.Text.Trim(),    // @Company_Name
                    Short_Name = txtShortName.Text.Trim(),      // @Short_Name
                    Supplier_Type = txtSupplierType.Text.Trim(),   // @Supplier_Type
                    Tax_Code = txtTaxCode.Text.Trim(),        // @Tax_Code
                    Contact_Person = txtContactPerson.Text.Trim(),  // @Contact_Person
                    Contact_Phone = txtContactPhone.Text.Trim(),   // @Contact_Phone
                    Email = txtEmail.Text.Trim(),          // @Email
                    Company_Address = txtAddress.Text.Trim(),        // @Company_Address
                    Bank_Account = txtBankAccount.Text.Trim(),    // @Bank_Account
                    Bank_Name = txtBankName.Text.Trim(),       // @Bank_Name
                    Website = txtWebsite.Text.Trim(),        // @Website
                    Cert = txtCert.Text.Trim(),           // @Cert
                    Notes = txtNotes.Text.Trim(),          // @Notes
                    IsActive = chkIsActive.Checked            // @IsActive
                };

                if (_selectedSupplierID == 0)
                {
                    _service.Insert(s, _currentUser);
                    MessageBox.Show("✅ Thêm nhà cung cấp thành công!", "Thành công",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    _service.Update(s, _currentUser);
                    MessageBox.Show("✅ Cập nhật nhà cung cấp thành công!", "Thành công",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                LoadSuppliers();
                ClearForm();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi lưu: " + ex.Message, "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =================================================================
        // XÓA — dùng _service.Delete
        // =================================================================
        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (_selectedSupplierID == 0)
            {
                MessageBox.Show("Vui lòng chọn nhà cung cấp cần xóa!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string name = txtCompanyName.Text.Trim();
            if (MessageBox.Show(
                    $"Bạn có chắc muốn xóa nhà cung cấp '{name}'?\nHành động này không thể hoàn tác!",
                    "Xác nhận xóa",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2)
                == DialogResult.Yes)
            {
                try
                {
                    _service.Delete(_selectedSupplierID, _currentUser);
                    MessageBox.Show("✅ Xóa thành công!", "Thông báo",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadSuppliers();
                    ClearForm();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi xóa: " + ex.Message, "Lỗi",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // =================================================================
        // LÀM MỚI
        // =================================================================
        private void BtnClear_Click(object sender, EventArgs e)
        {
            ClearForm();
            LoadSuppliers();
        }

        // =================================================================
        // CLEAR FORM
        // =================================================================
        private void ClearForm()
        {
            _selectedSupplierID = 0;
            txtCompanyName.Text = "";
            txtShortName.Text = "";
            txtSupplierType.Text = "";
            txtTaxCode.Text = "";
            txtContactPerson.Text = "";
            txtContactPhone.Text = "";
            txtEmail.Text = "";
            txtAddress.Text = "";
            txtBankAccount.Text = "";
            txtBankName.Text = "";
            txtWebsite.Text = "";
            txtCert.Text = "";
            txtNotes.Text = "";
            chkIsActive.Checked = true;
            txtSearch.Text = "";
            lblStatus.Text = "";
        }
    }
}