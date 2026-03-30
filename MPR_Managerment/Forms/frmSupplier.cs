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
        private SupplierService _service = new SupplierService();
        private List<Supplier> _suppliers = new List<Supplier>();
        private int _selectedSupplierID = 0;
        private string _currentUser = "Admin";

        private DataGridView dgvSuppliers;
        private TextBox txtSearch, txtCompanyName, txtShortName, txtSupplierType;
        private TextBox txtCert, txtEmail, txtContactPerson, txtContactPhone;
        private TextBox txtAddress, txtBankAccount, txtBankName, txtTaxCode;
        private TextBox txtWebsite, txtNotes;
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

        private void BuildUI()
        {
            this.Text = "Quản lý Nhà Cung Cấp";
            this.Size = new Size(1200, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(245, 245, 245);

            panelLeft = new Panel { Location = new Point(10, 10), Size = new Size(580, 640), BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle };
            this.Controls.Add(panelLeft);

            panelLeft.Controls.Add(new Label { Text = "DANH SÁCH NHÀ CUNG CẤP", Font = new Font("Segoe UI", 11, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212), Location = new Point(10, 10), Size = new Size(400, 30) });

            txtSearch = new TextBox { Location = new Point(10, 50), Size = new Size(380, 30), Font = new Font("Segoe UI", 10), PlaceholderText = "Tìm kiếm theo tên..." };
            panelLeft.Controls.Add(txtSearch);
            txtSearch.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) BtnSearch_Click(null, null); };

            btnSearch = new Button { Text = "Tìm", Location = new Point(400, 49), Size = new Size(70, 32), BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Cursor = Cursors.Hand };
            btnSearch.Click += BtnSearch_Click;
            panelLeft.Controls.Add(btnSearch);

            dgvSuppliers = new DataGridView
            {
                Location = new Point(10, 95),
                Size = new Size(555, 500),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            dgvSuppliers.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvSuppliers.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvSuppliers.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvSuppliers.EnableHeadersVisualStyles = false;
            dgvSuppliers.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvSuppliers.SelectionChanged += DgvSuppliers_SelectionChanged;
            panelLeft.Controls.Add(dgvSuppliers);

            panelRight = new Panel { Location = new Point(600, 10), Size = new Size(580, 640), BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle };
            this.Controls.Add(panelRight);

            panelRight.Controls.Add(new Label { Text = "THÔNG TIN NHÀ CUNG CẤP", Font = new Font("Segoe UI", 11, FontStyle.Bold), ForeColor = Color.FromArgb(0, 120, 212), Location = new Point(10, 10), Size = new Size(400, 30) });

            int y = 50;
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

            chkIsActive = new CheckBox { Text = "Đang hoạt động", Location = new Point(150, y), Size = new Size(200, 25), Font = new Font("Segoe UI", 9), Checked = true };
            panelRight.Controls.Add(chkIsActive);
            y += 35;

            btnNew = CreateButton("+ Thêm mới", Color.FromArgb(40, 167, 69), new Point(10, y));
            btnSave = CreateButton("Lưu", Color.FromArgb(0, 120, 212), new Point(130, y));
            btnDelete = CreateButton("Xóa", Color.FromArgb(220, 53, 69), new Point(250, y));
            btnClear = CreateButton("Làm mới", Color.FromArgb(108, 117, 125), new Point(370, y));

            btnNew.Click += BtnNew_Click;
            btnSave.Click += BtnSave_Click;
            btnDelete.Click += BtnDelete_Click;
            btnClear.Click += BtnClear_Click;

            panelRight.Controls.Add(btnNew);
            panelRight.Controls.Add(btnSave);
            panelRight.Controls.Add(btnDelete);
            panelRight.Controls.Add(btnClear);

            lblStatus = new Label { Location = new Point(10, y + 45), Size = new Size(550, 25), Font = new Font("Segoe UI", 9), ForeColor = Color.Gray };
            panelRight.Controls.Add(lblStatus);
        }

        private TextBox AddField(Panel panel, string label, ref int y)
        {
            panel.Controls.Add(new Label { Text = label, Location = new Point(10, y + 3), Size = new Size(135, 20), Font = new Font("Segoe UI", 9) });
            var txt = new TextBox { Location = new Point(150, y), Size = new Size(400, 25), Font = new Font("Segoe UI", 9) };
            panel.Controls.Add(txt);
            y += 35;
            return txt;
        }

        private Button CreateButton(string text, Color color, Point location)
        {
            return new Button { Text = text, Location = location, Size = new Size(110, 35), BackColor = color, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Cursor = Cursors.Hand };
        }
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
                MessageBox.Show("Lỗi tải dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BindGrid(List<Supplier> list)
        {
            dgvSuppliers.DataSource = list.ConvertAll(s => new
            {
                ID = s.Supplier_ID,
                Ten_Cong_Ty = s.Company_Name,
                Viet_Tat = s.Short_Name,
                Loai = s.Supplier_Type,
                Lien_He = s.Contact_Person,
                SDT = s.Contact_Phone,
                Trang_Thai = s.IsActive ? "Hoat dong" : "Ngung"
            });
        }

        private void BtnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtSearch.Text))
                    LoadSuppliers();
                else
                {
                    var result = _service.Search(txtSearch.Text.Trim());
                    BindGrid(result);
                    lblStatus.Text = $"Tìm thấy: {result.Count} nhà cung cấp";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tìm kiếm: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DgvSuppliers_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvSuppliers.SelectedRows.Count == 0) return;
            var row = dgvSuppliers.SelectedRows[0];
            _selectedSupplierID = Convert.ToInt32(row.Cells["ID"].Value);
            var s = _suppliers.Find(x => x.Supplier_ID == _selectedSupplierID);
            if (s == null) return;

            txtCompanyName.Text = s.Company_Name;
            txtShortName.Text = s.Short_Name;
            txtSupplierType.Text = s.Supplier_Type;
            txtTaxCode.Text = s.Tax_Code;
            txtContactPerson.Text = s.Contact_Person;
            txtContactPhone.Text = s.Contact_Phone;
            txtEmail.Text = s.Email;
            txtAddress.Text = s.Company_Address;
            txtBankAccount.Text = s.Bank_Account;
            txtBankName.Text = s.Bank_Name;
            txtWebsite.Text = s.Website;
            txtCert.Text = s.Cert;
            txtNotes.Text = s.Notes;
            chkIsActive.Checked = s.IsActive;
        }

        private void BtnNew_Click(object sender, EventArgs e)
        {
            ClearForm();
            _selectedSupplierID = 0;
            txtCompanyName.Focus();
            lblStatus.Text = "Đang thêm mới...";
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtCompanyName.Text))
            {
                MessageBox.Show("Vui lòng nhập Tên công ty!", "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtCompanyName.Focus();
                return;
            }
            try
            {
                var s = new Supplier
                {
                    Supplier_ID = _selectedSupplierID,
                    Company_Name = txtCompanyName.Text.Trim(),
                    Short_Name = txtShortName.Text.Trim(),
                    Supplier_Type = txtSupplierType.Text.Trim(),
                    Tax_Code = txtTaxCode.Text.Trim(),
                    Contact_Person = txtContactPerson.Text.Trim(),
                    Contact_Phone = txtContactPhone.Text.Trim(),
                    Email = txtEmail.Text.Trim(),
                    Company_Address = txtAddress.Text.Trim(),
                    Bank_Account = txtBankAccount.Text.Trim(),
                    Bank_Name = txtBankName.Text.Trim(),
                    Website = txtWebsite.Text.Trim(),
                    Cert = txtCert.Text.Trim(),
                    Notes = txtNotes.Text.Trim(),
                    IsActive = chkIsActive.Checked
                };

                if (_selectedSupplierID == 0)
                    _service.Insert(s, _currentUser);
                else
                    _service.Update(s, _currentUser);

                MessageBox.Show("Lưu thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadSuppliers();
                ClearForm();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi lưu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (_selectedSupplierID == 0)
            {
                MessageBox.Show("Vui lòng chọn nhà cung cấp cần xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (MessageBox.Show("Bạn có chắc muốn xóa nhà cung cấp này?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    _service.Delete(_selectedSupplierID, _currentUser);
                    MessageBox.Show("Xóa thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadSuppliers();
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
            ClearForm();
            LoadSuppliers();
        }

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