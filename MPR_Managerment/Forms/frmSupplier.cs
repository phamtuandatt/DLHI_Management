using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using MPR_Managerment.Models;
using MPR_Managerment.Services;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace MPR_Managerment.Forms
{
    public partial class frmSupplier : Form
    {
        // ── Service (đã sửa Search dùng SQL thẳng, không dùng SP lỗi) ──
        private readonly SupplierService _service = new SupplierService();

        private List<Supplier> _suppliers = new List<Supplier>();
        private int _selectedSupplierID = 0;
        private string _currentUser = "Admin";
        private bool _suppressSelection = false;

        // ── Controls ──────────────────────────────────────────────────
        private DataGridView dgvSuppliers;
        private TextBox txtSearch;
        private TextBox txtCompanyName, txtShortName, txtSupplierType;
        private TextBox txtTaxCode, txtContactPerson, txtContactPhone;
        private TextBox txtEmail, txtAddress;
        private TextBox txtBankAccount, txtBankName;
        private TextBox txtWebsite, txtCert, txtNotes;
        private CheckBox chkIsActive;
        private Button btnSearch, btnNew, btnSave, btnDelete, btnClear, btnExport, btnIso;
        private Label lblStatus;
        private Panel panelLeft, panelRight;
        private Form TopOwner => (this.TopLevelControl as Form) ?? this;

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
            this.Resize += (s, e) => LayoutPanels();

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
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom
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

            // Dòng 1: thêm / lưu / xóa
            btnNew    = MkBtn("+ Thêm mới",    Color.FromArgb(40, 167, 69),   new Point(10,  y),      120, 32);
            btnSave   = MkBtn("💾 Lưu",         Color.FromArgb(0, 120, 212),   new Point(140, y),      110, 32);
            btnDelete = MkBtn("🗑 Xóa",         Color.FromArgb(220, 53, 69),   new Point(260, y),      100, 32);
            // Dòng 2: làm mới / xuất excel / xuất ISO
            int y2 = y + 40;
            btnClear  = MkBtn("🔄 Làm mới",    Color.FromArgb(108, 117, 125), new Point(10,  y2),     110, 32);
            btnExport = MkBtn("📊 Xuất Excel",  Color.FromArgb(33, 115, 70),   new Point(130, y2),     130, 32);
            btnIso    = MkBtn("📋 Xuất ISO",    Color.FromArgb(142, 68, 173),  new Point(270, y2),     120, 32);

            btnNew.Click    += BtnNew_Click;
            btnSave.Click   += BtnSave_Click;
            btnDelete.Click += BtnDelete_Click;
            btnExport.Click += (s, e) => ExportToExcel();
            btnIso.Click    += (s, e) => ExportISO();
            btnClear.Click  += BtnClear_Click;

            panelRight.Controls.Add(btnNew);
            panelRight.Controls.Add(btnSave);
            panelRight.Controls.Add(btnDelete);
            panelRight.Controls.Add(btnClear);
            panelRight.Controls.Add(btnExport);
            panelRight.Controls.Add(btnIso);

            lblStatus = new Label
            {
                Location = new Point(10, y2 + 40),
                Size = new Size(430, 25),
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.Gray
            };
            panelRight.Controls.Add(lblStatus);

            LayoutPanels();

            // Click vào khoảng trắng bất kỳ → bỏ chọn
            this.MouseClick       += (s, e) => DeselectGrid();
            panelLeft.MouseClick  += (s, e) => DeselectGrid();
            panelRight.MouseClick += (s, e) => DeselectGrid();
        }

        private void DeselectGrid()
        {
            if (dgvSuppliers.SelectedRows.Count > 0)
                dgvSuppliers.ClearSelection();
        }

        private void LayoutPanels()
        {
            const int margin = 10;
            const int gap = 10;
            int available = this.ClientSize.Width - margin * 2 - gap;
            if (available < 200) return;
            int leftW = (int)(available * 0.6);
            int rightW = available - leftW;
            panelLeft.Width = leftW;
            panelRight.Left = margin + leftW + gap;
            panelRight.Width = rightW;
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
                MessageBox.Show(TopOwner, "Lỗi tải dữ liệu: " + ex.Message, "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =================================================================
        // BIND GRID
        // =================================================================
        private void BindGrid(List<Supplier> list)
        {
            _suppressSelection = true;
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
            dgvSuppliers.ClearSelection();
            _suppressSelection = false;
            ClearDetailFields();
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
            if (_suppressSelection) return;
            if (dgvSuppliers.SelectedRows.Count == 0) { ClearDetailFields(); return; }
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
                MessageBox.Show(TopOwner, "Vui lòng nhập Tên công ty!", "Thiếu thông tin",
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
                    MessageBox.Show(TopOwner, "✅ Thêm nhà cung cấp thành công!", "Thành công",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    _service.Update(s, _currentUser);
                    MessageBox.Show(TopOwner, "✅ Cập nhật nhà cung cấp thành công!", "Thành công",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                LoadSuppliers();
                ClearForm();
            }
            catch (Exception ex)
            {
                MessageBox.Show(TopOwner, "Lỗi khi lưu: " + ex.Message, "Lỗi",
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
                MessageBox.Show(TopOwner, "Vui lòng chọn nhà cung cấp cần xóa!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string name = txtCompanyName.Text.Trim();
            if (MessageBox.Show(TopOwner,
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
                    MessageBox.Show(TopOwner, "✅ Xóa thành công!", "Thông báo",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadSuppliers();
                    ClearForm();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(TopOwner, "Lỗi khi xóa: " + ex.Message, "Lỗi",
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
        private void ClearDetailFields()
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
        }

        private void ClearForm()
        {
            ClearDetailFields();
            txtSearch.Text = "";
            lblStatus.Text = "";
        }
        // =================================================================
        // XUẤT EXCEL
        // =================================================================
        private void ExportToExcel()
        {
            if (_suppliers == null || _suppliers.Count == 0)
            {
                MessageBox.Show(TopOwner, "Không có dữ liệu để xuất.", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var dlg = new SaveFileDialog
            {
                Title = "Lưu file Excel",
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = $"DanhSachNCC_{DateTime.Now:yyyyMMdd_HHmm}.xlsx",
                DefaultExt = "xlsx",
                OverwritePrompt = true,
                InitialDirectory = Directory.Exists(@"D:\RÁC") ? @"D:\RÁC" : Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using var pkg = new ExcelPackage();
                var ws = pkg.Workbook.Worksheets.Add("Nhà cung cấp");

                string[] headers = { "STT", "Tên công ty", "Tên viết tắt", "Loại NCC",
                    "Mã số thuế", "Người liên hệ", "Số điện thoại", "Email",
                    "Địa chỉ", "Số tài khoản", "Ngân hàng",
                    "Website", "Giấy chứng nhận", "Ghi chú", "Hoạt động" };

                for (int c = 0; c < headers.Length; c++)
                {
                    var cell = ws.Cells[1, c + 1];
                    cell.Value = headers[c];
                    cell.Style.Font.Bold = true;
                    cell.Style.Font.Color.SetColor(Color.White);
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 120, 212));
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }

                int row = 2;
                foreach (var s in _suppliers)
                {
                    Color rowBg = (row % 2 == 0)
                        ? Color.FromArgb(240, 248, 255) : Color.White;
                    object[] values = {
                        row - 1,
                        s.Company_Name    ?? "",
                        s.Short_Name      ?? "",
                        s.Supplier_Type   ?? "",
                        s.Tax_Code        ?? "",
                        s.Contact_Person  ?? "",
                        s.Contact_Phone   ?? "",
                        s.Email           ?? "",
                        s.Company_Address ?? "",
                        s.Bank_Account    ?? "",
                        s.Bank_Name       ?? "",
                        s.Website         ?? "",
                        s.Cert            ?? "",
                        s.Notes           ?? "",
                        s.IsActive ? "Có" : "Không"
                    };
                    for (int c = 0; c < values.Length; c++)
                    {
                        var cell = ws.Cells[row, c + 1];
                        cell.Value = values[c];
                        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(rowBg);
                        cell.Style.Border.BorderAround(ExcelBorderStyle.Thin,
                            Color.FromArgb(200, 200, 200));
                    }
                    ws.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[row, 15].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    row++;
                }

                ws.Cells[ws.Dimension.Address].AutoFitColumns(8, 50);
                ws.View.FreezePanes(2, 1);
                pkg.SaveAs(new FileInfo(dlg.FileName));

                MessageBox.Show(TopOwner,
                    $"✅ Đã xuất {_suppliers.Count} nhà cung cấp.\nFile: {dlg.FileName}",
                    "Hoàn thành", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                { FileName = dlg.FileName, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show(TopOwner, "Lỗi xuất Excel: " + ex.Message, "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =================================================================
        // XUẤT ISO TEMPLATE
        // =================================================================
        // ── Helper: tìm file template ────────────────────────────────────
        private string GetTemplatePath()
        {
            string[] candidates = {
                Path.Combine(
                    Path.GetDirectoryName(
                        System.Reflection.Assembly.GetExecutingAssembly().Location) ?? "",
                    "ISO_template.xlsx"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ISO_template.xlsx")
            };
            foreach (var p in candidates)
                if (File.Exists(p)) return p;
            return null;
        }

        // ── Helper: điền thông tin NCC vào 1 worksheet ───────────────────
        private void FillSupplierSheet(OfficeOpenXml.ExcelWorksheet ws, Supplier s)
        {
            void Fill(string coord, string placeholder, string value)
            {
                var cell = ws.Cells[coord];
                if (cell.Value?.ToString()?.Contains(placeholder) == true)
                    cell.Value = cell.Value.ToString().Replace(placeholder, value ?? "");
            }
            Fill("A2", "<<Tên công ty>>", s.Company_Name ?? "");
            Fill("D3", "<<Địa chỉ>>", s.Company_Address ?? "");
            Fill("D4", "<<Số điện thoại>>", s.Contact_Phone ?? "");
            Fill("J4", "<<Email>>", s.Email ?? "");
            Fill("D5", "<<Người liên hệ>>", s.Contact_Person ?? "");
            Fill("J5", "<<Phone>>", s.Contact_Phone ?? "");
        }

        // ── Xuất ISO — chọn nhiều NCC + 2 chế độ ────────────────────────
        private void ExportISO()
        {
            if (_suppliers == null || _suppliers.Count == 0)
            {
                MessageBox.Show(TopOwner, "Không có dữ liệu nhà cung cấp.",
                    "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ── Bước 1: Chọn nhà cung cấp từ danh sách ──────────────────
            var selected = ShowSupplierPickerDialog();
            if (selected == null || selected.Count == 0) return;

            // ── Bước 2: Kiểm tra template ────────────────────────────────
            string templatePath = GetTemplatePath();
            if (templatePath == null)
            {
                MessageBox.Show(TopOwner,
                    "Không tìm thấy file ISO_template.xlsx.\n" +
                    "Vui lòng đặt file vào:\n" +
                    AppDomain.CurrentDomain.BaseDirectory,
                    "Thiếu template", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ── Bước 3: Chọn chế độ xuất ─────────────────────────────────
            var modeResult = MessageBox.Show(TopOwner,
                $"Đã chọn {selected.Count} nhà cung cấp.\n\n" +
                "Chọn cách xuất:\n" +
                "• YES  → 1 file duy nhất (mỗi NCC 1 sheet)\n" +
                "• NO   → Nhiều file riêng lẻ (mỗi NCC 1 file)",
                "Chọn chế độ xuất",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question);

            if (modeResult == DialogResult.Cancel) return;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            if (modeResult == DialogResult.Yes)
                ExportISOSingleFile(selected, templatePath);
            else
                ExportISOMultipleFiles(selected, templatePath);
        }

        // ── Xuất 1 file — mỗi NCC 1 sheet ───────────────────────────────
        private void ExportISOSingleFile(List<Supplier> suppliers, string templatePath)
        {
            using var save = new SaveFileDialog
            {
                Title = "Lưu file ISO tổng hợp",
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = $"ISO_TongHop_{DateTime.Now:yyyyMMdd_HHmm}.xlsx",
                DefaultExt = "xlsx",
                OverwritePrompt = true,
                InitialDirectory = Directory.Exists(@"D:\RÁC") ? @"D:\RÁC" : Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };
            if (save.ShowDialog() != DialogResult.OK) return;

            try
            {
                // Dùng file đầu tiên làm base
                File.Copy(templatePath, save.FileName, overwrite: true);

                using var pkg = new ExcelPackage(new FileInfo(save.FileName));

                // Sheet đầu tiên: rename theo NCC đầu tiên
                var templateWs = pkg.Workbook.Worksheets[0];
                string firstName = suppliers[0].Short_Name ?? suppliers[0].Company_Name ?? "NCC1";
                templateWs.Name = firstName.Length > 31
                    ? firstName.Substring(0, 31) : firstName;
                FillSupplierSheet(templateWs, suppliers[0]);

                // Các NCC còn lại: copy sheet template, đổi tên, điền thông tin
                for (int i = 1; i < suppliers.Count; i++)
                {
                    var s = suppliers[i];
                    string sheetName = (s.Short_Name ?? s.Company_Name ?? $"NCC{i + 1}");
                    if (sheetName.Length > 31) sheetName = sheetName.Substring(0, 31);
                    // Đảm bảo tên sheet không trùng
                    int suffix = 1;
                    string baseName = sheetName;
                    while (pkg.Workbook.Worksheets[sheetName] != null)
                        sheetName = $"{baseName}_{suffix++}";

                    pkg.Workbook.Worksheets.Copy(templateWs.Name, sheetName);
                    var newWs = pkg.Workbook.Worksheets[sheetName];

                    // Reset về template rồi điền lại
                    // (copy từ template đã điền NCC1 → cần reset placeholders trước)
                    // Thay vì reset, load lại từ template gốc cho sheet mới
                    // Cách đơn giản: copy từ file template gốc
                    using var tplPkg = new ExcelPackage(new FileInfo(templatePath));
                    var tplWs = tplPkg.Workbook.Worksheets[0];

                    // Sao chép từng cell có placeholder từ template gốc
                    foreach (var cell in tplWs.Cells)
                        if (cell.Value?.ToString()?.Contains("<<") == true)
                            newWs.Cells[cell.Address].Value = cell.Value;

                    FillSupplierSheet(newWs, s);
                }

                pkg.Save();

                MessageBox.Show(TopOwner,
                    $"✅ Đã xuất {suppliers.Count} sheet vào 1 file:\n{save.FileName}",
                    "Hoàn thành", MessageBoxButtons.OK, MessageBoxIcon.Information);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                { FileName = save.FileName, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show(TopOwner, "Lỗi xuất file: " + ex.Message, "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ── Xuất nhiều file riêng lẻ ─────────────────────────────────────
        private void ExportISOMultipleFiles(List<Supplier> suppliers, string templatePath)
        {
            using var folderDlg = new FolderBrowserDialog
            {
                Description = "Chọn thư mục lưu các file ISO",
                UseDescriptionForTitle = true
            };
            if (folderDlg.ShowDialog() != DialogResult.OK) return;

            string outDir = folderDlg.SelectedPath;
            int ok = 0, fail = 0;
            var failList = new System.Text.StringBuilder();

            foreach (var s in suppliers)
            {
                try
                {
                    string safeName = string.Concat(
                        (s.Short_Name ?? s.Company_Name ?? "NCC")
                        .Split(Path.GetInvalidFileNameChars()));
                    string outPath = Path.Combine(outDir,
                        $"ISO_{safeName}_{DateTime.Now:yyyyMMdd}.xlsx");

                    // Tránh ghi đè — thêm số nếu trùng
                    int idx = 1;
                    while (File.Exists(outPath))
                        outPath = Path.Combine(outDir,
                            $"ISO_{safeName}_{DateTime.Now:yyyyMMdd}_{idx++}.xlsx");

                    File.Copy(templatePath, outPath, overwrite: false);

                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using var pkg = new ExcelPackage(new FileInfo(outPath));
                    var ws = pkg.Workbook.Worksheets["Table 1"]
                          ?? pkg.Workbook.Worksheets[0];
                    FillSupplierSheet(ws, s);
                    pkg.Save();
                    ok++;
                }
                catch (Exception ex)
                {
                    fail++;
                    failList.AppendLine($"• {s.Company_Name}: {ex.Message}");
                }
            }

            string msg = $"✅ Đã xuất {ok} file vào:\n{outDir}";
            if (fail > 0) msg += $"\n\n⚠️ {fail} file lỗi:\n{failList}";

            { var f = TopOwner; f.BringToFront(); f.Activate(); MessageBox.Show(f, msg, "Hoàn thành", MessageBoxButtons.OK, fail > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information); }

            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            { FileName = outDir, UseShellExecute = true });
        }

        // ── Dialog chọn nhiều nhà cung cấp ───────────────────────────────
        private List<Supplier> ShowSupplierPickerDialog()
        {
            var dlg = new Form
            {
                Text = "Chọn nhà cung cấp để xuất ISO",
                Size = new Size(480, 520),
                StartPosition = FormStartPosition.CenterParent,
                MinimizeBox = false,
                MaximizeBox = false
            };

            var lblHint = new Label
            {
                Text = "Giữ Ctrl để chọn nhiều, Ctrl+A để chọn tất cả:",
                Location = new Point(12, 12),
                AutoSize = true
            };

            var clb = new CheckedListBox
            {
                Location = new Point(12, 36),
                Size = new Size(440, 380),
                CheckOnClick = true,
                Font = new Font("Segoe UI", 9)
            };

            foreach (var s in _suppliers)
                clb.Items.Add(s.Company_Name +
                    (string.IsNullOrEmpty(s.Short_Name) ? "" : $" ({s.Short_Name})"),
                    false);

            var btnAll = new Button
            {
                Text = "Chọn tất cả",
                Location = new Point(12, 426),
                Size = new Size(100, 30),
                FlatStyle = FlatStyle.Flat
            };
            btnAll.Click += (s, e) =>
            {
                for (int i = 0; i < clb.Items.Count; i++)
                    clb.SetItemChecked(i, true);
            };

            var btnNone = new Button
            {
                Text = "Bỏ chọn tất cả",
                Location = new Point(120, 426),
                Size = new Size(110, 30),
                FlatStyle = FlatStyle.Flat
            };
            btnNone.Click += (s, e) =>
            {
                for (int i = 0; i < clb.Items.Count; i++)
                    clb.SetItemChecked(i, false);
            };

            var btnOK = new Button
            {
                Text = "✅ Xuất",
                Location = new Point(254, 426),
                Size = new Size(90, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                DialogResult = DialogResult.OK
            };

            var btnCancel = new Button
            {
                Text = "Hủy",
                Location = new Point(352, 426),
                Size = new Size(80, 30),
                FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.Cancel
            };

            dlg.Controls.AddRange(new Control[]
                { lblHint, clb, btnAll, btnNone, btnOK, btnCancel });
            dlg.AcceptButton = btnOK;
            dlg.CancelButton = btnCancel;

            if (dlg.ShowDialog() != DialogResult.OK) return null;

            var result = new List<Supplier>();
            for (int i = 0; i < clb.CheckedIndices.Count; i++)
                result.Add(_suppliers[clb.CheckedIndices[i]]);

            if (result.Count == 0)
            {
                MessageBox.Show(TopOwner, "Chưa chọn nhà cung cấp nào.",
                    "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }
            return result;
        }

    }
}