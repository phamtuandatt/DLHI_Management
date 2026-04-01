using MPR_Managerment.Models;
using MPR_Managerment.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace MPR_Managerment.Forms
{
    public partial class frmPO : Form
    {
        private POService _service = new POService();
        private List<POHead> _poList = new List<POHead>();
        private List<PODetail> _details = new List<PODetail>();
        private int _selectedPO_ID = 0;
        private string _currentUser = "Admin";

        // Biến nhận mã PO từ form khác truyền sang
        private string _targetPoNo = "";

        private DataGridView dgvPO;
        private TextBox txtSearch;
        private Button btnSearch, btnNewPO, btnDeletePO, btnClearHeader, btnExport, btnSavePO;
        private Label lblStatus;

        private TextBox txtPONo, txtProjectName, txtWorkorderNo, txtMPRNo;
        private TextBox txtPrepared, txtReviewed, txtAgreement, txtApproved, txtNotes;
        private DateTimePicker dtpPODate;
        private ComboBox cboStatus;
        private NumericUpDown nudRevise;

        private DataGridView dgvDetails;
        private Button btnAddDetail, btnDeleteDetail;
        private Label lblTotal;
        private Panel panelTop, panelHeader, panelDetail;
        private ComboBox cboSupplier;
        private System.Data.DataTable _supplierTable;
        private bool _isSearching = false;

        // Khai báo biến tạm lưu ProjectCode khi Import MPR
        private string _projectCodeImport = string.Empty;

        public frmPO(string poNo = "")
        {
            _targetPoNo = poNo;
            InitializeComponent();
            BuildUI();
            LoadPO();
            this.Resize += FrmPO_Resize;

            // Tự động tìm và bôi xanh PO nếu được truyền mã PO vào
            if (!string.IsNullOrEmpty(_targetPoNo))
            {
                SelectPOByNo(_targetPoNo);
            }
        }

        private void SelectPOByNo(string poNo)
        {
            var targetPO = _poList.Find(p => p.PONo == poNo);
            if (targetPO != null)
            {
                txtSearch.Text = targetPO.PONo;
                BtnSearch_Click(null, null);
            }

            foreach (DataGridViewRow row in dgvPO.Rows)
            {
                if (row.Cells["PO_No"].Value?.ToString() == poNo)
                {
                    dgvPO.ClearSelection();
                    row.Selected = true;

                    if (row.Index >= 0)
                        dgvPO.FirstDisplayedScrollingRowIndex = row.Index;

                    break;
                }
            }
        }

        private void BuildUI()
        {
            this.Text = "Quản lý Đơn Đặt Hàng (PO)";
            this.Size = new Size(1300, 780);
            this.MinimumSize = new Size(1000, 650);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(245, 245, 245);

            // ===== PANEL TOP =====
            panelTop = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(1260, 210),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(panelTop);

            panelTop.Controls.Add(new Label
            {
                Text = "DANH SÁCH ĐƠN ĐẶT HÀNG (PO)",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 10),
                Size = new Size(450, 30)
            });

            txtSearch = new TextBox { Location = new Point(10, 48), Size = new Size(300, 28), Font = new Font("Segoe UI", 10), PlaceholderText = "Tìm theo PO No, MPR No, tên dự án..." };
            panelTop.Controls.Add(txtSearch);
            txtSearch.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) BtnSearch_Click(null, null); };

            btnSearch = CreateButton("Tìm", Color.FromArgb(0, 120, 212), new Point(320, 47), 70, 30);
            btnNewPO = CreateButton("+ Tạo PO", Color.FromArgb(40, 167, 69), new Point(400, 47), 100, 30);
            btnDeletePO = CreateButton("Xóa PO", Color.FromArgb(220, 53, 69), new Point(510, 47), 90, 30);

            btnSearch.Click += BtnSearch_Click;
            btnNewPO.Click += BtnNewPO_Click;
            btnDeletePO.Click += BtnDeletePO_Click;

            panelTop.Controls.Add(btnSearch);
            panelTop.Controls.Add(btnNewPO);
            panelTop.Controls.Add(btnDeletePO);

            lblStatus = new Label { Location = new Point(620, 52), Size = new Size(500, 25), Font = new Font("Segoe UI", 9), ForeColor = Color.Gray };
            panelTop.Controls.Add(lblStatus);

            dgvPO = new DataGridView
            {
                Location = new Point(10, 85),
                Size = new Size(1235, 115),
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
            dgvPO.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvPO.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPO.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvPO.EnableHeadersVisualStyles = false;
            dgvPO.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgvPO.SelectionChanged += DgvPO_SelectionChanged;
            panelTop.Controls.Add(dgvPO);

            // ===== PANEL HEADER =====
            panelHeader = new Panel
            {
                Location = new Point(10, 230),
                Size = new Size(1260, 245),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(panelHeader);

            panelHeader.Controls.Add(new Label
            {
                Text = "THÔNG TIN ĐƠN ĐẶT HÀNG",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 8),
                Size = new Size(350, 25)
            });

            // Row 1
            int y = 38;
            AddLabel(panelHeader, "PO No (*):", 10, y);
            txtPONo = AddTxt(panelHeader, 90, y, 140);

            AddLabel(panelHeader, "Tên dự án:", 240, y);
            txtProjectName = AddTxt(panelHeader, 320, y, 200);

            AddLabel(panelHeader, "Workorder:", 530, y);
            txtWorkorderNo = AddTxt(panelHeader, 610, y, 160);

            AddLabel(panelHeader, "MPR No:", 780, y);
            txtMPRNo = AddTxt(panelHeader, 865, y, 250);

            // Row 2 - Supplier
            y += 38;
            AddLabel(panelHeader, "Nhà cung cấp:", 10, y);
            cboSupplier = new ComboBox
            {
                Location = new Point(115, y),
                Size = new Size(280, 25),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteMode = AutoCompleteMode.None
            };
            panelHeader.Controls.Add(cboSupplier);
            cboSupplier.Validating += CboSupplier_Validating;
            cboSupplier.SelectedIndexChanged += CboSupplier_SelectedIndexChanged;
            cboSupplier.TextChanged += CboSupplier_TextChanged;
            cboSupplier.KeyDown += CboSupplier_KeyDown;
            LoadSupplierCombo();

            // Row 2 continued
            y += 38;
            AddLabel(panelHeader, "Ngày PO:", 10, y);
            dtpPODate = new DateTimePicker { Location = new Point(90, y), Size = new Size(140, 25), Font = new Font("Segoe UI", 9), Format = DateTimePickerFormat.Short };
            panelHeader.Controls.Add(dtpPODate);

            AddLabel(panelHeader, "Trạng thái:", 240, y);
            cboStatus = new ComboBox { Location = new Point(320, y), Size = new Size(130, 25), Font = new Font("Segoe UI", 9), DropDownStyle = ComboBoxStyle.DropDownList };
            cboStatus.Items.AddRange(new[] { "Draft", "Pending", "Approved", "In Progress", "Completed", "Cancelled" });
            cboStatus.SelectedIndex = 0;
            panelHeader.Controls.Add(cboStatus);

            AddLabelCus(panelHeader, "Revise:", 480, y, 40, 20);
            nudRevise = new NumericUpDown { Location = new Point(525, y), Size = new Size(60, 25), Font = new Font("Segoe UI", 9), Minimum = 0, Maximum = 99 };
            nudRevise.BringToFront();
            panelHeader.Controls.Add(nudRevise);

            AddLabel(panelHeader, "Ghi chú:", 610, y);
            txtNotes = AddTxt(panelHeader, 680, y, 380); txtNotes.BringToFront();
            txtNotes.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            // Row 3
            y += 38;
            AddLabel(panelHeader, "Prepared:", 10, y);
            txtPrepared = AddTxt(panelHeader, 90, y, 140);

            AddLabel(panelHeader, "Reviewed:", 240, y);
            txtReviewed = AddTxt(panelHeader, 320, y, 140);

            AddLabel(panelHeader, "Agreement:", 470, y);
            txtAgreement = AddTxt(panelHeader, 555, y, 140);

            AddLabel(panelHeader, "Approved:", 705, y);
            txtApproved = AddTxt(panelHeader, 785, y, 140);

            // Buttons
            y += 45;
            btnSavePO = CreateButton("💾 Lưu Toàn Bộ PO", Color.FromArgb(0, 120, 212), new Point(10, y), 150, 32);
            btnSavePO.Click += BtnSavePO_Click;
            panelHeader.Controls.Add(btnSavePO);

            btnClearHeader = CreateButton("Làm mới", Color.FromArgb(108, 117, 125), new Point(170, y), 100, 32);
            btnClearHeader.Click += BtnClearHeader_Click;
            panelHeader.Controls.Add(btnClearHeader);

            var btnImportMPR = CreateButton("Import MPR", Color.FromArgb(255, 140, 0), new Point(280, y), 120, 32);
            btnImportMPR.Click += BtnImportMPR_Click;
            panelHeader.Controls.Add(btnImportMPR);

            var btnHistory = CreateButton("Revise History", Color.FromArgb(102, 51, 153), new Point(410, y), 130, 32);
            btnHistory.Click += (s, e) =>
            {
                if (string.IsNullOrEmpty(txtPONo.Text))
                {
                    MessageBox.Show("Vui lòng chọn một PO trước!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                new frmReviseHistory(txtPONo.Text).ShowDialog();
            };
            panelHeader.Controls.Add(btnHistory);

            // ===== PANEL DETAIL =====
            panelDetail = new Panel
            {
                Location = new Point(10, 500),
                Size = new Size(1260, 285),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            this.Controls.Add(panelDetail);

            panelDetail.Controls.Add(new Label
            {
                Text = "CHI TIẾT ĐƠN HÀNG",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212),
                Location = new Point(10, 8),
                Size = new Size(300, 25)
            });

            btnAddDetail = CreateButton("+ Thêm dòng", Color.FromArgb(40, 167, 69), new Point(10, 38), 120, 30);
            btnDeleteDetail = CreateButton("Xóa dòng", Color.FromArgb(220, 53, 69), new Point(140, 38), 100, 30);
            btnExport = CreateButton("📄 Xuất Excel", Color.FromArgb(0, 150, 100), new Point(250, 38), 130, 30);

            btnAddDetail.Click += BtnAddDetail_Click;
            btnDeleteDetail.Click += BtnDeleteDetail_Click;
            btnExport.Click += BtnExport_Click;

            panelDetail.Controls.Add(btnAddDetail);
            panelDetail.Controls.Add(btnDeleteDetail);
            panelDetail.Controls.Add(btnExport);

            lblTotal = new Label
            {
                Location = new Point(390, 45),
                Size = new Size(500, 22),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 120, 212)
            };
            panelDetail.Controls.Add(lblTotal);

            dgvDetails = new DataGridView
            {
                Location = new Point(10, 75),
                Size = new Size(1235, 195),
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            dgvDetails.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgvDetails.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvDetails.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvDetails.EnableHeadersVisualStyles = false;
            dgvDetails.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);

            dgvDetails.CellEndEdit += DgvDetails_CellEndEdit;
            dgvDetails.KeyDown += DgvDetails_KeyDown; // SỰ KIỆN COPY PASTE TỪ EXCEL

            BuildDetailColumns();
            panelDetail.Controls.Add(dgvDetails);
        }

        public static void AutoCompleteComboboxValidating(ComboBox sender, CancelEventArgs e)
        {
            var cb = sender as ComboBox;
            string typedText = cb.Text?.Trim();

            if (string.IsNullOrEmpty(typedText))
            {
                cb.SelectedIndex = 0;
                return;
            }

            bool matched = false;
            string displayMember = cb.DisplayMember;

            foreach (var item in cb.Items)
            {
                if (item is DataRowView drv)
                {
                    string value = drv[displayMember]?.ToString();
                    if (value != null && value.Equals(typedText, StringComparison.OrdinalIgnoreCase))
                    {
                        cb.SelectedItem = item;
                        matched = true;
                        break;
                    }
                }
            }

            if (!matched)
            {
                cb.SelectedIndex = 0;
            }
        }

        private void LoadSupplierCombo()
        {
            try
            {
                _supplierTable = new SupplierService().GetForCombo();
                BindSupplierCombo(_supplierTable);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải nhà cung cấp: " + ex.Message, "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BindSupplierCombo(System.Data.DataTable dt)
        {
            _isSearching = true;
            string currentText = cboSupplier.Text;
            cboSupplier.DataSource = null;
            cboSupplier.DataSource = dt;
            cboSupplier.DisplayMember = "Name";
            cboSupplier.ValueMember = "ID";
            cboSupplier.Text = currentText;
            _isSearching = false;
        }

        private void CboSupplier_TextChanged(object sender, EventArgs e)
        {
            if (_isSearching) return;
            string keyword = cboSupplier.Text.Trim();

            if (string.IsNullOrEmpty(keyword))
            {
                BindSupplierCombo(_supplierTable);
                cboSupplier.DroppedDown = false;
                return;
            }

            string keywordNorm = RemoveDiacritics(keyword).ToLower();
            var filtered = new System.Data.DataTable();
            filtered.Columns.Add("ID", typeof(int));
            filtered.Columns.Add("Name", typeof(string));

            foreach (System.Data.DataRow row in _supplierTable.Rows)
            {
                string name = row["Name"].ToString();
                string nameNorm = RemoveDiacritics(name).ToLower();

                if (nameNorm.Contains(keywordNorm) ||
                    name.ToLower().Contains(keyword.ToLower()))
                {
                    filtered.Rows.Add(row["ID"], row["Name"]);
                }
            }

            if (filtered.Rows.Count == 0)
            {
                var empty = new System.Data.DataTable();
                empty.Columns.Add("ID", typeof(int));
                empty.Columns.Add("Name", typeof(string));
                empty.Rows.Add(0, "-- Không tìm thấy --");
                BindSupplierCombo(empty);
            }
            else
            {
                BindSupplierCombo(filtered);
            }

            _isSearching = true;
            cboSupplier.Text = keyword;
            cboSupplier.SelectionStart = keyword.Length;
            cboSupplier.DroppedDown = true;
            _isSearching = false;
        }

        private string RemoveDiacritics(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            try
            {
                string normalized = text.Normalize(System.Text.NormalizationForm.FormD);
                var sb = new System.Text.StringBuilder();
                foreach (char c in normalized)
                {
                    var category = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c);
                    if (category != System.Globalization.UnicodeCategory.NonSpacingMark)
                        sb.Append(c);
                }
                return sb.ToString().Normalize(System.Text.NormalizationForm.FormC);
            }
            catch { return text; }
        }

        private void CboSupplier_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                if (cboSupplier.DroppedDown && cboSupplier.Items.Count > 0)
                {
                    if (cboSupplier.SelectedIndex >= 0)
                    {
                        int selectedId = Convert.ToInt32(cboSupplier.SelectedValue ?? 0);
                        if (selectedId > 0)
                        {
                            string selectedName = cboSupplier.Text;
                            cboSupplier.DroppedDown = false;

                            _isSearching = true;
                            BindSupplierCombo(_supplierTable);
                            cboSupplier.SelectedValue = selectedId;
                            _isSearching = false;

                            cboSupplier.BackColor = Color.White;
                            e.SuppressKeyPress = true;
                            e.Handled = true;
                            return;
                        }
                    }

                    string keyword = cboSupplier.Text.Trim();
                    string keywordNorm = RemoveDiacritics(keyword).ToLower();
                    int matchId = 0;

                    foreach (System.Data.DataRowView drv in cboSupplier.Items)
                    {
                        string name = drv["Name"].ToString();
                        string nameNorm = RemoveDiacritics(name).ToLower();
                        int id = Convert.ToInt32(drv["ID"]);

                        if (id > 0 && (nameNorm.Contains(keywordNorm) ||
                            name.ToLower().Contains(keyword.ToLower())))
                        {
                            matchId = id;
                            break;
                        }
                    }

                    if (matchId > 0)
                    {
                        cboSupplier.DroppedDown = false;
                        _isSearching = true;
                        BindSupplierCombo(_supplierTable);
                        cboSupplier.SelectedValue = matchId;
                        _isSearching = false;

                        cboSupplier.BackColor = Color.White;
                    }
                    else
                    {
                        cboSupplier.BackColor = Color.FromArgb(255, 230, 230);
                    }
                }

                e.SuppressKeyPress = true;
                e.Handled = true;
            }

            if (e.KeyCode == Keys.Escape)
            {
                _isSearching = true;
                BindSupplierCombo(_supplierTable);
                cboSupplier.Text = "";
                cboSupplier.DroppedDown = false;
                cboSupplier.BackColor = Color.White;
                _isSearching = false;
            }
        }

        private void CboSupplier_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            AutoCompleteComboboxValidating(sender as ComboBox, e);
        }

        private void CboSupplier_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_isSearching) return;
            if (cboSupplier.SelectedValue == null) return;
            int supplierId = Convert.ToInt32(cboSupplier.SelectedValue);

            if (supplierId == 0)
            {
                cboSupplier.BackColor = Color.White;
                return;
            }

            try
            {
                cboSupplier.BackColor = Color.White;
                _isSearching = true;
                BindSupplierCombo(_supplierTable);
                cboSupplier.SelectedValue = supplierId;
                _isSearching = false;
            }
            catch { }
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (_selectedPO_ID == 0)
            {
                MessageBox.Show("Vui lòng chọn PO cần xuất!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var po = _poList.Find(p => p.PO_ID == _selectedPO_ID);
                var details = _service.GetDetails(_selectedPO_ID);
                if (po == null) return;

                var suppliers = new SupplierService().GetAll();
                var supplier = suppliers.Find(s => s.Supplier_ID == Convert.ToInt32(cboSupplier.SelectedValue?.ToString() ?? "0"));

                var projects = new ProjectService().GetAll();
                var project = projects.Find(p => p.WorkorderNo == po.WorkorderNo);

                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "po_template.xlsx");
                if (!File.Exists(templatePath))
                {
                    MessageBox.Show($"Lỗi: Không tìm thấy file template!\nĐường dẫn dự kiến: {templatePath}",
                                    "Lỗi hệ thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var saveDialog = new SaveFileDialog
                {
                    Title = "Lưu file PO",
                    Filter = "Excel Files|*.xlsx",
                    FileName = $"PO_{po.PONo}_{DateTime.Now:ddMMyyyy}",
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                };
                if (saveDialog.ShowDialog() != DialogResult.OK) return;

                File.Copy(templatePath, saveDialog.FileName, true);
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(saveDialog.FileName)))
                {
                    var ws = package.Workbook.Worksheets[0];

                    // ===== HEADER =====
                    ReplaceCell(ws, "<<PROJECT_NAME>>", project?.ProjectName ?? po.Project_Name ?? "");
                    ReplaceCell(ws, "<<WO-NO>>", po.WorkorderNo ?? "");
                    ReplaceCell(ws, "<<REV.NUM>>", po.Revise.ToString() ?? "0");
                    ReplaceCell(ws, "<<DATE>>", po.PO_Date.HasValue ? po.PO_Date.Value.ToString("dd/MM/yyyy") : DateTime.Today.ToString("dd/MM/yyyy"));
                    ReplaceCell(ws, "<<MPR-NO>>", po.MPR_No ?? "");
                    ReplaceCell(ws, "<<PO-NO>>", po.PONo ?? "");

                    // Supplier info
                    string supplierInfo = supplier != null
                        ? $"{supplier.Company_Name}\nCert: {supplier.Cert ?? ""}\nEmail: {supplier.Email}"
                        : "";
                    ReplaceCell(ws, "<<SUPPLIER-INFO>>", supplierInfo);

                    // ===== DETAIL ROWS =====
                    int startRow = 8;
                    int detailCount = details.Count;

                    if (detailCount > 1)
                        ws.InsertRow(startRow + 1, detailCount - 1);
                    for (int i = 0; i < detailCount; i++)
                    {
                        var d = details[i];
                        int row = startRow + i;

                        ws.Cells[row, 1].Value = i + 1;
                        ws.Cells[row, 2].Value = d.Item_Name ?? "";
                        ws.Cells[row, 3].Value = d.Material ?? "";
                        ws.Cells[row, 4].Value = d.Asize;
                        ws.Cells[row, 5].Value = d.Bsize;
                        ws.Cells[row, 6].Value = d.Csize;
                        ws.Cells[row, 7].Value = d.Qty_Per_Sheet;
                        ws.Cells[row, 8].Value = d.UNIT ?? "";
                        ws.Cells[row, 9].Value = d.Weight_kg;
                        ws.Cells[row, 10].Value = d.MPSNo ?? "";
                        ws.Cells[row, 11].Value = d.RequestDay;
                        ws.Cells[row, 12].Value = "Kho DLHI";
                        ws.Cells[row, 13].Value = d.Price;
                        ws.Cells[row, 14].Value = d.Amount;
                        ws.Cells[row, 16].Value = d.Remarks ?? "";

                        if (i > 0)
                        {
                            for (int col = 1; col <= 16; col++)
                            {
                                ws.Cells[row, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                ws.Cells[row, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                ws.Cells[row, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                ws.Cells[row, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                ws.Cells[row, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                ws.Cells[row, col].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                ws.Cells[row, col].Style.Font.Name = "Arial";
                                ws.Cells[row, col].Style.Font.Size = 9;
                            }
                            ws.Cells[row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        }
                    }

                    // ===== SUB-TOTAL & VAT =====
                    int subTotalRow = startRow + detailCount;
                    int vatRow = subTotalRow + 1;

                    ws.Cells[subTotalRow, 3].Value = "SUB-TOTAL";
                    ws.Cells[subTotalRow, 9].Value = details.Sum(d => (double)d.Weight_kg);
                    ws.Cells[subTotalRow, 14].Formula = $"=SUM(N{startRow}:N{startRow + detailCount - 1})";

                    ws.Cells[vatRow, 3].Value = "Final Price Requested (Included 10% VAT)";
                    ws.Cells[vatRow, 14].Formula = $"=N{subTotalRow}*1.1";

                    for (int r = 1; r <= ws.Dimension.End.Row; r++)
                        for (int c = 1; c <= ws.Dimension.End.Column; c++)
                            if (ws.Cells[r, c].Value?.ToString() == "<<DATE>>")
                                ws.Cells[r, c].Value = DateTime.Today.ToString("dd/MM/yyyy");
                    package.Save();
                }

                var result = MessageBox.Show(
                    $"✅ Xuất Excel thành công!\nFile: {saveDialog.FileName}\n\nBạn có muốn mở file không?",
                    "Thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (result == DialogResult.Yes)
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = saveDialog.FileName,
                        UseShellExecute = true
                    });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xuất Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ReplaceCell(OfficeOpenXml.ExcelWorksheet ws, string placeholder, string value)
        {
            for (int r = 1; r <= ws.Dimension.End.Row; r++)
                for (int c = 1; c <= ws.Dimension.End.Column; c++)
                    if (ws.Cells[r, c].Value?.ToString() == placeholder)
                        ws.Cells[r, c].Value = value;
        }

        private void BuildDetailColumns()
        {
            dgvDetails.Columns.Clear();
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "DeliveryLocation", HeaderText = "Nơi giao", Width = 120 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Item_No", HeaderText = "STT", Width = 45, ReadOnly = true });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Item_Name", HeaderText = "Tên hàng", Width = 180 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Material", HeaderText = "Vật liệu", Width = 90 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Asize", HeaderText = "A(mm)", Width = 65 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Bsize", HeaderText = "B(mm)", Width = 65 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Csize", HeaderText = "C(mm)", Width = 65 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Qty", HeaderText = "SL", Width = 55 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "UNIT", HeaderText = "ĐVT", Width = 55 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Weight", HeaderText = "KG", Width = 60 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Price", HeaderText = "Đơn giá", Width = 90 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "VAT", HeaderText = "VAT(%)", Width = 65 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Amount", HeaderText = "Thành tiền", Width = 100, ReadOnly = true });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Received", HeaderText = "Đã nhận", Width = 70 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "MPSNo", HeaderText = "MPS No", Width = 90 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "Remarks", HeaderText = "Ghi chú", FillWeight = 100 });
            dgvDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "PO_Detail_ID", HeaderText = "PO_ID", FillWeight = 100, Visible = false });
        }

        private void AddLabel(Panel p, string text, int x, int y)
        {
            p.Controls.Add(new Label { Text = text, Location = new Point(x, y + 3), Size = new Size(80, 20), Font = new Font("Segoe UI", 9) });
        }

        private void AddLabelCus(Panel p, string text, int x, int y, int w, int h)
        {
            p.Controls.Add(new Label { Text = text, Location = new Point(x, y + 3), Size = new Size(w, h), Font = new Font("Segoe UI", 9) });
        }

        private TextBox AddTxt(Panel p, int x, int y, int width)
        {
            var txt = new TextBox { Location = new Point(x, y), Size = new Size(width, 25), Font = new Font("Segoe UI", 9) };
            p.Controls.Add(txt);
            return txt;
        }

        private Button CreateButton(string text, Color color, Point loc, int w, int h)
        {
            return new Button { Text = text, Location = loc, Size = new Size(w, h), BackColor = color, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold), Cursor = Cursors.Hand };
        }

        // ===== RESIZE =====
        private void FrmPO_Resize(object sender, EventArgs e)
        {
            int w = this.ClientSize.Width - 20;
            int h = this.ClientSize.Height;

            panelTop.Width = w;
            panelHeader.Width = w;
            panelDetail.Width = w;

            panelHeader.Top = panelTop.Bottom + 10;
            panelDetail.Top = panelHeader.Bottom + 10;
            panelDetail.Height = h - panelDetail.Top - 10;

            dgvPO.Width = panelTop.Width - 20;
            dgvDetails.Width = panelDetail.Width - 20;
            dgvDetails.Height = panelDetail.Height - 80;

            txtNotes.Width = panelHeader.Width - txtNotes.Left - 20;
        }

        // ===== LOAD DỮ LIỆU =====
        private void LoadPO()
        {
            try
            {
                _poList = _service.GetAll();
                BindPOGrid(_poList);
                lblStatus.Text = $"Tổng: {_poList.Count} đơn PO";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BindPOGrid(List<POHead> list)
        {
            dgvPO.DataSource = list.ConvertAll(h => new
            {
                ID = h.PO_ID,
                PO_No = h.PONo,
                Du_An = h.Project_Name,
                MPR_No = h.MPR_No,
                Workorder = h.WorkorderNo,
                Ngay_PO = h.PO_Date.HasValue ? h.PO_Date.Value.ToString("dd/MM/yyyy") : "",
                Trang_Thai = h.Status,
                Tong_Tien = h.Total_Amount.ToString("N0"),
                Revise = h.Revise,
                Ngay_Tao = h.Created_Date.HasValue ? h.Created_Date.Value.ToString("dd/MM/yyyy") : ""
            });
        }

        private void LoadDetails(int poId)
        {
            try
            {
                _details = new POService().GetDetails(poId);
                dgvDetails.Rows.Clear();

                foreach (var d in _details)
                {
                    int idx = dgvDetails.Rows.Add();
                    var row = dgvDetails.Rows[idx];

                    row.Cells["PO_Detail_ID"].Value = d.PO_Detail_ID;
                    row.Cells["Item_No"].Value = d.Item_No;
                    row.Cells["Item_Name"].Value = d.Item_Name;
                    row.Cells["Material"].Value = d.Material;
                    row.Cells["Asize"].Value = d.Asize;
                    row.Cells["Bsize"].Value = d.Bsize;
                    row.Cells["Csize"].Value = d.Csize;
                    row.Cells["Qty"].Value = d.Qty_Per_Sheet;
                    row.Cells["UNIT"].Value = d.UNIT;
                    row.Cells["Weight"].Value = d.Weight_kg;
                    row.Cells["Price"].Value = d.Price;
                    row.Cells["VAT"].Value = d.VAT;
                    row.Cells["Amount"].Value = d.Amount;
                    row.Cells["Received"].Value = d.Received;
                    row.Cells["MPSNo"].Value = d.MPSNo;
                    row.Cells["DeliveryLocation"].Value = d.DeliveryLocation;
                    row.Cells["Remarks"].Value = d.Remarks;
                }

                UpdateTotal();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tải chi tiết PO: " + ex.Message, "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateTotal()
        {
            decimal total = 0;
            foreach (DataGridViewRow row in dgvDetails.Rows)
            {
                decimal.TryParse(row.Cells["Amount"].Value?.ToString(), out decimal amt);
                total += amt;
            }
            lblTotal.Text = $"Tổng tiền: {total:N0} VND";
        }

        // ===== SỰ KIỆN =====

        // Bắt sự kiện bấm Ctrl + V
        private void DgvDetails_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V)
            {
                PasteFromExcel();
                e.Handled = true;
            }
        }

        // Hàm xử lý cắt dán từ Clipboard
        private void PasteFromExcel()
        {
            try
            {
                string copiedData = Clipboard.GetText();
                if (string.IsNullOrEmpty(copiedData)) return;

                string[] lines = copiedData.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                if (lines.Length == 0) return;

                int startRow = 0;
                int startCol = 0;

                if (dgvDetails.CurrentCell != null)
                {
                    startRow = dgvDetails.CurrentCell.RowIndex;
                    startCol = dgvDetails.CurrentCell.ColumnIndex;
                }

                foreach (string line in lines)
                {
                    string[] cells = line.Split('\t');

                    if (startRow >= dgvDetails.Rows.Count)
                    {
                        int nextItem = dgvDetails.Rows.Count + 1;
                        int newIdx = dgvDetails.Rows.Add();
                        var r = dgvDetails.Rows[newIdx];

                        r.Cells["Item_No"].Value = nextItem;
                        r.Cells["PO_Detail_ID"].Value = 0;
                        r.Cells["UNIT"].Value = "PCS";
                        r.Cells["Qty"].Value = 0;
                        r.Cells["Weight"].Value = 0;
                        r.Cells["Price"].Value = 0;
                        r.Cells["VAT"].Value = 0;
                        r.Cells["Amount"].Value = 0;
                        r.Cells["Received"].Value = 0;
                    }

                    int colIndex = startCol;
                    for (int i = 0; i < cells.Length; i++)
                    {
                        if (colIndex < dgvDetails.Columns.Count)
                        {
                            if (dgvDetails.Columns[colIndex].Visible && !dgvDetails.Columns[colIndex].ReadOnly)
                            {
                                dgvDetails.Rows[startRow].Cells[colIndex].Value = cells[i].Trim();
                            }
                            colIndex++;
                        }
                    }

                    RecalculateAmount(startRow);
                    startRow++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi dán dữ liệu: " + ex.Message, "Lỗi Paste", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RecalculateAmount(int rowIndex)
        {
            var row = dgvDetails.Rows[rowIndex];
            decimal.TryParse(row.Cells["Qty"].Value?.ToString(), out decimal qty);
            decimal.TryParse(row.Cells["Price"].Value?.ToString(), out decimal price);
            decimal.TryParse(row.Cells["VAT"].Value?.ToString(), out decimal vat);

            decimal amount = qty * price * (1 + vat / 100);
            row.Cells["Amount"].Value = Math.Round(amount, 0);

            UpdateTotal();
        }

        private void DgvDetails_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            RecalculateAmount(e.RowIndex);
        }

        private void BtnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtSearch.Text))
                    LoadPO();
                else
                {
                    var result = _service.Search(txtSearch.Text.Trim());
                    BindPOGrid(result);
                    lblStatus.Text = $"Tìm thấy: {result.Count} đơn PO";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tìm kiếm: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DgvPO_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvPO.SelectedRows.Count == 0) return;
            var row = dgvPO.SelectedRows[0];
            _selectedPO_ID = Convert.ToInt32(row.Cells["ID"].Value);
            var h = _poList.Find(x => x.PO_ID == _selectedPO_ID);
            if (h == null) return;

            txtPONo.Text = h.PONo;
            txtProjectName.Text = h.Project_Name;
            txtWorkorderNo.Text = h.WorkorderNo;
            txtMPRNo.Text = h.MPR_No;
            txtPrepared.Text = h.Prepared;
            txtReviewed.Text = h.Reviewed;
            txtAgreement.Text = h.Agreement;
            txtApproved.Text = h.Approved;
            txtNotes.Text = h.Notes;
            nudRevise.Value = h.Revise;

            if (h.PO_Date.HasValue) dtpPODate.Value = h.PO_Date.Value;
            var idx = cboStatus.Items.IndexOf(h.Status);
            cboStatus.SelectedIndex = idx >= 0 ? idx : 0;

            LoadDetails(_selectedPO_ID);
        }

        private void BtnNewPO_Click(object sender, EventArgs e)
        {
            ClearHeader();
            _selectedPO_ID = 0;
            dgvDetails.Rows.Clear();
            lblTotal.Text = "";
            txtPONo.Focus();
            lblStatus.Text = "Đang tạo đơn PO mới...";
        }

        private void BtnSavePO_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtPONo.Text))
            {
                MessageBox.Show("Vui lòng nhập PO No!", "Thiếu thông tin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtPONo.Focus();
                return;
            }

            if (dgvDetails.Rows.Count == 0 && MessageBox.Show("Đơn hàng này chưa có chi tiết vật tư nào.\nBạn có chắc chắn muốn lưu chỉ với Header không?", "Cảnh báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
            {
                return;
            }

            try
            {
                // ==========================================
                // PHẦN 1: LƯU HEADER
                // ==========================================
                string basePONo = txtPONo.Text.Trim();
                int revIdx = basePONo.LastIndexOf("_Rev");
                if (revIdx > 0) basePONo = basePONo.Substring(0, revIdx);

                string finalPONo = basePONo;
                bool isBaseDuplicate = _poList.Exists(p => p.PONo == basePONo && p.PO_ID != _selectedPO_ID);

                if (isBaseDuplicate || nudRevise.Value > 0)
                {
                    if (nudRevise.Value == 0)
                    {
                        MessageBox.Show("Số PO này đã tồn tại!\nVui lòng tăng số Revise để tạo bản sửa đổi.", "Trùng lặp", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        nudRevise.Focus();
                        return;
                    }
                    finalPONo = $"{basePONo}_Rev{nudRevise.Value}";
                    if (_poList.Exists(p => p.PONo == finalPONo && p.PO_ID != _selectedPO_ID))
                    {
                        MessageBox.Show($"Bản '{finalPONo}' cũng đã tồn tại!\nVui lòng tăng Revise lên mức cao hơn.", "Trùng lặp", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        nudRevise.Focus();
                        return;
                    }
                }

                var h = new POHead
                {
                    PO_ID = _selectedPO_ID,
                    PONo = finalPONo,
                    Project_Name = txtProjectName.Text.Trim(),
                    WorkorderNo = txtWorkorderNo.Text.Trim(),
                    MPR_No = txtMPRNo.Text.Trim(),
                    Prepared = txtPrepared.Text.Trim(),
                    Reviewed = txtReviewed.Text.Trim(),
                    Agreement = txtAgreement.Text.Trim(),
                    Approved = txtApproved.Text.Trim(),
                    Notes = txtNotes.Text.Trim(),
                    PO_Date = dtpPODate.Value,
                    Status = cboStatus.SelectedItem?.ToString() ?? "Draft",
                    Revise = (int)nudRevise.Value,

                    Supplier_ID = Convert.ToInt32(cboSupplier.SelectedValue ?? 0),
                    ProjectCode = _projectCodeImport
                };

                if (_selectedPO_ID == 0)
                {
                    _selectedPO_ID = _service.InsertHead(h, _currentUser);
                }
                else
                {
                    _service.UpdateHead(h, _currentUser);
                }
                txtPONo.Text = finalPONo;

                // ==========================================
                // PHẦN 2: LƯU CHI TIẾT ĐƠN HÀNG
                // ==========================================
                var oldDetails = _service.GetDetails(_selectedPO_ID);
                foreach (var d in oldDetails)
                {
                    _service.DeleteDetail(d.PO_Detail_ID);
                }

                int itemNo = 1;
                foreach (DataGridViewRow row in dgvDetails.Rows)
                {
                    var detail = new PODetail
                    {
                        Item_No = itemNo++,
                        Item_Name = row.Cells["Item_Name"].Value?.ToString() ?? "",
                        Material = row.Cells["Material"].Value?.ToString() ?? "",
                        Asize = row.Cells["Asize"].Value?.ToString() ?? "",
                        Bsize = row.Cells["Bsize"].Value?.ToString() ?? "",
                        Csize = row.Cells["Csize"].Value?.ToString() ?? "",
                        Qty_Per_Sheet = int.TryParse(row.Cells["Qty"].Value?.ToString(), out int q) ? q : 0,
                        UNIT = row.Cells["UNIT"].Value?.ToString() ?? "",
                        Weight_kg = decimal.TryParse(row.Cells["Weight"].Value?.ToString(), out decimal wk) ? wk : 0,
                        Price = decimal.TryParse(row.Cells["Price"].Value?.ToString(), out decimal p) ? p : 0,
                        VAT = decimal.TryParse(row.Cells["VAT"].Value?.ToString(), out decimal vat) ? vat : 0,
                        Received = int.TryParse(row.Cells["Received"].Value?.ToString(), out int rec) ? rec : 0,
                        MPSNo = row.Cells["MPSNo"].Value?.ToString() ?? "",
                        DeliveryLocation = row.Cells["DeliveryLocation"].Value?.ToString() ?? "",
                        Remarks = row.Cells["Remarks"].Value?.ToString() ?? ""
                    };
                    _service.InsertDetail(detail, _selectedPO_ID);
                }

                MessageBox.Show($"Đã lưu toàn bộ PO thành công!\n- Số PO: {finalPONo}\n- Số dòng vật tư: {dgvDetails.Rows.Count}",
                                "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadPO();
                LoadDetails(_selectedPO_ID);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi lưu PO: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnDeletePO_Click(object sender, EventArgs e)
        {
            if (_selectedPO_ID == 0)
            {
                MessageBox.Show("Vui lòng chọn đơn PO cần xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBox.Show("Xóa đơn PO này và toàn bộ chi tiết?", "Xác nhận",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    _service.DeletePO(_selectedPO_ID);
                    MessageBox.Show("Xóa thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ClearHeader();
                    dgvDetails.Rows.Clear();
                    lblTotal.Text = "";
                    _selectedPO_ID = 0;
                    LoadPO();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi xóa: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnAddDetail_Click(object sender, EventArgs e)
        {
            int nextItem = dgvDetails.Rows.Count + 1;
            dgvDetails.Rows.Add(0, nextItem, "", "", 0, 0, 0, 0, "PCS", 0, 0, 0, 0, 0, "", "", "");
            dgvDetails.CurrentCell = dgvDetails.Rows[dgvDetails.Rows.Count - 1].Cells["Item_Name"];
        }

        private void BtnDeleteDetail_Click(object sender, EventArgs e)
        {
            if (dgvDetails.SelectedRows.Count == 0) return;
            if (MessageBox.Show("Xóa dòng này?", "Xác nhận",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    var row = dgvDetails.SelectedRows[0];
                    dgvDetails.Rows.Remove(row);
                    UpdateTotal();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi xóa: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnImportMPR_Click(object sender, EventArgs e)
        {
            using (var dlg = new frmSelectMPR())
            {
                if (dlg.ShowDialog() == DialogResult.OK && dlg.SelectedMPR != null)
                {
                    ClearHeader();
                    _selectedPO_ID = 0;
                    _details.Clear();
                    dgvDetails.Rows.Clear();

                    var mpr = dlg.SelectedMPR;
                    var details = dlg.SelectedDetails;

                    txtProjectName.Text = mpr.Project_Name;
                    txtMPRNo.Text = mpr.MPR_No;
                    try
                    {
                        var projects = new ProjectService().GetAll();
                        var project = projects.Find(p =>
                            !string.IsNullOrEmpty(p.ProjectName) &&
                            p.ProjectName.Equals(mpr.Project_Name, StringComparison.OrdinalIgnoreCase));

                        if (project == null)
                            project = projects.Find(p =>
                                !string.IsNullOrEmpty(p.ProjectName) &&
                                (p.ProjectName.Contains(mpr.Project_Name, StringComparison.OrdinalIgnoreCase) ||
                                 mpr.Project_Name.Contains(p.ProjectName, StringComparison.OrdinalIgnoreCase)));

                        if (project != null)
                        {
                            txtWorkorderNo.Text = project.WorkorderNo ?? "";
                            _projectCodeImport = project.ProjectCode;
                            string poCode = project.POCode ?? project.ProjectCode ?? "";
                            string autoPoNo = GenerateAutoPoNo(poCode);
                            txtPONo.Text = autoPoNo;
                        }
                        else
                        {
                            MessageBox.Show(
                                $"Không tìm thấy dự án khớp với tên \"{mpr.Project_Name}\".\n" +
                                "Vui lòng kiểm tra lại thông tin Workorder và PO No.",
                                "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("Lỗi tìm project: " + ex.Message);
                    }

                    int itemNo = 1;
                    foreach (var d in details)
                    {
                        dgvDetails.Rows.Add(
                            0, itemNo++, d.Item_Name, d.Material,
                            d.Thickness_mm, d.C_Width_mm, d.F_Length_mm,
                            d.Qty_Per_Sheet, d.UNIT, d.Weight_kg,
                            0, 0, 0, 0,
                            d.MPS_Info, d.Usage_Location, d.Remarks
                        );
                    }

                    MessageBox.Show(
                        $"✅ Đã import {details.Count} dòng từ MPR {mpr.MPR_No}!\n" +
                        $"PO No dự kiến: {txtPONo.Text}\n" +
                        $"Workorder: {txtWorkorderNo.Text}",
                        "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private string GenerateAutoPoNo(string poCode)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(poCode))
                    poCode = "PRJ";
                string prefix = $"DV-{poCode}-PC-";

                using (var conn = MPR_Managerment.Helpers.DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    var cmd = new Microsoft.Data.SqlClient.SqlCommand(
                        "SELECT COUNT(*) FROM PO_head WHERE PONo LIKE @prefix", conn);
                    cmd.Parameters.AddWithValue("@prefix", prefix + "%");
                    int count = Convert.ToInt32(cmd.ExecuteScalar());

                    int inMemory = _poList.FindAll(p =>
                        (p.PONo ?? "").StartsWith(prefix, StringComparison.OrdinalIgnoreCase)).Count;
                    int nextNo = Math.Max(count, inMemory) + 1;
                    return $"{prefix}{nextNo:D3}";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("GenerateAutoPoNo error: " + ex.Message);
                return $"DV-{poCode}-PC-{DateTime.Now:ddMMHH}";
            }
        }

        private void BtnClearHeader_Click(object sender, EventArgs e)
        {
            ClearHeader();
            dgvDetails.Rows.Clear();
            lblTotal.Text = "";
            _selectedPO_ID = 0;
            LoadPO();
        }

        private void ClearHeader()
        {
            txtPONo.Text = "";
            txtProjectName.Text = "";
            txtWorkorderNo.Text = "";
            txtMPRNo.Text = "";
            txtPrepared.Text = "";
            txtReviewed.Text = "";
            txtAgreement.Text = "";
            txtApproved.Text = "";
            txtNotes.Text = "";
            nudRevise.Value = 0;
            dtpPODate.Value = DateTime.Today;
            cboStatus.SelectedIndex = 0;
        }
    }
}