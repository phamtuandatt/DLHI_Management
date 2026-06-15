using MPR_Managerment.Models;
using MPR_Managerment.Services;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace MPR_Managerment.Forms
{
    /// <summary>
    /// Pop-up tìm kiếm PO theo Nhà cung cấp.
    /// Tính năng:
    ///   - Tìm real-time theo tên NCC (hỗ trợ không dấu)
    ///   - Bộ lọc theo Mã dự án (ComboBox)
    ///   - Bộ lọc theo Trạng thái PO (ComboBox)
    ///   - Cột % Tiến độ giao hàng (Received / Qty * 100), tô màu
    ///   - Double-click hoặc nút "Chọn PO" để trả về PO đã chọn
    /// </summary>
    public partial class frmSearchPOBySupplier : Form
    {
        // ── Kết quả trả về cho frmPO ──────────────────────────────────────────
        public string SelectedPONo { get; private set; } = "";

        // ── Data ──────────────────────────────────────────────────────────────
        private readonly List<POHead> _allPO;          // danh sách đầy đủ được truyền từ frmPO
        private List<Supplier> _suppliers;
        private Dictionary<int, double> _progressMap = new Dictionary<int, double>();
        private Dictionary<string, string> _paidDateMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // ── Controls ──────────────────────────────────────────────────────────
        private TextBox txtSearchNCC;
        private ComboBox cboProjectFilter;
        private ComboBox cboStatusFilter;
        private Label lblCount;
        private DataGridView dgv;
        private Button btnSelect, btnCancel, btnRefresh;

        // ── Trạng thái lọc ────────────────────────────────────────────────────
        private bool _loading = false;
        private Form TopOwner { get { if (!IsDisposed) { BringToFront(); Activate(); } return this; } }

        // =====================================================================
        public frmSearchPOBySupplier(List<POHead> allPO)
        {
            // Gắn AI Trợ lý — hiển thị nút floating + chat panel
            frmAIChat.Attach(this);
            _allPO = allPO ?? new List<POHead>();
            InitForm();
            LoadData();
        }

        // ── Khởi tạo giao diện ───────────────────────────────────────────────
        private void InitForm()
        {
            this.Text = "Tìm kiếm PO theo Nhà cung cấp";
            this.Size = new Size(1100, 620);
            this.MinimumSize = new Size(900, 500);
            this.StartPosition = FormStartPosition.CenterParent;
            this.BackColor = Color.FromArgb(245, 245, 245);
            this.FormBorderStyle = FormBorderStyle.Sizable;

            // ── Tiêu đề ──────────────────────────────────────────────────────
            var lblTitle = new Label
            {
                Text = "🔍  TÌM KIẾM ĐƠN ĐẶT HÀNG THEO NHÀ CUNG CẤP",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                ForeColor = Color.FromArgb(102, 51, 153),
                Location = new Point(12, 12),
                Size = new Size(750, 30),
                AutoSize = false
            };
            this.Controls.Add(lblTitle);

            // ── Panel bộ lọc ─────────────────────────────────────────────────
            var panelFilter = new Panel
            {
                Location = new Point(12, 50),
                Size = new Size(1060, 60),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
            this.Controls.Add(panelFilter);

            // Tìm NCC
            panelFilter.Controls.Add(MkLabel("Nhà cung cấp:", 10, 20));
            txtSearchNCC = new TextBox
            {
                Location = new Point(105, 17),
                Size = new Size(250, 26),
                Font = new Font("Segoe UI", 10),
                PlaceholderText = "Gõ để lọc ngay..."
            };
            panelFilter.Controls.Add(txtSearchNCC);
            txtSearchNCC.TextChanged += (s, e) => ApplyFilter();

            // Lọc dự án
            panelFilter.Controls.Add(MkLabel("Dự án:", 370, 20));
            cboProjectFilter = new ComboBox
            {
                Location = new Point(420, 17),
                Size = new Size(220, 26),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            panelFilter.Controls.Add(cboProjectFilter);
            cboProjectFilter.SelectedIndexChanged += (s, e) => { if (!_loading) ApplyFilter(); };

            // Lọc trạng thái
            panelFilter.Controls.Add(MkLabel("Trạng thái:", 655, 20));
            cboStatusFilter = new ComboBox
            {
                Location = new Point(735, 17),
                Size = new Size(140, 26),
                Font = new Font("Segoe UI", 9),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboStatusFilter.Items.Add("-- Tất cả --");
            cboStatusFilter.Items.AddRange(new[] { "Draft", "Pending", "Approved", "In Progress", "Completed", "Cancelled" });
            cboStatusFilter.SelectedIndex = 0;
            panelFilter.Controls.Add(cboStatusFilter);
            cboStatusFilter.SelectedIndexChanged += (s, e) => { if (!_loading) ApplyFilter(); };

            // Nút làm mới
            btnRefresh = MkButton("↺ Làm mới", Color.FromArgb(108, 117, 125), new Point(890, 14), 130, 30);
            btnRefresh.Click += (s, e) =>
            {
                _loading = true;
                txtSearchNCC.Text = "";
                cboProjectFilter.SelectedIndex = 0;
                cboStatusFilter.SelectedIndex = 0;
                _loading = false;
                ApplyFilter();
            };
            panelFilter.Controls.Add(btnRefresh);

            // ── Label đếm kết quả ────────────────────────────────────────────
            lblCount = new Label
            {
                Location = new Point(12, 118),
                Size = new Size(500, 22),
                Font = new Font("Segoe UI", 9, FontStyle.Italic),
                ForeColor = Color.FromArgb(0, 120, 212)
            };
            this.Controls.Add(lblCount);

            // ── DataGridView ─────────────────────────────────────────────────
            dgv = new DataGridView
            {
                Location = new Point(12, 144),
                Size = new Size(1060, 390),
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("Segoe UI", 9),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(102, 51, 153);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.EnableHeadersVisualStyles = false;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 244, 255);
            dgv.CellFormatting += Dgv_CellFormatting;
            dgv.CellDoubleClick += Dgv_CellDoubleClick;
            this.Controls.Add(dgv);

            BuildColumns();

            // ── Nút dưới cùng ────────────────────────────────────────────────
            btnSelect = MkButton("✔ Chọn PO này", Color.FromArgb(40, 167, 69), new Point(12, 548), 150, 34);
            btnCancel = MkButton("Đóng", Color.FromArgb(220, 53, 69), new Point(172, 548), 90, 34);

            btnSelect.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;

            btnSelect.Click += BtnSelect_Click;
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            this.Controls.Add(btnSelect);
            this.Controls.Add(btnCancel);

            // Resize
            this.Resize += (s, e) =>
            {
                int w = this.ClientSize.Width - 24;
                panelFilter.Width = w;
                dgv.Width = w;
                dgv.Height = this.ClientSize.Height - dgv.Top - 55;
                btnSelect.Top = dgv.Bottom + 10;
                btnCancel.Top = dgv.Bottom + 10;
            };
        }

        // ── Cột DataGridView ─────────────────────────────────────────────────
        private void BuildColumns()
        {
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "PO_ID", HeaderText = "ID", Visible = false });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "PO_No", HeaderText = "PO No", FillWeight = 120 });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "NCC", HeaderText = "Nhà cung cấp", FillWeight = 180 });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Du_An", HeaderText = "Dự án", FillWeight = 200 });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "MPR_No", HeaderText = "MPR No", FillWeight = 110 });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Tien_Do", HeaderText = "% Tiến độ", FillWeight = 80 });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Ngay_PO", HeaderText = "Ngày PO", FillWeight = 90 });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Trang_Thai", HeaderText = "Trạng thái", FillWeight = 90 });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Tong_Tien", HeaderText = "Tổng tiền", FillWeight = 110 });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Revise", HeaderText = "Rev", FillWeight = 50 });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Paid_Date", HeaderText = "Paid Date", FillWeight = 90 });
        }

        // ── Tải dữ liệu lần đầu ──────────────────────────────────────────────
        private void LoadData()
        {
            _loading = true;
            try
            {
                _suppliers = new SupplierService().GetAll();

                // Tính % tiến độ cho tất cả PO
                var poSvc = new POService();
                foreach (var po in _allPO)
                {
                    try
                    {
                        var dets = poSvc.GetDetails(po.PO_ID);
                        double totalQty = dets.Sum(d => (double)d.Qty_Per_Sheet);
                        double totalRcv = dets.Sum(d => (double)d.Received);
                        _progressMap[po.PO_ID] = totalQty > 0
                            ? Math.Round(totalRcv / totalQty * 100, 1) : 0;
                    }
                    catch { _progressMap[po.PO_ID] = 0; }
                }

                // Nạp combo Dự án
                cboProjectFilter.Items.Clear();
                cboProjectFilter.Items.Add("-- Tất cả --");
                var projectNames = _allPO
                    .Select(p => p.Project_Name ?? "")
                    .Where(n => !string.IsNullOrWhiteSpace(n))
                    .Distinct()
                    .OrderBy(n => n)
                    .ToList();
                foreach (var n in projectNames)
                    cboProjectFilter.Items.Add(n);
                cboProjectFilter.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(TopOwner, "Lỗi tải dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _loading = false;
            }

            LoadPaidDates();
            ApplyFilter();
        }

        private string RobustNormalize(string input)
        {
            if (string.IsNullOrEmpty(input)) return "";
            // Loại bỏ tất cả ký tự không phải chữ hoặc số, chuyển về chữ hoa
            // Ví dụ: DV-WOL-PC-010 -> DVWOLPC010, PO 24-001 -> PO24001
            return new string(input.Where(char.IsLetterOrDigit).ToArray()).ToUpperInvariant();
        }

        private void LoadPaidDates()
        {
            _paidDateMap.Clear();
            int recordCount = 0;
            try
            {
                using var conn = Helpers.DatabaseHelper.GetConnection();
                conn.Open();
                
                // Lấy tất cả bản ghi có Paid_Date. Mở rộng lấy thêm Title_EN và GW_No để tăng khả năng khớp
                         using var cmd = new Microsoft.Data.SqlClient.SqlCommand(@"
                         SELECT ISNULL(PO_No, '') as PO_No, ISNULL(Ecount_No, '') as Ecount_No, 
                            ISNULL(Title_EN, '') as Title_EN, ISNULL(GW_No, '') as GW_No, Paid_Date 
                         FROM Zalo_PaymentImport 
                         WHERE Paid_Date IS NOT NULL", conn);
 
                 var raw = new Dictionary<string, List<DateTime>>(StringComparer.OrdinalIgnoreCase);
                 using (var rdr = cmd.ExecuteReader())
                 {
                     while (rdr.Read())
                     {
                         recordCount++;
                         DateTime dt = Convert.ToDateTime(rdr["Paid_Date"]);
                         string dbPo = rdr["PO_No"].ToString().Trim();
                         string dbEc = rdr["Ecount_No"].ToString().Trim();
                         string dbTi = rdr["Title_EN"].ToString().Trim();
                         string dbGw = rdr["GW_No"].ToString().Trim();
 
                         // Danh sách các key có thể dùng để khớp cho bản ghi này
                         var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                         
                         void AddKey(string k) {
                             if (string.IsNullOrEmpty(k)) return;
                             keys.Add(k);
                             string norm = RobustNormalize(k);
                             if (!string.IsNullOrEmpty(norm)) keys.Add(norm);
                         }
 
                         AddKey(dbPo);
                         AddKey(dbEc);
                         AddKey(dbGw);
 
                         // Đối với Title_EN, nếu chứa số PO (ví dụ "Thanh toán cho PO-123") thì cũng cố gắng bóc tách
                         if (!string.IsNullOrEmpty(dbTi)) {
                             keys.Add(dbTi);
                             // Nếu trong Title có chứa số PO dạng DV-...
                             if (dbTi.Contains("DV-")) {
                                 int idx = dbTi.IndexOf("DV-");
                                 string sub = dbTi.Substring(idx).Split(' ')[0]; // lấy chuỗi liên tục từ DV-
                                 AddKey(sub);
                             }
                         }
                         if (!string.IsNullOrEmpty(dbEc))
                         {
                             keys.Add(dbEc);
                             keys.Add(RobustNormalize(dbEc));
                         }
 
                        // NEW: Add digit‑only key for PO_No to improve matching (e.g., “12345”)
                        if (!string.IsNullOrEmpty(dbPo)) {
                            string digitsOnly = new string(dbPo.Where(char.IsDigit).ToArray());
                            if (digitsOnly.Length >= 3) AddKey(digitsOnly);
                        }

                         foreach (var k in keys)
                         {
                             if (!raw.ContainsKey(k)) raw[k] = new List<DateTime>();
                             raw[k].Add(dt);
                         }
                     }
                 }

                foreach (var kv in raw)
                {
                    var distinctDates = kv.Value.Select(d => d.Date).Distinct().OrderBy(d => d).Select(d => d.ToString("dd/MM/yyyy"));
                    _paidDateMap[kv.Key] = string.Join(", ", distinctDates);
                }
                
                System.Diagnostics.Debug.WriteLine($"[LoadPaidDates] Loaded {recordCount} records, map size: {_paidDateMap.Count}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Lỗi tải Paid Date:\n" + ex.Message,
                    "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

        }

        // ── Lọc và bind lại grid ─────────────────────────────────────────────
        private void ApplyFilter()
        {
            string keyword = txtSearchNCC.Text.Trim();
            string kwNorm = RemoveDiacritics(keyword).ToLower();
            string projectSel = cboProjectFilter.SelectedIndex > 0 ? cboProjectFilter.SelectedItem?.ToString() ?? "" : "";
            string statusSel = cboStatusFilter.SelectedIndex > 0 ? cboStatusFilter.SelectedItem?.ToString() ?? "" : "";

            // Lấy Supplier_ID khớp từ khóa NCC
            HashSet<int> matchedSupplierIds = null;
            if (!string.IsNullOrEmpty(keyword))
            {
                matchedSupplierIds = _suppliers
                    .Where(s =>
                        RemoveDiacritics(s.Short_Name ?? "").ToLower().Contains(kwNorm) ||
                        RemoveDiacritics(s.Company_Name ?? "").ToLower().Contains(kwNorm) ||
                        (s.Short_Name ?? "").ToLower().Contains(keyword.ToLower()) ||
                        (s.Company_Name ?? "").ToLower().Contains(keyword.ToLower()))
                    .Select(s => s.Supplier_ID)
                    .ToHashSet();
            }

            var filtered = _allPO.Where(po =>
            {
                if (matchedSupplierIds != null && !matchedSupplierIds.Contains(po.Supplier_ID)) return false;
                if (!string.IsNullOrEmpty(projectSel) && (po.Project_Name ?? "") != projectSel) return false;
                if (!string.IsNullOrEmpty(statusSel) && (po.Status ?? "") != statusSel) return false;
                return true;
            }).ToList();

            BindGrid(filtered);
        }

        private void BindGrid(List<POHead> list)
        {
            dgv.Rows.Clear();
            foreach (var h in list)
            {
                var supplier = _suppliers?.Find(s => s.Supplier_ID == h.Supplier_ID);
                double pct = _progressMap.ContainsKey(h.PO_ID) ? _progressMap[h.PO_ID] : 0;

                int idx = dgv.Rows.Add();
                var row = dgv.Rows[idx];
                row.Cells["PO_ID"].Value = h.PO_ID;
                row.Cells["PO_No"].Value = h.PONo;
                row.Cells["NCC"].Value = supplier?.Short_Name ?? supplier?.Company_Name ?? "";
                row.Cells["Du_An"].Value = h.Project_Name;
                row.Cells["MPR_No"].Value = h.MPR_No;
                row.Cells["Tien_Do"].Value = $"{pct:F1}%";
                row.Cells["Ngay_PO"].Value = h.PO_Date.HasValue ? h.PO_Date.Value.ToString("dd/MM/yyyy") : "";
                row.Cells["Trang_Thai"].Value = h.Status;
                row.Cells["Tong_Tien"].Value = h.Total_Amount.ToString("N0");
                row.Cells["Revise"].Value = h.Revise;

                // Tìm kiếm Paid Date từ Map
                string displayPaidDate = "";
                string poNo = (h.PONo ?? "").Trim();

                if (!string.IsNullOrEmpty(poNo))
                {
                    // Ưu tiên 1: Khớp theo chuỗi chuẩn hóa Alphanumeric (Ví dụ: DV-WOL-01 -> DVWOL01)
                    // Cách này hiệu quả nhất vì ZaloImportService thường xóa khoảng trắng khi lưu.
                    string norm = RobustNormalize(poNo);
                    
                    if (_paidDateMap.TryGetValue(poNo, out string d1))
                    {
                        displayPaidDate = d1;
                    }
                    else if (!string.IsNullOrEmpty(norm) && _paidDateMap.TryGetValue(norm, out string d2))
                    {
                        displayPaidDate = d2;
                    }
                    else
                    {
                        // Ưu tiên 2: Khớp theo phần số (nếu có ít nhất 3 chữ số)
                        string digitsOnly = new string(poNo.Where(char.IsDigit).ToArray());
                        if (digitsOnly.Length >= 3 && _paidDateMap.TryGetValue(digitsOnly, out string d3))
                        {
                            displayPaidDate = d3;
                        }
                    }
                }

                row.Cells["Paid_Date"].Value = displayPaidDate;
            }

            lblCount.Text = $"Tìm thấy {list.Count} đơn PO (Đã tải {_paidDateMap.Count} đầu mã thanh toán)";
        }

        // ── Tô màu cột % Tiến độ ─────────────────────────────────────────────
        private void Dgv_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (dgv.Columns[e.ColumnIndex].Name != "Tien_Do") return;

            string raw = e.Value?.ToString()?.Replace("%", "").Trim() ?? "0";
            if (!double.TryParse(raw,
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out double pct)) return;

            if (pct >= 100)
            {
                e.CellStyle.BackColor = Color.FromArgb(198, 239, 206);
                e.CellStyle.ForeColor = Color.FromArgb(0, 97, 0);
                e.CellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }
            else if (pct >= 30)
            {
                e.CellStyle.BackColor = Color.FromArgb(255, 235, 156);
                e.CellStyle.ForeColor = Color.FromArgb(156, 87, 0);
            }
            else if (pct > 0)
            {
                e.CellStyle.BackColor = Color.FromArgb(255, 199, 206);
                e.CellStyle.ForeColor = Color.FromArgb(156, 0, 6);
            }
        }

        // ── Chọn PO ──────────────────────────────────────────────────────────
        private void Dgv_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0) ConfirmSelection(e.RowIndex);
        }

        private void BtnSelect_Click(object sender, EventArgs e)
        {
            if (dgv.SelectedRows.Count == 0)
            { MessageBox.Show(TopOwner, "Vui lòng chọn một đơn PO!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            ConfirmSelection(dgv.SelectedRows[0].Index);
        }

        private void ConfirmSelection(int rowIndex)
        {
            SelectedPONo = dgv.Rows[rowIndex].Cells["PO_No"].Value?.ToString() ?? "";
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        // ── Helpers ──────────────────────────────────────────────────────────
        private static string RemoveDiacritics(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            try
            {
                string norm = text.Normalize(System.Text.NormalizationForm.FormD);
                var sb = new System.Text.StringBuilder();
                foreach (char c in norm)
                    if (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c) != System.Globalization.UnicodeCategory.NonSpacingMark)
                        sb.Append(c);
                return sb.ToString().Normalize(System.Text.NormalizationForm.FormC);
            }
            catch { return text; }
        }

        private static Label MkLabel(string text, int x, int y) =>
            new Label { Text = text, Location = new Point(x, y + 3), AutoSize = true, Font = new Font("Segoe UI", 9) };

        private static Button MkButton(string text, Color color, Point loc, int w, int h) =>
            new Button
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
    }
}