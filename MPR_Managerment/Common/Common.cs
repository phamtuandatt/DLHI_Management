using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MPR_Managerment.Common
{
    public static class Common
    {
        private static readonly System.Globalization.CultureInfo _numCulture = new System.Globalization.CultureInfo("vi-VN");

        public static void HideColumnDataGridView(DataGridView dgv, List<string> lstColumn)
        {
            if (dgv.Rows.Count == 0 || lstColumn == null) return;

            foreach (DataGridViewColumn column in dgv.Columns)
            {
                if (lstColumn.Contains(column.Name))
                {
                    column.Visible = false;
                }
            }
        }

        public static void ComboBoxTextUpdateForListItem(object? sender, EventArgs e, List<string> itemList)
        {
            ComboBox cb = sender as ComboBox;
            string typedText = cb.Text;

            // Nếu trống thì hiện lại toàn bộ danh sách
            if (string.IsNullOrWhiteSpace(typedText) || typedText == "-- Chọn PO --")
            {
                cb.Items.Clear();
                cb.Items.Add("-- Chọn PO --");
                cb.Items.AddRange(itemList.ToArray());
            }
            else
            {
                // Lọc danh sách gốc theo điều kiện "Contains"
                var filtered = itemList
                    .Where(x => x.IndexOf(typedText, StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();

                cb.Items.Clear();
                cb.Items.AddRange(filtered.ToArray());
            }

            // Giữ trạng thái hiển thị
            cb.DroppedDown = true;
            cb.Text = typedText;
            cb.SelectionStart = typedText.Length; // Đưa con trỏ về cuối văn bản
            Cursor.Current = Cursors.Default;
        }

        public static ComboBox CreateCombox(string cboName, Point location, int width = 200)
        {
            ComboBox cb = new ComboBox()
            {
                Name = cboName,
                Location = location,
                Width = width,
                DropDownStyle = ComboBoxStyle.DropDown,
                // Quan trọng: Để tự xử lý search "Contains", ta tắt AutoComplete mặc định
                AutoCompleteMode = AutoCompleteMode.None
            };
            return cb;
        }

        public static void ComboBoxTextUpdateForDataTable(object? sender, EventArgs e)
        {
            ComboBox cb = sender as ComboBox;
            string filterText = cb.Text;

            if (string.IsNullOrWhiteSpace(filterText))
            {
                (cb.DataSource as DataView).RowFilter = "";
                cb.SelectedIndex = 0;
            }
            else
            {
                // Lọc các item có chứa (Contains) chuỗi nhập vào
                (cb.DataSource as DataView).RowFilter = $"{cb.DisplayMember} LIKE '%{filterText}%'";
            }

            // Giữ cho Dropdown luôn mở và cập nhật text người dùng đang gõ
            cb.DroppedDown = true;
            cb.Text = filterText;
            cb.SelectionStart = filterText.Length;
            Cursor.Current = Cursors.Default;
        }

        public static ComboBox CreateComboxUsingDataTable(DataTable dtSource, string displayMember, string valueMember, string cboName, Point location, int width = 200)
        {
            ComboBox cb = new ComboBox()
            {
                Name = cboName,
                Location = location,
                Width = width,
                DropDownStyle = ComboBoxStyle.DropDown,
                // Lưu ý: Tắt AutoComplete mặc định của Windows để dùng Custom Filter bên dưới
                AutoCompleteMode = AutoCompleteMode.None,
                DataSource = dtSource.DefaultView,
                DisplayMember = displayMember,
                ValueMember = valueMember
            };
            return cb;
        }

        public static void SetupComboBoxSearchUsingDataTable(ComboBox cb, DataTable dtSource, string displayMember, string valueMember)
        {
            cb.DropDownStyle = ComboBoxStyle.DropDown;
            // Lưu ý: Tắt AutoComplete mặc định của Windows để dùng Custom Filter bên dưới
            cb.AutoCompleteMode = AutoCompleteMode.None;
            cb.DataSource = dtSource.DefaultView;
            cb.DisplayMember = displayMember;
            cb.ValueMember = valueMember;
        }

        public static void AutoAdjustColumnWidths(DataGridView dgvDetails)
        {
            if (dgvDetails.Columns.Count == 0) return;
            dgvDetails.SuspendLayout();
            dgvDetails.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            using (Graphics g = dgvDetails.CreateGraphics())
            {
                Font headerFont = dgvDetails.ColumnHeadersDefaultCellStyle.Font ?? dgvDetails.Font;
                foreach (DataGridViewColumn col in dgvDetails.Columns)
                {
                    if (!col.Visible) continue;
                    col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    SizeF headerSize = g.MeasureString(col.HeaderText, headerFont);
                    int minWidth = (int)Math.Ceiling(headerSize.Width) + 20; int maxWidth = minWidth;
                    foreach (DataGridViewRow row in dgvDetails.Rows)
                    {
                        if (row.IsNewRow || row.Tag?.ToString() == "TOTAL") continue;
                        string cellValue = row.Cells[col.Index].Value?.ToString() ?? "";
                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            SizeF size = g.MeasureString(cellValue, dgvDetails.Font); int cw = (int)Math.Ceiling(size.Width) + 15;
                            if (cw > maxWidth) maxWidth = cw;
                        }
                    }
                    if (maxWidth > 200) { col.Width = 200; col.DefaultCellStyle.WrapMode = DataGridViewTriState.True; }
                    else { col.Width = maxWidth; col.DefaultCellStyle.WrapMode = DataGridViewTriState.False; }
                }
            }
            dgvDetails.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            dgvDetails.ResumeLayout();
        }

        public static void ColorRowsByIdGroups(DataGridView dgv, string idColumnName)
        {
            // 1. Đếm số lần xuất hiện của từng ID
            // Key: Giá trị ID, Value: Số lần xuất hiện
            Dictionary<string, int> idCounts = new Dictionary<string, int>();

            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow) continue;
                string idValue = row.Cells[idColumnName].Value?.ToString();
                if (string.IsNullOrEmpty(idValue)) continue;

                if (idCounts.ContainsKey(idValue))
                    idCounts[idValue]++;
                else
                    idCounts[idValue] = 1;
            }

            // 2. Bảng màu đậm rõ nét cho các nhóm trùng
            Color[] groupColors = new Color[]
            {
                Color.FromArgb(210, 230, 255), // Xanh dương
                Color.FromArgb(200, 245, 210), // Xanh lá
                Color.FromArgb(255, 235, 190), // Vàng cam
                Color.FromArgb(235, 210, 255), // Tím
                Color.FromArgb(255, 210, 210)  // Hồng
            };

            Dictionary<string, Color> idColorMap = new Dictionary<string, Color>();
            int colorIndex = 0;

            dgv.SuspendLayout();

            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow) continue;
                string idValue = row.Cells[idColumnName].Value?.ToString();

                // Kiểm tra: Chỉ tô màu nếu ID này xuất hiện từ 2 lần trở lên
                if (!string.IsNullOrEmpty(idValue) && idCounts[idValue] > 1)
                {
                    if (!idColorMap.ContainsKey(idValue))
                    {
                        idColorMap[idValue] = groupColors[colorIndex % groupColors.Length];
                        colorIndex++;
                    }
                    row.DefaultCellStyle.BackColor = idColorMap[idValue];
                }
                else
                {
                    // Nếu không trùng, trả về màu mặc định (Trắng hoặc màu xen kẽ của Grid)
                    row.DefaultCellStyle.BackColor = Color.White;
                }
            }

            dgv.ResumeLayout();
        }

        public static void CreateDataGridView(DataGridView dgv)
        {
            //dgv.ReadOnly = false; // Phải để false để có thể click vào CheckBox
            //dgv.RowHeadersVisible = false;
            dgv.AllowUserToAddRows = false;
            dgv.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dgv.BackgroundColor = Color.White;
            dgv.BorderStyle = BorderStyle.FixedSingle;
            dgv.Font = new Font("Segoe UI", 9);
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgv.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            dgv.Margin = new Padding(0, 100, 0, 0);

            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.EnableHeadersVisualStyles = false;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgv.Dock = DockStyle.Fill;
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(204, 232, 255);
            dgv.DefaultCellStyle.SelectionForeColor = Color.Black;

            dgv.RowPostPaint += (s, e) =>
            {
                RenderNumbering(s, e);
            };
        }

        public static void CreateDataGridView_Hide_RowHeader(DataGridView dgv)
        {
            dgv.ReadOnly = false; // Phải để false để có thể click vào CheckBox
            dgv.RowHeadersVisible = false;
            dgv.AllowUserToAddRows = false;
            dgv.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dgv.BackgroundColor = Color.White;
            dgv.BorderStyle = BorderStyle.FixedSingle;
            dgv.Font = new Font("Segoe UI", 9);
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgv.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            dgv.Margin = new Padding(0, 100, 0, 0);

            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.EnableHeadersVisualStyles = false;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            dgv.Dock = DockStyle.Fill;
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(204, 232, 255);
            dgv.DefaultCellStyle.SelectionForeColor = Color.Black;
        }

        public static void CreateButtonSave(Button btn, string text)
        {
            btn.Text = text ?? "💾 Lưu phiếu";
            btn.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            btn.Cursor = Cursors.Hand;
            btn.BackColor = Color.FromArgb(0, 120, 212);
            btn.ForeColor = Color.White;
            btn.FlatStyle = FlatStyle.Flat;
        }

        public static void CreateButtonAdd(Button btn, string text)
        {
            btn.Text = text ?? "✔ Thêm vào phiếu";
            btn.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            btn.Cursor = Cursors.Hand;
            btn.BackColor = Color.FromArgb(40, 167, 69);
            btn.ForeColor = Color.White;
            btn.FlatStyle = FlatStyle.Flat;
        }

        public static void CreateButtonCancel(Button btn, string text)
        {
            btn.Text = "🆕 Phiếu mới";
            btn.BackColor = Color.FromArgb(108, 117, 125);
            btn.ForeColor = Color.White;
            btn.FlatStyle = FlatStyle.Flat;
            btn.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            btn.Cursor = Cursors.Hand;
        }

        public static void CreateButtonSearch(Button btn, string text)
        {
            btn.Text = text ?? "🔍 Tìm";
            btn.Cursor = Cursors.Hand;
            btn.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            btn.BackColor = Color.FromArgb(0, 120, 212);
            btn.ForeColor = Color.White;
            btn.FlatStyle = FlatStyle.Flat;
        }

        public static void CreateButtonDelete(Button btn, string text = "🗑 Xóa dòng")
        {
            btn.Text = text;
            btn.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            btn.BackColor = Color.FromArgb(232, 17, 35);
            btn.ForeColor = Color.White;
            btn.FlatStyle = FlatStyle.Flat;
            btn.Cursor = Cursors.Hand;
        }

        public static void CreateButtonPrint(Button btn, string text = "")
        {
            btn.Text = text ?? "🖨 In";
            btn.BackColor = Color.FromArgb(33, 115, 70);
            btn.ForeColor = Color.White;
            btn.FlatStyle = FlatStyle.Flat;
            btn.Font = new Font("Segoe UI", 9, FontStyle.Bold);
        }

        public static void CreateButtonRefresh(Button btn)
        {
            btn.Text = "🔄 Làm mới";
            btn.BackColor = Color.FromArgb(108, 117, 125);
            btn.ForeColor = Color.White;
            btn.FlatStyle = FlatStyle.Flat;
            btn.Font = new Font("Segoe UI", 9, FontStyle.Bold);
        }

        public static void Column_KeyPress_Digital(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        public static Label AddCard(Panel p, string title, string value, Color color, int x)
        {
            var card = new Panel { BackColor = color };
            p.Controls.Add(card);
            card.Controls.Add(new Label { Text = title, Font = new Font("Segoe UI", 8, FontStyle.Bold), ForeColor = Color.White });
            var lbl = new Label { Text = value, Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = Color.White };
            card.Controls.Add(lbl);
            return lbl;
        }

        public static void UpdateSelectionSum(DataGridView dgv, Label lblStatus)
        {
            decimal totalSum = 0;
            int count = 0;
            bool hasNumber = false;

            // Duyệt qua tất cả các ô đang được bôi đen (SelectedCells)
            foreach (DataGridViewCell cell in dgv.SelectedCells)
            {
                if (cell.Value != null && cell.Value != DBNull.Value)
                {
                    // Sử dụng hàm SafeParse (đã xây dựng ở các bước trước) để đọc số an toàn
                    string cellValue = cell.Value.ToString().Replace(",", "").Trim();

                    if (decimal.TryParse(cellValue, System.Globalization.NumberStyles.Any,
                                         System.Globalization.CultureInfo.InvariantCulture, out decimal val))
                    {
                        totalSum += val;
                        count++;
                        hasNumber = true;
                    }
                }
            }

            // Hiển thị kết quả lên Label hoặc StatusStrip
            if (hasNumber && count > 1) // Chỉ hiện khi chọn từ 2 ô số trở lên
            {
                lblStatus.Text = $"Count: {count}  |  Sum: {totalSum:N0}"; // N0 để định dạng dấu phẩy hàng nghìn
            }
            else
            {
                lblStatus.Text = "Ready";
            }
        }

        public static decimal ParseDecimalRaw(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return 0;
            raw = raw.Trim();

            // Co ca "." va ","
            if (raw.Contains(".") && raw.Contains(","))
            {
                if (raw.IndexOf(".") < raw.IndexOf(","))
                {
                    // "1.234,56" -> vi-VN -> bo . -> "1234,56"
                    raw = raw.Replace(".", "");
                }
                else
                {
                    // "1,234.56" -> InvariantCulture -> bo , -> "1234.56" -> doi . thanh ,
                    raw = raw.Replace(",", "").Replace(".", ",");
                }
            }
            else if (raw.Contains(".") && !raw.Contains(","))
            {
                // Chi co "."
                var parts = raw.Split('.');
                // Neu tat ca phan sau dau . deu co 3 chu so -> day la ngan separator
                bool allThousand = parts.Length > 1 &&
                    parts.Skip(1).All(p => p.Length == 3);
                if (allThousand)
                    raw = raw.Replace(".", "");        // bo ngan
                else
                    raw = raw.Replace(".", ",");        // doi . -> , (thap phan vi-VN)
            }
            // Chi co "," hoac so nguyen: vi-VN hieu "," la thap phan -> giu nguyen

            decimal.TryParse(raw,
                System.Globalization.NumberStyles.Number,
                _numCulture, out decimal result);
            return result;
        }


        public static DataView SearchDate(DateTime FromDate, DateTime ToDate, DataTable dtSource, List<string> lstProperties)
        {
            DataView dv = dtSource.DefaultView;
            string filter = "";
            foreach (var item in lstProperties)
            {
                if (string.IsNullOrEmpty(filter))
                {
                    filter = $"{item} >= '{FromDate:dd/MM/yyyy}' ";
                }
                else
                {
                    filter += $"AND {item} <= '{ToDate:dd/MM/yyyy}' ";
                }
            }

            dv.RowFilter = filter;

            return dv;
        }
        public static Dictionary<int, decimal> GetDictionaryDifferences(Dictionary<int, decimal> dict1, Dictionary<int, decimal> dict2)
        {
            // Kiểm tra null để tránh lỗi Runtime
            if (dict1 == null) return dict2 ?? new Dictionary<int, decimal>();
            if (dict2 == null) return dict1;

            // Cách 1: Sử dụng KeyValuePair tường minh để fix lỗi CS0103
            var diff1 = dict1.Where((KeyValuePair<int, decimal> kvp) =>
                !dict2.ContainsKey(kvp.Key) || dict2[kvp.Key] != kvp.Value);

            var diff2 = dict2.Where((KeyValuePair<int, decimal> kvp) =>
                !dict1.ContainsKey(kvp.Key));

            // Kết hợp và chuyển về Dictionary
            return diff1.Concat(diff2)
                        .ToDictionary(x => x.Key, x => x.Value);
        }

        public static DataView Search(string search, DataTable dtSource, List<string> lstProperty)
        {
            DataView dv = dtSource.DefaultView;
            string filter = "";
            foreach (var item in lstProperty)
            {
                if (string.IsNullOrEmpty(filter))
                {
                    filter = $"{item} LIKE '%{search}%' ";
                }
                else
                {
                    filter += $"OR {item} LIKE '%{search}%' ";
                }
            }
            dv.RowFilter = filter;

            return dv;
        }

        public static bool IsDataGridViewValid(DataGridView dgv, string gridName = "Danh sách")
        {
            // 1. Kiểm tra Grid có bị null (chưa khởi tạo) không
            if (dgv == null)
            {
                MessageBoxHelper.ShowError(dgv, $"{gridName} chưa được khởi tạo!", "Lỗi hệ thống");
                return false;
            }

            // 2. Kiểm tra xem Grid có dữ liệu hay không
            // (Kiểm tra cả DataSource và số lượng dòng thực tế)
            if (dgv.Rows.Count == 0 || (dgv.DataSource != null && ((DataTable)dgv.DataSource).Rows.Count == 0))
            {
                MessageBoxHelper.ShowWarning(dgv, $"{gridName} hiện đang trống, không có dữ liệu để xử lý!", "Thông báo");
                return false;
            }

            // 3. Kiểm tra xem có dòng nào đang được chọn hay không
            if (dgv.CurrentRow == null || dgv.CurrentRow.Index < 0)
            {
                MessageBoxHelper.ShowInfo(dgv, $"Vui lòng chọn một dòng trong {gridName} để tiếp tục!", "Thông báo");
                return false;
            }

            return true;
        }

        public static bool IsComboBoxValid(ComboBox cbo, string displayName = "Dữ liệu")
        {
            // 1. Kiểm tra ComboBox có bị null về mặt khởi tạo không
            if (cbo == null)
            {
                MessageBoxHelper.ShowError(cbo, $"{displayName} chưa được khởi tạo!", "Lỗi hệ thống");
                return false;
            }

            // 2. Kiểm tra DataSource (nếu bạn dùng Binding dữ liệu)
            if (cbo.DataSource == null && cbo.Items.Count == 0)
            {
                MessageBoxHelper.ShowWarning(cbo, $"Danh sách {displayName} đang trống. Vui lòng kiểm tra lại kết nối dữ liệu!", "Thông báo");
                return false;
            }

            // 3. Kiểm tra xem có Item nào đang được chọn không
            if (cbo.SelectedIndex == -1)
            {
                MessageBoxHelper.ShowInfo(cbo, $"Vui lòng chọn một giá trị từ {displayName}!", "Thông báo");
                cbo.Focus(); // Đưa con trỏ vào combobox để người dùng chọn luôn
                return false;
            }

            return true;
        }

        public static void AutoBringToFontControl(Array panels)
        {
            foreach (Panel panel in panels)
            {
                foreach (Control c in panel.Controls)
                {
                    if (c is TextBox || c is ComboBox || c is DateTimePicker || c is NumericUpDown)
                        c.BringToFront();
                }
            }
        }

//ok

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

            //if (!matched &&
            //    cb.SelectedItem is DataRowView selected &&
            //    selected[displayMember]?.ToString() != typedText)
            //{
            //    cb.SelectedIndex = 0;
            //}
            if (!matched)
            {
                cb.SelectedIndex = 0;
            }
        }

        public static void RenderNumbering(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var grid = sender as DataGridView;
            var rowIdx = (e.RowIndex + 1).ToString();

            var centerFormat = new StringFormat()
            {
                // right alignment might actually make more sense for numbers
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            var headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);
            e.Graphics.DrawString(rowIdx, new Font("Segoe UI", 9.0f), SystemBrushes.ControlText, headerBounds, centerFormat);
        }

        public static void ApplyCustomFormatting(DataGridViewCellFormattingEventArgs e, DataGridView dgv, string targetColumn, List<StringRule> strRules = null, List<NumericRule> numRules = null)
        {
            // Kiểm tra nếu đúng cột cần định dạng
            if (dgv.Columns[e.ColumnIndex].Name != targetColumn && e.ColumnIndex != dgv.Columns[targetColumn].Index)
                return;

            if (e.Value == null || e.Value == DBNull.Value) return;

            // --- TRƯỜNG HỢP 1: ĐỊNH DẠNG SỐ ---
            if (numRules != null && decimal.TryParse(e.Value.ToString(), out decimal numValue))
            {
                foreach (var rule in numRules)
                {
                    if (numValue >= rule.MinValue && numValue <= rule.MaxValue)
                    {
                        e.CellStyle.ForeColor = rule.CellColor;
                        e.CellStyle.Font = new Font(dgv.Font, FontStyle.Bold); // In đậm cho nổi bật
                        break;
                    }
                }
            }

            // --- TRƯỜNG HỢP 2: ĐỊNH DẠNG CHUỖI ---
            if (strRules != null)
            {
                string cellText = e.Value.ToString().ToLower().Trim();
                foreach (var rule in strRules)
                {
                    bool isMatch = rule.IsFullMatch ? cellText == rule.Value.ToLower() : cellText.Contains(rule.Value.ToLower());

                    if (isMatch)
                    {
                        e.CellStyle.ForeColor = rule.CellColor;
                        e.CellStyle.Font = new Font(dgv.Font, FontStyle.Bold);
                        break;
                    }
                }
            }
        }
    }

    // Quy tắc cho chuỗi (String Mapping)
    public class StringRule
    {
        public string Value { get; set; }
        public Color CellColor { get; set; }
        public bool IsFullMatch { get; set; } = true; // True: khớp hoàn toàn, False: chỉ cần chứa từ đó
    }

    // Quy tắc cho số (Numeric Range)
    public class NumericRule
    {
        public decimal MinValue { get; set; }
        public decimal MaxValue { get; set; }
        public Color CellColor { get; set; }
    }
}
