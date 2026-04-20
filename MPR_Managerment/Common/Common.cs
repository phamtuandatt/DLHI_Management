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
                MessageBox.Show($"{gridName} chưa được khởi tạo!", "Lỗi hệ thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            // 2. Kiểm tra xem Grid có dữ liệu hay không
            // (Kiểm tra cả DataSource và số lượng dòng thực tế)
            if (dgv.Rows.Count == 0 || (dgv.DataSource != null && ((DataTable)dgv.DataSource).Rows.Count == 0))
            {
                MessageBox.Show($"{gridName} hiện đang trống, không có dữ liệu để xử lý!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            // 3. Kiểm tra xem có dòng nào đang được chọn hay không
            if (dgv.CurrentRow == null || dgv.CurrentRow.Index < 0)
            {
                MessageBox.Show($"Vui lòng chọn một dòng trong {gridName} để tiếp tục!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            return true;
        }

        public static bool IsComboBoxValid(ComboBox cbo, string displayName = "Dữ liệu")
        {
            // 1. Kiểm tra ComboBox có bị null về mặt khởi tạo không
            if (cbo == null)
            {
                MessageBox.Show($"{displayName} chưa được khởi tạo!", "Lỗi hệ thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            // 2. Kiểm tra DataSource (nếu bạn dùng Binding dữ liệu)
            if (cbo.DataSource == null && cbo.Items.Count == 0)
            {
                MessageBox.Show($"Danh sách {displayName} đang trống. Vui lòng kiểm tra lại kết nối dữ liệu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            // 3. Kiểm tra xem có Item nào đang được chọn không
            if (cbo.SelectedIndex == -1)
            {
                MessageBox.Show($"Vui lòng chọn một giá trị từ {displayName}!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            e.Graphics.DrawString(rowIdx, new Font("Arial", 12.0f), SystemBrushes.ControlText, headerBounds, centerFormat);
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
