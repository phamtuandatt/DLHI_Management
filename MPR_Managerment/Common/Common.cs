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
