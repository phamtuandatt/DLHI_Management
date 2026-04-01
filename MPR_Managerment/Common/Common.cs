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
    }
}
