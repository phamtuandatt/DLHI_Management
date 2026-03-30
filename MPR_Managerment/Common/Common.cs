using System;
using System.Collections.Generic;
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
    }
}
