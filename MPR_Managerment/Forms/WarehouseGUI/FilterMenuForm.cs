using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace MPR_Managerment.Forms.WarehouseGUI
{
    public class FilterMenuForm : Form
    {
        private CheckedListBox clbValues;
        private TextBox txtSearch;
        private Button btnOk, btnCancel;
        private Button btnSortAZ, btnSortZA;
        public List<string> SelectedValues { get; private set; } = new List<string>();
        public bool IsSorted { get; private set; }
        public ListSortDirection SortDirection { get; private set; }
        public bool Canceled { get; private set; } = true;

        private Color _menuBackColor = Color.FromArgb(45, 45, 48); // Dark background
        private Color _controlBackColor = Color.FromArgb(60, 60, 63); // Slightly lighter for controls
        private Color _textColor = Color.White;
        private Color _borderColor = Color.FromArgb(80, 80, 83);
        private Color _buttonBackColor = Color.FromArgb(70, 70, 73);
        private Color _buttonHoverColor = Color.FromArgb(90, 90, 93);
        private Color _buttonTextColor = Color.White;

        public FilterMenuForm(string columnName, DataTable table, List<string> currentSelected)
        {
            InitializeComponent(columnName, table, currentSelected);
            PopulateList(columnName, table, currentSelected);
        }

        private void InitializeComponent(string columnName, DataTable table, List<string> currentSelected)
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.Size = new Size(280, 400); // Increased size for better layout
            this.BackColor = _menuBackColor;
            this.StartPosition = FormStartPosition.Manual;
            this.DoubleBuffered = true; // For smoother rendering

            // TableLayoutPanel for overall layout
            var mainLayoutPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                BackColor = _menuBackColor,
                Padding = new Padding(5)
            };
            mainLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            mainLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F)); // Sort buttons
            mainLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F)); // Search box
            mainLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // CheckedListBox
            mainLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F)); // OK/Cancel buttons

            //// Sort Buttons Panel
            //var sortPanel = new Panel { Dock = DockStyle.Fill, BackColor = _menuBackColor };
            //btnSortAZ = CreateModernButton("Sort A to Z", Image.FromFile("C:\\Users\\TUAN DAT\\OneDrive\\Working\\Final_Project_DLHI_Management\\MPR_Managerment\\Properties\\sort_az_icon.png"), new Point(5, 5)); // Assuming you have icons
            //btnSortAZ.Click += (s, e) => { IsSorted = true; SortDirection = ListSortDirection.Ascending; this.Close(); };
            //btnSortZA = CreateModernButton("Sort Z to A", Image.FromFile("C:\\Users\\TUAN DAT\\OneDrive\\Working\\Final_Project_DLHI_Management\\MPR_Managerment\\Properties\\sort_za_icon.png"), new Point(90, 5)); // Assuming you have icons
            //btnSortZA.Click += (s, e) => { IsSorted = true; SortDirection = ListSortDirection.Descending; this.Close(); };
            //sortPanel.Controls.AddRange(new Control[] { btnSortAZ, btnSortZA });

            // Search TextBox
            txtSearch = new TextBox
            {
                Dock = DockStyle.Fill,
                BackColor = _controlBackColor,
                ForeColor = _textColor,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9),
                Padding = new Padding(3, 5, 0, 3),
                PlaceholderText = "Search...",
                TextAlign = HorizontalAlignment.Left,
                Top = (this.ClientSize.Height - this.Height) / 2
            };
            txtSearch.TextChanged += (s, e) => FilterList(columnName, table, currentSelected);

            // CheckedListBox
            clbValues = new CheckedListBox
            {
                Dock = DockStyle.Fill,
                CheckOnClick = true,
                BackColor = _controlBackColor,
                ForeColor = _textColor,
                BorderStyle = BorderStyle.None, // Let the panel handle border
                Font = new Font("Segoe UI", 9),
                Padding = new Padding(3)
            };
            clbValues.ItemCheck += (s, e) => {
                if (e.Index == 0) { // "Select All" checkbox
                    for (int i = 1; i < clbValues.Items.Count; i++) clbValues.SetItemChecked(i, e.NewValue == CheckState.Checked);
                }
            };

            // OK/Cancel Buttons Panel
            var pnlButtons = new Panel { Dock = DockStyle.Fill, BackColor = _menuBackColor };
            btnOk = CreateModernButton("OK", null, new Point(pnlButtons.Width / 2 - 7, 5));
            btnOk.Click += (s, e) => {
                SelectedValues = clbValues.CheckedItems.Cast<string>().ToList();
                Canceled = false;
                this.Close();
            };
            btnCancel = CreateModernButton("Cancel", null, new Point(pnlButtons.Width / 2 + 83, 5));
            btnCancel.Click += (s, e) => this.Close();
            pnlButtons.Controls.AddRange(new Control[] { btnOk, btnCancel });

            //mainLayoutPanel.Controls.Add(sortPanel, 0, 0);
            mainLayoutPanel.Controls.Add(txtSearch, 0, 1);
            mainLayoutPanel.Controls.Add(clbValues, 0, 2);
            mainLayoutPanel.Controls.Add(pnlButtons, 0, 3);

            this.Controls.Add(mainLayoutPanel);

            // Add a border to the form
            this.Paint += (sender, e) => {
                using (var pen = new Pen(_borderColor, 1))
                {
                    e.Graphics.DrawRectangle(pen, new Rectangle(0, 0, this.Width - 1, this.Height - 1));
                }
            };
        }

        private Button CreateModernButton(string text, Image? icon, Point location)
        {
            var btn = new Button
            {
                Text = $" {text}", // Add space for icon
                Image = icon,
                ImageAlign = icon != null ? ContentAlignment.MiddleLeft : ContentAlignment.MiddleCenter,
                TextImageRelation = icon != null ? TextImageRelation.ImageBeforeText : TextImageRelation.Overlay,
                BackColor = _buttonBackColor,
                ForeColor = _buttonTextColor,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 0, MouseOverBackColor = _buttonHoverColor },
                Font = new Font("Segoe UI", 9),
                Location = location,
                Size = new Size(80, 25), // Default size, adjust as needed
                Cursor = Cursors.Hand
            };
            return btn;
        }

        private void PopulateList(string columnName, DataTable table, List<string> currentSelected)
        {
            var values = table.AsEnumerable().Select(r => r[columnName]?.ToString() ?? "").Distinct().OrderBy(x => x).ToList();
            
            clbValues.Items.Clear();
            clbValues.Items.Add("(Select All)", true); // Default to selected
            
            foreach (var v in values)
            {
                clbValues.Items.Add(v, currentSelected.Contains(v));
            }
        }

        private void FilterList(string columnName, DataTable table, List<string> currentSelected)
        {
            var searchText = txtSearch.Text.ToLower();
            var values = table.AsEnumerable().Select(r => r[columnName]?.ToString() ?? "").Distinct().OrderBy(x => x).ToList();
            
            clbValues.ItemCheck -= null; // Temporarily detach event handler
            clbValues.Items.Clear();
            clbValues.Items.Add("(Select All)", true); // Add "Select All" back
            
            foreach (var v in values)
            {
                if (string.IsNullOrEmpty(searchText) || v.ToLower().Contains(searchText))
                {
                    clbValues.Items.Add(v, currentSelected.Contains(v));
                }
            }
            // Re-attach event handler if needed, or handle it in constructor
        }
    }
}