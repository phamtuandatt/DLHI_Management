using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace MPR_Managerment.Forms
{
    /// <summary>Dialog chọn file PDF từ thư mục để thêm vào danh sách phân loại.</summary>
    public class frmPickFiles : Form
    {
        public List<string> SelectedFiles { get; private set; } = new();

        private readonly List<string> _allFiles;
        private CheckedListBox clbFiles = null!;
        private Button btnOK = null!;
        private Button btnCancel = null!;
        private Button btnSelectAll = null!;
        private Button btnDeselectAll = null!;
        private Label lblInfo = null!;

        public frmPickFiles(List<string> files)
        {
            _allFiles = files;
            InitializeComponent();
            PopulateList();
        }

        private void InitializeComponent()
        {
            Text = "Chọn file để phân loại";
            Size = new Size(560, 500);
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            FormBorderStyle = FormBorderStyle.FixedDialog;

            var pnlTop = new Panel { Dock = DockStyle.Top, Height = 40, Padding = new Padding(8, 8, 8, 0) };
            lblInfo = new Label { Location = new Point(8, 12), AutoSize = true, ForeColor = Color.DimGray };
            pnlTop.Controls.Add(lblInfo);

            clbFiles = new CheckedListBox
            {
                Dock = DockStyle.Fill,
                CheckOnClick = true,
                Font = new Font("Segoe UI", 9),
                IntegralHeight = false
            };
            clbFiles.ItemCheck += (s, e) => UpdateInfo();

            var pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 48, Padding = new Padding(8, 8, 8, 8) };

            btnSelectAll = new Button { Text = "Chọn tất cả", Location = new Point(8, 10), Width = 100, Height = 28 };
            btnSelectAll.Click += (s, e) => { for (int i = 0; i < clbFiles.Items.Count; i++) clbFiles.SetItemChecked(i, true); };

            btnDeselectAll = new Button { Text = "Bỏ chọn tất cả", Location = new Point(116, 10), Width = 110, Height = 28 };
            btnDeselectAll.Click += (s, e) => { for (int i = 0; i < clbFiles.Items.Count; i++) clbFiles.SetItemChecked(i, false); };

            btnOK = new Button
            {
                Text = "✔ Thêm vào danh sách",
                Location = new Point(270, 10), Width = 160, Height = 28,
                BackColor = Color.FromArgb(39, 174, 96), ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat, DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;

            btnCancel = new Button
            {
                Text = "Hủy", Location = new Point(440, 10), Width = 80, Height = 28,
                DialogResult = DialogResult.Cancel
            };

            pnlBottom.Controls.AddRange(new Control[] { btnSelectAll, btnDeselectAll, btnOK, btnCancel });

            Controls.Add(clbFiles);
            Controls.Add(pnlTop);
            Controls.Add(pnlBottom);

            AcceptButton = btnOK;
            CancelButton = btnCancel;
        }

        private void PopulateList()
        {
            clbFiles.Items.Clear();
            foreach (var f in _allFiles)
                clbFiles.Items.Add(Path.GetFileName(f), false);
            UpdateInfo();
        }

        private void UpdateInfo()
        {
            // ItemCheck fires before the check state is updated — read after via BeginInvoke only when handle exists
            void doUpdate()
            {
                int checked_ = clbFiles.CheckedItems.Count;
                lblInfo.Text = $"{_allFiles.Count} file trong thư mục — đã chọn: {checked_}";
            }

            if (IsHandleCreated)
                BeginInvoke(doUpdate);
            else
                doUpdate();
        }

        private void BtnOK_Click(object? sender, EventArgs e)
        {
            SelectedFiles = clbFiles.CheckedIndices
                .Cast<int>()
                .Select(i => _allFiles[i])
                .ToList();

            if (SelectedFiles.Count == 0)
            {
                MessageBox.Show("Chưa chọn file nào.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }
            DialogResult = DialogResult.OK;
        }
    }
}
