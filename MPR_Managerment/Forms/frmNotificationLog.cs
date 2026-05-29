using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using MPR_Managerment.Common;
using MPR_Managerment.Models;
using MPR_Managerment.Services;

namespace MPR_Managerment.Forms
{
    public class frmNotificationLog : frmBase
    {
        private Panel pnlFilter;
        private TextBox txtUser;
        private DateTimePicker dtpFrom;
        private DateTimePicker dtpTo;
        private Button btnSearch;
        private Button btnClear;
        private DataGridView dgvLogs;
        private NotificationLogService _logService = new();
        private System.Windows.Forms.Timer _refreshTimer;

        public frmNotificationLog()
        {
            InitializeComponent();
            SetupAutoRefresh();
        }

        private void SetupAutoRefresh()
        {
            _refreshTimer = new System.Windows.Forms.Timer();
            _refreshTimer.Interval = 60000; // 1 minute
            _refreshTimer.Tick += async (s, e) => await LoadLogsAsync(isAutoRefresh: true);
            _refreshTimer.Start();
        }

        private void InitializeComponent()
        {
            this.Text = "Lịch sử Thông báo (Email & Zalo)";
            this.Size = new Size(1100, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(245, 245, 245);

            // Filter Panel
            pnlFilter = new Panel
            {
                Dock = DockStyle.Top,
                Height = 80,
                BackColor = Color.White,
                Padding = new Padding(10)
            };

            Label lblUser = new Label { Text = "Người dùng:", Location = new Point(20, 25), Size = new Size(80, 20) };
            txtUser = new TextBox { Location = new Point(100, 22), Size = new Size(150, 25) };

            Label lblFrom = new Label { Text = "Từ ngày:", Location = new Point(270, 25), Size = new Size(60, 20) };
            dtpFrom = new DateTimePicker { Location = new Point(330, 22), Size = new Size(150, 25), Format = DateTimePickerFormat.Short };
            dtpFrom.Value = DateTime.Now.AddDays(-7);

            Label lblTo = new Label { Text = "Đến ngày:", Location = new Point(490, 25), Size = new Size(60, 20) };
            dtpTo = new DateTimePicker { Location = new Point(550, 22), Size = new Size(150, 25), Format = DateTimePickerFormat.Short };
            dtpTo.Value = DateTime.Now;

            btnSearch = new Button
            {
                Text = "🔍 Tìm kiếm",
                Location = new Point(720, 20),
                Size = new Size(100, 30),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnSearch.Click += async (s, e) => await LoadLogsAsync();

            btnClear = new Button
            {
                Text = "🔄 Xóa lọc",
                Location = new Point(830, 20),
                Size = new Size(100, 30),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnClear.Click += async (s, e) => {
                txtUser.Clear();
                dtpFrom.Value = DateTime.Now.AddDays(-7);
                dtpTo.Value = DateTime.Now;
                await LoadLogsAsync();
            };

            pnlFilter.Controls.AddRange(new Control[] { lblUser, txtUser, lblFrom, dtpFrom, lblTo, dtpTo, btnSearch, btnClear });

            // Data Grid
            dgvLogs = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                EnableHeadersVisualStyles = false
            };
            dgvLogs.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);
            dgvLogs.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvLogs.ColumnHeadersHeight = 35;

            this.Controls.Add(dgvLogs);
            this.Controls.Add(pnlFilter);

            // Initial load
            this.Load += async (s, e) => await LoadLogsAsync();
        }

        private async Task LoadLogsAsync(bool isAutoRefresh = false)
        {
            try
            {
                string user = txtUser.Text;
                DateTime from = dtpFrom.Value;
                DateTime to = dtpTo.Value;

                // Execute query in background
                var logs = await Task.Run(() => _logService.GetLogs(user, from, to));

                if (this.IsDisposed) return;

                if (this.IsHandleCreated && !this.IsDisposed)
                {
                    this.Invoke(new MethodInvoker(() =>
                    {
                        if (this.IsDisposed) return;
                        int scrollIdx = dgvLogs.FirstDisplayedScrollingRowIndex;

                        dgvLogs.DataSource = null;
                        dgvLogs.DataSource = logs;

                        if (dgvLogs.Columns["Log_ID"] != null) dgvLogs.Columns["Log_ID"].Visible = false;

                        // Rename columns and fix widths
                        dgvLogs.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

                        if (dgvLogs.Columns["Sent_At"] != null) { dgvLogs.Columns["Sent_At"].HeaderText = "Thời gian"; dgvLogs.Columns["Sent_At"].Width = 140; }
                        if (dgvLogs.Columns["Sent_By"] != null) { dgvLogs.Columns["Sent_By"].HeaderText = "Người gửi"; dgvLogs.Columns["Sent_By"].Width = 150; }
                        if (dgvLogs.Columns["Recipient"] != null) { dgvLogs.Columns["Recipient"].HeaderText = "Người nhận/Nhóm"; dgvLogs.Columns["Recipient"].Width = 150; }
                        if (dgvLogs.Columns["Type"] != null) { dgvLogs.Columns["Type"].HeaderText = "Loại"; dgvLogs.Columns["Type"].Width = 60; }
                        if (dgvLogs.Columns["Content"] != null) { dgvLogs.Columns["Content"].HeaderText = "Nội dung"; dgvLogs.Columns["Content"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; }
                        if (dgvLogs.Columns["Status"] != null) { dgvLogs.Columns["Status"].HeaderText = "Trạng thái"; dgvLogs.Columns["Status"].Width = 80; }
                        if (dgvLogs.Columns["Error_Message"] != null) { dgvLogs.Columns["Error_Message"].HeaderText = "Lỗi"; dgvLogs.Columns["Error_Message"].Width = 100; }
                        if (dgvLogs.Columns["Project_Code"] != null) { dgvLogs.Columns["Project_Code"].HeaderText = "Mã dự án"; dgvLogs.Columns["Project_Code"].Width = 80; }

                        if (scrollIdx >= 0 && scrollIdx < dgvLogs.RowCount)
                        {
                            dgvLogs.FirstDisplayedScrollingRowIndex = scrollIdx;
                        }
                    }));
                }
            }
            catch (Exception ex)
            {
                if (!isAutoRefresh)
                {
                    MessageBox.Show($"Lỗi khi tải lịch sử: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _refreshTimer?.Stop();
            _refreshTimer?.Dispose();
            base.OnFormClosing(e);
        }
    }
}
