namespace MPR_Managerment.Forms.ExportGUI
{
    partial class ucExportWarehouse_V2
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            tableLayoutPanel1 = new TableLayoutPanel();
            pnAction = new Panel();
            flpAction = new FlowLayoutPanel();
            label1 = new Label();
            cboProject = new ComboBox();
            label2 = new Label();
            txtSearch = new TextBox();
            label3 = new Label();
            dtpFromDate = new DateTimePicker();
            label5 = new Label();
            dtpToDate = new DateTimePicker();
            label4 = new Label();
            cboStatus = new ComboBox();
            btnSearch = new Button();
            btnRefresh = new Button();
            tableLayoutPanel2 = new TableLayoutPanel();
            groupBox1 = new GroupBox();
            dgvHisExport = new DataGridView();
            flowLayoutPanel1 = new FlowLayoutPanel();
            btnAddXK = new Button();
            btnInXK = new Button();
            btnUpdateStatus = new Button();
            btnExportExcel = new Button();
            _statusMenu = new ContextMenuStrip(components);
            tsAccept = new ToolStripMenuItem();
            tsCancel = new ToolStripMenuItem();
            tsKeep = new ToolStripMenuItem();
            tableLayoutPanel1.SuspendLayout();
            pnAction.SuspendLayout();
            flpAction.SuspendLayout();
            tableLayoutPanel2.SuspendLayout();
            groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvHisExport).BeginInit();
            flowLayoutPanel1.SuspendLayout();
            _statusMenu.SuspendLayout();
            SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.ColumnCount = 1;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.Controls.Add(pnAction, 0, 0);
            tableLayoutPanel1.Controls.Add(tableLayoutPanel2, 0, 2);
            tableLayoutPanel1.Controls.Add(flowLayoutPanel1, 0, 1);
            tableLayoutPanel1.Dock = DockStyle.Fill;
            tableLayoutPanel1.Location = new Point(0, 0);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 3;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 39F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 42F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.Size = new Size(1359, 697);
            tableLayoutPanel1.TabIndex = 0;
            // 
            // pnAction
            // 
            pnAction.Controls.Add(flpAction);
            pnAction.Dock = DockStyle.Fill;
            pnAction.Location = new Point(3, 3);
            pnAction.Name = "pnAction";
            pnAction.Size = new Size(1353, 33);
            pnAction.TabIndex = 0;
            // 
            // flpAction
            // 
            flpAction.Controls.Add(label1);
            flpAction.Controls.Add(cboProject);
            flpAction.Controls.Add(label2);
            flpAction.Controls.Add(txtSearch);
            flpAction.Controls.Add(label3);
            flpAction.Controls.Add(dtpFromDate);
            flpAction.Controls.Add(label5);
            flpAction.Controls.Add(dtpToDate);
            flpAction.Controls.Add(label4);
            flpAction.Controls.Add(cboStatus);
            flpAction.Controls.Add(btnSearch);
            flpAction.Controls.Add(btnRefresh);
            flpAction.Dock = DockStyle.Fill;
            flpAction.Location = new Point(0, 0);
            flpAction.Name = "flpAction";
            flpAction.Size = new Size(1353, 33);
            flpAction.TabIndex = 0;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Dock = DockStyle.Fill;
            label1.Location = new Point(3, 0);
            label1.Name = "label1";
            label1.Size = new Size(31, 33);
            label1.TabIndex = 6;
            label1.Text = "Kho:";
            label1.TextAlign = ContentAlignment.MiddleLeft;
            label1.Visible = false;
            // 
            // cboProject
            // 
            cboProject.Dock = DockStyle.Fill;
            cboProject.DropDownStyle = ComboBoxStyle.DropDownList;
            cboProject.Location = new Point(40, 3);
            cboProject.Name = "cboProject";
            cboProject.Size = new Size(150, 23);
            cboProject.TabIndex = 0;
            cboProject.Visible = false;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Dock = DockStyle.Fill;
            label2.Location = new Point(196, 0);
            label2.Name = "label2";
            label2.Size = new Size(59, 33);
            label2.TabIndex = 7;
            label2.Text = "Tìm kiếm:";
            label2.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // txtSearch
            // 
            txtSearch.Dock = DockStyle.Fill;
            txtSearch.Location = new Point(261, 3);
            txtSearch.Name = "txtSearch";
            txtSearch.Size = new Size(200, 23);
            txtSearch.TabIndex = 1;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Dock = DockStyle.Fill;
            label3.Location = new Point(467, 0);
            label3.Name = "label3";
            label3.Size = new Size(38, 33);
            label3.TabIndex = 8;
            label3.Text = "Ngày:";
            label3.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // dtpFromDate
            // 
            dtpFromDate.Dock = DockStyle.Fill;
            dtpFromDate.Format = DateTimePickerFormat.Short;
            dtpFromDate.Location = new Point(511, 3);
            dtpFromDate.Name = "dtpFromDate";
            dtpFromDate.Size = new Size(150, 23);
            dtpFromDate.TabIndex = 2;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Dock = DockStyle.Fill;
            label5.Location = new Point(667, 0);
            label5.Name = "label5";
            label5.Size = new Size(55, 33);
            label5.TabIndex = 9;
            label5.Text = "Tới ngày:";
            label5.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // dtpToDate
            // 
            dtpToDate.Dock = DockStyle.Fill;
            dtpToDate.Format = DateTimePickerFormat.Short;
            dtpToDate.Location = new Point(728, 3);
            dtpToDate.Name = "dtpToDate";
            dtpToDate.Size = new Size(150, 23);
            dtpToDate.TabIndex = 2;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Dock = DockStyle.Fill;
            label4.Location = new Point(884, 0);
            label4.Name = "label4";
            label4.Size = new Size(62, 33);
            label4.TabIndex = 8;
            label4.Text = "Trạng thái:";
            label4.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // cboStatus
            // 
            cboStatus.DropDownStyle = ComboBoxStyle.DropDownList;
            cboStatus.Location = new Point(952, 3);
            cboStatus.Name = "cboStatus";
            cboStatus.Size = new Size(120, 23);
            cboStatus.TabIndex = 3;
            // 
            // btnSearch
            // 
            btnSearch.Location = new Point(1078, 3);
            btnSearch.Name = "btnSearch";
            btnSearch.Size = new Size(100, 27);
            btnSearch.TabIndex = 4;
            btnSearch.Text = "Search";
            btnSearch.Click += BtnSearch_Click;
            // 
            // btnRefresh
            // 
            btnRefresh.Location = new Point(1184, 3);
            btnRefresh.Name = "btnRefresh";
            btnRefresh.Size = new Size(100, 27);
            btnRefresh.TabIndex = 5;
            btnRefresh.Text = "Refresh";
            btnRefresh.Click += BtnRefresh_Click;
            // 
            // tableLayoutPanel2
            // 
            tableLayoutPanel2.ColumnCount = 1;
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel2.Controls.Add(groupBox1, 0, 0);
            tableLayoutPanel2.Dock = DockStyle.Fill;
            tableLayoutPanel2.Location = new Point(3, 84);
            tableLayoutPanel2.Name = "tableLayoutPanel2";
            tableLayoutPanel2.RowCount = 1;
            tableLayoutPanel2.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel2.Size = new Size(1353, 610);
            tableLayoutPanel2.TabIndex = 2;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(dgvHisExport);
            groupBox1.Dock = DockStyle.Fill;
            groupBox1.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            groupBox1.Location = new Point(3, 3);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(1347, 604);
            groupBox1.TabIndex = 1;
            groupBox1.TabStop = false;
            groupBox1.Text = "Danh sách xuất kho";
            // 
            // dgvHisExport
            // 
            dgvHisExport.BackgroundColor = SystemColors.Control;
            dgvHisExport.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvHisExport.Dock = DockStyle.Fill;
            dgvHisExport.Location = new Point(3, 23);
            dgvHisExport.Name = "dgvHisExport";
            dgvHisExport.Size = new Size(1341, 578);
            dgvHisExport.TabIndex = 0;
            // 
            // flowLayoutPanel1
            // 
            flowLayoutPanel1.Controls.Add(btnAddXK);
            flowLayoutPanel1.Controls.Add(btnInXK);
            flowLayoutPanel1.Controls.Add(btnUpdateStatus);
            flowLayoutPanel1.Controls.Add(btnExportExcel);
            flowLayoutPanel1.Dock = DockStyle.Fill;
            flowLayoutPanel1.Location = new Point(3, 42);
            flowLayoutPanel1.Name = "flowLayoutPanel1";
            flowLayoutPanel1.Size = new Size(1353, 36);
            flowLayoutPanel1.TabIndex = 3;
            // 
            // btnAddXK
            // 
            btnAddXK.Location = new Point(3, 3);
            btnAddXK.Name = "btnAddXK";
            btnAddXK.Size = new Size(155, 33);
            btnAddXK.TabIndex = 0;
            btnAddXK.Text = "Thêm mới";
            btnAddXK.UseVisualStyleBackColor = true;
            btnAddXK.Click += btnAddXK_Click;
            // 
            // btnInXK
            // 
            btnInXK.Location = new Point(164, 3);
            btnInXK.Name = "btnInXK";
            btnInXK.Size = new Size(184, 33);
            btnInXK.TabIndex = 0;
            btnInXK.Text = "In";
            btnInXK.UseVisualStyleBackColor = true;
            btnInXK.Click += btnInXK_Click;
            // 
            // btnUpdateStatus
            // 
            btnUpdateStatus.Location = new Point(354, 3);
            btnUpdateStatus.Name = "btnUpdateStatus";
            btnUpdateStatus.Size = new Size(193, 33);
            btnUpdateStatus.TabIndex = 0;
            btnUpdateStatus.Text = "Cập nhật trạng thái           ⏷";
            btnUpdateStatus.UseVisualStyleBackColor = true;
            btnUpdateStatus.Click += btnUpdateStatus_Click;
            // 
            // btnExportExcel
            // 
            btnExportExcel.Location = new Point(553, 3);
            btnExportExcel.Name = "btnExportExcel";
            btnExportExcel.Size = new Size(184, 33);
            btnExportExcel.TabIndex = 0;
            btnExportExcel.Text = "Xuất lịch sử XK";
            btnExportExcel.UseVisualStyleBackColor = true;
            btnExportExcel.Click += btnExportExcel_Click;
            // 
            // _statusMenu
            // 
            _statusMenu.AutoSize = false;
            _statusMenu.BackColor = Color.DodgerBlue;
            _statusMenu.DropShadowEnabled = false;
            _statusMenu.Items.AddRange(new ToolStripItem[] { tsAccept, tsCancel, tsKeep });
            _statusMenu.Name = "contextMenuStrip1";
            _statusMenu.Size = new Size(181, 70);
            _statusMenu.ItemClicked += StatusMenu_ItemClicked;
            // 
            // tsAccept
            // 
            tsAccept.AutoSize = false;
            tsAccept.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            tsAccept.ForeColor = Color.White;
            tsAccept.Name = "tsAccept";
            tsAccept.Size = new Size(180, 22);
            tsAccept.Text = "Xác nhận";
            // 
            // tsCancel
            // 
            tsCancel.AutoSize = false;
            tsCancel.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            tsCancel.ForeColor = Color.White;
            tsCancel.Name = "tsCancel";
            tsCancel.Size = new Size(180, 22);
            tsCancel.Text = "Hủy";
            // 
            // tsKeep
            // 
            tsKeep.AutoSize = false;
            tsKeep.ForeColor = Color.White;
            tsKeep.Name = "tsKeep";
            tsKeep.Size = new Size(180, 22);
            tsKeep.Text = "Giữ";
            // 
            // ucExportWarehouse_V2
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.White;
            Controls.Add(tableLayoutPanel1);
            Name = "ucExportWarehouse_V2";
            Size = new Size(1359, 697);
            tableLayoutPanel1.ResumeLayout(false);
            pnAction.ResumeLayout(false);
            flpAction.ResumeLayout(false);
            flpAction.PerformLayout();
            tableLayoutPanel2.ResumeLayout(false);
            groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvHisExport).EndInit();
            flowLayoutPanel1.ResumeLayout(false);
            _statusMenu.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private TableLayoutPanel tableLayoutPanel1;
        private Panel pnAction;
        private FlowLayoutPanel flpAction;
        private ComboBox cboProject;
        private TextBox txtSearch;
        private DateTimePicker dtpFromDate;
        private ComboBox cboStatus;
        private Button btnSearch;
        private Button btnRefresh;
        private GroupBox groupBox1;
        private DataGridView dgvHisExport;
        private TableLayoutPanel tableLayoutPanel2;
        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
        private Label label5;
        private DateTimePicker dtpToDate;
        private FlowLayoutPanel flowLayoutPanel1;
        private Button btnAddXK;
        private Button btnInXK;
        private Button btnUpdateStatus;
        private ContextMenuStrip _statusMenu;
        private ToolStripMenuItem tsAccept;
        private ToolStripMenuItem tsCancel;
        private ToolStripMenuItem tsKeep;
        private Button btnExportExcel;
    }
}