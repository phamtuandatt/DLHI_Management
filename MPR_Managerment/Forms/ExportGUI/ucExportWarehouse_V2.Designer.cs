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
            tableLayoutPanel1 = new TableLayoutPanel();
            pnAction = new Panel();
            flpAction = new FlowLayoutPanel();
            cboProject = new ComboBox();
            txtSearch = new TextBox();
            dtpDate = new DateTimePicker();
            cboStatus = new ComboBox();
            btnSearch = new Button();
            btnRefresh = new Button();
            groupBox1 = new GroupBox();
            dgvHisExport = new DataGridView();
            tableLayoutPanel2 = new TableLayoutPanel();
            tableLayoutPanel1.SuspendLayout();
            pnAction.SuspendLayout();
            flpAction.SuspendLayout();
            groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvHisExport).BeginInit();
            SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.ColumnCount = 1;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.Controls.Add(pnAction, 0, 0);
            tableLayoutPanel1.Controls.Add(groupBox1, 0, 1);
            tableLayoutPanel1.Controls.Add(tableLayoutPanel2, 0, 2);
            tableLayoutPanel1.Dock = DockStyle.Fill;
            tableLayoutPanel1.Location = new Point(0, 0);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 3;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));
            tableLayoutPanel1.Size = new Size(1359, 697);
            tableLayoutPanel1.TabIndex = 0;
            // 
            // pnAction
            // 
            pnAction.Controls.Add(flpAction);
            pnAction.Dock = DockStyle.Fill;
            pnAction.Location = new Point(3, 3);
            pnAction.Name = "pnAction";
            pnAction.Size = new Size(1353, 44);
            pnAction.TabIndex = 0;
            // 
            // flpAction
            // 
            flpAction.Controls.Add(cboProject);
            flpAction.Controls.Add(txtSearch);
            flpAction.Controls.Add(dtpDate);
            flpAction.Controls.Add(cboStatus);
            flpAction.Controls.Add(btnSearch);
            flpAction.Controls.Add(btnRefresh);
            flpAction.Dock = DockStyle.Fill;
            flpAction.Location = new Point(0, 0);
            flpAction.Name = "flpAction";
            flpAction.Size = new Size(1353, 44);
            flpAction.TabIndex = 0;
            // 
            // cboProject
            // 
            cboProject.Name = "cboProject";
            cboProject.Size = new Size(150, 23);
            // 
            // txtSearch
            // 
            txtSearch.Name = "txtSearch";
            txtSearch.Size = new Size(200, 23);
            // 
            // dtpDate
            // 
            dtpDate.Name = "dtpDate";
            dtpDate.Size = new Size(150, 23);
            // 
            // cboStatus
            // 
            cboStatus.Name = "cboStatus";
            cboStatus.Size = new Size(120, 23);
            // 
            // btnSearch
            // 
            btnSearch.Name = "btnSearch";
            btnSearch.Text = "Search";
            btnSearch.Size = new Size(75, 23);
            // 
            // btnRefresh
            // 
            btnRefresh.Name = "btnRefresh";
            btnRefresh.Text = "Refresh";
            btnRefresh.Size = new Size(75, 23);
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(dgvHisExport);
            groupBox1.Dock = DockStyle.Fill;
            groupBox1.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            groupBox1.Location = new Point(3, 53);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(1353, 591);
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
            dgvHisExport.Size = new Size(1347, 565);
            dgvHisExport.TabIndex = 0;
            // 
            // tableLayoutPanel2
            // 
            tableLayoutPanel2.ColumnCount = 6;
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel2.Dock = DockStyle.Fill;
            tableLayoutPanel2.Location = new Point(3, 650);
            tableLayoutPanel2.Name = "tableLayoutPanel2";
            tableLayoutPanel2.RowCount = 1;
            tableLayoutPanel2.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel2.Size = new Size(1353, 44);
            tableLayoutPanel2.TabIndex = 2;
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
            groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvHisExport).EndInit();
            ResumeLayout(false);
        }

        #endregion

        private TableLayoutPanel tableLayoutPanel1;
        private Panel pnAction;
        private FlowLayoutPanel flpAction;
        private ComboBox cboProject;
        private TextBox txtSearch;
        private DateTimePicker dtpDate;
        private ComboBox cboStatus;
        private Button btnSearch;
        private Button btnRefresh;
        private GroupBox groupBox1;
        private DataGridView dgvHisExport;
        private TableLayoutPanel tableLayoutPanel2;
    }
}