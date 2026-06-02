namespace MPR_Managerment.Forms.WarehouseGUI
{
    partial class ucImportWarehouse
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
            groupBox1 = new GroupBox();
            tableLayoutPanel2 = new TableLayoutPanel();
            label1 = new Label();
            cboProjectForImport = new ComboBox();
            label2 = new Label();
            cboPONoForImport = new ComboBox();
            btnSearchItemPO = new Button();
            btnRefresh = new Button();
            btnDeleteRow = new Button();
            btnSaveImport = new Button();
            groupBox2 = new GroupBox();
            dgvImportQueue = new DataGridView();
            groupBox3 = new GroupBox();
            tableLayoutPanel1.SuspendLayout();
            groupBox1.SuspendLayout();
            tableLayoutPanel2.SuspendLayout();
            groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvImportQueue).BeginInit();
            SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.ColumnCount = 1;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20F));
            tableLayoutPanel1.Controls.Add(groupBox1, 0, 0);
            tableLayoutPanel1.Controls.Add(groupBox2, 0, 1);
            tableLayoutPanel1.Controls.Add(groupBox3, 0, 2);
            tableLayoutPanel1.Dock = DockStyle.Fill;
            tableLayoutPanel1.Location = new Point(0, 0);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 4;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 63F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));
            tableLayoutPanel1.Size = new Size(1267, 707);
            tableLayoutPanel1.TabIndex = 0;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(tableLayoutPanel2);
            groupBox1.Dock = DockStyle.Fill;
            groupBox1.Location = new Point(3, 3);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(1261, 57);
            groupBox1.TabIndex = 0;
            groupBox1.TabStop = false;
            groupBox1.Text = "Thông tin dự án";
            // 
            // tableLayoutPanel2
            // 
            tableLayoutPanel2.ColumnCount = 9;
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 215F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel2.Controls.Add(label1, 0, 0);
            tableLayoutPanel2.Controls.Add(cboProjectForImport, 1, 0);
            tableLayoutPanel2.Controls.Add(label2, 2, 0);
            tableLayoutPanel2.Controls.Add(cboPONoForImport, 3, 0);
            tableLayoutPanel2.Controls.Add(btnSearchItemPO, 4, 0);
            tableLayoutPanel2.Controls.Add(btnRefresh, 5, 0);
            tableLayoutPanel2.Controls.Add(btnDeleteRow, 7, 0);
            tableLayoutPanel2.Controls.Add(btnSaveImport, 8, 0);
            tableLayoutPanel2.Dock = DockStyle.Fill;
            tableLayoutPanel2.Location = new Point(3, 19);
            tableLayoutPanel2.Name = "tableLayoutPanel2";
            tableLayoutPanel2.RowCount = 1;
            tableLayoutPanel2.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel2.Size = new Size(1255, 35);
            tableLayoutPanel2.TabIndex = 0;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Dock = DockStyle.Fill;
            label1.Location = new Point(3, 0);
            label1.Name = "label1";
            label1.Size = new Size(64, 35);
            label1.TabIndex = 0;
            label1.Text = "Dự án:";
            label1.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // cboProjectForImport
            // 
            cboProjectForImport.Dock = DockStyle.Fill;
            cboProjectForImport.FormattingEnabled = true;
            cboProjectForImport.Location = new Point(73, 3);
            cboProjectForImport.Name = "cboProjectForImport";
            cboProjectForImport.Size = new Size(154, 23);
            cboProjectForImport.TabIndex = 1;
            cboProjectForImport.SelectedIndexChanged += cboProjectForImport_SelectedIndexChanged;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Dock = DockStyle.Fill;
            label2.Location = new Point(233, 0);
            label2.Name = "label2";
            label2.Size = new Size(64, 35);
            label2.TabIndex = 2;
            label2.Text = "PO No:";
            label2.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // cboPONoForImport
            // 
            cboPONoForImport.Dock = DockStyle.Fill;
            cboPONoForImport.FormattingEnabled = true;
            cboPONoForImport.Location = new Point(303, 3);
            cboPONoForImport.Name = "cboPONoForImport";
            cboPONoForImport.Size = new Size(209, 23);
            cboPONoForImport.TabIndex = 3;
            cboPONoForImport.TextUpdate += cboPONoForImport_TextUpdate;
            // 
            // btnSearchItemPO
            // 
            btnSearchItemPO.Dock = DockStyle.Fill;
            btnSearchItemPO.FlatStyle = FlatStyle.Flat;
            btnSearchItemPO.Location = new Point(518, 3);
            btnSearchItemPO.Name = "btnSearchItemPO";
            btnSearchItemPO.Size = new Size(94, 29);
            btnSearchItemPO.TabIndex = 4;
            btnSearchItemPO.Text = "Search";
            btnSearchItemPO.UseVisualStyleBackColor = true;
            btnSearchItemPO.Click += btnSearchItemPO_Click;
            // 
            // btnRefresh
            // 
            btnRefresh.Dock = DockStyle.Fill;
            btnRefresh.FlatStyle = FlatStyle.Flat;
            btnRefresh.Location = new Point(618, 3);
            btnRefresh.Name = "btnRefresh";
            btnRefresh.Size = new Size(94, 29);
            btnRefresh.TabIndex = 4;
            btnRefresh.Text = "Refresh";
            btnRefresh.UseVisualStyleBackColor = true;
            // 
            // btnDeleteRow
            // 
            btnDeleteRow.Dock = DockStyle.Fill;
            btnDeleteRow.FlatStyle = FlatStyle.Flat;
            btnDeleteRow.Location = new Point(1058, 3);
            btnDeleteRow.Name = "btnDeleteRow";
            btnDeleteRow.Size = new Size(94, 29);
            btnDeleteRow.TabIndex = 4;
            btnDeleteRow.Text = "Delete";
            btnDeleteRow.UseVisualStyleBackColor = true;
            btnDeleteRow.Click += btnDeleteRow_Click;
            // 
            // btnSaveImport
            // 
            btnSaveImport.Dock = DockStyle.Fill;
            btnSaveImport.FlatStyle = FlatStyle.Flat;
            btnSaveImport.Location = new Point(1158, 3);
            btnSaveImport.Name = "btnSaveImport";
            btnSaveImport.Size = new Size(94, 29);
            btnSaveImport.TabIndex = 4;
            btnSaveImport.Text = "Save";
            btnSaveImport.UseVisualStyleBackColor = true;
            btnSaveImport.Click += btnSaveImport_Click;
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(dgvImportQueue);
            groupBox2.Dock = DockStyle.Fill;
            groupBox2.Location = new Point(3, 66);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new Size(1261, 291);
            groupBox2.TabIndex = 0;
            groupBox2.TabStop = false;
            groupBox2.Text = "Thông tin dự án";
            // 
            // dgvImportQueue
            // 
            dgvImportQueue.BackgroundColor = Color.White;
            dgvImportQueue.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvImportQueue.Dock = DockStyle.Fill;
            dgvImportQueue.Location = new Point(3, 19);
            dgvImportQueue.Name = "dgvImportQueue";
            dgvImportQueue.Size = new Size(1255, 269);
            dgvImportQueue.TabIndex = 0;
            // 
            // groupBox3
            // 
            groupBox3.Dock = DockStyle.Fill;
            groupBox3.Location = new Point(3, 363);
            groupBox3.Name = "groupBox3";
            groupBox3.Size = new Size(1261, 291);
            groupBox3.TabIndex = 0;
            groupBox3.TabStop = false;
            groupBox3.Text = "Thông tin dự án";
            // 
            // ucImportWarehouse
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.White;
            Controls.Add(tableLayoutPanel1);
            Name = "ucImportWarehouse";
            Size = new Size(1267, 707);
            Load += ucImportWarehouse_Load;
            tableLayoutPanel1.ResumeLayout(false);
            groupBox1.ResumeLayout(false);
            tableLayoutPanel2.ResumeLayout(false);
            tableLayoutPanel2.PerformLayout();
            groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvImportQueue).EndInit();
            ResumeLayout(false);
        }

        #endregion

        private TableLayoutPanel tableLayoutPanel1;
        private GroupBox groupBox1;
        private TableLayoutPanel tableLayoutPanel2;
        private Label label1;
        private ComboBox cboProjectForImport;
        private Label label2;
        private ComboBox cboPONoForImport;
        private GroupBox groupBox2;
        private GroupBox groupBox3;
        private Button btnSearchItemPO;
        private Button btnRefresh;
        private DataGridView dgvImportQueue;
        private Button btnDeleteRow;
        private Button btnSaveImport;
    }
}
