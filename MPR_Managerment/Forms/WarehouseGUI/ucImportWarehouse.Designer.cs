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
            gbPrintImportNotes = new GroupBox();
            groupBox3 = new GroupBox();
            dgvHisImport = new DataGridView();
            tableLayoutPanel3 = new TableLayoutPanel();
            label3 = new Label();
            cboProjectForHisImport = new ComboBox();
            label4 = new Label();
            dtpFrom = new DateTimePicker();
            label5 = new Label();
            dtpTo = new DateTimePicker();
            btnSearchHisImport = new Button();
            btnRefeshHisImport = new Button();
            btnPrintImportHis = new Button();
            tableLayoutPanel1.SuspendLayout();
            groupBox1.SuspendLayout();
            tableLayoutPanel2.SuspendLayout();
            groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvImportQueue).BeginInit();
            gbPrintImportNotes.SuspendLayout();
            groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvHisImport).BeginInit();
            tableLayoutPanel3.SuspendLayout();
            SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.ColumnCount = 1;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.Controls.Add(groupBox1, 0, 0);
            tableLayoutPanel1.Controls.Add(groupBox2, 0, 1);
            tableLayoutPanel1.Controls.Add(gbPrintImportNotes, 0, 2);
            tableLayoutPanel1.Dock = DockStyle.Fill;
            tableLayoutPanel1.Location = new Point(0, 0);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 4;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 65F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 350F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 700F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.Size = new Size(1267, 1141);
            tableLayoutPanel1.TabIndex = 0;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(tableLayoutPanel2);
            groupBox1.Dock = DockStyle.Fill;
            groupBox1.Location = new Point(3, 3);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(1261, 59);
            groupBox1.TabIndex = 0;
            groupBox1.TabStop = false;
            groupBox1.Text = "Thông tin dự án";
            // 
            // tableLayoutPanel2
            // 
            tableLayoutPanel2.ColumnCount = 11;
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 215F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20F));
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
            tableLayoutPanel2.Size = new Size(1255, 37);
            tableLayoutPanel2.TabIndex = 0;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Dock = DockStyle.Fill;
            label1.Location = new Point(3, 0);
            label1.Name = "label1";
            label1.Size = new Size(64, 37);
            label1.TabIndex = 0;
            label1.Text = "Dự án:";
            label1.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // cboProjectForImport
            // 
            cboProjectForImport.Dock = DockStyle.Fill;
            cboProjectForImport.FormattingEnabled = true;
            cboProjectForImport.Location = new Point(73, 7);
            cboProjectForImport.Margin = new Padding(3, 7, 3, 3);
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
            label2.Size = new Size(64, 37);
            label2.TabIndex = 2;
            label2.Text = "PO No:";
            label2.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // cboPONoForImport
            // 
            cboPONoForImport.Dock = DockStyle.Fill;
            cboPONoForImport.FormattingEnabled = true;
            cboPONoForImport.Location = new Point(303, 7);
            cboPONoForImport.Margin = new Padding(3, 7, 3, 3);
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
            btnSearchItemPO.Size = new Size(94, 31);
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
            btnRefresh.Size = new Size(94, 31);
            btnRefresh.TabIndex = 4;
            btnRefresh.Text = "Refresh";
            btnRefresh.UseVisualStyleBackColor = true;
            btnRefresh.Click += btnRefresh_Click;
            // 
            // btnDeleteRow
            // 
            btnDeleteRow.Dock = DockStyle.Fill;
            btnDeleteRow.FlatStyle = FlatStyle.Flat;
            btnDeleteRow.Location = new Point(1018, 3);
            btnDeleteRow.Name = "btnDeleteRow";
            btnDeleteRow.Size = new Size(94, 31);
            btnDeleteRow.TabIndex = 4;
            btnDeleteRow.Text = "Delete";
            btnDeleteRow.UseVisualStyleBackColor = true;
            btnDeleteRow.Click += btnDeleteRow_Click;
            // 
            // btnSaveImport
            // 
            btnSaveImport.Dock = DockStyle.Fill;
            btnSaveImport.FlatStyle = FlatStyle.Flat;
            btnSaveImport.Location = new Point(1118, 3);
            btnSaveImport.Name = "btnSaveImport";
            btnSaveImport.Size = new Size(94, 31);
            btnSaveImport.TabIndex = 4;
            btnSaveImport.Text = "Save";
            btnSaveImport.UseVisualStyleBackColor = true;
            btnSaveImport.Click += btnSaveImport_Click;
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(dgvImportQueue);
            groupBox2.Dock = DockStyle.Fill;
            groupBox2.Location = new Point(3, 68);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new Size(1261, 344);
            groupBox2.TabIndex = 0;
            groupBox2.TabStop = false;
            groupBox2.Text = "Thông tin phiếu xuất kho";
            // 
            // dgvImportQueue
            // 
            dgvImportQueue.BackgroundColor = Color.White;
            dgvImportQueue.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvImportQueue.Dock = DockStyle.Fill;
            dgvImportQueue.Location = new Point(3, 19);
            dgvImportQueue.Name = "dgvImportQueue";
            dgvImportQueue.Size = new Size(1255, 322);
            dgvImportQueue.TabIndex = 0;
            // 
            // gbPrintImportNotes
            // 
            gbPrintImportNotes.Controls.Add(groupBox3);
            gbPrintImportNotes.Controls.Add(tableLayoutPanel3);
            gbPrintImportNotes.Dock = DockStyle.Fill;
            gbPrintImportNotes.Location = new Point(3, 418);
            gbPrintImportNotes.Name = "gbPrintImportNotes";
            gbPrintImportNotes.Size = new Size(1261, 694);
            gbPrintImportNotes.TabIndex = 0;
            gbPrintImportNotes.TabStop = false;
            gbPrintImportNotes.Text = "Thông tin dự án";
            // 
            // groupBox3
            // 
            groupBox3.Controls.Add(dgvHisImport);
            groupBox3.Dock = DockStyle.Fill;
            groupBox3.Location = new Point(3, 56);
            groupBox3.Name = "groupBox3";
            groupBox3.Size = new Size(1255, 635);
            groupBox3.TabIndex = 1;
            groupBox3.TabStop = false;
            groupBox3.Text = "Lịch sử nhập kho";
            // 
            // dgvHisImport
            // 
            dgvHisImport.BackgroundColor = Color.White;
            dgvHisImport.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvHisImport.Dock = DockStyle.Fill;
            dgvHisImport.Location = new Point(3, 19);
            dgvHisImport.Name = "dgvHisImport";
            dgvHisImport.Size = new Size(1249, 613);
            dgvHisImport.TabIndex = 0;
            // 
            // tableLayoutPanel3
            // 
            tableLayoutPanel3.ColumnCount = 11;
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 143F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20F));
            tableLayoutPanel3.Controls.Add(label3, 0, 0);
            tableLayoutPanel3.Controls.Add(cboProjectForHisImport, 1, 0);
            tableLayoutPanel3.Controls.Add(label4, 2, 0);
            tableLayoutPanel3.Controls.Add(dtpFrom, 3, 0);
            tableLayoutPanel3.Controls.Add(label5, 4, 0);
            tableLayoutPanel3.Controls.Add(dtpTo, 5, 0);
            tableLayoutPanel3.Controls.Add(btnSearchHisImport, 6, 0);
            tableLayoutPanel3.Controls.Add(btnRefeshHisImport, 7, 0);
            tableLayoutPanel3.Controls.Add(btnPrintImportHis, 8, 0);
            tableLayoutPanel3.Dock = DockStyle.Top;
            tableLayoutPanel3.Location = new Point(3, 19);
            tableLayoutPanel3.Name = "tableLayoutPanel3";
            tableLayoutPanel3.RowCount = 1;
            tableLayoutPanel3.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel3.Size = new Size(1255, 37);
            tableLayoutPanel3.TabIndex = 0;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Dock = DockStyle.Fill;
            label3.Location = new Point(3, 0);
            label3.Name = "label3";
            label3.Size = new Size(64, 37);
            label3.TabIndex = 0;
            label3.Text = "Dự án:";
            label3.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // cboProjectForHisImport
            // 
            cboProjectForHisImport.Dock = DockStyle.Fill;
            cboProjectForHisImport.FormattingEnabled = true;
            cboProjectForHisImport.Location = new Point(73, 7);
            cboProjectForHisImport.Margin = new Padding(3, 7, 3, 3);
            cboProjectForHisImport.Name = "cboProjectForHisImport";
            cboProjectForHisImport.Size = new Size(154, 23);
            cboProjectForHisImport.TabIndex = 1;
            cboProjectForHisImport.SelectedIndexChanged += cboProjectForImport_SelectedIndexChanged;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Dock = DockStyle.Fill;
            label4.Location = new Point(233, 0);
            label4.Name = "label4";
            label4.Size = new Size(64, 37);
            label4.TabIndex = 2;
            label4.Text = "From";
            label4.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // dtpFrom
            // 
            dtpFrom.CustomFormat = "dd/MM/yyyy";
            dtpFrom.Dock = DockStyle.Fill;
            dtpFrom.Format = DateTimePickerFormat.Short;
            dtpFrom.Location = new Point(303, 7);
            dtpFrom.Margin = new Padding(3, 7, 3, 3);
            dtpFrom.Name = "dtpFrom";
            dtpFrom.Size = new Size(154, 23);
            dtpFrom.TabIndex = 5;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Dock = DockStyle.Fill;
            label5.Location = new Point(463, 0);
            label5.Name = "label5";
            label5.Size = new Size(64, 37);
            label5.TabIndex = 2;
            label5.Text = "To:";
            label5.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // dtpTo
            // 
            dtpTo.CustomFormat = "dd/MM/yyyy";
            dtpTo.Dock = DockStyle.Fill;
            dtpTo.Format = DateTimePickerFormat.Short;
            dtpTo.Location = new Point(533, 7);
            dtpTo.Margin = new Padding(3, 7, 3, 3);
            dtpTo.Name = "dtpTo";
            dtpTo.Size = new Size(154, 23);
            dtpTo.TabIndex = 6;
            // 
            // btnSearchHisImport
            // 
            btnSearchHisImport.Dock = DockStyle.Fill;
            btnSearchHisImport.FlatStyle = FlatStyle.Flat;
            btnSearchHisImport.Location = new Point(693, 3);
            btnSearchHisImport.Name = "btnSearchHisImport";
            btnSearchHisImport.Size = new Size(94, 31);
            btnSearchHisImport.TabIndex = 4;
            btnSearchHisImport.Text = "Search";
            btnSearchHisImport.UseVisualStyleBackColor = true;
            btnSearchHisImport.Click += btnSearchHisImport_Click;
            // 
            // btnRefeshHisImport
            // 
            btnRefeshHisImport.Dock = DockStyle.Fill;
            btnRefeshHisImport.FlatStyle = FlatStyle.Flat;
            btnRefeshHisImport.Location = new Point(793, 3);
            btnRefeshHisImport.Name = "btnRefeshHisImport";
            btnRefeshHisImport.Size = new Size(94, 31);
            btnRefeshHisImport.TabIndex = 4;
            btnRefeshHisImport.Text = "Refresh";
            btnRefeshHisImport.UseVisualStyleBackColor = true;
            btnRefeshHisImport.Click += btnRefeshHisImport_Click;
            // 
            // btnPrintImportHis
            // 
            btnPrintImportHis.Dock = DockStyle.Fill;
            btnPrintImportHis.FlatStyle = FlatStyle.Flat;
            btnPrintImportHis.Location = new Point(893, 3);
            btnPrintImportHis.Name = "btnPrintImportHis";
            btnPrintImportHis.Size = new Size(137, 31);
            btnPrintImportHis.TabIndex = 4;
            btnPrintImportHis.Text = "Print";
            btnPrintImportHis.UseVisualStyleBackColor = true;
            btnPrintImportHis.Click += btnPrintImportHis_Click;
            // 
            // ucImportWarehouse
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.White;
            Controls.Add(tableLayoutPanel1);
            Name = "ucImportWarehouse";
            Size = new Size(1267, 1141);
            Load += ucImportWarehouse_Load;
            tableLayoutPanel1.ResumeLayout(false);
            groupBox1.ResumeLayout(false);
            tableLayoutPanel2.ResumeLayout(false);
            tableLayoutPanel2.PerformLayout();
            groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvImportQueue).EndInit();
            gbPrintImportNotes.ResumeLayout(false);
            groupBox3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvHisImport).EndInit();
            tableLayoutPanel3.ResumeLayout(false);
            tableLayoutPanel3.PerformLayout();
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
        private GroupBox gbPrintImportNotes;
        private Button btnSearchItemPO;
        private Button btnRefresh;
        private DataGridView dgvImportQueue;
        private Button btnDeleteRow;
        private Button btnSaveImport;
        private GroupBox groupBox3;
        private TableLayoutPanel tableLayoutPanel3;
        private Label label3;
        private ComboBox cboProjectForHisImport;
        private Label label4;
        private Button btnSearchHisImport;
        private Button btnRefeshHisImport;
        private DateTimePicker dtpFrom;
        private Label label5;
        private DateTimePicker dtpTo;
        private Button btnPrintImportHis;
        private DataGridView dgvHisImport;
    }
}
