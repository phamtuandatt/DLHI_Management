namespace MPR_Managerment.Forms.ImportWarehouseGUI
{
    partial class ucFillInvoiceNo
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
            panel1 = new Panel();
            tableLayoutPanel1 = new TableLayoutPanel();
            groupBox2 = new GroupBox();
            panel3 = new Panel();
            dgvList = new DataGridView();
            panel2 = new Panel();
            btnFill = new Button();
            dtDate = new DateTimePicker();
            txtInvoiceNo = new TextBox();
            btnClear = new Button();
            label3 = new Label();
            lblStatus = new Label();
            btnSaveInvoice = new Button();
            groupBox3 = new GroupBox();
            panel4 = new Panel();
            dgvInvoices = new DataGridView();
            panel5 = new Panel();
            txtInvoiceSearch = new TextBox();
            label5 = new Label();
            groupBox1 = new GroupBox();
            btnSearch = new Button();
            cboPO = new ComboBox();
            label1 = new Label();
            cboProject = new ComboBox();
            label2 = new Label();
            panel1.SuspendLayout();
            tableLayoutPanel1.SuspendLayout();
            groupBox2.SuspendLayout();
            panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvList).BeginInit();
            panel2.SuspendLayout();
            groupBox3.SuspendLayout();
            panel4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvInvoices).BeginInit();
            panel5.SuspendLayout();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // panel1
            // 
            panel1.BackColor = Color.White;
            panel1.Controls.Add(tableLayoutPanel1);
            panel1.Controls.Add(groupBox1);
            panel1.Dock = DockStyle.Fill;
            panel1.Location = new Point(0, 0);
            panel1.Name = "panel1";
            panel1.Size = new Size(1149, 658);
            panel1.TabIndex = 0;
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.ColumnCount = 1;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tableLayoutPanel1.Controls.Add(groupBox2, 0, 0);
            tableLayoutPanel1.Controls.Add(groupBox3, 0, 1);
            tableLayoutPanel1.Dock = DockStyle.Fill;
            tableLayoutPanel1.Location = new Point(0, 43);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 2;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 40.8130074F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 59.1869926F));
            tableLayoutPanel1.Size = new Size(1149, 615);
            tableLayoutPanel1.TabIndex = 1;
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(panel3);
            groupBox2.Controls.Add(panel2);
            groupBox2.Dock = DockStyle.Fill;
            groupBox2.Location = new Point(3, 3);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new Size(1143, 245);
            groupBox2.TabIndex = 0;
            groupBox2.TabStop = false;
            // 
            // panel3
            // 
            panel3.Controls.Add(dgvList);
            panel3.Dock = DockStyle.Fill;
            panel3.Location = new Point(3, 48);
            panel3.Name = "panel3";
            panel3.Size = new Size(1137, 194);
            panel3.TabIndex = 8;
            // 
            // dgvList
            // 
            dgvList.AllowUserToAddRows = false;
            dgvList.AllowUserToDeleteRows = false;
            dgvList.AllowUserToOrderColumns = true;
            dgvList.BackgroundColor = Color.White;
            dgvList.BorderStyle = BorderStyle.None;
            dgvList.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Sunken;
            dgvList.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvList.Dock = DockStyle.Fill;
            dgvList.Location = new Point(0, 0);
            dgvList.Margin = new Padding(3, 2, 3, 2);
            dgvList.Name = "dgvList";
            dgvList.ReadOnly = true;
            dgvList.RowHeadersWidth = 51;
            dgvList.Size = new Size(1137, 194);
            dgvList.TabIndex = 6;
            dgvList.CellClick += dgvList_CellClick;
            dgvList.CellContentClick += dgvList_CellContentClick;
            dgvList.EditingControlShowing += dgvList_EditingControlShowing;
            dgvList.Scroll += dgvList_Scroll;
            // 
            // panel2
            // 
            panel2.Controls.Add(btnFill);
            panel2.Controls.Add(dtDate);
            panel2.Controls.Add(txtInvoiceNo);
            panel2.Controls.Add(btnClear);
            panel2.Controls.Add(label3);
            panel2.Controls.Add(lblStatus);
            panel2.Controls.Add(btnSaveInvoice);
            panel2.Dock = DockStyle.Top;
            panel2.Location = new Point(3, 19);
            panel2.Name = "panel2";
            panel2.Size = new Size(1137, 29);
            panel2.TabIndex = 7;
            // 
            // btnFill
            // 
            btnFill.BackColor = Color.LimeGreen;
            btnFill.Dock = DockStyle.Left;
            btnFill.FlatStyle = FlatStyle.Flat;
            btnFill.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnFill.ForeColor = Color.White;
            btnFill.Location = new Point(0, 0);
            btnFill.Margin = new Padding(3, 2, 3, 2);
            btnFill.Name = "btnFill";
            btnFill.Size = new Size(76, 29);
            btnFill.TabIndex = 11;
            btnFill.Text = "🔄 Fill";
            btnFill.UseVisualStyleBackColor = false;
            btnFill.Click += btnFill_Click;
            // 
            // dtDate
            // 
            dtDate.Format = DateTimePickerFormat.Short;
            dtDate.Location = new Point(436, 1);
            dtDate.Name = "dtDate";
            dtDate.Size = new Size(140, 23);
            dtDate.TabIndex = 10;
            // 
            // txtInvoiceNo
            // 
            txtInvoiceNo.Location = new Point(179, 1);
            txtInvoiceNo.Margin = new Padding(3, 2, 3, 2);
            txtInvoiceNo.Name = "txtInvoiceNo";
            txtInvoiceNo.Size = new Size(167, 23);
            txtInvoiceNo.TabIndex = 9;
            // 
            // btnClear
            // 
            btnClear.BackColor = Color.IndianRed;
            btnClear.Dock = DockStyle.Right;
            btnClear.FlatStyle = FlatStyle.Flat;
            btnClear.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnClear.ForeColor = Color.White;
            btnClear.Location = new Point(887, 0);
            btnClear.Margin = new Padding(3, 2, 3, 2);
            btnClear.Name = "btnClear";
            btnClear.Size = new Size(101, 29);
            btnClear.TabIndex = 8;
            btnClear.Text = "🔄 Làm mới";
            btnClear.UseVisualStyleBackColor = false;
            btnClear.Click += btnClear_Click;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.BackColor = Color.Transparent;
            label3.Location = new Point(355, 5);
            label3.Name = "label3";
            label3.Size = new Size(75, 15);
            label3.TabIndex = 6;
            label3.Text = "Invoice Date:";
            // 
            // lblStatus
            // 
            lblStatus.AutoSize = true;
            lblStatus.Location = new Point(96, 5);
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(67, 15);
            lblStatus.TabIndex = 6;
            lblStatus.Text = "Invoice No:";
            lblStatus.Visible = false;
            // 
            // btnSaveInvoice
            // 
            btnSaveInvoice.BackColor = Color.FromArgb(52, 152, 219);
            btnSaveInvoice.Dock = DockStyle.Right;
            btnSaveInvoice.FlatStyle = FlatStyle.Flat;
            btnSaveInvoice.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnSaveInvoice.ForeColor = Color.White;
            btnSaveInvoice.Location = new Point(988, 0);
            btnSaveInvoice.Margin = new Padding(3, 2, 3, 2);
            btnSaveInvoice.Name = "btnSaveInvoice";
            btnSaveInvoice.Size = new Size(149, 29);
            btnSaveInvoice.TabIndex = 7;
            btnSaveInvoice.Text = "💾 Cập nhật hóa đơn";
            btnSaveInvoice.UseVisualStyleBackColor = false;
            btnSaveInvoice.Click += btnSaveInvoice_Click;
            // 
            // groupBox3
            // 
            groupBox3.Controls.Add(panel4);
            groupBox3.Controls.Add(panel5);
            groupBox3.Dock = DockStyle.Fill;
            groupBox3.Location = new Point(3, 254);
            groupBox3.Name = "groupBox3";
            groupBox3.Size = new Size(1143, 358);
            groupBox3.TabIndex = 0;
            groupBox3.TabStop = false;
            // 
            // panel4
            // 
            panel4.Controls.Add(dgvInvoices);
            panel4.Dock = DockStyle.Fill;
            panel4.Location = new Point(3, 48);
            panel4.Name = "panel4";
            panel4.Size = new Size(1137, 307);
            panel4.TabIndex = 8;
            // 
            // dgvInvoices
            // 
            dgvInvoices.AllowUserToAddRows = false;
            dgvInvoices.AllowUserToDeleteRows = false;
            dgvInvoices.AllowUserToOrderColumns = true;
            dgvInvoices.BackgroundColor = Color.White;
            dgvInvoices.BorderStyle = BorderStyle.None;
            dgvInvoices.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Sunken;
            dgvInvoices.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvInvoices.Dock = DockStyle.Fill;
            dgvInvoices.Location = new Point(0, 0);
            dgvInvoices.Margin = new Padding(3, 2, 3, 2);
            dgvInvoices.Name = "dgvInvoices";
            dgvInvoices.ReadOnly = true;
            dgvInvoices.RowHeadersWidth = 51;
            dgvInvoices.Size = new Size(1137, 307);
            dgvInvoices.TabIndex = 6;
            dgvInvoices.CellContentDoubleClick += dgvInvoices_CellContentDoubleClick;
            dgvInvoices.EditingControlShowing += dgvList_EditingControlShowing;
            dgvInvoices.Scroll += dgvList_Scroll;
            // 
            // panel5
            // 
            panel5.Controls.Add(txtInvoiceSearch);
            panel5.Controls.Add(label5);
            panel5.Dock = DockStyle.Top;
            panel5.Location = new Point(3, 19);
            panel5.Name = "panel5";
            panel5.Size = new Size(1137, 29);
            panel5.TabIndex = 7;
            // 
            // txtInvoiceSearch
            // 
            txtInvoiceSearch.Location = new Point(96, 1);
            txtInvoiceSearch.Margin = new Padding(3, 2, 3, 2);
            txtInvoiceSearch.Name = "txtInvoiceSearch";
            txtInvoiceSearch.Size = new Size(167, 23);
            txtInvoiceSearch.TabIndex = 9;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(13, 5);
            label5.Name = "label5";
            label5.Size = new Size(67, 15);
            label5.TabIndex = 6;
            label5.Text = "Invoice No:";
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(btnSearch);
            groupBox1.Controls.Add(cboPO);
            groupBox1.Controls.Add(label1);
            groupBox1.Controls.Add(cboProject);
            groupBox1.Controls.Add(label2);
            groupBox1.Dock = DockStyle.Top;
            groupBox1.Location = new Point(0, 0);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(1149, 43);
            groupBox1.TabIndex = 0;
            groupBox1.TabStop = false;
            // 
            // btnSearch
            // 
            btnSearch.BackColor = Color.FromArgb(0, 120, 212);
            btnSearch.FlatStyle = FlatStyle.Flat;
            btnSearch.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnSearch.ForeColor = Color.White;
            btnSearch.Location = new Point(673, 8);
            btnSearch.Margin = new Padding(3, 2, 3, 2);
            btnSearch.Name = "btnSearch";
            btnSearch.Size = new Size(101, 33);
            btnSearch.TabIndex = 7;
            btnSearch.Text = "🔍 Tìm kiếm";
            btnSearch.UseVisualStyleBackColor = false;
            btnSearch.Click += btnSearch_Click;
            // 
            // cboPO
            // 
            cboPO.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cboPO.AutoCompleteSource = AutoCompleteSource.ListItems;
            cboPO.FormattingEnabled = true;
            cboPO.Location = new Point(416, 14);
            cboPO.Margin = new Padding(3, 2, 3, 2);
            cboPO.Name = "cboPO";
            cboPO.Size = new Size(238, 23);
            cboPO.TabIndex = 6;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(350, 17);
            label1.Name = "label1";
            label1.Size = new Size(42, 15);
            label1.TabIndex = 5;
            label1.Text = "PO.No";
            // 
            // cboProject
            // 
            cboProject.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cboProject.AutoCompleteSource = AutoCompleteSource.ListItems;
            cboProject.FormattingEnabled = true;
            cboProject.Location = new Point(74, 14);
            cboProject.Margin = new Padding(3, 2, 3, 2);
            cboProject.Name = "cboProject";
            cboProject.Size = new Size(238, 23);
            cboProject.TabIndex = 6;
            cboProject.SelectedIndexChanged += cboProject_SelectedIndexChanged;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(8, 17);
            label2.Name = "label2";
            label2.Size = new Size(41, 15);
            label2.TabIndex = 5;
            label2.Text = "Dự án:";
            // 
            // ucFillInvoiceNo
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            Controls.Add(panel1);
            Name = "ucFillInvoiceNo";
            Size = new Size(1149, 658);
            Load += ucFillInvoiceNo_Load;
            panel1.ResumeLayout(false);
            tableLayoutPanel1.ResumeLayout(false);
            groupBox2.ResumeLayout(false);
            panel3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvList).EndInit();
            panel2.ResumeLayout(false);
            panel2.PerformLayout();
            groupBox3.ResumeLayout(false);
            panel4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvInvoices).EndInit();
            panel5.ResumeLayout(false);
            panel5.PerformLayout();
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private Panel panel1;
        private GroupBox groupBox2;
        private GroupBox groupBox1;
        private Button btnSearch;
        private ComboBox cboPO;
        private Label label1;
        private ComboBox cboProject;
        private Label label2;
        private DataGridView dgvList;
        private Panel panel3;
        private Panel panel2;
        private Label lblStatus;
        private Button btnClear;
        private Button btnSaveInvoice;
        private DateTimePicker dtDate;
        private TextBox txtInvoiceNo;
        private Label label3;
        private Button btnFill;
        private TableLayoutPanel tableLayoutPanel1;
        private GroupBox groupBox3;
        private Panel panel4;
        private DataGridView dgvInvoices;
        private Panel panel5;
        private TextBox txtInvoiceSearch;
        private Label label5;
    }
}
