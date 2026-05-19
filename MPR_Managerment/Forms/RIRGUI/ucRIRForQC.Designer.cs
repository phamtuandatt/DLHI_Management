namespace MPR_Managerment.Forms.RIRGUI
{
    partial class ucRIRForQC
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
            cboProjectMaterial = new ComboBox();
            btnSearch = new Button();
            btnPrintReportMaterial = new Button();
            cboRIRs = new ComboBox();
            lblCountRIR = new Label();
            label1 = new Label();
            label2 = new Label();
            groupBox2 = new GroupBox();
            panel3 = new Panel();
            tableLayoutPanel1 = new TableLayoutPanel();
            dgvRIR = new DataGridView();
            tableLayoutPanel2 = new TableLayoutPanel();
            groupBox1 = new GroupBox();
            dgvPaint = new DataGridView();
            groupBox3 = new GroupBox();
            dgvWelding = new DataGridView();
            panel2 = new Panel();
            btnSave = new Button();
            label5 = new Label();
            btnExport = new Button();
            btnClear = new Button();
            lblStatus = new Label();
            btnXoaRow = new Button();
            panel1.SuspendLayout();
            groupBox2.SuspendLayout();
            panel3.SuspendLayout();
            tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvRIR).BeginInit();
            tableLayoutPanel2.SuspendLayout();
            groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvPaint).BeginInit();
            groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvWelding).BeginInit();
            panel2.SuspendLayout();
            SuspendLayout();
            // 
            // panel1
            // 
            panel1.Controls.Add(cboProjectMaterial);
            panel1.Controls.Add(btnSearch);
            panel1.Controls.Add(btnPrintReportMaterial);
            panel1.Controls.Add(cboRIRs);
            panel1.Controls.Add(lblCountRIR);
            panel1.Controls.Add(label1);
            panel1.Controls.Add(label2);
            panel1.Dock = DockStyle.Top;
            panel1.Location = new Point(0, 0);
            panel1.Name = "panel1";
            panel1.Size = new Size(1220, 38);
            panel1.TabIndex = 0;
            // 
            // cboProjectMaterial
            // 
            cboProjectMaterial.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cboProjectMaterial.AutoCompleteSource = AutoCompleteSource.ListItems;
            cboProjectMaterial.FormattingEnabled = true;
            cboProjectMaterial.Location = new Point(71, 8);
            cboProjectMaterial.Margin = new Padding(3, 2, 3, 2);
            cboProjectMaterial.Name = "cboProjectMaterial";
            cboProjectMaterial.Size = new Size(243, 23);
            cboProjectMaterial.TabIndex = 8;
            cboProjectMaterial.SelectedIndexChanged += cboProjectMaterial_SelectedIndexChanged;
            // 
            // btnSearch
            // 
            btnSearch.BackColor = Color.FromArgb(0, 120, 212);
            btnSearch.FlatStyle = FlatStyle.Flat;
            btnSearch.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnSearch.ForeColor = Color.White;
            btnSearch.Location = new Point(327, 3);
            btnSearch.Margin = new Padding(3, 2, 3, 2);
            btnSearch.Name = "btnSearch";
            btnSearch.Size = new Size(101, 33);
            btnSearch.TabIndex = 7;
            btnSearch.Text = "🔍 Tìm kiếm";
            btnSearch.UseVisualStyleBackColor = false;
            btnSearch.Click += btnSearch_Click;
            // 
            // btnPrintReportMaterial
            // 
            btnPrintReportMaterial.BackColor = Color.ForestGreen;
            btnPrintReportMaterial.FlatStyle = FlatStyle.Flat;
            btnPrintReportMaterial.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnPrintReportMaterial.ForeColor = Color.White;
            btnPrintReportMaterial.Location = new Point(902, 5);
            btnPrintReportMaterial.Margin = new Padding(3, 2, 3, 2);
            btnPrintReportMaterial.Name = "btnPrintReportMaterial";
            btnPrintReportMaterial.Size = new Size(159, 29);
            btnPrintReportMaterial.TabIndex = 4;
            btnPrintReportMaterial.Text = "📄 In báo cáo vật tư";
            btnPrintReportMaterial.UseVisualStyleBackColor = false;
            btnPrintReportMaterial.Click += btnPrintReportMaterial_Click;
            // 
            // cboRIRs
            // 
            cboRIRs.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cboRIRs.AutoCompleteSource = AutoCompleteSource.ListItems;
            cboRIRs.FormattingEnabled = true;
            cboRIRs.Location = new Point(505, 8);
            cboRIRs.Margin = new Padding(3, 2, 3, 2);
            cboRIRs.Name = "cboRIRs";
            cboRIRs.Size = new Size(238, 23);
            cboRIRs.TabIndex = 6;
            cboRIRs.SelectedIndexChanged += cboRIRs_SelectedIndexChanged;
            // 
            // lblCountRIR
            // 
            lblCountRIR.AutoSize = true;
            lblCountRIR.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            lblCountRIR.ForeColor = Color.LimeGreen;
            lblCountRIR.Location = new Point(749, 12);
            lblCountRIR.Name = "lblCountRIR";
            lblCountRIR.Size = new Size(42, 15);
            lblCountRIR.TabIndex = 1;
            lblCountRIR.Text = "Status";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(456, 12);
            label1.Name = "label1";
            label1.Size = new Size(43, 15);
            label1.TabIndex = 5;
            label1.Text = "RIR No";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(10, 12);
            label2.Name = "label2";
            label2.Size = new Size(41, 15);
            label2.TabIndex = 5;
            label2.Text = "Dự án:";
            // 
            // groupBox2
            // 
            groupBox2.BackColor = Color.White;
            groupBox2.Controls.Add(panel3);
            groupBox2.Controls.Add(panel2);
            groupBox2.Controls.Add(btnXoaRow);
            groupBox2.Dock = DockStyle.Fill;
            groupBox2.Location = new Point(0, 38);
            groupBox2.Margin = new Padding(3, 2, 3, 2);
            groupBox2.Name = "groupBox2";
            groupBox2.Padding = new Padding(3, 2, 3, 2);
            groupBox2.Size = new Size(1220, 633);
            groupBox2.TabIndex = 2;
            groupBox2.TabStop = false;
            // 
            // panel3
            // 
            panel3.Controls.Add(tableLayoutPanel1);
            panel3.Dock = DockStyle.Fill;
            panel3.Location = new Point(3, 56);
            panel3.Name = "panel3";
            panel3.Size = new Size(1214, 575);
            panel3.TabIndex = 7;
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.ColumnCount = 1;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tableLayoutPanel1.Controls.Add(dgvRIR, 0, 0);
            tableLayoutPanel1.Controls.Add(tableLayoutPanel2, 0, 1);
            tableLayoutPanel1.Dock = DockStyle.Fill;
            tableLayoutPanel1.Location = new Point(0, 0);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 2;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 39.47826F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 60.52174F));
            tableLayoutPanel1.Size = new Size(1214, 575);
            tableLayoutPanel1.TabIndex = 6;
            // 
            // dgvRIR
            // 
            dgvRIR.AllowUserToAddRows = false;
            dgvRIR.AllowUserToDeleteRows = false;
            dgvRIR.AllowUserToOrderColumns = true;
            dgvRIR.BackgroundColor = Color.White;
            dgvRIR.BorderStyle = BorderStyle.None;
            dgvRIR.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Sunken;
            dgvRIR.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvRIR.Dock = DockStyle.Fill;
            dgvRIR.Location = new Point(3, 2);
            dgvRIR.Margin = new Padding(3, 2, 3, 2);
            dgvRIR.Name = "dgvRIR";
            dgvRIR.RowHeadersWidth = 51;
            dgvRIR.Size = new Size(1208, 223);
            dgvRIR.TabIndex = 5;
            dgvRIR.CellContentClick += dgvRIR_CellContentClick;
            dgvRIR.CellEndEdit += dgvRIR_CellEndEdit;
            dgvRIR.CellFormatting += dgvRIR_CellFormatting;
            dgvRIR.EditingControlShowing += dgvRIR_EditingControlShowing;
            // 
            // tableLayoutPanel2
            // 
            tableLayoutPanel2.ColumnCount = 2;
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tableLayoutPanel2.Controls.Add(groupBox1, 0, 0);
            tableLayoutPanel2.Controls.Add(groupBox3, 1, 0);
            tableLayoutPanel2.Dock = DockStyle.Fill;
            tableLayoutPanel2.Location = new Point(3, 230);
            tableLayoutPanel2.Name = "tableLayoutPanel2";
            tableLayoutPanel2.RowCount = 1;
            tableLayoutPanel2.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel2.Size = new Size(1208, 342);
            tableLayoutPanel2.TabIndex = 6;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(dgvPaint);
            groupBox1.Dock = DockStyle.Fill;
            groupBox1.Location = new Point(3, 3);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(598, 336);
            groupBox1.TabIndex = 0;
            groupBox1.TabStop = false;
            groupBox1.Text = "Danh sách sơn";
            // 
            // dgvPaint
            // 
            dgvPaint.AllowUserToAddRows = false;
            dgvPaint.AllowUserToDeleteRows = false;
            dgvPaint.AllowUserToOrderColumns = true;
            dgvPaint.BackgroundColor = Color.White;
            dgvPaint.BorderStyle = BorderStyle.None;
            dgvPaint.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Sunken;
            dgvPaint.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvPaint.Dock = DockStyle.Fill;
            dgvPaint.Location = new Point(3, 19);
            dgvPaint.Margin = new Padding(3, 2, 3, 2);
            dgvPaint.Name = "dgvPaint";
            dgvPaint.RowHeadersWidth = 51;
            dgvPaint.Size = new Size(592, 314);
            dgvPaint.TabIndex = 6;
            // 
            // groupBox3
            // 
            groupBox3.Controls.Add(dgvWelding);
            groupBox3.Dock = DockStyle.Fill;
            groupBox3.Location = new Point(607, 3);
            groupBox3.Name = "groupBox3";
            groupBox3.Size = new Size(598, 336);
            groupBox3.TabIndex = 0;
            groupBox3.TabStop = false;
            groupBox3.Text = "Danh sách que hàn";
            // 
            // dgvWelding
            // 
            dgvWelding.AllowUserToAddRows = false;
            dgvWelding.AllowUserToDeleteRows = false;
            dgvWelding.AllowUserToOrderColumns = true;
            dgvWelding.BackgroundColor = Color.White;
            dgvWelding.BorderStyle = BorderStyle.None;
            dgvWelding.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Sunken;
            dgvWelding.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvWelding.Dock = DockStyle.Fill;
            dgvWelding.Location = new Point(3, 19);
            dgvWelding.Margin = new Padding(3, 2, 3, 2);
            dgvWelding.Name = "dgvWelding";
            dgvWelding.RowHeadersWidth = 51;
            dgvWelding.Size = new Size(592, 314);
            dgvWelding.TabIndex = 6;
            // 
            // panel2
            // 
            panel2.Controls.Add(btnSave);
            panel2.Controls.Add(label5);
            panel2.Controls.Add(btnExport);
            panel2.Controls.Add(btnClear);
            panel2.Controls.Add(lblStatus);
            panel2.Dock = DockStyle.Top;
            panel2.Location = new Point(3, 18);
            panel2.Name = "panel2";
            panel2.Size = new Size(1214, 38);
            panel2.TabIndex = 6;
            // 
            // btnSave
            // 
            btnSave.BackColor = Color.FromArgb(255, 128, 0);
            btnSave.FlatStyle = FlatStyle.Flat;
            btnSave.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnSave.ForeColor = Color.White;
            btnSave.Location = new Point(202, 2);
            btnSave.Margin = new Padding(3, 2, 3, 2);
            btnSave.Name = "btnSave";
            btnSave.Size = new Size(109, 29);
            btnSave.TabIndex = 4;
            btnSave.Text = "💾 Lưu RIR";
            btnSave.UseVisualStyleBackColor = false;
            btnSave.Click += btnSave_Click;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.BackColor = Color.Transparent;
            label5.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            label5.ForeColor = Color.FromArgb(220, 53, 69);
            label5.Location = new Point(3, 6);
            label5.Name = "label5";
            label5.Size = new Size(182, 21);
            label5.TabIndex = 0;
            label5.Text = "THÔNG TIN XUẤT KHO";
            // 
            // btnExport
            // 
            btnExport.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;
            btnExport.BackColor = Color.ForestGreen;
            btnExport.FlatStyle = FlatStyle.Flat;
            btnExport.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnExport.ForeColor = Color.White;
            btnExport.Location = new Point(1097, 4);
            btnExport.Margin = new Padding(3, 2, 3, 2);
            btnExport.Name = "btnExport";
            btnExport.Size = new Size(109, 29);
            btnExport.TabIndex = 4;
            btnExport.Text = "📄 Xuất Excel";
            btnExport.UseVisualStyleBackColor = false;
            btnExport.Click += btnExport_Click;
            // 
            // btnClear
            // 
            btnClear.BackColor = Color.FromArgb(108, 117, 125);
            btnClear.FlatStyle = FlatStyle.Flat;
            btnClear.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnClear.ForeColor = Color.White;
            btnClear.Location = new Point(321, 2);
            btnClear.Margin = new Padding(3, 2, 3, 2);
            btnClear.Name = "btnClear";
            btnClear.Size = new Size(109, 29);
            btnClear.TabIndex = 4;
            btnClear.Text = "🔄 Xóa form";
            btnClear.UseVisualStyleBackColor = false;
            btnClear.Click += btnClear_Click;
            // 
            // lblStatus
            // 
            lblStatus.AutoSize = true;
            lblStatus.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            lblStatus.ForeColor = Color.LimeGreen;
            lblStatus.Location = new Point(445, 10);
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(42, 15);
            lblStatus.TabIndex = 1;
            lblStatus.Text = "Status";
            // 
            // btnXoaRow
            // 
            btnXoaRow.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnXoaRow.BackColor = Color.FromArgb(220, 53, 69);
            btnXoaRow.FlatStyle = FlatStyle.Flat;
            btnXoaRow.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnXoaRow.ForeColor = Color.White;
            btnXoaRow.Location = new Point(2121, 15);
            btnXoaRow.Margin = new Padding(3, 2, 3, 2);
            btnXoaRow.Name = "btnXoaRow";
            btnXoaRow.Size = new Size(109, 29);
            btnXoaRow.TabIndex = 4;
            btnXoaRow.Text = "🗑 Xóa dòng";
            btnXoaRow.UseVisualStyleBackColor = false;
            // 
            // ucRIRForQC
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            Controls.Add(groupBox2);
            Controls.Add(panel1);
            Name = "ucRIRForQC";
            Size = new Size(1220, 671);
            Load += ucRIRForQC_Load;
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            groupBox2.ResumeLayout(false);
            panel3.ResumeLayout(false);
            tableLayoutPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvRIR).EndInit();
            tableLayoutPanel2.ResumeLayout(false);
            groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvPaint).EndInit();
            groupBox3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvWelding).EndInit();
            panel2.ResumeLayout(false);
            panel2.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private Panel panel1;
        private Button btnSearch;
        private ComboBox cboRIRs;
        private Label label1;
        private Label label2;
        private GroupBox groupBox2;
        private DataGridView dgvRIR;
        private Button btnClear;
        private Button btnSave;
        private Button btnXoaRow;
        private Label lblStatus;
        private Label label5;
        private Panel panel3;
        private Panel panel2;
        private Label lblCountRIR;
        private TableLayoutPanel tableLayoutPanel1;
        private TableLayoutPanel tableLayoutPanel2;
        private GroupBox groupBox1;
        private DataGridView dgvPaint;
        private GroupBox groupBox3;
        private DataGridView dgvWelding;
        private Button btnExport;
        private ComboBox cboProjectMaterial;
        private Button btnPrintReportMaterial;
    }
}
