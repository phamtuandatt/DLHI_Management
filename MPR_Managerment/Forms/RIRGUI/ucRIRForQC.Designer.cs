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
            DataGridViewCellStyle dataGridViewCellStyle1 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle2 = new DataGridViewCellStyle();
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
            tableLayoutPanel2 = new TableLayoutPanel();
            panel4 = new Panel();
            panel6 = new Panel();
            groupBox1 = new GroupBox();
            dgvPaint = new DataGridView();
            panel5 = new Panel();
            cboProjectForPain = new ComboBox();
            btnShowAllPaint = new Button();
            btnReportPaintOfProject = new Button();
            cboRIRNoPaint = new ComboBox();
            btnPrintRPPaint = new Button();
            label6 = new Label();
            label3 = new Label();
            panel7 = new Panel();
            panel9 = new Panel();
            groupBox3 = new GroupBox();
            dgvWelding = new DataGridView();
            panel8 = new Panel();
            cboProjectForWelding = new ComboBox();
            btnShowAllWelding = new Button();
            label4 = new Label();
            btnReportWeldingOfProject = new Button();
            cboRIRNoWelding = new ComboBox();
            label7 = new Label();
            btnPrinRPWelding = new Button();
            groupBox4 = new GroupBox();
            dgvRIR = new DataGridView();
            panel2 = new Panel();
            btnSave = new Button();
            label5 = new Label();
            btnExportListItem = new Button();
            btnExport = new Button();
            btnClear = new Button();
            lblStatus = new Label();
            btnXoaRow = new Button();
            btnUpdateIDCode_QC = new Button();
            panel1.SuspendLayout();
            groupBox2.SuspendLayout();
            panel3.SuspendLayout();
            tableLayoutPanel1.SuspendLayout();
            tableLayoutPanel2.SuspendLayout();
            panel4.SuspendLayout();
            panel6.SuspendLayout();
            groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvPaint).BeginInit();
            panel5.SuspendLayout();
            panel7.SuspendLayout();
            panel9.SuspendLayout();
            groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvWelding).BeginInit();
            panel8.SuspendLayout();
            groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvRIR).BeginInit();
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
            panel1.Size = new Size(1630, 38);
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
            groupBox2.Size = new Size(1630, 762);
            groupBox2.TabIndex = 2;
            groupBox2.TabStop = false;
            // 
            // panel3
            // 
            panel3.Controls.Add(tableLayoutPanel1);
            panel3.Dock = DockStyle.Fill;
            panel3.Location = new Point(3, 56);
            panel3.Name = "panel3";
            panel3.Size = new Size(1624, 704);
            panel3.TabIndex = 7;
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.ColumnCount = 1;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tableLayoutPanel1.Controls.Add(tableLayoutPanel2, 0, 1);
            tableLayoutPanel1.Controls.Add(groupBox4, 0, 0);
            tableLayoutPanel1.Dock = DockStyle.Fill;
            tableLayoutPanel1.Location = new Point(0, 0);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 2;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 39.47826F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 60.52174F));
            tableLayoutPanel1.Size = new Size(1624, 704);
            tableLayoutPanel1.TabIndex = 6;
            // 
            // tableLayoutPanel2
            // 
            tableLayoutPanel2.ColumnCount = 2;
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tableLayoutPanel2.Controls.Add(panel4, 0, 0);
            tableLayoutPanel2.Controls.Add(panel7, 1, 0);
            tableLayoutPanel2.Dock = DockStyle.Fill;
            tableLayoutPanel2.Location = new Point(3, 280);
            tableLayoutPanel2.Name = "tableLayoutPanel2";
            tableLayoutPanel2.RowCount = 1;
            tableLayoutPanel2.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel2.Size = new Size(1618, 421);
            tableLayoutPanel2.TabIndex = 6;
            // 
            // panel4
            // 
            panel4.Controls.Add(panel6);
            panel4.Controls.Add(panel5);
            panel4.Dock = DockStyle.Fill;
            panel4.Location = new Point(3, 3);
            panel4.Name = "panel4";
            panel4.Size = new Size(803, 415);
            panel4.TabIndex = 1;
            // 
            // panel6
            // 
            panel6.Controls.Add(groupBox1);
            panel6.Dock = DockStyle.Fill;
            panel6.Location = new Point(0, 42);
            panel6.Name = "panel6";
            panel6.Size = new Size(803, 373);
            panel6.TabIndex = 2;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(dgvPaint);
            groupBox1.Dock = DockStyle.Fill;
            groupBox1.Location = new Point(0, 0);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(803, 373);
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
            dataGridViewCellStyle1.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = Color.DodgerBlue;
            dataGridViewCellStyle1.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            dataGridViewCellStyle1.ForeColor = Color.White;
            dataGridViewCellStyle1.Padding = new Padding(5);
            dataGridViewCellStyle1.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = DataGridViewTriState.True;
            dgvPaint.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            dgvPaint.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvPaint.Dock = DockStyle.Fill;
            dgvPaint.EnableHeadersVisualStyles = false;
            dgvPaint.Location = new Point(3, 19);
            dgvPaint.Margin = new Padding(3, 2, 3, 2);
            dgvPaint.Name = "dgvPaint";
            dgvPaint.RowHeadersWidth = 51;
            dgvPaint.Size = new Size(797, 351);
            dgvPaint.TabIndex = 6;
            // 
            // panel5
            // 
            panel5.BorderStyle = BorderStyle.FixedSingle;
            panel5.Controls.Add(cboProjectForPain);
            panel5.Controls.Add(btnShowAllPaint);
            panel5.Controls.Add(btnReportPaintOfProject);
            panel5.Controls.Add(cboRIRNoPaint);
            panel5.Controls.Add(btnPrintRPPaint);
            panel5.Controls.Add(label6);
            panel5.Controls.Add(label3);
            panel5.Dock = DockStyle.Top;
            panel5.Location = new Point(0, 0);
            panel5.Name = "panel5";
            panel5.Size = new Size(803, 42);
            panel5.TabIndex = 1;
            // 
            // cboProjectForPain
            // 
            cboProjectForPain.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cboProjectForPain.AutoCompleteSource = AutoCompleteSource.ListItems;
            cboProjectForPain.FormattingEnabled = true;
            cboProjectForPain.Location = new Point(51, 9);
            cboProjectForPain.Margin = new Padding(3, 2, 3, 2);
            cboProjectForPain.Name = "cboProjectForPain";
            cboProjectForPain.Size = new Size(132, 23);
            cboProjectForPain.TabIndex = 8;
            cboProjectForPain.SelectedIndexChanged += cboProjectForPain_SelectedIndexChanged;
            // 
            // btnShowAllPaint
            // 
            btnShowAllPaint.BackColor = Color.FromArgb(0, 120, 212);
            btnShowAllPaint.FlatStyle = FlatStyle.Flat;
            btnShowAllPaint.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnShowAllPaint.ForeColor = Color.White;
            btnShowAllPaint.Location = new Point(407, 6);
            btnShowAllPaint.Margin = new Padding(3, 2, 3, 2);
            btnShowAllPaint.Name = "btnShowAllPaint";
            btnShowAllPaint.Size = new Size(101, 29);
            btnShowAllPaint.TabIndex = 7;
            btnShowAllPaint.Text = "🔍 Tất cả";
            btnShowAllPaint.UseVisualStyleBackColor = false;
            btnShowAllPaint.Click += btnShowAllPaint_Click;
            // 
            // btnReportPaintOfProject
            // 
            btnReportPaintOfProject.BackColor = Color.ForestGreen;
            btnReportPaintOfProject.FlatStyle = FlatStyle.Flat;
            btnReportPaintOfProject.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnReportPaintOfProject.ForeColor = Color.White;
            btnReportPaintOfProject.Location = new Point(641, 6);
            btnReportPaintOfProject.Margin = new Padding(3, 2, 3, 2);
            btnReportPaintOfProject.Name = "btnReportPaintOfProject";
            btnReportPaintOfProject.Size = new Size(82, 29);
            btnReportPaintOfProject.TabIndex = 4;
            btnReportPaintOfProject.Text = "📄 Report";
            btnReportPaintOfProject.UseVisualStyleBackColor = false;
            btnReportPaintOfProject.Click += btnReportPaintOfProject_Click;
            // 
            // cboRIRNoPaint
            // 
            cboRIRNoPaint.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cboRIRNoPaint.AutoCompleteSource = AutoCompleteSource.ListItems;
            cboRIRNoPaint.FormattingEnabled = true;
            cboRIRNoPaint.Location = new Point(234, 9);
            cboRIRNoPaint.Margin = new Padding(3, 2, 3, 2);
            cboRIRNoPaint.Name = "cboRIRNoPaint";
            cboRIRNoPaint.Size = new Size(167, 23);
            cboRIRNoPaint.TabIndex = 6;
            cboRIRNoPaint.SelectedIndexChanged += cboRIRNoPaint_SelectedIndexChanged;
            // 
            // btnPrintRPPaint
            // 
            btnPrintRPPaint.BackColor = Color.ForestGreen;
            btnPrintRPPaint.FlatStyle = FlatStyle.Flat;
            btnPrintRPPaint.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnPrintRPPaint.ForeColor = Color.White;
            btnPrintRPPaint.Location = new Point(514, 6);
            btnPrintRPPaint.Margin = new Padding(3, 2, 3, 2);
            btnPrintRPPaint.Name = "btnPrintRPPaint";
            btnPrintRPPaint.Size = new Size(121, 29);
            btnPrintRPPaint.TabIndex = 4;
            btnPrintRPPaint.Text = "📄 In báo cáo sơn";
            btnPrintRPPaint.UseVisualStyleBackColor = false;
            btnPrintRPPaint.Click += btnPrintRPPaint_Click;
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(186, 13);
            label6.Name = "label6";
            label6.Size = new Size(46, 15);
            label6.TabIndex = 5;
            label6.Text = "RIR No:";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(6, 13);
            label3.Name = "label3";
            label3.Size = new Size(41, 15);
            label3.TabIndex = 5;
            label3.Text = "Dự án:";
            // 
            // panel7
            // 
            panel7.Controls.Add(panel9);
            panel7.Controls.Add(panel8);
            panel7.Dock = DockStyle.Fill;
            panel7.Location = new Point(812, 3);
            panel7.Name = "panel7";
            panel7.Size = new Size(803, 415);
            panel7.TabIndex = 2;
            // 
            // panel9
            // 
            panel9.Controls.Add(groupBox3);
            panel9.Dock = DockStyle.Fill;
            panel9.Location = new Point(0, 42);
            panel9.Name = "panel9";
            panel9.Size = new Size(803, 373);
            panel9.TabIndex = 1;
            // 
            // groupBox3
            // 
            groupBox3.Controls.Add(dgvWelding);
            groupBox3.Dock = DockStyle.Fill;
            groupBox3.Location = new Point(0, 0);
            groupBox3.Name = "groupBox3";
            groupBox3.Size = new Size(803, 373);
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
            dataGridViewCellStyle2.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = Color.DodgerBlue;
            dataGridViewCellStyle2.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            dataGridViewCellStyle2.ForeColor = Color.White;
            dataGridViewCellStyle2.Padding = new Padding(5);
            dataGridViewCellStyle2.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = DataGridViewTriState.True;
            dgvWelding.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            dgvWelding.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvWelding.Dock = DockStyle.Fill;
            dgvWelding.EnableHeadersVisualStyles = false;
            dgvWelding.Location = new Point(3, 19);
            dgvWelding.Margin = new Padding(3, 2, 3, 2);
            dgvWelding.Name = "dgvWelding";
            dgvWelding.RowHeadersWidth = 51;
            dgvWelding.Size = new Size(797, 351);
            dgvWelding.TabIndex = 6;
            // 
            // panel8
            // 
            panel8.BorderStyle = BorderStyle.FixedSingle;
            panel8.Controls.Add(cboProjectForWelding);
            panel8.Controls.Add(btnShowAllWelding);
            panel8.Controls.Add(label4);
            panel8.Controls.Add(btnReportWeldingOfProject);
            panel8.Controls.Add(cboRIRNoWelding);
            panel8.Controls.Add(label7);
            panel8.Controls.Add(btnPrinRPWelding);
            panel8.Dock = DockStyle.Top;
            panel8.Location = new Point(0, 0);
            panel8.Name = "panel8";
            panel8.Size = new Size(803, 42);
            panel8.TabIndex = 0;
            // 
            // cboProjectForWelding
            // 
            cboProjectForWelding.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cboProjectForWelding.AutoCompleteSource = AutoCompleteSource.ListItems;
            cboProjectForWelding.FormattingEnabled = true;
            cboProjectForWelding.Location = new Point(59, 9);
            cboProjectForWelding.Margin = new Padding(3, 2, 3, 2);
            cboProjectForWelding.Name = "cboProjectForWelding";
            cboProjectForWelding.Size = new Size(132, 23);
            cboProjectForWelding.TabIndex = 8;
            cboProjectForWelding.SelectedIndexChanged += cboProjectForWelding_SelectedIndexChanged;
            // 
            // btnShowAllWelding
            // 
            btnShowAllWelding.BackColor = Color.FromArgb(0, 120, 212);
            btnShowAllWelding.FlatStyle = FlatStyle.Flat;
            btnShowAllWelding.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnShowAllWelding.ForeColor = Color.White;
            btnShowAllWelding.Location = new Point(420, 6);
            btnShowAllWelding.Margin = new Padding(3, 2, 3, 2);
            btnShowAllWelding.Name = "btnShowAllWelding";
            btnShowAllWelding.Size = new Size(101, 29);
            btnShowAllWelding.TabIndex = 7;
            btnShowAllWelding.Text = "🔍 Tất cả";
            btnShowAllWelding.UseVisualStyleBackColor = false;
            btnShowAllWelding.Click += btnShowAllWelding_Click;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(14, 13);
            label4.Name = "label4";
            label4.Size = new Size(41, 15);
            label4.TabIndex = 5;
            label4.Text = "Dự án:";
            // 
            // btnReportWeldingOfProject
            // 
            btnReportWeldingOfProject.BackColor = Color.ForestGreen;
            btnReportWeldingOfProject.FlatStyle = FlatStyle.Flat;
            btnReportWeldingOfProject.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnReportWeldingOfProject.ForeColor = Color.White;
            btnReportWeldingOfProject.Location = new Point(689, 5);
            btnReportWeldingOfProject.Margin = new Padding(3, 2, 3, 2);
            btnReportWeldingOfProject.Name = "btnReportWeldingOfProject";
            btnReportWeldingOfProject.Size = new Size(82, 29);
            btnReportWeldingOfProject.TabIndex = 4;
            btnReportWeldingOfProject.Text = "📄 Report";
            btnReportWeldingOfProject.UseVisualStyleBackColor = false;
            btnReportWeldingOfProject.Click += btnReportWeldingOfProject_Click;
            // 
            // cboRIRNoWelding
            // 
            cboRIRNoWelding.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cboRIRNoWelding.AutoCompleteSource = AutoCompleteSource.ListItems;
            cboRIRNoWelding.FormattingEnabled = true;
            cboRIRNoWelding.Location = new Point(247, 9);
            cboRIRNoWelding.Margin = new Padding(3, 2, 3, 2);
            cboRIRNoWelding.Name = "cboRIRNoWelding";
            cboRIRNoWelding.Size = new Size(167, 23);
            cboRIRNoWelding.TabIndex = 6;
            cboRIRNoWelding.SelectedIndexChanged += cboRIRNoWelding_SelectedIndexChanged;
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new Point(195, 13);
            label7.Name = "label7";
            label7.Size = new Size(46, 15);
            label7.TabIndex = 5;
            label7.Text = "RIR No:";
            // 
            // btnPrinRPWelding
            // 
            btnPrinRPWelding.BackColor = Color.ForestGreen;
            btnPrinRPWelding.FlatStyle = FlatStyle.Flat;
            btnPrinRPWelding.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnPrinRPWelding.ForeColor = Color.White;
            btnPrinRPWelding.Location = new Point(524, 6);
            btnPrinRPWelding.Margin = new Padding(3, 2, 3, 2);
            btnPrinRPWelding.Name = "btnPrinRPWelding";
            btnPrinRPWelding.Size = new Size(159, 29);
            btnPrinRPWelding.TabIndex = 4;
            btnPrinRPWelding.Text = "📄 In báo cáo vật tư hàn";
            btnPrinRPWelding.UseVisualStyleBackColor = false;
            btnPrinRPWelding.Click += btnPrinRPWelding_Click;
            // 
            // groupBox4
            // 
            groupBox4.Controls.Add(dgvRIR);
            groupBox4.Dock = DockStyle.Fill;
            groupBox4.Location = new Point(3, 3);
            groupBox4.Name = "groupBox4";
            groupBox4.Size = new Size(1618, 271);
            groupBox4.TabIndex = 7;
            groupBox4.TabStop = false;
            groupBox4.Text = "Danh sách vật tư";
            // 
            // dgvRIR
            // 
            dgvRIR.AllowUserToAddRows = false;
            dgvRIR.AllowUserToDeleteRows = false;
            dgvRIR.AllowUserToOrderColumns = true;
            dgvRIR.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvRIR.BackgroundColor = Color.White;
            dgvRIR.BorderStyle = BorderStyle.None;
            dgvRIR.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Sunken;
            dgvRIR.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvRIR.Dock = DockStyle.Fill;
            dgvRIR.Location = new Point(3, 19);
            dgvRIR.Margin = new Padding(3, 2, 3, 2);
            dgvRIR.Name = "dgvRIR";
            dgvRIR.RowHeadersWidth = 51;
            dgvRIR.Size = new Size(1612, 249);
            dgvRIR.TabIndex = 5;
            dgvRIR.CellContentClick += dgvRIR_CellContentClick;
            dgvRIR.CellEndEdit += dgvRIR_CellEndEdit;
            dgvRIR.CellFormatting += dgvRIR_CellFormatting;
            dgvRIR.EditingControlShowing += dgvRIR_EditingControlShowing;
            // 
            // panel2
            // 
            panel2.Controls.Add(btnSave);
            panel2.Controls.Add(label5);
            panel2.Controls.Add(btnUpdateIDCode_QC);
            panel2.Controls.Add(btnExportListItem);
            panel2.Controls.Add(btnExport);
            panel2.Controls.Add(btnClear);
            panel2.Controls.Add(lblStatus);
            panel2.Dock = DockStyle.Top;
            panel2.Location = new Point(3, 18);
            panel2.Name = "panel2";
            panel2.Size = new Size(1624, 38);
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
            // btnExportListItem
            // 
            btnExportListItem.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;
            btnExportListItem.BackColor = Color.ForestGreen;
            btnExportListItem.FlatStyle = FlatStyle.Flat;
            btnExportListItem.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnExportListItem.ForeColor = Color.White;
            btnExportListItem.Location = new Point(1392, 4);
            btnExportListItem.Margin = new Padding(3, 2, 3, 2);
            btnExportListItem.Name = "btnExportListItem";
            btnExportListItem.Size = new Size(109, 29);
            btnExportListItem.TabIndex = 4;
            btnExportListItem.Text = "📄 In báo cáo";
            btnExportListItem.UseVisualStyleBackColor = false;
            btnExportListItem.Click += btnExportListItem_Click;
            // 
            // btnExport
            // 
            btnExport.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;
            btnExport.BackColor = Color.ForestGreen;
            btnExport.FlatStyle = FlatStyle.Flat;
            btnExport.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnExport.ForeColor = Color.White;
            btnExport.Location = new Point(1507, 4);
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
            btnXoaRow.Location = new Point(2531, 15);
            btnXoaRow.Margin = new Padding(3, 2, 3, 2);
            btnXoaRow.Name = "btnXoaRow";
            btnXoaRow.Size = new Size(109, 29);
            btnXoaRow.TabIndex = 4;
            btnXoaRow.Text = "🗑 Xóa dòng";
            btnXoaRow.UseVisualStyleBackColor = false;
            // 
            // btnUpdateIDCode_QC
            // 
            btnUpdateIDCode_QC.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;
            btnUpdateIDCode_QC.BackColor = Color.Purple;
            btnUpdateIDCode_QC.FlatStyle = FlatStyle.Flat;
            btnUpdateIDCode_QC.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnUpdateIDCode_QC.ForeColor = Color.White;
            btnUpdateIDCode_QC.Location = new Point(1236, 4);
            btnUpdateIDCode_QC.Margin = new Padding(3, 2, 3, 2);
            btnUpdateIDCode_QC.Name = "btnUpdateIDCode_QC";
            btnUpdateIDCode_QC.Size = new Size(150, 29);
            btnUpdateIDCode_QC.TabIndex = 4;
            btnUpdateIDCode_QC.Text = "📄 Cập nhật ID Code";
            btnUpdateIDCode_QC.UseVisualStyleBackColor = false;
            btnUpdateIDCode_QC.Click += btnUpdateIDCode_QC_Click;
            // 
            // ucRIRForQC
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            Controls.Add(groupBox2);
            Controls.Add(panel1);
            Name = "ucRIRForQC";
            Size = new Size(1630, 800);
            Load += ucRIRForQC_Load;
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            groupBox2.ResumeLayout(false);
            panel3.ResumeLayout(false);
            tableLayoutPanel1.ResumeLayout(false);
            tableLayoutPanel2.ResumeLayout(false);
            panel4.ResumeLayout(false);
            panel6.ResumeLayout(false);
            groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvPaint).EndInit();
            panel5.ResumeLayout(false);
            panel5.PerformLayout();
            panel7.ResumeLayout(false);
            panel9.ResumeLayout(false);
            groupBox3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvWelding).EndInit();
            panel8.ResumeLayout(false);
            panel8.PerformLayout();
            groupBox4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvRIR).EndInit();
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
        private Button btnExportListItem;
        private Panel panel4;
        private Panel panel6;
        private Panel panel5;
        private ComboBox cboProjectForPain;
        private Button btnShowAllPaint;
        private Button btnPrintRPPaint;
        private Label label3;
        private Panel panel7;
        private Panel panel9;
        private Panel panel8;
        private ComboBox cboProjectForWelding;
        private Label label4;
        private Button btnPrinRPWelding;
        private Label label6;
        private Label label7;
        private GroupBox groupBox4;
        private Button btnReportPaintOfProject;
        private ComboBox cboRIRNoPaint;
        private ComboBox cboRIRNoWelding;
        private Button btnReportWeldingOfProject;
        private Button btnShowAllWelding;
        private Button btnUpdateIDCode_QC;
    }
}
