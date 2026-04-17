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
            txtSearch = new TextBox();
            btnSearch = new Button();
            cboRIRs = new ComboBox();
            lblCountRIR = new Label();
            label1 = new Label();
            label2 = new Label();
            groupBox2 = new GroupBox();
            panel3 = new Panel();
            dgvRIR = new DataGridView();
            panel2 = new Panel();
            btnSave = new Button();
            label5 = new Label();
            btnClear = new Button();
            lblStatus = new Label();
            btnXoaRow = new Button();
            panel1.SuspendLayout();
            groupBox2.SuspendLayout();
            panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvRIR).BeginInit();
            panel2.SuspendLayout();
            SuspendLayout();
            // 
            // panel1
            // 
            panel1.Controls.Add(txtSearch);
            panel1.Controls.Add(btnSearch);
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
            // txtSearch
            // 
            txtSearch.Location = new Point(70, 8);
            txtSearch.Margin = new Padding(3, 2, 3, 2);
            txtSearch.Name = "txtSearch";
            txtSearch.Size = new Size(236, 23);
            txtSearch.TabIndex = 8;
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
            panel3.Controls.Add(dgvRIR);
            panel3.Dock = DockStyle.Fill;
            panel3.Location = new Point(3, 56);
            panel3.Name = "panel3";
            panel3.Size = new Size(1214, 575);
            panel3.TabIndex = 7;
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
            dgvRIR.Location = new Point(0, 0);
            dgvRIR.Margin = new Padding(3, 2, 3, 2);
            dgvRIR.Name = "dgvRIR";
            dgvRIR.RowHeadersWidth = 51;
            dgvRIR.Size = new Size(1214, 575);
            dgvRIR.TabIndex = 5;
            dgvRIR.CellEndEdit += dgvRIR_CellEndEdit;
            dgvRIR.CellFormatting += dgvRIR_CellFormatting;
            dgvRIR.EditingControlShowing += dgvRIR_EditingControlShowing;
            // 
            // panel2
            // 
            panel2.Controls.Add(btnSave);
            panel2.Controls.Add(label5);
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
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            groupBox2.ResumeLayout(false);
            panel3.ResumeLayout(false);
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
        private TextBox txtSearch;
        private Label lblCountRIR;
    }
}
