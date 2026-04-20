namespace MPR_Managerment.Forms.ExportGUI
{
    partial class ucExportWarehouse
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
            groupBox1 = new GroupBox();
            dgvKho = new DataGridView();
            btnCancelSer = new Button();
            btnSearch = new Button();
            cboProject = new ComboBox();
            lblStatus = new Label();
            label2 = new Label();
            label1 = new Label();
            groupBox2 = new GroupBox();
            dgvExportQue = new DataGridView();
            btnClear = new Button();
            button1 = new Button();
            btnXoaRow = new Button();
            lblInfoKho = new Label();
            label5 = new Label();
            groupBox3 = new GroupBox();
            txtSearch = new TextBox();
            dtpTo = new DateTimePicker();
            dtpStart = new DateTimePicker();
            dgvHis = new DataGridView();
            btnClearSearch = new Button();
            btnSearchHis = new Button();
            lblInfoXK = new Label();
            label6 = new Label();
            label3 = new Label();
            cboProjectCheck = new ComboBox();
            label7 = new Label();
            groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvKho).BeginInit();
            groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvExportQue).BeginInit();
            groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvHis).BeginInit();
            SuspendLayout();
            // 
            // groupBox1
            // 
            groupBox1.BackColor = Color.White;
            groupBox1.Controls.Add(dgvKho);
            groupBox1.Controls.Add(btnCancelSer);
            groupBox1.Controls.Add(btnSearch);
            groupBox1.Controls.Add(cboProject);
            groupBox1.Controls.Add(lblStatus);
            groupBox1.Controls.Add(label2);
            groupBox1.Controls.Add(label1);
            groupBox1.Dock = DockStyle.Top;
            groupBox1.Location = new Point(0, 0);
            groupBox1.Margin = new Padding(3, 2, 3, 2);
            groupBox1.Name = "groupBox1";
            groupBox1.Padding = new Padding(3, 2, 3, 2);
            groupBox1.Size = new Size(1600, 369);
            groupBox1.TabIndex = 0;
            groupBox1.TabStop = false;
            // 
            // dgvKho
            // 
            dgvKho.AllowUserToAddRows = false;
            dgvKho.AllowUserToDeleteRows = false;
            dgvKho.AllowUserToOrderColumns = true;
            dgvKho.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dgvKho.BackgroundColor = Color.White;
            dgvKho.BorderStyle = BorderStyle.None;
            dgvKho.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Sunken;
            dgvKho.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvKho.Location = new Point(14, 95);
            dgvKho.Margin = new Padding(3, 2, 3, 2);
            dgvKho.Name = "dgvKho";
            dgvKho.ReadOnly = true;
            dgvKho.RowHeadersWidth = 51;
            dgvKho.Size = new Size(1570, 262);
            dgvKho.TabIndex = 5;
            dgvKho.CellClick += dgvKho_CellClick;
            dgvKho.CellDoubleClick += dgvKho_CellDoubleClick;
            dgvKho.CellFormatting += dgvKho_CellFormatting;
            // 
            // btnCancelSer
            // 
            btnCancelSer.BackColor = Color.DarkGray;
            btnCancelSer.FlatStyle = FlatStyle.Flat;
            btnCancelSer.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnCancelSer.ForeColor = Color.White;
            btnCancelSer.Location = new Point(438, 42);
            btnCancelSer.Margin = new Padding(3, 2, 3, 2);
            btnCancelSer.Name = "btnCancelSer";
            btnCancelSer.Size = new Size(101, 33);
            btnCancelSer.TabIndex = 4;
            btnCancelSer.Text = "✖ Xóa lọc";
            btnCancelSer.UseVisualStyleBackColor = false;
            // 
            // btnSearch
            // 
            btnSearch.BackColor = Color.FromArgb(0, 120, 212);
            btnSearch.FlatStyle = FlatStyle.Flat;
            btnSearch.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnSearch.ForeColor = Color.White;
            btnSearch.Location = new Point(336, 42);
            btnSearch.Margin = new Padding(3, 2, 3, 2);
            btnSearch.Name = "btnSearch";
            btnSearch.Size = new Size(101, 33);
            btnSearch.TabIndex = 3;
            btnSearch.Text = "🔍 Tìm kiếm";
            btnSearch.UseVisualStyleBackColor = false;
            btnSearch.Click += btnSearch_Click;
            // 
            // cboProject
            // 
            cboProject.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cboProject.AutoCompleteSource = AutoCompleteSource.ListItems;
            cboProject.FormattingEnabled = true;
            cboProject.Location = new Point(80, 48);
            cboProject.Margin = new Padding(3, 2, 3, 2);
            cboProject.Name = "cboProject";
            cboProject.Size = new Size(238, 23);
            cboProject.TabIndex = 2;
            cboProject.Validating += cboProject_Validating;
            // 
            // lblStatus
            // 
            lblStatus.AutoSize = true;
            lblStatus.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            lblStatus.ForeColor = Color.LimeGreen;
            lblStatus.Location = new Point(14, 77);
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(42, 15);
            lblStatus.TabIndex = 1;
            lblStatus.Text = "Status";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(14, 51);
            label2.Name = "label2";
            label2.Size = new Size(41, 15);
            label2.TabIndex = 1;
            label2.Text = "Dự án:";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.BackColor = Color.Transparent;
            label1.Dock = DockStyle.Top;
            label1.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            label1.ForeColor = Color.Blue;
            label1.Location = new Point(3, 18);
            label1.Name = "label1";
            label1.Size = new Size(441, 21);
            label1.TabIndex = 0;
            label1.Text = "CHỌN VẬT TƯ TỪ KHO ĐỂ XUẤT - Click vào dòng để xuất";
            // 
            // groupBox2
            // 
            groupBox2.BackColor = Color.White;
            groupBox2.Controls.Add(dgvExportQue);
            groupBox2.Controls.Add(btnClear);
            groupBox2.Controls.Add(button1);
            groupBox2.Controls.Add(btnXoaRow);
            groupBox2.Controls.Add(lblInfoKho);
            groupBox2.Controls.Add(label5);
            groupBox2.Dock = DockStyle.Top;
            groupBox2.Location = new Point(0, 369);
            groupBox2.Margin = new Padding(3, 2, 3, 2);
            groupBox2.Name = "groupBox2";
            groupBox2.Padding = new Padding(3, 2, 3, 2);
            groupBox2.Size = new Size(1600, 493);
            groupBox2.TabIndex = 1;
            groupBox2.TabStop = false;
            // 
            // dgvExportQue
            // 
            dgvExportQue.AllowUserToAddRows = false;
            dgvExportQue.AllowUserToDeleteRows = false;
            dgvExportQue.AllowUserToOrderColumns = true;
            dgvExportQue.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dgvExportQue.BackgroundColor = Color.White;
            dgvExportQue.BorderStyle = BorderStyle.None;
            dgvExportQue.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Sunken;
            dgvExportQue.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvExportQue.Location = new Point(14, 49);
            dgvExportQue.Margin = new Padding(3, 2, 3, 2);
            dgvExportQue.Name = "dgvExportQue";
            dgvExportQue.RowHeadersWidth = 51;
            dgvExportQue.Size = new Size(1570, 440);
            dgvExportQue.TabIndex = 5;
            dgvExportQue.CellFormatting += dgvExportQue_CellFormatting;
            // 
            // btnClear
            // 
            btnClear.BackColor = Color.FromArgb(108, 117, 125);
            btnClear.FlatStyle = FlatStyle.Flat;
            btnClear.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnClear.ForeColor = Color.White;
            btnClear.Location = new Point(336, 16);
            btnClear.Margin = new Padding(3, 2, 3, 2);
            btnClear.Name = "btnClear";
            btnClear.Size = new Size(109, 29);
            btnClear.TabIndex = 4;
            btnClear.Text = "🔄 Xóa form";
            btnClear.UseVisualStyleBackColor = false;
            btnClear.Click += btnClear_Click;
            // 
            // button1
            // 
            button1.BackColor = Color.FromArgb(255, 128, 0);
            button1.FlatStyle = FlatStyle.Flat;
            button1.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            button1.ForeColor = Color.White;
            button1.Location = new Point(217, 16);
            button1.Margin = new Padding(3, 2, 3, 2);
            button1.Name = "button1";
            button1.Size = new Size(109, 29);
            button1.TabIndex = 4;
            button1.Text = "💾 Lưu xuất kho";
            button1.UseVisualStyleBackColor = false;
            button1.Click += btnXuatKHO_Click;
            // 
            // btnXoaRow
            // 
            btnXoaRow.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnXoaRow.BackColor = Color.FromArgb(220, 53, 69);
            btnXoaRow.FlatStyle = FlatStyle.Flat;
            btnXoaRow.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnXoaRow.ForeColor = Color.White;
            btnXoaRow.Location = new Point(1475, 16);
            btnXoaRow.Margin = new Padding(3, 2, 3, 2);
            btnXoaRow.Name = "btnXoaRow";
            btnXoaRow.Size = new Size(109, 29);
            btnXoaRow.TabIndex = 4;
            btnXoaRow.Text = "🗑 Xóa dòng";
            btnXoaRow.UseVisualStyleBackColor = false;
            btnXoaRow.Click += btnXoaRow_Click;
            // 
            // lblInfoKho
            // 
            lblInfoKho.AutoSize = true;
            lblInfoKho.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            lblInfoKho.ForeColor = Color.LimeGreen;
            lblInfoKho.Location = new Point(460, 24);
            lblInfoKho.Name = "lblInfoKho";
            lblInfoKho.Size = new Size(42, 15);
            lblInfoKho.TabIndex = 1;
            lblInfoKho.Text = "Status";
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.BackColor = Color.Transparent;
            label5.Dock = DockStyle.Top;
            label5.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            label5.ForeColor = Color.FromArgb(220, 53, 69);
            label5.Location = new Point(3, 18);
            label5.Name = "label5";
            label5.Size = new Size(182, 21);
            label5.TabIndex = 0;
            label5.Text = "THÔNG TIN XUẤT KHO";
            // 
            // groupBox3
            // 
            groupBox3.BackColor = Color.White;
            groupBox3.Controls.Add(txtSearch);
            groupBox3.Controls.Add(dtpTo);
            groupBox3.Controls.Add(dtpStart);
            groupBox3.Controls.Add(dgvHis);
            groupBox3.Controls.Add(btnClearSearch);
            groupBox3.Controls.Add(btnSearchHis);
            groupBox3.Controls.Add(lblInfoXK);
            groupBox3.Controls.Add(label6);
            groupBox3.Controls.Add(label3);
            groupBox3.Controls.Add(cboProjectCheck);
            groupBox3.Controls.Add(label7);
            groupBox3.Dock = DockStyle.Top;
            groupBox3.Location = new Point(0, 862);
            groupBox3.Margin = new Padding(3, 2, 3, 2);
            groupBox3.Name = "groupBox3";
            groupBox3.Padding = new Padding(3, 2, 3, 2);
            groupBox3.Size = new Size(1600, 549);
            groupBox3.TabIndex = 1;
            groupBox3.TabStop = false;
            groupBox3.Enter += groupBox3_Enter;
            // 
            // txtSearch
            // 
            txtSearch.Location = new Point(820, 21);
            txtSearch.Margin = new Padding(3, 2, 3, 2);
            txtSearch.Name = "txtSearch";
            txtSearch.Size = new Size(134, 23);
            txtSearch.TabIndex = 7;
            // 
            // dtpTo
            // 
            dtpTo.Format = DateTimePickerFormat.Short;
            dtpTo.Location = new Point(425, 21);
            dtpTo.Margin = new Padding(3, 2, 3, 2);
            dtpTo.Name = "dtpTo";
            dtpTo.Size = new Size(114, 23);
            dtpTo.TabIndex = 6;
            // 
            // dtpStart
            // 
            dtpStart.Format = DateTimePickerFormat.Short;
            dtpStart.Location = new Point(241, 21);
            dtpStart.Margin = new Padding(3, 2, 3, 2);
            dtpStart.Name = "dtpStart";
            dtpStart.Size = new Size(114, 23);
            dtpStart.TabIndex = 6;
            // 
            // dgvHis
            // 
            dgvHis.AllowUserToAddRows = false;
            dgvHis.AllowUserToDeleteRows = false;
            dgvHis.AllowUserToOrderColumns = true;
            dgvHis.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dgvHis.BackgroundColor = Color.White;
            dgvHis.BorderStyle = BorderStyle.None;
            dgvHis.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Sunken;
            dgvHis.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvHis.Location = new Point(14, 74);
            dgvHis.Margin = new Padding(3, 2, 3, 2);
            dgvHis.Name = "dgvHis";
            dgvHis.ReadOnly = true;
            dgvHis.RowHeadersWidth = 51;
            dgvHis.Size = new Size(1570, 463);
            dgvHis.TabIndex = 5;
            dgvHis.CellFormatting += dgvHis_CellFormatting;
            // 
            // btnClearSearch
            // 
            btnClearSearch.BackColor = Color.FromArgb(192, 0, 0);
            btnClearSearch.FlatStyle = FlatStyle.Flat;
            btnClearSearch.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnClearSearch.ForeColor = Color.White;
            btnClearSearch.Location = new Point(1072, 18);
            btnClearSearch.Margin = new Padding(3, 2, 3, 2);
            btnClearSearch.Name = "btnClearSearch";
            btnClearSearch.Size = new Size(94, 29);
            btnClearSearch.TabIndex = 4;
            btnClearSearch.Text = "🗑 Xóa lọc";
            btnClearSearch.UseVisualStyleBackColor = false;
            // 
            // btnSearchHis
            // 
            btnSearchHis.BackColor = Color.RoyalBlue;
            btnSearchHis.FlatStyle = FlatStyle.Flat;
            btnSearchHis.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnSearchHis.ForeColor = Color.White;
            btnSearchHis.Location = new Point(973, 18);
            btnSearchHis.Margin = new Padding(3, 2, 3, 2);
            btnSearchHis.Name = "btnSearchHis";
            btnSearchHis.Size = new Size(94, 29);
            btnSearchHis.TabIndex = 3;
            btnSearchHis.Text = "🔍 Tìm kiếm";
            btnSearchHis.UseVisualStyleBackColor = false;
            btnSearchHis.Click += btnSearchHis_Click;
            // 
            // lblInfoXK
            // 
            lblInfoXK.AutoSize = true;
            lblInfoXK.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            lblInfoXK.ForeColor = Color.FromArgb(0, 120, 212);
            lblInfoXK.Location = new Point(14, 52);
            lblInfoXK.Name = "lblInfoXK";
            lblInfoXK.Size = new Size(42, 15);
            lblInfoXK.TabIndex = 1;
            lblInfoXK.Text = "Status";
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.BackColor = Color.Transparent;
            label6.Dock = DockStyle.Top;
            label6.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            label6.ForeColor = Color.FromArgb(220, 53, 69);
            label6.Location = new Point(3, 18);
            label6.Name = "label6";
            label6.Size = new Size(155, 21);
            label6.TabIndex = 0;
            label6.Text = "LỊCH SỬ XUẤT KHO";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(372, 23);
            label3.Name = "label3";
            label3.Size = new Size(41, 15);
            label3.TabIndex = 1;
            label3.Text = "Dự án:";
            // 
            // cboProjectCheck
            // 
            cboProjectCheck.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cboProjectCheck.AutoCompleteSource = AutoCompleteSource.ListItems;
            cboProjectCheck.FormattingEnabled = true;
            cboProjectCheck.Location = new Point(559, 20);
            cboProjectCheck.Margin = new Padding(3, 2, 3, 2);
            cboProjectCheck.Name = "cboProjectCheck";
            cboProjectCheck.Size = new Size(238, 23);
            cboProjectCheck.TabIndex = 2;
            cboProjectCheck.Validating += cboProjectCheck_Validating;
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new Point(190, 23);
            label7.Name = "label7";
            label7.Size = new Size(41, 15);
            label7.TabIndex = 1;
            label7.Text = "Dự án:";
            // 
            // ucExportWarehouse
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.White;
            Controls.Add(groupBox3);
            Controls.Add(groupBox2);
            Controls.Add(groupBox1);
            Margin = new Padding(3, 2, 3, 2);
            Name = "ucExportWarehouse";
            Size = new Size(1600, 1429);
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)dgvKho).EndInit();
            groupBox2.ResumeLayout(false);
            groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)dgvExportQue).EndInit();
            groupBox3.ResumeLayout(false);
            groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)dgvHis).EndInit();
            ResumeLayout(false);
        }

        #endregion

        private GroupBox groupBox1;
        private Button btnCancelSer;
        private Button btnSearch;
        private ComboBox cboProject;
        private Label label2;
        private Label label1;
        private Label lblStatus;
        private GroupBox groupBox2;
        private DataGridView dgvExportQue;
        private Label lblInfoKho;
        private Label label5;
        private GroupBox groupBox3;
        private DataGridView dgvHis;
        private Label lblInfoXK;
        private Label label6;
        private Button btnClearSearch;
        private Button btnSearchHis;
        private ComboBox cboProjectCheck;
        private Label label7;
        private DateTimePicker dtpTo;
        private DateTimePicker dtpStart;
        private Label label3;
        private TextBox txtSearch;
        private Button btnXoaRow;
        private DataGridView dgvKho;
        private Button btnClear;
        private Button button1;
    }
}
