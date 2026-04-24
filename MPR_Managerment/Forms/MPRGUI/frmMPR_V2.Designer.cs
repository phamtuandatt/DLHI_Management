namespace MPR_Managerment.Forms.MPRGUI
{
    partial class frmMPR_V2
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            panel1 = new Panel();
            groupBox2 = new GroupBox();
            dgvItems = new DataGridView();
            groupBox1 = new GroupBox();
            panel2 = new Panel();
            tableLayoutPanel2 = new TableLayoutPanel();
            btnSave = new Button();
            btnCancel = new Button();
            btnDeleteRow = new Button();
            btnAddRow = new Button();
            lblTotal = new Label();
            tableLayoutPanel1 = new TableLayoutPanel();
            dateTimePicker2 = new DateTimePicker();
            label1 = new Label();
            label2 = new Label();
            label3 = new Label();
            label4 = new Label();
            label5 = new Label();
            label6 = new Label();
            label7 = new Label();
            label8 = new Label();
            dateTimePicker1 = new DateTimePicker();
            comboBox1 = new ComboBox();
            comboBox2 = new ComboBox();
            btnPrint = new Button();
            panel1.SuspendLayout();
            groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvItems).BeginInit();
            groupBox1.SuspendLayout();
            panel2.SuspendLayout();
            tableLayoutPanel2.SuspendLayout();
            tableLayoutPanel1.SuspendLayout();
            SuspendLayout();
            // 
            // panel1
            // 
            panel1.BackColor = Color.White;
            panel1.Controls.Add(groupBox2);
            panel1.Controls.Add(groupBox1);
            panel1.Dock = DockStyle.Fill;
            panel1.Location = new Point(0, 0);
            panel1.Margin = new Padding(10);
            panel1.Name = "panel1";
            panel1.Padding = new Padding(10);
            panel1.Size = new Size(1271, 726);
            panel1.TabIndex = 0;
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(dgvItems);
            groupBox2.Dock = DockStyle.Fill;
            groupBox2.Location = new Point(10, 162);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new Size(1251, 554);
            groupBox2.TabIndex = 1;
            groupBox2.TabStop = false;
            groupBox2.Text = "Danh sách vật tư";
            // 
            // dgvItems
            // 
            dgvItems.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvItems.Dock = DockStyle.Fill;
            dgvItems.Location = new Point(3, 19);
            dgvItems.Name = "dgvItems";
            dgvItems.Size = new Size(1245, 532);
            dgvItems.TabIndex = 0;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(panel2);
            groupBox1.Controls.Add(tableLayoutPanel1);
            groupBox1.Dock = DockStyle.Top;
            groupBox1.Location = new Point(10, 10);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(1251, 152);
            groupBox1.TabIndex = 0;
            groupBox1.TabStop = false;
            groupBox1.Text = "★ Yêu cầu mua hàng mới";
            // 
            // panel2
            // 
            panel2.Controls.Add(tableLayoutPanel2);
            panel2.Dock = DockStyle.Fill;
            panel2.Location = new Point(3, 112);
            panel2.Name = "panel2";
            panel2.Size = new Size(1245, 37);
            panel2.TabIndex = 1;
            // 
            // tableLayoutPanel2
            // 
            tableLayoutPanel2.ColumnCount = 7;
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 229F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel2.Controls.Add(btnPrint, 0, 0);
            tableLayoutPanel2.Controls.Add(btnSave, 1, 0);
            tableLayoutPanel2.Controls.Add(btnCancel, 2, 0);
            tableLayoutPanel2.Controls.Add(btnDeleteRow, 4, 0);
            tableLayoutPanel2.Controls.Add(btnAddRow, 3, 0);
            tableLayoutPanel2.Controls.Add(lblTotal, 5, 0);
            tableLayoutPanel2.Dock = DockStyle.Fill;
            tableLayoutPanel2.Location = new Point(0, 0);
            tableLayoutPanel2.Name = "tableLayoutPanel2";
            tableLayoutPanel2.RowCount = 1;
            tableLayoutPanel2.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel2.Size = new Size(1245, 37);
            tableLayoutPanel2.TabIndex = 5;
            // 
            // btnSave
            // 
            btnSave.Dock = DockStyle.Fill;
            btnSave.Location = new Point(153, 3);
            btnSave.Name = "btnSave";
            btnSave.Size = new Size(144, 31);
            btnSave.TabIndex = 0;
            btnSave.Text = "Lưu";
            btnSave.UseVisualStyleBackColor = true;
            btnSave.Click += btnSave_Click;
            // 
            // btnCancel
            // 
            btnCancel.Dock = DockStyle.Fill;
            btnCancel.Location = new Point(303, 3);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new Size(144, 31);
            btnCancel.TabIndex = 0;
            btnCancel.Text = "Hủy";
            btnCancel.UseVisualStyleBackColor = true;
            btnCancel.Click += btnCancel_Click;
            // 
            // btnDeleteRow
            // 
            btnDeleteRow.Dock = DockStyle.Fill;
            btnDeleteRow.Location = new Point(603, 3);
            btnDeleteRow.Name = "btnDeleteRow";
            btnDeleteRow.Size = new Size(144, 31);
            btnDeleteRow.TabIndex = 0;
            btnDeleteRow.Text = "Xóa dòng";
            btnDeleteRow.UseVisualStyleBackColor = true;
            btnDeleteRow.Click += btnDeleteRow_Click;
            // 
            // btnAddRow
            // 
            btnAddRow.Dock = DockStyle.Fill;
            btnAddRow.Location = new Point(453, 3);
            btnAddRow.Name = "btnAddRow";
            btnAddRow.Size = new Size(144, 31);
            btnAddRow.TabIndex = 0;
            btnAddRow.Text = "Thêm dòng";
            btnAddRow.UseVisualStyleBackColor = true;
            btnAddRow.Click += btnAddRow_Click;
            // 
            // lblTotal
            // 
            lblTotal.AutoSize = true;
            lblTotal.Dock = DockStyle.Fill;
            lblTotal.Location = new Point(753, 3);
            lblTotal.Margin = new Padding(3);
            lblTotal.Name = "lblTotal";
            lblTotal.Size = new Size(223, 31);
            lblTotal.TabIndex = 1;
            lblTotal.Text = "TỔNG";
            lblTotal.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.ColumnCount = 5;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 250F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 250F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.Controls.Add(dateTimePicker2, 3, 0);
            tableLayoutPanel1.Controls.Add(label1, 2, 0);
            tableLayoutPanel1.Controls.Add(label2, 0, 0);
            tableLayoutPanel1.Controls.Add(label3, 0, 1);
            tableLayoutPanel1.Controls.Add(label4, 0, 2);
            tableLayoutPanel1.Controls.Add(label5, 0, 3);
            tableLayoutPanel1.Controls.Add(label6, 2, 1);
            tableLayoutPanel1.Controls.Add(label7, 2, 2);
            tableLayoutPanel1.Controls.Add(label8, 2, 3);
            tableLayoutPanel1.Controls.Add(dateTimePicker1, 1, 0);
            tableLayoutPanel1.Controls.Add(comboBox1, 1, 1);
            tableLayoutPanel1.Controls.Add(comboBox2, 3, 1);
            tableLayoutPanel1.Dock = DockStyle.Top;
            tableLayoutPanel1.Location = new Point(3, 19);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 4;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.Size = new Size(1245, 93);
            tableLayoutPanel1.TabIndex = 0;
            // 
            // dateTimePicker2
            // 
            dateTimePicker2.Dock = DockStyle.Fill;
            dateTimePicker2.Format = DateTimePickerFormat.Short;
            dateTimePicker2.Location = new Point(492, 5);
            dateTimePicker2.Margin = new Padding(2, 5, 5, 5);
            dateTimePicker2.Name = "dateTimePicker2";
            dateTimePicker2.Size = new Size(243, 23);
            dateTimePicker2.TabIndex = 2;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Dock = DockStyle.Fill;
            label1.Location = new Point(373, 5);
            label1.Margin = new Padding(3, 5, 15, 5);
            label1.Name = "label1";
            label1.Size = new Size(102, 20);
            label1.TabIndex = 0;
            label1.Text = "Ngày giao:";
            label1.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Dock = DockStyle.Fill;
            label2.Location = new Point(3, 5);
            label2.Margin = new Padding(3, 5, 15, 5);
            label2.Name = "label2";
            label2.Size = new Size(102, 20);
            label2.TabIndex = 0;
            label2.Text = "Ngày tạo:";
            label2.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Dock = DockStyle.Fill;
            label3.Location = new Point(3, 35);
            label3.Margin = new Padding(3, 5, 15, 5);
            label3.Name = "label3";
            label3.Size = new Size(102, 20);
            label3.TabIndex = 0;
            label3.Text = "Bộ phận:";
            label3.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Dock = DockStyle.Fill;
            label4.Location = new Point(3, 65);
            label4.Margin = new Padding(3, 5, 15, 5);
            label4.Name = "label4";
            label4.Size = new Size(102, 20);
            label4.TabIndex = 0;
            label4.Text = "label1";
            label4.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Dock = DockStyle.Fill;
            label5.Location = new Point(3, 95);
            label5.Margin = new Padding(3, 5, 15, 5);
            label5.Name = "label5";
            label5.Size = new Size(102, 1);
            label5.TabIndex = 0;
            label5.Text = "label1";
            label5.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Dock = DockStyle.Fill;
            label6.Location = new Point(373, 35);
            label6.Margin = new Padding(3, 5, 15, 5);
            label6.Name = "label6";
            label6.Size = new Size(102, 20);
            label6.TabIndex = 0;
            label6.Text = "Dự án:";
            label6.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Dock = DockStyle.Fill;
            label7.Location = new Point(373, 65);
            label7.Margin = new Padding(3, 5, 15, 5);
            label7.Name = "label7";
            label7.Size = new Size(102, 20);
            label7.TabIndex = 0;
            label7.Text = "MPR No:";
            label7.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // label8
            // 
            label8.AutoSize = true;
            label8.Dock = DockStyle.Fill;
            label8.Location = new Point(373, 95);
            label8.Margin = new Padding(3, 5, 15, 5);
            label8.Name = "label8";
            label8.Size = new Size(102, 1);
            label8.TabIndex = 0;
            label8.Text = "label1";
            label8.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // dateTimePicker1
            // 
            dateTimePicker1.Dock = DockStyle.Fill;
            dateTimePicker1.Format = DateTimePickerFormat.Short;
            dateTimePicker1.Location = new Point(122, 5);
            dateTimePicker1.Margin = new Padding(2, 5, 5, 5);
            dateTimePicker1.Name = "dateTimePicker1";
            dateTimePicker1.Size = new Size(243, 23);
            dateTimePicker1.TabIndex = 1;
            // 
            // comboBox1
            // 
            comboBox1.Dock = DockStyle.Fill;
            comboBox1.FormattingEnabled = true;
            comboBox1.Location = new Point(122, 35);
            comboBox1.Margin = new Padding(2, 5, 5, 5);
            comboBox1.Name = "comboBox1";
            comboBox1.Size = new Size(243, 23);
            comboBox1.TabIndex = 3;
            // 
            // comboBox2
            // 
            comboBox2.Dock = DockStyle.Fill;
            comboBox2.FormattingEnabled = true;
            comboBox2.Location = new Point(492, 35);
            comboBox2.Margin = new Padding(2, 5, 5, 5);
            comboBox2.Name = "comboBox2";
            comboBox2.Size = new Size(243, 23);
            comboBox2.TabIndex = 4;
            // 
            // btnPrint
            // 
            btnPrint.Dock = DockStyle.Fill;
            btnPrint.Location = new Point(3, 3);
            btnPrint.Name = "btnPrint";
            btnPrint.Size = new Size(144, 31);
            btnPrint.TabIndex = 2;
            btnPrint.Text = "Xuất File";
            btnPrint.UseVisualStyleBackColor = true;
            // 
            // frmMPR_V2
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1271, 726);
            Controls.Add(panel1);
            Name = "frmMPR_V2";
            Text = "frmMPR_V2";
            panel1.ResumeLayout(false);
            groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvItems).EndInit();
            groupBox1.ResumeLayout(false);
            panel2.ResumeLayout(false);
            tableLayoutPanel2.ResumeLayout(false);
            tableLayoutPanel2.PerformLayout();
            tableLayoutPanel1.ResumeLayout(false);
            tableLayoutPanel1.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private Panel panel1;
        private GroupBox groupBox1;
        private TableLayoutPanel tableLayoutPanel1;
        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
        private Label label5;
        private Label label6;
        private Label label7;
        private Label label8;
        private GroupBox groupBox2;
        private Panel panel2;
        private DataGridView dgvItems;
        private DateTimePicker dateTimePicker2;
        private DateTimePicker dateTimePicker1;
        private ComboBox comboBox1;
        private ComboBox comboBox2;
        private TableLayoutPanel tableLayoutPanel2;
        private Button btnSave;
        private Button btnCancel;
        private Button btnDeleteRow;
        private Button btnAddRow;
        private Label lblTotal;
        private Button btnPrint;
    }
}