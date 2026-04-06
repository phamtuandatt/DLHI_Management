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
            groupBox2 = new GroupBox();
            dgvList = new DataGridView();
            groupBox1 = new GroupBox();
            btnCancelSer = new Button();
            btnSearch = new Button();
            cboPO = new ComboBox();
            label1 = new Label();
            cboProject = new ComboBox();
            label2 = new Label();
            panel1.SuspendLayout();
            groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvList).BeginInit();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // panel1
            // 
            panel1.BackColor = Color.White;
            panel1.Controls.Add(groupBox2);
            panel1.Controls.Add(groupBox1);
            panel1.Dock = DockStyle.Fill;
            panel1.Location = new Point(0, 0);
            panel1.Name = "panel1";
            panel1.Size = new Size(989, 624);
            panel1.TabIndex = 0;
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(dgvList);
            groupBox2.Dock = DockStyle.Fill;
            groupBox2.Location = new Point(0, 43);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new Size(989, 581);
            groupBox2.TabIndex = 0;
            groupBox2.TabStop = false;
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
            dgvList.Location = new Point(3, 19);
            dgvList.Margin = new Padding(3, 2, 3, 2);
            dgvList.Name = "dgvList";
            dgvList.ReadOnly = true;
            dgvList.RowHeadersWidth = 51;
            dgvList.Size = new Size(983, 559);
            dgvList.TabIndex = 6;
            dgvList.CellClick += dgvList_CellClick;
            dgvList.Scroll += dgvList_Scroll;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(btnCancelSer);
            groupBox1.Controls.Add(btnSearch);
            groupBox1.Controls.Add(cboPO);
            groupBox1.Controls.Add(label1);
            groupBox1.Controls.Add(cboProject);
            groupBox1.Controls.Add(label2);
            groupBox1.Dock = DockStyle.Top;
            groupBox1.Location = new Point(0, 0);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(989, 43);
            groupBox1.TabIndex = 0;
            groupBox1.TabStop = false;
            // 
            // btnCancelSer
            // 
            btnCancelSer.BackColor = Color.DarkGray;
            btnCancelSer.FlatStyle = FlatStyle.Flat;
            btnCancelSer.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnCancelSer.ForeColor = Color.White;
            btnCancelSer.Location = new Point(775, 8);
            btnCancelSer.Margin = new Padding(3, 2, 3, 2);
            btnCancelSer.Name = "btnCancelSer";
            btnCancelSer.Size = new Size(101, 33);
            btnCancelSer.TabIndex = 8;
            btnCancelSer.Text = "✖ Xóa lọc";
            btnCancelSer.UseVisualStyleBackColor = false;
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
            Size = new Size(989, 624);
            Load += ucFillInvoiceNo_Load;
            panel1.ResumeLayout(false);
            groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvList).EndInit();
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private Panel panel1;
        private GroupBox groupBox2;
        private GroupBox groupBox1;
        private Button btnCancelSer;
        private Button btnSearch;
        private ComboBox cboPO;
        private Label label1;
        private ComboBox cboProject;
        private Label label2;
        private DataGridView dgvList;
    }
}
