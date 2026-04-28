namespace MPR_Managerment.Forms.MPRGUI
{
    partial class frmSelectItem
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
            tableLayoutPanel1 = new TableLayoutPanel();
            tableLayoutPanel2 = new TableLayoutPanel();
            label2 = new Label();
            label1 = new Label();
            txtSearch = new TextBox();
            btnSearch = new Button();
            btnRefresh = new Button();
            cboMaterial = new ComboBox();
            tableLayoutPanel3 = new TableLayoutPanel();
            lblStatus = new Label();
            groupBox1 = new GroupBox();
            dgvItems = new DataGridView();
            tableLayoutPanel4 = new TableLayoutPanel();
            btnSelect = new Button();
            btnDelete = new Button();
            btnCancels = new Button();
            button6 = new Button();
            button7 = new Button();
            tableLayoutPanel1.SuspendLayout();
            tableLayoutPanel2.SuspendLayout();
            tableLayoutPanel3.SuspendLayout();
            groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvItems).BeginInit();
            tableLayoutPanel4.SuspendLayout();
            SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.ColumnCount = 1;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.Controls.Add(tableLayoutPanel2, 0, 0);
            tableLayoutPanel1.Controls.Add(tableLayoutPanel3, 0, 1);
            tableLayoutPanel1.Controls.Add(groupBox1, 0, 2);
            tableLayoutPanel1.Controls.Add(tableLayoutPanel4, 0, 3);
            tableLayoutPanel1.Dock = DockStyle.Fill;
            tableLayoutPanel1.Location = new Point(0, 0);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 5;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 38F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 26F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 20F));
            tableLayoutPanel1.Size = new Size(1207, 634);
            tableLayoutPanel1.TabIndex = 0;
            // 
            // tableLayoutPanel2
            // 
            tableLayoutPanel2.ColumnCount = 7;
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 156F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 294F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 249F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel2.Controls.Add(label2, 2, 0);
            tableLayoutPanel2.Controls.Add(label1, 0, 0);
            tableLayoutPanel2.Controls.Add(txtSearch, 1, 0);
            tableLayoutPanel2.Controls.Add(btnSearch, 4, 0);
            tableLayoutPanel2.Controls.Add(btnRefresh, 5, 0);
            tableLayoutPanel2.Controls.Add(cboMaterial, 3, 0);
            tableLayoutPanel2.Dock = DockStyle.Fill;
            tableLayoutPanel2.Location = new Point(3, 3);
            tableLayoutPanel2.Name = "tableLayoutPanel2";
            tableLayoutPanel2.RowCount = 1;
            tableLayoutPanel2.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel2.Size = new Size(1201, 32);
            tableLayoutPanel2.TabIndex = 0;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Dock = DockStyle.Fill;
            label2.Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point, 0);
            label2.Location = new Point(453, 0);
            label2.Name = "label2";
            label2.Size = new Size(54, 32);
            label2.TabIndex = 3;
            label2.Text = "Loại:";
            label2.TextAlign = ContentAlignment.MiddleRight;
            label2.Visible = false;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Dock = DockStyle.Fill;
            label1.Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point, 0);
            label1.Location = new Point(3, 0);
            label1.Name = "label1";
            label1.Size = new Size(150, 32);
            label1.TabIndex = 0;
            label1.Text = "Tìm kiếm mặt hàng:";
            label1.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // txtSearch
            // 
            txtSearch.Dock = DockStyle.Fill;
            txtSearch.Location = new Point(159, 5);
            txtSearch.Margin = new Padding(3, 5, 3, 3);
            txtSearch.Name = "txtSearch";
            txtSearch.Size = new Size(288, 23);
            txtSearch.TabIndex = 1;
            txtSearch.KeyDown += txtSearch_KeyDown;
            // 
            // btnSearch
            // 
            btnSearch.Dock = DockStyle.Fill;
            btnSearch.Location = new Point(762, 3);
            btnSearch.Name = "btnSearch";
            btnSearch.Size = new Size(144, 26);
            btnSearch.TabIndex = 2;
            btnSearch.Text = "Tìm";
            btnSearch.UseVisualStyleBackColor = true;
            btnSearch.Visible = false;
            btnSearch.Click += btnSearch_Click;
            // 
            // btnRefresh
            // 
            btnRefresh.Dock = DockStyle.Fill;
            btnRefresh.Location = new Point(912, 3);
            btnRefresh.Name = "btnRefresh";
            btnRefresh.Size = new Size(144, 26);
            btnRefresh.TabIndex = 2;
            btnRefresh.Text = "Làm mới";
            btnRefresh.UseVisualStyleBackColor = true;
            btnRefresh.Visible = false;
            // 
            // cboMaterial
            // 
            cboMaterial.Dock = DockStyle.Fill;
            cboMaterial.FormattingEnabled = true;
            cboMaterial.Location = new Point(513, 5);
            cboMaterial.Margin = new Padding(3, 5, 3, 3);
            cboMaterial.Name = "cboMaterial";
            cboMaterial.Size = new Size(243, 23);
            cboMaterial.TabIndex = 4;
            cboMaterial.Visible = false;
            // 
            // tableLayoutPanel3
            // 
            tableLayoutPanel3.ColumnCount = 11;
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 48F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 11F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20F));
            tableLayoutPanel3.Controls.Add(lblStatus, 0, 0);
            tableLayoutPanel3.Dock = DockStyle.Fill;
            tableLayoutPanel3.Location = new Point(3, 41);
            tableLayoutPanel3.Name = "tableLayoutPanel3";
            tableLayoutPanel3.RowCount = 1;
            tableLayoutPanel3.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel3.Size = new Size(1201, 20);
            tableLayoutPanel3.TabIndex = 1;
            // 
            // lblStatus
            // 
            lblStatus.AutoSize = true;
            lblStatus.BackColor = Color.White;
            lblStatus.Dock = DockStyle.Fill;
            lblStatus.Font = new Font("Segoe UI", 11.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            lblStatus.ForeColor = Color.FromArgb(48, 200, 82);
            lblStatus.Location = new Point(3, 0);
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(976, 20);
            lblStatus.TabIndex = 1;
            lblStatus.Text = "Đã chọn: 0 vật tư";
            lblStatus.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(dgvItems);
            groupBox1.Dock = DockStyle.Fill;
            groupBox1.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            groupBox1.Location = new Point(3, 67);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(1201, 494);
            groupBox1.TabIndex = 2;
            groupBox1.TabStop = false;
            groupBox1.Text = "Danh sách vật tư";
            // 
            // dgvItems
            // 
            dgvItems.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvItems.Dock = DockStyle.Fill;
            dgvItems.Location = new Point(3, 23);
            dgvItems.Name = "dgvItems";
            dgvItems.RowTemplate.Height = 30;
            dgvItems.Size = new Size(1195, 468);
            dgvItems.TabIndex = 0;
            // 
            // tableLayoutPanel4
            // 
            tableLayoutPanel4.ColumnCount = 7;
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100F));
            tableLayoutPanel4.Controls.Add(btnSelect, 0, 0);
            tableLayoutPanel4.Controls.Add(btnDelete, 1, 0);
            tableLayoutPanel4.Controls.Add(btnCancels, 2, 0);
            tableLayoutPanel4.Controls.Add(button6, 3, 0);
            tableLayoutPanel4.Controls.Add(button7, 4, 0);
            tableLayoutPanel4.Dock = DockStyle.Fill;
            tableLayoutPanel4.Location = new Point(3, 567);
            tableLayoutPanel4.Name = "tableLayoutPanel4";
            tableLayoutPanel4.RowCount = 1;
            tableLayoutPanel4.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel4.Size = new Size(1201, 44);
            tableLayoutPanel4.TabIndex = 3;
            // 
            // btnSelect
            // 
            btnSelect.Dock = DockStyle.Fill;
            btnSelect.Location = new Point(3, 3);
            btnSelect.Name = "btnSelect";
            btnSelect.Size = new Size(94, 38);
            btnSelect.TabIndex = 0;
            btnSelect.Text = "Xác nhận";
            btnSelect.UseVisualStyleBackColor = true;
            btnSelect.Click += btnSelect_Click;
            // 
            // btnDelete
            // 
            btnDelete.Dock = DockStyle.Fill;
            btnDelete.Location = new Point(103, 3);
            btnDelete.Name = "btnDelete";
            btnDelete.Size = new Size(94, 38);
            btnDelete.TabIndex = 0;
            btnDelete.Text = "Bỏ chọn";
            btnDelete.UseVisualStyleBackColor = true;
            btnDelete.Click += btnDelete_Click;
            // 
            // btnCancels
            // 
            btnCancels.Dock = DockStyle.Fill;
            btnCancels.Location = new Point(203, 3);
            btnCancels.Name = "btnCancels";
            btnCancels.Size = new Size(94, 38);
            btnCancels.TabIndex = 0;
            btnCancels.Text = "Đóng";
            btnCancels.UseVisualStyleBackColor = true;
            btnCancels.Click += btnCancels_Click;
            // 
            // button6
            // 
            button6.Dock = DockStyle.Fill;
            button6.Location = new Point(303, 3);
            button6.Name = "button6";
            button6.Size = new Size(94, 38);
            button6.TabIndex = 0;
            button6.Text = "button3";
            button6.UseVisualStyleBackColor = true;
            button6.Visible = false;
            // 
            // button7
            // 
            button7.Dock = DockStyle.Fill;
            button7.Location = new Point(403, 3);
            button7.Name = "button7";
            button7.Size = new Size(94, 38);
            button7.TabIndex = 0;
            button7.Text = "button3";
            button7.UseVisualStyleBackColor = true;
            button7.Visible = false;
            // 
            // frmSelectItem
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.White;
            ClientSize = new Size(1207, 634);
            Controls.Add(tableLayoutPanel1);
            MinimumSize = new Size(900, 620);
            Name = "frmSelectItem";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Tìm kiếm mặt hàng";
            Load += frmSelectItem_Load_1;
            tableLayoutPanel1.ResumeLayout(false);
            tableLayoutPanel2.ResumeLayout(false);
            tableLayoutPanel2.PerformLayout();
            tableLayoutPanel3.ResumeLayout(false);
            tableLayoutPanel3.PerformLayout();
            groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvItems).EndInit();
            tableLayoutPanel4.ResumeLayout(false);
            ResumeLayout(false);
        }

        private System.Windows.Forms.Button CreatePagerButton(string text)
        {
            var button = new System.Windows.Forms.Button();
            button.AutoSize = true;
            button.FlatStyle = System.Windows.Forms.FlatStyle.System;
            button.Font = new System.Drawing.Font("Segoe UI", 8.25F);
            button.Margin = new System.Windows.Forms.Padding(0, 0, 6, 0);
            button.MinimumSize = new System.Drawing.Size(30, 28);
            button.Size = new System.Drawing.Size(30, 28);
            button.Text = text;
            return button;
        }

        #endregion

        private TableLayoutPanel tableLayoutPanel1;
        private TableLayoutPanel tableLayoutPanel2;
        private Label label1;
        private TextBox txtSearch;
        private Button btnSearch;
        private Button btnRefresh;
        private TableLayoutPanel tableLayoutPanel3;
        private GroupBox groupBox1;
        private DataGridView dgvItems;
        private TableLayoutPanel tableLayoutPanel4;
        private Button btnSelect;
        private Button btnDelete;
        private Button btnCancels;
        private Button button6;
        private Button button7;
        private Label label2;
        private ComboBox cboMaterial;
        private Label lblStatus;
    }
}