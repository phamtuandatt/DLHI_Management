namespace MPR_Managerment.Forms.ItemCodeGUI
{
    partial class frmAddMaterialDetail
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
            lblName = new Label();
            tableLayoutPanel5 = new TableLayoutPanel();
            btnGenerate = new Button();
            label6 = new Label();
            txtCode = new TextBox();
            tableLayoutPanel3 = new TableLayoutPanel();
            label4 = new Label();
            cboOriginal = new ComboBox();
            tableLayoutPanel5.SuspendLayout();
            tableLayoutPanel3.SuspendLayout();
            SuspendLayout();
            // 
            // lblName
            // 
            lblName.BackColor = Color.Blue;
            lblName.Dock = DockStyle.Top;
            lblName.Font = new Font("Segoe UI", 20F, FontStyle.Bold);
            lblName.ForeColor = Color.White;
            lblName.Location = new Point(0, 0);
            lblName.Name = "lblName";
            lblName.Size = new Size(514, 39);
            lblName.TabIndex = 1;
            lblName.Text = "Thêm vật tư";
            lblName.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // tableLayoutPanel5
            // 
            tableLayoutPanel5.ColumnCount = 3;
            tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 18.1347141F));
            tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 81.86529F));
            tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 113F));
            tableLayoutPanel5.Controls.Add(btnGenerate, 2, 0);
            tableLayoutPanel5.Controls.Add(label6, 0, 0);
            tableLayoutPanel5.Controls.Add(txtCode, 1, 0);
            tableLayoutPanel5.Location = new Point(3, 72);
            tableLayoutPanel5.Margin = new Padding(3, 2, 3, 2);
            tableLayoutPanel5.Name = "tableLayoutPanel5";
            tableLayoutPanel5.RowCount = 1;
            tableLayoutPanel5.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel5.Size = new Size(507, 29);
            tableLayoutPanel5.TabIndex = 25;
            // 
            // btnGenerate
            // 
            btnGenerate.BackColor = Color.Navy;
            btnGenerate.Dock = DockStyle.Fill;
            btnGenerate.FlatStyle = FlatStyle.Flat;
            btnGenerate.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnGenerate.ForeColor = Color.White;
            btnGenerate.Location = new Point(396, 2);
            btnGenerate.Margin = new Padding(3, 2, 3, 2);
            btnGenerate.Name = "btnGenerate";
            btnGenerate.Size = new Size(108, 25);
            btnGenerate.TabIndex = 4;
            btnGenerate.Text = "➕ THÊM";
            btnGenerate.UseVisualStyleBackColor = false;
            btnGenerate.Click += btnGenerate_Click;
            // 
            // label6
            // 
            label6.Dock = DockStyle.Fill;
            label6.Location = new Point(3, 0);
            label6.Name = "label6";
            label6.Size = new Size(65, 29);
            label6.TabIndex = 2;
            label6.Text = "Tên vật tư:";
            label6.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // txtCode
            // 
            txtCode.Dock = DockStyle.Fill;
            txtCode.Location = new Point(74, 4);
            txtCode.Margin = new Padding(3, 4, 3, 2);
            txtCode.Name = "txtCode";
            txtCode.Size = new Size(316, 23);
            txtCode.TabIndex = 30;
            // 
            // tableLayoutPanel3
            // 
            tableLayoutPanel3.ColumnCount = 2;
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 13.81323F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 86.18677F));
            tableLayoutPanel3.Controls.Add(label4, 0, 0);
            tableLayoutPanel3.Controls.Add(cboOriginal, 1, 0);
            tableLayoutPanel3.Location = new Point(3, 38);
            tableLayoutPanel3.Margin = new Padding(3, 2, 3, 2);
            tableLayoutPanel3.Name = "tableLayoutPanel3";
            tableLayoutPanel3.RowCount = 1;
            tableLayoutPanel3.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel3.Size = new Size(507, 29);
            tableLayoutPanel3.TabIndex = 26;
            // 
            // label4
            // 
            label4.Dock = DockStyle.Fill;
            label4.Location = new Point(3, 0);
            label4.Name = "label4";
            label4.Size = new Size(64, 29);
            label4.TabIndex = 2;
            label4.Text = "Original:";
            label4.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // cboOriginal
            // 
            cboOriginal.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cboOriginal.AutoCompleteSource = AutoCompleteSource.ListItems;
            cboOriginal.Dock = DockStyle.Fill;
            cboOriginal.FormattingEnabled = true;
            cboOriginal.Location = new Point(73, 6);
            cboOriginal.Margin = new Padding(3, 6, 3, 2);
            cboOriginal.Name = "cboOriginal";
            cboOriginal.Size = new Size(431, 23);
            cboOriginal.TabIndex = 28;
            // 
            // frmAddMaterialDetail
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(514, 105);
            Controls.Add(tableLayoutPanel3);
            Controls.Add(tableLayoutPanel5);
            Controls.Add(lblName);
            Name = "frmAddMaterialDetail";
            StartPosition = FormStartPosition.CenterParent;
            Text = "Thêm vật tư";
            tableLayoutPanel5.ResumeLayout(false);
            tableLayoutPanel5.PerformLayout();
            tableLayoutPanel3.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private Label lblName;
        private TableLayoutPanel tableLayoutPanel5;
        private Button btnGenerate;
        private Label label6;
        private TextBox txtCode;
        private TableLayoutPanel tableLayoutPanel3;
        private Label label4;
        private ComboBox cboOriginal;
    }
}