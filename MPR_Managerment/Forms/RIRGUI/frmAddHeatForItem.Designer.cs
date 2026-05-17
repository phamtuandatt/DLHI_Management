namespace MPR_Managerment.Forms.RIRGUI
{
    partial class frmAddHeatForItem
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
            label1 = new Label();
            txtQty = new TextBox();
            tableLayoutPanel3 = new TableLayoutPanel();
            txtMTR = new TextBox();
            label2 = new Label();
            tableLayoutPanel4 = new TableLayoutPanel();
            txtHeat = new TextBox();
            label3 = new Label();
            tableLayoutPanel5 = new TableLayoutPanel();
            txtIDCode = new TextBox();
            label4 = new Label();
            tableLayoutPanel6 = new TableLayoutPanel();
            btnCancel = new Button();
            btnSave = new Button();
            tableLayoutPanel1.SuspendLayout();
            tableLayoutPanel2.SuspendLayout();
            tableLayoutPanel3.SuspendLayout();
            tableLayoutPanel4.SuspendLayout();
            tableLayoutPanel5.SuspendLayout();
            tableLayoutPanel6.SuspendLayout();
            SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.ColumnCount = 1;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.Controls.Add(tableLayoutPanel2, 0, 0);
            tableLayoutPanel1.Controls.Add(tableLayoutPanel3, 0, 1);
            tableLayoutPanel1.Controls.Add(tableLayoutPanel4, 0, 2);
            tableLayoutPanel1.Controls.Add(tableLayoutPanel5, 0, 3);
            tableLayoutPanel1.Controls.Add(tableLayoutPanel6, 0, 4);
            tableLayoutPanel1.Dock = DockStyle.Fill;
            tableLayoutPanel1.Location = new Point(0, 0);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 6;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.Size = new Size(312, 181);
            tableLayoutPanel1.TabIndex = 0;
            // 
            // tableLayoutPanel2
            // 
            tableLayoutPanel2.ColumnCount = 2;
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel2.Controls.Add(label1, 0, 0);
            tableLayoutPanel2.Controls.Add(txtQty, 1, 0);
            tableLayoutPanel2.Dock = DockStyle.Fill;
            tableLayoutPanel2.Location = new Point(3, 3);
            tableLayoutPanel2.Name = "tableLayoutPanel2";
            tableLayoutPanel2.RowCount = 1;
            tableLayoutPanel2.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel2.Size = new Size(306, 29);
            tableLayoutPanel2.TabIndex = 1;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Dock = DockStyle.Left;
            label1.Location = new Point(3, 0);
            label1.Name = "label1";
            label1.Size = new Size(57, 29);
            label1.TabIndex = 0;
            label1.Text = "Số lượng:";
            label1.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // txtQty
            // 
            txtQty.Dock = DockStyle.Fill;
            txtQty.Location = new Point(73, 3);
            txtQty.Name = "txtQty";
            txtQty.Size = new Size(230, 23);
            txtQty.TabIndex = 1;
            txtQty.KeyPress += txtQty_KeyPress;
            // 
            // tableLayoutPanel3
            // 
            tableLayoutPanel3.ColumnCount = 2;
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel3.Controls.Add(txtMTR, 1, 0);
            tableLayoutPanel3.Controls.Add(label2, 0, 0);
            tableLayoutPanel3.Dock = DockStyle.Fill;
            tableLayoutPanel3.Location = new Point(3, 38);
            tableLayoutPanel3.Name = "tableLayoutPanel3";
            tableLayoutPanel3.RowCount = 1;
            tableLayoutPanel3.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel3.Size = new Size(306, 29);
            tableLayoutPanel3.TabIndex = 2;
            // 
            // txtMTR
            // 
            txtMTR.Dock = DockStyle.Fill;
            txtMTR.Location = new Point(73, 3);
            txtMTR.Name = "txtMTR";
            txtMTR.Size = new Size(230, 23);
            txtMTR.TabIndex = 2;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Dock = DockStyle.Left;
            label2.Location = new Point(3, 0);
            label2.Name = "label2";
            label2.Size = new Size(53, 29);
            label2.TabIndex = 0;
            label2.Text = "MTR No:";
            label2.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // tableLayoutPanel4
            // 
            tableLayoutPanel4.ColumnCount = 2;
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70F));
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel4.Controls.Add(txtHeat, 1, 0);
            tableLayoutPanel4.Controls.Add(label3, 0, 0);
            tableLayoutPanel4.Dock = DockStyle.Fill;
            tableLayoutPanel4.Location = new Point(3, 73);
            tableLayoutPanel4.Name = "tableLayoutPanel4";
            tableLayoutPanel4.RowCount = 1;
            tableLayoutPanel4.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel4.Size = new Size(306, 29);
            tableLayoutPanel4.TabIndex = 3;
            // 
            // txtHeat
            // 
            txtHeat.Dock = DockStyle.Fill;
            txtHeat.Location = new Point(73, 3);
            txtHeat.Name = "txtHeat";
            txtHeat.Size = new Size(230, 23);
            txtHeat.TabIndex = 2;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Dock = DockStyle.Left;
            label3.Location = new Point(3, 0);
            label3.Name = "label3";
            label3.Size = new Size(54, 29);
            label3.TabIndex = 0;
            label3.Text = "Heat No:";
            label3.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // tableLayoutPanel5
            // 
            tableLayoutPanel5.ColumnCount = 2;
            tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70F));
            tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel5.Controls.Add(txtIDCode, 1, 0);
            tableLayoutPanel5.Controls.Add(label4, 0, 0);
            tableLayoutPanel5.Dock = DockStyle.Fill;
            tableLayoutPanel5.Location = new Point(3, 108);
            tableLayoutPanel5.Name = "tableLayoutPanel5";
            tableLayoutPanel5.RowCount = 1;
            tableLayoutPanel5.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel5.Size = new Size(306, 29);
            tableLayoutPanel5.TabIndex = 4;
            // 
            // txtIDCode
            // 
            txtIDCode.Dock = DockStyle.Fill;
            txtIDCode.Location = new Point(73, 3);
            txtIDCode.Name = "txtIDCode";
            txtIDCode.Size = new Size(230, 23);
            txtIDCode.TabIndex = 2;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Dock = DockStyle.Left;
            label4.Location = new Point(3, 0);
            label4.Name = "label4";
            label4.Size = new Size(52, 29);
            label4.TabIndex = 0;
            label4.Text = "ID Code:";
            label4.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // tableLayoutPanel6
            // 
            tableLayoutPanel6.ColumnCount = 2;
            tableLayoutPanel6.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tableLayoutPanel6.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tableLayoutPanel6.Controls.Add(btnCancel, 0, 0);
            tableLayoutPanel6.Controls.Add(btnSave, 1, 0);
            tableLayoutPanel6.Dock = DockStyle.Fill;
            tableLayoutPanel6.Location = new Point(3, 143);
            tableLayoutPanel6.Name = "tableLayoutPanel6";
            tableLayoutPanel6.RowCount = 1;
            tableLayoutPanel6.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel6.Size = new Size(306, 29);
            tableLayoutPanel6.TabIndex = 5;
            // 
            // btnCancel
            // 
            btnCancel.Dock = DockStyle.Fill;
            btnCancel.Location = new Point(3, 3);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new Size(147, 23);
            btnCancel.TabIndex = 0;
            btnCancel.Text = "THOÁT";
            btnCancel.UseVisualStyleBackColor = true;
            btnCancel.Click += btnCancel_Click;
            // 
            // btnSave
            // 
            btnSave.Dock = DockStyle.Fill;
            btnSave.Location = new Point(156, 3);
            btnSave.Name = "btnSave";
            btnSave.Size = new Size(147, 23);
            btnSave.TabIndex = 1;
            btnSave.Text = "XÁC NHẬN";
            btnSave.UseVisualStyleBackColor = true;
            btnSave.Click += btnSave_Click;
            // 
            // frmAddHeatForItem
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(312, 181);
            Controls.Add(tableLayoutPanel1);
            Name = "frmAddHeatForItem";
            Text = "frmAddHeatForItem";
            tableLayoutPanel1.ResumeLayout(false);
            tableLayoutPanel2.ResumeLayout(false);
            tableLayoutPanel2.PerformLayout();
            tableLayoutPanel3.ResumeLayout(false);
            tableLayoutPanel3.PerformLayout();
            tableLayoutPanel4.ResumeLayout(false);
            tableLayoutPanel4.PerformLayout();
            tableLayoutPanel5.ResumeLayout(false);
            tableLayoutPanel5.PerformLayout();
            tableLayoutPanel6.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private TableLayoutPanel tableLayoutPanel1;
        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
        private TableLayoutPanel tableLayoutPanel2;
        private TableLayoutPanel tableLayoutPanel3;
        private TextBox txtQty;
        private TextBox txtMTR;
        private TableLayoutPanel tableLayoutPanel4;
        private TextBox txtHeat;
        private TableLayoutPanel tableLayoutPanel5;
        private TextBox txtIDCode;
        private TableLayoutPanel tableLayoutPanel6;
        private Button btnCancel;
        private Button btnSave;
    }
}