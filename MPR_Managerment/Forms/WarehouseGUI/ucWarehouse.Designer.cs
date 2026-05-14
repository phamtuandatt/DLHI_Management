namespace MPR_Managerment.Forms.WarehouseGUI
{
    partial class ucWarehouse
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
            pnContent = new TableLayoutPanel();
            tableLayoutPanel2 = new TableLayoutPanel();
            tableLayoutPanel3 = new TableLayoutPanel();
            pnHeader = new Panel();
            pnAction = new Panel();
            pnHisTransfer = new Panel();
            panel1.SuspendLayout();
            pnContent.SuspendLayout();
            tableLayoutPanel2.SuspendLayout();
            tableLayoutPanel3.SuspendLayout();
            SuspendLayout();
            // 
            // panel1
            // 
            panel1.BackColor = Color.White;
            panel1.Controls.Add(pnContent);
            panel1.Dock = DockStyle.Fill;
            panel1.Location = new Point(0, 0);
            panel1.Name = "panel1";
            panel1.Size = new Size(1353, 714);
            panel1.TabIndex = 0;
            // 
            // pnContent
            // 
            pnContent.ColumnCount = 1;
            pnContent.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            pnContent.Controls.Add(tableLayoutPanel2, 0, 0);
            pnContent.Dock = DockStyle.Fill;
            pnContent.Location = new Point(0, 0);
            pnContent.Name = "pnContent";
            pnContent.RowCount = 3;
            pnContent.RowStyles.Add(new RowStyle(SizeType.Absolute, 176F));
            pnContent.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            pnContent.RowStyles.Add(new RowStyle(SizeType.Absolute, 20F));
            pnContent.RowStyles.Add(new RowStyle(SizeType.Absolute, 20F));
            pnContent.Size = new Size(1353, 714);
            pnContent.TabIndex = 0;
            // 
            // tableLayoutPanel2
            // 
            tableLayoutPanel2.ColumnCount = 2;
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 780F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel2.Controls.Add(tableLayoutPanel3, 0, 0);
            tableLayoutPanel2.Controls.Add(pnHisTransfer, 1, 0);
            tableLayoutPanel2.Dock = DockStyle.Fill;
            tableLayoutPanel2.Location = new Point(3, 3);
            tableLayoutPanel2.Name = "tableLayoutPanel2";
            tableLayoutPanel2.RowCount = 1;
            tableLayoutPanel2.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel2.Size = new Size(1347, 170);
            tableLayoutPanel2.TabIndex = 0;
            // 
            // tableLayoutPanel3
            // 
            tableLayoutPanel3.ColumnCount = 1;
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tableLayoutPanel3.Controls.Add(pnHeader, 0, 0);
            tableLayoutPanel3.Controls.Add(pnAction, 0, 1);
            tableLayoutPanel3.Dock = DockStyle.Fill;
            tableLayoutPanel3.Location = new Point(3, 3);
            tableLayoutPanel3.Name = "tableLayoutPanel3";
            tableLayoutPanel3.RowCount = 2;
            tableLayoutPanel3.RowStyles.Add(new RowStyle(SizeType.Percent, 43.2926826F));
            tableLayoutPanel3.RowStyles.Add(new RowStyle(SizeType.Percent, 56.7073174F));
            tableLayoutPanel3.Size = new Size(774, 164);
            tableLayoutPanel3.TabIndex = 0;
            // 
            // pnHeader
            // 
            pnHeader.Dock = DockStyle.Fill;
            pnHeader.Location = new Point(3, 3);
            pnHeader.Name = "pnHeader";
            pnHeader.Size = new Size(768, 65);
            pnHeader.TabIndex = 0;
            // 
            // pnAction
            // 
            pnAction.Dock = DockStyle.Fill;
            pnAction.Location = new Point(3, 74);
            pnAction.Name = "pnAction";
            pnAction.Size = new Size(768, 87);
            pnAction.TabIndex = 1;
            // 
            // pnHisTransfer
            // 
            pnHisTransfer.Dock = DockStyle.Fill;
            pnHisTransfer.Location = new Point(783, 3);
            pnHisTransfer.Name = "pnHisTransfer";
            pnHisTransfer.Size = new Size(561, 164);
            pnHisTransfer.TabIndex = 1;
            // 
            // ucWarehouse
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            Controls.Add(panel1);
            Name = "ucWarehouse";
            Size = new Size(1353, 714);
            Load += ucWarehouse_Load;
            panel1.ResumeLayout(false);
            pnContent.ResumeLayout(false);
            tableLayoutPanel2.ResumeLayout(false);
            tableLayoutPanel3.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private Panel panel1;
        private TableLayoutPanel pnContent;
        private TableLayoutPanel tableLayoutPanel2;
        private TableLayoutPanel tableLayoutPanel3;
        private Panel pnHeader;
        private Panel pnAction;
        private Panel pnHisTransfer;
    }
}
