namespace MPR_Managerment.Forms.ExportGUI
{
    partial class frmPreviewExportWarehouse
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null)) components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            txtToProject = new TextBox();
            txtCreateBy = new TextBox();
            dgvDetails = new DataGridView();
            btnSave = new Button();
            tableLayoutPanel2 = new TableLayoutPanel();
            tableLayoutPanel3 = new TableLayoutPanel();
            flowLayoutPanel3 = new FlowLayoutPanel();
            btnCancel = new Button();
            flowLayoutPanel1 = new FlowLayoutPanel();
            tableLayoutPanel1 = new TableLayoutPanel();
            panel1 = new Panel();
            txtExportNo = new TextBox();
            label1 = new Label();
            panel2 = new Panel();
            txtFromProject = new TextBox();
            label2 = new Label();
            panel3 = new Panel();
            label3 = new Label();
            panel4 = new Panel();
            label4 = new Label();
            panel5 = new Panel();
            panel6 = new Panel();
            flowLayoutPanel2 = new FlowLayoutPanel();
            btnDelete = new Button();
            btnAddRow = new Button();
            ((System.ComponentModel.ISupportInitialize)dgvDetails).BeginInit();
            tableLayoutPanel2.SuspendLayout();
            tableLayoutPanel3.SuspendLayout();
            flowLayoutPanel3.SuspendLayout();
            flowLayoutPanel1.SuspendLayout();
            tableLayoutPanel1.SuspendLayout();
            panel1.SuspendLayout();
            panel2.SuspendLayout();
            panel3.SuspendLayout();
            panel4.SuspendLayout();
            flowLayoutPanel2.SuspendLayout();
            SuspendLayout();
            // 
            // txtToProject
            // 
            txtToProject.Dock = DockStyle.Right;
            txtToProject.Location = new Point(117, 0);
            txtToProject.Name = "txtToProject";
            txtToProject.Size = new Size(236, 23);
            txtToProject.TabIndex = 5;
            // 
            // txtCreateBy
            // 
            txtCreateBy.Dock = DockStyle.Right;
            txtCreateBy.Location = new Point(119, 0);
            txtCreateBy.Name = "txtCreateBy";
            txtCreateBy.Size = new Size(219, 23);
            txtCreateBy.TabIndex = 7;
            // 
            // dgvDetails
            // 
            dgvDetails.Dock = DockStyle.Fill;
            dgvDetails.Location = new Point(3, 126);
            dgvDetails.Name = "dgvDetails";
            dgvDetails.Size = new Size(947, 442);
            dgvDetails.TabIndex = 0;
            dgvDetails.CellContentDoubleClick += dgvDetails_CellContentDoubleClick;
            // 
            // btnSave
            // 
            btnSave.Location = new Point(3, 3);
            btnSave.Name = "btnSave";
            btnSave.Size = new Size(150, 36);
            btnSave.TabIndex = 2;
            btnSave.Text = "Lưu";
            btnSave.Click += btnSave_Click;
            // 
            // tableLayoutPanel2
            // 
            tableLayoutPanel2.ColumnCount = 1;
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel2.Controls.Add(tableLayoutPanel3, 0, 2);
            tableLayoutPanel2.Controls.Add(dgvDetails, 0, 1);
            tableLayoutPanel2.Controls.Add(flowLayoutPanel1, 0, 0);
            tableLayoutPanel2.Dock = DockStyle.Fill;
            tableLayoutPanel2.Location = new Point(0, 0);
            tableLayoutPanel2.Name = "tableLayoutPanel2";
            tableLayoutPanel2.RowCount = 3;
            tableLayoutPanel2.RowStyles.Add(new RowStyle(SizeType.Percent, 21.5411568F));
            tableLayoutPanel2.RowStyles.Add(new RowStyle(SizeType.Percent, 78.45885F));
            tableLayoutPanel2.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));
            tableLayoutPanel2.Size = new Size(953, 622);
            tableLayoutPanel2.TabIndex = 4;
            // 
            // tableLayoutPanel3
            // 
            tableLayoutPanel3.ColumnCount = 2;
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tableLayoutPanel3.Controls.Add(flowLayoutPanel3, 0, 0);
            tableLayoutPanel3.Dock = DockStyle.Fill;
            tableLayoutPanel3.Location = new Point(3, 574);
            tableLayoutPanel3.Name = "tableLayoutPanel3";
            tableLayoutPanel3.RowCount = 1;
            tableLayoutPanel3.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel3.Size = new Size(947, 45);
            tableLayoutPanel3.TabIndex = 0;
            // 
            // flowLayoutPanel3
            // 
            flowLayoutPanel3.Controls.Add(btnSave);
            flowLayoutPanel3.Controls.Add(btnCancel);
            flowLayoutPanel3.Dock = DockStyle.Fill;
            flowLayoutPanel3.Location = new Point(3, 3);
            flowLayoutPanel3.Name = "flowLayoutPanel3";
            flowLayoutPanel3.Size = new Size(467, 39);
            flowLayoutPanel3.TabIndex = 5;
            // 
            // btnCancel
            // 
            btnCancel.Location = new Point(159, 3);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new Size(89, 36);
            btnCancel.TabIndex = 2;
            btnCancel.Text = "Thoát";
            btnCancel.Click += btnSave_Click;
            // 
            // flowLayoutPanel1
            // 
            flowLayoutPanel1.Controls.Add(tableLayoutPanel1);
            flowLayoutPanel1.Controls.Add(flowLayoutPanel2);
            flowLayoutPanel1.Dock = DockStyle.Left;
            flowLayoutPanel1.Location = new Point(3, 3);
            flowLayoutPanel1.Name = "flowLayoutPanel1";
            flowLayoutPanel1.Size = new Size(706, 117);
            flowLayoutPanel1.TabIndex = 1;
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.ColumnCount = 2;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 51.0668564F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 48.9331436F));
            tableLayoutPanel1.Controls.Add(panel1, 0, 0);
            tableLayoutPanel1.Controls.Add(panel2, 1, 0);
            tableLayoutPanel1.Controls.Add(panel3, 0, 1);
            tableLayoutPanel1.Controls.Add(panel4, 1, 1);
            tableLayoutPanel1.Controls.Add(panel5, 0, 2);
            tableLayoutPanel1.Controls.Add(panel6, 1, 2);
            tableLayoutPanel1.Location = new Point(3, 3);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 3;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.Size = new Size(703, 72);
            tableLayoutPanel1.TabIndex = 12;
            // 
            // panel1
            // 
            panel1.Controls.Add(txtExportNo);
            panel1.Controls.Add(label1);
            panel1.Dock = DockStyle.Fill;
            panel1.Location = new Point(3, 3);
            panel1.Name = "panel1";
            panel1.Size = new Size(353, 29);
            panel1.TabIndex = 0;
            // 
            // txtExportNo
            // 
            txtExportNo.Dock = DockStyle.Right;
            txtExportNo.Location = new Point(117, 0);
            txtExportNo.Name = "txtExportNo";
            txtExportNo.Size = new Size(236, 23);
            txtExportNo.TabIndex = 1;
            // 
            // label1
            // 
            label1.Dock = DockStyle.Left;
            label1.Location = new Point(0, 0);
            label1.Name = "label1";
            label1.Size = new Size(109, 29);
            label1.TabIndex = 8;
            label1.Text = "Kho xuất HH/NVL:";
            label1.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // panel2
            // 
            panel2.Controls.Add(txtFromProject);
            panel2.Controls.Add(label2);
            panel2.Dock = DockStyle.Fill;
            panel2.Location = new Point(362, 3);
            panel2.Name = "panel2";
            panel2.Size = new Size(338, 29);
            panel2.TabIndex = 0;
            // 
            // txtFromProject
            // 
            txtFromProject.Dock = DockStyle.Right;
            txtFromProject.Location = new Point(119, 0);
            txtFromProject.Name = "txtFromProject";
            txtFromProject.Size = new Size(219, 23);
            txtFromProject.TabIndex = 3;
            // 
            // label2
            // 
            label2.Dock = DockStyle.Left;
            label2.Location = new Point(0, 0);
            label2.Name = "label2";
            label2.Size = new Size(109, 29);
            label2.TabIndex = 9;
            label2.Text = "Kho nhập HH/NVL:";
            label2.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // panel3
            // 
            panel3.Controls.Add(txtToProject);
            panel3.Controls.Add(label3);
            panel3.Dock = DockStyle.Fill;
            panel3.Location = new Point(3, 38);
            panel3.Name = "panel3";
            panel3.Size = new Size(353, 29);
            panel3.TabIndex = 0;
            // 
            // label3
            // 
            label3.Dock = DockStyle.Left;
            label3.Location = new Point(0, 0);
            label3.Name = "label3";
            label3.Size = new Size(109, 29);
            label3.TabIndex = 10;
            label3.Text = "Người phụ trách:";
            label3.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // panel4
            // 
            panel4.Controls.Add(txtCreateBy);
            panel4.Controls.Add(label4);
            panel4.Dock = DockStyle.Fill;
            panel4.Location = new Point(362, 38);
            panel4.Name = "panel4";
            panel4.Size = new Size(338, 29);
            panel4.TabIndex = 0;
            // 
            // label4
            // 
            label4.Dock = DockStyle.Left;
            label4.Location = new Point(0, 0);
            label4.Name = "label4";
            label4.Size = new Size(100, 29);
            label4.TabIndex = 11;
            label4.Text = "Tình trạng:";
            label4.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // panel5
            // 
            panel5.Dock = DockStyle.Fill;
            panel5.Location = new Point(3, 73);
            panel5.Name = "panel5";
            panel5.Size = new Size(353, 1);
            panel5.TabIndex = 0;
            // 
            // panel6
            // 
            panel6.Dock = DockStyle.Fill;
            panel6.Location = new Point(362, 73);
            panel6.Name = "panel6";
            panel6.Size = new Size(338, 1);
            panel6.TabIndex = 0;
            // 
            // flowLayoutPanel2
            // 
            flowLayoutPanel2.Controls.Add(btnDelete);
            flowLayoutPanel2.Controls.Add(btnAddRow);
            flowLayoutPanel2.Location = new Point(3, 81);
            flowLayoutPanel2.Name = "flowLayoutPanel2";
            flowLayoutPanel2.Size = new Size(703, 36);
            flowLayoutPanel2.TabIndex = 13;
            // 
            // btnDelete
            // 
            btnDelete.Enabled = false;
            btnDelete.Location = new Point(3, 3);
            btnDelete.Name = "btnDelete";
            btnDelete.Size = new Size(150, 30);
            btnDelete.TabIndex = 4;
            btnDelete.Text = "Delete";
            btnDelete.Click += btnDelete_Click;
            // 
            // btnAddRow
            // 
            btnAddRow.Location = new Point(159, 3);
            btnAddRow.Name = "btnAddRow";
            btnAddRow.Size = new Size(150, 30);
            btnAddRow.TabIndex = 4;
            btnAddRow.Text = "Thêm dòng";
            btnAddRow.Click += btnDelete_Click;
            // 
            // frmPreviewExportWarehouse
            // 
            ClientSize = new Size(953, 622);
            Controls.Add(tableLayoutPanel2);
            Name = "frmPreviewExportWarehouse";
            Text = "Preview Export Warehouse";
            Load += frmPreviewExportWarehouse_Load;
            ((System.ComponentModel.ISupportInitialize)dgvDetails).EndInit();
            tableLayoutPanel2.ResumeLayout(false);
            tableLayoutPanel3.ResumeLayout(false);
            flowLayoutPanel3.ResumeLayout(false);
            flowLayoutPanel1.ResumeLayout(false);
            tableLayoutPanel1.ResumeLayout(false);
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            panel2.ResumeLayout(false);
            panel2.PerformLayout();
            panel3.ResumeLayout(false);
            panel3.PerformLayout();
            panel4.ResumeLayout(false);
            panel4.PerformLayout();
            flowLayoutPanel2.ResumeLayout(false);
            ResumeLayout(false);
        }

        private System.Windows.Forms.TextBox txtToProject;
        private System.Windows.Forms.TextBox txtCreateBy;
        private System.Windows.Forms.DataGridView dgvDetails;
        private System.Windows.Forms.Button btnSave;
        private TableLayoutPanel tableLayoutPanel2;
        private TableLayoutPanel tableLayoutPanel3;
        private FlowLayoutPanel flowLayoutPanel1;
        private Label label4;
        private TableLayoutPanel tableLayoutPanel1;
        private Panel panel1;
        private Label label1;
        private Panel panel2;
        private Label label2;
        private Panel panel3;
        private Label label3;
        private Panel panel4;
        private Panel panel5;
        private Panel panel6;
        private FlowLayoutPanel flowLayoutPanel2;
        private TextBox txtExportNo;
        private TextBox txtFromProject;
        private Button btnDelete;
        private FlowLayoutPanel flowLayoutPanel3;
        private Button btnCancel;
        private Button btnAddRow;
    }
}