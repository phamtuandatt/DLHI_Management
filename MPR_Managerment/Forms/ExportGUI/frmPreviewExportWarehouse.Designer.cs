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
            txtFromProject = new TextBox();
            dgvDetails = new DataGridView();
            btnSave = new Button();
            tableLayoutPanel2 = new TableLayoutPanel();
            tableLayoutPanel3 = new TableLayoutPanel();
            flowLayoutPanel3 = new FlowLayoutPanel();
            btnCancel = new Button();
            btnUpdateServer = new Button();
            flowLayoutPanel1 = new FlowLayoutPanel();
            tableLayoutPanel1 = new TableLayoutPanel();
            panel1 = new Panel();
            txtExportNo = new TextBox();
            label1 = new Label();
            panel3 = new Panel();
            btnFromWarehouse = new Button();
            label3 = new Label();
            panel5 = new Panel();
            txtCreatedBy = new TextBox();
            label5 = new Label();
            panel4 = new Panel();
            cboStatus = new ComboBox();
            label4 = new Label();
            panel2 = new Panel();
            btnToWarehosue = new Button();
            txtToProject = new TextBox();
            label2 = new Label();
            panel6 = new Panel();
            dtCreate = new DateTimePicker();
            label6 = new Label();
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
            panel3.SuspendLayout();
            panel5.SuspendLayout();
            panel4.SuspendLayout();
            panel2.SuspendLayout();
            panel6.SuspendLayout();
            flowLayoutPanel2.SuspendLayout();
            SuspendLayout();
            // 
            // txtFromProject
            // 
            txtFromProject.Location = new Point(117, 0);
            txtFromProject.Name = "txtFromProject";
            txtFromProject.Size = new Size(197, 23);
            txtFromProject.TabIndex = 5;
            txtFromProject.TextChanged += txtFromProject_TextChanged;
            // 
            // dgvDetails
            // 
            dgvDetails.Dock = DockStyle.Fill;
            dgvDetails.Location = new Point(3, 152);
            dgvDetails.Name = "dgvDetails";
            dgvDetails.Size = new Size(947, 416);
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
            tableLayoutPanel2.RowStyles.Add(new RowStyle(SizeType.Percent, 26.09457F));
            tableLayoutPanel2.RowStyles.Add(new RowStyle(SizeType.Percent, 73.905426F));
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
            tableLayoutPanel3.Controls.Add(btnUpdateServer, 1, 0);
            tableLayoutPanel3.Dock = DockStyle.Fill;
            tableLayoutPanel3.Location = new Point(3, 574);
            tableLayoutPanel3.Name = "tableLayoutPanel3";
            tableLayoutPanel3.RowCount = 1;
            tableLayoutPanel3.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel3.Size = new Size(947, 45);
            tableLayoutPanel3.TabIndex = 0;
            tableLayoutPanel3.Paint += tableLayoutPanel3_Paint;
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
            btnCancel.Click += btnCancel_Click;
            // 
            // btnUpdateServer
            // 
            btnUpdateServer.Dock = DockStyle.Right;
            btnUpdateServer.Location = new Point(793, 3);
            btnUpdateServer.Name = "btnUpdateServer";
            btnUpdateServer.Size = new Size(151, 39);
            btnUpdateServer.TabIndex = 2;
            btnUpdateServer.Text = "Cập nhật server";
            btnUpdateServer.Click += btnUpdateServer_Click;
            // 
            // flowLayoutPanel1
            // 
            flowLayoutPanel1.Controls.Add(tableLayoutPanel1);
            flowLayoutPanel1.Controls.Add(flowLayoutPanel2);
            flowLayoutPanel1.Dock = DockStyle.Left;
            flowLayoutPanel1.Location = new Point(3, 3);
            flowLayoutPanel1.Name = "flowLayoutPanel1";
            flowLayoutPanel1.Size = new Size(706, 143);
            flowLayoutPanel1.TabIndex = 1;
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.ColumnCount = 2;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 51.06686F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 48.9331436F));
            tableLayoutPanel1.Controls.Add(panel1, 0, 0);
            tableLayoutPanel1.Controls.Add(panel3, 0, 1);
            tableLayoutPanel1.Controls.Add(panel5, 0, 2);
            tableLayoutPanel1.Controls.Add(panel4, 1, 2);
            tableLayoutPanel1.Controls.Add(panel2, 1, 1);
            tableLayoutPanel1.Controls.Add(panel6, 1, 0);
            tableLayoutPanel1.Location = new Point(3, 3);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 3;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
            tableLayoutPanel1.Size = new Size(703, 100);
            tableLayoutPanel1.TabIndex = 12;
            tableLayoutPanel1.Paint += tableLayoutPanel1_Paint;
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
            label1.Text = "Số phiếu xuất kho";
            label1.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // panel3
            // 
            panel3.Controls.Add(btnFromWarehouse);
            panel3.Controls.Add(txtFromProject);
            panel3.Controls.Add(label3);
            panel3.Location = new Point(3, 38);
            panel3.Name = "panel3";
            panel3.Size = new Size(353, 29);
            panel3.TabIndex = 0;
            // 
            // btnFromWarehouse
            // 
            btnFromWarehouse.Location = new Point(313, 0);
            btnFromWarehouse.Name = "btnFromWarehouse";
            btnFromWarehouse.Size = new Size(40, 23);
            btnFromWarehouse.TabIndex = 1;
            btnFromWarehouse.Text = "🔎";
            btnFromWarehouse.UseVisualStyleBackColor = true;
            btnFromWarehouse.Click += btnFromWarehouse_Click;
            // 
            // label3
            // 
            label3.Dock = DockStyle.Left;
            label3.Location = new Point(0, 0);
            label3.Name = "label3";
            label3.Size = new Size(109, 29);
            label3.TabIndex = 10;
            label3.Text = "Kho xuất HH/NVL:";
            label3.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // panel5
            // 
            panel5.Controls.Add(txtCreatedBy);
            panel5.Controls.Add(label5);
            panel5.Location = new Point(3, 73);
            panel5.Name = "panel5";
            panel5.Size = new Size(353, 29);
            panel5.TabIndex = 0;
            // 
            // txtCreatedBy
            // 
            txtCreatedBy.Dock = DockStyle.Right;
            txtCreatedBy.Location = new Point(117, 0);
            txtCreatedBy.Name = "txtCreatedBy";
            txtCreatedBy.Size = new Size(236, 23);
            txtCreatedBy.TabIndex = 5;
            // 
            // label5
            // 
            label5.Dock = DockStyle.Left;
            label5.Location = new Point(0, 0);
            label5.Name = "label5";
            label5.Size = new Size(109, 29);
            label5.TabIndex = 10;
            label5.Text = "Người phụ trách:";
            label5.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // panel4
            // 
            panel4.Controls.Add(cboStatus);
            panel4.Controls.Add(label4);
            panel4.Dock = DockStyle.Fill;
            panel4.Location = new Point(362, 73);
            panel4.Name = "panel4";
            panel4.Size = new Size(338, 29);
            panel4.TabIndex = 0;
            // 
            // cboStatus
            // 
            cboStatus.Dock = DockStyle.Right;
            cboStatus.DropDownStyle = ComboBoxStyle.DropDownList;
            cboStatus.FormattingEnabled = true;
            cboStatus.Items.AddRange(new object[] { "Chưa xác nhận", "Xác nhận", "Hủy bỏ" });
            cboStatus.Location = new Point(119, 0);
            cboStatus.Name = "cboStatus";
            cboStatus.Size = new Size(219, 23);
            cboStatus.TabIndex = 12;
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
            // panel2
            // 
            panel2.Controls.Add(btnToWarehosue);
            panel2.Controls.Add(txtToProject);
            panel2.Controls.Add(label2);
            panel2.Location = new Point(362, 38);
            panel2.Name = "panel2";
            panel2.Size = new Size(338, 29);
            panel2.TabIndex = 0;
            // 
            // btnToWarehosue
            // 
            btnToWarehosue.Location = new Point(296, 0);
            btnToWarehosue.Name = "btnToWarehosue";
            btnToWarehosue.Size = new Size(40, 23);
            btnToWarehosue.TabIndex = 1;
            btnToWarehosue.Text = "🔎";
            btnToWarehosue.UseVisualStyleBackColor = true;
            btnToWarehosue.Click += btnToWarehosue_Click;
            // 
            // txtToProject
            // 
            txtToProject.Location = new Point(119, 0);
            txtToProject.Name = "txtToProject";
            txtToProject.Size = new Size(178, 23);
            txtToProject.TabIndex = 3;
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
            // panel6
            // 
            panel6.Controls.Add(dtCreate);
            panel6.Controls.Add(label6);
            panel6.Location = new Point(362, 3);
            panel6.Name = "panel6";
            panel6.Size = new Size(338, 29);
            panel6.TabIndex = 0;
            // 
            // dtCreate
            // 
            dtCreate.Dock = DockStyle.Right;
            dtCreate.Format = DateTimePickerFormat.Short;
            dtCreate.Location = new Point(119, 0);
            dtCreate.Name = "dtCreate";
            dtCreate.Size = new Size(219, 23);
            dtCreate.TabIndex = 12;
            // 
            // label6
            // 
            label6.Dock = DockStyle.Left;
            label6.Location = new Point(0, 0);
            label6.Name = "label6";
            label6.Size = new Size(100, 29);
            label6.TabIndex = 11;
            label6.Text = "Tình trạng:";
            label6.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // flowLayoutPanel2
            // 
            flowLayoutPanel2.Controls.Add(btnDelete);
            flowLayoutPanel2.Controls.Add(btnAddRow);
            flowLayoutPanel2.Location = new Point(3, 109);
            flowLayoutPanel2.Name = "flowLayoutPanel2";
            flowLayoutPanel2.Size = new Size(703, 34);
            flowLayoutPanel2.TabIndex = 13;
            // 
            // btnDelete
            // 
            btnDelete.BackColor = SystemColors.Control;
            btnDelete.Enabled = false;
            btnDelete.ForeColor = Color.White;
            btnDelete.Location = new Point(3, 3);
            btnDelete.Name = "btnDelete";
            btnDelete.Size = new Size(150, 30);
            btnDelete.TabIndex = 4;
            btnDelete.Text = "Delete";
            btnDelete.UseVisualStyleBackColor = false;
            btnDelete.Click += btnDelete_Click;
            // 
            // btnAddRow
            // 
            btnAddRow.Location = new Point(159, 3);
            btnAddRow.Name = "btnAddRow";
            btnAddRow.Size = new Size(150, 30);
            btnAddRow.TabIndex = 4;
            btnAddRow.Text = "Thêm dòng";
            btnAddRow.Click += btnAddRow_Click;
            // 
            // frmPreviewExportWarehouse
            // 
            ClientSize = new Size(953, 622);
            Controls.Add(tableLayoutPanel2);
            Name = "frmPreviewExportWarehouse";
            StartPosition = FormStartPosition.CenterScreen;
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
            panel3.ResumeLayout(false);
            panel3.PerformLayout();
            panel5.ResumeLayout(false);
            panel5.PerformLayout();
            panel4.ResumeLayout(false);
            panel2.ResumeLayout(false);
            panel2.PerformLayout();
            panel6.ResumeLayout(false);
            flowLayoutPanel2.ResumeLayout(false);
            ResumeLayout(false);
        }

        private System.Windows.Forms.TextBox txtFromProject;
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
        private FlowLayoutPanel flowLayoutPanel2;
        private TextBox txtExportNo;
        private TextBox txtToProject;
        private Button btnDelete;
        private FlowLayoutPanel flowLayoutPanel3;
        private Button btnCancel;
        private Button btnAddRow;
        private Button btnFromWarehouse;
        private Button btnToWarehosue;
        private Panel panel5;
        private TextBox txtCreatedBy;
        private Label label5;
        private Panel panel6;
        private DateTimePicker dtCreate;
        private Label label6;
        private ComboBox cboStatus;
        private Button btnUpdateServer;
    }
}