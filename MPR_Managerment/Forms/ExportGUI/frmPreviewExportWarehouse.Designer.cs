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
            tableLayoutPanel1 = new TableLayoutPanel();
            txtExportNo = new TextBox();
            txtFromProject = new TextBox();
            txtToProject = new TextBox();
            txtCreateBy = new TextBox();
            dgvDetails = new DataGridView();
            btnSave = new Button();
            btnDelete = new Button();
            tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvDetails).BeginInit();
            SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.ColumnCount = 2;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70F));
            tableLayoutPanel1.Controls.Add(txtExportNo, 1, 0);
            tableLayoutPanel1.Controls.Add(txtFromProject, 1, 1);
            tableLayoutPanel1.Controls.Add(txtToProject, 1, 2);
            tableLayoutPanel1.Controls.Add(txtCreateBy, 1, 3);
            tableLayoutPanel1.Dock = DockStyle.Top;
            tableLayoutPanel1.Location = new Point(0, 0);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 20F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 20F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 20F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 20F));
            tableLayoutPanel1.Size = new Size(772, 150);
            tableLayoutPanel1.TabIndex = 1;
            // 
            // txtExportNo
            // 
            txtExportNo.Location = new Point(234, 3);
            txtExportNo.Name = "txtExportNo";
            txtExportNo.Size = new Size(100, 23);
            txtExportNo.TabIndex = 1;
            // 
            // txtFromProject
            // 
            txtFromProject.Location = new Point(234, 23);
            txtFromProject.Name = "txtFromProject";
            txtFromProject.Size = new Size(100, 23);
            txtFromProject.TabIndex = 3;
            // 
            // txtToProject
            // 
            txtToProject.Location = new Point(234, 43);
            txtToProject.Name = "txtToProject";
            txtToProject.Size = new Size(100, 23);
            txtToProject.TabIndex = 5;
            // 
            // txtCreateBy
            // 
            txtCreateBy.Location = new Point(234, 63);
            txtCreateBy.Name = "txtCreateBy";
            txtCreateBy.Size = new Size(100, 23);
            txtCreateBy.TabIndex = 7;
            // 
            // dgvDetails
            // 
            dgvDetails.Dock = DockStyle.Fill;
            dgvDetails.Location = new Point(0, 150);
            dgvDetails.Name = "dgvDetails";
            dgvDetails.Size = new Size(772, 63);
            dgvDetails.TabIndex = 0;
            // 
            // btnSave
            // 
            btnSave.Dock = DockStyle.Bottom;
            btnSave.Location = new Point(0, 213);
            btnSave.Name = "btnSave";
            btnSave.Size = new Size(772, 23);
            btnSave.TabIndex = 2;
            btnSave.Text = "Save all";
            btnSave.Click += btnSave_Click;
            // 
            // btnDelete
            // 
            btnDelete.Dock = DockStyle.Bottom;
            btnDelete.Location = new Point(0, 236);
            btnDelete.Name = "btnDelete";
            btnDelete.Size = new Size(772, 23);
            btnDelete.TabIndex = 3;
            btnDelete.Text = "Delete";
            btnDelete.Visible = false;
            btnDelete.Click += btnDelete_Click;
            // 
            // frmPreviewExportWarehouse
            // 
            ClientSize = new Size(772, 259);
            Controls.Add(dgvDetails);
            Controls.Add(tableLayoutPanel1);
            Controls.Add(btnSave);
            Controls.Add(btnDelete);
            Name = "frmPreviewExportWarehouse";
            Text = "Preview Export Warehouse";
            tableLayoutPanel1.ResumeLayout(false);
            tableLayoutPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)dgvDetails).EndInit();
            ResumeLayout(false);
        }

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.TextBox txtExportNo;
        private System.Windows.Forms.TextBox txtFromProject;
        private System.Windows.Forms.TextBox txtToProject;
        private System.Windows.Forms.TextBox txtCreateBy;
        private System.Windows.Forms.DataGridView dgvDetails;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnDelete;
    }
}