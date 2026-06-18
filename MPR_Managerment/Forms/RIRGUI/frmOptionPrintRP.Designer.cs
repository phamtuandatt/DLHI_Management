namespace MPR_Managerment.Forms.RIRGUI
{
    partial class frmOptionPrintRP
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
            btnProject = new Button();
            btnRIRNo = new Button();
            tableLayoutPanel1 = new TableLayoutPanel();
            tableLayoutPanel1.SuspendLayout();
            SuspendLayout();
            //
            // btnProject
            //
            btnProject.BackColor = Color.FromArgb(0, 120, 212);
            btnProject.Dock = DockStyle.Fill;
            btnProject.FlatStyle = FlatStyle.Flat;
            btnProject.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnProject.ForeColor = Color.White;
            btnProject.Location = new Point(3, 2);
            btnProject.Margin = new Padding(3, 2, 3, 2);
            btnProject.Name = "btnProject";
            btnProject.Size = new Size(195, 45);
            btnProject.TabIndex = 9;
            btnProject.Text = "🎨 Project";
            btnProject.UseVisualStyleBackColor = false;
            btnProject.Click += btnProject_Click;
            //
            // btnRIRNo
            //
            btnRIRNo.BackColor = Color.ForestGreen;
            btnRIRNo.Dock = DockStyle.Fill;
            btnRIRNo.FlatStyle = FlatStyle.Flat;
            btnRIRNo.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnRIRNo.ForeColor = Color.White;
            btnRIRNo.Location = new Point(204, 2);
            btnRIRNo.Margin = new Padding(3, 2, 3, 2);
            btnRIRNo.Name = "btnRIRNo";
            btnRIRNo.Size = new Size(195, 45);
            btnRIRNo.TabIndex = 8;
            btnRIRNo.Text = "RIR No";
            btnRIRNo.UseVisualStyleBackColor = false;
            btnRIRNo.Click += btnRIRNo_Click;
            //
            // tableLayoutPanel1
            //
            tableLayoutPanel1.ColumnCount = 2;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tableLayoutPanel1.Controls.Add(btnRIRNo, 1, 0);
            tableLayoutPanel1.Controls.Add(btnProject, 0, 0);
            tableLayoutPanel1.Dock = DockStyle.Fill;
            tableLayoutPanel1.Location = new Point(0, 0);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 1;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel1.Size = new Size(402, 49);
            tableLayoutPanel1.TabIndex = 1;
            //
            // frmOptionPrintRP
            //
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(402, 49);
            Controls.Add(tableLayoutPanel1);
            try
            {
                string iconPath = System.IO.Path.Combine(
                    System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)!,
                    "Resources", "icon.ico");
                if (System.IO.File.Exists(iconPath))
                    Icon = new Icon(iconPath);
            }
            catch { }
            Text = "Option";
            tableLayoutPanel1.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private Button btnProject;
        private Button btnRIRNo;
        private TableLayoutPanel tableLayoutPanel1;
    }
}
