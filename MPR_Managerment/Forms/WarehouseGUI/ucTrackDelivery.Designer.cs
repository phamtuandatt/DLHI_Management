namespace MPR_Managerment.Forms.WarehouseGUI
{
    partial class ucTrackDelivery
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
            tableLayoutPanel1 = new TableLayoutPanel();
            pnFilter = new Panel();
            btnSearch = new Button();
            btnClear = new Button();
            btnSave = new Button();
            cboProjectCode = new ComboBox();
            lblProjectCode = new Label();
            dtpToDate = new DateTimePicker();
            lblToDate = new Label();
            dtpFromDate = new DateTimePicker();
            lblFromDate = new Label();
            cboSupplier = new ComboBox();
            lblSupplier = new Label();
            pnGrid = new Panel();
            dgvTrackDelivery = new DataGridView();
            tableLayoutPanel1.SuspendLayout();
            pnFilter.SuspendLayout();
            pnGrid.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(dgvTrackDelivery)).BeginInit();
            SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.BackColor = Color.White;
            tableLayoutPanel1.ColumnCount = 1;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.Controls.Add(pnFilter, 0, 0);
            tableLayoutPanel1.Controls.Add(pnGrid, 0, 1);
            tableLayoutPanel1.Dock = DockStyle.Fill;
            tableLayoutPanel1.Location = new Point(0, 0);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 2;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 100F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.Size = new Size(950, 550);
            tableLayoutPanel1.TabIndex = 0;
            // 
            // pnFilter
            // 
            pnFilter.Controls.Add(btnSave);
            pnFilter.Controls.Add(btnClear);
            pnFilter.Controls.Add(btnSearch);
            pnFilter.Controls.Add(cboProjectCode);
            pnFilter.Controls.Add(lblProjectCode);
            pnFilter.Controls.Add(dtpToDate);
            pnFilter.Controls.Add(lblToDate);
            pnFilter.Controls.Add(dtpFromDate);
            pnFilter.Controls.Add(lblFromDate);
            pnFilter.Controls.Add(cboSupplier);
            pnFilter.Controls.Add(lblSupplier);
            pnFilter.Dock = DockStyle.Fill;
            pnFilter.Location = new Point(3, 3);
            pnFilter.Name = "pnFilter";
            pnFilter.Size = new Size(944, 54);
            pnFilter.TabIndex = 0;
            // 
            // lblSupplier
            // 
            lblSupplier.AutoSize = true;
            lblSupplier.Font = new Font("Segoe UI", 9F);
            lblSupplier.Location = new Point(10, 18);
            lblSupplier.Name = "lblSupplier";
            lblSupplier.Size = new Size(55, 15);
            lblSupplier.TabIndex = 0;
            lblSupplier.Text = "Supplier:";
            // 
            // cboSupplier
            // 
            cboSupplier.DropDownStyle = ComboBoxStyle.DropDownList;
            cboSupplier.Font = new Font("Segoe UI", 9F);
            cboSupplier.Location = new Point(70, 14);
            cboSupplier.Name = "cboSupplier";
            cboSupplier.Size = new Size(200, 23);
            cboSupplier.TabIndex = 1;
            // 
            // lblFromDate
            // 
            lblFromDate.AutoSize = true;
            lblFromDate.Font = new Font("Segoe UI", 9F);
            lblFromDate.Location = new Point(285, 18);
            lblFromDate.Name = "lblFromDate";
            lblFromDate.Size = new Size(66, 15);
            lblFromDate.TabIndex = 2;
            lblFromDate.Text = "From Date:";
            // 
            // dtpFromDate
            // 
            dtpFromDate.Font = new Font("Segoe UI", 9F);
            dtpFromDate.Format = DateTimePickerFormat.Short;
            dtpFromDate.Location = new Point(355, 14);
            dtpFromDate.Name = "dtpFromDate";
            dtpFromDate.Size = new Size(120, 23);
            dtpFromDate.TabIndex = 3;
            // 
            // lblToDate
            // 
            lblToDate.AutoSize = true;
            lblToDate.Font = new Font("Segoe UI", 9F);
            lblToDate.Location = new Point(490, 18);
            lblToDate.Name = "lblToDate";
            lblToDate.Size = new Size(50, 15);
            lblToDate.TabIndex = 4;
            lblToDate.Text = "To Date:";
            // 
            // dtpToDate
            // 
            dtpToDate.Font = new Font("Segoe UI", 9F);
            dtpToDate.Format = DateTimePickerFormat.Short;
            dtpToDate.Location = new Point(545, 14);
            dtpToDate.Name = "dtpToDate";
            dtpToDate.Size = new Size(120, 23);
            dtpToDate.TabIndex = 5;
            // 
            // lblProjectCode
            // 
            lblProjectCode.AutoSize = true;
            lblProjectCode.Font = new Font("Segoe UI", 9F);
            lblProjectCode.Location = new Point(10, 58);
            lblProjectCode.Name = "lblProjectCode";
            lblProjectCode.Size = new Size(79, 15);
            lblProjectCode.TabIndex = 8;
            lblProjectCode.Text = "Project Code:";
            // 
            // cboProjectCode
            // 
            cboProjectCode.DropDownStyle = ComboBoxStyle.DropDown;
            cboProjectCode.Font = new Font("Segoe UI", 9F);
            cboProjectCode.Location = new Point(95, 54);
            cboProjectCode.Name = "cboProjectCode";
            cboProjectCode.Size = new Size(175, 23);
            cboProjectCode.TabIndex = 9;
            // 
            // btnSave
            // 
            btnSave.BackColor = Color.FromArgb(40, 167, 69);
            btnSave.FlatStyle = FlatStyle.Flat;
            btnSave.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnSave.ForeColor = Color.White;
            btnSave.Location = new Point(475, 51);
            btnSave.Name = "btnSave";
            btnSave.Size = new Size(110, 30);
            btnSave.TabIndex = 12;
            btnSave.Text = "💾 Save";
            btnSave.UseVisualStyleBackColor = false;
            btnSave.Click += new EventHandler(btnSave_Click);
            // 
            // btnSearch
            // 
            btnSearch.BackColor = Color.FromArgb(0, 122, 204);
            btnSearch.FlatStyle = FlatStyle.Flat;
            btnSearch.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnSearch.ForeColor = Color.White;
            btnSearch.Location = new Point(290, 51);
            btnSearch.Name = "btnSearch";
            btnSearch.Size = new Size(90, 30);
            btnSearch.TabIndex = 10;
            btnSearch.Text = "🔍 Search";
            btnSearch.UseVisualStyleBackColor = false;
            btnSearch.Click += new EventHandler(btnSearch_Click);
            // 
            // btnClear
            // 
            btnClear.BackColor = Color.Gray;
            btnClear.FlatStyle = FlatStyle.Flat;
            btnClear.Font = new Font("Segoe UI", 9F);
            btnClear.ForeColor = Color.White;
            btnClear.Location = new Point(390, 51);
            btnClear.Name = "btnClear";
            btnClear.Size = new Size(75, 30);
            btnClear.TabIndex = 11;
            btnClear.Text = "Clear";
            btnClear.UseVisualStyleBackColor = false;
            btnClear.Click += new EventHandler(btnClear_Click);
            // 
            // pnGrid
            // 
            pnGrid.Controls.Add(dgvTrackDelivery);
            pnGrid.Dock = DockStyle.Fill;
            pnGrid.Location = new Point(3, 63);
            pnGrid.Name = "pnGrid";
            pnGrid.Size = new Size(944, 484);
            pnGrid.TabIndex = 1;
            // 
            // dgvTrackDelivery
            // 
            dgvTrackDelivery.AllowUserToAddRows = false;
            dgvTrackDelivery.AllowUserToDeleteRows = false;
            dgvTrackDelivery.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvTrackDelivery.BackgroundColor = Color.White;
            dgvTrackDelivery.BorderStyle = BorderStyle.Fixed3D;
            dgvTrackDelivery.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvTrackDelivery.Dock = DockStyle.Fill;
            dgvTrackDelivery.Location = new Point(0, 0);
            dgvTrackDelivery.Name = "dgvTrackDelivery";
            dgvTrackDelivery.ReadOnly = true;
            dgvTrackDelivery.RowHeadersWidth = 51;
            dgvTrackDelivery.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvTrackDelivery.Size = new Size(944, 484);
            dgvTrackDelivery.TabIndex = 0;
            // 
            // ucTrackDelivery
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            Controls.Add(tableLayoutPanel1);
            Name = "ucTrackDelivery";
            Size = new Size(950, 550);
            tableLayoutPanel1.ResumeLayout(false);
            pnFilter.ResumeLayout(false);
            pnFilter.PerformLayout();
            pnGrid.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(dgvTrackDelivery)).EndInit();
            ResumeLayout(false);
        }

        #endregion

        private TableLayoutPanel tableLayoutPanel1;
        private Panel pnFilter;
        private Panel pnGrid;
        private Label lblSupplier;
        private ComboBox cboSupplier;
        private Label lblFromDate;
        private DateTimePicker dtpFromDate;
        private Label lblToDate;
        private DateTimePicker dtpToDate;
        private Label lblProjectCode;
        private ComboBox cboProjectCode;
        private Button btnSearch;
        private Button btnClear;
        private Button btnSave;
        private DataGridView dgvTrackDelivery;
    }
}