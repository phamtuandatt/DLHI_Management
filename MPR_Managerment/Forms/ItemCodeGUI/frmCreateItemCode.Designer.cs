namespace MPR_Managerment.Forms.ItemCodeGUI
{
    partial class frmCreateItemCode
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
            DataGridViewCellStyle dataGridViewCellStyle1 = new DataGridViewCellStyle();
            panel1 = new Panel();
            groupBox1 = new GroupBox();
            dgvItemExist = new DataGridView();
            material_detail_id = new DataGridViewTextBoxColumn();
            material_detail_number = new DataGridViewTextBoxColumn();
            material_detail_name = new DataGridViewTextBoxColumn();
            material_detail_code = new DataGridViewTextBoxColumn();
            item_code_existed = new DataGridViewTextBoxColumn();
            tableLayoutPanel9 = new TableLayoutPanel();
            btnSave = new Button();
            btnCancel = new Button();
            tableLayoutPanel5 = new TableLayoutPanel();
            btnGenerate = new Button();
            label6 = new Label();
            txtCode = new TextBox();
            tableLayoutPanel4 = new TableLayoutPanel();
            label5 = new Label();
            cboStandard = new ComboBox();
            tableLayoutPanel3 = new TableLayoutPanel();
            label4 = new Label();
            cboOriginal = new ComboBox();
            tableLayoutPanel2 = new TableLayoutPanel();
            label3 = new Label();
            cboMaterial = new ComboBox();
            tableLayoutPanel1 = new TableLayoutPanel();
            label2 = new Label();
            cboMaterialCate = new ComboBox();
            label1 = new Label();
            panel1.SuspendLayout();
            groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dgvItemExist).BeginInit();
            tableLayoutPanel9.SuspendLayout();
            tableLayoutPanel5.SuspendLayout();
            tableLayoutPanel4.SuspendLayout();
            tableLayoutPanel3.SuspendLayout();
            tableLayoutPanel2.SuspendLayout();
            tableLayoutPanel1.SuspendLayout();
            SuspendLayout();
            // 
            // panel1
            // 
            panel1.Controls.Add(groupBox1);
            panel1.Controls.Add(tableLayoutPanel9);
            panel1.Controls.Add(tableLayoutPanel5);
            panel1.Controls.Add(tableLayoutPanel4);
            panel1.Controls.Add(tableLayoutPanel3);
            panel1.Controls.Add(tableLayoutPanel2);
            panel1.Controls.Add(tableLayoutPanel1);
            panel1.Controls.Add(label1);
            panel1.Dock = DockStyle.Fill;
            panel1.Location = new Point(0, 0);
            panel1.Name = "panel1";
            panel1.Size = new Size(1417, 352);
            panel1.TabIndex = 0;
            panel1.Paint += panel1_Paint;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(dgvItemExist);
            groupBox1.Location = new Point(606, 69);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(799, 272);
            groupBox1.TabIndex = 25;
            groupBox1.TabStop = false;
            groupBox1.Text = "Items";
            // 
            // dgvItemExist
            // 
            dgvItemExist.AllowUserToAddRows = false;
            dgvItemExist.AllowUserToDeleteRows = false;
            dgvItemExist.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvItemExist.BackgroundColor = Color.White;
            dataGridViewCellStyle1.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = Color.Blue;
            dataGridViewCellStyle1.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            dataGridViewCellStyle1.ForeColor = Color.White;
            dataGridViewCellStyle1.SelectionBackColor = SystemColors.GradientInactiveCaption;
            dataGridViewCellStyle1.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = DataGridViewTriState.True;
            dgvItemExist.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            dgvItemExist.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvItemExist.Columns.AddRange(new DataGridViewColumn[] { material_detail_id, material_detail_number, material_detail_name, material_detail_code, item_code_existed });
            dgvItemExist.Dock = DockStyle.Fill;
            dgvItemExist.Location = new Point(3, 23);
            dgvItemExist.Name = "dgvItemExist";
            dgvItemExist.ReadOnly = true;
            dgvItemExist.RowHeadersWidth = 51;
            dgvItemExist.Size = new Size(793, 246);
            dgvItemExist.TabIndex = 0;
            dgvItemExist.CellClick += dgvItemExist_CellClick;
            dgvItemExist.RowPostPaint += dgvItemExist_RowPostPaint;
            dgvItemExist.RowPrePaint += dgvItemExist_RowPrePaint;
            // 
            // material_detail_id
            // 
            material_detail_id.DataPropertyName = "material_detail_id";
            material_detail_id.HeaderText = "ID";
            material_detail_id.MinimumWidth = 6;
            material_detail_id.Name = "material_detail_id";
            material_detail_id.ReadOnly = true;
            material_detail_id.Visible = false;
            // 
            // material_detail_number
            // 
            material_detail_number.DataPropertyName = "material_detail_number";
            material_detail_number.FillWeight = 20F;
            material_detail_number.HeaderText = "Number";
            material_detail_number.MinimumWidth = 6;
            material_detail_number.Name = "material_detail_number";
            material_detail_number.ReadOnly = true;
            // 
            // material_detail_name
            // 
            material_detail_name.DataPropertyName = "material_detail_name";
            material_detail_name.FillWeight = 49.5095673F;
            material_detail_name.HeaderText = "Name";
            material_detail_name.MinimumWidth = 6;
            material_detail_name.Name = "material_detail_name";
            material_detail_name.ReadOnly = true;
            // 
            // material_detail_code
            // 
            material_detail_code.DataPropertyName = "material_detail_code";
            material_detail_code.HeaderText = "Material ID";
            material_detail_code.MinimumWidth = 6;
            material_detail_code.Name = "material_detail_code";
            material_detail_code.ReadOnly = true;
            material_detail_code.Visible = false;
            // 
            // item_code_existed
            // 
            item_code_existed.DataPropertyName = "item_code_existed";
            item_code_existed.HeaderText = "Code";
            item_code_existed.MinimumWidth = 6;
            item_code_existed.Name = "item_code_existed";
            item_code_existed.ReadOnly = true;
            item_code_existed.Visible = false;
            // 
            // tableLayoutPanel9
            // 
            tableLayoutPanel9.ColumnCount = 2;
            tableLayoutPanel9.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 49.91511F));
            tableLayoutPanel9.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50.08489F));
            tableLayoutPanel9.Controls.Add(btnSave, 1, 0);
            tableLayoutPanel9.Controls.Add(btnCancel, 0, 0);
            tableLayoutPanel9.Location = new Point(6, 291);
            tableLayoutPanel9.Name = "tableLayoutPanel9";
            tableLayoutPanel9.RowCount = 1;
            tableLayoutPanel9.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel9.Size = new Size(594, 53);
            tableLayoutPanel9.TabIndex = 1;
            // 
            // btnSave
            // 
            btnSave.BackColor = Color.Lime;
            btnSave.Dock = DockStyle.Fill;
            btnSave.FlatStyle = FlatStyle.Flat;
            btnSave.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnSave.ForeColor = SystemColors.ButtonFace;
            btnSave.Location = new Point(299, 3);
            btnSave.Name = "btnSave";
            btnSave.Size = new Size(292, 47);
            btnSave.TabIndex = 0;
            btnSave.Text = "💾 SAVE";
            btnSave.UseVisualStyleBackColor = false;
            btnSave.Click += btnSave_Click;
            // 
            // btnCancel
            // 
            btnCancel.BackColor = SystemColors.ActiveBorder;
            btnCancel.Dock = DockStyle.Fill;
            btnCancel.FlatStyle = FlatStyle.Flat;
            btnCancel.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnCancel.ForeColor = Color.Snow;
            btnCancel.Location = new Point(3, 3);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new Size(290, 47);
            btnCancel.TabIndex = 0;
            btnCancel.Text = "🆕 Cancel";
            btnCancel.UseVisualStyleBackColor = false;
            btnCancel.Click += btnCancel_Click;
            // 
            // tableLayoutPanel5
            // 
            tableLayoutPanel5.ColumnCount = 3;
            tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 31.8906612F));
            tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 68.10934F));
            tableLayoutPanel5.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 151F));
            tableLayoutPanel5.Controls.Add(btnGenerate, 2, 0);
            tableLayoutPanel5.Controls.Add(label6, 0, 0);
            tableLayoutPanel5.Controls.Add(txtCode, 1, 0);
            tableLayoutPanel5.Location = new Point(6, 245);
            tableLayoutPanel5.Name = "tableLayoutPanel5";
            tableLayoutPanel5.RowCount = 1;
            tableLayoutPanel5.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel5.Size = new Size(594, 39);
            tableLayoutPanel5.TabIndex = 24;
            // 
            // btnGenerate
            // 
            btnGenerate.BackColor = Color.FromArgb(255, 128, 0);
            btnGenerate.Dock = DockStyle.Fill;
            btnGenerate.FlatStyle = FlatStyle.Flat;
            btnGenerate.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnGenerate.ForeColor = Color.White;
            btnGenerate.Location = new Point(445, 3);
            btnGenerate.Name = "btnGenerate";
            btnGenerate.Size = new Size(146, 33);
            btnGenerate.TabIndex = 4;
            btnGenerate.Text = "➕ Generate";
            btnGenerate.UseVisualStyleBackColor = false;
            btnGenerate.Click += btnGenerate_Click;
            // 
            // label6
            // 
            label6.Dock = DockStyle.Fill;
            label6.Location = new Point(3, 0);
            label6.Name = "label6";
            label6.Size = new Size(135, 39);
            label6.TabIndex = 2;
            label6.Text = "Code:";
            label6.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // txtCode
            // 
            txtCode.Dock = DockStyle.Fill;
            txtCode.Location = new Point(144, 5);
            txtCode.Margin = new Padding(3, 5, 3, 3);
            txtCode.Name = "txtCode";
            txtCode.Size = new Size(295, 27);
            txtCode.TabIndex = 30;
            // 
            // tableLayoutPanel4
            // 
            tableLayoutPanel4.ColumnCount = 2;
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 23.7691F));
            tableLayoutPanel4.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 76.2309F));
            tableLayoutPanel4.Controls.Add(label5, 0, 0);
            tableLayoutPanel4.Controls.Add(cboStandard, 1, 0);
            tableLayoutPanel4.Location = new Point(6, 200);
            tableLayoutPanel4.Name = "tableLayoutPanel4";
            tableLayoutPanel4.RowCount = 1;
            tableLayoutPanel4.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel4.Size = new Size(594, 39);
            tableLayoutPanel4.TabIndex = 23;
            // 
            // label5
            // 
            label5.Dock = DockStyle.Fill;
            label5.Location = new Point(3, 0);
            label5.Name = "label5";
            label5.Size = new Size(135, 39);
            label5.TabIndex = 2;
            label5.Text = "Standard:";
            label5.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // cboStandard
            // 
            cboStandard.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cboStandard.AutoCompleteSource = AutoCompleteSource.ListItems;
            cboStandard.Dock = DockStyle.Fill;
            cboStandard.FormattingEnabled = true;
            cboStandard.Location = new Point(144, 8);
            cboStandard.Margin = new Padding(3, 8, 3, 3);
            cboStandard.Name = "cboStandard";
            cboStandard.Size = new Size(447, 28);
            cboStandard.TabIndex = 29;
            cboStandard.Validating += cboStandard_Validating;
            // 
            // tableLayoutPanel3
            // 
            tableLayoutPanel3.ColumnCount = 2;
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 23.7691F));
            tableLayoutPanel3.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 76.2309F));
            tableLayoutPanel3.Controls.Add(label4, 0, 0);
            tableLayoutPanel3.Controls.Add(cboOriginal, 1, 0);
            tableLayoutPanel3.Location = new Point(6, 155);
            tableLayoutPanel3.Name = "tableLayoutPanel3";
            tableLayoutPanel3.RowCount = 1;
            tableLayoutPanel3.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel3.Size = new Size(594, 39);
            tableLayoutPanel3.TabIndex = 22;
            // 
            // label4
            // 
            label4.Dock = DockStyle.Fill;
            label4.Location = new Point(3, 0);
            label4.Name = "label4";
            label4.Size = new Size(135, 39);
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
            cboOriginal.Location = new Point(144, 8);
            cboOriginal.Margin = new Padding(3, 8, 3, 3);
            cboOriginal.Name = "cboOriginal";
            cboOriginal.Size = new Size(447, 28);
            cboOriginal.TabIndex = 28;
            cboOriginal.Validating += cboOriginal_Validating;
            // 
            // tableLayoutPanel2
            // 
            tableLayoutPanel2.ColumnCount = 2;
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 23.7691F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 76.2309F));
            tableLayoutPanel2.Controls.Add(label3, 0, 0);
            tableLayoutPanel2.Controls.Add(cboMaterial, 1, 0);
            tableLayoutPanel2.Location = new Point(6, 113);
            tableLayoutPanel2.Name = "tableLayoutPanel2";
            tableLayoutPanel2.RowCount = 1;
            tableLayoutPanel2.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel2.Size = new Size(594, 39);
            tableLayoutPanel2.TabIndex = 21;
            // 
            // label3
            // 
            label3.Dock = DockStyle.Fill;
            label3.Location = new Point(3, 0);
            label3.Name = "label3";
            label3.Size = new Size(135, 39);
            label3.TabIndex = 2;
            label3.Text = "Material:";
            label3.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // cboMaterial
            // 
            cboMaterial.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cboMaterial.AutoCompleteSource = AutoCompleteSource.ListItems;
            cboMaterial.Dock = DockStyle.Fill;
            cboMaterial.FormattingEnabled = true;
            cboMaterial.Location = new Point(144, 8);
            cboMaterial.Margin = new Padding(3, 8, 3, 3);
            cboMaterial.Name = "cboMaterial";
            cboMaterial.Size = new Size(447, 28);
            cboMaterial.TabIndex = 27;
            cboMaterial.SelectedIndexChanged += cboMaterial_SelectedIndexChanged;
            cboMaterial.Validating += cboMaterial_Validating;
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.ColumnCount = 2;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 23.7691F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 76.2309F));
            tableLayoutPanel1.Controls.Add(label2, 0, 0);
            tableLayoutPanel1.Controls.Add(cboMaterialCate, 1, 0);
            tableLayoutPanel1.Location = new Point(6, 68);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 1;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            tableLayoutPanel1.Size = new Size(594, 39);
            tableLayoutPanel1.TabIndex = 20;
            // 
            // label2
            // 
            label2.Dock = DockStyle.Fill;
            label2.Location = new Point(3, 0);
            label2.Name = "label2";
            label2.Size = new Size(135, 39);
            label2.TabIndex = 2;
            label2.Text = "Material Category:";
            label2.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // cboMaterialCate
            // 
            cboMaterialCate.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cboMaterialCate.AutoCompleteSource = AutoCompleteSource.ListItems;
            cboMaterialCate.Dock = DockStyle.Fill;
            cboMaterialCate.FormattingEnabled = true;
            cboMaterialCate.Location = new Point(144, 8);
            cboMaterialCate.Margin = new Padding(3, 8, 3, 3);
            cboMaterialCate.Name = "cboMaterialCate";
            cboMaterialCate.Size = new Size(447, 28);
            cboMaterialCate.TabIndex = 26;
            cboMaterialCate.SelectedIndexChanged += cboMaterialCate_SelectedIndexChanged;
            cboMaterialCate.Validating += cboMaterialCate_Validating;
            // 
            // label1
            // 
            label1.BackColor = Color.Blue;
            label1.Dock = DockStyle.Top;
            label1.Font = new Font("Segoe UI", 20F, FontStyle.Bold);
            label1.ForeColor = Color.White;
            label1.Location = new Point(0, 0);
            label1.Name = "label1";
            label1.Size = new Size(1417, 65);
            label1.TabIndex = 0;
            label1.Text = "TẠO CODE";
            label1.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // frmCreateItemCode
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1417, 352);
            Controls.Add(panel1);
            Name = "frmCreateItemCode";
            Text = "frmCreateItemCode";
            Load += frmCreateItemCode_Load;
            panel1.ResumeLayout(false);
            groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dgvItemExist).EndInit();
            tableLayoutPanel9.ResumeLayout(false);
            tableLayoutPanel5.ResumeLayout(false);
            tableLayoutPanel5.PerformLayout();
            tableLayoutPanel4.ResumeLayout(false);
            tableLayoutPanel3.ResumeLayout(false);
            tableLayoutPanel2.ResumeLayout(false);
            tableLayoutPanel1.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private Panel panel1;
        private Label label1;
        private TableLayoutPanel tableLayoutPanel1;
        private Label label2;
        private ComboBox cboMaterialCate;
        private TableLayoutPanel tableLayoutPanel9;
        private Button btnSave;
        private Button btnCancel;
        private TableLayoutPanel tableLayoutPanel5;
        private Button btnGenerate;
        private Label label6;
        private TextBox txtCode;
        private TableLayoutPanel tableLayoutPanel4;
        private Label label5;
        private ComboBox cboStandard;
        private TableLayoutPanel tableLayoutPanel3;
        private Label label4;
        private ComboBox cboOriginal;
        private TableLayoutPanel tableLayoutPanel2;
        private Label label3;
        private ComboBox cboMaterial;
        private GroupBox groupBox1;
        private DataGridView dgvItemExist;
        private DataGridViewTextBoxColumn material_detail_id;
        private DataGridViewTextBoxColumn material_detail_number;
        private DataGridViewTextBoxColumn material_detail_name;
        private DataGridViewTextBoxColumn material_detail_code;
        private DataGridViewTextBoxColumn item_code_existed;
    }
}