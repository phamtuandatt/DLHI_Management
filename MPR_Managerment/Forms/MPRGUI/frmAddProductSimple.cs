using System;
using System.Drawing;
using System.Windows.Forms;
using MPR_Managerment.Models;

namespace MPR_Managerment.Forms.MPRGUI
{
    public class frmAddProductSimple : Form
    {
        private TextBox txtName, txtDes2, txtCode, txtMaterial, txtThickness, txtDepth, txtWidth, txtWeb, txtFlag, txtLength, txtWeight, txtUnit;
        private Button btnSave, btnCancel;

        public ProductModel NewProduct { get; private set; }

        public frmAddProductSimple()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Thêm vật tư mới";
            this.Size = new Size(800, 225);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            var table = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 6,
                Padding = new Padding(10),
                Margin = new Padding(0)
            };

            AddRow(table, "Tên vật tư:", out txtName, "Mô tả (Des2):", out txtDes2, "Mã vật tư:", out txtCode);
            AddRow(table, "Mã vật liệu:", out txtMaterial, "Độ dày (A):", out txtThickness, "Chiều sâu (B):", out txtDepth);
            AddRow(table, "Chiều rộng (C):", out txtWidth, "Web (D):", out txtWeb, "Flag (E):", out txtFlag);
            AddRow(table, "Chiều dài (F):", out txtLength, "Trọng lượng (G):", out txtWeight, "Đơn vị:", out txtUnit);

            //btnSave = new Button { Text = "💾 Lưu", Size = new Size(120, 40), FlatStyle = FlatStyle.Flat, BackColor = (40, 167, 69), Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            btnSave = new Button { Size = new Size(120, 40) };
            Common.Common.CreateButtonSave(btnSave, "💾 Lưu");
            btnSave.Click += BtnSave_Click;

            //btnCancel = new Button { Text = "❌ Hủy", Size = new Size(120, 40), FlatStyle = FlatStyle.Flat, BackColor = Color.LightCoral, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            btnCancel = new Button() { Size = new Size(120, 40) };
            Common.Common.CreateButtonCancel(btnCancel, "❌ Hủy");
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            var btnPanel = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 60, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(10, 5, 10, 5) };
            btnPanel.Controls.Add(btnCancel);
            btnPanel.Controls.Add(btnSave);

            this.Controls.Add(table);
            this.Controls.Add(btnPanel);
        }

        private void AddRow(TableLayoutPanel table, string l1, out TextBox t1, string l2, out TextBox t2, string l3, out TextBox t3)
        {
            table.Controls.Add(new Label { Text = l1, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 5, 5, 5) });
            table.Controls.Add(t1 = new TextBox { Width = 150, Margin = new Padding(0, 2, 10, 2) });
            table.Controls.Add(new Label { Text = l2, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 5, 5, 5) });
            table.Controls.Add(t2 = new TextBox { Width = 150, Margin = new Padding(0, 2, 10, 2) });
            table.Controls.Add(new Label { Text = l3, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 5, 5, 5) });
            table.Controls.Add(t3 = new TextBox { Width = 150, Margin = new Padding(0, 2, 0, 2) });
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            NewProduct = new ProductModel
            {
                Name = txtName.Text,
                Des2 = txtDes2.Text,
                Code = txtCode.Text,
                ProdMaterialCode = txtMaterial.Text,
                A_Thickness = txtThickness.Text,
                B_Depth = txtDepth.Text,
                C_Width = txtWidth.Text,
                D_Web = txtWeb.Text,
                E_Flag = txtFlag.Text,
                F_Length = txtLength.Text,
                G_Weight = txtWeight.Text,
                Unit = txtUnit.Text
            };
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}