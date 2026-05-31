namespace MPR_Managerment.Forms
{
    partial class frmShowImage
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.ListBox listBoxImages;
        private System.Windows.Forms.PictureBox pictureBoxPreview;
        private System.Windows.Forms.TextBox txtFilter;
        private System.Windows.Forms.Button btnZoomIn;
        private System.Windows.Forms.Button btnZoomOut;
        private System.Windows.Forms.Panel panelPreview;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            listBoxImages = new ListBox();
            pictureBoxPreview = new PictureBox();
            txtFilter = new TextBox();
            btnZoomIn = new Button();
            btnZoomOut = new Button();
            panelPreview = new Panel();
            tableLayoutPanel1 = new TableLayoutPanel();
            panel1 = new Panel();
            panel2 = new Panel();
            ((System.ComponentModel.ISupportInitialize)pictureBoxPreview).BeginInit();
            panelPreview.SuspendLayout();
            tableLayoutPanel1.SuspendLayout();
            panel1.SuspendLayout();
            panel2.SuspendLayout();
            SuspendLayout();
            // 
            // listBoxImages
            // 
            listBoxImages.Dock = DockStyle.Fill;
            listBoxImages.FormattingEnabled = true;
            listBoxImages.ItemHeight = 15;
            listBoxImages.Location = new Point(3, 33);
            listBoxImages.Name = "listBoxImages";
            listBoxImages.Size = new Size(194, 604);
            listBoxImages.TabIndex = 1;
            listBoxImages.SelectedIndexChanged += listBoxImages_SelectedIndexChanged;
            // 
            // pictureBoxPreview
            // 
            pictureBoxPreview.Dock = DockStyle.Fill;
            pictureBoxPreview.Location = new Point(0, 0);
            pictureBoxPreview.Name = "pictureBoxPreview";
            pictureBoxPreview.Size = new Size(946, 604);
            pictureBoxPreview.TabIndex = 0;
            pictureBoxPreview.TabStop = false;
            pictureBoxPreview.MouseDown += pictureBoxPreview_MouseDown;
            pictureBoxPreview.MouseMove += pictureBoxPreview_MouseMove;
            pictureBoxPreview.MouseUp += pictureBoxPreview_MouseUp;
            // 
            // txtFilter
            // 
            txtFilter.Dock = DockStyle.Fill;
            txtFilter.Location = new Point(0, 0);
            txtFilter.Name = "txtFilter";
            txtFilter.Size = new Size(194, 23);
            txtFilter.TabIndex = 0;
            txtFilter.TextChanged += txtFilter_TextChanged;
            // 
            // btnZoomIn
            // 
            btnZoomIn.Location = new Point(4, 1);
            btnZoomIn.Name = "btnZoomIn";
            btnZoomIn.Size = new Size(75, 23);
            btnZoomIn.TabIndex = 3;
            btnZoomIn.Text = "Zoom In";
            btnZoomIn.Click += btnZoomIn_Click;
            // 
            // btnZoomOut
            // 
            btnZoomOut.Location = new Point(85, 1);
            btnZoomOut.Name = "btnZoomOut";
            btnZoomOut.Size = new Size(75, 23);
            btnZoomOut.TabIndex = 4;
            btnZoomOut.Text = "Zoom Out";
            btnZoomOut.Click += btnZoomOut_Click;
            // 
            // panelPreview
            // 
            panelPreview.AutoScroll = true;
            panelPreview.Controls.Add(pictureBoxPreview);
            panelPreview.Dock = DockStyle.Fill;
            panelPreview.Location = new Point(203, 33);
            panelPreview.Name = "panelPreview";
            panelPreview.Size = new Size(946, 604);
            panelPreview.TabIndex = 2;
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.ColumnCount = 2;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.Controls.Add(panelPreview, 1, 1);
            tableLayoutPanel1.Controls.Add(listBoxImages, 0, 1);
            tableLayoutPanel1.Controls.Add(panel1, 0, 0);
            tableLayoutPanel1.Controls.Add(panel2, 1, 0);
            tableLayoutPanel1.Dock = DockStyle.Fill;
            tableLayoutPanel1.Location = new Point(0, 0);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 2;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.Size = new Size(1152, 640);
            tableLayoutPanel1.TabIndex = 5;
            // 
            // panel1
            // 
            panel1.BackColor = Color.White;
            panel1.Controls.Add(txtFilter);
            panel1.Dock = DockStyle.Fill;
            panel1.Location = new Point(3, 3);
            panel1.Name = "panel1";
            panel1.Size = new Size(194, 24);
            panel1.TabIndex = 3;
            // 
            // panel2
            // 
            panel2.BackColor = Color.White;
            panel2.Controls.Add(btnZoomOut);
            panel2.Controls.Add(btnZoomIn);
            panel2.Dock = DockStyle.Fill;
            panel2.Location = new Point(203, 3);
            panel2.Name = "panel2";
            panel2.Size = new Size(946, 24);
            panel2.TabIndex = 3;
            // 
            // frmShowImage
            // 
            ClientSize = new Size(1152, 640);
            Controls.Add(tableLayoutPanel1);
            Name = "frmShowImage";
            Text = "Xem trước hình ảnh";
            ((System.ComponentModel.ISupportInitialize)pictureBoxPreview).EndInit();
            panelPreview.ResumeLayout(false);
            tableLayoutPanel1.ResumeLayout(false);
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            panel2.ResumeLayout(false);
            ResumeLayout(false);
        }
    }
}
