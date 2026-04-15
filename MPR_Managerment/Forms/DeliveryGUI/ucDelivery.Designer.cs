namespace MPR_Managerment.Forms.DeliveryGUI
{
    partial class ucDelivery
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
            pHead = new Panel();
            pDeliveryLeft = new Panel();
            pDeliveryRight = new GroupBox();
            pTutorial = new Panel();
            pGrid = new Panel();
            pDeliveryLeft.SuspendLayout();
            SuspendLayout();
            // 
            // pHead
            // 
            pHead.Dock = DockStyle.Top;
            pHead.Location = new Point(0, 0);
            pHead.Name = "pHead";
            pHead.Size = new Size(1360, 38);
            pHead.TabIndex = 0;
            // 
            // pDeliveryLeft
            // 
            pDeliveryLeft.Controls.Add(pGrid);
            pDeliveryLeft.Controls.Add(pTutorial);
            pDeliveryLeft.Dock = DockStyle.Left;
            pDeliveryLeft.Location = new Point(0, 38);
            pDeliveryLeft.Name = "pDeliveryLeft";
            pDeliveryLeft.Size = new Size(375, 758);
            pDeliveryLeft.TabIndex = 1;
            // 
            // pDeliveryRight
            // 
            pDeliveryRight.Dock = DockStyle.Fill;
            pDeliveryRight.Location = new Point(375, 38);
            pDeliveryRight.Name = "pDeliveryRight";
            pDeliveryRight.Size = new Size(985, 758);
            pDeliveryRight.TabIndex = 2;
            pDeliveryRight.TabStop = false;
            pDeliveryRight.Text = "PREVIEW";
            // 
            // pTutorial
            // 
            pTutorial.Dock = DockStyle.Top;
            pTutorial.Location = new Point(0, 0);
            pTutorial.Name = "pTutorial";
            pTutorial.Size = new Size(375, 26);
            pTutorial.TabIndex = 0;
            // 
            // pGrid
            // 
            pGrid.Dock = DockStyle.Fill;
            pGrid.Location = new Point(0, 26);
            pGrid.Name = "pGrid";
            pGrid.Size = new Size(375, 732);
            pGrid.TabIndex = 1;
            // 
            // ucDelivery
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            Controls.Add(pDeliveryRight);
            Controls.Add(pDeliveryLeft);
            Controls.Add(pHead);
            Name = "ucDelivery";
            Size = new Size(1360, 796);
            pDeliveryLeft.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private Panel pHead;
        private Panel pDeliveryLeft;
        private GroupBox pDeliveryRight;
        private Panel pGrid;
        private Panel pTutorial;
    }
}
