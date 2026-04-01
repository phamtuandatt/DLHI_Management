namespace MPR_Managerment.Forms.ItemCodeGUI
{
    partial class frmOptions
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
            btnGetCode = new Button();
            btnCreateCode = new Button();
            SuspendLayout();
            // 
            // btnGetCode
            // 
            btnGetCode.Location = new Point(12, 12);
            btnGetCode.Name = "btnGetCode";
            btnGetCode.Size = new Size(191, 40);
            btnGetCode.TabIndex = 0;
            btnGetCode.Text = "GET CODE";
            btnGetCode.UseVisualStyleBackColor = true;
            btnGetCode.Click += btnGetCode_Click;
            // 
            // btnCreateCode
            // 
            btnCreateCode.Location = new Point(209, 12);
            btnCreateCode.Name = "btnCreateCode";
            btnCreateCode.Size = new Size(191, 40);
            btnCreateCode.TabIndex = 0;
            btnCreateCode.Text = "GENERATE CODE";
            btnCreateCode.UseVisualStyleBackColor = true;
            btnCreateCode.Click += btnCreateCode_Click;
            // 
            // frmOptions
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(416, 60);
            Controls.Add(btnCreateCode);
            Controls.Add(btnGetCode);
            Name = "frmOptions";
            Text = "frmOptions";
            ResumeLayout(false);
        }

        #endregion

        private Button btnGetCode;
        private Button btnCreateCode;
    }
}