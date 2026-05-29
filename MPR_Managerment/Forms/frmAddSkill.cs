using System;
using System.Drawing;
using System.Windows.Forms;
using MPR_Managerment.Services;

namespace MPR_Managerment.Forms
{
    public class frmAddSkill : Form
    {
        public string SkillName { get; private set; } = "";

        private TextBox txtName;
        private ComboBox cboType;
        private TextBox txtSlash;
        private Label lblSlash;
        private TextBox txtIcon;
        private TextBox txtDesc;
        private TextBox txtTemplate;
        private Button btnSave;
        private Button btnCancel;

        private static readonly Color C_BG     = Color.FromArgb(26, 30, 42);
        private static readonly Color C_HEADER  = Color.FromArgb(35, 40, 58);
        private static readonly Color C_INPUT   = Color.FromArgb(44, 50, 72);
        private static readonly Color C_TEXT    = Color.FromArgb(226, 232, 240);
        private static readonly Color C_SUBTEXT = Color.FromArgb(148, 163, 184);
        private static readonly Color C_ACCENT  = Color.FromArgb(99, 102, 241);

        public frmAddSkill()
        {
            Text = "Thêm AI Skill mới";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false; MinimizeBox = false;
            Size = new Size(440, 400);
            BackColor = C_BG;
            ForeColor = C_TEXT;

            BuildUI();
        }

        private void BuildUI()
        {
            int y = 14, lw = 110, cx = 124, cw = 288;

            Label Lbl(string text) => new Label
            {
                Text = text, Location = new Point(14, y + 3),
                Size = new Size(lw, 20), ForeColor = C_SUBTEXT,
                Font = new Font("Segoe UI", 9f)
            };
            TextBox Txt(bool multiline = false) => new TextBox
            {
                Location = new Point(cx, y), Width = cw,
                BackColor = C_INPUT, ForeColor = C_TEXT,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9.5f),
                Multiline = multiline,
                Height = multiline ? 72 : 24
            };

            // Tên skill
            Controls.Add(Lbl("Tên skill:"));
            txtName = Txt(); Controls.Add(txtName); y += 32;

            // Loại
            Controls.Add(Lbl("Loại:"));
            cboType = new ComboBox
            {
                Location = new Point(cx, y), Width = cw,
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = C_INPUT, ForeColor = C_TEXT,
                Font = new Font("Segoe UI", 9.5f), Height = 24,
                FlatStyle = FlatStyle.Flat
            };
            cboType.Items.AddRange(new object[] { "⚡ Câu hỏi nhanh (quick)", "/ Lệnh gõ tắt (slash)", "🔬 Workflow (workflow)" });
            cboType.SelectedIndex = 0;
            cboType.SelectedIndexChanged += (s, e) => lblSlash.Visible = txtSlash.Visible = cboType.SelectedIndex == 1;
            Controls.Add(cboType); y += 32;

            // Lệnh tắt (chỉ visible khi slash)
            lblSlash = Lbl("Lệnh /xxx:"); lblSlash.Visible = false; Controls.Add(lblSlash);
            txtSlash = Txt(); txtSlash.Visible = false;
            txtSlash.PlaceholderText = "/po, /mpr, /tonkho...";
            Controls.Add(txtSlash); y += 32;

            // Icon
            Controls.Add(Lbl("Icon (emoji):"));
            txtIcon = Txt();
            txtIcon.Text = "⚡"; txtIcon.Width = 60;
            Controls.Add(txtIcon); y += 32;

            // Mô tả
            Controls.Add(Lbl("Mô tả ngắn:"));
            txtDesc = Txt();
            txtDesc.PlaceholderText = "Tooltip hiển thị khi hover (tuỳ chọn)";
            Controls.Add(txtDesc); y += 32;

            // Template
            Controls.Add(new Label
            {
                Text = "Template câu hỏi:", Location = new Point(14, y),
                Size = new Size(lw, 20), ForeColor = C_SUBTEXT,
                Font = new Font("Segoe UI", 9f)
            });
            txtTemplate = Txt(multiline: true);
            txtTemplate.PlaceholderText = "Câu hỏi gửi AI. Dùng {param} để nhận tham số từ lệnh /xxx hoặc workflow.";
            Controls.Add(txtTemplate); y += 84;

            // Buttons
            btnSave = new Button
            {
                Text = "✅ Lưu skill", Location = new Point(cx, y),
                Size = new Size(136, 32), FlatStyle = FlatStyle.Flat,
                BackColor = C_ACCENT, ForeColor = Color.White,
                Font = new Font("Segoe UI Semibold", 9.5f), Cursor = Cursors.Hand
            };
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Click += BtnSave_Click;
            Controls.Add(btnSave);

            btnCancel = new Button
            {
                Text = "Hủy", Location = new Point(cx + 144, y),
                Size = new Size(72, 32), FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 65, 90), ForeColor = C_SUBTEXT,
                Font = new Font("Segoe UI", 9.5f), Cursor = Cursors.Hand
            };
            btnCancel.FlatAppearance.BorderSize = 0;
            btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
            Controls.Add(btnCancel);

            ClientSize = new Size(440, y + 56);
            AcceptButton = btnSave;
            CancelButton = btnCancel;
        }

        private string GetType() => cboType.SelectedIndex switch
        {
            1 => "slash",
            2 => "workflow",
            _ => "quick"
        };

        private void BtnSave_Click(object sender, EventArgs e)
        {
            string name = txtName.Text.Trim();
            string template = txtTemplate.Text.Trim();
            string type = GetType();
            string slash = type == "slash" ? txtSlash.Text.Trim() : "";
            string icon = txtIcon.Text.Trim();
            string desc = txtDesc.Text.Trim();

            if (string.IsNullOrEmpty(name)) { Warn("Vui lòng nhập tên skill."); return; }
            if (string.IsNullOrEmpty(template)) { Warn("Vui lòng nhập template câu hỏi."); return; }
            if (type == "slash" && (string.IsNullOrEmpty(slash) || !slash.StartsWith("/")))
            { Warn("Lệnh tắt phải bắt đầu bằng '/' (ví dụ: /po)."); return; }

            string result = AISearchService.AddSkill(name, type, template, icon, desc, slash);
            if (result.StartsWith("⚠️")) { Warn(result); return; }

            SkillName = name;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void Warn(string msg) =>
            MessageBox.Show(this, msg, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }
}
