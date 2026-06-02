using System;
using System.Windows.Forms;

namespace MPR_Managerment.Common
{
    public static class MessageBoxHelper
    {
        private static Form? GetOwnerForm(Control? owner)
        {
            try
            {
                if (owner is Form ownerForm)
                    return ownerForm;
                if (owner != null)
                {
                    var f = owner.FindForm();
                    if (f != null)
                        return f;
                }
                if (Form.ActiveForm != null)
                    return Form.ActiveForm;
                if (Application.OpenForms.Count > 0)
                    return Application.OpenForms[0];
            }
            catch
            {
                // fallback to null
            }
            return null;
        }

        public static DialogResult ShowInfo(Control? owner, string message, string title = "Thông báo")
        {
            var f = GetOwnerForm(owner);
            if (f != null && !f.IsDisposed)
            {
                f.BringToFront();
                f.Activate();
                return MessageBox.Show(f, message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            return MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static DialogResult ShowWarning(Control? owner, string message, string title = "Cảnh báo")
        {
            var f = GetOwnerForm(owner);
            if (f != null && !f.IsDisposed)
            {
                f.BringToFront();
                f.Activate();
                return MessageBox.Show(f, message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        public static DialogResult ShowError(Control? owner, string message, string title = "Lỗi")
        {
            var f = GetOwnerForm(owner);
            if (f != null && !f.IsDisposed)
            {
                f.BringToFront();
                f.Activate();
                return MessageBox.Show(f, message, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public static DialogResult ShowQuestion(Control? owner, string message, string title = "Xác nhận")
        {
            var f = GetOwnerForm(owner);
            if (f != null && !f.IsDisposed)
            {
                f.BringToFront();
                f.Activate();
                return MessageBox.Show(f, message, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            }
            return MessageBox.Show(message, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        }
    }
}