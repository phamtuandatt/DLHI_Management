using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;

namespace MPR_Managerment.Forms
{
    public partial class frmShowImage : Form
    {
        private List<string> _allFiles = new List<string>();
        private float _zoomFactor = 1.0f;
        private bool _isDragging = false;
        private Point _lastMousePos;

        public frmShowImage(string folderPath)
        {
            InitializeComponent();
            // Thiết lập PictureBox để hỗ trợ zoom tốt hơn
            pictureBoxPreview.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBoxPreview.Dock = DockStyle.None;
            pictureBoxPreview.Location = new Point(0, 0);
            
            LoadImages(folderPath);
            this.MouseWheel += new MouseEventHandler(frmShowImage_MouseWheel);
        }

        private void LoadImages(string folderPath)
        {
            if (!Directory.Exists(folderPath))
            {
                MessageBox.Show("Thư mục không tồn tại!");
                return;
            }

            string[] extensions = { "*.png", "*.jpg", "*.jpeg", "*.bmp", "*.gif" };
            _allFiles = extensions.SelectMany(ext => Directory.GetFiles(folderPath, ext)).ToList();
            UpdateListBox(_allFiles);
        }

        private void UpdateListBox(List<string> files)
        {
            listBoxImages.Items.Clear();
            foreach (var file in files)
            {
                listBoxImages.Items.Add(file);
            }
        }

        private void txtFilter_TextChanged(object sender, EventArgs e)
        {
            string filter = txtFilter.Text.ToLower();
            var filtered = _allFiles.Where(f => Path.GetFileName(f).ToLower().Contains(filter)).ToList();
            UpdateListBox(filtered);
        }

        private void listBoxImages_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxImages.SelectedItem != null)
            {
                string path = listBoxImages.SelectedItem.ToString();
                if (File.Exists(path))
                {
                    pictureBoxPreview.Image = Image.FromFile(path);
                    _zoomFactor = 1.0f;
                    ApplyZoom();
                }
            }
        }

        private void btnZoomIn_Click(object sender, EventArgs e)
        {
            _zoomFactor += 0.1f;
            ApplyZoom();
        }

        private void btnZoomOut_Click(object sender, EventArgs e)
        {
            if (_zoomFactor > 0.2f)
            {
                _zoomFactor -= 0.1f;
                ApplyZoom();
            }
        }

        private void ApplyZoom()
        {
            if (pictureBoxPreview.Image != null)
            {
                // Thay đổi kích thước PictureBox để ảnh tự co giãn theo SizeMode.Zoom
                int newWidth = (int)(pictureBoxPreview.Image.Width * _zoomFactor);
                int newHeight = (int)(pictureBoxPreview.Image.Height * _zoomFactor);
                
                pictureBoxPreview.Size = new Size(newWidth, newHeight);
                panelPreview.Invalidate();
            }
        }

        private void frmShowImage_MouseWheel(object sender, MouseEventArgs e)
        {
            if (Control.ModifierKeys == Keys.Control)
            {
                if (e.Delta > 0)
                    _zoomFactor += 0.1f;
                else if (e.Delta < 0 && _zoomFactor > 0.2f)
                    _zoomFactor -= 0.1f;

                ApplyZoom();
            }
        }

        private void pictureBoxPreview_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                _isDragging = true;
                _lastMousePos = e.Location;
            }
        }

        private void pictureBoxPreview_MouseMove(object sender, MouseEventArgs e)
        {
            if (_isDragging)
            {
                int deltaX = e.X - _lastMousePos.X;
                int deltaY = e.Y - _lastMousePos.Y;
                panelPreview.AutoScrollPosition = new Point(-panelPreview.AutoScrollPosition.X - deltaX, -panelPreview.AutoScrollPosition.Y - deltaY);
            }
        }

        private void pictureBoxPreview_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                _isDragging = false;
            }
        }
    }
}
