using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace MPR_Managerment.Forms
{
    public partial class frmShowImage : Form
    {
        private List<string> _allFiles = new List<string>();
        private float _zoomFactor = 1.0f;
        private Point _panOffset = new Point(0, 0);
        private Point _lastMousePos;
        private bool _isDragging = false;

        public frmShowImage(string folderPath)
        {
            InitializeComponent();
            
            pictureBoxPreview.SizeMode = PictureBoxSizeMode.Normal;
            pictureBoxPreview.Paint += PictureBoxPreview_Paint;
            pictureBoxPreview.MouseDown += PictureBoxPreview_MouseDown;
            pictureBoxPreview.MouseMove += PictureBoxPreview_MouseMove;
            pictureBoxPreview.MouseUp += PictureBoxPreview_MouseUp;
            pictureBoxPreview.MouseWheel += PictureBoxPreview_MouseWheel;
            
            LoadImages(folderPath);
        }

        private void LoadImages(string folderPath)
        {
            if (!Directory.Exists(folderPath)) return;
            string[] extensions = { "*.png", "*.jpg", "*.jpeg", "*.bmp", "*.gif" };
            _allFiles = extensions.SelectMany(ext => Directory.GetFiles(folderPath, ext)).ToList();
            listBoxImages.Items.AddRange(_allFiles.Select(Path.GetFileName).ToArray());
        }

        private void PictureBoxPreview_Paint(object sender, PaintEventArgs e)
        {
            if (pictureBoxPreview.Image == null) return;

            e.Graphics.Clear(pictureBoxPreview.BackColor);
            e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;

            // Tính toán vị trí trung tâm
            float scaledWidth = pictureBoxPreview.Image.Width * _zoomFactor;
            float scaledHeight = pictureBoxPreview.Image.Height * _zoomFactor;
            
            float centerX = (pictureBoxPreview.Width - scaledWidth) / 2f;
            float centerY = (pictureBoxPreview.Height - scaledHeight) / 2f;

            e.Graphics.TranslateTransform(centerX + _panOffset.X, centerY + _panOffset.Y);
            e.Graphics.ScaleTransform(_zoomFactor, _zoomFactor);
            
            e.Graphics.DrawImage(pictureBoxPreview.Image, 0, 0);
        }

        private void PictureBoxPreview_MouseWheel(object sender, MouseEventArgs e)
        {
            if (Control.ModifierKeys == Keys.Control)
            {
                if (e.Delta > 0) _zoomFactor *= 1.1f;
                else _zoomFactor /= 1.1f;
                
                _zoomFactor = Math.Max(0.1f, Math.Min(_zoomFactor, 10.0f));
                pictureBoxPreview.Invalidate();
            }
        }

        private void PictureBoxPreview_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left || e.Button == MouseButtons.Right)
            {
                _isDragging = true;
                _lastMousePos = e.Location;
            }
        }

        private void PictureBoxPreview_MouseMove(object sender, MouseEventArgs e)
        {
            if (_isDragging)
            {
                _panOffset.X += (e.X - _lastMousePos.X);
                _panOffset.Y += (e.Y - _lastMousePos.Y);
                _lastMousePos = e.Location;
                pictureBoxPreview.Invalidate();
            }
        }

        private void PictureBoxPreview_MouseUp(object sender, MouseEventArgs e) => _isDragging = false;

        private void listBoxImages_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxImages.SelectedIndex >= 0)
            {
                pictureBoxPreview.Image?.Dispose();
                pictureBoxPreview.Image = Image.FromFile(_allFiles[listBoxImages.SelectedIndex]);
                
                // Reset zoom để ảnh khớp với PictureBox ban đầu
                float ratioW = (float)pictureBoxPreview.Width / pictureBoxPreview.Image.Width;
                float ratioH = (float)pictureBoxPreview.Height / pictureBoxPreview.Image.Height;
                _zoomFactor = Math.Min(ratioW, ratioH);
                
                _panOffset = new Point(0, 0);
                pictureBoxPreview.Invalidate();
            }
        }

        private void btnZoomIn_Click(object sender, EventArgs e) { _zoomFactor *= 1.2f; pictureBoxPreview.Invalidate(); }
        private void btnZoomOut_Click(object sender, EventArgs e) { _zoomFactor /= 1.2f; pictureBoxPreview.Invalidate(); }
        private void txtFilter_TextChanged(object sender, EventArgs e) { }
    }
}
