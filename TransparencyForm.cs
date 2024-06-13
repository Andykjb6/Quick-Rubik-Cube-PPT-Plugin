using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Threading.Tasks;

namespace 课件帮PPT助手
{
    public partial class TransparencyForm : Form
    {
        private Bitmap currentImage;
        private Bitmap processedImage;
        private Color overlayColor = Color.Transparent;

        public TransparencyForm()
        {
            InitializeComponent();
            this.TopMost = true;
            this.ShowInTaskbar = false;
            horizontalRadioButton.Checked = true;
            flipComboBox.SelectedIndex = 0;
            this.pictureBox.Paint += new PaintEventHandler(this.PictureBox_Paint);
            this.colorOptionsPanel.Visible = false;
        }

        private void PictureBox_Paint(object sender, PaintEventArgs e)
        {
            DrawCheckerboard(e.Graphics, pictureBox.ClientRectangle);

            if (pictureBox.Image != null)
            {
                Rectangle targetRect = CalculateAspectRatioRectangle(pictureBox.ClientRectangle, pictureBox.Image.Size);
                e.Graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                e.Graphics.DrawImage(pictureBox.Image, targetRect);
            }
        }

        private void DrawCheckerboard(Graphics g, Rectangle rect)
        {
            int cellSize = 10;
            for (int y = 0; y < rect.Height; y += cellSize)
            {
                for (int x = 0; x < rect.Width; x += cellSize)
                {
                    bool isWhite = ((x / cellSize) % 2 == (y / cellSize) % 2);
                    using (Brush brush = new SolidBrush(isWhite ? Color.White : Color.LightGray))
                    {
                        g.FillRectangle(brush, x, y, cellSize, cellSize);
                    }
                }
            }
        }

        private Rectangle CalculateAspectRatioRectangle(Rectangle container, Size imageSize)
        {
            float containerAspectRatio = (float)container.Width / container.Height;
            float imageAspectRatio = (float)imageSize.Width / imageSize.Height;

            if (imageAspectRatio > containerAspectRatio)
            {
                int width = container.Width;
                int height = (int)(width / imageAspectRatio);
                int x = container.X;
                int y = container.Y + (container.Height - height) / 2;
                return new Rectangle(x, y, width, height);
            }
            else
            {
                int height = container.Height;
                int width = (int)(height * imageAspectRatio);
                int x = container.X + (container.Width - width) / 2;
                int y = container.Y;
                return new Rectangle(x, y, width, height);
            }
        }

        private void RadioButton_CheckedChanged(object sender, EventArgs e)
        {
            ApplyTransparency();
        }

        private void TrackBar_Scroll(object sender, EventArgs e)
        {
            this.transparencyLabel.Text = this.transparencyTrackBar.Value.ToString();
            ApplyTransparency();
        }

        private void FlipComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyTransparency();
        }

        private void GrayscaleCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            ApplyTransparency();
        }

        private void ColorOverlayButton_Click(object sender, EventArgs e)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    overlayColor = Color.FromArgb(128, colorDialog.Color);
                    ApplyTransparency();
                }
            }
        }

        private void ResetColorButton_Click(object sender, EventArgs e)
        {
            overlayColor = Color.Transparent;
            ApplyTransparency();
        }

        private void ColorOptionsButton_Click(object sender, EventArgs e)
        {
            this.colorOptionsPanel.Visible = !this.colorOptionsPanel.Visible;
            if (this.colorOptionsPanel.Visible)
            {
                this.Height += this.colorOptionsPanel.Height;
            }
            else
            {
                this.Height -= this.colorOptionsPanel.Height;
            }
        }

        private void ApplyTransparency()
        {
            if (currentImage == null) return;

            int transparencyValue = this.transparencyTrackBar.Value;
            int width = currentImage.Width;
            int height = currentImage.Height;

            processedImage = new Bitmap(width, height, PixelFormat.Format32bppArgb);
            Color[,] pixelData = new Color[width, height];
            Color[,] resultData = new Color[width, height];

            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    pixelData[x, y] = currentImage.GetPixel(x, y);
                }
            }

            bool isFlippedHorizontally = false;
            bool isFlippedVertically = false;
            switch (this.flipComboBox.SelectedItem.ToString())
            {
                case "水平翻转":
                    isFlippedHorizontally = true;
                    break;
                case "垂直翻转":
                    isFlippedVertically = true;
                    break;
            }

            Parallel.For(0, height, y =>
            {
                for (int x = 0; x < width; x++)
                {
                    Color pixelColor = pixelData[x, y];
                    int alpha = pixelColor.A;

                    int xPos = isFlippedHorizontally ? width - x - 1 : x;
                    int yPos = isFlippedVertically ? height - y - 1 : y;

                    if (this.horizontalRadioButton.Checked)
                    {
                        alpha = Math.Min(alpha, 255 - (transparencyValue * 255 / 100 * (width - xPos) / width));
                    }
                    else if (this.verticalRadioButton.Checked)
                    {
                        alpha = Math.Min(alpha, 255 - (transparencyValue * 255 / 100 * (height - yPos) / height));
                    }
                    else if (this.fullTransparencyRadioButton.Checked)
                    {
                        alpha = Math.Min(alpha, 255 - (transparencyValue * 255 / 100));
                    }
                    else if (this.radialTransparencyRadioButton.Checked)
                    {
                        double distance = Math.Sqrt(Math.Pow(xPos - width / 2, 2) + Math.Pow(yPos - height / 2, 2));
                        double maxDistance = Math.Sqrt(Math.Pow(width / 2, 2) + Math.Pow(height / 2, 2));
                        alpha = Math.Min(alpha, 255 - (int)(transparencyValue * 255 / 100 * distance / maxDistance));
                    }
                    else if (this.diagonalTransparencyRadioButton.Checked)
                    {
                        double distance = Math.Sqrt(Math.Pow(xPos, 2) + Math.Pow(yPos, 2));
                        double maxDistance = Math.Sqrt(Math.Pow(width, 2) + Math.Pow(height, 2));
                        alpha = Math.Min(alpha, 255 - (int)(transparencyValue * 255 / 100 * distance / maxDistance));
                    }

                    if (this.grayscaleCheckBox.Checked)
                    {
                        int gray = (int)(pixelColor.R * 0.3 + pixelColor.G * 0.59 + pixelColor.B * 0.11);
                        pixelColor = Color.FromArgb(pixelColor.A, gray, gray, gray);
                    }

                    if (overlayColor != Color.Transparent)
                    {
                        int r = (pixelColor.R * (255 - overlayColor.A) + overlayColor.R * overlayColor.A) / 255;
                        int g = (pixelColor.G * (255 - overlayColor.A) + overlayColor.G * overlayColor.A) / 255;
                        int b = (pixelColor.B * (255 - overlayColor.A) + overlayColor.B * overlayColor.A) / 255;
                        pixelColor = Color.FromArgb(pixelColor.A, r, g, b);
                    }

                    resultData[x, y] = Color.FromArgb(alpha, pixelColor.R, pixelColor.G, pixelColor.B);
                }
            });

            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    processedImage.SetPixel(x, y, resultData[x, y]);
                }
            }

            if (isFlippedHorizontally)
            {
                processedImage.RotateFlip(RotateFlipType.RotateNoneFlipX);
            }

            if (isFlippedVertically)
            {
                processedImage.RotateFlip(RotateFlipType.RotateNoneFlipY);
            }

            this.pictureBox.Image = processedImage;
            this.pictureBox.Invalidate();
        }

        private void ImportButton_Click(object sender, EventArgs e)
        {
            var application = 课件帮PPT助手.Globals.ThisAddIn.Application;
            var slide = application.ActiveWindow.View.Slide;
            PowerPoint.ShapeRange selectedShapes = application.ActiveWindow.Selection.ShapeRange;

            if (selectedShapes != null && selectedShapes.Count == 1 && selectedShapes[1].Type == Office.MsoShapeType.msoPicture)
            {
                var selectedShape = selectedShapes[1];
                string tempPath = Path.GetTempFileName();
                selectedShape.Export(tempPath, PowerPoint.PpShapeFormat.ppShapeFormatPNG);

                using (var tempImage = Image.FromFile(tempPath))
                {
                    currentImage = new Bitmap(tempImage.Width, tempImage.Height, PixelFormat.Format32bppArgb);
                    using (Graphics g = Graphics.FromImage(currentImage))
                    {
                        g.DrawImage(tempImage, new Rectangle(0, 0, tempImage.Width, tempImage.Height));
                    }
                }

                this.pictureBox.Image = currentImage;
                this.pictureBox.Invalidate();
                File.Delete(tempPath);
            }
            else
            {
                MessageBox.Show("请先选择一张图片。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ExportButton_Click(object sender, EventArgs e)
        {
            var application = 课件帮PPT助手.Globals.ThisAddIn.Application;
            var slide = application.ActiveWindow.View.Slide;
            PowerPoint.ShapeRange selectedShapes = application.ActiveWindow.Selection.ShapeRange;

            if (processedImage != null && selectedShapes != null && selectedShapes.Count == 1 && selectedShapes[1].Type == Office.MsoShapeType.msoPicture)
            {
                var selectedShape = selectedShapes[1];
                string newTempPath = Path.GetTempFileName() + ".png";
                processedImage.Save(newTempPath, ImageFormat.Png);

                float left = selectedShape.Left;
                float top = selectedShape.Top;
                float width = selectedShape.Width;
                float height = selectedShape.Height;

                selectedShape.Delete();
                slide.Shapes.AddPicture(newTempPath, Office.MsoTriState.msoFalse,
                                        Office.MsoTriState.msoCTrue, left, top, width, height);

                File.Delete(newTempPath);
            }
            else
            {
                MessageBox.Show("请先选择一张图片。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
