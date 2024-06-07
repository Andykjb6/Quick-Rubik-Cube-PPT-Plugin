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
            this.TopMost = true; // 窗口总在最前
            this.ShowInTaskbar = false; // 不在任务栏中显示
            horizontalRadioButton.Checked = true; // 默认选中水平选项
            flipComboBox.SelectedIndex = 0; // 默认选中无翻转

            // 订阅 PictureBox 的 Paint 事件
            this.pictureBox.Paint += new PaintEventHandler(this.PictureBox_Paint);

            // 初始化颜色选项面板为不可见
            this.colorOptionsPanel.Visible = false;
        }

        private void PictureBox_Paint(object sender, PaintEventArgs e)
        {
            // 绘制表示透明像素的棋盘背景
            DrawCheckerboard(e.Graphics, pictureBox.ClientRectangle);

            if (pictureBox.Image != null)
            {
                // 保持图像的纵横比计算目标矩形
                Rectangle targetRect = CalculateAspectRatioRectangle(pictureBox.ClientRectangle, pictureBox.Image.Size);

                // 绘制图像
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
                // 图像比容器宽
                int width = container.Width;
                int height = (int)(width / imageAspectRatio);
                int x = container.X;
                int y = container.Y + (container.Height - height) / 2;
                return new Rectangle(x, y, width, height);
            }
            else
            {
                // 图像比容器高
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
                    overlayColor = Color.FromArgb(128, colorDialog.Color); // Adjust alpha to mix with the original image
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
            // 切换颜色选项面板的可见性
            this.colorOptionsPanel.Visible = !this.colorOptionsPanel.Visible;

            // 根据面板的可见性动态调整窗体的高度
            if (this.colorOptionsPanel.Visible)
            {
                // 如果面板可见，则增加窗口高度以容纳面板
                this.Height += this.colorOptionsPanel.Height;
            }
            else
            {
                // 如果面板不可见，则减小窗口高度以隐藏面板
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

            // 将 currentImage 的像素数据复制到一个数组中
            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    pixelData[x, y] = currentImage.GetPixel(x, y);
                }
            }

            // 使用并行处理加速像素操作
            Parallel.For(0, height, y =>
            {
                for (int x = 0; x < width; x++)
                {
                    Color pixelColor = pixelData[x, y];
                    int alpha = pixelColor.A;

                    if (this.horizontalRadioButton.Checked)
                    {
                        alpha = Math.Min(alpha, 255 - (transparencyValue * 255 / 100 * (width - x) / width));
                    }
                    else if (this.verticalRadioButton.Checked)
                    {
                        alpha = Math.Min(alpha, 255 - (transparencyValue * 255 / 100 * (height - y) / height));
                    }
                    else if (this.fullTransparencyRadioButton.Checked)
                    {
                        alpha = Math.Min(alpha, 255 - (transparencyValue * 255 / 100));
                    }
                    else if (this.radialTransparencyRadioButton.Checked)
                    {
                        double distance = Math.Sqrt(Math.Pow(x - width / 2, 2) + Math.Pow(y - height / 2, 2));
                        double maxDistance = Math.Sqrt(Math.Pow(width / 2, 2) + Math.Pow(height / 2, 2));
                        alpha = Math.Min(alpha, 255 - (int)(transparencyValue * 255 / 100 * distance / maxDistance));
                    }
                    else if (this.diagonalTransparencyRadioButton.Checked)
                    {
                        double distance = Math.Sqrt(Math.Pow(x, 2) + Math.Pow(y, 2));
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

            // 将结果数据应用到 processedImage
            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    processedImage.SetPixel(x, y, resultData[x, y]);
                }
            }

            // Apply flip based on the selected option
            switch (this.flipComboBox.SelectedItem.ToString())
            {
                case "水平翻转":
                    processedImage.RotateFlip(RotateFlipType.RotateNoneFlipX);
                    break;
                case "垂直翻转":
                    processedImage.RotateFlip(RotateFlipType.RotateNoneFlipY);
                    break;
            }

            this.pictureBox.Image = processedImage;
            this.pictureBox.Invalidate(); // Force repaint to show background
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
                this.pictureBox.Invalidate(); // Force repaint to show background
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

