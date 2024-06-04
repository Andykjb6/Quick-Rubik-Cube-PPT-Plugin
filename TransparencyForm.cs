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
    public class TransparencyForm : Form
    {
        private PictureBox pictureBox;
        private RadioButton horizontalRadioButton;
        private RadioButton verticalRadioButton;
        private RadioButton fullTransparencyRadioButton;
        private RadioButton radialTransparencyRadioButton;
        private RadioButton diagonalTransparencyRadioButton;
        private TrackBar transparencyTrackBar;
        private Label transparencyLabel;
        private ComboBox flipComboBox;
        private Button importButton;
        private Button exportButton;
        private Button colorOptionsButton;
        private Panel colorOptionsPanel;
        private CheckBox grayscaleCheckBox;
        private Button colorOverlayButton;
        private Button resetColorButton;
        private Bitmap currentImage;
        private Bitmap processedImage;
        private Color overlayColor = Color.Transparent;

        public TransparencyForm()
        {
            InitializeComponent();
            this.TopMost = true; // 窗口总在最前
            this.ShowInTaskbar = false; // 不在任务栏中显示
        }

        private void InitializeComponent()
        {
            this.pictureBox = new System.Windows.Forms.PictureBox();
            this.horizontalRadioButton = new System.Windows.Forms.RadioButton();
            this.verticalRadioButton = new System.Windows.Forms.RadioButton();
            this.fullTransparencyRadioButton = new System.Windows.Forms.RadioButton();
            this.radialTransparencyRadioButton = new System.Windows.Forms.RadioButton();
            this.diagonalTransparencyRadioButton = new System.Windows.Forms.RadioButton();
            this.transparencyTrackBar = new System.Windows.Forms.TrackBar();
            this.transparencyLabel = new System.Windows.Forms.Label();
            this.flipComboBox = new System.Windows.Forms.ComboBox();
            this.importButton = new System.Windows.Forms.Button();
            this.exportButton = new System.Windows.Forms.Button();
            this.colorOptionsButton = new System.Windows.Forms.Button();
            this.colorOptionsPanel = new System.Windows.Forms.Panel();
            this.grayscaleCheckBox = new System.Windows.Forms.CheckBox();
            this.colorOverlayButton = new System.Windows.Forms.Button();
            this.resetColorButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.transparencyTrackBar)).BeginInit();
            this.colorOptionsPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // pictureBox
            // 
            this.pictureBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pictureBox.Location = new System.Drawing.Point(0, 0);
            this.pictureBox.Name = "pictureBox";
            this.pictureBox.Size = new System.Drawing.Size(100, 50);
            this.pictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox.TabIndex = 0;
            this.pictureBox.TabStop = false;
            this.pictureBox.Click += new System.EventHandler(this.pictureBox_Click);
            this.pictureBox.Paint += new System.Windows.Forms.PaintEventHandler(this.PictureBox_Paint);
            // 
            // horizontalRadioButton
            // 
            this.horizontalRadioButton.Checked = true;
            this.horizontalRadioButton.Location = new System.Drawing.Point(0, 0);
            this.horizontalRadioButton.Name = "horizontalRadioButton";
            this.horizontalRadioButton.Size = new System.Drawing.Size(104, 24);
            this.horizontalRadioButton.TabIndex = 4;
            this.horizontalRadioButton.TabStop = true;
            this.horizontalRadioButton.Text = "水平";
            this.horizontalRadioButton.CheckedChanged += new System.EventHandler(this.RadioButton_CheckedChanged);
            // 
            // verticalRadioButton
            // 
            this.verticalRadioButton.Location = new System.Drawing.Point(0, 0);
            this.verticalRadioButton.Name = "verticalRadioButton";
            this.verticalRadioButton.Size = new System.Drawing.Size(104, 24);
            this.verticalRadioButton.TabIndex = 5;
            this.verticalRadioButton.Text = "垂直";
            this.verticalRadioButton.CheckedChanged += new System.EventHandler(this.RadioButton_CheckedChanged);
            // 
            // fullTransparencyRadioButton
            // 
            this.fullTransparencyRadioButton.Location = new System.Drawing.Point(0, 0);
            this.fullTransparencyRadioButton.Name = "fullTransparencyRadioButton";
            this.fullTransparencyRadioButton.Size = new System.Drawing.Size(104, 24);
            this.fullTransparencyRadioButton.TabIndex = 6;
            this.fullTransparencyRadioButton.Text = "整体";
            this.fullTransparencyRadioButton.CheckedChanged += new System.EventHandler(this.RadioButton_CheckedChanged);
            // 
            // radialTransparencyRadioButton
            // 
            this.radialTransparencyRadioButton.Location = new System.Drawing.Point(0, 0);
            this.radialTransparencyRadioButton.Name = "radialTransparencyRadioButton";
            this.radialTransparencyRadioButton.Size = new System.Drawing.Size(104, 24);
            this.radialTransparencyRadioButton.TabIndex = 7;
            this.radialTransparencyRadioButton.Text = "径向";
            this.radialTransparencyRadioButton.CheckedChanged += new System.EventHandler(this.RadioButton_CheckedChanged);
            // 
            // diagonalTransparencyRadioButton
            // 
            this.diagonalTransparencyRadioButton.Location = new System.Drawing.Point(0, 0);
            this.diagonalTransparencyRadioButton.Name = "diagonalTransparencyRadioButton";
            this.diagonalTransparencyRadioButton.Size = new System.Drawing.Size(104, 24);
            this.diagonalTransparencyRadioButton.TabIndex = 8;
            this.diagonalTransparencyRadioButton.Text = "对角";
            this.diagonalTransparencyRadioButton.CheckedChanged += new System.EventHandler(this.RadioButton_CheckedChanged);
            // 
            // transparencyTrackBar
            // 
            this.transparencyTrackBar.Location = new System.Drawing.Point(0, 0);
            this.transparencyTrackBar.Maximum = 100;
            this.transparencyTrackBar.Name = "transparencyTrackBar";
            this.transparencyTrackBar.Size = new System.Drawing.Size(104, 90);
            this.transparencyTrackBar.TabIndex = 1;
            this.transparencyTrackBar.TickFrequency = 10;
            this.transparencyTrackBar.Scroll += new System.EventHandler(this.TrackBar_Scroll);
            // 
            // transparencyLabel
            // 
            this.transparencyLabel.AutoSize = true;
            this.transparencyLabel.Location = new System.Drawing.Point(0, 0);
            this.transparencyLabel.Name = "transparencyLabel";
            this.transparencyLabel.Size = new System.Drawing.Size(22, 24);
            this.transparencyLabel.TabIndex = 2;
            this.transparencyLabel.Text = "0";
            // 
            // flipComboBox
            // 
            this.flipComboBox.Items.AddRange(new object[] {
            "无翻转",
            "水平翻转",
            "垂直翻转"});
            this.flipComboBox.Location = new System.Drawing.Point(0, 0);
            this.flipComboBox.Name = "flipComboBox";
            this.flipComboBox.Size = new System.Drawing.Size(121, 32);
            this.flipComboBox.TabIndex = 3;
            this.flipComboBox.SelectedIndexChanged += new System.EventHandler(this.FlipComboBox_SelectedIndexChanged);
            // 设置默认选项为 "无翻转"
            this.flipComboBox.SelectedIndex = 0;
            // 
            // importButton
            // 
            this.importButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(85)))), ((int)(((byte)(151)))));
            this.importButton.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.importButton.Location = new System.Drawing.Point(0, 0);
            this.importButton.Name = "importButton";
            this.importButton.Size = new System.Drawing.Size(120, 60);
            this.importButton.TabIndex = 9;
            this.importButton.Text = "所选导入";
            this.importButton.UseVisualStyleBackColor = false;
            this.importButton.Click += new System.EventHandler(this.ImportButton_Click);
            // 
            // exportButton
            // 
            this.exportButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(85)))), ((int)(((byte)(151)))));
            this.exportButton.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.exportButton.Location = new System.Drawing.Point(0, 0);
            this.exportButton.Name = "exportButton";
            this.exportButton.Size = new System.Drawing.Size(120, 60);
            this.exportButton.TabIndex = 10;
            this.exportButton.Text = "导出至幻灯片";
            this.exportButton.UseVisualStyleBackColor = false;
            this.exportButton.Click += new System.EventHandler(this.ExportButton_Click);
            // 
            // colorOptionsButton
            // 
            this.colorOptionsButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(85)))), ((int)(((byte)(151)))));
            this.colorOptionsButton.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.colorOptionsButton.Location = new System.Drawing.Point(0, 0);
            this.colorOptionsButton.Name = "colorOptionsButton";
            this.colorOptionsButton.Size = new System.Drawing.Size(75, 23);
            this.colorOptionsButton.TabIndex = 11;
            this.colorOptionsButton.Text = "颜色选项";
            this.colorOptionsButton.UseVisualStyleBackColor = false;
            this.colorOptionsButton.Click += new System.EventHandler(this.ColorOptionsButton_Click);
            // 
            // colorOptionsPanel
            // 
            this.colorOptionsPanel.Controls.Add(this.grayscaleCheckBox);
            this.colorOptionsPanel.Controls.Add(this.colorOverlayButton);
            this.colorOptionsPanel.Controls.Add(this.resetColorButton);
            this.colorOptionsPanel.Location = new System.Drawing.Point(0, 0);
            this.colorOptionsPanel.Name = "colorOptionsPanel";
            this.colorOptionsPanel.Size = new System.Drawing.Size(200, 100);
            this.colorOptionsPanel.TabIndex = 12;
            this.colorOptionsPanel.Visible = false;
            this.colorOptionsPanel.Paint += new System.Windows.Forms.PaintEventHandler(this.colorOptionsPanel_Paint);
            // 
            // grayscaleCheckBox
            // 
            this.grayscaleCheckBox.Location = new System.Drawing.Point(0, 0);
            this.grayscaleCheckBox.Name = "grayscaleCheckBox";
            this.grayscaleCheckBox.Size = new System.Drawing.Size(104, 24);
            this.grayscaleCheckBox.TabIndex = 0;
            this.grayscaleCheckBox.Text = "灰度";
            this.grayscaleCheckBox.CheckedChanged += new System.EventHandler(this.GrayscaleCheckBox_CheckedChanged);
            // 
            // colorOverlayButton
            // 
            this.colorOverlayButton.Location = new System.Drawing.Point(0, 0);
            this.colorOverlayButton.Name = "colorOverlayButton";
            this.colorOverlayButton.Size = new System.Drawing.Size(75, 23);
            this.colorOverlayButton.TabIndex = 1;
            this.colorOverlayButton.Text = "颜色";
            this.colorOverlayButton.Click += new System.EventHandler(this.ColorOverlayButton_Click);
            // 
            // resetColorButton
            // 
            this.resetColorButton.Location = new System.Drawing.Point(0, 0);
            this.resetColorButton.Name = "resetColorButton";
            this.resetColorButton.Size = new System.Drawing.Size(75, 23);
            this.resetColorButton.TabIndex = 2;
            this.resetColorButton.Text = "重置颜色叠加";
            this.resetColorButton.Click += new System.EventHandler(this.ResetColorButton_Click);
            // 
            // TransparencyForm
            // 
            this.ClientSize = new System.Drawing.Size(600, 675);
            this.Controls.Add(this.pictureBox);
            this.Controls.Add(this.transparencyTrackBar);
            this.Controls.Add(this.transparencyLabel);
            this.Controls.Add(this.flipComboBox);
            this.Controls.Add(this.horizontalRadioButton);
            this.Controls.Add(this.verticalRadioButton);
            this.Controls.Add(this.fullTransparencyRadioButton);
            this.Controls.Add(this.radialTransparencyRadioButton);
            this.Controls.Add(this.diagonalTransparencyRadioButton);
            this.Controls.Add(this.importButton);
            this.Controls.Add(this.exportButton);
            this.Controls.Add(this.colorOptionsButton);
            this.Controls.Add(this.colorOptionsPanel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "TransparencyForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "图片透明PRO";
            this.Load += new System.EventHandler(this.Form_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.transparencyTrackBar)).EndInit();
            this.colorOptionsPanel.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void Form_Load(object sender, EventArgs e)
        {
            LayoutControls();
        }

        private void LayoutControls()
        {
            int margin = 10;
            int controlHeight = 40;
            int labelWidth = 30;
            int comboBoxWidth = 110;

            // PictureBox
            this.pictureBox.Location = new Point(margin, margin);
            this.pictureBox.Size = new Size(this.ClientSize.Width - 2 * margin, 350);

            // Transparency TrackBar
            this.transparencyTrackBar.Location = new Point(margin, this.pictureBox.Bottom + margin);
            this.transparencyTrackBar.Size = new Size(this.ClientSize.Width - 3 * margin - labelWidth - comboBoxWidth, controlHeight);

            // Transparency Label
            this.transparencyLabel.Location = new Point(this.transparencyTrackBar.Right + margin, this.transparencyTrackBar.Top);

            // Flip ComboBox
            this.flipComboBox.Location = new Point(this.transparencyLabel.Right + margin, this.transparencyTrackBar.Top);
            this.flipComboBox.Size = new Size(comboBoxWidth, controlHeight);

            // RadioButtons
            int radioButtonTop = this.transparencyTrackBar.Bottom + margin;
            int radioButtonWidth = (this.ClientSize.Width - 6 * margin) / 5;
            this.horizontalRadioButton.Location = new Point(margin, radioButtonTop);
            this.horizontalRadioButton.Size = new Size(radioButtonWidth, controlHeight);
            this.verticalRadioButton.Location = new Point(this.horizontalRadioButton.Right + margin, radioButtonTop);
            this.verticalRadioButton.Size = new Size(radioButtonWidth, controlHeight);
            this.fullTransparencyRadioButton.Location = new Point(this.verticalRadioButton.Right + margin, radioButtonTop);
            this.fullTransparencyRadioButton.Size = new Size(radioButtonWidth, controlHeight);
            this.radialTransparencyRadioButton.Location = new Point(this.fullTransparencyRadioButton.Right + margin, radioButtonTop);
            this.radialTransparencyRadioButton.Size = new Size(radioButtonWidth, controlHeight);
            this.diagonalTransparencyRadioButton.Location = new Point(this.radialTransparencyRadioButton.Right + margin, radioButtonTop);
            this.diagonalTransparencyRadioButton.Size = new Size(radioButtonWidth, controlHeight);

            // Import Button
            this.importButton.Location = new Point(margin, this.horizontalRadioButton.Bottom + margin);
            this.importButton.Size = new Size((this.ClientSize.Width - 3 * margin) / 2, controlHeight + 5);

            // Export Button
            this.exportButton.Location = new Point(this.importButton.Right + margin, this.importButton.Top);
            this.exportButton.Size = new Size((this.ClientSize.Width - 3 * margin) / 2, controlHeight + 5);

            // Color Options Button
            this.colorOptionsButton.Location = new Point(margin, this.importButton.Bottom + margin);
            this.colorOptionsButton.Size = new Size(this.ClientSize.Width - 2 * margin, controlHeight + 5);

            // Color Options Panel
            this.colorOptionsPanel.Location = new Point(margin, this.colorOptionsButton.Bottom + margin);
            this.colorOptionsPanel.Size = new Size(this.ClientSize.Width - 2 * margin, 3 * controlHeight + 2 * margin);

            // Grayscale CheckBox
            this.grayscaleCheckBox.Location = new Point(margin, margin);
            this.grayscaleCheckBox.Size = new Size(this.colorOptionsPanel.ClientSize.Width - 2 * margin, controlHeight);

            // Color Overlay Button
            this.colorOverlayButton.Location = new Point(margin, this.grayscaleCheckBox.Bottom + margin);
            this.colorOverlayButton.Size = new Size(this.colorOptionsPanel.ClientSize.Width - 2 * margin, controlHeight);

            // Reset Color Button
            this.resetColorButton.Location = new Point(margin, this.colorOverlayButton.Bottom + margin);
            this.resetColorButton.Size = new Size(this.colorOptionsPanel.ClientSize.Width - 2 * margin, controlHeight);
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
            if (slide.Shapes.Count > 0)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Type == Office.MsoShapeType.msoPicture)
                    {
                        string tempPath = Path.GetTempFileName();
                        shape.Export(tempPath, PowerPoint.PpShapeFormat.ppShapeFormatPNG);

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
                        break;
                    }
                }
            }
        }

        private void ExportButton_Click(object sender, EventArgs e)
        {
            if (processedImage != null)
            {
                string newTempPath = Path.GetTempFileName() + ".png";
                processedImage.Save(newTempPath, ImageFormat.Png);

                var application = 课件帮PPT助手.Globals.ThisAddIn.Application;
                var slide = application.ActiveWindow.View.Slide;
                PowerPoint.Shape selectedShape = null;

                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Type == Office.MsoShapeType.msoPicture)
                    {
                        selectedShape = shape;
                        break;
                    }
                }

                if (selectedShape != null)
                {
                    float left = selectedShape.Left;
                    float top = selectedShape.Top;
                    float width = selectedShape.Width;
                    float height = selectedShape.Height;

                    selectedShape.Delete();
                    slide.Shapes.AddPicture(newTempPath, Office.MsoTriState.msoFalse,
                                            Office.MsoTriState.msoCTrue, left, top, width, height);
                }

                File.Delete(newTempPath);
            }
        }

        private void PictureBox_Paint(object sender, PaintEventArgs e)
        {
            // Draw checkerboard pattern to indicate transparency
            DrawCheckerboard(e.Graphics, pictureBox.ClientRectangle);

            if (pictureBox.Image != null)
            {
                // Calculate the target rectangle to keep the aspect ratio
                Rectangle targetRect = CalculateAspectRatioRectangle(pictureBox.ClientRectangle, pictureBox.Image.Size);

                // Draw the image
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
                // Image is wider than the container
                int width = container.Width;
                int height = (int)(width / imageAspectRatio);
                int x = container.X;
                int y = container.Y + (container.Height - height) / 2;
                return new Rectangle(x, y, width, height);
            }
            else
            {
                // Image is taller than the container
                int height = container.Height;
                int width = (int)(height * imageAspectRatio);
                int x = container.X + (container.Width - width) / 2;
                int y = container.Y;
                return new Rectangle(x, y, width, height);
            }
        }

        private void colorOptionsPanel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void pictureBox_Click(object sender, EventArgs e)
        {

        }
    }

    public class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            TransparencyForm form = new TransparencyForm();
            form.Show();
            Application.Run(new ApplicationContext()); // 使用 ApplicationContext 来防止窗口成为模式对话框
        }
    }
}
