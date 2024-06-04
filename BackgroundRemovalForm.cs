using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using OpenCvSharp;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Size = OpenCvSharp.Size;

public class BackgroundRemovalForm : Form
{
    private PictureBox originalPictureBox;
    private PictureBox processedPictureBox;
    private Button removeBackgroundButton;
    private Button exportButton;
    private Button resetButton;
    private TrackBar smoothnessTrackBar;
    private Label smoothnessLabel;
    private TrackBar sharpnessTrackBar;
    private Label sharpnessLabel;
    private ProgressBar progressBar;
    private TrackBar thresholdTrackBar;
    private Label thresholdLabel;
    private TrackBar iterationsTrackBar;
    private Label iterationsLabel;
    private Mat originalImage;
    private Mat processedImage;
    private string exportedImagePath;
    private PowerPoint.Application pptApp;
    private PowerPoint.Slide currentSlide;
    private string tempImagePath;
    private int thresholdValue = 128;
    private int iterations = 5;
    private int smoothnessValue = 0;
    private int sharpnessValue = 0;

    public BackgroundRemovalForm(Image image, PowerPoint.Application pptApp, PowerPoint.Slide currentSlide, string tempImagePath)
    {
        InitializeComponent();
        originalPictureBox.Image = AddPaddingToImage(image, 10);
        originalImage = BitmapToMat((Bitmap)originalPictureBox.Image);
        this.pptApp = pptApp;
        this.currentSlide = currentSlide;
        this.tempImagePath = tempImagePath;

        AutoAdjustParameters(originalImage);

        this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
        this.ShowInTaskbar = false;
    }

    public string ExportedImagePath => exportedImagePath;

    private void InitializeComponent()
    {
        this.originalPictureBox = new PictureBox();
        this.processedPictureBox = new PictureBox();
        this.removeBackgroundButton = new Button();
        this.exportButton = new Button();
        this.resetButton = new Button();
        this.progressBar = new ProgressBar();
        this.thresholdTrackBar = new TrackBar();
        this.thresholdLabel = new Label();
        this.iterationsTrackBar = new TrackBar();
        this.iterationsLabel = new Label();
        this.smoothnessTrackBar = new TrackBar();
        this.smoothnessLabel = new Label();
        this.sharpnessTrackBar = new TrackBar();
        this.sharpnessLabel = new Label();

        // originalPictureBox
        this.originalPictureBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
        this.originalPictureBox.Location = new System.Drawing.Point(12, 12);
        this.originalPictureBox.Size = new System.Drawing.Size(256, 256);
        this.originalPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;

        // processedPictureBox
        this.processedPictureBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
        this.processedPictureBox.Location = new System.Drawing.Point(284, 12);
        this.processedPictureBox.Size = new System.Drawing.Size(256, 256);
        this.processedPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;

        // removeBackgroundButton
        this.removeBackgroundButton.Location = new System.Drawing.Point(12, 650);
        this.removeBackgroundButton.Size = new System.Drawing.Size(100, 45);
        this.removeBackgroundButton.Text = "去背景";
        this.removeBackgroundButton.Click += new System.EventHandler(this.RemoveBackgroundButton_Click);

        // exportButton
        this.exportButton.Location = new System.Drawing.Point(440, 650);
        this.exportButton.Size = new System.Drawing.Size(100, 45);
        this.exportButton.Text = "导出";
        this.exportButton.Click += new System.EventHandler(this.ExportButton_Click);

        // resetButton
        this.resetButton.Location = new System.Drawing.Point(270, 650);
        this.resetButton.Size = new System.Drawing.Size(100, 45);
        this.resetButton.Text = "重置";
        this.resetButton.Click += new System.EventHandler(this.ResetButton_Click);

        // smoothnessTrackBar
        this.smoothnessTrackBar.Location = new System.Drawing.Point(12, 460);
        this.smoothnessTrackBar.Size = new System.Drawing.Size(256, 45);
        this.smoothnessTrackBar.Minimum = 0;
        this.smoothnessTrackBar.Maximum = 100;
        this.smoothnessTrackBar.Value = smoothnessValue;
        this.smoothnessTrackBar.Scroll += new System.EventHandler(this.SmoothnessTrackBar_Scroll);

        // smoothnessLabel
        this.smoothnessLabel.Location = new System.Drawing.Point(274, 460);
        this.smoothnessLabel.Size = new System.Drawing.Size(140, 45);
        this.smoothnessLabel.Text = "平滑度: " + smoothnessValue;

        // sharpnessTrackBar
        this.sharpnessTrackBar.Location = new System.Drawing.Point(12, 550);
        this.sharpnessTrackBar.Size = new System.Drawing.Size(256, 45);
        this.sharpnessTrackBar.Minimum = 0;
        this.sharpnessTrackBar.Maximum = 100;
        this.sharpnessTrackBar.Value = sharpnessValue;
        this.sharpnessTrackBar.Scroll += new System.EventHandler(this.SharpnessTrackBar_Scroll);

        // sharpnessLabel
        this.sharpnessLabel.Location = new System.Drawing.Point(274, 550);
        this.sharpnessLabel.Size = new System.Drawing.Size(140, 45);
        this.sharpnessLabel.Text = "锐化度: " + sharpnessValue;

        // progressBar
        this.progressBar.Location = new System.Drawing.Point(120, 650);
        this.progressBar.Size = new System.Drawing.Size(140, 45);

        // thresholdTrackBar
        this.thresholdTrackBar.Location = new System.Drawing.Point(12, 284);
        this.thresholdTrackBar.Size = new System.Drawing.Size(256, 45);
        this.thresholdTrackBar.Minimum = 0;
        this.thresholdTrackBar.Maximum = 255;
        this.thresholdTrackBar.Value = thresholdValue;
        this.thresholdTrackBar.Scroll += new System.EventHandler(this.ThresholdTrackBar_Scroll);

        // thresholdLabel
        this.thresholdLabel.Location = new System.Drawing.Point(274, 284);
        this.thresholdLabel.Size = new System.Drawing.Size(140, 45);
        this.thresholdLabel.Text = "阈值: " + thresholdValue;

        // iterationsTrackBar
        this.iterationsTrackBar.Location = new System.Drawing.Point(12, 370);
        this.iterationsTrackBar.Size = new System.Drawing.Size(256, 45);
        this.iterationsTrackBar.Minimum = 1;
        this.iterationsTrackBar.Maximum = 10;
        this.iterationsTrackBar.Value = iterations;
        this.iterationsTrackBar.Scroll += new System.EventHandler(this.IterationsTrackBar_Scroll);

        // iterationsLabel
        this.iterationsLabel.Location = new System.Drawing.Point(274, 370);
        this.iterationsLabel.Size = new System.Drawing.Size(140, 45);
        this.iterationsLabel.Text = "迭代: " + iterations;

        // BackgroundRemovalForm
        this.ClientSize = new System.Drawing.Size(564, 800);
        this.Controls.Add(this.originalPictureBox);
        this.Controls.Add(this.processedPictureBox);
        this.Controls.Add(this.removeBackgroundButton);
        this.Controls.Add(this.exportButton);
        this.Controls.Add(this.resetButton);
        this.Controls.Add(this.progressBar);
        this.Controls.Add(this.thresholdTrackBar);
        this.Controls.Add(this.thresholdLabel);
        this.Controls.Add(this.iterationsTrackBar);
        this.Controls.Add(this.iterationsLabel);
        this.Controls.Add(this.smoothnessTrackBar);
        this.Controls.Add(this.smoothnessLabel);
        this.Controls.Add(this.sharpnessTrackBar);
        this.Controls.Add(this.sharpnessLabel);
        this.Text = "便捷抠图Beta版本";
        this.FormClosed += new FormClosedEventHandler(this.BackgroundRemovalForm_FormClosed);
    }

    private void ThresholdTrackBar_Scroll(object sender, EventArgs e)
    {
        thresholdValue = thresholdTrackBar.Value;
        thresholdLabel.Text = "阈值: " + thresholdValue;
    }

    private void IterationsTrackBar_Scroll(object sender, EventArgs e)
    {
        iterations = iterationsTrackBar.Value;
        iterationsLabel.Text = "迭代: " + iterations;
    }

    private void SmoothnessTrackBar_Scroll(object sender, EventArgs e)
    {
        smoothnessValue = smoothnessTrackBar.Value;
        smoothnessLabel.Text = "平滑度: " + smoothnessValue;
        ApplyEdgeProcessing();
    }

    private void SharpnessTrackBar_Scroll(object sender, EventArgs e)
    {
        sharpnessValue = sharpnessTrackBar.Value;
        sharpnessLabel.Text = "锐化度: " + sharpnessValue;
        ApplyEdgeProcessing();
    }

    private void ApplyBackgroundRemoval()
    {
        try
        {
            // 使用原始图像进行处理，确保图像类型正确
            Mat tempImage = originalImage.Clone();

            // 对图像进行平滑处理，减少噪声的影响
            Cv2.GaussianBlur(tempImage, tempImage, new Size(5, 5), 0);

            Rect rectangle = new Rect(10, 10, tempImage.Width - 20, tempImage.Height - 20);
            Mat mask = new Mat(tempImage.Size(), MatType.CV_8UC1, new Scalar(2));
            Mat bgdModel = new Mat();
            Mat fgdModel = new Mat();

            Cv2.GrabCut(tempImage, mask, rectangle, bgdModel, fgdModel, iterations, GrabCutModes.InitWithRect);

            mask = mask & 1;

            Mat foreground = new Mat(tempImage.Size(), MatType.CV_8UC4, new Scalar(0, 0, 0, 0));
            for (int y = 0; y < tempImage.Rows; y++)
            {
                for (int x = 0; x < tempImage.Cols; x++)
                {
                    if (mask.At<byte>(y, x) == 1)
                    {
                        Vec3b color = tempImage.At<Vec3b>(y, x);
                        foreground.Set(y, x, new Vec4b(color.Item0, color.Item1, color.Item2, 255)); // 确保颜色通道正确
                    }
                }
            }

            // 确保处理后的图像没有被压缩或缩小
            processedImage = foreground.Clone();
            UpdateProcessedPictureBox(foreground);

            // 释放资源
            tempImage.Dispose();
            mask.Dispose();
            bgdModel.Dispose();
            fgdModel.Dispose();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"图像处理时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }



    private void ApplyEdgeProcessing()
    {
        if (processedImage == null) return;

        Mat tempImage = processedImage.Clone();

        if (smoothnessValue > 0)
        {
            Cv2.GaussianBlur(tempImage, tempImage, new Size(0, 0), smoothnessValue);
        }

        if (sharpnessValue > 0)
        {
            Mat kernel = new Mat(3, 3, MatType.CV_32F, new float[]
            {
                -1, -1, -1,
                -1,  9, -1,
                -1, -1, -1
            });
            Cv2.Filter2D(tempImage, tempImage, -1, kernel);
        }

        UpdateProcessedPictureBox(tempImage);
        processedImage = tempImage;
    }

    private void UpdateProcessedPictureBox(Mat image)
    {
        Invoke((Action)(() =>
        {
            if (processedPictureBox.Image != null)
            {
                processedPictureBox.Image.Dispose();
            }

            // 创建棋盘背景
            Bitmap bitmapWithBackground = CreateCheckerboardBackground(image.Width, image.Height);
            using (Graphics g = Graphics.FromImage(bitmapWithBackground))
            {
                Bitmap imageBitmap = MatToBitmap(image);
                g.DrawImage(imageBitmap, 0, 0);
            }

            processedPictureBox.Image = bitmapWithBackground;
        }));
    }

    private Bitmap CreateCheckerboardBackground(int width, int height)
    {
        Bitmap background = new Bitmap(width, height);
        using (Graphics g = Graphics.FromImage(background))
        {
            int cellSize = 10;
            Color color1 = Color.LightGray;
            Color color2 = Color.White;
            for (int y = 0; y < height; y += cellSize)
            {
                for (int x = 0; x < width; x += cellSize)
                {
                    bool isWhite = ((x / cellSize) % 2 == (y / cellSize) % 2);
                    Brush brush = new SolidBrush(isWhite ? color1 : color2);
                    g.FillRectangle(brush, x, y, cellSize, cellSize);
                }
            }
        }
        return background;
    }


    private async void RemoveBackgroundButton_Click(object sender, EventArgs e)
    {
        progressBar.Style = ProgressBarStyle.Marquee;
        progressBar.MarqueeAnimationSpeed = 30;

        await System.Threading.Tasks.Task.Run(() =>
        {
            ApplyBackgroundRemoval();

            Invoke(new Action(() =>
            {
                progressBar.Style = ProgressBarStyle.Blocks;
                progressBar.MarqueeAnimationSpeed = 0;
                progressBar.Value = 100;
            }));
        });
    }

    private void ExportButton_Click(object sender, EventArgs e)
    {
        try
        {
            if (processedImage != null)
            {
                // 使用高分辨率保存图像
                Bitmap bitmap = MatToBitmap(processedImage);

                // 保存为PNG格式，确保高分辨率和清晰度
                tempImagePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.png");
                bitmap.Save(tempImagePath, ImageFormat.Png);

                // 获取原图的尺寸和位置
                float left = 0, top = 0, width = bitmap.Width, height = bitmap.Height;
                foreach (PowerPoint.Shape shape in currentSlide.Shapes)
                {
                    if (shape.Type == Office.MsoShapeType.msoPicture)
                    {
                        left = shape.Left;
                        top = shape.Top;
                        width = shape.Width;
                        height = shape.Height;
                        shape.Delete();
                        break;
                    }
                }

                // 将图片插入到原图位置并保持尺寸
                currentSlide.Shapes.AddPicture(tempImagePath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoCTrue, left, top, width, height);
                MessageBox.Show("图片已插入幻灯片", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                bitmap.Dispose();
            }
            else
            {
                MessageBox.Show("请先去除背景后再插入图片", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"导出图片时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }




    private void ResetButton_Click(object sender, EventArgs e)
    {
        if (processedPictureBox.Image != null)
        {
            processedPictureBox.Image.Dispose();
        }
        processedPictureBox.Image = originalPictureBox.Image;
        processedImage = originalImage.Clone();
        thresholdTrackBar.Value = 128;
        iterationsTrackBar.Value = 5;
        thresholdLabel.Text = "阈值: 128";
        iterationsLabel.Text = "迭代: 5";
        smoothnessTrackBar.Value = 0;
        smoothnessLabel.Text = "平滑度: 0";
        sharpnessTrackBar.Value = 0;
        sharpnessLabel.Text = "锐化度: 0";
    }

    private Mat BitmapToMat(Bitmap bitmap)
    {
        Mat mat = new Mat(bitmap.Height, bitmap.Width, MatType.CV_8UC3);
        for (int y = 0; y < bitmap.Height; y++)
        {
            for (int x = 0; x < bitmap.Width; x++)
            {
                Color color = bitmap.GetPixel(x, y);
                Vec3b pixel = new Vec3b(color.B, color.G, color.R);
                mat.Set(y, x, pixel);
            }
        }
        return mat;
    }

    private Bitmap MatToBitmap(Mat mat)
    {
        Bitmap bitmap = new Bitmap(mat.Width, mat.Height, PixelFormat.Format32bppArgb);
        var data = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.WriteOnly, bitmap.PixelFormat);
        using (var tempMat = new Mat(mat.Rows, mat.Cols, MatType.CV_8UC4, data.Scan0))
        {
            Cv2.CvtColor(mat, tempMat, ColorConversionCodes.BGR2BGRA);
        }
        bitmap.UnlockBits(data);
        return bitmap;
    }

    private Bitmap AddPaddingToImage(Image image, int padding)
    {
        int width = image.Width;
        int height = image.Height + padding;
        Bitmap paddedImage = new Bitmap(width, height);

        using (Graphics g = Graphics.FromImage(paddedImage))
        {
            g.Clear(Color.Transparent);
            g.DrawImage(image, 0, 0, width, image.Height);
        }

        return paddedImage;
    }

    private Mat RemovePaddingFromImage(Mat image, int padding)
    {
        return new Mat(image, new Rect(0, 0, image.Width, image.Height - padding));
    }

    private void BackgroundRemovalForm_FormClosed(object sender, FormClosedEventArgs e)
    {
        originalImage.Dispose();
        processedImage?.Dispose();
        processedPictureBox.Image?.Dispose();
        removeBackgroundButton.Dispose();
        exportButton.Dispose();
        smoothnessTrackBar.Dispose();
        smoothnessLabel.Dispose();
        sharpnessTrackBar.Dispose();
        sharpnessLabel.Dispose();
    }

    private void AutoAdjustParameters(Mat image)
    {
        Mat grayImage = new Mat();
        Cv2.CvtColor(image, grayImage, ColorConversionCodes.BGR2GRAY);

        int[] histSize = { 256 };
        Rangef[] ranges = { new Rangef(0, 256) };
        Mat hist = new Mat();
        Cv2.CalcHist(new Mat[] { grayImage }, new int[] { 0 }, null, hist, 1, histSize, ranges);

        double totalPixels = image.Rows * image.Cols;
        double sum = 0;
        int threshold = 0;
        for (int i = 0; i < hist.Rows; i++)
        {
            sum += hist.At<float>(i);
            if (sum / totalPixels > 0.5)
            {
                threshold = i;
                break;
            }
        }

        thresholdValue = threshold;
        iterations = Math.Min(10, Math.Max(1, (int)(sum / totalPixels * 10)));

        thresholdTrackBar.Value = thresholdValue;
        thresholdLabel.Text = "阈值: " + thresholdValue;

        iterationsTrackBar.Value = iterations;
        iterationsLabel.Text = "迭代: " + iterations;
    }
}
