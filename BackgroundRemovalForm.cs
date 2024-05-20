using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using OpenCvSharp;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

public class BackgroundRemovalForm : Form
{
    private PictureBox originalPictureBox;
    private PictureBox processedPictureBox;
    private Button removeBackgroundButton;
    private Button exportButton;
    private ProgressBar progressBar;
    private TrackBar thresholdTrackBar;
    private Label thresholdLabel;
    private TrackBar iterationsTrackBar;
    private Label iterationsLabel;
    private Mat originalImage;
    private string exportedImagePath;
    private PowerPoint.Application pptApp;
    private PowerPoint.Slide currentSlide;
    private string tempImagePath;
    private int thresholdValue = 128;  // 初始阈值
    private int iterations = 5;  // 初始迭代次数

    public BackgroundRemovalForm(Image image, PowerPoint.Application pptApp, PowerPoint.Slide currentSlide, string tempImagePath)
    {
        InitializeComponent();
        originalPictureBox.Image = image;
        originalImage = BitmapToMat((Bitmap)image);
        this.pptApp = pptApp;
        this.currentSlide = currentSlide;
        this.tempImagePath = tempImagePath;
    }

    public string ExportedImagePath => exportedImagePath;

    private void InitializeComponent()
    {
        this.originalPictureBox = new PictureBox();
        this.processedPictureBox = new PictureBox();
        this.removeBackgroundButton = new Button();
        this.exportButton = new Button();
        this.progressBar = new ProgressBar();
        this.thresholdTrackBar = new TrackBar();
        this.thresholdLabel = new Label();
        this.iterationsTrackBar = new TrackBar();
        this.iterationsLabel = new Label();

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
        this.removeBackgroundButton.Location = new System.Drawing.Point(12, 440);
        this.removeBackgroundButton.Size = new System.Drawing.Size(100, 40);
        this.removeBackgroundButton.Text = "去背景";
        this.removeBackgroundButton.Click += new System.EventHandler(this.RemoveBackgroundButton_Click);

        // exportButton
        this.exportButton.Location = new System.Drawing.Point(440, 440);
        this.exportButton.Size = new System.Drawing.Size(100, 40);
        this.exportButton.Text = "导出";
        this.exportButton.Click += new System.EventHandler(this.ExportButton_Click);

        // progressBar
        this.progressBar.Location = new System.Drawing.Point(130, 440);
        this.progressBar.Size = new System.Drawing.Size(160, 30);

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
        this.ClientSize = new System.Drawing.Size(564, 520);
        this.Controls.Add(this.originalPictureBox);
        this.Controls.Add(this.processedPictureBox);
        this.Controls.Add(this.removeBackgroundButton);
        this.Controls.Add(this.exportButton);
        this.Controls.Add(this.progressBar);
        this.Controls.Add(this.thresholdTrackBar);
        this.Controls.Add(this.thresholdLabel);
        this.Controls.Add(this.iterationsTrackBar);
        this.Controls.Add(this.iterationsLabel);
        this.Text = "便捷抠图Beta版本";
        this.FormClosed += new FormClosedEventHandler(this.BackgroundRemovalForm_FormClosed);
    }

    private void ThresholdTrackBar_Scroll(object sender, EventArgs e)
    {
        thresholdValue = thresholdTrackBar.Value;
        thresholdLabel.Text = "阈值: " + thresholdValue;
        ApplyBackgroundRemoval();
    }

    private void IterationsTrackBar_Scroll(object sender, EventArgs e)
    {
        iterations = iterationsTrackBar.Value;
        iterationsLabel.Text = "迭代: " + iterations;
        ApplyBackgroundRemoval();
    }

    private void ApplyBackgroundRemoval()
    {
        Rect rectangle = new Rect(10, 10, originalImage.Width - 20, originalImage.Height - 20);
        Mat mask = new Mat(originalImage.Size(), MatType.CV_8UC1, new Scalar(2));
        Mat bgdModel = new Mat();
        Mat fgdModel = new Mat();

        Cv2.GrabCut(originalImage, mask, rectangle, bgdModel, fgdModel, iterations, GrabCutModes.InitWithRect);

        mask = mask & 1;

        Mat foreground = new Mat(originalImage.Size(), MatType.CV_8UC4, new Scalar(0, 0, 0, 0));
        for (int y = 0; y < originalImage.Rows; y++)
        {
            for (int x = 0; x < originalImage.Cols; x++)
            {
                if (mask.At<byte>(y, x) == 1)
                {
                    Vec3b color = originalImage.At<Vec3b>(y, x);
                    foreground.Set(y, x, new Vec4b(color.Item0, color.Item1, color.Item2, 255));
                }
            }
        }

        processedPictureBox.Image = MatToBitmap(foreground);
    }

    private async void RemoveBackgroundButton_Click(object sender, EventArgs e)
    {
        progressBar.Style = ProgressBarStyle.Marquee;
        progressBar.MarqueeAnimationSpeed = 30;

        await System.Threading.Tasks.Task.Run(() =>
        {
            ApplyBackgroundRemoval();

            progressBar.Invoke(new Action(() =>
            {
                progressBar.Style = ProgressBarStyle.Blocks;
                progressBar.MarqueeAnimationSpeed = 0;
            }));
        });
    }

    private void ExportButton_Click(object sender, EventArgs e)
    {
        if (processedPictureBox.Image != null)
        {
            string tempImagePath = Path.Combine(Path.GetTempPath(), "processed_image_" + Guid.NewGuid().ToString() + ".png");
            processedPictureBox.Image.Save(tempImagePath, ImageFormat.Png);

            currentSlide.Shapes.AddPicture(tempImagePath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoCTrue, 0, 0, -1, -1);
            MessageBox.Show("图片已插入幻灯片", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        else
        {
            MessageBox.Show("请先去除背景后再插入图片", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private Mat BitmapToMat(Bitmap bitmap)
    {
        Mat mat = new Mat(bitmap.Height, bitmap.Width, MatType.CV_8UC3);
        for (int y = 0; y < bitmap.Height; y++)
        {
            for (int x = 0; x < bitmap.Width; x++)
            {
                Color color = bitmap.GetPixel(x, y);
                Vec3b pixel = new Vec3b(color.R, color.G, color.B);
                mat.Set<Vec3b>(y, x, pixel);
            }
        }
        Cv2.CvtColor(mat, mat, ColorConversionCodes.BGR2RGB);
        return mat;
    }

    private Bitmap MatToBitmap(Mat mat)
    {
        Bitmap bitmap = new Bitmap(mat.Width, mat.Height, PixelFormat.Format32bppArgb);
        var data = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.WriteOnly, bitmap.PixelFormat);
        using (var tempMat = new Mat(mat.Rows, mat.Cols, MatType.CV_8UC4, data.Scan0))
        {
            mat.CopyTo(tempMat);
        }
        bitmap.UnlockBits(data);
        return bitmap;
    }

    private void BackgroundRemovalForm_FormClosed(object sender, FormClosedEventArgs e)
    {
        // 确保资源被正确释放
        originalImage.Dispose();
        processedPictureBox.Image?.Dispose();
        removeBackgroundButton.Dispose();
        exportButton.Dispose();
    }
}
