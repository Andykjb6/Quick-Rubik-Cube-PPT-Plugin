using System;
using System.Drawing;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using 课件帮PPT助手;
using System.Runtime.InteropServices;


namespace 课件帮PPT助手 { }


public class SplitterTool
{
    private Form overlayForm;
    private Point startPoint;
    private Rectangle splitRectangle;
    private PowerPoint.Shape splitShape;
    private PowerPoint.Application pptApp;
    private float screenDpiX;
    private float screenDpiY;

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    private static extern bool ScreenToClient(IntPtr hWnd, ref Point lpPoint);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    public SplitterTool()
    {
        pptApp = Globals.ThisAddIn.Application;
        using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
        {
            screenDpiX = g.DpiX;
            screenDpiY = g.DpiY;
        }
    }

    public void StartSplitting()
    {
        CreateOverlayForm();
    }

    private void CreateOverlayForm()
    {
        overlayForm = new Form();
        overlayForm.FormBorderStyle = FormBorderStyle.None;
        overlayForm.BackColor = Color.White;
        overlayForm.Opacity = 0.2;
        overlayForm.WindowState = FormWindowState.Maximized;
        overlayForm.TopMost = true;
        overlayForm.ShowInTaskbar = false;

        overlayForm.MouseDown += new MouseEventHandler(OverlayForm_MouseDown);
        overlayForm.MouseMove += new MouseEventHandler(OverlayForm_MouseMove);
        overlayForm.MouseUp += new MouseEventHandler(OverlayForm_MouseUp);
        overlayForm.Paint += new PaintEventHandler(OverlayForm_Paint);

        overlayForm.Show();
    }

    private void OverlayForm_MouseDown(object sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            startPoint = new Point(e.X, e.Y);
        }
    }

    private void OverlayForm_MouseMove(object sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            Point endPoint = new Point(e.X, e.Y);
            splitRectangle = new Rectangle(
                Math.Min(startPoint.X, endPoint.X),
                Math.Min(startPoint.Y, endPoint.Y),
                Math.Abs(startPoint.X - endPoint.X),
                Math.Abs(startPoint.Y - endPoint.Y));

            overlayForm.Invalidate();
        }
    }

    private void OverlayForm_MouseUp(object sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            overlayForm.Close();
            InsertSplitRectangle();
        }
    }

    private void OverlayForm_Paint(object sender, PaintEventArgs e)
    {
        if (splitRectangle != Rectangle.Empty)
        {
            e.Graphics.DrawRectangle(Pens.Red, splitRectangle);
        }
    }

    private void InsertSplitRectangle()
    {
        PowerPoint.Slide slide = pptApp.ActiveWindow.View.Slide;
        float zoom = pptApp.ActiveWindow.View.Zoom / 100f;
        IntPtr pptHandle = new IntPtr(pptApp.ActiveWindow.HWND);

        GetWindowRect(pptHandle, out RECT pptWindowRect);

        Point clientStartPoint = new Point(splitRectangle.Left, splitRectangle.Top);
        ScreenToClient(pptHandle, ref clientStartPoint);

        clientStartPoint.X -= pptWindowRect.Left;
        clientStartPoint.Y -= pptWindowRect.Top;

        // 使用参考值并根据比例调整
        int referenceTitleBarHeight = -670;
        int referenceBorderWidth = 309;
        float baseZoom = 0.67f; // 67%

        int titleBarHeight = (int)(referenceTitleBarHeight * (zoom / baseZoom));
        int borderWidth = (int)(referenceBorderWidth * (zoom / baseZoom));

        // 调整坐标
        clientStartPoint.X -= borderWidth;
        clientStartPoint.Y -= titleBarHeight;

        // 动态调整微调参数
        int baseFineTuneX = 37; // 基准缩放比例(67%)下的X微调参数
        int baseFineTuneY = -310; // 基准缩放比例(67%)下的Y微调参数

        int fineTuneX = (int)(baseFineTuneX * (zoom / baseZoom));
        int fineTuneY = (int)(baseFineTuneY * (zoom / baseZoom));

        clientStartPoint.X += fineTuneX;
        clientStartPoint.Y += fineTuneY;

        // 转换像素到点并调整缩放比例
        float left = ConvertPixelsToPoints(clientStartPoint.X, screenDpiX) / zoom;
        float top = ConvertPixelsToPoints(clientStartPoint.Y, screenDpiY) / zoom;
        float width = ConvertPixelsToPoints(splitRectangle.Width, screenDpiX) / zoom;
        float height = ConvertPixelsToPoints(splitRectangle.Height, screenDpiY) / zoom;

        if (width <= 0 || height <= 0)
        {
            MessageBox.Show("插入矩形的尺寸无效，请重新选择", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            throw new InvalidOperationException("插入矩形的尺寸无效");
        }

        PowerPoint.Shape selectedShape = pptApp.ActiveWindow.Selection.ShapeRange[1];
        splitShape = slide.Shapes.AddShape(
            Office.MsoAutoShapeType.msoShapeRectangle,
            left, top, width, height);

        CopyShapeProperties(selectedShape, splitShape);
        splitShape.Fill.Transparency = selectedShape.Fill.Transparency;

        FragmentShapes(slide, selectedShape);
    }

    private void CopyShapeProperties(PowerPoint.Shape sourceShape, PowerPoint.Shape targetShape)
    {
        targetShape.Fill.ForeColor.RGB = sourceShape.Fill.ForeColor.RGB;
        targetShape.Line.ForeColor.RGB = sourceShape.Line.ForeColor.RGB;

        float lineWeight = sourceShape.Line.Weight;
        if (lineWeight > 0 && lineWeight < 100)
        {
            targetShape.Line.Weight = lineWeight;
        }
        else
        {
            targetShape.Line.Weight = 1;
        }

        targetShape.TextFrame.TextRange.Text = sourceShape.TextFrame.TextRange.Text;
        targetShape.TextFrame.TextRange.Font.Name = sourceShape.TextFrame.TextRange.Font.Name;
        targetShape.TextFrame.TextRange.Font.Size = sourceShape.TextFrame.TextRange.Font.Size;
        targetShape.TextFrame.TextRange.Font.Color.RGB = sourceShape.TextFrame.TextRange.Font.Color.RGB;
    }

    private void FragmentShapes(PowerPoint.Slide slide, PowerPoint.Shape selectedShape)
    {
        PowerPoint.ShapeRange shapeRange = slide.Shapes.Range(new[] { splitShape.Name, selectedShape.Name });
        shapeRange.MergeShapes(Office.MsoMergeCmd.msoMergeFragment);

        foreach (PowerPoint.Shape shape in shapeRange)
        {
            try
            {
                if (shape.Line.ForeColor.RGB == Color.White.ToArgb())
                {
                    shape.Line.Visible = Office.MsoTriState.msoFalse;
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // 忽略无法访问的属性
            }
        }

        MessageBox.Show("形状已成功拆分", "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private float ConvertPixelsToPoints(float pixels, float dpi)
    {
        return pixels * 72f / dpi;
    }
}


