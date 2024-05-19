using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

// 测试
namespace 课件帮PPT助手
{
    public class CustomCloudTextGenerator
    {
        private PowerPoint.Shape textBoxShape3; // 添加字段
        private IntPtr pptHandle; // PowerPoint 窗口句柄
        private CheckBox shadowCheckBox; // 添加文本阴影开关

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll")]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        private const int GWL_HWNDPARENT = -8;
        private const int HWND_TOPMOST = -1;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOACTIVATE = 0x0010;

        public void InitializeForm()
        {
            var pptApplication = Globals.ThisAddIn.Application;
            pptHandle = (IntPtr)pptApplication.HWND; // 获取 PowerPoint 窗口句柄

            var form = new Form();
            form.Text = "自定义云朵字生成";
            form.Size = new System.Drawing.Size(600, 900); // 调整窗口大小
            form.FormBorderStyle = FormBorderStyle.SizableToolWindow; // 设置边框样式为可调整大小的工具窗口
            form.ShowInTaskbar = false; // 不在任务栏显示窗口

            // 添加文本框，确保文本不会自动换行
            var textBox = new TextBox();
            textBox.Location = new System.Drawing.Point(150, 50); // 调整文本框位置
            textBox.Size = new System.Drawing.Size(350, 30); // 增加文本框宽度
            textBox.Font = new System.Drawing.Font(textBox.Font.FontFamily, 10); // 设置文本框字号
            textBox.TextAlign = HorizontalAlignment.Center; // 文本居中对齐
            form.Controls.Add(textBox);

            shadowCheckBox = new CheckBox();
            shadowCheckBox.Text = "阴影开关（生成后可调）";
            shadowCheckBox.Location = new System.Drawing.Point(50, 750);
            shadowCheckBox.Size = new System.Drawing.Size(500, 40);

            shadowCheckBox.CheckedChanged += (sender, e) =>
            {
                if (shadowCheckBox.Checked)
                {
                    // 关闭文本阴影
                    textBoxShape3.TextFrame2.TextRange.Font.Shadow.Visible = Office.MsoTriState.msoFalse;
                }
                else
                {
                    // 开启文本阴影
                    textBoxShape3.TextFrame2.TextRange.Font.Shadow.Visible = Office.MsoTriState.msoTrue;
                }
            };
            form.Controls.Add(shadowCheckBox);

            Label fontLabel = new Label();
            fontLabel.Text = "选择字体：";
            fontLabel.Location = new System.Drawing.Point(50, 100); // 调整位置
            fontLabel.Size = new System.Drawing.Size(150, 30); // 增加标签宽度
            form.Controls.Add(fontLabel);

            ComboBox fontComboBox = new ComboBox();
            fontComboBox.Location = new System.Drawing.Point(200, 100);
            fontComboBox.Size = new System.Drawing.Size(300, 30); // 增加宽度
            foreach (System.Drawing.FontFamily fontFamily in System.Drawing.FontFamily.Families)
            {
                fontComboBox.Items.Add(fontFamily.Name);
            }
            fontComboBox.SelectedIndex = 0;
            form.Controls.Add(fontComboBox);

            Label topColorLabel = new Label();
            topColorLabel.Text = "顶层颜色：";
            topColorLabel.Location = new System.Drawing.Point(50, 150); // 调整位置
            topColorLabel.Size = new System.Drawing.Size(150, 40); // 增加标签宽度
            form.Controls.Add(topColorLabel);

            Button topColorButton = new Button();
            topColorButton.Text = "自定义";
            topColorButton.Location = new System.Drawing.Point(200, 150);
            topColorButton.Size = new System.Drawing.Size(100, 40);
            topColorButton.BackColor = System.Drawing.Color.Black;
            topColorButton.Click += (s, args) =>
            {
                ColorDialog colorDialog = new ColorDialog();
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    topColorButton.BackColor = colorDialog.Color;
                }
            };
            form.Controls.Add(topColorButton);

            Label middleColorLabel = new Label(); // 添加中间层颜色标签
            middleColorLabel.Text = "中层颜色：";
            middleColorLabel.Location = new System.Drawing.Point(50, 200); // 调整位置
            middleColorLabel.Size = new System.Drawing.Size(150, 30); // 增加标签宽度
            form.Controls.Add(middleColorLabel);

            Button middleColorButton = new Button(); // 添加中间层颜色按钮
            middleColorButton.Text = "自定义";
            middleColorButton.Location = new System.Drawing.Point(200, 200);
            middleColorButton.Size = new System.Drawing.Size(100, 40);
            middleColorButton.BackColor = System.Drawing.Color.White;
            middleColorButton.Click += (s, args) =>
            {
                ColorDialog colorDialog = new ColorDialog();
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    middleColorButton.BackColor = colorDialog.Color;
                }
            };
            form.Controls.Add(middleColorButton);

            Label bottomColorLabel = new Label(); // 添加最底层颜色标签
            bottomColorLabel.Text = "底层颜色：";
            bottomColorLabel.Location = new System.Drawing.Point(50, 250); // 调整位置
            bottomColorLabel.Size = new System.Drawing.Size(150, 30); // 增加标签宽度
            form.Controls.Add(bottomColorLabel);

            Button bottomColorButton = new Button(); // 添加最底层颜色按钮
            bottomColorButton.Text = "自定义";
            bottomColorButton.Location = new System.Drawing.Point(200, 250);
            bottomColorButton.Size = new System.Drawing.Size(100, 40);
            bottomColorButton.BackColor = System.Drawing.Color.Blue;
            bottomColorButton.Click += (s, args) =>
            {
                ColorDialog colorDialog = new ColorDialog();
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    bottomColorButton.BackColor = colorDialog.Color;
                    // 输出选择的颜色和调整后的颜色
                    int selectedColor = colorDialog.Color.ToArgb();
                    int adjustedColor = SwapRedBlue(selectedColor);
                    Debug.WriteLine($"Selected Color ARGB: {selectedColor:X}");
                    Debug.WriteLine($"Adjusted Color ARGB: {adjustedColor:X}");
                }
            };
            form.Controls.Add(bottomColorButton);

            Label middleOutlineLabel = new Label(); // 添加中间层文字轮廓宽度标签
            middleOutlineLabel.Text = "中层轮廓："; // 修正文字
            middleOutlineLabel.Location = new System.Drawing.Point(50, 300); // 调整位置
            middleOutlineLabel.Size = new System.Drawing.Size(150, 30); // 增加标签宽度
            form.Controls.Add(middleOutlineLabel);

            NumericUpDown middleOutlineNumericUpDown = new NumericUpDown(); // 添加中间层文字轮廓宽度控件
            middleOutlineNumericUpDown.Location = new System.Drawing.Point(200, 300);
            middleOutlineNumericUpDown.Size = new System.Drawing.Size(100, 30); // 增加宽度
            middleOutlineNumericUpDown.Minimum = 0;
            middleOutlineNumericUpDown.Maximum = 100;
            middleOutlineNumericUpDown.Value = 6;
            form.Controls.Add(middleOutlineNumericUpDown);

            Label bottomOutlineLabel = new Label();
            bottomOutlineLabel.Text = "底层轮廓：";
            bottomOutlineLabel.Location = new System.Drawing.Point(50, 350); // 调整位置
            bottomOutlineLabel.Size = new System.Drawing.Size(150, 30); // 增加标签宽度
            form.Controls.Add(bottomOutlineLabel);

            NumericUpDown bottomOutlineNumericUpDown = new NumericUpDown();
            bottomOutlineNumericUpDown.Location = new System.Drawing.Point(200, 350);
            bottomOutlineNumericUpDown.Size = new System.Drawing.Size(100, 30); // 增加宽度
            bottomOutlineNumericUpDown.Minimum = 0;
            bottomOutlineNumericUpDown.Maximum = 100;
            bottomOutlineNumericUpDown.Value = 12;
            form.Controls.Add(bottomOutlineNumericUpDown);

            Button shadowColorButton = new Button();
            shadowColorButton.Text = "更改阴影颜色";
            shadowColorButton.Location = new System.Drawing.Point(50, 400);
            shadowColorButton.Size = new System.Drawing.Size(200, 40);
            shadowColorButton.Click += (s, args) =>
            {
                ColorDialog colorDialog = new ColorDialog();
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    int argbColor = colorDialog.Color.ToArgb();
                    int correctedArgbColor = SwapRedBlue(argbColor); // 正确调整颜色值
                    textBoxShape3.TextFrame2.TextRange.Font.Shadow.ForeColor.RGB = correctedArgbColor;
                }
            };
            form.Controls.Add(shadowColorButton);

            // 添加阴影模糊度标签和滑块控件
            Label shadowBlurLabel = new Label();
            shadowBlurLabel.Text = "阴影模糊（生成后可调）：";
            shadowBlurLabel.Location = new System.Drawing.Point(50, 470); // 调整位置
            shadowBlurLabel.Size = new System.Drawing.Size(500, 40); // 调整大小
            form.Controls.Add(shadowBlurLabel);

            TrackBar shadowBlurTrackBar = new TrackBar();
            shadowBlurTrackBar.Minimum = 0;
            shadowBlurTrackBar.Maximum = 100;
            shadowBlurTrackBar.Value = 25; // 设置默认值
            shadowBlurTrackBar.Location = new System.Drawing.Point(50, 520); // 调整位置，增加垂直间距
            shadowBlurTrackBar.Size = new System.Drawing.Size(200, 40);
            form.Controls.Add(shadowBlurTrackBar);

            // 添加阴影透明度标签和滑块控件
            Label shadowTransparencyLabel = new Label();
            shadowTransparencyLabel.Text = "阴影透明（生成后可调）：";
            shadowTransparencyLabel.Location = new System.Drawing.Point(50, 620); // 调整位置，增加垂直间距
            shadowTransparencyLabel.Size = new System.Drawing.Size(500, 40); // 调整大小
            form.Controls.Add(shadowTransparencyLabel);

            TrackBar shadowTransparencyTrackBar = new TrackBar();
            shadowTransparencyTrackBar.Minimum = 0;
            shadowTransparencyTrackBar.Maximum = 100;
            shadowTransparencyTrackBar.Value = 65; // 设置默认值
            shadowTransparencyTrackBar.Location = new System.Drawing.Point(50, 670);
            shadowTransparencyTrackBar.Size = new System.Drawing.Size(200, 40);
            form.Controls.Add(shadowTransparencyTrackBar);

            // 在滑块的 Scroll 事件处理程序中更新底层文字阴影的模糊度和透明度参数
            shadowBlurTrackBar.Scroll += (s, args) =>
            {
                int blurValue = shadowBlurTrackBar.Value;
                textBoxShape3.TextFrame2.TextRange.Font.Shadow.Blur = blurValue;
            };

            shadowTransparencyTrackBar.Scroll += (s, args) =>
            {
                float transparencyValue = shadowTransparencyTrackBar.Value / 100f; // 将值转换为百分比
                textBoxShape3.TextFrame2.TextRange.Font.Shadow.Transparency = transparencyValue;
            };

            Button generateButton = new Button();
            generateButton.Text = "生成";
            generateButton.Location = new System.Drawing.Point(250, 400); // 调整按钮位置
            generateButton.Size = new System.Drawing.Size(100, 40);
            generateButton.Click += (s, args) =>
            {
                PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
                PowerPoint.Slide currentSlide = pptApp.ActiveWindow.View.Slide;

                // 创建一个文本框并设置样式
                PowerPoint.Shape textBoxShape = currentSlide.Shapes.AddTextbox(
                    Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                    100, 100, 200, 100);
                textBoxShape.TextFrame.TextRange.Text = "Your text here";
                textBoxShape.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText; // 自适应文本长度
                textBoxShape.TextFrame.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse; // 禁用文字包裹，确保文本始终在一行显示

                // 获取文本的宽度和高度
                float textWidth = textBoxShape.TextFrame2.TextRange.BoundWidth;
                float textHeight = textBoxShape.TextFrame2.TextRange.BoundHeight;

                // 调整文本框的大小以适应文本内容
                textBoxShape.Width = textWidth + 10; // 添加一些额外空间，确保文本不被截断
                textBoxShape.Height = textHeight + 10; // 添加一些额外空间，确保文本不被截断
                textBoxShape.TextFrame.TextRange.Text = textBox.Text;
                textBoxShape.TextFrame.TextRange.Font.NameFarEast = fontComboBox.SelectedItem.ToString(); // 设置中文字体
                textBoxShape.TextFrame.TextRange.Font.Name = fontComboBox.SelectedItem.ToString(); // 设置字体
                textBoxShape.TextFrame2.TextRange.Font.Size = 130; // 设置字号为130
                textBoxShape.TextFrame2.TextRange.ParagraphFormat.Alignment = Office.MsoParagraphAlignment.msoAlignCenter; // 文本居中对齐

                // 设置最顶层文本框的样式
                int topColor = SwapRedBlue(topColorButton.BackColor.ToArgb());
                textBoxShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = topColor;
                textBoxShape.TextFrame2.TextRange.Font.Shadow.Visible = Office.MsoTriState.msoFalse;
                textBoxShape.TextFrame2.TextRange.Font.Line.Visible = Office.MsoTriState.msoFalse;

                // 输出调试信息
                Debug.WriteLine($"Top Color: {topColor:X}");

                // 复制文本框并调整位置
                PowerPoint.Shape textBoxShape2 = textBoxShape.Duplicate()[1];
                textBoxShape2.Left = textBoxShape.Left;
                textBoxShape2.Top = textBoxShape.Top;
                textBoxShape3 = textBoxShape.Duplicate()[1]; // 修改此处，使用类的字段而不是创建新的局部变量
                textBoxShape3.Left = textBoxShape.Left;
                textBoxShape3.Top = textBoxShape.Top;

                // 设置中间层文本框的样式
                int middleColor = SwapRedBlue(middleColorButton.BackColor.ToArgb());
                textBoxShape2.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = middleColor;
                textBoxShape2.TextFrame2.TextRange.Font.Shadow.Visible = Office.MsoTriState.msoFalse;
                textBoxShape2.TextFrame2.TextRange.Font.Line.ForeColor.RGB = middleColor;
                textBoxShape2.TextFrame2.TextRange.Font.Line.Visible = Office.MsoTriState.msoTrue;
                textBoxShape2.TextFrame2.TextRange.Font.Line.Weight = Convert.ToSingle(middleOutlineNumericUpDown.Value);

                // 输出调试信息
                Debug.WriteLine($"Middle Color: {middleColor:X}");

                // 设置最底层文本框的样式
                int bottomColor = SwapRedBlue(bottomColorButton.BackColor.ToArgb());
                textBoxShape3.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = bottomColor;
                textBoxShape3.TextFrame2.TextRange.Font.Shadow.Visible = Office.MsoTriState.msoTrue;
                textBoxShape3.TextFrame2.TextRange.Font.Line.ForeColor.RGB = bottomColor;
                textBoxShape3.TextFrame2.TextRange.Font.Line.Visible = Office.MsoTriState.msoTrue;
                textBoxShape3.TextFrame2.TextRange.Font.Line.Weight = Convert.ToSingle(bottomOutlineNumericUpDown.Value);
                // 设置最底层文本框的阴影模糊度和透明度
                textBoxShape3.TextFrame2.TextRange.Font.Shadow.Blur = 15;
                textBoxShape3.TextFrame2.TextRange.Font.Shadow.Transparency = 0.50f;

                // 输出调试信息
                Debug.WriteLine($"Bottom Color: {bottomColor:X}");

                // 将文字轮廓颜色设置为与填充颜色相同
                textBoxShape2.TextFrame2.TextRange.Font.Line.ForeColor.RGB = middleColor;
                textBoxShape3.TextFrame2.TextRange.Font.Line.ForeColor.RGB = bottomColor;

                // 设置文本框的层级顺序
                textBoxShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                textBoxShape2.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
                textBoxShape3.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
            };
            form.Controls.Add(generateButton);

            // 将窗口设置为 PowerPoint 的子窗口，使其始终悬停在 PowerPoint 上
            SetWindowLong(form.Handle, GWL_HWNDPARENT, pptHandle.ToInt32());

            // 将窗口设置为顶级窗口，并且不会被激活
            SetWindowPos(form.Handle, (IntPtr)HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

            // 注册鼠标移动事件
            form.MouseMove += (s, args) =>
            {
                // 检查鼠标是否在窗口内
                if (args.Button == MouseButtons.None && args.X >= 0 && args.Y >= 0 && args.X < form.Width && args.Y < form.Height)
                {
                    // 当鼠标在窗口内移动时，不改变窗口位置
                    return;
                }

                // 当鼠标在窗口外移动时，将窗口移动到 PowerPoint 主窗口附近
                var left = (int)pptApplication.Left;
                var top = (int)pptApplication.Top;
                SetWindowPos(form.Handle, (IntPtr)HWND_TOPMOST, left + 20, top + 20, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
            };

            // 显示窗口，但不阻止其他操作
            form.Show();

            // 将窗口置于最前面，以防止被 PowerPoint 窗口遮挡
            form.TopMost = true;
        }

        // 辅助方法：交换红色和蓝色值
        private int SwapRedBlue(int argbColor)
        {
            byte[] colorBytes = BitConverter.GetBytes(argbColor);
            byte alpha = colorBytes[3];
            byte red = colorBytes[2];
            byte green = colorBytes[1];
            byte blue = colorBytes[0];

            return (alpha << 24) | (blue << 16) | (green << 8) | red;
        }
    }
}
