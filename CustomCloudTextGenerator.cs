using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace 课件帮PPT助手
{
    public class CustomCloudTextGenerator
    {
        private PowerPoint.Shape textBoxShape;
        private PowerPoint.Shape textBoxShape2;
        private PowerPoint.Shape textBoxShape3;
        private IntPtr pptHandle;
        private CheckBox shadowCheckBox;

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
            pptHandle = (IntPtr)pptApplication.HWND;

            var form = new Form
            {
                Text = "自定义云朵字生成",
                Size = new System.Drawing.Size(600, 900),
                FormBorderStyle = FormBorderStyle.SizableToolWindow,
                ShowInTaskbar = false
            };

            var textBox = new TextBox
            {
                Location = new System.Drawing.Point(150, 50),
                Size = new System.Drawing.Size(350, 30),
                Font = new System.Drawing.Font(FontFamily.GenericSansSerif, 10, FontStyle.Regular),
                TextAlign = HorizontalAlignment.Center
            };
            form.Controls.Add(textBox);

            shadowCheckBox = new CheckBox
            {
                Text = "阴影开关（生成后可调）",
                Location = new System.Drawing.Point(50, 750),
                Size = new System.Drawing.Size(500, 40)
            };
            shadowCheckBox.CheckedChanged += (sender, e) =>
            {
                textBoxShape3.TextFrame2.TextRange.Font.Shadow.Visible = shadowCheckBox.Checked ? Office.MsoTriState.msoFalse : Office.MsoTriState.msoTrue;
            };
            form.Controls.Add(shadowCheckBox);

            var fontLabel = new Label
            {
                Text = "选择字体：",
                Location = new System.Drawing.Point(50, 100),
                Size = new System.Drawing.Size(150, 30)
            };
            form.Controls.Add(fontLabel);

            var fontComboBox = new ComboBox
            {
                Location = new System.Drawing.Point(200, 100),
                Size = new System.Drawing.Size(300, 30)
            };
            foreach (System.Drawing.FontFamily fontFamily in System.Drawing.FontFamily.Families)
            {
                fontComboBox.Items.Add(fontFamily.Name);
            }
            fontComboBox.SelectedIndex = 0;
            fontComboBox.SelectedIndexChanged += (s, e) =>
            {
                UpdateFont(fontComboBox.SelectedItem.ToString());
            };
            form.Controls.Add(fontComboBox);

            var topColorLabel = new Label
            {
                Text = "顶层颜色：",
                Location = new System.Drawing.Point(50, 150),
                Size = new System.Drawing.Size(150, 40)
            };
            form.Controls.Add(topColorLabel);

            var topColorButton = new Button
            {
                Text = "自定义",
                Location = new System.Drawing.Point(200, 150),
                Size = new System.Drawing.Size(100, 40),
                BackColor = System.Drawing.Color.Black
            };
            topColorButton.Click += (s, args) =>
            {
                ColorDialog colorDialog = new ColorDialog();
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    topColorButton.BackColor = colorDialog.Color;
                    UpdateTopColor(SwapRedBlue(colorDialog.Color.ToArgb()));
                }
            };
            form.Controls.Add(topColorButton);

            var middleColorLabel = new Label
            {
                Text = "中层颜色：",
                Location = new System.Drawing.Point(50, 200),
                Size = new System.Drawing.Size(150, 30)
            };
            form.Controls.Add(middleColorLabel);

            var middleColorButton = new Button
            {
                Text = "自定义",
                Location = new System.Drawing.Point(200, 200),
                Size = new System.Drawing.Size(100, 40),
                BackColor = System.Drawing.Color.White
            };
            middleColorButton.Click += (s, args) =>
            {
                ColorDialog colorDialog = new ColorDialog();
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    middleColorButton.BackColor = colorDialog.Color;
                    UpdateMiddleColor(SwapRedBlue(colorDialog.Color.ToArgb()));
                }
            };
            form.Controls.Add(middleColorButton);

            var bottomColorLabel = new Label
            {
                Text = "底层颜色：",
                Location = new System.Drawing.Point(50, 250),
                Size = new System.Drawing.Size(150, 30)
            };
            form.Controls.Add(bottomColorLabel);

            var bottomColorButton = new Button
            {
                Text = "自定义",
                Location = new System.Drawing.Point(200, 250),
                Size = new System.Drawing.Size(100, 40),
                BackColor = System.Drawing.Color.Blue
            };
            bottomColorButton.Click += (s, args) =>
            {
                ColorDialog colorDialog = new ColorDialog();
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    bottomColorButton.BackColor = colorDialog.Color;
                    UpdateBottomColor(SwapRedBlue(colorDialog.Color.ToArgb()));
                }
            };
            form.Controls.Add(bottomColorButton);

            var middleOutlineLabel = new Label
            {
                Text = "中层轮廓：",
                Location = new System.Drawing.Point(50, 300),
                Size = new System.Drawing.Size(150, 30)
            };
            form.Controls.Add(middleOutlineLabel);

            var middleOutlineNumericUpDown = new NumericUpDown
            {
                Location = new System.Drawing.Point(200, 300),
                Size = new System.Drawing.Size(100, 30),
                Minimum = 0,
                Maximum = 100,
                Value = 45
            };
            middleOutlineNumericUpDown.ValueChanged += (s, e) =>
            {
                UpdateMiddleOutline(Convert.ToSingle(middleOutlineNumericUpDown.Value));
            };
            form.Controls.Add(middleOutlineNumericUpDown);

            var bottomOutlineLabel = new Label
            {
                Text = "底层轮廓：",
                Location = new System.Drawing.Point(50, 350),
                Size = new System.Drawing.Size(150, 30)
            };
            form.Controls.Add(bottomOutlineLabel);

            var bottomOutlineNumericUpDown = new NumericUpDown
            {
                Location = new System.Drawing.Point(200, 350),
                Size = new System.Drawing.Size(100, 30),
                Minimum = 0,
                Maximum = 100,
                Value = 55
            };
            bottomOutlineNumericUpDown.ValueChanged += (s, e) =>
            {
                UpdateBottomOutline(Convert.ToSingle(bottomOutlineNumericUpDown.Value));
            };
            form.Controls.Add(bottomOutlineNumericUpDown);

            var shadowColorButton = new Button
            {
                Text = "更改阴影颜色",
                Location = new System.Drawing.Point(50, 400),
                Size = new System.Drawing.Size(200, 40)
            };
            shadowColorButton.Click += (s, args) =>
            {
                ColorDialog colorDialog = new ColorDialog();
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    int argbColor = colorDialog.Color.ToArgb();
                    int correctedArgbColor = SwapRedBlue(argbColor);
                    textBoxShape3.TextFrame2.TextRange.Font.Shadow.ForeColor.RGB = correctedArgbColor;
                }
            };
            form.Controls.Add(shadowColorButton);

            var shadowBlurLabel = new Label
            {
                Text = "阴影模糊（生成后可调）：",
                Location = new System.Drawing.Point(50, 470),
                Size = new System.Drawing.Size(500, 40)
            };
            form.Controls.Add(shadowBlurLabel);

            var shadowBlurTrackBar = new TrackBar
            {
                Minimum = 0,
                Maximum = 100,
                Value = 25,
                Location = new System.Drawing.Point(50, 520),
                Size = new System.Drawing.Size(200, 40)
            };
            shadowBlurTrackBar.Scroll += (s, args) =>
            {
                int blurValue = shadowBlurTrackBar.Value;
                textBoxShape3.TextFrame2.TextRange.Font.Shadow.Blur = blurValue;
            };
            form.Controls.Add(shadowBlurTrackBar);

            var shadowTransparencyLabel = new Label
            {
                Text = "阴影透明（生成后可调）：",
                Location = new System.Drawing.Point(50, 620),
                Size = new System.Drawing.Size(500, 40)
            };
            form.Controls.Add(shadowTransparencyLabel);

            var shadowTransparencyTrackBar = new TrackBar
            {
                Minimum = 0,
                Maximum = 100,
                Value = 65,
                Location = new System.Drawing.Point(50, 670),
                Size = new System.Drawing.Size(200, 40)
            };
            shadowTransparencyTrackBar.Scroll += (s, args) =>
            {
                float transparencyValue = shadowTransparencyTrackBar.Value / 100f;
                textBoxShape3.TextFrame2.TextRange.Font.Shadow.Transparency = transparencyValue;
            };
            form.Controls.Add(shadowTransparencyTrackBar);

            var generateButton = new Button
            {
                Text = "生成",
                Location = new System.Drawing.Point(250, 400),
                Size = new System.Drawing.Size(100, 40)
            };
            generateButton.Click += (s, args) =>
            {
                PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
                PowerPoint.Slide currentSlide = pptApp.ActiveWindow.View.Slide;

                textBoxShape = currentSlide.Shapes.AddTextbox(
                    Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                    100, 100, 200, 100);
                textBoxShape.TextFrame.TextRange.Text = textBox.Text;
                textBoxShape.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                textBoxShape.TextFrame.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;

                float textWidth = textBoxShape.TextFrame2.TextRange.BoundWidth;
                float textHeight = textBoxShape.TextFrame2.TextRange.BoundHeight;

                textBoxShape.Width = textWidth + 10;
                textBoxShape.Height = textHeight + 10;
                textBoxShape.TextFrame.TextRange.Font.NameFarEast = fontComboBox.SelectedItem.ToString();
                textBoxShape.TextFrame.TextRange.Font.Name = fontComboBox.SelectedItem.ToString();
                textBoxShape.TextFrame2.TextRange.Font.Size = 130;
                textBoxShape.TextFrame2.TextRange.ParagraphFormat.Alignment = Office.MsoParagraphAlignment.msoAlignCenter;

                int topColor = SwapRedBlue(topColorButton.BackColor.ToArgb());
                textBoxShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = topColor;
                textBoxShape.TextFrame2.TextRange.Font.Shadow.Visible = Office.MsoTriState.msoFalse;
                textBoxShape.TextFrame2.TextRange.Font.Line.Visible = Office.MsoTriState.msoFalse;

                Debug.WriteLine($"Top Color: {topColor:X}");

                textBoxShape2 = textBoxShape.Duplicate()[1];
                textBoxShape2.Left = textBoxShape.Left;
                textBoxShape2.Top = textBoxShape.Top;
                textBoxShape3 = textBoxShape.Duplicate()[1];
                textBoxShape3.Left = textBoxShape.Left;
                textBoxShape3.Top = textBoxShape.Top;

                int middleColor = SwapRedBlue(middleColorButton.BackColor.ToArgb());
                textBoxShape2.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = middleColor;
                textBoxShape2.TextFrame2.TextRange.Font.Shadow.Visible = Office.MsoTriState.msoFalse;
                textBoxShape2.TextFrame2.TextRange.Font.Line.ForeColor.RGB = middleColor;
                textBoxShape2.TextFrame2.TextRange.Font.Line.Visible = Office.MsoTriState.msoTrue;
                textBoxShape2.TextFrame2.TextRange.Font.Line.Weight = Convert.ToSingle(middleOutlineNumericUpDown.Value);

                Debug.WriteLine($"Middle Color: {middleColor:X}");

                int bottomColor = SwapRedBlue(bottomColorButton.BackColor.ToArgb());
                textBoxShape3.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = bottomColor;
                textBoxShape3.TextFrame2.TextRange.Font.Shadow.Visible = Office.MsoTriState.msoTrue;
                textBoxShape3.TextFrame2.TextRange.Font.Line.ForeColor.RGB = bottomColor;
                textBoxShape3.TextFrame2.TextRange.Font.Line.Visible = Office.MsoTriState.msoTrue;
                textBoxShape3.TextFrame2.TextRange.Font.Line.Weight = Convert.ToSingle(bottomOutlineNumericUpDown.Value);
                textBoxShape3.TextFrame2.TextRange.Font.Shadow.Blur = 15;
                textBoxShape3.TextFrame2.TextRange.Font.Shadow.Transparency = 0.50f;

                Debug.WriteLine($"Bottom Color: {bottomColor:X}");

                textBoxShape2.TextFrame2.TextRange.Font.Line.ForeColor.RGB = middleColor;
                textBoxShape3.TextFrame2.TextRange.Font.Line.ForeColor.RGB = bottomColor;

                textBoxShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                textBoxShape2.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
                textBoxShape3.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
            };
            form.Controls.Add(generateButton);

            SetWindowLong(form.Handle, GWL_HWNDPARENT, pptHandle.ToInt32());
            SetWindowPos(form.Handle, (IntPtr)HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

            form.MouseMove += (s, args) =>
            {
                if (args.Button == MouseButtons.None && args.X >= 0 && args.Y >= 0 && args.X < form.Width && args.Y < form.Height)
                {
                    return;
                }

                var left = (int)pptApplication.Left;
                var top = (int)pptApplication.Top;
                SetWindowPos(form.Handle, (IntPtr)HWND_TOPMOST, left + 20, top + 20, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
            };

            form.Show();
            form.TopMost = true;
        }

        private void UpdateFont(string fontName)
        {
            if (textBoxShape != null)
            {
                textBoxShape.TextFrame2.TextRange.Font.NameFarEast = fontName;
                textBoxShape.TextFrame2.TextRange.Font.Name = fontName;
                textBoxShape2.TextFrame2.TextRange.Font.NameFarEast = fontName;
                textBoxShape2.TextFrame2.TextRange.Font.Name = fontName;
                textBoxShape3.TextFrame2.TextRange.Font.NameFarEast = fontName;
                textBoxShape3.TextFrame2.TextRange.Font.Name = fontName;
            }
        }

        private void UpdateTopColor(int color)
        {
            if (textBoxShape != null)
            {
                textBoxShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = color;
            }
        }

        private void UpdateMiddleColor(int color)
        {
            if (textBoxShape2 != null)
            {
                textBoxShape2.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = color;
                textBoxShape2.TextFrame2.TextRange.Font.Line.ForeColor.RGB = color;
            }
        }

        private void UpdateBottomColor(int color)
        {
            if (textBoxShape3 != null)
            {
                textBoxShape3.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = color;
                textBoxShape3.TextFrame2.TextRange.Font.Line.ForeColor.RGB = color;
            }
        }

        private void UpdateMiddleOutline(float weight)
        {
            if (textBoxShape2 != null)
            {
                textBoxShape2.TextFrame2.TextRange.Font.Line.Weight = weight;
            }
        }

        private void UpdateBottomOutline(float weight)
        {
            if (textBoxShape3 != null)
            {
                textBoxShape3.TextFrame2.TextRange.Font.Line.Weight = weight;
            }
        }

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
