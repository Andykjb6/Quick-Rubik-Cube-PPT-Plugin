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
        private TrackBar letterSpacingTrackBar;

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
                Size = new System.Drawing.Size(470, 600),
                FormBorderStyle = FormBorderStyle.SizableToolWindow,
                ShowInTaskbar = false
            };

            var tabControl = new TabControl
            {
                Dock = DockStyle.Fill
            };
            form.Controls.Add(tabControl);

            var textSettingsPage = new TabPage("文本设置");
            var shadowSettingsPage = new TabPage("阴影设置");
            var colorSettingsPage = new TabPage("颜色设置");
            var spacingSettingsPage = new TabPage("轮廓设置");

            tabControl.TabPages.Add(textSettingsPage);
            tabControl.TabPages.Add(colorSettingsPage);
            tabControl.TabPages.Add(shadowSettingsPage);
            tabControl.TabPages.Add(spacingSettingsPage);

            var textBox = new TextBox
            {
                Location = new Point(20, 20),
                Size = new Size(380, 80),
                Font = new Font(FontFamily.GenericSansSerif, 10, FontStyle.Regular),
                TextAlign = HorizontalAlignment.Center
            };
            textSettingsPage.Controls.Add(textBox);

            var fontLabel = new Label
            {
                Text = "选择字体：",
                Location = new Point(20, 90),
                Size = new Size(150, 30)
            };
            textSettingsPage.Controls.Add(fontLabel);

            var fontComboBox = new ComboBox
            {
                Location = new Point(20, 120),
                Size = new Size(380, 30)
            };
            foreach (FontFamily fontFamily in FontFamily.Families)
            {
                fontComboBox.Items.Add(fontFamily.Name);
            }
            fontComboBox.SelectedIndex = 0;
            fontComboBox.SelectedIndexChanged += (s, e) =>
            {
                UpdateFont(fontComboBox.SelectedItem.ToString());
            };
            textSettingsPage.Controls.Add(fontComboBox);

            var topColorLabel = new Label
            {
                Text = "顶层颜色：",
                Location = new Point(20, 20),
                Size = new Size(150, 30)
            };
            colorSettingsPage.Controls.Add(topColorLabel);

            var topColorButton = new Button
            {
                Text = "自定义",
                Location = new Point(20, 50),
                Size = new Size(380, 40),
                BackColor = Color.Black
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
            colorSettingsPage.Controls.Add(topColorButton);

            var middleColorLabel = new Label
            {
                Text = "中层颜色：",
                Location = new Point(20, 100),
                Size = new Size(150, 30)
            };
            colorSettingsPage.Controls.Add(middleColorLabel);

            var middleColorButton = new Button
            {
                Text = "自定义",
                Location = new Point(20, 130),
                Size = new Size(380, 40),
                BackColor = Color.White
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
            colorSettingsPage.Controls.Add(middleColorButton);

            var bottomColorLabel = new Label
            {
                Text = "底层颜色：",
                Location = new Point(20, 180),
                Size = new Size(150, 30)
            };
            colorSettingsPage.Controls.Add(bottomColorLabel);

            var bottomColorButton = new Button
            {
                Text = "自定义",
                Location = new Point(20, 210),
                Size = new Size(380, 40),
                BackColor = Color.Blue
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
            colorSettingsPage.Controls.Add(bottomColorButton);

            shadowCheckBox = new CheckBox
            {
                Text = "阴影开关",
                Location = new Point(20, 20),
                Size = new Size(380, 40)
            };
            shadowCheckBox.CheckedChanged += (sender, e) =>
            {
                if (textBoxShape3 != null && textBoxShape3.TextFrame2 != null)
                {
                    textBoxShape3.TextFrame2.TextRange.Font.Shadow.Visible = shadowCheckBox.Checked ? Office.MsoTriState.msoFalse : Office.MsoTriState.msoTrue;
                }
            };
            shadowSettingsPage.Controls.Add(shadowCheckBox);

            var shadowColorButton = new Button
            {
                Text = "更改阴影颜色",
                Location = new Point(20, 70),
                Size = new Size(380, 40),
                BackColor = Color.FromArgb(47, 85, 151),
                ForeColor = Color.White
            };
            shadowColorButton.Click += (s, args) =>
            {
                ColorDialog colorDialog = new ColorDialog();
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    int argbColor = colorDialog.Color.ToArgb();
                    int correctedArgbColor = SwapRedBlue(argbColor);
                    if (textBoxShape3 != null && textBoxShape3.TextFrame2 != null)
                    {
                        textBoxShape3.TextFrame2.TextRange.Font.Shadow.ForeColor.RGB = correctedArgbColor;
                    }
                }
            };
            shadowSettingsPage.Controls.Add(shadowColorButton);

            var shadowBlurLabel = new Label
            {
                Text = "阴影模糊：",
                Location = new Point(20, 130),
                Size = new Size(380, 30)
            };
            shadowSettingsPage.Controls.Add(shadowBlurLabel);

            var shadowBlurTrackBar = new TrackBar
            {
                Minimum = 0,
                Maximum = 100,
                Value = 25,
                Location = new Point(20, 170),
                Size = new Size(380, 30)
            };
            shadowBlurTrackBar.Scroll += (s, args) =>
            {
                int blurValue = shadowBlurTrackBar.Value;
                if (textBoxShape3 != null && textBoxShape3.TextFrame2 != null)
                {
                    textBoxShape3.TextFrame2.TextRange.Font.Shadow.Blur = blurValue;
                }
            };
            shadowSettingsPage.Controls.Add(shadowBlurTrackBar);

            var shadowTransparencyLabel = new Label
            {
                Text = "阴影透明：",
                Location = new Point(20, 260),
                Size = new Size(380, 30)
            };
            shadowSettingsPage.Controls.Add(shadowTransparencyLabel);

            var shadowTransparencyTrackBar = new TrackBar
            {
                Minimum = 0,
                Maximum = 100,
                Value = 65,
                Location = new Point(20, 300),
                Size = new Size(380, 30)
            };
            shadowTransparencyTrackBar.Scroll += (s, args) =>
            {
                float transparencyValue = shadowTransparencyTrackBar.Value / 100f;
                if (textBoxShape3 != null && textBoxShape3.TextFrame2 != null)
                {
                    textBoxShape3.TextFrame2.TextRange.Font.Shadow.Transparency = transparencyValue;
                }
            };
            shadowSettingsPage.Controls.Add(shadowTransparencyTrackBar);

            var letterSpacingLabel = new Label
            {
                Text = "字符间距：",
                Location = new Point(20, 260),
                Size = new Size(380, 30)
            };
            textSettingsPage.Controls.Add(letterSpacingLabel);

            letterSpacingTrackBar = new TrackBar
            {
                Minimum = -50,
                Maximum = 50,
                Value = 0,
                Location = new Point(20, 300),
                Size = new Size(380, 30)
            };
            letterSpacingTrackBar.Scroll += (s, args) =>
            {
                float spacingValue = letterSpacingTrackBar.Value * 0.1f;
                UpdateLetterSpacing(spacingValue);
            };
            textSettingsPage.Controls.Add(letterSpacingTrackBar);

            var middleOutlineLabel = new Label
            {
                Text = "中层轮廓：",
                Location = new Point(20, 20),
                Size = new Size(150, 30)
            };
            spacingSettingsPage.Controls.Add(middleOutlineLabel);

            var middleOutlineNumericUpDown = new NumericUpDown
            {
                Location = new Point(20, 60),
                Size = new Size(380, 50),
                Minimum = 0,
                Maximum = 100,
                Value = 45
            };
            middleOutlineNumericUpDown.ValueChanged += (s, e) =>
            {
                UpdateMiddleOutline(Convert.ToSingle(middleOutlineNumericUpDown.Value));
            };
            spacingSettingsPage.Controls.Add(middleOutlineNumericUpDown);

            var bottomOutlineLabel = new Label
            {
                Text = "底层轮廓：",
                Location = new Point(20, 140),
                Size = new Size(150, 30)
            };
            spacingSettingsPage.Controls.Add(bottomOutlineLabel);

            var bottomOutlineNumericUpDown = new NumericUpDown
            {
                Location = new Point(20, 180),
                Size = new Size(380, 50),
                Minimum = 0,
                Maximum = 100,
                Value = 55
            };
            bottomOutlineNumericUpDown.ValueChanged += (s, e) =>
            {
                UpdateBottomOutline(Convert.ToSingle(bottomOutlineNumericUpDown.Value));
            };
            spacingSettingsPage.Controls.Add(bottomOutlineNumericUpDown);

            var generateButton = new Button
            {
                Text = "生成",
                Location = new Point(20, 180),
                Size = new Size(380, 50),
                BackColor = Color.FromArgb(47, 85, 151),
                ForeColor = Color.White
            };
            generateButton.Click += (s, args) =>
            {
                TryGenerateTextBox(textBox, fontComboBox, topColorButton, middleColorButton, bottomColorButton, middleOutlineNumericUpDown, bottomOutlineNumericUpDown, letterSpacingTrackBar.Value * 0.1f);
            };
            textSettingsPage.Controls.Add(generateButton);

            SetWindowLong(form.Handle, GWL_HWNDPARENT, pptHandle.ToInt32());
            SetWindowPos(form.Handle, (IntPtr)HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

            form.Show();
            form.TopMost = true;
        }

        private void TryGenerateTextBox(TextBox textBox, ComboBox fontComboBox, Button topColorButton, Button middleColorButton, Button bottomColorButton, NumericUpDown middleOutlineNumericUpDown, NumericUpDown bottomOutlineNumericUpDown, float letterSpacing)
        {
            try
            {
                GenerateTextBox(textBox, fontComboBox, topColorButton, middleColorButton, bottomColorButton, middleOutlineNumericUpDown, bottomOutlineNumericUpDown, letterSpacing);
            }
            catch (COMException ex)
            {
                if ((uint)ex.ErrorCode == 0x800A01A8)
                {
                    Debug.WriteLine($"Ignored COM Exception: {ex.Message}");
                }
                else
                {
                    MessageBox.Show($"COM Exception: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Debug.WriteLine($"COM Exception in Generate Button Click: {ex}");
                }
            }
        }

        private void GenerateTextBox(TextBox textBox, ComboBox fontComboBox, Button topColorButton, Button middleColorButton, Button bottomColorButton, NumericUpDown middleOutlineNumericUpDown, NumericUpDown bottomOutlineNumericUpDown, float letterSpacing)
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
            textBoxShape.TextFrame2.TextRange.Font.Spacing = letterSpacing;

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
        }

        private void UpdateFont(string fontName)
        {
            try
            {
                if (textBoxShape != null && textBoxShape.TextFrame2 != null)
                {
                    textBoxShape.TextFrame2.TextRange.Font.NameFarEast = fontName;
                    textBoxShape.TextFrame2.TextRange.Font.Name = fontName;
                    textBoxShape2.TextFrame2.TextRange.Font.NameFarEast = fontName;
                    textBoxShape2.TextFrame2.TextRange.Font.Name = fontName;
                    textBoxShape3.TextFrame2.TextRange.Font.NameFarEast = fontName;
                    textBoxShape3.TextFrame2.TextRange.Font.Name = fontName;
                }
            }
            catch (COMException ex)
            {
                if ((uint)ex.ErrorCode == 0x800A01A8)
                {
                    Debug.WriteLine($"Ignored COM Exception: {ex.Message}");
                }
                else
                {
                    MessageBox.Show($"COM Exception: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Debug.WriteLine($"COM Exception in UpdateFont: {ex}");
                }
            }
        }

        private void UpdateTopColor(int color)
        {
            try
            {
                if (textBoxShape != null && textBoxShape.TextFrame2 != null)
                {
                    textBoxShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = color;
                }
            }
            catch (COMException ex)
            {
                if ((uint)ex.ErrorCode == 0x800A01A8)
                {
                    Debug.WriteLine($"Ignored COM Exception: {ex.Message}");
                }
                else
                {
                    MessageBox.Show($"COM Exception: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Debug.WriteLine($"COM Exception in UpdateTopColor: {ex}");
                }
            }
        }

        private void UpdateMiddleColor(int color)
        {
            try
            {
                if (textBoxShape2 != null && textBoxShape2.TextFrame2 != null)
                {
                    textBoxShape2.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = color;
                    textBoxShape2.TextFrame2.TextRange.Font.Line.ForeColor.RGB = color;
                }
            }
            catch (COMException ex)
            {
                if ((uint)ex.ErrorCode == 0x800A01A8)
                {
                    Debug.WriteLine($"Ignored COM Exception: {ex.Message}");
                }
                else
                {
                    MessageBox.Show($"COM Exception: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Debug.WriteLine($"COM Exception in UpdateMiddleColor: {ex}");
                }
            }
        }

        private void UpdateBottomColor(int color)
        {
            try
            {
                if (textBoxShape3 != null && textBoxShape3.TextFrame2 != null)
                {
                    textBoxShape3.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = color;
                    textBoxShape3.TextFrame2.TextRange.Font.Line.ForeColor.RGB = color;
                }
            }
            catch (COMException ex)
            {
                if ((uint)ex.ErrorCode == 0x800A01A8)
                {
                    Debug.WriteLine($"Ignored COM Exception: {ex.Message}");
                }
                else
                {
                    MessageBox.Show($"COM Exception: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Debug.WriteLine($"COM Exception in UpdateBottomColor: {ex}");
                }
            }
        }

        private void UpdateMiddleOutline(float weight)
        {
            try
            {
                if (textBoxShape2 != null && textBoxShape2.TextFrame2 != null)
                {
                    textBoxShape2.TextFrame2.TextRange.Font.Line.Weight = weight;
                }
            }
            catch (COMException ex)
            {
                if ((uint)ex.ErrorCode == 0x800A01A8)
                {
                    Debug.WriteLine($"Ignored COM Exception: {ex.Message}");
                }
                else
                {
                    MessageBox.Show($"COM Exception: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Debug.WriteLine($"COM Exception in UpdateMiddleOutline: {ex}");
                }
            }
        }

        private void UpdateBottomOutline(float weight)
        {
            try
            {
                if (textBoxShape3 != null && textBoxShape3.TextFrame2 != null)
                {
                    textBoxShape3.TextFrame2.TextRange.Font.Line.Weight = weight;
                }
            }
            catch (COMException ex)
            {
                if ((uint)ex.ErrorCode == 0x800A01A8)
                {
                    Debug.WriteLine($"Ignored COM Exception: {ex.Message}");
                }
                else
                {
                    MessageBox.Show($"COM Exception: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Debug.WriteLine($"COM Exception in UpdateBottomOutline: {ex}");
                }
            }
        }

        private void UpdateLetterSpacing(float spacing)
        {
            try
            {
                if (textBoxShape != null && textBoxShape.TextFrame2 != null)
                {
                    textBoxShape.TextFrame2.TextRange.Font.Spacing = spacing;
                    textBoxShape2.TextFrame2.TextRange.Font.Spacing = spacing;
                    textBoxShape3.TextFrame2.TextRange.Font.Spacing = spacing;
                }
            }
            catch (COMException ex)
            {
                if ((uint)ex.ErrorCode == 0x800A01A8)
                {
                    Debug.WriteLine($"Ignored COM Exception: {ex.Message}");
                }
                else
                {
                    MessageBox.Show($"COM Exception: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Debug.WriteLine($"COM Exception in UpdateLetterSpacing: {ex}");
                }
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
