using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace 课件帮PPT助手
{
    public partial class CustomCloudTextGeneratorForm : Form
    {
        private PowerPoint.Shape textBoxShape;
        private PowerPoint.Shape textBoxShape2;
        private PowerPoint.Shape textBoxShape3;

        public CustomCloudTextGeneratorForm()
        {
            InitializeComponent();
            InitializeFontComboBox();
        }

        public void InitializeForm()
        {
            InitializeFontComboBox();
        }

        private void InitializeFontComboBox()
        {
            foreach (FontFamily fontFamily in FontFamily.Families)
            {
                this.fontComboBox.Items.Add(fontFamily.Name);
            }
            this.fontComboBox.SelectedIndex = 0;
            this.fontComboBox.SelectedIndexChanged += new System.EventHandler(this.fontComboBox_SelectedIndexChanged);
        }

       

        private void topColorButton_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                topColorButton.BackColor = colorDialog.Color;
                UpdateTopColor(SwapRedBlue(colorDialog.Color.ToArgb()));
            }
        }

        private void middleColorButton_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                middleColorButton.BackColor = colorDialog.Color;
                UpdateMiddleColor(SwapRedBlue(colorDialog.Color.ToArgb()));
            }
        }

        private void bottomColorButton_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                bottomColorButton.BackColor = colorDialog.Color;
                UpdateBottomColor(SwapRedBlue(colorDialog.Color.ToArgb()));
            }
        }

        private void fontComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateFont(fontComboBox.SelectedItem.ToString());
        }

        private void fontSizeNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            UpdateFontSize(Convert.ToSingle(fontSizeNumericUpDown.Value));
        }

        private void generateButton_Click(object sender, EventArgs e)
        {
            TryGenerateTextBox(textBox, fontComboBox, topColorButton, middleColorButton, bottomColorButton, middleOutlineNumericUpDown, bottomOutlineNumericUpDown, letterSpacingTrackBar.Value * 0.1f);
        }

        private void letterSpacingTrackBar_Scroll(object sender, EventArgs e)
        {
            float spacingValue = letterSpacingTrackBar.Value * 0.1f;
            UpdateLetterSpacing(spacingValue);
        }

        private void middleOutlineNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            UpdateMiddleOutline(Convert.ToSingle(middleOutlineNumericUpDown.Value));
        }

        private void bottomOutlineNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            UpdateBottomOutline(Convert.ToSingle(bottomOutlineNumericUpDown.Value));
        }

        private void shadowCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (textBoxShape3 != null && textBoxShape3.TextFrame2 != null)
            {
                textBoxShape3.TextFrame2.TextRange.Font.Shadow.Visible = shadowCheckBox.Checked ? Office.MsoTriState.msoFalse : Office.MsoTriState.msoTrue;
            }
        }


        private void shadowColorButton_Click(object sender, EventArgs e)
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
        }

        private void shadowBlurTrackBar_Scroll(object sender, EventArgs e)
        {
            int blurValue = shadowBlurTrackBar.Value;
            if (textBoxShape3 != null && textBoxShape3.TextFrame2 != null)
            {
                textBoxShape3.TextFrame2.TextRange.Font.Shadow.Blur = blurValue;
            }
        }

        private void shadowTransparencyTrackBar_Scroll(object sender, EventArgs e)
        {
            float transparencyValue = shadowTransparencyTrackBar.Value / 100f;
            if (textBoxShape3 != null && textBoxShape3.TextFrame2 != null)
            {
                textBoxShape3.TextFrame2.TextRange.Font.Shadow.Transparency = transparencyValue;
            }
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
            textBoxShape.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            textBoxShape.TextFrame.WordWrap = Office.MsoTriState.msoFalse;

            float textWidth = textBoxShape.TextFrame2.TextRange.BoundWidth;
            float textHeight = textBoxShape.TextFrame2.TextRange.BoundHeight;

            textBoxShape.Width = textWidth + 10;
            textBoxShape.Height = textHeight + 10;
            textBoxShape.TextFrame.TextRange.Font.NameFarEast = fontComboBox.SelectedItem.ToString();
            textBoxShape.TextFrame.TextRange.Font.Name = fontComboBox.SelectedItem.ToString();
            textBoxShape.TextFrame2.TextRange.Font.Size = Convert.ToSingle(fontSizeNumericUpDown.Value);
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
                MessageBox.Show($"COM Exception: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debug.WriteLine($"COM Exception in UpdateFont: {ex}");
            }
        }

        private void UpdateFontSize(float fontSize)
        {
            if (textBoxShape != null && textBoxShape.TextFrame2 != null)
            {
                textBoxShape.TextFrame2.TextRange.Font.Size = fontSize;
                if (textBoxShape2 != null)
                    textBoxShape2.TextFrame2.TextRange.Font.Size = fontSize;
                if (textBoxShape3 != null)
                    textBoxShape3.TextFrame2.TextRange.Font.Size = fontSize;
            }
        }

        private void UpdateTopColor(int color)
        {
            if (textBoxShape != null && textBoxShape.TextFrame2 != null)
            {
                textBoxShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = color;
            }
        }

        private void UpdateMiddleColor(int color)
        {
            if (textBoxShape2 != null && textBoxShape2.TextFrame2 != null)
            {
                textBoxShape2.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = color;
                textBoxShape2.TextFrame2.TextRange.Font.Line.ForeColor.RGB = color;
            }
        }

        private void UpdateBottomColor(int color)
        {
            if (textBoxShape3 != null && textBoxShape3.TextFrame2 != null)
            {
                textBoxShape3.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = color;
                textBoxShape3.TextFrame2.TextRange.Font.Line.ForeColor.RGB = color;
            }
        }

        private void UpdateLetterSpacing(float spacing)
        {
            if (textBoxShape != null && textBoxShape.TextFrame2 != null)
            {
                textBoxShape.TextFrame2.TextRange.Font.Spacing = spacing;
                if (textBoxShape2 != null)
                    textBoxShape2.TextFrame2.TextRange.Font.Spacing = spacing;
                if (textBoxShape3 != null)
                    textBoxShape3.TextFrame2.TextRange.Font.Spacing = spacing;
            }
        }

        private void UpdateMiddleOutline(float outlineWidth)
        {
            if (textBoxShape2 != null && textBoxShape2.TextFrame2 != null)
            {
                textBoxShape2.TextFrame2.TextRange.Font.Line.Weight = outlineWidth;
            }
        }

        private void UpdateBottomOutline(float outlineWidth)
        {
            if (textBoxShape3 != null && textBoxShape3.TextFrame2 != null)
            {
                textBoxShape3.TextFrame2.TextRange.Font.Line.Weight = outlineWidth;
            }
        }

        private int SwapRedBlue(int color)
        {
            int a = (color >> 24) & 0xFF;
            int r = (color >> 16) & 0xFF;
            int g = (color >> 8) & 0xFF;
            int b = color & 0xFF;

            return (a << 24) | (b << 16) | (g << 8) | r;
        }
    }
}
