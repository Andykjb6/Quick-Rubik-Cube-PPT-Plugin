using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace 课件帮PPT助手
{
    public partial class SmartScalingForm : Form
    {
        private Dictionary<int, (float Width, float Height, float Left, float Top, float LineWeight, float FontSize, float BevelTopWidth, float BevelTopHeight, float BevelBottomWidth, float BevelBottomHeight, float Depth, float ContourWidth, float ShadowBlur, float ShadowOffsetX, float ShadowOffsetY, float GlowRadius, float TextOutlineWeight, float TextBevelTopWidth, float TextBevelTopHeight, float TextBevelBottomWidth, float TextBevelBottomHeight, float TextDepth, float TextContourWidth, float TableBorderWidth)> originalSizes;
        private PowerPoint.Selection selection;
        private PointF scaleCenter; // 用于存储当前缩放中心
        private bool isHandlingCheckBoxEvent = false; // 防止递归调用

        public SmartScalingForm()
        {
            InitializeComponent();
            SaveOriginalSizes();
            TogglePropertySettings();
            scaleCenter = new PointF(0.5f, 0.5f); // 初始设置为中心点
        }

        private void SaveOriginalSizes()
        {
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            selection = pptApp.ActiveWindow.Selection;
            originalSizes = new Dictionary<int, (float Width, float Height, float Left, float Top, float LineWeight, float FontSize, float BevelTopWidth, float BevelTopHeight, float BevelBottomWidth, float BevelBottomHeight, float Depth, float ContourWidth, float ShadowBlur, float ShadowOffsetX, float ShadowOffsetY, float GlowRadius, float TextOutlineWeight, float TextBevelTopWidth, float TextBevelTopHeight, float TextBevelBottomWidth, float TextBevelBottomHeight, float TextDepth, float TextContourWidth, float TableBorderWidth)>();

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    SaveShapeOriginalSize(shape);
                }
            }
        }

        private void SaveShapeOriginalSize(PowerPoint.Shape shape)
        {
            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                foreach (PowerPoint.Shape subShape in shape.GroupItems)
                {
                    SaveShapeOriginalSize(subShape);
                }
            }
            else
            {
                float lineWeight = 0;
                float fontSize = 0;
                float bevelTopWidth = 0;
                float bevelTopHeight = 0;
                float bevelBottomWidth = 0;
                float bevelBottomHeight = 0;
                float depth = 0;
                float contourWidth = 0;
                float shadowBlur = 0;
                float shadowOffsetX = 0;
                float shadowOffsetY = 0;
                float glowRadius = 0;
                float textOutlineWeight = 0;
                float textBevelTopWidth = 0;
                float textBevelTopHeight = 0;
                float textBevelBottomWidth = 0;
                float textBevelBottomHeight = 0;
                float textDepth = 0;
                float textContourWidth = 0;
                float tableBorderWidth = 0;

                if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    var textRange = shape.TextFrame2.TextRange;
                    fontSize = textRange.Font.Size;
                    if (textRange.Font.Line.Visible == Office.MsoTriState.msoTrue)
                    {
                        textOutlineWeight = textRange.Font.Line.Weight;
                    }
                    var threeDFormat = shape.TextFrame2.ThreeD;
                    textBevelTopWidth = threeDFormat.BevelTopInset;
                    textBevelTopHeight = threeDFormat.BevelTopDepth;
                    textBevelBottomWidth = threeDFormat.BevelBottomInset;
                    textBevelBottomHeight = threeDFormat.BevelBottomDepth;
                    textDepth = threeDFormat.Depth;
                    textContourWidth = threeDFormat.ContourWidth;

                    if (textRange.Font.Shadow.Visible == Office.MsoTriState.msoTrue)
                    {
                        shadowOffsetX = textRange.Font.Shadow.OffsetX;
                        shadowOffsetY = textRange.Font.Shadow.OffsetY;
                    }
                }

                if (shape.HasTable == Office.MsoTriState.msoTrue)
                {
                    // 获取表格边框宽度
                    PowerPoint.Table table = shape.Table;
                    tableBorderWidth = table.Cell(1, 1).Borders[PowerPoint.PpBorderType.ppBorderBottom].Weight;
                }
                else
                {
                    if (shape.Line != null && shape.Line.Visible == Office.MsoTriState.msoTrue)
                    {
                        lineWeight = shape.Line.Weight;
                    }

                    if (shape.ThreeD != null && shape.ThreeD.Visible == Office.MsoTriState.msoTrue)
                    {
                        bevelTopWidth = shape.ThreeD.BevelTopInset;
                        bevelTopHeight = shape.ThreeD.BevelTopDepth;
                        bevelBottomWidth = shape.ThreeD.BevelBottomInset;
                        bevelBottomHeight = shape.ThreeD.BevelBottomDepth;
                        depth = shape.ThreeD.Depth;
                        contourWidth = shape.ThreeD.ContourWidth;
                    }

                    if (shape.Shadow != null && shape.Shadow.Visible == Office.MsoTriState.msoTrue)
                    {
                        shadowBlur = shape.Shadow.Blur;
                        shadowOffsetX = shape.Shadow.OffsetX;
                        shadowOffsetY = shape.Shadow.OffsetY;
                    }

                    if (shape.Glow != null && shape.Glow.Radius > 0)
                    {
                        glowRadius = shape.Glow.Radius;
                    }
                }

                originalSizes[shape.Id] = (
                    shape.Width,
                    shape.Height,
                    shape.Left,
                    shape.Top,
                    lineWeight,
                    fontSize,
                    bevelTopWidth,
                    bevelTopHeight,
                    bevelBottomWidth,
                    bevelBottomHeight,
                    depth,
                    contourWidth,
                    shadowBlur,
                    shadowOffsetX,
                    shadowOffsetY,
                    glowRadius,
                    textOutlineWeight,
                    textBevelTopWidth,
                    textBevelTopHeight,
                    textBevelBottomWidth,
                    textBevelBottomHeight,
                    textDepth,
                    textContourWidth,
                    tableBorderWidth
                );
            }
        }

        private void trackBar_Scroll(object sender, EventArgs e)
        {
            numericUpDown.Value = trackBar.Value;
            ApplyScaling(trackBar.Value / 100.0f);
        }

        private void numericUpDown_ValueChanged(object sender, EventArgs e)
        {
            trackBar.Value = (int)numericUpDown.Value;
            ApplyScaling((float)numericUpDown.Value / 100.0f);
        }

        private void ApplyScaling(float scaleFactor)
        {
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                // 获取选中形状整体的边界框
                float totalLeft = float.MaxValue;
                float totalTop = float.MaxValue;
                float totalRight = float.MinValue;
                float totalBottom = float.MinValue;

                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    totalLeft = Math.Min(totalLeft, shape.Left);
                    totalTop = Math.Min(totalTop, shape.Top);
                    totalRight = Math.Max(totalRight, shape.Left + shape.Width);
                    totalBottom = Math.Max(totalBottom, shape.Top + shape.Height);
                }

                // 计算中心点位置
                float centerX = totalLeft + (totalRight - totalLeft) * scaleCenter.X;
                float centerY = totalTop + (totalBottom - totalTop) * scaleCenter.Y;

                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    ScaleShape(shape, scaleFactor, centerX, centerY);
                }
            }
        }

        private void ScaleShape(PowerPoint.Shape shape, float scaleFactor, float centerX, float centerY)
        {
            // 递归处理组合中的每个子对象
            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                foreach (PowerPoint.Shape subShape in shape.GroupItems)
                {
                    ScaleShape(subShape, scaleFactor, centerX, centerY);
                }
            }
            else
            {
                // 获取原始大小和位置
                if (!originalSizes.ContainsKey(shape.Id))
                {
                    // 如果没有找到，说明此形状是新添加的，忽略
                    return;
                }

                var originalSize = originalSizes[shape.Id];

                // 缩放形状的大小
                shape.Width = originalSize.Width * scaleFactor;
                shape.Height = originalSize.Height * scaleFactor;

                // 缩放位置
                try
                {
                    shape.Left = centerX + (originalSize.Left + originalSize.Width / 2 - centerX) * scaleFactor - shape.Width / 2;
                    shape.Top = centerY + (originalSize.Top + originalSize.Height / 2 - centerY) * scaleFactor - shape.Height / 2;
                }
                catch (ArgumentException)
                {
                    MessageBox.Show("缩放超出范围，无法应用。");
                    return;
                }

                if (shape.HasTable == Office.MsoTriState.msoTrue)
                {
                    // 缩放表格边框宽度
                    PowerPoint.Table table = shape.Table;
                    float newBorderWidth = originalSize.TableBorderWidth * scaleFactor;
                    for (int i = 1; i <= table.Rows.Count; i++)
                    {
                        for (int j = 1; j <= table.Columns.Count; j++)
                        {
                            table.Cell(i, j).Borders[PowerPoint.PpBorderType.ppBorderBottom].Weight = Math.Max(0.25f, Math.Min(6.0f, newBorderWidth));
                            table.Cell(i, j).Borders[PowerPoint.PpBorderType.ppBorderLeft].Weight = Math.Max(0.25f, Math.Min(6.0f, newBorderWidth));
                            table.Cell(i, j).Borders[PowerPoint.PpBorderType.ppBorderRight].Weight = Math.Max(0.25f, Math.Min(6.0f, newBorderWidth));
                            table.Cell(i, j).Borders[PowerPoint.PpBorderType.ppBorderTop].Weight = Math.Max(0.25f, Math.Min(6.0f, newBorderWidth));
                        }
                    }
                }
                else
                {
                    // 缩放线条属性
                    if (shape.Line != null && shape.Line.Visible == Office.MsoTriState.msoTrue)
                    {
                        shape.Line.Weight = originalSize.LineWeight * scaleFactor;
                    }

                    // 缩放字体大小及其相关属性
                    if (shape.TextFrame.HasText == Office.MsoTriState.msoTrue && checkBoxText.Checked)
                    {
                        var textRange = shape.TextFrame2.TextRange;
                        textRange.Font.Size = originalSize.FontSize * scaleFactor;

                        // 缩放文本轮廓
                        if (originalSize.TextOutlineWeight > 0 && textRange.Font.Line.Visible == Office.MsoTriState.msoTrue)
                        {
                            textRange.Font.Line.Weight = originalSize.TextOutlineWeight * scaleFactor;
                        }

                        // 缩放文本三维效果（如果存在）
                        var threeDFormat = shape.TextFrame2.ThreeD;
                        if (threeDFormat.Visible == Office.MsoTriState.msoTrue)
                        {
                            threeDFormat.BevelTopInset = originalSize.TextBevelTopWidth * scaleFactor;
                            threeDFormat.BevelTopDepth = originalSize.TextBevelTopHeight * scaleFactor;
                            threeDFormat.BevelBottomInset = originalSize.TextBevelBottomWidth * scaleFactor;
                            threeDFormat.BevelBottomDepth = originalSize.TextBevelBottomHeight * scaleFactor;
                            threeDFormat.Depth = originalSize.TextDepth * scaleFactor;
                            threeDFormat.ContourWidth = originalSize.TextContourWidth * scaleFactor;
                        }

                        // 缩放文本阴影距离
                        if (textRange.Font.Shadow.Visible == Office.MsoTriState.msoTrue)
                        {
                            textRange.Font.Shadow.OffsetX = Math.Max(0, Math.Min(200, originalSize.ShadowOffsetX * scaleFactor));
                            textRange.Font.Shadow.OffsetY = Math.Max(0, Math.Min(200, originalSize.ShadowOffsetY * scaleFactor));
                        }
                    }

                    // 缩放形状阴影模糊度
                    if (checkBoxShadow.Checked && shape.Shadow != null && shape.Shadow.Visible == Office.MsoTriState.msoTrue)
                    {
                        float newBlur = originalSize.ShadowBlur * scaleFactor;
                        if (newBlur >= 0 && newBlur <= 100) // 设置合理范围
                        {
                            shape.Shadow.Blur = newBlur;
                        }

                        // 缩放形状阴影距离
                        shape.Shadow.OffsetX = Math.Max(0, Math.Min(200, originalSize.ShadowOffsetX * scaleFactor));
                        shape.Shadow.OffsetY = Math.Max(0, Math.Min(200, originalSize.ShadowOffsetY * scaleFactor));
                    }

                    // 缩放形状发光效果
                    if (checkBoxGlow.Checked && shape.Glow != null && shape.Glow.Radius > 0)
                    {
                        shape.Glow.Radius = originalSize.GlowRadius * scaleFactor;
                    }

                    // 缩放形状三维效果
                    if (checkBox3D.Checked && shape.ThreeD != null && shape.ThreeD.Visible == Office.MsoTriState.msoTrue)
                    {
                        shape.ThreeD.BevelTopInset = originalSize.BevelTopWidth * scaleFactor;
                        shape.ThreeD.BevelTopDepth = originalSize.BevelTopHeight * scaleFactor;
                        shape.ThreeD.BevelBottomInset = originalSize.BevelBottomWidth * scaleFactor;
                        shape.ThreeD.BevelBottomDepth = originalSize.BevelBottomHeight * scaleFactor;
                        shape.ThreeD.Depth = originalSize.Depth * scaleFactor;
                        shape.ThreeD.ContourWidth = originalSize.ContourWidth * scaleFactor;
                    }
                }
            }
        }

        private void applyButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void resetButton_Click(object sender, EventArgs e)
        {
            trackBar.Value = 100;
            numericUpDown.Value = 100;
            RestoreOriginalSizes();
        }

        private void RestoreOriginalSizes()
        {
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    if (!originalSizes.ContainsKey(shape.Id))
                    {
                        continue;
                    }

                    var originalSize = originalSizes[shape.Id];

                    shape.Width = originalSize.Width;
                    shape.Height = originalSize.Height;
                    shape.Left = originalSize.Left;
                    shape.Top = originalSize.Top;

                    if (shape.HasTable == Office.MsoTriState.msoTrue)
                    {
                        // 恢复表格边框宽度
                        PowerPoint.Table table = shape.Table;
                        for (int i = 1; i <= table.Rows.Count; i++)
                        {
                            for (int j = 1; i <= table.Columns.Count; j++)
                            {
                                table.Cell(i, j).Borders[PowerPoint.PpBorderType.ppBorderBottom].Weight = originalSize.TableBorderWidth;
                                table.Cell(i, j).Borders[PowerPoint.PpBorderType.ppBorderLeft].Weight = originalSize.TableBorderWidth;
                                table.Cell(i, j).Borders[PowerPoint.PpBorderType.ppBorderRight].Weight = originalSize.TableBorderWidth;
                                table.Cell(i, j).Borders[PowerPoint.PpBorderType.ppBorderTop].Weight = originalSize.TableBorderWidth;
                            }
                        }
                    }
                    else
                    {
                        if (shape.Line != null && shape.Line.Visible == Office.MsoTriState.msoTrue)
                        {
                            shape.Line.Weight = originalSize.LineWeight;
                        }

                        if (shape.ThreeD != null && shape.ThreeD.Visible == Office.MsoTriState.msoTrue)
                        {
                            shape.ThreeD.BevelTopInset = originalSize.BevelTopWidth;
                            shape.ThreeD.BevelTopDepth = originalSize.BevelTopHeight;
                            shape.ThreeD.BevelBottomInset = originalSize.BevelBottomWidth;
                            shape.ThreeD.BevelBottomDepth = originalSize.BevelBottomHeight;
                            shape.ThreeD.Depth = originalSize.Depth;
                            shape.ThreeD.ContourWidth = originalSize.ContourWidth;
                        }

                        if (shape.Shadow != null && shape.Shadow.Visible == Office.MsoTriState.msoTrue)
                        {
                            shape.Shadow.Blur = originalSize.ShadowBlur;
                            shape.Shadow.OffsetX = originalSize.ShadowOffsetX;
                            shape.Shadow.OffsetY = originalSize.ShadowOffsetY;
                        }

                        if (shape.Glow != null && shape.Glow.Radius > 0)
                        {
                            shape.Glow.Radius = originalSize.GlowRadius;
                        }
                    }

                    if (shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        var textRange = shape.TextFrame2.TextRange;
                        textRange.Font.Size = originalSize.FontSize;

                        if (originalSize.TextOutlineWeight > 0 && textRange.Font.Line.Visible == Office.MsoTriState.msoTrue)
                        {
                            textRange.Font.Line.Weight = originalSize.TextOutlineWeight;
                        }

                        var threeDFormat = shape.TextFrame2.ThreeD;
                        threeDFormat.BevelTopInset = originalSize.TextBevelTopWidth;
                        threeDFormat.BevelTopDepth = originalSize.TextBevelTopHeight;
                        threeDFormat.BevelBottomInset = originalSize.TextBevelBottomWidth;
                        threeDFormat.BevelBottomDepth = originalSize.TextBevelBottomHeight;
                        threeDFormat.Depth = originalSize.TextDepth;
                        threeDFormat.ContourWidth = originalSize.TextContourWidth;

                        if (textRange.Font.Shadow.Visible == Office.MsoTriState.msoTrue)
                        {
                            textRange.Font.Shadow.OffsetX = originalSize.ShadowOffsetX;
                            textRange.Font.Shadow.OffsetY = originalSize.ShadowOffsetY;
                        }
                    }
                }
            }
        }

        private void checkBoxPropertySettings_CheckedChanged(object sender, EventArgs e)
        {
            TogglePropertySettings();
        }

        private void TogglePropertySettings()
        {
            bool isChecked = checkBoxPropertySettings.Checked;
            groupBoxShapeAttributes.Visible = isChecked;
            groupBoxTextAttributes.Visible = isChecked;

            if (isChecked)
            {
                this.ClientSize = new System.Drawing.Size(482, 670); // 展开
            }
            else
            {
                this.ClientSize = new System.Drawing.Size(482, 135); // 折叠
            }
        }

        // 缩放中心复选框的事件处理
        private void checkBoxCenter_CheckedChanged(object sender, EventArgs e)
        {
            if (isHandlingCheckBoxEvent) return;
            if (checkBoxCenter.Checked)
            {
                isHandlingCheckBoxEvent = true;
                ResetCenterSelection();
                checkBoxCenter.Checked = true;
                scaleCenter = new PointF(0.5f, 0.5f); // 中心点
                isHandlingCheckBoxEvent = false;
            }
        }

        private void checkBoxTopLeft_CheckedChanged(object sender, EventArgs e)
        {
            if (isHandlingCheckBoxEvent) return;
            if (checkBoxTopLeft.Checked)
            {
                isHandlingCheckBoxEvent = true;
                ResetCenterSelection();
                checkBoxTopLeft.Checked = true;
                scaleCenter = new PointF(0, 0); // 左上角
                isHandlingCheckBoxEvent = false;
            }
        }

        private void checkBoxTopRight_CheckedChanged(object sender, EventArgs e)
        {
            if (isHandlingCheckBoxEvent) return;
            if (checkBoxTopRight.Checked)
            {
                isHandlingCheckBoxEvent = true;
                ResetCenterSelection();
                checkBoxTopRight.Checked = true;
                scaleCenter = new PointF(1, 0); // 右上角
                isHandlingCheckBoxEvent = false;
            }
        }

        private void checkBoxBottomLeft_CheckedChanged(object sender, EventArgs e)
        {
            if (isHandlingCheckBoxEvent) return;
            if (checkBoxBottomLeft.Checked)
            {
                isHandlingCheckBoxEvent = true;
                ResetCenterSelection();
                checkBoxBottomLeft.Checked = true;
                scaleCenter = new PointF(0, 1); // 左下角
                isHandlingCheckBoxEvent = false;
            }
        }

        private void checkBoxBottomRight_CheckedChanged(object sender, EventArgs e)
        {
            if (isHandlingCheckBoxEvent) return;
            if (checkBoxBottomRight.Checked)
            {
                isHandlingCheckBoxEvent = true;
                ResetCenterSelection();
                checkBoxBottomRight.Checked = true;
                scaleCenter = new PointF(1, 1); // 右下角
                isHandlingCheckBoxEvent = false;
            }
        }

        private void ResetCenterSelection()
        {
            checkBoxCenter.Checked = false;
            checkBoxTopLeft.Checked = false;
            checkBoxTopRight.Checked = false;
            checkBoxBottomLeft.Checked = false;
            checkBoxBottomRight.Checked = false;
        }
    }
}
