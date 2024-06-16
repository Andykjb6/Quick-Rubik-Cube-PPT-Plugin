using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace 课件帮PPT助手
{
    public partial class SmartScalingForm : Form
    {
        private Dictionary<int, (float Width, float Height, float Left, float Top, float LineWeight, float FontSize, float BevelTopWidth, float BevelTopHeight, float BevelBottomWidth, float BevelBottomHeight, float Depth, float ContourWidth, float ShadowBlur, float GlowRadius, float TextOutlineWeight, float TextBevelTopWidth, float TextBevelTopHeight, float TextBevelBottomWidth, float TextBevelBottomHeight, float TextDepth, float TextContourWidth)> originalSizes;
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
            originalSizes = new Dictionary<int, (float Width, float Height, float Left, float Top, float LineWeight, float FontSize, float BevelTopWidth, float BevelTopHeight, float BevelBottomWidth, float BevelBottomHeight, float Depth, float ContourWidth, float ShadowBlur, float GlowRadius, float TextOutlineWeight, float TextBevelTopWidth, float TextBevelTopHeight, float TextBevelBottomWidth, float TextBevelBottomHeight, float TextDepth, float TextContourWidth)>();

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
            if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
            {
                foreach (PowerPoint.Shape subShape in shape.GroupItems)
                {
                    SaveShapeOriginalSize(subShape);
                }
            }
            else
            {
                float textOutlineWeight = 0;
                float textBevelTopWidth = 0;
                float textBevelTopHeight = 0;
                float textBevelBottomWidth = 0;
                float textBevelBottomHeight = 0;
                float textDepth = 0;
                float textContourWidth = 0;

                if (shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    var textRange = shape.TextFrame.TextRange;
                    textOutlineWeight = shape.Line.Weight > 0 ? shape.Line.Weight : 0; // 检查 shape.Line.Weight 是否有效
                    textBevelTopWidth = shape.ThreeD.BevelTopInset;
                    textBevelTopHeight = shape.ThreeD.BevelTopDepth;
                    textBevelBottomWidth = shape.ThreeD.BevelBottomInset;
                    textBevelBottomHeight = shape.ThreeD.BevelBottomDepth;
                    textDepth = shape.ThreeD.Depth;
                    textContourWidth = shape.ThreeD.ContourWidth;
                }

                originalSizes[shape.Id] = (
                    shape.Width,
                    shape.Height,
                    shape.Left,
                    shape.Top,
                    shape.Line.Weight,
                    shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue ? shape.TextFrame.TextRange.Font.Size : 0,
                    shape.ThreeD.BevelTopInset,
                    shape.ThreeD.BevelTopDepth,
                    shape.ThreeD.BevelBottomInset,
                    shape.ThreeD.BevelBottomDepth,
                    shape.ThreeD.Depth,
                    shape.ThreeD.ContourWidth,
                    shape.Shadow.Visible == Microsoft.Office.Core.MsoTriState.msoTrue ? shape.Shadow.Blur : 0,
                    shape.Glow.Radius,
                    textOutlineWeight,
                    textBevelTopWidth,
                    textBevelTopHeight,
                    textBevelBottomWidth,
                    textBevelBottomHeight,
                    textDepth,
                    textContourWidth
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
            if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
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

                // 缩放线条属性
                if (shape.Line.Weight > 0)
                {
                    shape.Line.Weight = originalSize.LineWeight * scaleFactor;
                }

                // 缩放字体大小及其相关属性
                if (shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue && checkBoxText.Checked)
                {
                    var textRange = shape.TextFrame.TextRange;
                    textRange.Font.Size = originalSize.FontSize * scaleFactor;

                    // 缩放文本轮廓
                    if (originalSize.TextOutlineWeight > 0)
                    {
                        shape.Line.Weight = originalSize.TextOutlineWeight * scaleFactor; // 检查 TextOutlineWeight 是否有效
                    }

                    // 缩放文本三维效果
                    shape.ThreeD.BevelTopInset = originalSize.TextBevelTopWidth * scaleFactor;
                    shape.ThreeD.BevelTopDepth = originalSize.TextBevelTopHeight * scaleFactor;
                    shape.ThreeD.BevelBottomInset = originalSize.TextBevelBottomWidth * scaleFactor;
                    shape.ThreeD.BevelBottomDepth = originalSize.TextBevelBottomHeight * scaleFactor;
                    shape.ThreeD.Depth = originalSize.TextDepth * scaleFactor;
                    shape.ThreeD.ContourWidth = originalSize.TextContourWidth * scaleFactor;
                }

                // 缩放形状阴影模糊度
                if (checkBoxShadow.Checked && shape.Shadow.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    float newBlur = originalSize.ShadowBlur * scaleFactor;
                    if (newBlur >= 0 && newBlur <= 100) // 设置合理范围
                    {
                        shape.Shadow.Blur = newBlur;
                    }
                }

                // 缩放形状发光效果
                if (checkBoxGlow.Checked)
                {
                    shape.Glow.Radius = originalSize.GlowRadius * scaleFactor;
                }

                // 缩放形状三维效果
                if (checkBox3D.Checked && shape.ThreeD.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)
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

        private void applyButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void resetButton_Click(object sender, EventArgs e)
        {
            trackBar.Value = 100;
            numericUpDown.Value = 100;
            ApplyScaling(1.0f); // 恢复到原始大小
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
                this.ClientSize = new System.Drawing.Size(482, 600); // 展开
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

        private void groupBoxTextAttributes_Enter(object sender, EventArgs e)
        {

        }
    }
}
