using System;
using System.Collections.Generic;
using System.Windows;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using System.Linq;

namespace 课件帮PPT助手
{
    public partial class RingDistribution : Window
    {
        private Dictionary<int, (float width, float height)> originalSizes = new Dictionary<int, (float width, float height)>();
        private List<Shape> ringShapes = new List<Shape>(); // 保存环形中的形状

        public RingDistribution()
        {
            InitializeComponent();

            QuantitySlider.ValueChanged += Slider_ValueChanged;
            RadiusSlider.ValueChanged += Slider_ValueChanged;
            AutoRotateCheckbox.Checked += Checkbox_Changed;
            AutoRotateCheckbox.Unchecked += Checkbox_Changed;
            ScaleCheckbox.Checked += Checkbox_Changed;
            ScaleCheckbox.Unchecked += Checkbox_Changed;

            // 绑定 TextBox 的 TextChanged 事件
            QuantityValue.TextChanged += QuantityValue_TextChanged;
            RadiusValue.TextChanged += RadiusValue_TextChanged;

            InitializeSlider(); // 初始化滑块设置
            UpdateUI(); // 初始化时更新UI
        }

        private void InitializeSlider()
        {
            var application = Globals.ThisAddIn.Application;
            var selection = application.ActiveWindow.Selection;

            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                int selectedCount = selection.ShapeRange.Count;

                if (selectedCount > 1)
                {
                    QuantitySlider.IsEnabled = false;
                    QuantitySlider.Value = selectedCount;
                }
                else
                {
                    QuantitySlider.IsEnabled = true;
                    QuantitySlider.Minimum = 1;
                    QuantitySlider.Maximum = 100;
                    QuantitySlider.Value = 1;
                }

                QuantityValue.Text = QuantitySlider.Value.ToString();

                // 记录初始选择的形状并保存到ringShapes列表
                ringShapes.Clear();
                foreach (Shape shape in selection.ShapeRange)
                {
                    ringShapes.Add(shape);
                    if (!originalSizes.ContainsKey(shape.Id))
                    {
                        originalSizes[shape.Id] = (shape.Width, shape.Height);
                    }
                }
            }
            else
            {
                this.Close(); // 如果没有选中对象，关闭窗体
            }
        }

        private void Slider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            UpdateUI();
        }

        private void QuantityValue_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (int.TryParse(QuantityValue.Text, out int value))
            {
                QuantitySlider.Value = Math.Max(QuantitySlider.Minimum, Math.Min(QuantitySlider.Maximum, value));
            }
        }

        private void RadiusValue_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (int.TryParse(RadiusValue.Text, out int value))
            {
                RadiusSlider.Value = Math.Max(RadiusSlider.Minimum, Math.Min(RadiusSlider.Maximum, value));
            }
        }

        private void Checkbox_Changed(object sender, RoutedEventArgs e)
        {
            ApplyRingDistribution();
        }

        private void UpdateUI()
        {
            int quantity = (int)QuantitySlider.Value;
            int radius = (int)RadiusSlider.Value;

            QuantityValue.Text = quantity.ToString();
            RadiusValue.Text = radius.ToString();
            ApplyRingDistribution(); // 更新UI时应用环形分布
        }

        private void ApplyRingDistribution()
        {
            int quantity = (int)QuantitySlider.Value;
            double radius = RadiusSlider.Value;
            bool autoRotate = AutoRotateCheckbox.IsChecked == true;
            bool scale = ScaleCheckbox.IsChecked == true;

            var application = Globals.ThisAddIn.Application;
            var slide = application.ActiveWindow.View.Slide;

            if (ringShapes.Count == 0)
            {
                return; // 如果没有选中的形状，直接返回，避免错误提示
            }

            int currentShapeCount = ringShapes.Count;

            // 如果需要更多的形状，进行复制
            if (quantity > currentShapeCount)
            {
                AddShapes(quantity - currentShapeCount);
            }
            // 如果不需要这么多形状，进行删除
            else if (quantity < currentShapeCount)
            {
                RemoveShapes(currentShapeCount - quantity);
            }

            // 重新计算所有对象的等距角度
            double angleStep = 360.0 / quantity;
            for (int i = 0; i < ringShapes.Count; i++)
            {
                double angle = i * angleStep;
                double x = Math.Cos(angle * Math.PI / 180) * radius + slide.Design.SlideMaster.Width / 2 - ringShapes[i].Width / 2;
                double y = Math.Sin(angle * Math.PI / 180) * radius + slide.Design.SlideMaster.Height / 2 - ringShapes[i].Height / 2;

                ringShapes[i].Left = (float)x;
                ringShapes[i].Top = (float)y;

                if (autoRotate)
                {
                    ringShapes[i].Rotation = (float)angle;
                }

                if (scale)
                {
                    float scaleFactor = 1 + (i * 0.1f);
                    ringShapes[i].ScaleWidth(scaleFactor, MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromTopLeft);
                    ringShapes[i].ScaleHeight(scaleFactor, MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromTopLeft);
                }
                else
                {
                    // 恢复所有形状的原始大小
                    if (originalSizes.ContainsKey(ringShapes[i].Id))
                    {
                        var originalSize = originalSizes[ringShapes[i].Id];
                        ringShapes[i].Width = originalSize.width;
                        ringShapes[i].Height = originalSize.height;
                    }
                }
            }
        }

        private void AddShapes(int count)
        {
            var application = Globals.ThisAddIn.Application;
            var originalShape = ringShapes.FirstOrDefault();

            if (originalShape == null)
            {
                return; // 直接返回，不执行后续操作
            }

            for (int i = 0; i < count; i++)
            {
                var newShape = originalShape.Duplicate()[1] as Shape;
                newShape.Tags.Add("RingDistributionCopy", "True");
                ringShapes.Add(newShape);

                if (!originalSizes.ContainsKey(newShape.Id))
                {
                    originalSizes[newShape.Id] = (newShape.Width, newShape.Height);
                }
            }
        }

        private void RemoveShapes(int count)
        {
            for (int i = 0; i < count; i++)
            {
                var shapeToRemove = ringShapes.LastOrDefault();
                if (shapeToRemove != null && shapeToRemove.Tags["RingDistributionCopy"] == "True")
                {
                    shapeToRemove.Delete();
                    ringShapes.Remove(shapeToRemove);
                }
            }
        }
    }
}
