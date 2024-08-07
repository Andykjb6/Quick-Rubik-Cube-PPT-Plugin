using System;
using System.Windows;
using Microsoft.Office.Interop.PowerPoint;
using System.Linq;
using System.Collections.Generic;

namespace 课件帮PPT助手
{
    public partial class RingDistribution : Window
    {
        private List<Shape> ringShapes = new List<Shape>(); // 保存环形中的形状

        public RingDistribution()
        {
            InitializeComponent();

            QuantitySlider.ValueChanged += Slider_ValueChanged;
            RadiusSlider.ValueChanged += Slider_ValueChanged;

            // 绑定 TextBox 的 TextChanged 事件
            QuantityValue.TextChanged += QuantityValue_TextChanged;
            RadiusValue.TextChanged += RadiusValue_TextChanged;

            InitializeShapes(); // 初始化选中的形状
            UpdateUI(); // 初始化时更新UI
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
                ApplyRingDistribution();
            }
        }

        private void RadiusValue_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (int.TryParse(RadiusValue.Text, out int value))
            {
                RadiusSlider.Value = Math.Max(RadiusSlider.Minimum, Math.Min(RadiusSlider.Maximum, value));
                ApplyRingDistribution();
            }
        }

        private void QuantityIncrease_Click(object sender, RoutedEventArgs e)
        {
            QuantitySlider.Value = Math.Min(QuantitySlider.Value + 1, QuantitySlider.Maximum);
            QuantityValue.Text = QuantitySlider.Value.ToString();
        }

        private void QuantityDecrease_Click(object sender, RoutedEventArgs e)
        {
            QuantitySlider.Value = Math.Max(QuantitySlider.Value - 1, QuantitySlider.Minimum);
            QuantityValue.Text = QuantitySlider.Value.ToString();
        }

        private void RadiusIncrease_Click(object sender, RoutedEventArgs e)
        {
            RadiusSlider.Value = Math.Min(RadiusSlider.Value + 1, RadiusSlider.Maximum);
            RadiusValue.Text = RadiusSlider.Value.ToString();
        }

        private void RadiusDecrease_Click(object sender, RoutedEventArgs e)
        {
            RadiusSlider.Value = Math.Max(RadiusSlider.Value - 1, RadiusSlider.Minimum);
            RadiusValue.Text = RadiusSlider.Value.ToString();
        }

        private void InitializeShapes()
        {
            var application = Globals.ThisAddIn.Application;
            var selection = application.ActiveWindow.Selection;

            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                // 清空之前的形状列表
                ringShapes.Clear();

                foreach (Shape shape in selection.ShapeRange)
                {
                    ringShapes.Add(shape);
                }
            }
            else
            {
                MessageBox.Show("请先选择要进行环形分布的形状。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                this.Close(); // 如果没有选中对象，关闭窗体
            }
        }

        private void UpdateUI()
        {
            QuantityValue.Text = QuantitySlider.Value.ToString();
            RadiusValue.Text = RadiusSlider.Value.ToString();
            ApplyRingDistribution();
        }

        private void ApplyRingDistribution()
        {
            int quantity = (int)QuantitySlider.Value;
            double radius = RadiusSlider.Value;

            var application = Globals.ThisAddIn.Application;
            var slide = application.ActiveWindow.View.Slide;

            if (ringShapes.Count == 0 || slide == null)
            {
                return; // 如果没有选中的形状或幻灯片，直接返回
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
                double x = Math.Cos(angle * Math.PI / 180) * radius + slide.Design.SlideMaster.Width / 2;
                double y = Math.Sin(angle * Math.PI / 180) * radius + slide.Design.SlideMaster.Height / 2;

                // 确保索引有效
                if (i < slide.Shapes.Count)
                {
                    var shape = ringShapes[i];
                    shape.Left = (float)x - shape.Width / 2;
                    shape.Top = (float)y - shape.Height / 2;
                    shape.Rotation = (float)angle;
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
