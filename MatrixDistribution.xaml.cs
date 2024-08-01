using Microsoft.Office.Interop.PowerPoint;
using System.Collections.Generic;
using System.Windows;
using System;

namespace 课件帮PPT助手
{
    public partial class MatrixDistribution : Window
    {
        private List<Shape> matrixShapes = new List<Shape>();

        public MatrixDistribution()
        {
            InitializeComponent();

            // 绑定事件
            ColumnsSlider.ValueChanged += SliderOrTextBox_ValueChanged;
            RowSpacingSlider.ValueChanged += SliderOrTextBox_ValueChanged;
            ColumnSpacingSlider.ValueChanged += SliderOrTextBox_ValueChanged;

            ColumnsValue.TextChanged += TextBox_ValueChanged;
            RowSpacingValue.TextChanged += TextBox_ValueChanged;
            ColumnSpacingValue.TextChanged += TextBox_ValueChanged;

            OrientationComboBox.SelectionChanged += OrientationComboBox_SelectionChanged;

            InitializeSlider();
        }

        private void InitializeSlider()
        {
            var application = Globals.ThisAddIn.Application;
            var selection = application.ActiveWindow.Selection;

            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 1)
            {
                matrixShapes.Clear();
                foreach (Shape shape in selection.ShapeRange)
                {
                    matrixShapes.Add(shape);
                }

                int shapeCount = matrixShapes.Count;

                // 根据选中的对象数量，计算初始横向数量（假设为大致的平方根）
                int initialColumnsCount = (int)Math.Ceiling(Math.Sqrt(shapeCount));
                ColumnsSlider.Maximum = 100; // 设置最大值为100
                ColumnsSlider.Value = initialColumnsCount; // 动态设置初始横向数量
                ColumnsValue.Text = initialColumnsCount.ToString(); // 更新文本框值

                ApplyMatrixDistribution();
            }
            else
            {
                Close();
            }
        }


        private void SliderOrTextBox_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            UpdateUI();
        }

        private void TextBox_ValueChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (sender is System.Windows.Controls.TextBox textBox)
            {
                if (int.TryParse(textBox.Text, out int value))
                {
                    if (textBox == ColumnsValue && value >= ColumnsSlider.Minimum && value <= ColumnsSlider.Maximum)
                    {
                        ColumnsSlider.Value = value;
                    }
                    else if (textBox == RowSpacingValue && value >= RowSpacingSlider.Minimum && value <= RowSpacingSlider.Maximum)
                    {
                        RowSpacingSlider.Value = value;
                    }
                    else if (textBox == ColumnSpacingValue && value >= ColumnSpacingSlider.Minimum && value <= ColumnSpacingSlider.Maximum)
                    {
                        ColumnSpacingSlider.Value = value;
                    }
                }
            }
            UpdateUI();
        }

        private void OrientationComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (OrientationComboBox.SelectedIndex == 1) // 选择了“纵向”
            {
                PrimaryCountLabel.Text = "纵向数量："; // 更新标签文本
                ColumnsSlider.Maximum = 100; // 设置最大值为100
                ColumnsSlider.Value = Math.Min(ColumnsSlider.Value, 100);
                ColumnsValue.Text = ColumnsSlider.Value.ToString();
                ColumnsSlider.ToolTip = "调整纵向数量";
            }
            else
            {
                PrimaryCountLabel.Text = "横向数量："; // 更新标签文本
                ColumnsSlider.Maximum = 100; // 设置最大值为100
                ColumnsSlider.Value = Math.Min(ColumnsSlider.Value, 100);
                ColumnsValue.Text = ColumnsSlider.Value.ToString();
                ColumnsSlider.ToolTip = "调整横向数量";
            }

            UpdateUI();
        }

        private void UpdateUI()
        {
            int primaryCount = (int)ColumnsSlider.Value; // 横向或纵向数量
            int rowSpacing = (int)RowSpacingSlider.Value;
            int columnSpacing = (int)ColumnSpacingSlider.Value;

            ColumnsValue.Text = primaryCount.ToString();
            RowSpacingValue.Text = rowSpacing.ToString();
            ColumnSpacingValue.Text = columnSpacing.ToString();

            ApplyMatrixDistribution();
        }

        private void ApplyMatrixDistribution()
        {
            int rowCount = (int)ColumnsSlider.Value;  // 行数（即纵向数量）
            int rowSpacing = (int)RowSpacingSlider.Value;
            int columnSpacing = (int)ColumnSpacingSlider.Value;

            // 计算列数
            int columnCount = (int)Math.Ceiling((double)matrixShapes.Count / rowCount);

            // 获取基准形状的位置
            var firstShape = matrixShapes[0];
            double baseX = firstShape.Left;
            double baseY = firstShape.Top;

            for (int col = 0; col < columnCount; col++)
            {
                for (int row = 0; row < rowCount; row++)
                {
                    int index = col * rowCount + row;
                    if (index >= matrixShapes.Count)
                        break;

                    double x, y;
                    if (OrientationComboBox.SelectedIndex == 1) // 纵向排列
                    {
                        x = baseX + col * (columnSpacing + matrixShapes[index].Width);
                        y = baseY + row * (rowSpacing + matrixShapes[index].Height);
                    }
                    else // 横向排列
                    {
                        x = baseX + row * (columnSpacing + matrixShapes[index].Width);
                        y = baseY + col * (rowSpacing + matrixShapes[index].Height);
                    }

                    matrixShapes[index].Left = (float)x;
                    matrixShapes[index].Top = (float)y;
                }
            }
        }
    }
}
