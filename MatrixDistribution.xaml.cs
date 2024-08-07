using Microsoft.Office.Interop.PowerPoint;
using System.Collections.Generic;
using System.Windows;
using System.Linq;
using System;

namespace 课件帮PPT助手
{
    public partial class MatrixDistribution : Window
    {
        private List<Shape> matrixShapes = new List<Shape>();
        private bool isUpdating = false;

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

            // 增量和减量按钮事件绑定
            ColumnsIncrease.Click += ColumnsIncrease_Click;
            ColumnsDecrease.Click += ColumnsDecrease_Click;
            RowSpacingIncrease.Click += RowSpacingIncrease_Click;
            RowSpacingDecrease.Click += RowSpacingDecrease_Click;
            ColumnSpacingIncrease.Click += ColumnSpacingIncrease_Click;
            ColumnSpacingDecrease.Click += ColumnSpacingDecrease_Click;

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
            if (!isUpdating)
            {
                isUpdating = true;
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
                isUpdating = false;
            }
        }

        private void OrientationComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (OrientationComboBox.SelectedIndex == 1) // 选择了“纵向”
            {
                PrimaryCountLabel.Text = "纵向数量："; // 更新标签文本
                ColumnsSlider.Maximum = 100; // 设置最大值为100
                ColumnsSlider.Value = Math.Min(ColumnsSlider.Value, 100);
                ColumnsValue.Text = ColumnsSlider.Value.ToString();
            }
            else
            {
                PrimaryCountLabel.Text = "横向数量："; // 更新标签文本
                ColumnsSlider.Maximum = 100; // 设置最大值为100
                ColumnsSlider.Value = Math.Min(ColumnsSlider.Value, 100);
                ColumnsValue.Text = ColumnsSlider.Value.ToString();
            }

            UpdateUI();
        }

        private void ColumnsIncrease_Click(object sender, RoutedEventArgs e)
        {
            ColumnsSlider.Value = Math.Min(ColumnsSlider.Value + 1, ColumnsSlider.Maximum);
            ColumnsValue.Text = ColumnsSlider.Value.ToString();
        }

        private void ColumnsDecrease_Click(object sender, RoutedEventArgs e)
        {
            ColumnsSlider.Value = Math.Max(ColumnsSlider.Value - 1, ColumnsSlider.Minimum);
            ColumnsValue.Text = ColumnsSlider.Value.ToString();
        }

        private void RowSpacingIncrease_Click(object sender, RoutedEventArgs e)
        {
            RowSpacingSlider.Value = Math.Min(RowSpacingSlider.Value + 1, RowSpacingSlider.Maximum);
            RowSpacingValue.Text = RowSpacingSlider.Value.ToString();
        }

        private void RowSpacingDecrease_Click(object sender, RoutedEventArgs e)
        {
            RowSpacingSlider.Value = Math.Max(RowSpacingSlider.Value - 1, RowSpacingSlider.Minimum);
            RowSpacingValue.Text = RowSpacingSlider.Value.ToString();
        }

        private void ColumnSpacingIncrease_Click(object sender, RoutedEventArgs e)
        {
            ColumnSpacingSlider.Value = Math.Min(ColumnSpacingSlider.Value + 1, ColumnSpacingSlider.Maximum);
            ColumnSpacingValue.Text = ColumnSpacingSlider.Value.ToString();
        }

        private void ColumnSpacingDecrease_Click(object sender, RoutedEventArgs e)
        {
            ColumnSpacingSlider.Value = Math.Max(ColumnSpacingSlider.Value - 1, ColumnSpacingSlider.Minimum);
            ColumnSpacingValue.Text = ColumnSpacingSlider.Value.ToString();
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

            int columnCount = (int)Math.Ceiling((double)matrixShapes.Count / rowCount);

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
