using System.Collections.Generic;
using System.Windows;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using System.Windows.Controls;
using System;
using System.Linq;

namespace 课件帮PPT助手
{
    public partial class MatrixCopy : Window
    {
        private Dictionary<int, (float width, float height)> originalSizes = new Dictionary<int, (float width, float height)>();
        private List<Shape> matrixShapes = new List<Shape>(); // 保存矩阵中的形状
        private bool isUpdating = false;

        public MatrixCopy()
        {
            InitializeComponent();

            RowsSlider.ValueChanged += SliderOrTextBox_ValueChanged;
            ColumnsSlider.ValueChanged += SliderOrTextBox_ValueChanged;
            RowSpacingSlider.ValueChanged += SliderOrTextBox_ValueChanged;
            ColumnSpacingSlider.ValueChanged += SliderOrTextBox_ValueChanged;

            RowsValue.TextChanged += SliderOrTextBox_ValueChanged;
            ColumnsValue.TextChanged += SliderOrTextBox_ValueChanged;
            RowSpacingValue.TextChanged += SliderOrTextBox_ValueChanged;
            ColumnSpacingValue.TextChanged += SliderOrTextBox_ValueChanged;

            InitializeSlider();
        }

        private void InitializeSlider()
        {
            var application = Globals.ThisAddIn.Application;
            var selection = application.ActiveWindow.Selection;

            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                int selectedCount = selection.ShapeRange.Count;

                int initialRows = (int)Math.Sqrt(selectedCount);
                int initialColumns = (int)Math.Ceiling((double)selectedCount / initialRows);

                RowsSlider.IsEnabled = true;
                ColumnsSlider.IsEnabled = true;
                RowsSlider.Value = initialRows;
                ColumnsSlider.Value = initialColumns;

                RowsValue.Text = RowsSlider.Value.ToString();
                ColumnsValue.Text = ColumnsSlider.Value.ToString();

                // 记录初始选择的形状并保存到matrixShapes列表
                matrixShapes.Clear();
                foreach (Shape shape in selection.ShapeRange)
                {
                    matrixShapes.Add(shape);
                    if (!originalSizes.ContainsKey(shape.Id))
                    {
                        originalSizes[shape.Id] = (shape.Width, shape.Height);
                    }
                }
            }
            else
            {
                Close(); // 直接关闭窗口，不提示信息
            }
        }

        private void SliderOrTextBox_ValueChanged(object sender, RoutedEventArgs e)
        {
            if (!isUpdating)
            {
                isUpdating = true;

                if (sender is TextBox textBox)
                {
                    if (double.TryParse(textBox.Text, out double value))
                    {
                        if (textBox == RowsValue && value >= RowsSlider.Minimum && value <= RowsSlider.Maximum)
                            RowsSlider.Value = value;
                        else if (textBox == ColumnsValue && value >= ColumnsSlider.Minimum && value <= ColumnsSlider.Maximum)
                            ColumnsSlider.Value = value;
                        else if (textBox == RowSpacingValue && value >= RowSpacingSlider.Minimum && value <= RowSpacingSlider.Maximum)
                            RowSpacingSlider.Value = value;
                        else if (textBox == ColumnSpacingValue && value >= ColumnSpacingSlider.Minimum && value <= ColumnSpacingSlider.Maximum)
                            ColumnSpacingSlider.Value = value;
                    }
                }

                UpdateUI();
                isUpdating = false;
            }
        }

        private void UpdateUI()
        {
            if (RowsSlider == null || ColumnsSlider == null || RowSpacingSlider == null || ColumnSpacingSlider == null)
            {
                return;
            }

            int rows = (int)RowsSlider.Value;
            int columns = (int)ColumnsSlider.Value;
            int rowSpacing = (int)RowSpacingSlider.Value;
            int columnSpacing = (int)ColumnSpacingSlider.Value;

            RowsValue.Text = rows.ToString();
            ColumnsValue.Text = columns.ToString();
            RowSpacingValue.Text = rowSpacing.ToString();
            ColumnSpacingValue.Text = columnSpacing.ToString();

            ApplyMatrixCopy();
        }

        private void ApplyMatrixCopy()
        {
            int rows = (int)RowsSlider.Value;
            int columns = (int)ColumnsSlider.Value;
            int rowSpacing = (int)RowSpacingSlider.Value;
            int columnSpacing = (int)ColumnSpacingSlider.Value;

            var application = Globals.ThisAddIn.Application;
            _ = application.ActiveWindow.View.Slide;

            int totalShapesNeeded = rows * columns;

            // 确保 matrixShapes 中至少有一个形状，避免索引超出范围的错误
            if (matrixShapes.Count == 0)
            {
                return; // 直接跳过，不执行后续操作
            }

            // 调整形状数量：添加或移除形状
            if (matrixShapes.Count < totalShapesNeeded)
            {
                AddShapes(totalShapesNeeded - matrixShapes.Count);
            }
            else if (matrixShapes.Count > totalShapesNeeded)
            {
                RemoveShapes(matrixShapes.Count - totalShapesNeeded);
            }

            // 获取基准形状的位置
            var firstShape = matrixShapes[0];
            double baseX = firstShape.Left;
            double baseY = firstShape.Top;

            // 重新排列矩阵中的形状
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < columns; c++)
                {
                    int index = r * columns + c;
                    if (index >= matrixShapes.Count)
                        break;

                    double x = baseX + c * (columnSpacing + matrixShapes[index].Width);
                    double y = baseY + r * (rowSpacing + matrixShapes[index].Height);

                    matrixShapes[index].Left = (float)x;
                    matrixShapes[index].Top = (float)y;

                    // 恢复原始大小
                    if (originalSizes.ContainsKey(matrixShapes[index].Id))
                    {
                        var originalSize = originalSizes[matrixShapes[index].Id];
                        matrixShapes[index].Width = originalSize.width;
                        matrixShapes[index].Height = originalSize.height;
                    }
                }
            }
        }

        private void AddShapes(int count)
        {
            _ = Globals.ThisAddIn.Application;
            var originalShape = matrixShapes.FirstOrDefault();

            if (originalShape == null)
            {
                return; // 直接返回，不执行后续操作
            }

            for (int i = 0; i < count; i++)
            {
                var newShape = originalShape.Duplicate()[1];
                newShape.Tags.Add("MatrixCopyCopy", "True");
                matrixShapes.Add(newShape);

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
                var shapeToRemove = matrixShapes.LastOrDefault();
                if (shapeToRemove != null && shapeToRemove.Tags["MatrixCopyCopy"] == "True")
                {
                    shapeToRemove.Delete();
                    matrixShapes.Remove(shapeToRemove);
                }
            }
        }
    }
}
