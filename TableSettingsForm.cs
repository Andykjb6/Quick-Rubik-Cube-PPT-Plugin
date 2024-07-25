using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace 课件帮PPT助手
{
    public partial class TableSettingsForm : Form
    {
        private Color borderColor = Color.Black;

        public TableSettingsForm()
        {
            InitializeComponent();
        }

        private void ButtonChooseColor_Click(object sender, EventArgs e)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    borderColor = colorDialog.Color;
                    ApplySettings();
                }
            }
        }

        private void ButtonOK_Click(object sender, EventArgs e)
        {
            GenerateTable();
        }

        private void ApplySettings()
        {
            // Your logic to apply settings based on user input can go here.
        }

        private void GenerateTable()
        {
            int rows = (int)numericUpDownRows.Value;
            int columns = (int)numericUpDownColumns.Value;
            float rowSpacing = (float)numericUpDownRowSpacing.Value;
            float columnSpacing = (float)numericUpDownColumnSpacing.Value;
            float borderWidth = (float)numericUpDownBorderWidth.Value;
            float scale = trackBarScale.Value / 100f;

            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Slide activeSlide = app.ActiveWindow.View.Slide as PowerPoint.Slide;

            float startX = 100; // 初始X位置，可以根据需要调整
            float startY = 100; // 初始Y位置，可以根据需要调整
            float squareSize = 100 * scale; // 根据缩放比例调整每个正方形表格的大小

            // 保存现有对象的位置和顺序
            var originalShapes = new List<PowerPoint.Shape>();
            for (int i = 1; i <= activeSlide.Shapes.Count; i++)
            {
                originalShapes.Add(activeSlide.Shapes[i]);
            }

            // 生成表格
            List<PowerPoint.Shape> newTableShapes = new List<PowerPoint.Shape>();
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    float left = startX + j * (squareSize + columnSpacing);
                    float top = startY + i * (squareSize + rowSpacing);

                    PowerPoint.Shape tableShape = activeSlide.Shapes.AddTable(2, 2, left, top, squareSize, squareSize);
                    tableShape.LockAspectRatio = Office.MsoTriState.msoTrue; // 锁定纵横比

                    PowerPoint.Table table = tableShape.Table;
                    SetTableProperties(table, borderWidth, borderColor);

                    newTableShapes.Add(tableShape);
                }
            }

            // 检查是否有选中的对象
            try
            {
                var selection = app.ActiveWindow.Selection;
                if (selection.ShapeRange.Count > 0)
                {
                    int shapeIndex = 0;
                    foreach (PowerPoint.Shape selectedShape in selection.ShapeRange)
                    {
                        float left = startX + (shapeIndex % columns) * (squareSize + columnSpacing);
                        float top = startY + (shapeIndex / columns) * (squareSize + rowSpacing);

                        selectedShape.Left = left + (squareSize - selectedShape.Width) / 2;
                        selectedShape.Top = top + (squareSize - selectedShape.Height) / 2;

                        shapeIndex++;
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // 没有选中任何对象
            }

            // 恢复原始对象位置和顺序
            foreach (var shape in originalShapes)
            {
                shape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            }

            // 确保新表格在最前面
            foreach (var tableShape in newTableShapes)
            {
                tableShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            }
        }

        private void SetTableProperties(PowerPoint.Table table, float borderWidth, Color borderColor)
        {
            int colorRgb = ConvertColor(borderColor);

            for (int i = 1; i <= table.Rows.Count; i++)
            {
                for (int j = 1; j <= table.Columns.Count; j++)
                {
                    PowerPoint.Cell cell = table.Cell(i, j);

                    cell.Shape.Fill.Transparency = 1;
                    cell.Shape.TextFrame.TextRange.Font.Size = 1;

                    if (i == 1)
                    {
                        SetCellBorder(cell.Borders[PowerPoint.PpBorderType.ppBorderTop], borderWidth, colorRgb, true);
                    }
                    if (i == table.Rows.Count)
                    {
                        SetCellBorder(cell.Borders[PowerPoint.PpBorderType.ppBorderBottom], borderWidth, colorRgb, true);
                    }
                    if (j == 1)
                    {
                        SetCellBorder(cell.Borders[PowerPoint.PpBorderType.ppBorderLeft], borderWidth, colorRgb, true);
                    }
                    if (j == table.Columns.Count)
                    {
                        SetCellBorder(cell.Borders[PowerPoint.PpBorderType.ppBorderRight], borderWidth, colorRgb, true);
                    }

                    if (i < table.Rows.Count)
                    {
                        SetCellBorder(cell.Borders[PowerPoint.PpBorderType.ppBorderBottom], borderWidth, colorRgb, false);
                    }
                    if (j < table.Columns.Count)
                    {
                        SetCellBorder(cell.Borders[PowerPoint.PpBorderType.ppBorderRight], borderWidth, colorRgb, false);
                    }
                }
            }
        }

        private void SetCellBorder(PowerPoint.LineFormat border, float borderWidth, int colorRgb, bool isOuterCell)
        {
            border.Weight = borderWidth;
            border.ForeColor.RGB = colorRgb;
            border.Visible = Office.MsoTriState.msoTrue;
            border.DashStyle = isOuterCell ? Office.MsoLineDashStyle.msoLineSolid : Office.MsoLineDashStyle.msoLineDash;
        }

        private int ConvertColor(Color color)
        {
            return (color.B << 16) | (color.G << 8) | color.R;
        }
    }
}
