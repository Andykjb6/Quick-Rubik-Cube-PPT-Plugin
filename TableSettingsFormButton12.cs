using System;
using System.Drawing;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace 课件帮PPT助手
{
    public partial class TableSettingsFormButton12 : Form
    {
        private Color borderColor = Color.Black;

        public TableSettingsFormButton12()
        {
            InitializeComponent();
            checkBoxShape.Checked = true; // 默认勾选表格
        }

        private void ButtonChooseColor_Click(object sender, EventArgs e)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    borderColor = colorDialog.Color;
                    buttonChooseColor.BackColor = borderColor; // 设置按钮的背景颜色为用户选择的颜色
                }
            }
        }

        private void ButtonOK_Click(object sender, EventArgs e)
        {
            bool ctrlPressed = (ModifierKeys & Keys.Control) == Keys.Control;

            if (checkBoxTable.Checked)
            {
                if (ctrlPressed)
                {
                    GenerateTableWithoutLayout();
                }
                else
                {
                    GenerateTable();
                }
            }
            else if (checkBoxShape.Checked)
            {
                if (ctrlPressed)
                {
                    GenerateShapeWithoutLayout();
                }
                else
                {
                    GenerateShape();
                }
            }
        }

        private void ButtonApply_Click(object sender, EventArgs e)
        {
            ApplySettings();
        }

        private void GenerateTable()
        {
            float borderWidth = (float)numericUpDownBorderWidth.Value;
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Slide activeSlide = app.ActiveWindow.View.Slide as PowerPoint.Slide;

            if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange selectedShapes = app.ActiveWindow.Selection.ShapeRange;

                float initialLeft = selectedShapes[1].Left;
                float initialTop = selectedShapes[1].Top;
                float currentLeft = initialLeft;
                float currentTop = initialTop;
                float maxHeightInRow = 0;
                float rowStartTop = initialTop; // 记录当前行起始位置
                float rowSpacing = 10; // 行与行之间的间距

                foreach (PowerPoint.Shape selectedShape in selectedShapes)
                {
                    float selectedSize = Math.Max(selectedShape.Width, selectedShape.Height) + 18;
                    maxHeightInRow = Math.Max(maxHeightInRow, selectedSize);

                    // 如果当前对象的 top 位置与 rowStartTop 的差值大于一个阈值（例如20），说明是新的一行
                    if (Math.Abs(selectedShape.Top - rowStartTop) > 20)
                    {
                        currentLeft = initialLeft; // 重置为行起始位置
                        rowStartTop = selectedShape.Top; // 更新当前行的起始位置
                        currentTop += maxHeightInRow + rowSpacing; // 更新到下一行的顶部位置，并加上行间距
                        maxHeightInRow = selectedSize; // 更新当前行的最大高度
                    }

                    float left = currentLeft;
                    float top = currentTop + (maxHeightInRow - selectedSize) / 2;

                    PowerPoint.Shape tableShape = activeSlide.Shapes.AddTable(2, 2, left, top, selectedSize, selectedSize);
                    tableShape.LockAspectRatio = Office.MsoTriState.msoTrue; // 锁定纵横比

                    PowerPoint.Table table = tableShape.Table;

                    SetTableProperties(table, borderWidth, borderColor);

                    // 将表格置于选中对象的底层
                    tableShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);

                    // 调整选中对象的位置以居中
                    selectedShape.Left = left + (selectedSize - selectedShape.Width) / 2;
                    selectedShape.Top = top + (selectedSize - selectedShape.Height) / 2;

                    // 更新当前 left 位置以紧挨着放置下一个田字格
                    currentLeft += selectedSize;

                    // 确保田字格在选中对象的后面
                    tableShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);

                    // 将田字格置于当前选中对象的底层
                    tableShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                    selectedShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                }
            }
        }

        private void GenerateTableWithoutLayout()
        {
            float borderWidth = (float)numericUpDownBorderWidth.Value;
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Slide activeSlide = app.ActiveWindow.View.Slide as PowerPoint.Slide;

            if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange selectedShapes = app.ActiveWindow.Selection.ShapeRange;

                foreach (PowerPoint.Shape selectedShape in selectedShapes)
                {
                    float selectedSize = Math.Max(selectedShape.Width, selectedShape.Height) + 18;

                    float left = selectedShape.Left;
                    float top = selectedShape.Top;

                    PowerPoint.Shape tableShape = activeSlide.Shapes.AddTable(2, 2, left, top, selectedSize, selectedSize);
                    tableShape.LockAspectRatio = Office.MsoTriState.msoTrue; // 锁定纵横比

                    PowerPoint.Table table = tableShape.Table;

                    SetTableProperties(table, borderWidth, borderColor);

                    // 将表格置于选中对象的底层
                    tableShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);

                    // 调整选中对象的位置以居中
                    selectedShape.Left = left + (selectedSize - selectedShape.Width) / 2;
                    selectedShape.Top = top + (selectedSize - selectedShape.Height) / 2;

                    // 确保田字格在选中对象的后面
                    tableShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);

                    // 将田字格置于当前选中对象的底层
                    tableShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                    selectedShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                }
            }
        }

        private void GenerateShape()
        {
            float borderWidth = (float)numericUpDownBorderWidth.Value;
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Slide activeSlide = app.ActiveWindow.View.Slide as PowerPoint.Slide;

            if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange selectedShapes = app.ActiveWindow.Selection.ShapeRange;

                float initialLeft = selectedShapes[1].Left;
                float initialTop = selectedShapes[1].Top;
                float currentLeft = initialLeft;
                float currentTop = initialTop;
                float maxHeightInRow = 0;
                float rowStartTop = initialTop; // 记录当前行起始位置
                float rowSpacing = 10; // 行与行之间的间距

                foreach (PowerPoint.Shape selectedShape in selectedShapes)
                {
                    float selectedSize = Math.Max(selectedShape.Width, selectedShape.Height) + 18;
                    maxHeightInRow = Math.Max(maxHeightInRow, selectedSize);

                    // 如果当前对象的 top 位置与 rowStartTop 的差值大于一个阈值（例如20），说明是新的一行
                    if (Math.Abs(selectedShape.Top - rowStartTop) > 20)
                    {
                        currentLeft = initialLeft; // 重置为行起始位置
                        rowStartTop = selectedShape.Top; // 更新当前行的起始位置
                        currentTop += maxHeightInRow + rowSpacing; // 更新到下一行的顶部位置，并加上行间距
                        maxHeightInRow = selectedSize; // 更新当前行的最大高度
                    }

                    float left = currentLeft;
                    float top = currentTop + (maxHeightInRow - selectedSize) / 2;

                    // 创建正方形
                    PowerPoint.Shape squareShape = activeSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, selectedSize, selectedSize);
                    squareShape.Line.Weight = borderWidth;
                    squareShape.Line.ForeColor.RGB = ConvertColor(borderColor);
                    squareShape.Fill.Transparency = 1; // 确保填充透明度

                    // 创建两条虚线
                    float halfSize = selectedSize / 2;
                    PowerPoint.Shape verticalLine = activeSlide.Shapes.AddLine(left + halfSize, top, left + halfSize, top + selectedSize);
                    PowerPoint.Shape horizontalLine = activeSlide.Shapes.AddLine(left, top + halfSize, left + selectedSize, top + halfSize);

                    verticalLine.Line.Weight = borderWidth;
                    verticalLine.Line.ForeColor.RGB = ConvertColor(borderColor);
                    verticalLine.Line.DashStyle = Office.MsoLineDashStyle.msoLineDash;

                    horizontalLine.Line.Weight = borderWidth;
                    horizontalLine.Line.ForeColor.RGB = ConvertColor(borderColor);
                    horizontalLine.Line.DashStyle = Office.MsoLineDashStyle.msoLineDash;

                    // 编组形状
                    PowerPoint.ShapeRange shapeRange = activeSlide.Shapes.Range(new string[] { squareShape.Name, verticalLine.Name, horizontalLine.Name });
                    PowerPoint.Shape groupShape = shapeRange.Group();

                    // 将形状置于选中对象的底层
                    groupShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);

                    // 调整选中对象的位置以居中
                    selectedShape.Left = left + (selectedSize - selectedShape.Width) / 2;
                    selectedShape.Top = top + (selectedSize - selectedShape.Height) / 2;

                    // 更新当前 left 位置以紧挨着放置下一个田字格
                    currentLeft += selectedSize;

                    // 确保田字格在选中对象的后面
                    groupShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);

                    // 将田字格置于当前选中对象的底层
                    groupShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                    selectedShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                }
            }
        }

        private void GenerateShapeWithoutLayout()
        {
            float borderWidth = (float)numericUpDownBorderWidth.Value;
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Slide activeSlide = app.ActiveWindow.View.Slide as PowerPoint.Slide;

            if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange selectedShapes = app.ActiveWindow.Selection.ShapeRange;

                foreach (PowerPoint.Shape selectedShape in selectedShapes)
                {
                    float selectedSize = Math.Max(selectedShape.Width, selectedShape.Height) + 18;

                    float left = selectedShape.Left;
                    float top = selectedShape.Top;

                    // 创建正方形
                    PowerPoint.Shape squareShape = activeSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, selectedSize, selectedSize);
                    squareShape.Line.Weight = borderWidth;
                    squareShape.Line.ForeColor.RGB = ConvertColor(borderColor);
                    squareShape.Fill.Transparency = 1; // 确保填充透明度

                    // 创建两条虚线
                    float halfSize = selectedSize / 2;
                    PowerPoint.Shape verticalLine = activeSlide.Shapes.AddLine(left + halfSize, top, left + halfSize, top + selectedSize);
                    PowerPoint.Shape horizontalLine = activeSlide.Shapes.AddLine(left, top + halfSize, left + selectedSize, top + halfSize);

                    verticalLine.Line.Weight = borderWidth;
                    verticalLine.Line.ForeColor.RGB = ConvertColor(borderColor);
                    verticalLine.Line.DashStyle = Office.MsoLineDashStyle.msoLineDash;

                    horizontalLine.Line.Weight = borderWidth;
                    horizontalLine.Line.ForeColor.RGB = ConvertColor(borderColor);
                    horizontalLine.Line.DashStyle = Office.MsoLineDashStyle.msoLineDash;

                    // 编组形状
                    PowerPoint.ShapeRange shapeRange = activeSlide.Shapes.Range(new string[] { squareShape.Name, verticalLine.Name, horizontalLine.Name });
                    PowerPoint.Shape groupShape = shapeRange.Group();

                    // 将形状置于选中对象的底层
                    groupShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);

                    // 调整选中对象的位置以居中
                    selectedShape.Left = left + (selectedSize - selectedShape.Width) / 2;
                    selectedShape.Top = top + (selectedSize - selectedShape.Height) / 2;

                    // 确保田字格在选中对象的后面
                    groupShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);

                    // 将田字格置于当前选中对象的底层
                    groupShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                    selectedShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                }
            }
        }

        private void ApplySettings()
        {
            float borderWidth = (float)numericUpDownBorderWidth.Value;
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Slide activeSlide = app.ActiveWindow.View.Slide as PowerPoint.Slide;

            if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange selectedShapes = app.ActiveWindow.Selection.ShapeRange;

                foreach (PowerPoint.Shape selectedShape in selectedShapes)
                {
                    if (selectedShape.Type == Office.MsoShapeType.msoTable)
                    {
                        PowerPoint.Table table = selectedShape.Table;
                        SetTableProperties(table, borderWidth, borderColor);
                    }
                    else if (selectedShape.Type == Office.MsoShapeType.msoGroup)
                    {
                        foreach (PowerPoint.Shape shape in selectedShape.GroupItems)
                        {
                            if (shape.Type == Office.MsoShapeType.msoAutoShape || shape.Type == Office.MsoShapeType.msoLine)
                            {
                                shape.Line.Weight = borderWidth;
                                shape.Line.ForeColor.RGB = ConvertColor(borderColor);
                            }
                        }
                    }
                }
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

                    cell.Shape.Fill.Transparency = 1; // 确保填充透明度
                    cell.Shape.TextFrame.TextRange.Font.Size = 1; // 设置字号为1

                    SetCellBorder(cell.Borders[PowerPoint.PpBorderType.ppBorderTop], borderWidth, colorRgb);
                    SetCellBorder(cell.Borders[PowerPoint.PpBorderType.ppBorderBottom], borderWidth, colorRgb);
                    SetCellBorder(cell.Borders[PowerPoint.PpBorderType.ppBorderLeft], borderWidth, colorRgb);
                    SetCellBorder(cell.Borders[PowerPoint.PpBorderType.ppBorderRight], borderWidth, colorRgb);
                }
            }

            table.Cell(1, 1).Borders[PowerPoint.PpBorderType.ppBorderBottom].DashStyle = Office.MsoLineDashStyle.msoLineDash;
            table.Cell(1, 1).Borders[PowerPoint.PpBorderType.ppBorderRight].DashStyle = Office.MsoLineDashStyle.msoLineDash;
            table.Cell(1, 2).Borders[PowerPoint.PpBorderType.ppBorderBottom].DashStyle = Office.MsoLineDashStyle.msoLineDash;
            table.Cell(2, 1).Borders[PowerPoint.PpBorderType.ppBorderRight].DashStyle = Office.MsoLineDashStyle.msoLineDash;
        }

        private void SetCellBorder(PowerPoint.LineFormat border, float borderWidth, int colorRgb)
        {
            border.Weight = borderWidth;
            border.ForeColor.RGB = colorRgb;
            border.Visible = Office.MsoTriState.msoTrue;
        }

        private int ConvertColor(Color color)
        {
            return (color.B << 16) | (color.G << 8) | color.R;
        }

        private void CheckBoxTable_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxTable.Checked)
            {
                checkBoxShape.Checked = false;
            }
        }

        private void CheckBoxShape_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxShape.Checked)
            {
                checkBoxTable.Checked = false;
            }
        }
    }
}
