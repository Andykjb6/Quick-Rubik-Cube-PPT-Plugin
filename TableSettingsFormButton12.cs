using System;
using System.Collections.Generic;
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
        }

        private void ButtonChooseColor_Click(object sender, EventArgs e)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    borderColor = colorDialog.Color;
                }
            }
        }

        private void ButtonOK_Click(object sender, EventArgs e)
        {
            if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
            {
                AddGridUnderShape();
            }
            else
            {
                GenerateTable();
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

                    // 将表格置于底层
                    tableShape.ZOrder(Office.MsoZOrderCmd.msoSendToBack);

                    // 调整选中对象的位置以居中
                    selectedShape.Left = left + (selectedSize - selectedShape.Width) / 2;
                    selectedShape.Top = top + (selectedSize - selectedShape.Height) / 2;

                    // 更新当前 left 位置以紧挨着放置下一个田字格
                    currentLeft += selectedSize;
                }
            }
        }

        private void AddGridUnderShape()
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

                    // 计算田字格的左上角位置，使其中心与选中对象中心重合
                    float left = selectedShape.Left + (selectedShape.Width / 2) - (selectedSize / 2);
                    float top = selectedShape.Top + (selectedShape.Height / 2) - (selectedSize / 2);

                    PowerPoint.Shape tableShape = activeSlide.Shapes.AddTable(2, 2, left, top, selectedSize, selectedSize);
                    tableShape.LockAspectRatio = Office.MsoTriState.msoTrue; // 锁定纵横比

                    PowerPoint.Table table = tableShape.Table;

                    SetTableProperties(table, borderWidth, borderColor);

                    // 将表格置于底层
                    tableShape.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
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

                    cell.Shape.Fill.Transparency = 1;
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
    }
}
