using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Color = System.Windows.Media.Color;
using System.Collections.Generic;

namespace 课件帮PPT助手
{
    public partial class TableSettingsForm : Window
    {
        private Color borderColor = Colors.Black;
        private int tianZiGeCounter = 1;

        public TableSettingsForm()
        {
            InitializeComponent();

            // 检查当前选中的对象
            Microsoft.Office.Interop.PowerPoint.Application app = Globals.ThisAddIn.Application;
            if (app.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionShapes)
            {
                ShapeRange selectedShapes = app.ActiveWindow.Selection.ShapeRange;
                bool allAreTianZiGe = true;

                float? borderWidth = null;
                float? widthDifference = null;
                string borderColorString = null;
                float? brightnessDifference = null;

                int tianZiGeCount = 0;

                foreach (Shape groupShape in selectedShapes)
                {
                    if (groupShape.Tags["IsTianZiGe"] == "True" ||
                        (groupShape.Type == Office.MsoShapeType.msoGroup && groupShape.GroupItems[1].Tags["IsTianZiGe"] == "True"))
                    {
                        float currentBorderWidth = float.Parse(groupShape.Tags["BorderWidth"]);
                        float currentWidthDifference = float.Parse(groupShape.Tags["WidthDifference"]);
                        string currentBorderColorString = groupShape.Tags["BorderColor"];
                        float currentBrightnessDifference = float.Parse(groupShape.Tags["BrightnessDifference"]);

                        if (borderWidth == null)
                        {
                            // 初始化参数
                            borderWidth = currentBorderWidth;
                            widthDifference = currentWidthDifference;
                            borderColorString = currentBorderColorString;
                            brightnessDifference = currentBrightnessDifference;
                        }
                        else
                        {
                            // 检查参数是否一致
                            if (borderWidth != currentBorderWidth ||
                                widthDifference != currentWidthDifference ||
                                borderColorString != currentBorderColorString ||
                                brightnessDifference != currentBrightnessDifference)
                            {
                                allAreTianZiGe = false;
                                break;
                            }
                        }

                        tianZiGeCount++;
                    }
                }

                if (tianZiGeCount > 0 && allAreTianZiGe && borderWidth != null)
                {
                    // 同步参数到窗体
                    TextBoxBorderWidth.Text = borderWidth.Value.ToString();
                    TextBoxWidthDifference.Text = widthDifference.Value.ToString();
                    borderColor = (Color)ColorConverter.ConvertFromString(borderColorString);
                    ButtonChooseColor.Background = new SolidColorBrush(borderColor);

                    TextBoxBrightnessDifference.Text = brightnessDifference.Value.ToString();
                }
                else
                {
                    // 如果没有选中任何田字格对象或参数不一致，显示默认参数
                    if (tianZiGeCount == 0)
                    {
                        // 选中对象中没有田字格，显示默认参数
                        TextBoxBorderWidth.Text = "1.25";
                        TextBoxWidthDifference.Text = "0";
                        borderColor = Colors.Black;
                        ButtonChooseColor.Background = new SolidColorBrush(borderColor);
                        TextBoxBrightnessDifference.Text = "0";
                    }
                    else
                    {
                        MessageBox.Show("选中的田字格参数不一致，无法批量修改。");
                    }
                }
            }
        }


        private void ButtonChooseColor_Click(object sender, RoutedEventArgs e)
        {
            var colorDialog = new System.Windows.Forms.ColorDialog();
            if (colorDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                borderColor = Color.FromArgb(colorDialog.Color.A, colorDialog.Color.R, colorDialog.Color.G, colorDialog.Color.B);
                ButtonChooseColor.Background = new SolidColorBrush(borderColor);
            }
        }

        private void ButtonOK_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.PowerPoint.Application app = Globals.ThisAddIn.Application;

            // 检查当前选中的对象类型
            if (app.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionShapes)
            {
                ShapeRange selectedShapes = app.ActiveWindow.Selection.ShapeRange;

                foreach (Shape selectedShape in selectedShapes)
                {
                    // 检查所选对象是否已包含田字格
                    if (selectedShape.Tags["IsTianZiGe"] == "True" ||
                        (selectedShape.Type == Office.MsoShapeType.msoGroup && selectedShape.GroupItems[1].Tags["IsTianZiGe"] == "True"))
                    {
                        MessageBoxResult result = MessageBox.Show("所选对象已有田字格，是否需要再次添加？", "确认", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                        if (result == MessageBoxResult.No)
                        {
                            return; // 如果用户选择不添加，直接返回
                        }
                        break; // 只要有一个形状符合条件，就不再继续检查
                    }

                    // 检查所选对象是否与图层名称前缀为“田字格”的对象重叠
                    foreach (Shape shape in app.ActiveWindow.View.Slide.Shapes)
                    {
                        if (shape.Name.StartsWith("田字格") && AreShapesOverlapping(selectedShape, shape))
                        {
                            MessageBoxResult result = MessageBox.Show("所选对象与已有的田字格存在重叠，是否需要再次添加？", "确认", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                            if (result == MessageBoxResult.No)
                            {
                                return; // 如果用户选择不添加，直接返回
                            }
                            break; // 只要有一个形状符合条件，就不再继续检查
                        }
                    }
                }

                bool ctrlPressed = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;
                if (ctrlPressed)
                {
                    GenerateShapeWithoutLayout();
                }
                else
                {
                    GenerateShape();
                }
            }
            else
            {
                MessageBox.Show("请先选中一个或多个对象。");
            }
        }

        // 辅助方法：检查两个形状是否重叠
        private bool AreShapesOverlapping(Shape shape1, Shape shape2)
        {
            float shape1Left = shape1.Left;
            float shape1Top = shape1.Top;
            float shape1Right = shape1Left + shape1.Width;
            float shape1Bottom = shape1Top + shape1.Height;

            float shape2Left = shape2.Left;
            float shape2Top = shape2.Top;
            float shape2Right = shape2Left + shape2.Width;
            float shape2Bottom = shape2Top + shape2.Height;

            // 检查是否有重叠
            bool isOverlapping = !(shape1Left >= shape2Right ||
                                   shape1Right <= shape2Left ||
                                   shape1Top >= shape2Bottom ||
                                   shape1Bottom <= shape2Top);

            return isOverlapping;
        }


        private void ButtonApply_Click(object sender, RoutedEventArgs e)
        {
            if (!float.TryParse(TextBoxBorderWidth.Text, out float borderWidth))
            {
                MessageBox.Show("无效的边框宽度");
                return;
            }

            float widthDifference = float.Parse(TextBoxWidthDifference.Text);
            float brightnessDifference = float.Parse(TextBoxBrightnessDifference.Text);

            Microsoft.Office.Interop.PowerPoint.Application app = Globals.ThisAddIn.Application;
            if (app.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionShapes)
            {
                ShapeRange selectedShapes = app.ActiveWindow.Selection.ShapeRange;

                foreach (Shape groupShape in selectedShapes)
                {
                    // 检查是否为田字格对象
                    if (groupShape.Tags["IsTianZiGe"] == "True" ||
                        (groupShape.Type == Office.MsoShapeType.msoGroup && groupShape.GroupItems[1].Tags["IsTianZiGe"] == "True"))
                    {
                        // 更新田字格的外部边框
                        foreach (Shape shape in groupShape.GroupItems)
                        {
                            if (shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeRectangle)
                            {
                                shape.Line.Weight = borderWidth;
                                shape.Line.ForeColor.RGB = ConvertColor(borderColor);
                            }
                            else if (shape.Type == Office.MsoShapeType.msoLine)
                            {
                                // 更新田字格的内部边框
                                float innerBorderWidth = borderWidth - widthDifference;
                                shape.Line.Weight = innerBorderWidth;
                                shape.Line.ForeColor.RGB = AdjustColorBrightness(ConvertColor(borderColor), brightnessDifference);
                            }
                        }

                        // 删除并添加新的标签参数值
                        groupShape.Tags.Delete("BorderWidth");
                        groupShape.Tags.Add("BorderWidth", borderWidth.ToString());

                        groupShape.Tags.Delete("WidthDifference");
                        groupShape.Tags.Add("WidthDifference", widthDifference.ToString());

                        groupShape.Tags.Delete("BorderColor");
                        groupShape.Tags.Add("BorderColor", borderColor.ToString());

                        groupShape.Tags.Delete("BrightnessDifference");
                        groupShape.Tags.Add("BrightnessDifference", brightnessDifference.ToString());
                    }
                    else
                    {
                        // 忽略非田字格对象
                        continue;
                    }
                }
            }
        }
        private void ButtonRead_Click(object sender, RoutedEventArgs e)
        {
            // 检查当前选中的对象
            Microsoft.Office.Interop.PowerPoint.Application app = Globals.ThisAddIn.Application;
            if (app.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionShapes)
            {
                ShapeRange selectedShapes = app.ActiveWindow.Selection.ShapeRange;
                bool allAreTianZiGe = true;

                float? borderWidth = null;
                float? widthDifference = null;
                string borderColorString = null;
                float? brightnessDifference = null;

                int tianZiGeCount = 0;

                foreach (Shape groupShape in selectedShapes)
                {
                    if (groupShape.Tags["IsTianZiGe"] == "True" ||
                        (groupShape.Type == Office.MsoShapeType.msoGroup && groupShape.GroupItems[1].Tags["IsTianZiGe"] == "True"))
                    {
                        float currentBorderWidth = float.Parse(groupShape.Tags["BorderWidth"]);
                        float currentWidthDifference = float.Parse(groupShape.Tags["WidthDifference"]);
                        string currentBorderColorString = groupShape.Tags["BorderColor"];
                        float currentBrightnessDifference = float.Parse(groupShape.Tags["BrightnessDifference"]);

                        if (borderWidth == null)
                        {
                            // 初始化参数
                            borderWidth = currentBorderWidth;
                            widthDifference = currentWidthDifference;
                            borderColorString = currentBorderColorString;
                            brightnessDifference = currentBrightnessDifference;
                        }
                        else
                        {
                            // 检查参数是否一致
                            if (borderWidth != currentBorderWidth ||
                                widthDifference != currentWidthDifference ||
                                borderColorString != currentBorderColorString ||
                                brightnessDifference != currentBrightnessDifference)
                            {
                                allAreTianZiGe = false;
                                break;
                            }
                        }

                        tianZiGeCount++;
                    }
                }

                if (tianZiGeCount > 0 && allAreTianZiGe && borderWidth != null)
                {
                    // 同步参数到窗体
                    TextBoxBorderWidth.Text = borderWidth.Value.ToString();
                    TextBoxWidthDifference.Text = widthDifference.Value.ToString();
                    _ = new ColorConverter();
                    borderColor = (Color)ColorConverter.ConvertFromString(borderColorString);
                    ButtonChooseColor.Background = new SolidColorBrush(borderColor);

                    TextBoxBrightnessDifference.Text = brightnessDifference.Value.ToString();
                }
                else
                {
                    // 如果没有选中任何田字格对象或参数不一致，显示默认参数
                    if (tianZiGeCount == 0)
                    {
                        // 选中对象中没有田字格，显示默认参数
                        TextBoxBorderWidth.Text = "1.25";
                        TextBoxWidthDifference.Text = "0";
                        borderColor = Colors.Black;
                        ButtonChooseColor.Background = new SolidColorBrush(borderColor);
                        TextBoxBrightnessDifference.Text = "0";
                    }
                    else
                    {
                        MessageBox.Show("选中的田字格参数不一致，无法批量修改。");
                    }
                }
            }
        }


        private void ButtonIncrease_Click(object sender, RoutedEventArgs e)
        {
            AdjustBorderWidth(0.25m);
        }

        private void ButtonDecrease_Click(object sender, RoutedEventArgs e)
        {
            AdjustBorderWidth(-0.25m);
        }

        private void AdjustBorderWidth(decimal adjustment)
        {
            if (decimal.TryParse(TextBoxBorderWidth.Text, out decimal currentValue))
            {
                currentValue = Math.Max(0, currentValue + adjustment);
                TextBoxBorderWidth.Text = currentValue.ToString("0.00");
            }
        }

        private void GenerateShape()
        {
            if (!float.TryParse(TextBoxBorderWidth.Text, out float borderWidth))
            {
                MessageBox.Show("无效的边框宽度");
                return;
            }

            float widthDifference = float.Parse(TextBoxWidthDifference.Text);
            float brightnessDifference = float.Parse(TextBoxBrightnessDifference.Text);

            Microsoft.Office.Interop.PowerPoint.Application app = Globals.ThisAddIn.Application;
            Slide activeSlide = app.ActiveWindow.View.Slide as Slide;

            if (app.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionShapes)
            {
                ShapeRange selectedShapes = app.ActiveWindow.Selection.ShapeRange;

                List<string> newShapeNames = new List<string>();
                foreach (Shape selectedShape in selectedShapes)
                {
                    if (selectedShape.Type == Office.MsoShapeType.msoTextBox && selectedShape.TextFrame.TextRange.Text.Length > 1)
                    {
                        var splitShapeNames = SplitTextBoxIntoCharacters(activeSlide, selectedShape);
                        newShapeNames.AddRange(splitShapeNames);
                    }
                    else
                    {
                        newShapeNames.Add(selectedShape.Name);
                    }
                }

                selectedShapes = activeSlide.Shapes.Range(newShapeNames.ToArray());

                float initialLeft = selectedShapes[1].Left;
                float initialTop = selectedShapes[1].Top;
                float currentLeft = initialLeft;
                float currentTop = initialTop;
                float maxHeightInRow = 0;
                float rowStartTop = initialTop;
                float rowSpacing = 10;

                foreach (Shape selectedShape in selectedShapes)
                {
                    float selectedSize = Math.Max(selectedShape.Width, selectedShape.Height) + 18;
                    maxHeightInRow = Math.Max(maxHeightInRow, selectedSize);

                    if (Math.Abs(selectedShape.Top - rowStartTop) > 20)
                    {
                        currentLeft = initialLeft;
                        rowStartTop = selectedShape.Top;
                        currentTop += maxHeightInRow + rowSpacing;
                        maxHeightInRow = selectedSize;
                    }

                    float left = currentLeft;
                    float top = currentTop + (maxHeightInRow - selectedSize) / 2;

                    // 创建外部边框
                    Shape squareShape = activeSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, selectedSize, selectedSize);
                    squareShape.Line.Weight = borderWidth;
                    squareShape.Line.ForeColor.RGB = ConvertColor(borderColor);
                    squareShape.Fill.Transparency = 1;

                    // 创建内部边框
                    float innerBorderWidth = borderWidth - widthDifference;
                    Shape verticalLine = activeSlide.Shapes.AddLine(left + selectedSize / 2, top, left + selectedSize / 2, top + selectedSize);
                    Shape horizontalLine = activeSlide.Shapes.AddLine(left, top + selectedSize / 2, left + selectedSize, top + selectedSize / 2);

                    verticalLine.Line.Weight = innerBorderWidth;
                    horizontalLine.Line.Weight = innerBorderWidth;

                    verticalLine.Line.ForeColor.RGB = AdjustColorBrightness(ConvertColor(borderColor), brightnessDifference);
                    horizontalLine.Line.ForeColor.RGB = AdjustColorBrightness(ConvertColor(borderColor), brightnessDifference);

                    verticalLine.Line.DashStyle = Office.MsoLineDashStyle.msoLineDash;
                    horizontalLine.Line.DashStyle = Office.MsoLineDashStyle.msoLineDash;

                    // 先调整外部边框的层次，将其置于顶层
                    squareShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);

                    // 将所有形状组合成一个田字格
                    ShapeRange shapeRange = activeSlide.Shapes.Range(new string[] { squareShape.Name, verticalLine.Name, horizontalLine.Name });
                    Shape groupShape = shapeRange.Group();
                    groupShape.Name = $"田字格{tianZiGeCounter++}";
                    groupShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);

                    // 为组合形状添加标签
                    groupShape.Tags.Add("IsTianZiGe", "True");
                    groupShape.Tags.Add("BorderWidth", borderWidth.ToString());
                    groupShape.Tags.Add("WidthDifference", widthDifference.ToString());
                    groupShape.Tags.Add("BorderColor", borderColor.ToString());
                    groupShape.Tags.Add("BrightnessDifference", brightnessDifference.ToString());

                    // 确保每个子形状也带有相应的标签
                    foreach (Shape shape in groupShape.GroupItems)
                    {
                        shape.Tags.Add("IsTianZiGe", "True");
                        shape.Tags.Add("BorderWidth", borderWidth.ToString());
                        shape.Tags.Add("WidthDifference", widthDifference.ToString());
                        shape.Tags.Add("BorderColor", borderColor.ToString());
                        shape.Tags.Add("BrightnessDifference", brightnessDifference.ToString());
                    }

                    groupShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);

                    if (selectedShape.Type == Office.MsoShapeType.msoTextBox)
                    {
                        selectedShape.Width = selectedSize;
                        selectedShape.Height = selectedSize;
                        AdjustFontSizeToFit(selectedShape);

                        selectedShape.Left = left + (selectedSize - selectedShape.Width) / 2;
                        selectedShape.Top = top + (selectedSize - selectedShape.Height) / 2;
                    }
                    else
                    {
                        selectedShape.Left = left + (selectedSize - selectedShape.Width) / 2;
                        selectedShape.Top = top + (selectedSize - selectedShape.Height) / 2;
                    }

                    currentLeft += selectedSize;

                    groupShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                    selectedShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                }
            }
        }

        //田字格布局算法
        private void GenerateShapeWithoutLayout()
        {
            if (!float.TryParse(TextBoxBorderWidth.Text, out float borderWidth))
            {
                MessageBox.Show("无效的边框宽度");
                return;
            }

            float widthDifference = float.Parse(TextBoxWidthDifference.Text);
            float brightnessDifference = float.Parse(TextBoxBrightnessDifference.Text);

            Microsoft.Office.Interop.PowerPoint.Application app = Globals.ThisAddIn.Application;
            Slide activeSlide = app.ActiveWindow.View.Slide as Slide;

            if (app.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionShapes)
            {
                ShapeRange selectedShapes = app.ActiveWindow.Selection.ShapeRange;

                List<string> newShapeNames = new List<string>();
                foreach (Shape selectedShape in selectedShapes)
                {
                    if (selectedShape.Type == Office.MsoShapeType.msoTextBox && selectedShape.TextFrame.TextRange.Text.Length > 1)
                    {
                        var splitShapeNames = SplitTextBoxIntoCharacters(activeSlide, selectedShape);
                        newShapeNames.AddRange(splitShapeNames);
                    }
                    else
                    {
                        newShapeNames.Add(selectedShape.Name);
                    }
                }

                selectedShapes = activeSlide.Shapes.Range(newShapeNames.ToArray());

                foreach (Shape selectedShape in selectedShapes)
                {
                    float selectedSize = Math.Max(selectedShape.Width, selectedShape.Height) + 18;

                    float left = selectedShape.Left;
                    float top = selectedShape.Top;

                    // 创建外部边框
                    Shape squareShape = activeSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, selectedSize, selectedSize);
                    squareShape.Line.Weight = borderWidth;
                    squareShape.Line.ForeColor.RGB = ConvertColor(borderColor);
                    squareShape.Fill.Transparency = 1;

                    // 创建内部边框
                    float innerBorderWidth = borderWidth - widthDifference;
                    Shape verticalLine = activeSlide.Shapes.AddLine(left + selectedSize / 2, top, left + selectedSize / 2, top + selectedSize);
                    Shape horizontalLine = activeSlide.Shapes.AddLine(left, top + selectedSize / 2, left + selectedSize, top + selectedSize / 2);

                    verticalLine.Line.Weight = innerBorderWidth;
                    horizontalLine.Line.Weight = innerBorderWidth;

                    verticalLine.Line.ForeColor.RGB = AdjustColorBrightness(ConvertColor(borderColor), brightnessDifference);
                    horizontalLine.Line.ForeColor.RGB = AdjustColorBrightness(ConvertColor(borderColor), brightnessDifference);

                    verticalLine.Line.DashStyle = Office.MsoLineDashStyle.msoLineDash;
                    horizontalLine.Line.DashStyle = Office.MsoLineDashStyle.msoLineDash;

                    // 先调整外部边框的层次，将其置于顶层
                    squareShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);

                    // 将所有形状组合成一个田字格
                    ShapeRange shapeRange = activeSlide.Shapes.Range(new string[] { squareShape.Name, verticalLine.Name, horizontalLine.Name });
                    Shape groupShape = shapeRange.Group();

                    groupShape.Name = $"田字格{tianZiGeCounter++}";

                    // 添加标签，方便识别
                    groupShape.Tags.Add("IsTianZiGe", "True");
                    groupShape.Tags.Add("BorderWidth", borderWidth.ToString());
                    groupShape.Tags.Add("WidthDifference", widthDifference.ToString());
                    groupShape.Tags.Add("BorderColor", borderColor.ToString());
                    groupShape.Tags.Add("BrightnessDifference", brightnessDifference.ToString());

                    groupShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);

                    if (selectedShape.Type == Office.MsoShapeType.msoTextBox)
                    {
                        selectedShape.Width = selectedSize;
                        selectedShape.Height = selectedSize;
                        AdjustFontSizeToFit(selectedShape);

                        selectedShape.Left = left + (selectedSize - selectedShape.Width) / 2;
                        selectedShape.Top = top + (selectedSize - selectedShape.Height) / 2;
                    }
                    else
                    {
                        selectedShape.Left = left + (selectedSize - selectedShape.Width) / 2;
                        selectedShape.Top = top + (selectedSize - selectedShape.Height) / 2;
                    }

                    groupShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                    selectedShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                }
            }
        }

        private List<string> SplitTextBoxIntoCharacters(Slide slide, Shape textBox)
        {
            List<string> newShapeNames = new List<string>();
            string text = textBox.TextFrame.TextRange.Text;
            float left = textBox.Left;
            float top = textBox.Top;
            float spacing = 5f;

            float originalFontSize = textBox.TextFrame.TextRange.Font.Size;

            for (int i = 0; i < text.Length; i++)
            {
                string character = text[i].ToString();
                Shape charShape = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left, top, 10, 20);
                charShape.TextFrame.TextRange.Text = character;

                charShape.TextFrame.TextRange.Font.Size = originalFontSize;

                charShape.TextFrame.HorizontalAnchor = Office.MsoHorizontalAnchor.msoAnchorCenter;
                charShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                charShape.TextFrame.MarginLeft = 0;
                charShape.TextFrame.MarginRight = 0;
                charShape.TextFrame.MarginTop = 0;
                charShape.TextFrame.MarginBottom = 0;

                newShapeNames.Add(charShape.Name);

                left += charShape.Width + spacing;
            }

            textBox.Delete();
            return newShapeNames;
        }

        //使字号大小适应田字格大小
        private void AdjustFontSizeToFit(Shape textBox)
        {
            float maxSize = textBox.Width - 2;
            float fontSize = textBox.TextFrame.TextRange.Font.Size + 14;

            textBox.TextFrame.TextRange.Font.Size = fontSize;

            while (fontSize > 1 && textBox.TextFrame.TextRange.BoundWidth > maxSize)
            {
                fontSize -= 1;
                textBox.TextFrame.TextRange.Font.Size = fontSize;
            }
        }

        private void ButtonExit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }

        private int ConvertColor(Color color)
        {
            return (color.B << 16) | (color.G << 8) | color.R;
        }
        
        //调整颜色明亮度
        private int AdjustColorBrightness(int rgb, float brightnessDifference)
        {
            Color color = Color.FromRgb((byte)(rgb & 0xFF), (byte)((rgb >> 8) & 0xFF), (byte)((rgb >> 16) & 0xFF));
            ColorToHSL(color, out float h, out float s, out float l);

            l = Clamp(l + brightnessDifference, 0, 100);

            return HSLToColor(h, s, l);
        }

        //RGB转HSL
        private void ColorToHSL(Color color, out float h, out float s, out float l)
        {
            float r = color.R / 255f;
            float g = color.G / 255f;
            float b = color.B / 255f;

            float max = Math.Max(Math.Max(r, g), b);
            float min = Math.Min(Math.Min(r, g), b);

            h = 0f;
            s = 0f;
            l = (max + min) / 2f;

            if (max != min)
            {
                float delta = max - min;

                s = l > 0.5f ? delta / (2f - max - min) : delta / (max + min);

                if (max == r)
                {
                    h = (g - b) / delta + (g < b ? 6f : 0f);
                }
                else if (max == g)
                {
                    h = (b - r) / delta + 2f;
                }
                else if (max == b)
                {
                    h = (r - g) / delta + 4f;
                }

                h /= 6f;
            }

            h *= 360f;
            s *= 100f;
            l *= 100f;
        }

        //HSL转RGB
        private int HSLToColor(float h, float s, float l)
        {
            h /= 360f;
            s /= 100f;
            l /= 100f;

            float r, g, b;

            if (s == 0)
            {
                r = g = b = l; // achromatic
            }
            else
            {
                Func<float, float, float, float> hue2rgb = (p, q, t) =>
                {
                    if (t < 0) t += 1;
                    if (t > 1) t -= 1;
                    if (t < 1 / 6f) return p + (q - p) * 6f * t;
                    if (t < 1 / 2f) return q;
                    if (t < 2 / 3f) return p + (q - p) * (2 / 3f - t) * 6f;
                    return p;
                };

                float q1 = l < 0.5f ? l * (1 + s) : l + s - l * s;
                float p1 = 2 * l - q1;

                r = hue2rgb(p1, q1, h + 1 / 3f);
                g = hue2rgb(p1, q1, h);
                b = hue2rgb(p1, q1, h - 1 / 3f);
            }

            return ((int)(r * 255)) | ((int)(g * 255) << 8) | ((int)(b * 255) << 16);
        }

        private float Clamp(float value, float min, float max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }
    }
}
