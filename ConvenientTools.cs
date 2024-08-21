using HtmlAgilityPack;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using VBIDE = Microsoft.Vbe.Interop;

namespace 课件帮PPT助手
{
    public partial class DesignTools : UserControl
    {
        private ToolTip toolTip1;

        public DesignTools()
        {
            InitializeComponent();
            InitializeToolTips();
        }

        private void ToolTipsInitializeComponent()
        {
            this.笔画拆分 = new Button();
            this.分解笔顺 = new Button();
            this.挖词填空 = new Button();
            this.SuspendLayout();

            // 笔画拆分按钮
            this.笔画拆分.Location = new System.Drawing.Point(50, 130);
            this.笔画拆分.Name = "笔画拆分";
            this.笔画拆分.Size = new Size(100, 30);
            this.笔画拆分.TabIndex = 3;
            this.笔画拆分.Text = "笔画拆分";
            this.笔画拆分.UseVisualStyleBackColor = true;

            // 分解笔顺按钮
            this.分解笔顺.Location = new System.Drawing.Point(50, 170);
            this.分解笔顺.Name = "分解笔顺";
            this.分解笔顺.Size = new Size(100, 30);
            this.分解笔顺.TabIndex = 4;
            this.分解笔顺.Text = "分解笔顺";
            this.分解笔顺.UseVisualStyleBackColor = true;
            this.分解笔顺.MouseHover += new EventHandler(this.分解笔顺_MouseHover);

            // 挖词填空按钮
            this.挖词填空.Location = new System.Drawing.Point(50, 210);
            this.挖词填空.Name = "挖词填空";
            this.挖词填空.Size = new Size(100, 30);
            this.挖词填空.TabIndex = 5;
            this.挖词填空.Text = "挖词填空";
            this.挖词填空.UseVisualStyleBackColor = true;
            this.挖词填空.Click += new EventHandler(this.挖词填空_Click);
            this.挖词填空.MouseHover += new EventHandler(this.挖词填空_MouseHover);

            // DesignTools
            this.Controls.Add(this.笔画拆分);
            this.Controls.Add(this.分解笔顺);
            this.Controls.Add(this.挖词填空);
            this.Name = "DesignTools";
            this.Size = new Size(200, 300);
            this.ResumeLayout(false);
        }

        private void InitializeToolTips()
        {
            toolTip1 = new ToolTip
            {
                IsBalloon = false, // 不使用气泡形式显示提示
                AutoPopDelay = 5000, // 提示显示时间，单位为毫秒
                InitialDelay = 1000, // 鼠标悬停后显示提示的时间延迟，单位为毫秒
                ReshowDelay = 500, // 从一个控件移到另一个控件时，提示再次显示的时间延迟，单位为毫秒
                ShowAlways = true // 总是显示提示
            };

            // 设置多个按钮的ToolTip提示
            toolTip1.SetToolTip(this.笔画拆分, "选中文本拆分该字笔画。");

            string 分解笔顺ToolTipText = "选中文本可分解该字笔顺，默认按一行排列；\n按住Ctrl键单击，可按两行排列。";
            toolTip1.SetToolTip(this.分解笔顺, 分解笔顺ToolTipText);

            string 挖词填空ToolTipText = "①默认单击，使用划线挖空。\n" +
                                    "②按住Shift单击则使用括号挖空。";
                                    
            toolTip1.SetToolTip(this.挖词填空, 挖词填空ToolTipText);
        }

        private void 分解笔顺_MouseHover(object sender, EventArgs e)
        {
            Button button = sender as Button;
            toolTip1.Show(toolTip1.GetToolTip(button), button, 0, button.Height + 5, 5000); // 显示在按钮下方
        }

        private void 挖词填空_MouseHover(object sender, EventArgs e)
        {
            Button button = sender as Button;
            toolTip1.Show(toolTip1.GetToolTip(button), button, 0, button.Height + 5, 5000); // 显示在按钮下方
        }
    

         private void DesignTools_Load(object sender, EventArgs e)
        {
            // Any initialization code if necessary

        }

        private void 文字标注_Click(object sender, EventArgs e)
        {
            AnnotationToolForm form = new AnnotationToolForm();
            form.SetDefaultValues();
            form.AnnotationApplied += ApplyAnnotation; // 订阅事件
            form.Show();  // 以非模态方式显示表单
        }

        private void ApplyAnnotation(string annotationType, Color color, bool isBold, bool isItalic, Color highlightColor, Color textColor)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            Selection sel = app.ActiveWindow.Selection;

            if (sel.Type == PpSelectionType.ppSelectionText)
            {
                TextRange textRange = sel.TextRange;
                Slide slide = (Slide)app.ActiveWindow.View.Slide;

                // 获取所选文本的字体大小
                float baseFontSize = textRange.Font.Size;

                // Apply text properties
                textRange.Font.Bold = isBold ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                textRange.Font.Italic = isItalic ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                textRange.Font.Color.RGB = ColorTranslator.ToOle(textColor);

                float lineSpacing = textRange.ParagraphFormat.SpaceWithin; // 获取行间距
                float topTextAdjustment = lineSpacing <= 1.0f ? 0 : -(lineSpacing - 1) * 10;
                float topText = textRange.BoundTop + textRange.BoundHeight + topTextAdjustment;

                float leftText = textRange.BoundLeft;
                float widthText = textRange.BoundWidth;
                int charCount = textRange.Text.Length;

                // Add highlight by creating a rectangle behind the text
                if (highlightColor != Color.Empty)
                {
                    float left = textRange.BoundLeft;
                    float width = textRange.BoundWidth;
                    float height = textRange.BoundHeight;

                    // Adjust the height based on line spacing
                    float adjustedHeight = height * (lineSpacing < 1 ? 1 : 1 / lineSpacing);

                    // Position the rectangle to the top of the text with additional offset
                    float additionalOffset = 4; // 固定的偏移量，可以根据需要调整
                    float highlightTop = topText - adjustedHeight - 5 + additionalOffset;

                    var highlightRect = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, left, highlightTop, width, adjustedHeight);
                    highlightRect.Name = "Annotation_Highlight";
                    highlightRect.Fill.ForeColor.RGB = ColorTranslator.ToOle(highlightColor);
                    highlightRect.Line.Visible = MsoTriState.msoFalse;
                    highlightRect.ZOrder(MsoZOrderCmd.msoSendBackward);
                }

                if (annotationType.Contains("(") && annotationType.Contains(" - "))
                {
                    string[] parts = annotationType.Split(new string[] { "(", ")", " - " }, StringSplitOptions.RemoveEmptyEntries);
                    string name = parts[0].Trim();
                    string symbol = parts[1].Trim();
                    string position = parts[2].Trim();

                    string[] symbols = symbol.Split(',');

                    switch (position)
                    {
                        case "底部":
                            AddRepeatedSymbols(slide, leftText, topText + 5, widthText, color, symbols[0], charCount, baseFontSize);
                            break;
                        case "开头和末尾":
                            string startSymbol = symbols.Length > 0 ? symbols[0] : string.Empty;
                            string endSymbol = symbols.Length > 1 ? symbols[1] : string.Empty;
                            textRange.Text = $"{startSymbol}{textRange.Text}{endSymbol}";
                            break;
                        case "末尾":
                            textRange.Text = $"{textRange.Text}{symbols[0]}";
                            break;
                    }
                }
                else
                {
                    switch (annotationType)
                    {
                        case "横线":
                            AddLine(slide, leftText, topText, widthText, color);
                            break;
                        case "双横线":
                            AddLine(slide, leftText, topText, widthText, color);
                            AddLine(slide, leftText, topText + 3, widthText, color);
                            break;
                        case "波浪线":
                            AddWavyLine(slide, leftText, topText, widthText, color);
                            break;
                        case "重读符号":
                            AddRepeatedSymbols(slide, leftText, topText, widthText, color, "●", charCount, baseFontSize);
                            break;
                        case "轻读符号":
                            AddRepeatedSymbols(slide, leftText, topText, widthText, color, "○", charCount, baseFontSize);
                            break;
                        case "着重符号":
                            AddRepeatedSymbols(slide, leftText, topText, widthText, color, "▲", charCount, baseFontSize);
                            break;
                        case "大括号":
                            textRange.Text = "[" + textRange.Text + "]";
                            break;
                        case "层级符":
                            textRange.Text = textRange.Text + "/";
                            break;
                        case "段落符":
                            textRange.Text = textRange.Text + "//";
                            break;
                    }
                }
            }
        }

        private void AddLine(Slide slide, float left, float top, float width, Color color)
        {
            var line = slide.Shapes.AddLine(left, top, left + width, top);
            line.Name = "Annotation_Line";
            line.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
            line.Line.Weight = 1.5f;
        }

        private void AddWavyLine(Slide slide, float left, float top, float width, Color color)
        {
            float step = 5f; // 每个波浪的宽度
            float amplitude = 2f; // 波浪的高度

            List<float> points = new List<float>();

            // 计算波浪线的点
            bool goingUp = true;
            for (float x = left; x <= left + width; x += step)
            {
                points.Add(x);
                points.Add(top + (goingUp ? amplitude : -amplitude));
                goingUp = !goingUp;
            }

            // 确保波浪线在结束时回到中间水平线
            points.Add(left + width);
            points.Add(top);

            // 将计算的点转换为 PowerPoint 安全数组
            Array safeArray = Array.CreateInstance(typeof(float), new int[] { points.Count / 2, 2 }, new int[] { 1, 1 });
            for (int i = 0; i < points.Count; i += 2)
            {
                safeArray.SetValue(points[i], i / 2 + 1, 1);
                safeArray.SetValue(points[i + 1], i / 2 + 1, 2);
            }

            // 创建波浪线形状
            PowerPoint.Shape waveShape = slide.Shapes.AddPolyline(safeArray);
            waveShape.Name = "Annotation_WavyLine";
            waveShape.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
            waveShape.Line.Weight = 1.5f;
            waveShape.Line.DashStyle = MsoLineDashStyle.msoLineSolid;
            waveShape.Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadNone;
            waveShape.Line.BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadNone;
        }

        private void AddRepeatedSymbols(Slide slide, float left, float top, float width, Color color, string symbol, int count, float baseFontSize)
        {
            float step = width / count;
            float symbolFontSize = baseFontSize / 3; // 动态调整符号字号

            for (int i = 0; i < count; i++)
            {
                var textBox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, left + i * step, top, step, symbolFontSize);
                textBox.Name = "Annotation_Symbol_" + i;
                var textRange = textBox.TextFrame.TextRange;
                textRange.Text = symbol;
                textRange.Font.Color.RGB = ColorTranslator.ToOle(color);
                textRange.Font.Size = symbolFontSize; // 设置动态符号字号
                textBox.TextFrame.HorizontalAnchor = Office.MsoHorizontalAnchor.msoAnchorCenter;
                textBox.TextFrame.MarginLeft = 0;
                textBox.TextFrame.MarginRight = 0;
                textBox.TextFrame.MarginTop = 0;
                textBox.TextFrame.MarginBottom = 0;
            }
        }






        private void 书写动画_Click(object sender, EventArgs ev)
        {
            AnimationForm animationForm = new AnimationForm();
            animationForm.Show();
        }


        private void 汉字字典_Click(object sender, EventArgs e)
        {
            HanziDictionaryForm dictionaryForm = new HanziDictionaryForm();
            dictionaryForm.Show();
        }

        private void 字源字形_Click(object sender, EventArgs e)
        {
            string inputChar = Microsoft.VisualBasic.Interaction.InputBox("请输入目标汉字（需联网，且一次仅支持查询单个汉字）:", "一键获取汉字字源字形图", "");
            if (!string.IsNullOrWhiteSpace(inputChar))
            {
                string url = $"https://www.zdic.net/hans/{inputChar}";
                ExtractSVGFromWebpage(url, inputChar);
            }
        }

        private void ExtractSVGFromWebpage(string url, string inputChar)
        {
            try
            {
                HtmlWeb web = new HtmlWeb();
                HtmlAgilityPack.HtmlDocument doc = web.Load(url);
                var sectionNode = doc.DocumentNode.SelectSingleNode("//div[@class='content definitions znr']");
                if (sectionNode != null)
                {
                    var headerNodes = sectionNode.SelectNodes(".//table[@class='zyyb']//tr[1]//td[not(@style='display:none')]");
                    var rowNodes = sectionNode.SelectNodes(".//table[@class='zyyb']//tr[position()>1]");
                    if (rowNodes != null && headerNodes != null)
                    {
                        List<List<KeyValuePair<string, string>>> svgMatrix = new List<List<KeyValuePair<string, string>>>();
                        List<string> headers = new List<string>();

                        foreach (var headerNode in headerNodes)
                        {
                            headers.Add(headerNode.InnerText.Trim());
                        }

                        foreach (var rowNode in rowNodes)
                        {
                            var cellNodes = rowNode.SelectNodes(".//td[not(@style='display:none')]");
                            List<KeyValuePair<string, string>> rowList = new List<KeyValuePair<string, string>>();

                            for (int i = 0; i < headers.Count; i++)
                            {
                                if (cellNodes != null && i < cellNodes.Count)
                                {
                                    var imgNode = cellNodes[i].SelectSingleNode(".//img[@class='lazy ypic']");
                                    if (imgNode != null)
                                    {
                                        string svgUrl = imgNode.GetAttributeValue("data-original", string.Empty);
                                        if (string.IsNullOrWhiteSpace(svgUrl))
                                        {
                                            svgUrl = imgNode.GetAttributeValue("src", string.Empty);
                                        }
                                        if (!svgUrl.StartsWith("http"))
                                        {
                                            svgUrl = "https:" + svgUrl;
                                        }
                                        var descriptionNodes = imgNode.ParentNode.SelectNodes("span");
                                        string description = descriptionNodes != null ? string.Join(" ", descriptionNodes.Select(node => node.InnerText.Trim())) : "";
                                        string header = headers[i];
                                        rowList.Add(new KeyValuePair<string, string>($"{header}|{description}", svgUrl));
                                    }
                                    else
                                    {
                                        rowList.Add(new KeyValuePair<string, string>("", "")); // 占位
                                    }
                                }
                                else
                                {
                                    rowList.Add(new KeyValuePair<string, string>("", "")); // 占位
                                }
                            }
                            svgMatrix.Add(rowList);
                        }
                        ShowSVGSelectionWindow(svgMatrix, inputChar, headers);
                    }
                    else
                    {

                    }
                }
                else
                {
                    MessageBox.Show("未查询到字源字形部分！");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("出现错误：" + ex.Message);
            }
        }

        private void ShowSVGSelectionWindow(List<List<KeyValuePair<string, string>>> svgMatrix, string inputChar, List<string> headers)
        {
            SVGSettingsForm svgSelectionForm = new SVGSettingsForm(inputChar, svgMatrix, headers);
            svgSelectionForm.ShowDialog();
        }

        private void 分解笔顺_Click(object sender, EventArgs e)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                Selection sel = app.ActiveWindow.Selection;

                // 确定选中的文本
                string selectedText = null;
                PowerPoint.Shape selectedShape = null;  // 定义 selectedShape 变量

                if (sel.Type == PpSelectionType.ppSelectionText || sel.Type == PpSelectionType.ppSelectionShapes)
                {
                    TextRange textRange = null;
                    if (sel.Type == PpSelectionType.ppSelectionText)
                    {
                        textRange = sel.TextRange;
                    }
                    else if (sel.Type == PpSelectionType.ppSelectionShapes)
                    {
                        if (sel.ShapeRange.Count == 1 && sel.ShapeRange[1].HasTextFrame == MsoTriState.msoTrue && sel.ShapeRange[1].TextFrame.HasText == MsoTriState.msoTrue)
                        {
                            textRange = sel.ShapeRange[1].TextFrame.TextRange;
                            selectedShape = sel.ShapeRange[1];  // 赋值给 selectedShape
                        }
                    }

                    if (textRange != null)
                    {
                        selectedText = textRange.Text.Trim();
                        if (selectedText.Length != 1)
                        {
                            MessageBox.Show("请选择包含一个汉字的文本框。");
                            return;
                        }
                    }
                }

                if (string.IsNullOrEmpty(selectedText) || selectedShape == null)
                {
                    MessageBox.Show("请选择包含一个汉字的文本框。");
                    return;
                }

                Slide slide = app.ActiveWindow.View.Slide;

                // 获取选中文本框的初始位置
                float initialLeft = selectedShape.Left;
                float initialTop = selectedShape.Top + selectedShape.Height; // 计算底部位置

                // 添加标识符“&”到当前幻灯片上符合条件的形状
                List<PowerPoint.Shape> shapesToRename = new List<PowerPoint.Shape>();
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Name.StartsWith($"{selectedText}-笔画"))
                    {
                        shape.Name = $"&{shape.Name}";
                        shapesToRename.Add(shape);
                    }
                }

                // 调用笔画拆分的点击事件
                笔画拆分_Click(sender, e);

                // 收集由“笔画拆分”事件产生的形状
                List<PowerPoint.Shape> shapesToGroup = new List<PowerPoint.Shape>();
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Name.StartsWith($"{selectedText}-笔画") && !shape.Name.StartsWith("&"))
                    {
                        shape.Name = "※" + shape.Name;
                        shapesToGroup.Add(shape);
                    }
                }

                // 对这些形状进行组合
                if (shapesToGroup.Count > 0)
                {
                    PowerPoint.ShapeRange shapeRange = slide.Shapes.Range(shapesToGroup.Select(s => s.Name).ToArray());
                    PowerPoint.Shape groupShape = null;

                    try
                    {
                        groupShape = shapeRange.Group();
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("您之前使用笔画拆分的汉字和现在要分解笔顺的汉字相同，导致程序无法判断该分解哪个笔画，您可以先对拆分的笔画进行批量命名或者组合或直接使用旧版的分解笔顺功能。");
                        return;
                    }

                    // 删除组合中子形状图层名称的字符“※”
                    foreach (PowerPoint.Shape shape in groupShape.GroupItems)
                    {
                        if (shape.Name.StartsWith("※"))
                        {
                            shape.Name = shape.Name.Substring(1);
                        }
                    }

                    // 手动缩放编组后的形状
                    float scaleFactor = 0.5f;
                    groupShape.Width *= scaleFactor;
                    groupShape.Height *= scaleFactor;

                    // 继续执行分解笔顺的逻辑
                    if (groupShape.Type == MsoShapeType.msoGroup)
                    {
                        PowerPoint.GroupShapes groupItems = groupShape.GroupItems;
                        int itemCount = groupItems.Count;

                        // Create new groups based on the number of items in the original group
                        List<PowerPoint.Shape> newGroups = new List<PowerPoint.Shape>();
                        for (int i = 0; i < itemCount; i++)
                        {
                            // Duplicate the original group
                            PowerPoint.Shape newGroup = groupShape.Duplicate()[1];
                            newGroup.Left += (i + 1) * (groupShape.Width + 10); // Adjust position
                            newGroup.Tags.Add("FenBu", "True");  // 添加 "FenBu" Tag
                            newGroups.Add(newGroup);
                        }

                        // Check if Ctrl key is pressed
                        bool isCtrlPressed = (ModifierKeys & Keys.Control) == Keys.Control;

                        // Set colors and remove borders based on the pattern
                        for (int i = 0; i < newGroups.Count; i++)
                        {
                            PowerPoint.Shape newGroup = newGroups[i];
                            PowerPoint.GroupShapes newGroupItems = newGroup.GroupItems;

                            for (int j = 1; j <= itemCount; j++)
                            {
                                newGroupItems[j].Line.Visible = MsoTriState.msoFalse; // Remove border

                                if (j <= i + 1)
                                {
                                    newGroupItems[j].Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);
                                }
                                if (j == i + 1)
                                {
                                    newGroupItems[j].Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Red);
                                }
                                if (j > i + 1)
                                {
                                    newGroupItems[j].Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray);
                                }
                            }

                            // 命名新组合
                            newGroup.Name = $"【{selectedText}】：分步第{i + 1}笔";

                            // Adjust positions for two-row layout if Ctrl key is pressed
                            if (isCtrlPressed)
                            {
                                int columns = (int)Math.Ceiling(newGroups.Count / 2.0);
                                int row = i / columns;
                                int column = i % columns;

                                newGroup.Left = initialLeft + column * (groupShape.Width + 10);
                                newGroup.Top = initialTop + row * (groupShape.Height + 10); // Adjust the top position
                            }
                        }

                        // 删除原来的组合形状
                        groupShape.Delete();

                        // 收集所有带有 "FenBu" Tag 的形状
                        List<PowerPoint.Shape> finalGroupsToAlign = newGroups.Where(g => g.Tags["FenBu"] == "True").ToList();

                        // 确保不影响其他已存在的形状
                        if (finalGroupsToAlign.Count > 0)
                        {
                            PowerPoint.ShapeRange newShapeRange = slide.Shapes.Range(finalGroupsToAlign.Select(s => s.Name).ToArray());
                            PowerPoint.Shape newGroupShape = newShapeRange.Group();

                            // 对新组合执行水平位置对齐
                            newGroupShape.Left = initialLeft;
                            newGroupShape.Top = initialTop;  // Set the top position

                            // 取消组合
                            newGroupShape.Ungroup();

                            // 删除 "FenBu" Tag
                            foreach (PowerPoint.Shape shape in finalGroupsToAlign)
                            {
                                shape.Tags.Delete("FenBu");
                            }
                        }
                    }
                }

                // 删除标识符“&”以恢复原始形状名称
                foreach (PowerPoint.Shape shape in shapesToRename)
                {
                    if (shape.Name.StartsWith($"&{selectedText}-笔画"))
                    {
                        shape.Name = shape.Name.Substring(1);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"发生错误：{ex.Message}");
            }
        }

        private void 挖词填空_Click(object sender, EventArgs e)
        {
            // 获取当前幻灯片
            var slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            // 获取选中的文本框
            if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionText)
            {
                var textRange = Globals.ThisAddIn.Application.ActiveWindow.Selection.TextRange;

                // 获取选中的文本
                string selectedText = textRange.Text;

                // 如果有选中的文本
                if (!string.IsNullOrEmpty(selectedText))
                {
                    // 获取原文本框的属性
                    var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                    var originalShape = selection.ShapeRange[1];

                    // 获取选中文字的位置和大小
                    float originalLeft = textRange.BoundLeft;
                    float originalTop = textRange.BoundTop;

                    // 测量选中文本的宽度
                    float textWidth = MeasureTextWidth(selectedText, textRange.Font.Size, textRange.Font.Name);

                    // 创建一个新的文本框，并设置其内容为选中的文本
                    var newTextBox = slide.Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        originalLeft, originalTop, textWidth, originalShape.Height);

                    var newTextRange = newTextBox.TextFrame.TextRange;
                    newTextRange.Text = selectedText;

                    // 设置新文本框的字体大小与选中文本一致
                    float originalFontSize = textRange.Font.Size;
                    newTextRange.Font.Size = originalFontSize;

                    ApplyFormatting(originalShape, newTextBox);

                    // 再次设置新文本框的字体大小，确保不被格式刷覆盖
                    newTextRange.Font.Size = originalFontSize;

                    // 确保新文本框不带有PPT自带的下划线
                    newTextRange.Font.Underline = MsoTriState.msoFalse;

                    // 设置新文本框的字体颜色为红色并加粗
                    newTextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Red);
                    newTextRange.Font.Bold = MsoTriState.msoTrue;

                    // 设置新文本框不自动换行
                    newTextBox.TextFrame.WordWrap = MsoTriState.msoFalse;

                    // 设置文本框的文本对齐方式为居中
                    newTextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;

                    // 添加“擦除”动画，并设置在“单击”时播放
                    PowerPoint.TimeLine timeLine = slide.TimeLine;
                    PowerPoint.Sequence sequence = timeLine.MainSequence;
                    PowerPoint.Effect effect = sequence.AddEffect(newTextBox, MsoAnimEffect.msoAnimEffectFade, (MsoAnimateByLevel)MsoAnimTriggerType.msoAnimTriggerOnShapeClick);
                    effect.Timing.Duration = 0.5f; // 动画持续时间为2秒

                    // 初始化 newLeft, newTop 和 newWidth 变量
                    float newLeft;
                    float newTop;
                    float newWidth;

                    // 检查是否按住Shift键
                    PowerPoint.Shape rectangle = null;
                    if ((ModifierKeys & Keys.Shift) == Keys.Shift)
                    {
                        HandleShiftKeyPress(textRange, textWidth, out newLeft, out newTop, out newWidth, out rectangle);
                        newTextBox.Width = newWidth; // 设置新文本框的宽度
                    }
                    else
                    {
                        // 默认情况下的文本框位置调整
                        HandlePress(textRange, textWidth, originalLeft, originalTop, out newLeft, out newTop);
                    }

                    // 应用新的位置
                    newTextBox.Left = newLeft;
                    newTextBox.Top = newTop;

                    // 删除辅助矩形
                    rectangle?.Delete();
                }
                else
                {
                    MessageBox.Show("请选中文本框内的文本！");
                }
            }
            else
            {
                MessageBox.Show("请选中文本框内的文本！");
            }
        }

        private void ApplyFormatting(PowerPoint.Shape originalShape, PowerPoint.Shape newTextBox)
        {
            try
            {
                // 应用格式刷
                originalShape.PickUp();
                newTextBox.Apply();

                // 获取新文本框的段落格式
                var paragraphFormat = newTextBox.TextFrame2.TextRange.ParagraphFormat;

                // 重置缩进和间距，确保没有首行缩进
                paragraphFormat.LeftIndent = 0;   // 左缩进
                paragraphFormat.RightIndent = 0;  // 右缩进
                paragraphFormat.FirstLineIndent = 0; // 首行缩进

                // 重置段前段后的间距
                paragraphFormat.SpaceBefore = 0; // 段前间距
                paragraphFormat.SpaceAfter = 0;  // 段后间距

                // 忽略字号大小的复制，保持新文本框的字号不变
                newTextBox.TextFrame.TextRange.Font.Size = originalShape.TextFrame.TextRange.Font.Size;
            }
            catch (Exception ex)
            {
                MessageBox.Show("使用格式刷复制字体属性时出错：" + ex.Message);
            }
        }


        private void HandlePress(TextRange textRange, float textWidth, float originalLeft, float originalTop, out float newLeft, out float newTop)
        {
            bool containsChinese = textRange.Text.Any(c => c >= 0x4e00 && c <= 0x9fff);

            string underlineText;
            if (containsChinese)
            {
                underlineText = new string('\u3000', textRange.Text.Length);
            }
            else
            {
                float spaceWidth = MeasureTextWidth(" ", textRange.Font.Size, textRange.Font.Name);
                int numSpaces = (int)Math.Ceiling(textWidth / spaceWidth) + 2;
                underlineText = new string(' ', numSpaces);
            }

            textRange.Text = underlineText;
            textRange.Font.Underline = MsoTriState.msoTrue;

            newLeft = originalLeft - 7;
            newTop = originalTop - 6;
        }

        private void HandleShiftKeyPress(TextRange textRange, float textWidth, out float newLeft, out float newTop, out float newWidth, out PowerPoint.Shape rectangle)
        {
            bool containsChinese = textRange.Text.Any(c => c >= 0x4e00 && c <= 0x9fff);

            string leftBracket = containsChinese ? "（" : "(";
            string rightBracket = containsChinese ? "）" : ")";
            string placeholders;

            if (containsChinese)
            {
                placeholders = new string('\u3000', textRange.Text.Length);
            }
            else
            {
                float placeholderWidth = MeasureTextWidth(" ", textRange.Font.Size, textRange.Font.Name);
                int placeholderCount = (int)Math.Ceiling(textWidth / placeholderWidth) + 2;
                placeholders = new string('\u0020', placeholderCount);
            }

            textRange.Text = $"{leftBracket}{placeholders}{rightBracket}";

            float selectionLeft = textRange.BoundLeft;
            float selectionTop = textRange.BoundTop;
            float selectionWidth = textRange.BoundWidth;
            float selectionHeight = textRange.BoundHeight;

            var slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            rectangle = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, selectionLeft, selectionTop, selectionWidth, selectionHeight);

            rectangle.Line.Visible = MsoTriState.msoFalse;
            rectangle.Fill.Transparency = 1.0f;

            newLeft = rectangle.Left;
            newTop = rectangle.Top - 3;
            newWidth = rectangle.Width;
        }

        private float MeasureTextWidth(string text, float fontSize, string fontName)
        {
            using (var bmp = new Bitmap(1, 1))
            {
                using (var g = Graphics.FromImage(bmp))
                {
                    var font = new System.Drawing.Font(fontName, fontSize);
                    var size = g.MeasureString(text, font);
                    return size.Width;
                }
            }
        }

         public void 笔画拆分_Click(object sender, EventArgs e)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                Selection sel = app.ActiveWindow.Selection;

                if (sel.Type == PpSelectionType.ppSelectionText || sel.Type == PpSelectionType.ppSelectionShapes)
                {
                    TextRange textRange = null;

                    if (sel.Type == PpSelectionType.ppSelectionText)
                    {
                        textRange = sel.TextRange;
                    }
                    else if (sel.Type == PpSelectionType.ppSelectionShapes)
                    {
                        if (sel.ShapeRange.Count == 1 && sel.ShapeRange[1].HasTextFrame == MsoTriState.msoTrue && sel.ShapeRange[1].TextFrame.HasText == MsoTriState.msoTrue)
                        {
                            textRange = sel.ShapeRange[1].TextFrame.TextRange;
                        }
                    }

                    if (textRange != null)
                    {
                        string selectedText = textRange.Text.Trim();

                        if (selectedText.Length == 1)
                        {
                            string svgContent = GetSVGContent(selectedText);

                            if (!string.IsNullOrEmpty(svgContent))
                            {
                                Slide slide = app.ActiveWindow.View.Slide;
                                PowerPoint.Shape svgShape = InsertSVGIntoSlide(svgContent, slide);
                                SelectShape(app, svgShape);
                                AddAndRunVBA(app, selectedText);
                            }
                            else
                            {
                                MessageBox.Show("未找到对应的SVG文件。");
                            }
                        }
                        else
                        {
                            MessageBox.Show("请选择包含一个汉字的文本框。");
                        }
                    }
                    else
                    {
                        MessageBox.Show("请选择包含一个汉字的文本框。");
                    }
                }
                else
                {
                    MessageBox.Show("请选择一个文本框。");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"发生错误：{ex.Message}");
            }
        }

        private string GetSVGContent(string character)
        {
            string resourceName = $"课件帮PPT助手.汉字笔画.{character}.svg";
            Assembly assembly = Assembly.GetExecutingAssembly();

            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream != null)
                {
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            return null;
        }

        private PowerPoint.Shape InsertSVGIntoSlide(string svgContent, Slide slide)
        {
            string tempSvgPath = Path.Combine(Path.GetTempPath(), "temp.svg");
            File.WriteAllText(tempSvgPath, svgContent);

            float left = 100;  // 可以根据需求调整
            float top = 100;   // 可以根据需求调整

            PowerPoint.Shape svgShape = slide.Shapes.AddPicture(tempSvgPath, MsoTriState.msoFalse, MsoTriState.msoCTrue, left, top);

            // 放大SVG
            svgShape.Width *= 1.5f;
            svgShape.Height *= 1.5f;

            File.Delete(tempSvgPath);

            return svgShape;
        }

        private void SelectShape(PowerPoint.Application app, PowerPoint.Shape shape)
        {
            shape.Select();
        }

        private void AddAndRunVBA(PowerPoint.Application app, string svgFileName)
        {
            string guid = Guid.NewGuid().ToString("N"); // 生成唯一的GUID标识符
            string vbaCode = $@"
Sub ConvertSVGToShape()
    ' Ensure a shape is selected
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox ""Please select an SVG shape to convert."", vbExclamation
        Exit Sub
    End If
    
    Dim shp As Shape
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    
    ' Convert the SVG to a shape by copying and pasting it as an EMF
    shp.Copy
    
    Dim slide As slide
    Set slide = ActiveWindow.View.slide
    Dim newShape As Shape
    Set newShape = slide.Shapes.PasteSpecial(DataType:=ppPasteEnhancedMetafile)(1)
    
    ' Delete the original SVG shape
    shp.Delete
    
    ' Ungroup the new shape multiple times to fully convert it to individual shapes
    On Error Resume Next
    Dim i As Integer
    For i = 1 To 5
        newShape.Ungroup
        Set newShape = slide.Shapes(slide.Shapes.Count) ' Re-select the shape after ungrouping
    Next i
    
    ' Find and delete the shape named ""AutoShape""
    Dim shapeItem As Shape
    For Each shapeItem In slide.Shapes
        If InStr(shapeItem.Name, ""AutoShape"") > 0 Then
            shapeItem.Delete
        End If
    Next shapeItem
    
    ' Remove the border from each ungrouped shape and rename them, but without changing the naming pattern
    Dim count As Integer
    count = 1
    For Each shapeItem In slide.Shapes
        If Left(shapeItem.Name, 8) = ""Freeform"" Then
            ' Assign the name as per your naming pattern
            shapeItem.Name = ""{svgFileName}-笔画"" & count
            ' Add a tag to uniquely identify this batch of shapes
            shapeItem.Tags.Add ""BatchIdentifier"", ""{guid}""
            ' Remove the border
            shapeItem.Line.Visible = msoFalse
            count = count + 1
        End If
    Next shapeItem
End Sub";

            VBIDE.VBProject vbProject = app.ActivePresentation.VBProject;
            VBIDE.VBComponent vbModule = vbProject.VBComponents.Add(VBIDE.vbext_ComponentType.vbext_ct_StdModule);
            vbModule.CodeModule.AddFromString(vbaCode);

            app.Run("ConvertSVGToShape");

            vbProject.VBComponents.Remove(vbModule);
        }


        public void PerformStrokeSplit(string selectedText)
        {
            笔画拆分_Click(this, EventArgs.Empty);
        }
    }
}
















