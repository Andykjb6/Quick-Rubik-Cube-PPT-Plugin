using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using HtmlAgilityPack;
using System.IO;
using System.Text.RegularExpressions;
using TheArtOfDev.HtmlRenderer.WinForms;
using System.Net;
using Newtonsoft.Json;
using System.Runtime.InteropServices;
using System.Reflection;
using Microsoft.Office.Core;
using VBIDE = Microsoft.Vbe.Interop;
using Microsoft.Office.Interop.PowerPoint;

namespace 课件帮PPT助手
{
    public partial class DesignTools : UserControl
    {
       

        public DesignTools()
        {
            InitializeComponent();
            
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
            PowerPoint.Selection sel = app.ActiveWindow.Selection;

            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
            {
                PowerPoint.TextRange textRange = sel.TextRange;
                PowerPoint.Slide slide = (PowerPoint.Slide)app.ActiveWindow.View.Slide;

                // 获取所选文本的字体大小
                float baseFontSize = textRange.Font.Size;

                // Apply text properties
                textRange.Font.Bold = isBold ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                textRange.Font.Italic = isItalic ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
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

                    var highlightRect = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, highlightTop, width, adjustedHeight);
                    highlightRect.Name = "Annotation_Highlight";
                    highlightRect.Fill.ForeColor.RGB = ColorTranslator.ToOle(highlightColor);
                    highlightRect.Line.Visible = Office.MsoTriState.msoFalse;
                    highlightRect.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
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

        private void AddLine(PowerPoint.Slide slide, float left, float top, float width, Color color)
        {
            var line = slide.Shapes.AddLine(left, top, left + width, top);
            line.Name = "Annotation_Line";
            line.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
            line.Line.Weight = 1.5f;
        }

        private void AddWavyLine(PowerPoint.Slide slide, float left, float top, float width, Color color)
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
            waveShape.Line.DashStyle = Office.MsoLineDashStyle.msoLineSolid;
            waveShape.Line.EndArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadNone;
            waveShape.Line.BeginArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadNone;
        }

        private void AddRepeatedSymbols(PowerPoint.Slide slide, float left, float top, float width, Color color, string symbol, int count, float baseFontSize)
        {
            float step = width / count;
            float symbolFontSize = baseFontSize / 3; // 动态调整符号字号

            for (int i = 0; i < count; i++)
            {
                var textBox = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left + i * step, top, step, symbolFontSize);
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


        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void DesignTools_Load_1(object sender, EventArgs e)
        {

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
            // 调用笔画拆分的点击事件
            笔画拆分_Click(sender, e);

            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

            // 收集所有前缀名为“Freeform”且填充颜色为“0,112,192”的形状
            List<PowerPoint.Shape> shapesToGroup = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                if (shape.Name.StartsWith("Freeform") && shape.Fill.ForeColor.RGB == System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 112, 192)))
                {
                    shapesToGroup.Add(shape);
                }
            }

            // 对这些形状进行组合
            if (shapesToGroup.Count > 0)
            {
                PowerPoint.ShapeRange shapeRange = slide.Shapes.Range(shapesToGroup.Select(s => s.Name).ToArray());
                PowerPoint.Shape groupShape = shapeRange.Group();

                // 手动缩放编组后的形状
                float scaleFactor = 0.26f;
                groupShape.Width *= scaleFactor;
                groupShape.Height *= scaleFactor;

                // 继续执行分解笔顺的逻辑
                if (groupShape.Type == Office.MsoShapeType.msoGroup)
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
                        newGroups.Add(newGroup);
                    }

                    // Check if Ctrl key is pressed
                    bool isCtrlPressed = (Control.ModifierKeys & Keys.Control) == Keys.Control;

                    // Set colors and remove borders based on the pattern
                    for (int i = 0; i < newGroups.Count; i++)
                    {
                        PowerPoint.Shape newGroup = newGroups[i];
                        PowerPoint.GroupShapes newGroupItems = newGroup.GroupItems;

                        for (int j = 1; j <= itemCount; j++)
                        {
                            newGroupItems[j].Line.Visible = Office.MsoTriState.msoFalse; // Remove border

                            if (j <= i + 1)
                            {
                                newGroupItems[j].Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                            }
                            if (j == i + 1)
                            {
                                newGroupItems[j].Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                            if (j > i + 1)
                            {
                                newGroupItems[j].Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                            }
                        }

                        // Adjust positions for two-row layout if Ctrl key is pressed
                        if (isCtrlPressed)
                        {
                            int columns = (int)Math.Ceiling(newGroups.Count / 2.0);
                            int row = i / columns;
                            int column = i % columns;

                            newGroup.Left = groupShape.Left + column * (groupShape.Width + 10);
                            newGroup.Top = groupShape.Top + row * (groupShape.Height + 10);
                        }
                    }

                    // 删除原来的组合形状
                    groupShape.Delete();

                    // 将所有新组合再次进行组合
                    PowerPoint.ShapeRange newShapeRange = slide.Shapes.Range(newGroups.Select(s => s.Name).ToArray());
                    PowerPoint.Shape newGroupShape = newShapeRange.Group();

                    // 对新组合执行水平居中对齐
                    float slideCenter = slide.Master.Width / 2;
                    newGroupShape.Left = slideCenter - newGroupShape.Width / 2;

                    // 取消组合
                    newGroupShape.Ungroup();
                }
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
                        Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                        originalLeft, originalTop, textWidth, originalShape.Height);

                    var newTextRange = newTextBox.TextFrame.TextRange;
                    newTextRange.Text = selectedText;

                    // 使用格式刷复制字体属性
                    try
                    {
                        originalShape.PickUp();
                        newTextBox.Apply();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("使用格式刷复制字体属性时出错：" + ex.Message);
                    }

                    // 设置新文本框的字体颜色为红色并加粗
                    newTextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Red);
                    newTextRange.Font.Bold = MsoTriState.msoTrue;

                    // 设置新文本框不自动换行
                    newTextBox.TextFrame.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;

                    // 检查是否按住Ctrl键或Shift键
                    if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
                    {
                        // 动态计算需要的空格字符数量
                        float spaceWidth = MeasureTextWidth(" ", textRange.Font.Size, textRange.Font.Name);
                        int numSpaces = (int)Math.Ceiling((textWidth * 1.1) / spaceWidth); // 1.1倍宽度以确保足够长
                        string underlineText = new string(' ', numSpaces);

                        // 确保下划线长度适中
                        textRange.Text = underlineText;
                        textRange.Font.Underline = MsoTriState.msoTrue;
                    }
                    else if ((Control.ModifierKeys & Keys.Shift) == Keys.Shift)
                    {
                        // 动态计算需要的空格字符数量
                        float spaceWidth = MeasureTextWidth(" ", textRange.Font.Size, textRange.Font.Name);
                        int numSpaces = (int)Math.Ceiling(textWidth / spaceWidth);
                        string spaces = new string(' ', numSpaces);

                        // 使用括号和空格替换选中的文本
                        textRange.Text = $"({spaces})";
                        textRange.Font.Underline = MsoTriState.msoFalse; // 取消下划线
                    }
                    else
                    {
                        // 使用“_”字符替换选中的文本
                        string underline = new string('_', (int)(selectedText.Length * 2.5)); // 动态生成下划线
                        textRange.Text = underline;
                    }

                    // 设置新文本框的位置与被选中的文本相同
                    newTextBox.Left = originalLeft;
                    newTextBox.Top = originalTop - (originalShape.Height - textRange.Font.Size) / 2; // 调整文本框位置
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
                PowerPoint.Selection sel = app.ActiveWindow.Selection;

                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText || sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    PowerPoint.TextRange textRange = null;

                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                    {
                        textRange = sel.TextRange;
                    }
                    else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
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
                                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                                PowerPoint.Shape svgShape = InsertSVGIntoSlide(svgContent, slide);
                                SelectShape(app, svgShape);
                                AddAndRunVBA(app);
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

        private PowerPoint.Shape InsertSVGIntoSlide(string svgContent, PowerPoint.Slide slide)
        {
            string tempSvgPath = Path.Combine(Path.GetTempPath(), "temp.svg");
            File.WriteAllText(tempSvgPath, svgContent);

            float left = 100;  // 可以根据需求调整
            float top = 100;   // 可以根据需求调整

            PowerPoint.Shape svgShape = slide.Shapes.AddPicture(tempSvgPath, MsoTriState.msoFalse, MsoTriState.msoCTrue, left, top);

            // 放大SVG
            svgShape.Width *= 2;
            svgShape.Height *= 2;

            File.Delete(tempSvgPath);

            return svgShape;
        }

        private void SelectShape(PowerPoint.Application app, PowerPoint.Shape shape)
        {
            shape.Select();
        }

        private void AddAndRunVBA(PowerPoint.Application app)
        {
            string vbaCode = @"
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
    
    ' Loop through shapes to find and delete the shape with name containing ""AutoShape""
    Dim shapeItem As Shape
    For Each shapeItem In slide.Shapes
        If InStr(shapeItem.Name, ""AutoShape"") > 0 Then
            shapeItem.Delete
        End If
    Next shapeItem
End Sub";

            VBIDE.VBProject vbProject = app.ActivePresentation.VBProject;
            VBIDE.VBComponent vbModule = vbProject.VBComponents.Add(VBIDE.vbext_ComponentType.vbext_ct_StdModule);
            vbModule.CodeModule.AddFromString(vbaCode);

            app.Run("ConvertSVGToShape");

            vbProject.VBComponents.Remove(vbModule);
        }

        public void PerformStrokeSplit()
        {
            笔画拆分_Click(this, EventArgs.Empty);
        }
    }
}
















