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

namespace 课件帮PPT助手
{
    public partial class DesignTools : UserControl
    {
        private SplitterTool splitterTool;

        public DesignTools()
        {
            InitializeComponent();
            splitterTool = new SplitterTool();
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


        private void 笔画拆分_Click(object sender, EventArgs e)
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
                        if (sel.ShapeRange.Count == 1 && sel.ShapeRange[1].HasTextFrame == Office.MsoTriState.msoTrue && sel.ShapeRange[1].TextFrame.HasText == Office.MsoTriState.msoTrue)
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
                                InsertSVGIntoSlide(svgContent, app.ActiveWindow.View.Slide);
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

        private void InsertSVGIntoSlide(string svgContent, PowerPoint.Slide slide)
        {
            string tempSvgPath = Path.Combine(Path.GetTempPath(), "temp.svg");
            File.WriteAllText(tempSvgPath, svgContent);

            float left = 100;  // 可以根据需求调整
            float top = 100;   // 可以根据需求调整

            PowerPoint.Shape svgShape = slide.Shapes.AddPicture(tempSvgPath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoCTrue, left, top);

            // 放大SVG
            svgShape.Width *= 2;
            svgShape.Height *= 2;

            File.Delete(tempSvgPath);
        }


        private void 书写动画_Click(object sender, EventArgs ev)
        {
            AnimationForm animationForm = new AnimationForm();
            animationForm.Show();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            splitterTool.StartSplitting();
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
    }
}


















