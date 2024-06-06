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
                        MessageBox.Show("未查询到对应SVG字源字形图或表头！");
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
            Form svgSelectionForm = new Form();
            svgSelectionForm.Text = "请选择字源字形SVG图形";
            svgSelectionForm.Size = new System.Drawing.Size(800, 600);

            ListBox svgListBox = new ListBox();
            svgListBox.Dock = DockStyle.Fill;

            foreach (var row in svgMatrix)
            {
                foreach (var svgItem in row)
                {
                    if (!string.IsNullOrEmpty(svgItem.Value))
                    {
                        svgListBox.Items.Add(svgItem);
                    }
                }
            }

            svgListBox.DisplayMember = "Key";
            svgListBox.ValueMember = "Value";
            svgListBox.SelectionMode = SelectionMode.MultiExtended;
            svgSelectionForm.Controls.Add(svgListBox);

            Button selectButton = new Button();
            selectButton.Text = "确认插入";
            selectButton.Size = new System.Drawing.Size(150, 50);
            selectButton.Dock = DockStyle.Bottom;
            selectButton.Click += (sender, e) =>
            {
                List<KeyValuePair<string, string>> selectedSVGs = new List<KeyValuePair<string, string>>();
                foreach (var selectedItem in svgListBox.SelectedItems)
                {
                    var kvp = (KeyValuePair<string, string>)selectedItem;
                    selectedSVGs.Add(kvp);
                }
                svgSelectionForm.Close();
                InsertSVGsIntoPresentation(svgMatrix, inputChar, headers);
            };
            svgSelectionForm.Controls.Add(selectButton);
            svgSelectionForm.ShowDialog();
        }

        private void InsertSVGsIntoPresentation(List<List<KeyValuePair<string, string>>> svgMatrix, string inputChar, List<string> headers)
        {
            try
            {
                PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
                PowerPoint.Slide slide = pptApp.ActiveWindow.View.Slide;

                int xOffset = 50;  // 初始 x 坐标
                int yOffset = 75;  // 初始 y 坐标
                int xSpacing = 130; // 每个 SVG 之间的水平间隔
                int ySpacing = 150; // 每行之间的垂直间隔

                // 插入顶部的分类描述
                int headerXOffset = xOffset;
                foreach (var header in headers)
                {
                    var headerTextBox = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, headerXOffset, yOffset - 50, 100, 20);
                    headerTextBox.TextFrame.TextRange.Text = header;
                    headerTextBox.TextFrame.TextRange.Font.Size = 12;
                    headerTextBox.TextFrame.TextRange.Font.Name = "Arial";
                    headerXOffset += xSpacing;
                }

                foreach (var row in svgMatrix)
                {
                    int columnCount = 0;
                    foreach (var svgItem in row)
                    {
                        if (string.IsNullOrEmpty(svgItem.Value))
                        {
                            columnCount++;
                            xOffset += xSpacing; // 更新 x 坐标
                            continue;
                        }

                        string svgUrl = svgItem.Value;
                        string description = svgItem.Key.Split('|')[1];

                        string tempSvgPath = Path.Combine(Path.GetTempPath(), $"{inputChar}-{columnCount}.svg");
                        using (var client = new System.Net.WebClient())
                        {
                            client.DownloadFile(svgUrl, tempSvgPath);
                        }

                        var picture = slide.Shapes.AddPicture(tempSvgPath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoCTrue, xOffset, yOffset);
                        picture.LockAspectRatio = Office.MsoTriState.msoTrue;
                        picture.Width *= 0.5f; // 缩放50%
                        picture.Height *= 0.5f;

                        var textBox = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, xOffset, yOffset + (int)picture.Height + 5, picture.Width, 50);
                        textBox.TextFrame.TextRange.Text = description;
                        textBox.TextFrame.TextRange.Font.Size = 12;
                        textBox.TextFrame.TextRange.Font.Name = "Arial";

                        File.Delete(tempSvgPath);

                        // 更新坐标位置
                        columnCount++;
                        xOffset += xSpacing;
                    }
                    xOffset = 50;
                    yOffset += ySpacing;
                }

                MessageBox.Show("成功插入字源字形SVG到PPT中！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("出现错误：" + ex.Message);
            }
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

        public class AnnotationToolForm : Form
        {
            public event Action<string, Color, bool, bool, Color, Color> AnnotationApplied;

            private ComboBox annotationTypeComboBox;
            private Button annotationColorButton;
            private Button confirmButton;
            private Button clearButton;
            private Button deleteCustomAnnotationButton;
            private CheckBox boldCheckBox;
            private CheckBox italicCheckBox;
            private CheckBox highlightCheckBox;
            private Button highlightColorButton;
            private Button textColorButton;
            private Label annotationTypeLabel;
            private Label annotationColorLabel;
            private Label textColorLabel;
            private Label highlightColorLabel;
            private Label textSettingsLabel;

            private ContextMenuStrip contextMenuStrip;
            private ToolStripMenuItem customizeAnnotationMenuItem;
            private ToolStripMenuItem deleteAnnotationMenuItem;

            private const string CustomAnnotationsFile = "custom_annotations.json";
            private const string CustomAnnotationPrefix = "[自定义] ";

            public string SelectedAnnotationType { get; private set; }
            public Color AnnotationColor { get; private set; }
            public bool IsBold { get; private set; }
            public bool IsItalic { get; private set; }
            public bool IsHighlight { get; private set; }
            public Color HighlightColor { get; private set; }
            public Color TextColor { get; private set; }

            public AnnotationToolForm()
            {
                InitializeComponent();
                SetDefaultValues();
                this.TopMost = true;
                this.FormBorderStyle = FormBorderStyle.FixedToolWindow;

                confirmButton.BackColor = Color.FromArgb(47, 85, 151);
                confirmButton.ForeColor = Color.White;
                clearButton.BackColor = Color.FromArgb(47, 85, 151);
                clearButton.ForeColor = Color.White;

                InitializeContextMenu();
                LoadCustomAnnotations();
            }

            private void SaveCustomAnnotations()
            {
                var customAnnotations = new List<string>();
                foreach (var item in annotationTypeComboBox.Items)
                {
                    if (item.ToString().StartsWith(CustomAnnotationPrefix))
                    {
                        customAnnotations.Add(item.ToString());
                    }
                }

                var json = JsonConvert.SerializeObject(customAnnotations);
                File.WriteAllText(CustomAnnotationsFile, json);
            }

            private void LoadCustomAnnotations()
            {
                if (File.Exists(CustomAnnotationsFile))
                {
                    var json = File.ReadAllText(CustomAnnotationsFile);
                    var customAnnotations = JsonConvert.DeserializeObject<List<string>>(json);
                    foreach (var annotation in customAnnotations)
                    {
                        annotationTypeComboBox.Items.Add(annotation);
                    }
                }
            }

            protected override void OnFormClosing(FormClosingEventArgs e)
            {
                base.OnFormClosing(e);
                SaveCustomAnnotations();
            }

            private void InitializeContextMenu()
            {
                contextMenuStrip = new ContextMenuStrip();
                customizeAnnotationMenuItem = new ToolStripMenuItem("自定义标注");
                customizeAnnotationMenuItem.Click += CustomizeAnnotationMenuItem_Click;
                deleteAnnotationMenuItem = new ToolStripMenuItem("删除标注");
                deleteAnnotationMenuItem.Click += DeleteAnnotationMenuItem_Click;
                contextMenuStrip.Items.Add(customizeAnnotationMenuItem);
                contextMenuStrip.Items.Add(deleteAnnotationMenuItem);

                annotationTypeComboBox.ContextMenuStrip = contextMenuStrip;
                annotationTypeComboBox.DrawMode = DrawMode.OwnerDrawFixed;
                annotationTypeComboBox.DrawItem += AnnotationTypeComboBox_DrawItem;
                annotationTypeComboBox.MouseDown += AnnotationTypeComboBox_MouseDown;
            }

            private void AnnotationTypeComboBox_DrawItem(object sender, DrawItemEventArgs e)
            {
                if (e.Index < 0)
                    return;

                e.DrawBackground();
                e.Graphics.DrawString(annotationTypeComboBox.Items[e.Index].ToString(), e.Font, Brushes.Black, e.Bounds);

                e.DrawFocusRectangle();
            }

            private void AnnotationTypeComboBox_MouseDown(object sender, MouseEventArgs e)
            {
                if (e.Button == MouseButtons.Right)
                {
                    int index = annotationTypeComboBox.SelectedIndex;
                    contextMenuStrip.Show(Cursor.Position);
                }
            }

            private void DeleteAnnotationMenuItem_Click(object sender, EventArgs e)
            {
                if (annotationTypeComboBox.SelectedIndex >= 0)
                {
                    string selectedItem = annotationTypeComboBox.SelectedItem.ToString();
                    if (selectedItem.StartsWith(CustomAnnotationPrefix))
                    {
                        annotationTypeComboBox.Items.RemoveAt(annotationTypeComboBox.SelectedIndex);
                        SaveCustomAnnotations();
                        deleteCustomAnnotationButton.Enabled = false;
                    }
                    else
                    {
                        MessageBox.Show("无法删除默认标注。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }

            private void CustomizeAnnotationMenuItem_Click(object sender, EventArgs e)
            {
                this.Hide();

                CustomizeAnnotationForm customizeForm = new CustomizeAnnotationForm();
                customizeForm.AnnotationSaved += (symbol, name, position) =>
                {
                    annotationTypeComboBox.Items.Add($"{CustomAnnotationPrefix}{name} ({symbol}) - {position}");
                };

                customizeForm.FormClosed += (s, args) =>
                {
                    this.Show();
                };

                customizeForm.ShowDialog();
            }

            public void SetDefaultValues()
            {
                annotationTypeComboBox.SelectedItem = "横线";
                annotationColorButton.BackColor = Color.Red;
                boldCheckBox.Checked = true;
                italicCheckBox.Checked = false;
                highlightCheckBox.Checked = false;
                highlightColorButton.BackColor = SystemColors.Control;
                textColorButton.BackColor = Color.Red;
            }

            private void InitializeComponent()
            {
                this.annotationTypeComboBox = new ComboBox();
                this.annotationColorButton = new Button();
                this.confirmButton = new Button();
                this.clearButton = new Button();
                this.deleteCustomAnnotationButton = new Button();
                this.boldCheckBox = new CheckBox();
                this.italicCheckBox = new CheckBox();
                this.highlightCheckBox = new CheckBox();
                this.highlightColorButton = new Button();
                this.textColorButton = new Button();
                this.annotationTypeLabel = new Label();
                this.annotationColorLabel = new Label();
                this.textColorLabel = new Label();
                this.highlightColorLabel = new Label();
                this.textSettingsLabel = new Label();
                this.SuspendLayout();

                this.annotationTypeLabel.AutoSize = true;
                this.annotationTypeLabel.Location = new Point(20, 20);
                this.annotationTypeLabel.Name = "annotationTypeLabel";
                this.annotationTypeLabel.Size = new Size(40, 30);
                this.annotationTypeLabel.Text = "标注：";

                this.annotationTypeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
                this.annotationTypeComboBox.FormattingEnabled = true;
                this.annotationTypeComboBox.Location = new Point(85, 20);
                this.annotationTypeComboBox.Name = "annotationTypeComboBox";
                this.annotationTypeComboBox.Size = new Size(210, 30);
                this.annotationTypeComboBox.Items.AddRange(new string[] { "横线", "双横线", "波浪线", "重读符号", "轻读符号", "着重符号", "大括号", "层级符", "段落符" });
                this.annotationTypeComboBox.SelectedIndexChanged += AnnotationTypeComboBox_SelectedIndexChanged;

                this.annotationColorLabel.AutoSize = true;
                this.annotationColorLabel.Location = new Point(310, 20);
                this.annotationColorLabel.Name = "annotationColorLabel";
                this.annotationColorLabel.Size = new Size(67, 30);
                this.annotationColorLabel.Text = "标注颜色：";

                this.annotationColorButton.Location = new Point(430, 20);
                this.annotationColorButton.Name = "annotationColorButton";
                this.annotationColorButton.Size = new Size(50, 30);
                this.annotationColorButton.Text = "";
                this.annotationColorButton.UseVisualStyleBackColor = true;
                this.annotationColorButton.Click += new EventHandler(this.AnnotationColorButton_Click);
                this.annotationColorButton.BackColor = Color.Red;

                this.textSettingsLabel.AutoSize = true;
                this.textSettingsLabel.Location = new Point(20, 80);
                this.textSettingsLabel.Name = "textSettingsLabel";
                this.textSettingsLabel.Size = new Size(43, 30);
                this.textSettingsLabel.Text = "文字：";

                this.boldCheckBox.AutoSize = true;
                this.boldCheckBox.Location = new Point(90, 80);
                this.boldCheckBox.Name = "boldCheckBox";
                this.boldCheckBox.Size = new Size(50, 30);
                this.boldCheckBox.Text = "加粗";
                this.boldCheckBox.UseVisualStyleBackColor = true;
                this.boldCheckBox.Checked = true;

                this.italicCheckBox.AutoSize = true;
                this.italicCheckBox.Location = new Point(185, 80);
                this.italicCheckBox.Name = "italicCheckBox";
                this.italicCheckBox.Size = new Size(50, 30);
                this.italicCheckBox.Text = "倾斜";
                this.italicCheckBox.UseVisualStyleBackColor = true;

                this.highlightCheckBox.AutoSize = true;
                this.highlightCheckBox.Location = new Point(280, 80);
                this.highlightCheckBox.Name = "highlightCheckBox";
                this.highlightCheckBox.Size = new Size(50, 30);
                this.highlightCheckBox.Text = "高亮";
                this.highlightCheckBox.UseVisualStyleBackColor = true;
                this.highlightCheckBox.CheckedChanged += new EventHandler(this.HighlightCheckBox_CheckedChanged);

                this.textColorLabel.AutoSize = true;
                this.textColorLabel.Location = new Point(20, 140);
                this.textColorLabel.Name = "textColorLabel";
                this.textColorLabel.Size = new Size(67, 30);
                this.textColorLabel.Text = "文字颜色：";

                this.textColorButton.Location = new Point(135, 140);
                this.textColorButton.Name = "textColorButton";
                this.textColorButton.Size = new Size(50, 30);
                this.textColorButton.Text = "";
                this.textColorButton.UseVisualStyleBackColor = true;
                this.textColorButton.Click += new EventHandler(this.TextColorButton_Click);
                this.textColorButton.BackColor = Color.Red;

                this.highlightColorLabel.AutoSize = true;
                this.highlightColorLabel.Location = new Point(230, 140);
                this.highlightColorLabel.Name = "highlightColorLabel";
                this.highlightColorLabel.Size = new Size(67, 30);
                this.highlightColorLabel.Text = "高亮颜色：";

                this.highlightColorButton.Location = new Point(340, 140);
                this.highlightColorButton.Name = "highlightColorButton";
                this.highlightColorButton.Size = new Size(50, 30);
                this.highlightColorButton.Text = "";
                this.highlightColorButton.UseVisualStyleBackColor = true;
                this.highlightColorButton.Click += new EventHandler(this.HighlightColorButton_Click);
                this.highlightColorButton.Enabled = false;

                this.confirmButton.Location = new Point(50, 210);
                this.confirmButton.Name = "confirmButton";
                this.confirmButton.Size = new Size(120, 60);
                this.confirmButton.Text = "标注所选";
                this.confirmButton.UseVisualStyleBackColor = true;
                this.confirmButton.Click += new EventHandler(this.ConfirmButton_Click);

                this.clearButton.Location = new Point(180, 210);
                this.clearButton.Name = "clearButton";
                this.clearButton.Size = new Size(140, 60);
                this.clearButton.Text = "清除标注";
                this.clearButton.UseVisualStyleBackColor = true;
                this.clearButton.Click += new EventHandler(this.ClearButton_Click);

                this.deleteCustomAnnotationButton.Location = new Point(330, 210);
                this.deleteCustomAnnotationButton.Name = "deleteCustomAnnotationButton";
                this.deleteCustomAnnotationButton.Size = new Size(120, 60);
                this.deleteCustomAnnotationButton.Text = "删除自定义标注";
                this.deleteCustomAnnotationButton.UseVisualStyleBackColor = true;
                this.deleteCustomAnnotationButton.Click += new EventHandler(this.DeleteCustomAnnotationButton_Click);
                this.deleteCustomAnnotationButton.Enabled = false;

                this.ClientSize = new Size(500, 340);
                this.Controls.Add(this.annotationTypeComboBox);
                this.Controls.Add(this.annotationColorButton);
                this.Controls.Add(this.boldCheckBox);
                this.Controls.Add(this.italicCheckBox);
                this.Controls.Add(this.highlightCheckBox);
                this.Controls.Add(this.highlightColorButton);
                this.Controls.Add(this.textColorButton);
                this.Controls.Add(this.confirmButton);
                this.Controls.Add(this.clearButton);
                this.Controls.Add(this.deleteCustomAnnotationButton);
                this.Controls.Add(this.annotationTypeLabel);
                this.Controls.Add(this.annotationColorLabel);
                this.Controls.Add(this.textColorLabel);
                this.Controls.Add(this.highlightColorLabel);
                this.Controls.Add(this.textSettingsLabel);
                this.Name = "AnnotationToolForm";
                this.Text = "文字标注工具";
                this.ResumeLayout(false);
                this.PerformLayout();
            }

            private void AnnotationTypeComboBox_SelectedIndexChanged(object sender, EventArgs e)
            {
                string selectedItem = annotationTypeComboBox.SelectedItem.ToString();
                deleteCustomAnnotationButton.Enabled = selectedItem.StartsWith(CustomAnnotationPrefix);
            }

            private void AnnotationColorButton_Click(object sender, EventArgs e)
            {
                using (ColorDialog colorDialog = new ColorDialog())
                {
                    if (colorDialog.ShowDialog() == DialogResult.OK)
                    {
                        AnnotationColor = colorDialog.Color;
                        annotationColorButton.BackColor = AnnotationColor;
                    }
                }
            }

            private void HighlightColorButton_Click(object sender, EventArgs e)
            {
                using (ColorDialog colorDialog = new ColorDialog())
                {
                    if (colorDialog.ShowDialog() == DialogResult.OK)
                    {
                        HighlightColor = colorDialog.Color;
                        highlightColorButton.BackColor = HighlightColor;
                    }
                }
            }

            private void TextColorButton_Click(object sender, EventArgs e)
            {
                using (ColorDialog colorDialog = new ColorDialog())
                {
                    if (colorDialog.ShowDialog() == DialogResult.OK)
                    {
                        TextColor = colorDialog.Color;
                        textColorButton.BackColor = TextColor;
                    }
                }
            }

            private void HighlightCheckBox_CheckedChanged(object sender, EventArgs e)
            {
                highlightColorButton.Enabled = highlightCheckBox.Checked;
                if (!highlightCheckBox.Checked)
                {
                    highlightColorButton.BackColor = SystemColors.Control;
                }
            }

            private void ConfirmButton_Click(object sender, EventArgs e)
            {
                SelectedAnnotationType = annotationTypeComboBox.SelectedItem.ToString();
                AnnotationColor = annotationColorButton.BackColor;
                IsBold = boldCheckBox.Checked;
                IsItalic = italicCheckBox.Checked;
                IsHighlight = highlightCheckBox.Checked;
                HighlightColor = highlightCheckBox.Checked ? highlightColorButton.BackColor : Color.Empty;
                TextColor = textColorButton.BackColor;
                AnnotationApplied?.Invoke(SelectedAnnotationType, AnnotationColor, IsBold, IsItalic, HighlightColor, TextColor);
            }

            private void ClearButton_Click(object sender, EventArgs e)
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Selection sel = app.ActiveWindow.Selection;

                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    PowerPoint.TextRange textRange = sel.TextRange;

                    // Clear text properties only for the selected text
                    textRange.Font.Bold = Office.MsoTriState.msoFalse;
                    textRange.Font.Italic = Office.MsoTriState.msoFalse;
                    textRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black);

                    // Remove annotations only for the selected text
                    string text = textRange.Text;
                    text = RemoveAnnotations(text);
                    textRange.Text = text;

                    // Remove shapes behind the selected text range if they overlap
                    PowerPoint.Slide slide = (PowerPoint.Slide)app.ActiveWindow.View.Slide;
                    List<PowerPoint.Shape> shapesToDelete = new List<PowerPoint.Shape>();

                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        if (shape.Name.StartsWith("Annotation_") && IsShapeOverlappingTextRange(shape, textRange))
                        {
                            shapesToDelete.Add(shape);
                        }
                    }

                    foreach (var shape in shapesToDelete)
                    {
                        shape.Delete();
                    }
                }
            }

            private bool IsShapeOverlappingTextRange(PowerPoint.Shape shape, PowerPoint.TextRange textRange)
            {
                // Check if the shape overlaps with the text range bounds
                float textLeft = textRange.BoundLeft;
                float textTop = textRange.BoundTop;
                float textWidth = textRange.BoundWidth;
                float textHeight = textRange.BoundHeight;

                float shapeLeft = shape.Left;
                float shapeTop = shape.Top;
                float shapeWidth = shape.Width;
                float shapeHeight = shape.Height;

                return !(shapeLeft + shapeWidth < textLeft ||
                         shapeLeft > textLeft + textWidth ||
                         shapeTop + shapeHeight < textTop ||
                         shapeTop > textTop + textHeight);
            }
            private string RemoveAnnotations(string text)
            {
                // Define both default and custom symbols to remove
                string[] defaultSymbols = new string[] { "{", "}", "※", "/", "//", "[", "]", "*", "(", ")", "▲", "○", "●" };

                // Load custom symbols from file
                string filePath = "custom_symbols.json";
                List<string> customSymbols = new List<string>();

                if (File.Exists(filePath))
                {
                    string json = File.ReadAllText(filePath);
                    customSymbols = JsonConvert.DeserializeObject<List<string>>(json);
                }

                // Combine default and custom symbols
                HashSet<string> symbolsToRemove = new HashSet<string>(defaultSymbols);
                foreach (string symbol in customSymbols)
                {
                    symbolsToRemove.Add(symbol);
                }

                // Remove all symbols from text
                foreach (string symbol in symbolsToRemove)
                {
                    text = text.Replace(symbol, "");
                }

                return text;
            }

            private void DeleteCustomAnnotationButton_Click(object sender, EventArgs e)
            {
                if (annotationTypeComboBox.SelectedIndex >= 0)
                {
                    string selectedItem = annotationTypeComboBox.SelectedItem.ToString();
                    if (selectedItem.StartsWith(CustomAnnotationPrefix))
                    {
                        annotationTypeComboBox.Items.RemoveAt(annotationTypeComboBox.SelectedIndex);
                        SaveCustomAnnotations();
                        deleteCustomAnnotationButton.Enabled = false;
                    }
                }
            }
        }

        public class CustomizeAnnotationForm : Form
        {
            public event Action<string, string, string> AnnotationSaved;

            private TextBox symbolTextBox;
            private TextBox nameTextBox;
            private RadioButton bottomRadioButton;
            private RadioButton startEndRadioButton;
            private RadioButton endRadioButton;
            private Button saveButton;
            private Button cancelButton;

            public CustomizeAnnotationForm()
            {
                InitializeComponent();
                this.Load += CustomizeAnnotationForm_Load;
            }

            private void InitializeComponent()
            {
                this.Text = "自定义标注";
                this.Size = new Size(525, 440);

                Label symbolLabel = new Label { Text = "标注符号：", Location = new Point(20, 20), Width = 140 };
                symbolTextBox = new TextBox { Location = new Point(160, 20), Width = 300 };

                Label nameLabel = new Label { Text = "符号名称：", Location = new Point(20, 70), Width = 140 };
                nameTextBox = new TextBox { Location = new Point(160, 60), Width = 300 };

                Label positionLabel = new Label { Text = "标注位置：", Location = new Point(20, 120), Width = 140 };

                bottomRadioButton = new RadioButton { Text = "所选文本的底部", Location = new Point(160, 120), Width = 220, Height = 40, Checked = true };
                startEndRadioButton = new RadioButton { Text = "所选文本的开头和末尾", Location = new Point(160, 170), Width = 300, Height = 40 };
                endRadioButton = new RadioButton { Text = "所选文本的末尾", Location = new Point(160, 220), Width = 220, Height = 40 };

                saveButton = new Button { Text = "保存", Location = new Point(140, 290), Width = 100, Height = 50 };
                cancelButton = new Button { Text = "退出", Location = new Point(270, 290), Width = 100, Height = 50 };

                saveButton.Click += SaveButton_Click;
                cancelButton.Click += (s, e) => this.Close();

                this.Controls.Add(symbolLabel);
                this.Controls.Add(symbolTextBox);
                this.Controls.Add(nameLabel);
                this.Controls.Add(nameTextBox);
                this.Controls.Add(positionLabel);
                this.Controls.Add(bottomRadioButton);
                this.Controls.Add(startEndRadioButton);
                this.Controls.Add(endRadioButton);
                this.Controls.Add(saveButton);
                this.Controls.Add(cancelButton);
            }

            private void CustomizeAnnotationForm_Load(object sender, EventArgs e)
            {
            }

            private void SaveButton_Click(object sender, EventArgs e)
            {
                string symbol = symbolTextBox.Text;
                string name = nameTextBox.Text;
                string position = bottomRadioButton.Checked ? "底部" :
                                  startEndRadioButton.Checked ? "开头和末尾" : "末尾";

                if (!string.IsNullOrEmpty(symbol) && !string.IsNullOrEmpty(name))
                {
                    SaveCustomSymbol(symbol); // Save the custom symbol
                    AnnotationSaved?.Invoke(symbol, name, position);
                    this.Close();
                }
                else
                {
                    MessageBox.Show("请填写符号和名称。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            private void SaveCustomSymbol(string symbol)
            {
                string filePath = "custom_symbols.json";
                List<string> symbols = new List<string>();

                if (File.Exists(filePath))
                {
                    string json = File.ReadAllText(filePath);
                    symbols = JsonConvert.DeserializeObject<List<string>>(json);
                }

                if (!symbols.Contains(symbol))
                {
                    symbols.Add(symbol);
                }

                string outputJson = JsonConvert.SerializeObject(symbols);
                File.WriteAllText(filePath, outputJson);
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


        private Form animationForm;
        private List<PowerPoint.Shape> selectedShapes;
        private Panel adjustPanel;
        private ListBox listBox;
        private Dictionary<string, NumericUpDown> durationControls;
        private Label durationLabel;
        private NumericUpDown durationControl;


        private void 书写动画_Click(object sender, EventArgs ev)
        {
            if (animationForm != null)
            {
                animationForm.Dispose();
            }

            animationForm = new Form();
            animationForm.Text = "书写动画辅助生成";
            animationForm.Size = new System.Drawing.Size(520, 555);
            animationForm.TopMost = true;

            TabControl tabControl = new TabControl();
            tabControl.Dock = DockStyle.Fill;

            TabPage tabPage1 = new TabPage("第一步");
            TabPage tabPage2 = new TabPage("第二步");

            // 第一页内容
            Label label1 = new Label();
            label1.Text = "提示：请按照笔画顺序依次选中所有笔画。";
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point(10, 40);

            Label inputLabel = new Label();
            inputLabel.Text = "①请输入对应汉字：";
            inputLabel.AutoSize = true;
            inputLabel.Location = new System.Drawing.Point(10, 90);

            TextBox textBox = new TextBox();
            textBox.Name = "prefixTextBox";
            textBox.Location = new System.Drawing.Point(255, 90);
            textBox.Width = 190;
            textBox.KeyDown += TextBox_KeyDown;

            tabPage1.Controls.Add(label1);
            tabPage1.Controls.Add(inputLabel);
            tabPage1.Controls.Add(textBox);

            // 第二页内容
            Label label2 = new Label();
            label2.Text = "提示：“智能全选”→“智能动画”。";
            label2.AutoSize = true;
            label2.Location = new System.Drawing.Point(10, 20);

            Button selectAllButton = new Button();
            selectAllButton.Text = "②智能全选";
            selectAllButton.BackColor = System.Drawing.Color.FromArgb(47, 85, 151);
            selectAllButton.ForeColor = System.Drawing.Color.White;
            selectAllButton.Width = 220;
            selectAllButton.Height = 40;
            selectAllButton.Location = new System.Drawing.Point(10, 70);
            selectAllButton.Click += SelectAllButton_Click;

            Button animateButton = new Button();
            animateButton.Text = "③智能动画";
            animateButton.BackColor = System.Drawing.Color.FromArgb(47, 85, 151);
            animateButton.ForeColor = System.Drawing.Color.White;
            animateButton.Width = 220;
            animateButton.Height = 40;
            animateButton.Location = new System.Drawing.Point(250, 70);
            animateButton.Click += AnimateButton_Click;

            Button adjustAnimationButton = new Button();
            adjustAnimationButton.Text = "动画调整";
            adjustAnimationButton.BackColor = System.Drawing.Color.FromArgb(47, 85, 151);
            adjustAnimationButton.ForeColor = System.Drawing.Color.White;
            adjustAnimationButton.Width = 458;
            adjustAnimationButton.Height = 40;
            adjustAnimationButton.Location = new System.Drawing.Point(10, 120);
            adjustAnimationButton.Click += AdjustAnimationButton_Click;

            adjustPanel = new Panel();
            adjustPanel.Size = new System.Drawing.Size(460, 510);
            adjustPanel.Location = new System.Drawing.Point(10, 180);
            adjustPanel.Visible = false;

            listBox = new ListBox();
            listBox.SelectionMode = SelectionMode.MultiExtended;
            listBox.Location = new System.Drawing.Point(10, 10);
            listBox.Size = new System.Drawing.Size(200, 200);

            durationControls = new Dictionary<string, NumericUpDown>();

            listBox.SelectedIndexChanged += (s, ev1) =>
            {
                foreach (Control control in adjustPanel.Controls.OfType<NumericUpDown>())
                {
                    adjustPanel.Controls.Remove(control);
                }

                PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
                pptApp.ActiveWindow.Selection.Unselect();

                if (listBox.SelectedItems.Count > 1)
                {
                    foreach (var selectedItem in listBox.SelectedItems)
                    {
                        string shapeName = selectedItem.ToString();
                        var shape = pptApp.ActiveWindow.View.Slide.Shapes[shapeName];
                        shape.Select(Office.MsoTriState.msoFalse);
                    }

                    NumericUpDown multiDurationControl = new NumericUpDown();
                    multiDurationControl.Minimum = 0.1m;
                    multiDurationControl.Maximum = 10m;
                    multiDurationControl.DecimalPlaces = 2;
                    multiDurationControl.Increment = 0.1m;
                    multiDurationControl.Value = 0.50m;
                    multiDurationControl.ValueChanged += MultiDurationControl_ValueChanged;

                    multiDurationControl.Location = new System.Drawing.Point(270, 170);
                    adjustPanel.Controls.Add(multiDurationControl);
                }
                else
                {
                    foreach (var selectedItem in listBox.SelectedItems)
                    {
                        string shapeName = selectedItem.ToString();
                        var shape = pptApp.ActiveWindow.View.Slide.Shapes[shapeName];
                        shape.Select(Office.MsoTriState.msoFalse);

                        NumericUpDown durationControl;
                        if (!durationControls.TryGetValue(shapeName, out durationControl))
                        {
                            durationControl = new NumericUpDown();
                            durationControl.Minimum = 0.1m;
                            durationControl.Maximum = 10m;
                            durationControl.DecimalPlaces = 2;
                            durationControl.Increment = 0.1m;
                            durationControl.Value = 0.50m;
                            durationControl.Tag = shapeName;
                            durationControl.ValueChanged += DurationControl_ValueChanged;
                            durationControls[shapeName] = durationControl;
                        }

                        durationControl.Location = new System.Drawing.Point(270, 170);
                        adjustPanel.Controls.Add(durationControl);
                    }
                }
            };

            Button upButton = new Button();
            upButton.Text = "↑";
            upButton.Size = new System.Drawing.Size(50, 50);
            upButton.Location = new System.Drawing.Point(320, 10);
            upButton.Click += (s, ev2) => AdjustAnimationDirection(listBox, PowerPoint.MsoAnimDirection.msoAnimDirectionBottom);

            Button downButton = new Button();
            downButton.Text = "↓";
            downButton.Size = new System.Drawing.Size(50, 50);
            downButton.Location = new System.Drawing.Point(320, 70);
            downButton.Click += (s, ev3) => AdjustAnimationDirection(listBox, PowerPoint.MsoAnimDirection.msoAnimDirectionTop);

            Button leftButton = new Button();
            leftButton.Text = "←";
            leftButton.Size = new System.Drawing.Size(50, 50);
            leftButton.Location = new System.Drawing.Point(270, 40);
            leftButton.Click += (s, ev4) => AdjustAnimationDirection(listBox, PowerPoint.MsoAnimDirection.msoAnimDirectionRight);

            Button rightButton = new Button();
            rightButton.Text = "→";
            rightButton.Size = new System.Drawing.Size(50, 50);
            rightButton.Location = new System.Drawing.Point(370, 40);
            rightButton.Click += (s, ev5) => AdjustAnimationDirection(listBox, PowerPoint.MsoAnimDirection.msoAnimDirectionLeft);

            durationLabel = new Label();
            durationLabel.Text = "动画持续时间：";
            durationLabel.AutoSize = true;
            durationLabel.Location = new System.Drawing.Point(270, 135);
            adjustPanel.Controls.Add(durationLabel);

            adjustPanel.Controls.Add(listBox);
            adjustPanel.Controls.Add(upButton);
            adjustPanel.Controls.Add(downButton);
            adjustPanel.Controls.Add(leftButton);
            adjustPanel.Controls.Add(rightButton);

            tabPage2.Controls.Add(label2);
            tabPage2.Controls.Add(selectAllButton);
            tabPage2.Controls.Add(animateButton);
            tabPage2.Controls.Add(adjustAnimationButton);
            tabPage2.Controls.Add(adjustPanel);

            tabControl.TabPages.Add(tabPage1);
            tabControl.TabPages.Add(tabPage2);

            animationForm.Controls.Add(tabControl);
            animationForm.Show();
        }

        private void TextBox_KeyDown(object sender, KeyEventArgs ev)
        {
            if (ev.KeyCode == Keys.Enter)
            {
                PowerPoint.Application pptApplication = Globals.ThisAddIn.Application;
                PowerPoint.DocumentWindow activeWindow = pptApplication.ActiveWindow;
                PowerPoint.Selection selection = activeWindow.Selection;

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    TextBox textBox = sender as TextBox;
                    string prefix = textBox.Text;

                    if (!string.IsNullOrEmpty(prefix))
                    {
                        int counter = 1;
                        selectedShapes = new List<PowerPoint.Shape>();
                        foreach (PowerPoint.Shape shape in selection.ShapeRange)
                        {
                            shape.Name = $"{prefix}-{counter}";
                            selectedShapes.Add(shape);
                            counter++;
                        }

                        activeWindow.View.GotoSlide(activeWindow.View.Slide.SlideIndex);
                    }
                    else
                    {
                        MessageBox.Show("命名前缀不能为空。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("请选择一个或多个对象。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void SelectAllButton_Click(object sender, EventArgs ev)
        {
            if (selectedShapes == null || !selectedShapes.Any())
            {
                MessageBox.Show("请先完成第一步的批量命名。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Slide slide = pptApp.ActiveWindow.View.Slide;

            pptApp.ActiveWindow.Selection.Unselect();
            foreach (var shape in selectedShapes)
            {
                shape.Select(Office.MsoTriState.msoFalse);
            }
        }

        private void AnimateButton_Click(object sender, EventArgs ev)
        {
            if (selectedShapes == null || !selectedShapes.Any())
            {
                MessageBox.Show("请先完成第一步的批量命名。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Slide slide = pptApp.ActiveWindow.View.Slide;

            PowerPoint.TimeLine timeLine = slide.TimeLine;
            bool isFirstEffect = true;
            foreach (PowerPoint.Shape shape in selectedShapes)
            {
                PowerPoint.Effect effect = timeLine.MainSequence.AddEffect(
                    shape,
                    PowerPoint.MsoAnimEffect.msoAnimEffectWipe,
                    PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                    isFirstEffect ? PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick : PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious
                );

                if (shape.Width > shape.Height)
                {
                    effect.EffectParameters.Direction = PowerPoint.MsoAnimDirection.msoAnimDirectionLeft;
                }
                else
                {
                    effect.EffectParameters.Direction = PowerPoint.MsoAnimDirection.msoAnimDirectionUp;
                }

                NumericUpDown durationControl;
                if (durationControls.TryGetValue(shape.Name, out durationControl))
                {
                    effect.Timing.Duration = (float)durationControl.Value;
                }

                isFirstEffect = false;
            }
        }

        private void AdjustAnimationButton_Click(object sender, EventArgs ev)
        {
            if (selectedShapes != null)
            {
                listBox.Items.Clear();
                string currentPrefix = selectedShapes[0].Name.Split('-')[0]; // 获取前缀

                foreach (var shape in selectedShapes)
                {
                    if (shape.Name.StartsWith(currentPrefix))
                    {
                        listBox.Items.Add(shape.Name);
                    }
                }
            }
            adjustPanel.Visible = !adjustPanel.Visible;
        }

        private void AdjustAnimationDirection(ListBox listBox, PowerPoint.MsoAnimDirection direction)
        {
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Slide slide = pptApp.ActiveWindow.View.Slide;

            foreach (string shapeName in listBox.SelectedItems)
            {
                var shape = slide.Shapes[shapeName];
                var effect = slide.TimeLine.MainSequence.Cast<PowerPoint.Effect>().FirstOrDefault(e => e.Shape.Name == shapeName);
                if (effect != null)
                {
                    effect.EffectParameters.Direction = direction;
                }
            }
        }

        private void DurationControl_ValueChanged(object sender, EventArgs ev)
        {
            NumericUpDown durationControl = sender as NumericUpDown;
            string shapeName = durationControl.Tag as string;

            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Slide slide = pptApp.ActiveWindow.View.Slide;

            var effect = slide.TimeLine.MainSequence.Cast<PowerPoint.Effect>().FirstOrDefault(e => e.Shape.Name == shapeName);
            if (effect != null)
            {
                effect.Timing.Duration = (float)durationControl.Value;
            }
        }

        private void MultiDurationControl_ValueChanged(object sender, EventArgs ev)
        {
            NumericUpDown multiDurationControl = sender as NumericUpDown;
            float newDuration = (float)multiDurationControl.Value;

            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Slide slide = pptApp.ActiveWindow.View.Slide;

            foreach (string shapeName in listBox.SelectedItems)
            {
                var effect = slide.TimeLine.MainSequence.Cast<PowerPoint.Effect>().FirstOrDefault(e => e.Shape.Name == shapeName);
                if (effect != null)
                {
                    effect.Timing.Duration = newDuration;
                }

                NumericUpDown durationControl;
                if (durationControls.TryGetValue(shapeName, out durationControl))
                {
                    durationControl.Value = (decimal)newDuration;
                }
            }
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
    }
}



















