using System;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlAgilityPack;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using OfficeOpenXml;
using System.IO;
using System.Reflection;

namespace 课件帮PPT助手
{
    public partial class 快捷盒子 : Form
    {
        public 快捷盒子()
        {
            InitializeComponent();
            comboBox1.SelectedIndex = 0; // 默认选择第一个选项
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedType = comboBox1.SelectedItem.ToString();
            if (selectedType == "同义查询" || selectedType == "反义查询")
            {
                this.Height = 460;
                synonymsPanel.Visible = true;
            }
            else
            {
                this.Height = 160;
                synonymsPanel.Visible = false;
            }
        }

        private async void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string selectedType = comboBox1.SelectedItem.ToString();
                string input = textBox1.Text;

                PowerPoint.Application pptApp = Globals.ThisAddIn.Application;

                if (selectedType == "同义查询")
                {
                    await QueryWords(input, "同义词", pptApp.ActiveWindow.View.Slide);
                    return;
                }
                else if (selectedType == "反义查询")
                {
                    await QueryWords(input, "反义词", pptApp.ActiveWindow.View.Slide);
                    return;
                }
                if (selectedType == "生字模板")
                {
                    ApplyCharacterTemplate(input);
                    return;
                }
                if (selectedType == "生字方格")
                {
                    ApplyCharacterGrid(input);
                    return;
                }

                PowerPoint.Selection selection = pptApp.ActiveWindow.Selection;

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    switch (selectedType)
                    {
                        case "批量命名":
                            BatchRename(selection, input);
                            break;
                        case "原位复制":
                            BatchDuplicate(selection, input);
                            break;
                        case "尺寸缩放":
                            BatchScale(selection, input);
                            break;
                        case "分割词语":
                            if (int.TryParse(input, out int numRows))
                            {
                                SplitWords(selection, numRows);
                            }
                            else
                            {
                                MessageBox.Show("请输入一个有效的行数。");
                            }
                            break;
                        case "水平间距":
                            if (float.TryParse(input, out float spacing))
                            {
                                AdjustHorizontalSpacing(selection, spacing);
                            }
                            else
                            {
                                MessageBox.Show("请输入一个有效的间距数值。");
                            }
                            break;
                        case "垂直间距":
                            if (float.TryParse(input, out float verticalSpacing))
                            {
                                AdjustVerticalSpacing(selection, verticalSpacing);
                            }
                            else
                            {
                                MessageBox.Show("请输入一个有效的间距数值。");
                            }
                            break;
                        default:
                            MessageBox.Show("未知操作类型。");
                            break;
                    }
                }
                else
                {
                    MessageBox.Show("请选择一个或多个对象。");
                }
            }
        }
        private void ApplyCharacterGrid(string input)
        {
            var inputs = input.Split(',');
            if (inputs.Length == 2 && int.TryParse(inputs[0], out int rows) && int.TryParse(inputs[1], out int columns))
            {
                if (rows < 1 || columns < 1)
                {
                    MessageBox.Show("行数和列数必须大于0。");
                    return;
                }

                string pptPath = ExtractCharacterGridResource("课件帮PPT助手.Resources.生字格.pptx");
                CopyCharacterGrid(pptPath, rows, columns);
                DeleteTemporaryFile(pptPath);  // 删除临时文件
            }
            else
            {
                MessageBox.Show("请输入有效的行数和列数，用逗号分隔。");
            }
        }

        private string ExtractCharacterGridResource(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceStream = assembly.GetManifestResourceStream(resourceName);
            if (resourceStream == null)
            {
                throw new Exception("无法找到资源文件。");
            }

            var tempFilePath = Path.Combine(Path.GetTempPath(), Path.GetFileName(resourceName));
            using (var fileStream = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write))
            {
                resourceStream.CopyTo(fileStream);
            }

            return tempFilePath;
        }

        private void CopyCharacterGrid(string pptPath, int rows, int columns)
        {
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Presentation presentation = pptApp.Presentations.Open(pptPath, WithWindow: Office.MsoTriState.msoFalse);
            PowerPoint.Slide sourceSlide = presentation.Slides[1];
            PowerPoint.Shape shapeToCopy = sourceSlide.Shapes["生字格"];

            PowerPoint.Slide currentSlide = pptApp.ActiveWindow.View.Slide;
            float startX = 100;
            float startY = 100;
            float spacingX = shapeToCopy.Width + 10;
            float spacingY = shapeToCopy.Height + 10;

            // 首先复制一次形状到当前幻灯片，以便后续的重复操作基于此新形状
            shapeToCopy.Copy();
            PowerPoint.Shape newShape = currentSlide.Shapes.Paste()[1];

            // 调整原始复制形状的位置
            newShape.Left = startX;
            newShape.Top = startY;

            // 从该新形状进行复制并调整位置
            for (int row = 0; row < rows; row++)
            {
                for (int col = 0; col < columns; col++)
                {
                    if (row == 0 && col == 0) continue; // 跳过已经复制的第一个形状

                    PowerPoint.Shape duplicatedShape = newShape.Duplicate()[1];
                    duplicatedShape.Left = startX + (col * spacingX);
                    duplicatedShape.Top = startY + (row * spacingY);
                }
            }

            presentation.Close();
        }

        private void DeleteTemporaryFile(string filePath)
        {
            if (File.Exists(filePath))
            {
                try
                {
                    File.Delete(filePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("删除临时文件时发生错误：" + ex.Message);
                }
            }
        }


        private void AdjustVerticalSpacing(PowerPoint.Selection selection, float spacing)
        {
            if (selection.ShapeRange.Count < 2)
            {
                MessageBox.Show("请选择两个或多个对象。");
                return;
            }

            var shapesByX = selection.ShapeRange.Cast<PowerPoint.Shape>()
                .GroupBy(s => s.Left)
                .Where(g => g.Count() > 1)
                .ToArray();

            foreach (var group in shapesByX)
            {
                var shapes = group.OrderBy(s => s.Top).ToArray();
                for (int i = 1; i < shapes.Length; i++)
                {
                    shapes[i].Top = shapes[i - 1].Top + shapes[i - 1].Height + spacing;
                }
            }
        }


        private void AdjustHorizontalSpacing(PowerPoint.Selection selection, float spacing)
        {
            if (selection.ShapeRange.Count < 2)
            {
                MessageBox.Show("请选择两个或多个对象。");
                return;
            }

            var shapesByY = selection.ShapeRange.Cast<PowerPoint.Shape>()
                .GroupBy(s => s.Top)
                .Where(g => g.Count() > 1)
                .ToArray();

            foreach (var group in shapesByY)
            {
                var shapes = group.OrderBy(s => s.Left).ToArray();
                for (int i = 1; i < shapes.Length; i++)
                {
                    shapes[i].Left = shapes[i - 1].Left + shapes[i - 1].Width + spacing;
                }
            }
        }

        private void ApplyCharacterTemplate(string character)
        {
            string pptPath = null;
            PowerPoint.Presentation sourcePresentation = null;
            try
            {
                pptPath = ExtractResourceFile("课件帮PPT助手.Resources.生字教学.pptx");
                if (string.IsNullOrEmpty(pptPath))
                {
                    MessageBox.Show("无法提取PPT资源。");
                    return;
                }

                PowerPoint.Application app = Globals.ThisAddIn.Application;
                sourcePresentation = app.Presentations.Open(pptPath, WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);

                PowerPoint.Slide sourceSlide = sourcePresentation.Slides[1];

                string filePath = ExtractResourceFile("课件帮PPT助手.汉字字典.汉字字典.xlsx");
                if (string.IsNullOrEmpty(filePath))
                {
                    MessageBox.Show("无法提取汉字字典资源。");
                    return;
                }

                string pinyin = GetPinyinResult(filePath, character);
                if (string.IsNullOrEmpty(pinyin))
                {
                    MessageBox.Show("未找到拼音信息。");
                    return;
                }

                PowerPoint.Shape pinyinTextBox = FindShapeByName(sourceSlide, "拼音返回文本框");
                if (pinyinTextBox != null)
                {
                    pinyinTextBox.TextFrame.TextRange.Text = pinyin;
                }
                else
                {
                    MessageBox.Show("未找到名为‘拼音返回文本框’的形状。");
                    return;
                }

                PowerPoint.Shape strokeReplaceTextBox = FindShapeByName(sourceSlide, "[笔画拆分]替换");
                if (strokeReplaceTextBox != null)
                {
                    strokeReplaceTextBox.TextFrame.TextRange.Text = character;
                }
                else
                {
                    MessageBox.Show("未找到名为‘[笔画拆分]替换’的形状。");
                    return;
                }

                var characterInfo = GetCharacterInfo(filePath, character);
                if (characterInfo == null)
                {
                    MessageBox.Show("未找到相关汉字信息。");
                    return;
                }

                PowerPoint.Shape tableShape = FindShapeByName(sourceSlide, "表格（部首-结构-笔画）");
                if (tableShape != null)
                {
                    var table = tableShape.Table;
                    if (table.Columns.Count >= 2 && table.Rows.Count >= 3)
                    {
                        table.Cell(1, 2).Shape.TextFrame.TextRange.Text = characterInfo.Radical;
                        table.Cell(2, 2).Shape.TextFrame.TextRange.Text = characterInfo.Structure;
                        table.Cell(3, 2).Shape.TextFrame.TextRange.Text = characterInfo.Strokes.ToString();
                    }
                    else
                    {
                        MessageBox.Show("表格的行数或列数不足。");
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("未找到名为‘表格（部首-结构-笔画）’的形状。");
                    return;
                }

                for (int i = 0; i < 8; i++)
                {
                    PowerPoint.Shape groupWordShape = FindShapeByName(sourceSlide, $"相关组词-{i + 1}");
                    if (groupWordShape != null)
                    {
                        groupWordShape.TextFrame.TextRange.Text = characterInfo.RelatedWords.ElementAtOrDefault(i) ?? "";
                    }
                    else
                    {
                        MessageBox.Show($"未找到名为‘相关组词-{i + 1}’的形状。");
                    }
                }

                PowerPoint.Presentation currentPresentation = app.ActivePresentation;
                PowerPoint.Slide currentSlide = app.ActiveWindow.View.Slide;

                sourceSlide.Copy();
                currentPresentation.Slides.Paste(currentSlide.SlideIndex + 1);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"发生错误：{ex.Message}");
            }
            finally
            {
                if (sourcePresentation != null)
                {
                    sourcePresentation.Close();
                }
                if (pptPath != null && File.Exists(pptPath))
                {
                    try
                    {
                        File.Delete(pptPath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"无法删除临时文件：{ex.Message}");
                    }
                }
            }
        }

        private string GetPinyinResult(string filePath, string character)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                if (package.Workbook.Worksheets.Count == 0)
                {
                    throw new Exception("没有查到相关信息。");
                }

                var worksheet = package.Workbook.Worksheets[0];

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    if (worksheet.Cells[row, 1].Text == character)
                    {
                        return worksheet.Cells[row, 2].Text;
                    }
                }
            }

            return null;
        }

        private PowerPoint.Shape FindShapeByName(PowerPoint.Slide slide, string shapeName)
        {
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                if (shape.Name == shapeName)
                {
                    return shape;
                }
            }
            return null;
        }

        private HanziInfo GetCharacterInfo(string filePath, string character)
        {
            var info = new HanziInfo();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                if (package.Workbook.Worksheets.Count == 0)
                {
                    throw new Exception("没有查到相关信息。");
                }

                var worksheet = package.Workbook.Worksheets[0];

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    if (worksheet.Cells[row, 1].Text == character)
                    {
                        info.Radical = worksheet.Cells[row, 3].Text;
                        info.Structure = worksheet.Cells[row, 4].Text;
                        info.Strokes = int.Parse(worksheet.Cells[row, 5].Text);
                        info.RelatedWords = worksheet.Cells[row, 6].Text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                    .Select(word => word.Trim())
                                                    .Where(word => !string.IsNullOrEmpty(word))
                                                    .ToArray();
                        break;
                    }
                }
            }

            return info;
        }

        private string ExtractResourceFile(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (Stream resourceStream = assembly.GetManifestResourceStream(resourceName))
            {
                if (resourceStream == null)
                    throw new Exception($"未找到嵌入资源 {resourceName}.");

                string tempFilePath = Path.Combine(Path.GetTempPath(), resourceName);

                using (var fileStream = new FileStream(tempFilePath, FileMode.Create))
                {
                    resourceStream.CopyTo(fileStream);
                }

                return tempFilePath;
            }
        }

        public class HanziInfo
        {
            public string Radical { get; set; }
            public string Structure { get; set; }
            public int Strokes { get; set; }
            public string[] RelatedWords { get; set; }
        }

        private void BatchRename(PowerPoint.Selection selection, string prefix)
        {
            if (string.IsNullOrEmpty(prefix))
            {
                MessageBox.Show("请输入一个前缀。");
                return;
            }

            int counter = 1;
            foreach (PowerPoint.Shape shape in selection.ShapeRange)
            {
                RenameShape(shape, prefix, ref counter);
            }

            // 刷新视图
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex);
        }

        private void RenameShape(PowerPoint.Shape shape, string prefix, ref int counter)
        {
            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                foreach (PowerPoint.Shape childShape in shape.GroupItems)
                {
                    RenameShape(childShape, prefix, ref counter);
                }
            }
            else
            {
                shape.Name = $"{prefix}-{counter}";
                counter++;
            }
        }

        private void BatchDuplicate(PowerPoint.Selection selection, string input)
        {
            if (!int.TryParse(input, out int copyCount) || copyCount < 1)
            {
                MessageBox.Show("请输入一个大于0的整数。");
                return;
            }

            for (int i = 0; i < copyCount; i++)
            {
                DuplicateSelectedShapes(selection);
            }
        }

        private void DuplicateSelectedShapes(PowerPoint.Selection selection)
        {
            foreach (PowerPoint.Shape shape in selection.ShapeRange)
            {
                PowerPoint.Shape copiedShape = shape.Duplicate()[1];
                copiedShape.Left = shape.Left;
                copiedShape.Top = shape.Top;
                copiedShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward); // 确保复制出来的对象是所选对象的上一层
            }
        }

        private void BatchScale(PowerPoint.Selection selection, string input)
        {
            string[] scaleValues = input.Split(',');

            bool isArithmetic = scaleValues.Length == 2;

            float commonDifference = 0;
            if (isArithmetic)
            {
                if (!float.TryParse(scaleValues[0], out float startScale) || !float.TryParse(scaleValues[1], out float endScale))
                {
                    MessageBox.Show("请输入有效的缩放比例。");
                    return;
                }

                commonDifference = (endScale - startScale) / (selection.ShapeRange.Count - 1);
            }

            if (!float.TryParse(scaleValues[0], out float currentScale))
            {
                MessageBox.Show("请输入有效的缩放比例。");
                return;
            }

            foreach (PowerPoint.Shape shape in selection.ShapeRange)
            {
                ScaleShape(shape, currentScale);

                if (isArithmetic)
                {
                    currentScale += commonDifference;
                }
            }
        }

        private void ScaleShape(PowerPoint.Shape shape, float scale)
        {
            float newWidth = shape.Width * scale / 100;
            float newHeight = shape.Height * scale / 100;

            float newX = shape.Left + (shape.Width - newWidth) / 2;
            float newY = shape.Top + (shape.Height - newHeight) / 2;

            shape.LockAspectRatio = Office.MsoTriState.msoTrue;
            shape.Width = newWidth;
            shape.Height = newHeight;
            shape.Left = newX;
            shape.Top = newY;
        }

        private void SplitWords(PowerPoint.Selection selection, int numRows)
        {
            if (selection.ShapeRange.Count == 1 && selection.ShapeRange[1].Type == Office.MsoShapeType.msoTextBox)
            {
                PowerPoint.Shape shape = selection.ShapeRange[1];
                string text = shape.TextFrame.TextRange.Text;
                float originalFontSize = shape.TextFrame.TextRange.Font.Size;

                char[] punctuation = { '.', ',', '!', '?', ';', ':', '，', '。', '！', '？', '；', '：' };
                var segments = text.Split(punctuation, StringSplitOptions.RemoveEmptyEntries)
                                   .Select(s => s.Trim())
                                   .Where(s => !string.IsNullOrEmpty(s))
                                   .ToArray();

                PowerPoint.Slide slide = shape.Parent;

                float startX = shape.Left;
                float startY = shape.Top;
                float offsetX = 0;
                float offsetY = 0;
                float boxHeight = shape.Height;
                float spacing = 10;

                shape.Delete();

                int numCols = (int)Math.Ceiling((double)segments.Length / numRows);

                for (int i = 0; i < segments.Length; i++)
                {
                    var segment = segments[i];

                    PowerPoint.Shape newTextBox = slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        startX + offsetX,
                        startY + offsetY,
                        0,
                        boxHeight);

                    newTextBox.TextFrame.TextRange.Text = segment;
                    newTextBox.TextFrame.TextRange.Font.Size = originalFontSize;
                    newTextBox.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                    newTextBox.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                    newTextBox.TextFrame.WordWrap = Office.MsoTriState.msoFalse;

                    float segmentWidth = newTextBox.TextFrame.TextRange.BoundWidth;

                    newTextBox.Width = segmentWidth;

                    offsetY += boxHeight + spacing;

                    if ((i + 1) % numRows == 0)
                    {
                        offsetY = 0;
                        offsetX += segmentWidth + spacing;
                    }
                }
            }
        }

        private async Task QueryWords(string word, string queryType, PowerPoint.Slide slide)
        {
            string url = $"https://hanyu.baidu.com/s?wd={word}&ptype=zici&tn";
            HttpClient client = new HttpClient();
            string html = await client.GetStringAsync(url);

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(html);

            string xpath = queryType == "同义词" ?
                "//div[@id='synonym']//div[@class='block']/a" :
                "//div[@id='antonym']//div[@class='block']/a";

            var wordNodes = doc.DocumentNode.SelectNodes(xpath);

            if (wordNodes != null && wordNodes.Count > 0)
            {
                string[] wordsArray = wordNodes.Select(node => node.InnerText).ToArray();

                Invoke(new Action(() =>
                {
                    synonymsPanel.Controls.Clear();

                    int x = 10;
                    int y = 10;
                    int buttonWidth = 80;
                    int buttonHeight = 50;
                    int spacing = 10;
                    int maxColumns = 4;
                    int currentColumn = 0;

                    foreach (string synonym in wordsArray)
                    {
                        Button wordButton = new Button();
                        wordButton.Text = synonym;
                        wordButton.Width = buttonWidth;
                        wordButton.Height = buttonHeight;
                        wordButton.Left = x;
                        wordButton.Top = y;
                        wordButton.Click += wordButton_Click;

                        synonymsPanel.Controls.Add(wordButton);

                        x += buttonWidth + spacing;
                        currentColumn++;

                        if (currentColumn >= maxColumns)
                        {
                            currentColumn = 0;
                            x = 10;
                            y += buttonHeight + spacing;
                        }
                    }
                }));
            }
            else
            {
                MessageBox.Show($"未找到{queryType}。");
            }
        }

        private void wordButton_Click(object sender, EventArgs e)
        {
            var button = sender as Button;
            string word = button.Text;

            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = pptApp.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count == 1)
            {
                PowerPoint.Shape selectedShape = selection.ShapeRange[1];

                if (selectedShape.Type == Office.MsoShapeType.msoTextBox)
                {
                    selectedShape.TextFrame.TextRange.Text = word;
                }
                else
                {
                    Clipboard.SetText(word);
                    MessageBox.Show("已复制到剪贴板: " + word);
                }
            }
            else
            {
                Clipboard.SetText(word);
                MessageBox.Show("已复制到剪贴板: " + word);
            }
        }
    }
}