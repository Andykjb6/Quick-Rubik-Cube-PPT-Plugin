﻿using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using Microsoft.Win32;
using OfficeOpenXml;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Drawing;
using Microsoft.Office.Core;
using System;
using System.Threading.Tasks;

namespace 课件帮PPT助手
{
    public partial class ZhuYinEditor : Window
    {
        private const int DefaultMaxCharsPerLine = 20;
        private const double DefaultOddLineSpacing = 1.2;
        private int MaxCharsPerLine = DefaultMaxCharsPerLine;
        private double OddLineSpacing = DefaultOddLineSpacing;
        private Dictionary<string, List<string>> hanziPinyinDict;
        private Dictionary<string, string> multiPronunciationDict;
        private Dictionary<int, string> correctedPinyinDict = new Dictionary<int, string>();
        private List<TextBlock> multiPronunciationTextBlocks = new List<TextBlock>();
        private bool isTextChangedEventHandlerActive = true;

        public ZhuYinEditor()
        {
            InitializeComponent();
            LoadHanziPinyinDict();
            LoadMultiPronunciationDict();
            RichTextBoxLeft.KeyDown += RichTextBoxLeft_KeyDown;
            RichTextBoxLeft.TextChanged += RichTextBoxLeft_TextChanged;

            var contextMenu = new ContextMenu();

            var setMaxCharsMenuItem = new MenuItem { Header = "设置每行字符数" };
            var maxCharsPanel = new StackPanel { Orientation = Orientation.Horizontal };
            var maxCharsSlider = new Slider
            {
                Minimum = 1,
                Maximum = 100,
                Value = MaxCharsPerLine,
                Width = 100,
                Margin = new Thickness(5, 0, 5, 0)
            };
            var maxCharsValueText = new TextBlock { Text = MaxCharsPerLine.ToString(), VerticalAlignment = VerticalAlignment.Center };
            maxCharsSlider.ValueChanged += (sender, e) =>
            {
                MaxCharsPerLine = (int)e.NewValue;
                maxCharsValueText.Text = MaxCharsPerLine.ToString();
                UpdateStackPanelContent(GetPlainTextFromRichTextBox());
                SyncAlignmentWithPinyin();
            };
            maxCharsPanel.Children.Add(maxCharsSlider);
            maxCharsPanel.Children.Add(maxCharsValueText);
            setMaxCharsMenuItem.Items.Add(maxCharsPanel);

            var setOddLineSpacingMenuItem = new MenuItem { Header = "设置文本行间距" };
            var oddLineSpacingPanel = new StackPanel { Orientation = Orientation.Horizontal };
            var oddLineSpacingSlider = new Slider
            {
                Minimum = 0.5,
                Maximum = 3,
                Value = OddLineSpacing,
                Width = 100,
                Margin = new Thickness(5, 0, 5, 0)
            };
            var oddLineSpacingValueText = new TextBlock { Text = OddLineSpacing.ToString("F1"), VerticalAlignment = VerticalAlignment.Center };
            oddLineSpacingSlider.ValueChanged += (sender, e) =>
            {
                OddLineSpacing = e.NewValue;
                oddLineSpacingValueText.Text = OddLineSpacing.ToString("F1");
                UpdateStackPanelContent(GetPlainTextFromRichTextBox());
                SyncAlignmentWithPinyin();
            };
            oddLineSpacingPanel.Children.Add(oddLineSpacingSlider);
            oddLineSpacingPanel.Children.Add(oddLineSpacingValueText);
            setOddLineSpacingMenuItem.Items.Add(oddLineSpacingPanel);

            contextMenu.Items.Add(setMaxCharsMenuItem);
            contextMenu.Items.Add(setOddLineSpacingMenuItem);

            ContextMenu = contextMenu;
            MouseRightButtonUp += ZhuYinEditor_MouseRightButtonUp;
        }

        private void ZhuYinEditor_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            ContextMenu.IsOpen = true;
        }

        private void LoadHanziPinyinDict()
        {
            hanziPinyinDict = new Dictionary<string, List<string>>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string filePath = ExtractEmbeddedResource("课件帮PPT助手.汉字字典.汉字拼音信息库.xlsx");

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    string hanzi = worksheet.Cells[row, 1].Text;
                    string[] pinyins = worksheet.Cells[row, 2].Text.Split(',');
                    hanziPinyinDict[hanzi] = pinyins.ToList();
                }
            }
        }

        private void LoadMultiPronunciationDict()
        {
            multiPronunciationDict = new Dictionary<string, string>();
            string filePath = ExtractEmbeddedResource("课件帮PPT助手.汉字字典.多音字词语.txt");

            foreach (var line in File.ReadLines(filePath))
            {
                var parts = line.Split(new[] { '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 2)
                {
                    string word = parts[0].Trim();
                    string pinyin = parts[1].Trim();
                    multiPronunciationDict[word] = pinyin;
                }
            }
        }

        private string ExtractEmbeddedResource(string resourceName)
        {
            string tempFile = Path.Combine(Path.GetTempPath(), resourceName);
            using (var resource = typeof(ZhuYinEditor).Assembly.GetManifestResourceStream(resourceName))
            {
                using (var file = new FileStream(tempFile, FileMode.Create, FileAccess.Write))
                {
                    resource.CopyTo(file);
                }
            }
            return tempFile;
        }

        private void BtnDetectMultiPronunciations_Click(object sender, RoutedEventArgs e)
        {
            DetectMultiPronunciations();
        }

        private void DetectMultiPronunciations()
        {
            multiPronunciationTextBlocks.Clear();
            foreach (var child in StackPanelContent.Children)
            {
                if (child is StackPanel sp)
                {
                    foreach (var element in sp.Children)
                    {
                        if (element is StackPanel charPanel)
                        {
                            var hanziBlock = charPanel.Children[1] as TextBlock;
                            string hanzi = hanziBlock.Text;
                            if (hanziPinyinDict.ContainsKey(hanzi) && hanziPinyinDict[hanzi].Count > 1)
                            {
                                hanziBlock.Background = System.Windows.Media.Brushes.Yellow;
                                hanziBlock.MouseLeftButtonUp += HanziBlock_MouseLeftButtonUp;
                                multiPronunciationTextBlocks.Add(hanziBlock);
                            }
                        }
                    }
                }
            }
        }

        private void HanziBlock_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (sender is TextBlock hanziBlock)
            {
                string hanzi = hanziBlock.Text;
                if (hanziPinyinDict.ContainsKey(hanzi))
                {
                    var menu = new ContextMenu();
                    foreach (var pinyin in hanziPinyinDict[hanzi])
                    {
                        var menuItem = new MenuItem
                        {
                            Header = pinyin,
                            Tag = hanziBlock
                        };
                        menuItem.Click += MenuItem_Click;
                        menu.Items.Add(menuItem);
                    }

                    // 添加儿化音选项
                    if (hanziBlock.Parent is StackPanel charPanel && charPanel.Parent is StackPanel linePanel)
                    {
                        int charPanelIndex = linePanel.Children.IndexOf(charPanel);
                        if (charPanelIndex + 1 < linePanel.Children.Count)
                        {
                            var nextCharPanel = linePanel.Children[charPanelIndex + 1] as StackPanel;
                            var nextHanziBlock = nextCharPanel.Children[1] as TextBlock;
                            if (nextHanziBlock.Text == "儿")
                            {
                                var erPinyin = hanziPinyinDict[hanzi][0] + "r";
                                var erMenuItem = new MenuItem
                                {
                                    Header = erPinyin,
                                    Tag = hanziBlock
                                };
                                erMenuItem.Click += MenuItem_Click;
                                menu.Items.Add(erMenuItem);
                            }
                        }
                    }

                    hanziBlock.ContextMenu = menu;
                    hanziBlock.ContextMenu.IsOpen = true;
                }
            }
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem menuItem)
            {
                if (menuItem.Tag is TextBlock hanziBlock)
                {
                    var charPanel = hanziBlock.Parent as StackPanel;
                    if (charPanel != null)
                    {
                        var pinyinBlock = charPanel.Children[0] as TextBlock;
                        pinyinBlock.Text = menuItem.Header.ToString();

                        // 保存纠正后的拼音到字典
                        int index = GetCharacterIndex(charPanel);
                        if (index >= 0)
                        {
                            correctedPinyinDict[index] = pinyinBlock.Text;
                        }

                        // 如果选择了儿化音，则清空"儿"字的拼音
                        if (pinyinBlock.Text.EndsWith("r") && charPanel.Parent is StackPanel linePanel)
                        {
                            int charPanelIndex = linePanel.Children.IndexOf(charPanel);
                            if (charPanelIndex + 1 < linePanel.Children.Count)
                            {
                                var nextCharPanel = linePanel.Children[charPanelIndex + 1] as StackPanel;
                                var nextPinyinBlock = nextCharPanel.Children[0] as TextBlock;
                                var nextHanziBlock = nextCharPanel.Children[1] as TextBlock;
                                if (nextHanziBlock.Text == "儿")
                                {
                                    nextPinyinBlock.Text = string.Empty;
                                }
                            }
                        }
                        else if (!pinyinBlock.Text.EndsWith("r") && charPanel.Parent is StackPanel linePanelForEr)
                        {
                            int charPanelIndex = linePanelForEr.Children.IndexOf(charPanel);
                            if (charPanelIndex + 1 < linePanelForEr.Children.Count)
                            {
                                var nextCharPanel = linePanelForEr.Children[charPanelIndex + 1] as StackPanel;
                                var nextPinyinBlock = nextCharPanel.Children[0] as TextBlock;
                                var nextHanziBlock = nextCharPanel.Children[1] as TextBlock;
                                if (nextHanziBlock.Text == "儿")
                                {
                                    nextPinyinBlock.Text = "ér"; // 恢复"儿"字的拼音
                                }
                            }
                        }
                    }
                }
            }
        }

        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "文本文件|*.txt",
                Title = "选择文本文件"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                string text = File.ReadAllText(openFileDialog.FileName);
                RichTextBoxLeft.Document.Blocks.Clear();
                RichTextBoxLeft.Document.Blocks.Add(new Paragraph(new Run(text)));
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            ExportToTable();
        }

        private void BtnCorrectPronunciations_Click(object sender, RoutedEventArgs e)
        {
            CorrectPronunciations();
            SyncAlignmentWithPinyin();
        }

        private void CorrectPronunciations()
        {
            string text = GetPlainTextFromRichTextBox();
            UpdateStackPanelContentWithCorrection(text);
            SyncAlignmentWithPinyin();
            HighlightErhuaHanzi();
        }

        private void UpdateStackPanelContentWithCorrection(string text)
        {
            StackPanelContent.Children.Clear();
            StackPanel currentLinePanel = CreateNewLinePanel();
            int currentCharCount = 0; // 当前行字符计数
            int i = 0;

            while (i < text.Length)
            {
                if (text[i] == '\n' || currentCharCount >= MaxCharsPerLine)
                {
                    StackPanelContent.Children.Add(currentLinePanel);
                    currentLinePanel = CreateNewLinePanel();
                    currentCharCount = 0; // 重置当前行字符计数
                    if (text[i] == '\n')
                    {
                        i++;
                        continue; // 跳过换行符
                    }
                }

                string wordToCheck = GetWordToCheck(text, i);
                if (multiPronunciationDict.ContainsKey(wordToCheck))
                {
                    string[] pinyinArray = multiPronunciationDict[wordToCheck].Split(' ');
                    for (int j = 0; j < wordToCheck.Length; j++)
                    {
                        if (currentCharCount >= MaxCharsPerLine)
                        {
                            StackPanelContent.Children.Add(currentLinePanel);
                            currentLinePanel = CreateNewLinePanel();
                            currentCharCount = 0;
                        }

                        StackPanel sp = CreateCharacterPanel(wordToCheck[j], pinyinArray[j]);
                        currentLinePanel.Children.Add(sp);

                        // 保存纠正后的拼音到字典
                        int index = GetCharacterIndex(sp);
                        correctedPinyinDict[index] = pinyinArray[j];
                        currentCharCount++; // 更新当前行字符计数
                    }
                    i += wordToCheck.Length - 1;
                }
                else
                {
                    char currentChar = text[i];
                    string pinyin = GetCorrectedPinyin(text, i, i == text.Length - 1);

                    if (correctedPinyinDict.ContainsKey(i))
                    {
                        // 使用纠正后的拼音
                        pinyin = correctedPinyinDict[i];
                    }

                    if (currentCharCount >= MaxCharsPerLine)
                    {
                        StackPanelContent.Children.Add(currentLinePanel);
                        currentLinePanel = CreateNewLinePanel();
                        currentCharCount = 0;
                    }

                    StackPanel sp = CreateCharacterPanel(currentChar, pinyin);
                    currentLinePanel.Children.Add(sp);
                    currentCharCount++; // 更新当前行字符计数

                    // 保存纠正后的拼音到字典
                    int index = GetCharacterIndex(sp);
                    correctedPinyinDict[index] = pinyin;

                    // 如果当前汉字为“儿”，且前一个汉字的拼音末尾包含字符“r”，则将当前汉字拼音留空
                    if (currentChar == '儿' && i > 0)
                    {
                        var prevCharPanel = currentLinePanel.Children[currentLinePanel.Children.Count - 2] as StackPanel;
                        var prevPinyinBlock = prevCharPanel.Children[0] as TextBlock;
                        if (prevPinyinBlock.Text.EndsWith("r"))
                        {
                            (sp.Children[0] as TextBlock).Text = string.Empty;
                        }
                    }
                }
                i++;
            }

            if (currentLinePanel.Children.Count > 0)
            {
                StackPanelContent.Children.Add(currentLinePanel);
            }
        }


        private string GetCorrectedPinyin(string text, int index, bool isLastChar)
        {
            char currentChar = text[index];

            if (currentChar == '哇')
            {
                if (index > 0 && index < text.Length - 1 && text[index - 1] == text[index + 1])
                {
                    return "wa";
                }
            }

            if (currentChar == '一')
            {
                if (index < text.Length - 1)
                {
                    string nextCharPinyin = hanziPinyinDict.ContainsKey(text[index + 1].ToString()) ?
                                             hanziPinyinDict[text[index + 1].ToString()][0] : string.Empty;
                    if (nextCharPinyin.IndexOfAny(new char[] { 'ā', 'ō', 'ē', 'ī', 'ū', 'ǖ', 'á', 'ó', 'é', 'ú', 'ǘ', 'ǎ', 'ǒ', 'ě', 'ǐ', 'ǔ', 'ǚ' }) >= 0)
                    {
                        return "yì";
                    }
                    else if (nextCharPinyin.IndexOfAny(new char[] { 'à', 'ò', 'è', 'ì', 'ù', 'ǜ' }) >= 0)
                    {
                        return "yí";
                    }
                }

                if (index > 0)
                {
                    char prevChar = text[index - 1];
                    if (new string[] { "第", "其", "专", "任", "唯", "无", "万", "不", "如", "非", "为", "若", "归", "说", "十", "合", "惟", "当", "失", "挂" }.Contains(prevChar.ToString()))
                    {
                        return "yī";
                    }
                }
            }

            // 默认处理其他汉字
            return hanziPinyinDict.ContainsKey(currentChar.ToString()) ? hanziPinyinDict[currentChar.ToString()][0] : string.Empty;
        }

        private string GetWordToCheck(string text, int startIndex)
        {
            int maxLength = multiPronunciationDict.Keys.Max(k => k.Length);
            for (int length = maxLength; length > 0; length--)
            {
                if (startIndex + length <= text.Length)
                {
                    string substring = text.Substring(startIndex, length);
                    if (multiPronunciationDict.ContainsKey(substring))
                    {
                        return substring;
                    }
                }
            }
            return text[startIndex].ToString();
        }

        private void HighlightErhuaHanzi()
        {
            foreach (var child in StackPanelContent.Children)
            {
                if (child is StackPanel sp)
                {
                    for (int i = 1; i < sp.Children.Count; i++)
                    {
                        var charPanel = sp.Children[i] as StackPanel;
                        var hanziBlock = charPanel.Children[1] as TextBlock;
                        if (hanziBlock.Text == "儿")
                        {
                            var prevCharPanel = sp.Children[i - 1] as StackPanel;
                            var prevHanziBlock = prevCharPanel.Children[1] as TextBlock;
                            prevHanziBlock.Background = System.Windows.Media.Brushes.LightGreen;
                            prevHanziBlock.MouseLeftButtonUp += HanziBlock_MouseLeftButtonUp;
                        }
                    }
                }
            }
        }

        private async void ExportToTable()
        {
            var pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Slide activeSlide = pptApp.ActiveWindow.View.Slide as PowerPoint.Slide;

            List<List<string>> pinyinLines = new List<List<string>>();
            List<List<string>> hanziLines = new List<List<string>>();

            foreach (var child in StackPanelContent.Children)
            {
                if (child is StackPanel sp)
                {
                    List<string> pinyinList = new List<string>();
                    List<string> hanziList = new List<string>();

                    foreach (var element in sp.Children)
                    {
                        if (element is StackPanel charPanel)
                        {
                            var pinyinBlock = charPanel.Children[0] as TextBlock;
                            var hanziBlock = charPanel.Children[1] as TextBlock;

                            pinyinList.Add(pinyinBlock.Text);
                            hanziList.Add(hanziBlock.Text);
                        }
                    }

                    // 补足每行的单元格数量
                    while (pinyinList.Count < MaxCharsPerLine)
                    {
                        pinyinList.Add(string.Empty);
                        hanziList.Add(string.Empty);
                    }

                    pinyinLines.Add(pinyinList);
                    hanziLines.Add(hanziList);
                }
            }

            int rowCount = pinyinLines.Count * 2;
            int columnCount = MaxCharsPerLine;

            PowerPoint.Table table = activeSlide.Shapes.AddTable(rowCount, columnCount).Table;

            float hanziFontSize = 20;
            float pinyinFontSize = hanziFontSize * 0.5f;

            var content = new string[rowCount, columnCount];

            for (int i = 0; i < pinyinLines.Count; i++)
            {
                for (int j = 0; j < MaxCharsPerLine; j++)
                {
                    content[i * 2, j] = pinyinLines[i][j];
                    content[i * 2 + 1, j] = hanziLines[i][j];

                    if (j > 0 && hanziLines[i][j] == "儿" && !string.IsNullOrEmpty(content[i * 2, j - 1]) && content[i * 2, j - 1].EndsWith("r"))
                    {
                        content[i * 2, j] = string.Empty; // "儿"字不添加拼音
                    }
                }
            }

            ProgressBarExport.Visibility = Visibility.Visible;
            TextBlockProgress.Visibility = Visibility.Visible;
            ProgressBarExport.Value = 0;
            TextBlockProgress.Text = "正在导出...";

            for (int row = 0; row < rowCount; row++)
            {
                for (int col = 0; col < columnCount; col++)
                {
                    var cell = table.Cell(row + 1, col + 1);
                    if (!string.IsNullOrEmpty(content[row, col]))
                    {
                        cell.Shape.TextFrame.TextRange.Text = content[row, col];
                        cell.Shape.TextFrame.TextRange.Font.Size = row % 2 == 0 ? pinyinFontSize : hanziFontSize;
                        if (row % 2 == 0)
                        {
                            cell.Shape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(System.Drawing.Color.Black);
                        }
                    }
                    // Update progress
                    double progress = ((double)(row * columnCount + col) / (rowCount * columnCount)) * 100;
                    ProgressBarExport.Value = progress;
                    TextBlockProgress.Text = $"导出进度: {progress:F2}%";
                    await Task.Delay(10); // Simulate some delay to visualize progress
                }
            }

            AdjustTableSize(table);

            for (int col = columnCount; col >= 1; col--)
            {
                bool isColumnEmpty = true;
                for (int row = 1; row <= rowCount; row++)
                {
                    if (!string.IsNullOrWhiteSpace(table.Cell(row, col).Shape.TextFrame.TextRange.Text))
                    {
                        isColumnEmpty = false;
                        break;
                    }
                }
                if (isColumnEmpty)
                {
                    table.Columns[col].Delete();
                    columnCount--;
                }
            }

            for (int row = rowCount; row >= 1; row--)
            {
                bool isRowEmpty = true;
                for (int col = 1; col <= columnCount; col++)
                {
                    if (!string.IsNullOrWhiteSpace(table.Cell(row, col).Shape.TextFrame.TextRange.Text))
                    {
                        isRowEmpty = false;
                        break;
                    }
                }
                if (isRowEmpty)
                {
                    table.Rows[row].Delete();
                    rowCount--;
                }
            }

            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= columnCount; col++)
                {
                    var cell = table.Cell(row, col);
                    cell.Borders[PowerPoint.PpBorderType.ppBorderTop].Weight = 0;
                    cell.Borders[PowerPoint.PpBorderType.ppBorderBottom].Weight = 0;
                    cell.Borders[PowerPoint.PpBorderType.ppBorderLeft].Weight = 0;
                    cell.Borders[PowerPoint.PpBorderType.ppBorderRight].Weight = 0;
                    cell.Shape.Fill.Transparency = 1;
                    cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;

                    cell.Shape.TextFrame.MarginTop = 0;
                    cell.Shape.TextFrame.MarginBottom = (float)(row % 2 == 0 ? 0.5 : 0);
                    cell.Shape.TextFrame.MarginLeft = 0;
                    cell.Shape.TextFrame.MarginRight = 0;

                    if (row % 2 != 0)
                    {
                        var textRange = cell.Shape.TextFrame.TextRange;
                        textRange.ParagraphFormat.SpaceWithin = (float)OddLineSpacing;
                        textRange.Font.Bold = MsoTriState.msoFalse;
                        cell.Shape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorBottom;
                    }
                }
            }

            ProgressBarExport.Value = 100;
            TextBlockProgress.Text = "导出完成!";
            await Task.Delay(2000); // Show completion for 2 seconds
            ProgressBarExport.Visibility = Visibility.Collapsed;
            TextBlockProgress.Visibility = Visibility.Collapsed;
        }

        private void AdjustTableSize(PowerPoint.Table table)
        {
            float maxWidth = 0;

            for (int i = 1; i <= table.Rows.Count; i++)
            {
                float maxHeight = 0;
                for (int j = 1; j <= table.Columns.Count; j++)
                {
                    PowerPoint.Cell cell = table.Cell(i, j);
                    float height = cell.Shape.TextFrame2.TextRange.BoundHeight;
                    float width = cell.Shape.TextFrame2.TextRange.BoundWidth;
                    if (height > maxHeight)
                    {
                        maxHeight = height;
                    }
                    if (width > maxWidth)
                    {
                        maxWidth = width;
                    }
                }
                table.Rows[i].Height = maxHeight + 2; // 增加一点高度防止拥挤
            }

            // 设置所有列的宽度为最大宽度
            for (int j = 1; j <= table.Columns.Count; j++)
            {
                table.Columns[j].Width = maxWidth + 2; // 增加一点宽度防止拥挤
            }
        }

        private void RichTextBoxLeft_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Tab)
            {
                e.Handled = true;
                RichTextBoxLeft.CaretPosition.InsertTextInRun("　　");
            }
            else if (e.Key == Key.Enter)
            {
                UpdateStackPanelContent(GetPlainTextFromRichTextBox());
                SyncAlignmentWithPinyin();
            }
        }

        private void RichTextBoxLeft_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (isTextChangedEventHandlerActive)
            {
                isTextChangedEventHandlerActive = false;

                correctedPinyinDict.Clear(); // 清空纠正的拼音字典
                UpdateStackPanelContent(GetPlainTextFromRichTextBox());
                SyncAlignmentWithPinyin();

                isTextChangedEventHandlerActive = true;
            }
        }

        private string GetPlainTextFromRichTextBox()
        {
            return new TextRange(RichTextBoxLeft.Document.ContentStart, RichTextBoxLeft.Document.ContentEnd).Text;
        }

        private void UpdateStackPanelContent(string text)
        {
            StackPanelContent.Children.Clear();
            StackPanel currentLinePanel = CreateNewLinePanel();
            int currentCharCount = 0; // 当前行字符计数

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '\n' || currentCharCount >= MaxCharsPerLine)
                {
                    StackPanelContent.Children.Add(currentLinePanel);
                    currentLinePanel = CreateNewLinePanel();
                    currentCharCount = 0; // 重置当前行字符计数

                    if (c == '\n')
                    {
                        continue; // 跳过换行符
                    }
                }

                StackPanel sp = CreateCharacterPanel(c);

                // 使用纠正后的拼音
                if (correctedPinyinDict.ContainsKey(i))
                {
                    (sp.Children[0] as TextBlock).Text = correctedPinyinDict[i];
                }
                else
                {
                    string pinyin = GetCorrectedPinyin(text, i, i == text.Length - 1);
                    (sp.Children[0] as TextBlock).Text = pinyin;
                }

                currentLinePanel.Children.Add(sp);
                currentCharCount++; // 更新当前行字符计数
            }

            if (currentLinePanel.Children.Count > 0)
            {
                StackPanelContent.Children.Add(currentLinePanel);
            }

            SyncAlignmentWithPinyin();
            HighlightErhuaHanzi();
        }

        private StackPanel CreateNewLinePanel()
        {
            return new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(0, 5, 0, 0)
            };
        }

        private StackPanel CreateCharacterPanel(char c, string pinyin = null)
        {
            StackPanel sp = new StackPanel
            {
                Orientation = Orientation.Vertical,
                Margin = new Thickness(5, 0, 5, 0),
                HorizontalAlignment = HorizontalAlignment.Center
            };

            sp.Children.Add(new TextBlock
            {
                Text = pinyin ?? string.Empty,
                FontSize = 10,
                TextAlignment = TextAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center
            });

            sp.Children.Add(new TextBlock
            {
                Text = c.ToString(),
                FontSize = 20,
                TextAlignment = TextAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center
            });

            return sp;
        }

        private int GetCharacterIndex(StackPanel charPanel)
        {
            var parentPanel = charPanel.Parent as StackPanel;
            int parentIndex = StackPanelContent.Children.IndexOf(parentPanel);
            int charIndex = parentPanel.Children.IndexOf(charPanel);
            return parentIndex * MaxCharsPerLine + charIndex;
        }

        private void AlignText(TextAlignment alignment)
        {
            TextSelection selection = RichTextBoxLeft.Selection;
            if (selection.IsEmpty)
            {
                foreach (Paragraph paragraph in RichTextBoxLeft.Document.Blocks.OfType<Paragraph>())
                {
                    paragraph.TextAlignment = alignment;
                }
                foreach (StackPanel linePanel in StackPanelContent.Children)
                {
                    linePanel.HorizontalAlignment = ConvertToHorizontalAlignment(alignment);
                }
            }
            else
            {
                ApplyAlignmentToParagraphs(alignment);
            }
            SyncAlignmentWithPinyin();
        }

        private void ApplyAlignmentToParagraphs(TextAlignment alignment)
        {
            TextPointer start = RichTextBoxLeft.Selection.Start;
            TextPointer end = RichTextBoxLeft.Selection.End;

            while (start.CompareTo(end) < 0)
            {
                Paragraph paragraph = start.Paragraph;
                if (paragraph != null)
                {
                    paragraph.TextAlignment = alignment;
                }
                start = start.GetNextContextPosition(LogicalDirection.Forward);
            }
        }

        private HorizontalAlignment ConvertToHorizontalAlignment(TextAlignment alignment)
        {
            switch (alignment)
            {
                case TextAlignment.Left:
                    return HorizontalAlignment.Left;
                case TextAlignment.Center:
                    return HorizontalAlignment.Center;
                case TextAlignment.Right:
                    return HorizontalAlignment.Right;
                case TextAlignment.Justify:
                    return HorizontalAlignment.Stretch;
                default:
                    return HorizontalAlignment.Left;
            }
        }

        private void SyncAlignmentWithPinyin()
        {
            // 根据左侧RichTextBox的段落对齐方式，同步右侧拼音区的对齐方式
            var leftParagraphs = RichTextBoxLeft.Document.Blocks.OfType<Paragraph>().ToList();
            var rightPanels = StackPanelContent.Children.OfType<StackPanel>().ToList();

            for (int i = 0; i < leftParagraphs.Count && i < rightPanels.Count; i++)
            {
                var alignment = leftParagraphs[i].TextAlignment;
                rightPanels[i].HorizontalAlignment = ConvertToHorizontalAlignment(alignment);
            }
        }

        private void BtnAlignLeft_Click(object sender, RoutedEventArgs e)
        {
            AlignText(TextAlignment.Left);
        }

        private void BtnAlignCenter_Click(object sender, RoutedEventArgs e)
        {
            AlignText(TextAlignment.Center);
        }

        private void BtnAlignJustify_Click(object sender, RoutedEventArgs e)
        {
            AlignText(TextAlignment.Justify);
        }
    }
}
