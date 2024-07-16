using System.Collections.Generic;
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
        private List<TextBlock> multiPronunciationTextBlocks = new List<TextBlock>();

        public ZhuYinEditor()
        {
            InitializeComponent();
            LoadHanziPinyinDict();
            LoadMultiPronunciationDict();
            TextBoxLeft.KeyDown += TextBoxLeft_KeyDown;
            TextBoxLeft.TextChanged += TextBoxLeft_TextChanged;

            // 添加右键菜单
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
                UpdateStackPanelContent(GetPlainTextFromTextBox());
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
                UpdateStackPanelContent(GetPlainTextFromTextBox());
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
                TextBoxLeft.Text = text;
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            ExportToTable();
        }

        private void BtnCorrectPronunciations_Click(object sender, RoutedEventArgs e)
        {
            CorrectPronunciations();
        }

        private void CorrectPronunciations()
        {
            string text = GetPlainTextFromTextBox();
            UpdateStackPanelContentWithCorrection(text);
        }

        private void UpdateStackPanelContentWithCorrection(string text)
        {
            StackPanelContent.Children.Clear();
            StackPanel currentLinePanel = CreateNewLinePanel();

            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == '\n')
                {
                    StackPanelContent.Children.Add(currentLinePanel);
                    currentLinePanel = CreateNewLinePanel();
                }
                else
                {
                    if (currentLinePanel.Children.Count >= MaxCharsPerLine)
                    {
                        StackPanelContent.Children.Add(currentLinePanel);
                        currentLinePanel = CreateNewLinePanel();
                    }

                    string wordToCheck = GetWordToCheck(text, i);
                    if (multiPronunciationDict.ContainsKey(wordToCheck))
                    {
                        string[] pinyinArray = multiPronunciationDict[wordToCheck].Split(' ');
                        for (int j = 0; j < wordToCheck.Length; j++)
                        {
                            StackPanel sp = CreateCharacterPanel(wordToCheck[j], null, text.Length > i + wordToCheck.Length ? text[i + wordToCheck.Length].ToString() : null);
                            currentLinePanel.Children.Add(sp);
                        }
                        i += wordToCheck.Length - 1;
                    }
                    else
                    {
                        StackPanel sp = CreateCharacterPanel(text[i], i > 0 ? text[i - 1].ToString() : null, i < text.Length - 1 ? text[i + 1].ToString() : null);
                        currentLinePanel.Children.Add(sp);
                    }
                }
            }

            StackPanelContent.Children.Add(currentLinePanel);
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

        private void ExportToTable()
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

                    pinyinLines.Add(pinyinList);
                    hanziLines.Add(hanziList);
                }
            }

            int rowCount = pinyinLines.Count * 2;
            int columnCount = pinyinLines.Max(line => line.Count);

            PowerPoint.Table table = activeSlide.Shapes.AddTable(rowCount, columnCount).Table;

            float hanziFontSize = 20;
            float pinyinFontSize = hanziFontSize * 0.5f;

            // 创建一个二维数组来存储表格内容
            var content = new string[rowCount, columnCount];

            for (int i = 0; i < pinyinLines.Count; i++)
            {
                for (int j = 0; j < pinyinLines[i].Count; j++)
                {
                    content[i * 2, j] = pinyinLines[i][j];
                    content[i * 2 + 1, j] = hanziLines[i][j];
                }
            }

            // 批量设置表格内容
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
                }
            }

            // 删除空白列
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

            // 删除空白行
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

            // 应用样式和属性
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
        }

        private void TextBoxLeft_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Tab)
            {
                e.Handled = true;
                int caretIndex = TextBoxLeft.CaretIndex;
                TextBoxLeft.Text = TextBoxLeft.Text.Insert(caretIndex, "　　"); // 插入中文全角空格
                TextBoxLeft.CaretIndex = caretIndex + 2;
            }
            else if (e.Key == Key.Enter)
            {
                UpdateStackPanelContent(GetPlainTextFromTextBox());
            }
        }

        private void TextBoxLeft_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdateStackPanelContent(GetPlainTextFromTextBox());
        }

        private string GetPlainTextFromTextBox()
        {
            return TextBoxLeft.Text;
        }

        private void UpdateStackPanelContent(string text)
        {
            StackPanelContent.Children.Clear();
            StackPanel currentLinePanel = CreateNewLinePanel();

            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == '\n')
                {
                    StackPanelContent.Children.Add(currentLinePanel);
                    currentLinePanel = CreateNewLinePanel();
                }
                else
                {
                    if (currentLinePanel.Children.Count >= MaxCharsPerLine)
                    {
                        StackPanelContent.Children.Add(currentLinePanel);
                        currentLinePanel = CreateNewLinePanel();
                    }

                    string previousChar = i > 0 ? text[i - 1].ToString() : null;
                    string nextChar = i < text.Length - 1 ? text[i + 1].ToString() : null;

                    StackPanel sp = CreateCharacterPanel(text[i], previousChar, nextChar);
                    currentLinePanel.Children.Add(sp);
                }
            }

            StackPanelContent.Children.Add(currentLinePanel);
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

        private StackPanel CreateCharacterPanel(char c, string previousChar = null, string nextChar = null)
        {
            StackPanel sp = new StackPanel
            {
                Orientation = Orientation.Vertical,
                Margin = new Thickness(5, 0, 5, 0),
                HorizontalAlignment = HorizontalAlignment.Center
            };

            // 特殊处理汉字“一”
            string pinyin = null;
            if (c == '一')
            {
                // 判断前一个字符是否为“第、其、专、任”中的一个
                if (previousChar != null && "第其专任唯无万不如非为若归说十合惟当失挂".Contains(previousChar))
                {
                    pinyin = "yī";
                }
                else if (nextChar != null)
                {
                    pinyin = GetCorrectPinyinForYi(nextChar);
                }
            }

            // 处理汉字和标点符号
            if (pinyin != null || hanziPinyinDict.ContainsKey(c.ToString()))
            {
                sp.Children.Add(new TextBlock
                {
                    Text = pinyin ?? hanziPinyinDict[c.ToString()][0],
                    FontSize = 10,
                    TextAlignment = TextAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center
                });
            }
            else
            {
                sp.Children.Add(new TextBlock
                {
                    Text = string.Empty,
                    FontSize = 10,
                    TextAlignment = TextAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center
                });
            }

            sp.Children.Add(new TextBlock
            {
                Text = c.ToString(),
                FontSize = 20,
                TextAlignment = TextAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center
            });

            return sp;
        }

        private string GetCorrectPinyinForYi(string nextChar)
        {
            string pinyin = "yi";
            if (hanziPinyinDict.ContainsKey(nextChar))
            {
                var pinyins = hanziPinyinDict[nextChar];
                foreach (var p in pinyins)
                {
                    if ("āēīōūǖáéíóúǘǎěǐǒǔǚ".Any(x => p.Contains(x)))
                    {
                        return "yì";
                    }
                    if ("àèìòùǜ".Any(x => p.Contains(x)))
                    {
                        return "yí";
                    }
                }
            }
            return pinyin;
        }
    }
}
