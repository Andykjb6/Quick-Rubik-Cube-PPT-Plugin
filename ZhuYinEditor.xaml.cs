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

namespace 课件帮PPT助手
{
    public partial class ZhuYinEditor : Window
    {
        private const int DefaultMaxCharsPerLine = 20;
        private const double DefaultOddLineSpacing = 1.2;
        private int MaxCharsPerLine = DefaultMaxCharsPerLine;
        private double OddLineSpacing = DefaultOddLineSpacing;
        private Dictionary<string, List<string>> hanziPinyinDict;
        private List<TextBlock> multiPronunciationTextBlocks = new List<TextBlock>();

        public ZhuYinEditor()
        {
            InitializeComponent();
            LoadHanziPinyinDict();
            RichTextBoxContent.KeyDown += RichTextBoxContent_KeyDown;

            // 添加右键菜单
            var contextMenu = new ContextMenu();

            var setMaxCharsItem = new MenuItem { Header = "设置每行字符数" };
            setMaxCharsItem.Click += SetMaxCharsItem_Click;
            contextMenu.Items.Add(setMaxCharsItem);

            var setOddLineSpacingItem = new MenuItem { Header = "设置文本行间距" };
            setOddLineSpacingItem.Click += SetOddLineSpacingItem_Click;
            contextMenu.Items.Add(setOddLineSpacingItem);

            ContextMenu = contextMenu;
            MouseRightButtonUp += ZhuYinEditor_MouseRightButtonUp;
        }

        private void ZhuYinEditor_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            ContextMenu.IsOpen = true;
        }

        private void SetMaxCharsItem_Click(object sender, RoutedEventArgs e)
        {
            var inputBox = new InputBox("设置每行最大字符数", MaxCharsPerLine.ToString());
            if (inputBox.ShowDialog() == true)
            {
                if (int.TryParse(inputBox.InputText, out int maxChars))
                {
                    MaxCharsPerLine = maxChars;
                    UpdateStackPanelContent(GetPlainTextFromRichTextBox());
                }
            }
        }

        private void SetOddLineSpacingItem_Click(object sender, RoutedEventArgs e)
        {
            var inputBox = new InputBox("设置段落文本行间距", OddLineSpacing.ToString());
            if (inputBox.ShowDialog() == true)
            {
                if (double.TryParse(inputBox.InputText, out double spacing))
                {
                    OddLineSpacing = spacing;
                    UpdateStackPanelContent(GetPlainTextFromRichTextBox());
                }
            }
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

        private void BtnAddPronunciations_Click(object sender, RoutedEventArgs e)
        {
            UpdateStackPanelContent(GetPlainTextFromRichTextBox());
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
                RichTextBoxContent.Document.Blocks.Clear();
                RichTextBoxContent.Document.Blocks.Add(new Paragraph(new Run(text)));
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            ExportToTable();
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

            // 填充表格内容
            for (int i = 0; i < pinyinLines.Count; i++)
            {
                for (int j = 0; j < pinyinLines[i].Count; j++)
                {
                    table.Cell(i * 2 + 1, j + 1).Shape.TextFrame.TextRange.Text = pinyinLines[i][j];
                    table.Cell(i * 2 + 1, j + 1).Shape.TextFrame.TextRange.Font.Size = pinyinFontSize;
                    table.Cell(i * 2 + 1, j + 1).Shape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(System.Drawing.Color.Black);

                    table.Cell(i * 2 + 2, j + 1).Shape.TextFrame.TextRange.Text = hanziLines[i][j];
                    table.Cell(i * 2 + 2, j + 1).Shape.TextFrame.TextRange.Font.Size = hanziFontSize;
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

        private void RichTextBoxContent_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Tab)
            {
                e.Handled = true;
                var caretPosition = RichTextBoxContent.CaretPosition;
                var tabRun = new Run("　　"); // 中文全角空格，等宽于汉字
                RichTextBoxContent.CaretPosition.InsertTextInRun(tabRun.Text);
                RichTextBoxContent.CaretPosition = caretPosition.GetNextInsertionPosition(LogicalDirection.Forward);
            }
            else if (e.Key == Key.Enter)
            {
                UpdateStackPanelContent(GetPlainTextFromRichTextBox());
            }
        }

        private void RichTextBoxContent_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            UpdateStackPanelContent(GetPlainTextFromRichTextBox());
        }

        private string GetPlainTextFromRichTextBox()
        {
            System.Windows.Documents.TextRange textRange = new System.Windows.Documents.TextRange(RichTextBoxContent.Document.ContentStart, RichTextBoxContent.Document.ContentEnd);
            return textRange.Text;
        }

        private void UpdateStackPanelContent(string text)
        {
            StackPanelContent.Children.Clear();
            StackPanel currentLinePanel = CreateNewLinePanel();

            foreach (char c in text)
            {
                if (c == '\n')
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

                    StackPanel sp = CreateCharacterPanel(c);
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

        private StackPanel CreateCharacterPanel(char c)
        {
            StackPanel sp = new StackPanel
            {
                Orientation = Orientation.Vertical,
                Margin = new Thickness(5, 0, 5, 0),
                HorizontalAlignment = HorizontalAlignment.Center
            };

            // 处理汉字和标点符号
            if (hanziPinyinDict.ContainsKey(c.ToString()))
            {
                sp.Children.Add(new TextBlock
                {
                    Text = hanziPinyinDict[c.ToString()][0],
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
    }
}

// InputBox 窗口类，用于输入配置值
public class InputBox : Window
{
    public string InputText { get; private set; }

    public InputBox(string title, string defaultText)
    {
        Title = title;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        Width = 300;
        Height = 150;

        var stackPanel = new StackPanel();
        var textBox = new TextBox { Text = defaultText, Margin = new Thickness(10) };
        var button = new Button { Content = "确定", Margin = new Thickness(10), IsDefault = true };

        button.Click += (sender, e) =>
        {
            InputText = textBox.Text;
            DialogResult = true;
        };

        stackPanel.Children.Add(textBox);
        stackPanel.Children.Add(button);

        Content = stackPanel;
    }
}