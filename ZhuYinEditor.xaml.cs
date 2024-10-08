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
using Microsoft.Office.Interop.PowerPoint;

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
  
        private Dictionary<string, string> lightToneDict; // 新的轻声词语库字典
        private const int MaxCharCountThreshold = 300; // 设置字符数阈值

        public ZhuYinEditor()
        {
            InitializeComponent();
            LoadAdjectiveDict(); // 确保在窗口初始化时加载形容词字典
            LoadVerbDict(); // 加载动词字典
            LoadLightToneDict(); // 加载轻声词语库
            // 在初始化时应用默认的字体设置
            ApplyDefaultFontSettings();

            LoadHanziPinyinDict();
            LoadMultiPronunciationDict();
            LoadVerbDict();  // 加载动词字典
            LoadErhuaRules(); // 加载儿化音处理规则（前缀和后缀）
            LoadDePrefixAndSuffix();  // 加载得字前后缀字典
            RichTextBoxLeft.KeyDown += RichTextBoxLeft_KeyDown;
            RichTextBoxLeft.TextChanged += RichTextBoxLeft_TextChanged;

            InitializeContextMenu();
        }

        private HashSet<string> erPrefixSet;
        private HashSet<string> erSuffixSet;

        public void LoadErhuaRules()
        {
            string prefixFilePath = ExtractEmbeddedResource("课件帮PPT助手.汉字字典.儿字前缀.txt");
            string suffixFilePath = ExtractEmbeddedResource("课件帮PPT助手.汉字字典.儿字后缀.txt");

            erPrefixSet = LoadErhuaExceptions(prefixFilePath);
            erSuffixSet = LoadErhuaExceptions(suffixFilePath);
        }
        private HashSet<string> verbSet;

        private HashSet<string> dePrefixSet;
        private HashSet<string> deSuffixSet;

        private void LoadDePrefixAndSuffix()
        {
            dePrefixSet = LoadWordSetFromEmbeddedResource("课件帮PPT助手.汉字字典.得字前缀.txt");
            deSuffixSet = LoadWordSetFromEmbeddedResource("课件帮PPT助手.汉字字典.得字后缀.txt");
        }

        private HashSet<string> LoadWordSetFromEmbeddedResource(string resourceName)
        {
            var wordSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string filePath = ExtractEmbeddedResource(resourceName);

            if (!string.IsNullOrEmpty(filePath))
            {
                foreach (var line in File.ReadLines(filePath))
                {
                    string word = line.Trim();
                    if (!string.IsNullOrEmpty(word))
                    {
                        wordSet.Add(word);
                    }
                }
            }

            return wordSet;
        }

        

        private void LoadVerbDict()
        {
            verbSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase); // 使用不区分大小写的比较器
            string filePath = ExtractEmbeddedResource("课件帮PPT助手.汉字字典.汉字动词大全.txt");

            if (File.Exists(filePath))
            {
                foreach (var line in File.ReadLines(filePath))
                {
                    string verb = line.Trim();
                    if (!string.IsNullOrEmpty(verb))
                    {
                        verbSet.Add(verb); // 将每个非空的动词加入集合中
                    }
                }
            }
        }

        private HashSet<string> adjectiveSet; // 定义形容词集合

        private void LoadAdjectiveDict()
        {
            adjectiveSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string filePath = ExtractEmbeddedResource("课件帮PPT助手.汉字字典.汉字形容词大全.txt");

            foreach (var line in File.ReadLines(filePath))
            {
                string adjective = line.Trim();
                if (!string.IsNullOrEmpty(adjective))
                {
                    adjectiveSet.Add(adjective);
                }
            }
        }
        private void LoadLightToneDict()
        {
            lightToneDict = new Dictionary<string, string>();
            string filePath = ExtractEmbeddedResource("课件帮PPT助手.汉字字典.轻声词语库.txt");

            foreach (var line in File.ReadLines(filePath))
            {
                var parts = line.Split(new[] { '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 2)
                {
                    string word = parts[0].Trim();
                    string pinyin = parts[1].Trim();
                    lightToneDict[word] = pinyin; // 存储到轻声词语库字典
                }
            }
        }
        private void InitializeContextMenu()
        {
            var contextMenu = new ContextMenu();

            // 创建并配置“设置每行字符数”菜单项
            var setMaxCharsMenuItem = new MenuItem { Header = "设置每行字符数" };
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

            // 将滑块和文本块添加到菜单项
            var maxCharsPanel = new StackPanel { Orientation = Orientation.Horizontal };
            maxCharsPanel.Children.Add(maxCharsSlider);
            maxCharsPanel.Children.Add(maxCharsValueText);
            setMaxCharsMenuItem.Items.Add(maxCharsPanel);

            // 创建并配置“设置文本行间距”菜单项
            var setOddLineSpacingMenuItem = new MenuItem { Header = "设置文本行间距" };
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

            // 将滑块和文本块添加到菜单项
            var oddLineSpacingPanel = new StackPanel { Orientation = Orientation.Horizontal };
            oddLineSpacingPanel.Children.Add(oddLineSpacingSlider);
            oddLineSpacingPanel.Children.Add(oddLineSpacingValueText);
            setOddLineSpacingMenuItem.Items.Add(oddLineSpacingPanel);

            // 添加其他菜单项到 contextMenu
            contextMenu.Items.Add(setMaxCharsMenuItem);
            contextMenu.Items.Add(setOddLineSpacingMenuItem);
            ContextMenu = contextMenu;
            MouseRightButtonUp += ZhuYinEditor_MouseRightButtonUp;
        }
        private void BtnFontSettings_Click(object sender, RoutedEventArgs e)
        {
            // 打开字体设置窗口
            FontSettingsWindow fontSettingsWindow = new FontSettingsWindow();
            if (fontSettingsWindow.ShowDialog() == true)
            {
                // 如果用户点击确认，则应用设置
                ApplyFontSettings(fontSettingsWindow.PinyinFont, fontSettingsWindow.PinyinFontSize,
                                  fontSettingsWindow.HanziFont, fontSettingsWindow.HanziFontSize);
            }
        }

        private void ApplyFontSettings(string pinyinFont, double pinyinFontSize, string hanziFont, double hanziFontSize)
        {
            // 遍历 StackPanelContent 中的所有 TextBlock，应用字体设置
            foreach (var child in StackPanelContent.Children)
            {
                if (child is StackPanel sp)
                {
                    foreach (var element in sp.Children)
                    {
                        if (element is StackPanel charPanel)
                        {
                            var pinyinBlock = charPanel.Children[0] as TextBlock;
                            var hanziBlock = charPanel.Children[1] as TextBlock;

                            if (pinyinBlock != null && hanziBlock != null)
                            {
                                pinyinBlock.FontFamily = new System.Windows.Media.FontFamily(pinyinFont);
                                pinyinBlock.FontSize = pinyinFontSize;
                                hanziBlock.FontFamily = new System.Windows.Media.FontFamily(hanziFont);
                                hanziBlock.FontSize = hanziFontSize;
                            }
                        }
                    }
                }
            }
        }

        private void ApplyDefaultFontSettings()
        {
            // 获取并应用默认设置
            string pinyinFont = Properties.Settings.Default.PinyinFont;
            double pinyinFontSize = Properties.Settings.Default.PinyinFontSize;
            string hanziFont = Properties.Settings.Default.HanziFont;
            double hanziFontSize = Properties.Settings.Default.HanziFontSize;

            ApplyFontSettings(pinyinFont, pinyinFontSize, hanziFont, hanziFontSize);
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

                    // 添加拼音选项到菜单
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

                    // 检查是否为儿化音情况
                    if (hanziBlock.Parent is StackPanel charPanel && charPanel.Parent is StackPanel linePanel)
                    {
                        int charPanelIndex = linePanel.Children.IndexOf(charPanel);
                        if (charPanelIndex + 1 < linePanel.Children.Count)
                        {
                            var nextCharPanel = linePanel.Children[charPanelIndex + 1] as StackPanel;
                            var nextHanziBlock = nextCharPanel.Children[1] as TextBlock;
                            if (nextHanziBlock.Text == "儿")
                            {
                                // 添加儿化音选项
                                var erPinyin = hanziPinyinDict[hanzi][0] + "r";
                                var erMenuItem = new MenuItem
                                {
                                    Header = erPinyin,
                                    Tag = hanziBlock
                                };
                                erMenuItem.Click += (s, ev) =>
                                {
                                    MenuItem_Click(s, ev);
                                };
                                menu.Items.Add(erMenuItem);
                            }
                        }

                        // 检查是否为叠词且背景为粉色高亮
                        int currentIndex = linePanel.Children.IndexOf(charPanel);
                        if (currentIndex > 0)
                        {
                            var prevCharPanel = linePanel.Children[currentIndex - 1] as StackPanel;
                            var prevHanziBlock = prevCharPanel?.Children[1] as TextBlock;

                            if (prevHanziBlock != null && prevHanziBlock.Text == hanziBlock.Text && hanziBlock.Background == System.Windows.Media.Brushes.LightPink)
                            {
                                // 添加轻声选项
                                var lightToneMenuItem = new MenuItem
                                {
                                    Header = RemoveTone(hanziPinyinDict[hanzi][0]), // 假设第一个拼音为标准拼音
                                    Tag = hanziBlock
                                };
                                lightToneMenuItem.Click += MenuItem_Click;
                                menu.Items.Add(lightToneMenuItem);
                            }
                        }
                    }

                    hanziBlock.ContextMenu = menu;
                    hanziBlock.ContextMenu.IsOpen = true;
                }
            }
        }

        private string RemoveTone(string pinyin)
        {
            // 将拼音的声调去除，返回轻声拼音
            return pinyin
                .Replace("ā", "a")
                .Replace("á", "a")
                .Replace("ǎ", "a")
                .Replace("à", "a")
                .Replace("ē", "e")
                .Replace("é", "e")
                .Replace("ě", "e")
                .Replace("è", "e")
                .Replace("ī", "i")
                .Replace("í", "i")
                .Replace("ǐ", "i")
                .Replace("ì", "i")
                .Replace("ō", "o")
                .Replace("ó", "o")
                .Replace("ǒ", "o")
                .Replace("ò", "o")
                .Replace("ū", "u")
                .Replace("ú", "u")
                .Replace("ǔ", "u")
                .Replace("ù", "u")
                .Replace("ǖ", "ü")
                .Replace("ǘ", "ü")
                .Replace("ǚ", "ü")
                .Replace("ǜ", "ü");
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

                        // 处理纠正拼音后的逻辑
                        int index = GetCharacterIndex(charPanel);
                        if (index >= 0)
                        {
                            correctedPinyinDict[index] = pinyinBlock.Text;
                        }

                        var parentPanel = charPanel.Parent as StackPanel;
                        if (parentPanel != null)
                        {
                            int charPanelIndex = parentPanel.Children.IndexOf(charPanel);
                            if (charPanelIndex + 1 < parentPanel.Children.Count)
                            {
                                var nextCharPanel = parentPanel.Children[charPanelIndex + 1] as StackPanel;
                                var nextPinyinBlock = nextCharPanel.Children[0] as TextBlock;
                                var nextHanziBlock = nextCharPanel.Children[1] as TextBlock;

                                if (nextHanziBlock.Text == "儿")
                                {
                                    if (pinyinBlock.Text.EndsWith("r"))
                                    {
                                        nextPinyinBlock.Text = string.Empty; // 清空“儿”字的拼音
                                    }
                                    else
                                    {
                                        nextPinyinBlock.Text = "ér"; // 恢复“儿”字的拼音为“ér”
                                    }
                                }
                            }
                        }

                        // 更新后调用高亮函数
                        HighlightErhuaHanzi();
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
            ContextMenu exportMenu = new ContextMenu();

            MenuItem exportTableMenuItem = new MenuItem() { Header = "导出表格" };
            exportTableMenuItem.Click += (s, args) => ExportToTable();

            MenuItem exportTextMenuItem = new MenuItem() { Header = "导出文本" };
            exportTextMenuItem.Click += (s, args) => ExportToTextBox();

            exportMenu.Items.Add(exportTableMenuItem);
            exportMenu.Items.Add(exportTextMenuItem);

            BtnExport.ContextMenu = exportMenu;
            BtnExport.ContextMenu.IsOpen = true;
        }


        private void ExportTable_Click(object sender, RoutedEventArgs e)
        {
            ExportToTable();
        }

        private void ExportText_Click(object sender, RoutedEventArgs e)
        {
            ExportToTextBox();
        }


        private void BtnCorrectPronunciations_Click(object sender, RoutedEventArgs e)
        {
            CorrectPronunciations();
            SyncAlignmentWithPinyin();
        }

        // 在所有内容更新后调用字体应用
        private void CorrectPronunciations()
        {
            string text = GetPlainTextFromRichTextBox();
            UpdateStackPanelContentWithCorrection(text);
            SyncAlignmentWithPinyin();
            HighlightErhuaHanzi();
            HighlightReduplicatedWords(); // 检测并高亮叠词
            ApplyFontSettings(); // 这里再次应用字体设置
        }

        private void UpdateStackPanelContentWithCorrection(string text)
        {
            StackPanelContent.Children.Clear();
            StackPanel mainCurrentLinePanel = CreateNewLinePanel();
            int currentCharCount = 0;
            int i = 0;

            while (i < text.Length)
            {
                if (text[i] == '\n' || currentCharCount >= MaxCharsPerLine)
                {
                    StackPanelContent.Children.Add(mainCurrentLinePanel);
                    mainCurrentLinePanel = CreateNewLinePanel();
                    currentCharCount = 0;
                    if (text[i] == '\n')
                    {
                        i++;
                        continue;
                    }
                }

                // 跳过空格符
                if (char.IsWhiteSpace(text[i]))
                {
                    // 为空格符创建一个没有拼音的面板
                    StackPanel sp = CreateCharacterPanel(text[i], string.Empty);

                    // 检查是否有误加的拼音，如果有则清除
                    var pinyinBlock = sp.Children[0] as TextBlock;
                    if (!string.IsNullOrEmpty(pinyinBlock.Text))
                    {
                        pinyinBlock.Text = string.Empty;
                    }

                    mainCurrentLinePanel.Children.Add(sp);
                    currentCharCount++;
                    i++;
                    continue;
                }

                // 优先检查轻声词语库
                string wordToCheck = GetLightToneWord(text, i);
                if (lightToneDict.ContainsKey(wordToCheck))
                {
                    i = ProcessWord(wordToCheck, lightToneDict[wordToCheck], ref currentCharCount, ref mainCurrentLinePanel, i);
                }
                // 检查多音字词语库
                else
                {
                    wordToCheck = GetMultiPronunciationWord(text, i);
                    if (multiPronunciationDict.ContainsKey(wordToCheck))
                    {
                        i = ProcessWord(wordToCheck, multiPronunciationDict[wordToCheck], ref currentCharCount, ref mainCurrentLinePanel, i);
                    }
                    else
                    {
                        // 处理其他拼音逻辑
                        char currentChar = text[i];
                        string pinyin = GetCorrectedPinyin(text, i, i == text.Length - 1);

                        if (correctedPinyinDict.ContainsKey(i))
                        {
                            pinyin = correctedPinyinDict[i];
                        }

                        if (currentCharCount >= MaxCharsPerLine)
                        {
                            StackPanelContent.Children.Add(mainCurrentLinePanel);
                            mainCurrentLinePanel = CreateNewLinePanel();
                            currentCharCount = 0;
                        }

                        StackPanel sp = CreateCharacterPanel(currentChar, pinyin);
                        mainCurrentLinePanel.Children.Add(sp);
                        currentCharCount++;
                    }
                }
                i++;
            }

            if (mainCurrentLinePanel.Children.Count > 0)
            {
                StackPanelContent.Children.Add(mainCurrentLinePanel);
            }

            HighlightErhuaHanzi();
            HighlightReduplicatedWords();

           

            // 应用儿化音标注
            foreach (StackPanel linePanel in StackPanelContent.Children)
            {
                ApplyErhuaPinyin(linePanel);
                ApplyDePinyin(linePanel);
            }

            ApplyFontSettings();
        }
        private HashSet<string> LoadErhuaExceptions(string filePath)
        {
            HashSet<string> exceptionSet = new HashSet<string>();
            if (File.Exists(filePath))
            {
                string[] lines = File.ReadAllLines(filePath);
                foreach (var line in lines)
                {
                    exceptionSet.Add(line.Trim());
                }
            }
            return exceptionSet;
        }
       


        private void ApplyErhuaPinyin(StackPanel linePanel)
        {
            for (int i = 0; i < linePanel.Children.Count - 1; i++)
            {
                StackPanel charPanel = linePanel.Children[i] as StackPanel;
                StackPanel nextCharPanel = linePanel.Children[i + 1] as StackPanel;

                if (charPanel != null && nextCharPanel != null)
                {
                    TextBlock hanziBlock = charPanel.Children[1] as TextBlock;
                    TextBlock nextHanziBlock = nextCharPanel.Children[1] as TextBlock;

                    if (hanziBlock != null && nextHanziBlock != null && nextHanziBlock.Text == "儿")
                    {
                        // 检查“儿”字前面的字符
                        if (erPrefixSet.Contains(hanziBlock.Text))
                        {
                            continue; // 跳过儿化音处理
                        }

                        // 检查“儿”字后面的字符
                        if (i + 2 < linePanel.Children.Count)
                        {
                            StackPanel afterNextCharPanel = linePanel.Children[i + 2] as StackPanel;
                            TextBlock afterNextHanziBlock = afterNextCharPanel?.Children[1] as TextBlock;

                            if (afterNextHanziBlock != null && erSuffixSet.Contains(afterNextHanziBlock.Text))
                            {
                                continue; // 跳过儿化音处理
                            }
                        }

                        TextBlock pinyinBlock = charPanel.Children[0] as TextBlock;

                        // 只在拼音不以"r"结尾时才添加儿化音
                        if (!pinyinBlock.Text.EndsWith("r"))
                        {
                            pinyinBlock.Text += "r";
                        }

                        // 清除"儿"字的拼音
                        TextBlock nextPinyinBlock = nextCharPanel.Children[0] as TextBlock;
                        nextPinyinBlock.Text = string.Empty;
                    }
                }
            }
        }


        private void EnsureErhuaConsistency()
        {
            foreach (var linePanel in StackPanelContent.Children)
            {
                ApplyErhuaPinyin(linePanel as StackPanel);
            }
        }

        private int ProcessWord(string word, string pinyinData, ref int currentCharCount, ref StackPanel mainCurrentLinePanel, int i)
        {
            string[] pinyinArray = pinyinData.Split(' ');
            for (int j = 0; j < word.Length; j++)
            {
                if (currentCharCount >= MaxCharsPerLine)
                {
                    StackPanelContent.Children.Add(mainCurrentLinePanel);
                    mainCurrentLinePanel = CreateNewLinePanel();
                    currentCharCount = 0;
                }

                StackPanel sp = CreateCharacterPanel(word[j], pinyinArray[j]);
                mainCurrentLinePanel.Children.Add(sp);

                int index = GetCharacterIndex(sp);
                if (index >= 0)
                {
                    correctedPinyinDict[index] = pinyinArray[j];
                }
                currentCharCount++;
            }
            return i + word.Length - 1;
        }
        private void ApplyDePinyin(StackPanel linePanel)
        {
            for (int i = 0; i < linePanel.Children.Count - 1; i++)
            {
                StackPanel charPanel = linePanel.Children[i] as StackPanel;
                StackPanel nextCharPanel = linePanel.Children[i + 1] as StackPanel;

                if (charPanel != null && nextCharPanel != null)
                {
                    TextBlock hanziBlock = charPanel.Children[1] as TextBlock;
                    TextBlock nextHanziBlock = nextCharPanel.Children[1] as TextBlock;

                    if (hanziBlock != null && hanziBlock.Text == "得")
                    {
                        // 检查“得”字后面的字符是否是动词
                        if (verbSet.Contains(nextHanziBlock.Text))
                        {
                            TextBlock pinyinBlock = charPanel.Children[0] as TextBlock;

                            // 将“得”的拼音设置为“děi”
                            pinyinBlock.Text = "děi";
                        }
                    }
                }
            }
        }

        private string GetCorrectedPinyin(string text, int index, bool isLastChar)
        {
            char currentChar = text[index];
            bool fixDePronunciation = false;

            // 处理特殊的汉字拼音
            if (currentChar == '哇')
            {
                if ((index > 0 && index < text.Length - 1 && text[index - 1] == text[index + 1]) ||
                    (index < text.Length - 1 && char.IsPunctuation(text[index + 1])))
                {
                    return "wa";
                }
            }
            else if (currentChar == '啊')
            {
                if ((index > 0 && index < text.Length - 1 && text[index - 1] == text[index + 1]) ||
                    (index < text.Length - 1 && char.IsPunctuation(text[index + 1])))
                {
                    return "a";
                }
            }
            else if (currentChar == '呀')
            {
                if ((index > 0 && index < text.Length - 1 && text[index - 1] == text[index + 1]) ||
                    (index < text.Length - 1 && char.IsPunctuation(text[index + 1])))
                {
                    return "ya";
                }
            }
            else if (currentChar == '哼')
            {
                if (index < text.Length - 1 && char.IsPunctuation(text[index + 1]))
                {
                    return "hng";
                }
            }
            else if (currentChar == '不')
            {
                if ((index > 0 && index < text.Length - 1 && text[index - 1] == text[index + 1]) ||
                    (index < text.Length - 1 && char.IsPunctuation(text[index + 1])))
                {
                    return "bu";
                }
                else if (index < text.Length - 1)
                {
                    char nextChar = text[index + 1];
                    string nextCharPinyin = hanziPinyinDict.ContainsKey(nextChar.ToString()) ? hanziPinyinDict[nextChar.ToString()][0] : string.Empty;

                    if (nextCharPinyin.IndexOfAny(new char[] { 'à', 'ò', 'è', 'ì', 'ù', 'ǜ' }) >= 0)
                    {
                        return "bú";
                    }
                }
            }
            else if (currentChar == '地')
            {
                if (index < text.Length - 1)
                {
                    for (int i = index + 1; i < text.Length; i++)
                    {
                        string nextWord = GetNextWord(text, i);

                        if (verbSet != null && verbSet.Contains(nextWord))
                        {
                            fixDePronunciation = true;
                            break;
                        }
                        else if (hanziPinyinDict.ContainsKey(text[i].ToString()))
                        {
                            continue;
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }

            // 如果标志被触发，将拼音固定为"de"
            if (fixDePronunciation)
            {
                return "de";
            }

            else if (currentChar == '得')
            {
                // 检查“得”字前面是否匹配前缀
                if (index > 0)
                {
                    string prevWord = text.Substring(Math.Max(0, index - 4), Math.Min(4, index)); // 获取可能的前缀
                    if (dePrefixSet.Any(prefix => prevWord.EndsWith(prefix)))
                    {
                        return "dé";
                    }
                }

                // 检查“得”字后面是否匹配后缀
                string followingText = text.Substring(index + 1); // 确保followingText在需要时已经定义
                if (index < text.Length - 1)
                {
                    if (deSuffixSet.Any(suffix => followingText.StartsWith(suffix)))
                    {
                        return "dé";
                    }
                }

                // 默认逻辑
                bool hasVerbAfter = verbSet != null && verbSet.Any(verb => followingText.StartsWith(verb));
                if (hasVerbAfter)
                {
                    return "děi";
                }
                else
                {
                    return "de";
                }
            }


            else if (currentChar == '更')
            {
                if (index < text.Length - 1 && adjectiveSet != null)
                {
                    for (int i = index + 1; i < text.Length; i++)
                    {
                        string nextWord = GetNextWord(text, i);
                        if (adjectiveSet != null && adjectiveSet.Contains(nextWord))
                        {
                            return "gèng";
                        }
                        else if (hanziPinyinDict.ContainsKey(text[i].ToString()))
                        {
                            continue;
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }
            else if (currentChar == '一')
            {
                if (index < text.Length - 1)
                {
                    string nextChar = text[index + 1].ToString();
                    string nextCharPinyin = hanziPinyinDict.ContainsKey(nextChar) ? hanziPinyinDict[nextChar][0] : string.Empty;

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

        private string GetNextWord(string text, int startIndex)
        {
            // 获取词库中最长的词语长度
            int maxLength = Math.Min(Math.Max(
                adjectiveSet?.Max(w => w.Length) ?? 0,
                Math.Max(
                    verbSet?.Max(w => w.Length) ?? 0,
                    multiPronunciationDict.Keys.Max(k => k.Length)
                )
            ), text.Length - startIndex); // 确保最大长度不超过文本剩余长度

            for (int length = maxLength; length > 0; length--)
            {
                if (startIndex + length <= text.Length)
                {
                    string potentialWord = text.Substring(startIndex, length);

                    // 忽略空格和标点符号
                    if (string.IsNullOrWhiteSpace(potentialWord) || char.IsPunctuation(potentialWord[0]))
                    {
                        continue;
                    }

                    // 优先检查内置形容词库
                    if (adjectiveSet != null && adjectiveSet.Contains(potentialWord))
                    {
                        return potentialWord;
                    }

                    // 然后检查内置动词库
                    if (verbSet != null && verbSet.Contains(potentialWord))
                    {
                        return potentialWord;
                    }

                    // 最后检查多音字词库
                    if (multiPronunciationDict.ContainsKey(potentialWord))
                    {
                        return potentialWord;
                    }
                }
            }

            // 如果没有匹配，返回单个字符
            return text[startIndex].ToString();
        }


        private string GetMultiPronunciationWord(string text, int startIndex)
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
        private string GetLightToneWord(string text, int startIndex)
        {
            int maxLength = lightToneDict.Keys.Max(k => k.Length);
            for (int length = maxLength; length > 0; length--)
            {
                if (startIndex + length <= text.Length)
                {
                    string substring = text.Substring(startIndex, length);
                    if (lightToneDict.ContainsKey(substring))
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
                        var hanziBlock = charPanel?.Children[1] as TextBlock;

                        if (hanziBlock != null && hanziBlock.Text == "儿")
                        {
                            var prevCharPanel = sp.Children[i - 1] as StackPanel;
                            var prevHanziBlock = prevCharPanel?.Children[1] as TextBlock;

                            if (prevHanziBlock != null)
                            {
                                prevHanziBlock.Background = System.Windows.Media.Brushes.LightGreen;
                                prevHanziBlock.MouseLeftButtonUp += HanziBlock_MouseLeftButtonUp;
                            }
                        }
                    }
                }
            }
        }

        private void HighlightReduplicatedWords()
        {
            foreach (var child in StackPanelContent.Children)
            {
                if (child is StackPanel sp)
                {
                    for (int i = 1; i < sp.Children.Count; i++)
                    {
                        var charPanel = sp.Children[i] as StackPanel;
                        var hanziBlock = charPanel?.Children[1] as TextBlock;

                        if (hanziBlock != null && IsChineseCharacter(hanziBlock.Text))
                        {
                            var prevCharPanel = sp.Children[i - 1] as StackPanel;
                            var prevHanziBlock = prevCharPanel?.Children[1] as TextBlock;

                            if (prevHanziBlock != null && prevHanziBlock.Text == hanziBlock.Text && IsChineseCharacter(prevHanziBlock.Text))
                            {
                                hanziBlock.Background = System.Windows.Media.Brushes.LightPink;
                                hanziBlock.MouseLeftButtonUp += HanziBlock_MouseLeftButtonUp; // 添加选项菜单
                            }
                        }
                    }
                }
            }
        }

        private bool IsChineseCharacter(string text)
        {
            return !string.IsNullOrEmpty(text) && text[0] >= 0x4E00 && text[0] <= 0x9FFF;
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

                    // 如果当前是偶数行（汉字行），并且该单元格是中文标点符号
                    if (IsChinesePunctuation(hanziLines[i][j]) && string.IsNullOrEmpty(content[i * 2, j]))
                    {
                        content[i * 2, j] = "\u3000"; // 在对应的奇数行（拼音行）中使用全角空格符占位
                    }

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

        private readonly HashSet<char> chinesePunctuation = new HashSet<char>
        {
            '。', '，', '、', '；', '：', '！', '？', '“', '”', '‘', '’', '（', '）', '【', '】', '《', '》', '—', '…', '『', '』', '「', '」'
        };

        private bool IsChinesePunctuation(string text)
        {
            return text.Length == 1 && chinesePunctuation.Contains(text[0]);
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

        private void ExportToTextBox()
        {
            var pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Slide activeSlide = pptApp.ActiveWindow.View.Slide as PowerPoint.Slide;

            // 获取拼音和汉字文本
            string hanziText = string.Empty;
            List<PinyinData> pinyinDataList = new List<PinyinData>();

            foreach (var child in StackPanelContent.Children)
            {
                if (child is StackPanel sp)
                {
                    foreach (var element in sp.Children)
                    {
                        if (element is StackPanel charPanel)
                        {
                            var pinyinBlock = charPanel.Children[0] as TextBlock;
                            var hanziBlock = charPanel.Children[1] as TextBlock;

                            hanziText += hanziBlock.Text;
                            pinyinDataList.Add(new PinyinData
                            {
                                Pinyin = pinyinBlock.Text,
                                CharIndex = hanziText.Length, // 记录字符的索引位置
                                HanziTextBoxName = "HanziTextBox" // 暂时设置为默认值
                            });
                        }
                    }
                }
            }

            // 创建一个大的文本框来容纳汉字和标点符号
            float fontSize = 28; // 汉字的字体大小
            float pinyinFontSize = fontSize * 0.5f; // 拼音的字体大小为汉字的一半
                                                    // 计算文本框的宽度
            float averageCharWidth = fontSize * 0.6f; // 假设平均字符宽度为字体大小的0.6倍
            float calculatedWidth = hanziText.Length * averageCharWidth;

            // 限制宽度最大值为幻灯片宽度，最小值为一定宽度
            float slideWidth = pptApp.ActivePresentation.PageSetup.SlideWidth;
            float textBoxWidth = Math.Min(Math.Max(calculatedWidth, 100), slideWidth - 20); // 最小宽度100，最大宽度为幻灯片宽度减去20

            PowerPoint.Shape hanziTextBox = activeSlide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal, 100, 100, textBoxWidth, 100);
            hanziTextBox.TextFrame2.TextRange.Text = hanziText;
            hanziTextBox.TextFrame2.TextRange.Font.Size = fontSize;
            hanziTextBox.TextFrame2.TextRange.ParagraphFormat.SpaceWithin = 1.5f; // 行间距
            hanziTextBox.TextFrame2.TextRange.Font.Spacing = 6; // 字间距设为加宽
            hanziTextBox.TextFrame2.WordWrap = MsoTriState.msoTrue;
            hanziTextBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            // 设置文本对齐方式
            TextAlignment currentAlignment = RichTextBoxLeft.Document.Blocks.OfType<Paragraph>().FirstOrDefault()?.TextAlignment ?? TextAlignment.Left;
            switch (currentAlignment)
            {
                case TextAlignment.Left:
                    hanziTextBox.TextFrame2.TextRange.ParagraphFormat.Alignment = (MsoParagraphAlignment)PowerPoint.PpParagraphAlignment.ppAlignLeft;
                    break;
                case TextAlignment.Center:
                    hanziTextBox.TextFrame2.TextRange.ParagraphFormat.Alignment = (MsoParagraphAlignment)PowerPoint.PpParagraphAlignment.ppAlignCenter;
                    break;
                case TextAlignment.Right:
                    hanziTextBox.TextFrame2.TextRange.ParagraphFormat.Alignment = (MsoParagraphAlignment)PowerPoint.PpParagraphAlignment.ppAlignRight;
                    break;
                case TextAlignment.Justify:
                    hanziTextBox.TextFrame2.TextRange.ParagraphFormat.Alignment = (MsoParagraphAlignment)PowerPoint.PpParagraphAlignment.ppAlignJustify;
                    break;
            }


            // 更新每个拼音数据的文本框名称
            for (int i = 0; i < pinyinDataList.Count; i++)
            {
                pinyinDataList[i].HanziTextBoxName = hanziTextBox.Name;
            }

            // 获取每个字符的位置和大小
            for (int i = 1; i <= hanziTextBox.TextFrame.TextRange.Text.Length; i++)
            {
                Microsoft.Office.Interop.PowerPoint.TextRange charRange = hanziTextBox.TextFrame.TextRange.Characters(i, 1);
                char currentChar = hanziTextBox.TextFrame.TextRange.Text[i - 1];

                // 跳过标点符号和空格
                if (char.IsPunctuation(currentChar) || char.IsWhiteSpace(currentChar))
                {
                    continue;
                }

                // 获取选中文字的位置和大小并进行调整
                float charTop = charRange.BoundTop - fontSize / 2; // 初始位置调整

                // 计算行间距调整
                float extraLineSpacing = Math.Max(0, charRange.BoundHeight - fontSize);
                const float downwardShiftRatio = 1.2f;  // 基于额外行间距的向下调整比例
                charTop += extraLineSpacing * downwardShiftRatio;

                // 动态计算偏移量
                float lineSpacingMultiplier = charRange.ParagraphFormat.SpaceWithin; // 获取行间距倍数
                float baseOffset = 8; // 基础偏移量（对应1.5倍行间距）
                float additionalOffset = (lineSpacingMultiplier - 1.5f) * 12; // 每增加0.25行间距，增加3个单位的偏移量
                float adjustedPinyinTop = charTop - pinyinFontSize - (baseOffset + additionalOffset);

                // 计算动态左移偏移量，基于28字号时左移3个单位
                float charLeft = charRange.BoundLeft - ((charRange.Font.Size / 28) * 3);
                float charWidth = charRange.BoundWidth;

                // 创建拼音文本框并设置位置
                PowerPoint.Shape pinyinTextBox = activeSlide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal, charLeft, adjustedPinyinTop, charWidth, pinyinFontSize);
                pinyinTextBox.TextFrame.TextRange.Text = pinyinDataList[i - 1].Pinyin;
                pinyinTextBox.TextFrame.TextRange.Font.Size = pinyinFontSize;
                pinyinTextBox.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                pinyinTextBox.TextFrame.WordWrap = MsoTriState.msoFalse;

                // 添加 Tags
                pinyinTextBox.Tags.Add("PinYin", "True");
                pinyinTextBox.Tags.Add("ParentCharIndex", (i).ToString());
                pinyinTextBox.Tags.Add("ParentTextBoxName", hanziTextBox.Name);
            }

            // 检查并移除汉字文本框末尾的空白行
            if (hanziText.EndsWith("\r\n") || hanziText.EndsWith("\n") || hanziText.EndsWith("\r"))
            {
                // 移除末尾的空白行
                hanziTextBox.TextFrame2.TextRange.Text = hanziText.TrimEnd('\r', '\n');
            }
            // 在所有拼音文本框创建完成后弹出一次成功提示
            MessageBox.Show("已成功导出注音文本。", "导出成功");

        }

        private class PinyinData
        {
            public string HanziTextBoxName { get; set; }
            public int CharIndex { get; set; }
            public string Pinyin { get; set; }
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

                TextPointer caretPosition = RichTextBoxLeft.CaretPosition;

                ResetRichTextBoxFormatting();

                RichTextBoxLeft.CaretPosition = caretPosition;

                string text = GetPlainTextFromRichTextBox();
                string plainText = text.Replace(Environment.NewLine, "").Replace(" ", ""); // 去除空格和换行符

                int charCount = plainText.Length;
                UpdateCharacterCount(charCount);

                correctedPinyinDict.Clear();
                UpdateStackPanelContentWithCorrection(text);

                // 确保儿化音标注一致性
                EnsureErhuaConsistency();

                SyncAlignmentWithPinyin();

                isTextChangedEventHandlerActive = true;
            }
        }

        private void ResetRichTextBoxFormatting()
        {
            var textRange = new System.Windows.Documents.TextRange(RichTextBoxLeft.Document.ContentStart, RichTextBoxLeft.Document.ContentEnd);
            textRange.ClearAllProperties();

            // 设置默认字体和字号
            textRange.ApplyPropertyValue(TextElement.FontFamilyProperty, new System.Windows.Media.FontFamily("Microsoft YaHei UI"));
            textRange.ApplyPropertyValue(TextElement.FontSizeProperty, 20.0); // 默认字号
            // 设置行间距
            foreach (var block in RichTextBoxLeft.Document.Blocks.OfType<Paragraph>())
            {
                block.LineHeight = 10; // 设置行间距
            }
        }


        private void UpdateCharacterCount(int charCount)
        {
            TextBlockCharCount.Text = $"当前字符数: {charCount}";

            if (charCount > MaxCharCountThreshold)
            {
                BtnCharCountWarning.Visibility = Visibility.Visible;
            }
            else
            {
                BtnCharCountWarning.Visibility = Visibility.Collapsed;
            }
        }

        private string GetPlainTextFromRichTextBox()
        {
            // 获取 RichTextBox 中的纯文本
            return new System.Windows.Documents.TextRange(RichTextBoxLeft.Document.ContentStart, RichTextBoxLeft.Document.ContentEnd).Text;
        }

        private void UpdateStackPanelContent(string text)
        {
            StackPanelContent.Children.Clear();
            StackPanel mainCurrentLinePanel = CreateNewLinePanel();
            int currentCharCount = 0;

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];

                // 正常处理字符
                StackPanel charPanel = CreateCharacterPanel(c);
                string pinyin = GetCorrectedPinyin(text, i, i == text.Length - 1);
                (charPanel.Children[0] as TextBlock).Text = pinyin;

                mainCurrentLinePanel.Children.Add(charPanel);
                currentCharCount++;

                if (currentCharCount >= MaxCharsPerLine)
                {
                    StackPanelContent.Children.Add(mainCurrentLinePanel);
                    mainCurrentLinePanel = CreateNewLinePanel();
                    currentCharCount = 0;
                }
            }

            if (mainCurrentLinePanel.Children.Count > 0)
            {
                StackPanelContent.Children.Add(mainCurrentLinePanel);
            }

            SyncAlignmentWithPinyin();
            HighlightErhuaHanzi();
            HighlightReduplicatedWords();
            ApplyFontSettings();
        }

        private void ApplyFontSettings()
        {
            string pinyinFont = Properties.Settings.Default.PinyinFont;
            double pinyinFontSize = Properties.Settings.Default.PinyinFontSize;
            string hanziFont = Properties.Settings.Default.HanziFont;
            double hanziFontSize = Properties.Settings.Default.HanziFontSize;

            foreach (var child in StackPanelContent.Children)
            {
                if (child is StackPanel sp)
                {
                    foreach (var element in sp.Children)
                    {
                        if (element is StackPanel charPanel)
                        {
                            var pinyinBlock = charPanel.Children[0] as TextBlock;
                            var hanziBlock = charPanel.Children[1] as TextBlock;

                            if (pinyinBlock != null)
                            {
                                pinyinBlock.FontFamily = new System.Windows.Media.FontFamily(pinyinFont);  // System.Windows.Media.FontFamily
                                pinyinBlock.FontSize = pinyinFontSize;
                            }

                            if (hanziBlock != null)
                            {
                                hanziBlock.FontFamily = new System.Windows.Media.FontFamily(hanziFont);  // System.Windows.Media.FontFamily
                                hanziBlock.FontSize = hanziFontSize;
                            }
                        }
                    }
                }
            }
        }

        private void BtnChangeFontSettings_Click(object sender, RoutedEventArgs e)
        {
            var fontSettingsWindow = new FontSettingsWindow();
            if (fontSettingsWindow.ShowDialog() == true)
            {
                ApplyFontSettings(); // 用户确认设置后应用新的字体设置
            }
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
            if (charPanel.Parent is StackPanel parentPanel)
            {
                int parentIndex = StackPanelContent.Children.IndexOf(parentPanel);
                if (parentIndex >= 0)
                {
                    int charIndex = parentPanel.Children.IndexOf(charPanel);
                    if (charIndex >= 0)
                    {
                        return parentIndex * MaxCharsPerLine + charIndex;
                    }
                    else
                    {
                        return -1;
                    }
                }
                else
                {
                    return -1;
                }
            }
            else
            {
                return -1;
            }
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
            }
            else
            {
                ApplyAlignmentToParagraphs(alignment);
            }

            foreach (StackPanel linePanel in StackPanelContent.Children.OfType<StackPanel>())
            {
                linePanel.HorizontalAlignment = ConvertToHorizontalAlignment(alignment);
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
            var leftParagraphs = RichTextBoxLeft.Document.Blocks.OfType<Paragraph>().ToList();
            var rightPanels = StackPanelContent.Children.OfType<StackPanel>().ToList();

            for (int i = 0; i < leftParagraphs.Count && i < rightPanels.Count; i++)
            {
                var alignment = leftParagraphs[i].TextAlignment;
                rightPanels[i].HorizontalAlignment = ConvertToHorizontalAlignment(alignment);
            }
        }

        //左对齐
        private void BtnAlignLeft_Click(object sender, RoutedEventArgs e)
        {
            AlignText(TextAlignment.Left);
        }
        //居中对齐
        private void BtnAlignCenter_Click(object sender, RoutedEventArgs e)
        {
            AlignText(TextAlignment.Center);
        }

        //两端对齐
        private void BtnAlignJustify_Click(object sender, RoutedEventArgs e)
        {
            AlignText(TextAlignment.Justify);
        }
    }
}
