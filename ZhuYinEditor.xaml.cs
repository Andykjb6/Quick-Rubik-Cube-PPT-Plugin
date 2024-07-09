using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;
using OfficeOpenXml;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace 课件帮PPT助手
{
    public partial class ZhuYinEditor : Window
    {
        private const int MaxCharsPerLine = 30;
        private Dictionary<string, List<string>> hanziPinyinDict;

        public ZhuYinEditor()
        {
            InitializeComponent();
            LoadHanziPinyinDict();
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
            // 检测多音字逻辑
        }

        private void BtnAddPronunciations_Click(object sender, RoutedEventArgs e)
        {
            UpdateStackPanelContent(InputTextBox.Text);
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
                InputTextBox.Text = text;
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            var pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Slide activeSlide = pptApp.ActiveWindow.View.Slide as PowerPoint.Slide;

            PowerPoint.Shape textBox = activeSlide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                100, 100, 500, 400
            );

            foreach (var child in StackPanelContent.Children)
            {
                if (child is StackPanel sp)
                {
                    foreach (var element in sp.Children)
                    {
                        if (element is TextBlock tb)
                        {
                            textBox.TextFrame.TextRange.Text += tb.Text;
                        }
                    }
                }
            }
        }

        private void InputTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            var textBox = sender as TextBox;
            string text = textBox.Text;
            string wrappedText = WrapText(text, MaxCharsPerLine);
            textBox.TextChanged -= InputTextBox_TextChanged; // 防止无限递归
            textBox.Text = wrappedText;
            textBox.CaretIndex = wrappedText.Length; // 设置光标位置到文本末尾
            textBox.TextChanged += InputTextBox_TextChanged;

            UpdateStackPanelContent(textBox.Text.Replace("\r\n", ""));
        }

        private string WrapText(string text, int maxCharsPerLine)
        {
            string result = "";
            int currentLength = 0;
            foreach (char c in text)
            {
                if (currentLength >= maxCharsPerLine && c != '\n')
                {
                    result += "\n";
                    currentLength = 0;
                }
                result += c;
                currentLength++;
            }
            return result;
        }

        private void UpdateStackPanelContent(string text)
        {
            StackPanelContent.Children.Clear();
            foreach (char c in text)
            {
                StackPanel sp = new StackPanel
                {
                    Orientation = Orientation.Vertical,
                    Margin = new Thickness(5, 0, 5, 0),
                    HorizontalAlignment = HorizontalAlignment.Center
                };

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

                StackPanelContent.Children.Add(sp);
            }
        }
    }
}
