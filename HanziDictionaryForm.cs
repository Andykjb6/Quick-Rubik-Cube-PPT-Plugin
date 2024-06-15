using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace 课件帮PPT助手
{
    public partial class HanziDictionaryForm : Form
    {
        private Dictionary<string, HanziInfo> hanziDictionary;
        private HanziInfo currentHanziInfo;

        public HanziDictionaryForm()
        {
            InitializeComponent();
            LoadHanziDictionary();
            this.TopMost = true;
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
        }

        private void HanziDictionaryForm_Load(object sender, EventArgs e)
        {
            searchButton.Click += SearchButton_Click;
            searchTextBox.Text = "字";
            SearchButton_Click(this, EventArgs.Empty);

            pinyinLabel.Click += (s, ev) => Clipboard.SetText(pinyinLabel.Text);
            radicalLabel.Click += (s, ev) => Clipboard.SetText(radicalLabel.Text);
            strokesLabel.Click += (s, ev) => Clipboard.SetText(strokesLabel.Text);
            structureLabel.Click += (s, ev) => Clipboard.SetText(structureLabel.Text);
        }

        private void LoadHanziDictionary()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            hanziDictionary = new Dictionary<string, HanziInfo>();
            string filePath = ExtractEmbeddedResource("课件帮PPT助手.汉字字典.汉字字典.xlsx");

            if (!File.Exists(filePath))
            {
                MessageBox.Show($"未找到文件：{filePath}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                if (package.Workbook.Worksheets.Count == 0)
                {
                    MessageBox.Show("没有查到相关信息。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var worksheet = package.Workbook.Worksheets[0];

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    string hanzi = worksheet.Cells[row, 1].Text;
                    string pinyin = worksheet.Cells[row, 2].Text;
                    string radical = worksheet.Cells[row, 3].Text;
                    string structure = worksheet.Cells[row, 4].Text;
                    int strokes = int.Parse(worksheet.Cells[row, 5].Text);
                    string relatedWords = worksheet.Cells[row, 6].Text;

                    hanziDictionary[hanzi] = new HanziInfo
                    {
                        Pinyin = pinyin,
                        Radical = radical,
                        Structure = structure,
                        Strokes = strokes,
                        RelatedWords = relatedWords.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                    .Select(word => word.Trim())
                                                    .Where(word => !string.IsNullOrEmpty(word)) // 移除空白词语
                                                    .ToArray()
                    };
                }
            }
        }

        private string ExtractEmbeddedResource(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceStream = assembly.GetManifestResourceStream(resourceName);

            if (resourceStream == null)
                throw new Exception($"Embedded resource {resourceName} not found.");

            string tempFilePath = Path.Combine(Path.GetTempPath(), resourceName);

            using (var fileStream = new FileStream(tempFilePath, FileMode.Create))
            {
                resourceStream.CopyTo(fileStream);
            }

            return tempFilePath;
        }

        private void SearchButton_Click(object sender, EventArgs e)
        {
            string hanzi = searchTextBox.Text.Trim();
            if (hanziDictionary.ContainsKey(hanzi))
            {
                currentHanziInfo = hanziDictionary[hanzi];
                hanziLabel.Text = hanzi;
                pinyinLabel.Text = $"拼音: {currentHanziInfo.Pinyin}";
                radicalLabel.Text = $"部首: {currentHanziInfo.Radical}";
                strokesLabel.Text = $"笔画: {currentHanziInfo.Strokes}";
                structureLabel.Text = $"结构: {currentHanziInfo.Structure}";

                wordsPanel.Controls.Clear();
                int wordCount = 0;
                bool hasLongWord = currentHanziInfo.RelatedWords.Any(word => word.Length > 2);
                int wordsPerRow = hasLongWord ? 4 : 5;

                foreach (var word in currentHanziInfo.RelatedWords)
                {
                    var button = new Button
                    {
                        Text = word,
                        AutoSize = true,
                        Margin = new Padding(5),
                        AutoSizeMode = AutoSizeMode.GrowAndShrink
                    };

                    button.Click += (s, ev) => Clipboard.SetText(word);

                    wordsPanel.Controls.Add(button);
                    wordCount++;

                    if (wordCount % wordsPerRow == 0)
                    {
                        wordsPanel.SetFlowBreak(button, true);
                    }
                }
            }
            else
            {
                MessageBox.Show("未找到该汉字的信息。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public class HanziInfo
        {
            public string Pinyin { get; set; }
            public string Radical { get; set; }
            public string Structure { get; set; }
            public int Strokes { get; set; }
            public string[] RelatedWords { get; set; }
        }

        private void 导出_Click(object sender, EventArgs e)
        {
            if (currentHanziInfo == null)
            {
                MessageBox.Show("请先搜索一个汉字。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                var powerpointApp = new PowerPoint.Application();
                var presentation = powerpointApp.ActivePresentation;
                var slide = presentation.Slides[presentation.Slides.Count];

                var text = $"汉字: {hanziLabel.Text}\n" +
                           $"拼音: {currentHanziInfo.Pinyin}\n" +
                           $"部首: {currentHanziInfo.Radical}\n" +
                           $"笔画: {currentHanziInfo.Strokes}\n" +
                           $"结构: {currentHanziInfo.Structure}\n" +
                           $"相关词语: {string.Join(", ", currentHanziInfo.RelatedWords)}";

                var textBox = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 100, 100, 500, 200);
                textBox.TextFrame.TextRange.Text = text;

            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
