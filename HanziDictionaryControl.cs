using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace 课件帮PPT助手
{
    public partial class HanziDictionaryControl : UserControl
    {
        private Dictionary<string, HanziInfo> hanziDictionary;

        public HanziDictionaryControl()
        {
            InitializeComponent();
            LoadHanziDictionary();
        }

        private void HanziDictionaryControl_Load(object sender, EventArgs e)
        {
            searchButton.Click += SearchButton_Click;
        }

        private void LoadHanziDictionary()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // 设置许可证上下文

            hanziDictionary = new Dictionary<string, HanziInfo>();
            string filePath = @"C:\Users\Andy\source\repos\课件帮PPT助手\汉字字典\汉字字典.xlsx";

            if (!File.Exists(filePath))
            {
                MessageBox.Show($"未找到文件：{filePath}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                if (package.Workbook.Worksheets.Count == 0)
                {
                    MessageBox.Show("Excel文件中没有工作表。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                                                    .ToArray()
                    };
                }
            }
        }

        private void SearchButton_Click(object sender, EventArgs e)
        {
            string hanzi = searchTextBox.Text.Trim();
            if (hanziDictionary.ContainsKey(hanzi))
            {
                var info = hanziDictionary[hanzi];
                hanziLabel.Text = hanzi;
                pinyinLabel.Text = $"拼音: {info.Pinyin}";
                radicalLabel.Text = $"部首: {info.Radical}";
                strokesLabel.Text = $"笔画: {info.Strokes}";
                structureLabel.Text = $"结构: {info.Structure}";

                wordsPanel.Controls.Clear();
                for (int i = 0; i < info.RelatedWords.Length; i++)
                {
                    var word = info.RelatedWords[i];
                    var button = new Button
                    {
                        Text = word,
                        AutoSize = true,
                        Margin = new Padding(5)
                    };

                    wordsPanel.Controls.Add(button);

                    if ((i + 1) % 5 == 0)
                    {
                        // Set the FlowBreak property to true for every 5th word
                        wordsPanel.SetFlowBreak(button, true);
                    }
                }
            }
            else
            {
                MessageBox.Show("未找到该汉字的信息。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void hanziLabel_Click(object sender, EventArgs e)
        {

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
}
