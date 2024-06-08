using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using HtmlAgilityPack;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace 课件帮PPT助手
{
    public partial class SvgSelectionForm : Form
    {
        private string inputChar;
        private Dictionary<string, string> svgDictionary;

        public SvgSelectionForm(HtmlNodeCollection svgNodes, string inputChar)
        {
            InitializeComponent();
            this.inputChar = inputChar;
            svgDictionary = new Dictionary<string, string>();
            PopulateSvgListBox(svgNodes);
        }

        private void PopulateSvgListBox(HtmlNodeCollection svgNodes)
        {
            int count = 1;
            foreach (var svgNode in svgNodes)
            {
                string scientificName = $"{inputChar}-第{count}笔";
                svgListBox.Items.Add(scientificName);
                svgDictionary[scientificName] = svgNode.OuterHtml;
                count++;
            }
        }

        private void SelectButton_Click(object sender, EventArgs e)
        {
            List<string> selectedSVGs = new List<string>();
            foreach (var selectedItem in svgListBox.SelectedItems)
            {
                string key = selectedItem.ToString();
                if (svgDictionary.ContainsKey(key))
                {
                    selectedSVGs.Add(svgDictionary[key]);
                }
            }
            this.Close();
            InsertSVGsIntoPresentation(selectedSVGs, inputChar);
        }

        private void InsertSVGsIntoPresentation(List<string> svgContents, string inputChar)
        {
            try
            {
                PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
                PowerPoint.Slide slide = pptApp.ActiveWindow.View.Slide;

                int count = 1;
                int xOffset = 100;  // 初始 x 坐标
                int yOffset = 100;  // 初始 y 坐标
                int xSpacing = 2; // 插入PPT的每个 SVG 之间的水平间隔

                foreach (var svgContent in svgContents)
                {
                    string pattern = @"width:(\s*\d+)px;\s*height:(\s*\d+)px;";
                    Match match = Regex.Match(svgContent, pattern);

                    string width = "100";  // 默认宽度
                    string height = "100"; // 默认高度

                    if (match.Success)
                    {
                        width = match.Groups[1].Value.Trim();
                        height = match.Groups[2].Value.Trim();
                    }

                    string updatedSvg = InsertSvgAttributesWithDimensions(svgContent, width, height);
                    string tempSvgPath = Path.Combine(Path.GetTempPath(), $"{inputChar}-第{count}笔.svg");
                    File.WriteAllText(tempSvgPath, updatedSvg);
                    slide.Shapes.AddPicture(tempSvgPath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoCTrue, xOffset, yOffset);

                    File.Delete(tempSvgPath);

                    // 更新 xOffset，以便下一个 SVG 水平排列
                    xOffset += int.Parse(width) + xSpacing;

                    count++;
                }

                MessageBox.Show("成功插入SVG到PPT中！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("出现错误：" + ex.Message);
            }
        }

        private string InsertSvgAttributesWithDimensions(string svg, string width, string height)
        {
            int index = svg.IndexOf("<svg ");
            if (index != -1)
            {
                int spaceIndex = svg.IndexOf(' ', index);
                if (spaceIndex != -1)
                {
                    string attributes = $"width='{width}' height='{height}' ";
                    return svg.Insert(spaceIndex + 1, attributes);
                }
            }
            return svg;
        }
    }
}
