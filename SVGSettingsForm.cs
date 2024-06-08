using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Windows.Forms;
using HtmlAgilityPack;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace 课件帮PPT助手
{
    public partial class SVGSettingsForm : Form
    {
        private string inputChar;
        private List<List<KeyValuePair<string, string>>> svgMatrix;
        private List<string> headers;

        public SVGSettingsForm(string inputChar, List<List<KeyValuePair<string, string>>> svgMatrix, List<string> headers)
        {
            InitializeComponent();
            this.inputChar = inputChar;
            this.svgMatrix = svgMatrix;
            this.headers = headers;
            PopulateListBox();
        }

        private void PopulateListBox()
        {
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
        }

        private void SelectButton_Click(object sender, EventArgs e)
        {
            List<KeyValuePair<string, string>> selectedSVGs = new List<KeyValuePair<string, string>>();
            foreach (var selectedItem in svgListBox.SelectedItems)
            {
                var kvp = (KeyValuePair<string, string>)selectedItem;
                selectedSVGs.Add(kvp);
            }
            this.Close();
            InsertSVGsIntoPresentation(selectedSVGs);
        }

        private void InsertSVGsIntoPresentation(List<KeyValuePair<string, string>> selectedSVGs)
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
                        using (var client = new WebClient())
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
            }
            catch (Exception ex)
            {
                MessageBox.Show("出现错误：" + ex.Message);
            }
        }
    }
}
