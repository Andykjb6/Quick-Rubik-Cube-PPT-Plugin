using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace 课件帮PPT助手
{
    public partial class FontDetectionForm : Form
    {
        public FontDetectionForm(List<string> usedFonts, List<string> unusedFonts, PowerPoint.Presentation presentation)
        {
            InitializeComponent();
            listBoxUsed.Items.AddRange(usedFonts.ToArray());
            listBoxUnused.Items.AddRange(unusedFonts.ToArray());
            this.presentation = presentation;
            this.usedFonts = usedFonts;
        }

        private readonly PowerPoint.Presentation presentation;
        private readonly List<string> usedFonts;

        private void ClearButton_Click(object sender, EventArgs e)
        {
            var unusedFonts = listBoxUnused.Items.Cast<string>().ToList();
            if (unusedFonts.Count == 0 || usedFonts.Count == 0)
            {
                MessageBox.Show("没有未使用的字体或没有已使用的字体进行替换。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string replacementFont = usedFonts[0]; // 使用已使用字体中的第一个进行替换

            foreach (string font in unusedFonts)
            {
                try
                {
                    foreach (PowerPoint.Slide slide in presentation.Slides)
                    {
                        foreach (PowerPoint.Shape shape in slide.Shapes)
                        {
                            ReplaceFontInShape(shape, font, replacementFont);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"替换字体 {font} 过程中出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            listBoxUnused.Items.Clear(); // 清空未使用字体列表
            MessageBox.Show("未使用的字体已清除并替换为已使用的字体。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ReplaceFontInShape(PowerPoint.Shape shape, string targetFont, string replacementFont)
        {
            if (shape.HasTextFrame == MsoTriState.msoTrue)
            {
                PowerPoint.TextRange textRange = shape.TextFrame.TextRange;
                foreach (PowerPoint.TextRange run in textRange.Runs(1, textRange.Text.Length))
                {
                    if (run.Font.Name == targetFont)
                    {
                        run.Font.Name = replacementFont;
                    }
                }
            }

            // 替换没有文本但有字体设置的形状
            if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoFalse)
            {
                var fonts = shape.TextFrame.TextRange.Font;
                if (fonts.Name == targetFont)
                {
                    fonts.Name = replacementFont;
                }
            }

            if (shape.Type == MsoShapeType.msoGroup)
            {
                foreach (PowerPoint.Shape groupedShape in shape.GroupItems)
                {
                    ReplaceFontInShape(groupedShape, targetFont, replacementFont);
                }
            }
        }

        private void ExportFontsButton_Click(object sender, EventArgs e)
        {
            string folderPath;
            if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
            {
                using (FolderBrowserDialog dialog = new FolderBrowserDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        folderPath = dialog.SelectedPath;
                    }
                    else
                    {
                        return; // 用户取消操作
                    }
                }
            }
            else
            {
                string presentationName = presentation.Name;
                folderPath = Path.Combine("C:\\", presentationName + "（所需字体）");

                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }
            }

            foreach (string fontName in usedFonts)
            {
                string fontFilePath = GetFontFilePath(fontName);

                if (!string.IsNullOrEmpty(fontFilePath))
                {
                    try
                    {
                        string destFontPath = Path.Combine(folderPath, Path.GetFileName(fontFilePath));
                        File.Copy(fontFilePath, destFontPath, true);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"字体 {fontName} 复制失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show($"未找到字体文件: {fontName}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            MessageBox.Show("字体导出完成！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            System.Diagnostics.Process.Start("explorer.exe", folderPath);
        }

        private void PackageDocumentButton_Click(object sender, EventArgs e)
        {
            string folderPath;
            if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
            {
                using (FolderBrowserDialog dialog = new FolderBrowserDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        folderPath = dialog.SelectedPath;
                    }
                    else
                    {
                        return; // 用户取消操作
                    }
                }
            }
            else
            {
                string presentationName = Path.GetFileNameWithoutExtension(presentation.FullName);
                folderPath = Path.Combine("C:\\", presentationName);

                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }
            }

            string presentationNameForPath = Path.GetFileNameWithoutExtension(presentation.FullName);
            string presentationPath = Path.Combine(folderPath, presentationNameForPath + ".pptx");
            presentation.SaveCopyAs(presentationPath);

            string fontsFolderPath = Path.Combine(folderPath, "文档所用字体");

            if (!Directory.Exists(fontsFolderPath))
            {
                Directory.CreateDirectory(fontsFolderPath);
            }

            foreach (string fontName in usedFonts)
            {
                string fontFilePath = GetFontFilePath(fontName);

                if (!string.IsNullOrEmpty(fontFilePath))
                {
                    try
                    {
                        string destFontPath = Path.Combine(fontsFolderPath, Path.GetFileName(fontFilePath));
                        File.Copy(fontFilePath, destFontPath, true);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"字体 {fontName} 复制失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show($"未找到字体文件: {fontName}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            MessageBox.Show("打包完成！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            System.Diagnostics.Process.Start("explorer.exe", folderPath);
        }

        private string GetFontFilePath(string fontName)
        {
            // 处理常见字体名称
            switch (fontName)
            {
                case "宋体":
                    fontName = "SimSun";
                    break;
                case "新宋体":
                    fontName = "NSimSun";
                    break;
                case "黑体":
                    fontName = "SimHei";
                    break;
                case "楷体":
                    fontName = "KaiTi";
                    break;
                case "微软雅黑":
                    fontName = "Microsoft YaHei";
                    break;
                case "等线":
                    fontName = "Deng";
                    break;
                case "等线Light":
                    fontName = "Deng1";
                    break;
                case "仿宋":
                    fontName = "FangSong";
                    break;
                case "华文楷体":
                    fontName = "STKaiti";
                    break;
                case "华文宋体":
                    fontName = "STSong";
                    break;
                case "华文中宋":
                    fontName = "STZhongsong";
                    break;
                case "华文细黑":
                    fontName = "STXIHEI";
                    break;
                case "华文仿宋":
                    fontName = "STFANGSO";
                    break;
                case "行楷":
                    fontName = "STXINGKA";
                    break;
                case "华文新魏":
                    fontName = "STXinwei";
                    break;
                case "华文彩云":
                    fontName = "STCaiyun";
                    break;
                case "华文琥珀":
                    fontName = "STHupo";
                    break;
                case "华文隶书":
                    fontName = "STLiti";
                    break;
                case "华文黑体":
                    fontName = "STHeiti";
                    break;
                case "方正舒体":
                    fontName = "FZSTK";
                    break;
                case "方正姚体":
                    fontName = "FZYTK";
                    break;
                case "幼圆":
                    fontName = "YouYuan";
                    break;
                case "隶书":
                    fontName = "LiSu";
                    break;
                case "方正黑体":
                    fontName = "FZHeiTi";
                    break;
                case "方正仿宋":
                    fontName = "FZFangSong";
                    break;
                    // 添加其他常见字体的特殊处理
            }

            string fontFilePath = FindFontFilePathInRegistry(fontName, Microsoft.Win32.Registry.LocalMachine);
            if (!string.IsNullOrEmpty(fontFilePath))
            {
                return fontFilePath;
            }

            fontFilePath = FindFontFilePathInRegistry(fontName, Microsoft.Win32.Registry.CurrentUser);
            if (!string.IsNullOrEmpty(fontFilePath))
            {
                return fontFilePath;
            }

            fontFilePath = FindFontFilePathInDirectory(fontName, Environment.GetFolderPath(Environment.SpecialFolder.Fonts));
            if (!string.IsNullOrEmpty(fontFilePath))
            {
                return fontFilePath;
            }

            string userFontDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft\\Windows\\Fonts");
            fontFilePath = FindFontFilePathInDirectory(fontName, userFontDir);
            if (!string.IsNullOrEmpty(fontFilePath))
            {
                return fontFilePath;
            }

            return null;
        }

        private string FindFontFilePathInRegistry(string fontName, Microsoft.Win32.RegistryKey registryKey)
        {
            string fontFilePath = null;
            string fontsRegistryPath = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts";

            using (Microsoft.Win32.RegistryKey key = registryKey.OpenSubKey(fontsRegistryPath, false))
            {
                if (key != null)
                {
                    foreach (string fontRegName in key.GetValueNames())
                    {
                        if (System.Globalization.CultureInfo.CurrentCulture.CompareInfo.IndexOf(fontRegName, fontName, System.Globalization.CompareOptions.IgnoreCase) >= 0)
                        {
                            fontFilePath = key.GetValue(fontRegName) as string;
                            if (!string.IsNullOrEmpty(fontFilePath))
                            {
                                if (!Path.IsPathRooted(fontFilePath))
                                {
                                    fontFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), fontFilePath);
                                }
                                break;
                            }
                        }
                    }
                }
            }

            return fontFilePath;
        }

        private string FindFontFilePathInDirectory(string fontName, string directoryPath)
        {
            if (Directory.Exists(directoryPath))
            {
                var fontFiles = Directory.GetFiles(directoryPath, "*.*", SearchOption.TopDirectoryOnly)
                                          .Where(f => f.EndsWith(".ttf", StringComparison.OrdinalIgnoreCase) ||
                                                      f.EndsWith(".otf", StringComparison.OrdinalIgnoreCase));

                foreach (string fontFile in fontFiles)
                {
                    if (System.Globalization.CultureInfo.CurrentCulture.CompareInfo.IndexOf(Path.GetFileNameWithoutExtension(fontFile), fontName, System.Globalization.CompareOptions.IgnoreCase) >= 0)
                    {
                        return fontFile;
                    }
                }
            }
            return null;
        }
    }
}
