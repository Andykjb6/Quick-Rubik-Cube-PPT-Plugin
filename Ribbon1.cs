using System;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;  // 指定Shape引用的命名空间
using NPinyin;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing.Imaging;
using System.Net;
using System.Threading.Tasks;
using DrawingPoint = System.Drawing.Point;
using System.Drawing.Drawing2D;
using System.Numerics;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using System.Reflection;
using Microsoft.Office.Tools;
using System.Drawing.Text;
using NStandard;
using System.Web.UI.WebControls;
using Microsoft.Win32;
using System.Globalization;
using Core = Microsoft.Office.Core;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using Word = Microsoft.Office.Interop.Word;
using OfficeOpenXml;
using System.Net.Http;
using System.Collections.Concurrent;


namespace 课件帮PPT助手
{

    public partial class Ribbon1 : Office.IRibbonExtensibility
    {
        private CustomCloudTextGeneratorForm cloudTextGeneratorForm;

        public Ribbon1(RibbonFactory factory) : base(factory)
        {
            Debug.WriteLine("Ribbon1 constructor called.");
            InitializeComponent();
        }

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("课件帮PPT助手.Ribbon1.xml");
        }

        private static string GetResourceText(string resourceName)
        {
            var asm = typeof(Ribbon1).Assembly;
            using (var stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    return null;
                }
                using (var reader = new System.IO.StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        public void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            Debug.WriteLine("Ribbon_Load called.");
        }

        private void button5_Click_1(object sender, RibbonControlEventArgs e)
        {
            if (cloudTextGeneratorForm == null || cloudTextGeneratorForm.IsDisposed)
            {
                cloudTextGeneratorForm = new CustomCloudTextGeneratorForm();
            }

            cloudTextGeneratorForm.InitializeForm();

            // 设置窗体总在最前并激活
            cloudTextGeneratorForm.TopMost = true;
            cloudTextGeneratorForm.BringToFront();
            cloudTextGeneratorForm.Show();
        }



        // 定义全局变量
        private string selectedSVG;
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                foreach (Shape shape in selection.ShapeRange)
                {
                    AdjustShapeSizeAndAlign(shape);
                }
            }
            else if (selection.Type == PpSelectionType.ppSelectionSlides && selection.SlideRange.Count == 1)
            {
                Slide slide = selection.SlideRange[1];
                foreach (Shape shape in slide.Shapes)
                {
                    AdjustShapeSizeAndAlign(shape);
                }
            }
        }

        private void AdjustShapeSizeAndAlign(Shape shape)
        {
            float slideWidth = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth;
            float slideHeight = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight;
            float slideAspectRatio = slideWidth / slideHeight;

            if (shape.Type == MsoShapeType.msoPicture)
            {
                float pictureWidth = shape.Width;
                float pictureHeight = shape.Height;
                float pictureAspectRatio = pictureWidth / pictureHeight;

                shape.LockAspectRatio = MsoTriState.msoFalse; // 关闭锁定比例

                if (pictureAspectRatio > slideAspectRatio)
                {
                    // 图片过宽，调整宽度和裁剪
                    float newWidth = pictureHeight * slideAspectRatio;
                    shape.PictureFormat.CropLeft = (pictureWidth - newWidth) / 2;
                    shape.PictureFormat.CropRight = (pictureWidth - newWidth) / 2;
                    shape.Width = slideWidth;  // 直接设定宽度为幻灯片宽度
                    shape.Height = slideHeight; // 高度设为幻灯片高度
                }
                else
                {
                    // 图片过高，调整高度和裁剪
                    float newHeight = pictureWidth / slideAspectRatio;
                    shape.PictureFormat.CropTop = (pictureHeight - newHeight) / 2;
                    shape.PictureFormat.CropBottom = (pictureHeight - newHeight) / 2;
                    shape.Width = slideWidth;
                    shape.Height = slideHeight;
                }

                shape.Left = 0;
                shape.Top = 0;
            }
            else
            {
                // 如果是形状，直接调整大小以覆盖整个幻灯片
                shape.Width = slideWidth;
                shape.Height = slideHeight;
                shape.Left = 0;
                shape.Top = 0;
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 1)
            {
                Shape referenceShape = selection.ShapeRange[1];  // 第一个选中的形状或图片作为参考
                bool otherShapesExist = false;

                // 检查后续是否还有其他形状
                for (int i = 2; i <= selection.ShapeRange.Count; i++)
                {
                    Shape shape = selection.ShapeRange[i];
                    if (shape.Type == MsoShapeType.msoAutoShape ||
                        shape.Type == MsoShapeType.msoFreeform ||
                        shape.Type == MsoShapeType.msoGroup)
                    {
                        otherShapesExist = true;
                        break;
                    }
                }

                if (otherShapesExist && (referenceShape.Type == MsoShapeType.msoAutoShape ||
                                         referenceShape.Type == MsoShapeType.msoFreeform ||
                                         referenceShape.Type == MsoShapeType.msoGroup))
                {
                    MessageBox.Show("参考裁剪仅支持图片裁剪，第一个被选对象是图片大小裁剪的参考，可以是形状或图片", "裁剪信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                float referenceWidth = referenceShape.Width;
                float referenceHeight = referenceShape.Height;
                float referenceAspectRatio = referenceWidth / referenceHeight;

                // 遍历除第一个外的其他选中的形状或图片
                for (int i = 2; i <= selection.ShapeRange.Count; i++)
                {
                    Shape shape = selection.ShapeRange[i];
                    AdjustShapeToReference(referenceShape, shape, referenceAspectRatio);
                }
            }
        }

        private void AdjustShapeToReference(Shape referenceShape, Shape shapeToAdjust, float referenceAspectRatio)
        {
            float referenceHeight = referenceShape.Height;
            float referenceWidth = referenceShape.Width;

            // 设置第一个被选中的对象的高度为参考高度
            if (shapeToAdjust == referenceShape)
            {
                referenceShape.LockAspectRatio = MsoTriState.msoTrue;
                referenceShape.Height = referenceHeight;
            }
            else
            {
                float shapeWidth = shapeToAdjust.Width;
                float shapeHeight = shapeToAdjust.Height;
                float shapeAspectRatio = shapeWidth / shapeHeight;

                shapeToAdjust.LockAspectRatio = MsoTriState.msoFalse; // 允许改变宽高比

                // 调整形状尺寸以匹配参考形状的纵横比
                if (shapeAspectRatio > referenceAspectRatio)
                {
                    // 如果当前形状比例宽于参考比例，调整宽度
                    float newWidth = shapeHeight * referenceAspectRatio;  // 调整宽度以匹配宽高比
                    float cropWidth = (shapeWidth - newWidth) / 2;  // 计算需要裁剪的宽度
                    shapeToAdjust.PictureFormat.CropLeft = cropWidth;
                    shapeToAdjust.PictureFormat.CropRight = cropWidth;
                }
                else if (shapeAspectRatio < referenceAspectRatio)
                {
                    // 如果当前形状比例窄于参考比例，调整高度
                    float newHeight = shapeWidth / referenceAspectRatio;  // 调整高度以匹配宽高比
                    float cropHeight = (shapeHeight - newHeight) / 2;  // 计算需要裁剪的高度
                    shapeToAdjust.PictureFormat.CropTop = cropHeight;
                    shapeToAdjust.PictureFormat.CropBottom = cropHeight;
                }

                // 设置形状大小与参考形状一致
                shapeToAdjust.Width = referenceWidth;
                shapeToAdjust.Height = referenceHeight;
            }
        }


        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            GeneratePinyinForSelectedText();
        }

        private void GeneratePinyinForSelectedText()
        {
            try
            {
                Selection Sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                if (Sel.Type == PpSelectionType.ppSelectionShapes)
                {
                    Microsoft.Office.Interop.PowerPoint.ShapeRange sr = Sel.ShapeRange;
                    foreach (Shape shape in sr)
                    {
                        if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoTrue)
                        {
                            string text = shape.TextFrame.TextRange.Text;
                            string pinyin = ConvertToPinyin(text);
                            CreatePinyinShape(shape, pinyin);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error: " + ex.Message);
            }
        }

        private string ConvertToPinyin(string text)
        {
            // 简单的调用 NPinyin 获取拼音，假设无需额外参数处理声调
            return Pinyin.GetPinyin(text); // 获取拼音
        }

        private void CreatePinyinShape(Shape originShape, string pinyin)
        {
            Slide slide = originShape.Parent as Slide;
            float originFontSize = originShape.TextFrame.TextRange.Font.Size; // 获取原文本框的字号
            float newFontSize = originFontSize / 3; // 新文本框字号为原字号的二分之一

            float newShapeTop = originShape.Top - (originShape.Height / 4) - (newFontSize / 4); // 新文本框放置在原文本框的顶部，距离原文本框中心线一半字号的高度
            if (newShapeTop < 0) newShapeTop = originShape.Top + originShape.Height; // 如果超出幻灯片顶部，则放在下方

            Shape pinyinShape = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                originShape.Left,
                newShapeTop,
                originShape.Width,
                newFontSize); // 新文本框的高度与字号相同
            pinyinShape.TextFrame.TextRange.Text = pinyin;
            pinyinShape.TextFrame.TextRange.Font.Size = newFontSize; // 设置字体大小为原字号的二分之一
            pinyinShape.TextFrame.TextRange.Font.Name = "Arial"; // 设置字体，确保支持拼音符号

            // 设置新文本框的水平对齐方式与原文本框相同
            pinyinShape.TextFrame.TextRange.ParagraphFormat.Alignment = originShape.TextFrame.TextRange.ParagraphFormat.Alignment;
        }
        private Form form; // 保持 form 的引用，以便可以随时调用或修改它

        

        /// <summary>
        /// 在给定的SVG字符串中插入新的宽度和高度属性。
        /// </summary>
        /// <param name="svg">原始的SVG字符串</param>
        /// <param name="width">要插入的宽度值</param>
        /// <param name="height">要插入的高度值</param>
        /// <returns>带有新属性的SVG字符串</returns>
        private string InsertSvgAttributes(string svg, string width, string height)
        {
            int index = svg.IndexOf("<svg ");
            if (index != -1)
            {
                // 找到<svg后的第一个空格位置
                int spaceIndex = svg.IndexOf(' ', index);
                if (spaceIndex != -1)
                {
                    string attributes = "width='" + width + "' height='" + height + "' ";
                    return svg.Insert(spaceIndex + 1, attributes);
                }
            }
            return svg;
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 获取剪贴板内容
                if (Clipboard.ContainsText())
                {
                    string clipboardText = Clipboard.GetText();

                    string pattern = @"width:(\s*\d+)px;\s*height:(\s*\d+)px;"; // 修改正则表达式，去掉匹配px并且调整捕获组

                    Match match = Regex.Match(clipboardText, pattern);

                    // 检查是否包含SVG代码
                    if (Regex.IsMatch(clipboardText, "<svg", RegexOptions.IgnoreCase))
                    {
                        // 含有style的
                        if (match.Success)
                        {
                            string width = match.Groups[1].Value.Trim(); // 去掉空格
                            string height = match.Groups[2].Value.Trim(); // 去掉空格

                            // 插入新属性到SVG标签中
                            string updatedSvg = InsertSvgAttributes(clipboardText, width, height);

                            // 保存SVG代码到文件
                            string tempSvgPath = Path.GetTempFileName() + ".svg";
                            File.WriteAllText(tempSvgPath, updatedSvg);

                            // 插入SVG到PPT
                            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
                            PowerPoint.Slide slide = pptApp.ActiveWindow.View.Slide;
                            slide.Shapes.AddPicture(tempSvgPath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoCTrue, 100, 100);

                            // 删除临时文件
                            File.Delete(tempSvgPath);
                        }
                        // 如果没有style的
                        else
                        {
                            // 保存SVG代码到文件
                            string tempSvgPath = Path.GetTempFileName() + ".svg";
                            File.WriteAllText(tempSvgPath, clipboardText);

                            // 插入SVG到PPT
                            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
                            PowerPoint.Slide slide = pptApp.ActiveWindow.View.Slide;
                            slide.Shapes.AddPicture(tempSvgPath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoCTrue, 100, 100, 50, 50);

                            // 删除临时文件
                            File.Delete(tempSvgPath);
                        }
                    }
                    else
                    {
                        MessageBox.Show("剪贴板内容不包含SVG代码。");
                    }


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("出现错误：" + ex.Message);
            }
        }

        #region Helper methods

        private static string getResourceText(string resourceName)
        {
            var asm = System.Reflection.Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            foreach (string resource in resourceNames)
            {
                if (string.Compare(resourceName, resource, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resource)))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            OpenWebPage("https://flowus.cn/andyblog/share/6da481ac-a57b-4214-9ce8-94273bbf2f45?code=GEH4ZC");
        }

        private void OpenWebPage(string url)
        {
            try
            {
                Process.Start(url);
            }
            catch (Exception ex)
            {
                Console.WriteLine("打开网页时出错：" + ex.Message);
            }
        }
      


        private void button10_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = pptApp.ActiveWindow.Selection;

            // 检查是否至少选择了两个对象
            if (selection.ShapeRange.Count < 2)
            {
                // 提示用户至少选择两个对象
                System.Windows.Forms.MessageBox.Show("请选择至少两个对象。");
                return;
            }

            // 获取第一个被选中的对象
            PowerPoint.Shape firstShape = selection.ShapeRange[1];

            // 记录第一个对象的位置
            float referenceLeft = firstShape.Left;
            float referenceTop = firstShape.Top;

            // 循环对齐后续被选中的对象
            for (int i = 2; i <= selection.ShapeRange.Count; i++)
            {
                PowerPoint.Shape currentShape = selection.ShapeRange[i];

                // 计算居中对齐的位置
                float newLeft = referenceLeft + (firstShape.Width - currentShape.Width) / 2;
                float newTop = referenceTop + (firstShape.Height - currentShape.Height) / 2;

                // 设置对象位置
                currentShape.Left = newLeft;
                currentShape.Top = newTop;
            }
        }


        private void button11_Click_1(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = pptApp.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count >= 2)
            {
                PowerPoint.ShapeRange selectedShapes = selection.ShapeRange;

                // 获取第一个形状所在的幻灯片
                PowerPoint.Slide firstSlide = selectedShapes[1].Parent;

                // 检查所有选定的形状是否都在同一个幻灯片上
                foreach (PowerPoint.Shape shape in selectedShapes)
                {
                    if (shape.Parent != firstSlide)
                    {
                        System.Windows.Forms.MessageBox.Show("所有形状必须在同一个幻灯片上。");
                        return;
                    }
                }

                // 对选定的形状进行排序
                selectedShapes.ZOrder(Office.MsoZOrderCmd.msoSendToBack);

                // 获取第一个被选中的对象的位置
                PowerPoint.Shape firstShape = selectedShapes[1];
                float firstLeft = firstShape.Left;
                float firstTop = firstShape.Top;
                float firstHeight = firstShape.Height;

                // 计算每个形状的新位置
                float currentLeft = firstLeft;
                float currentTop = firstTop;
                foreach (PowerPoint.Shape shape in selectedShapes)
                {
                    shape.Left = currentLeft;
                    shape.Top = currentTop;
                    currentLeft += shape.Width;
                }
            }
        }

       

        private void button13_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = pptApp.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count >= 2)
            {
                PowerPoint.ShapeRange selectedShapes = selection.ShapeRange;

                // 获取第一个形状所在的幻灯片
                PowerPoint.Slide firstSlide = selectedShapes[1].Parent;

                // 检查所有选定的形状是否都在同一个幻灯片上
                foreach (PowerPoint.Shape shape in selectedShapes)
                {
                    if (shape.Parent != firstSlide)
                    {
                        System.Windows.Forms.MessageBox.Show("所有形状必须在同一个幻灯片上。");
                        return;
                    }
                }

                // 获取第一个被选中的对象的位置
                PowerPoint.Shape firstShape = selectedShapes[1];
                float firstLeft = firstShape.Left;
                float firstTop = firstShape.Top;
                float lastBottom = firstTop + firstShape.Height;

                // 根据第一个对象的位置对后续对象进行对齐
                foreach (PowerPoint.Shape shape in selectedShapes)
                {
                    if (shape != firstShape && shape.Top != firstTop)
                    {
                        float topGap = shape.Top - lastBottom;
                        shape.Top = lastBottom + (topGap > 0 ? topGap : 0);
                    }
                    lastBottom = shape.Top + shape.Height;
                }

                // 对所有对象进行左对齐
                float minLeft = float.MaxValue;
                foreach (PowerPoint.Shape shape in selectedShapes)
                {
                    if (shape.Left < minLeft)
                    {
                        minLeft = shape.Left;
                    }
                }
                foreach (PowerPoint.Shape shape in selectedShapes)
                {
                    shape.Left = minLeft;
                }
            }
        }

        private void button19_Click(object sender, RibbonControlEventArgs e)
        {
            string url = "https://miankoutupian.com/ai/cutout";
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法打开链接: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

       


        [DllImport("user32.dll")]
        private static extern short GetKeyState(int keyCode);

        private void Replaceimage_Click(object sender, RibbonControlEventArgs e)
        {
            const int VK_CONTROL = 0x11;
            bool isCtrlPressed = (GetKeyState(VK_CONTROL) & 0x8000) != 0;

            var application = Globals.ThisAddIn.Application;
            var activeWindow = application.ActiveWindow;

            if (activeWindow.Selection.Type != PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("请选择一张或多张图片");
                return;
            }

            var selectedShapes = activeWindow.Selection.ShapeRange.Cast<Shape>().ToList();

            if (!selectedShapes.Any())
            {
                MessageBox.Show("请选择一张或多张图片");
                return;
            }

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif",
                Title = "Select images to replace the selected shapes"
            };

            if (openFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            var selectedFiles = openFileDialog.FileNames;
            int fileIndex = 0;

            foreach (var shape in selectedShapes)
            {
                ReplaceShapeImage(shape, selectedFiles, ref fileIndex, isCtrlPressed);
            }

            // Handle additional images
            while (fileIndex < selectedFiles.Length)
            {
                string filePath = selectedFiles[fileIndex++];
                activeWindow.View.Slide.Shapes.AddPicture(filePath, MsoTriState.msoFalse, MsoTriState.msoCTrue, 0, 0, -1, -1);
            }
        }

        private void ReplaceShapeImage(Shape shape, string[] selectedFiles, ref int fileIndex, bool isCtrlPressed)
        {
            if (fileIndex >= selectedFiles.Length)
                return;

            if (shape.Type == MsoShapeType.msoGroup)
            {
                foreach (Shape subShape in shape.GroupItems)
                {
                    ReplaceShapeImage(subShape, selectedFiles, ref fileIndex, isCtrlPressed);
                }
            }
            else
            {
                var application = Globals.ThisAddIn.Application;
                var filePath = selectedFiles[fileIndex++];

                float left = shape.Left;
                float top = shape.Top;
                float originalWidth = shape.Width;
                float originalHeight = shape.Height;
                float originalAspectRatio = originalWidth / originalHeight;

                Shape newShape = null;
                if (shape.Type == MsoShapeType.msoPicture)
                {
                    var rotation = shape.Rotation;
                    var zOrderPosition = shape.ZOrderPosition;
                    var fill = shape.Fill;
                    var line = shape.Line;
                    var shadow = shape.Shadow;
                    var lockAspectRatio = shape.LockAspectRatio;

                    shape.Delete();

                    newShape = application.ActiveWindow.View.Slide.Shapes.AddPicture(filePath, MsoTriState.msoFalse, MsoTriState.msoCTrue, left, top, -1, -1);

                    // Get new shape dimensions and aspect ratio
                    float newWidth = newShape.Width;
                    float newHeight = newShape.Height;
                    float newAspectRatio = newWidth / newHeight;

                    // Apply cropping to maintain the original aspect ratio
                    if (newAspectRatio > originalAspectRatio)
                    {
                        float cropAmount = (newWidth - originalAspectRatio * newHeight) / 2;
                        newShape.PictureFormat.CropLeft = cropAmount;
                        newShape.PictureFormat.CropRight = cropAmount;
                    }
                    else
                    {
                        float cropAmount = (newHeight - newWidth / originalAspectRatio) / 2;
                        newShape.PictureFormat.CropTop = cropAmount;
                        newShape.PictureFormat.CropBottom = cropAmount;
                    }

                    // Resize new shape to match original dimensions
                    newShape.LockAspectRatio = MsoTriState.msoTrue;
                    newShape.Width = originalWidth;
                    newShape.Height = originalHeight;
                    newShape.Left = left;
                    newShape.Top = top;
                    newShape.Rotation = rotation;
                    newShape.ZOrder(MsoZOrderCmd.msoSendBackward);

                    // Copy Fill properties
                    if (newShape.Fill != null && fill != null)
                    {
                        try
                        {
                            newShape.Fill.BackColor.RGB = fill.BackColor.RGB;
                            newShape.Fill.ForeColor.RGB = fill.ForeColor.RGB;
                            newShape.Fill.Transparency = fill.Transparency;
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            // Handle the exception if necessary
                        }
                    }

                    // Copy Line properties
                    if (newShape.Line != null && line != null)
                    {
                        try
                        {
                            newShape.Line.ForeColor.RGB = line.ForeColor.RGB;
                            newShape.Line.BackColor.RGB = line.BackColor.RGB;
                            if (line.Weight >= 0.25f && line.Weight <= 6.0f) // Example valid range
                            {
                                newShape.Line.Weight = line.Weight;
                            }
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            // Handle the exception if necessary
                        }
                        catch (System.ArgumentException)
                        {
                            // Handle the exception if necessary
                        }
                    }

                    // Copy Shadow properties
                    if (newShape.Shadow != null && shadow != null)
                    {
                        try
                        {
                            newShape.Shadow.ForeColor.RGB = shadow.ForeColor.RGB;
                            newShape.Shadow.Visible = shadow.Visible;
                            newShape.Shadow.OffsetX = shadow.OffsetX;
                            newShape.Shadow.OffsetY = shadow.OffsetY;
                            newShape.Shadow.Transparency = shadow.Transparency;
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            // Handle the exception if necessary
                        }
                    }

                    newShape.LockAspectRatio = lockAspectRatio;

                    // Ensure no distortion
                    if (!isCtrlPressed)
                    {
                        // Adjust aspect ratio without cropping
                        newWidth = newShape.Width;
                        newHeight = newShape.Height;
                        newAspectRatio = newWidth / newHeight;

                        if (newAspectRatio > originalAspectRatio)
                        {
                            newHeight = originalHeight;
                            newWidth = newHeight * newAspectRatio;
                        }
                        else
                        {
                            newWidth = originalWidth;
                            newHeight = newWidth / newAspectRatio;
                        }

                        newShape.Width = newWidth;
                        newShape.Height = newHeight;
                    }
                }
                else
                {
                    shape.Fill.UserPicture(filePath);
                    newShape = shape;
                }
            }
        }

        private void Replaceaudio_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 获取当前的PPT应用实例
                Microsoft.Office.Interop.PowerPoint.Application pptApplication = Globals.ThisAddIn.Application;
                // 获取当前的幻灯片
                Slide activeSlide = pptApplication.ActiveWindow.View.Slide;

                // 检查是否有选中的形状
                if (pptApplication.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionShapes)
                {
                    Shape selectedShape = pptApplication.ActiveWindow.Selection.ShapeRange[1];

                    // 检查选中的形状是否为媒体对象
                    if (selectedShape.MediaType == PpMediaType.ppMediaTypeSound)
                    {
                        // 显示文件对话框，选择新音频文件
                        OpenFileDialog openFileDialog = new OpenFileDialog
                        {
                            Filter = "音频文件 (*.mp3;*.wav)|*.mp3;*.wav",
                            Title = "选择音频文件"
                        };

                        if (openFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            string newAudioFilePath = openFileDialog.FileName;
                            float left = selectedShape.Left;
                            float top = selectedShape.Top;

                            // 保存原音频的动作设置
                            TimeLine timeLine = activeSlide.TimeLine;
                            Sequence sequence = timeLine.MainSequence;
                            List<Effect> originalEffects = new List<Effect>();

                            foreach (Effect effect in sequence)
                            {
                                if (effect.Shape == selectedShape)
                                {
                                    originalEffects.Add(effect);
                                }
                            }

                            // 插入新音频文件
                            Shape newAudioShape = activeSlide.Shapes.AddMediaObject2(newAudioFilePath, MsoTriState.msoFalse, MsoTriState.msoCTrue, left, top);

                            // 删除原来的音频图标
                            selectedShape.Delete();

                            // 将原音频的动作设置应用到新音频上
                            foreach (Effect originalEffect in originalEffects)
                            {
                                Effect newEffect = sequence.AddEffect(newAudioShape, originalEffect.EffectType);

                                // 复制效果的时间设置
                                newEffect.Timing.Duration = originalEffect.Timing.Duration;
                                newEffect.Timing.TriggerDelayTime = originalEffect.Timing.TriggerDelayTime;
                                newEffect.Timing.RepeatCount = originalEffect.Timing.RepeatCount;
                                newEffect.Timing.Speed = originalEffect.Timing.Speed;
                            }

                            MessageBox.Show("音频替换成功，并保留了原音频的动作设置！", "课件帮PPT助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        MessageBox.Show("请选择一个音频图标进行替换。", "课件帮PPT助手", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    MessageBox.Show("请先选中一个音频图标。", "课件帮PPT助手", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception)
            {
              
            }
        }

        private void Fillblank_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 获取当前的PPT应用实例
                Microsoft.Office.Interop.PowerPoint.Application pptApplication = Globals.ThisAddIn.Application;
                // 获取当前的幻灯片
                Slide activeSlide = pptApplication.ActiveWindow.View.Slide;

                // 检查是否有选中的文本框
                if (pptApplication.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionText)
                {
                    TextRange selectedTextRange = pptApplication.ActiveWindow.Selection.TextRange;
                    Shape textBoxShape = pptApplication.ActiveWindow.Selection.ShapeRange[1];

                    // 获取选中文本的位置和大小
                    float left = textBoxShape.Left + selectedTextRange.BoundLeft;
                    float top = textBoxShape.Top + selectedTextRange.BoundTop;
                    float width = selectedTextRange.BoundWidth;
                    float height = selectedTextRange.BoundHeight;

                    // 复制选中的文本内容
                    string selectedText = selectedTextRange.Text;

                    // 替换选中的文本为下划线
                    string underlineText = new string('_', selectedText.Length + 2); // 确保下划线稍长
                    selectedTextRange.Text = underlineText;
                    selectedTextRange.Font.Underline = MsoTriState.msoTrue;

                    // 在临时形状的位置创建一个新的文本框，包含选中的文本
                    Shape newTextBox = activeSlide.Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        left,
                        top,
                        width,
                        height);
                    TextRange newTextRange = newTextBox.TextFrame.TextRange;
                    newTextRange.Text = selectedText;
                    newTextRange.Font.Name = selectedTextRange.Font.Name;
                    newTextRange.Font.Size = selectedTextRange.Font.Size;
                    newTextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    newTextRange.Font.Bold = MsoTriState.msoTrue;

                    // 将新的文本框置于前面，以确保它显示在下划线上方
                    newTextBox.ZOrder(MsoZOrderCmd.msoBringToFront);

                    MessageBox.Show("文本替换成功，并创建了一个新的文本框！", "课件帮PPT助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("请先选中文本框内的文本。", "课件帮PPT助手", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"发生错误: {ex.Message}", "课件帮PPT助手", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void comboBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {
           
                // 获取当前活动窗口
                PowerPoint.DocumentWindow activeWindow = Globals.ThisAddIn.Application.ActiveWindow;
                if (activeWindow != null)
                {
                    PowerPoint.Presentation presentation = activeWindow.Presentation;

                    // 根据用户选择的页面尺寸更改PPT页面设置
                    switch (comboBox1.Text)
                    {
                        case "A1":
                            SetPageSize(presentation, 841, 594); // A1: 841mm x 594mm
                            break;
                        case "A2":
                            SetPageSize(presentation, 594, 420); // A2: 594mm x 420mm
                            break;
                        case "A3":
                            SetPageSize(presentation, 420, 297); // A3: 420mm x 297mm
                            break;
                        case "A4":
                            SetPageSize(presentation, 297, 210); // A4: 297mm x 210mm
                            break;
                        case "16:9":
                            SetAspectRatio(presentation, 1920, 1080); // 16:9 比例
                            break;
                        case "4:3":
                            SetAspectRatio(presentation, 1280, 960); // 4:3 比例
                            break;
                        case "公众号封面":
                            SetAspectRatio(presentation, 2.35f, 1); // 2.35:1 比例
                            break;
                        case "小红书图文":
                            SetAspectRatio(presentation, 3, 4); // 3:4 比例
                            break;
                        default:
                            MessageBox.Show("未知页面尺寸。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            break;
                    }
                }
            }

            private void SetPageSize(PowerPoint.Presentation presentation, float width, float height)
            {
                // 将毫米转换为英寸（1英寸 = 25.4毫米）
                float widthInches = width / 25.4f;
                float heightInches = height / 25.4f;

                // 设置页面尺寸
                presentation.PageSetup.SlideWidth = widthInches * 72; // 1英寸 = 72点
                presentation.PageSetup.SlideHeight = heightInches * 72;
            }

            private void SetAspectRatio(PowerPoint.Presentation presentation, float widthRatio, float heightRatio)
            {
                // 获取当前页面宽度并计算相应的高度
                float currentWidth = presentation.PageSetup.SlideWidth;
                float calculatedHeight = (currentWidth / widthRatio) * heightRatio;

                // 设置页面尺寸
                presentation.PageSetup.SlideHeight = calculatedHeight;
            }






            private void comboBox2_TextChanged(object sender, RibbonControlEventArgs e)
        {
            // 获取当前活动窗口
            PowerPoint.DocumentWindow activeWindow = Globals.ThisAddIn.Application.ActiveWindow;
            if (activeWindow != null)
            {
                PowerPoint.Presentation presentation = activeWindow.Presentation;

                // 获取当前页面尺寸
                float slideWidth = presentation.PageSetup.SlideWidth;
                float slideHeight = presentation.PageSetup.SlideHeight;

                // 根据用户选择的页面方向更改PPT页面设置
                switch (comboBox2.Text)
                {
                    case "纵向":
                        if (slideWidth > slideHeight) // 目前是横向，需要调整为纵向
                        {
                            SetPageOrientation(presentation, slideHeight, slideWidth); // 交换宽度和高度
                        }
                        break;
                    case "横向":
                        if (slideWidth < slideHeight) // 目前是纵向，需要调整为横向
                        {
                            SetPageOrientation(presentation, slideHeight, slideWidth); // 交换宽度和高度
                        }
                        break;
                    default:
                        MessageBox.Show("未知页面方向。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                }
            }
        }

        private void SetPageOrientation(PowerPoint.Presentation presentation, float width, float height)
        {
            // 设置页面尺寸
            presentation.PageSetup.SlideWidth = width;
            presentation.PageSetup.SlideHeight = height;
        }

      

        private void Gradientrectangle_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前的PowerPoint应用程序实例
            PowerPoint.Application application = Globals.ThisAddIn.Application;

            // 检查当前视图是否支持选择操作
            if (application.ActiveWindow.ViewType == PowerPoint.PpViewType.ppViewNormal)
            {
                // 获取当前的演示文稿
                PowerPoint.Presentation presentation = application.ActivePresentation;

                // 检查是否有选中的对象
                if (application.ActiveWindow.Selection.Type != PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    PowerPoint.Shape selectedShape = application.ActiveWindow.Selection.ShapeRange[1];

                    // 创建渐变透明矩形
                    PowerPoint.Shape rectangle = null;

                    if (selectedShape.Type == Office.MsoShapeType.msoPlaceholder)
                    {
                        // 如果选中的是幻灯片
                        PowerPoint.Slide slide = selectedShape.Parent as PowerPoint.Slide;
                        if (slide != null)
                        {
                            rectangle = slide.Shapes.AddShape(
                                Office.MsoAutoShapeType.msoShapeRectangle,
                                0, 0, slide.Master.Width, slide.Master.Height);
                        }
                    }
                    else
                    {
                        // 如果选中的是其他对象
                        rectangle = presentation.Slides[selectedShape.Parent.SlideIndex].Shapes.AddShape(
                            Office.MsoAutoShapeType.msoShapeRectangle,
                            selectedShape.Left, selectedShape.Top, selectedShape.Width, selectedShape.Height);
                    }

                    // 设置边框为不可见
                    rectangle.Line.Visible = Office.MsoTriState.msoFalse;

                    // 设置填充为渐变
                    rectangle.Fill.OneColorGradient(Office.MsoGradientStyle.msoGradientHorizontal, 1, 1);

                    // 设置渐变方向
                    if (Control.ModifierKeys == System.Windows.Forms.Keys.Control) // 如果按下了Ctrl键
                    {
                        rectangle.Fill.GradientAngle = 90; // 从上往下
                    }
                    else if (Control.ModifierKeys == System.Windows.Forms.Keys.Shift) // 如果按下了Shift键
                    {
                        rectangle.Fill.GradientAngle = 45; // 对角线方向
                    }
                    else
                    {
                        rectangle.Fill.GradientAngle = 0; // 默认从左到右
                    }

                    // 设置渐变色标
                    rectangle.Fill.GradientStops[1].Color.RGB = 0x000000; // 黑色
                    rectangle.Fill.GradientStops[1].Transparency = 1.0f; // 透明度100%
                    rectangle.Fill.GradientStops[2].Color.RGB = 0x000000; // 黑色
                    rectangle.Fill.GradientStops[2].Transparency = 0.0f; // 透明度0%

                    // 将矩形置于选中对象的顶层
                    rectangle.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                }
                else
                {
                    // 当没有选中对象时，默认在当前幻灯片上插入与幻灯片等大的渐变透明矩形
                    PowerPoint.Slide slide = application.ActiveWindow.View.Slide as PowerPoint.Slide;

                    if (slide != null)
                    {
                        // 插入一个与幻灯片等大的矩形
                        PowerPoint.Shape rectangle = slide.Shapes.AddShape(
                            Office.MsoAutoShapeType.msoShapeRectangle,
                            0, 0, slide.Master.Width, slide.Master.Height);

                        // 设置边框为不可见
                        rectangle.Line.Visible = Office.MsoTriState.msoFalse;

                        // 设置填充为渐变
                        rectangle.Fill.OneColorGradient(Office.MsoGradientStyle.msoGradientHorizontal, 1, 1);

                        // 设置渐变方向
                        if (Control.ModifierKeys == System.Windows.Forms.Keys.Control) // 如果按下了Ctrl键
                        {
                            rectangle.Fill.GradientAngle = 90; // 从上往下
                        }
                        else if (Control.ModifierKeys == System.Windows.Forms.Keys.Shift) // 如果按下了Shift键
                        {
                            rectangle.Fill.GradientAngle = 45; // 对角线方向
                        }
                        else
                        {
                            rectangle.Fill.GradientAngle = 0; // 默认从左到右
                        }

                        // 设置渐变色标
                        rectangle.Fill.GradientStops[1].Color.RGB = 0x000000; // 黑色
                        rectangle.Fill.GradientStops[1].Transparency = 1.0f; // 透明度100%
                        rectangle.Fill.GradientStops[2].Color.RGB = 0x000000; // 黑色
                        rectangle.Fill.GradientStops[2].Transparency = 0.0f; // 透明度0%

                        // 将矩形置于幻灯片的底层
                        rectangle.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
                    }
                    else
                    {
                        MessageBox.Show("当前没有选中的幻灯片。");
                    }
                }
            }
            else
            {
                MessageBox.Show("请切换到Normal视图以进行选择操作。");
            }
        }

        private void Pagecentered_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前活动的PowerPoint应用程序
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;

            // 获取当前选中的幻灯片
            PowerPoint.Slide currentSlide = pptApp.ActiveWindow.View.Slide;

            // 获取当前选中的对象
            PowerPoint.Selection selection = pptApp.ActiveWindow.Selection;

            // 计算幻灯片的中心位置
            float slideCenterX = currentSlide.Master.Width / 2;
            float slideCenterY = currentSlide.Master.Height / 2;

            // 确保至少选中了一个对象
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionNone)
            {
                // 遍历选中的每一个对象
                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    // 计算对象的中心位置
                    float shapeCenterX = shape.Left + shape.Width / 2;
                    float shapeCenterY = shape.Top + shape.Height / 2;

                    // 计算需要移动的距离
                    float deltaX = slideCenterX - shapeCenterX;
                    float deltaY = slideCenterY - shapeCenterY;

                    // 移动对象到幻灯片的中心位置
                    shape.Left += deltaX;
                    shape.Top += deltaY;
                }
            }
        }

       

        private static List<Shape> hiddenShapes = new List<Shape>();
        private void button20_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前活动的PPT应用程序
            Application pptApplication = Globals.ThisAddIn.Application;
            // 获取当前活动的窗口
            DocumentWindow activeWindow = pptApplication.ActiveWindow;
            // 获取当前选中的对象
            Selection selection = activeWindow.Selection;

            if (hiddenShapes.Count > 0)
            {
                // 恢复隐藏的对象
                foreach (Shape shape in hiddenShapes)
                {
                    shape.Visible = Office.MsoTriState.msoTrue;
                }
                hiddenShapes.Clear();
            }
            else
            {
                if (selection.Type == PpSelectionType.ppSelectionShapes)
                {
                    foreach (Shape shape in selection.ShapeRange)
                    {
                        if (shape.Visible == Office.MsoTriState.msoTrue)
                        {
                            shape.Visible = Office.MsoTriState.msoFalse;
                            hiddenShapes.Add(shape);
                        }
                    }

                    // 刷新视图
                    activeWindow.View.GotoSlide(activeWindow.View.Slide.SlideIndex);
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("请选择一个或多个对象。");
                }
            }
        }

   

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }
        private void toggleTaskPane_Click(object sender, RibbonControlEventArgs e)
        {
            var addIn = Globals.ThisAddIn;
            addIn.ToggleTaskPaneVisibility();
        }


        private void 匹配对齐_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            AlignToolWindow alignToolWindow = new AlignToolWindow(app);
            alignToolWindow.Show();
        }

        private void 平移居中_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            CentralalignmentForm form = new CentralalignmentForm(pptApp);
            form.Show();
        }

        private void Masking_Click_1(object sender, RibbonControlEventArgs e)
        {
            TransparencyForm transparencyForm = new TransparencyForm();
            transparencyForm.Show(); // 使用 Show 方法以非模态方式显示窗体
        }

       

        private InputForm inputForm;
        private void 批量改字_Click(object sender, RibbonControlEventArgs e)
        {
            if (inputForm == null || inputForm.IsDisposed)
            {
                inputForm = new InputForm();
                inputForm.TextConfirmed += OnTextConfirmed;
                inputForm.Show(); // 非模式对话框
            }
            else
            {
                inputForm.BringToFront();
            }
        }

        private void OnTextConfirmed(string replacementText)
        {
            if (string.IsNullOrEmpty(replacementText))
            {
                return; // 如果用户未输入任何文本，则不执行替换操作
            }

            var shapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in shapes)
            {
                if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoTextBox)
                {
                    shape.TextFrame.TextRange.Text = replacementText;
                }
            }
        }

        private PinyinSelectorForm pinyinForm;
        private void 便捷注音_Click(object sender, RibbonControlEventArgs e)
        {
            if (pinyinForm == null || pinyinForm.IsDisposed)
            {
                pinyinForm = new PinyinSelectorForm();
                pinyinForm.Show();
            }
            else
            {
                pinyinForm.BringToFront();
            }

            pinyinForm.UpdateComboBoxOptions();
        }

        public void PinyinSelectorFormClosed()
        {
            pinyinForm = null;
        }

        //据字查笔顺
        private void 笔顺图解_Click(object sender, RibbonControlEventArgs e)
        {
            string inputChar = Microsoft.VisualBasic.Interaction.InputBox("请输入目标汉字（需联网，且一次仅支持查询单个汉字）:", "一键获取汉字笔顺图解", "");
            if (!string.IsNullOrWhiteSpace(inputChar))
            {
                string url = $"https://hanyu.baidu.com/s?wd={inputChar}&ptype=zici";
                ExtractSVGFromWebpage(url, inputChar);
            }
        }

        private void ExtractSVGFromWebpage(string url, string inputChar)
        {
            try
            {
                HtmlWeb web = new HtmlWeb();
                HtmlAgilityPack.HtmlDocument doc = web.Load(url);
                var svgNodes = doc.DocumentNode.SelectNodes("//svg");
                if (svgNodes != null)
                {
                    ShowSVGSelectionWindow(svgNodes, inputChar);
                }
                else
                {
                    MessageBox.Show("未查询到对应SVG笔顺图！");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("出现错误：" + ex.Message);
            }
        }

        private void ShowSVGSelectionWindow(HtmlNodeCollection svgNodes, string inputChar)
        {
            SvgSelectionForm svgSelectionForm = new SvgSelectionForm(svgNodes, inputChar);
            svgSelectionForm.ShowDialog();
        }
        
        //田字格生成
        private TableSettingsForm settingsForm;
        private void 生字格子_Click(object sender, RibbonControlEventArgs e)
        {
            if (settingsForm == null || settingsForm.IsDisposed)
            {
                settingsForm = new TableSettingsForm();
            }

            settingsForm.Show();
            settingsForm.TopMost = true;
        }

        //给生字创建田字格
        private TableSettingsFormButton12 settingsFormButton12;
        private Color borderColorButton12 = Color.Black;
        private void 生字赋格_Click(object sender, RibbonControlEventArgs e)
        {
            if (settingsFormButton12 == null || settingsFormButton12.IsDisposed)
            {
                settingsFormButton12 = new TableSettingsFormButton12();
            }

            settingsFormButton12.Show();
            settingsFormButton12.TopMost = true;
        }

       
        private void 原位复制_TextChanged(object sender, RibbonControlEventArgs e)
        {
            // 获取当前活动的PowerPoint应用程序
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;

            // 获取当前选中的幻灯片
            PowerPoint.Slide currentSlide = pptApp.ActiveWindow.View.Slide;

            // 获取当前选中的对象
            PowerPoint.Selection selection = pptApp.ActiveWindow.Selection;

            // 确保至少选中了一个对象
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionNone)
            {
                // 解析用户输入的复制次数
                string input = ((RibbonEditBox)sender).Text.Trim();
                int copyCount;

                if (!int.TryParse(input, out copyCount) || copyCount < 1)
                {
                    MessageBox.Show("请输入一个大于0的整数。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 复制选中的对象指定次数
                for (int i = 0; i < copyCount; i++)
                {
                    DuplicateSelectedShapes(selection);
                }

                // 清空输入框内容
                ((RibbonEditBox)sender).Text = string.Empty;
            }
            else
            {
                MessageBox.Show("请选择一个或多个对象。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // 复制选中的对象并置于原对象的上一层
        private void DuplicateSelectedShapes(PowerPoint.Selection selection)
        {
            // 确保至少选中了一个对象
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionNone)
            {
                // 遍历选中的每一个对象
                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    // 复制选中的对象
                    PowerPoint.Shape copiedShape = shape.Duplicate()[1];

                    // 将复制的对象置于选中对象的上一层
                    copiedShape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringForward);

                    // 将复制的对象移动到与选中对象相同的位置
                    copiedShape.Left = shape.Left;
                    copiedShape.Top = shape.Top;
                }
            }
        }

        private void 尺寸缩放_TextChanged(object sender, RibbonControlEventArgs e)
        {
            // 获取当前活动的PowerPoint应用程序
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;

            // 获取当前选中的对象
            PowerPoint.Selection selection = pptApp.ActiveWindow.Selection;

            // 确保至少选中了一个对象
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionNone)
            {
                // 解析用户输入的缩放比例
                string input = ((RibbonEditBox)sender).Text.Trim();
                string[] scaleValues = input.Split(',');

                // 确定缩放方式
                bool isArithmetic = scaleValues.Length == 2;

                // 计算等差缩放的公差
                float commonDifference = 0;
                if (isArithmetic)
                {
                    float startScale, endScale;
                    if (!float.TryParse(scaleValues[0], out startScale) || !float.TryParse(scaleValues[1], out endScale))
                    {
                        MessageBox.Show("请输入有效的缩放比例。");
                        return;
                    }

                    // 计算等差缩放的公差
                    commonDifference = (endScale - startScale) / (selection.ShapeRange.Count - 1);
                }

                // 记录当前缩放比例
                float currentScale = 0;
                if (!float.TryParse(scaleValues[0], out currentScale))
                {
                    MessageBox.Show("请输入有效的缩放比例。");
                    return;
                }

                // 遍历选中的每一个对象
                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    // 执行缩放
                    ScaleShape(shape, currentScale);

                    // 更新缩放比例
                    if (isArithmetic)
                    {
                        currentScale += commonDifference;
                    }
                }

                // 清空输入框内容
                ((RibbonEditBox)sender).Text = string.Empty;
            }
            else
            {
                MessageBox.Show("请选择一个或多个对象。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // 缩放指定的形状
        private void ScaleShape(PowerPoint.Shape shape, float scale)
        {
            // 计算缩放后的宽度和高度
            float newWidth = shape.Width * scale / 100;
            float newHeight = shape.Height * scale / 100;

            // 计算缩放后的左上角位置
            float newX = shape.Left + (shape.Width - newWidth) / 2;
            float newY = shape.Top + (shape.Height - newHeight) / 2;

            // 执行缩放
            shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
            shape.Width = newWidth;
            shape.Height = newHeight;
            shape.Left = newX;
            shape.Top = newY;
        }

        private void 批量命名_TextChanged(object sender, RibbonControlEventArgs e)
        {
            RibbonEditBox editBox = (RibbonEditBox)sender;
            string prefix = editBox.Text;

            if (string.IsNullOrEmpty(prefix))
            {
                return; // 如果用户未输入任何前缀，则不执行批量命名操作
            }

            // 获取当前活动的PPT应用程序
            Application pptApplication = Globals.ThisAddIn.Application;
            // 获取当前活动的窗口
            DocumentWindow activeWindow = pptApplication.ActiveWindow;
            // 获取当前选中的对象
            Selection selection = activeWindow.Selection;

            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                int counter = 1;
                foreach (Shape shape in selection.ShapeRange)
                {
                    RenameShape(shape, prefix, ref counter);
                }

                // 刷新视图
                activeWindow.View.GotoSlide(activeWindow.View.Slide.SlideIndex);
            }
            else
            {
                MessageBox.Show("请选择一个或多个对象。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void RenameShape(Shape shape, string prefix, ref int counter)
        {
            if (shape.Type == MsoShapeType.msoGroup)
            {
                // 如果形状是组合，则递归处理组合内的每个子形状
                foreach (Shape childShape in shape.GroupItems)
                {
                    RenameShape(childShape, prefix, ref counter);
                }
            }
            else
            {
                // 如果形状不是组合，直接重命名
                shape.Name = $"{prefix}-{counter}";
                counter++;
            }
        }

        private bool isRecording = false;
        private List<(PowerPoint.Shape Shape, PowerPoint.Shape Marker)> recordedShapes = new List<(PowerPoint.Shape, PowerPoint.Shape)>();

        private void 选择增强_Click(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            var selection = application.ActiveWindow.Selection;
            var button = sender as RibbonButton;

            if (isRecording)
            {
                // 结束记录
                isRecording = false;
                application.WindowSelectionChange -= Application_WindowSelectionChange;

                // 删除标记并选中记录的对象
                foreach (var (shape, marker) in recordedShapes)
                {
                    marker.Delete();
                }

                if (recordedShapes.Count > 0)
                {
                    var shapeNames = recordedShapes.Select(tuple => tuple.Shape.Name).ToArray();
                    application.ActiveWindow.View.Slide.Shapes.Range(shapeNames).Select();
                }

                // 恢复按钮的原始显示
                button.Label = "选择增强";
                button.Image = Properties.Resources.选择结束; // 恢复到“选择结束”图标
            }
            else
            {
                // 开始记录选择
                isRecording = true;
                recordedShapes.Clear();
                application.WindowSelectionChange += Application_WindowSelectionChange;

                // 突出显示按钮
                button.Label = "选择增强 (记录中...)";
                button.Image = Properties.Resources.选择记录中; // 更改为“选择记录中”图标
            }
        }

        private void Application_WindowSelectionChange(PowerPoint.Selection Sel)
        {
            if (isRecording && Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape shape in Sel.ShapeRange)
                {
                    if (!recordedShapes.Any(tuple => tuple.Shape.Name == shape.Name))
                    {
                        // 记录形状
                        var marker = AddCheckMark(shape);
                        recordedShapes.Add((shape, marker));
                    }
                }
            }
        }

        private PowerPoint.Shape AddCheckMark(PowerPoint.Shape shape)
        {
            var slide = shape.Parent;
            float markerSize = 10f; // 标记的大小
            float left = shape.Left + shape.Width - markerSize;
            float top = shape.Top - markerSize;

            var marker = slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval,
                left,
                top,
                markerSize,
                markerSize);

            marker.Fill.ForeColor.RGB = ToRGB(255, 0, 0); // 红色填充
            marker.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse; // 无边框

            var textFrame = marker.TextFrame;
            textFrame.TextRange.Text = "√";
            textFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
            textFrame.TextRange.Font.Size = 8; // 适当调整字符大小

            return marker;
        }

        private int ToRGB(int red, int green, int blue)
        {
            return (blue << 16) | (green << 8) | red;
        }


        private void 沿线分布_Click(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            var selection = application.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 1)
            {
                var lineShape = selection.ShapeRange[1];
                if (lineShape.Type == MsoShapeType.msoLine || lineShape.Type == MsoShapeType.msoFreeform)
                {
                    List<PowerPoint.Shape> shapesToDistribute = new List<PowerPoint.Shape>();
                    for (int i = 2; i <= selection.ShapeRange.Count; i++)
                    {
                        var shape = selection.ShapeRange[i];
                        if (shape.Type != MsoShapeType.msoLine && shape.Type != MsoShapeType.msoFreeform)
                        {
                            shapesToDistribute.Add(shape);
                        }
                    }

                    if (shapesToDistribute.Count > 0)
                    {
                        DistributeShapesAlongLine(lineShape, shapesToDistribute);
                    }
                    else
                    {
                        MessageBox.Show("没有其他对象可以分布。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("第一个选择的对象必须是线段或自由绘制的曲线。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("请至少选择一条线段或曲线和一个对象。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DistributeShapesAlongLine(PowerPoint.Shape lineShape, List<PowerPoint.Shape> shapesToDistribute)
        {
            var nodes = lineShape.Nodes;
            if (nodes.Count < 2)
            {
                MessageBox.Show("线段或曲线必须至少有两个节点。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 获取线段或曲线的所有节点坐标
            List<(float X, float Y)> nodePoints = new List<(float X, float Y)>();
            for (int i = 1; i <= nodes.Count; i++)
            {
                var point = nodes[i].Points;
                nodePoints.Add((point[1, 1], point[1, 2]));
            }

            // 计算每个对象的间距
            int count = shapesToDistribute.Count;
            float totalLength = GetTotalLength(nodePoints);
            float stepLength = totalLength / (count + 1);

            // 沿线段或曲线分布对象
            float currentLength = 0;
            for (int i = 0; i < count; i++)
            {
                currentLength += stepLength;
                var (newX, newY) = GetPointAtLength(nodePoints, currentLength);
                var shape = shapesToDistribute[i];
                shape.Left = newX - shape.Width / 2;
                shape.Top = newY - shape.Height / 2;

                // 调整对象使得曲线穿过它们的中心点
                if (lineShape.Type == MsoShapeType.msoFreeform)
                {
                    AdjustShapeToLineCenter(shape, lineShape, newX, newY);
                }
            }
        }

        private void AdjustShapeToLineCenter(PowerPoint.Shape shape, PowerPoint.Shape lineShape, float centerX, float centerY)
        {
            float shapeCenterX = shape.Left + shape.Width / 2;
            float shapeCenterY = shape.Top + shape.Height / 2;
            float offsetX = centerX - shapeCenterX;
            float offsetY = centerY - shapeCenterY;

            shape.Left += offsetX;
            shape.Top += offsetY;
        }

        private float GetTotalLength(List<(float X, float Y)> points)
        {
            float length = 0;
            for (int i = 1; i < points.Count; i++)
            {
                length += GetDistance(points[i - 1], points[i]);
            }
            return length;
        }

        private (float X, float Y) GetPointAtLength(List<(float X, float Y)> points, float targetLength)
        {
            float accumulatedLength = 0;
            for (int i = 1; i < points.Count; i++)
            {
                float segmentLength = GetDistance(points[i - 1], points[i]);
                if (accumulatedLength + segmentLength >= targetLength)
                {
                    float ratio = (targetLength - accumulatedLength) / segmentLength;
                    float newX = points[i - 1].X + ratio * (points[i].X - points[i - 1].X);
                    float newY = points[i - 1].Y + ratio * (points[i].Y - points[i - 1].Y);
                    return (newX, newY);
                }
                accumulatedLength += segmentLength;
            }
            return points.Last();
        }

        private float GetDistance((float X, float Y) point1, (float X, float Y) point2)
        {
            return (float)Math.Sqrt(Math.Pow(point2.X - point1.X, 2) + Math.Pow(point2.Y - point1.Y, 2));
        }

        private void 板贴辅助_Click(object sender, RibbonControlEventArgs e)
        {
            // 检查 Ctrl 键是否被按下
            bool isCtrlPressed = (Control.ModifierKeys & System.Windows.Forms.Keys.Control) == System.Windows.Forms.Keys.Control;

            string[] lines = null;

            if (isCtrlPressed)
            {
                // 打开文件选择对话框
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // 读取文件内容
                    lines = File.ReadAllLines(openFileDialog.FileName);
                }
            }
            else
            {
                // 创建并显示输入文本的窗口
                BoardInputTextForm inputForm = new BoardInputTextForm();
                inputForm.Text = "请输入分行文本";
                DialogResult result = inputForm.ShowDialog();

                // 如果用户点击了确定按钮
                if (result == DialogResult.OK)
                {
                    // 获取用户输入的文本
                    lines = inputForm.TextLines;
                }
            }

            // 如果获取了文本行
            if (lines != null && lines.Length > 0)
            {
                // 获取当前活动窗口
                PowerPoint.DocumentWindow activeWindow = Globals.ThisAddIn.Application.ActiveWindow;
                if (activeWindow != null)
                {
                    // 获取当前页幻灯片
                    PowerPoint.Slide currentSlide = activeWindow.View.Slide;

                    // 记录当前幻灯片的索引
                    int currentSlideIndex = currentSlide.SlideIndex;

                    // 获取当前选中的文本框
                    PowerPoint.Selection selection = activeWindow.Selection;
                    if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        PowerPoint.ShapeRange selectedShapes = selection.ShapeRange;

                        // 创建一个字典来存储相同文本内容的文本框
                        Dictionary<string, List<PowerPoint.Shape>> textBoxGroups = new Dictionary<string, List<PowerPoint.Shape>>();

                        foreach (PowerPoint.Shape shape in selectedShapes)
                        {
                            if (shape.Type == Office.MsoShapeType.msoTextBox)
                            {
                                string text = shape.TextFrame.TextRange.Text;
                                if (!textBoxGroups.ContainsKey(text))
                                {
                                    textBoxGroups[text] = new List<PowerPoint.Shape>();
                                }
                                textBoxGroups[text].Add(shape);
                            }
                        }

                        // 计算需要复制的幻灯片次数
                        int groupCount = textBoxGroups.Count;
                        int slidesNeeded = (int)Math.Ceiling((double)lines.Length / groupCount);

                        // 复制当前页幻灯片，复制次数为计算得到的slidesNeeded
                        for (int i = 0; i < slidesNeeded; i++)
                        {
                            currentSlide.Duplicate();
                        }

                        // 获取所有幻灯片
                        PowerPoint.Slides slides = activeWindow.Presentation.Slides;

                        int lineIndex = 0;
                        int slideOffset = 1; // 从下一页开始

                        // 逐页替换文本框内容
                        while (lineIndex < lines.Length && (currentSlideIndex + slideOffset) <= slides.Count)
                        {
                            PowerPoint.Slide slide = slides[currentSlideIndex + slideOffset];
                            PowerPoint.Shapes shapes = slide.Shapes;

                            foreach (var group in textBoxGroups)
                            {
                                foreach (PowerPoint.Shape shape in group.Value)
                                {
                                    if (lineIndex >= lines.Length)
                                        break;

                                    // 查找匹配的文本框并替换内容
                                    foreach (PowerPoint.Shape s in shapes)
                                    {
                                        if (s.Type == Office.MsoShapeType.msoTextBox && s.TextFrame.TextRange.Text == group.Key)
                                        {
                                            if (lineIndex < lines.Length)
                                            {
                                                s.TextFrame.TextRange.Text = lines[lineIndex];
                                            }
                                        }
                                    }
                                }
                                lineIndex++;
                            }
                            slideOffset++;
                        }
                    }
                    else
                    {
                        MessageBox.Show("请选择多个文本框来替换内容。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }


        private void 去除边距_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
            if (slide != null)
            {
                PowerPoint.ShapeRange shapeRange = app.ActiveWindow.Selection.ShapeRange;
                foreach (PowerPoint.Shape shape in shapeRange)
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        // 设置边距为0
                        shape.TextFrame.MarginLeft = 0;
                        shape.TextFrame.MarginRight = 0;
                        shape.TextFrame.MarginTop = 0;
                        shape.TextFrame.MarginBottom = 0;

                        // 调整文本框的宽度使其与文本长度相匹配
                        string text = shape.TextFrame.TextRange.Text;
                        if (!string.IsNullOrEmpty(text))
                        {
                            // 计算文本长度
                            float textLength = shape.TextFrame.TextRange.BoundWidth;
                            // 设置文本框宽度
                            shape.Width = textLength;
                        }
                    }
                }
            }
        }


        private void 单字拆分_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = app.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange shapeRange = selection.ShapeRange;
                if (shapeRange.Count == 1)
                {
                    PowerPoint.Shape shape = shapeRange[1] as PowerPoint.Shape;
                    if (shape != null && shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        string text = shape.TextFrame.TextRange.Text;
                        float left = shape.Left;
                        float top = shape.Top;
                        float width = shape.Width / text.Length;
                        float height = shape.Height;

                        for (int i = 0; i < text.Length; i++)
                        {
                            PowerPoint.Shape newShape = shape.Duplicate()[1];
                            newShape.Left = left + i * width;
                            newShape.Top = top;
                            newShape.Width = width;
                            newShape.Height = height;
                            newShape.TextFrame.TextRange.Text = text[i].ToString();
                        }

                        // 删除原有的文本框
                        shape.Delete();
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("请选择一个包含文本的文本框。");
                    }
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("请选择一个文本框。");
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("请选择一个文本框。");
            }
        }

        private void 拆分段落_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = app.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange shapeRange = selection.ShapeRange;
                if (shapeRange.Count == 1)
                {
                    PowerPoint.Shape shape = shapeRange[1] as PowerPoint.Shape;
                    if (shape != null && shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        PowerPoint.TextRange textRange = shape.TextFrame.TextRange;
                        int paragraphCount = textRange.Paragraphs().Count;
                        float left = shape.Left;
                        float top = shape.Top;
                        float width = shape.Width;

                        // 计算有效段落的数量和高度
                        int validParagraphCount = 0;
                        for (int i = 1; i <= paragraphCount; i++)
                        {
                            if (!string.IsNullOrWhiteSpace(textRange.Paragraphs(i).Text.Trim()))
                            {
                                validParagraphCount++;
                            }
                        }
                        float height = shape.Height / validParagraphCount;

                        int validIndex = 0;
                        for (int i = 1; i <= paragraphCount; i++)
                        {
                            PowerPoint.TextRange paragraph = textRange.Paragraphs(i);
                            string trimmedText = paragraph.Text.Trim();
                            if (!string.IsNullOrWhiteSpace(trimmedText))
                            {
                                validIndex++;
                                PowerPoint.Shape newShape = shape.Duplicate()[1];
                                newShape.Left = left;
                                newShape.Top = top + (validIndex - 1) * height;
                                newShape.Width = width;
                                newShape.Height = height;
                                newShape.TextFrame.TextRange.Text = trimmedText;
                            }
                        }

                        // 删除原有的文本框
                        shape.Delete();
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("请选择一个包含多个段落的文本框。");
                    }
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("请选择一个文本框。");
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("请选择一个文本框。");
            }
        }


        private Dictionary<string, float> originalIndents = new Dictionary<string, float>();
        private bool isAdjusted = false;
        private void 首行缩进_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Presentation presentation = app.ActivePresentation;

            string adjustRulerVbaCode = @"
Sub AdjustRuler()
    Dim shp As Shape
    Dim tr As TextRange2
    Dim para As ParagraphFormat2
    Dim fontSize As Single
    Dim baseFontSize As Single
    Dim baseIndent As Single
    Dim indentSize As Single

    baseFontSize = 18
    baseIndent = 1.27 * 28.3465

    For Each shp In Application.ActiveWindow.Selection.ShapeRange
        If shp.HasTextFrame Then
            Set tr = shp.TextFrame2.TextRange

            For i = 1 To tr.Paragraphs.Count
                Set para = tr.ParagraphFormat
                fontSize = tr.Paragraphs(i).Font.Size
                indentSize = baseIndent * (fontSize / baseFontSize)
                para.LeftIndent = 0
                para.FirstLineIndent = indentSize
            Next i
        End If
    Next shp
End Sub
";

            string restoreRulerVbaCode = @"
Sub RestoreRuler()
    Dim shp As Shape
    Dim tr As TextRange2
    Dim para As ParagraphFormat2

    For Each shp In Application.ActiveWindow.Selection.ShapeRange
        If shp.HasTextFrame Then
            Set tr = shp.TextFrame2.TextRange

            For i = 1 To tr.Paragraphs.Count
                Set para = tr.ParagraphFormat
                para.LeftIndent = 0
                para.FirstLineIndent = 0
            Next i
        End If
    Next shp
End Sub
";

            string moduleName = "DynamicModule";

            if (!isAdjusted)
            {
                InsertVbaCode(presentation, adjustRulerVbaCode, moduleName);

                try
                {
                    app.Run("AdjustRuler");
                    isAdjusted = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error running macro: " + ex.Message);
                }
            }
            else
            {
                InsertVbaCode(presentation, restoreRulerVbaCode, moduleName);

                try
                {
                    app.Run("RestoreRuler");
                    isAdjusted = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error running macro: " + ex.Message);
                }
            }

            DeleteVbaModule(presentation, moduleName);
        }

        private void InsertVbaCode(PowerPoint.Presentation presentation, string code, string moduleName)
        {
            var vbaProject = presentation.VBProject;
            var vbaModule = vbaProject.VBComponents.Add(Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule);
            vbaModule.Name = moduleName;
            vbaModule.CodeModule.AddFromString(code);
        }

        private void DeleteVbaModule(PowerPoint.Presentation presentation, string moduleName)
        {
            var vbaProject = presentation.VBProject;
            var vbaModule = vbaProject.VBComponents.Item(moduleName);
            vbaProject.VBComponents.Remove(vbaModule);
        }

        private void 音频导出_Click(object sender, RibbonControlEventArgs e)
        {
           
        }

        private void 动画触发_Click(object sender, RibbonControlEventArgs e)
        {
           
        }


        private List<PowerPoint.Shape> copiedShapes = new List<PowerPoint.Shape>();
        private Dictionary<int, (float Width, float Height)> initialSizes = new Dictionary<int, (float Width, float Height)>();

        private void 环形分布_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = pptApp.ActiveWindow.Selection;
            PowerPoint.ShapeRange selectedShapes = selection.ShapeRange;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes && selectedShapes.Count >= 1)
            {
                float radius = 100;
                float initialRotation = 0;
                float finalRotation = 0;
                float sizeIncrement = 0;
                int copyCount = 0;

                if (selectedShapes.Count == 1)
                {
                    ShowSingleObjectForm(pptApp, selectedShapes, radius, initialRotation, finalRotation, sizeIncrement, copyCount);
                }
                else
                {
                    PerformCircularDistribution(pptApp, selectedShapes, radius, initialRotation, finalRotation, sizeIncrement, false);
                    ShowMultipleObjectsForm(pptApp, selectedShapes, radius, initialRotation, finalRotation, sizeIncrement);
                }
            }
            else
            {
                MessageBox.Show("请选择至少一个对象。");
            }
        }

        private void PerformCircularDistribution(PowerPoint.Application pptApp, PowerPoint.ShapeRange shapes, float radius, float initialRotation, float finalRotation, float sizeIncrement, bool isCopyMode, int copyCount = 0)
        {
            if (isCopyMode)
            {
                foreach (PowerPoint.Shape shape in copiedShapes)
                {
                    shape.Delete();
                }
                copiedShapes.Clear();
            }

            int count = isCopyMode ? copyCount + 1 : shapes.Count; // 确保选中的对象也被包含
            float angleStep = 360.0f / count;
            float angleIncrement = (finalRotation - initialRotation) / count;

            float currentRadius = radius;

            for (int i = 0; i < count; i++)
            {
                float angle = initialRotation + i * angleStep;
                float radians = angle * (float)(Math.PI / 180.0);
                float newX = (float)(currentRadius * Math.Cos(radians));
                float newY = (float)(currentRadius * Math.Sin(radians));

                PowerPoint.Shape shape;
                if (isCopyMode)
                {
                    if (i == 0)
                    {
                        shape = shapes[1]; // 第一个形状是选中的对象
                    }
                    else
                    {
                        shape = shapes[1].Duplicate()[1];
                        copiedShapes.Add(shape);
                    }
                }
                else
                {
                    shape = shapes[i + 1];
                }

                shape.Left = newX + (pptApp.ActivePresentation.PageSetup.SlideWidth / 2) - (shape.Width / 2);
                shape.Top = newY + (pptApp.ActivePresentation.PageSetup.SlideHeight / 2) - (shape.Height / 2);
                shape.Rotation = initialRotation + i * angleIncrement;

                if (!initialSizes.ContainsKey(shape.Id))
                {
                    initialSizes[shape.Id] = (shape.Width, shape.Height);
                }

                if (sizeIncrement != 0)
                {
                    float newSize = initialSizes[shape.Id].Width * (1 + i * sizeIncrement / 100.0f);
                    shape.Width = newSize;
                    shape.Height = newSize;

                    // 增加当前半径以保持间距相等
                    currentRadius += sizeIncrement / 2.0f;
                }
            }
        }

        private void ShowSingleObjectForm(PowerPoint.Application pptApp, PowerPoint.ShapeRange shapes, float radius, float initialRotation, float finalRotation, float sizeIncrement, int copyCount)
        {
            SingleObjectForm form = new SingleObjectForm(pptApp, shapes, radius, initialRotation, finalRotation, sizeIncrement, copyCount);
            form.ShowDialog();
        }

        private void ShowMultipleObjectsForm(PowerPoint.Application pptApp, PowerPoint.ShapeRange shapes, float radius, float initialRotation, float finalRotation, float sizeIncrement)
        {
            MultipleObjectsForm form = new MultipleObjectsForm(pptApp, shapes, radius, initialRotation, finalRotation, sizeIncrement);
            form.ShowDialog();
        }



        private PowerPoint.ShapeRange selectedShapes;
        private MatrixDistributionForm matrixForm;

        // 保持原始尺寸和当前缩放比例
        private float[] originalWidths;
        private float[] originalHeights;
        private float currentScale = 100.0f;

        private void 矩阵分布_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取选中的对象
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

            // 检查选择是否有效
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 0)
            {
                selectedShapes = selection.ShapeRange;

                // 保存原始尺寸
                originalWidths = new float[selectedShapes.Count];
                originalHeights = new float[selectedShapes.Count];
                for (int i = 0; i < selectedShapes.Count; i++)
                {
                    originalWidths[i] = selectedShapes[i + 1].Width;
                    originalHeights[i] = selectedShapes[i + 1].Height;
                }

                // 初始化当前缩放比例
                currentScale = 100.0f;

                // 显示矩阵分布设置窗体
                if (matrixForm == null || matrixForm.IsDisposed)
                {
                    matrixForm = new MatrixDistributionForm();
                    matrixForm.ParametersChanged += Form_ParametersChanged;
                    matrixForm.FormClosed += Form_FormClosed;
                }

                if (selectedShapes.Count > 1)
                {
                    matrixForm.SetTotalCount(selectedShapes.Count);
                }
                else
                {
                    matrixForm.EnableTotalCountAdjustment();
                }

                matrixForm.Show();
                matrixForm.TopMost = true;  // 设置为顶层窗体
            }
            else
            {
                MessageBox.Show("请选择一个或多个对象");
            }
        }

        private void Form_ParametersChanged(object sender, EventArgs e)
        {
            var form = sender as MatrixDistributionForm;
            int totalCount = form.TotalCount;
            int horizontalCount = form.HorizontalCount;
            int rowSpacing = form.RowSpacing;
            int columnSpacing = form.ColumnSpacing;
            int scale = form.Scale;

            if (selectedShapes != null && selectedShapes.Count > 0)
            {
                var slide = selectedShapes[1].Parent;

                // 删除现有的复制对象
                DeleteExistingCopies(slide);

                if (selectedShapes.Count > 1)
                {
                    // 对多个对象进行排列
                    ArrangeShapes(selectedShapes, horizontalCount, rowSpacing, columnSpacing, scale);
                }
                else
                {
                    // 对单个对象进行复制和排列
                    var baseShape = selectedShapes[1];
                    CreateMatrix(baseShape, totalCount, horizontalCount, rowSpacing, columnSpacing, scale);
                }
            }
        }

        private void Form_FormClosed(object sender, FormClosedEventArgs e)
        {
            var form = sender as MatrixDistributionForm;
            form.ParametersChanged -= Form_ParametersChanged;
        }

        private void ArrangeShapes(PowerPoint.ShapeRange shapes, int horizontalCount, int rowSpacing, int columnSpacing, int scale)
        {
            // 计算每个形状的初始位置
            float initialLeft = shapes[1].Left;
            float initialTop = shapes[1].Top;

            // 计算缩放比例的变化
            float scaleFactor = scale / currentScale;

            // 更新 currentScale
            currentScale = scale;

            // 排列选中的对象
            for (int i = 0; i < shapes.Count; i++)
            {
                int row = i / horizontalCount;
                int column = i % horizontalCount;

                float left = initialLeft + column * (originalWidths[i] * scaleFactor + columnSpacing);
                float top = initialTop + row * (originalHeights[i] * scaleFactor + rowSpacing);

                var shape = shapes[i + 1];
                shape.Left = left;
                shape.Top = top;
                shape.Width *= scaleFactor; // 基于当前尺寸计算新的宽度
                shape.Height *= scaleFactor; // 基于当前尺寸计算新的高度

                // 更新原始尺寸为当前尺寸
                originalWidths[i] = shape.Width;
                originalHeights[i] = shape.Height;
            }
        }

        private void CreateMatrix(PowerPoint.Shape baseShape, int totalCount, int horizontalCount, int rowSpacing, int columnSpacing, int scale)
        {
            // 计算每个形状的初始位置
            float initialLeft = baseShape.Left;
            float initialTop = baseShape.Top;

            // 保存原始尺寸
            float originalWidth = baseShape.Width;
            float originalHeight = baseShape.Height;

            // 计算缩放比例的变化
            float scaleFactor = scale / currentScale;

            // 更新 currentScale
            currentScale = scale;

            // 创建矩阵
            for (int i = 0; i < totalCount; i++)
            {
                int row = i / horizontalCount;
                int column = i % horizontalCount;

                float left = initialLeft + column * (originalWidth * scaleFactor + columnSpacing);
                float top = initialTop + row * (originalHeight * scaleFactor + rowSpacing);

                // 只有在i大于0时才复制原始对象
                if (i > 0)
                {
                    var newShape = baseShape.Duplicate();
                    newShape.Left = left;
                    newShape.Top = top;
                    newShape.Width *= scaleFactor; // 基于当前尺寸计算新的宽度
                    newShape.Height *= scaleFactor; // 基于当前尺寸计算新的高度
                    newShape.Name = "Copy_of_" + baseShape.Name + "_" + i;
                }
                else
                {
                    baseShape.Left = left;
                    baseShape.Top = top;
                    baseShape.Width *= scaleFactor; // 基于当前尺寸计算新的宽度
                    baseShape.Height *= scaleFactor; // 基于当前尺寸计算新的高度

                    // 更新原始尺寸为当前尺寸
                    originalWidths[0] = baseShape.Width;
                    originalHeights[0] = baseShape.Height;
                }
            }
        }

        private void DeleteExistingCopies(PowerPoint.Slide slide)
        {
            for (int i = slide.Shapes.Count; i >= 1; i--)
            {
                var shape = slide.Shapes[i];
                if (shape.Name.StartsWith("Copy_of_"))
                {
                    shape.Delete();
                }
            }
        }

        private void Experte抠图_Click(object sender, RibbonControlEventArgs e)
        {
            string url = "https://quzuotu.com/home";
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法打开链接: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private SpecifyalignmentForm specifyalignmentForm;
        private void 指定对齐_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;

            if (specifyalignmentForm == null || specifyalignmentForm.IsDisposed)
            {
                specifyalignmentForm = new SpecifyalignmentForm(app);
            }

            specifyalignmentForm.Show();
            specifyalignmentForm.BringToFront();
        }


        private TimerForm timerForm;
        private void Timer_Click(object sender, RibbonControlEventArgs e)
        {
            ShowTimer();
        }

        private void ShowTimer()
        {
            if (timerForm == null || timerForm.IsDisposed)
            {
                timerForm = new TimerForm();
                timerForm.Show();
            }
            else
            {
                timerForm.BringToFront();
            }
        }

        private void 原位转图_Click(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            var presentation = application.ActivePresentation;
            var slide = application.ActiveWindow.View.Slide;

            if (application.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionShapes)
            {
                var selectedShapes = application.ActiveWindow.Selection.ShapeRange;
                foreach (Shape shape in selectedShapes)
                {
                    // 复制选定的形状
                    shape.Copy();

                    // 粘贴为图片
                    var pictureShape = slide.Shapes.PasteSpecial(PpPasteDataType.ppPastePNG)[1];

                    // 获取原始位置和大小
                    float left = shape.Left;
                    float top = shape.Top;
                    float width = shape.Width;
                    float height = shape.Height;

                    // 设置图片的位置和大小
                    pictureShape.Left = left;
                    pictureShape.Top = top;
                    pictureShape.Width = width;
                    pictureShape.Height = height;

                    // 删除原来的形状
                    shape.Delete();
                }
            }
            else
            {
                MessageBox.Show("请先选择一个或多个对象。");
            }
        }


        private void splitButton2_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void 位图转矢量图_Click(object sender, RibbonControlEventArgs e)
        {
            string url = "https://svg.tmttool.com/";
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法打开链接: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Bgsub_Click(object sender, RibbonControlEventArgs e)
        {
            string url = "https://bgsub.cn/webapp/";
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法打开链接: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //原位复制一个对象
        private void LCopy_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = pptApp.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    PowerPoint.Shape newShape = shape.Duplicate()[1]; // 复制形状
                    newShape.Left = shape.Left; // 保持原位
                    newShape.Top = shape.Top;   // 保持原位
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("请选择一个或多个对象进行复制。", "提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
        }


       
      

        private void 生成样机_Click(object sender, RibbonControlEventArgs e)
        {
            using (var form = new SampleGenerationForm())
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    var exportSelectedSlides = form.ExportSelectedSlides;
                    var exportAllSlides = form.ExportAllSlides;
                    var selectedSampleStyle = form.SelectedSampleStyle;
                    var selectedResolution = form.SelectedResolution;

                    var pptApp = Globals.ThisAddIn.Application;
                    var presentation = pptApp.ActivePresentation;

                    var tempFolder = Path.Combine(Path.GetTempPath(), "PPTImages");
                    if (!Directory.Exists(tempFolder))
                    {
                        Directory.CreateDirectory(tempFolder);
                    }

                    try
                    {
                        SlideRange selectedSlides = null;
                        if (exportSelectedSlides)
                        {
                            selectedSlides = pptApp.ActiveWindow.Selection.SlideRange;
                        }

                        if (exportAllSlides || (selectedSlides != null && selectedSlides.Count > 0))
                        {
                            int slideIndex = 1;
                            var resolution = GetResolution(selectedResolution);
                            if (exportAllSlides)
                            {
                                foreach (Slide slide in presentation.Slides)
                                {
                                    string imagePath = Path.Combine(tempFolder, $"样机填充-{slideIndex}.png");
                                    slide.Export(imagePath, "PNG", resolution.Width, resolution.Height);
                                    slideIndex++;
                                }
                            }
                            else if (selectedSlides != null)
                            {
                                foreach (Slide slide in selectedSlides)
                                {
                                    string imagePath = Path.Combine(tempFolder, $"样机填充-{slideIndex}.png");
                                    slide.Export(imagePath, "PNG", resolution.Width, resolution.Height);
                                    slideIndex++;
                                }
                            }

                            string samplePath = GetSamplePath(selectedSampleStyle, tempFolder);
                            var samplePresentation = pptApp.Presentations.Open(samplePath);

                            foreach (Slide slide in samplePresentation.Slides)
                            {
                                foreach (Shape shape in slide.Shapes)
                                {
                                    if (shape.Type == MsoShapeType.msoGroup)
                                    {
                                        FillGroupShapes(shape.GroupItems, tempFolder);
                                    }
                                    else if (shape.Name.StartsWith("样机填充-"))
                                    {
                                        string shapeIndex = shape.Name.Substring(5); // 获取形状的索引
                                        string imagePath = Path.Combine(tempFolder, $"样机填充-{shapeIndex}.png");
                                        if (File.Exists(imagePath))
                                        {
                                            shape.Fill.UserPicture(imagePath);
                                        }
                                    }
                                }
                            }

                            string savePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "生成的样机展示.pptx");
                            samplePresentation.SaveAs(savePath);
                            samplePresentation.Close();

                            MessageBox.Show($"样机展示已生成并保存在：{savePath}", "生成成功", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            // 打开生成的样机展示文档
                            pptApp.Presentations.Open(savePath);
                        }
                        else
                        {
                            MessageBox.Show("没有选择任何幻灯片", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"生成样机展示时发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        if (Directory.Exists(tempFolder))
                        {
                            Directory.Delete(tempFolder, true);
                        }
                    }
                }
            }
        }

        private (int Width, int Height) GetResolution(string selectedResolution)
        {
            switch (selectedResolution)
            {
                case "720x480 (标清)":
                    return (720, 480);
                case "1280x720 (高清)":
                    return (1280, 720);
                case "1920x1080 (全高清)":
                    return (1920, 1080);
                case "2048x1080 (2K)":
                    return (2048, 1080);
                case "3840x2160 (超高清)":
                    return (3840, 2160);
                case "4096x2160 (4K)":
                    return (4096, 2160);
                case "7680x4320 (8K)":
                    return (7680, 4320);
                default:
                    return (1920, 1080); // 默认分辨率
            }
        }

        private string GetSamplePath(int selectedSampleStyle, string tempFolder)
        {
            string samplePath = string.Empty;
            switch (selectedSampleStyle)
            {
                case 1:
                    samplePath = Path.Combine(tempFolder, "样机1.pptx");
                    File.WriteAllBytes(samplePath, Properties.Resources.样机样式1);
                    break;
                case 2:
                    samplePath = Path.Combine(tempFolder, "样机2.pptx");
                    File.WriteAllBytes(samplePath, Properties.Resources.样机样式2);
                    break;
                case 3:
                    samplePath = Path.Combine(tempFolder, "样机3.pptx");
                    File.WriteAllBytes(samplePath, Properties.Resources.样机样式3);
                    break;
                case 4:
                    samplePath = Path.Combine(tempFolder, "样机4.pptx");
                    File.WriteAllBytes(samplePath, Properties.Resources.样机样式4);
                    break;
                case 5:
                    samplePath = Path.Combine(tempFolder, "样机5.pptx");
                    File.WriteAllBytes(samplePath, Properties.Resources.样机样式5);
                    break;
                case 6:
                    samplePath = Path.Combine(tempFolder, "样机6.pptx");
                    File.WriteAllBytes(samplePath, Properties.Resources.样机样式6);
                    break;
            }

            return samplePath;
        }

        private void FillGroupShapes(Microsoft.Office.Interop.PowerPoint.GroupShapes groupShapes, string tempFolder)
        {
            foreach (Shape shape in groupShapes)
            {
                if (shape.Type == MsoShapeType.msoGroup)
                {
                    FillGroupShapes(shape.GroupItems, tempFolder);
                }
                else if (shape.Name.StartsWith("样机填充-"))
                {
                    string shapeIndex = shape.Name.Substring(5); // 获取形状的索引
                    string imagePath = Path.Combine(tempFolder, $"样机填充-{shapeIndex}.png");
                    if (File.Exists(imagePath))
                    {
                        shape.Fill.UserPicture(imagePath);
                        shape.Fill.RotateWithObject = MsoTriState.msoFalse; // 确保取消勾选“与形状一起旋转”
                    }
                }
            }
        }
 
      
        private void 图形修剪_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                TrimShapes();
            }
            catch (Exception ex)
            {
                MessageBox.Show("发生错误：" + ex.Message);
            }
        }

        private void TrimShapes()
        {
            Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            float slideWidth = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth;
            float slideHeight = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight;
            string command = "ShapesIntersect";

            if (sel.Type == PpSelectionType.ppSelectionShapes || sel.Type == PpSelectionType.ppSelectionText)
            {
                Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
                ProcessShapeRange((Microsoft.Office.Interop.PowerPoint.ShapeRange)sel.ShapeRange, slide, slideWidth, slideHeight, command);
            }
            else if (sel.Type == PpSelectionType.ppSelectionSlides)
            {
                foreach (Slide slide in sel.SlideRange)
                {
                    ProcessShapes(slide.Shapes, slide, slideWidth, slideHeight, command);
                }
            }
            else
            {
                MessageBox.Show("请选中图形");
            }
        }

        private void ProcessShapeRange(Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange, Slide slide, float slideWidth, float slideHeight, string command)
        {
            List<Microsoft.Office.Interop.PowerPoint.Shape> shapesToProcess = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in shapeRange)
            {
                if (shape.Type == MsoShapeType.msoAutoShape || shape.Type == MsoShapeType.msoFreeform || shape.Type == MsoShapeType.msoTextBox || shape.Type == MsoShapeType.msoPicture)
                {
                    shapesToProcess.Add(shape);
                }
            }

            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in shapesToProcess)
            {
                if (IsShapeOutOfBounds(shape, slideWidth, slideHeight))
                {
                    Microsoft.Office.Interop.PowerPoint.Shape clippingRectangle = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0f, 0f, slideWidth, slideHeight);
                    try
                    {
                        shape.Select(MsoTriState.msoTrue);
                        clippingRectangle.Select(MsoTriState.msoFalse);
                        Globals.ThisAddIn.Application.CommandBars.ExecuteMso(command);
                    }
                    catch (Exception ex)
                    {
                        // 忽略特定的异常
                        Console.WriteLine("忽略的错误：" + ex.Message);
                    }
                    finally
                    {
                        // 确保删除辅助矩形
                        try
                        {
                            clippingRectangle.Delete();
                        }
                        catch (Exception deleteEx)
                        {
                            Console.WriteLine("删除辅助矩形时忽略的错误：" + deleteEx.Message);
                        }
                    }
                }
            }
        }

        private void ProcessShapes(Microsoft.Office.Interop.PowerPoint.Shapes shapes, Slide slide, float slideWidth, float slideHeight, string command)
        {
            List<Microsoft.Office.Interop.PowerPoint.Shape> shapesToProcess = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in shapes)
            {
                if (shape.Type == MsoShapeType.msoAutoShape || shape.Type == MsoShapeType.msoFreeform || shape.Type == MsoShapeType.msoTextBox || shape.Type == MsoShapeType.msoPicture)
                {
                    shapesToProcess.Add(shape);
                }
            }

            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in shapesToProcess)
            {
                if (IsShapeOutOfBounds(shape, slideWidth, slideHeight))
                {
                    Microsoft.Office.Interop.PowerPoint.Shape clippingRectangle = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0f, 0f, slideWidth, slideHeight);
                    try
                    {
                        shape.Select(MsoTriState.msoTrue);
                        clippingRectangle.Select(MsoTriState.msoFalse);
                        Globals.ThisAddIn.Application.CommandBars.ExecuteMso(command);
                    }
                    catch (Exception ex)
                    {
                        // 忽略特定的异常
                        Console.WriteLine("忽略的错误：" + ex.Message);
                    }
                    finally
                    {
                        // 确保删除辅助矩形
                        try
                        {
                            clippingRectangle.Delete();
                        }
                        catch (Exception deleteEx)
                        {
                            Console.WriteLine("删除辅助矩形时忽略的错误：" + deleteEx.Message);
                        }
                    }
                }
            }
        }

        private bool IsShapeOutOfBounds(Microsoft.Office.Interop.PowerPoint.Shape shape, float slideWidth, float slideHeight)
        {
            return shape.Left < 0f || shape.Left + shape.Width > slideWidth || shape.Top < 0f || shape.Top + shape.Height > slideHeight;
        }

       

        private void 统一大小_Click(object sender, RibbonControlEventArgs e)
        {
            Application app = Globals.ThisAddIn.Application;
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("请至少选择2个形状或1个文本框");
                return;
            }

            Microsoft.Office.Interop.PowerPoint.ShapeRange rng = sel.ShapeRange;
            if (sel.HasChildShapeRange)
            {
                rng = sel.ChildShapeRange;
            }

            Shape shp = rng[1];
            int count = rng.Count;

            if (count == 1)
            {
                float pw = app.ActivePresentation.PageSetup.SlideWidth;
                float ph = app.ActivePresentation.PageSetup.SlideHeight;
                float pn = pw / ph;

                if (shp.Type == MsoShapeType.msoPicture)
                {
                    shp.ScaleWidth(1f, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromMiddle);
                    shp.ScaleHeight(1f, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromMiddle);
                    shp.PictureFormat.CropLeft = 0f;
                    shp.PictureFormat.CropRight = 0f;
                    shp.PictureFormat.CropTop = 0f;
                    shp.PictureFormat.CropBottom = 0f;
                    if (shp.Width >= shp.Height)
                    {
                        if (shp.Width - shp.Height * pn >= 0f)
                        {
                            float n2 = (shp.Width - shp.Height * pn) / 2f;
                            shp.PictureFormat.CropLeft = n2;
                            shp.PictureFormat.CropRight = n2;
                        }
                        else
                        {
                            float n2 = (shp.Height - shp.Width / pn) / 2f;
                            shp.PictureFormat.CropTop = n2;
                            shp.PictureFormat.CropBottom = n2;
                        }
                    }
                    else if (shp.Height - shp.Width / pn >= 0f)
                    {
                        float n2 = (shp.Height - shp.Width / pn) / 2f;
                        shp.PictureFormat.CropTop = n2;
                        shp.PictureFormat.CropBottom = n2;
                    }
                    else
                    {
                        float n2 = (shp.Width - shp.Height * pn) / 2f;
                        shp.PictureFormat.CropLeft = n2;
                        shp.PictureFormat.CropRight = n2;
                    }
                }
                shp.Width = pw;
                shp.Height = ph;
                shp.Left = pw / 2f - shp.Width / 2f;
                shp.Top = ph / 2f - shp.Height / 2f;
                return;
            }

            for (int i = 2; i <= count; i++)
            {
                if (rng[i].Type == MsoShapeType.msoPicture)
                {
                    float lm4 = rng[i].Left + rng[i].Width / 2f;
                    float tm4 = rng[i].Top + rng[i].Height / 2f;
                    rng[i].PictureFormat.Crop.ShapeHeight = rng[i].PictureFormat.Crop.PictureHeight;
                    rng[i].PictureFormat.Crop.ShapeWidth = rng[i].PictureFormat.Crop.PictureWidth;
                    rng[i].PictureFormat.Crop.PictureOffsetX = (rng[i].PictureFormat.Crop.ShapeWidth - rng[i].PictureFormat.Crop.PictureWidth) / 1024f;
                    rng[i].PictureFormat.Crop.PictureOffsetY = (rng[i].PictureFormat.Crop.ShapeHeight - rng[i].PictureFormat.Crop.PictureHeight) / 1024f;
                    rng[i].Top = tm4 - rng[i].Height / 2f;
                    rng[i].Left = lm4 - rng[i].Width / 2f;
                    rng[i].Width = shp.Width;
                    rng[i].Left = lm4 - rng[i].Width / 2f;
                    if (rng[i].Height > shp.Height)
                    {
                        rng[i].PictureFormat.Crop.ShapeHeight = shp.Height;
                        rng[i].Top = tm4 - rng[i].Height / 2f;
                        rng[i].PictureFormat.Crop.PictureOffsetY = (rng[i].PictureFormat.Crop.ShapeHeight - rng[i].PictureFormat.Crop.PictureHeight) / 2f / 1024f;
                    }
                    else if (rng[i].Height < shp.Height)
                    {
                        rng[i].Height = shp.Height;
                        rng[i].Top = tm4 - rng[i].Height / 2f;
                        rng[i].PictureFormat.Crop.ShapeWidth = rng[1].Width;
                        rng[i].Left = lm4 - rng[i].Width / 2f;
                        rng[i].PictureFormat.Crop.PictureOffsetX = (rng[i].PictureFormat.Crop.PictureWidth - rng[i].Width) / 2f / 1024f;
                    }
                    else
                    {
                        rng[i].Top = tm4 - rng[i].Height / 2f;
                    }
                }
                else
                {
                    float tm5 = rng[i].Top + rng[i].Height / 2f;
                    float lm5 = rng[i].Left + rng[i].Width / 2f;
                    rng[i].Height = shp.Height;
                    rng[i].Width = shp.Width;
                    rng[i].Left = lm5 - rng[i].Width / 2f;
                    rng[i].Top = tm5 - rng[i].Height / 2f;
                }
            }
        }

        private void 统一格式_Click(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            var presentation = application.ActivePresentation;
            var slide = application.ActiveWindow.View.Slide;

            if (application.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionShapes)
            {
                var selectedShapes = application.ActiveWindow.Selection.ShapeRange;

                if (application.ActiveWindow.Selection.HasChildShapeRange)
                {
                    selectedShapes = application.ActiveWindow.Selection.ChildShapeRange;
                }

                if (selectedShapes.Count > 1)
                {
                    Shape baseShape = selectedShapes[1];

                    // 遍历后续被选中的对象
                    ApplyFormatToShapes(baseShape, selectedShapes);
                }
                else
                {
                    MessageBox.Show("请至少选择两个对象。");
                }
            }
            else
            {
                MessageBox.Show("请先选择一个或多个对象。");
            }
        }

        private void ApplyFormatToShapes(Shape baseShape, Microsoft.Office.Interop.PowerPoint.ShapeRange selectedShapes)
        {
            for (int i = 2; i <= selectedShapes.Count; i++)
            {
                Shape shape = selectedShapes[i];

                if (shape.Type == MsoShapeType.msoGroup)
                {
                    // 对于组合形状，递归处理其子形状
                    ApplyFormatToShapes(baseShape, shape.GroupItems);
                }
                else
                {
                    try
                    {
                        // 使用格式刷功能复制格式
                        baseShape.PickUp();
                        shape.Apply();
                    }
                    catch (Exception ex)
                    {
                        // 忽略不支持的格式，并记录异常
                        Console.WriteLine($"应用格式时出错: {ex.Message}");
                    }
                }
            }
        }

        private void ApplyFormatToShapes(Shape baseShape, Microsoft.Office.Interop.PowerPoint.GroupShapes groupShapes)
        {
            for (int i = 1; i <= groupShapes.Count; i++)
            {
                Shape shape = groupShapes[i];

                if (shape.Type == MsoShapeType.msoGroup)
                {
                    // 对于组合形状，递归处理其子形状
                    ApplyFormatToShapes(baseShape, shape.GroupItems);
                }
                else
                {
                    try
                    {
                        // 使用格式刷功能复制格式
                        baseShape.PickUp();
                        shape.Apply();
                    }
                    catch (Exception ex)
                    {
                        // 忽略不支持的格式，并记录异常
                        Console.WriteLine($"应用格式时出错: {ex.Message}");
                    }
                }
            }
        }

        private void 交换位置_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 获取当前活动的PowerPoint应用程序
                var application = Globals.ThisAddIn.Application;

                // 获取当前选中的对象
                var selection = application.ActiveWindow.Selection;

                // 确保选中了两个对象
                if (selection.Type == PpSelectionType.ppSelectionShapes)
                {
                    var selectedShapes = selection.ShapeRange;

                    if (selection.HasChildShapeRange)
                    {
                        selectedShapes = selection.ChildShapeRange;
                    }

                    if (selectedShapes.Count == 2)
                    {
                        // 获取两个选中的形状
                        Shape shape1 = selectedShapes[1];
                        Shape shape2 = selectedShapes[2];

                        // 记录这两个形状的位置
                        float shape1Left = shape1.Left;
                        float shape1Top = shape1.Top;
                        float shape2Left = shape2.Left;
                        float shape2Top = shape2.Top;

                        // 交换位置
                        shape1.Left = shape2Left;
                        shape1.Top = shape2Top;
                        shape2.Left = shape1Left;
                        shape2.Top = shape1Top;
                    }
                    else
                    {
                        MessageBox.Show("请选中两个对象以交换它们的位置。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("请选中两个对象以交换它们的位置。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("交换位置时发生错误: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void 交换尺寸_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 获取当前活动的PowerPoint应用程序
                var application = Globals.ThisAddIn.Application;

                // 获取当前选中的对象
                var selection = application.ActiveWindow.Selection;

                // 确保选中了两个对象
                if (selection.Type == PpSelectionType.ppSelectionShapes)
                {
                    var selectedShapes = selection.ShapeRange;

                    if (selection.HasChildShapeRange)
                    {
                        selectedShapes = selection.ChildShapeRange;
                    }

                    if (selectedShapes.Count == 2)
                    {
                        // 获取两个选中的形状
                        Shape shape1 = selectedShapes[1];
                        Shape shape2 = selectedShapes[2];

                        // 记录这两个形状的原始尺寸
                        float shape1OriginalWidth = shape1.Width;
                        float shape1OriginalHeight = shape1.Height;
                        float shape2OriginalWidth = shape2.Width;
                        float shape2OriginalHeight = shape2.Height;

                        // 计算这两个形状的比例
                        float shape1AspectRatio = shape1OriginalWidth / shape1OriginalHeight;
                        float shape2AspectRatio = shape2OriginalWidth / shape2OriginalHeight;

                        // 交换彼此的比例进行裁剪和缩放
                        if (shape1.Type == MsoShapeType.msoPicture || shape1.Type == MsoShapeType.msoLinkedPicture)
                        {
                            ResizeAndCropShape(shape1, shape2AspectRatio, shape2OriginalWidth, shape2OriginalHeight);
                        }
                        else
                        {
                            shape1.Width = shape2OriginalWidth;
                            shape1.Height = shape2OriginalHeight;
                        }

                        if (shape2.Type == MsoShapeType.msoPicture || shape2.Type == MsoShapeType.msoLinkedPicture)
                        {
                            ResizeAndCropShape(shape2, shape1AspectRatio, shape1OriginalWidth, shape1OriginalHeight);
                        }
                        else
                        {
                            shape2.Width = shape1OriginalWidth;
                            shape2.Height = shape1OriginalHeight;
                        }
                    }
                    else
                    {
                        MessageBox.Show("请选中两个对象以交换它们的尺寸。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("请选中两个对象以交换它们的尺寸。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("交换尺寸时发生错误: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ResizeAndCropShape(Shape shape, float targetAspectRatio, float targetWidth, float targetHeight)
        {
            shape.LockAspectRatio = MsoTriState.msoFalse;

            float originalWidth = shape.Width;
            float originalHeight = shape.Height;
            float originalAspectRatio = originalWidth / originalHeight;

            if (originalAspectRatio > targetAspectRatio)
            {
                // 当前形状比例宽于目标比例，裁剪宽度
                float newWidth = originalHeight * targetAspectRatio;
                float cropWidth = (originalWidth - newWidth) / 2;
                shape.PictureFormat.CropLeft += cropWidth;
                shape.PictureFormat.CropRight += cropWidth;
            }
            else
            {
                // 当前形状比例窄于目标比例，裁剪高度
                float newHeight = originalWidth / targetAspectRatio;
                float cropHeight = (originalHeight - newHeight) / 2;
                shape.PictureFormat.CropTop += cropHeight;
                shape.PictureFormat.CropBottom += cropHeight;
            }

            // 缩放到目标尺寸
            shape.Width = targetWidth;
            shape.Height = targetHeight;
        }

        private void 交换图层_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var pptApp = Globals.ThisAddIn.Application;
                var selection = pptApp.ActiveWindow.Selection;

                // 获取选中的形状范围
                var selectedShapes = GetSelectedShapesFromSelection(selection);

                // 检查是否选中两个形状
                if (selectedShapes.Count == 2)
                {
                    var shape1 = selectedShapes[0];
                    var shape2 = selectedShapes[1];

                    // 保存shape1和shape2的图层顺序
                    int shape1ZOrderPosition = shape1.ZOrderPosition;
                    int shape2ZOrderPosition = shape2.ZOrderPosition;

                    // 交换图层顺序
                    ExchangeShapeZOrderPosition(shape1, shape2ZOrderPosition);
                    ExchangeShapeZOrderPosition(shape2, shape1ZOrderPosition);
                }
                else
                {
                    MessageBox.Show($"请选中两个形状进行互换。当前选中形状数量：{selectedShapes.Count}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("交换图层时发生错误: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<PowerPoint.Shape> GetSelectedShapesFromSelection(PowerPoint.Selection selection)
        {
            var selectedShapes = new List<PowerPoint.Shape>();

            // 检查选择的类型
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                // 获取选中的形状范围
                PowerPoint.ShapeRange selectedShapeRange = selection.ShapeRange;
                if (selection.HasChildShapeRange)
                {
                    selectedShapeRange = selection.ChildShapeRange;
                }

                // 遍历选中的形状范围
                foreach (PowerPoint.Shape shape in selectedShapeRange)
                {
                    selectedShapes.Add(shape);
                }
            }

            // 如果选中的形状超过2个，只保留前两个
            if (selectedShapes.Count > 2)
            {
                selectedShapes = selectedShapes.Take(2).ToList();
            }

            // 返回选中的形状列表
            return selectedShapes;
        }

        private void ExchangeShapeZOrderPosition(PowerPoint.Shape shape, int targetZOrderPosition)
        {
            while (shape.ZOrderPosition > targetZOrderPosition)
            {
                shape.ZOrder(MsoZOrderCmd.msoSendBackward);
            }

            while (shape.ZOrderPosition < targetZOrderPosition)
            {
                shape.ZOrder(MsoZOrderCmd.msoBringForward);
            }
        }

        private void 交换文字_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 获取当前活动的PowerPoint应用程序
                var application = Globals.ThisAddIn.Application;

                // 获取当前选中的对象
                var selection = application.ActiveWindow.Selection;

                // 获取选中的形状范围
                var selectedShapes = GetSelectedShapesForTextExchange(selection);

                // 确保选中了两个对象
                if (selectedShapes.Count == 2)
                {
                    // 获取两个选中的形状
                    var shape1 = selectedShapes[0];
                    var shape2 = selectedShapes[1];

                    // 检查这两个形状是否包含文本
                    if ((shape1.HasTextFrame == MsoTriState.msoTrue && shape1.TextFrame.HasText == MsoTriState.msoTrue) &&
                        (shape2.HasTextFrame == MsoTriState.msoTrue && shape2.TextFrame.HasText == MsoTriState.msoTrue))
                    {
                        // 记录这两个形状的文本内容
                        string text1 = shape1.TextFrame.TextRange.Text;
                        string text2 = shape2.TextFrame.TextRange.Text;

                        // 交换文本内容
                        shape1.TextFrame.TextRange.Text = text2;
                        shape2.TextFrame.TextRange.Text = text1;
                    }
                    else
                    {
                        MessageBox.Show("请确保选中的两个对象都是文本框或带有文本的形状。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("请选中两个文本框或带有文本的形状以交换它们的文字内容。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("交换文字内容时发生错误: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<PowerPoint.Shape> GetSelectedShapesForTextExchange(PowerPoint.Selection selection)
        {
            var selectedShapes = new List<PowerPoint.Shape>();

            // 检查选择的类型
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                // 获取选中的形状范围
                PowerPoint.ShapeRange selectedShapeRange = selection.ShapeRange;
                if (selection.HasChildShapeRange)
                {
                    selectedShapeRange = selection.ChildShapeRange;
                }

                // 遍历选中的形状范围
                foreach (PowerPoint.Shape shape in selectedShapeRange)
                {
                    selectedShapes.Add(shape);
                }
            }

            // 如果选中的形状超过2个，只保留前两个
            if (selectedShapes.Count > 2)
            {
                selectedShapes = selectedShapes.Take(2).ToList();
            }

            // 返回选中的形状列表
            return selectedShapes;
        }

       
        private void 交换格式_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var application = Globals.ThisAddIn.Application;
                var selection = application.ActiveWindow.Selection;

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes && GetSelectedShapeCount(selection) == 2)
                {
                    var selectedShapes = GetSelectedShapesForFormatSwap(selection);
                    PowerPoint.Shape shape1 = selectedShapes[0];
                    PowerPoint.Shape shape2 = selectedShapes[1];

                    // 创建临时形状以保存格式
                    var slide = shape1.Parent;
                    PowerPoint.Shape tempShape1 = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, 1, 1);
                    PowerPoint.Shape tempShape2 = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, 1, 1);

                    // 保存形状1的格式到临时形状1
                    shape1.PickUp();
                    tempShape1.Apply();

                    // 保存形状2的格式到临时形状2
                    shape2.PickUp();
                    tempShape2.Apply();

                    // 交换格式
                    tempShape1.PickUp();
                    shape2.Apply();

                    tempShape2.PickUp();
                    shape1.Apply();

                    // 删除临时形状
                    tempShape1.Delete();
                    tempShape2.Delete();
                }
                else
                {
                    MessageBox.Show("请选中两个对象以交换它们的格式。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("交换格式时发生错误: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<PowerPoint.Shape> GetSelectedShapesForFormatSwap(PowerPoint.Selection selection)
        {
            var selectedShapes = new List<PowerPoint.Shape>();

            if (selection.HasChildShapeRange)
            {
                selectedShapes.Add(selection.ChildShapeRange[1]);
                selectedShapes.Add(selection.ChildShapeRange[2]);
            }
            else
            {
                selectedShapes.Add(selection.ShapeRange[1]);
                selectedShapes.Add(selection.ShapeRange[2]);
            }

            return selectedShapes;
        }

        private int GetSelectedShapeCount(PowerPoint.Selection selection)
        {
            if (selection.HasChildShapeRange)
            {
                return selection.ChildShapeRange.Count;
            }
            return selection.ShapeRange.Count;
        }

       

        private void 四线三格_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Selection sel = app.ActiveWindow.Selection;
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

                float defaultWidth = 5.0f * 28.3465f; // cm to points
                float defaultHeight = 1.8f * 28.3465f; // cm to points
                float additionalHeight = 0.25f * 28.3465f; // 0.25 cm to points

                PowerPoint.Shape gridGroup = null;

                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    // Insert default "four-line three-grid" shape
                    gridGroup = InsertFourLineThreeGrid(slide, defaultWidth, defaultHeight);
                }
                else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    PowerPoint.Shape shape = sel.ShapeRange[1];
                    if (shape.Type == Office.MsoShapeType.msoTable)
                    {
                        // Insert above the selected table
                        float tableWidth = shape.Width;
                        float tableTop = shape.Top - 10 - defaultHeight; // Adjust the top position correctly
                        gridGroup = InsertFourLineThreeGrid(slide, tableWidth, defaultHeight);
                        gridGroup.Top = tableTop;
                        gridGroup.Left = shape.Left; // Align left
                    }
                    else if (shape.Type == Office.MsoShapeType.msoTextBox)
                    {
                        // Insert aligned with the top of the selected text box
                        float textBoxWidth = shape.Width;
                        float textBoxHeight = shape.Height;
                        float newHeight = textBoxHeight + additionalHeight; // Text box height + 0.25 cm
                        gridGroup = InsertFourLineThreeGrid(slide, textBoxWidth, newHeight);

                        // Ensure the four-line three-grid is centered horizontally with the text box
                        float textBoxCenter = shape.Left + (textBoxWidth / 2);
                        gridGroup.Left = textBoxCenter - (gridGroup.Width / 2);

                        // Align the top of the four-line three-grid with the top of the text box
                        gridGroup.Top = shape.Top;

                        // Bring the text box to the front
                        shape.ZOrder(MsoZOrderCmd.msoBringToFront);
                    }
                }
                else
                {
                    throw new InvalidOperationException("不支持的选中对象类型。");
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("发生错误：" + ex.Message);
            }
        }

        private PowerPoint.Shape InsertFourLineThreeGrid(PowerPoint.Slide slide, float width, float height)
        {
            float lineSpacing = height / 3.0f;
            PowerPoint.Shapes shapes = slide.Shapes;
            PowerPoint.Shape line1 = shapes.AddLine(0, 0, width, 0);
            PowerPoint.Shape line2 = shapes.AddLine(0, lineSpacing, width, lineSpacing);
            PowerPoint.Shape line3 = shapes.AddLine(0, lineSpacing * 2, width, lineSpacing * 2);
            PowerPoint.Shape line4 = shapes.AddLine(0, height, width, height);

            line1.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            line1.Line.Weight = 1.5f;
            line4.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            line4.Line.Weight = 1.5f;
            line2.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
            line2.Line.Weight = 1.0f;
            line3.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            line3.Line.Weight = 1.0f;

            PowerPoint.ShapeRange shapeRange = slide.Shapes.Range(new string[] { line1.Name, line2.Name, line3.Name, line4.Name });
            return shapeRange.Group();
        }

        private PowerPoint.Shape AdjustFourLineThreeGrid(PowerPoint.Shape gridGroup, float newSpacing)
        {
            PowerPoint.ShapeRange shapes = gridGroup.Ungroup();
            shapes[1].Top = newSpacing;
            shapes[2].Top = newSpacing * 2;
            shapes[3].Top = newSpacing * 3;
            return shapes.Group();
        }

        private float GetMinCharacterHeight(PowerPoint.Shape textBox)
        {
            PowerPoint.TextRange textRange = textBox.TextFrame.TextRange;
            float minHeight = float.MaxValue;
            for (int i = 1; i <= textRange.Length; i++)
            {
                float charHeight = textRange.Characters(i, 1).BoundHeight;
                if (charHeight < minHeight)
                {
                    minHeight = charHeight;
                }
            }
            return minHeight == float.MaxValue ? 0 : minHeight;
        }

        private void 完全交换_Click(object sender, RibbonControlEventArgs e)
        {
            var pptApp = Globals.ThisAddIn.Application;
            var selection = pptApp.ActiveWindow.Selection;

            var selectedShapes = GetSelectedShapes(selection);

            if (selectedShapes.Count == 2)
            {
                var shape1 = selectedShapes[0];
                var shape2 = selectedShapes[1];

                var shape1Properties = SaveShapeProperties(shape1);
                int shape1ZOrderPosition = shape1.ZOrderPosition;

                var shape2Properties = SaveShapeProperties(shape2);
                int shape2ZOrderPosition = shape2.ZOrderPosition;

                ApplyShapeProperties(shape1, shape2Properties);
                ApplyShapeProperties(shape2, shape1Properties);

                SetShapeZOrderPosition(shape1, shape2ZOrderPosition);
                SetShapeZOrderPosition(shape2, shape1ZOrderPosition);
            }
            else
            {
                MessageBox.Show($"请选中两个形状进行互换。当前选中形状数量：{selectedShapes.Count}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<PowerPoint.Shape> GetSelectedShapes(PowerPoint.Selection selection)
        {
            var selectedShapes = new List<PowerPoint.Shape>();

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    selectedShapes.Add(shape);
                }
            }

            if (selectedShapes.Count > 2)
            {
                selectedShapes = selectedShapes.Take(2).ToList();
            }

            if (selection.HasChildShapeRange)
            {
                selectedShapes.Clear();
                foreach (PowerPoint.Shape shape in selection.ChildShapeRange)
                {
                    selectedShapes.Add(shape);
                }
            }

            return selectedShapes;
        }

        private (float Left, float Top, float Width, float Height, float Rotation, float ThreeDRotationX, float ThreeDRotationY, float ThreeDRotationZ, string Text, float ShadowBlur, float Depth) SaveShapeProperties(PowerPoint.Shape shape)
        {
            float left = shape.Left;
            float top = shape.Top;
            float width = shape.Width;
            float height = shape.Height;
            float rotation = shape.Rotation;

            float threeDRotationX = 0;
            float threeDRotationY = 0;
            float threeDRotationZ = 0;
            float shadowBlur = 0;
            float depth = 0;
            string text = "";

            try { threeDRotationX = shape.ThreeD.RotationX; } catch { /* Ignore if not applicable */ }
            try { threeDRotationY = shape.ThreeD.RotationY; } catch { /* Ignore if not applicable */ }
            try { threeDRotationZ = shape.ThreeD.RotationZ; } catch { /* Ignore if not applicable */ }
            try { shadowBlur = shape.Shadow.Blur; } catch { /* Ignore if not applicable */ }
            try { depth = shape.ThreeD.Depth; } catch { /* Ignore if not applicable */ }
            try { text = shape.TextFrame2.TextRange.Text; } catch { /* Ignore if not applicable */ }

            return (
                left,
                top,
                width,
                height,
                rotation,
                threeDRotationX,
                threeDRotationY,
                threeDRotationZ,
                text,
                shadowBlur,
                depth
            );
        }

        private void ApplyShapeProperties(PowerPoint.Shape shape, (float Left, float Top, float Width, float Height, float Rotation, float ThreeDRotationX, float ThreeDRotationY, float ThreeDRotationZ, string Text, float ShadowBlur, float Depth) properties)
        {
            try
            {
                shape.Width = properties.Width;
                shape.Height = properties.Height;
                shape.Rotation = properties.Rotation;
                shape.ThreeD.RotationX = properties.ThreeDRotationX;
                shape.ThreeD.RotationY = properties.ThreeDRotationY;
                shape.ThreeD.RotationZ = properties.ThreeDRotationZ;

                shape.Left = properties.Left;
                shape.Top = properties.Top;
                shape.TextFrame2.TextRange.Text = properties.Text;
                shape.Shadow.Blur = properties.ShadowBlur;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"无法修改形状属性：{ex.Message}");
            }
        }

        private void SetShapeZOrderPosition(PowerPoint.Shape shape, int targetZOrderPosition)
        {
            while (shape.ZOrderPosition > targetZOrderPosition)
            {
                shape.ZOrder(Core.MsoZOrderCmd.msoSendBackward);
            }

            while (shape.ZOrderPosition < targetZOrderPosition)
            {
                shape.ZOrder(Core.MsoZOrderCmd.msoBringForward);
            }
        }

        private void 移动对齐_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            MovingAlignmentForm form = new MovingAlignmentForm(app);
            form.Show();
        }

        private void 智能缩放_Click(object sender, RibbonControlEventArgs e)
        {
            SmartScalingForm scalingForm = new SmartScalingForm();
            scalingForm.Show();
        }

        private void 一键注音_Click(object sender, RibbonControlEventArgs e)
        {
            // 设置EPPlus的许可上下文
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // 从嵌入资源中提取汉字字典Excel文件
            string filePath = ExtractEmbeddedResource("课件帮PPT助手.汉字字典.汉字字典.xlsx");

            // 加载汉字拼音字典
            Dictionary<string, string> hanziPinyinDictionary = LoadHanziPinyinDictionary(filePath);

            // 获取当前PPT应用和选中的文本框或文本
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Selection pptSelection = pptApp.ActiveWindow.Selection;

            if (pptSelection.Type == PowerPoint.PpSelectionType.ppSelectionText || pptSelection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape selectedShape in pptSelection.ShapeRange)
                {
                    PowerPoint.TextRange textRange = selectedShape.TextFrame.TextRange;
                    string selectedText = textRange.Text;
                    string annotatedText = GetPinyinForText(selectedText, hanziPinyinDictionary);

                    // 获取所选文本框的位置和大小
                    float left = selectedShape.Left;
                    float top = selectedShape.Top - (selectedShape.Height / 4) - (textRange.Font.Size / 4); // 新文本框放置在原文本框的顶部，距离原文本框中心线一半字号的高度
                    float width = selectedShape.Width;
                    float newFontSize = textRange.Font.Size / 2;

                    // 创建新的文本框并插入注音后的文本
                    PowerPoint.Shape newShape = pptSelection.SlideRange[1].Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        left, top, width, selectedShape.Height / 2);
                    newShape.TextFrame.TextRange.Text = annotatedText;

                    // 设置新文本框的字体大小为原文本框字体大小的一半
                    newShape.TextFrame.TextRange.Font.Size = newFontSize;

                    // 设置新文本框的对齐方式与原文本框一致
                    newShape.TextFrame.TextRange.ParagraphFormat.Alignment = textRange.ParagraphFormat.Alignment;

                    // 确保新文本框不自动换行
                    newShape.TextFrame.WordWrap = Office.MsoTriState.msoFalse;
                }
            }
            else
            {
                MessageBox.Show("请先选中一个或多个文本框。");
            }
        }

        private Dictionary<string, string> LoadHanziPinyinDictionary(string filePath)
        {
            var hanziPinyinDictionary = new Dictionary<string, string>();
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++) // Assuming the first row is header
                {
                    string hanzi = worksheet.Cells[row, 1].Text;
                    string pinyin = worksheet.Cells[row, 2].Text;
                    if (!string.IsNullOrWhiteSpace(hanzi) && !string.IsNullOrWhiteSpace(pinyin))
                    {
                        hanziPinyinDictionary[hanzi] = pinyin;
                    }
                }
            }
            return hanziPinyinDictionary;
        }

        private string ExtractEmbeddedResource(string resourceName)
        {
            string tempFilePath = Path.GetTempFileName();
            using (var stream = GetType().Assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null) throw new ArgumentException("Resource not found.", nameof(resourceName));
                using (var fileStream = new FileStream(tempFilePath, FileMode.Create))
                {
                    stream.CopyTo(fileStream);
                }
            }
            return tempFilePath;
        }

        private string GetPinyinForText(string text, Dictionary<string, string> hanziPinyinDictionary)
        {
            List<string> pinyinList = new List<string>();

            foreach (char c in text)
            {
                if (hanziPinyinDictionary.ContainsKey(c.ToString()))
                {
                    string pinyin = hanziPinyinDictionary[c.ToString()];
                    if (pinyin.Contains(","))
                    {
                        // 如果有多个拼音，弹出选择对话框
                        using (PinYinForm form = new PinYinForm(c.ToString(), text, pinyin.Split(',')))
                        {
                            if (form.ShowDialog() == DialogResult.OK)
                            {
                                pinyinList.Add(form.SelectedPinyin);
                            }
                            else
                            {
                                pinyinList.Add(pinyin.Split(',')[0]);
                            }
                        }
                    }
                    else
                    {
                        pinyinList.Add(pinyin);
                    }
                }
                else
                {
                    pinyinList.Add(c.ToString());
                }
            }

            return string.Join(" ", pinyinList);
        }


        private ConcurrentDictionary<string, string> pinyinCache = new ConcurrentDictionary<string, string>();
        private async void 提取拼音_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前PPT应用和选中的文本框
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Selection pptSelection = pptApp.ActiveWindow.Selection;

            if (pptSelection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange shapeRange = pptSelection.ShapeRange;

                foreach (PowerPoint.Shape selectedShape in shapeRange)
                {
                    if (selectedShape.HasTextFrame == Office.MsoTriState.msoTrue && selectedShape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        PowerPoint.TextRange textRange = selectedShape.TextFrame.TextRange;
                        string selectedText = textRange.Text.Trim();
                        string pinyinText = await GetPinyinFromWeb(selectedText);

                        // 获取所选文本框的位置和大小
                        float left = selectedShape.Left;
                        float top = selectedShape.Top - (selectedShape.Height / 4) - (textRange.Font.Size / 4); // 新文本框放置在原文本框的顶部，距离原文本框中心线一半字号的高度
                        float width = selectedShape.Width;
                        float newFontSize = textRange.Font.Size / 2;

                        // 创建新的文本框并插入注音后的文本
                        PowerPoint.Shape newShape = pptSelection.SlideRange[1].Shapes.AddTextbox(
                            Office.MsoTextOrientation.msoTextOrientationHorizontal,
                            left, top, width, selectedShape.Height / 2);
                        newShape.TextFrame.TextRange.Text = pinyinText;

                        // 设置新文本框的字体大小为原文本框字体大小的一半
                        newShape.TextFrame.TextRange.Font.Size = newFontSize;

                        // 设置新文本框的对齐方式与原文本框一致
                        newShape.TextFrame.TextRange.ParagraphFormat.Alignment = textRange.ParagraphFormat.Alignment;

                        // 确保拼音文本框取消自动换行
                        newShape.TextFrame.WordWrap = Office.MsoTriState.msoFalse;

                        // 设置拼音文本框置于最顶层
                        newShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                    }
                }
            }
            else
            {
                MessageBox.Show("请先选中一个或多个文本框。");
            }
        }

        private async Task<string> GetPinyinFromWeb(string text)
        {
            if (pinyinCache.TryGetValue(text, out string cachedPinyin))
            {
                return cachedPinyin;
            }

            string url = $"https://hanyu.baidu.com/s?wd={Uri.EscapeDataString(text)}&ptype=zici";
            HtmlWeb web = new HtmlWeb();
            HtmlAgilityPack.HtmlDocument doc = await web.LoadFromWebAsync(url);

            try
            {
                string pinyin = string.Empty;

                if (text.Length == 1)
                {
                    // 单字符处理
                    var pinyinNode = doc.DocumentNode.SelectSingleNode("//span/b");
                    if (pinyinNode != null)
                    {
                        pinyin = pinyinNode.InnerText;
                    }
                }
                else
                {
                    // 多字符处理
                    var pinyinNode = doc.DocumentNode.SelectSingleNode("//b[@class='pinyin-font']");
                    if (pinyinNode != null)
                    {
                        pinyin = pinyinNode.InnerText;
                    }
                }

                if (!string.IsNullOrEmpty(pinyin))
                {
                    pinyin = pinyin.Replace("[", "").Replace("]", "");
                    pinyinCache[text] = pinyin; // 缓存结果
                }

                return pinyin;
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }


        private async Task<string> GetPinyinText(string selectedText)
        {
            // 获取拼音的异步方法代码
            return await GetPinyinFromWeb(selectedText);
        }

        private async void Zici_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前PPT应用和选中的文本框
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Selection pptSelection = pptApp.ActiveWindow.Selection;

            if (pptSelection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape selectedShape in pptSelection.ShapeRange)
                {
                    if (selectedShape.HasTextFrame == Office.MsoTriState.msoTrue && selectedShape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        await ProcessShapeAsync(selectedShape);
                    }
                }
            }
        }

        private async Task ProcessShapeAsync(PowerPoint.Shape selectedShape)
        {
            PowerPoint.TextRange textRange = selectedShape.TextFrame.TextRange;
            if (textRange != null && !string.IsNullOrEmpty(textRange.Text))
            {
                string selectedText = textRange.Text;

                // 获取拼音
                string pinyin = await GetPinyinText(selectedText);

                // 创建拼音文本框
                PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
                PowerPoint.Shape pinyinShape = pptApp.ActivePresentation.Slides[pptApp.ActiveWindow.View.Slide.SlideIndex].Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    selectedShape.Left,
                    selectedShape.Top - 20, // 拼音文本框放在原文本框上方
                    selectedShape.Width,
                    20 // 高度设置为20，根据需要调整
                );

                pinyinShape.TextFrame.TextRange.Text = pinyin;
                pinyinShape.TextFrame.TextRange.Font.Size = textRange.Font.Size / 2; // 拼音字体大小为原字体的一半
                pinyinShape.TextFrame.TextRange.ParagraphFormat.Alignment = textRange.ParagraphFormat.Alignment;
                pinyinShape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorTop;
                pinyinShape.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
                pinyinShape.TextFrame.WordWrap = Office.MsoTriState.msoFalse; // 取消自动换行

                // 计算括号文本框需要的宽度
                int numSpaces = selectedText.Length * 4; // 简单计算所需空格数量，可根据需要调整
                string spaces = new string(' ', numSpaces);
                string parenthesesText = $"({spaces})";

                // 创建括号文本框
                PowerPoint.Shape parenthesesShape = pptApp.ActivePresentation.Slides[pptApp.ActiveWindow.View.Slide.SlideIndex].Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    selectedShape.Left,
                    selectedShape.Top,
                    selectedShape.Width,
                    selectedShape.Height
                );

                parenthesesShape.TextFrame.TextRange.Text = parenthesesText;
                parenthesesShape.TextFrame.TextRange.Font.Size = textRange.Font.Size;
                parenthesesShape.TextFrame.TextRange.ParagraphFormat.Alignment = textRange.ParagraphFormat.Alignment;
                parenthesesShape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                parenthesesShape.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
                parenthesesShape.TextFrame.WordWrap = Office.MsoTriState.msoFalse; // 取消自动换行

                // 调整括号文本框的位置，使其在水平和垂直方向上居中对齐
                parenthesesShape.Left = selectedShape.Left + (selectedShape.Width - parenthesesShape.Width) / 2;
                parenthesesShape.Top = selectedShape.Top + (selectedShape.Height - parenthesesShape.Height) / 2;

                // 修改用户所选文本的字体样式
                textRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Red);
                textRange.Font.Bold = Office.MsoTriState.msoTrue;
                selectedShape.TextFrame.WordWrap = Office.MsoTriState.msoFalse; // 取消自动换行
            }
        }

        private async Task Call提取拼音_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Selection pptSelection = pptApp.ActiveWindow.Selection;

            if (pptSelection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape selectedShape in pptSelection.ShapeRange)
                {
                    if (selectedShape.HasTextFrame == Office.MsoTriState.msoTrue && selectedShape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        await ProcessShapeAsync(selectedShape);
                    }
                }
            }
        }



        private async Task<string> GetPinyinTextAsync(string selectedText)
        {
            return await GetPinyinFromWeb(selectedText);
        }

        private async void WritePinyin_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Selection pptSelection = pptApp.ActiveWindow.Selection;

            if (pptSelection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape selectedShape in pptSelection.ShapeRange)
                {
                    if (selectedShape.HasTextFrame == Office.MsoTriState.msoTrue && selectedShape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        await ProcessShapeForWritePinyinAsync(selectedShape);
                    }
                }
            }
        }

        private async Task ProcessShapeForWritePinyinAsync(PowerPoint.Shape selectedShape)
        {
            PowerPoint.TextRange textRange = selectedShape.TextFrame.TextRange;
            if (textRange != null && !string.IsNullOrEmpty(textRange.Text))
            {
                string selectedText = textRange.Text;
                string pinyin = await GetPinyinTextAsync(selectedText);

                PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
                float originalLeft = selectedShape.Left;
                float originalTop = selectedShape.Top;
                float originalWidth = selectedShape.Width;
                float originalHeight = selectedShape.Height;

                float pinyinWidth = MeasureTextWidth(pinyin, textRange.Font.Size - 2, textRange.Font.Name);

                // 动态计算空格符数量
                float spaceWidth = MeasureTextWidth(" ", textRange.Font.Size, textRange.Font.Name);
                int additionalSpaces = 2; // 固定增加的空格符数量
                int numSpaces = (int)Math.Ceiling(pinyinWidth / spaceWidth) + additionalSpaces;
                string spaces = new string(' ', numSpaces);


                // 创建括号文本框
                PowerPoint.Shape parenthesesShape = pptApp.ActivePresentation.Slides[pptApp.ActiveWindow.View.Slide.SlideIndex].Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    originalLeft + originalWidth,
                    originalTop,
                    0,
                    originalHeight
                );

                parenthesesShape.TextFrame.TextRange.Text = $"({spaces})";
                parenthesesShape.TextFrame.TextRange.Font.Size = textRange.Font.Size;
                parenthesesShape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                parenthesesShape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                parenthesesShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black);
                parenthesesShape.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                parenthesesShape.TextFrame.WordWrap = Office.MsoTriState.msoFalse;
                parenthesesShape.ZOrder(Office.MsoZOrderCmd.msoSendToBack);

                // 获取括号文本框实际宽度
                float parenthesesWidth = parenthesesShape.Width;

                // 创建拼音文本框
                PowerPoint.Shape pinyinShape = pptApp.ActivePresentation.Slides[pptApp.ActiveWindow.View.Slide.SlideIndex].Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    originalLeft + originalWidth + parenthesesWidth,
                    originalTop,
                    pinyinWidth,
                    originalHeight
                );

                pinyinShape.TextFrame.TextRange.Text = pinyin;
                pinyinShape.TextFrame.TextRange.Font.Size = textRange.Font.Size - 2;
                pinyinShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Red);
                pinyinShape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                pinyinShape.TextFrame.TextRange.ParagraphFormat.Alignment = textRange.ParagraphFormat.Alignment;
                pinyinShape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                pinyinShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                pinyinShape.TextFrame.WordWrap = Office.MsoTriState.msoFalse;

                // 确保括号文本框与用户所选文本框紧紧挨着
                parenthesesShape.Left = originalLeft + originalWidth;

                // 调整拼音文本框的位置，使其水平居中对齐
                pinyinShape.Left = parenthesesShape.Left + (parenthesesShape.Width - pinyinShape.Width) / 2;
                pinyinShape.Top = parenthesesShape.Top + (parenthesesShape.Height - pinyinShape.Height) / 2;


                // 确保三个文本框在同一水平线上
                parenthesesShape.Top = selectedShape.Top;
                pinyinShape.Top = selectedShape.Top;
            }
        }

        private float MeasureTextWidth(string text, float fontSize, string fontName)
        {
            using (var bmp = new Bitmap(1, 1))
            {
                using (var g = Graphics.FromImage(bmp))
                {
                    var font = new System.Drawing.Font(fontName, fontSize);
                    var size = g.MeasureString(text, font);
                    return size.Width;
                }
            }
        }


      

        private void 生字教学_Click(object sender, RibbonControlEventArgs e)
        {
            string pptPath = null;
            PowerPoint.Presentation sourcePresentation = null;
            try
            {
                // 提取并打开嵌入的PPT资源
                pptPath = ExtractResourceFile("课件帮PPT助手.Resources.生字教学.pptx");
                if (string.IsNullOrEmpty(pptPath))
                {
                    MessageBox.Show("无法提取PPT资源。");
                    return;
                }

                PowerPoint.Application app = Globals.ThisAddIn.Application;
                sourcePresentation = app.Presentations.Open(pptPath, WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);

                // 获取第一个幻灯片
                PowerPoint.Slide sourceSlide = sourcePresentation.Slides[1];

                // 获取选中文字并获取拼音
                string character = GetSelectedCharacter(app);
                if (string.IsNullOrEmpty(character))
                {
                    MessageBox.Show("未找到选中的汉字。");
                    return;
                }
                string filePath = ExtractResourceFile("课件帮PPT助手.汉字字典.汉字字典.xlsx");
                if (string.IsNullOrEmpty(filePath))
                {
                    MessageBox.Show("无法提取汉字字典资源。");
                    return;
                }
                string pinyin = GetPinyinResult(filePath, character);
                if (string.IsNullOrEmpty(pinyin))
                {
                    MessageBox.Show("未找到拼音信息。");
                    return;
                }

                // 获取拼音返回文本框并替换内容
                PowerPoint.Shape pinyinTextBox = FindShapeByName(sourceSlide, "拼音返回文本框");
                if (pinyinTextBox != null)
                {
                    pinyinTextBox.TextFrame.TextRange.Text = pinyin;
                }
                else
                {
                    MessageBox.Show("未找到名为‘拼音返回文本框’的形状。");
                    return;
                }

                // 替换 [笔画拆分] 替换文本框内的文本
                PowerPoint.Shape strokeReplaceTextBox = FindShapeByName(sourceSlide, "[笔画拆分]替换");
                if (strokeReplaceTextBox != null)
                {
                    strokeReplaceTextBox.TextFrame.TextRange.Text = character; // 这里替换为用户选中的文本
                }
                else
                {
                    MessageBox.Show("未找到名为‘[笔画拆分]替换’的形状。");
                    return;
                }

                // 获取汉字信息
                var characterInfo = GetCharacterInfo(filePath, character);
                if (characterInfo == null)
                {
                    MessageBox.Show("未找到相关汉字信息。");
                    return;
                }

                // 更新部首、结构、笔画信息到表格
                PowerPoint.Shape tableShape = FindShapeByName(sourceSlide, "表格（部首-结构-笔画）");
                if (tableShape != null)
                {
                    var table = tableShape.Table;
                    if (table.Columns.Count >= 2 && table.Rows.Count >= 3)
                    {
                        table.Cell(1, 2).Shape.TextFrame.TextRange.Text = characterInfo.Radical;
                        table.Cell(2, 2).Shape.TextFrame.TextRange.Text = characterInfo.Structure;
                        table.Cell(3, 2).Shape.TextFrame.TextRange.Text = characterInfo.Strokes.ToString();
                    }
                    else
                    {
                        MessageBox.Show("表格的行数或列数不足。");
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("未找到名为‘表格（部首-结构-笔画）’的形状。");
                    return;
                }

                // 更新相关组词信息到指定形状
                for (int i = 0; i < 8; i++)
                {
                    PowerPoint.Shape groupWordShape = FindShapeByName(sourceSlide, $"相关组词-{i + 1}");
                    if (groupWordShape != null)
                    {
                        groupWordShape.TextFrame.TextRange.Text = characterInfo.RelatedWords.ElementAtOrDefault(i) ?? "";
                    }
                    else
                    {
                        MessageBox.Show($"未找到名为‘相关组词-{i + 1}’的形状。");
                    }
                }

                // 获取当前演示文稿和当前幻灯片
                PowerPoint.Presentation currentPresentation = app.ActivePresentation;
                PowerPoint.Slide currentSlide = app.ActiveWindow.View.Slide;

                // 复制修改后的源幻灯片到当前演示文稿的下一页
                sourceSlide.Copy();
                currentPresentation.Slides.Paste(currentSlide.SlideIndex + 1);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"发生错误：{ex.Message}");
            }
            finally
            {
                // 释放资源并删除临时文件
                if (sourcePresentation != null)
                {
                    sourcePresentation.Close();
                }
                if (pptPath != null && File.Exists(pptPath))
                {
                    try
                    {
                        File.Delete(pptPath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"无法删除临时文件：{ex.Message}");
                    }
                }
            }
        }

        private string GetPinyinResult(string filePath, string character)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                if (package.Workbook.Worksheets.Count == 0)
                {
                    throw new Exception("没有查到相关信息。");
                }

                var worksheet = package.Workbook.Worksheets[0];

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    if (worksheet.Cells[row, 1].Text == character)
                    {
                        return worksheet.Cells[row, 2].Text;
                    }
                }
            }

            return null;
        }

        private PowerPoint.Shape FindStrokeSplitShape(PowerPoint.Slide slide, string shapeName)
        {
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                if (shape.Name == shapeName)
                {
                    return shape;
                }
            }
            return null;
        }

        private string GetSelectedCharacter(PowerPoint.Application app)
        {
            PowerPoint.Selection selection = app.ActiveWindow.Selection;
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
            {
                return selection.TextRange.Text.Trim();
            }
            MessageBox.Show("请选中文字进行操作。");
            return null;
        }

        private HanziInfo GetCharacterInfo(string filePath, string character)
        {
            var info = new HanziInfo();

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                if (package.Workbook.Worksheets.Count == 0)
                {
                    throw new Exception("没有查到相关信息。");
                }

                var worksheet = package.Workbook.Worksheets[0];

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    if (worksheet.Cells[row, 1].Text == character)
                    {
                        info.Radical = worksheet.Cells[row, 3].Text;
                        info.Structure = worksheet.Cells[row, 4].Text;
                        info.Strokes = int.Parse(worksheet.Cells[row, 5].Text);
                        info.RelatedWords = worksheet.Cells[row, 6].Text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                    .Select(word => word.Trim())
                                                    .Where(word => !string.IsNullOrEmpty(word))
                                                    .ToArray();
                        break;
                    }
                }
            }

            return info;
        }

        private string ExtractResourceFile(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (Stream resourceStream = assembly.GetManifestResourceStream(resourceName))
            {
                if (resourceStream == null)
                    throw new Exception($"未找到嵌入资源 {resourceName}.");

                string tempFilePath = Path.Combine(Path.GetTempPath(), resourceName);

                using (var fileStream = new FileStream(tempFilePath, FileMode.Create))
                {
                    resourceStream.CopyTo(fileStream);
                }

                return tempFilePath;
            }
        }

        private PowerPoint.Shape FindShapeByName(PowerPoint.Slide slide, string shapeName)
        {
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                if (shape.Name == shapeName)
                {
                    return shape;
                }
            }
            return null;
        }

        public class HanziInfo
        {
            public string Radical { get; set; }
            public string Structure { get; set; }
            public int Strokes { get; set; }
            public string[] RelatedWords { get; set; }
        }

        private void 检测字体_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 获取当前演示文稿
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Presentation presentation = app.ActivePresentation;

                // 获取演示文稿中所有已使用的字体
                List<string> usedFonts = new List<string>();
                List<string> unusedFonts = new List<string>();

                foreach (PowerPoint.Slide slide in presentation.Slides)
                {
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        CollectFontsFromShape(shape, usedFonts, unusedFonts);
                    }
                }
                usedFonts = usedFonts.Distinct().ToList();
                unusedFonts = unusedFonts.Distinct().Except(usedFonts).ToList();

                FontDetectionForm form = new FontDetectionForm(usedFonts, unusedFonts, presentation);
                form.Show();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("检测过程中出错: " + ex.Message, "错误", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void CollectFontsFromShape(PowerPoint.Shape shape, List<string> usedFonts, List<string> unusedFonts)
        {
            if (shape.HasTextFrame == MsoTriState.msoTrue)
            {
                PowerPoint.TextRange textRange = shape.TextFrame.TextRange;
                foreach (PowerPoint.TextRange run in textRange.Runs(1, textRange.Text.Length))
                {
                    string fontName = run.Font.Name;
                    if (!usedFonts.Contains(fontName))
                    {
                        usedFonts.Add(fontName);
                    }
                }
            }

            // 检查没有文本但有字体设置的形状
            if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoFalse)
            {
                var fonts = shape.TextFrame.TextRange.Font;
                string fontName = fonts.Name;
                if (!usedFonts.Contains(fontName))
                {
                    unusedFonts.Add(fontName);
                }
            }

            if (shape.Type == MsoShapeType.msoGroup)
            {
                foreach (PowerPoint.Shape groupedShape in shape.GroupItems)
                {
                    CollectFontsFromShape(groupedShape, usedFonts, unusedFonts);
                }
            }
        }

        private List<string> GetUnusedFonts(PowerPoint.Presentation presentation, List<string> usedFonts)
        {
            List<string> unusedFonts = new List<string>();
            foreach (PowerPoint.Font font in presentation.Fonts)
            {
                if (!usedFonts.Contains(font.Name))
                {
                    unusedFonts.Add(font.Name);
                }
            }
            return unusedFonts;
        }

       

        private void Selectsize_Click(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = application.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                var selectedShape = GetFirstShapeFromSelectionForSize(selection);
                var selectedShapeWidth = selectedShape.Width; // 获取选中形状的宽度
                var selectedShapeHeight = selectedShape.Height; // 获取选中形状的高度
                List<PowerPoint.Shape> sameSizeShapes = new List<PowerPoint.Shape>();

                if (IsShapeInGroupForSize(selectedShape))
                {
                    var parentGroup = GetParentGroupForSize(selectedShape);
                    foreach (PowerPoint.Shape shape in GetAllShapesFromGroupForSize(parentGroup.GroupItems))
                    {
                        if (shape.Width == selectedShapeWidth && shape.Height == selectedShapeHeight)
                        {
                            sameSizeShapes.Add(shape);
                        }
                    }
                }
                else
                {
                    var slide = application.ActiveWindow.View.Slide;
                    foreach (PowerPoint.Shape shape in GetAllShapesForSize(slide.Shapes))
                    {
                        if (shape.Width == selectedShapeWidth && shape.Height == selectedShapeHeight && !IsShapeInGroupForSize(shape))
                        {
                            sameSizeShapes.Add(shape);
                        }
                    }
                }

                if (sameSizeShapes.Count > 0)
                {
                    var shapeNamesArray = sameSizeShapes.Select(shape => shape.Name).ToArray();
                    var shapeRange = application.ActiveWindow.View.Slide.Shapes.Range(shapeNamesArray);
                    shapeRange.Select();
                }
            }
            else
            {
                MessageBox.Show("请选择一个对象。");
            }
        }

        private PowerPoint.Shape GetFirstShapeFromSelectionForSize(PowerPoint.Selection selection)
        {
            if (selection.HasChildShapeRange)
            {
                return selection.ChildShapeRange[1];
            }
            return selection.ShapeRange[1];
        }

        private List<PowerPoint.Shape> GetAllShapesForSize(PowerPoint.Shapes shapes)
        {
            var allShapes = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in shapes)
            {
                if (shape.Type == MsoShapeType.msoGroup)
                {
                    allShapes.AddRange(GetAllShapesFromGroupForSize(shape.GroupItems));
                }
                else
                {
                    allShapes.Add(shape);
                }
            }
            return allShapes;
        }

        private List<PowerPoint.Shape> GetAllShapesFromGroupForSize(PowerPoint.GroupShapes groupShapes)
        {
            var allShapes = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in groupShapes)
            {
                if (shape.Type == MsoShapeType.msoGroup)
                {
                    allShapes.AddRange(GetAllShapesFromGroupForSize(shape.GroupItems));
                }
                else
                {
                    allShapes.Add(shape);
                }
            }
            return allShapes;
        }

        private bool IsShapeInGroupForSize(PowerPoint.Shape shape)
        {
            try
            {
                var parent = shape.ParentGroup;
                return parent != null;
            }
            catch
            {
                return false;
            }
        }

        private PowerPoint.Shape GetParentGroupForSize(PowerPoint.Shape shape)
        {
            return shape.ParentGroup;
        }

        private void SelectedColor_Click(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = application.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                var selectedShape = GetFirstShapeFromSelectionForColor(selection);
                var selectedShapeFillColor = selectedShape.Fill.ForeColor.RGB; // 获取选中形状的填充颜色
                List<PowerPoint.Shape> sameColorShapes = new List<PowerPoint.Shape>();

                if (IsShapeInGroupForColor(selectedShape))
                {
                    var parentGroup = GetParentGroupForColor(selectedShape);
                    foreach (PowerPoint.Shape shape in GetAllShapesFromGroupForColor(parentGroup.GroupItems))
                    {
                        if (shape.Fill.ForeColor.RGB == selectedShapeFillColor)
                        {
                            sameColorShapes.Add(shape);
                        }
                    }
                }
                else
                {
                    var slide = application.ActiveWindow.View.Slide;
                    foreach (PowerPoint.Shape shape in GetAllShapesForColor(slide.Shapes))
                    {
                        if (shape.Fill.ForeColor.RGB == selectedShapeFillColor && !IsShapeInGroupForColor(shape))
                        {
                            sameColorShapes.Add(shape);
                        }
                    }
                }

                if (sameColorShapes.Count > 0)
                {
                    var shapeNamesArray = sameColorShapes.Select(shape => shape.Name).ToArray();
                    var shapeRange = application.ActiveWindow.View.Slide.Shapes.Range(shapeNamesArray);
                    shapeRange.Select();
                }
            }
            else
            {
                MessageBox.Show("请选择一个对象。");
            }
        }

        private PowerPoint.Shape GetFirstShapeFromSelectionForColor(PowerPoint.Selection selection)
        {
            if (selection.HasChildShapeRange)
            {
                return selection.ChildShapeRange[1];
            }
            return selection.ShapeRange[1];
        }

        private List<PowerPoint.Shape> GetAllShapesForColor(PowerPoint.Shapes shapes)
        {
            var allShapes = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in shapes)
            {
                if (shape.Type == MsoShapeType.msoGroup)
                {
                    allShapes.AddRange(GetAllShapesFromGroupForColor(shape.GroupItems));
                }
                else
                {
                    allShapes.Add(shape);
                }
            }
            return allShapes;
        }

        private List<PowerPoint.Shape> GetAllShapesFromGroupForColor(PowerPoint.GroupShapes groupShapes)
        {
            var allShapes = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in groupShapes)
            {
                if (shape.Type == MsoShapeType.msoGroup)
                {
                    allShapes.AddRange(GetAllShapesFromGroupForColor(shape.GroupItems));
                }
                else
                {
                    allShapes.Add(shape);
                }
            }
            return allShapes;
        }

        private bool IsShapeInGroupForColor(PowerPoint.Shape shape)
        {
            try
            {
                var parent = shape.ParentGroup;
                return parent != null;
            }
            catch
            {
                return false;
            }
        }

        private PowerPoint.Shape GetParentGroupForColor(PowerPoint.Shape shape)
        {
            return shape.ParentGroup;
        }

        private void Selectedline_Click(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = application.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                var selectedShape = GetFirstShapeFromSelectionForLine(selection);
                var slide = application.ActiveWindow.View.Slide;
                List<PowerPoint.Shape> sameLineShapes = new List<PowerPoint.Shape>();

                if (IsShapeInGroupForLine(selectedShape))
                {
                    var parentGroup = GetParentGroupForLine(selectedShape);

                    if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
                    {
                        // 获取选中形状的轮廓颜色
                        var selectedLineColor = selectedShape.Line.ForeColor.RGB;

                        foreach (PowerPoint.Shape shape in GetAllShapesFromGroupForLine(parentGroup.GroupItems))
                        {
                            if (shape.Line.ForeColor.RGB == selectedLineColor)
                            {
                                sameLineShapes.Add(shape);
                            }
                        }
                    }
                    else if ((Control.ModifierKeys & Keys.Shift) == Keys.Shift)
                    {
                        // 获取选中形状的轮廓线条类型
                        var selectedLineDashStyle = selectedShape.Line.DashStyle;

                        foreach (PowerPoint.Shape shape in GetAllShapesFromGroupForLine(parentGroup.GroupItems))
                        {
                            if ((MsoLineDashStyle)shape.Line.DashStyle == selectedLineDashStyle)
                            {
                                sameLineShapes.Add(shape);
                            }
                        }
                    }
                    else
                    {
                        // 获取选中形状的轮廓宽度
                        var selectedLineWidth = selectedShape.Line.Weight;

                        foreach (PowerPoint.Shape shape in GetAllShapesFromGroupForLine(parentGroup.GroupItems))
                        {
                            if (shape.Line.Weight == selectedLineWidth)
                            {
                                sameLineShapes.Add(shape);
                            }
                        }
                    }
                }
                else
                {
                    if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
                    {
                        // 获取选中形状的轮廓颜色
                        var selectedLineColor = selectedShape.Line.ForeColor.RGB;

                        foreach (PowerPoint.Shape shape in GetAllShapesForLine(slide.Shapes))
                        {
                            if (shape.Line.ForeColor.RGB == selectedLineColor && !IsShapeInGroupForLine(shape))
                            {
                                sameLineShapes.Add(shape);
                            }
                        }
                    }
                    else if ((Control.ModifierKeys & Keys.Shift) == Keys.Shift)
                    {
                        // 获取选中形状的轮廓线条类型
                        var selectedLineDashStyle = selectedShape.Line.DashStyle;

                        foreach (PowerPoint.Shape shape in GetAllShapesForLine(slide.Shapes))
                        {
                            if ((MsoLineDashStyle)shape.Line.DashStyle == selectedLineDashStyle && !IsShapeInGroupForLine(shape))
                            {
                                sameLineShapes.Add(shape);
                            }
                        }
                    }
                    else
                    {
                        // 获取选中形状的轮廓宽度
                        var selectedLineWidth = selectedShape.Line.Weight;

                        foreach (PowerPoint.Shape shape in GetAllShapesForLine(slide.Shapes))
                        {
                            if (shape.Line.Weight == selectedLineWidth && !IsShapeInGroupForLine(shape))
                            {
                                sameLineShapes.Add(shape);
                            }
                        }
                    }
                }

                if (sameLineShapes.Count > 0)
                {
                    var shapeNamesArray = sameLineShapes.Select(shape => shape.Name).ToArray();
                    var shapeRange = slide.Shapes.Range(shapeNamesArray);
                    shapeRange.Select();
                }
            }
            else
            {
                MessageBox.Show("请选择一个对象。");
            }
        }

        private PowerPoint.Shape GetFirstShapeFromSelectionForLine(PowerPoint.Selection selection)
        {
            if (selection.HasChildShapeRange)
            {
                return selection.ChildShapeRange[1];
            }
            return selection.ShapeRange[1];
        }

        private List<PowerPoint.Shape> GetAllShapesForLine(PowerPoint.Shapes shapes)
        {
            var allShapes = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in shapes)
            {
                if (shape.Type == MsoShapeType.msoGroup)
                {
                    allShapes.AddRange(GetAllShapesFromGroupForLine(shape.GroupItems));
                }
                else
                {
                    allShapes.Add(shape);
                }
            }
            return allShapes;
        }

        private List<PowerPoint.Shape> GetAllShapesFromGroupForLine(PowerPoint.GroupShapes groupShapes)
        {
            var allShapes = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in groupShapes)
            {
                if (shape.Type == MsoShapeType.msoGroup)
                {
                    allShapes.AddRange(GetAllShapesFromGroupForLine(shape.GroupItems));
                }
                else
                {
                    allShapes.Add(shape);
                }
            }
            return allShapes;
        }

        private bool IsShapeInGroupForLine(PowerPoint.Shape shape)
        {
            try
            {
                var parent = shape.ParentGroup;
                return parent != null;
            }
            catch
            {
                return false;
            }
        }

        private PowerPoint.Shape GetParentGroupForLine(PowerPoint.Shape shape)
        {
            return shape.ParentGroup;
        }

        private void Selectfontsize_Click(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = application.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                var selectedShape = GetFirstShapeFromSelectionForFontSize(selection);
                var slide = application.ActiveWindow.View.Slide;
                List<PowerPoint.Shape> sameFontSizeShapes = new List<PowerPoint.Shape>();

                // 获取选中形状的字体大小
                float selectedFontSize = 0;
                if (selectedShape.TextFrame.HasText == MsoTriState.msoTrue)
                {
                    selectedFontSize = selectedShape.TextFrame.TextRange.Font.Size;
                }

                // 如果选中形状没有字体大小，则不进行后续操作
                if (selectedFontSize == 0)
                {
                    MessageBox.Show("请选择一个包含文本的对象。");
                    return;
                }

                if (IsShapeInGroupForFontSize(selectedShape))
                {
                    var parentGroup = GetParentGroupForFontSize(selectedShape);

                    foreach (PowerPoint.Shape shape in GetAllShapesFromGroupForFontSize(parentGroup.GroupItems))
                    {
                        if (shape.TextFrame.HasText == MsoTriState.msoTrue &&
                            shape.TextFrame.TextRange.Font.Size == selectedFontSize)
                        {
                            sameFontSizeShapes.Add(shape);
                        }
                    }
                }
                else
                {
                    foreach (PowerPoint.Shape shape in GetAllShapesForFontSize(slide.Shapes))
                    {
                        if (shape.TextFrame.HasText == MsoTriState.msoTrue &&
                            shape.TextFrame.TextRange.Font.Size == selectedFontSize && !IsShapeInGroupForFontSize(shape))
                        {
                            sameFontSizeShapes.Add(shape);
                        }
                    }
                }

                if (sameFontSizeShapes.Count > 0)
                {
                    var shapeNamesArray = sameFontSizeShapes.Select(shape => shape.Name).ToArray();
                    var shapeRange = slide.Shapes.Range(shapeNamesArray);
                    shapeRange.Select();
                }
            }
            else
            {
                MessageBox.Show("请选择一个对象。");
            }
        }

        private PowerPoint.Shape GetFirstShapeFromSelectionForFontSize(PowerPoint.Selection selection)
        {
            if (selection.HasChildShapeRange)
            {
                return selection.ChildShapeRange[1];
            }
            return selection.ShapeRange[1];
        }

        private List<PowerPoint.Shape> GetAllShapesForFontSize(PowerPoint.Shapes shapes)
        {
            var allShapes = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in shapes)
            {
                if (shape.Type == MsoShapeType.msoGroup)
                {
                    allShapes.AddRange(GetAllShapesFromGroupForFontSize(shape.GroupItems));
                }
                else
                {
                    allShapes.Add(shape);
                }
            }
            return allShapes;
        }

        private List<PowerPoint.Shape> GetAllShapesFromGroupForFontSize(PowerPoint.GroupShapes groupShapes)
        {
            var allShapes = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in groupShapes)
            {
                if (shape.Type == MsoShapeType.msoGroup)
                {
                    allShapes.AddRange(GetAllShapesFromGroupForFontSize(shape.GroupItems));
                }
                else
                {
                    allShapes.Add(shape);
                }
            }
            return allShapes;
        }

        private bool IsShapeInGroupForFontSize(PowerPoint.Shape shape)
        {
            try
            {
                var parent = shape.ParentGroup;
                return parent != null;
            }
            catch
            {
                return false;
            }
        }

        private PowerPoint.Shape GetParentGroupForFontSize(PowerPoint.Shape shape)
        {
            return shape.ParentGroup;
        }

        private void Type_Click(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = application.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                var selectedShape = GetFirstShapeFromSelection(selection);
                var selectedShapeType = (MsoShapeType)selectedShape.Type;
                List<PowerPoint.Shape> sameTypeShapes = new List<PowerPoint.Shape>();

                if (selectedShapeType == MsoShapeType.msoGroup)
                {
                    // 如果选中的是组合对象，对当前页幻灯片的所有组合对象进行筛选
                    var slide = application.ActiveWindow.View.Slide;
                    foreach (PowerPoint.Shape shape in GetAllShapes(slide.Shapes))
                    {
                        if (shape.Type == MsoShapeType.msoGroup)
                        {
                            sameTypeShapes.Add(shape);
                        }
                    }
                }
                else if (IsShapeInGroup(selectedShape))
                {
                    // 如果选中的是组合内部的子对象，只在该组合内部进行筛选
                    var parentGroup = GetParentGroup(selectedShape);
                    foreach (PowerPoint.Shape shape in GetAllShapes(parentGroup.GroupItems))
                    {
                        if (shape.Type == selectedShapeType)
                        {
                            sameTypeShapes.Add(shape);
                        }
                    }
                }
                else
                {
                    // 如果选中的是独立对象，只筛选当前页幻灯片中的所有独立对象
                    var slide = application.ActiveWindow.View.Slide;
                    foreach (PowerPoint.Shape shape in GetAllShapes(slide.Shapes))
                    {
                        if (shape.Type == selectedShapeType && !IsShapeInGroup(shape))
                        {
                            sameTypeShapes.Add(shape);
                        }
                    }
                }

                if (sameTypeShapes.Count > 0)
                {
                    var shapeNamesArray = sameTypeShapes.Select(shape => shape.Name).ToArray();
                    var shapeRange = application.ActiveWindow.View.Slide.Shapes.Range(shapeNamesArray);
                    shapeRange.Select();
                }
            }
            else
            {
                MessageBox.Show("请选择一个对象。");
            }
        }

        private PowerPoint.Shape GetFirstShapeFromSelection(PowerPoint.Selection selection)
        {
            if (selection.HasChildShapeRange)
            {
                return selection.ChildShapeRange[1];
            }
            return selection.ShapeRange[1];
        }

        private List<PowerPoint.Shape> GetAllShapes(PowerPoint.Shapes shapes)
        {
            var allShapes = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in shapes)
            {
                if (shape.Type == MsoShapeType.msoGroup)
                {
                    allShapes.Add(shape);
                    allShapes.AddRange(GetAllShapes(shape.GroupItems));
                }
                else
                {
                    allShapes.Add(shape);
                }
            }
            return allShapes;
        }

        private List<PowerPoint.Shape> GetAllShapes(PowerPoint.GroupShapes groupShapes)
        {
            var allShapes = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in groupShapes)
            {
                if (shape.Type == MsoShapeType.msoGroup)
                {
                    allShapes.Add(shape);
                    allShapes.AddRange(GetAllShapes(shape.GroupItems));
                }
                else
                {
                    allShapes.Add(shape);
                }
            }
            return allShapes;
        }

        private bool IsShapeInGroup(PowerPoint.Shape shape)
        {
            try
            {
                var parent = shape.ParentGroup;
                return parent != null;
            }
            catch
            {
                return false;
            }
        }

        private PowerPoint.Shape GetParentGroup(PowerPoint.Shape shape)
        {
            return shape.ParentGroup;
        }
    }
}