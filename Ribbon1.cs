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

        // 按类型筛选
        private void Type_Click(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = application.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                var selectedShape = selection.ShapeRange[1];
                var selectedShapeType = (Office.MsoShapeType)selectedShape.Type; // 获取选中形状的类型
                var slide = application.ActiveWindow.View.Slide;
                List<int> sameTypeShapeIndices = new List<int>();

                // 遍历当前幻灯片中的所有形状
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    var shape = slide.Shapes[i];
                    if ((Office.MsoShapeType)shape.Type == selectedShapeType)
                    {
                        sameTypeShapeIndices.Add(i);
                    }
                }

                // 创建一个ShapeRange对象来包含所有相同类型的形状
                var shapeIndicesArray = sameTypeShapeIndices.ToArray();
                PowerPoint.ShapeRange shapeRange = slide.Shapes.Range(shapeIndicesArray);

                // 选中所有相同类型的形状
                shapeRange.Select();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("请选择一个对象。");
            }
        }

        private void Selectsize_Click(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = application.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                var selectedShape = selection.ShapeRange[1];
                var selectedShapeType = (Office.MsoShapeType)selectedShape.Type; // 获取选中形状的类型
                var slide = application.ActiveWindow.View.Slide;
                List<int> sameTypeShapeIndices = new List<int>();

                // 遍历当前幻灯片中的所有形状
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    var shape = slide.Shapes[i];
                    if ((Office.MsoShapeType)shape.Type == selectedShapeType)
                    {
                        sameTypeShapeIndices.Add(i);
                    }
                }

                // 创建一个ShapeRange对象来包含所有相同类型的形状
                var shapeIndicesArray = sameTypeShapeIndices.ToArray();
                PowerPoint.ShapeRange shapeRange = slide.Shapes.Range(shapeIndicesArray);

                // 选中所有相同类型的形状
                shapeRange.Select();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("请选择一个对象。");
            }
        }

        //按尺寸筛选
        private void Selectsize_Click1(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = application.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                var selectedShape = selection.ShapeRange[1];
                var selectedShapeWidth = selectedShape.Width; // 获取选中形状的宽度
                var selectedShapeHeight = selectedShape.Height; // 获取选中形状的高度
                var slide = application.ActiveWindow.View.Slide;
                List<int> sameSizeShapeIndices = new List<int>();

                // 遍历当前幻灯片中的所有形状
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    var shape = slide.Shapes[i];
                    if (shape.Width == selectedShapeWidth && shape.Height == selectedShapeHeight)
                    {
                        sameSizeShapeIndices.Add(i);
                    }
                }

                // 创建一个ShapeRange对象来包含所有相同尺寸的形状
                var shapeIndicesArray = sameSizeShapeIndices.ToArray();
                PowerPoint.ShapeRange shapeRange = slide.Shapes.Range(shapeIndicesArray);

                // 选中所有相同尺寸的形状
                shapeRange.Select();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("请选择一个对象。");
            }
        }

        //按颜色筛选
        private void SelectedColor_Click(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = application.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                var selectedShape = selection.ShapeRange[1];
                var selectedShapeFillColor = selectedShape.Fill.ForeColor.RGB; // 获取选中形状的填充颜色
                var slide = application.ActiveWindow.View.Slide;
                List<int> sameColorShapeIndices = new List<int>();

                // 遍历当前幻灯片中的所有形状
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    var shape = slide.Shapes[i];
                    if (shape.Fill.ForeColor.RGB == selectedShapeFillColor)
                    {
                        sameColorShapeIndices.Add(i);
                    }
                }

                // 创建一个ShapeRange对象来包含所有相同颜色的形状
                var shapeIndicesArray = sameColorShapeIndices.ToArray();
                PowerPoint.ShapeRange shapeRange = slide.Shapes.Range(shapeIndicesArray);

                // 选中所有相同颜色的形状
                shapeRange.Select();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("请选择一个对象。");
            }
        }

        //按轮廓筛选
        private void Selectedline_Click(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = application.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                var selectedShape = selection.ShapeRange[1];
                var slide = application.ActiveWindow.View.Slide;
                List<int> sameLineShapeIndices = new List<int>();

                if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
                {
                    // 获取选中形状的轮廓颜色
                    var selectedLineColor = selectedShape.Line.ForeColor.RGB;

                    // 遍历当前幻灯片中的所有形状
                    for (int i = 1; i <= slide.Shapes.Count; i++)
                    {
                        var shape = slide.Shapes[i];
                        if (shape.Line.ForeColor.RGB == selectedLineColor)
                        {
                            sameLineShapeIndices.Add(i);
                        }
                    }
                }
                else if ((Control.ModifierKeys & Keys.Shift) == Keys.Shift)
                {
                    // 获取选中形状的轮廓线条类型
                    var selectedLineDashStyle = selectedShape.Line.DashStyle;

                    // 遍历当前幻灯片中的所有形状
                    for (int i = 1; i <= slide.Shapes.Count; i++)
                    {
                        var shape = slide.Shapes[i];
                        if ((Office.MsoLineDashStyle)shape.Line.DashStyle == selectedLineDashStyle)
                        {
                            sameLineShapeIndices.Add(i);
                        }
                    }
                }
                else
                {
                    // 获取选中形状的轮廓宽度
                    var selectedLineWidth = selectedShape.Line.Weight;

                    // 遍历当前幻灯片中的所有形状
                    for (int i = 1; i <= slide.Shapes.Count; i++)
                    {
                        var shape = slide.Shapes[i];
                        if (shape.Line.Weight == selectedLineWidth)
                        {
                            sameLineShapeIndices.Add(i);
                        }
                    }
                }

                // 创建一个ShapeRange对象来包含所有相同轮廓属性的形状
                var shapeIndicesArray = sameLineShapeIndices.ToArray();
                PowerPoint.ShapeRange shapeRange = slide.Shapes.Range(shapeIndicesArray);

                // 选中所有相同轮廓属性的形状
                shapeRange.Select();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("请选择一个对象。");
            }
        }

        //按字号筛选
        private void Selectfontsize_Click(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = application.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                var selectedShape = selection.ShapeRange[1];
                var slide = application.ActiveWindow.View.Slide;
                List<int> sameFontSizeShapeIndices = new List<int>();

                // 获取选中形状的字体大小
                float selectedFontSize = 0;
                if (selectedShape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    selectedFontSize = selectedShape.TextFrame.TextRange.Font.Size;
                }

                // 如果选中形状没有字体大小，则不进行后续操作
                if (selectedFontSize == 0)
                {
                    System.Windows.Forms.MessageBox.Show("请选择一个包含文本的对象。");
                    return;
                }

                // 遍历当前幻灯片中的所有形状
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    var shape = slide.Shapes[i];
                    if ((Office.MsoTriState)shape.TextFrame.HasText == Office.MsoTriState.msoTrue &&
                        shape.TextFrame.TextRange.Font.Size == selectedFontSize)
                    {
                        sameFontSizeShapeIndices.Add(i);
                    }
                }

                // 创建一个ShapeRange对象来包含所有相同字体大小的形状
                var shapeIndicesArray = sameFontSizeShapeIndices.ToArray();
                PowerPoint.ShapeRange shapeRange = slide.Shapes.Range(shapeIndicesArray);

                // 选中所有相同字体大小的形状
                shapeRange.Select();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("请选择一个对象。");
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
                    if (Control.ModifierKeys == Keys.Control) // 如果按下了Ctrl键
                    {
                        rectangle.Fill.GradientAngle = 90; // 从上往下
                    }
                    else if (Control.ModifierKeys == Keys.Shift) // 如果按下了Shift键
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
                        if (Control.ModifierKeys == Keys.Control) // 如果按下了Ctrl键
                        {
                            rectangle.Fill.GradientAngle = 90; // 从上往下
                        }
                        else if (Control.ModifierKeys == Keys.Shift) // 如果按下了Shift键
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

        private void 交换位置_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 获取当前活动的PowerPoint应用程序
                var application = Globals.ThisAddIn.Application;

                // 获取当前选中的对象
                var selection = application.ActiveWindow.Selection;

                // 确保选中了两个对象
                if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count == 2)
                {
                    // 获取两个选中的形状
                    Shape shape1 = selection.ShapeRange[1];
                    Shape shape2 = selection.ShapeRange[2];

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
            catch (Exception ex)
            {
                MessageBox.Show("交换位置时发生错误: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                // 确保选中了两个对象
                if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count == 2)
                {
                    // 获取两个选中的形状
                    Shape shape1 = selection.ShapeRange[1];
                    Shape shape2 = selection.ShapeRange[2];

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

        private void 交换格式_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 获取当前活动的PowerPoint应用程序
                var application = Globals.ThisAddIn.Application;

                // 获取当前选中的对象
                var selection = application.ActiveWindow.Selection;

                // 确保选中了两个对象
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count == 2)
                {
                    // 获取两个选中的形状
                    PowerPoint.Shape shape1 = selection.ShapeRange[1];
                    PowerPoint.Shape shape2 = selection.ShapeRange[2];

                    // 交换填充格式
                    SwapFill(shape1.Fill, shape2.Fill);

                    // 交换线条格式
                    SwapLine(shape1.Line, shape2.Line);

                    // 交换阴影格式
                    SwapShadow(shape1.Shadow, shape2.Shadow);

                    // 交换文本框格式
                    SwapTextFrame(shape1.TextFrame, shape2.TextFrame);

                    // 交换三维格式
                    SwapThreeD(shape1.ThreeD, shape2.ThreeD);
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

        private void SwapFill(PowerPoint.FillFormat fill1, PowerPoint.FillFormat fill2)
        {
            // 使用临时变量交换填充格式
            var tempBackColor = fill1.BackColor.RGB;
            fill1.BackColor.RGB = fill2.BackColor.RGB;
            fill2.BackColor.RGB = tempBackColor;

            var tempForeColor = fill1.ForeColor.RGB;
            fill1.ForeColor.RGB = fill2.ForeColor.RGB;
            fill2.ForeColor.RGB = tempForeColor;

            var tempTransparency = fill1.Transparency;
            fill1.Transparency = fill2.Transparency;
            fill2.Transparency = tempTransparency;

            var tempVisible = fill1.Visible;
            fill1.Visible = fill2.Visible;
            fill2.Visible = tempVisible;
        }

        private void SwapLine(PowerPoint.LineFormat line1, PowerPoint.LineFormat line2)
        {
            // 使用临时变量交换线条格式
            var tempForeColor = line1.ForeColor.RGB;
            line1.ForeColor.RGB = line2.ForeColor.RGB;
            line2.ForeColor.RGB = tempForeColor;

            var tempWeight = line1.Weight;
            line1.Weight = line2.Weight;
            line2.Weight = tempWeight;

            var tempDashStyle = line1.DashStyle;
            line1.DashStyle = line2.DashStyle;
            line2.DashStyle = tempDashStyle;

            var tempVisible = line1.Visible;
            line1.Visible = line2.Visible;
            line2.Visible = tempVisible;
        }

        private void SwapShadow(PowerPoint.ShadowFormat shadow1, PowerPoint.ShadowFormat shadow2)
        {
            // 使用临时变量交换阴影格式
            var tempForeColor = shadow1.ForeColor.RGB;
            shadow1.ForeColor.RGB = shadow2.ForeColor.RGB;
            shadow2.ForeColor.RGB = tempForeColor;

            var tempObscured = shadow1.Obscured;
            shadow1.Obscured = shadow2.Obscured;
            shadow2.Obscured = tempObscured;

            var tempOffsetX = shadow1.OffsetX;
            shadow1.OffsetX = shadow2.OffsetX;
            shadow2.OffsetX = tempOffsetX;

            var tempOffsetY = shadow1.OffsetY;
            shadow1.OffsetY = shadow2.OffsetY;
            shadow2.OffsetY = tempOffsetY;

            var tempTransparency = shadow1.Transparency;
            shadow1.Transparency = shadow2.Transparency;
            shadow2.Transparency = tempTransparency;

            var tempVisible = shadow1.Visible;
            shadow1.Visible = shadow2.Visible;
            shadow2.Visible = tempVisible;
        }

        private void SwapTextFrame(PowerPoint.TextFrame textFrame1, PowerPoint.TextFrame textFrame2)
        {
            // 使用临时变量交换文本框格式
            var tempMarginBottom = textFrame1.MarginBottom;
            textFrame1.MarginBottom = textFrame2.MarginBottom;
            textFrame2.MarginBottom = tempMarginBottom;

            var tempMarginLeft = textFrame1.MarginLeft;
            textFrame1.MarginLeft = textFrame2.MarginLeft;
            textFrame2.MarginLeft = tempMarginLeft;

            var tempMarginRight = textFrame1.MarginRight;
            textFrame1.MarginRight = textFrame2.MarginRight;
            textFrame2.MarginRight = tempMarginRight;

            var tempMarginTop = textFrame1.MarginTop;
            textFrame1.MarginTop = textFrame2.MarginTop;
            textFrame2.MarginTop = tempMarginTop;

            var tempVerticalAnchor = textFrame1.VerticalAnchor;
            textFrame1.VerticalAnchor = textFrame2.VerticalAnchor;
            textFrame2.VerticalAnchor = tempVerticalAnchor;

            var tempWordWrap = textFrame1.WordWrap;
            textFrame1.WordWrap = textFrame2.WordWrap;
            textFrame2.WordWrap = tempWordWrap;
        }

        private void SwapThreeD(PowerPoint.ThreeDFormat threeD1, PowerPoint.ThreeDFormat threeD2)
        {
            // 使用临时变量交换三维格式
            var tempDepth = threeD1.Depth;
            threeD1.Depth = threeD2.Depth;
            threeD2.Depth = tempDepth;

            // ExtrusionColor is read-only, so we cannot swap it directly

            var tempPresetMaterial = threeD1.PresetMaterial;
            threeD1.PresetMaterial = threeD2.PresetMaterial;
            threeD2.PresetMaterial = tempPresetMaterial;

            var tempPresetLighting = threeD1.PresetLighting;
            threeD1.PresetLighting = threeD2.PresetLighting;
            threeD2.PresetLighting = tempPresetLighting;

            var tempVisible = threeD1.Visible;
            threeD1.Visible = threeD2.Visible;
            threeD2.Visible = tempVisible;
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
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count == 2)
                {
                    // 获取两个选中的形状
                    PowerPoint.Shape shape1 = selection.ShapeRange[1];
                    PowerPoint.Shape shape2 = selection.ShapeRange[2];

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
            catch (Exception ex)
            {
                MessageBox.Show("交换尺寸时发生错误: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ResizeAndCropShape(PowerPoint.Shape shape, float targetAspectRatio, float targetWidth, float targetHeight)
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
            bool isCtrlPressed = (Control.ModifierKeys & Keys.Control) == Keys.Control;

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

     

        private void 统一大小_Click(object sender, RibbonControlEventArgs e)
        {
           
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
                        float height = shape.Height / paragraphCount;

                        for (int i = 1; i <= paragraphCount; i++)
                        {
                            PowerPoint.TextRange paragraph = textRange.Paragraphs(i);
                            PowerPoint.Shape newShape = shape.Duplicate()[1];
                            newShape.Left = left;
                            newShape.Top = top + (i - 1) * height;
                            newShape.Width = width;
                            newShape.Height = height;
                            newShape.TextFrame.TextRange.Text = paragraph.Text;

                            // 删除空白的文本框
                            if (string.IsNullOrWhiteSpace(paragraph.Text))
                            {
                                newShape.Delete();
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

        private void 统一大小_Click_1(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            var presentation = application.ActivePresentation;
            var slide = application.ActiveWindow.View.Slide;

            if (application.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionShapes)
            {
                var selectedShapes = application.ActiveWindow.Selection.ShapeRange;

                if (selectedShapes.Count > 1)
                {
                    // 获取第一个选中的对象的宽度和高度
                    float targetWidth = selectedShapes[1].Width;
                    float targetHeight = selectedShapes[1].Height;

                    // 遍历后续被选中的对象
                    for (int i = 2; i <= selectedShapes.Count; i++)
                    {
                        Shape shape = selectedShapes[i];

                        if (shape.Type == MsoShapeType.msoPicture)
                        {
                            // 处理透明背景图片
                            HandleTransparentImage(shape, targetWidth, targetHeight);
                        }
                        else
                        {
                            // 统一大小
                            shape.Width = targetWidth;
                            shape.Height = targetHeight;
                        }
                    }
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

        private void HandleTransparentImage(Shape shape, float targetWidth, float targetHeight)
        {
            // 为透明背景图片添加一个临时背景
            shape.Fill.Solid();
            shape.Fill.ForeColor.RGB = Color.White.ToArgb();
            shape.Fill.Transparency = 0.0f;

            // 裁剪和调整大小
            CropAndResizePicture(shape, targetWidth, targetHeight);

            // 去掉临时背景
            shape.Fill.Transparency = 1.0f;
        }

        private void CropAndResizePicture(Shape shape, float targetWidth, float targetHeight)
        {
            // 获取图片的原始宽高
            float originalWidth = shape.Width;
            float originalHeight = shape.Height;

            // 计算目标宽高比
            float targetRatio = targetWidth / targetHeight;

            // 计算裁剪区域，使其匹配目标宽高比
            float cropWidth = originalWidth;
            float cropHeight = originalHeight;

            if (originalWidth / originalHeight > targetRatio)
            {
                // 如果宽高比过大，需要裁剪宽度
                cropWidth = originalHeight * targetRatio;
            }
            else
            {
                // 如果宽高比过小，需要裁剪高度
                cropHeight = originalWidth / targetRatio;
            }

            // 计算裁剪区域的左上角坐标
            float cropLeft = (originalWidth - cropWidth) / 2;
            float cropTop = (originalHeight - cropHeight) / 2;

            // 设置裁剪
            shape.PictureFormat.CropLeft = cropLeft;
            shape.PictureFormat.CropRight = originalWidth - cropWidth - cropLeft;
            shape.PictureFormat.CropTop = cropTop;
            shape.PictureFormat.CropBottom = originalHeight - cropHeight - cropTop;

            // 调整大小，确保保持比例
            shape.LockAspectRatio = MsoTriState.msoTrue;
            shape.Width = targetWidth;
            shape.Height = targetHeight;

            // 修正最终尺寸，确保与目标尺寸一致
            shape.LockAspectRatio = MsoTriState.msoFalse;
            shape.Width = targetWidth;
            shape.Height = targetHeight;
        }

        private void 统一格式_Click(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            var presentation = application.ActivePresentation;
            var slide = application.ActiveWindow.View.Slide;

            if (application.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionShapes)
            {
                var selectedShapes = application.ActiveWindow.Selection.ShapeRange;

                if (selectedShapes.Count > 1)
                {
                    Shape baseShape = selectedShapes[1];

                    // 遍历后续被选中的对象
                    for (int i = 2; i <= selectedShapes.Count; i++)
                    {
                        Shape shape = selectedShapes[i];
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

        private void 智能缩放_Click(object sender, RibbonControlEventArgs e)
        {
            SmartScalingForm scalingForm = new SmartScalingForm();
            scalingForm.Show();
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


       
        private void 打包文档_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 获取当前演示文稿
                Microsoft.Office.Interop.PowerPoint.Application app = Globals.ThisAddIn.Application;
                Presentation presentation = app.ActivePresentation;

                // 获取演示文稿名称
                string presentationName = Path.GetFileNameWithoutExtension(presentation.FullName);

                // 创建文件夹路径
                string folderPath = Path.Combine("C:\\", presentationName);

                // 创建文件夹
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }

                // 保存演示文稿
                string presentationPath = Path.Combine(folderPath, presentationName + ".pptx");
                presentation.SaveCopyAs(presentationPath);

                // 创建“文档所用字体”子文件夹
                string fontsFolderPath = Path.Combine(folderPath, "文档所用字体");
                if (!Directory.Exists(fontsFolderPath))
                {
                    Directory.CreateDirectory(fontsFolderPath);
                }

                // 获取演示文稿所用的所有字体
                for (int i = 1; i <= presentation.Fonts.Count; i++)
                {
                    Microsoft.Office.Interop.PowerPoint.Font font = presentation.Fonts[i];
                    string fontFilePath = GetFontFilePath(font.Name);

                    if (!string.IsNullOrEmpty(fontFilePath))
                    {
                        try
                        {
                            // 复制字体文件到“文档所用字体”子文件夹
                            string destFontPath = Path.Combine(fontsFolderPath, Path.GetFileName(fontFilePath));
                            File.Copy(fontFilePath, destFontPath, true);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"字体 {font.Name} 复制失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show($"未找到字体文件: {font.Name}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }

                MessageBox.Show("打包完成！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("打包过程中出错: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetFontFilePath(string fontName)
        {
            // 在注册表中查找字体文件路径
            string fontFilePath = FindFontFilePathInRegistry(fontName, Registry.LocalMachine);
            if (!string.IsNullOrEmpty(fontFilePath))
            {
                return fontFilePath;
            }

            // 在用户注册表中查找字体文件路径
            fontFilePath = FindFontFilePathInRegistry(fontName, Registry.CurrentUser);
            if (!string.IsNullOrEmpty(fontFilePath))
            {
                return fontFilePath;
            }

            // 在系统字体文件夹中查找
            fontFilePath = FindFontFilePathInDirectory(fontName, Environment.GetFolderPath(Environment.SpecialFolder.Fonts));
            if (!string.IsNullOrEmpty(fontFilePath))
            {
                return fontFilePath;
            }

            // 在用户字体文件夹中查找（如果有）
            string userFontDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft\\Windows\\Fonts");
            fontFilePath = FindFontFilePathInDirectory(fontName, userFontDir);
            if (!string.IsNullOrEmpty(fontFilePath))
            {
                return fontFilePath;
            }

            return null;
        }

        private string FindFontFilePathInRegistry(string fontName, RegistryKey registryKey)
        {
            string fontFilePath = null;
            string fontsRegistryPath = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts";

            using (RegistryKey key = registryKey.OpenSubKey(fontsRegistryPath, false))
            {
                if (key != null)
                {
                    foreach (string fontRegName in key.GetValueNames())
                    {
                        if (CultureInfo.CurrentCulture.CompareInfo.IndexOf(fontRegName, fontName, CompareOptions.IgnoreCase) >= 0)
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
                    if (CultureInfo.CurrentCulture.CompareInfo.IndexOf(Path.GetFileNameWithoutExtension(fontFile), fontName, CompareOptions.IgnoreCase) >= 0)
                    {
                        return fontFile;
                    }
                }
            }
            return null;
        }
    }
}













