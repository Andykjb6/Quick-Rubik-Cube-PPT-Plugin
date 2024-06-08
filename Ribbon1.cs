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

        private void button14_Click(object sender, RibbonControlEventArgs e)
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

                // 初始间距设置为10像素
                float initialSpacing = 10f;

                // 计算每个形状的新位置
                ArrangeShapes(selectedShapes, firstLeft, initialSpacing);

                // 创建并显示窗体
                Form form = new Form();
                form.Text = "水平分布间距";
                form.Width = 520;
                form.Height = 180;
                form.StartPosition = FormStartPosition.CenterScreen;

                // 添加滑块控件
                TrackBar trackBar = new TrackBar();
                trackBar.Location = new System.Drawing.Point(20, 40);
                trackBar.Size = new System.Drawing.Size(480, 50);
                trackBar.Minimum = 0;
                trackBar.Maximum = 200;
                trackBar.Value = (int)initialSpacing;
                trackBar.LargeChange = 10;
                trackBar.SmallChange = 1;
                trackBar.TickStyle = TickStyle.TopLeft;
                trackBar.TickFrequency = 10;
                trackBar.Dock = DockStyle.Top;

                trackBar.ValueChanged += (s, ev) =>
                {
                    float spacing = trackBar.Value; // 将滑块值转换为间距值
                    ArrangeShapes(selectedShapes, firstLeft, spacing);
                };

                // 将滑块控件添加到窗体
                form.Controls.Add(trackBar);

                // 显示窗体
                form.ShowDialog();
            }
            else
            {
                MessageBox.Show("请选择至少两个形状。");
            }
        }

        /// <summary>
        /// 以指定间距水平排列形状。
        /// </summary>
        /// <param name="shapes">要排列的形状集合</param>
        /// <param name="startLeft">起始左坐标</param>
        /// <param name="spacing">间距</param>
        private void ArrangeShapes(PowerPoint.ShapeRange shapes, float startLeft, float spacing)
        {
            float currentLeft = startLeft;

            // 跳过第一个形状，从第二个形状开始排列
            for (int i = 2; i <= shapes.Count; i++)
            {
                PowerPoint.Shape shape = shapes[i];
                currentLeft += shapes[i - 1].Width + spacing;
                shape.Left = currentLeft;
                shape.Top = shapes[1].Top; // 保持所有形状的垂直对齐
            }
        }

        private void button15_Click(object sender, RibbonControlEventArgs e)
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
                        MessageBox.Show("所有形状必须在同一个幻灯片上。");
                        return;
                    }
                }

                // 获取第一个被选中的对象的位置
                PowerPoint.Shape firstShape = selectedShapes[1];
                float firstLeft = firstShape.Left;
                float firstTop = firstShape.Top;

                // 计算每个形状的新位置
                float currentTop = firstTop;
                float spacing = 10; // 初始间距为10

                foreach (PowerPoint.Shape shape in selectedShapes)
                {
                    shape.Left = firstLeft;
                    shape.Top = currentTop;
                    currentTop += shape.Height + spacing; // 从上到下排列，保持一定间距
                }

                // 创建并显示窗体
                Form form = new Form();
                form.Text = "垂直分布间距";
                form.Width = 520;
                form.Height = 200;
                form.StartPosition = FormStartPosition.CenterScreen;

                // 添加滑块控件
                TrackBar trackBar = new TrackBar();
                trackBar.Location = new System.Drawing.Point(25, 5);
                trackBar.Size = new System.Drawing.Size(440, 30);
                trackBar.Minimum = 0;
                trackBar.Maximum = 100;
                trackBar.Value = 10;
                trackBar.LargeChange = 10;
                trackBar.SmallChange = 1;
                trackBar.TickStyle = TickStyle.BottomRight;

                trackBar.ValueChanged += (s, ev) =>
                {
                    spacing = trackBar.Value;
                    float top = firstTop;
                    foreach (PowerPoint.Shape shape in selectedShapes)
                    {
                        shape.Top = top;
                        top += shape.Height + spacing; // 计算下一个形状的位置，包括间距
                    }
                };

                // 将滑块控件添加到窗体
                form.Controls.Add(trackBar);

                // 显示窗体
                form.ShowDialog();
            }
            else
            {
                MessageBox.Show("请选择至少两个形状。");
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
                MessageBox.Show("Please select one or more shapes with images.");
                return;
            }

            var selectedShapes = activeWindow.Selection.ShapeRange.Cast<Shape>().ToList();

            if (!selectedShapes.Any())
            {
                MessageBox.Show("Please select one or more shapes with images.");
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
            catch (Exception ex)
            {
                MessageBox.Show($"发生错误: {ex.Message}", "课件帮PPT助手", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        public class TimerForm : System.Windows.Forms.Form
        {
            private System.Windows.Forms.TextBox timeTextBox;
            private System.Windows.Forms.Button startButton;
            private System.Windows.Forms.Button stopButton;
            private System.Windows.Forms.Button resetButton;
            private System.Windows.Forms.Button closeButton;
            private System.Windows.Forms.Button settingsButton;
            private System.Windows.Forms.Button darkModeButton;
            private System.Windows.Forms.Timer timer;
            private DateTime targetTime;
            private bool isCountdown = true; // 默认倒计时
            private System.Drawing.Font currentFont = new System.Drawing.Font("Arial", 40, System.Drawing.FontStyle.Bold);
            private System.Drawing.Color backgroundColor = System.Drawing.Color.White;
            private System.Drawing.Color timeTextColor = System.Drawing.Color.Black;
            private System.Drawing.Color otherTextColor = System.Drawing.Color.Black;
            private System.Drawing.Color darkModeButtonColor = System.Drawing.Color.LightGray;

            public TimerForm()
            {
                InitializeComponents();
            }

            private void InitializeComponents()
            {
                this.Text = "计时器";
                this.Size = new System.Drawing.Size(520, 250);
                this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None; // 无边框
                this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                this.BackColor = backgroundColor;
                this.TopMost = true; // 窗口置顶

                // 绘制边框
                this.Paint += TimerForm_Paint;

                // 时间输入文本框
                timeTextBox = new System.Windows.Forms.TextBox
                {
                    Text = "00:00:00",
                    Font = currentFont,
                    TextAlign = System.Windows.Forms.HorizontalAlignment.Center,
                    Dock = System.Windows.Forms.DockStyle.None,
                    Height = 100,
                    Width = 455,
                    BorderStyle = System.Windows.Forms.BorderStyle.None,
                    BackColor = backgroundColor, // 同步背景色
                    ForeColor = timeTextColor,
                    ReadOnly = false // 可编辑
                };
                timeTextBox.Location = new System.Drawing.Point(30, 30);
                this.Controls.Add(timeTextBox);

                // 设置按钮
                settingsButton = new System.Windows.Forms.Button
                {
                    Text = "⚙",
                    Width = 35,
                    Height = 35,
                    FlatStyle = System.Windows.Forms.FlatStyle.Flat,
                    Location = new System.Drawing.Point(5, 5)
                };
                settingsButton.FlatAppearance.BorderSize = 0; // 去掉边框
                settingsButton.Click += SettingsButton_Click;
                this.Controls.Add(settingsButton);
                settingsButton.BringToFront(); // 将设置按钮置于顶层

                // 关闭按钮
                closeButton = new System.Windows.Forms.Button
                {
                    Text = "✖",
                    Width = 35,
                    Height = 35,
                    FlatStyle = System.Windows.Forms.FlatStyle.Flat,
                    Location = new System.Drawing.Point(this.Width - 62, 2),
                    Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right
                };
                closeButton.FlatAppearance.BorderSize = 0; // 去掉边框
                closeButton.Click += CloseButton_Click;
                this.Controls.Add(closeButton);
                closeButton.BringToFront(); // 将关闭按钮置于顶层

                // 开始按钮
                startButton = new System.Windows.Forms.Button
                {
                    Text = "▶",
                    Width = 35,
                    Height = 35,
                    FlatStyle = System.Windows.Forms.FlatStyle.Flat
                };
                startButton.Click += StartButton_Click;
                this.Controls.Add(startButton);
                startButton.BringToFront(); // 将开始按钮置于顶层

                // 暂停按钮
                stopButton = new System.Windows.Forms.Button
                {
                    Text = "⏸",
                    Width = 35,
                    Height = 35,
                    FlatStyle = System.Windows.Forms.FlatStyle.Flat,
                    Enabled = false
                };
                stopButton.Click += StopButton_Click;
                this.Controls.Add(stopButton);
                stopButton.BringToFront(); // 将暂停按钮置于顶层

                // 重置按钮
                resetButton = new System.Windows.Forms.Button
                {
                    Text = "⟳",
                    Width = 35,
                    Height = 35,
                    FlatStyle = System.Windows.Forms.FlatStyle.Flat,
                    Enabled = false
                };
                resetButton.Click += ResetButton_Click;
                this.Controls.Add(resetButton);
                resetButton.BringToFront(); // 将重置按钮置于顶层

                // 按钮布局
                System.Windows.Forms.FlowLayoutPanel buttonPanel = new System.Windows.Forms.FlowLayoutPanel
                {
                    Dock = System.Windows.Forms.DockStyle.None,
                    FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                    Height = 40,
                    Padding = new System.Windows.Forms.Padding(0),
                    Width = 150,
                    Location = new System.Drawing.Point((this.Width - 120) / 2, 165) // 居中
                };
                buttonPanel.Controls.Add(startButton);
                buttonPanel.Controls.Add(stopButton);
                buttonPanel.Controls.Add(resetButton);
                this.Controls.Add(buttonPanel);
                buttonPanel.BringToFront(); // 将 buttonPanel 置于顶层

                // 暗色模式按钮
                darkModeButton = new System.Windows.Forms.Button
                {
                    Text = "🌙 暗色模式",
                    Width = 100,
                    Height = 35,
                    FlatStyle = System.Windows.Forms.FlatStyle.Flat,
                    BackColor = darkModeButtonColor,
                    Dock = System.Windows.Forms.DockStyle.Bottom
                };
                darkModeButton.FlatAppearance.BorderSize = 0; // 去掉边框
                darkModeButton.Click += DarkModeButton_Click;
                this.Controls.Add(darkModeButton);

                // 定时器
                timer = new System.Windows.Forms.Timer
                {
                    Interval = 1000
                };
                timer.Tick += Timer_Tick;

                // 使窗口支持拖动
                this.MouseDown += (s, e) =>
                {
                    if (e.Button == System.Windows.Forms.MouseButtons.Left)
                    {
                        this.Capture = false;
                        System.Windows.Forms.Message m = System.Windows.Forms.Message.Create(this.Handle, 0xA1, new System.IntPtr(2), System.IntPtr.Zero);
                        this.WndProc(ref m);
                    }
                };
            }

            private void TimerForm_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
            {
                System.Windows.Forms.ControlPaint.DrawBorder(e.Graphics, this.ClientRectangle, System.Drawing.Color.LightGray, System.Windows.Forms.ButtonBorderStyle.Solid);
            }

            private void StartButton_Click(object sender, System.EventArgs e)
            {
                if (System.TimeSpan.TryParseExact(timeTextBox.Text, @"hh\:mm\:ss", null, out System.TimeSpan timeSpan))
                {
                    if (isCountdown)
                    {
                        targetTime = System.DateTime.Now.Add(timeSpan);
                    }
                    else
                    {
                        targetTime = System.DateTime.Now.AddHours(-timeSpan.TotalHours).AddMinutes(-timeSpan.TotalMinutes).AddSeconds(-timeSpan.TotalSeconds);
                    }

                    startButton.Enabled = false;
                    stopButton.Enabled = true;
                    resetButton.Enabled = true;
                    timeTextBox.ReadOnly = true;

                    timer.Start();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("请输入有效的时间格式（hh:mm:ss）。");
                }
            }

            private void StopButton_Click(object sender, System.EventArgs e)
            {
                timer.Stop();
                startButton.Enabled = true;
                stopButton.Enabled = false;
                resetButton.Enabled = true;
            }

            private void ResetButton_Click(object sender, System.EventArgs e)
            {
                timer.Stop();
                timeTextBox.Text = "00:00:00";
                startButton.Enabled = true;
                stopButton.Enabled = false;
                resetButton.Enabled = false;
                timeTextBox.ReadOnly = false;
            }

            private void CloseButton_Click(object sender, System.EventArgs e)
            {
                this.Close();
            }

            private void SettingsButton_Click(object sender, System.EventArgs e)
            {
                this.Hide(); // 隐藏计时器窗口
                using (SettingsForm settingsForm = new SettingsForm(currentFont, timeTextColor, otherTextColor, isCountdown, backgroundColor, darkModeButtonColor))
                {
                    if (settingsForm.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        currentFont = settingsForm.SelectedFont;
                        timeTextColor = settingsForm.TimeTextColor;
                        otherTextColor = settingsForm.OtherTextColor;
                        isCountdown = settingsForm.IsCountdown;
                        backgroundColor = settingsForm.BackgroundColor;
                        darkModeButtonColor = settingsForm.DarkModeButtonColor;

                        timeTextBox.Font = currentFont;
                        timeTextBox.ForeColor = timeTextColor;
                        this.BackColor = backgroundColor;
                        timeTextBox.BackColor = backgroundColor; // 同步背景色
                        darkModeButton.BackColor = darkModeButtonColor;
                    }
                }
                this.Show(); // 显示计时器窗口
            }

            private void DarkModeButton_Click(object sender, System.EventArgs e)
            {
                if (this.BackColor == System.Drawing.Color.White)
                {
                    // 切换到深色模式
                    this.BackColor = System.Drawing.Color.Black;
                    timeTextBox.BackColor = System.Drawing.Color.Black;
                    timeTextBox.ForeColor = System.Drawing.Color.White;
                    darkModeButton.Text = "浅色模式";
                    darkModeButton.BackColor = System.Drawing.Color.Gray;

                    // 设置按钮颜色
                    SetButtonColors(System.Drawing.Color.White, System.Drawing.Color.Black);

                    // 设置 设置按钮和关闭按钮颜色
                    settingsButton.BackColor = System.Drawing.Color.Black;
                    settingsButton.ForeColor = System.Drawing.Color.White;
                    closeButton.BackColor = System.Drawing.Color.Black;
                    closeButton.ForeColor = System.Drawing.Color.White;
                }
                else
                {
                    // 切换到浅色模式
                    this.BackColor = System.Drawing.Color.White;
                    timeTextBox.BackColor = System.Drawing.Color.White;
                    timeTextBox.ForeColor = System.Drawing.Color.Black;
                    darkModeButton.Text = "暗色模式";
                    darkModeButton.BackColor = System.Drawing.Color.LightGray;

                    // 设置按钮颜色
                    SetButtonColors(System.Drawing.Color.White, System.Drawing.Color.Black);
                }
            }

            private void SetButtonColors(System.Drawing.Color backColor, System.Drawing.Color foreColor)
            {
                startButton.BackColor = backColor;
                startButton.ForeColor = foreColor;
                stopButton.BackColor = backColor;
                stopButton.ForeColor = foreColor;
                resetButton.BackColor = backColor;
                resetButton.ForeColor = foreColor;
                closeButton.BackColor = backColor;
                closeButton.ForeColor = foreColor;
                settingsButton.BackColor = backColor;
                settingsButton.ForeColor = foreColor;
            }

            private void Timer_Tick(object sender, System.EventArgs e)
            {
                System.TimeSpan remainingTime = isCountdown ? targetTime - System.DateTime.Now : System.DateTime.Now - targetTime;

                if (remainingTime.TotalSeconds <= 0)
                {
                    timer.Stop();
                    timeTextBox.Text = "00:00:00";
                    System.Windows.Forms.MessageBox.Show("时间到！");
                    startButton.Enabled = true;
                    stopButton.Enabled = false;
                    resetButton.Enabled = true;
                    timeTextBox.ReadOnly = false;
                    // 播放音效
                    System.Media.SystemSounds.Exclamation.Play();
                }
                else
                {
                    timeTextBox.Text = remainingTime.ToString(@"hh\:mm\:ss");
                }
            }
        }

        public class SettingsForm : System.Windows.Forms.Form
        {
            public System.Drawing.Font SelectedFont { get; private set; }
            public System.Drawing.Color TimeTextColor { get; private set; }
            public System.Drawing.Color OtherTextColor { get; private set; }
            public bool IsCountdown { get; private set; }
            public System.Drawing.Color BackgroundColor { get; private set; }
            public System.Drawing.Color DarkModeButtonColor { get; private set; }

            private System.Windows.Forms.FontDialog fontDialog;
            private System.Windows.Forms.ColorDialog colorDialog;
            private System.Windows.Forms.Label timeFontLabel;
            private System.Windows.Forms.Label timeTextColorLabel;
            private System.Windows.Forms.Label otherTextColorLabel;
            private System.Windows.Forms.Label countdownLabel;
            private System.Windows.Forms.Label backgroundColorLabel;
            private System.Windows.Forms.Label darkModeButtonColorLabel;
            private System.Windows.Forms.ComboBox fontComboBox;
            private System.Windows.Forms.TextBox timeTextColorBox;
            private System.Windows.Forms.TextBox otherTextColorBox;
            private System.Windows.Forms.RadioButton countdownRadioButton;
            private System.Windows.Forms.RadioButton stopwatchRadioButton;
            private System.Windows.Forms.TextBox backgroundColorBox;
            private System.Windows.Forms.TextBox darkModeButtonColorBox;
            private System.Windows.Forms.Button timeTextColorButton;
            private System.Windows.Forms.Button otherTextColorButton;
            private System.Windows.Forms.Button backgroundColorButton;
            private System.Windows.Forms.Button darkModeButtonColorButton;
            private System.Windows.Forms.Button okButton;
            private System.Windows.Forms.Button cancelButton;

            public SettingsForm(System.Drawing.Font currentFont, System.Drawing.Color timeTextColor, System.Drawing.Color otherTextColor, bool isCountdown, System.Drawing.Color backgroundColor, System.Drawing.Color darkModeButtonColor)
            {
                SelectedFont = currentFont;
                TimeTextColor = timeTextColor;
                OtherTextColor = otherTextColor;
                IsCountdown = isCountdown;
                BackgroundColor = backgroundColor;
                DarkModeButtonColor = darkModeButtonColor;

                InitializeComponents();
            }

            private void InitializeComponents()
            {
                this.Text = "设置面板";
                this.Size = new System.Drawing.Size(480, 500);
                this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;

                fontDialog = new System.Windows.Forms.FontDialog
                {
                    Font = SelectedFont
                };

                colorDialog = new System.Windows.Forms.ColorDialog();

                // 时间显示字体标签
                timeFontLabel = new System.Windows.Forms.Label
                {
                    Text = "时钟字体：",
                    Location = new System.Drawing.Point(10, 20),
                    AutoSize = true
                };
                this.Controls.Add(timeFontLabel);

                // 时间显示字体选择
                fontComboBox = new System.Windows.Forms.ComboBox
                {
                    Location = new System.Drawing.Point(150, 20),
                    Width = 150
                };
                foreach (System.Drawing.FontFamily font in System.Drawing.FontFamily.Families)
                {
                    fontComboBox.Items.Add(font.Name);
                }
                fontComboBox.SelectedItem = SelectedFont.Name;
                this.Controls.Add(fontComboBox);

                // 时间显示字体颜色标签
                timeTextColorLabel = new System.Windows.Forms.Label
                {
                    Text = "时钟字体颜色：",
                    Location = new System.Drawing.Point(10, 80),
                    AutoSize = true
                };
                this.Controls.Add(timeTextColorLabel);

                // 时间显示字体颜色选择
                timeTextColorBox = new System.Windows.Forms.TextBox
                {
                    Text = TimeTextColor.ToArgb().ToString("X"),
                    Location = new System.Drawing.Point(190, 80),
                    Width = 100,
                    ReadOnly = true
                };
                this.Controls.Add(timeTextColorBox);

                timeTextColorButton = new System.Windows.Forms.Button
                {
                    Text = "选择颜色",
                    Location = new System.Drawing.Point(310, 80),
                    Width = 130,
                    Height = 40
                };
                timeTextColorButton.Click += TimeTextColorButton_Click;
                this.Controls.Add(timeTextColorButton);

                // 其他字体颜色标签
                otherTextColorLabel = new System.Windows.Forms.Label
                {
                    Text = "其他字体颜色：",
                    Location = new System.Drawing.Point(10, 130),
                    AutoSize = true
                };
                this.Controls.Add(otherTextColorLabel);

                // 其他字体颜色选择
                otherTextColorBox = new System.Windows.Forms.TextBox
                {
                    Text = OtherTextColor.ToArgb().ToString("X"),
                    Location = new System.Drawing.Point(190, 130),
                    Width = 100,
                    ReadOnly = true
                };
                this.Controls.Add(otherTextColorBox);

                otherTextColorButton = new System.Windows.Forms.Button
                {
                    Text = "选择颜色",
                    Location = new System.Drawing.Point(310, 130),
                    Width = 130,
                    Height = 40
                };
                otherTextColorButton.Click += OtherTextColorButton_Click;
                this.Controls.Add(otherTextColorButton);

                // 计时模式标签
                countdownLabel = new System.Windows.Forms.Label
                {
                    Text = "计时模式：",
                    Location = new System.Drawing.Point(10, 180),
                    AutoSize = true
                };
                this.Controls.Add(countdownLabel);

                // 计时模式选择
                countdownRadioButton = new System.Windows.Forms.RadioButton
                {
                    Text = "倒计时",
                    Location = new System.Drawing.Point(150, 180),
                    Checked = IsCountdown,
                    Height = 40,
                    Width = 120
                };
                stopwatchRadioButton = new System.Windows.Forms.RadioButton
                {
                    Text = "顺计时",
                    Location = new System.Drawing.Point(285, 180),
                    Checked = !IsCountdown,
                    Height = 40,
                    Width = 120
                };
                this.Controls.Add(countdownRadioButton);
                this.Controls.Add(stopwatchRadioButton);

                // 背景颜色标签
                backgroundColorLabel = new System.Windows.Forms.Label
                {
                    Text = "背景颜色：",
                    Location = new System.Drawing.Point(10, 230),
                    AutoSize = true
                };
                this.Controls.Add(backgroundColorLabel);

                // 背景颜色选择
                backgroundColorBox = new System.Windows.Forms.TextBox
                {
                    Text = BackgroundColor.ToArgb().ToString("X"),
                    Location = new System.Drawing.Point(190, 230),
                    Width = 100,
                    ReadOnly = true
                };
                this.Controls.Add(backgroundColorBox);

                backgroundColorButton = new System.Windows.Forms.Button
                {
                    Text = "选择颜色",
                    Location = new System.Drawing.Point(310, 230),
                    Width = 130,
                    Height = 40
                };
                backgroundColorButton.Click += BackgroundColorButton_Click;
                this.Controls.Add(backgroundColorButton);

                // 深/浅模式按钮颜色标签
                darkModeButtonColorLabel = new System.Windows.Forms.Label
                {
                    Text = "深/浅模式：",
                    Location = new System.Drawing.Point(10, 280),
                    AutoSize = true
                };
                this.Controls.Add(darkModeButtonColorLabel);

                // 深/浅模式按钮颜色选择
                darkModeButtonColorBox = new System.Windows.Forms.TextBox
                {
                    Text = DarkModeButtonColor.ToArgb().ToString("X"),
                    Location = new System.Drawing.Point(190, 280),
                    Width = 100,
                    ReadOnly = true
                };
                this.Controls.Add(darkModeButtonColorBox);

                darkModeButtonColorButton = new System.Windows.Forms.Button
                {
                    Text = "选择颜色",
                    Location = new System.Drawing.Point(310, 280),
                    Width = 130,
                    Height = 40
                };
                darkModeButtonColorButton.Click += DarkModeButtonColorButton_Click;
                this.Controls.Add(darkModeButtonColorButton);

                // 确定和取消按钮
                okButton = new System.Windows.Forms.Button
                {
                    Text = "确定",
                    Location = new System.Drawing.Point(150, 350),
                    Width = 80,
                    Height = 50,
                    DialogResult = System.Windows.Forms.DialogResult.OK
                };
                okButton.Click += OkButton_Click;
                cancelButton = new System.Windows.Forms.Button
                {
                    Text = "取消",
                    Location = new System.Drawing.Point(240, 350),
                    Width = 80,
                    Height = 50,
                    DialogResult = System.Windows.Forms.DialogResult.Cancel
                };
                this.Controls.Add(okButton);
                this.Controls.Add(cancelButton);

                this.AcceptButton = okButton;
                this.CancelButton = cancelButton;
            }

            private void TimeTextColorButton_Click(object sender, System.EventArgs e)
            {
                if (colorDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    TimeTextColor = colorDialog.Color;
                    timeTextColorBox.Text = TimeTextColor.ToArgb().ToString("X");
                }
            }

            private void OtherTextColorButton_Click(object sender, System.EventArgs e)
            {
                if (colorDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    OtherTextColor = colorDialog.Color;
                    otherTextColorBox.Text = OtherTextColor.ToArgb().ToString("X");
                }
            }

            private void BackgroundColorButton_Click(object sender, System.EventArgs e)
            {
                if (colorDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    BackgroundColor = colorDialog.Color;
                    backgroundColorBox.Text = BackgroundColor.ToArgb().ToString("X");
                }
            }

            private void DarkModeButtonColorButton_Click(object sender, System.EventArgs e)
            {
                if (colorDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    DarkModeButtonColor = colorDialog.Color;
                    darkModeButtonColorBox.Text = DarkModeButtonColor.ToArgb().ToString("X");
                }
            }

            private void OkButton_Click(object sender, System.EventArgs e)
            {
                IsCountdown = countdownRadioButton.Checked;
                SelectedFont = new System.Drawing.Font(fontComboBox.SelectedItem.ToString(), SelectedFont.Size);
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


        private void Boardpasting_Click(object sender, RibbonControlEventArgs e)
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
                    lines = System.IO.File.ReadAllLines(openFileDialog.FileName);
                }
            }
            else
            {
                // 创建并显示输入文本的窗口
                InputTextForm inputForm = new InputTextForm();
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
                            if (shape.Type == MsoShapeType.msoTextBox)
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
                                        if (s.Type == MsoShapeType.msoTextBox && s.TextFrame.TextRange.Text == group.Key)
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

       

        private void Fillblank_Click_1(object sender, RibbonControlEventArgs e)
        {
            // 获取当前幻灯片
            var slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            // 获取选中的文本框
            if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText)
            {
                var textRange = Globals.ThisAddIn.Application.ActiveWindow.Selection.TextRange;

                // 获取选中的文本
                string selectedText = textRange.Text;

                // 如果有选中的文本
                if (!string.IsNullOrEmpty(selectedText))
                {
                    // 获取原文本框的属性
                    var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                    var originalShape = selection.ShapeRange[1];
                    float fontSize = textRange.Font.Size;
                    string fontName = textRange.Font.Name;

                    // 获取选中文字的位置和大小
                    float originalLeft = textRange.BoundLeft;
                    float originalTop = textRange.BoundTop;

                    // 测量选中文本的宽度
                    float textWidth = MeasureTextWidth(selectedText, fontSize, fontName);

                    // 创建一个新的文本框，并设置其内容为选中的文本
                    var newTextBox = slide.Shapes.AddTextbox(
                        Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                        originalLeft, originalTop, textWidth, originalShape.Height);

                    var newTextFrame = newTextBox.TextFrame2;
                    var newTextRange = newTextBox.TextFrame.TextRange;
                    newTextRange.Text = selectedText;
                    newTextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    newTextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
                    newTextRange.Font.Size = fontSize;
                    newTextRange.Font.Name = fontName;

                    // 设置文本框不自动换行
                    newTextFrame.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;

                    // 确保下划线长度足够长但不过长
                    string underline = new string('_', (int)(selectedText.Length * 2.5)); // 动态生成下划线

                    // 将选中的文本用下划线替换
                    textRange.Text = underline;

                    // 设置新文本框的位置与被选中的文本相同
                    newTextBox.Left = originalLeft;
                    newTextBox.Top = originalTop - (originalShape.Height - fontSize) / 2; // 调整文本框位置
                }
                else
                {
                    MessageBox.Show("请选中文本框内的文本！");
                }
            }
            else
            {
                MessageBox.Show("请选中文本框内的文本！");
            }
        }

        private float MeasureTextWidth(string text, float fontSize, string fontName)
        {
            using (var bmp = new System.Drawing.Bitmap(1, 1))
            {
                using (var g = System.Drawing.Graphics.FromImage(bmp))
                {
                    var font = new System.Drawing.Font(fontName, fontSize);
                    var size = g.MeasureString(text, font);
                    return size.Width;
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

       

        private void 功能选择_TextChanged(object sender, RibbonControlEventArgs e)
        {
        }
        private string selectedFunction = "";

        private void 功能选择_Changed(object sender, RibbonControlEventArgs e)
        {
            selectedFunction = ((RibbonComboBox)sender).Text;
        }

        private void 参数输入_TextChanged(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = pptApp.ActiveWindow.Selection;

            if (string.IsNullOrEmpty(selectedFunction) || selection.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                MessageBox.Show("请先选择一个功能并选中一个或多个对象。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string input = ((RibbonEditBox)sender).Text.Trim();

            switch (selectedFunction)
            {
                case "批量改名":
                    批量改名(input, selection);
                    break;
                case "批量原位":
                    批量原位复制(input, selection);
                    break;
                case "尺寸比例":
                    尺寸缩放(input, selection);
                    break;
                default:
                    MessageBox.Show("未知功能。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            }

            ((RibbonEditBox)sender).Text = string.Empty;
        }

        private void 批量改名(string prefix, PowerPoint.Selection selection)
        {
            if (!string.IsNullOrEmpty(prefix))
            {
                int counter = 1;
                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    shape.Name = $"{prefix}-{counter}";
                    counter++;
                }
                Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex);
            }
            else
            {
                MessageBox.Show("请输入命名前缀。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void 批量原位复制(string input, PowerPoint.Selection selection)
        {
            if (int.TryParse(input, out int copyCount) && copyCount > 0)
            {
                for (int i = 0; i < copyCount; i++)
                {
                    DuplicateSelectedShapes(selection);
                }
            }
            else
            {
                MessageBox.Show("请输入一个大于0的整数。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DuplicateSelectedShapes(PowerPoint.Selection selection)
        {
            foreach (PowerPoint.Shape shape in selection.ShapeRange)
            {
                PowerPoint.Shape copiedShape = shape.Duplicate()[1];
                copiedShape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringForward);
                copiedShape.Left = shape.Left;
                copiedShape.Top = shape.Top;
            }
        }

        private void 尺寸缩放(string input, PowerPoint.Selection selection)
        {
            string[] scaleValues = input.Split(',');
            bool isArithmetic = scaleValues.Length == 2;
            float commonDifference = 0;

            if (isArithmetic)
            {
                if (float.TryParse(scaleValues[0], out float startScale) && float.TryParse(scaleValues[1], out float endScale))
                {
                    commonDifference = (endScale - startScale) / (selection.ShapeRange.Count - 1);
                    float currentScale = startScale;

                    foreach (PowerPoint.Shape shape in selection.ShapeRange)
                    {
                        ScaleShape(shape, currentScale);
                        currentScale += commonDifference;
                    }
                }
                else
                {
                    MessageBox.Show("请输入有效的缩放比例。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else if (float.TryParse(input, out float scale))
            {
                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    ScaleShape(shape, scale);
                }
            }
            else
            {
                MessageBox.Show("请输入有效的缩放比例。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ScaleShape(PowerPoint.Shape shape, float scale)
        {
            float newWidth = shape.Width * scale / 100;
            float newHeight = shape.Height * scale / 100;
            float newX = shape.Left + (shape.Width - newWidth) / 2;
            float newY = shape.Top + (shape.Height - newHeight) / 2;

            shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
            shape.Width = newWidth;
            shape.Height = newHeight;
            shape.Left = newX;
            shape.Top = newY;
        }
    }
}




























