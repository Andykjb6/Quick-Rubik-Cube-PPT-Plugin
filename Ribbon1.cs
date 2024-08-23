using System;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;  // 指定Shape引用的命名空间
using NPinyin;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Drawing;
using System.Diagnostics;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using System.Threading.Tasks;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using NStandard;
using OfficeOpenXml;
using System.Reflection;
using System.Windows.Input;
using Excel = OfficeOpenXml.ExcelPackage;
using System.Text;
using System.Text.Json;



namespace 课件帮PPT助手
{

    public partial class Ribbon1 : Office.IRibbonExtensibility
    {
        private CustomCloudTextGeneratorForm cloudTextGeneratorForm;
        public PowerPoint.Application PptApplication { get; set; }


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
                            PowerPoint.TimeLine timeLine = activeSlide.TimeLine;
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
                    PowerPoint.TextRange selectedTextRange = pptApplication.ActiveWindow.Selection.TextRange;
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
                    PowerPoint.TextRange newTextRange = newTextBox.TextFrame.TextRange;
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
            Application application = Globals.ThisAddIn.Application;

            // 检查当前视图是否支持选择操作
            if (application.ActiveWindow.ViewType == PpViewType.ppViewNormal)
            {
                // 获取当前的演示文稿
                Presentation presentation = application.ActivePresentation;

                // 检查是否有选中的对象
                if (application.ActiveWindow.Selection.Type != PpSelectionType.ppSelectionNone)
                {
                    Shape selectedShape = application.ActiveWindow.Selection.ShapeRange[1];

                    // 创建渐变透明矩形
                    Shape rectangle = null;
                    Slide slide = selectedShape.Parent as Slide;

                    if (slide != null)
                    {
                        // 插入一个与选中对象等大的矩形
                        rectangle = slide.Shapes.AddShape(
                            MsoAutoShapeType.msoShapeRectangle,
                            selectedShape.Left, selectedShape.Top, selectedShape.Width, selectedShape.Height);

                        // 设置边框为不可见
                        rectangle.Line.Visible = MsoTriState.msoFalse;

                        // 设置填充为渐变
                        rectangle.Fill.OneColorGradient(MsoGradientStyle.msoGradientHorizontal, 1, 1);

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

                        // 获取所选对象的图层位置
                        int selectedShapeZOrder = selectedShape.ZOrderPosition;

                        // 将矩形置于所选对象的后一个图层
                        while (rectangle.ZOrderPosition > selectedShapeZOrder + 1)
                        {
                            rectangle.ZOrder(MsoZOrderCmd.msoSendBackward);
                        }
                    }
                }
                else
                {
                    // 当没有选中对象时，默认在当前幻灯片上插入与幻灯片等大的渐变透明矩形
                    Slide slide = application.ActiveWindow.View.Slide as Slide;

                    if (slide != null)
                    {
                        // 插入一个与幻灯片等大的矩形
                        Shape rectangle = slide.Shapes.AddShape(
                            MsoAutoShapeType.msoShapeRectangle,
                            0, 0, slide.Master.Width, slide.Master.Height);

                        // 设置边框为不可见
                        rectangle.Line.Visible = MsoTriState.msoFalse;

                        // 设置填充为渐变
                        rectangle.Fill.OneColorGradient(MsoGradientStyle.msoGradientHorizontal, 1, 1);

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
                        rectangle.ZOrder(MsoZOrderCmd.msoSendToBack);
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
                PowerPoint.ShapeRange shapeRange = selection.ShapeRange;

                // 初始化边界值
                float groupLeft = float.MaxValue;
                float groupTop = float.MaxValue;
                float groupRight = float.MinValue;
                float groupBottom = float.MinValue;

                // 计算选中对象的边界框
                foreach (PowerPoint.Shape shape in shapeRange)
                {
                    if (shape.Left < groupLeft)
                        groupLeft = shape.Left;
                    if (shape.Top < groupTop)
                        groupTop = shape.Top;
                    if (shape.Left + shape.Width > groupRight)
                        groupRight = shape.Left + shape.Width;
                    if (shape.Top + shape.Height > groupBottom)
                        groupBottom = shape.Top + shape.Height;
                }

                // 计算选中对象整体的中心位置
                float groupCenterX = (groupLeft + groupRight) / 2;
                float groupCenterY = (groupTop + groupBottom) / 2;

                // 检查是否按住Ctrl键
                if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
                {
                    // 计算需要水平移动的距离
                    float deltaX = slideCenterX - groupCenterX;

                    // 遍历选中的每一个对象，横向移动到水平线的中部
                    foreach (PowerPoint.Shape shape in shapeRange)
                    {
                        shape.Left += deltaX;
                    }
                }
                else
                {
                    // 计算需要移动的距离
                    float deltaX = slideCenterX - groupCenterX;
                    float deltaY = slideCenterY - groupCenterY;

                    // 遍历选中的每一个对象，平移到新的位置
                    foreach (PowerPoint.Shape shape in shapeRange)
                    {
                        shape.Left += deltaX;
                        shape.Top += deltaY;
                    }
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
        
        }

        private void 平移居中_Click(object sender, RibbonControlEventArgs e)
        {
            // 直接创建并显示WPF窗体
            AlignmentWindow alignmentWindow = new AlignmentWindow(Globals.ThisAddIn.Application);
            alignmentWindow.Show();
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
                // 检查形状是否包含文本框或文本
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    shape.TextFrame.TextRange.Text = replacementText;
                }
                else if (shape.Type == MsoShapeType.msoGroup)
                {
                    // 如果是分组形状，递归处理
                    ReplaceTextInGroup(shape.GroupItems, replacementText);
                }
            }
        }

        private void ReplaceTextInGroup(Microsoft.Office.Interop.PowerPoint.GroupShapes groupShapes, string replacementText)
        {
            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in groupShapes)
            {
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    shape.TextFrame.TextRange.Text = replacementText;
                }
                else if (shape.Type == MsoShapeType.msoGroup)
                {
                    // 递归处理嵌套的分组形状
                    ReplaceTextInGroup(shape.GroupItems, replacementText);
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
        private TableForm Form;
        private void 生字格子_Click(object sender, RibbonControlEventArgs e)
        {
            if (Form == null || Form.IsDisposed)
            {
                Form = new TableForm();
            }

            Form.Show();
            Form.TopMost = true;
        }

    
        private void 生字赋格_Click(object sender, RibbonControlEventArgs e)
        {
            TableSettingsForm form = new TableSettingsForm();
            form.Show();
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



        private void 板贴辅助_Click(object sender, RibbonControlEventArgs e)
        {
            string[] lines = null;

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
            Application app = Globals.ThisAddIn.Application;
            Selection selection = app.ActiveWindow.Selection;

            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange shapeRange = selection.ShapeRange;

                foreach (PowerPoint.Shape shape in shapeRange)
                {
                    if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
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
                        MessageBox.Show("请选择包含文本的文本框。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            else
            {
                MessageBox.Show("请选择一个或多个文本框。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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


       
        private void 指定对齐_Click(object sender, RibbonControlEventArgs e)
        {
            
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
            string url = "https://www.pngtosvg.com/";
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
            Application pptApp = Globals.ThisAddIn.Application;
            Selection selection = pptApp.ActiveWindow.Selection;

            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                foreach (Shape shape in selection.ShapeRange)
                {
                    Shape newShape = shape.Duplicate()[1]; // 复制形状
                    newShape.Left = shape.Left; // 保持原位
                    newShape.Top = shape.Top;   // 保持原位

                    // 获取所选对象的图层位置
                    int selectedShapeZOrder = shape.ZOrderPosition;

                    // 将新复制的形状置于所选对象的后一个图层
                    while (newShape.ZOrderPosition > selectedShapeZOrder + 1)
                    {
                        newShape.ZOrder(MsoZOrderCmd.msoSendBackward);
                    }
                }
            }
            else
            {
                MessageBox.Show("请选择一个或多个对象进行复制。", "提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
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

                        // 保存shape1和shape2的图层顺序
                        int shape1ZOrderPosition = shape1.ZOrderPosition;
                        int shape2ZOrderPosition = shape2.ZOrderPosition;

                        // 交换位置
                        shape1.Left = shape2Left;
                        shape1.Top = shape2Top;
                        shape2.Left = shape1Left;
                        shape2.Top = shape1Top;

                        // 交换图层顺序
                        ExchangeShapeZOrderPosition(shape1, shape2ZOrderPosition);
                        ExchangeShapeZOrderPosition(shape2, shape1ZOrderPosition);
                    }
                    else
                    {
                        MessageBox.Show("请选中两个对象以交换它们的位置和图层。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("请选中两个对象以交换它们的位置和图层。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("交换位置和图层时发生错误: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExchangeShapeZOrderPosition(Shape shape, int targetZOrderPosition)
        {
            while (shape.ZOrderPosition > targetZOrderPosition)
            {
                shape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
            }

            while (shape.ZOrderPosition < targetZOrderPosition)
            {
                shape.ZOrder(Office.MsoZOrderCmd.msoBringForward);
            }
        }

        private List<Shape> GetSelectedShapesFromSelection(Selection selection)
        {
            var selectedShapes = new List<Shape>();

            // 检查选择的类型
            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                // 获取选中的形状范围
                PowerPoint.ShapeRange selectedShapeRange = selection.ShapeRange;
                if (selection.HasChildShapeRange)
                {
                    selectedShapeRange = selection.ChildShapeRange;
                }

                // 遍历选中的形状范围
                foreach (Shape shape in selectedShapeRange)
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


        private void 交换尺寸_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var application = Globals.ThisAddIn.Application;
                var selection = application.ActiveWindow.Selection;

                if (selection.Type == PpSelectionType.ppSelectionShapes)
                {
                    var selectedShapes = selection.ShapeRange;

                    if (selection.HasChildShapeRange)
                    {
                        selectedShapes = selection.ChildShapeRange;
                    }

                    if (selectedShapes.Count == 2)
                    {
                        Shape shape1 = selectedShapes[1];
                        Shape shape2 = selectedShapes[2];

                        float shape1OriginalWidth = shape1.Width;
                        float shape1OriginalHeight = shape1.Height;
                        float shape2OriginalWidth = shape2.Width;
                        float shape2OriginalHeight = shape2.Height;

                        float shape1Bottom = shape1.Top + shape1.Height;
                        float shape2Bottom = shape2.Top + shape2.Height;

                        float shape1AspectRatio = shape1OriginalWidth / shape1OriginalHeight;
                        float shape2AspectRatio = shape2OriginalWidth / shape2OriginalHeight;

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

                        // 调整Top属性，使底部位置不变
                        shape1.Top = shape1Bottom - shape1.Height;
                        shape2.Top = shape2Bottom - shape2.Height;
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
                float newWidth = originalHeight * targetAspectRatio;
                float cropWidth = (originalWidth - newWidth) / 2;
                shape.PictureFormat.CropLeft += cropWidth;
                shape.PictureFormat.CropRight += cropWidth;
            }
            else
            {
                float newHeight = originalWidth / targetAspectRatio;
                float cropHeight = (originalHeight - newHeight) / 2;
                shape.PictureFormat.CropTop += cropHeight;
                shape.PictureFormat.CropBottom += cropHeight;
            }

            shape.Width = targetWidth;
            shape.Height = targetHeight;
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

                    // 记录这两个形状的文本内容
                    string text1 = shape1.HasTextFrame == MsoTriState.msoTrue && shape1.TextFrame.HasText == MsoTriState.msoTrue
                        ? shape1.TextFrame.TextRange.Text : null;
                    string text2 = shape2.HasTextFrame == MsoTriState.msoTrue && shape2.TextFrame.HasText == MsoTriState.msoTrue
                        ? shape2.TextFrame.TextRange.Text : null;

                    // 交换文本内容
                    shape1.TextFrame.TextRange.Text = text2 ?? string.Empty;
                    shape2.TextFrame.TextRange.Text = text1 ?? string.Empty;
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

        private List<Shape> GetSelectedShapesForTextExchange(Selection selection)
        {
            var selectedShapes = new List<Shape>();

            // 检查选择的类型
            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                // 获取选中的形状范围
                PowerPoint.ShapeRange selectedShapeRange = selection.ShapeRange;
                if (selection.HasChildShapeRange)
                {
                    selectedShapeRange = selection.ChildShapeRange;
                }

                // 遍历选中的形状范围
                foreach (Shape shape in selectedShapeRange)
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
                    // 如果没有选中任何形状，则插入默认的四线三格
                    gridGroup = InsertFourLineThreeGrid(slide, defaultWidth, defaultHeight);
                }
                else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    PowerPoint.Shape shape = sel.ShapeRange[1];
                    if (shape.Type == Office.MsoShapeType.msoTextBox)
                    {
                        // 首先移除文本框的边距
                        去除边距_Click(sender, e);

                        // 获取文本框的新尺寸
                        float textBoxWidth = shape.Width;
                        float textBoxHeight = shape.Height;
                        float newHeight = textBoxHeight + additionalHeight; // 文本框高度 + 0.25 cm

                        // 插入四线三格，宽度与文本框一致，高度为文本框高度+额外高度
                        gridGroup = InsertFourLineThreeGrid(slide, textBoxWidth, newHeight);

                        // 确保四线三格与文本框顶端对齐
                        gridGroup.Top = shape.Top;

                        // 水平居中四线三格与文本框
                        float textBoxCenter = shape.Left + (textBoxWidth / 2);
                        gridGroup.Left = textBoxCenter - (gridGroup.Width / 2);

                        // 确保文本框位于最前面
                        shape.ZOrder(MsoZOrderCmd.msoBringToFront);
                    }
                    else if (shape.Type == Office.MsoShapeType.msoTable)
                    {
                        // 插入在选中表格的上方
                        float tableWidth = shape.Width;
                        float tableTop = shape.Top - 10 - defaultHeight; // 正确调整顶部位置
                        gridGroup = InsertFourLineThreeGrid(slide, tableWidth, defaultHeight);
                        gridGroup.Top = tableTop;
                        gridGroup.Left = shape.Left; // 左对齐
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

        private Shape InsertFourLineThreeGrid(Slide slide, float width, float height)
        {
            float lineSpacing = height / 3.0f;
            PowerPoint.Shapes shapes = slide.Shapes;
            Shape line1 = shapes.AddLine(0, 0, width, 0);
            Shape line2 = shapes.AddLine(0, lineSpacing, width, lineSpacing);
            Shape line3 = shapes.AddLine(0, lineSpacing * 2, width, lineSpacing * 2);
            Shape line4 = shapes.AddLine(0, height, width, height);

            line1.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);
            line1.Line.Weight = 1.5f;
            line4.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);
            line4.Line.Weight = 1.5f;
            line2.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray);
            line2.Line.Weight = 1.0f;
            line3.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Red);
            line3.Line.Weight = 1.0f;

            PowerPoint.ShapeRange shapeRange = slide.Shapes.Range(new string[] { line1.Name, line2.Name, line3.Name, line4.Name });
            return shapeRange.Group();
        }


        private Shape AdjustFourLineThreeGrid(Shape gridGroup, float newSpacing)
        {
            PowerPoint.ShapeRange shapes = gridGroup.Ungroup();
            shapes[1].Top = newSpacing;
            shapes[2].Top = newSpacing * 2;
            shapes[3].Top = newSpacing * 3;
            return shapes.Group();
        }

        private float GetMinCharacterHeight(Shape textBox)
        {
            TextRange textRange = textBox.TextFrame.TextRange;
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


        private void 移动对齐_Click(object sender, RibbonControlEventArgs e)
        {
           
        }

        private void 智能缩放_Click(object sender, RibbonControlEventArgs e)
        {
            SmartScalingForm scalingForm = new SmartScalingForm();
            scalingForm.Show();
        }

        private void 字典注音_Click(object sender, RibbonControlEventArgs e)
        {
            // 检查是否已激活
            if (!Properties.Settings.Default.IsActivated)
            {
                // 插件未激活，提示用户激活
                System.Windows.Forms.MessageBox.Show("请激活插件以使用此功能。", "未激活", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                return; // 插件未激活，直接返回，不执行后续代码
            }

            // 插件已激活，允许使用功能
            // 设置EPPlus的许可上下文
            Excel.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            // 获取当前PPT应用和选中的文本框或文本
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Selection pptSelection = pptApp.ActiveWindow.Selection;

            if (pptSelection.Type == PowerPoint.PpSelectionType.ppSelectionText || pptSelection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape selectedShape in pptSelection.ShapeRange)
                {
                    // 只处理文本框
                    if (selectedShape.Type == Office.MsoShapeType.msoTextBox)
                    {
                        PowerPoint.TextRange textRange = selectedShape.TextFrame.TextRange;
                        string selectedText;
                        bool isCharacterSelected = false;

                        // 检查用户是否在文本框内选中了某些字符
                        if (pptSelection.Type == PowerPoint.PpSelectionType.ppSelectionText && textRange.Length > 0)
                        {
                            // 处理选中的单个字符或部分文本
                            selectedText = pptSelection.TextRange.Text;
                            textRange = pptSelection.TextRange;
                            isCharacterSelected = true;
                        }
                        else
                        {
                            // 如果没有选中字符，处理整个文本框中的文本
                            selectedText = textRange.Text;
                        }

                        // 从嵌入资源中提取汉字拼音信息库Excel文件
                        string filePath = ExtractEmbeddedResource("课件帮PPT助手.汉字字典.汉字拼音信息库.xlsx");
                        // 加载汉字拼音字典
                        Dictionary<string, string> hanziPinyinDictionary = LoadHanziPinyinDictionary(filePath);

                        // 从嵌入资源中提取汉字字典Excel文件
                        string hanziDictionaryFilePath = ExtractEmbeddedResource("课件帮PPT助手.汉字字典.汉字字典.xlsx");
                        // 加载汉字字典
                        Dictionary<string, string> hanziDictionary = LoadHanziPinyinDictionary(hanziDictionaryFilePath);

                        // 加载多音字词语库
                        string duoyinziFilePath = ExtractEmbeddedResource("课件帮PPT助手.汉字字典.多音字词语.txt");
                        Dictionary<string, string> duoyinziDictionary = LoadDuoyinziDictionary(duoyinziFilePath);

                        // 首先检查是否在多音字词语库中找到整个词语
                        if (duoyinziDictionary.ContainsKey(selectedText))
                        {
                            // 获取词语对应的拼音并分割
                            string[] pinyinArray = duoyinziDictionary[selectedText].Split(' ');

                            // 遍历每个字符，为其生成独立的拼音文本框
                            for (int i = 0; i < selectedText.Length; i++)
                            {
                                PowerPoint.TextRange charRange = textRange.Characters(i + 1, 1);
                                string annotatedText = pinyinArray[i]; // 使用从多音字词语库中获取的拼音

                                // 获取选中文字的位置和大小
                                float fontSize = charRange.Font.Size;
                                float charTop = charRange.BoundTop - fontSize / 2; // 初始位置调整

                                // 仅当行间距 > 1 且选中的是文本框内的某个字符时，调整拼音文本框的位置
                                if (isCharacterSelected && charRange.ParagraphFormat.SpaceWithin > 1)
                                {
                                    float extraLineSpacing = Math.Max(0, charRange.BoundHeight - fontSize);
                                    const float downwardShiftRatio = 0.5f;  // 基于额外行间距的向下调整比例
                                    charTop += extraLineSpacing * downwardShiftRatio;
                                }

                                float charLeft = charRange.BoundLeft;
                                float charWidth = charRange.BoundWidth;
                                float newFontSize = fontSize / 2;

                                // 获取当前形状所属的幻灯片
                                PowerPoint.Slide currentSlide = selectedShape.Parent;

                                // 创建新的文本框并插入注音后的文本
                                PowerPoint.Shape newShape = currentSlide.Shapes.AddTextbox(
                                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                                    charLeft, charTop - 5, charWidth, selectedShape.Height / 2);
                                newShape.TextFrame.TextRange.Text = annotatedText;

                                // 设置新文本框的字体大小为原文本框字体大小的一半
                                newShape.TextFrame.TextRange.Font.Size = newFontSize;

                                // 设置拼音文本框的文本对齐方式为居中
                                newShape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;

                                // 确保新文本框不自动换行
                                newShape.TextFrame.WordWrap = Office.MsoTriState.msoFalse;

                                // 检查是否有重叠的田字格对象
                                if (IsOverlappingWithTianZiGe(selectedShape))
                                {
                                    // 将拼音文本框移动到所选对象的顶部上方
                                    newShape.Top = selectedShape.Top - newShape.Height;
                                }
                            }
                        }
                        else
                        {
                            // 如果没有找到完整词语，逐个处理字符
                            for (int i = 1; i <= textRange.Text.Length; i++)
                            {
                                PowerPoint.TextRange charRange = textRange.Characters(i, 1);
                                string charText = charRange.Text;

                                // 如果字符是标点符号，跳过处理
                                if (char.IsPunctuation(charText, 0))
                                {
                                    continue;
                                }

                                // 获取拼音注音后的文本
                                string annotatedText = GetPinyinForText(charText, hanziPinyinDictionary, duoyinziDictionary, hanziDictionary, selectedText);

                                // 获取选中文字的位置和大小
                                float fontSize = charRange.Font.Size;
                                float charTop = charRange.BoundTop - fontSize / 2; // 初始位置调整

                                // 仅当行间距 > 1 且选中的是文本框内的某个字符时，调整拼音文本框的位置
                                if (isCharacterSelected && charRange.ParagraphFormat.SpaceWithin > 1)
                                {
                                    float extraLineSpacing = Math.Max(0, charRange.BoundHeight - fontSize);
                                    const float downwardShiftRatio = 0.5f;  // 基于额外行间距的向下调整比例
                                    charTop += extraLineSpacing * downwardShiftRatio;
                                }

                                float charLeft = charRange.BoundLeft;
                                float charWidth = charRange.BoundWidth;
                                float newFontSize = fontSize / 2;

                                // 获取当前形状所属的幻灯片
                                PowerPoint.Slide currentSlide = selectedShape.Parent;

                                // 创建新的文本框并插入注音后的文本
                                PowerPoint.Shape newShape = currentSlide.Shapes.AddTextbox(
                                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                                    charLeft, charTop - 5, charWidth, selectedShape.Height / 2);
                                newShape.TextFrame.TextRange.Text = annotatedText;

                                // 设置新文本框的字体大小为原文本框字体大小的一半
                                newShape.TextFrame.TextRange.Font.Size = newFontSize;

                                // 设置拼音文本框的文本对齐方式为居中
                                newShape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;

                                // 确保新文本框不自动换行
                                newShape.TextFrame.WordWrap = Office.MsoTriState.msoFalse;

                                // 检查是否有重叠的田字格对象
                                if (IsOverlappingWithTianZiGe(selectedShape))
                                {
                                    // 将拼音文本框移动到所选对象的顶部上方
                                    newShape.Top = selectedShape.Top - newShape.Height;
                                }
                            }
                        }
                    }
                    else
                    {
                        // 跳过非文本框对象
                        continue;
                    }
                }
            }
            else
            {
                MessageBox.Show("请先选中一个或多个文本框。");
            }
        }

        // 检查所选文本框底部是否存在重叠的对象且该对象图层名称以“田字格”开头
        private bool IsOverlappingWithTianZiGe(PowerPoint.Shape selectedShape)
        {
            PowerPoint.Slide currentSlide = selectedShape.Parent;
            foreach (PowerPoint.Shape shape in currentSlide.Shapes)
            {
                try
                {
                    // 检查是否是一个组合形状，且名称以“田字格”开头
                    if (shape.Type == Office.MsoShapeType.msoGroup && shape.Name.StartsWith("田字格"))
                    {
                        // 检查是否重叠
                        if (shape.Left < selectedShape.Left + selectedShape.Width &&
                            shape.Left + shape.Width > selectedShape.Left &&
                            shape.Top < selectedShape.Top + selectedShape.Height &&
                            shape.Top + shape.Height > selectedShape.Top)
                        {
                            return true;
                        }
                    }
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    // 捕获并处理COM异常
                    Debug.WriteLine($"COM Exception in IsOverlappingWithTianZiGe: {ex.Message}");
                    continue;
                }
            }
            return false;
        }

        private Dictionary<string, string> LoadHanziPinyinDictionary(string filePath)
        {
            var hanziPinyinDictionary = new Dictionary<string, string>();
            using (var package = new Excel(new FileInfo(filePath)))
            {
                foreach (var worksheet in package.Workbook.Worksheets)
                {
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
            }
            return hanziPinyinDictionary;
        }

        private Dictionary<string, string> LoadDuoyinziDictionary(string filePath)
        {
            var duoyinziDictionary = new Dictionary<string, string>();
            var lines = File.ReadAllLines(filePath);
            foreach (var line in lines)
            {
                var parts = line.Split(new[] { '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 2)
                {
                    string ci = parts[0].Trim();
                    string pinyin = parts[1].Trim();
                    duoyinziDictionary[ci] = pinyin;
                }
            }
            return duoyinziDictionary;
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

        private string GetPinyinForText(string text, Dictionary<string, string> hanziPinyinDictionary, Dictionary<string, string> duoyinziDictionary, Dictionary<string, string> hanziDictionary, string selectedText)
        {
            List<string> pinyinList = new List<string>();
            string remainingText = text;

            for (int index = 0; index < remainingText.Length; index++)
            {
                char currentChar = remainingText[index];
                string hanzi = currentChar.ToString();

                // 处理普通汉字
                if (hanziPinyinDictionary.ContainsKey(hanzi))
                {
                    string pinyin = hanziPinyinDictionary[hanzi];
                    pinyinList.Add(GetPinyinFromDictionary(pinyin, hanzi, text, selectedText));
                }
                else if (hanziDictionary.ContainsKey(hanzi))
                {
                    string pinyin = hanziDictionary[hanzi];
                    pinyinList.Add(GetPinyinFromDictionary(pinyin, hanzi, text, selectedText));
                }
                else
                {
                    pinyinList.Add(hanzi);
                }
            }

            return string.Join(" ", pinyinList);
        }

        private string GetPinyinFromDictionary(string pinyin, string hanzi, string context, string selectedText)
        {
            if (pinyin.Contains(","))
            {
                using (PinYinForm form = new PinYinForm(hanzi, selectedText, pinyin.Split(',')))
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        return form.SelectedPinyin;
                    }
                    else
                    {
                        return pinyin.Split(',')[0];
                    }
                }
            }
            else
            {
                return pinyin;
            }
        }

        private void Zici_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前PPT应用和选中的文本框
            Application pptApp = Globals.ThisAddIn.Application;
            Selection pptSelection = pptApp.ActiveWindow.Selection;

            // 设置EPPlus的许可上下文
            Excel.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            if (pptSelection.Type == PpSelectionType.ppSelectionShapes)
            {
                foreach (Shape selectedShape in pptSelection.ShapeRange)
                {
                    if (selectedShape.HasTextFrame == MsoTriState.msoTrue && selectedShape.TextFrame.HasText == MsoTriState.msoTrue)
                    {
                        ProcessShape(selectedShape);
                    }
                }
            }
        }

        private void ProcessShape(Shape selectedShape)
        {
            TextRange textRange = selectedShape.TextFrame.TextRange;
            if (textRange != null && !string.IsNullOrEmpty(textRange.Text))
            {
                string selectedText = textRange.Text;

                // 从嵌入资源中提取字典和词语库
                string filePath = ExtractEmbeddedResource("课件帮PPT助手.汉字字典.汉字拼音信息库.xlsx");
                Dictionary<string, string> hanziPinyinDictionary = LoadHanziPinyinDictionary(filePath);

                string duoyinziFilePath = ExtractEmbeddedResource("课件帮PPT助手.汉字字典.多音字词语.txt");
                Dictionary<string, string> duoyinziDictionary = LoadDuoyinziDictionary(duoyinziFilePath);

                PowerPoint.Slide currentSlide = selectedShape.Parent;

                StringBuilder pinyinTextBuilder = new StringBuilder();
                bool hasProcessedCompleteWord = false;

                // 优先处理多音字词语库中的完整词语
                if (selectedText.Length > 1)
                {
                    foreach (var kvp in duoyinziDictionary)
                    {
                        if (selectedText == kvp.Key)
                        {
                            // 找到匹配的词语，使用词语的拼音
                            string pinyin = kvp.Value;
                            pinyinTextBuilder.Append(pinyin).Append(" ");
                            hasProcessedCompleteWord = true;
                            break;
                        }
                    }
                }

                if (!hasProcessedCompleteWord)
                {
                    // 如果没有找到完整词语，逐个处理字符
                    for (int i = 1; i <= selectedText.Length; i++)
                    {
                        char c = selectedText[i - 1];
                        string pinyin = GetPinyinForCharacter(c.ToString(), hanziPinyinDictionary, duoyinziDictionary);

                        if (!string.IsNullOrEmpty(pinyin))
                        {
                            string[] pinyinOptions = pinyin.Split(',');

                            if (pinyinOptions.Length > 1)
                            {
                                // 多音字，弹出选择框
                                using (var form = new PinyinSelectionForm(c.ToString(), selectedText, pinyinOptions))
                                {
                                    if (form.ShowDialog() == DialogResult.OK)
                                    {
                                        pinyin = form.SelectedPinyin;
                                    }
                                    else
                                    {
                                        pinyin = pinyinOptions[0]; // 用户未选择时，默认使用第一个拼音
                                    }
                                }
                            }

                            pinyinTextBuilder.Append(pinyin).Append(" ");
                        }
                    }
                }

                string finalPinyinText = pinyinTextBuilder.ToString().Trim();

                // 创建拼音文本框
                float charTop = textRange.BoundTop - textRange.Font.Size / 2;
                float charLeft = textRange.BoundLeft;
                float charWidth = textRange.BoundWidth;

                CreatePinyinTextbox(currentSlide, selectedShape, finalPinyinText, charLeft, charTop, charWidth);

                // 创建括号文本框
                int chineseCharCount = selectedText.Count(c => c >= 0x4e00 && c <= 0x9fff);
                string spaces = new string('\u3000', chineseCharCount); // 使用全角空格符填充
                string parenthesesText = $"（{spaces}）";

                Shape parenthesesShape = currentSlide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    selectedShape.Left,
                    selectedShape.Top,
                    selectedShape.Width,
                    selectedShape.Height
                );

                parenthesesShape.TextFrame.TextRange.Text = parenthesesText;
                parenthesesShape.TextFrame.TextRange.Font.Size = textRange.Font.Size;
                parenthesesShape.TextFrame.TextRange.ParagraphFormat.Alignment = textRange.ParagraphFormat.Alignment;
                parenthesesShape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                parenthesesShape.ZOrder(MsoZOrderCmd.msoSendToBack);
                parenthesesShape.TextFrame.WordWrap = MsoTriState.msoFalse; // 取消自动换行

                // 调整括号文本框的位置，使其在水平和垂直方向上居中对齐
                parenthesesShape.Left = selectedShape.Left + (selectedShape.Width - parenthesesShape.Width) / 2;
                parenthesesShape.Top = selectedShape.Top + (selectedShape.Height - parenthesesShape.Height) / 2;

                // 修改用户所选文本的字体样式
                textRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Red);
                textRange.Font.Bold = MsoTriState.msoTrue;
                selectedShape.TextFrame.WordWrap = MsoTriState.msoFalse; // 取消自动换行
            }
        }

        private void CreatePinyinTextbox(Slide slide, Shape selectedShape, string pinyin, float left, float top, float width)
        {
            float fontSize = selectedShape.TextFrame.TextRange.Font.Size * 0.6f;

            // 创建拼音文本框
            Shape pinyinShape = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                left, top - 8, width, selectedShape.Height / 2);

            // 设置拼音文本框内容
            pinyinShape.TextFrame.TextRange.Text = pinyin;

            // 设置拼音文本框的字体大小和对齐方式
            pinyinShape.TextFrame.TextRange.Font.Size = fontSize;
            pinyinShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            pinyinShape.TextFrame.WordWrap = MsoTriState.msoFalse;

            // 调整拼音文本框的宽度以适应内容
            float pinyinWidth = pinyinShape.TextFrame.TextRange.BoundWidth;
            pinyinShape.Width = pinyinWidth;

            // 使拼音文本框水平居中对齐所选文本框
            pinyinShape.Left = selectedShape.Left + (selectedShape.Width - pinyinWidth) / 2;
        }

        // 获取单个字符的拼音
        private string GetPinyinForCharacter(string character, Dictionary<string, string> hanziPinyinDictionary, Dictionary<string, string> duoyinziDictionary)
        {
            if (hanziPinyinDictionary.ContainsKey(character))
            {
                return hanziPinyinDictionary[character];
            }
            else if (duoyinziDictionary.ContainsKey(character))
            {
                return duoyinziDictionary[character];
            }
            else
            {
                return character; // 如果字典中没有找到，则返回原字符
            }
        }


        private void WritePinyin_Click(object sender, RibbonControlEventArgs e)
        {
            // 设置EPPlus的许可上下文
            Excel.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            // 初始化汉字拼音字典和多音字词语库
            string hanziDictionaryPath = ExtractEmbeddedResource("课件帮PPT助手.汉字字典.汉字拼音信息库.xlsx");
            Dictionary<string, string> hanziPinyinDictionary = LoadHanziPinyinDictionary(hanziDictionaryPath);

            string duoyinziDictionaryPath = ExtractEmbeddedResource("课件帮PPT助手.汉字字典.多音字词语.txt");
            Dictionary<string, string> duoyinziDictionary = LoadDuoyinziDictionary(duoyinziDictionaryPath);

            // 获取当前PPT应用和选中的文本框或文本
            Application pptApp = Globals.ThisAddIn.Application;
            Selection pptSelection = pptApp.ActiveWindow.Selection;

            // 确保选择了文本框
            if (pptSelection.Type == PpSelectionType.ppSelectionText || pptSelection.Type == PpSelectionType.ppSelectionShapes)
            {
                float currentX = 0;
                float currentY = 0;
                float previousShapeWidth = 0;
                float previousParenthesesWidth = 0;

                foreach (Shape selectedShape in pptSelection.ShapeRange)
                {
                    // 只处理文本框
                    if (selectedShape.Type == Office.MsoShapeType.msoTextBox)
                    {
                        if (currentX == 0 && currentY == 0)
                        {
                            // 初始化第一个文本框的位置
                            currentX = selectedShape.Left;
                            currentY = selectedShape.Top;
                        }
                        else if (selectedShape.Top != currentY)
                        {
                            // 如果当前文本框在新的一行，则重置currentX，并更新currentY
                            currentX = selectedShape.Left;
                            currentY = selectedShape.Top;
                        }
                        else
                        {
                            // 更新currentX以在同一行继续排列文本框
                            currentX += previousShapeWidth + previousParenthesesWidth;
                        }

                        // 移动选中的文本框到计算好的位置
                        selectedShape.Left = currentX;
                        selectedShape.Top = currentY;

                        // 处理文本框内容
                        TextRange textRange = selectedShape.TextFrame.TextRange;
                        string selectedText = textRange.Text;

                        // 动态生成拼音
                        string annotatedText = PinyinForText(selectedText, hanziPinyinDictionary, duoyinziDictionary);

                        // 计算拼音文本框的英文字符数量
                        int englishCharCount = annotatedText.Count(c => c <= 0x7F);

                        // 生成括号内的英文半角空格符
                        string spaces = new string(' ', englishCharCount + 9);
                        string parenthesesText = $"（{spaces}）";

                        // 在原文本框的右侧创建括号文本框
                        Shape parenthesesShape = pptSelection.SlideRange[1].Shapes.AddTextbox(
                            MsoTextOrientation.msoTextOrientationHorizontal,
                            currentX + selectedShape.Width - 15,
                            currentY,
                            selectedShape.Width / 2,
                            selectedShape.Height
                        );

                        parenthesesShape.TextFrame.TextRange.Text = parenthesesText;
                        parenthesesShape.TextFrame.TextRange.Font.Size = textRange.Font.Size;
                        parenthesesShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                        parenthesesShape.TextFrame.WordWrap = MsoTriState.msoFalse; // 取消自动换行

                        // 调整括号文本框的宽度以适应文本内容
                        float parenthesesWidth = parenthesesShape.TextFrame.TextRange.BoundWidth;
                        parenthesesShape.Width = parenthesesWidth;

                        // 在括号区域中创建拼音文本框，确保重合
                        Shape pinyinShape = pptSelection.SlideRange[1].Shapes.AddTextbox(
                            MsoTextOrientation.msoTextOrientationHorizontal,
                            parenthesesShape.Left,
                            parenthesesShape.Top - (textRange.Font.Size / 2),
                            parenthesesShape.Width,
                            parenthesesShape.Height
                        );

                        pinyinShape.TextFrame.TextRange.Text = annotatedText; // 动态生成的拼音
                        pinyinShape.TextFrame.TextRange.Font.Size = textRange.Font.Size - 4; // 字号小2个字
                        pinyinShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Red); // 红色字体
                        pinyinShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                        pinyinShape.TextFrame.WordWrap = MsoTriState.msoFalse; // 取消自动换行
                        pinyinShape.ZOrder(MsoZOrderCmd.msoBringToFront); // 置于顶层

                        // 将拼音文本框的位置与括号文本框对齐
                        pinyinShape.Left = parenthesesShape.Left;
                        pinyinShape.Top = parenthesesShape.Top;
                        pinyinShape.Width = parenthesesShape.Width;

                        // 更新用于下一次循环的参数
                        previousShapeWidth = selectedShape.Width;
                        previousParenthesesWidth = parenthesesShape.Width;
                    }
                    else
                    {
                        // 跳过非文本框对象
                        continue;
                    }
                }
            }
            else
            {
                MessageBox.Show("请先选中一个或多个文本框。");
            }
        }



        // 下面的辅助方法保持不变
        private Dictionary<string, string> HanziPinyinDictionary(string filePath)
        {
            var hanziPinyinDictionary = new Dictionary<string, string>();
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++) // 假设第一行是标题
                    {
                        string hanzi = worksheet.Cells[row, 1].Text;
                        string pinyin = worksheet.Cells[row, 2].Text;
                        if (!string.IsNullOrWhiteSpace(hanzi) && !string.IsNullOrWhiteSpace(pinyin))
                        {
                            hanziPinyinDictionary[hanzi] = pinyin;
                        }
                    }
                }
            }
            return hanziPinyinDictionary;
        }

        private Dictionary<string, string> DuoyinziDictionary(string filePath)
        {
            var duoyinziDictionary = new Dictionary<string, string>();
            var lines = File.ReadAllLines(filePath);
            foreach (var line in lines)
            {
                var parts = line.Split(new[] { '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 2)
                {
                    string ci = parts[0].Trim();
                    string pinyin = parts[1].Trim();
                    duoyinziDictionary[ci] = pinyin;
                }
            }
            return duoyinziDictionary;
        }

        private string EmbeddedResource(string resourceName)
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

        private string PinyinForText(string text, Dictionary<string, string> hanziPinyinDictionary, Dictionary<string, string> duoyinziDictionary)
        {
            List<string> pinyinList = new List<string>();
            string remainingText = text;

            // 先查找多音字词语库
            foreach (var kvp in duoyinziDictionary)
            {
                if (remainingText.Contains(kvp.Key))
                {
                    pinyinList.Add(kvp.Value);
                    remainingText = remainingText.Replace(kvp.Key, string.Empty);
                }
            }

            // 处理剩余的字符
            foreach (char c in remainingText)
            {
                if (hanziPinyinDictionary.ContainsKey(c.ToString()))
                {
                    string pinyin = hanziPinyinDictionary[c.ToString()];
                    pinyinList.Add(pinyin);
                }
                else
                {
                    pinyinList.Add(c.ToString());
                }
            }

            return string.Join(" ", pinyinList);
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
            Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (sel.Type != PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("本功能可同时选中当前页面中与所选形状相同尺寸大小的形状");
                return;
            }

            Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            PowerPoint.ShapeRange range = sel.ShapeRange;

            bool isCtrlPressed = (Control.ModifierKeys & Keys.Control) == Keys.Control;
            bool isShiftPressed = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;

            if (sel.HasChildShapeRange)
            {
                Shape shape = sel.ChildShapeRange[1];

                for (int i = 1; i <= range[1].GroupItems.Count; i++)
                {
                    Shape item = range[1].GroupItems[i];
                    if (item.Type != MsoShapeType.msoGroup && item.Visible == MsoTriState.msoTrue)
                    {
                        if (isCtrlPressed && item.Width == shape.Width)
                        {
                            item.Select(MsoTriState.msoFalse);
                        }
                        else if (isShiftPressed && item.Height == shape.Height)
                        {
                            item.Select(MsoTriState.msoFalse);
                        }
                        else if (!isCtrlPressed && !isShiftPressed && item.Width == shape.Width && item.Height == shape.Height)
                        {
                            item.Select(MsoTriState.msoFalse);
                        }
                    }
                }
            }
            else
            {
                Shape shape2 = range[1];

                for (int j = 1; j <= slide.Shapes.Count; j++)
                {
                    PowerPoint.Shape item2 = slide.Shapes[j];
                    if (item2.Type != MsoShapeType.msoGroup && item2.Visible == MsoTriState.msoTrue)
                    {
                        if (isCtrlPressed && item2.Width == shape2.Width)
                        {
                            item2.Select(MsoTriState.msoFalse);
                        }
                        else if (isShiftPressed && item2.Height == shape2.Height)
                        {
                            item2.Select(MsoTriState.msoFalse);
                        }
                        else if (!isCtrlPressed && !isShiftPressed && item2.Width == shape2.Width && item2.Height == shape2.Height)
                        {
                            item2.Select(MsoTriState.msoFalse);
                        }
                    }
                }
            }
        }

        private void SelectedColor_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (sel.Type != PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("本功能可同时选中当前页面中与所选形状相同填充颜色的形状");
                return;
            }

            Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            PowerPoint.ShapeRange range = sel.ShapeRange;

            if (sel.HasChildShapeRange)
            {
                PowerPoint.Shape shape = sel.ChildShapeRange[1];
                if (shape.Fill.Type == Office.MsoFillType.msoFillSolid)
                {
                    for (int i = 1; i <= range[1].GroupItems.Count; i++)
                    {
                        PowerPoint.Shape item = range[1].GroupItems[i];
                        if (item.Type != Office.MsoShapeType.msoGroup && item.Visible == Office.MsoTriState.msoTrue && item.Fill.Type == shape.Fill.Type && item.Fill.ForeColor.RGB == shape.Fill.ForeColor.RGB)
                        {
                            item.Select(Office.MsoTriState.msoFalse);
                        }
                    }
                    return;
                }
                MessageBox.Show("所选形状不是纯色填充");
                return;
            }
            else
            {
                PowerPoint.Shape shape2 = range[1];
                if (shape2.Fill.Type == Office.MsoFillType.msoFillSolid)
                {
                    for (int j = 1; j <= slide.Shapes.Count; j++)
                    {
                        PowerPoint.Shape item2 = slide.Shapes[j];
                        if (item2.Type != Office.MsoShapeType.msoGroup && item2.Visible == Office.MsoTriState.msoTrue && item2.Fill.Type == shape2.Fill.Type && item2.Fill.ForeColor.RGB == shape2.Fill.ForeColor.RGB)
                        {
                            item2.Select(Office.MsoTriState.msoFalse);
                        }
                    }
                    return;
                }
                MessageBox.Show("所选形状不是纯色填充");
                return;
            }
        }

        private void Selectedline_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.PowerPoint.Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("本功能可同时选中当前页面中与所选形状相同线条颜色、宽度或类型的形状");
                return;
            }

            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            PowerPoint.ShapeRange range = sel.ShapeRange;

            bool isCtrlPressed = (Control.ModifierKeys & Keys.Control) == Keys.Control;
            bool isShiftPressed = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;

            if (isCtrlPressed)
            {
                // 按线条宽度筛选
                if (sel.HasChildShapeRange)
                {
                    PowerPoint.Shape shape = sel.ChildShapeRange[1];
                    if (shape.Line.Visible == Office.MsoTriState.msoTrue)
                    {
                        for (int i = 1; i <= range[1].GroupItems.Count; i++)
                        {
                            PowerPoint.Shape item = range[1].GroupItems[i];
                            if (item.Type != Office.MsoShapeType.msoGroup && item.Visible == Office.MsoTriState.msoTrue && item.Line.Visible == Office.MsoTriState.msoTrue && item.Line.Weight == shape.Line.Weight)
                            {
                                item.Select(Office.MsoTriState.msoFalse);
                            }
                        }
                        return;
                    }
                    MessageBox.Show("所选形状没有线条");
                    return;
                }
                else
                {
                    PowerPoint.Shape shape2 = range[1];
                    if (shape2.Line.Visible == Office.MsoTriState.msoTrue)
                    {
                        for (int j = 1; j <= slide.Shapes.Count; j++)
                        {
                            PowerPoint.Shape item2 = slide.Shapes[j];
                            if (item2.Type != Office.MsoShapeType.msoGroup && item2.Visible == Office.MsoTriState.msoTrue && item2.Line.Visible == Office.MsoTriState.msoTrue && item2.Line.Weight == shape2.Line.Weight)
                            {
                                item2.Select(Office.MsoTriState.msoFalse);
                            }
                        }
                        return;
                    }
                    MessageBox.Show("所选形状没有线条");
                    return;
                }
            }
            else if (isShiftPressed)
            {
                // 按线条类型筛选
                if (sel.HasChildShapeRange)
                {
                    PowerPoint.Shape shape = sel.ChildShapeRange[1];
                    if (shape.Line.Visible == Office.MsoTriState.msoTrue)
                    {
                        for (int i = 1; i <= range[1].GroupItems.Count; i++)
                        {
                            PowerPoint.Shape item = range[1].GroupItems[i];
                            if (item.Type != Office.MsoShapeType.msoGroup && item.Visible == Office.MsoTriState.msoTrue && item.Line.Visible == Office.MsoTriState.msoTrue && item.Line.DashStyle == shape.Line.DashStyle)
                            {
                                item.Select(Office.MsoTriState.msoFalse);
                            }
                        }
                        return;
                    }
                    MessageBox.Show("所选形状没有线条");
                    return;
                }
                else
                {
                    PowerPoint.Shape shape2 = range[1];
                    if (shape2.Line.Visible == Office.MsoTriState.msoTrue)
                    {
                        for (int j = 1; j <= slide.Shapes.Count; j++)
                        {
                            PowerPoint.Shape item2 = slide.Shapes[j];
                            if (item2.Type != Office.MsoShapeType.msoGroup && item2.Visible == Office.MsoTriState.msoTrue && item2.Line.Visible == Office.MsoTriState.msoTrue && item2.Line.DashStyle == shape2.Line.DashStyle)
                            {
                                item2.Select(Office.MsoTriState.msoFalse);
                            }
                        }
                        return;
                    }
                    MessageBox.Show("所选形状没有线条");
                    return;
                }
            }
            else
            {
                // 按线条颜色筛选
                if (sel.HasChildShapeRange)
                {
                    PowerPoint.Shape shape = sel.ChildShapeRange[1];
                    if (shape.Line.Visible == Office.MsoTriState.msoTrue)
                    {
                        for (int i = 1; i <= range[1].GroupItems.Count; i++)
                        {
                            PowerPoint.Shape item = range[1].GroupItems[i];
                            if (item.Type != Office.MsoShapeType.msoGroup && item.Visible == Office.MsoTriState.msoTrue && item.Line.Visible == Office.MsoTriState.msoTrue && item.Line.ForeColor.RGB == shape.Line.ForeColor.RGB)
                            {
                                item.Select(Office.MsoTriState.msoFalse);
                            }
                        }
                        return;
                    }
                    MessageBox.Show("所选形状没有线条");
                    return;
                }
                else
                {
                    PowerPoint.Shape shape2 = range[1];
                    if (shape2.Line.Visible == Office.MsoTriState.msoTrue)
                    {
                        for (int j = 1; j <= slide.Shapes.Count; j++)
                        {
                            PowerPoint.Shape item2 = slide.Shapes[j];
                            if (item2.Type != Office.MsoShapeType.msoGroup && item2.Visible == Office.MsoTriState.msoTrue && item2.Line.Visible == Office.MsoTriState.msoTrue && item2.Line.ForeColor.RGB == shape2.Line.ForeColor.RGB)
                            {
                                item2.Select(Office.MsoTriState.msoFalse);
                            }
                        }
                        return;
                    }
                    MessageBox.Show("所选形状没有线条");
                    return;
                }
            }
        }

        private void Selectfontsize_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.PowerPoint.Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("本功能可同时选中当前页面中与所选形状中相同字号大小的形状");
                return;
            }

            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            PowerPoint.ShapeRange range = sel.ShapeRange;

            if (sel.HasChildShapeRange)
            {
                PowerPoint.Shape shape = sel.ChildShapeRange[1];
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    float fontSize = shape.TextFrame.TextRange.Font.Size;
                    for (int i = 1; i <= range[1].GroupItems.Count; i++)
                    {
                        PowerPoint.Shape item = range[1].GroupItems[i];
                        if (item.HasTextFrame == Office.MsoTriState.msoTrue && item.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        {
                            if (item.TextFrame.TextRange.Font.Size == fontSize && item.Visible == Office.MsoTriState.msoTrue)
                            {
                                item.Select(Office.MsoTriState.msoFalse);
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("所选形状不包含文本");
                }
            }
            else
            {
                PowerPoint.Shape shape2 = range[1];
                if (shape2.HasTextFrame == Office.MsoTriState.msoTrue && shape2.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    float fontSize = shape2.TextFrame.TextRange.Font.Size;
                    for (int j = 1; j <= slide.Shapes.Count; j++)
                    {
                        PowerPoint.Shape item2 = slide.Shapes[j];
                        if (item2.HasTextFrame == Office.MsoTriState.msoTrue && item2.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        {
                            if (item2.TextFrame.TextRange.Font.Size == fontSize && item2.Visible == Office.MsoTriState.msoTrue)
                            {
                                item2.Select(Office.MsoTriState.msoFalse);
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("所选形状不包含文本");
                }
            }
        }

        private void Type_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.PowerPoint.Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("本功能可同时选中当前页面中与所选形状类型相同的形状");
                return;
            }

            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            PowerPoint.ShapeRange range = sel.ShapeRange;

            if (sel.HasChildShapeRange)
            {
                PowerPoint.Shape shape = sel.ChildShapeRange[1];
                for (int i = 1; i <= range[1].GroupItems.Count; i++)
                {
                    PowerPoint.Shape item = range[1].GroupItems[i];
                    if (item.Type == shape.Type && item.Visible == Office.MsoTriState.msoTrue)
                    {
                        item.Select(Office.MsoTriState.msoFalse);
                    }
                }
            }
            else
            {
                PowerPoint.Shape shape2 = range[1];
                for (int j = 1; j <= slide.Shapes.Count; j++)
                {
                    PowerPoint.Shape item2 = slide.Shapes[j];
                    if (item2.Type == shape2.Type && item2.Visible == Office.MsoTriState.msoTrue)
                    {
                        item2.Select(Office.MsoTriState.msoFalse);
                    }
                }
            }
        }

        private void 重叠组合_Click(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = application.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                var selectedShapes = GetShapesFromSelectionForOverlap(selection);
                var overlappingGroups = FindOverlappingGroups(selectedShapes);

                foreach (var group in overlappingGroups)
                {
                    if (group.Count > 1)
                    {
                        var shapeRange = application.ActiveWindow.Selection.SlideRange.Shapes.Range(group.Select(s => s.Name).ToArray());
                        shapeRange.Group();
                    }
                }
            }
            else
            {
                MessageBox.Show("请选择多个对象。");
            }
        }

        private List<PowerPoint.Shape> GetShapesFromSelectionForOverlap(PowerPoint.Selection selection)
        {
            List<PowerPoint.Shape> shapes = new List<PowerPoint.Shape>();
            for (int i = 1; i <= selection.ShapeRange.Count; i++)
            {
                shapes.Add(selection.ShapeRange[i]);
            }
            return shapes;
        }

        private List<List<PowerPoint.Shape>> FindOverlappingGroups(List<PowerPoint.Shape> shapes)
        {
            List<List<PowerPoint.Shape>> overlappingGroups = new List<List<PowerPoint.Shape>>();
            bool[] visited = new bool[shapes.Count];

            for (int i = 0; i < shapes.Count; i++)
            {
                if (!visited[i])
                {
                    List<PowerPoint.Shape> group = new List<PowerPoint.Shape>();
                    FindOverlappingShapes(shapes, visited, i, group);
                    overlappingGroups.Add(group);
                }
            }

            return overlappingGroups;
        }

        private void FindOverlappingShapes(List<PowerPoint.Shape> shapes, bool[] visited, int index, List<PowerPoint.Shape> group)
        {
            visited[index] = true;
            group.Add(shapes[index]);

            for (int i = 0; i < shapes.Count; i++)
            {
                if (!visited[i] && IsOverlapping(shapes[index], shapes[i]))
                {
                    FindOverlappingShapes(shapes, visited, i, group);
                }
            }
        }

        private bool IsOverlapping(PowerPoint.Shape shape1, PowerPoint.Shape shape2)
        {
            const float tolerance = 0.1f; // 极小的间距误差值

            float left1 = shape1.Left;
            float top1 = shape1.Top;
            float right1 = shape1.Left + shape1.Width;
            float bottom1 = shape1.Top + shape1.Height;

            float left2 = shape2.Left;
            float top2 = shape2.Top;
            float right2 = shape2.Left + shape2.Width;
            float bottom2 = shape2.Top + shape2.Height;

            // Check for overlap with tolerance
            return !(left1 >= right2 - tolerance || right1 <= left2 + tolerance || top1 >= bottom2 - tolerance || bottom1 <= top2 + tolerance);
        }

        private void 临近组合_Click(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = application.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                var selectedShapes = GetShapesFromSelectionForAdjacency(selection);
                var adjacentGroups = FindAdjacentGroups(selectedShapes);

                foreach (var group in adjacentGroups)
                {
                    if (group.Count > 1)
                    {
                        var shapeRange = application.ActiveWindow.Selection.SlideRange.Shapes.Range(group.Select(s => s.Name).ToArray());
                        shapeRange.Group();
                    }
                }
            }
            else
            {
                MessageBox.Show("请选择多个对象。");
            }
        }

        private List<PowerPoint.Shape> GetShapesFromSelectionForAdjacency(PowerPoint.Selection selection)
        {
            List<PowerPoint.Shape> shapes = new List<PowerPoint.Shape>();
            for (int i = 1; i <= selection.ShapeRange.Count; i++)
            {
                shapes.Add(selection.ShapeRange[i]);
            }
            return shapes;
        }

        private List<List<PowerPoint.Shape>> FindAdjacentGroups(List<PowerPoint.Shape> shapes)
        {
            List<List<PowerPoint.Shape>> adjacentGroups = new List<List<PowerPoint.Shape>>();
            bool[] visited = new bool[shapes.Count];

            for (int i = 0; i < shapes.Count; i++)
            {
                if (!visited[i])
                {
                    List<PowerPoint.Shape> group = new List<PowerPoint.Shape>();
                    FindAdjacentShapes(shapes, visited, i, group);
                    adjacentGroups.Add(group);
                }
            }

            return adjacentGroups;
        }

        private void FindAdjacentShapes(List<PowerPoint.Shape> shapes, bool[] visited, int index, List<PowerPoint.Shape> group)
        {
            visited[index] = true;
            group.Add(shapes[index]);

            for (int i = 0; i < shapes.Count; i++)
            {
                if (!visited[i] && IsAdjacent(shapes[index], shapes[i]))
                {
                    FindAdjacentShapes(shapes, visited, i, group);
                }
            }
        }

        private bool IsAdjacent(PowerPoint.Shape shape1, PowerPoint.Shape shape2)
        {
            const float tolerance = 0.1f; // 极小的间距误差值

            float left1 = shape1.Left;
            float top1 = shape1.Top;
            float right1 = shape1.Left + shape1.Width;
            float bottom1 = shape1.Top + shape1.Height;

            float left2 = shape2.Left;
            float top2 = shape2.Top;
            float right2 = shape2.Left + shape2.Width;
            float bottom2 = shape2.Top + shape2.Height;

            // Check for horizontal adjacency
            bool horizontallyAdjacent = (Math.Abs(right1 - left2) <= tolerance || Math.Abs(left1 - right2) <= tolerance) && !(top1 >= bottom2 || bottom1 <= top2);

            // Check for vertical adjacency
            bool verticallyAdjacent = (Math.Abs(bottom1 - top2) <= tolerance || Math.Abs(top1 - bottom2) <= tolerance) && !(left1 >= right2 || right1 <= left2);

            return horizontallyAdjacent || verticallyAdjacent;
        }
        private void 同色组合_Click(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = application.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                var allShapes = new List<PowerPoint.Shape>();
                var tag = Guid.NewGuid().ToString(); // 生成唯一的标签
                var groupedShapes = new List<PowerPoint.Shape>(); // 用于存储新的组合对象

                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    // 如果形状是一个组合，记录组合中的子形状，并添加Shape Tag
                    if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                    {
                        try
                        {
                            var ungroupedShapes = shape.Ungroup();

                            for (int i = 1; i <= ungroupedShapes.Count; i++)
                            {
                                var ungroupedShape = ungroupedShapes[i];
                                ungroupedShape.Tags.Add("GroupTag", tag);
                                allShapes.Add(ungroupedShape);
                            }
                        }
                        catch (System.Runtime.InteropServices.COMException ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Unable to ungroup shape: {ex.Message}");
                            continue;
                        }
                    }
                    else
                    {
                        shape.Tags.Add("GroupTag", tag);
                        allShapes.Add(shape);
                    }
                }

                // 按颜色分组
                var colorGroups = GroupShapesByFillColor(allShapes);

                // 对每组相同颜色的形状进行组合
                foreach (var group in colorGroups.Values)
                {
                    if (group.Count > 1)
                    {
                        var shapeRange = application.ActiveWindow.Selection.SlideRange.Shapes.Range(group.Select(s => s.Name).ToArray());
                        var newGroup = shapeRange.Group();

                        // 删除组合后的Shape Tag
                        newGroup.Tags.Delete("GroupTag");

                        // 将新组合的形状加入到groupedShapes列表中
                        groupedShapes.Add(newGroup);
                    }
                }

                // 删除未组合形状的Shape Tag
                foreach (var shape in allShapes)
                {
                    if (shape.Tags["GroupTag"] == tag)
                    {
                        shape.Tags.Delete("GroupTag");
                    }
                }

                // 如果有新的组合形状，自动选中这些组合
                if (groupedShapes.Count > 0)
                {
                    var newGroupNames = groupedShapes.Select(g => g.Name).ToArray();
                    var newSelectionRange = application.ActiveWindow.Selection.SlideRange.Shapes.Range(newGroupNames);
                    newSelectionRange.Select(); // 自动选中新组合的形状
                }
            }
            else
            {
                MessageBox.Show("请选择多个对象。");
            }
        }

        private List<PowerPoint.Shape> GetShapesFromSelectionForSameColor(PowerPoint.Selection selection)
        {
            List<PowerPoint.Shape> shapes = new List<PowerPoint.Shape>();
            for (int i = 1; i <= selection.ShapeRange.Count; i++)
            {
                shapes.Add(selection.ShapeRange[i]);
            }
            return shapes;
        }

        private Dictionary<int, List<PowerPoint.Shape>> GroupShapesByFillColor(List<PowerPoint.Shape> shapes)
        {
            Dictionary<int, List<PowerPoint.Shape>> colorGroups = new Dictionary<int, List<PowerPoint.Shape>>();

            foreach (var shape in shapes)
            {
                int color = shape.Fill.ForeColor.RGB;

                if (!colorGroups.ContainsKey(color))
                {
                    colorGroups[color] = new List<PowerPoint.Shape>();
                }
                colorGroups[color].Add(shape);
            }

            return colorGroups;
        }


        private void 沿线分布_Click(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            var selection = application.ActiveWindow.Selection;

            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 1)
            {
                var lineShape = selection.ShapeRange[1];
                if (lineShape.Type == MsoShapeType.msoLine || lineShape.Type == MsoShapeType.msoFreeform)
                {
                    List<Shape> shapesToDistribute = new List<Shape>();
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

        private void DistributeShapesAlongLine(Shape lineShape, List<Shape> shapesToDistribute)
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

      

        private void 图形分割_Click(object sender, RibbonControlEventArgs e)
        {
            图形分割 form = new 图形分割();
            form.Show();
        }

        private void 快捷盒子_Click(object sender, RibbonControlEventArgs e)
        {
            快捷盒子 form = new 快捷盒子();
            form.Show();
        }

      

        private void 形状填图_Click(object sender, RibbonControlEventArgs e)
        {
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
                Title = "Select images to fill the selected shapes"
            };

            if (openFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            var selectedFiles = openFileDialog.FileNames;
            int fileIndex = 0;

            foreach (var shape in selectedShapes)
            {
                FillShapeWithImage(shape, selectedFiles, ref fileIndex);
            }
        }

        private void FillShapeWithImage(Shape shape, string[] selectedFiles, ref int fileIndex)
        {
            if (fileIndex >= selectedFiles.Length)
                return;

            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                foreach (Shape subShape in shape.GroupItems)
                {
                    FillShapeWithImage(subShape, selectedFiles, ref fileIndex);
                }
            }
            else
            {
                var filePath = selectedFiles[fileIndex++];
                shape.Fill.UserPicture(filePath);
            }
        }

        private void 批量换图_Click(object sender, RibbonControlEventArgs e)
        {
            Application app = Globals.ThisAddIn.Application;
            Presentation presentation = app.ActivePresentation;
            Slide slide = app.ActiveWindow.View.Slide;

            // 获取当前选中的图片
            List<Shape> selectedShapes = new List<Shape>();
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange rng = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    rng = sel.ChildShapeRange;
                }

                foreach (Shape shape in rng)
                {
                    if (shape.Type == Office.MsoShapeType.msoPicture)
                    {
                        selectedShapes.Add(shape);
                    }
                }
            }

            if (selectedShapes.Count == 0)
            {
                MessageBox.Show("请先选择一张或多张图片。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 打开文件对话框让用户选择新图片
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "图片文件|*.jpg;*.jpeg;*.png;*.bmp;*.gif"
            };

            if (openFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            string[] selectedFiles = openFileDialog.FileNames;
            if (selectedFiles.Length != selectedShapes.Count)
            {
                MessageBox.Show("请确保选择的图片数量与原图片数量一致。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 插入新图片并替换原图片
            for (int i = 0; i < selectedShapes.Count; i++)
            {
                Shape originalShape = selectedShapes[i];
                string newImagePath = selectedFiles[i];

                // 插入新图片
                Shape newShape = slide.Shapes.AddPicture(newImagePath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoCTrue,
                    originalShape.Left, originalShape.Top, originalShape.Width, originalShape.Height);

                // 调整新图片大小以匹配原图片
                ResizeAndCropPicture(newShape, originalShape);

                // 复制格式
                originalShape.PickUp();
                newShape.Apply();

                // 删除原图片
                originalShape.Delete();
            }

            MessageBox.Show("图片替换完成。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ResizeAndCropPicture(Shape newShape, Shape originalShape)
        {
            Application app = Globals.ThisAddIn.Application;

            // 调整新图片大小
            newShape.LockAspectRatio = Office.MsoTriState.msoTrue;

            float pw = originalShape.Width;
            float ph = originalShape.Height;
            float pn = pw / ph;

            newShape.ScaleWidth(1f, Office.MsoTriState.msoTrue, Office.MsoScaleFrom.msoScaleFromMiddle);
            newShape.ScaleHeight(1f, Office.MsoTriState.msoTrue, Office.MsoScaleFrom.msoScaleFromMiddle);
            newShape.PictureFormat.CropLeft = 0f;
            newShape.PictureFormat.CropRight = 0f;
            newShape.PictureFormat.CropTop = 0f;
            newShape.PictureFormat.CropBottom = 0f;

            if (newShape.Width >= newShape.Height)
            {
                if (newShape.Width - newShape.Height * pn >= 0f)
                {
                    float n2 = (newShape.Width - newShape.Height * pn) / 2f;
                    newShape.PictureFormat.CropLeft = n2;
                    newShape.PictureFormat.CropRight = n2;
                }
                else
                {
                    float n2 = (newShape.Height - newShape.Width / pn) / 2f;
                    newShape.PictureFormat.CropTop = n2;
                    newShape.PictureFormat.CropBottom = n2;
                }
            }
            else if (newShape.Height - newShape.Width / pn >= 0f)
            {
                float n2 = (newShape.Height - newShape.Width / pn) / 2f;
                newShape.PictureFormat.CropTop = n2;
                newShape.PictureFormat.CropBottom = n2;
            }
            else
            {
                float n2 = (newShape.Width - newShape.Height * pn) / 2f;
                newShape.PictureFormat.CropLeft = n2;
                newShape.PictureFormat.CropRight = n2;
            }

            // 调整新图片位置和大小
            newShape.Width = pw;
            newShape.Height = ph;
            newShape.Left = originalShape.Left;
            newShape.Top = originalShape.Top;
        }

        private void 原位转JPG_Click(object sender, RibbonControlEventArgs e)
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
                    var pictureShape = slide.Shapes.PasteSpecial(PpPasteDataType.ppPasteJPG)[1];

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

       
        private void 图层_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 获取当前选中的形状对象
                Microsoft.Office.Interop.PowerPoint.Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                // 检查是否选中的是形状
                if (sel.Type != PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("本功能可同时选中当前页面中与所选形状前缀名相同的所有形状");
                    return;
                }

                Microsoft.Office.Interop.PowerPoint.ShapeRange range = sel.ShapeRange;

                // 检查是否选中了形状
                if (range.Count == 0)
                {
                    MessageBox.Show("请选择一个形状");
                    return;
                }

                // 获取选中形状的名称前缀
                string selectedShapeName = range[1].Name;
                string prefix = GetPrefix(selectedShapeName);

                // 遍历幻灯片中的所有形状
                Microsoft.Office.Interop.PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    Microsoft.Office.Interop.PowerPoint.Shape shape = slide.Shapes[i];
                    // 如果形状的名称前缀与选中形状的前缀相同且形状可见
                    if (GetPrefix(shape.Name) == prefix && shape.Visible == Office.MsoTriState.msoTrue)
                    {
                        shape.Select(Office.MsoTriState.msoFalse);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("发生错误: " + ex.Message);
            }
        }

        // 获取形状名称的前缀
        private string GetPrefix(string shapeName)
        {
            // 找到第一个数字的位置
            int index = shapeName.IndexOfAny("0123456789".ToCharArray());
            if (index > 0)
            {
                return shapeName.Substring(0, index);
            }
            return shapeName;
        }

        private void 统一控点_Click(object sender, RibbonControlEventArgs e)
        {
            Application pptApp = Globals.ThisAddIn.Application;
            Selection selection = pptApp.ActiveWindow.Selection;

            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange shapes = selection.ShapeRange;

                // 检查是否所有选中的形状都具有控点
                if (shapes.Cast<Shape>().All(shape => shape.Adjustments.Count > 0))
                {
                    // 获取第一个选中的对象的控点大小作为参考
                    var referenceAdjustments = new float[shapes[1].Adjustments.Count];
                    for (int i = 1; i <= shapes[1].Adjustments.Count; i++)
                    {
                        referenceAdjustments[i - 1] = shapes[1].Adjustments[i];
                    }

                    // 遍历所有选择的形状并设置控点大小
                    for (int i = 2; i <= shapes.Count; i++)
                    {
                        for (int j = 1; j <= shapes[i].Adjustments.Count; j++)
                        {
                            shapes[i].Adjustments[j] = referenceAdjustments[j - 1];
                        }
                    }
                }
                else
                {
                    MessageBox.Show("请选择多个具有控点的对象。");
                }
            }
            else
            {
                MessageBox.Show("请选择多个对象统一控点。");
            }
        }

        private void 文本矢量化_Click(object sender, RibbonControlEventArgs e)
        {
            Application pptApp = Globals.ThisAddIn.Application;
            Selection selection = pptApp.ActiveWindow.Selection;

            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count >= 1)
            {
                PowerPoint.ShapeRange selectedShapes = selection.ShapeRange;

                foreach (Shape selectedShape in selectedShapes)
                {
                    if (selectedShape.Type == MsoShapeType.msoTextBox)
                    {
                        Slide slide = pptApp.ActiveWindow.View.Slide;

                        // 在页面以外左上角插入一个小正方形
                        float squareSize = 50; // 正方形边长
                        float leftPosition = -squareSize; // 移动到页面以外
                        float topPosition = -squareSize;  // 移动到页面以外
                        Shape squareShape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, leftPosition, topPosition, squareSize, squareSize);

                        // 选中正方形和文本框
                        selectedShape.Select();
                        squareShape.Select(MsoTriState.msoFalse);

                        // 创建一个 ShapeRange 包含选中的形状
                        PowerPoint.ShapeRange shapeRange = slide.Shapes.Range(new object[] { selectedShape.Name, squareShape.Name });

                        // 执行“剪除”操作
                        shapeRange.MergeShapes(MsoMergeCmd.msoMergeSubtract);
                    }
                }
            }
            else
            {
                MessageBox.Show("请选择要转矢量的文本框。");
            }
        }


        private void 删除动画_Click(object sender, RibbonControlEventArgs e)
        {
           Application pptApp = Globals.ThisAddIn.Application;
           Selection selection = pptApp.ActiveWindow.Selection;

            if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
            {
                // 删除当前页所有对象的动画
                Slide currentSlide = pptApp.ActiveWindow.View.Slide;
                DeleteAnimationsFromSlide(currentSlide);
            }
            else if ((Control.ModifierKeys & Keys.Shift) == Keys.Shift)
            {
                // 删除所有幻灯片中的动画
                Presentation presentation = pptApp.ActivePresentation;
                foreach (Slide slide in presentation.Slides)
                {
                    DeleteAnimationsFromSlide(slide);
                }
            }
            else
            {
                // 删除选中对象的动画
                if (selection.Type == PpSelectionType.ppSelectionShapes)
                {
                    foreach (Shape shape in selection.ShapeRange)
                    {
                        DeleteAnimationsFromShape(shape);
                    }
                }
                else
                {
                    MessageBox.Show("请选择要删除动画的对象。");
                }
            }
        }

        private void DeleteAnimationsFromShape(Shape shape)
        {
            shape.AnimationSettings.Animate = MsoTriState.msoFalse;
        }

        private void DeleteAnimationsFromSlide(Slide slide)
        {
            foreach (Shape shape in slide.Shapes)
            {
                DeleteAnimationsFromShape(shape);
            }
        }

        private void 清空页外_Click(object sender, RibbonControlEventArgs e)
        {
            Application pptApp = Globals.ThisAddIn.Application;
            Selection selection = pptApp.ActiveWindow.Selection;

            if (selection.Type == PpSelectionType.ppSelectionSlides)
            {
                SlideRange slideRange = selection.SlideRange;
                int slidesProcessed = 0;

                foreach (Slide slide in slideRange)
                {
                    if (DeleteOffPageObjects(slide))
                    {
                        slidesProcessed++;
                    }
                }

                MessageBox.Show($"已成功删除 {slidesProcessed} 页的页面外部元素。");
            }
            else
            {
                MessageBox.Show("请选择要删除页面外部元素的幻灯片页面。");
            }
        }

        private bool DeleteOffPageObjects(Slide slide)
        {
            float slideWidth = slide.Parent.PageSetup.SlideWidth;
            float slideHeight = slide.Parent.PageSetup.SlideHeight;
            bool objectsDeleted = false;

            for (int i = slide.Shapes.Count; i >= 1; i--)
            {
                Shape shape = slide.Shapes[i];
                if (shape.Left + shape.Width < 0 || shape.Left > slideWidth || shape.Top + shape.Height < 0 || shape.Top > slideHeight)
                {
                    shape.Delete();
                    objectsDeleted = true;
                }
            }

            return objectsDeleted;
        }
    

        private void 清除备注_Click(object sender, RibbonControlEventArgs e)
        {
            Application pptApp = Globals.ThisAddIn.Application;
            Selection selection = pptApp.ActiveWindow.Selection;

            if (selection.Type == PpSelectionType.ppSelectionSlides)
            {
                SlideRange selectedSlides = selection.SlideRange;
                foreach (Slide slide in selectedSlides)
                {
                    slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text = string.Empty;
                }
            }
            else
            {
                MessageBox.Show("请选择要清楚备注的幻灯片页面。");
            }
        }

        private void 清除超链接_Click(object sender, RibbonControlEventArgs e)
        {
            Application pptApp = Globals.ThisAddIn.Application;
            Selection sel = pptApp.ActiveWindow.Selection;
            bool isSuccessful = false;

            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                int count = range.Count;
                for (int i = 1; i <= count; i++)
                {
                    Shape shape = range[i];
                    if (shape.ActionSettings[PpMouseActivation.ppMouseClick].Action != PpActionType.ppActionNone)
                    {
                        shape.ActionSettings[PpMouseActivation.ppMouseClick].Hyperlink.Delete();
                        isSuccessful = true;
                    }
                    if (shape.ActionSettings[PpMouseActivation.ppMouseOver].Action != PpActionType.ppActionNone)
                    {
                        shape.ActionSettings[PpMouseActivation.ppMouseOver].Hyperlink.Delete();
                        isSuccessful = true;
                    }
                    if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                    {
                        if (shape.TextFrame.TextRange.ActionSettings[PpMouseActivation.ppMouseClick].Action != PpActionType.ppActionNone)
                        {
                            shape.TextFrame.TextRange.ActionSettings[PpMouseActivation.ppMouseClick].Hyperlink.Delete();
                            isSuccessful = true;
                        }
                        if (shape.TextFrame.TextRange.ActionSettings[PpMouseActivation.ppMouseOver].Action != PpActionType.ppActionNone)
                        {
                            shape.TextFrame.TextRange.ActionSettings[PpMouseActivation.ppMouseOver].Hyperlink.Delete();
                            isSuccessful = true;
                        }
                    }
                }
                if (isSuccessful)
                {
                    MessageBox.Show("已成功清除所选对象的超链接。");
                }
                return;
            }

            if (sel.Type == PpSelectionType.ppSelectionSlides)
            {
                SlideRange srange = sel.SlideRange;
                for (int j = 1; j <= srange.Count; j++)
                {
                    int count2 = srange[j].Hyperlinks.Count;
                    if (count2 > 0)
                    {
                        for (int k = count2; k >= 1; k--)
                        {
                            srange[j].Hyperlinks[k].Delete();
                            isSuccessful = true;
                        }
                    }
                }
                if (isSuccessful)
                {
                    MessageBox.Show("已成功清除所选页面的超链接。");
                }
                return;
            }

            if (sel.Type == PpSelectionType.ppSelectionText)
            {
                if (sel.TextRange.ActionSettings[PpMouseActivation.ppMouseClick].Action != PpActionType.ppActionNone)
                {
                    sel.TextRange.ActionSettings[PpMouseActivation.ppMouseClick].Hyperlink.Delete();
                    isSuccessful = true;
                }
                if (sel.TextRange.ActionSettings[PpMouseActivation.ppMouseOver].Action != PpActionType.ppActionNone)
                {
                    sel.TextRange.ActionSettings[PpMouseActivation.ppMouseOver].Hyperlink.Delete();
                    isSuccessful = true;
                }
                if (isSuccessful)
                {
                    MessageBox.Show("已成功清除所选对象的超链接。");
                }
                return;
            }

            MessageBox.Show("请先选中要删除超链接的幻灯片页面或对象。");
        }
    

         private void 删除未用版式_Click(object sender, RibbonControlEventArgs e)
        {
            Application pptApp = Globals.ThisAddIn.Application;
            Designs designs = pptApp.ActivePresentation.Designs;
            int deletedCount = 0;

            for (int j = designs.Count; j >= 1; j--)
            {
                Design design = designs[j];
                CustomLayouts customLayouts = design.SlideMaster.CustomLayouts;

                for (int k = customLayouts.Count; k >= 1; k--)
                {
                    CustomLayout layout = customLayouts[k];

                    if (!IsLayoutUsed(pptApp.ActivePresentation, layout))
                    {
                        try
                        {
                            layout.Delete();
                            deletedCount++;
                        }
                        catch
                        {
                            // 处理删除失败的情况
                        }
                    }
                }

                // 如果自定义版式全部删除后，删除设计
                if (design.SlideMaster.CustomLayouts.Count == 0)
                {
                    design.Delete();
                }
            }

            MessageBox.Show("已删除 " + deletedCount + " 个未使用版式");
        }

        private bool IsLayoutUsed(Presentation presentation, CustomLayout layout)
        {
            foreach (Slide slide in presentation.Slides)
            {
                if (slide.CustomLayout == layout)
                {
                    return true;
                }
            }
            return false;
        }

       

        private void 分解笔顺_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                Selection sel = app.ActiveWindow.Selection;

                // 确保用户选中一个组合
                if (sel.Type != PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count != 1 || sel.ShapeRange[1].Type != MsoShapeType.msoGroup)
                {
                    MessageBox.Show("请选择需要填色分解的汉字笔画组合。");
                    return;
                }

                PowerPoint.Shape groupShape = sel.ShapeRange[1];
                PowerPoint.GroupShapes groupItems = groupShape.GroupItems;
                int itemCount = groupItems.Count;

                // Create new groups based on the number of items in the original group
                List<PowerPoint.Shape> newGroups = new List<PowerPoint.Shape>();
                for (int i = 0; i < itemCount; i++)
                {
                    // Duplicate the original group
                    PowerPoint.Shape newGroup = groupShape.Duplicate()[1];
                    newGroup.Left += (i + 1) * (groupShape.Width + 10); // Adjust position
                    newGroups.Add(newGroup);
                }

                // Check if Ctrl key is pressed
                bool isCtrlPressed = (Control.ModifierKeys & Keys.Control) == Keys.Control;

                // Set colors and remove borders based on the pattern
                for (int i = 0; i < newGroups.Count; i++)
                {
                    PowerPoint.Shape newGroup = newGroups[i];
                    PowerPoint.GroupShapes newGroupItems = newGroup.GroupItems;

                    for (int j = 1; j <= itemCount; j++)
                    {
                        newGroupItems[j].Line.Visible = MsoTriState.msoFalse; // Remove border

                        if (j <= i + 1)
                        {
                            newGroupItems[j].Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);
                        }
                        if (j == i + 1)
                        {
                            newGroupItems[j].Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Red);
                        }
                        if (j > i + 1)
                        {
                            newGroupItems[j].Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray);
                        }
                    }

                    // 命名新组合
                    newGroup.Name = $"【G】：分步第{i + 1}笔";

                    // Adjust positions for two-row layout if Ctrl key is pressed
                    if (isCtrlPressed)
                    {
                        int columns = (int)Math.Ceiling(newGroups.Count / 2.0);
                        int row = i / columns;
                        int column = i % columns;

                        newGroup.Left = groupShape.Left + column * (groupShape.Width + 10);
                        newGroup.Top = groupShape.Top + row * (groupShape.Height + 10);
                    }
                }

                // 删除原来的组合形状
                groupShape.Delete();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"发生错误：{ex.Message}");
            }
        }

        private void 部首描红_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 设置EPPlus的许可证上下文
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // 加载嵌入资源的Excel文件
                string filePath = ExtractEmbeddedResource("课件帮PPT助手.汉字字典.汉字字典.xlsx");

                Application app = Globals.ThisAddIn.Application;
                Selection sel = app.ActiveWindow.Selection;

                if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                    Shape selectedShape;

                    // 如果所选对象不是组合对象，自动将其组合
                    if (sel.ShapeRange.Count > 1 || (sel.ShapeRange.Count == 1 && sel.ShapeRange[1].Type != MsoShapeType.msoGroup))
                    {
                        // 自动将所选对象组合起来
                        selectedShape = sel.ShapeRange.Group();
                    }
                    else
                    {
                        selectedShape = sel.ShapeRange[1];
                    }

                    // 检查组合对象并获取其中第一个形状的前缀名
                    if (selectedShape.Type == MsoShapeType.msoGroup)
                    {
                        var firstShapeName = selectedShape.GroupItems[1].Name;
                        var prefixName = firstShapeName.Split('-')[0].Trim();

                        // 加载 Excel 数据
                        var hanziStrokeOrderDictionary = new Dictionary<string, string>();

                        if (!File.Exists(filePath))
                        {
                            MessageBox.Show($"未找到文件：{filePath}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        using (var package = new ExcelPackage(new FileInfo(filePath)))
                        {
                            foreach (var worksheet in package.Workbook.Worksheets)
                            {
                                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                                {
                                    string hanziFromExcel = worksheet.Cells[row, 1].Text;
                                    string strokeOrderValue = worksheet.Cells[row, 8].Text;
                                    hanziStrokeOrderDictionary[hanziFromExcel] = strokeOrderValue;
                                }
                            }
                        }

                        if (hanziStrokeOrderDictionary.TryGetValue(prefixName, out string foundStrokeOrder))
                        {
                            var strokeIndices = foundStrokeOrder.Split(',').Select(int.Parse).ToList();
                            for (int i = 1; i <= selectedShape.GroupItems.Count; i++)
                            {
                                PowerPoint.Shape subShape = selectedShape.GroupItems[i];
                                if (strokeIndices.Contains(i))
                                {
                                    subShape.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Red);
                                }
                                else
                                {
                                    subShape.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("未找到该汉字的部首笔画序列。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        MessageBox.Show("请选中需要部首描红的汉字笔画组合。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("所选对象非组合对象。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string Resource(string resourceName)
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


        private void 关于我_Click(object sender, RibbonControlEventArgs e)
        {
            OpenWebPage("https://flowus.cn/andyblog/share/6da481ac-a57b-4214-9ce8-94273bbf2f45?code=GEH4ZC");
        }

        private void 检查更新_Click(object sender, RibbonControlEventArgs e)
        {
            OpenWebPage("https://flowus.cn/andyblog/share/d3ba4de8-3319-476e-ab7a-260bbf8add5b?code=GEH4ZC");
        }


        private void 文本居中_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = app.ActiveWindow.Selection;
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                var shape = selection.ShapeRange[1];
                if (shape.HasTable == Office.MsoTriState.msoTrue)
                {
                    PowerPoint.Table table = shape.Table;

                    // 遍历所有行
                    for (int i = 1; i <= table.Rows.Count; i++)
                    {
                        bool rowSelected = false;

                        // 检查当前行是否被选中
                        for (int j = 1; j <= table.Columns.Count; j++)
                        {
                            if ((bool)table.Cell(i, j).Selected)
                            {
                                rowSelected = true;
                                break;
                            }
                        }

                        if (rowSelected)
                        {
                            // 读取选中行第一个单元格的文本格式
                            var firstCellFormat = table.Cell(i, 1).Shape.TextFrame.TextRange.Font;

                            // 读取选中行的内容
                            var contents = new System.Collections.Generic.List<string>();
                            for (int j = 1; j <= table.Columns.Count; j++)
                            {
                                var cell = table.Cell(i, j);
                                if (!string.IsNullOrEmpty(cell.Shape.TextFrame.TextRange.Text))
                                {
                                    contents.Add(cell.Shape.TextFrame.TextRange.Text);
                                    cell.Shape.TextFrame.TextRange.Text = string.Empty; // 清空单元格内容
                                }
                            }

                            if (contents.Count > 0)
                            {
                                // 计算中间位置
                                int totalCells = table.Columns.Count;
                                int emptyCells = totalCells - contents.Count;
                                int startCol = (emptyCells / 2) + 1;

                                // 将内容写入到中间位置，并应用格式
                                for (int j = 0; j < contents.Count; j++)
                                {
                                    var targetCell = table.Cell(i, startCol + j);
                                    targetCell.Shape.TextFrame.TextRange.Text = contents[j];
                                    
                                }
                            }
                        }
                    }
                }
            }
        }

        private void 自动补齐_Click(object sender, RibbonControlEventArgs e)
        {
            Application app = Globals.ThisAddIn.Application;
            Selection selection = app.ActiveWindow.Selection;
            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                var shape = selection.ShapeRange[1];
                if (shape.HasTable == MsoTriState.msoTrue)
                {
                    Table table = shape.Table;

                    // 获取选中的行
                    int[] selectedRows = GetSelectedRows(table);

                    foreach (int rowIndex in selectedRows)
                    {
                        Row row = table.Rows[rowIndex];

                        // 找到第一个有内容的单元格，用于复制文字格式
                        TextRange firstNonEmptyTextRange = null;
                        foreach (Cell cell in row.Cells)
                        {
                            if (!string.IsNullOrEmpty(cell.Shape.TextFrame.TextRange.Text.Trim()))
                            {
                                firstNonEmptyTextRange = cell.Shape.TextFrame.TextRange;
                                break;
                            }
                        }

                        if (firstNonEmptyTextRange != null)
                        {
                            // 计算前面连续空白单元格数量
                            int consecutiveEmptyCells = 0;
                            foreach (Cell cell in row.Cells)
                            {
                                if (string.IsNullOrEmpty(cell.Shape.TextFrame.TextRange.Text.Trim()))
                                {
                                    consecutiveEmptyCells++;
                                }
                                else
                                {
                                    break;
                                }
                            }

                            // 为非前几个连续空白格子添加零宽度空格符，并复制文字格式
                            bool inNonLeadingEmptyCells = false;
                            foreach (Cell cell in row.Cells)
                            {
                                if (string.IsNullOrEmpty(cell.Shape.TextFrame.TextRange.Text))
                                {
                                    if (consecutiveEmptyCells > 0)
                                    {
                                        consecutiveEmptyCells--;
                                    }
                                    else
                                    {
                                        inNonLeadingEmptyCells = true;
                                    }

                                    if (inNonLeadingEmptyCells)
                                    {
                                        cell.Shape.TextFrame.TextRange.Text = "\u200B"; // 添加零宽度空格符
                                        cell.Shape.TextFrame.TextRange.Font.Name = firstNonEmptyTextRange.Font.Name;
                                        cell.Shape.TextFrame.TextRange.Font.Size = firstNonEmptyTextRange.Font.Size;
                                        cell.Shape.TextFrame.TextRange.Font.Bold = firstNonEmptyTextRange.Font.Bold;
                                        cell.Shape.TextFrame.TextRange.Font.Italic = firstNonEmptyTextRange.Font.Italic;
                                        cell.Shape.TextFrame.TextRange.Font.Underline = firstNonEmptyTextRange.Font.Underline;
                                    }
                                }
                            }

                            // 查找行最前面的第一个空单元格
                            int firstEmptyCell = -1;
                            for (int j = 1; j <= table.Columns.Count; j++)
                            {
                                if (string.IsNullOrEmpty(row.Cells[j].Shape.TextFrame.TextRange.Text.Trim()))
                                {
                                    firstEmptyCell = j;
                                    break;
                                }
                            }

                            if (firstEmptyCell != -1)
                            {
                                // 将内容整体前移，填补空白格子，但保留有内容格子之间的空白格子
                                int currentIndex = firstEmptyCell;
                                for (int j = firstEmptyCell; j <= table.Columns.Count; j++)
                                {
                                    if (!string.IsNullOrEmpty(row.Cells[j].Shape.TextFrame.TextRange.Text.Trim()))
                                    {
                                        row.Cells[currentIndex].Shape.TextFrame.TextRange.Text = row.Cells[j].Shape.TextFrame.TextRange.Text;
                                        if (currentIndex != j)
                                        {
                                            row.Cells[j].Shape.TextFrame.TextRange.Text = string.Empty;
                                        }
                                        currentIndex++;
                                    }
                                }
                            }

                            // 设置行对齐方式和文字格式
                            for (int j = 1; j <= table.Columns.Count; j++)
                            {
                                TextRange textRange = row.Cells[j].Shape.TextFrame.TextRange;

                                textRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter; // 设置居中对齐

                                if (rowIndex % 2 == 1) // 奇数行
                                {
                                    row.Cells[j].Shape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorBottom;
                                }

                                textRange.Font.Name = firstNonEmptyTextRange.Font.Name;
                                textRange.Font.Size = firstNonEmptyTextRange.Font.Size;
                                textRange.Font.Bold = firstNonEmptyTextRange.Font.Bold;
                            }
                        }
                    }
                }
            }
        }

        private int[] GetSelectedRows(Table table)
        {
            var selectedRows = new List<int>();
            for (int i = 1; i <= table.Rows.Count; i++)
            {
                for (int j = 1; j <= table.Columns.Count; j++)
                {
                    if (table.Cell(i, j).Selected)
                    {
                        selectedRows.Add(i);
                        break;
                    }
                }
            }
            return selectedRows.Distinct().ToArray();
        }


        private void 注音编辑_Click(object sender, RibbonControlEventArgs e)
        {
            // 检查是否已激活
            if (Properties.Settings.Default.IsActivated)
            {
                // 插件已激活，允许使用功能
                ZhuYinEditor editor = new ZhuYinEditor();
                editor.Show();
            }
            else
            {
                // 插件未激活，提示用户激活
                System.Windows.Forms.MessageBox.Show("请激活插件以使用此功能。", "未激活", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        private void 删列补行_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前激活的PPT应用程序实例
            Application application = Globals.ThisAddIn.Application;

            // 获取当前选中的对象
            Selection selection = application.ActiveWindow.Selection;

            // 检查是否按下Ctrl键
            bool isCtrlPressed = (Control.ModifierKeys & Keys.Control) == Keys.Control;

            int columnsToDelete = 1;

            // 如果按下Ctrl键，显示输入窗口
            if (isCtrlPressed)
            {
                using (FormInputColumns form = new FormInputColumns())
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        columnsToDelete = form.ColumnsToDelete;
                    }
                    else
                    {
                        return;
                    }
                }
            }

            // 确保选中的是表格
            if (selection.Type == PpSelectionType.ppSelectionShapes &&
                selection.ShapeRange.Count == 1 &&
                selection.ShapeRange[1].HasTable == MsoTriState.msoTrue)
            {
                PowerPoint.Table table = selection.ShapeRange[1].Table;

                // 获取表格的当前行数和列数
                int rowCount = table.Rows.Count;
                int colCount = table.Columns.Count;

                // 确保表格至少有两列且待删除列数不超过现有列数
                if (colCount > columnsToDelete)
                {
                    // 创建两个临时列表来存储拼音和汉字内容及其样式
                    List<(string text, float fontSize, string fontName, MsoTriState bold)> pinyinList = new List<(string, float, string, MsoTriState)>();
                    List<(string text, float fontSize, string fontName, MsoTriState bold)> hanziList = new List<(string, float, string, MsoTriState)>();

                    // 将表格中的拼音和汉字内容及其样式分别存储到临时列表中
                    for (int i = 1; i <= rowCount; i += 2)
                    {
                        for (int j = 1; j <= colCount; j++)
                        {
                            var pinyinCell = table.Cell(i, j).Shape.TextFrame.TextRange;
                            var hanziCell = table.Cell(i + 1, j).Shape.TextFrame.TextRange;

                            pinyinList.Add((pinyinCell.Text, pinyinCell.Font.Size, pinyinCell.Font.Name, pinyinCell.Font.Bold));
                            hanziList.Add((hanziCell.Text, hanziCell.Font.Size, hanziCell.Font.Name, hanziCell.Font.Bold));
                        }
                    }

                    // 删除指定数量的列
                    for (int k = 0; k < columnsToDelete; k++)
                    {
                        table.Columns[colCount - k].Delete();
                    }
                    colCount -= columnsToDelete;

                    // 计算新的行数
                    int newRowCount = (pinyinList.Count + colCount - 1) / colCount * 2;

                    // 确保表格有足够的行（两倍于新的行数以存储拼音和汉字）
                    while (table.Rows.Count < newRowCount)
                    {
                        table.Rows.Add();
                    }

                    // 清空所有单元格内容
                    for (int i = 1; i <= table.Rows.Count; i++)
                    {
                        for (int j = 1; j <= table.Columns.Count; j++)
                        {
                            table.Cell(i, j).Shape.TextFrame.TextRange.Text = string.Empty;
                        }
                    }

                    // 重新排列表格内容并应用样式
                    int index = 0;
                    for (int i = 1; i <= newRowCount; i += 2)
                    {
                        for (int j = 1; j <= colCount; j++)
                        {
                            if (index < pinyinList.Count)
                            {
                                var pinyinCell = table.Cell(i, j).Shape.TextFrame.TextRange;
                                var hanziCell = table.Cell(i + 1, j).Shape.TextFrame.TextRange;
                                pinyinCell.Text = pinyinList[index].text;
                                table.Cell(i, j).Shape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorBottom;
                                hanziCell.Text = hanziList[index].text;

                                // 确保奇数行的字号大小为偶数行字号大小的一半
                                pinyinCell.Font.Size = hanziCell.Font.Size / 2;

                                index++;
                            }
                        }
                    }

                    // 设置所有奇数行的高度与第一行保持一致，偶数行的高度与第二行保持一致
                    float firstRowHeight = table.Rows[1].Height;
                    float secondRowHeight = table.Rows[2].Height;
                    for (int i = 3; i <= table.Rows.Count; i += 2)
                    {
                        table.Rows[i].Height = firstRowHeight;
                        if (i + 1 <= table.Rows.Count)
                        {
                            table.Rows[i + 1].Height = secondRowHeight;
                        }
                    }

                    // 删除多余的空行，保留非空行上方的空行
                    for (int i = table.Rows.Count; i > 1; i--)
                    {
                        bool isEmpty = true;
                        for (int j = 1; j <= table.Columns.Count; j++)
                        {
                            if (!string.IsNullOrEmpty(table.Cell(i, j).Shape.TextFrame.TextRange.Text))
                            {
                                isEmpty = false;
                                break;
                            }
                        }

                        bool isNextRowEmpty = true;
                        if (i < table.Rows.Count)
                        {
                            for (int j = 1; j <= table.Columns.Count; j++)
                            {
                                if (!string.IsNullOrEmpty(table.Cell(i + 1, j).Shape.TextFrame.TextRange.Text))
                                {
                                    isNextRowEmpty = false;
                                    break;
                                }
                            }
                        }

                        if (isEmpty && isNextRowEmpty)
                        {
                            table.Rows[i].Delete();
                        }
                    }
                }
                else
                {
                    MessageBox.Show("表格必须至少包含更多列。", "重排表格", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show("请选择一个包含表格的形状。", "重排表格", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }


        private void 合并段落_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前激活的PPT应用程序实例
            Application application = Globals.ThisAddIn.Application;

            // 获取当前选中的对象
            Selection selection = application.ActiveWindow.Selection;

            // 确保选中的是多个表格
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 1)
            {
                List<Table> tables = new List<Table>();
                Table mainTable = null;

                foreach (Shape shape in selection.ShapeRange)
                {
                    if (shape.HasTable == MsoTriState.msoTrue)
                    {
                        if (mainTable == null)
                        {
                            mainTable = shape.Table;
                        }
                        else
                        {
                            tables.Add(shape.Table);
                        }
                    }
                }

                if (mainTable != null && tables.Count > 0)
                {
                    foreach (var table in tables)
                    {
                        int rowCount = table.Rows.Count;

                        // 在主表格末尾插入新行
                        for (int i = 1; i <= rowCount; i++)
                        {
                            mainTable.Rows.Add();
                        }

                        // 计算主表格新行的起始行索引
                        int startRow = mainTable.Rows.Count - rowCount + 1;
                        int colCount = Math.Min(mainTable.Columns.Count, table.Columns.Count);

                        // 将内容从次表格剪切到主表格
                        for (int i = 1; i <= rowCount; i++)
                        {
                            for (int j = 1; j <= colCount; j++)
                            {
                                var sourceCell = table.Cell(i, j).Shape.TextFrame.TextRange;
                                var targetCell = mainTable.Cell(startRow + i - 1, j).Shape.TextFrame.TextRange;

                                targetCell.Text = sourceCell.Text;
                            }
                        }

                        // 删除次表格
                        table.Parent.Delete();
                    }

                    // 应用奇数行字号大小为偶数行字号大小的50%
                    for (int i = 1; i <= mainTable.Rows.Count; i++)
                    {
                        for (int j = 1; j <= mainTable.Columns.Count; j++)
                        {
                            var cell = mainTable.Cell(i, j).Shape.TextFrame.TextRange;
                            if (i % 2 == 1)
                            {
                                cell.Font.Size = mainTable.Cell(i + 1, j).Shape.TextFrame.TextRange.Font.Size * 0.5f;
                                mainTable.Cell(i, j).Shape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorBottom;
                            }
                            cell.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("请选择多个表格进行合并。", "合并段落", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show("请选择多个表格进行合并。", "合并段落", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void 重设表格_Click(object sender, RibbonControlEventArgs e)
        {
            Application application = Globals.ThisAddIn.Application;
            Selection selection = application.ActiveWindow.Selection;

            // 检查是否选中了表格
            if (selection.Type == PpSelectionType.ppSelectionShapes &&
                selection.ShapeRange.Count == 1 &&
                selection.ShapeRange[1].HasTable == Office.MsoTriState.msoTrue)
            {
                Table table = selection.ShapeRange[1].Table;

                for (int i = 1; i <= table.Rows.Count; i++)
                {
                    for (int j = 1; j <= table.Columns.Count; j++)
                    {
                        Cell cell = table.Cell(i, j);
                        PowerPoint.TextFrame2 textFrame = cell.Shape.TextFrame2;

                        // 设置字体颜色为黑色
                        textFrame.TextRange.Font.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);

                        // 偶数行设置为顶端对齐
                        if (i % 2 == 0)
                        {
                            textFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorTop;
                        }
                        else // 奇数行设置字号为偶数行的50%
                        {
                            float evenRowFontSize = 12; // 默认值

                            // 获取偶数行字号
                            if (i < table.Rows.Count)
                            {
                                PowerPoint.TextFrame2 evenTextFrame = table.Cell(i + 1, j).Shape.TextFrame2;
                                if (evenTextFrame.HasText == MsoTriState.msoTrue)
                                {
                                    evenRowFontSize = evenTextFrame.TextRange.Font.Size;
                                }
                            }

                            // 设置奇数行字号为偶数行的50%
                            textFrame.TextRange.Font.Size = evenRowFontSize * 0.5f;
                        }
                    }
                }

                // 调整行高和列宽
                AdjustTableSize(table);
            }
            else
            {
                MessageBox.Show("请选中一个表格", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AdjustTableSize(Table table)
        {
            float maxWidth = 0;

            for (int i = 1; i <= table.Rows.Count; i++)
            {
                float maxHeight = 0;
                for (int j = 1; j <= table.Columns.Count; j++)
                {
                    Cell cell = table.Cell(i, j);
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

        private async void 文转表格_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Application pptApp = Globals.ThisAddIn.Application;
                Presentation presentation = pptApp.ActivePresentation;
                Slide slide = presentation.Slides[presentation.Slides.Count];

                slide.Shapes.PasteSpecial(PpPasteDataType.ppPasteHTML);
                DeleteEmptyRectangles(slide);
                var tableContents = ExtractHtmlTablesContent(slide);
                DeleteHtmlTables(slide);

                using (ProgressForm progressForm = new ProgressForm())
                {
                    progressForm.Show();

                    foreach (var tableContent in tableContents)
                    {
                        var formattedContent = await Task.Run(() => FormatTableContent(pptApp, tableContent, progressForm));
                        InsertFormattedTable(slide, formattedContent, pptApp);
                    }

                    progressForm.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("粘贴网页表格时发生错误: " + ex.Message);
            }
        }

        private void DeleteEmptyRectangles(Slide slide)
        {
            for (int i = slide.Shapes.Count; i >= 1; i--)
            {
                Shape shape = slide.Shapes[i];
                if (shape.Name.StartsWith("Rectangle"))
                {
                    shape.Delete();
                }
            }
        }

        private List<string[,]> ExtractHtmlTablesContent(Slide slide)
        {
            List<string[,]> tableContents = new List<string[,]>();

            for (int i = slide.Shapes.Count; i >= 1; i--)
            {
                Shape shape = slide.Shapes[i];
                if (shape.Type == Office.MsoShapeType.msoTable)
                {
                    Table htmlTable = shape.Table;
                    int rows = htmlTable.Rows.Count;
                    int cols = htmlTable.Columns.Count;
                    string[,] content = new string[rows, cols];

                    for (int r = 1; r <= rows; r++)
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            content[r - 1, c - 1] = htmlTable.Cell(r, c).Shape.TextFrame.TextRange.Text;
                        }
                    }

                    tableContents.Add(content);
                }
            }

            return tableContents;
        }

        private void DeleteHtmlTables(Slide slide)
        {
            for (int i = slide.Shapes.Count; i >= 1; i--)
            {
                Shape shape = slide.Shapes[i];
                if (shape.Type == Office.MsoShapeType.msoTable)
                {
                    shape.Delete();
                }
            }
        }

        private string[,] FormatTableContent(Application pptApp, string[,] content, ProgressForm progressForm)
        {
            int rows = content.GetLength(0);
            int cols = content.GetLength(1);
            int maxCols = 20;
            int totalRows = (cols / maxCols + 1) * rows;
            string[,] formattedContent = new string[totalRows, maxCols];

            int totalCells = rows * cols;
            int processedCells = 0;

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    int targetRow = (j / maxCols) * rows + i;
                    int targetCol = j % maxCols;
                    formattedContent[targetRow, targetCol] = content[i, j];

                    processedCells++;
                    int progress = (int)((double)processedCells / totalCells * 100);
                    progressForm.Invoke(new Action(() =>
                    {
                        progressForm.ProgressBar.Value = progress;
                        progressForm.ProgressBar.Refresh();
                        progressForm.ProgressLabel.Text = $"表格写入... {progress}%";
                    }));

                    System.Threading.Thread.Sleep(30);
                }
            }

            return formattedContent;
        }

        private void InsertFormattedTable(Slide slide, string[,] content, Application pptApp)
        {
            int rows = content.GetLength(0);
            int cols = content.GetLength(1);
            Shape pptTableShape = slide.Shapes.AddTable(rows, cols);
            Table pptTable = pptTableShape.Table;

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    pptTable.Cell(i + 1, j + 1).Shape.TextFrame.TextRange.Text = content[i, j];
                    FormatDoubleQuotes(pptTable.Cell(i + 1, j + 1));
                }
            }

            ResetTableFormat(pptApp, pptTable);
        }

        private void ResetTableFormat(Application pptApp, Table table)
        {
            int rowCount = table.Rows.Count;
            int columnCount = table.Columns.Count;
            float pinyinFontSize = 20;
            float hanziFontSize = 10;

            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= columnCount; col++)
                {
                    var cell = table.Cell(row, col);
                    PowerPoint.TextFrame textFrame = cell.Shape.TextFrame;

                    textFrame.MarginBottom = 0;
                    textFrame.MarginTop = row % 2 == 0 ? 0.5f : 0;
                    textFrame.MarginLeft = 0;
                    textFrame.MarginRight = 0;

                    textFrame.TextRange.ParagraphFormat.Alignment = (PpParagraphAlignment)Office.MsoParagraphAlignment.msoAlignCenter;
                    textFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(System.Drawing.Color.Black);
                    textFrame.TextRange.Font.Bold = Office.MsoTriState.msoFalse;
                    cell.Shape.Fill.Transparency = 1;

                    cell.Borders[PowerPoint.PpBorderType.ppBorderTop].Weight = 0;
                    cell.Borders[PowerPoint.PpBorderType.ppBorderBottom].Weight = 0;
                    cell.Borders[PowerPoint.PpBorderType.ppBorderLeft].Weight = 0;
                    cell.Borders[PowerPoint.PpBorderType.ppBorderRight].Weight = 0;

                    textFrame.VerticalAnchor = row % 2 == 0 ? Office.MsoVerticalAnchor.msoAnchorTop : Office.MsoVerticalAnchor.msoAnchorBottom;

                    textFrame.TextRange.Font.Size = row % 2 == 0 ? pinyinFontSize : hanziFontSize;
                }
            }

            AdjustTableDimensions(table);
        }

        private void AdjustTableDimensions(Table table)
        {
            float[] colMaxWidths = new float[table.Columns.Count];

            for (int j = 1; j <= table.Columns.Count; j++)
            {
                colMaxWidths[j - 1] = 0;
                for (int i = 1; i <= table.Rows.Count; i++)
                {
                    Cell cell = table.Cell(i, j);
                    float width = cell.Shape.TextFrame.TextRange.BoundWidth;
                    if (width > colMaxWidths[j - 1])
                    {
                        colMaxWidths[j - 1] = width;
                    }
                }
            }

            for (int j = 1; j <= table.Columns.Count; j++)
            {
                table.Columns[j].Width = colMaxWidths[j - 1] + 2;
            }

            for (int i = 1; i <= table.Rows.Count; i++)
            {
                float maxHeight = 0;
                for (int j = 1; j <= table.Columns.Count; j++)
                {
                    Cell cell = table.Cell(i, j);
                    float height = cell.Shape.TextFrame.TextRange.BoundHeight;
                    if (height > maxHeight)
                    {
                        maxHeight = height;
                    }
                }
                table.Rows[i].Height = maxHeight + 2;
            }
        }

        private void FormatDoubleQuotes(Cell cell)
        {
            PowerPoint.TextRange textRange = cell.Shape.TextFrame.TextRange;
            int startPos = 0;
            while ((startPos = textRange.Text.IndexOf('“', startPos)) != -1)
            {
                if (startPos + 1 < textRange.Text.Length)
                {
                    textRange.Characters(startPos + 1, 1).Font.Superscript = Office.MsoTriState.msoTrue;
                }
                startPos++;
            }
            startPos = 0;
            while ((startPos = textRange.Text.IndexOf('”', startPos)) != -1)
            {
                if (startPos + 1 < textRange.Text.Length)
                {
                    textRange.Characters(startPos + 1, 1).Font.Superscript = Office.MsoTriState.msoTrue;
                }
                startPos++;
            }
        }

        private void 在线注音编辑器_Click(object sender, RibbonControlEventArgs e)
        {
            OpenWebPage("https://toneoz.com/ime/?fnt=1");
        }


        private void 左右镜像_Click(object sender, RibbonControlEventArgs e)
        {
            Application app = Globals.ThisAddIn.Application;
            _ = app.ActiveWindow.View.Slide;
            Selection selection = app.ActiveWindow.Selection;

            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange shapeRange = selection.ShapeRange;

                foreach (Shape shape in shapeRange)
                {
                    float originalLeft = shape.Left;
                    float originalTop = shape.Top;
                    float originalWidth = shape.Width;

                    // 创建镜像副本
                   Shape mirroredShape = shape.Duplicate()[1];

                    // 水平翻转副本
                    mirroredShape.Flip(MsoFlipCmd.msoFlipHorizontal);

                    // 计算镜像位置
                    float mirroredLeft = originalLeft + originalWidth;
                    mirroredShape.Left = mirroredLeft;
                    mirroredShape.Top = originalTop;
                }
            }
            else
            {
                MessageBox.Show("请选择一个或多个形状进行镜像操作。", "错误", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void 上下镜像_Click(object sender, RibbonControlEventArgs e)
        {
            Application app = Globals.ThisAddIn.Application;
            _ = app.ActiveWindow.View.Slide;
            Selection selection = app.ActiveWindow.Selection;

            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange shapeRange = selection.ShapeRange;

                foreach (Shape shape in shapeRange)
                {
                    float originalLeft = shape.Left;
                    float originalTop = shape.Top;
                    float originalHeight = shape.Height;

                    // 创建镜像副本
                    Shape mirroredShape = shape.Duplicate()[1];

                    // 垂直翻转副本
                    mirroredShape.Flip(MsoFlipCmd.msoFlipVertical);

                    // 计算镜像位置
                    float mirroredTop = originalTop + originalHeight;
                    mirroredShape.Left = originalLeft;
                    mirroredShape.Top = mirroredTop;
                }
            }
            else
            {
                MessageBox.Show("请选择一个或多个形状进行镜像操作。", "错误", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void 分解拼音_Click(object sender, RibbonControlEventArgs e)
        {
            Application app = Globals.ThisAddIn.Application;
            Selection selection = app.ActiveWindow.Selection;
            string pinyinFilePath = ExtractEmbeddedResource("课件帮PPT助手.汉字字典.拆分拼音.txt");

            if (selection.Type == PpSelectionType.ppSelectionText || selection.Type == PpSelectionType.ppSelectionShapes)
            {
                List<TextRange> textRanges = new List<TextRange>();

                if (selection.Type == PpSelectionType.ppSelectionText)
                {
                    textRanges.Add(selection.TextRange);
                }
                else if (selection.Type == PpSelectionType.ppSelectionShapes)
                {
                    foreach (Shape shape in selection.ShapeRange)
                    {
                        if (shape.HasTextFrame == MsoTriState.msoTrue)
                        {
                            shape.TextFrame.WordWrap = MsoTriState.msoFalse; // 取消自动换行
                            textRanges.Add(shape.TextFrame.TextRange);
                        }
                    }
                }

                foreach (PowerPoint.TextRange textRange in textRanges)
                {
                    string originalText = RemoveAllWhitespace(textRange.Text); // 移除所有空白字符
                    string pinyinWithoutTone = RemoveTone(originalText);
                    Dictionary<string, string> pinyinMap = LoadPinyinMap(pinyinFilePath);

                    if (pinyinMap.TryGetValue(pinyinWithoutTone, out string splitPinyin))
                    {
                        string splitPinyinWithTone = AssignTone(originalText, splitPinyin);
                        string formattedText = $"{splitPinyinWithTone.Replace("+", "–")}→{originalText}";
                        textRange.Text = RemoveAllWhitespace(formattedText); // 移除结果中的所有空白字符
                    }
                    else
                    {
                        MessageBox.Show("请确保这是单个拼音（检查是否存在大写、拼音是否正确等）。", "错误", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    }
                }
            }

            else
                {
                    MessageBox.Show("请选择至少一个带拼音的文本框进行拼音分解操作。", "错误", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private string RemoveAllWhitespace(string input)
        {
            return new string(input.Where(c => !char.IsWhiteSpace(c)).ToArray());
        }

        private string PinyinResource(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            using (StreamReader reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }

        private string RemoveTone(string pinyin)
        {
            Dictionary<char, char> toneMap = new Dictionary<char, char>
            {
                {'ā', 'a'}, {'á', 'a'}, {'ǎ', 'a'}, {'à', 'a'},
                {'ē', 'e'}, {'é', 'e'}, {'ě', 'e'}, {'è', 'e'},
                {'ī', 'i'}, {'í', 'i'}, {'ǐ', 'i'}, {'ì', 'i'},
                {'ō', 'o'}, {'ó', 'o'}, {'ǒ', 'o'}, {'ò', 'o'},
                {'ū', 'u'}, {'ú', 'u'}, {'ǔ', 'u'}, {'ù', 'u'},
                {'ǖ', 'ü'}, {'ǘ', 'ü'}, {'ǚ', 'ü'}, {'ǜ', 'ü'}
            };

            char[] result = pinyin.Select(c => toneMap.ContainsKey(c) ? toneMap[c] : c).ToArray();
            return new string(result);
        }

        private string AssignTone(string original, string splitPinyin)
        {
            Dictionary<char, char> toneMap = new Dictionary<char, char>
            {
                {'ā', 'a'}, {'á', 'a'}, {'ǎ', 'a'}, {'à', 'a'},
                {'ē', 'e'}, {'é', 'e'}, {'ě', 'e'}, {'è', 'e'},
                {'ī', 'i'}, {'í', 'i'}, {'ǐ', 'i'}, {'ì', 'i'},
                {'ō', 'o'}, {'ó', 'o'}, {'ǒ', 'o'}, {'ò', 'o'},
                {'ū', 'u'}, {'ú', 'u'}, {'ǔ', 'u'}, {'ù', 'u'},
                {'ǖ', 'ü'}, {'ǘ', 'ü'}, {'ǚ', 'ü'}, {'ǜ', 'ü'}
            };

            char[] result = splitPinyin.ToCharArray();
            foreach (char c in original)
            {
                if (toneMap.ContainsKey(c))
                {
                    char toneChar = toneMap[c];
                    for (int i = 0; i < result.Length; i++)
                    {
                        if (result[i] == toneChar)
                        {
                            result[i] = c;
                            break;
                        }
                    }
                    break;
                }
            }

            return new string(result);
        }

        private Dictionary<string, string> LoadPinyinMap(string filePath)
        {
            var map = new Dictionary<string, string>();
            string[] lines = File.ReadAllLines(filePath);

            foreach (string line in lines)
            {
                string[] parts = line.Split('=');
                if (parts.Length == 2)
                {
                    map[parts[0]] = parts[1];
                }
            }

            return map;
        }

        private void 环形分布_Click(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            var selection = application.ActiveWindow.Selection;

            // 检查是否至少选中了一个对象
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 0)
            {
                RingDistribution distribution = new RingDistribution();
                distribution.Show();
            }
            else
            {
                MessageBox.Show("请至少选中一个对象！");
            }
        }

        private void 矩阵分布_Click(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            var selection = application.ActiveWindow.Selection;

            // 检查是否至少选中了一个对象
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 0)
            {
                if (selection.ShapeRange.Count == 1)
                {
                    // 选择了一个对象，弹出“矩阵复制”窗体
                    MatrixCopy copyWindow = new MatrixCopy();
                    copyWindow.Show();
                }
                else
                {
                    // 选择了多个对象，弹出“矩阵分布”窗体
                    MatrixDistribution distributionWindow = new MatrixDistribution();
                    distributionWindow.Show();
                }
            }
            else
            {
                MessageBox.Show("请至少选中一个对象！");
            }
        }

        private LayoutForm settingsForm;
        private string currentTag;

        private void 辐射连线_Click(object sender, RibbonControlEventArgs e)
        {
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange shapes = selection.ShapeRange;

                if (shapes.Count > 1)
                {
                    settingsForm = new LayoutForm();
                    currentTag = Guid.NewGuid().ToString();

                    settingsForm.LayoutUpdated += (s, args) =>
                    {
                        ApplyLayout(shapes, (float)settingsForm.Distance, settingsForm.Compactness, settingsForm.StartAngle);
                    };

                    // 初始化布局
                    ApplyLayout(shapes, (float)settingsForm.Distance, settingsForm.Compactness, settingsForm.StartAngle);
                    settingsForm.Show();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("请至少选择两个对象。");
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("请选中要连接的对象。");
            }
        }

        private void ApplyLayout(PowerPoint.ShapeRange shapes, float radius, int compactness, int startAngle)
        {
            if (shapes == null || shapes.Parent?.Shapes == null)
            {
                return;
            }

            // 获取第一个选中的形状
            PowerPoint.Shape centerShape = shapes[1];
            float centerX = centerShape.Left + (centerShape.Width / 2);
            float centerY = centerShape.Top + (centerShape.Height / 2);

            // 确定线段颜色逻辑
            int lineColor;
            if (centerShape.Line.Visible == Office.MsoTriState.msoTrue)
            {
                lineColor = centerShape.Line.ForeColor.RGB;
            }
            else
            {
                lineColor = centerShape.Fill.ForeColor.RGB;
            }

            double angleIncrement = (compactness > 0 ? (compactness / 100.0 * 360.0) / (shapes.Count - 1) : 360.0 / (shapes.Count - 1));
            angleIncrement = settingsForm.IsClockwise ? angleIncrement : -angleIncrement;

            // 删除旧的线段
            foreach (var line in settingsForm.LayoutLines)
            {
                line.Delete();
            }
            settingsForm.LayoutLines.Clear();

            for (int i = 2; i <= shapes.Count; i++)
            {
                PowerPoint.Shape targetShape = shapes[i];

                double angle = startAngle + (i - 2) * angleIncrement;

                // 计算目标对象的新位置
                float targetX = centerX + (float)(radius * Math.Cos(angle * Math.PI / 180)) - (targetShape.Width / 2);
                float targetY = centerY + (float)(radius * Math.Sin(angle * Math.PI / 180)) - (targetShape.Height / 2);

                // 更新目标对象位置
                targetShape.Left = targetX;
                targetShape.Top = targetY;

                // 计算目标对象的中心点
                float targetCenterX = targetShape.Left + (targetShape.Width / 2);
                float targetCenterY = targetShape.Top + (targetShape.Height / 2);

                // 创建并更新线段
                PowerPoint.Shape line = shapes.Parent.Shapes.AddLine(centerX, centerY, targetCenterX, targetCenterY);
                line.Line.ForeColor.RGB = lineColor; // 使用确定的颜色
                line.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
                line.Line.Weight = 2;
                line.Tags.Add("LayoutTag", currentTag);
                settingsForm.LayoutLines.Add(line);
            }
        }
        // 查找与目标对象关联的线段
        private PowerPoint.Shape FindExistingLine(PowerPoint.Shape targetShape)
        {
            foreach (PowerPoint.Shape line in settingsForm.LayoutLines)
            {
                if (line.Tags["TargetShapeName"] == targetShape.Name)
                {
                    return line;
                }
            }
            return null;
        }

        // 更新现有线段的位置
        private void UpdateLinePosition(PowerPoint.Shape line, float startX, float startY, float endX, float endY)
        {
            // 更新线段的位置
            line.Left = Math.Min(startX, endX);
            line.Top = Math.Min(startY, endY);

            // 更新线段的尺寸
            line.Width = Math.Abs(endX - startX);
            line.Height = Math.Abs(endY - startY);

            // 更新线段的方向
            if (endX >= startX && endY >= startY)
            {
                line.Line.BeginArrowheadLength = Office.MsoArrowheadLength.msoArrowheadShort;
                line.Line.EndArrowheadLength = Office.MsoArrowheadLength.msoArrowheadShort;
            }
            else if (endX < startX && endY >= startY)
            {
                line.Line.BeginArrowheadLength = Office.MsoArrowheadLength.msoArrowheadShort;
                line.Line.EndArrowheadLength = Office.MsoArrowheadLength.msoArrowheadShort;
            }
            else if (endX >= startX && endY < startY)
            {
                line.Line.BeginArrowheadLength = Office.MsoArrowheadLength.msoArrowheadShort;
                line.Line.EndArrowheadLength = Office.MsoArrowheadLength.msoArrowheadShort;
            }
            else
            {
                line.Line.BeginArrowheadLength = Office.MsoArrowheadLength.msoArrowheadShort;
                line.Line.EndArrowheadLength = Office.MsoArrowheadLength.msoArrowheadShort;
            }
        }

        private const float MaxDiameter = 200f; // 设置圆形的最大直径
        private void 文形互转_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Slide activeSlide = app.ActiveWindow.View.Slide as PowerPoint.Slide;

            if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange selectedShapes = app.ActiveWindow.Selection.ShapeRange;

                // 调用消除边距的点击事件，移除所有选中文本框的边距
                去除边距_Click(sender, e);

                foreach (PowerPoint.Shape shape in selectedShapes)
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        PowerPoint.TextFrame textFrame = shape.TextFrame;
                        PowerPoint.TextRange textRange = textFrame.TextRange;

                        // 获取原文本框的相关属性
                        float originalFontSize = textRange.Font.Size;
                        string originalText = textRange.Text;
                        string originalFontName = textRange.Font.Name;

                        bool ctrlPressed = (Control.ModifierKeys & Keys.Control) == Keys.Control;

                        float left = shape.Left;
                        float top = shape.Top;

                        PowerPoint.Shape newShape;

                        if (ctrlPressed)
                        {
                            // 以文本框宽高的最大值作为初始直径
                            float diameter = Math.Max(shape.Width+3, shape.Height+3);
                            newShape = activeSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, left, top, diameter, diameter);
                            newShape.TextFrame.TextRange.Text = originalText;
                            newShape.TextFrame.TextRange.Font.Size = originalFontSize;
                            newShape.TextFrame.TextRange.Font.Name = originalFontName;
                            newShape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                            newShape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                            newShape.TextFrame.WordWrap = Office.MsoTriState.msoFalse; // 取消自动换行

                            // 调整圆形大小以适应文本
                            AdjustShapeToFitText(newShape, originalText, originalFontSize);
                        }
                        else
                        {
                            // 没有按下 Ctrl 时，创建圆角矩形
                            float width = shape.Width;
                            float height = shape.Height;
                            newShape = activeSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, left, top, width, height);
                            newShape.TextFrame.TextRange.Text = originalText;
                            newShape.TextFrame.TextRange.Font.Size = originalFontSize;
                            newShape.TextFrame.TextRange.Font.Name = originalFontName;
                            newShape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                            newShape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                            newShape.TextFrame.WordWrap = Office.MsoTriState.msoFalse; // 取消自动换行
                        }

                        // 将新形状放置到与原文本框相同的位置
                        newShape.Left = left;
                        newShape.Top = top;

                        // 删除原文本框
                        shape.Delete();
                    }
                }
            }
            else
            {
                MessageBox.Show("请先选择一个或多个文本框", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void AdjustShapeToFitText(PowerPoint.Shape shape, string text, float fontSize)
        {
            PowerPoint.TextRange textRange = shape.TextFrame.TextRange;
            float currentDiameter = shape.Width;

            while (true)
            {
                if ((textRange.BoundWidth > currentDiameter || textRange.BoundHeight > currentDiameter) && currentDiameter < MaxDiameter)
                {
                    currentDiameter += 5;
                    shape.Width = currentDiameter;
                    shape.Height = currentDiameter;
                }
                else if (currentDiameter >= MaxDiameter)
                {
                    // 减小字体以适应最大直径内的文本
                    fontSize -= 1;
                    textRange.Font.Size = fontSize;
                    if (textRange.BoundWidth <= currentDiameter && textRange.BoundHeight <= currentDiameter)
                    {
                        break;
                    }
                }
                else
                {
                    break;
                }
            }
        }

        private void 层级关系_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = app.ActiveWindow.Selection;

            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange.Count < 2)
            {
                System.Windows.Forms.MessageBox.Show("请至少选中两个对象进行操作。", "提示", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                return;
            }

            List<PowerPoint.Shape> shapes = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in selection.ShapeRange)
            {
                shapes.Add(shape);
            }

            PowerPoint.Shape mainShape = shapes[0];
            List<PowerPoint.Shape> branchShapes = shapes.Skip(1).ToList();

            // 横向布局
            float totalHeight = branchShapes.Sum(s => s.Height) + (branchShapes.Count - 1) * 10; // 10 为间距
            float startY = mainShape.Top + (mainShape.Height - totalHeight) / 2;

            for (int i = 0; i < branchShapes.Count; i++)
            {
                PowerPoint.Shape branchShape = branchShapes[i];
                float newTop = startY + i * (branchShape.Height + 10); // 10 为间距
                branchShape.Top = newTop;
                branchShape.Left = mainShape.Left + mainShape.Width + 50; // 50 为水平间距

                // 判断是否使用直线连接符
                PowerPoint.Shape connector;
                if (branchShape.Top == mainShape.Top)
                {
                    connector = selection.SlideRange[1].Shapes.AddConnector(
                        Office.MsoConnectorType.msoConnectorStraight,
                        0, 0, 0, 0);
                }
                else
                {
                    connector = selection.SlideRange[1].Shapes.AddConnector(
                        Office.MsoConnectorType.msoConnectorElbow,
                        0, 0, 0, 0);
                }

                // 主干右侧连接分支左侧
                connector.ConnectorFormat.BeginConnect(mainShape, 3); // 主干的右侧连接点
                connector.ConnectorFormat.EndConnect(branchShape, 1); // 分支的左侧连接点
                connector.RerouteConnections(); // 自动调整连接器的路径
            }
        }

        private void 括弧关系_Click(object sender, RibbonControlEventArgs e)
        {
            Application app = Globals.ThisAddIn.Application;
            Selection selection = app.ActiveWindow.Selection;

            if (selection.Type != PpSelectionType.ppSelectionShapes || selection.ShapeRange.Count < 2)
            {
                MessageBox.Show("请至少选中两个对象进行操作。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            List<Shape> shapes = new List<Shape>();
            foreach (Shape shape in selection.ShapeRange)
            {
                shapes.Add(shape);
            }

            // 检查是否有文本框，并移除其边距
            bool hasTextBox = shapes.Any(shape => shape.HasTextFrame == MsoTriState.msoTrue);
            if (hasTextBox)
            {
                // 调用去除边距的点击事件，移除所有选中文本框的边距
                去除边距_Click(sender, e);
            }

            Shape mainShape = shapes[0]; // 获取第一个选中的形状或文本框
            List<Shape> branchShapes = shapes.Skip(1).ToList();

            // 计算分支的总高度和最大宽度
            float totalHeight = branchShapes.Sum(s => s.Height) + (branchShapes.Count - 1) * 20; // 20 为间距
            float maxWidth = branchShapes.Max(s => s.Width);

            // 计算主干、左大括号和分支整体的垂直居中起始位置
            float centerY = mainShape.Top + mainShape.Height / 2;
            float startY = centerY - totalHeight / 2;

            // 布局分支
            float currentY = startY;
            for (int i = 0; i < branchShapes.Count; i++)
            {
                PowerPoint.Shape branchShape = branchShapes[i];
                branchShape.Left = mainShape.Left + mainShape.Width + 50; // 50 为水平间距
                branchShape.Top = currentY;
                currentY += branchShape.Height + 20; // 20 为垂直间距
            }

            // 在分支左侧添加一个“左大括号”
            Shape leftBracket = selection.SlideRange[1].Shapes.AddShape(
                MsoAutoShapeType.msoShapeLeftBrace,
                branchShapes.Min(s => s.Left) - 50, // 大括号的水平位置
                startY - 10, // 大括号的垂直位置
                20, // 大括号的宽度
                totalHeight - branchShapes[0].Height + 10 // 大括号的高度调整，减去一个分支高度
            );

            // 设置大括号的粗细为1.5磅
            leftBracket.Line.Weight = 1.5f;
            // 设置大括号的弧度，使其更美观
            leftBracket.Adjustments[1] = 0.2f; // 调整大括号的弧度

            // 如果第一个选中的对象是文本框，则将大括号的颜色设置为黑色
            if (mainShape.HasTextFrame == MsoTriState.msoTrue && mainShape.TextFrame.HasText == MsoTriState.msoTrue)
            {
                leftBracket.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            }
            else
            {
                // 否则，使用形状的填充颜色设置大括号的颜色
                leftBracket.Line.ForeColor.RGB = mainShape.Fill.ForeColor.RGB;
            }

            // 将分支对象组合在一起
            Shape[] shapesToGroup = branchShapes.ToArray();
            Shape groupedBranches = selection.SlideRange[1].Shapes.Range(shapesToGroup.Select(s => s.Name).ToArray()).Group();

            // 调整分支和左大括号的垂直居中位置
            float newTop = centerY - (groupedBranches.Height / 2);
            groupedBranches.Top = newTop;
            // 向左移动 groupedBranches 20 个单位
            groupedBranches.Left -= 20;

            // 调整大括号的位置
            leftBracket.Top = groupedBranches.Top;
            leftBracket.Height = groupedBranches.Height - branchShapes[0].Height; // 调整大括号的高度，减去一个分支对象的高度

            // 计算主干的垂直中心位置，考虑分支的总高度
            float mainShapeTop = newTop + (groupedBranches.Height - mainShape.Height) / 2;

            // 设置主干的位置
            mainShape.Top = mainShapeTop;

            // 重新调整大括号的位置，使其与分支整体垂直居中对齐
            leftBracket.Top = groupedBranches.Top;
            leftBracket.Top = groupedBranches.Top + (groupedBranches.Height - leftBracket.Height) / 2;
            groupedBranches.Ungroup();
        }

        private const string TagIdentifier = "MultiToneTag"; // 定义一个唯一的标记标识符
        private void 多音字词填空_Click(object sender, RibbonControlEventArgs e)
        {
            // 设置EPPlus的许可上下文
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // 获取当前选中的文本框
            Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                foreach (Shape shape in selection.ShapeRange)
                {
                    if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoTrue)
                    {
                        string selectedCharacter = shape.TextFrame.TextRange.Text;

                        if (selectedCharacter.Length == 1)
                        {
                            // 获取用户所选原始文本框的字号大小
                            float originalFontSize = shape.TextFrame.TextRange.Font.Size;
                            float newFontSize = (float)(originalFontSize * 0.6);

                            // 读取嵌入的Excel文件
                            string filePath = ExtractEmbeddedResourceFilePath("课件帮PPT助手.汉字字典.汉字拼音信息库.xlsx");
                            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                            {
                                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // 假设数据在第一个工作表

                                // 查找汉字
                                bool found = false;
                                for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                                {
                                    string character = worksheet.Cells[row, 1].Text;
                                    if (character == selectedCharacter)
                                    {
                                        found = true;
                                        string pinyin = worksheet.Cells[row, 2].Text;
                                        string[] pinyinArray = pinyin.Split(',');

                                        Slide slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

                                        // 给原始文本框添加标记
                                        shape.Tags.Add(TagIdentifier, "Original");

                                        // 创建拼音文本框并排列，并给每个拼音文本框添加相同的标记
                                        float leftPosition = shape.Left + shape.Width + 10;
                                        var pinyinTextBoxes = new Shape[pinyinArray.Length];
                                        for (int i = 0; i < pinyinArray.Length; i++)
                                        {
                                            Shape textBox = slide.Shapes.AddTextbox(
                                                MsoTextOrientation.msoTextOrientationHorizontal,
                                                leftPosition,
                                                shape.Top,
                                                50, 50);
                                            textBox.TextFrame.TextRange.Text = pinyinArray[i].Trim();
                                            textBox.TextFrame.TextRange.Font.Size = newFontSize; // 设置拼音文本框的字号
                                            textBox.TextFrame.WordWrap = MsoTriState.msoFalse; // 取消自动换行
                                            textBox.Tags.Add(TagIdentifier, "Pinyin"); // 添加标记
                                            pinyinTextBoxes[i] = textBox; // 保存拼音文本框引用
                                            leftPosition += textBox.Width + 10;
                                        }

                                        // 选择用户的原始文本框和刚刚创建的拼音文本框
                                        var shapesToSelect = new Shape[pinyinTextBoxes.Length + 1];
                                        shapesToSelect[0] = shape; // 原始文本框
                                        pinyinTextBoxes.CopyTo(shapesToSelect, 1);

                                        PowerPoint.ShapeRange shapeRange = slide.Shapes.Range(shapesToSelect.Select(s => s.Name).ToArray());
                                        shapeRange.Select(); // 仅选择带有标记的所有形状

                                        // 调用括弧关系的点击事件
                                        括弧关系_Click(sender, e);

                                        // 在括弧关系操作完成后，在每个拼音文本框右侧添加“（）”
                                        AddBracketsToPinyinShapes(pinyinTextBoxes, newFontSize);

                                        // 删除标记
                                        RemoveTagsFromShapes(slide, TagIdentifier);

                                        break;
                                    }
                                }

                                if (!found)
                                {
                                    MessageBox.Show($"未在字典中找到汉字: {selectedCharacter}");
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show($"文本框内容不止一个汉字: {selectedCharacter}");
                        }
                    }
                    else
                    {
                        MessageBox.Show("所选文本框中没有有效的汉字文本。");
                    }
                }
            }
            else
            {
                MessageBox.Show("请选择包含汉字的文本框。");
            }
        }

        // 方法：从嵌入的资源中提取文件路径
        private string ExtractEmbeddedResourceFilePath(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                string tempFile = Path.GetTempFileName();
                using (var fileStream = new FileStream(tempFile, FileMode.Create))
                {
                    stream.CopyTo(fileStream);
                }
                return tempFile;
            }
        }

        // 方法：在拼音文本框右侧添加“（）”
        private void AddBracketsToPinyinShapes(Shape[] pinyinTextBoxes, float fontSize)
        {
            foreach (var pinyinShape in pinyinTextBoxes)
            {
                Shape bracketBox = pinyinShape.Parent.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    pinyinShape.Left + pinyinShape.Width,
                    pinyinShape.Top,
                    50, 50);
                bracketBox.TextFrame.TextRange.Text = "（　　）"; // 两个全角空格
                bracketBox.TextFrame.TextRange.Font.Size = fontSize; // 设置括号文本框的字号
                bracketBox.TextFrame.WordWrap = MsoTriState.msoFalse; // 取消自动换行

                // 取消文本框边距
                bracketBox.TextFrame.MarginLeft = 0;
                bracketBox.TextFrame.MarginRight = 0;
                bracketBox.TextFrame.MarginTop = 0;
                bracketBox.TextFrame.MarginBottom = 0;
            }
        }

        // 方法：删除形状的标记
        private void RemoveTagsFromShapes(Slide slide, string tag)
        {
            foreach (Shape shape in slide.Shapes)
            {
                if (shape.Tags[tag] != null)
                {
                    shape.Tags.Delete(tag); // 删除标记
                }
            }
        }

        private void 字号调整_Click(object sender, RibbonControlEventArgs e)
        {
            var pptApp = Globals.ThisAddIn.Application;
            var activeWindow = pptApp.ActiveWindow;
            var selection = activeWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                var shape = selection.ShapeRange[1];
                if (shape.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    var table = shape.Table;

                    // 获取当前拼音行和汉字行的字号
                    double currentPinyinFontSize = 10;  // 默认值
                    double currentHanziFontSize = 20;  // 默认值

                    if (table.Rows.Count > 1 && table.Columns.Count > 0)
                    {
                        // 获取拼音行和汉字行的字号
                        currentPinyinFontSize = table.Cell(1, 1).Shape.TextFrame.TextRange.Font.Size;
                        currentHanziFontSize = table.Cell(2, 1).Shape.TextFrame.TextRange.Font.Size;
                    }

                    // 显示字号调整窗体
                    using (FontSizeAdjustForm fontSizeForm = new FontSizeAdjustForm((float)currentPinyinFontSize, (float)currentHanziFontSize))
                    {
                        if (fontSizeForm.ShowDialog() == DialogResult.OK)
                        {
                            float newPinyinFontSize = fontSizeForm.PinyinFontSize;
                            float newHanziFontSize = fontSizeForm.HanziFontSize;

                            // 更新表格中的字号
                            for (int row = 1; row <= table.Rows.Count; row++)
                            {
                                for (int col = 1; col <= table.Columns.Count; col++)
                                {
                                    var cell = table.Cell(row, col);
                                    if (row % 2 != 0) // 奇数行为拼音行
                                    {
                                        cell.Shape.TextFrame.TextRange.Font.Size = newPinyinFontSize;
                                    }
                                    else // 偶数行为汉字行
                                    {
                                        cell.Shape.TextFrame.TextRange.Font.Size = newHanziFontSize;
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("请选择一个注音文本布局表格进行字号调整。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                MessageBox.Show("请选择一个注音文本布局表格进行字号调整。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void 奇数行行高设置_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var application = Globals.ThisAddIn.Application;
                var selection = application.ActiveWindow.Selection;

                // 检查用户是否选中了一个表格
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count == 1)
                {
                    var shape = selection.ShapeRange[1];
                    if (shape.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        var table = shape.Table;

                        // 显示设置奇数行行高的窗体
                        OddRowHeightForm form = new OddRowHeightForm(table);
                        if (form.ShowDialog() == DialogResult.OK)
                        {
                            // 已在NumericUpDown的事件处理程序中应用行高
                        }
                    }
                    else
                    {
                        MessageBox.Show("请选中一个包含表格的对象。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("请选中一个包含表格的对象。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("设置奇数行行高时发生错误: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void 字词加点_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = app.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
            {
                PowerPoint.TextRange textRange = selection.TextRange;
                PowerPoint.Shape textBox = selection.ShapeRange[1];
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

                // 定义比例系数，用于调整圆点大小和位置
                const float dotSizeRatio = 0.2f;  // 圆点大小占字体大小的比例
                const float verticalAdjustmentRatio = 0.1f;  // 垂直调整比例，基于字体大小
                const float upwardShiftRatio = 0.5f;  // 向上调整比例，基于额外行间距

                // 遍历每个字符并在其下方添加圆点
                for (int i = 1; i <= textRange.Text.Count(); i++)
                {
                    var charRange = textRange.Characters(i, 1);
                    float fontSize = charRange.Font.Size;
                    float boundHeight = charRange.BoundHeight;
                    float dotSize = fontSize * dotSizeRatio;

                    // 行高的估算，如果行高明显大于字体大小，则认为行间距＞1
                    float extraLineSpacing = Math.Max(0, boundHeight - fontSize);

                    // 计算圆点的位置，基于行间距进行向上调整
                    float left = charRange.BoundLeft + (charRange.BoundWidth / 2) - (dotSize / 2); // 圆点居中于字符
                    float top = charRange.BoundTop + charRange.BoundHeight
                                + (fontSize * verticalAdjustmentRatio)
                                - (extraLineSpacing * upwardShiftRatio); // 调整圆点的垂直位置

                    // 添加圆点
                    var dot = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, left, top, dotSize, dotSize);
                    dot.Fill.ForeColor.RGB = charRange.Font.Color.RGB; // 设置圆点颜色
                    dot.Line.Visible = Office.MsoTriState.msoFalse; // 隐藏圆点的边框
                }
            }
            else
            {
                MessageBox.Show("请在文本框内选择需要加点的字词。");
            }
        }

        private void 拼音升调_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = app.ActiveWindow.Selection;

            // 定义拼音声调映射字典
            Dictionary<char, char> toneMap = new Dictionary<char, char>()
    {
        {'ā', 'á'}, {'á', 'ǎ'}, {'ǎ', 'à'}, {'à', 'ā'}, // 'a' 声调转换
        {'ē', 'é'}, {'é', 'ě'}, {'ě', 'è'}, {'è', 'ē'}, // 'e' 声调转换
        {'ī', 'í'}, {'í', 'ǐ'}, {'ǐ', 'ì'}, {'ì', 'ī'}, // 'i' 声调转换
        {'ō', 'ó'}, {'ó', 'ǒ'}, {'ǒ', 'ò'}, {'ò', 'ō'}, // 'o' 声调转换
        {'ū', 'ú'}, {'ú', 'ǔ'}, {'ǔ', 'ù'}, {'ù', 'ū'}, // 'u' 声调转换
        {'ǖ', 'ǘ'}, {'ǘ', 'ǚ'}, {'ǚ', 'ǜ'}, {'ǜ', 'ǖ'}, // 'ü' 声调转换
    };

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
            {
                // 如果选择的是文本，处理选中的文本
                PowerPoint.TextRange textRange = selection.TextRange;
                ApplyToneChange(textRange, toneMap);
            }
            else if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                // 如果选择的是形状（文本框），处理每个文本框中的文本
                PowerPoint.ShapeRange shapeRange = selection.ShapeRange;
                foreach (PowerPoint.Shape shape in shapeRange)
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        PowerPoint.TextRange textRange = shape.TextFrame.TextRange;
                        ApplyToneChange(textRange, toneMap);
                    }
                }
            }
            else
            {
                MessageBox.Show("请先选择带有拼音的文本框或文本。");
            }
        }

        private void ApplyToneChange(PowerPoint.TextRange textRange, Dictionary<char, char> toneMap)
        {
            string text = textRange.Text;
            char[] newTextArray = text.ToCharArray();

            // 遍历文本并调整声调
            for (int i = 0; i < newTextArray.Length; i++)
            {
                if (toneMap.ContainsKey(newTextArray[i]))
                {
                    newTextArray[i] = toneMap[newTextArray[i]];
                }
            }

            // 将调整后的文本设置回文本框
            textRange.Text = new string(newTextArray);
        }

        private void 增加行宽_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前激活的PPT应用程序实例
            Application application = Globals.ThisAddIn.Application;

            // 获取当前选中的对象
            Selection selection = application.ActiveWindow.Selection;

            // 检查是否按下Ctrl键
            bool isCtrlPressed = (Control.ModifierKeys & Keys.Control) == Keys.Control;

            int columnsToAdd = 1;

            // 如果按下Ctrl键，显示输入窗口
            if (isCtrlPressed)
            {
                using (FormAddColumns form = new FormAddColumns())
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        columnsToAdd = form.ColumnsToAdd;  // 改为添加列的数量
                    }
                    else
                    {
                        return;
                    }
                }
            }

            // 确保选中的是表格
            if (selection.Type == PpSelectionType.ppSelectionShapes &&
                selection.ShapeRange.Count == 1 &&
                selection.ShapeRange[1].HasTable == MsoTriState.msoTrue)
            {
                PowerPoint.Table table = selection.ShapeRange[1].Table;

                // 获取表格的当前行数和列数
                int rowCount = table.Rows.Count;
                int colCount = table.Columns.Count;

                // 新增指定数量的列
                for (int k = 0; k < columnsToAdd; k++)
                {
                    table.Columns.Add();
                }
                colCount += columnsToAdd;

                // 创建两个临时列表来存储拼音和汉字内容及其样式
                List<(string text, float fontSize, string fontName, MsoTriState bold)> pinyinList = new List<(string, float, string, MsoTriState)>();
                List<(string text, float fontSize, string fontName, MsoTriState bold)> hanziList = new List<(string, float, string, MsoTriState)>();

                // 将表格中的拼音和汉字内容及其样式分别存储到临时列表中
                for (int i = 1; i <= rowCount; i += 2)
                {
                    for (int j = 1; j <= colCount - columnsToAdd; j++)  // 只存储原始列数的数据
                    {
                        var pinyinCell = table.Cell(i, j).Shape.TextFrame.TextRange;
                        var hanziCell = table.Cell(i + 1, j).Shape.TextFrame.TextRange;

                        pinyinList.Add((pinyinCell.Text, pinyinCell.Font.Size, pinyinCell.Font.Name, pinyinCell.Font.Bold));
                        hanziList.Add((hanziCell.Text, hanziCell.Font.Size, hanziCell.Font.Name, hanziCell.Font.Bold));
                    }
                }

                // 清空所有单元格内容
                for (int i = 1; i <= table.Rows.Count; i++)
                {
                    for (int j = 1; j <= table.Columns.Count; j++)
                    {
                        table.Cell(i, j).Shape.TextFrame.TextRange.Text = string.Empty;
                    }
                }

                // 重新排列表格内容并应用样式
                int index = 0;
                for (int i = 1; i <= rowCount; i += 2)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (index < pinyinList.Count)
                        {
                            var pinyinCell = table.Cell(i, j).Shape.TextFrame.TextRange;
                            var hanziCell = table.Cell(i + 1, j).Shape.TextFrame.TextRange;
                            pinyinCell.Text = pinyinList[index].text;
                            table.Cell(i, j).Shape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorBottom;
                            hanziCell.Text = hanziList[index].text;

                            // 确保奇数行的字号大小为偶数行字号大小的一半
                            pinyinCell.Font.Size = hanziCell.Font.Size / 2;

                            index++;
                        }
                    }
                }

                // 删除多余的空行，保留非空行上方的空行
                for (int i = table.Rows.Count; i > 1; i--)
                {
                    bool isEmpty = true;
                    for (int j = 1; j <= table.Columns.Count; j++)
                    {
                        if (!string.IsNullOrEmpty(table.Cell(i, j).Shape.TextFrame.TextRange.Text))
                        {
                            isEmpty = false;
                            break;
                        }
                    }

                    if (isEmpty)
                    {
                        table.Rows[i].Delete();
                    }
                }

                // 设置所有奇数行的高度与第一行保持一致，偶数行的高度与第二行保持一致
                float firstRowHeight = table.Rows[1].Height;
                float secondRowHeight = table.Rows[2].Height;
                for (int i = 3; i <= table.Rows.Count; i += 2)
                {
                    table.Rows[i].Height = firstRowHeight;
                    if (i + 1 <= table.Rows.Count)
                    {
                        table.Rows[i + 1].Height = secondRowHeight;
                    }
                }

                // 设置所有奇数行的字号与第一行一致，偶数行的字号与第二行一致
                float firstRowFontSize = table.Cell(1, 1).Shape.TextFrame.TextRange.Font.Size;
                float secondRowFontSize = table.Cell(2, 1).Shape.TextFrame.TextRange.Font.Size;

                for (int i = 3; i <= table.Rows.Count; i += 2)
                {
                    for (int j = 1; j <= table.Columns.Count; j++)
                    {
                        table.Cell(i, j).Shape.TextFrame.TextRange.Font.Size = firstRowFontSize;
                        if (i + 1 <= table.Rows.Count)
                        {
                            table.Cell(i + 1, j).Shape.TextFrame.TextRange.Font.Size = secondRowFontSize;
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("请选择一个包含表格的形状。", "增列减行", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void 插入矩形_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.PowerPoint.Application app = Globals.ThisAddIn.Application;
            Slide activeSlide = app.ActiveWindow.View.Slide as Slide;

            // 存储新插入的矩形的名称列表
            List<string> newShapeNames = new List<string>();

            if (app.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionShapes)
            {
                // 选中形状时的操作
                PowerPoint.ShapeRange selectedShapes = app.ActiveWindow.Selection.ShapeRange;

                foreach (Shape selectedShape in selectedShapes)
                {
                    float left = selectedShape.Left;
                    float top = selectedShape.Top;
                    float width = selectedShape.Width;
                    float height = selectedShape.Height;

                    // 插入矩形，并确保位置重合
                    Shape rectangle = activeSlide.Shapes.AddShape(
                        Office.MsoAutoShapeType.msoShapeRectangle,
                        left,
                        top,
                        width,
                        height);
                    // 设置矩形外观为无边框
                    rectangle.Line.Visible = Office.MsoTriState.msoFalse; // 隐藏边框

                    // 判断是否按住了 Ctrl 键
                    bool ctrlPressed = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;

                    if (ctrlPressed)
                    {
                        // 获取选中形状的层次位置
                        int selectedShapeZOrder = selectedShape.ZOrderPosition;

                        // 将新插入的矩形移动到所选形状的下方
                        while (rectangle.ZOrderPosition > selectedShapeZOrder)
                        {
                            rectangle.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                        }
                    }
                    // 将新插入的矩形名称加入列表
                    newShapeNames.Add(rectangle.Name);
                }
            }
            else
            {
                // 没有选中任何形状时，在幻灯片上插入一个与幻灯片等大的矩形
                float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
                float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;

                // 插入矩形
                Shape rectangle = activeSlide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    0,
                    0,
                    slideWidth,
                    slideHeight);

                // 将新插入的矩形名称加入列表
                newShapeNames.Add(rectangle.Name);
                // 设置矩形外观为无边框
                rectangle.Line.Visible = Office.MsoTriState.msoFalse; // 隐藏边框

            }

            // 自动选中新插入的所有矩形
            if (newShapeNames.Count > 0)
            {
                PowerPoint.ShapeRange newShapes = activeSlide.Shapes.Range(newShapeNames.ToArray());
                newShapes.Select(); // 选中新插入的矩形
            }
        }

        private void 调整宽度_Click(object sender, RibbonControlEventArgs e)
        {
            // 检查是否按下了Ctrl键
            if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
            {
                // 打开拼音比例设置窗体
                PinyinRatioForm pinyinRatioForm = new PinyinRatioForm();
                if (pinyinRatioForm.ShowDialog() == true)
                {
                    // 将新的拼音比例和偏移量保存到设置中
                    Properties.Settings.Default.PinyinRatio = (double)pinyinRatioForm.PinyinRatio;
                    Properties.Settings.Default.OffsetValue = pinyinRatioForm.OffsetValue; // 保存偏移量
                    Properties.Settings.Default.Save();
                }
                return;
            }

            // 如果没有按下Ctrl键，继续执行调整宽度的操作
            var pptApp = Globals.ThisAddIn.Application;
            var selection = pptApp.ActiveWindow.Selection;

            if (selection.Type != PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("请先选择文本框和拼音文本框。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            // 从设置中加载已保存的拼音比例和偏移量
            float pinyinRatio = (float)Properties.Settings.Default.PinyinRatio;
            float offsetValue = (float)Properties.Settings.Default.OffsetValue; // 加载偏移量

            // 用于收集错误信息
            int indexErrorCount = 0;

            // 遍历所有选中的文本框
            foreach (Shape hanziTextBox in selection.ShapeRange)
            {
                if (hanziTextBox.Tags["PinYin"] == "True")
                {
                    // 跳过已经是拼音文本框的形状
                    continue;
                }

                try
                {
                    // 遍历每个字符并调整对应的拼音文本框
                    for (int charIndex = 1; charIndex <= hanziTextBox.TextFrame.TextRange.Text.Length; charIndex++)
                    {
                        var pinyinTextBox = GetPinyinTextBoxByCharIndex(selection.ShapeRange, hanziTextBox.Name, charIndex);
                        if (pinyinTextBox == null)
                        {
                            indexErrorCount++;
                            continue;
                        }

                        // 首先检查文本框末尾是否存在空白行
                        bool addedNewLine = false;
                        string originalText = hanziTextBox.TextFrame.TextRange.Text;

                        // 如果文本框末尾没有空白行，则添加一个空白行
                        if (!originalText.EndsWith("\r"))
                        {
                            hanziTextBox.TextFrame.TextRange.Text += "\r";
                            addedNewLine = true;
                        }

                        // 获取当前汉字字符的位置
                        Microsoft.Office.Interop.PowerPoint.TextRange charRange = hanziTextBox.TextFrame.TextRange.Characters(charIndex, 1);

                        // 检查字符是否是标点符号或空格符，跳过这些字符
                        char currentChar = hanziTextBox.TextFrame.TextRange.Text[charIndex - 1];
                        if (char.IsPunctuation(currentChar) || char.IsWhiteSpace(currentChar))
                        {
                            // 如果添加了空白行，现在删除它
                            if (addedNewLine)
                            {
                                hanziTextBox.TextFrame.TextRange.Text = originalText;
                            }
                            continue;  // 跳过标点符号和空格符
                        }

                        // 获取字符基线的位置
                        float baselinePosition = charRange.BoundTop + charRange.BoundHeight - charRange.Font.Size;

                        // 当行间距大于1.5时，减小拼音文本框的上移量，避免其距离字符过远
                        float adjustedPinyinTop;

                        if (charRange.ParagraphFormat.SpaceWithin > 2.0f)
                        {
                            adjustedPinyinTop = baselinePosition - charRange.Font.Size * 1.4f; // 更大行间距时减小更多
                        }
                        else if (charRange.ParagraphFormat.SpaceWithin > 1.5f)
                        {
                            adjustedPinyinTop = baselinePosition - charRange.Font.Size * 1.2f; // 中等行间距时适度减小
                        }
                        else
                        {
                            adjustedPinyinTop = baselinePosition - charRange.Font.Size * 1.1f; // 默认小行间距
                        }

                        // 动态调整左移偏移量，基于字号28时左移3个单位
                        float charLeft = charRange.BoundLeft - ((charRange.Font.Size / 28) * 3);
                        float charWidth = charRange.BoundWidth;

                        // 确保拼音文本框居中于汉字字符
                        float adjustedCharLeft = charLeft + (charWidth - charRange.BoundWidth) / 2;

                        // 调整现有拼音文本框的位置和大小
                        pinyinTextBox.Left = adjustedCharLeft;
                        pinyinTextBox.Top = adjustedPinyinTop + offsetValue; // 应用偏移量
                        pinyinTextBox.Width = charWidth;
                        // 使用保存的拼音比例来调整拼音文本框的字号大小
                        pinyinTextBox.TextFrame.TextRange.Font.Size = charRange.Font.Size * pinyinRatio;

                        // 如果添加了空白行，现在删除它
                        if (addedNewLine)
                        {
                            hanziTextBox.TextFrame.TextRange.Text = originalText;
                        }
                    }
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    // 记录日志或显示警告信息
                    Debug.WriteLine($"字符索引访问失败: {ex.Message}");

                    // 如果选择忽略此错误，可以继续执行后续代码
                    continue;
                }
            }
        }

        private Shape GetPinyinTextBoxByCharIndex(PowerPoint.ShapeRange shapeRange, string hanziTextBoxName, int charIndex)
        {
            foreach (Shape shape in shapeRange)
            {
                if (shape.Tags["PinYin"] == "True" &&
                    shape.Tags["ParentTextBoxName"] == hanziTextBoxName &&
                    int.Parse(shape.Tags["ParentCharIndex"]) == charIndex)
                {
                    return shape;
                }
            }
            return null;
        }

        private void 查看版本_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前插件的版本号
            string version = Assembly.GetExecutingAssembly().GetName().Version.ToString();

            // 显示版本号
            MessageBox.Show($"当前PPT插件的版本号是：{version}", "插件版本", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void 激活插件_Click(object sender, RibbonControlEventArgs e)
        {
            // 提示用户输入激活码
            string inputActivationCode = Microsoft.VisualBasic.Interaction.InputBox("请输入激活码", "插件激活", "");

            // 获取用户的硬件ID
            string hardwareId = HardwareInfoHelper.GetHardwareId();

            // 使用硬件ID生成激活码
            string correctActivationCode = HardwareInfoHelper.GenerateActivationCode(hardwareId);

            // 检查用户输入的激活码是否正确
            if (inputActivationCode == correctActivationCode)
            {
                // 激活成功，可以解锁插件功能
                System.Windows.Forms.MessageBox.Show("激活成功！", "激活", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);

                // 设置插件激活状态为 true，并保存
                Properties.Settings.Default.IsActivated = true;
                Properties.Settings.Default.Save();

                // 启用插件功能
                EnableAllPluginFeatures();  // 添加这行代码，确保插件功能被启用
            }
            else
            {
                // 激活失败
                System.Windows.Forms.MessageBox.Show("激活码无效，请重新输入。", "激活失败", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void EnableAllPluginFeatures()
        {
            // 启用插件的所有功能
            注音编辑.Enabled = true;
            字典注音.Enabled = true;

            // 如果有其他需要启用的功能，可以在这里添加
        }

        private void 获取ID_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取硬件ID
            string hardwareId = HardwareInfoHelper.GetHardwareId();

            // 将硬件ID复制到剪贴板
            Clipboard.SetText(hardwareId);

            // 显示硬件ID给用户，并提示已复制到剪贴板
            System.Windows.Forms.MessageBox.Show($"您的硬件ID: {hardwareId}\n该硬件ID已复制到剪贴板，请将此硬件ID发送给我们以获取激活码。", "硬件ID", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        }

        private void 清除激活_Click(object sender, RibbonControlEventArgs e)
        {
            // 重置激活相关的设置
            Properties.Settings.Default.IsActivated = false;
            // 保存设置
            Properties.Settings.Default.Save();

            // 提示用户激活状态已清除
            System.Windows.Forms.MessageBox.Show("激活状态已清除。插件功能已禁用。", "清除激活", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);

            // 禁用插件功能
            DisableAllPluginFeatures();
        }

        private void DisableAllPluginFeatures()
        {
            // 在这里禁用插件的所有功能
            // 例如：禁用菜单项、按钮等

            注音编辑.Enabled = false;
            字典注音.Enabled = false;

            // 如果有其他需要禁用的功能，可以在这里添加
        }

        private void 汉字拆字_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}