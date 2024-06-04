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
        private CustomCloudTextGenerator cloudTextGenerator;
        

        public Ribbon1(RibbonFactory factory) : base(factory)
        {
            Debug.WriteLine("Ribbon1 constructor called.");
            InitializeCloudTextGenerator();
        }

        private void InitializeCloudTextGenerator()
        {
            if (cloudTextGenerator == null)
            {
                cloudTextGenerator = new CustomCloudTextGenerator();
                Debug.WriteLine("cloudTextGenerator has been initialized.");
            }
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
            InitializeCloudTextGenerator();
        }

        private void button5_Click_1(object sender, RibbonControlEventArgs e)
        {
            Debug.WriteLine("button5_Click_1 called.");
            InitializeCloudTextGenerator();

            if (cloudTextGenerator != null)
            {
                Debug.WriteLine("cloudTextGenerator is not null, calling InitializeForm.");
                cloudTextGenerator.InitializeForm();
            }
            else
            {
                MessageBox.Show("cloudTextGenerator 未被初始化.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debug.WriteLine("cloudTextGenerator is null.");
            }
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

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            // 检查窗口是否已存在并未关闭
            if (form == null || form.IsDisposed)
            {
                CreateSelectionForm();
            }

            UpdateComboBoxOptions(); // 更新下拉框内容
            form.Show(); // 显示窗口
            form.BringToFront(); // 确保窗口在前台
        }

        private void CreateSelectionForm()
        {
            form = new Form()
            {
                Text = "Pinyin Selector",
                Size = new System.Drawing.Size(400, 200),
                StartPosition = FormStartPosition.CenterScreen,
                TopMost = true // 使窗口始终位于最顶部
            };

            ComboBox comboBox = new ComboBox()
            {
                Location = new System.Drawing.Point(20, 20),
                Width = 360,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            form.Controls.Add(comboBox);

            Button refreshButton = new Button()
            {
                Text = "刷新",
                Location = new System.Drawing.Point(20, 60),
                Width = 100,
                Height = 40 // 设置按钮高度为40像素
        };
            refreshButton.Click += (sender, e) =>
            {
                UpdateComboBoxOptions();
            };
            form.Controls.Add(refreshButton);

            Button replaceButton = new Button()
            {
                Text = "注音",
                Location = new System.Drawing.Point(140, 60),
                Width = 100,
                Height = 40 // 设置按钮高度为40像素
            };
            replaceButton.Click += (sender, e) =>
            {
                var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                if (selection.Type == PpSelectionType.ppSelectionText && comboBox.SelectedItem != null)
                {
                    selection.TextRange2.Text = comboBox.SelectedItem.ToString();
                }
            };
            form.Controls.Add(replaceButton);

            // Optional: Add a Close button to manually close the form
            Button closeButton = new Button()
            {
                Text = "退出",
                Location = new System.Drawing.Point(260, 60),
                Width = 100,
                Height = 40 // 设置按钮高度为40像素
            };
            closeButton.Click += (sender, e) => { form.Close(); };
            form.Controls.Add(closeButton);

            form.FormClosing += (sender, e) => { form = null; }; // Reset the form reference when closed
        }

        private void UpdateComboBoxOptions()
        {
            if (form != null && Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionText)
            {
                ComboBox comboBox = form.Controls.OfType<ComboBox>().FirstOrDefault();
                if (comboBox != null)
                {
                    string selectedText = Globals.ThisAddIn.Application.ActiveWindow.Selection.TextRange2.Text.Trim().ToLower();
                    string[] pinyinOptions;

                    if (selectedText == "yu") // 如果选中的文本是 "yu"
                    {
                        pinyinOptions = new string[] { "yū", "yú", "yǔ", "yù" }; // 直接设置拼音选项
                    }
                    else
                    {
                        pinyinOptions = FindClosestPinyin(selectedText); // 否则，按照原来的逻辑查找最接近的拼音选项
                    }

                    comboBox.DataSource = pinyinOptions;
                }
            }
        }

        private string[] FindClosestPinyin(string inputText)
        {
            string[] pinyinLibrary = { "ā", "á", "ǎ", "à", "āi", "ái", "ǎi", "ài", "ān", "án", "ǎn", "àn", "ānɡ", "ánɡ", "ǎnɡ", "ànɡ", "āo", "áo", "ǎo", "ào", "bā", "bá", "bǎ", "bà", "bāi", "bái", "bǎi", "bài", "bān", "bán", "bǎn", "bàn", "bānɡ", "bánɡ", "bǎnɡ", "bànɡ", "bāo", "báo", "bǎo", "bào", "bēi", "béi", "běi", "bèi", "bēn", "bén", "běn", "bèn", "bēnɡ", "bénɡ", "běnɡ", "bènɡ", "bī", "bí", "bǐ", "bì", "biān", "bián", "biǎn", "biàn", "biāo", "biáo", "biǎo", "biào", "biē", "bié", "biě", "biè", "bīn", "bín", "bǐn", "bìn", "bīnɡ", "bínɡ", "bǐnɡ", "bìnɡ", "bō", "bó", "bǒ", "bò", "bū", "bú", "bǔ", "bù", "cā", "cá", "cǎ", "cà", "cāi", "cái", "cǎi", "cài", "cān", "cán", "cǎn", "càn", "cāng", "cáng", "cǎng", "càng", "cāo", "cáo", "cǎo", "cào", "cē", "cé", "cě", "cè", "cēn", "cén", "cěn", "cèn", "cēng", "céng", "cěng", "cèng", "chā", "chá", "chǎ", "chà", "chāi", "chái", "chǎi", "chài", "chān", "chán", "chǎn", "chàn", "chāng", "cháng", "chǎng", "chàng", "chāo", "cháo", "chǎo", "chào", "chē", "ché", "chě", "chè", "chēn", "chén", "chěn", "chèn", "chēng", "chéng", "chěng", "chèng", "chī", "chí", "chǐ", "chì", "chōng", "chóng", "chǒng", "chòng", "chōu", "chóu", "chǒu", "chòu", "chū", "chú", "chǔ", "chù", "chuā", "chuá", "chuǎ", "chuà", "chuāi", "chuái", "chuǎi", "chuài", "chuān", "chuán", "chuǎn", "chuàn", "chuāng", "chuáng", "chuǎng", "chuàng", "chuī", "chuí", "chuǐ", "chuì", "chūn", "chún", "chǔn", "chùn", "chuō", "chuó", "chuǒ", "chuò", "cī", "cí", "cǐ", "cì", "cōng", "cóng", "cǒng", "còng", "cōu", "cóu", "cǒu", "còu", "cū", "cú", "cǔ", "cù", "cuān", "cuán", "cuǎn", "cuàn", "cuī", "cuí", "cuǐ", "cuì", "cūn", "cún", "cǔn", "cùn", "cuō", "cuó", "cuǒ", "cuò", "dā", "dá", "dǎ", "dà", "dāi", "dái", "dǎi", "dài", "dān", "dán", "dǎn", "dàn", "dāng", "dáng", "dǎng", "dàng", "dāo", "dáo", "dǎo", "dào", "dē", "dé", "dě", "dè", "dēi", "déi", "děi", "dèi", "dēn", "dén", "děn", "dèn", "dēng", "déng", "děng", "dèng", "dī", "dí", "dǐ", "dì", "diān", "dián", "diǎn", "diàn", "diāo", "diáo", "diǎo", "diào", "diē", "dié", "diě", "diè", "dīng", "díng", "dǐng", "dìng", "diū", "diú", "diǔ", "diù", "dōng", "dóng", "dǒng", "dòng", "dōu", "dóu", "dǒu", "dòu", "dū", "dú", "dǔ", "dù", "duān", "duán", "duǎn", "duàn", "duī", "duí", "duǐ", "duì", "dūn", "dún", "dǔn", "dùn", "duō", "duó", "duǒ", "duò", "ē", "é", "ě", "è", "ēi", "éi", "ěi", "èi", "ēn", "én", "ěn", "èn", "ēnɡ", "énɡ", "ěnɡ", "ènɡ", "ēr", "ér", "ěr", "èr", "fā", "fá", "fǎ", "fà", "fān", "fán", "fǎn", "fàn", "fāng", "fáng", "fǎng", "fàng", "fēi", "féi", "fěi", "fèi", "fēn", "fén", "fěn", "fèn", "fēng", "féng", "fěng", "fèng", "fō", "fó", "fǒ", "fò", "fōu", "fóu", "fǒu", "fòu", "fū", "fú", "fǔ", "fù", "gā", "gá", "gǎ", "gà", "gāi", "gái", "gǎi", "gài", "gān", "gán", "gǎn", "gàn", "gāng", "gáng", "gǎng", "gàng", "gāo", "gáo", "gǎo", "gào", "gē", "gé", "gě", "gè", "gēi", "géi", "gěi", "gèi", "gēn", "gén", "gěn", "gèn", "gēng", "géng", "gěng", "gèng", "gōng", "góng", "gǒng", "gòng", "gōu", "góu", "gǒu", "gòu", "gū", "gú", "gǔ", "gù", "guā", "guá", "guǎ", "guà", "guāi", "guái", "guǎi", "guài", "guān", "guán", "guǎn", "guàn", "guāng", "guáng", "guǎng", "guàng", "guī", "guí", "guǐ", "guì", "gūn", "gún", "gǔn", "gùn", "guō", "guó", "guǒ", "gùo", "hā", "há", "hǎ", "hà", "hāi", "hái", "hǎi", "hài", "hān", "hán", "hǎn", "hàn", "hāng", "háng", "hǎng", "hàng", "hāo", "háo", "hǎo", "hào", "hē", "hé", "hě", "hè", "hēi", "héi", "hěi", "hèi", "hēn", "hén", "hěn", "hèn", "hēng", "héng", "hěng", "hèng", "hōng", "hóng", "hǒng", "hòng", "hōu", "hóu", "hǒu", "hòu", "hū", "hú", "hǔ", "hù", "huā", "huá", "huǎ", "huà", "huāi", "huái", "huǎi", "huài", "huān", "huán", "huǎn", "huàn", "huāng", "huáng", "huǎng", "huàng", "huī", "huí", "huǐ", "huì", "hūn", "hún", "hǔn", "hùn", "huō", "huó", "huǒ", "huò", "ī", "í", "ǐ", "ì", "iē", "ié", "iě", "iè", "īn", "ín", "ǐn", "ìn", "īnɡ", "ínɡ", "ǐnɡ", "ìnɡ", "iū", "iú", "iǔ", "iù", "jī", "jí", "jǐ", "jì", "jiā", "jiá", "jiǎ", "jià", "jiāi", "jiái", "jiǎi", "jiài", "jiān", "jián", "jiǎn", "jiàn", "jiāng", "jiáng", "jiǎng", "jiàng", "jiāo", "jiáo", "jiǎo", "jiào", "jiē", "jié", "jiě", "jiè", "jīn", "jín", "jǐn", "jìn", "jīng", "jíng", "jǐng", "jìng", "jiōng", "jióng", "jiǒng", "jiòng", "jiū", "jiú", "jiǔ", "jiù", "jū", "jú", "jǔ", "jù", "juān", "juán", "juǎn", "juàn", "juē", "jué", "juě", "juè", "jūn", "jún", "jǔn", "jùn", "kā", "ká", "kǎ", "kà", "kāi", "kái", "kǎi", "kài", "kān", "kán", "kǎn", "kàn", "kāng", "káng", "kǎng", "kàng", "kāo", "káo", "kǎo", "kào", "kē", "ké", "kě", "kè", "kēn", "kén", "kěn", "kèn", "kēng", "kéng", "kěng", "kèng", "kōng", "kóng", "kǒng", "kòng", "kōu", "kóu", "kǒu", "kòu", "kū", "kú", "kǔ", "kù", "kuā", "kuá", "kuǎ", "kuà", "kuāi", "kuái", "kuǎi", "kuài", "kuān", "kuán", "kuǎn", "kuàn", "kuāng", "kuáng", "kuǎng", "kuàng", "kuī", "kuí", "kuǐ", "kuì", "kūn", "kún", "kǔn", "kùn", "kuō", "kuó", "kuǒ", "kuò", "lā", "lá", "lǎ", "là", "laī", "laí", "lǎi", "laì", "lān", "lán", "lǎn", "làn", "lāng", "láng", "lǎng", "làng", "lāo", "láo", "lǎo", "lào", "lē", "lé", "lě", "lè", "lēi", "léi", "lěi", "lèi", "lēng", "léng", "lěng", "lèng", "lī", "lí", "lǐ", "lì", "liān", "lián", "liǎn", "liàn", "liāng", "liáng", "liǎng", "liàng", "liāo", "liáo", "liǎo", "liào", "liē", "lié", "liě", "liè", "līn", "lín", "lǐn", "lìn", "līng", "líng", "lǐng", "lìng", "liū", "liú", "liǔ", "liù", "lōng", "lóng", "lǒng", "lòng", "lōu", "lóu", "lǒu", "lòu", "lū", "lú", "lǔ", "lù", "luān", "luán", "luǎn", "luàn", "luē", "lué", "luě", "luè", "lūn", "lún", "lǔn", "lùn", "luō", "luó", "luǒ", "luò", "mā", "má", "mǎ", "mà", "maī", "maí", "mǎi", "maì", "mān", "mán", "mǎn", "màn", "māng", "máng", "mǎng", "màng", "māo", "máo", "mǎo", "mào", "mē", "mé", "mě", "mè", "mēi", "méi", "měi", "mèi", "mēn", "mén", "měn", "mèn", "mēng", "méng", "měng", "mèng", "mī", "mí", "mǐ", "mì", "miān", "mián", "miǎn", "miàn", "miāo", "miáo", "miǎo", "miào", "miē", "mié", "miě", "miè", "mīn", "mín", "mǐn", "mìn", "mīng", "míng", "mǐng", "mìng", "miū", "miú", "miǔ", "miù", "mō", "mó", "mǒ", "mò", "mōu", "móu", "mǒu", "mòu", "mū", "mú", "mǔ", "mù", "nā", "ná", "nǎ", "nà", "naī", "naí", "nǎi", "naì", "nān", "nán", "nǎn", "nàn", "nāng", "náng", "nǎng", "nàng", "nāo", "náo", "nǎo", "nào", "nē", "né", "ně", "nè", "neī", "neí", "něi", "neì", "nēn", "nén", "něn", "nèn", "nēng", "néng", "něng", "nèng", "nī", "ní", "nǐ", "nì", "niān", "nián", "niǎn", "niàn", "niāng", "niáng", "niǎng", "niàng", "niāo", "niáo", "niǎo", "niào", "niē", "nié", "niě", "niè", "nīn", "nín", "nǐn", "nìn", "nīng", "níng", "nǐng", "nìng", "niū", "niú", "niǔ", "niù", "nōng", "nóng", "nǒng", "nòng", "nū", "nú", "nǔ", "nù", "nuān", "nuán", "nuǎn", "nuàn", "nuē", "nué", "nuě", "nuè", "nuō", "nuó", "nuǒ", "nuò", "nǘ", "nǚ", "nǜ", "nǘè", "ō", "ó", "ǒ", "ò", "ōnɡ", "ónɡ", "ǒnɡ", "ònɡ", "ōu", "óu", "ǒu", "òu", "pā", "pá", "pǎ", "pà", "pāi", "pái", "pǎi", "pài", "pān", "pán", "pǎn", "pàn", "pāng", "páng", "pǎng", "pàng", "pāo", "páo", "pǎo", "pào", "pēi", "péi", "pěi", "pèi", "pēn", "pén", "pěn", "pèn", "pēng", "péng", "pěng", "pèng", "pī", "pí", "pǐ", "pì", "piān", "pián", "piǎn", "piàn", "piāo", "piáo", "piǎo", "piào", "piē", "pié", "piě", "piè", "pīn", "pín", "pǐn", "pìn", "pīng", "píng", "pǐng", "pìng", "pō", "pó", "pǒ", "pò", "pōu", "póu", "pǒu", "pòu", "pū", "pú", "pǔ", "pù", "qī", "qí", "qǐ", "qì", "qiā", "qiá", "qǐa", "qìa", "qiān", "qián", "qǐan", "qìan", "qiāng", "qiáng", "qǐang", "qìang", "qiāo", "qiáo", "qǐao", "qìao", "qiē", "qié", "qǐe", "qìe", "qīn", "qín", "qǐn", "qìn", "qīng", "qíng", "qǐng", "qìng", "qiōnɡ", "qiónɡ", "qiǒnɡ", "qiònɡ", "qiū", "qiú", "qiǔ", "qiù", "qū", "qú", "qǔ", "qù", "quān", "quán", "quǎn", "quàn", "quē", "qué", "quě", "què", "qūn", "qún", "qǔn", "qùn", "rān", "rán", "rǎn", "ràn", "rāng", "ráng", "rǎng", "ràng", "rāo", "ráo", "rǎo", "rào", "rē", "ré", "rě", "rè", "rēn", "rén", "rěn", "rèn", "rēng", "réng", "rěng", "rèng", "rī", "rí", "rǐ", "rì", "rōng", "róng", "rǒng", "ròng", "rōu", "róu", "rǒu", "ròu", "rū", "rú", "rǔ", "rù", "ruān", "ruán", "ruǎn", "ruàn", "ruī", "ruí", "ruǐ", "ruì", "rūn", "rún", "rǔn", "rùn", "ruō", "ruó", "ruǒ", "ruò", "sā", "sá", "sǎ", "sà", "sāi", "sái", "sǎi", "sài", "sān", "sán", "sǎn", "sàn", "sāng", "sáng", "sǎng", "sàng", "sāo", "sáo", "sǎo", "sào", "sē", "sé", "sě", "sè", "sēn", "sén", "sěn", "sèn", "sēng", "séng", "sěng", "sèng", "shā", "shá", "shǎ", "shà", "shāi", "shái", "shǎi", "shài", "shān", "shán", "shǎn", "shàn", "shāng", "sháng", "shǎng", "shàng", "shāo", "sháo", "shǎo", "shào", "shē", "shé", "shě", "shè", "shēi", "shéi", "shěi", "shèi", "shēn", "shén", "shěn", "shèn", "shēng", "shéng", "shěng", "shèng", "shī", "shí", "shǐ", "shì", "shōu", "shóu", "shǒu", "shòu", "shū", "shú", "shǔ", "shù", "shuā", "shuá", "shuǎ", "shuà", "shuāi", "shuái", "shuǎi", "shuài", "shuān", "shuán", "shuǎn", "shuàn", "shuāng", "shuáng", "shuǎng", "shuàng", "shuī", "shuí", "shuǐ", "shuì", "shu?n", "shú", "shǔn", "shùn", "shuō", "shuó", "shuǒ", "shuò", "sī", "sí", "sǐ", "sì", "sōng", "sóng", "sǒng", "sòng", "sū", "sú", "sǔ", "sù", "suān", "suán", "suǎn", "suàn", "suī", "suí", "suǐ", "suì", "su?n", "sún", "sǔn", "sùn", "suō", "suó", "suǒ", "suò", "tā", "tá", "tǎ", "tà", "tāi", "tái", "tǎi", "tài", "tān", "tán", "tǎn", "tàn", "tāng", "táng", "tǎng", "tàng", "tāo", "táo", "tǎo", "tào", "tē", "té", "tě", "tè", "tēng", "téng", "těng", "tèng", "tī", "tí", "tǐ", "tì", "tiān", "tián", "tiǎn", "tiàn", "tiāo", "tiáo", "tiǎo", "tiào", "tiē", "tié", "tiě", "tiè", "tīng", "tíng", "tǐng", "tìng", "tōng", "tóng", "tǒng", "tòng", "tōu", "tóu", "tǒu", "tòu", "tū", "tú", "tǔ", "tù", "tuān", "tuán", "tuǎn", "tuàn", "tuī", "tuí", "tuǐ", "tuì", "tūn", "tún", "tǔn", "tùn", "tuō", "tuó", "tuǒ", "tuò", "ū", "ú", "ǔ", "ù", "uē", "ué", "uě", "uè", "uī", "uí", "uǐ", "uì", "ūn", "ún", "ǔn", "ùn", "ǖ", "ǘ", "ǚ", "ǜ", "ǖn", "ǘn", "ǚn", "ǜn", "wā", "wá", "wǎ", "wà", "wāi", "wái", "wǎi", "wài", "wān", "wán", "wǎn", "wàn", "wāng", "wáng", "wǎng", "wàng", "wēi", "wéi", "wěi", "wèi", "wēn", "wén", "wěn", "wèn", "wēng", "wéng", "wěng", "wèng", "wō", "wó", "wǒ", "wò", "wū", "wú", "wǔ", "wù", "xī", "xí", "xǐ", "xì", "xiā", "xiá", "xiǎ", "xià", "xiān", "xián", "xiǎn", "xiàn", "xiāng", "xiáng", "xiǎng", "xiàng", "xiāo", "xiáo", "xiǎo", "xiào", "xiē", "xié", "xiě", "xiè", "xīn", "xín", "xǐn", "xìn", "xīng", "xíng", "xǐng", "xìng", "xiōng", "xióng", "xiǒng", "xiòng", "xiū", "xiú", "xiǔ", "xiù", "xū", "xú", "xǔ", "xù", "xuān", "xuán", "xuǎn", "xuàn", "xuē", "xué", "xuě", "xuè", "xūn", "xún", "xǔn", "xùn", "yā", "yá", "yǎ", "yà", "yān", "yán", "yǎn", "yàn", "yāng", "yáng", "yǎng", "yàng", "yāo", "yáo", "yǎo", "yào", "yē", "yé", "yě", "yè", "yī", "yí", "yǐ", "yì", "yīn", "yín", "yǐn", "yìn", "yīng", "yíng", "yǐng", "yìng", "yō", "yó", "yǒ", "yò", "yōng", "yóng", "yǒng", "yòng", "yōu", "yóu", "yǒu", "yòu", "yū", "yú", "yǔ", "yù", "yuān", "yuán", "yuǎn", "yuàn", "yuē", "yué", "yuě", "yuè", "yūn", "yún", "yǔn", "yùn", "zā", "zá", "zǎ", "zà", "zāi", "zái", "zǎi", "zài", "zān", "zán", "zǎn", "zàn", "zāng", "záng", "zǎng", "zàng", "zāo", "záo", "zǎo", "zào", "zē", "zé", "zě", "zè", "zēi", "zéi", "zěi", "zèi", "zēn", "zén", "zěn", "zèn", "zēng", "zéng", "zěng", "zèng", "zhā", "zhá", "zhǎ", "zhà", "zhāi", "zhái", "zhǎi", "zhài", "zhān", "zhán", "zhǎn", "zhàn", "zhāng", "zháng", "zhǎng", "zhàng", "zhāo", "zháo", "zhǎo", "zhào", "zhē", "zhé", "zhě", "zhè", "zhēi", "zhéi", "zhěi", "zhèi", "zhēn", "zhén", "zhěn", "zhèn", "zhēng", "zhéng", "zhěng", "zhèng", "zhī", "zhí", "zhǐ", "zhì", "zhōng", "zhóng", "zhǒng", "zhòng", "zhōu", "zhóu", "zhǒu", "zhòu", "zhū", "zhú", "zhǔ", "zhù", "zhuā", "zhuá", "zhuǎ", "zhuà", "zhuāi", "zhuái", "zhuǎi", "zhuài", "zhuān", "zhuán", "zhuǎn", "zhuàn", "zhuānɡ", "zhuánɡ", "zhuǎnɡ", "zhuàng", "zhuī", "zhuí", "zhuǐ", "zhuì", "zhūn", "zhún", "zhǔn", "zhùn", "zī", "zí", "zǐ", "zì", "zōng", "zóng", "zǒng", "zòng", "zōu", "zóu", "zǒu", "zòu", "zū", "zú", "zǔ", "zù", "zuān", "zuán", "zuǎn", "zuàn", "zuī", "zuí", "zuǐ", "zuì", "zūn", "zún", "zǔn", "zùn", "zuō", "zuó", "zuǒ", "zuò", };
            return pinyinLibrary.Where(p => NormalizePinyin(p).StartsWith(NormalizePinyin(inputText))).Take(4).ToArray();
        }


        private string NormalizePinyin(string pinyin)
        {
            Dictionary<char, char> accentMapping = new Dictionary<char, char>
            {
                { 'ā', 'a' }, { 'á', 'a' }, { 'ǎ', 'a' }, { 'à', 'a' },
                { 'ō', 'o' }, { 'ó', 'o' }, { 'ǒ', 'o' }, { 'ò', 'o' },
                { 'ē', 'e' }, { 'é', 'e' }, { 'ě', 'e' }, { 'è', 'e' },
                { 'ī', 'i' }, { 'í', 'i' }, { 'ǐ', 'i' }, { 'ì', 'i' },
                { 'ū', 'u' }, { 'ú', 'u' }, { 'ǔ', 'u' }, { 'ù', 'u' },
                { 'ǖ', 'ü' }, { 'ǘ', 'ü' }, { 'ǚ', 'ü' }, { 'ǜ', 'ü' }
               
            };
            return new string(pinyin.Select(c => accentMapping.ContainsKey(c) ? accentMapping[c] : c).ToArray());
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
        private List<string> selectedSVGs = new List<string>();
        private const int MaxPerRow = 5;
        private const int Margin = 10;
        private int currentX = 50;
        private int currentY = 50;



        private void button8_Click(object sender, RibbonControlEventArgs e)
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
            Form svgSelectionForm = new Form();
            svgSelectionForm.Text = "请选择生字笔顺SVG分步图";
            svgSelectionForm.Size = new System.Drawing.Size(600, 400);

            ListBox svgListBox = new ListBox();
            svgListBox.Dock = DockStyle.Fill;

            int count = 1;
            foreach (var svgNode in svgNodes)
            {
                string scientificName = $"{inputChar}-第{count}笔";
                svgListBox.Items.Add(new KeyValuePair<string, string>(scientificName, svgNode.OuterHtml));
                count++;
            }

            svgListBox.SelectionMode = SelectionMode.MultiExtended;
            svgSelectionForm.Controls.Add(svgListBox);

            Button selectButton = new Button();
            selectButton.Text = "确认插入";
            selectButton.Size = new System.Drawing.Size(150, 50);
            selectButton.Dock = DockStyle.Bottom;
            selectButton.Click += (sender, e) =>
            {
                List<string> selectedSVGs = new List<string>();
                foreach (var selectedItem in svgListBox.SelectedItems)
                {
                    var kvp = (KeyValuePair<string, string>)selectedItem;
                    selectedSVGs.Add(kvp.Value);
                }
                svgSelectionForm.Close();
                InsertSVGsIntoPresentation(selectedSVGs, inputChar);
            };
            svgSelectionForm.Controls.Add(selectButton);
            svgSelectionForm.ShowDialog();
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

        /// <summary>
        /// 在给定的SVG字符串中插入新的宽度和高度属性。
        /// </summary>
        /// <param name="svg">原始的SVG字符串</param>
        /// <param name="width">要插入的宽度值</param>
        /// <param name="height">要插入的高度值</param>
        /// <returns>带有新属性的SVG字符串</returns>
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

        private PowerPoint.Shape tableShape;
        private Form settingsForm;
        private Color borderColor = Color.Black;

        private void button9_Click(object sender, RibbonControlEventArgs e)
        {
            if (settingsForm == null || settingsForm.IsDisposed)
            {
                settingsForm = new Form
                {
                    Text = "设置表格边框",
                    Size = new Size(400, 500),
                    StartPosition = FormStartPosition.CenterScreen
                };

                Label labelRows = new Label { Text = "行数:", Location = new DrawingPoint(65, 40), AutoSize = true };
                NumericUpDown numericUpDownRows = new NumericUpDown { Location = new DrawingPoint(190, 40), Minimum = 1, Maximum = 10, Value = 2 };
                numericUpDownRows.ValueChanged += ApplySettings;

                Label labelColumns = new Label { Text = "列数:", Location = new DrawingPoint(65, 80), AutoSize = true };
                NumericUpDown numericUpDownColumns = new NumericUpDown { Location = new DrawingPoint(190, 80), Minimum = 1, Maximum = 10, Value = 2 };
                numericUpDownColumns.ValueChanged += ApplySettings;

                Label labelRowSpacing = new Label { Text = "行间距:", Location = new DrawingPoint(65, 120), AutoSize = true };
                NumericUpDown numericUpDownRowSpacing = new NumericUpDown { Location = new DrawingPoint(190, 120), Minimum = 0, Maximum = 100, Value = 10 };
                numericUpDownRowSpacing.ValueChanged += ApplySettings;

                Label labelColumnSpacing = new Label { Text = "列间距:", Location = new DrawingPoint(65, 160), AutoSize = true };
                NumericUpDown numericUpDownColumnSpacing = new NumericUpDown { Location = new DrawingPoint(190, 160), Minimum = 0, Maximum = 100, Value = 10 };
                numericUpDownColumnSpacing.ValueChanged += ApplySettings;

                Label labelWidth = new Label { Text = "边框宽度:", Location = new DrawingPoint(65, 200), AutoSize = true };
                NumericUpDown numericUpDownBorderWidth = new NumericUpDown { Location = new DrawingPoint(190, 200), Minimum = 0, Maximum = 10, DecimalPlaces = 2, Value = 1.25m };
                numericUpDownBorderWidth.ValueChanged += ApplySettings;

                Label labelScale = new Label { Text = "缩放比例:", Location = new DrawingPoint(65, 300), AutoSize = true };
                TrackBar trackBarScale = new TrackBar { Location = new DrawingPoint(190, 300), Minimum = 50, Maximum = 200, Value = 100, TickFrequency = 10, Width = 120 };
                trackBarScale.ValueChanged += ApplySettings;

                Label labelColor = new Label { Text = "边框颜色:", Location = new DrawingPoint(65, 240), AutoSize = true };
                Button buttonChooseColor = new Button { Text = "自定义", Location = new DrawingPoint(190, 240), Size = new Size(120, 40) };

                buttonChooseColor.Click += (s, args) =>
                {
                    using (ColorDialog colorDialog = new ColorDialog())
                    {
                        if (colorDialog.ShowDialog() == DialogResult.OK)
                        {
                            borderColor = colorDialog.Color;
                            ApplySettings(s, args);
                        }
                    }
                };

                Button buttonOK = new Button { Text = "生成", Location = new DrawingPoint(65, 350), Size = new Size(100, 40) };

                settingsForm.Controls.Add(labelRows);
                settingsForm.Controls.Add(numericUpDownRows);
                settingsForm.Controls.Add(labelColumns);
                settingsForm.Controls.Add(numericUpDownColumns);
                settingsForm.Controls.Add(labelRowSpacing);
                settingsForm.Controls.Add(numericUpDownRowSpacing);
                settingsForm.Controls.Add(labelColumnSpacing);
                settingsForm.Controls.Add(numericUpDownColumnSpacing);
                settingsForm.Controls.Add(labelWidth);
                settingsForm.Controls.Add(numericUpDownBorderWidth);
                settingsForm.Controls.Add(labelScale);
                settingsForm.Controls.Add(trackBarScale);
                settingsForm.Controls.Add(labelColor);
                settingsForm.Controls.Add(buttonChooseColor);
                settingsForm.Controls.Add(buttonOK);

                buttonOK.Click += GenerateTable;
            }

            settingsForm.Show();
            settingsForm.TopMost = true;
        }

        private void GenerateTable(object sender, EventArgs e)
        {
            int rows = (int)((NumericUpDown)settingsForm.Controls[1]).Value;
            int columns = (int)((NumericUpDown)settingsForm.Controls[3]).Value;
            float rowSpacing = (float)((NumericUpDown)settingsForm.Controls[5]).Value;
            float columnSpacing = (float)((NumericUpDown)settingsForm.Controls[7]).Value;
            float borderWidth = (float)((NumericUpDown)settingsForm.Controls[9]).Value;
            float scale = ((TrackBar)settingsForm.Controls[11]).Value / 100f;

            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Slide activeSlide = app.ActiveWindow.View.Slide as PowerPoint.Slide;

            float startX = 100; // 初始X位置，可以根据需要调整
            float startY = 100; // 初始Y位置，可以根据需要调整
            float squareSize = 100 * scale; // 根据缩放比例调整每个正方形表格的大小

            // 保存现有对象的位置和顺序
            var originalShapes = new List<PowerPoint.Shape>();
            for (int i = 1; i <= activeSlide.Shapes.Count; i++)
            {
                originalShapes.Add(activeSlide.Shapes[i]);
            }

            // 生成表格
            List<PowerPoint.Shape> newTableShapes = new List<PowerPoint.Shape>();
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    float left = startX + j * (squareSize + columnSpacing);
                    float top = startY + i * (squareSize + rowSpacing);

                    PowerPoint.Shape tableShape = activeSlide.Shapes.AddTable(2, 2, left, top, squareSize, squareSize);
                    tableShape.LockAspectRatio = Office.MsoTriState.msoTrue; // 锁定纵横比

                    PowerPoint.Table table = tableShape.Table;
                    SetTableProperties(table, borderWidth, borderColor);

                    newTableShapes.Add(tableShape);
                }
            }

            // 检查是否有选中的对象
            try
            {
                var selection = app.ActiveWindow.Selection;
                if (selection.ShapeRange.Count > 0)
                {
                    int shapeIndex = 0;
                    foreach (PowerPoint.Shape selectedShape in selection.ShapeRange)
                    {
                        float left = startX + (shapeIndex % columns) * (squareSize + columnSpacing);
                        float top = startY + (shapeIndex / columns) * (squareSize + rowSpacing);

                        selectedShape.Left = left + (squareSize - selectedShape.Width) / 2;
                        selectedShape.Top = top + (squareSize - selectedShape.Height) / 2;

                        shapeIndex++;
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // 没有选中任何对象
            }

            // 恢复原始对象位置和顺序
            foreach (var shape in originalShapes)
            {
                shape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            }

            // 确保新表格在最前面
            foreach (var tableShape in newTableShapes)
            {
                tableShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            }
        }

        private void ApplySettings(object sender, EventArgs e)
        {
            int rows = (int)((NumericUpDown)settingsForm.Controls[1]).Value;
            int columns = (int)((NumericUpDown)settingsForm.Controls[3]).Value;
            float borderWidth = (float)((NumericUpDown)settingsForm.Controls[9]).Value;
            float scale = ((TrackBar)settingsForm.Controls[11]).Value / 100f;
            float rowSpacing = (float)((NumericUpDown)settingsForm.Controls[5]).Value;
            float columnSpacing = (float)((NumericUpDown)settingsForm.Controls[7]).Value;

            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Slide activeSlide = app.ActiveWindow.View.Slide as PowerPoint.Slide;

            // 检查是否有选中的对象
            try
            {
                var selection = app.ActiveWindow.Selection;
                if (selection.ShapeRange.Count == 0)
                {
                    // 没有选中对象，不执行对齐和大小调整
                    return;
                }

                float startX = 100; // 初始X位置，可以根据需要调整
                float startY = 100; // 初始Y位置，可以根据需要调整
                float squareSize = 100 * scale; // 根据缩放比例调整每个正方形表格的大小

                int shapeIndex = 0;

                foreach (PowerPoint.Shape selectedShape in selection.ShapeRange)
                {
                    float left = startX + (shapeIndex % columns) * (squareSize + columnSpacing);
                    float top = startY + (shapeIndex / columns) * (squareSize + rowSpacing);

                    selectedShape.Left = left + (squareSize - selectedShape.Width) / 2;
                    selectedShape.Top = top + (squareSize - selectedShape.Height) / 2;

                    if (selectedShape.Type == Office.MsoShapeType.msoTable)
                    {
                        PowerPoint.Table table = selectedShape.Table;
                        SetTableProperties(table, borderWidth, borderColor);
                    }

                    shapeIndex++;
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // 没有选中任何对象
            }
        }

        private void SetTableProperties(PowerPoint.Table table, float borderWidth, Color borderColor)
        {
            int colorRgb = ConvertColor(borderColor);

            for (int i = 1; i <= table.Rows.Count; i++)
            {
                for (int j = 1; j <= table.Columns.Count; j++)
                {
                    PowerPoint.Cell cell = table.Cell(i, j);

                    cell.Shape.Fill.Transparency = 1;
                    cell.Shape.TextFrame.TextRange.Font.Size = 1;

                    if (i == 1)
                    {
                        SetCellBorder(cell.Borders[PowerPoint.PpBorderType.ppBorderTop], borderWidth, colorRgb, true);
                    }
                    if (i == table.Rows.Count)
                    {
                        SetCellBorder(cell.Borders[PowerPoint.PpBorderType.ppBorderBottom], borderWidth, colorRgb, true);
                    }
                    if (j == 1)
                    {
                        SetCellBorder(cell.Borders[PowerPoint.PpBorderType.ppBorderLeft], borderWidth, colorRgb, true);
                    }
                    if (j == table.Columns.Count)
                    {
                        SetCellBorder(cell.Borders[PowerPoint.PpBorderType.ppBorderRight], borderWidth, colorRgb, true);
                    }

                    if (i < table.Rows.Count)
                    {
                        SetCellBorder(cell.Borders[PowerPoint.PpBorderType.ppBorderBottom], borderWidth, colorRgb, false);
                    }
                    if (j < table.Columns.Count)
                    {
                        SetCellBorder(cell.Borders[PowerPoint.PpBorderType.ppBorderRight], borderWidth, colorRgb, false);
                    }
                }
            }
        }

        private void SetCellBorder(PowerPoint.LineFormat border, float borderWidth, int colorRgb, bool isOuterCell)
        {
            border.Weight = borderWidth;
            border.ForeColor.RGB = colorRgb;
            border.Visible = Office.MsoTriState.msoTrue;
            border.DashStyle = isOuterCell ? Office.MsoLineDashStyle.msoLineSolid : Office.MsoLineDashStyle.msoLineDash;
        }

        private int ConvertColor(Color color)
        {
            return (color.B << 16) | (color.G << 8) | color.R;
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

        private void button12_Click(object sender, RibbonControlEventArgs e)
        {
            if (settingsFormButton12 == null || settingsFormButton12.IsDisposed)
            {
                settingsFormButton12 = new Form
                {
                    Text = "设置表格边框",
                    Size = new Size(400, 300),
                    StartPosition = FormStartPosition.CenterScreen
                };

                Label labelWidth = new Label { Text = "边框宽度:", Location = new System.Drawing.Point(65, 40), AutoSize = true };
                NumericUpDown numericUpDownBorderWidth = new NumericUpDown { Location = new System.Drawing.Point(190, 40), Minimum = 0, Maximum = 10, DecimalPlaces = 2, Value = 1.25m };

                Label labelColor = new Label { Text = "边框颜色:", Location = new System.Drawing.Point(65, 90), AutoSize = true };
                Button buttonChooseColor = new Button { Text = "自定义", Location = new System.Drawing.Point(190, 90), Size = new Size(120, 40) };

                buttonChooseColor.Click += (s, args) =>
                {
                    using (ColorDialog colorDialog = new ColorDialog())
                    {
                        if (colorDialog.ShowDialog() == DialogResult.OK)
                        {
                            borderColorButton12 = colorDialog.Color;
                        }
                    }
                };

                Button buttonOK = new Button { Text = "生成", Location = new System.Drawing.Point(75, 155), Size = new Size(100, 40) };
                Button buttonApply = new Button { Text = "应用", Location = new System.Drawing.Point(200, 155), Size = new Size(100, 40) };

                settingsFormButton12.Controls.Add(labelWidth);
                settingsFormButton12.Controls.Add(numericUpDownBorderWidth);
                settingsFormButton12.Controls.Add(labelColor);
                settingsFormButton12.Controls.Add(buttonChooseColor);
                settingsFormButton12.Controls.Add(buttonOK);
                settingsFormButton12.Controls.Add(buttonApply);

                buttonOK.Click += GenerateTableButton12;
                buttonApply.Click += ApplySettingsButton12;
            }

            settingsFormButton12.Show();
            settingsFormButton12.TopMost = true;
        }

        private void GenerateTableButton12(object sender, EventArgs e)
        {
            float borderWidth = (float)((NumericUpDown)settingsFormButton12.Controls[1]).Value;
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Slide activeSlide = app.ActiveWindow.View.Slide as PowerPoint.Slide;

            if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange selectedShapes = app.ActiveWindow.Selection.ShapeRange;

                foreach (PowerPoint.Shape selectedShape in selectedShapes)
                {
                    float selectedSize = Math.Max(selectedShape.Width, selectedShape.Height) + 18;
                    float left = selectedShape.Left + (selectedShape.Width - selectedSize) / 2;
                    float top = selectedShape.Top + (selectedShape.Height - selectedSize) / 2;

                    PowerPoint.Shape tableShapeButton12 = activeSlide.Shapes.AddTable(2, 2, left, top, selectedSize, selectedSize);
                    tableShapeButton12.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue; // 锁定纵横比

                    PowerPoint.Table table = tableShapeButton12.Table;

                    SetTablePropertiesButton12(table, borderWidth, borderColorButton12);

                    // 将表格置于底层
                    tableShapeButton12.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack);
                }
            }
        }

        private PowerPoint.Shape tableShapeButton12;
        private Form settingsFormButton12;
        private Color borderColorButton12 = Color.Black;

        private void ApplySettingsButton12(object sender, EventArgs e)
        {
            float borderWidth = (float)((NumericUpDown)settingsFormButton12.Controls[1]).Value;
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Slide activeSlide = app.ActiveWindow.View.Slide as PowerPoint.Slide;

            if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange selectedShapes = app.ActiveWindow.Selection.ShapeRange;

                foreach (PowerPoint.Shape selectedShape in selectedShapes)
                {
                    if (selectedShape.Type == Office.MsoShapeType.msoTable)
                    {
                        PowerPoint.Table table = selectedShape.Table;
                        SetTablePropertiesButton12(table, borderWidth, borderColorButton12);
                    }
                }
            }
        }

        private void SetTablePropertiesButton12(PowerPoint.Table table, float borderWidth, Color borderColor)
        {
            int colorRgb = ConvertColorButton12(borderColor);

            for (int i = 1; i <= table.Rows.Count; i++)
            {
                for (int j = 1; j <= table.Columns.Count; j++)
                {
                    PowerPoint.Cell cell = table.Cell(i, j);

                    cell.Shape.Fill.Transparency = 1;
                    cell.Shape.TextFrame.TextRange.Font.Size = 1; // 设置字号为1

                    SetCellBorderButton12(cell.Borders[PowerPoint.PpBorderType.ppBorderTop], borderWidth, colorRgb);
                    SetCellBorderButton12(cell.Borders[PowerPoint.PpBorderType.ppBorderBottom], borderWidth, colorRgb);
                    SetCellBorderButton12(cell.Borders[PowerPoint.PpBorderType.ppBorderLeft], borderWidth, colorRgb);
                    SetCellBorderButton12(cell.Borders[PowerPoint.PpBorderType.ppBorderRight], borderWidth, colorRgb);
                }
            }

            table.Cell(1, 1).Borders[PowerPoint.PpBorderType.ppBorderBottom].DashStyle = Office.MsoLineDashStyle.msoLineDash;
            table.Cell(1, 1).Borders[PowerPoint.PpBorderType.ppBorderRight].DashStyle = Office.MsoLineDashStyle.msoLineDash;
            table.Cell(1, 2).Borders[PowerPoint.PpBorderType.ppBorderBottom].DashStyle = Office.MsoLineDashStyle.msoLineDash;
            table.Cell(2, 1).Borders[PowerPoint.PpBorderType.ppBorderRight].DashStyle = Office.MsoLineDashStyle.msoLineDash;
        }

        private void SetCellBorderButton12(PowerPoint.LineFormat border, float borderWidth, int colorRgb)
        {
            border.Weight = borderWidth;
            border.ForeColor.RGB = colorRgb;
            border.Visible = Office.MsoTriState.msoTrue;
        }

        private int ConvertColorButton12(Color color)
        {
            return (color.B << 16) | (color.G << 8) | color.R;
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

        private List<PowerPoint.Shape> copiedShapes = new List<PowerPoint.Shape>();
        private Dictionary<int, (float Width, float Height)> initialSizes = new Dictionary<int, (float Width, float Height)>();
        private void button16_Click(object sender, RibbonControlEventArgs e)
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
                    ShowCircularDistributionForm(true, pptApp, selectedShapes, radius, initialRotation, finalRotation, sizeIncrement, copyCount);
                }
                else
                {
                    PerformCircularDistribution(pptApp, selectedShapes, radius, initialRotation, finalRotation, sizeIncrement, false);
                    ShowCircularDistributionForm(false, pptApp, selectedShapes, radius, initialRotation, finalRotation, sizeIncrement, copyCount);
                }
            }
            else
            {
                MessageBox.Show("请选择至少一个对象。");
            }
        }

        void PerformCircularDistribution(PowerPoint.Application pptApp, PowerPoint.ShapeRange shapes, float radius, float initialRotation, float finalRotation, float sizeIncrement, bool isCopyMode, int copyCount = 0)
        {
            if (isCopyMode)
            {
                foreach (PowerPoint.Shape shape in copiedShapes)
                {
                    shape.Delete();
                }
                copiedShapes.Clear();
            }

            int count = isCopyMode ? copyCount : shapes.Count;
            float angleStep = 360.0f / count;
            float angleIncrement = (finalRotation - initialRotation) / count;

            for (int i = 0; i < count; i++)
            {
                float angle = initialRotation + i * angleStep;
                float radians = angle * (float)(Math.PI / 180.0);
                float newX = (float)(radius * Math.Cos(radians));
                float newY = (float)(radius * Math.Sin(radians));

                PowerPoint.Shape shape;
                if (isCopyMode)
                {
                    shape = shapes[1].Duplicate()[1];
                    copiedShapes.Add(shape);
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
                    float newSize = shape.Width + i * sizeIncrement;
                    shape.Width = newSize;
                    shape.Height = newSize;
                }
            }
        }

        void ShowCircularDistributionForm(bool isCopyMode, PowerPoint.Application pptApp, PowerPoint.ShapeRange shapes, float radius, float initialRotation, float finalRotation, float sizeIncrement, int copyCount)
        {
            Form form = new Form
            {
                Text = isCopyMode ? "环形-复制分布模式" : "环形分布模式",
                Size = new System.Drawing.Size(600, 500),
                FormBorderStyle = FormBorderStyle.SizableToolWindow,
                StartPosition = FormStartPosition.CenterScreen
            };

            int yPosition = 20;

            Label radiusLabel = new Label { Location = new System.Drawing.Point(20, yPosition), Size = new System.Drawing.Size(150, 30), Text = "环形半径：" };
            TrackBar radiusTrackBar = new TrackBar { Location = new System.Drawing.Point(200, yPosition), Size = new System.Drawing.Size(350, 45), Minimum = 10, Maximum = 500, Value = (int)radius };
            yPosition += 100;

            Label initialRotationLabel = new Label { Location = new System.Drawing.Point(20, yPosition), Size = new System.Drawing.Size(210, 30), Text = "旋转递进起始角度：" };
            NumericUpDown initialRotationUpDown = new NumericUpDown { Location = new System.Drawing.Point(235, yPosition), Size = new System.Drawing.Size(80, 30), Minimum = 0, Maximum = 360, Value = (int)initialRotation };

            Label finalRotationLabel = new Label { Location = new System.Drawing.Point(320, yPosition), Size = new System.Drawing.Size(120, 30), Text = "~终点角度：" };
            NumericUpDown finalRotationUpDown = new NumericUpDown { Location = new System.Drawing.Point(450, yPosition), Size = new System.Drawing.Size(80, 30), Minimum = 0, Maximum = 360, Value = (int)finalRotation };
            yPosition += 100;

            Label sizeIncrementLabel = new Label { Location = new System.Drawing.Point(20, yPosition), Size = new System.Drawing.Size(150, 30), Text = "尺寸递增：" };
            TrackBar sizeIncrementTrackBar = new TrackBar { Location = new System.Drawing.Point(200, yPosition), Size = new System.Drawing.Size(350, 45), Minimum = 0, Maximum = 100, Value = (int)sizeIncrement };
            yPosition += 100;

            Label copyCountLabel = null;
            TrackBar copyCountTrackBar = null;
            if (isCopyMode)
            {
                copyCountLabel = new Label { Location = new System.Drawing.Point(20, yPosition), Size = new System.Drawing.Size(150, 30), Text = "复制数量：" };
                copyCountTrackBar = new TrackBar { Location = new System.Drawing.Point(200, yPosition), Size = new System.Drawing.Size(350, 45), Minimum = 0, Maximum = 50, Value = copyCount };
                yPosition += 100;
            }

            Button resetButton = new Button { Location = new System.Drawing.Point(20, yPosition), Size = new System.Drawing.Size(150, 50), Text = "重置大小" };
            yPosition += 30;

            void UpdateShapes()
            {
                radius = radiusTrackBar.Value;
                initialRotation = (float)initialRotationUpDown.Value;
                finalRotation = (float)finalRotationUpDown.Value;
                sizeIncrement = sizeIncrementTrackBar.Value;
                if (isCopyMode) copyCount = copyCountTrackBar.Value;

                PerformCircularDistribution(pptApp, shapes, radius, initialRotation, finalRotation, sizeIncrement, isCopyMode, copyCount);
            }

            void ResetParameters()
            {
                radiusTrackBar.Value = 100;
                initialRotationUpDown.Value = 0;
                finalRotationUpDown.Value = 0;
                sizeIncrementTrackBar.Value = 0;
                if (isCopyMode) copyCountTrackBar.Value = 0;

                foreach (PowerPoint.Shape shape in shapes)
                {
                    if (initialSizes.ContainsKey(shape.Id))
                    {
                        var size = initialSizes[shape.Id];
                        shape.Width = size.Width;
                        shape.Height = size.Height;
                    }
                }

                UpdateShapes();
            }

            radiusTrackBar.ValueChanged += (s, ev) => UpdateShapes();
            initialRotationUpDown.ValueChanged += (s, ev) => UpdateShapes();
            finalRotationUpDown.ValueChanged += (s, ev) => UpdateShapes();
            sizeIncrementTrackBar.ValueChanged += (s, ev) => UpdateShapes();
            if (isCopyMode) copyCountTrackBar.ValueChanged += (s, ev) => UpdateShapes();
            resetButton.Click += (s, ev) => ResetParameters();

            form.Controls.Add(radiusLabel);
            form.Controls.Add(radiusTrackBar);
            form.Controls.Add(initialRotationLabel);
            form.Controls.Add(initialRotationUpDown);
            form.Controls.Add(finalRotationLabel);
            form.Controls.Add(finalRotationUpDown);
            form.Controls.Add(sizeIncrementLabel);
            form.Controls.Add(sizeIncrementTrackBar);
            if (isCopyMode)
            {
                form.Controls.Add(copyCountLabel);
                form.Controls.Add(copyCountTrackBar);
            }
            form.Controls.Add(resetButton);

            UpdateShapes();
            form.ShowDialog();
        }

        


        private void button17_Click(object sender, RibbonControlEventArgs e)
            {
            var pptApp = Globals.ThisAddIn.Application;
            var slide = pptApp.ActiveWindow.View.Slide;
            var selection = pptApp.ActiveWindow.Selection;

            if (selection.Type == PpSelectionType.ppSelectionShapes &&
                selection.ShapeRange.Count == 1 &&
                selection.ShapeRange[1].Type == Microsoft.Office.Core.MsoShapeType.msoPicture)
            {
                var shape = selection.ShapeRange[1];
                var tempImagePath = Path.Combine(Path.GetTempPath(), "temp_image_" + Guid.NewGuid().ToString() + ".png");

                try
                {
                    // 使用文件路径直接导出图片
                    shape.Export(tempImagePath, PpShapeFormat.ppShapeFormatPNG);

                    using (var img = System.Drawing.Image.FromFile(tempImagePath))
                    {
                        using (var form = new BackgroundRemovalForm(img, pptApp, slide, tempImagePath))
                        {
                            form.ShowDialog();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导出图片失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("请选择一张图片进行操作", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        
        private void button18_Click(object sender, RibbonControlEventArgs e)
        {
            string tempPath = Path.GetTempPath(); // 获取系统临时文件夹路径
            string[] targetExtensions = { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tmp" }; // 常见的图片文件扩展名和TMP文件扩展名

            try
            {
                // 获取临时文件夹中的所有文件（仅一级目录）
                string[] tempFiles = Directory.GetFiles(tempPath);

                // 遍历所有文件并尝试删除匹配图片和TMP文件扩展名的文件
                foreach (string file in tempFiles)
                {
                    string extension = Path.GetExtension(file).ToLower();
                    if (targetExtensions.Contains(extension))
                    {
                        try
                        {
                            File.Delete(file);
                        }
                        catch (IOException ex) when ((ex.HResult & 0xFFFF) == 32) // 文件被占用
                        {
                            // 忽略文件被占用的异常，不提醒用户
                        }
                        catch (Exception ex)
                        {
                            // 其他异常情况，记录错误信息
                            MessageBox.Show($"无法删除文件: {file}\n错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }

                MessageBox.Show("所有临时图片文件和可删除的TMP文件已删除", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // 如果获取文件列表或删除文件过程中发生错误，记录错误信息
                MessageBox.Show($"删除临时图片文件和TMP文件时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void replacetextbutton_Click(object sender, RibbonControlEventArgs e)
        {
            string replacementText = GetUserInput();
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

        private string GetUserInput()
        {
            using (var form = new Form())
            {
                var textBox = new TextBox();
                var okButton = new Button();

                form.Text = "批量换字";
                textBox.Dock = DockStyle.Top;
                textBox.Height = 90;
              
                okButton.Height = 40;
                okButton.Dock = DockStyle.Bottom;
                okButton.Text = "确定";


                okButton.Click += (sender, e) => form.Close();
                form.Size = new System.Drawing.Size(500, 200);
                form.Controls.Add(textBox);
                form.Controls.Add(okButton);

                form.ShowDialog();

                return textBox.Text;
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



        private void Masking_Click(object sender, RibbonControlEventArgs e)
        {
            TransparencyForm transparencyForm = new TransparencyForm();
            transparencyForm.ShowDialog();
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

        private void Insertwebpage_Click(object sender, RibbonControlEventArgs e)
        {
            // 显示输入网址的窗口
            WebpageInputForm inputForm = new WebpageInputForm();
            DialogResult result = inputForm.ShowDialog();

            // 如果用户点击了嵌入按钮
            if (result == DialogResult.OK)
            {
                string url = inputForm.WebpageUrl;

                // 获取当前活动窗口
                PowerPoint.DocumentWindow activeWindow = Globals.ThisAddIn.Application.ActiveWindow;
                if (activeWindow != null)
                {
                    // 获取当前页幻灯片
                    PowerPoint.Slide currentSlide = activeWindow.View.Slide;

                    // 在当前页幻灯片中嵌入网页
                    PowerPoint.Shape LiveWebShape = currentSlide.Shapes.AddOLEObject(
                        Left: 100, // 可以根据需要调整位置
                        Top: 100, // 可以根据需要调整位置
                        Width: 600, // 可以根据需要调整大小
                        Height: 400, // 可以根据需要调整大小
                        ClassName: "Shell.Explorer",
                        FileName: "",
                        DisplayAsIcon: Office.MsoTriState.msoFalse
                    );

                    // 设置WebBrowser控件的网址
                    dynamic webBrowser = LiveWebShape.OLEFormat.Object;
                    webBrowser.Navigate(url);
                }
            }
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




        private void Supercopy_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前活动的PowerPoint应用程序
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;

            // 获取当前选中的幻灯片
            PowerPoint.Slide currentSlide = pptApp.ActiveWindow.View.Slide;

            // 获取当前选中的对象
            PowerPoint.Selection selection = pptApp.ActiveWindow.Selection;

            // 检查是否按住了Ctrl键
            bool ctrlPressed = (System.Windows.Forms.Control.ModifierKeys & System.Windows.Forms.Keys.Control) == System.Windows.Forms.Keys.Control;

            // 如果按住了Ctrl键
            if (ctrlPressed)
            {
                // 提示用户输入要复制的次数
                string input = Microsoft.VisualBasic.Interaction.InputBox("请输入要复制的次数:", "批量原位复制", "1");
                int copyCount;
                // 解析用户输入的次数
                if (!int.TryParse(input, out copyCount) || copyCount < 1)
                {
                    System.Windows.Forms.MessageBox.Show("请输入一个大于0的整数。", "错误", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    return;
                }

                // 复制选中的对象多次
                for (int i = 0; i < copyCount; i++)
                {
                    DuplicateSelectedShapes(selection);
                }
            }
            else // 如果没有按住Ctrl键
            {
                // 单次复制选中的对象
                DuplicateSelectedShapes(selection);
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

        private void Sizescaling_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前活动的PowerPoint应用程序
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;

            // 获取当前选中的对象
            PowerPoint.Selection selection = pptApp.ActiveWindow.Selection;

            // 确保至少选中了一个对象
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionNone)
            {
                // 解析用户输入的缩放比例
                string input = Microsoft.VisualBasic.Interaction.InputBox("请输入缩放比例（单位：%），等差缩放比例请用逗号分隔两个数值：", "尺寸缩放", "");
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
                       
                        return;
                    }

                    // 计算等差缩放的公差
                    commonDifference = (endScale - startScale) / (selection.ShapeRange.Count - 1);
                }

                // 记录当前缩放比例
                float currentScale = isArithmetic ? float.Parse(scaleValues[0]) : float.Parse(scaleValues[0]);

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

        private void button21_Click_1(object sender, RibbonControlEventArgs e)
        {
            // 获取当前活动的PPT应用程序
            Application pptApplication = Globals.ThisAddIn.Application;
            // 获取当前活动的窗口
            DocumentWindow activeWindow = pptApplication.ActiveWindow;
            // 获取当前选中的对象
            Selection selection = activeWindow.Selection;

            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                // 弹出输入框，让用户输入命名前缀
                string prefix = Microsoft.VisualBasic.Interaction.InputBox("请输入命名前缀:", "批量重命名");

                if (!string.IsNullOrEmpty(prefix))
                {
                    int counter = 1;
                    foreach (Shape shape in selection.ShapeRange)
                    {
                        shape.Name = $"{prefix}-{counter}";
                        counter++;
                    }

                    // 刷新视图
                    activeWindow.View.GotoSlide(activeWindow.View.Slide.SlideIndex);
                }
                else
                {
                    MessageBox.Show("命名前缀不能为空。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("请选择一个或多个对象。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
    }
}































