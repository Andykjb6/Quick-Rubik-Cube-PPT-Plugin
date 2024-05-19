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






namespace 课件帮PPT助手
{
    public partial class Ribbon1
    {
        private CustomCloudTextGenerator cloudTextGenerator;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            cloudTextGenerator = new CustomCloudTextGenerator();
        }

        private void button5_Click_1(object sender, RibbonControlEventArgs e)
        {
            cloudTextGenerator.InitializeForm();
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
            float newShapeTop = originShape.Top - 20; // 新文本框放置在原文本框顶部20点的位置
            if (newShapeTop < 0) newShapeTop = originShape.Top + originShape.Height; // 如果超出幻灯片顶部，则放在下方

            Shape pinyinShape = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                originShape.Left,
                newShapeTop,
                originShape.Width,
                20);
            pinyinShape.TextFrame.TextRange.Text = pinyin;
            pinyinShape.TextFrame.TextRange.Font.Size = 16; // 调整字体大小
            pinyinShape.TextFrame.TextRange.Font.Name = "Arial"; // 设置字体，确保支持拼音符号

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
            string[] pinyinLibrary = { "ā", "á", "ǎ", "à", "āi", "ái", "ǎi", "ài", "ān", "án", "ǎn", "àn", "āng", "áng", "ǎng", "àng", "āo", "áo", "ǎo", "ào", "bā", "bá", "bǎ", "bà", "bāi", "bái", "bǎi", "bài", "bān", "bán", "bǎn", "bàn", "bāng", "báng", "bǎng", "bàng", "bāo", "báo", "bǎo", "bào", "bēi", "béi", "běi", "bèi", "bēn", "bén", "běn", "bèn", "bēng", "béng", "běng", "bèng", "bī", "bí", "bǐ", "bì", "biān", "bián", "biǎn", "biàn", "biāo", "biáo", "biǎo", "biào", "biē", "bié", "biě", "biè", "bīn", "bín", "bǐn", "bìn", "bīng", "bíng", "bǐng", "bìng", "bō", "bó", "bǒ", "bò", "bū", "bú", "bǔ", "bù", "cā", "cá", "cǎ", "cà", "chā", "chá", "chǎ", "chà", "cāi", "cái", "cǎi", "cài", "chāi", "chái", "chǎi", "chài", "cān", "cán", "cǎn", "càn", "chān", "chán", "chǎn", "chàn", "cāng", "cáng", "cǎng", "càng", "chāng", "cháng", "chǎng", "chàng", "cāo", "cáo", "cǎo", "cào", "chāo", "cháo", "chǎo", "chào", "cē", "cé", "cě", "cè", "chē", "ché", "chě", "chè", "cēn", "cén", "cěn", "cèn", "chēn", "chén", "chěn", "chèn", "cēng", "céng", "cěng", "cèng", "chēng", "chéng", "chěng", "chèng", "cī", "cí", "cǐ", "cì", "chī", "chí", "chǐ", "chì", "cōng", "cóng", "cǒng", "còng", "chōng", "chóng", "chǒng", "chòng", "cōu", "cóu", "cǒu", "còu", "chōu", "chóu", "chǒu", "chòu", "cū", "cú", "cǔ", "cù", "chū", "chú", "chǔ", "chù", "chuā", "chuá", "chuǎ", "chuà", "chuāi", "chuái", "chuǎi", "chuài", "chuān", "chuán", "chuǎn", "chuàn", "chuāng", "chuáng", "chuǎng", "chuàng", "chuī", "chuí", "chuǐ", "chuì", "chūn", "chún", "chǔn", "chùn", "chuō", "chuó", "chuǒ", "chuò", "cōu", "coū", "cóu", "cǒu", "còu", "cuān", "cuán", "cuǎn", "cuàn", "cuī", "cuí", "cuǐ", "cuì", "cūn", "cún", "cǔn", "cùn", "cuō", "cuó", "cuǒ", "cuò", "dā", "dá", "dǎ", "dà", "dāi", "dái", "dǎi", "dài", "dān", "dán", "dǎn", "dàn", "dāng", "dáng", "dǎng", "dàng", "dāo", "dáo", "dǎo", "dào", "dē", "dé", "dě", "dè", "dēi", "déi", "děi", "dèi", "dēn", "dén", "děn", "dèn", "dēng", "déng", "děng", "dèng", "dī", "dí", "dǐ", "dì", "diān", "dián", "diǎn", "diàn", "diāo", "diáo", "diǎo", "diào", "diē", "dié", "diě", "diè", "dīng", "díng", "dǐng", "dìng", "diū", "diú", "diǔ", "diù", "dōng", "dóng", "dǒng", "dòng", "dōu", "dóu", "dǒu", "dòu", "dū", "dú", "dǔ", "dù", "duān", "duán", "duǎn", "duàn", "duī", "duí", "duǐ", "duì", "dūn", "dún", "dǔn", "dùn", "duō", "duó", "duǒ", "duò", "ē", "é", "ě", "è", "ēi", "éi", "ěi", "èi", "ēn", "én", "ěn", "èn", "ēng", "éng", "ěng", "èng", "ér", "èr", "èr", "ēr", "fā", "fá", "fǎ", "fà", "fān", "fán", "fǎn", "fàn", "fāng", "fáng", "fǎng", "fàng", "fāo", "fáo", "fǎo", "fào", "fēi", "féi", "fěi", "fèi", "fēn", "fén", "fěn", "fèn", "fēng", "féng", "fěng", "fèng", "fō", "fó", "fǒ", "fò", "fōu", "foū", "fóu", "fǒu", "fòu", "fū", "fú", "fǔ", "fù", "gā", "gá", "gǎ", "gà", "gāi", "gái", "gǎi", "gài", "gān", "gán", "gǎn", "gàn", "gāng", "gáng", "gǎng", "gàng", "gāo", "gáo", "gǎo", "gào", "gē", "gé", "gě", "gè", "gēi", "géi", "gěi", "gèi", "gēn", "gén", "gěn", "gèn", "gēng", "géng", "gěng", "gèng", "gōng", "góng", "gǒng", "gòng", "gōu", "góu", "gǒu", "gòu", "gū", "gú", "gǔ", "gù", "guā", "guá", "guǎ", "guà", "guāi", "guái", "guǎi", "guài", "guān", "guán", "guǎn", "guàn", "guāng", "guáng", "guǎng", "guàng", "guī", "guí", "guǐ", "guì", "gūn", "gún", "gǔn", "gùn", "guō", "guó", "guǒ", "guò", "hā", "há", "hǎ", "hà", "hāi", "hái", "hǎi", "hài", "hān", "hán", "hǎn", "hàn", "hāng", "háng", "hǎng", "hàng", "hāo", "háo", "hǎo", "hào", "hē", "hé", "hě", "hè", "hēn", "hén", "hěn", "hèn", "hēng", "héng", "hěng", "hèng", "hōng", "hóng", "hǒng", "hòng", "hōu", "hóu", "hǒu", "hòu", "hū", "hú", "hǔ", "hù", "huā", "huá", "huǎ", "huà", "huāi", "huái", "huǎi", "huài", "huān", "huán", "huǎn", "huàn", "huāng", "huáng", "huǎng", "huàng", "huī", "huí", "huǐ", "huì", "hūn", "hún", "hǔn", "hùn", "huō", "huó", "huǒ", "huò", "jī", "jí", "jǐ", "jì", "jiā", "jiá", "jiǎ", "jià", "jiān", "jián", "jiǎn", "jiàn", "jiāng", "jiáng", "jiǎng", "jiàng", "jiāo", "jiáo", "jiǎo", "jiào", "jiē", "jié", "jiě", "jiè", "jiu", "jiū", "jiú", "jiǔ", "jiù", "jū", "jú", "jǔ", "jù", "juān", "juán", "juǎn", "juàn", "juē", "jué", "juě", "juè", "jūn", "jún", "jǔn", "jùn", "kā", "ká", "kǎ", "kà", "kāi", "kái", "kǎi", "kài", "kān", "kán", "kǎn", "kàn", "kāng", "káng", "kǎng", "kàng", "kāo", "káo", "kǎo", "kào", "kē", "ké", "kě", "kè", "kēn", "kén", "kěn", "kèn", "kēng", "kéng", "kěng", "kèng", "kōng", "kóng", "kǒng", "kòng", "kōu", "kóu", "kǒu", "kòu", "kū", "kú", "kǔ", "kù", "kuā", "kuá", "kuǎ", "kuà", "kuāi", "kuái", "kuǎi", "kuài", "kuān", "kuán", "kuǎn", "kuàn", "kuāng", "kuáng", "kuǎng", "kuàng", "kuī", "kuí", "kuǐ", "kuì", "kūn", "kún", "kǔn", "kùn", "kuō", "kuó", "kuǒ", "kuò", "lā", "lá", "lǎ", "là", "lāi", "lái", "lǎi", "lài", "lān", "lán", "lǎn", "làn", "lāng", "láng", "lǎng", "làng", "lāo", "láo", "lǎo", "lào", "lē", "lé", "lě", "lè", "lēi", "léi", "lěi", "lèi", "lēn", "lén", "lěn", "lèn", "lēng", "léng", "lěng", "lèng", "lī", "lí", "lǐ", "lì", "lōng", "lóng", "lǒng", "lòng", "lōu", "lóu", "lǒu", "lòu", "lū", "lú", "lǔ", "lù", "luān", "luán", "luǎn", "luàn", "luē", "lué", "luě", "luè", "luō", "luó", "luǒ", "luò", "lun", "lūn", "lún", "lǔn", "lùn", "luōng", "luóng", "luǒng", "luòng", "mā", "má", "mǎ", "mà", "māi", "mái", "mǎi", "mài", "mān", "mán", "mǎn", "màn", "māng", "máng", "mǎng", "màng", "māo", "máo", "mǎo", "mào", "mē", "mé", "mě", "mè", "mēi", "méi", "měi", "mèi", "mēn", "mén", "měn", "mèn", "mēng", "méng", "měng", "mèng", "mī", "mí", "mǐ", "mì", "mō", "mó", "mǒ", "mò", "mōng", "móng", "mǒng", "mòng", "mōu", "móu", "mǒu", "mòu", "mū", "mú", "mǔ", "mù", "nuān", "nuán", "nuǎn", "nuàn", "nū", "nú", "nǔ", "nù", "nuē", "nué", "nuě", "nuè", "nuē", "nué", "nuě", "nuè", "nū", "nú", "nǔ", "nù", "nuān", "nuán", "nuǎn", "nuàn", "nǔ", "nù", "nuē", "nué", "nuě", "nuè", "nā", "ná", "nǎ", "nà", "nāi", "nái", "nǎi", "nài", "nān", "nán", "nǎn", "nàn", "nāng", "náng", "nǎng", "nàng", "nāo", "náo", "nǎo", "nào", "nē", "né", "ně", "nè", "nēi", "néi", "něi", "nèi", "nēn", "nén", "něn", "nèn", "nēng", "néng", "něng", "nèng", "nī", "ní", "nǐ", "nì", "nīng", "níng", "nǐng", "nìng", "niān", "nián", "niǎn", "niàn", "niāo", "niáo", "niǎo", "niào", "niē", "nié", "niě", "niè", "nín", "nín", "nǐn", "nìn", "nīng", "níng", "nǐng", "nìng", "niū", "niú", "niǔ", "niù", "nōng", "nóng", "nǒng", "nòng", "nóu", "nǒu", "nòu", "nū", "nú", "nǔ", "nù", "nuō", "nuó", "nuǒ", "nuò", "o", "ō", "ó", "ǒ", "ò", "ōng", "óng", "ǒng", "òng", "ōu", "óu", "ǒu", "òu", "pā", "pá", "pǎ", "pà", "pāi", "pái", "pǎi", "pài", "pān", "pán", "pǎn", "pàn", "pāng", "páng", "pǎng", "pàng", "pāo", "páo", "pǎo", "pào", "pēi", "péi", "pěi", "pèi", "pēn", "pén", "pěn", "pèn", "pēng", "péng", "pěng", "pèng", "pī", "pí", "pǐ", "pì", "pīn", "pín", "pǐn", "pìn", "pīng", "píng", "pǐng", "pìng", "pō", "pó", "pǒ", "pò", "pōng", "póng", "pǒng", "pòng", "pōu", "póu", "pǒu", "pòu", "pū", "pú", "pǔ", "pù", "qi", "qiā", "qiá", "qiǎ", "qià", "qiāi", "qiái", "qiǎi", "qiài", "qiān", "qián", "qiǎn", "qiàn", "qiāng", "qiáng", "qiǎng", "qiàng", "qiāo", "qiáo", "qiǎo", "qiào", "qiē", "qié", "qiě", "qiè", "qiū", "qiú", "qiǔ", "qiù", "qū", "qú", "qǔ", "qù", "quān", "quán", "quǎn", "quàn", "quē", "qué", "quě", "què", "qūn", "qún", "qǔn", "qùn", "ruān", "ruán", "ruǎn", "ruàn", "ruī", "ruí", "ruǐ", "ruì", "rū", "rú", "rǔ", "rù", "rān", "rán", "rǎn", "ràn", "rāng", "ráng", "rǎng", "ràng", "rāo", "ráo", "rǎo", "rào", "rē", "ré", "rě", "rè", "rēn", "rén", "rěn", "rèn", "rēng", "réng", "rěng", "rèng", "rī", "rí", "rǐ", "rì", "rōng", "róng", "rǒng", "ròng", "rōu", "róu", "rǒu", "ròu", "ruō", "ruó", "ruǒ", "ruò", "sā", "sá", "sǎ", "sà", "shā", "shá", "shǎ", "shà", "sāi", "sái", "sǎi", "sài", "shāi", "shái", "shǎi", "shài", "sān", "sán", "sǎn", "sàn", "shān", "shán", "shǎn", "shàn", "sāng", "sáng", "sǎng", "sàng", "shāng", "sháng", "shǎng", "shàng", "sāo", "sáo", "sǎo", "sào", "shāo", "sháo", "shǎo", "shào", "sē", "sé", "sě", "sè", "shē", "shé", "shě", "shè", "sēn", "sén", "sěn", "sèn", "shēn", "shén", "shěn", "shèn", "sēng", "séng", "sěng", "sèng", "shēng", "shéng", "shěng", "shèng", "sī", "sí", "sǐ", "sì", "shī", "shí", "shǐ", "shì", "sōng", "sóng", "sǒng", "sòng", "sōu", "sóu", "sǒu", "sòu", "shōu", "shóu", "shǒu", "shòu", "sū", "sú", "sǔ", "sù", "shū", "shú", "shǔ", "shù", "suān", "suán", "suǎn", "suàn", "shuān", "shuán", "shuǎn", "shuàn", "suī", "suí", "suǐ", "suì", "shuī", "shuí", "shuǐ", "shuì", "suō", "suó", "suǒ", "suò", "shuō", "shuó", "shuǒ", "shuò", "tā", "tá", "tǎ", "tà", "tāi", "tái", "tǎi", "tài", "tān", "tán", "tǎn", "tàn", "tāng", "táng", "tǎng", "tàng", "tāo", "táo", "tǎo", "tào", "tē", "té", "tě", "tè", "tēn", "tén", "těn", "tèn", "tēng", "téng", "těng", "tèng", "tī", "tí", "tǐ", "tì", "tiān", "tián", "tiǎn", "tiàn", "tōng", "tóng", "tǒng", "tòng", "tōu", "tóu", "tǒu", "tòu", "tū", "tú", "tǔ", "tù", "tuān", "tuán", "tuǎn", "tuàn", "tuē", "tué", "tuě", "tuè", "tuō", "tuó", "tuǒ", "tuò", "tuī", "tuí", "tuǐ", "tuì", "tūn", "tún", "tǔn", "tùn", "tuō", "tuó", "tuǒ", "tuò", "wā", "wá", "wǎ", "wà", "wō", "wó", "wǒ", "wò", "wāi", "wái", "wǎi", "wài", "wān", "wán", "wǎn", "wàn", "wāng", "wáng", "wǎng", "wàng", "wēi", "wéi", "wěi", "wèi", "wēn", "wén", "wěn", "wèn", "wēng", "wéng", "wěng", "wū", "wú", "wǔ", "wù", "xī", "xí", "xǐ", "xì", "xiān", "xián", "xiǎn", "xiàn", "xiāng", "xiáng", "xiǎng", "xiàng", "xiāo", "xiáo", "xiǎo", "xiào", "xiē", "xié", "xiě", "xiè", "xiū", "xiú", "xiǔ", "xiù", "xiāng", "xiáng", "xiǎng", "xiàng", "xū", "xú", "xǔ", "xù", "xuān", "xuán", "xuǎn", "xuàn", "xuē", "xué", "xuě", "xuè", "yā", "yá", "yǎ", "yà", "yān", "yán", "yǎn", "yàn", "yuān", "yuán", "yuǎn", "yuàn", "yāng", "yáng", "yǎng", "yàng", "yāo", "yáo", "yǎo", "yào", "yē", "yé", "yě", "yè", "yī", "yí", "yǐ", "yì", "yīn", "yín", "yǐn", "yìn", "yīng", "yíng", "yǐng", "yìng", "yō", "yó", "yǒ", "yò", "yōng", "yóng", "yǒng", "yòng", "yōu", "yóu", "yǒu", "yòu", "yū", "yú", "yǔ", "yù", "yuān", "yuán", "yuǎn", "yuàn", "yuē", "yué", "yuě", "yuè", "zā", "zá", "zǎ", "zà", "zhā", "zhá", "zhǎ", "zhà", "zāi", "zái", "zǎi", "zài", "zhāi", "zhái", "zhǎi", "zhài", "zān", "zán", "zǎn", "zàn", "zhān", "zhán", "zhǎn", "zhàn", "zāng", "záng", "zǎng", "zàng", "zhāng", "zháng", "zhǎng", "zhàng", "zāo", "záo", "zǎo", "zào", "zhāo", "zháo", "zhǎo", "zhào", "zē", "zé", "zě", "zè", "zhē", "zhé", "zhě", "zhè", "zēn", "zén", "zěn", "zèn", "zhēn", "zhén", "zhěn", "zhèn", "zēng", "zéng", "zěng", "zèng", "zhēng", "zhéng", "zhěng", "zhèng", "zī", "zí", "zǐ", "zì", "zhī", "zhí", "zhǐ", "zhì", "zōng", "zóng", "zǒng", "zòng", "zhōng", "zhóng", "zhǒng", "zhòng", "zōu", "zóu", "zǒu", "zòu", "zhōu", "zhóu", "zhǒu", "zhòu", "zū", "zú", "zǔ", "zù", "zhū", "zhú", "zhǔ", "zhù", "zuān", "zuán", "zuǎn", "zuàn", "zhuān", "zhuán", "zhuǎn", "zhuàn", "zuī", "zuí", "zuǐ", "zuì", "zhuī", "zhuí", "zhuǐ", "zhuì", "zuō", "zuó", "zuǒ", "zhuò", "zhuō", "zhuó", "zhuǒ", "zhuò", };
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

        private static string GetResourceText(string resourceName)
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
            string inputChar = Microsoft.VisualBasic.Interaction.InputBox("请输入目标汉字:", "输入目标汉字", "");
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

                Label labelRows = new Label { Text = "行数:", Location = new System.Drawing.Point(65, 40), AutoSize = true };
                NumericUpDown numericUpDownRows = new NumericUpDown { Location = new System.Drawing.Point(190, 40), Minimum = 1, Maximum = 10, Value = 2 };
                numericUpDownRows.ValueChanged += ApplySettings;

                Label labelColumns = new Label { Text = "列数:", Location = new System.Drawing.Point(65, 80), AutoSize = true };
                NumericUpDown numericUpDownColumns = new NumericUpDown { Location = new System.Drawing.Point(190, 80), Minimum = 1, Maximum = 10, Value = 2 };
                numericUpDownColumns.ValueChanged += ApplySettings;

                Label labelRowSpacing = new Label { Text = "行间距:", Location = new System.Drawing.Point(65, 120), AutoSize = true };
                NumericUpDown numericUpDownRowSpacing = new NumericUpDown { Location = new System.Drawing.Point(190, 120), Minimum = 0, Maximum = 100, Value = 10 };
                numericUpDownRowSpacing.ValueChanged += ApplySettings;

                Label labelColumnSpacing = new Label { Text = "列间距:", Location = new System.Drawing.Point(65, 160), AutoSize = true };
                NumericUpDown numericUpDownColumnSpacing = new NumericUpDown { Location = new System.Drawing.Point(190, 160), Minimum = 0, Maximum = 100, Value = 10 };
                numericUpDownColumnSpacing.ValueChanged += ApplySettings;

                Label labelWidth = new Label { Text = "边框宽度:", Location = new System.Drawing.Point(65, 200), AutoSize = true };
                NumericUpDown numericUpDownBorderWidth = new NumericUpDown { Location = new System.Drawing.Point(190, 200), Minimum = 0, Maximum = 10, DecimalPlaces = 2, Value = 1.25m };
                numericUpDownBorderWidth.ValueChanged += ApplySettings;

                Label labelScale = new Label { Text = "缩放比例:", Location = new System.Drawing.Point(65, 300), AutoSize = true };
                TrackBar trackBarScale = new TrackBar { Location = new System.Drawing.Point(190, 300), Minimum = 50, Maximum = 200, Value = 100, TickFrequency = 10, Width = 120 };
                trackBarScale.ValueChanged += ApplySettings;

                Label labelColor = new Label { Text = "边框颜色:", Location = new System.Drawing.Point(65, 240), AutoSize = true };
                Button buttonChooseColor = new Button { Text = "自定义", Location = new System.Drawing.Point(190, 240), Size = new Size(120, 40) };

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

                Button buttonOK = new Button { Text = "生成", Location = new System.Drawing.Point(65, 350), Size = new Size(100, 40) };

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

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    float left = startX + j * (squareSize + columnSpacing);
                    float top = startY + i * (squareSize + rowSpacing);

                    PowerPoint.Shape tableShape = activeSlide.Shapes.AddTable(2, 2, left, top, squareSize, squareSize);
                    tableShape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue; // 锁定纵横比

                    PowerPoint.Table table = tableShape.Table;

                    SetTableProperties(table, borderWidth, borderColor);
                }
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

            float startX = 100; // 初始X位置，可以根据需要调整
            float startY = 100; // 初始Y位置，可以根据需要调整
            float squareSize = 100 * scale; // 根据缩放比例调整每个正方形表格的大小

            int shapeIndex = 0;

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    if (shapeIndex < activeSlide.Shapes.Count)
                    {
                        PowerPoint.Shape tableShape = activeSlide.Shapes[shapeIndex + 1];
                        float left = startX + j * (squareSize + columnSpacing);
                        float top = startY + i * (squareSize + rowSpacing);

                        tableShape.Left = left;
                        tableShape.Top = top;
                        tableShape.Width = squareSize;
                        tableShape.Height = squareSize;

                        if (tableShape.Type == Microsoft.Office.Core.MsoShapeType.msoTable)
                        {
                            PowerPoint.Table table = tableShape.Table;
                            SetTableProperties(table, borderWidth, borderColor);
                        }

                        shapeIndex++;
                    }
                }
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

                    // 外部边框设置为实线
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

                    // 内部边框设置为虚线
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
            border.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            border.DashStyle = isOuterCell ? Microsoft.Office.Core.MsoLineDashStyle.msoLineSolid : Microsoft.Office.Core.MsoLineDashStyle.msoLineDash;
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

                    tableShapeButton12.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendBackward);
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
    }
}












