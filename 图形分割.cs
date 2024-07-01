using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace 课件帮PPT助手
{
    public partial class 图形分割 : Form
    {
        private PowerPoint.Application pptApp;

        public 图形分割()
        {
            InitializeComponent();
            comboBox1.SelectedIndex = 0; // 默认选择“圆形分割”
            pptApp = Globals.ThisAddIn.Application;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 如果不需要在Label显示当前选择的文本，可以注释掉以下行
            // label1.Text = comboBox1.SelectedItem.ToString();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string selectedType = comboBox1.SelectedItem.ToString();
                string input = textBox1.Text;

                PowerPoint.Selection selection = pptApp.ActiveWindow.Selection;

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count == 1)
                {
                    PowerPoint.Shape shape = selection.ShapeRange[1];
                    if (selectedType == "圆形分割" && shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeOval)
                    {
                        if (int.TryParse(input, out int numSectors) && numSectors > 0)
                        {
                            DivideCircle(shape, numSectors);
                        }
                        else
                        {
                            MessageBox.Show("请输入一个有效的正整数。");
                        }
                    }
                    else if (selectedType == "矩形分割")
                    {
                        ShpSplit(input);
                    }
                    else
                    {
                        MessageBox.Show("请选择一个圆形进行圆形分割。");
                    }
                }
                else
                {
                    MessageBox.Show("请选择一个形状进行分割。");
                }
            }
        }

        private void DivideCircle(PowerPoint.Shape circle, int numSectors)
        {
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Slide currentSlide = pptApp.ActiveWindow.View.Slide; // 获取当前活动的幻灯片

            float centerX = circle.Left + circle.Width / 2;
            float centerY = circle.Top + circle.Height / 2;
            float radius = circle.Width / 2;

            // 复制原始圆形的格式
            var fillColor = circle.Fill.ForeColor.RGB;
            var lineColor = circle.Line.ForeColor.RGB;
            var lineWeight = circle.Line.Weight;
            var lineStyle = circle.Line.DashStyle;
            var lineVisible = circle.Line.Visible;

            // 删除原始圆形
            circle.Delete();

            // 每个扇形的角度
            float angleIncrement = 360f / numSectors;

            for (int i = 0; i < numSectors; i++)
            {
                float startAngle = i * angleIncrement;
                float endAngle = startAngle + angleIncrement;

                // 创建扇形
                PowerPoint.Shape pieSlice = currentSlide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapePie,
                    centerX - radius,
                    centerY - radius,
                    radius * 2,
                    radius * 2
                );

                // 设置扇形的角度
                pieSlice.Adjustments[1] = startAngle;
                pieSlice.Adjustments[2] = endAngle;

                // 应用原始圆形的格式
                pieSlice.Fill.ForeColor.RGB = fillColor;
                pieSlice.Line.ForeColor.RGB = lineColor;
                pieSlice.Line.Weight = lineWeight;
                pieSlice.Line.DashStyle = lineStyle;
                pieSlice.Line.Visible = lineVisible;
            }
        }

        private void PicSplit(PowerPoint.Shape pic, int n1, int n2, float hp, float vp, int mode)
        {
            float rt = 0f;
            if (pic.Rotation != 0f)
            {
                rt = pic.Rotation;
                pic.Rotation = 0f;
            }
            float w = pic.Width / n2;
            float h = pic.Height / n1;
            float x = (pic.Width - w) / 2f;
            float y = (pic.Height - h) / 2f;
            PowerPoint.Selection sel = pptApp.ActiveWindow.Selection;
            sel.Unselect();

            for (int i = 0; i < n1; i++)
            {
                for (int j = 0; j < n2; j++)
                {
                    PowerPoint.Shape shape = pic.Duplicate()[1];
                    shape.PictureFormat.Crop.ShapeWidth = w;
                    shape.PictureFormat.Crop.ShapeHeight = h;
                    shape.Left = pic.Left + (w + hp) * j;
                    shape.Top = pic.Top + (h + vp) * i;
                    shape.PictureFormat.Crop.PictureOffsetX = x - w * j;
                    shape.PictureFormat.Crop.PictureOffsetY = y - h * i;
                    shape.Select(Office.MsoTriState.msoFalse);
                }
            }

            if (mode == 1 || rt != 0f)
            {
                PowerPoint.Shape shp = sel.ShapeRange.Group();
                if (rt != 0f)
                {
                    shp.Rotation = rt;
                }
                if (mode == 1)
                {
                    shp.Width = pic.Width;
                    shp.Height = pic.Height;
                    shp.Left = pic.Left;
                    shp.Top = pic.Top;
                }
                shp.Ungroup();
            }
            pic.Delete();
        }

        private void OtherSplit(PowerPoint.Slide slide, PowerPoint.Shape pic, int n1, int n2, float hp, float vp, int mode)
        {
            PowerPoint.Selection sel = pptApp.ActiveWindow.Selection;
            pic.Copy();
            PowerPoint.Shape npic = slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG, Office.MsoTriState.msoFalse, "", 0, "", Office.MsoTriState.msoFalse)[1];
            npic.Left = pic.Left + pic.Width / 2f - npic.Width / 2f;
            npic.Top = pic.Top + pic.Height / 2f - npic.Height / 2f;
            float w = npic.Width / n2;
            float h = npic.Height / n1;
            float x = (npic.Width - w) / 2f;
            float y = (npic.Height - h) / 2f;
            sel.Unselect();

            for (int i = 0; i < n1; i++)
            {
                for (int j = 0; j < n2; j++)
                {
                    PowerPoint.Shape shape = npic.Duplicate()[1];
                    shape.PictureFormat.Crop.ShapeWidth = w;
                    shape.PictureFormat.Crop.ShapeHeight = h;
                    shape.Left = npic.Left + (w + hp) * j;
                    shape.Top = npic.Top + (h + vp) * i;
                    shape.PictureFormat.Crop.PictureOffsetX = x - w * j;
                    shape.PictureFormat.Crop.PictureOffsetY = y - h * i;
                    shape.Select(Office.MsoTriState.msoFalse);
                }
            }

            if (mode == 1)
            {
                PowerPoint.Shape shape2 = sel.ShapeRange.Group();
                shape2.Width = npic.Width;
                shape2.Height = npic.Height;
                shape2.Left = npic.Left;
                shape2.Top = npic.Top;
                shape2.Ungroup();
            }
            npic.Delete();
            pic.Delete();
        }

        private void ShapeSplit(PowerPoint.Slide slide, PowerPoint.Shape pic, int n1, int n2, float hp, float vp, int mode)
        {
            PowerPoint.Selection sel = pptApp.ActiveWindow.Selection;
            List<PowerPoint.Shape> nshps = new List<PowerPoint.Shape>();
            float w = pic.Width / n2;
            float h = pic.Height / n1;
            float rt = 0f;
            if (pic.Rotation != 0f)
            {
                rt = pic.Rotation;
                pic.Rotation = 0f;
            }
            sel.Unselect();

            for (int i = 0; i < n1; i++)
            {
                for (int j = 0; j < n2; j++)
                {
                    PowerPoint.Shape shape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, pic.Left + w * j, pic.Top + h * i, w, h);
                    PowerPoint.Shape shape2 = pic.Duplicate()[1];
                    shape2.Left = pic.Left;
                    shape2.Top = pic.Top;
                    shape2.Select(Office.MsoTriState.msoTrue);
                    shape.Select(Office.MsoTriState.msoFalse);
                    pptApp.CommandBars.ExecuteMso("ShapesIntersect");
                    PowerPoint.Shape nshp = slide.Shapes[slide.Shapes.Count];
                    nshp.Left += hp * j;
                    nshp.Top += vp * i;
                    nshps.Add(nshp);
                }
            }

            if ((mode == 1 || rt != 0f) && nshps.Count > 1)
            {
                sel.Unselect();
                foreach (PowerPoint.Shape shape3 in nshps)
                {
                    shape3.Select(Office.MsoTriState.msoFalse);
                }
                PowerPoint.Shape shp = sel.ShapeRange.Group();
                if (mode == 1)
                {
                    shp.Width = pic.Width;
                    shp.Height = pic.Height;
                    shp.Left = pic.Left;
                    shp.Top = pic.Top;
                }
                if (rt != 0f)
                {
                    shp.Rotation = rt;
                }
                shp.Ungroup();
            }
            pic.Delete();
        }

        private void ShpSplit(string txt)
        {
            try
            {
                PowerPoint.Selection sel = pptApp.ActiveWindow.Selection;
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    PowerPoint.Slide slide = pptApp.ActiveWindow.View.Slide as PowerPoint.Slide;
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    int count = range.Count;
                    string[] arr = txt.Trim().Split(new char[] { ' ', '，', ',' }).ToArray();
                    int n = int.Parse(arr[0]);
                    int n2 = int.Parse(arr[1]);
                    float hp = 0f;
                    float vp = 0f;
                    int mode = 0;

                    if (arr.Length == 4)
                    {
                        hp = float.Parse(arr[2]) * 72f / 2.54f;
                        vp = float.Parse(arr[3]) * 72f / 2.54f;
                    }
                    else if (arr.Length == 5)
                    {
                        hp = float.Parse(arr[2]) * 72f / 2.54f;
                        vp = float.Parse(arr[3]) * 72f / 2.54f;
                        mode = int.Parse(arr[4]);
                    }

                    List<PowerPoint.Shape> oshapes = new List<PowerPoint.Shape>();
                    for (int i = 1; i <= count; i++)
                    {
                        oshapes.Add(range[i]);
                    }
                    sel.Unselect();

                    int ver = int.Parse(Globals.ThisAddIn.Application.Version.Split(new char[] { '.' }).ToArray()[0]);
                    foreach (PowerPoint.Shape pic in oshapes)
                    {
                        if (pic.Type == Office.MsoShapeType.msoAutoShape || pic.Type == Office.MsoShapeType.msoFreeform)
                        {
                            if (pic.Fill.Type == Office.MsoFillType.msoFillPicture)
                            {
                                if (pic.AutoShapeType == Office.MsoAutoShapeType.msoShapeRectangle)
                                {
                                    PicSplit(pic, n, n2, hp, vp, mode);
                                }
                                else
                                {
                                    OtherSplit(slide, pic, n, n2, hp, vp, mode);
                                }
                            }
                            else if (ver >= 10)
                            {
                                ShapeSplit(slide, pic, n, n2, hp, vp, mode);
                            }
                            else
                            {
                                OtherSplit(slide, pic, n, n2, hp, vp, mode);
                            }
                        }
                        else if (pic.Type == Office.MsoShapeType.msoPicture)
                        {
                            if (pic.AutoShapeType == Office.MsoAutoShapeType.msoShapeRectangle)
                            {
                                PicSplit(pic, n, n2, hp, vp, mode);
                            }
                            else
                            {
                                OtherSplit(slide, pic, n, n2, hp, vp, mode);
                            }
                        }
                        else
                        {
                            OtherSplit(slide, pic, n, n2, hp, vp, mode);
                        }
                    }
                }
            }
            catch 
            {
                MessageBox.Show("请输入正确的行列数，如“3，3”，表示将当前形状分割成3行3列。" );
            }
        }
    }
}
