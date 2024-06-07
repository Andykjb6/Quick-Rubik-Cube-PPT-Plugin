using System;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace 课件帮PPT助手
{
    public partial class AlignToolWindow : Form
    {
        private PowerPoint.Application app;

        public AlignToolWindow(PowerPoint.Application application)
        {
            InitializeComponent();
            app = application;
            this.TopMost = true; // 窗口始终在前
            this.FormClosing += new FormClosingEventHandler(AlignToolWindow_FormClosing); // 处理关闭事件
        }

        private void AlignToolWindow_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true; // 取消关闭操作
            this.Hide(); // 隐藏窗口而不是最小化
        }

        private void CenterAlignButton_Click(object sender, EventArgs e)
        {
            AlignShapes(Alignment.Center);
        }

        private void LeftAlignButton_Click(object sender, EventArgs e)
        {
            AlignShapes(Alignment.Left);
        }

        private void HorizontalCenterAlignButton_Click(object sender, EventArgs e)
        {
            AlignShapes(Alignment.HorizontalCenter);
        }

        private void RightAlignButton_Click(object sender, EventArgs e)
        {
            AlignShapes(Alignment.Right);
        }

        private void TopAlignButton_Click(object sender, EventArgs e)
        {
            AlignShapes(Alignment.Top);
        }

        private void VerticalCenterAlignButton_Click(object sender, EventArgs e)
        {
            AlignShapes(Alignment.VerticalCenter);
        }

        private void BottomAlignButton_Click(object sender, EventArgs e)
        {
            AlignShapes(Alignment.Bottom);
        }

        private enum Alignment
        {
            Center,
            Left,
            HorizontalCenter,
            Right,
            Top,
            VerticalCenter,
            Bottom
        }

        private void AlignShapes(Alignment alignment)
        {
            PowerPoint.Selection selection = app.ActiveWindow.Selection;

            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange.Count < 2)
            {
                MessageBox.Show("请至少选择两个对象进行匹配对齐。");
                return;
            }

            PowerPoint.ShapeRange selectedShapes = selection.ShapeRange;
            int count = selectedShapes.Count;
            int mid = count / 2;

            var referenceShapes = selectedShapes.Cast<PowerPoint.Shape>().Take(mid).ToList();
            var targetShapes = selectedShapes.Cast<PowerPoint.Shape>().Skip(mid).ToList();

            for (int i = 0; i < referenceShapes.Count; i++)
            {
                var referenceShape = referenceShapes[i];
                var targetShape = targetShapes[i];

                switch (alignment)
                {
                    case Alignment.Center:
                        targetShape.Left = referenceShape.Left + (referenceShape.Width - targetShape.Width) / 2;
                        targetShape.Top = referenceShape.Top + (referenceShape.Height - targetShape.Height) / 2;
                        break;
                    case Alignment.Left:
                        targetShape.Left = referenceShape.Left;
                        break;
                    case Alignment.HorizontalCenter:
                        targetShape.Left = referenceShape.Left + (referenceShape.Width - targetShape.Width) / 2;
                        break;
                    case Alignment.Right:
                        targetShape.Left = referenceShape.Left + referenceShape.Width - targetShape.Width;
                        break;
                    case Alignment.Top:
                        targetShape.Top = referenceShape.Top;
                        break;
                    case Alignment.VerticalCenter:
                        targetShape.Top = referenceShape.Top + (referenceShape.Height - targetShape.Height) / 2;
                        break;
                    case Alignment.Bottom:
                        targetShape.Top = referenceShape.Top + referenceShape.Height - targetShape.Height;
                        break;
                }
            }
        }
    }
}
