using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace 课件帮PPT助手
{
    public partial class SpecifyalignmentForm : Form
    {
        private PowerPoint.Application app;

        public SpecifyalignmentForm(PowerPoint.Application application)
        {
            InitializeComponent();
            app = application;
            this.TopMost = true; // 窗口始终在前
            this.FormClosing += new FormClosingEventHandler(SpecifyalignmentForm_FormClosing); // 处理关闭事件
        }

        private void SpecifyalignmentForm_FormClosing(object sender, FormClosingEventArgs e)
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
            List<PowerPoint.Shape> referenceShapes = new List<PowerPoint.Shape>();
            List<PowerPoint.Shape> targetShapes = new List<PowerPoint.Shape>();

            // 添加前缀
            for (int i = 0; i < selectedShapes.Count; i++)
            {
                PowerPoint.Shape shape = selectedShapes[i + 1]; // ShapeRange是1-based索引
                if (i % 2 == 0)
                {
                    referenceShapes.Add(shape);
                    shape.Name = $"参考{(i / 2) + 1}-{shape.Name}";
                }
                else
                {
                    targetShapes.Add(shape);
                    shape.Name = $"目标{(i / 2) + 1}-{shape.Name}";
                }
            }

            // 对齐目标对象到参考对象
            for (int i = 0; i < referenceShapes.Count && i < targetShapes.Count; i++)
            {
                AlignShape(referenceShapes[i], targetShapes[i], alignment);
            }

            // 移除前缀
            foreach (PowerPoint.Shape shape in selectedShapes)
            {
                int prefixIndex = shape.Name.IndexOf('-');
                if (prefixIndex != -1)
                {
                    shape.Name = shape.Name.Substring(prefixIndex + 1);
                }
            }
        }

        private void AlignShape(PowerPoint.Shape referenceShape, PowerPoint.Shape targetShape, Alignment alignment)
        {
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
