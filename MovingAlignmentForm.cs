using System;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace 课件帮PPT助手
{
    public partial class MovingAlignmentForm : Form
    {
        private PowerPoint.Application app;

        public MovingAlignmentForm(PowerPoint.Application application)
        {
            InitializeComponent();
            this.app = application;
        }

        private void AlignShapes(Action<PowerPoint.ShapeRange, float> alignAction)
        {
            PowerPoint.Selection selection = app.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 1)
            {
                PowerPoint.ShapeRange selectedShapes = selection.ShapeRange;

                // Reference shape is the first selected shape
                PowerPoint.Shape referenceShape = selectedShapes[1];
                float referencePosition = 0;

                if (alignAction == AlignHorizontally)
                {
                    referencePosition = referenceShape.Left + (referenceShape.Width / 2);
                }
                else if (alignAction == AlignVertically)
                {
                    referencePosition = referenceShape.Top + (referenceShape.Height / 2);
                }
                else
                {
                    referencePosition = 0; // Default case should not be hit
                }

                alignAction(selectedShapes, referencePosition);
            }
            else
            {
                MessageBox.Show("请至少选择两个对象进行对齐！");
            }
        }

        private void AlignHorizontally(PowerPoint.ShapeRange selectedShapes, float referenceCenterX)
        {
            float totalWidth = 0;
            float leftMost = float.MaxValue;
            float rightMost = float.MinValue;

            for (int i = 2; i <= selectedShapes.Count; i++)
            {
                PowerPoint.Shape shape = selectedShapes[i];
                totalWidth += shape.Width;
                if (shape.Left < leftMost)
                {
                    leftMost = shape.Left;
                }
                if (shape.Left + shape.Width > rightMost)
                {
                    rightMost = shape.Left + shape.Width;
                }
            }

            float otherCenterX = leftMost + ((rightMost - leftMost) / 2);
            float moveDistance = referenceCenterX - otherCenterX;

            for (int i = 2; i <= selectedShapes.Count; i++)
            {
                PowerPoint.Shape shape = selectedShapes[i];
                shape.Left += moveDistance;
            }
        }

        private void AlignVertically(PowerPoint.ShapeRange selectedShapes, float referenceCenterY)
        {
            float totalHeight = 0;
            float topMost = float.MaxValue;
            float bottomMost = float.MinValue;

            for (int i = 2; i <= selectedShapes.Count; i++)
            {
                PowerPoint.Shape shape = selectedShapes[i];
                totalHeight += shape.Height;
                if (shape.Top < topMost)
                {
                    topMost = shape.Top;
                }
                if (shape.Top + shape.Height > bottomMost)
                {
                    bottomMost = shape.Top + shape.Height;
                }
            }

            float otherCenterY = topMost + ((bottomMost - topMost) / 2);
            float moveDistance = referenceCenterY - otherCenterY;

            for (int i = 2; i <= selectedShapes.Count; i++)
            {
                PowerPoint.Shape shape = selectedShapes[i];
                shape.Top += moveDistance;
            }
        }

        private void btnLeftAlign_Click(object sender, EventArgs e)
        {
            AlignShapes((selectedShapes, referencePosition) =>
            {
                float referenceLeft = selectedShapes[1].Left;
                float leftMost = float.MaxValue;

                for (int i = 2; i <= selectedShapes.Count; i++)
                {
                    PowerPoint.Shape shape = selectedShapes[i];
                    if (shape.Left < leftMost)
                    {
                        leftMost = shape.Left;
                    }
                }

                float moveDistance = referenceLeft - leftMost;

                for (int i = 2; i <= selectedShapes.Count; i++)
                {
                    PowerPoint.Shape shape = selectedShapes[i];
                    shape.Left += moveDistance;
                }
            });
        }

        private void btnHorizontalCenter_Click(object sender, EventArgs e)
        {
            AlignShapes(AlignHorizontally);
        }

        private void btnRightAlign_Click(object sender, EventArgs e)
        {
            AlignShapes((selectedShapes, referencePosition) =>
            {
                float referenceRight = selectedShapes[1].Left + selectedShapes[1].Width;
                float rightMost = float.MinValue;

                for (int i = 2; i <= selectedShapes.Count; i++)
                {
                    PowerPoint.Shape shape = selectedShapes[i];
                    if (shape.Left + shape.Width > rightMost)
                    {
                        rightMost = shape.Left + shape.Width;
                    }
                }

                float moveDistance = referenceRight - rightMost;

                for (int i = 2; i <= selectedShapes.Count; i++)
                {
                    PowerPoint.Shape shape = selectedShapes[i];
                    shape.Left += moveDistance;
                }
            });
        }

        private void btnTopAlign_Click(object sender, EventArgs e)
        {
            AlignShapes((selectedShapes, referencePosition) =>
            {
                float referenceTop = selectedShapes[1].Top;
                float topMost = float.MaxValue;

                for (int i = 2; i <= selectedShapes.Count; i++)
                {
                    PowerPoint.Shape shape = selectedShapes[i];
                    if (shape.Top < topMost)
                    {
                        topMost = shape.Top;
                    }
                }

                float moveDistance = referenceTop - topMost;

                for (int i = 2; i <= selectedShapes.Count; i++)
                {
                    PowerPoint.Shape shape = selectedShapes[i];
                    shape.Top += moveDistance;
                }
            });
        }

        private void btnVerticalCenter_Click(object sender, EventArgs e)
        {
            AlignShapes(AlignVertically);
        }

        private void btnBottomAlign_Click(object sender, EventArgs e)
        {
            AlignShapes((selectedShapes, referencePosition) =>
            {
                float referenceBottom = selectedShapes[1].Top + selectedShapes[1].Height;
                float bottomMost = float.MinValue;

                for (int i = 2; i <= selectedShapes.Count; i++)
                {
                    PowerPoint.Shape shape = selectedShapes[i];
                    if (shape.Top + shape.Height > bottomMost)
                    {
                        bottomMost = shape.Top + shape.Height;
                    }
                }

                float moveDistance = referenceBottom - bottomMost;

                for (int i = 2; i <= selectedShapes.Count; i++)
                {
                    PowerPoint.Shape shape = selectedShapes[i];
                    shape.Top += moveDistance;
                }
            });
        }
    }
}
