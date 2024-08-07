using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointApplication = Microsoft.Office.Interop.PowerPoint.Application;

namespace 课件帮PPT助手
{
    public partial class AlignmentWindow : Window
    {
        private PowerPointApplication pptApp;

        public AlignmentWindow(PowerPointApplication app)
        {
            InitializeComponent();
            pptApp = app;
        }

        // 鼠标拖动窗体
        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                this.DragMove();
            }
        }

        // 右键菜单退出选项
        private void MenuItem_Exit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        // 第一组对齐功能：平移居中
        private void AlignShapes(AlignmentOptions alignment)
        {
            DocumentWindow activeWindow = pptApp.ActiveWindow;

            if (activeWindow.Selection.Type != PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("请先选择一个或多个形状。");
                return;
            }

            var selectedShapes = new List<Shape>();
            foreach (Shape shape in activeWindow.Selection.ShapeRange)
            {
                selectedShapes.Add(shape);
            }

            Shape targetShape = selectedShapes[0];
            var independentShapes = selectedShapes.GetRange(1, selectedShapes.Count - 1);

            float targetLeft = targetShape.Left;
            float targetTop = targetShape.Top;
            float targetRight = targetLeft + targetShape.Width;
            float targetBottom = targetTop + targetShape.Height;

            float shapesLeft = float.MaxValue;
            float shapesTop = float.MaxValue;
            float shapesRight = float.MinValue;
            float shapesBottom = float.MinValue;

            foreach (var shape in independentShapes)
            {
                if (shape.Left < shapesLeft) shapesLeft = shape.Left;
                if (shape.Top < shapesTop) shapesTop = shape.Top;
                if (shape.Left + shape.Width > shapesRight) shapesRight = shape.Left + shape.Width;
                if (shape.Top + shape.Height > shapesBottom) shapesBottom = shape.Top + shape.Height;
            }

            float offsetX = 0;
            float offsetY = 0;

            switch (alignment)
            {
                case AlignmentOptions.Center:
                    offsetX = (targetLeft + targetRight) / 2 - (shapesLeft + shapesRight) / 2;
                    offsetY = (targetTop + targetBottom) / 2 - (shapesTop + shapesBottom) / 2;
                    break;
                case AlignmentOptions.Left:
                    offsetX = targetLeft - shapesLeft;
                    offsetY = (targetTop + targetBottom) / 2 - (shapesTop + shapesBottom) / 2;
                    break;
                case AlignmentOptions.Right:
                    offsetX = targetRight - shapesRight;
                    offsetY = (targetTop + targetBottom) / 2 - (shapesTop + shapesBottom) / 2;
                    break;
                case AlignmentOptions.Top:
                    offsetX = (targetLeft + targetRight) / 2 - (shapesLeft + shapesRight) / 2;
                    offsetY = targetTop - shapesTop;
                    break;
                case AlignmentOptions.Bottom:
                    offsetX = (targetLeft + targetRight) / 2 - (shapesLeft + shapesRight) / 2;
                    offsetY = targetBottom - shapesBottom;
                    break;
            }

            foreach (var shape in independentShapes)
            {
                shape.Left += offsetX;
                shape.Top += offsetY;
            }
        }

        private void btnCenter_Click(object sender, RoutedEventArgs e)
        {
            AlignShapes(AlignmentOptions.Center);
        }

        private void btnLeft_Click(object sender, RoutedEventArgs e)
        {
            AlignShapes(AlignmentOptions.Left);
        }

        private void btnRight_Click(object sender, RoutedEventArgs e)
        {
            AlignShapes(AlignmentOptions.Right);
        }

        private void btnTop_Click(object sender, RoutedEventArgs e)
        {
            AlignShapes(AlignmentOptions.Top);
        }

        private void btnBottom_Click(object sender, RoutedEventArgs e)
        {
            AlignShapes(AlignmentOptions.Bottom);
        }

        // 第二组对齐功能：移动对齐
        private void AlignShapes(Action<ShapeRange, float> alignAction)
        {
            Selection selection = pptApp.ActiveWindow.Selection;

            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 1)
            {
                ShapeRange selectedShapes = selection.ShapeRange;

                Shape referenceShape = selectedShapes[1];
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
                    referencePosition = 0;
                }

                alignAction(selectedShapes, referencePosition);
            }
            else
            {
                MessageBox.Show("请至少选择两个对象进行对齐！");
            }
        }

        private void AlignHorizontally(ShapeRange selectedShapes, float referenceCenterX)
        {
            float totalWidth = 0;
            float leftMost = float.MaxValue;
            float rightMost = float.MinValue;

            for (int i = 2; i <= selectedShapes.Count; i++)
            {
                Shape shape = selectedShapes[i];
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
                Shape shape = selectedShapes[i];
                shape.Left += moveDistance;
            }
        }

        private void AlignVertically(ShapeRange selectedShapes, float referenceCenterY)
        {
            float totalHeight = 0;
            float topMost = float.MaxValue;
            float bottomMost = float.MinValue;

            for (int i = 2; i <= selectedShapes.Count; i++)
            {
                Shape shape = selectedShapes[i];
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
                Shape shape = selectedShapes[i];
                shape.Top += moveDistance;
            }
        }

        private void btnLeftAlign_Click(object sender, RoutedEventArgs e)
        {
            AlignShapes((selectedShapes, referencePosition) =>
            {
                float referenceLeft = selectedShapes[1].Left;
                float leftMost = float.MaxValue;

                for (int i = 2; i <= selectedShapes.Count; i++)
                {
                    Shape shape = selectedShapes[i];
                    if (shape.Left < leftMost)
                    {
                        leftMost = shape.Left;
                    }
                }

                float moveDistance = referenceLeft - leftMost;

                for (int i = 2; i <= selectedShapes.Count; i++)
                {
                    Shape shape = selectedShapes[i];
                    shape.Left += moveDistance;
                }
            });
        }

        private void btnHorizontalCenter_Click(object sender, RoutedEventArgs e)
        {
            AlignShapes(AlignHorizontally);
        }

        private void btnRightAlign_Click(object sender, RoutedEventArgs e)
        {
            AlignShapes((selectedShapes, referencePosition) =>
            {
                float referenceRight = selectedShapes[1].Left + selectedShapes[1].Width;
                float rightMost = float.MinValue;

                for (int i = 2; i <= selectedShapes.Count; i++)
                {
                    Shape shape = selectedShapes[i];
                    if (shape.Left + shape.Width > rightMost)
                    {
                        rightMost = shape.Left + shape.Width;
                    }
                }

                float moveDistance = referenceRight - rightMost;

                for (int i = 2; i <= selectedShapes.Count; i++)
                {
                    Shape shape = selectedShapes[i];
                    shape.Left += moveDistance;
                }
            });
        }

        private void btnTopAlign_Click(object sender, RoutedEventArgs e)
        {
            AlignShapes((selectedShapes, referencePosition) =>
            {
                float referenceTop = selectedShapes[1].Top;
                float topMost = float.MaxValue;

                for (int i = 2; i <= selectedShapes.Count; i++)
                {
                    Shape shape = selectedShapes[i];
                    if (shape.Top < topMost)
                    {
                        topMost = shape.Top;
                    }
                }

                float moveDistance = referenceTop - topMost;

                for (int i = 2; i <= selectedShapes.Count; i++)
                {
                    Shape shape = selectedShapes[i];
                    shape.Top += moveDistance;
                }
            });
        }

        private void btnVerticalCenter_Click(object sender, RoutedEventArgs e)
        {
            AlignShapes(AlignVertically);
        }

        private void btnBottomAlign_Click(object sender, RoutedEventArgs e)
        {
            AlignShapes((selectedShapes, referencePosition) =>
            {
                float referenceBottom = selectedShapes[1].Top + selectedShapes[1].Height;
                float bottomMost = float.MinValue;

                for (int i = 2; i <= selectedShapes.Count; i++)
                {
                    Shape shape = selectedShapes[i];
                    if (shape.Top + shape.Height > bottomMost)
                    {
                        bottomMost = shape.Top + shape.Height;
                    }
                }

                float moveDistance = referenceBottom - bottomMost;

                for (int i = 2; i <= selectedShapes.Count; i++)
                {
                    Shape shape = selectedShapes[i];
                    shape.Top += moveDistance;
                }
            });
        }

        // 第三组对齐功能：指定对齐
        private void CenterAlignButton_Click(object sender, RoutedEventArgs e)
        {
            AlignShapes(SpecifyAlignment.Center);
        }

        private void LeftAlignButton_Click(object sender, RoutedEventArgs e)
        {
            AlignShapes(SpecifyAlignment.Left);
        }

        private void HorizontalCenterAlignButton_Click(object sender, RoutedEventArgs e)
        {
            AlignShapes(SpecifyAlignment.HorizontalCenter);
        }

        private void RightAlignButton_Click(object sender, RoutedEventArgs e)
        {
            AlignShapes(SpecifyAlignment.Right);
        }

        private void TopAlignButton_Click(object sender, RoutedEventArgs e)
        {
            AlignShapes(SpecifyAlignment.Top);
        }

        private void VerticalCenterAlignButton_Click(object sender, RoutedEventArgs e)
        {
            AlignShapes(SpecifyAlignment.VerticalCenter);
        }

        private void BottomAlignButton_Click(object sender, RoutedEventArgs e)
        {
            AlignShapes(SpecifyAlignment.Bottom);
        }

        private void AlignShapes(SpecifyAlignment alignment)
        {
            Selection selection = pptApp.ActiveWindow.Selection;

            if (selection.Type != PpSelectionType.ppSelectionShapes || selection.ShapeRange.Count < 2)
            {
                MessageBox.Show("请至少选择两个对象进行匹配对齐。");
                return;
            }

            ShapeRange selectedShapes = selection.ShapeRange;
            List<Shape> referenceShapes = new List<Shape>();
            List<Shape> targetShapes = new List<Shape>();

            for (int i = 0; i < selectedShapes.Count; i++)
            {
                Shape shape = selectedShapes[i + 1]; // ShapeRange是1-based索引
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
            foreach (Shape shape in selectedShapes)
            {
                int prefixIndex = shape.Name.IndexOf('-');
                if (prefixIndex != -1)
                {
                    shape.Name = shape.Name.Substring(prefixIndex + 1);
                }
            }
        }

        private void AlignShape(Shape referenceShape, Shape targetShape, SpecifyAlignment alignment)
        {
            switch (alignment)
            {
                case SpecifyAlignment.Center:
                    targetShape.Left = referenceShape.Left + (referenceShape.Width - targetShape.Width) / 2;
                    targetShape.Top = referenceShape.Top + (referenceShape.Height - targetShape.Height) / 2;
                    break;
                case SpecifyAlignment.Left:
                    targetShape.Left = referenceShape.Left;
                    break;
                case SpecifyAlignment.HorizontalCenter:
                    targetShape.Left = referenceShape.Left + (referenceShape.Width - targetShape.Width) / 2;
                    break;
                case SpecifyAlignment.Right:
                    targetShape.Left = referenceShape.Left + referenceShape.Width - targetShape.Width;
                    break;
                case SpecifyAlignment.Top:
                    targetShape.Top = referenceShape.Top;
                    break;
                case SpecifyAlignment.VerticalCenter:
                    targetShape.Top = referenceShape.Top + (referenceShape.Height - targetShape.Height) / 2;
                    break;
                case SpecifyAlignment.Bottom:
                    targetShape.Top = referenceShape.Top + referenceShape.Height - targetShape.Height;
                    break;
            }
        }

        private enum SpecifyAlignment
        {
            Center,
            Left,
            HorizontalCenter,
            Right,
            Top,
            VerticalCenter,
            Bottom
        }
    }

    public enum AlignmentOptions
    {
        Center,
        Left,
        Right,
        Top,
        Bottom
    }
}
