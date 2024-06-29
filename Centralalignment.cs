using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace 课件帮PPT助手
{
    public partial class CentralalignmentForm : Form
    {
        
        private PowerPoint.Application pptApp;
        

        public CentralalignmentForm(PowerPoint.Application app)
        {
            InitializeComponent();
            pptApp = app;

            // 设置窗口位置到屏幕中心
            this.StartPosition = FormStartPosition.CenterScreen;
        }

        private void AlignShapes(AlignmentOptions alignment)
        {
            PowerPoint.DocumentWindow activeWindow = pptApp.ActiveWindow;

            if (activeWindow.Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("请先选择一个或多个形状。");
                return;
            }

            var selectedShapes = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in activeWindow.Selection.ShapeRange)
            {
                selectedShapes.Add(shape);
            }

            PowerPoint.Shape targetShape = selectedShapes[0];
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

        private void btnCenter_Click(object sender, EventArgs e)
        {
            AlignShapes(AlignmentOptions.Center);
        }

        private void btnLeft_Click(object sender, EventArgs e)
        {
            AlignShapes(AlignmentOptions.Left);
        }

        private void btnRight_Click(object sender, EventArgs e)
        {
            AlignShapes(AlignmentOptions.Right);
        }

        private void btnTop_Click(object sender, EventArgs e)
        {
            AlignShapes(AlignmentOptions.Top);
        }

        private void btnBottom_Click(object sender, EventArgs e)
        {
            AlignShapes(AlignmentOptions.Bottom);
        }

        private void CentralalignmentForm_Load(object sender, EventArgs e)
        {

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
