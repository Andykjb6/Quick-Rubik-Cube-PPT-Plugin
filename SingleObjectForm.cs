using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace 课件帮PPT助手
{
    public partial class SingleObjectForm : Form
    {
        private PowerPoint.Application pptApp;
        private PowerPoint.ShapeRange shapes;
        private float radius;
        private float initialRotation;
        private float finalRotation;
        private float sizeIncrement;
        private int copyCount;

        private List<PowerPoint.Shape> copiedShapes = new List<PowerPoint.Shape>();
        private Dictionary<int, (float Width, float Height)> initialSizes = new Dictionary<int, (float Width, float Height)>();

        public SingleObjectForm(PowerPoint.Application pptApp, PowerPoint.ShapeRange shapes, float radius, float initialRotation, float finalRotation, float sizeIncrement, int copyCount)
        {
            InitializeComponent();

            this.pptApp = pptApp;
            this.shapes = shapes;
            this.radius = radius;
            this.initialRotation = initialRotation;
            this.finalRotation = finalRotation;
            this.sizeIncrement = sizeIncrement;
            this.copyCount = copyCount;

            this.radiusTrackBar.Value = (int)radius;
            this.initialRotationUpDown.Value = (int)initialRotation;
            this.finalRotationUpDown.Value = (int)finalRotation;
            this.sizeIncrementTrackBar.Minimum = -50; // 调整最小值
            this.sizeIncrementTrackBar.Maximum = 50; // 调整最大值
            this.sizeIncrementTrackBar.Value = (int)sizeIncrement;
            this.copyCountTrackBar.Value = copyCount;

            this.radiusTrackBar.ValueChanged += (s, ev) => UpdateShapes();
            this.initialRotationUpDown.ValueChanged += (s, ev) => UpdateShapes();
            this.finalRotationUpDown.ValueChanged += (s, ev) => UpdateShapes();
            this.sizeIncrementTrackBar.ValueChanged += (s, ev) => UpdateShapes();
            this.copyCountTrackBar.ValueChanged += (s, ev) => UpdateShapes();
            this.resetButton.Click += (s, ev) => ResetParameters();

            UpdateShapes();
        }

        private void UpdateShapes()
        {
            radius = radiusTrackBar.Value;
            initialRotation = (float)initialRotationUpDown.Value;
            finalRotation = (float)finalRotationUpDown.Value;
            sizeIncrement = sizeIncrementTrackBar.Value * 0.2f; // 调整递进值
            copyCount = copyCountTrackBar.Value;

            PerformCircularDistribution(pptApp, shapes, radius, initialRotation, finalRotation, sizeIncrement, copyCount);
        }

        private void ResetParameters()
        {
            radiusTrackBar.Value = 100;
            initialRotationUpDown.Value = 0;
            finalRotationUpDown.Value = 0;
            sizeIncrementTrackBar.Value = 0;
            copyCountTrackBar.Value = 0;

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

        private void PerformCircularDistribution(PowerPoint.Application pptApp, PowerPoint.ShapeRange shapes, float radius, float initialRotation, float finalRotation, float sizeIncrement, int copyCount)
        {
            foreach (PowerPoint.Shape shape in copiedShapes)
            {
                shape.Delete();
            }
            copiedShapes.Clear();

            int count = copyCount + 1; // 确保选中的对象也被包含
            float angleStep = 360.0f / count;
            float angleIncrement = (finalRotation - initialRotation) / count;

            float currentRadius = radius;

            for (int i = 0; i < count; i++)
            {
                float angle = initialRotation + i * angleStep;
                float radians = angle * (float)(Math.PI / 180.0);
                float newX = (float)(currentRadius * Math.Cos(radians));
                float newY = (float)(currentRadius * Math.Sin(radians));

                PowerPoint.Shape shape;
                if (i == 0)
                {
                    shape = shapes[1]; // 第一个形状是选中的对象
                }
                else
                {
                    shape = shapes[1].Duplicate()[1];
                    copiedShapes.Add(shape);
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
                    float newSize = initialSizes[shape.Id].Width * (1 + i * sizeIncrement / 100.0f);
                    shape.Width = newSize;
                    shape.Height = newSize;

                    // 增加当前半径以保持间距相等
                    currentRadius += sizeIncrement / 2.0f;
                }
            }
        }

        private void finalRotationUpDown_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}
