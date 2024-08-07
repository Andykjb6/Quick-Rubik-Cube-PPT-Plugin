using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace 课件帮PPT助手
{
    public partial class LayoutForm : Form
    {
        public float Distance { get; private set; }
        public int Compactness { get; private set; }
        public int StartAngle { get; private set; }
        public bool IsClockwise { get; private set; }

        // 添加 LayoutLines 字段
        public List<PowerPoint.Shape> LayoutLines { get; private set; } = new List<PowerPoint.Shape>();

        public event EventHandler LayoutUpdated;

        private ContextMenuStrip contextMenu;
        private Point lastMousePosition;

        public LayoutForm()
        {
            InitializeComponent();

            Distance = 100; // 默认值
            Compactness = 50; // 默认值
            StartAngle = 0; // 默认起始角度为 0 度
            IsClockwise = true; // 默认顺时针

            InitializeContextMenu();

            this.MouseDown += LayoutForm_MouseDown;
            this.MouseMove += LayoutForm_MouseMove;
        }

        private void OnValueChanged(object sender, EventArgs e)
        {
            Distance = (float)numericUpDownDistance.Value;
            Compactness = trackBarCompactness.Value;
            StartAngle = (int)numericUpDownStartAngle.Value;
            LayoutUpdated?.Invoke(this, EventArgs.Empty);
        }

        private void OnDirectionChanged(object sender, EventArgs e)
        {
            IsClockwise = comboBoxDirection.SelectedIndex == 0;
            LayoutUpdated?.Invoke(this, EventArgs.Empty);
        }

        private void InitializeContextMenu()
        {
            contextMenu = new ContextMenuStrip();

            var exitMenuItem = new ToolStripMenuItem("退出");
            exitMenuItem.Click += ExitMenuItem_Click;
            contextMenu.Items.Add(exitMenuItem);

            this.ContextMenuStrip = contextMenu;
        }

        private void ExitMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void LayoutForm_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                lastMousePosition = e.Location;
            }
        }

        private void LayoutForm_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                int dx = e.X - lastMousePosition.X;
                int dy = e.Y - lastMousePosition.Y;
                this.Location = new Point(this.Location.X + dx, this.Location.Y + dy);
            }
        }
    }
}
