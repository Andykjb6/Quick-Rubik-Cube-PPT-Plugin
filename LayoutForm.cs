using System;
using System.Drawing;
using System.Windows.Forms;

namespace 课件帮PPT助手
{
    public partial class LayoutForm : Form
    {
        public float Distance { get; private set; }
        public int Compactness { get; private set; }
        public int StartAngle { get; private set; }
        public bool IsClockwise { get; private set; } // 新增属性：方向

        public event EventHandler LayoutUpdated;

        private ContextMenuStrip contextMenu; // 定义右键菜单
        private Point lastMousePosition; // 记录鼠标的位置

        public LayoutForm()
        {
            InitializeComponent();

            Distance = 100; // 默认值
            Compactness = 50; // 默认值
            StartAngle = 0; // 默认起始角度为 0 度
            IsClockwise = true; // 默认顺时针

            InitializeContextMenu(); // 初始化右键菜单

            // 添加鼠标事件处理，以实现无边框窗体的拖动
            this.MouseDown += LayoutForm_MouseDown;
            this.MouseMove += LayoutForm_MouseMove;
        }

        private void OnValueChanged(object sender, EventArgs e)
        {
            Distance = (float)numericUpDownDistance.Value; // 转换为 float
            Compactness = trackBarCompactness.Value; // 转换为 int
            StartAngle = (int)numericUpDownStartAngle.Value; // 转换为 int
            LayoutUpdated?.Invoke(this, EventArgs.Empty); // 触发事件以更新布局
        }

        private void OnDirectionChanged(object sender, EventArgs e)
        {
            IsClockwise = comboBoxDirection.SelectedIndex == 0; // 顺时针为 true，逆时针为 false
            LayoutUpdated?.Invoke(this, EventArgs.Empty); // 更新布局
        }

        private void InitializeContextMenu()
        {
            // 创建ContextMenuStrip对象
            contextMenu = new ContextMenuStrip();

            // 添加"退出"选项
            var exitMenuItem = new ToolStripMenuItem("退出");
            exitMenuItem.Click += ExitMenuItem_Click; // 绑定点击事件
            contextMenu.Items.Add(exitMenuItem);

            // 将右键菜单绑定到窗体上
            this.ContextMenuStrip = contextMenu;
        }

        private void ExitMenuItem_Click(object sender, EventArgs e)
        {
            this.Close(); // 关闭窗体
        }

        // 处理鼠标按下事件，记录当前位置
        private void LayoutForm_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                lastMousePosition = e.Location;
            }
        }

        // 处理鼠标移动事件，实现窗体拖动
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
