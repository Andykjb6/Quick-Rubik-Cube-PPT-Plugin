using System;
using System.Drawing;
using System.Windows.Forms;

namespace 课件帮PPT助手
{
    public partial class PinYinForm : Form
    {
        public string SelectedPinyin { get; private set; }

        public PinYinForm(string hanzi, string word, string[] pinyins)
        {
            InitializeComponent();

            // 设置窗体的标题
            this.Text = $"选择 [{hanzi}] 的拼音";

            // 填充 ComboBox 的拼音选项
            comboBox1.Items.AddRange(pinyins);
            comboBox1.SelectedIndex = 0;

            // 设置 Label 文本
            if (!string.IsNullOrEmpty(word) && word.Length > 1)
            {
                label1.Text = $"检测到 [{word}] 中的 [{hanzi}] 字是多音字，请选择正确的拼音：";
            }
            else
            {
                label1.Text = $"检测到 [{hanzi}] 是多音字，请选择正确的拼音：";
            }

            // 调整窗体布局
            AdjustFormLayout();

            // 将点击事件处理程序附加到 OK 按钮
            button1.Click += new EventHandler(buttonOk_Click);
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            SelectedPinyin = comboBox1.SelectedItem.ToString();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void AdjustFormLayout()
        {
            // 确保 Label 的 AutoSize 属性为 true
            label1.AutoSize = true;

            // 设置 Label 的最大宽度，超过这个宽度会自动换行
            label1.MaximumSize = new Size(this.ClientSize.Width - 40, 0); // 40 是左右的 padding

            // 动态计算 Label 的高度
            int labelHeight = TextRenderer.MeasureText(label1.Text, label1.Font, label1.MaximumSize, TextFormatFlags.WordBreak).Height;

            // 设置 Label 的高度
            label1.Height = labelHeight;

            // 重新定位 ComboBox 和 Button
            comboBox1.Top = label1.Bottom + 20; // 20 像素的间距
            button1.Top = comboBox1.Bottom + 20;

            // 根据内容动态调整窗体高度
            this.ClientSize = new Size(this.ClientSize.Width, button1.Bottom + 30); // 30 是底部 padding
        }
    }
}
