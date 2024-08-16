using System;
using System.Drawing;
using System.Windows.Forms;

namespace 课件帮PPT助手
{
    public partial class PinyinSelectionForm : Form
    {
        public string SelectedPinyin { get; private set; }

        public PinyinSelectionForm(string character, string word, string[] pinyins)
        {
            InitializeComponent();

            // Set the form's title
            this.Text = $"选择 [{character}] 的拼音";

            // Populate the combo box with the pinyin options
            comboBoxPinyins.Items.AddRange(pinyins);
            comboBoxPinyins.SelectedIndex = 0;

            // Set label text
            if (word.Length > 1)
            {
                label1.Text = $"检测到 [{word}] 中的 [{character}] 字是多音字，请选择正确的拼音：";
            }
            else
            {
                label1.Text = $"检测到 [{character}] 是多音字，请选择正确的拼音：";
            }

            // Adjust the height of the form and reposition controls
            AdjustFormLayout();

            // Attach the click event handler to the OK button
            buttonOk.Click += new EventHandler(buttonOk_Click);
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            SelectedPinyin = comboBoxPinyins.SelectedItem.ToString();
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
            comboBoxPinyins.Top = label1.Bottom + 20; // 20 像素的间距
            buttonOk.Top = comboBoxPinyins.Bottom + 20;

            // 根据内容动态调整窗体高度
            this.ClientSize = new Size(this.ClientSize.Width, buttonOk.Bottom + 30); // 30 是底部 padding
        }
    }
}
