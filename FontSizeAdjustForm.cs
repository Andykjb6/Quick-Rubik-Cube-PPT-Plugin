using System;
using System.Windows.Forms;

namespace 课件帮PPT助手
{
    public partial class FontSizeAdjustForm : Form
    {
        public float PinyinFontSize { get; private set; }
        public float HanziFontSize { get; private set; }

        public FontSizeAdjustForm(float currentPinyinFontSize, float currentHanziFontSize)
        {
            InitializeComponent();
            numericUpDownPinyin.Value = (decimal)currentPinyinFontSize;
            numericUpDownHanzi.Value = (decimal)currentHanziFontSize;
        }

        private void ConfirmButton_Click(object sender, EventArgs e)
        {
            PinyinFontSize = (float)numericUpDownPinyin.Value;
            HanziFontSize = (float)numericUpDownHanzi.Value;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
