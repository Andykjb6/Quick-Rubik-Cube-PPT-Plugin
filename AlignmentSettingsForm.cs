using System;
using System.Windows.Forms;

namespace 课件帮PPT助手
{
    public partial class AlignmentSettingsForm : Form
    {
        public Alignment SelectedAlignment { get; private set; }

        public AlignmentSettingsForm(Alignment currentAlignment)
        {
            InitializeComponent();

            // 根据当前对齐方式设置单选按钮
            switch (currentAlignment)
            {
                case Alignment.Center:
                    centerRadioButton.Checked = true;
                    break;
                case Alignment.Left:
                    leftRadioButton.Checked = true;
                    break;
                case Alignment.Right:
                    rightRadioButton.Checked = true;
                    break;
                case Alignment.Top:
                    topRadioButton.Checked = true;
                    break;
                case Alignment.Bottom:
                    bottomRadioButton.Checked = true;
                    break;
            }
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            if (centerRadioButton.Checked)
                SelectedAlignment = Alignment.Center;
            else if (leftRadioButton.Checked)
                SelectedAlignment = Alignment.Left;
            else if (rightRadioButton.Checked)
                SelectedAlignment = Alignment.Right;
            else if (topRadioButton.Checked)
                SelectedAlignment = Alignment.Top;
            else if (bottomRadioButton.Checked)
                SelectedAlignment = Alignment.Bottom;

            // 保存到注册表
            SettingsHelper.SaveAlignmentSetting(SelectedAlignment);

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void ExitButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void rightRadioButton_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
