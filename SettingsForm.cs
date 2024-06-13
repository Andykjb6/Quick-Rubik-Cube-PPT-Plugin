using System;
using System.Drawing;
using System.Windows.Forms;

namespace 课件帮PPT助手
{
    public partial class SettingsForm : Form
    {
        public Font SelectedFont { get; private set; }
        public Color TimeTextColor { get; private set; }
        public bool IsCountdown { get; private set; }
        public Color BackgroundColor { get; private set; }
        public Color DarkModeButtonColor { get; private set; }

        public SettingsForm(Font currentFont, Color timeTextColor, bool isCountdown, Color backgroundColor, Color darkModeButtonColor)
        {
            SelectedFont = currentFont;
            TimeTextColor = timeTextColor;
            IsCountdown = isCountdown;
            BackgroundColor = backgroundColor;
            DarkModeButtonColor = darkModeButtonColor;

            InitializeComponent();
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {
            foreach (System.Drawing.FontFamily font in System.Drawing.FontFamily.Families)
            {
                this.fontComboBox.Items.Add(font.Name);
            }
            this.fontComboBox.SelectedItem = this.SelectedFont.Name;
        }

        private void TimeTextColorButton_Click(object sender, EventArgs e)
        {
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                TimeTextColor = colorDialog.Color;
                timeTextColorBox.Text = TimeTextColor.ToArgb().ToString("X");
            }
        }

        private void BackgroundColorButton_Click(object sender, EventArgs e)
        {
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                BackgroundColor = colorDialog.Color;
                backgroundColorBox.Text = BackgroundColor.ToArgb().ToString("X");
            }
        }

        private void DarkModeButtonColorButton_Click(object sender, EventArgs e)
        {
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                DarkModeButtonColor = colorDialog.Color;
                darkModeButtonColorBox.Text = DarkModeButtonColor.ToArgb().ToString("X");
            }
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            IsCountdown = countdownRadioButton.Checked;
            SelectedFont = new Font(fontComboBox.SelectedItem.ToString(), SelectedFont.Size);
        }
    }
}
