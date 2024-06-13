using System;
using System.Drawing;
using System.Windows.Forms;

namespace 课件帮PPT助手
{
    public partial class TimerForm : Form
    {
        private DateTime targetTime;
        private bool isCountdown = true; // 默认倒计时
        private Font currentFont = new Font("Arial", 40, FontStyle.Bold);
        private Color backgroundColor = Color.White;
        private Color timeTextColor = Color.Black;
        private Color darkModeButtonColor = Color.LightGray;

        public TimerForm()
        {
            InitializeComponent();
            SetButtonStyles();
        }

        private void SetButtonStyles()
        {
            // 设置按钮的初始颜色
            SetButtonStyle(startButton, Color.White, Color.Black);
            SetButtonStyle(stopButton, Color.White, Color.Black);
            SetButtonStyle(resetButton, Color.White, Color.Black);
            SetButtonStyle(closeButton, backgroundColor, Color.Black);
            SetButtonStyle(settingsButton, backgroundColor, Color.Black);
        }

        private void SetButtonStyle(Button button, Color backColor, Color foreColor)
        {
            button.BackColor = backColor;
            button.ForeColor = foreColor;
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
        }

        private void TimerForm_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, this.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid);
        }

        private void StartButton_Click(object sender, EventArgs e)
        {
            if (TimeSpan.TryParseExact(timeTextBox.Text, @"hh\:mm\:ss", null, out TimeSpan timeSpan))
            {
                if (isCountdown)
                {
                    targetTime = DateTime.Now.Add(timeSpan);
                }
                else
                {
                    targetTime = DateTime.Now.AddHours(-timeSpan.TotalHours).AddMinutes(-timeSpan.TotalMinutes).AddSeconds(-timeSpan.TotalSeconds);
                }

                startButton.Enabled = false;
                stopButton.Enabled = true;
                resetButton.Enabled = true;
                timeTextBox.ReadOnly = true;

                timer.Start();
            }
            else
            {
                MessageBox.Show("请输入有效的时间格式（hh:mm:ss）。");
            }
        }

        private void StopButton_Click(object sender, EventArgs e)
        {
            timer.Stop();
            startButton.Enabled = true;
            stopButton.Enabled = false;
            resetButton.Enabled = true;
        }

        private void ResetButton_Click(object sender, EventArgs e)
        {
            timer.Stop();
            timeTextBox.Text = "00:00:00";
            startButton.Enabled = true;
            stopButton.Enabled = false;
            resetButton.Enabled = false;
            timeTextBox.ReadOnly = false;
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void SettingsButton_Click(object sender, EventArgs e)
        {
            this.Hide(); // 隐藏计时器窗口
            using (SettingsForm settingsForm = new SettingsForm(currentFont, timeTextColor, isCountdown, backgroundColor, darkModeButtonColor))
            {
                if (settingsForm.ShowDialog() == DialogResult.OK)
                {
                    currentFont = settingsForm.SelectedFont;
                    timeTextColor = settingsForm.TimeTextColor;
                    isCountdown = settingsForm.IsCountdown;
                    backgroundColor = settingsForm.BackgroundColor;
                    darkModeButtonColor = settingsForm.DarkModeButtonColor;

                    timeTextBox.Font = currentFont;
                    timeTextBox.ForeColor = timeTextColor;
                    this.BackColor = backgroundColor;
                    timeTextBox.BackColor = backgroundColor; // 同步背景色
                    darkModeButton.BackColor = darkModeButtonColor;

                    // 同步设置按钮和关闭按钮的背景颜色
                    SetButtonStyle(closeButton, backgroundColor, closeButton.ForeColor);
                    SetButtonStyle(settingsButton, backgroundColor, settingsButton.ForeColor);
                }
            }
            this.Show(); // 显示计时器窗口
        }

        private void DarkModeButton_Click(object sender, EventArgs e)
        {
            if (this.BackColor == Color.White)
            {
                this.BackColor = Color.Black;
                timeTextBox.BackColor = Color.Black;
                timeTextBox.ForeColor = Color.White;
                darkModeButton.Text = "浅色模式";
                SetButtonStyle(closeButton, Color.Black, Color.White);
                SetButtonStyle(settingsButton, Color.Black, Color.White);
            }
            else
            {
                this.BackColor = Color.White;
                timeTextBox.BackColor = Color.White;
                timeTextBox.ForeColor = Color.Black;
                darkModeButton.Text = "暗色模式";
                SetButtonStyle(closeButton, Color.White, Color.Black);
                SetButtonStyle(settingsButton, Color.White, Color.Black);
            }
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            TimeSpan remainingTime = isCountdown ? targetTime - DateTime.Now : DateTime.Now - targetTime;

            if (remainingTime.TotalSeconds <= 0)
            {
                timer.Stop();
                timeTextBox.Text = "00:00:00";
                MessageBox.Show("时间到！");
                startButton.Enabled = true;
                stopButton.Enabled = false;
                resetButton.Enabled = true;
                timeTextBox.ReadOnly = false;
                System.Media.SystemSounds.Exclamation.Play();
            }
            else
            {
                timeTextBox.Text = remainingTime.ToString(@"hh\:mm\:ss");
            }
        }
    }
}
