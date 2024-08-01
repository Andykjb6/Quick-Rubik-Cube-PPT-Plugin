using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace 课件帮PPT助手
{
    public partial class TimerWindow : Window
    {
        private DispatcherTimer timer;
        private TimeSpan time;
        private bool isCountingDown = true; // 默认倒计时
        private bool isRunning = false;
        private bool isDarkTheme = false;

        public TimerWindow()
        {
            InitializeComponent();

            // 初始化计时器
            timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(1);
            timer.Tick += Timer_Tick;

            // 初始化事件处理
            StartPauseButton.Click += StartPauseButton_Click;
            ResetButton.Click += ResetButton_Click;
            ThemeToggleButton.Click += ThemeToggleButton_Click;
            TimerModeComboBox.SelectionChanged += TimerModeComboBox_SelectionChanged;

            // 设置初始时间
            SetTimeDisplay(TimeSpan.FromHours(1));
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            if (isCountingDown)
            {
                if (time.TotalSeconds > 0)
                {
                    time = time.Add(TimeSpan.FromSeconds(-1));
                }
                else
                {
                    timer.Stop();
                    isRunning = false;
                    StartPauseButton.Content = "开始";
                }
            }
            else
            {
                time = time.Add(TimeSpan.FromSeconds(1));
            }
            SetTimeDisplay(time);
        }

        private void StartPauseButton_Click(object sender, RoutedEventArgs e)
        {
            if (isRunning)
            {
                timer.Stop();
                StartPauseButton.Content = "开始";
            }
            else
            {
                if (TryGetTimeFromInput(out time))
                {
                    timer.Start();
                    StartPauseButton.Content = "暂停";
                }
                else
                {
                    MessageBox.Show("请输入有效的时间格式（例如 00:15:00）", "无效的时间", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            isRunning = !isRunning;
        }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            timer.Stop();
            isRunning = false;
            StartPauseButton.Content = "开始";
            time = isCountingDown ? TimeSpan.FromHours(1) : TimeSpan.Zero; // 倒计时默认1小时，顺计时从0开始
            SetTimeDisplay(time);
        }

        private void ThemeToggleButton_Click(object sender, RoutedEventArgs e)
        {
            if (isDarkTheme)
            {
                this.Background = SystemColors.ControlBrush;
                SetTextColor(SystemColors.ControlTextBrush);
            }
            else
            {
                this.Background = SystemColors.ControlDarkDarkBrush;
                SetTextColor(SystemColors.ControlLightLightBrush);
            }
            isDarkTheme = !isDarkTheme;
        }

        private void TimerModeComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (TimerModeComboBox.SelectedIndex == 0) // 倒计时
            {
                isCountingDown = true;
                time = TimeSpan.FromHours(1); // 默认倒计时1小时
            }
            else // 顺计时
            {
                isCountingDown = false;
                time = TimeSpan.Zero;
            }
            SetTimeDisplay(time);
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void TimeBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // 只允许输入数字
            e.Handled = !int.TryParse(e.Text, out _);
        }

        private void SetTimeDisplay(TimeSpan time)
        {
            HourBox.Text = time.Hours.ToString("D2");
            MinuteBox.Text = time.Minutes.ToString("D2");
            SecondBox.Text = time.Seconds.ToString("D2");
        }

        private bool TryGetTimeFromInput(out TimeSpan time)
        {
            time = TimeSpan.Zero;
            try
            {
                int hours = int.Parse(HourBox.Text);
                int minutes = int.Parse(MinuteBox.Text);
                int seconds = int.Parse(SecondBox.Text);

                time = new TimeSpan(hours, minutes, seconds);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void SetTextColor(System.Windows.Media.Brush color)
        {
            HourBox.Foreground = color;
            MinuteBox.Foreground = color;
            SecondBox.Foreground = color;
        }
    }
}
