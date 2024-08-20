using System;
using System.Windows;
using System.Windows.Input;

namespace 课件帮PPT助手
{
    public partial class PinyinRatioForm : Window
    {
        public double PinyinRatio { get; set; }
        public double OffsetValue { get; set; }

        public PinyinRatioForm()
        {
            InitializeComponent();
            PinyinRatio = Properties.Settings.Default.PinyinRatio;
            OffsetValue = Properties.Settings.Default.OffsetValue; // 添加偏移量的读取

            TextBoxPinyinRatio.Text = PinyinRatio.ToString();
            TextBoxOffsetValue.Text = OffsetValue.ToString(); // 设置偏移量的TextBox
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            if (double.TryParse(TextBoxPinyinRatio.Text, out double ratio))
            {
                PinyinRatio = ratio;
            }

            if (double.TryParse(TextBoxOffsetValue.Text, out double offset))
            {
                OffsetValue = offset;
            }

            this.DialogResult = true;
            this.Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        // 增加偏移量的调整
        private void ButtonIncreaseOffset_Click(object sender, RoutedEventArgs e)
        {
            if (double.TryParse(TextBoxOffsetValue.Text, out double offset))
            {
                offset += 1;
                TextBoxOffsetValue.Text = offset.ToString();
            }
        }

        private void ButtonDecreaseOffset_Click(object sender, RoutedEventArgs e)
        {
            if (double.TryParse(TextBoxOffsetValue.Text, out double offset))
            {
                offset -= 1;
                TextBoxOffsetValue.Text = offset.ToString();
            }
        }

        private void ButtonIncrease_Click(object sender, RoutedEventArgs e)
        {
            if (double.TryParse(TextBoxPinyinRatio.Text, out double ratio))
            {
                ratio += 0.1;
                TextBoxPinyinRatio.Text = ratio.ToString("0.0");
            }
        }
        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }
        private void ButtonDecrease_Click(object sender, RoutedEventArgs e)
        {
            if (double.TryParse(TextBoxPinyinRatio.Text, out double ratio))
            {
                ratio -= 0.1;
                TextBoxPinyinRatio.Text = ratio.ToString("0.0");
            }
        }
    }
}
