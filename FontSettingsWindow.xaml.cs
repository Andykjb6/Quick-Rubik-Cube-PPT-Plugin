using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace 课件帮PPT助手
{
    public partial class FontSettingsWindow : Window
    {
        public string PinyinFont { get; private set; }
        public double PinyinFontSize { get; private set; }
        public string HanziFont { get; private set; }
        public double HanziFontSize { get; private set; }
        public bool IsDefaultFont { get; private set; }

        public FontSettingsWindow()
        {
            InitializeComponent();
            LoadFontsIntoComboBox();

            // 默认勾选“设置为默认显示字体”选项
            DefaultFontCheckBox.IsChecked = true;

            // 显示默认的字号参数
            PinyinFontSizeTextBox.Text = Properties.Settings.Default.PinyinFontSize.ToString();
            HanziFontSizeTextBox.Text = Properties.Settings.Default.HanziFontSize.ToString();
        }

        private void LoadFontsIntoComboBox()
        {
            var fontFamilies = Fonts.SystemFontFamilies.OrderBy(f => f.Source);
            foreach (var fontFamily in fontFamilies)
            {
                PinyinFontComboBox.Items.Add(fontFamily.Source);
                HanziFontComboBox.Items.Add(fontFamily.Source);
            }

            // 设置初始选择项为当前的字体
            PinyinFontComboBox.SelectedItem = Properties.Settings.Default.PinyinFont;
            HanziFontComboBox.SelectedItem = Properties.Settings.Default.HanziFont;
        }

        private void ConfirmButton_Click(object sender, RoutedEventArgs e)
        {
            // 获取用户输入
            PinyinFont = PinyinFontComboBox.SelectedItem as string;
            HanziFont = HanziFontComboBox.SelectedItem as string;
            double.TryParse(PinyinFontSizeTextBox.Text, out double pinyinSize);
            double.TryParse(HanziFontSizeTextBox.Text, out double hanziSize);

            PinyinFontSize = pinyinSize;
            HanziFontSize = hanziSize;
            IsDefaultFont = DefaultFontCheckBox.IsChecked == true;

            if (IsDefaultFont)
            {
                // 保存用户设置到应用程序的设置文件，包括字体和字号
                Properties.Settings.Default.PinyinFont = PinyinFont;
                Properties.Settings.Default.PinyinFontSize = PinyinFontSize;
                Properties.Settings.Default.HanziFont = HanziFont;
                Properties.Settings.Default.HanziFontSize = HanziFontSize;
                Properties.Settings.Default.Save(); // 保存设置
            }

            this.DialogResult = true; // 关闭窗口并返回 DialogResult 为 true
        }

        private void PinyinIncreaseButton_Click(object sender, RoutedEventArgs e)
        {
            if (double.TryParse(PinyinFontSizeTextBox.Text, out double currentSize))
            {
                PinyinFontSizeTextBox.Text = (currentSize + 1).ToString();
            }
        }

        private void PinyinDecreaseButton_Click(object sender, RoutedEventArgs e)
        {
            if (double.TryParse(PinyinFontSizeTextBox.Text, out double currentSize) && currentSize > 1)
            {
                PinyinFontSizeTextBox.Text = (currentSize - 1).ToString();
            }
        }

        private void HanziIncreaseButton_Click(object sender, RoutedEventArgs e)
        {
            if (double.TryParse(HanziFontSizeTextBox.Text, out double currentSize))
            {
                HanziFontSizeTextBox.Text = (currentSize + 1).ToString();
            }
        }

        private void HanziDecreaseButton_Click(object sender, RoutedEventArgs e)
        {
            if (double.TryParse(HanziFontSizeTextBox.Text, out double currentSize) && currentSize > 1)
            {
                HanziFontSizeTextBox.Text = (currentSize - 1).ToString();
            }
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // 使窗口支持拖动
            this.DragMove();
        }
    }
}
