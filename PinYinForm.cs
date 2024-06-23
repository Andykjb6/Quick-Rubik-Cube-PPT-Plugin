using System;
using System.Linq;
using System.Windows.Forms;

namespace 课件帮PPT助手
{
    public partial class PinYinForm : Form
    {
        public string SelectedPinyin { get; private set; }

        public PinYinForm(string hanzi, string word, string[] pinyins)
        {
            InitializeComponent();
            comboBox1.Items.AddRange(pinyins.Select(p => p.Trim()).ToArray());
            comboBox1.SelectedIndex = 0;

            label1.Text = $"[{word}]中的[{hanzi}]检测到多音字：";

            button1.Click += (sender, e) =>
            {
                SelectedPinyin = comboBox1.SelectedItem.ToString();
                DialogResult = DialogResult.OK;
                Close();
            };
        }
    }
}
