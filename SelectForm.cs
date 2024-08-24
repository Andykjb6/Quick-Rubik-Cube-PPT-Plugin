using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 课件帮PPT助手
{
    public partial class SelectForm : Form
    {
        public string SelectedOption { get; private set; }

        public SelectForm(string hanzi, string[] options)
        {
            InitializeComponent();
            labelPrompt.Text = $"“{hanzi}”字有以下拆法，请选择其一：";
            comboBoxOptions.Items.AddRange(options);
            comboBoxOptions.SelectedIndex = 0;  // 默认选中第一项
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            SelectedOption = comboBoxOptions.SelectedItem.ToString();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}