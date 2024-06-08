using System;
using System.Windows.Forms;

namespace 课件帮PPT助手
{
    public partial class InputForm : Form
    {
        public string ReplacementText => textBox.Text;

        public event Action<string> TextConfirmed;

        public InputForm()
        {
            InitializeComponent();
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            TextConfirmed?.Invoke(textBox.Text);
        }
    }
}
