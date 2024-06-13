using System;
using System.Windows.Forms;

namespace 课件帮PPT助手
{
    public partial class BoardInputTextForm : Form
    {
        public string[] TextLines { get; private set; }

        public BoardInputTextForm()
        {
            InitializeComponent();
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            TextLines = textBox.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
