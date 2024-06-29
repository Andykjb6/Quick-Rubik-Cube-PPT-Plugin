using System;
using System.IO;
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

        private void importButton_Click(object sender, EventArgs e) // 新增
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // 读取文件内容并设置到文本框中
                string fileContent = File.ReadAllText(openFileDialog.FileName);
                textBox.Text = fileContent;
            }
        }
    }
}
