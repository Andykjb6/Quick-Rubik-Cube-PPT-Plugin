using System;
using System.Windows.Forms;

namespace 课件帮PPT助手
{
    public partial class FormAddColumns : Form
    {
        public int ColumnsToAdd { get; private set; }

        public FormAddColumns()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            // 尝试将用户输入的内容转换为整数
            if (int.TryParse(txtColumns.Text, out int columns))
            {
                // 确保输入的列数为正整数
                if (columns > 0)
                {
                    ColumnsToAdd = columns;
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("请输入一个正整数列数。", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show("请输入有效的列数。", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
