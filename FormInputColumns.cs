using System;
using System.Windows.Forms;

namespace 课件帮PPT助手
{
    public partial class FormInputColumns : Form
    {
        public int ColumnsToDelete { get; private set; }

        public FormInputColumns()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (int.TryParse(txtColumns.Text, out int columns))
            {
                ColumnsToDelete = columns;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                MessageBox.Show("请输入有效的数字。");
            }
        }
    }
}
