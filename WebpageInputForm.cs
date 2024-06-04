using System;
using System.Windows.Forms;

namespace 课件帮PPT助手
{
    public partial class WebpageInputForm : Form
    {
        public string WebpageUrl { get; private set; }

        public WebpageInputForm()
        {
            InitializeComponent();
        }

        private void btnEmbed_Click(object sender, EventArgs e)
        {
            WebpageUrl = txtUrl.Text;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
