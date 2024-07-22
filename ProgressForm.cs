using System;
using System.Windows.Forms;

namespace 课件帮PPT助手
{
    public partial class ProgressForm : Form
    {
        public ProgressForm()
        {
            InitializeComponent();
        }

        public ProgressBar ProgressBar => progressBar;
        public Label ProgressLabel => label;
    }
}
