using System;
using System.Windows.Forms;

namespace 课件帮PPT助手
{
    public partial class SampleGenerationForm : Form
    {
        public bool ExportSelectedSlides { get; private set; }
        public bool ExportAllSlides { get; private set; }
        public int SelectedSampleStyle { get; private set; }

        public SampleGenerationForm()
        {
            InitializeComponent();
        }

        private void PictureBoxStyle1_Click(object sender, EventArgs e)
        {
            SelectedSampleStyle = 1;
            ResetPictureBoxBorders();
            pictureBoxStyle1.BorderStyle = BorderStyle.Fixed3D;
        }

        private void PictureBoxStyle2_Click(object sender, EventArgs e)
        {
            SelectedSampleStyle = 2;
            ResetPictureBoxBorders();
            pictureBoxStyle2.BorderStyle = BorderStyle.Fixed3D;
        }

        private void PictureBoxStyle3_Click(object sender, EventArgs e)
        {
            SelectedSampleStyle = 3;
            ResetPictureBoxBorders();
            pictureBoxStyle3.BorderStyle = BorderStyle.Fixed3D;
        }

        private void PictureBoxStyle4_Click(object sender, EventArgs e)
        {
            SelectedSampleStyle = 4;
            ResetPictureBoxBorders();
            pictureBoxStyle4.BorderStyle = BorderStyle.Fixed3D;
        }

        private void PictureBoxStyle5_Click(object sender, EventArgs e)
        {
            SelectedSampleStyle = 5;
            ResetPictureBoxBorders();
            pictureBoxStyle5.BorderStyle = BorderStyle.Fixed3D;
        }

        private void PictureBoxStyle6_Click(object sender, EventArgs e)
        {
            SelectedSampleStyle = 6;
            ResetPictureBoxBorders();
            pictureBoxStyle6.BorderStyle = BorderStyle.Fixed3D;
        }

        private void ButtonGenerate_Click(object sender, EventArgs e)
        {
            ExportSelectedSlides = checkBoxSelectedSlides.Checked;
            ExportAllSlides = checkBoxAllSlides.Checked;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void ResetPictureBoxBorders()
        {
            pictureBoxStyle1.BorderStyle = BorderStyle.None;
            pictureBoxStyle2.BorderStyle = BorderStyle.None;
            pictureBoxStyle3.BorderStyle = BorderStyle.None;
            pictureBoxStyle4.BorderStyle = BorderStyle.None;
            pictureBoxStyle5.BorderStyle = BorderStyle.None;
            pictureBoxStyle6.BorderStyle = BorderStyle.None;
        }
    }
}
