using System.Windows.Forms;
using System;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace 课件帮PPT助手
{
    public partial class SampleGenerationForm : Form
    {
        public bool ExportSelectedSlides { get; private set; }
        public bool ExportAllSlides { get; private set; }
        public int SelectedSampleStyle { get; private set; }
        public string SelectedResolution { get; private set; }

        public SampleGenerationForm()
        {
            InitializeComponent();
            // 设置初始分辨率
            comboBoxResolution.SelectedIndex = 2; // 默认选择1920x1080 (全高清)
            UpdateSelectedSlidesCount();
        }

        private void ButtonGenerate_Click(object sender, EventArgs e)
        {
            ExportSelectedSlides = checkBoxSelectedSlides.Checked;
            ExportAllSlides = checkBoxAllSlides.Checked;

            if (pictureBoxStyle1.BorderStyle == BorderStyle.FixedSingle)
                SelectedSampleStyle = 1;
            else if (pictureBoxStyle2.BorderStyle == BorderStyle.FixedSingle)
                SelectedSampleStyle = 2;
            else if (pictureBoxStyle3.BorderStyle == BorderStyle.FixedSingle)
                SelectedSampleStyle = 3;
            else if (pictureBoxStyle4.BorderStyle == BorderStyle.FixedSingle)
                SelectedSampleStyle = 4;
            else if (pictureBoxStyle5.BorderStyle == BorderStyle.FixedSingle)
                SelectedSampleStyle = 5;
            else if (pictureBoxStyle6.BorderStyle == BorderStyle.FixedSingle)
                SelectedSampleStyle = 6;

            SelectedResolution = comboBoxResolution.SelectedItem.ToString();

            DialogResult = DialogResult.OK;
            Close();
        }

        private void PictureBoxStyle1_Click(object sender, EventArgs e)
        {
            ResetPictureBoxBorders();
            pictureBoxStyle1.BorderStyle = BorderStyle.FixedSingle;
            pictureBoxStyle1.BackgroundImage = global::课件帮PPT助手.Properties.Resources.样机1;
        }

        private void PictureBoxStyle2_Click(object sender, EventArgs e)
        {
            ResetPictureBoxBorders();
            pictureBoxStyle2.BorderStyle = BorderStyle.FixedSingle;
            pictureBoxStyle2.BackgroundImage = global::课件帮PPT助手.Properties.Resources.样机2;
        }

        private void PictureBoxStyle3_Click(object sender, EventArgs e)
        {
            ResetPictureBoxBorders();
            pictureBoxStyle3.BorderStyle = BorderStyle.FixedSingle;
            pictureBoxStyle3.BackgroundImage = global::课件帮PPT助手.Properties.Resources.样机3;
        }

        private void PictureBoxStyle4_Click(object sender, EventArgs e)
        {
            ResetPictureBoxBorders();
            pictureBoxStyle4.BorderStyle = BorderStyle.FixedSingle;
            pictureBoxStyle4.BackgroundImage = global::课件帮PPT助手.Properties.Resources.样机4;
        }

        private void PictureBoxStyle5_Click(object sender, EventArgs e)
        {
            ResetPictureBoxBorders();
            pictureBoxStyle5.BorderStyle = BorderStyle.FixedSingle;
            pictureBoxStyle5.BackgroundImage = global::课件帮PPT助手.Properties.Resources.样机5;
        }

        private void PictureBoxStyle6_Click(object sender, EventArgs e)
        {
            ResetPictureBoxBorders();
            pictureBoxStyle6.BorderStyle = BorderStyle.FixedSingle;
            pictureBoxStyle6.BackgroundImage = global::课件帮PPT助手.Properties.Resources.样机6;
        }

        private void CheckBoxSelectedSlides_CheckedChanged(object sender, EventArgs e)
        {
            UpdateSelectedSlidesCount();
        }

        private void CheckBoxAllSlides_CheckedChanged(object sender, EventArgs e)
        {
            UpdateSelectedSlidesCount();
        }

        private void UpdateSelectedSlidesCount()
        {
            var pptApp = Globals.ThisAddIn.Application;
            var selectedSlidesCount = pptApp.ActiveWindow.Selection.SlideRange?.Count ?? 0;
            labelSelectedSlidesCount.Text = $"已选中幻灯片数量：{selectedSlidesCount}";
        }

        private void ResetPictureBoxBorders()
        {
            pictureBoxStyle1.BorderStyle = BorderStyle.None;
            pictureBoxStyle2.BorderStyle = BorderStyle.None;
            pictureBoxStyle3.BorderStyle = BorderStyle.None;
            pictureBoxStyle4.BorderStyle = BorderStyle.None;
            pictureBoxStyle5.BorderStyle = BorderStyle.None;
            pictureBoxStyle6.BorderStyle = BorderStyle.None;

            pictureBoxStyle1.BackgroundImage = global::课件帮PPT助手.Properties.Resources.原始样机1;
            pictureBoxStyle2.BackgroundImage = global::课件帮PPT助手.Properties.Resources.原始样机2;
            pictureBoxStyle3.BackgroundImage = global::课件帮PPT助手.Properties.Resources.原始样机3;
            pictureBoxStyle4.BackgroundImage = global::课件帮PPT助手.Properties.Resources.原始样机4;
            pictureBoxStyle5.BackgroundImage = global::课件帮PPT助手.Properties.Resources.原始样机5;
            pictureBoxStyle6.BackgroundImage = global::课件帮PPT助手.Properties.Resources.原始样机6;
        }
    }
}
