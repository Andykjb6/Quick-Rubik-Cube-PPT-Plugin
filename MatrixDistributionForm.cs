// MatrixDistributionForm.cs
using System;
using System.Windows.Forms;

namespace 课件帮PPT助手
{
    public partial class MatrixDistributionForm : Form
    {
        public int TotalCount { get; private set; }
        public int HorizontalCount { get; private set; }
        public int RowSpacing { get; private set; }
        public int ColumnSpacing { get; private set; }
        public new int Scale { get; private set; }  // 使用 new 关键字

        public event EventHandler ParametersChanged;

        public MatrixDistributionForm()
        {
            InitializeComponent();
        }

        private void totalCountTrackBar_Scroll(object sender, EventArgs e)
        {
            totalCountNumericUpDown.Value = totalCountTrackBar.Value;
            OnParametersChanged();
        }

        private void totalCountNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            totalCountTrackBar.Value = (int)totalCountNumericUpDown.Value;
            OnParametersChanged();
        }

        private void horizontalCountTrackBar_Scroll(object sender, EventArgs e)
        {
            horizontalCountNumericUpDown.Value = horizontalCountTrackBar.Value;
            OnParametersChanged();
        }

        private void horizontalCountNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            horizontalCountTrackBar.Value = (int)horizontalCountNumericUpDown.Value;
            OnParametersChanged();
        }

        private void rowSpacingTrackBar_Scroll(object sender, EventArgs e)
        {
            rowSpacingNumericUpDown.Value = rowSpacingTrackBar.Value;
            OnParametersChanged();
        }

        private void rowSpacingNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            rowSpacingTrackBar.Value = (int)rowSpacingNumericUpDown.Value;
            OnParametersChanged();
        }

        private void columnSpacingTrackBar_Scroll(object sender, EventArgs e)
        {
            columnSpacingNumericUpDown.Value = columnSpacingTrackBar.Value;
            OnParametersChanged();
        }

        private void columnSpacingNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            columnSpacingTrackBar.Value = (int)columnSpacingNumericUpDown.Value;
            OnParametersChanged();
        }

        private void scaleTrackBar_Scroll(object sender, EventArgs e)
        {
            scaleNumericUpDown.Value = scaleTrackBar.Value;
            OnParametersChanged();
        }

        private void scaleNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            scaleTrackBar.Value = (int)scaleNumericUpDown.Value;
            OnParametersChanged();
        }

        protected virtual void OnParametersChanged()
        {
            TotalCount = (int)totalCountNumericUpDown.Value;
            HorizontalCount = (int)horizontalCountNumericUpDown.Value;
            RowSpacing = (int)rowSpacingNumericUpDown.Value;
            ColumnSpacing = (int)columnSpacingNumericUpDown.Value;
            Scale = (int)scaleNumericUpDown.Value;

            ParametersChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}
