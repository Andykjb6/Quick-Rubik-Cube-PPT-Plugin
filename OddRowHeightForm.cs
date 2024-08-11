using System;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace 课件帮PPT助手
{
    public partial class OddRowHeightForm : Form
    {
        private PowerPoint.Table table;
        private float originalHeightCm;

        public OddRowHeightForm(PowerPoint.Table selectedTable)
        {
            InitializeComponent();

            this.table = selectedTable;
            this.originalHeightCm = PointsToCm(selectedTable.Rows[1].Height); // 将第一行高度从点转换为厘米

            // 仅调整非第一行的其他奇数行的高度，使其与第一行一致
            float firstRowHeightPt = selectedTable.Rows[1].Height;
            for (int i = 3; i <= table.Rows.Count; i += 2)
            {
                table.Rows[i].Height = firstRowHeightPt;
            }

            // 设置NumericUpDown的范围
            this.numericUpDown.Minimum = 0.1M; // 设置一个合理的最小值
            this.numericUpDown.Maximum = 50.0M; // 设置一个合理的最大值

            // 确保originalHeightCm在范围内
            if ((decimal)this.originalHeightCm < this.numericUpDown.Minimum)
            {
                this.numericUpDown.Value = this.numericUpDown.Minimum;
            }
            else if ((decimal)this.originalHeightCm > this.numericUpDown.Maximum)
            {
                this.numericUpDown.Value = this.numericUpDown.Maximum;
            }
            else
            {
                this.numericUpDown.Value = (decimal)this.originalHeightCm;
            }

            // 事件绑定
            this.numericUpDown.ValueChanged += NumericUpDown_ValueChanged;
            this.btnOK.Click += BtnOK_Click;
            this.btnCancel.Click += BtnCancel_Click;
        }

        private void NumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            float newHeightCm = (float)this.numericUpDown.Value;
            float newHeightPt = CmToPoints(newHeightCm); // 将厘米转换为点

            // 动态预览非第一行的其他奇数行高度变化
            for (int i = 3; i <= table.Rows.Count; i += 2)
            {
                table.Rows[i].Height = newHeightPt;
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            float originalHeightPt = CmToPoints(originalHeightCm); // 将厘米转换为点

            // 恢复非第一行的其他奇数行的原始高度
            for (int i = 3; i <= table.Rows.Count; i += 2)
            {
                table.Rows[i].Height = originalHeightPt;
            }
            this.Close();
        }

        private float PointsToCm(float points)
        {
            return points / 28.3465f; // 1厘米 = 28.3465点
        }

        private float CmToPoints(float cm)
        {
            return cm * 28.3465f;
        }
    }
}
