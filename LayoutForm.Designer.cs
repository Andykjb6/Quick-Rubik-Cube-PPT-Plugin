using System.Windows.Forms;

namespace 课件帮PPT助手
{
    partial class LayoutForm
    {
        private System.ComponentModel.IContainer components = null;
        private NumericUpDown numericUpDownDistance;
        private TrackBar trackBarCompactness;
        private NumericUpDown numericUpDownStartAngle;
        private ComboBox comboBoxDirection; // 新增方向选择 ComboBox
        private Label label1;
        private Label label2;
        private Label label3;
        private Label labelDirection; // 新增方向标签

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(LayoutForm));
            this.numericUpDownDistance = new System.Windows.Forms.NumericUpDown();
            this.trackBarCompactness = new System.Windows.Forms.TrackBar();
            this.numericUpDownStartAngle = new System.Windows.Forms.NumericUpDown();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.comboBoxDirection = new System.Windows.Forms.ComboBox();
            this.labelDirection = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownDistance)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.trackBarCompactness)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownStartAngle)).BeginInit();
            this.SuspendLayout();
            // 
            // numericUpDownDistance
            // 
            this.numericUpDownDistance.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.numericUpDownDistance.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.numericUpDownDistance.Location = new System.Drawing.Point(123, 32);
            this.numericUpDownDistance.Maximum = new decimal(new int[] {
            300,
            0,
            0,
            0});
            this.numericUpDownDistance.Minimum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numericUpDownDistance.Name = "numericUpDownDistance";
            this.numericUpDownDistance.Size = new System.Drawing.Size(120, 44);
            this.numericUpDownDistance.TabIndex = 0;
            this.numericUpDownDistance.Value = new decimal(new int[] {
            100,
            0,
            0,
            0});
            this.numericUpDownDistance.ValueChanged += new System.EventHandler(this.OnValueChanged);
            // 
            // trackBarCompactness
            // 
            this.trackBarCompactness.Location = new System.Drawing.Point(666, 23);
            this.trackBarCompactness.Maximum = 100;
            this.trackBarCompactness.Minimum = 10;
            this.trackBarCompactness.Name = "trackBarCompactness";
            this.trackBarCompactness.Size = new System.Drawing.Size(255, 90);
            this.trackBarCompactness.TabIndex = 1;
            this.trackBarCompactness.TickStyle = System.Windows.Forms.TickStyle.TopLeft;
            this.trackBarCompactness.Value = 50;
            this.trackBarCompactness.ValueChanged += new System.EventHandler(this.OnValueChanged);
            // 
            // numericUpDownStartAngle
            // 
            this.numericUpDownStartAngle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.numericUpDownStartAngle.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.numericUpDownStartAngle.Location = new System.Drawing.Point(410, 32);
            this.numericUpDownStartAngle.Maximum = new decimal(new int[] {
            360,
            0,
            0,
            0});
            this.numericUpDownStartAngle.Name = "numericUpDownStartAngle";
            this.numericUpDownStartAngle.Size = new System.Drawing.Size(120, 44);
            this.numericUpDownStartAngle.TabIndex = 2;
            this.numericUpDownStartAngle.ValueChanged += new System.EventHandler(this.OnValueChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(33)))), ((int)(((byte)(33)))), ((int)(((byte)(213)))));
            this.label1.Location = new System.Drawing.Point(34, 34);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(101, 37);
            this.label1.TabIndex = 3;
            this.label1.Text = "半径：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold);
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(33)))), ((int)(((byte)(33)))), ((int)(((byte)(213)))));
            this.label2.Location = new System.Drawing.Point(269, 34);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(157, 37);
            this.label2.TabIndex = 4;
            this.label2.Text = "起始角度：";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold);
            this.label3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(33)))), ((int)(((byte)(33)))), ((int)(((byte)(213)))));
            this.label3.Location = new System.Drawing.Point(561, 34);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(129, 37);
            this.label3.TabIndex = 5;
            this.label3.Text = "紧凑度：";
            // 
            // comboBoxDirection
            // 
            this.comboBoxDirection.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxDirection.Font = new System.Drawing.Font("微软雅黑", 10.5F);
            this.comboBoxDirection.FormattingEnabled = true;
            this.comboBoxDirection.Items.AddRange(new object[] {
            "顺时针",
            "逆时针"});
            this.comboBoxDirection.Location = new System.Drawing.Point(1015, 32);
            this.comboBoxDirection.Name = "comboBoxDirection";
            this.comboBoxDirection.Size = new System.Drawing.Size(150, 44);
            this.comboBoxDirection.TabIndex = 6;
            // 设置 ComboBox 的默认选项为 "顺时针"
            this.comboBoxDirection.SelectedIndex = 0; // 0 对应 "顺时针"
            this.comboBoxDirection.SelectedIndexChanged += new System.EventHandler(this.OnDirectionChanged);
            // 
            // labelDirection
            // 
            this.labelDirection.AutoSize = true;
            this.labelDirection.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold);
            this.labelDirection.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(33)))), ((int)(((byte)(33)))), ((int)(((byte)(213)))));
            this.labelDirection.Location = new System.Drawing.Point(927, 36);
            this.labelDirection.Name = "labelDirection";
            this.labelDirection.Size = new System.Drawing.Size(101, 37);
            this.labelDirection.TabIndex = 7;
            this.labelDirection.Text = "方向：";
            // 
            // LayoutForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(192F, 192F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(238)))), ((int)(((byte)(254)))));
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(1217, 121);
            this.Controls.Add(this.trackBarCompactness);
            this.Controls.Add(this.numericUpDownDistance);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.numericUpDownStartAngle);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.comboBoxDirection);
            this.Controls.Add(this.labelDirection);
            this.DoubleBuffered = true;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "LayoutForm";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.TopMost = true;
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownDistance)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.trackBarCompactness)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownStartAngle)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
    }
}
