namespace 课件帮PPT助手
{
    partial class TableForm
    {
        private System.Windows.Forms.Label labelRows;
        private System.Windows.Forms.NumericUpDown numericUpDownRows;
        private System.Windows.Forms.Label labelColumns;
        private System.Windows.Forms.NumericUpDown numericUpDownColumns;
        private System.Windows.Forms.Label labelRowSpacing;
        private System.Windows.Forms.NumericUpDown numericUpDownRowSpacing;
        private System.Windows.Forms.Label labelColumnSpacing;
        private System.Windows.Forms.NumericUpDown numericUpDownColumnSpacing;
        private System.Windows.Forms.Label labelWidth;
        private System.Windows.Forms.NumericUpDown numericUpDownBorderWidth;
        private System.Windows.Forms.Label labelScale;
        private System.Windows.Forms.TrackBar trackBarScale;
        private System.Windows.Forms.Label labelColor;
        private System.Windows.Forms.Button buttonChooseColor;
        private System.Windows.Forms.Button buttonOK;

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TableForm));
            this.labelRows = new System.Windows.Forms.Label();
            this.numericUpDownRows = new System.Windows.Forms.NumericUpDown();
            this.labelColumns = new System.Windows.Forms.Label();
            this.numericUpDownColumns = new System.Windows.Forms.NumericUpDown();
            this.labelRowSpacing = new System.Windows.Forms.Label();
            this.numericUpDownRowSpacing = new System.Windows.Forms.NumericUpDown();
            this.labelColumnSpacing = new System.Windows.Forms.Label();
            this.numericUpDownColumnSpacing = new System.Windows.Forms.NumericUpDown();
            this.labelWidth = new System.Windows.Forms.Label();
            this.numericUpDownBorderWidth = new System.Windows.Forms.NumericUpDown();
            this.labelScale = new System.Windows.Forms.Label();
            this.trackBarScale = new System.Windows.Forms.TrackBar();
            this.labelColor = new System.Windows.Forms.Label();
            this.buttonChooseColor = new System.Windows.Forms.Button();
            this.buttonOK = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownRows)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownColumns)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownRowSpacing)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownColumnSpacing)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownBorderWidth)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.trackBarScale)).BeginInit();
            this.SuspendLayout();
            // 
            // labelRows
            // 
            this.labelRows.AutoSize = true;
            this.labelRows.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.labelRows.Location = new System.Drawing.Point(32, 29);
            this.labelRows.Name = "labelRows";
            this.labelRows.Size = new System.Drawing.Size(68, 31);
            this.labelRows.TabIndex = 0;
            this.labelRows.Text = "行数:";
            // 
            // numericUpDownRows
            // 
            this.numericUpDownRows.Font = new System.Drawing.Font("宋体", 11F);
            this.numericUpDownRows.Location = new System.Drawing.Point(175, 27);
            this.numericUpDownRows.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numericUpDownRows.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDownRows.Name = "numericUpDownRows";
            this.numericUpDownRows.Size = new System.Drawing.Size(203, 41);
            this.numericUpDownRows.TabIndex = 1;
            this.numericUpDownRows.Value = new decimal(new int[] {
            2,
            0,
            0,
            0});
            // 
            // labelColumns
            // 
            this.labelColumns.AutoSize = true;
            this.labelColumns.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.labelColumns.Location = new System.Drawing.Point(32, 74);
            this.labelColumns.Name = "labelColumns";
            this.labelColumns.Size = new System.Drawing.Size(68, 31);
            this.labelColumns.TabIndex = 2;
            this.labelColumns.Text = "列数:";
            // 
            // numericUpDownColumns
            // 
            this.numericUpDownColumns.Font = new System.Drawing.Font("宋体", 11F);
            this.numericUpDownColumns.Location = new System.Drawing.Point(175, 74);
            this.numericUpDownColumns.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numericUpDownColumns.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDownColumns.Name = "numericUpDownColumns";
            this.numericUpDownColumns.Size = new System.Drawing.Size(203, 41);
            this.numericUpDownColumns.TabIndex = 3;
            this.numericUpDownColumns.Value = new decimal(new int[] {
            2,
            0,
            0,
            0});
            // 
            // labelRowSpacing
            // 
            this.labelRowSpacing.AutoSize = true;
            this.labelRowSpacing.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.labelRowSpacing.Location = new System.Drawing.Point(32, 121);
            this.labelRowSpacing.Name = "labelRowSpacing";
            this.labelRowSpacing.Size = new System.Drawing.Size(92, 31);
            this.labelRowSpacing.TabIndex = 4;
            this.labelRowSpacing.Text = "行间距:";
            // 
            // numericUpDownRowSpacing
            // 
            this.numericUpDownRowSpacing.Font = new System.Drawing.Font("宋体", 11F);
            this.numericUpDownRowSpacing.Location = new System.Drawing.Point(175, 121);
            this.numericUpDownRowSpacing.Name = "numericUpDownRowSpacing";
            this.numericUpDownRowSpacing.Size = new System.Drawing.Size(203, 41);
            this.numericUpDownRowSpacing.TabIndex = 5;
            this.numericUpDownRowSpacing.Value = new decimal(new int[] {
            10,
            0,
            0,
            0});
            // 
            // labelColumnSpacing
            // 
            this.labelColumnSpacing.AutoSize = true;
            this.labelColumnSpacing.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.labelColumnSpacing.Location = new System.Drawing.Point(32, 169);
            this.labelColumnSpacing.Name = "labelColumnSpacing";
            this.labelColumnSpacing.Size = new System.Drawing.Size(92, 31);
            this.labelColumnSpacing.TabIndex = 6;
            this.labelColumnSpacing.Text = "列间距:";
            // 
            // numericUpDownColumnSpacing
            // 
            this.numericUpDownColumnSpacing.Font = new System.Drawing.Font("宋体", 11F);
            this.numericUpDownColumnSpacing.Location = new System.Drawing.Point(175, 168);
            this.numericUpDownColumnSpacing.Name = "numericUpDownColumnSpacing";
            this.numericUpDownColumnSpacing.Size = new System.Drawing.Size(203, 41);
            this.numericUpDownColumnSpacing.TabIndex = 7;
            this.numericUpDownColumnSpacing.Value = new decimal(new int[] {
            10,
            0,
            0,
            0});
            // 
            // labelWidth
            // 
            this.labelWidth.AutoSize = true;
            this.labelWidth.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.labelWidth.Location = new System.Drawing.Point(32, 217);
            this.labelWidth.Name = "labelWidth";
            this.labelWidth.Size = new System.Drawing.Size(116, 31);
            this.labelWidth.TabIndex = 8;
            this.labelWidth.Text = "边框宽度:";
            // 
            // numericUpDownBorderWidth
            // 
            this.numericUpDownBorderWidth.DecimalPlaces = 2;
            this.numericUpDownBorderWidth.Font = new System.Drawing.Font("宋体", 11F);
            this.numericUpDownBorderWidth.Location = new System.Drawing.Point(175, 215);
            this.numericUpDownBorderWidth.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numericUpDownBorderWidth.Name = "numericUpDownBorderWidth";
            this.numericUpDownBorderWidth.Size = new System.Drawing.Size(203, 41);
            this.numericUpDownBorderWidth.TabIndex = 9;
            this.numericUpDownBorderWidth.Value = new decimal(new int[] {
            125,
            0,
            0,
            131072});
            // 
            // labelScale
            // 
            this.labelScale.AutoSize = true;
            this.labelScale.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.labelScale.Location = new System.Drawing.Point(32, 317);
            this.labelScale.Name = "labelScale";
            this.labelScale.Size = new System.Drawing.Size(116, 31);
            this.labelScale.TabIndex = 10;
            this.labelScale.Text = "缩放比例:";
            // 
            // trackBarScale
            // 
            this.trackBarScale.Location = new System.Drawing.Point(175, 323);
            this.trackBarScale.Maximum = 200;
            this.trackBarScale.Minimum = 50;
            this.trackBarScale.Name = "trackBarScale";
            this.trackBarScale.Size = new System.Drawing.Size(209, 90);
            this.trackBarScale.TabIndex = 11;
            this.trackBarScale.TickFrequency = 10;
            this.trackBarScale.Value = 100;
            // 
            // labelColor
            // 
            this.labelColor.AutoSize = true;
            this.labelColor.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.labelColor.Location = new System.Drawing.Point(32, 265);
            this.labelColor.Name = "labelColor";
            this.labelColor.Size = new System.Drawing.Size(116, 31);
            this.labelColor.TabIndex = 12;
            this.labelColor.Text = "边框颜色:";
            // 
            // buttonChooseColor
            // 
            this.buttonChooseColor.Location = new System.Drawing.Point(175, 262);
            this.buttonChooseColor.Name = "buttonChooseColor";
            this.buttonChooseColor.Size = new System.Drawing.Size(203, 40);
            this.buttonChooseColor.TabIndex = 13;
            this.buttonChooseColor.Text = "自定义";
            this.buttonChooseColor.Click += new System.EventHandler(this.ButtonChooseColor_Click);
            // 
            // buttonOK
            // 
            this.buttonOK.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(89)))), ((int)(((byte)(239)))));
            this.buttonOK.Font = new System.Drawing.Font("微软雅黑", 11F, System.Drawing.FontStyle.Bold);
            this.buttonOK.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.buttonOK.Location = new System.Drawing.Point(26, 372);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(143, 59);
            this.buttonOK.TabIndex = 14;
            this.buttonOK.Text = "生成";
            this.buttonOK.UseVisualStyleBackColor = false;
            this.buttonOK.Click += new System.EventHandler(this.ButtonOK_Click);
            // 
            // TableSettingsForm
            // 
            this.ClientSize = new System.Drawing.Size(419, 465);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.labelRows);
            this.Controls.Add(this.numericUpDownRows);
            this.Controls.Add(this.labelColumns);
            this.Controls.Add(this.numericUpDownColumns);
            this.Controls.Add(this.labelRowSpacing);
            this.Controls.Add(this.numericUpDownRowSpacing);
            this.Controls.Add(this.labelColumnSpacing);
            this.Controls.Add(this.numericUpDownColumnSpacing);
            this.Controls.Add(this.labelWidth);
            this.Controls.Add(this.numericUpDownBorderWidth);
            this.Controls.Add(this.labelScale);
            this.Controls.Add(this.trackBarScale);
            this.Controls.Add(this.labelColor);
            this.Controls.Add(this.buttonChooseColor);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "TableSettingsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "田字格";
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownRows)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownColumns)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownRowSpacing)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownColumnSpacing)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownBorderWidth)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.trackBarScale)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
    }
}
