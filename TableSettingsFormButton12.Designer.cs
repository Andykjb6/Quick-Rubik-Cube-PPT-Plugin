namespace 课件帮PPT助手
{
    partial class TableSettingsFormButton12
    {
        private System.Windows.Forms.Label labelWidth;
        private System.Windows.Forms.NumericUpDown numericUpDownBorderWidth;
        private System.Windows.Forms.Label labelColor;
        private System.Windows.Forms.Button buttonChooseColor;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.Button buttonApply;

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TableSettingsFormButton12));
            this.labelWidth = new System.Windows.Forms.Label();
            this.numericUpDownBorderWidth = new System.Windows.Forms.NumericUpDown();
            this.labelColor = new System.Windows.Forms.Label();
            this.buttonChooseColor = new System.Windows.Forms.Button();
            this.buttonOK = new System.Windows.Forms.Button();
            this.buttonApply = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownBorderWidth)).BeginInit();
            this.SuspendLayout();
            // 
            // labelWidth
            // 
            this.labelWidth.AutoSize = true;
            this.labelWidth.Font = new System.Drawing.Font("微软雅黑", 10.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelWidth.Location = new System.Drawing.Point(38, 40);
            this.labelWidth.Name = "labelWidth";
            this.labelWidth.Size = new System.Drawing.Size(140, 38);
            this.labelWidth.TabIndex = 0;
            this.labelWidth.Text = "边框宽度:";
            // 
            // numericUpDownBorderWidth
            // 
            this.numericUpDownBorderWidth.DecimalPlaces = 2;
            this.numericUpDownBorderWidth.Font = new System.Drawing.Font("宋体", 11F);
            this.numericUpDownBorderWidth.Location = new System.Drawing.Point(192, 40);
            this.numericUpDownBorderWidth.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numericUpDownBorderWidth.Name = "numericUpDownBorderWidth";
            this.numericUpDownBorderWidth.Size = new System.Drawing.Size(216, 41);
            this.numericUpDownBorderWidth.TabIndex = 1;
            this.numericUpDownBorderWidth.Value = new decimal(new int[] {
            125,
            0,
            0,
            131072});
            // 
            // labelColor
            // 
            this.labelColor.AutoSize = true;
            this.labelColor.Font = new System.Drawing.Font("微软雅黑", 10.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelColor.Location = new System.Drawing.Point(38, 90);
            this.labelColor.Name = "labelColor";
            this.labelColor.Size = new System.Drawing.Size(140, 38);
            this.labelColor.TabIndex = 2;
            this.labelColor.Text = "边框颜色:";
            // 
            // buttonChooseColor
            // 
            this.buttonChooseColor.Font = new System.Drawing.Font("宋体", 9F);
            this.buttonChooseColor.Location = new System.Drawing.Point(192, 90);
            this.buttonChooseColor.Name = "buttonChooseColor";
            this.buttonChooseColor.Size = new System.Drawing.Size(216, 40);
            this.buttonChooseColor.TabIndex = 3;
            this.buttonChooseColor.Text = "自定义";
            this.buttonChooseColor.Click += new System.EventHandler(this.ButtonChooseColor_Click);
            // 
            // buttonOK
            // 
            this.buttonOK.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(89)))), ((int)(((byte)(239)))));
            this.buttonOK.Font = new System.Drawing.Font("微软雅黑", 10.875F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.buttonOK.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.buttonOK.Location = new System.Drawing.Point(116, 155);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(100, 55);
            this.buttonOK.TabIndex = 4;
            this.buttonOK.Text = "生成";
            this.buttonOK.UseVisualStyleBackColor = false;
            this.buttonOK.Click += new System.EventHandler(this.ButtonOK_Click);
            // 
            // buttonApply
            // 
            this.buttonApply.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(89)))), ((int)(((byte)(239)))));
            this.buttonApply.Font = new System.Drawing.Font("微软雅黑", 10.875F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.buttonApply.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.buttonApply.Location = new System.Drawing.Point(241, 155);
            this.buttonApply.Name = "buttonApply";
            this.buttonApply.Size = new System.Drawing.Size(100, 55);
            this.buttonApply.TabIndex = 5;
            this.buttonApply.Text = "应用";
            this.buttonApply.UseVisualStyleBackColor = false;
            this.buttonApply.Click += new System.EventHandler(this.ButtonApply_Click);
            // 
            // TableSettingsFormButton12
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(192F, 192F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(463, 245);
            this.Controls.Add(this.labelWidth);
            this.Controls.Add(this.numericUpDownBorderWidth);
            this.Controls.Add(this.labelColor);
            this.Controls.Add(this.buttonChooseColor);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.buttonApply);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "TableSettingsFormButton12";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "生字赋格";
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownBorderWidth)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
    }
}
