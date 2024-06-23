namespace 课件帮PPT助手
{
    partial class SampleGenerationForm
    {
        private System.ComponentModel.IContainer components = null;

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SampleGenerationForm));
            this.checkBoxSelectedSlides = new System.Windows.Forms.CheckBox();
            this.checkBoxAllSlides = new System.Windows.Forms.CheckBox();
            this.buttonGenerate = new System.Windows.Forms.Button();
            this.pictureBoxStyle2 = new System.Windows.Forms.PictureBox();
            this.pictureBoxStyle1 = new System.Windows.Forms.PictureBox();
            this.pictureBoxStyle3 = new System.Windows.Forms.PictureBox();
            this.pictureBoxStyle4 = new System.Windows.Forms.PictureBox();
            this.pictureBoxStyle5 = new System.Windows.Forms.PictureBox();
            this.pictureBoxStyle6 = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.groupBoxTextAttributes = new System.Windows.Forms.GroupBox();
            this.comboBoxResolution = new System.Windows.Forms.ComboBox();
            this.labelResolution = new System.Windows.Forms.Label();
            this.labelSelectedSlidesCount = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle6)).BeginInit();
            this.groupBoxTextAttributes.SuspendLayout();
            this.SuspendLayout();
            // 
            // checkBoxSelectedSlides
            // 
            this.checkBoxSelectedSlides.AutoSize = true;
            this.checkBoxSelectedSlides.BackColor = System.Drawing.Color.Transparent;
            this.checkBoxSelectedSlides.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(33)))), ((int)(((byte)(230)))));
            this.checkBoxSelectedSlides.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(240)))), ((int)(((byte)(254)))));
            this.checkBoxSelectedSlides.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(158)))), ((int)(((byte)(166)))), ((int)(((byte)(255)))));
            this.checkBoxSelectedSlides.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(210)))), ((int)(((byte)(214)))), ((int)(((byte)(255)))));
            this.checkBoxSelectedSlides.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxSelectedSlides.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBoxSelectedSlides.Location = new System.Drawing.Point(58, 84);
            this.checkBoxSelectedSlides.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.checkBoxSelectedSlides.Name = "checkBoxSelectedSlides";
            this.checkBoxSelectedSlides.Size = new System.Drawing.Size(137, 35);
            this.checkBoxSelectedSlides.TabIndex = 0;
            this.checkBoxSelectedSlides.Text = "所选页面";
            this.checkBoxSelectedSlides.UseVisualStyleBackColor = false;
            this.checkBoxSelectedSlides.CheckedChanged += new System.EventHandler(this.CheckBoxSelectedSlides_CheckedChanged);
            // 
            // checkBoxAllSlides
            // 
            this.checkBoxAllSlides.AutoSize = true;
            this.checkBoxAllSlides.BackColor = System.Drawing.Color.Transparent;
            this.checkBoxAllSlides.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(13)))), ((int)(((byte)(33)))), ((int)(((byte)(230)))));
            this.checkBoxAllSlides.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(240)))), ((int)(((byte)(254)))));
            this.checkBoxAllSlides.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(158)))), ((int)(((byte)(166)))), ((int)(((byte)(255)))));
            this.checkBoxAllSlides.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(210)))), ((int)(((byte)(214)))), ((int)(((byte)(255)))));
            this.checkBoxAllSlides.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxAllSlides.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBoxAllSlides.Location = new System.Drawing.Point(222, 84);
            this.checkBoxAllSlides.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.checkBoxAllSlides.Name = "checkBoxAllSlides";
            this.checkBoxAllSlides.Size = new System.Drawing.Size(137, 35);
            this.checkBoxAllSlides.TabIndex = 1;
            this.checkBoxAllSlides.Text = "全部页面";
            this.checkBoxAllSlides.UseVisualStyleBackColor = false;
            this.checkBoxAllSlides.CheckedChanged += new System.EventHandler(this.CheckBoxAllSlides_CheckedChanged);
            // 
            // buttonGenerate
            // 
            this.buttonGenerate.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(51)))), ((int)(((byte)(242)))));
            this.buttonGenerate.FlatAppearance.BorderSize = 0;
            this.buttonGenerate.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(9)))), ((int)(((byte)(9)))), ((int)(((byte)(211)))));
            this.buttonGenerate.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(91)))), ((int)(((byte)(91)))), ((int)(((byte)(243)))));
            this.buttonGenerate.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonGenerate.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.buttonGenerate.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.buttonGenerate.Location = new System.Drawing.Point(300, 530);
            this.buttonGenerate.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.buttonGenerate.Name = "buttonGenerate";
            this.buttonGenerate.Size = new System.Drawing.Size(200, 61);
            this.buttonGenerate.TabIndex = 4;
            this.buttonGenerate.Text = "生成样机展示";
            this.buttonGenerate.UseVisualStyleBackColor = false;
            this.buttonGenerate.Click += new System.EventHandler(this.ButtonGenerate_Click);
            // 
            // pictureBoxStyle2
            // 
            this.pictureBoxStyle2.BackgroundImage = global::课件帮PPT助手.Properties.Resources.原始样机2;
            this.pictureBoxStyle2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBoxStyle2.Location = new System.Drawing.Point(300, 179);
            this.pictureBoxStyle2.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.pictureBoxStyle2.Name = "pictureBoxStyle2";
            this.pictureBoxStyle2.Size = new System.Drawing.Size(198, 115);
            this.pictureBoxStyle2.TabIndex = 3;
            this.pictureBoxStyle2.TabStop = false;
            this.pictureBoxStyle2.Click += new System.EventHandler(this.PictureBoxStyle2_Click);
            // 
            // pictureBoxStyle1
            // 
            this.pictureBoxStyle1.BackgroundImage = global::课件帮PPT助手.Properties.Resources.原始样机1;
            this.pictureBoxStyle1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBoxStyle1.Location = new System.Drawing.Point(60, 179);
            this.pictureBoxStyle1.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.pictureBoxStyle1.Name = "pictureBoxStyle1";
            this.pictureBoxStyle1.Size = new System.Drawing.Size(198, 115);
            this.pictureBoxStyle1.TabIndex = 2;
            this.pictureBoxStyle1.TabStop = false;
            this.pictureBoxStyle1.Click += new System.EventHandler(this.PictureBoxStyle1_Click);
            // 
            // pictureBoxStyle3
            // 
            this.pictureBoxStyle3.BackgroundImage = global::课件帮PPT助手.Properties.Resources.原始样机3;
            this.pictureBoxStyle3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBoxStyle3.Location = new System.Drawing.Point(492, 45);
            this.pictureBoxStyle3.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.pictureBoxStyle3.Name = "pictureBoxStyle3";
            this.pictureBoxStyle3.Size = new System.Drawing.Size(198, 115);
            this.pictureBoxStyle3.TabIndex = 9;
            this.pictureBoxStyle3.TabStop = false;
            this.pictureBoxStyle3.Click += new System.EventHandler(this.PictureBoxStyle3_Click);
            // 
            // pictureBoxStyle4
            // 
            this.pictureBoxStyle4.BackgroundImage = global::课件帮PPT助手.Properties.Resources.原始样机4;
            this.pictureBoxStyle4.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBoxStyle4.Location = new System.Drawing.Point(21, 206);
            this.pictureBoxStyle4.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.pictureBoxStyle4.Name = "pictureBoxStyle4";
            this.pictureBoxStyle4.Size = new System.Drawing.Size(198, 115);
            this.pictureBoxStyle4.TabIndex = 10;
            this.pictureBoxStyle4.TabStop = false;
            this.pictureBoxStyle4.Click += new System.EventHandler(this.PictureBoxStyle4_Click);
            // 
            // pictureBoxStyle5
            // 
            this.pictureBoxStyle5.BackgroundImage = global::课件帮PPT助手.Properties.Resources.原始样机5;
            this.pictureBoxStyle5.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBoxStyle5.Location = new System.Drawing.Point(261, 206);
            this.pictureBoxStyle5.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.pictureBoxStyle5.Name = "pictureBoxStyle5";
            this.pictureBoxStyle5.Size = new System.Drawing.Size(198, 115);
            this.pictureBoxStyle5.TabIndex = 11;
            this.pictureBoxStyle5.TabStop = false;
            this.pictureBoxStyle5.Click += new System.EventHandler(this.PictureBoxStyle5_Click);
            // 
            // pictureBoxStyle6
            // 
            this.pictureBoxStyle6.BackgroundImage = global::课件帮PPT助手.Properties.Resources.原始样机6;
            this.pictureBoxStyle6.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBoxStyle6.Location = new System.Drawing.Point(492, 206);
            this.pictureBoxStyle6.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.pictureBoxStyle6.Name = "pictureBoxStyle6";
            this.pictureBoxStyle6.Size = new System.Drawing.Size(198, 115);
            this.pictureBoxStyle6.TabIndex = 12;
            this.pictureBoxStyle6.TabStop = false;
            this.pictureBoxStyle6.Click += new System.EventHandler(this.PictureBoxStyle6_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(116, 299);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 24);
            this.label1.TabIndex = 5;
            this.label1.Text = "样式1";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(357, 299);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(70, 24);
            this.label2.TabIndex = 6;
            this.label2.Text = "样式2";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(548, 164);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(70, 24);
            this.label3.TabIndex = 13;
            this.label3.Text = "样式3";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(78, 326);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(70, 24);
            this.label4.TabIndex = 14;
            this.label4.Text = "样式4";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.Location = new System.Drawing.Point(317, 326);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(70, 24);
            this.label5.TabIndex = 15;
            this.label5.Text = "样式5";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.Location = new System.Drawing.Point(549, 326);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(70, 24);
            this.label6.TabIndex = 16;
            this.label6.Text = "样式6";
            // 
            // groupBoxTextAttributes
            // 
            this.groupBoxTextAttributes.Controls.Add(this.pictureBoxStyle3);
            this.groupBoxTextAttributes.Controls.Add(this.label3);
            this.groupBoxTextAttributes.Controls.Add(this.label6);
            this.groupBoxTextAttributes.Controls.Add(this.pictureBoxStyle4);
            this.groupBoxTextAttributes.Controls.Add(this.pictureBoxStyle6);
            this.groupBoxTextAttributes.Controls.Add(this.label5);
            this.groupBoxTextAttributes.Controls.Add(this.label4);
            this.groupBoxTextAttributes.Controls.Add(this.pictureBoxStyle5);
            this.groupBoxTextAttributes.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBoxTextAttributes.Location = new System.Drawing.Point(39, 134);
            this.groupBoxTextAttributes.Name = "groupBoxTextAttributes";
            this.groupBoxTextAttributes.Size = new System.Drawing.Size(719, 378);
            this.groupBoxTextAttributes.TabIndex = 7;
            this.groupBoxTextAttributes.TabStop = false;
            this.groupBoxTextAttributes.Text = "样机样式";
            // 
            // comboBoxResolution
            // 
            this.comboBoxResolution.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxResolution.Font = new System.Drawing.Font("宋体", 9F);
            this.comboBoxResolution.FormattingEnabled = true;
            this.comboBoxResolution.Items.AddRange(new object[] {
            "720x480 (标清)",
            "1280x720 (高清)",
            "1920x1080 (全高清)",
            "2048x1080 (2K)",
            "3840x2160 (超高清)",
            "4096x2160 (4K)",
            "7680x4320 (8K)"});
            this.comboBoxResolution.Location = new System.Drawing.Point(489, 87);
            this.comboBoxResolution.Name = "comboBoxResolution";
            this.comboBoxResolution.Size = new System.Drawing.Size(267, 32);
            this.comboBoxResolution.TabIndex = 17;
            // 
            // labelResolution
            // 
            this.labelResolution.AutoSize = true;
            this.labelResolution.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelResolution.Location = new System.Drawing.Point(403, 86);
            this.labelResolution.Name = "labelResolution";
            this.labelResolution.Size = new System.Drawing.Size(86, 31);
            this.labelResolution.TabIndex = 18;
            this.labelResolution.Text = "分辨率";
            // 
            // labelSelectedSlidesCount
            // 
            this.labelSelectedSlidesCount.AutoSize = true;
            this.labelSelectedSlidesCount.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelSelectedSlidesCount.Location = new System.Drawing.Point(54, 35);
            this.labelSelectedSlidesCount.Name = "labelSelectedSlidesCount";
            this.labelSelectedSlidesCount.Size = new System.Drawing.Size(238, 24);
            this.labelSelectedSlidesCount.TabIndex = 19;
            this.labelSelectedSlidesCount.Text = "已选中幻灯片数量：0";
            // 
            // SampleGenerationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(802, 624);
            this.Controls.Add(this.comboBoxResolution);
            this.Controls.Add(this.labelSelectedSlidesCount);
            this.Controls.Add(this.buttonGenerate);
            this.Controls.Add(this.labelResolution);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.pictureBoxStyle2);
            this.Controls.Add(this.pictureBoxStyle1);
            this.Controls.Add(this.checkBoxAllSlides);
            this.Controls.Add(this.checkBoxSelectedSlides);
            this.Controls.Add(this.groupBoxTextAttributes);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.Name = "SampleGenerationForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "生成样机";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle6)).EndInit();
            this.groupBoxTextAttributes.ResumeLayout(false);
            this.groupBoxTextAttributes.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.Windows.Forms.CheckBox checkBoxSelectedSlides;
        private System.Windows.Forms.CheckBox checkBoxAllSlides;
        private System.Windows.Forms.PictureBox pictureBoxStyle1;
        private System.Windows.Forms.PictureBox pictureBoxStyle2;
        private System.Windows.Forms.PictureBox pictureBoxStyle3;
        private System.Windows.Forms.PictureBox pictureBoxStyle4;
        private System.Windows.Forms.PictureBox pictureBoxStyle5;
        private System.Windows.Forms.PictureBox pictureBoxStyle6;
        private System.Windows.Forms.Button buttonGenerate;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.GroupBox groupBoxTextAttributes;
        private System.Windows.Forms.ComboBox comboBoxResolution;
        private System.Windows.Forms.Label labelResolution;
        private System.Windows.Forms.Label labelSelectedSlidesCount;
    }
}
