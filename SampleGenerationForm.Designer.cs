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
            this.pictureBoxStyle1 = new System.Windows.Forms.PictureBox();
            this.pictureBoxStyle2 = new System.Windows.Forms.PictureBox();
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
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle6)).BeginInit();
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
            this.checkBoxSelectedSlides.Location = new System.Drawing.Point(60, 55);
            this.checkBoxSelectedSlides.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.checkBoxSelectedSlides.Name = "checkBoxSelectedSlides";
            this.checkBoxSelectedSlides.Size = new System.Drawing.Size(137, 35);
            this.checkBoxSelectedSlides.TabIndex = 0;
            this.checkBoxSelectedSlides.Text = "所选页面";
            this.checkBoxSelectedSlides.UseVisualStyleBackColor = false;
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
            this.checkBoxAllSlides.Location = new System.Drawing.Point(300, 55);
            this.checkBoxAllSlides.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.checkBoxAllSlides.Name = "checkBoxAllSlides";
            this.checkBoxAllSlides.Size = new System.Drawing.Size(137, 35);
            this.checkBoxAllSlides.TabIndex = 1;
            this.checkBoxAllSlides.Text = "全部页面";
            this.checkBoxAllSlides.UseVisualStyleBackColor = false;
            // 
            // buttonGenerate
            // 
            this.buttonGenerate.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(20)))), ((int)(((byte)(51)))), ((int)(((byte)(242)))));
            this.buttonGenerate.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.buttonGenerate.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.buttonGenerate.Location = new System.Drawing.Point(177, 680);
            this.buttonGenerate.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.buttonGenerate.Name = "buttonGenerate";
            this.buttonGenerate.Size = new System.Drawing.Size(200, 61);
            this.buttonGenerate.TabIndex = 4;
            this.buttonGenerate.Text = "生成样机展示";
            this.buttonGenerate.UseVisualStyleBackColor = false;
            this.buttonGenerate.Click += new System.EventHandler(this.ButtonGenerate_Click);
            // 
            // pictureBoxStyle1
            // 
            this.pictureBoxStyle1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBoxStyle1.BackgroundImage")));
            this.pictureBoxStyle1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBoxStyle1.Location = new System.Drawing.Point(60, 158);
            this.pictureBoxStyle1.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.pictureBoxStyle1.Name = "pictureBoxStyle1";
            this.pictureBoxStyle1.Size = new System.Drawing.Size(198, 115);
            this.pictureBoxStyle1.TabIndex = 2;
            this.pictureBoxStyle1.TabStop = false;
            this.pictureBoxStyle1.Click += new System.EventHandler(this.PictureBoxStyle1_Click);
            // 
            // pictureBoxStyle2
            // 
            this.pictureBoxStyle2.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBoxStyle2.BackgroundImage")));
            this.pictureBoxStyle2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBoxStyle2.Location = new System.Drawing.Point(300, 158);
            this.pictureBoxStyle2.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.pictureBoxStyle2.Name = "pictureBoxStyle2";
            this.pictureBoxStyle2.Size = new System.Drawing.Size(198, 115);
            this.pictureBoxStyle2.TabIndex = 3;
            this.pictureBoxStyle2.TabStop = false;
            this.pictureBoxStyle2.Click += new System.EventHandler(this.PictureBoxStyle2_Click);
            // 
            // pictureBoxStyle3
            // 
            this.pictureBoxStyle3.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBoxStyle3.BackgroundImage")));
            this.pictureBoxStyle3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBoxStyle3.Location = new System.Drawing.Point(60, 320);
            this.pictureBoxStyle3.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.pictureBoxStyle3.Name = "pictureBoxStyle3";
            this.pictureBoxStyle3.Size = new System.Drawing.Size(198, 115);
            this.pictureBoxStyle3.TabIndex = 8;
            this.pictureBoxStyle3.TabStop = false;
            this.pictureBoxStyle3.Click += new System.EventHandler(this.PictureBoxStyle3_Click);
            // 
            // pictureBoxStyle4
            // 
            this.pictureBoxStyle4.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBoxStyle4.BackgroundImage")));
            this.pictureBoxStyle4.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBoxStyle4.Location = new System.Drawing.Point(300, 320);
            this.pictureBoxStyle4.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.pictureBoxStyle4.Name = "pictureBoxStyle4";
            this.pictureBoxStyle4.Size = new System.Drawing.Size(198, 115);
            this.pictureBoxStyle4.TabIndex = 9;
            this.pictureBoxStyle4.TabStop = false;
            this.pictureBoxStyle4.Click += new System.EventHandler(this.PictureBoxStyle4_Click);
            // 
            // pictureBoxStyle5
            // 
            this.pictureBoxStyle5.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBoxStyle5.BackgroundImage")));
            this.pictureBoxStyle5.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBoxStyle5.Location = new System.Drawing.Point(60, 480);
            this.pictureBoxStyle5.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.pictureBoxStyle5.Name = "pictureBoxStyle5";
            this.pictureBoxStyle5.Size = new System.Drawing.Size(198, 115);
            this.pictureBoxStyle5.TabIndex = 10;
            this.pictureBoxStyle5.TabStop = false;
            this.pictureBoxStyle5.Click += new System.EventHandler(this.PictureBoxStyle5_Click);
            // 
            // pictureBoxStyle6
            // 
            this.pictureBoxStyle6.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBoxStyle6.BackgroundImage")));
            this.pictureBoxStyle6.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBoxStyle6.Location = new System.Drawing.Point(300, 480);
            this.pictureBoxStyle6.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.pictureBoxStyle6.Name = "pictureBoxStyle6";
            this.pictureBoxStyle6.Size = new System.Drawing.Size(198, 115);
            this.pictureBoxStyle6.TabIndex = 11;
            this.pictureBoxStyle6.TabStop = false;
            this.pictureBoxStyle6.Click += new System.EventHandler(this.PictureBoxStyle6_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(116, 279);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 24);
            this.label1.TabIndex = 5;
            this.label1.Text = "样式1";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(357, 279);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(70, 24);
            this.label2.TabIndex = 6;
            this.label2.Text = "样式2";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(116, 438);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(70, 24);
            this.label3.TabIndex = 12;
            this.label3.Text = "样式3";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(357, 438);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(70, 24);
            this.label4.TabIndex = 13;
            this.label4.Text = "样式4";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(116, 598);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(70, 24);
            this.label5.TabIndex = 14;
            this.label5.Text = "样式5";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(357, 598);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(70, 24);
            this.label6.TabIndex = 15;
            this.label6.Text = "样式6";
            // 
            // groupBoxTextAttributes
            // 
            this.groupBoxTextAttributes.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBoxTextAttributes.Location = new System.Drawing.Point(39, 113);
            this.groupBoxTextAttributes.Name = "groupBoxTextAttributes";
            this.groupBoxTextAttributes.Size = new System.Drawing.Size(482, 615);
            this.groupBoxTextAttributes.TabIndex = 7;
            this.groupBoxTextAttributes.TabStop = false;
            this.groupBoxTextAttributes.Text = "样机样式";
            // 
            // SampleGenerationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(568, 755);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.pictureBoxStyle6);
            this.Controls.Add(this.pictureBoxStyle5);
            this.Controls.Add(this.pictureBoxStyle4);
            this.Controls.Add(this.pictureBoxStyle3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.buttonGenerate);
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
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStyle6)).EndInit();
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
    }
}
