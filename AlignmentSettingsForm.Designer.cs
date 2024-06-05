namespace 课件帮PPT助手
{
    partial class AlignmentSettingsForm
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
            this.centerRadioButton = new System.Windows.Forms.RadioButton();
            this.leftRadioButton = new System.Windows.Forms.RadioButton();
            this.rightRadioButton = new System.Windows.Forms.RadioButton();
            this.topRadioButton = new System.Windows.Forms.RadioButton();
            this.bottomRadioButton = new System.Windows.Forms.RadioButton();
            this.saveButton = new System.Windows.Forms.Button();
            this.exitButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // centerRadioButton
            // 
            this.centerRadioButton.AutoSize = true;
            this.centerRadioButton.Checked = true;
            this.centerRadioButton.Location = new System.Drawing.Point(26, 24);
            this.centerRadioButton.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.centerRadioButton.Name = "centerRadioButton";
            this.centerRadioButton.Size = new System.Drawing.Size(137, 28);
            this.centerRadioButton.TabIndex = 0;
            this.centerRadioButton.TabStop = true;
            this.centerRadioButton.Text = "中心对齐";
            this.centerRadioButton.UseVisualStyleBackColor = true;
            // 
            // leftRadioButton
            // 
            this.leftRadioButton.AutoSize = true;
            this.leftRadioButton.Location = new System.Drawing.Point(190, 24);
            this.leftRadioButton.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.leftRadioButton.Name = "leftRadioButton";
            this.leftRadioButton.Size = new System.Drawing.Size(113, 28);
            this.leftRadioButton.TabIndex = 1;
            this.leftRadioButton.Text = "左对齐";
            this.leftRadioButton.UseVisualStyleBackColor = true;
            // 
            // rightRadioButton
            // 
            this.rightRadioButton.AutoSize = true;
            this.rightRadioButton.Location = new System.Drawing.Point(26, 77);
            this.rightRadioButton.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.rightRadioButton.Name = "rightRadioButton";
            this.rightRadioButton.Size = new System.Drawing.Size(113, 28);
            this.rightRadioButton.TabIndex = 2;
            this.rightRadioButton.Text = "右对齐";
            this.rightRadioButton.UseVisualStyleBackColor = true;
            this.rightRadioButton.CheckedChanged += new System.EventHandler(this.rightRadioButton_CheckedChanged);
            // 
            // topRadioButton
            // 
            this.topRadioButton.AutoSize = true;
            this.topRadioButton.Location = new System.Drawing.Point(190, 77);
            this.topRadioButton.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.topRadioButton.Name = "topRadioButton";
            this.topRadioButton.Size = new System.Drawing.Size(137, 28);
            this.topRadioButton.TabIndex = 3;
            this.topRadioButton.Text = "顶部对齐";
            this.topRadioButton.UseVisualStyleBackColor = true;
            // 
            // bottomRadioButton
            // 
            this.bottomRadioButton.AutoSize = true;
            this.bottomRadioButton.Location = new System.Drawing.Point(26, 132);
            this.bottomRadioButton.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.bottomRadioButton.Name = "bottomRadioButton";
            this.bottomRadioButton.Size = new System.Drawing.Size(137, 28);
            this.bottomRadioButton.TabIndex = 4;
            this.bottomRadioButton.Text = "底部对齐";
            this.bottomRadioButton.UseVisualStyleBackColor = true;
            // 
            // saveButton
            // 
            this.saveButton.Location = new System.Drawing.Point(26, 197);
            this.saveButton.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.saveButton.Name = "saveButton";
            this.saveButton.Size = new System.Drawing.Size(150, 42);
            this.saveButton.TabIndex = 5;
            this.saveButton.Text = "保存";
            this.saveButton.UseVisualStyleBackColor = true;
            this.saveButton.Click += new System.EventHandler(this.SaveButton_Click);
            // 
            // exitButton
            // 
            this.exitButton.Location = new System.Drawing.Point(190, 197);
            this.exitButton.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.exitButton.Name = "exitButton";
            this.exitButton.Size = new System.Drawing.Size(150, 42);
            this.exitButton.TabIndex = 6;
            this.exitButton.Text = "退出";
            this.exitButton.UseVisualStyleBackColor = true;
            this.exitButton.Click += new System.EventHandler(this.ExitButton_Click);
            // 
            // AlignmentSettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(371, 275);
            this.Controls.Add(this.exitButton);
            this.Controls.Add(this.saveButton);
            this.Controls.Add(this.bottomRadioButton);
            this.Controls.Add(this.topRadioButton);
            this.Controls.Add(this.rightRadioButton);
            this.Controls.Add(this.leftRadioButton);
            this.Controls.Add(this.centerRadioButton);
            this.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.Name = "AlignmentSettingsForm";
            this.Text = "设置";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.Windows.Forms.RadioButton centerRadioButton;
        private System.Windows.Forms.RadioButton leftRadioButton;
        private System.Windows.Forms.RadioButton rightRadioButton;
        private System.Windows.Forms.RadioButton topRadioButton;
        private System.Windows.Forms.RadioButton bottomRadioButton;
        private System.Windows.Forms.Button saveButton;
        private System.Windows.Forms.Button exitButton;
    }
}
