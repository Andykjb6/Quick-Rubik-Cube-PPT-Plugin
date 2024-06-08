namespace 课件帮PPT助手
{
    partial class PinyinSelectorForm
    {
        private System.Windows.Forms.ComboBox comboBox;
        private System.Windows.Forms.Button refreshButton;
        private System.Windows.Forms.Button replaceButton;
        private System.Windows.Forms.Button closeButton;

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PinyinSelectorForm));
            this.comboBox = new System.Windows.Forms.ComboBox();
            this.refreshButton = new System.Windows.Forms.Button();
            this.replaceButton = new System.Windows.Forms.Button();
            this.closeButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // comboBox
            // 
            this.comboBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(226)))), ((int)(((byte)(234)))), ((int)(((byte)(255)))));
            this.comboBox.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.comboBox.Font = new System.Drawing.Font("Pinyinia_a", 11F);
            this.comboBox.Location = new System.Drawing.Point(24, 20);
            this.comboBox.Name = "comboBox";
            this.comboBox.Size = new System.Drawing.Size(410, 51);
            this.comboBox.TabIndex = 0;
            // 
            // refreshButton
            // 
            this.refreshButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(33)))), ((int)(((byte)(92)))), ((int)(((byte)(242)))));
            this.refreshButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(81)))), ((int)(((byte)(225)))));
            this.refreshButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(82)))), ((int)(((byte)(126)))), ((int)(((byte)(236)))));
            this.refreshButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.refreshButton.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.refreshButton.Location = new System.Drawing.Point(58, 95);
            this.refreshButton.Name = "refreshButton";
            this.refreshButton.Size = new System.Drawing.Size(100, 45);
            this.refreshButton.TabIndex = 1;
            this.refreshButton.Text = "刷新";
            this.refreshButton.UseVisualStyleBackColor = false;
            this.refreshButton.Click += new System.EventHandler(this.RefreshButton_Click);
            // 
            // replaceButton
            // 
            this.replaceButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(33)))), ((int)(((byte)(92)))), ((int)(((byte)(242)))));
            this.replaceButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(81)))), ((int)(((byte)(225)))));
            this.replaceButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(82)))), ((int)(((byte)(126)))), ((int)(((byte)(236)))));
            this.replaceButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.replaceButton.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.replaceButton.Location = new System.Drawing.Point(184, 95);
            this.replaceButton.Name = "replaceButton";
            this.replaceButton.Size = new System.Drawing.Size(100, 45);
            this.replaceButton.TabIndex = 2;
            this.replaceButton.Text = "注音";
            this.replaceButton.UseVisualStyleBackColor = false;
            this.replaceButton.Click += new System.EventHandler(this.ReplaceButton_Click);
            // 
            // closeButton
            // 
            this.closeButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(33)))), ((int)(((byte)(92)))), ((int)(((byte)(242)))));
            this.closeButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(81)))), ((int)(((byte)(225)))));
            this.closeButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(82)))), ((int)(((byte)(126)))), ((int)(((byte)(236)))));
            this.closeButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.closeButton.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.closeButton.Location = new System.Drawing.Point(304, 95);
            this.closeButton.Name = "closeButton";
            this.closeButton.Size = new System.Drawing.Size(100, 45);
            this.closeButton.TabIndex = 3;
            this.closeButton.Text = "退出";
            this.closeButton.UseVisualStyleBackColor = false;
            this.closeButton.Click += new System.EventHandler(this.CloseButton_Click);
            // 
            // PinyinSelectorForm
            // 
            this.ClientSize = new System.Drawing.Size(458, 168);
            this.Controls.Add(this.comboBox);
            this.Controls.Add(this.refreshButton);
            this.Controls.Add(this.replaceButton);
            this.Controls.Add(this.closeButton);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "PinyinSelectorForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "便捷注音";
            this.TopMost = true;
            this.ResumeLayout(false);

        }
    }
}
