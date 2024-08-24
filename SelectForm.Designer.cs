namespace 课件帮PPT助手
{
    partial class SelectForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.ComboBox comboBoxOptions;
        private System.Windows.Forms.Label labelPrompt;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.Button buttonCancel;

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SelectForm));
            this.comboBoxOptions = new System.Windows.Forms.ComboBox();
            this.labelPrompt = new System.Windows.Forms.Label();
            this.buttonOK = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // comboBoxOptions
            // 
            this.comboBoxOptions.Font = new System.Drawing.Font("Microsoft New Tai Lue", 10F);
            this.comboBoxOptions.FormattingEnabled = true;
            this.comboBoxOptions.Location = new System.Drawing.Point(24, 76);
            this.comboBoxOptions.Margin = new System.Windows.Forms.Padding(6);
            this.comboBoxOptions.Name = "comboBoxOptions";
            this.comboBoxOptions.Size = new System.Drawing.Size(516, 43);
            this.comboBoxOptions.TabIndex = 0;
            // 
            // labelPrompt
            // 
            this.labelPrompt.AutoSize = true;
            this.labelPrompt.Location = new System.Drawing.Point(24, 28);
            this.labelPrompt.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.labelPrompt.Name = "labelPrompt";
            this.labelPrompt.Size = new System.Drawing.Size(0, 24);
            this.labelPrompt.TabIndex = 1;
            // 
            // buttonOK
            // 
            this.buttonOK.Font = new System.Drawing.Font("思源宋体 CN", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonOK.Location = new System.Drawing.Point(394, 137);
            this.buttonOK.Margin = new System.Windows.Forms.Padding(6);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(150, 47);
            this.buttonOK.TabIndex = 2;
            this.buttonOK.Text = "确定";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.Font = new System.Drawing.Font("思源宋体 CN", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonCancel.Location = new System.Drawing.Point(232, 137);
            this.buttonCancel.Margin = new System.Windows.Forms.Padding(6);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(150, 47);
            this.buttonCancel.TabIndex = 3;
            this.buttonCancel.Text = "取消";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // SelectForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(568, 217);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.labelPrompt);
            this.Controls.Add(this.comboBoxOptions);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(6);
            this.Name = "SelectForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "选择拆字方式";
            this.ResumeLayout(false);
            this.PerformLayout();

        }
    }
}