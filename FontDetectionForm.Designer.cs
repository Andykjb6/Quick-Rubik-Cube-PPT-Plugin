namespace 课件帮PPT助手
{
    partial class FontDetectionForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.ListBox listBoxUsed;
        private System.Windows.Forms.ListBox listBoxUnused;
        private System.Windows.Forms.Button clearButton;
        private System.Windows.Forms.Button exportFontsButton;
        private System.Windows.Forms.Button packageDocumentButton;

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FontDetectionForm));
            this.listBoxUsed = new System.Windows.Forms.ListBox();
            this.listBoxUnused = new System.Windows.Forms.ListBox();
            this.clearButton = new System.Windows.Forms.Button();
            this.exportFontsButton = new System.Windows.Forms.Button();
            this.packageDocumentButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // listBoxUsed
            // 
            this.listBoxUsed.FormattingEnabled = true;
            this.listBoxUsed.ItemHeight = 24;
            this.listBoxUsed.Location = new System.Drawing.Point(17, 61);
            this.listBoxUsed.Name = "listBoxUsed";
            this.listBoxUsed.Size = new System.Drawing.Size(219, 268);
            this.listBoxUsed.TabIndex = 0;
            // 
            // listBoxUnused
            // 
            this.listBoxUnused.FormattingEnabled = true;
            this.listBoxUnused.ItemHeight = 24;
            this.listBoxUnused.Location = new System.Drawing.Point(251, 61);
            this.listBoxUnused.Name = "listBoxUnused";
            this.listBoxUnused.Size = new System.Drawing.Size(219, 268);
            this.listBoxUnused.TabIndex = 1;
            // 
            // clearButton
            // 
            this.clearButton.BackColor = System.Drawing.Color.White;
            this.clearButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(56)))), ((int)(((byte)(56)))), ((int)(((byte)(238)))));
            this.clearButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(195)))), ((int)(((byte)(195)))), ((int)(((byte)(241)))));
            this.clearButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(247)))), ((int)(((byte)(247)))), ((int)(((byte)(255)))));
            this.clearButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.clearButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.clearButton.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(56)))), ((int)(((byte)(56)))), ((int)(((byte)(238)))));
            this.clearButton.Location = new System.Drawing.Point(251, 339);
            this.clearButton.Name = "clearButton";
            this.clearButton.Size = new System.Drawing.Size(219, 46);
            this.clearButton.TabIndex = 2;
            this.clearButton.Text = "清除未用字体";
            this.clearButton.UseVisualStyleBackColor = false;
            this.clearButton.Click += new System.EventHandler(this.ClearButton_Click);
            // 
            // exportFontsButton
            // 
            this.exportFontsButton.BackColor = System.Drawing.Color.White;
            this.exportFontsButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(56)))), ((int)(((byte)(56)))), ((int)(((byte)(238)))));
            this.exportFontsButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(195)))), ((int)(((byte)(195)))), ((int)(((byte)(241)))));
            this.exportFontsButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(247)))), ((int)(((byte)(247)))), ((int)(((byte)(255)))));
            this.exportFontsButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.exportFontsButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.exportFontsButton.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(56)))), ((int)(((byte)(56)))), ((int)(((byte)(238)))));
            this.exportFontsButton.Location = new System.Drawing.Point(17, 339);
            this.exportFontsButton.Name = "exportFontsButton";
            this.exportFontsButton.Size = new System.Drawing.Size(219, 46);
            this.exportFontsButton.TabIndex = 3;
            this.exportFontsButton.Text = "导出已用字体";
            this.exportFontsButton.UseVisualStyleBackColor = false;
            this.exportFontsButton.Click += new System.EventHandler(this.ExportFontsButton_Click);
            // 
            // packageDocumentButton
            // 
            this.packageDocumentButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(56)))), ((int)(((byte)(56)))), ((int)(((byte)(238)))));
            this.packageDocumentButton.FlatAppearance.BorderSize = 0;
            this.packageDocumentButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(26)))), ((int)(((byte)(26)))), ((int)(((byte)(214)))));
            this.packageDocumentButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(79)))), ((int)(((byte)(79)))), ((int)(((byte)(232)))));
            this.packageDocumentButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.packageDocumentButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.packageDocumentButton.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.packageDocumentButton.Location = new System.Drawing.Point(17, 397);
            this.packageDocumentButton.Name = "packageDocumentButton";
            this.packageDocumentButton.Size = new System.Drawing.Size(453, 48);
            this.packageDocumentButton.TabIndex = 4;
            this.packageDocumentButton.Text = "打包文档";
            this.packageDocumentButton.UseVisualStyleBackColor = false;
            this.packageDocumentButton.Click += new System.EventHandler(this.PackageDocumentButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(18, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(160, 24);
            this.label1.TabIndex = 5;
            this.label1.Text = "已使用字体：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(258, 22);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(160, 24);
            this.label2.TabIndex = 6;
            this.label2.Text = "未使用字体：";
            // 
            // FontDetectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(192F, 192F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(487, 468);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.packageDocumentButton);
            this.Controls.Add(this.exportFontsButton);
            this.Controls.Add(this.clearButton);
            this.Controls.Add(this.listBoxUnused);
            this.Controls.Add(this.listBoxUsed);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FontDetectionForm";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "检测结果";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}
