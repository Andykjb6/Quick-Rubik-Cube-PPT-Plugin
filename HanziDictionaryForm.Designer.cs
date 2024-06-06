namespace 课件帮PPT助手
{
    partial class HanziDictionaryForm
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

        #region Windows 窗体设计器生成的代码

        private void InitializeComponent()
        {
            this.hanziDictionaryControl = new 课件帮PPT助手.HanziDictionaryControl();
            this.SuspendLayout();
            // 
            // hanziDictionaryControl
            // 
            this.hanziDictionaryControl.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(244)))), ((int)(((byte)(244)))), ((int)(((byte)(254)))));
            this.hanziDictionaryControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.hanziDictionaryControl.Location = new System.Drawing.Point(0, 0);
            this.hanziDictionaryControl.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.hanziDictionaryControl.Name = "hanziDictionaryControl";
            this.hanziDictionaryControl.Size = new System.Drawing.Size(720, 960);
            this.hanziDictionaryControl.TabIndex = 0;
            this.hanziDictionaryControl.Load += new System.EventHandler(this.hanziDictionaryControl_Load);
            // 
            // HanziDictionaryForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(720, 960);
            this.Controls.Add(this.hanziDictionaryControl);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "HanziDictionaryForm";
            this.Text = "汉字字典";
            this.ResumeLayout(false);

        }

        #endregion

        private 课件帮PPT助手.HanziDictionaryControl hanziDictionaryControl;
    }
}
