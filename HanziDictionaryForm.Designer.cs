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
            this.hanziDictionaryControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.hanziDictionaryControl.Location = new System.Drawing.Point(0, 0);
            this.hanziDictionaryControl.Name = "hanziDictionaryControl";
            this.hanziDictionaryControl.Size = new System.Drawing.Size(600, 800);
            this.hanziDictionaryControl.TabIndex = 0;
            // 
            // HanziDictionaryForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(600, 800);
            this.Controls.Add(this.hanziDictionaryControl);
            this.Name = "HanziDictionaryForm";
            this.Text = "汉字字典";
            this.ResumeLayout(false);
        }

        #endregion

        private 课件帮PPT助手.HanziDictionaryControl hanziDictionaryControl;
    }
}
