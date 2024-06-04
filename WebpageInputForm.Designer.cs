namespace 课件帮PPT助手
{
    partial class WebpageInputForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.TextBox txtUrl;
        private System.Windows.Forms.Label lblUrl;
        private System.Windows.Forms.Button btnEmbed;
        private System.Windows.Forms.Button btnCancel;

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
            this.txtUrl = new System.Windows.Forms.TextBox();
            this.lblUrl = new System.Windows.Forms.Label();
            this.btnEmbed = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtUrl
            // 
            this.txtUrl.Location = new System.Drawing.Point(150, 37);
            this.txtUrl.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.txtUrl.Name = "txtUrl";
            this.txtUrl.Size = new System.Drawing.Size(396, 35);
            this.txtUrl.TabIndex = 0;
            // 
            // lblUrl
            // 
            this.lblUrl.AutoSize = true;
            this.lblUrl.Location = new System.Drawing.Point(40, 42);
            this.lblUrl.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.lblUrl.Name = "lblUrl";
            this.lblUrl.Size = new System.Drawing.Size(70, 24);
            this.lblUrl.TabIndex = 1;
            this.lblUrl.Text = "网址:";
            // 
            // btnEmbed
            // 
            this.btnEmbed.Location = new System.Drawing.Point(150, 111);
            this.btnEmbed.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.btnEmbed.Name = "btnEmbed";
            this.btnEmbed.Size = new System.Drawing.Size(150, 42);
            this.btnEmbed.TabIndex = 2;
            this.btnEmbed.Text = "嵌入";
            this.btnEmbed.UseVisualStyleBackColor = true;
            this.btnEmbed.Click += new System.EventHandler(this.btnEmbed_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(396, 111);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(150, 42);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // WebpageInputForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(600, 185);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnEmbed);
            this.Controls.Add(this.lblUrl);
            this.Controls.Add(this.txtUrl);
            this.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
            this.Name = "WebpageInputForm";
            this.Text = "嵌入网页";
            this.ResumeLayout(false);
            this.PerformLayout();

        }
    }
}
