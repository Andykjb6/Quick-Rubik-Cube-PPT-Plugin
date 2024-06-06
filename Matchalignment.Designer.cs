namespace 课件帮PPT助手
{
    partial class Matchalignment
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Matchalignment));
            this.中心对齐 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // 中心对齐
            // 
            this.中心对齐.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("中心对齐.BackgroundImage")));
            this.中心对齐.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.中心对齐.Location = new System.Drawing.Point(52, 56);
            this.中心对齐.Name = "中心对齐";
            this.中心对齐.Size = new System.Drawing.Size(74, 73);
            this.中心对齐.TabIndex = 0;
            this.中心对齐.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button1.BackgroundImage")));
            this.button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button1.Location = new System.Drawing.Point(158, 56);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(74, 73);
            this.button1.TabIndex = 1;
            this.button1.UseVisualStyleBackColor = true;
            // 
            // Matchalignment
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.中心对齐);
            this.Name = "Matchalignment";
            this.Text = "匹配对齐";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button 中心对齐;
        private System.Windows.Forms.Button button1;
    }
}