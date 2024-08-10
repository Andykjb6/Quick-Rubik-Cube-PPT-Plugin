using System;
using System.Windows.Forms;  // 添加对 WinForms 的引用


namespace 课件帮PPT助手
{
    partial class FontSizeAdjustForm
    {
        private System.ComponentModel.IContainer components = null;
        private NumericUpDown numericUpDownPinyin;
        private NumericUpDown numericUpDownHanzi;
        private Button confirmButton;
        private Button cancelButton;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.numericUpDownPinyin = new System.Windows.Forms.NumericUpDown();
            this.numericUpDownHanzi = new System.Windows.Forms.NumericUpDown();
            this.confirmButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownPinyin)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownHanzi)).BeginInit();
            this.SuspendLayout();
            // 
            // numericUpDownPinyin
            // 
            this.numericUpDownPinyin.DecimalPlaces = 1;
            this.numericUpDownPinyin.Font = new System.Drawing.Font("Microsoft YaHei UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.numericUpDownPinyin.Location = new System.Drawing.Point(211, 36);
            this.numericUpDownPinyin.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDownPinyin.Name = "numericUpDownPinyin";
            this.numericUpDownPinyin.Size = new System.Drawing.Size(120, 48);
            this.numericUpDownPinyin.TabIndex = 0;
            this.numericUpDownPinyin.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // numericUpDownHanzi
            // 
            this.numericUpDownHanzi.DecimalPlaces = 1;
            this.numericUpDownHanzi.Font = new System.Drawing.Font("Microsoft YaHei UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.numericUpDownHanzi.Location = new System.Drawing.Point(211, 122);
            this.numericUpDownHanzi.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDownHanzi.Name = "numericUpDownHanzi";
            this.numericUpDownHanzi.Size = new System.Drawing.Size(120, 48);
            this.numericUpDownHanzi.TabIndex = 1;
            this.numericUpDownHanzi.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // confirmButton
            // 
            this.confirmButton.Font = new System.Drawing.Font("Microsoft YaHei UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.confirmButton.Location = new System.Drawing.Point(76, 207);
            this.confirmButton.Name = "confirmButton";
            this.confirmButton.Size = new System.Drawing.Size(117, 50);
            this.confirmButton.TabIndex = 2;
            this.confirmButton.Text = "确认";
            this.confirmButton.Click += new System.EventHandler(this.ConfirmButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.Font = new System.Drawing.Font("Microsoft YaHei UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cancelButton.Location = new System.Drawing.Point(211, 207);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(117, 50);
            this.cancelButton.TabIndex = 3;
            this.cancelButton.Text = "取消";
            this.cancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft YaHei UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(76, 41);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(114, 42);
            this.label1.TabIndex = 4;
            this.label1.Text = "拼音：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft YaHei UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(76, 127);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(114, 42);
            this.label2.TabIndex = 5;
            this.label2.Text = "汉字：";
            // 
            // FontSizeAdjustForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(192F, 192F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(406, 312);
            this.Controls.Add(this.numericUpDownPinyin);
            this.Controls.Add(this.numericUpDownHanzi);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.confirmButton);
            this.Controls.Add(this.cancelButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.Name = "FontSizeAdjustForm";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "字号调整";
            this.TopMost = true;
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownPinyin)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownHanzi)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Label label1;
        private Label label2;
    }
}
