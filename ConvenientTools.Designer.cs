﻿namespace 课件帮PPT助手
{
    partial class DesignTools
    {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region 组件设计器生成的代码

        /// <summary> 
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DesignTools));
            this.label1 = new System.Windows.Forms.Label();
            this.字源字形 = new System.Windows.Forms.Button();
            this.文字标注 = new System.Windows.Forms.Button();
            this.笔画拆分 = new System.Windows.Forms.Button();
            this.书写动画 = new System.Windows.Forms.Button();
            this.矩形拆分 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.ForeColor = System.Drawing.SystemColors.Highlight;
            this.label1.Name = "label1";
            // 
            // 字源字形
            // 
            this.字源字形.BackColor = System.Drawing.SystemColors.ControlLightLight;
            resources.ApplyResources(this.字源字形, "字源字形");
            this.字源字形.ForeColor = System.Drawing.SystemColors.Highlight;
            this.字源字形.Name = "字源字形";
            this.字源字形.UseVisualStyleBackColor = false;
            this.字源字形.Click += new System.EventHandler(this.字源字形_Click);
            // 
            // 文字标注
            // 
            this.文字标注.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.文字标注.ForeColor = System.Drawing.SystemColors.Highlight;
            resources.ApplyResources(this.文字标注, "文字标注");
            this.文字标注.Name = "文字标注";
            this.文字标注.UseVisualStyleBackColor = false;
            this.文字标注.Click += new System.EventHandler(this.文字标注_Click);
            // 
            // 笔画拆分
            // 
            this.笔画拆分.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.笔画拆分.ForeColor = System.Drawing.SystemColors.Highlight;
            resources.ApplyResources(this.笔画拆分, "笔画拆分");
            this.笔画拆分.Name = "笔画拆分";
            this.笔画拆分.UseVisualStyleBackColor = false;
            this.笔画拆分.Click += new System.EventHandler(this.笔画拆分_Click);
            // 
            // 书写动画
            // 
            resources.ApplyResources(this.书写动画, "书写动画");
            this.书写动画.Name = "书写动画";
            this.书写动画.UseVisualStyleBackColor = true;
            this.书写动画.Click += new System.EventHandler(this.书写动画_Click);
            // 
            // 矩形拆分
            // 
            resources.ApplyResources(this.矩形拆分, "矩形拆分");
            this.矩形拆分.Name = "矩形拆分";
            this.矩形拆分.UseVisualStyleBackColor = true;
            this.矩形拆分.Click += new System.EventHandler(this.button1_Click);
            // 
            // DesignTools
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.Controls.Add(this.矩形拆分);
            this.Controls.Add(this.书写动画);
            this.Controls.Add(this.笔画拆分);
            this.Controls.Add(this.文字标注);
            this.Controls.Add(this.字源字形);
            this.Controls.Add(this.label1);
            this.Name = "DesignTools";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button 字源字形;
        private System.Windows.Forms.Button 文字标注;
        private System.Windows.Forms.Button 笔画拆分;
        private System.Windows.Forms.Button 书写动画;
        private System.Windows.Forms.Button 矩形拆分;
    }
}
