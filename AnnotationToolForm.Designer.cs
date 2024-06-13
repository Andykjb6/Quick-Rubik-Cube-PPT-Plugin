using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using static 课件帮PPT助手.DesignTools;

namespace 课件帮PPT助手
{
    partial class AnnotationToolForm
    {
        private ComboBox annotationTypeComboBox;
        private Button annotationColorButton;
        private Button confirmButton;
        private Button clearButton;
        private Button deleteCustomAnnotationButton;
        private CheckBox boldCheckBox;
        private CheckBox italicCheckBox;
        private CheckBox highlightCheckBox;
        private Button highlightColorButton;
        private Button textColorButton;
        private Label annotationTypeLabel;
        private Label annotationColorLabel;
        private Label textColorLabel;
        private Label highlightColorLabel;
        private Label textSettingsLabel;
        private ContextMenuStrip contextMenuStrip;
        private ToolStripMenuItem customizeAnnotationMenuItem;
        private ToolStripMenuItem deleteAnnotationMenuItem;

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AnnotationToolForm));
            this.annotationTypeComboBox = new System.Windows.Forms.ComboBox();
            this.annotationColorButton = new System.Windows.Forms.Button();
            this.confirmButton = new System.Windows.Forms.Button();
            this.clearButton = new System.Windows.Forms.Button();
            this.deleteCustomAnnotationButton = new System.Windows.Forms.Button();
            this.boldCheckBox = new System.Windows.Forms.CheckBox();
            this.italicCheckBox = new System.Windows.Forms.CheckBox();
            this.highlightCheckBox = new System.Windows.Forms.CheckBox();
            this.highlightColorButton = new System.Windows.Forms.Button();
            this.textColorButton = new System.Windows.Forms.Button();
            this.annotationTypeLabel = new System.Windows.Forms.Label();
            this.annotationColorLabel = new System.Windows.Forms.Label();
            this.textColorLabel = new System.Windows.Forms.Label();
            this.highlightColorLabel = new System.Windows.Forms.Label();
            this.textSettingsLabel = new System.Windows.Forms.Label();
            this.contextMenuStrip = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.customizeAnnotationMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.deleteAnnotationMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.SuspendLayout();
            // 
            // annotationTypeComboBox
            // 
            this.annotationTypeComboBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(226)))), ((int)(((byte)(234)))), ((int)(((byte)(255)))));
            this.annotationTypeComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.annotationTypeComboBox.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.annotationTypeComboBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.annotationTypeComboBox.ForeColor = System.Drawing.SystemColors.WindowText;
            this.annotationTypeComboBox.FormattingEnabled = true;
            this.annotationTypeComboBox.Items.AddRange(new object[] {
            "横线",
            "双横线",
            "波浪线",
            "重读符号",
            "轻读符号",
            "着重符号",
            "大括号",
            "层级符",
            "段落符"});
            this.annotationTypeComboBox.Location = new System.Drawing.Point(93, 20);
            this.annotationTypeComboBox.Name = "annotationTypeComboBox";
            this.annotationTypeComboBox.Size = new System.Drawing.Size(210, 39);
            this.annotationTypeComboBox.TabIndex = 1;
            // 
            // annotationColorButton
            // 
            this.annotationColorButton.BackColor = System.Drawing.Color.Red;
            this.annotationColorButton.Location = new System.Drawing.Point(430, 20);
            this.annotationColorButton.Name = "annotationColorButton";
            this.annotationColorButton.Size = new System.Drawing.Size(50, 30);
            this.annotationColorButton.TabIndex = 2;
            this.annotationColorButton.UseVisualStyleBackColor = false;
            this.annotationColorButton.Click += new System.EventHandler(this.AnnotationColorButton_Click);
            // 
            // confirmButton
            // 
            this.confirmButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(19)))), ((int)(((byte)(84)))), ((int)(((byte)(247)))));
            this.confirmButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(14)))), ((int)(((byte)(76)))), ((int)(((byte)(231)))));
            this.confirmButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(108)))), ((int)(((byte)(150)))), ((int)(((byte)(255)))));
            this.confirmButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.confirmButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.confirmButton.ForeColor = System.Drawing.Color.White;
            this.confirmButton.Location = new System.Drawing.Point(50, 210);
            this.confirmButton.Name = "confirmButton";
            this.confirmButton.Size = new System.Drawing.Size(120, 60);
            this.confirmButton.TabIndex = 8;
            this.confirmButton.Text = "标注所选";
            this.confirmButton.UseVisualStyleBackColor = false;
            this.confirmButton.Click += new System.EventHandler(this.ConfirmButton_Click);
            // 
            // clearButton
            // 
            this.clearButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(19)))), ((int)(((byte)(84)))), ((int)(((byte)(247)))));
            this.clearButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(14)))), ((int)(((byte)(76)))), ((int)(((byte)(231)))));
            this.clearButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(108)))), ((int)(((byte)(150)))), ((int)(((byte)(255)))));
            this.clearButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.clearButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.clearButton.ForeColor = System.Drawing.Color.White;
            this.clearButton.Location = new System.Drawing.Point(190, 210);
            this.clearButton.Name = "clearButton";
            this.clearButton.Size = new System.Drawing.Size(120, 60);
            this.clearButton.TabIndex = 9;
            this.clearButton.Text = "清除所选";
            this.clearButton.UseVisualStyleBackColor = false;
            this.clearButton.Click += new System.EventHandler(this.ClearButton_Click);
            // 
            // deleteCustomAnnotationButton
            // 
            this.deleteCustomAnnotationButton.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.deleteCustomAnnotationButton.Enabled = false;
            this.deleteCustomAnnotationButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(40)))), ((int)(((byte)(99)))), ((int)(((byte)(245)))));
            this.deleteCustomAnnotationButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(199)))), ((int)(((byte)(215)))), ((int)(((byte)(255)))));
            this.deleteCustomAnnotationButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(226)))), ((int)(((byte)(234)))), ((int)(((byte)(252)))));
            this.deleteCustomAnnotationButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.deleteCustomAnnotationButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.deleteCustomAnnotationButton.Location = new System.Drawing.Point(330, 210);
            this.deleteCustomAnnotationButton.Name = "deleteCustomAnnotationButton";
            this.deleteCustomAnnotationButton.Size = new System.Drawing.Size(120, 60);
            this.deleteCustomAnnotationButton.TabIndex = 10;
            this.deleteCustomAnnotationButton.Text = "删除标注";
            this.deleteCustomAnnotationButton.UseVisualStyleBackColor = false;
            this.deleteCustomAnnotationButton.Click += new System.EventHandler(this.DeleteCustomAnnotationButton_Click);
            // 
            // boldCheckBox
            // 
            this.boldCheckBox.AutoSize = true;
            this.boldCheckBox.Checked = true;
            this.boldCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.boldCheckBox.Location = new System.Drawing.Point(90, 79);
            this.boldCheckBox.Name = "boldCheckBox";
            this.boldCheckBox.Size = new System.Drawing.Size(90, 28);
            this.boldCheckBox.TabIndex = 3;
            this.boldCheckBox.Text = "加粗";
            this.boldCheckBox.UseVisualStyleBackColor = true;
            // 
            // italicCheckBox
            // 
            this.italicCheckBox.AutoSize = true;
            this.italicCheckBox.Location = new System.Drawing.Point(195, 79);
            this.italicCheckBox.Name = "italicCheckBox";
            this.italicCheckBox.Size = new System.Drawing.Size(90, 28);
            this.italicCheckBox.TabIndex = 4;
            this.italicCheckBox.Text = "倾斜";
            this.italicCheckBox.UseVisualStyleBackColor = true;
            // 
            // highlightCheckBox
            // 
            this.highlightCheckBox.AutoSize = true;
            this.highlightCheckBox.Location = new System.Drawing.Point(300, 79);
            this.highlightCheckBox.Name = "highlightCheckBox";
            this.highlightCheckBox.Size = new System.Drawing.Size(90, 28);
            this.highlightCheckBox.TabIndex = 5;
            this.highlightCheckBox.Text = "高亮";
            this.highlightCheckBox.UseVisualStyleBackColor = true;
            this.highlightCheckBox.CheckedChanged += new System.EventHandler(this.HighlightCheckBox_CheckedChanged);
            // 
            // highlightColorButton
            // 
            this.highlightColorButton.Enabled = false;
            this.highlightColorButton.Location = new System.Drawing.Point(344, 140);
            this.highlightColorButton.Name = "highlightColorButton";
            this.highlightColorButton.Size = new System.Drawing.Size(50, 30);
            this.highlightColorButton.TabIndex = 6;
            this.highlightColorButton.UseVisualStyleBackColor = true;
            this.highlightColorButton.Click += new System.EventHandler(this.HighlightColorButton_Click);
            // 
            // textColorButton
            // 
            this.textColorButton.BackColor = System.Drawing.Color.Red;
            this.textColorButton.Location = new System.Drawing.Point(139, 140);
            this.textColorButton.Name = "textColorButton";
            this.textColorButton.Size = new System.Drawing.Size(50, 30);
            this.textColorButton.TabIndex = 7;
            this.textColorButton.UseVisualStyleBackColor = false;
            this.textColorButton.Click += new System.EventHandler(this.TextColorButton_Click);
            // 
            // annotationTypeLabel
            // 
            this.annotationTypeLabel.AutoSize = true;
            this.annotationTypeLabel.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.annotationTypeLabel.Location = new System.Drawing.Point(20, 20);
            this.annotationTypeLabel.Name = "annotationTypeLabel";
            this.annotationTypeLabel.Size = new System.Drawing.Size(86, 31);
            this.annotationTypeLabel.TabIndex = 11;
            this.annotationTypeLabel.Text = "标注：";
            // 
            // annotationColorLabel
            // 
            this.annotationColorLabel.AutoSize = true;
            this.annotationColorLabel.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.annotationColorLabel.Location = new System.Drawing.Point(310, 20);
            this.annotationColorLabel.Name = "annotationColorLabel";
            this.annotationColorLabel.Size = new System.Drawing.Size(134, 31);
            this.annotationColorLabel.TabIndex = 12;
            this.annotationColorLabel.Text = "标注颜色：";
            // 
            // textColorLabel
            // 
            this.textColorLabel.AutoSize = true;
            this.textColorLabel.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textColorLabel.Location = new System.Drawing.Point(20, 140);
            this.textColorLabel.Name = "textColorLabel";
            this.textColorLabel.Size = new System.Drawing.Size(134, 31);
            this.textColorLabel.TabIndex = 13;
            this.textColorLabel.Text = "文字颜色：";
            // 
            // highlightColorLabel
            // 
            this.highlightColorLabel.AutoSize = true;
            this.highlightColorLabel.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.highlightColorLabel.Location = new System.Drawing.Point(230, 140);
            this.highlightColorLabel.Name = "highlightColorLabel";
            this.highlightColorLabel.Size = new System.Drawing.Size(134, 31);
            this.highlightColorLabel.TabIndex = 14;
            this.highlightColorLabel.Text = "高亮颜色：";
            // 
            // textSettingsLabel
            // 
            this.textSettingsLabel.AutoSize = true;
            this.textSettingsLabel.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textSettingsLabel.Location = new System.Drawing.Point(20, 80);
            this.textSettingsLabel.Name = "textSettingsLabel";
            this.textSettingsLabel.Size = new System.Drawing.Size(86, 31);
            this.textSettingsLabel.TabIndex = 15;
            this.textSettingsLabel.Text = "文字：";
            // 
            // contextMenuStrip
            // 
            this.contextMenuStrip.ImageScalingSize = new System.Drawing.Size(32, 32);
            this.contextMenuStrip.Name = "contextMenuStrip";
            this.contextMenuStrip.Size = new System.Drawing.Size(61, 4);
            // 
            // customizeAnnotationMenuItem
            // 
            this.customizeAnnotationMenuItem.Name = "customizeAnnotationMenuItem";
            this.customizeAnnotationMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // deleteAnnotationMenuItem
            // 
            this.deleteAnnotationMenuItem.Name = "deleteAnnotationMenuItem";
            this.deleteAnnotationMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // AnnotationToolForm
            // 
            this.ClientSize = new System.Drawing.Size(500, 311);
            this.Controls.Add(this.annotationTypeComboBox);
            this.Controls.Add(this.annotationColorButton);
            this.Controls.Add(this.boldCheckBox);
            this.Controls.Add(this.italicCheckBox);
            this.Controls.Add(this.highlightCheckBox);
            this.Controls.Add(this.highlightColorButton);
            this.Controls.Add(this.textColorButton);
            this.Controls.Add(this.confirmButton);
            this.Controls.Add(this.clearButton);
            this.Controls.Add(this.deleteCustomAnnotationButton);
            this.Controls.Add(this.annotationTypeLabel);
            this.Controls.Add(this.annotationColorLabel);
            this.Controls.Add(this.textColorLabel);
            this.Controls.Add(this.highlightColorLabel);
            this.Controls.Add(this.textSettingsLabel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "AnnotationToolForm";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "文字标注工具";
            this.Load += new System.EventHandler(this.AnnotationToolForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.ComponentModel.IContainer components;
    }
}
