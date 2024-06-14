using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace 课件帮PPT助手
{
    partial class CustomCloudTextGeneratorForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.TrackBar letterSpacingTrackBar;
        private System.Windows.Forms.CheckBox shadowCheckBox;
        private System.Windows.Forms.TextBox textBox;
        private System.Windows.Forms.ComboBox fontComboBox;
        private System.Windows.Forms.Button topColorButton;
        private System.Windows.Forms.Button middleColorButton;
        private System.Windows.Forms.Button bottomColorButton;
        private System.Windows.Forms.NumericUpDown middleOutlineNumericUpDown;
        private System.Windows.Forms.NumericUpDown bottomOutlineNumericUpDown;
        private System.Windows.Forms.NumericUpDown fontSizeNumericUpDown;
        private System.Windows.Forms.Button generateButton;
        private System.Windows.Forms.Button shadowColorButton;
        private System.Windows.Forms.TrackBar shadowBlurTrackBar;
        private System.Windows.Forms.TrackBar shadowTransparencyTrackBar;
        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage textSettingsPage;
        private System.Windows.Forms.TabPage shadowSettingsPage;
        private System.Windows.Forms.TabPage colorSettingsPage;
        private System.Windows.Forms.TabPage spacingSettingsPage;

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
            this.tabControl = new System.Windows.Forms.TabControl();
            this.textSettingsPage = new System.Windows.Forms.TabPage();
            this.fontComboBox = new System.Windows.Forms.ComboBox();
            this.letterSpacingTrackBar = new System.Windows.Forms.TrackBar();
            this.fontSizeNumericUpDown = new System.Windows.Forms.NumericUpDown();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox = new System.Windows.Forms.TextBox();
            this.generateButton = new System.Windows.Forms.Button();
            this.shadowSettingsPage = new System.Windows.Forms.TabPage();
            this.shadowBlurTrackBar = new System.Windows.Forms.TrackBar();
            this.shadowTransparencyTrackBar = new System.Windows.Forms.TrackBar();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.shadowCheckBox = new System.Windows.Forms.CheckBox();
            this.shadowColorButton = new System.Windows.Forms.Button();
            this.colorSettingsPage = new System.Windows.Forms.TabPage();
            this.topColorButton = new System.Windows.Forms.Button();
            this.middleColorButton = new System.Windows.Forms.Button();
            this.bottomColorButton = new System.Windows.Forms.Button();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.spacingSettingsPage = new System.Windows.Forms.TabPage();
            this.middleOutlineNumericUpDown = new System.Windows.Forms.NumericUpDown();
            this.bottomOutlineNumericUpDown = new System.Windows.Forms.NumericUpDown();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.tabControl.SuspendLayout();
            this.textSettingsPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.letterSpacingTrackBar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.fontSizeNumericUpDown)).BeginInit();
            this.shadowSettingsPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.shadowBlurTrackBar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.shadowTransparencyTrackBar)).BeginInit();
            this.colorSettingsPage.SuspendLayout();
            this.spacingSettingsPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.middleOutlineNumericUpDown)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bottomOutlineNumericUpDown)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl
            // 
            this.tabControl.Controls.Add(this.textSettingsPage);
            this.tabControl.Controls.Add(this.shadowSettingsPage);
            this.tabControl.Controls.Add(this.colorSettingsPage);
            this.tabControl.Controls.Add(this.spacingSettingsPage);
            this.tabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl.Location = new System.Drawing.Point(0, 0);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(443, 600);
            this.tabControl.TabIndex = 0;
            // 
            // textSettingsPage
            // 
            this.textSettingsPage.Controls.Add(this.fontComboBox);
            this.textSettingsPage.Controls.Add(this.letterSpacingTrackBar);
            this.textSettingsPage.Controls.Add(this.fontSizeNumericUpDown);
            this.textSettingsPage.Controls.Add(this.label3);
            this.textSettingsPage.Controls.Add(this.label2);
            this.textSettingsPage.Controls.Add(this.label1);
            this.textSettingsPage.Controls.Add(this.textBox);
            this.textSettingsPage.Controls.Add(this.generateButton);
            this.textSettingsPage.Location = new System.Drawing.Point(8, 39);
            this.textSettingsPage.Name = "textSettingsPage";
            this.textSettingsPage.Padding = new System.Windows.Forms.Padding(3);
            this.textSettingsPage.Size = new System.Drawing.Size(427, 553);
            this.textSettingsPage.TabIndex = 0;
            this.textSettingsPage.Text = "文本设置";
            this.textSettingsPage.UseVisualStyleBackColor = true;
            // 
            // fontComboBox
            // 
            this.fontComboBox.Font = new System.Drawing.Font("宋体", 10F);
            this.fontComboBox.Location = new System.Drawing.Point(20, 143);
            this.fontComboBox.Name = "fontComboBox";
            this.fontComboBox.Size = new System.Drawing.Size(380, 35);
            this.fontComboBox.TabIndex = 1;
            // 
            // letterSpacingTrackBar
            // 
            this.letterSpacingTrackBar.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(249)))), ((int)(((byte)(249)))), ((int)(((byte)(249)))));
            this.letterSpacingTrackBar.Location = new System.Drawing.Point(20, 444);
            this.letterSpacingTrackBar.Maximum = 100;
            this.letterSpacingTrackBar.Minimum = -100;
            this.letterSpacingTrackBar.Name = "letterSpacingTrackBar";
            this.letterSpacingTrackBar.Size = new System.Drawing.Size(380, 90);
            this.letterSpacingTrackBar.TabIndex = 0;
            this.letterSpacingTrackBar.Scroll += new System.EventHandler(this.letterSpacingTrackBar_Scroll);
            // 
            // fontSizeNumericUpDown
            // 
            this.fontSizeNumericUpDown.Font = new System.Drawing.Font("微软雅黑", 10.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.fontSizeNumericUpDown.Location = new System.Drawing.Point(20, 330);
            this.fontSizeNumericUpDown.Maximum = new decimal(new int[] {
            600,
            0,
            0,
            0});
            this.fontSizeNumericUpDown.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.fontSizeNumericUpDown.Name = "fontSizeNumericUpDown";
            this.fontSizeNumericUpDown.Size = new System.Drawing.Size(380, 46);
            this.fontSizeNumericUpDown.TabIndex = 3;
            this.fontSizeNumericUpDown.Value = new decimal(new int[] {
            130,
            0,
            0,
            0});
            this.fontSizeNumericUpDown.ValueChanged += new System.EventHandler(this.fontSizeNumericUpDown_ValueChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(17, 401);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(134, 31);
            this.label3.TabIndex = 6;
            this.label3.Text = "字符间距：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(17, 293);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(134, 31);
            this.label2.TabIndex = 5;
            this.label2.Text = "字号大小：";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(17, 105);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(134, 31);
            this.label1.TabIndex = 4;
            this.label1.Text = "选择字体：";
            // 
            // textBox
            // 
            this.textBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(241)))), ((int)(((byte)(255)))));
            this.textBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F);
            this.textBox.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(3)))), ((int)(((byte)(46)))), ((int)(((byte)(195)))));
            this.textBox.Location = new System.Drawing.Point(20, 20);
            this.textBox.Name = "textBox";
            this.textBox.Size = new System.Drawing.Size(380, 68);
            this.textBox.TabIndex = 0;
            this.textBox.Text = "请输入文本";
            this.textBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // generateButton
            // 
            this.generateButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(40)))), ((int)(((byte)(87)))), ((int)(((byte)(247)))));
            this.generateButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(18)))), ((int)(((byte)(69)))), ((int)(((byte)(243)))));
            this.generateButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(72)))), ((int)(((byte)(105)))), ((int)(((byte)(218)))));
            this.generateButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.generateButton.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.generateButton.ForeColor = System.Drawing.Color.White;
            this.generateButton.Location = new System.Drawing.Point(20, 202);
            this.generateButton.Name = "generateButton";
            this.generateButton.Size = new System.Drawing.Size(380, 61);
            this.generateButton.TabIndex = 2;
            this.generateButton.Text = "生成";
            this.generateButton.UseVisualStyleBackColor = false;
            this.generateButton.Click += new System.EventHandler(this.generateButton_Click);
            // 
            // shadowSettingsPage
            // 
            this.shadowSettingsPage.Controls.Add(this.shadowBlurTrackBar);
            this.shadowSettingsPage.Controls.Add(this.shadowTransparencyTrackBar);
            this.shadowSettingsPage.Controls.Add(this.label5);
            this.shadowSettingsPage.Controls.Add(this.label4);
            this.shadowSettingsPage.Controls.Add(this.shadowCheckBox);
            this.shadowSettingsPage.Controls.Add(this.shadowColorButton);
            this.shadowSettingsPage.Location = new System.Drawing.Point(8, 39);
            this.shadowSettingsPage.Name = "shadowSettingsPage";
            this.shadowSettingsPage.Padding = new System.Windows.Forms.Padding(3);
            this.shadowSettingsPage.Size = new System.Drawing.Size(427, 553);
            this.shadowSettingsPage.TabIndex = 1;
            this.shadowSettingsPage.Text = "阴影设置";
            this.shadowSettingsPage.UseVisualStyleBackColor = true;
            // 
            // shadowBlurTrackBar
            // 
            this.shadowBlurTrackBar.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(249)))), ((int)(((byte)(249)))), ((int)(((byte)(249)))));
            this.shadowBlurTrackBar.Location = new System.Drawing.Point(20, 203);
            this.shadowBlurTrackBar.Maximum = 100;
            this.shadowBlurTrackBar.Name = "shadowBlurTrackBar";
            this.shadowBlurTrackBar.Size = new System.Drawing.Size(380, 90);
            this.shadowBlurTrackBar.TabIndex = 3;
            this.shadowBlurTrackBar.Value = 25;
            this.shadowBlurTrackBar.Scroll += new System.EventHandler(this.shadowBlurTrackBar_Scroll);
            // 
            // shadowTransparencyTrackBar
            // 
            this.shadowTransparencyTrackBar.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(249)))), ((int)(((byte)(249)))), ((int)(((byte)(249)))));
            this.shadowTransparencyTrackBar.Location = new System.Drawing.Point(20, 335);
            this.shadowTransparencyTrackBar.Maximum = 100;
            this.shadowTransparencyTrackBar.Name = "shadowTransparencyTrackBar";
            this.shadowTransparencyTrackBar.Size = new System.Drawing.Size(380, 90);
            this.shadowTransparencyTrackBar.TabIndex = 4;
            this.shadowTransparencyTrackBar.Value = 65;
            this.shadowTransparencyTrackBar.Scroll += new System.EventHandler(this.shadowTransparencyTrackBar_Scroll);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.Location = new System.Drawing.Point(16, 304);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(134, 31);
            this.label5.TabIndex = 6;
            this.label5.Text = "阴影透明：";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(20, 171);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(134, 31);
            this.label4.TabIndex = 5;
            this.label4.Text = "阴影模糊：";
            // 
            // shadowCheckBox
            // 
            this.shadowCheckBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.shadowCheckBox.Location = new System.Drawing.Point(20, 20);
            this.shadowCheckBox.Name = "shadowCheckBox";
            this.shadowCheckBox.Size = new System.Drawing.Size(380, 40);
            this.shadowCheckBox.TabIndex = 1;
            this.shadowCheckBox.Text = "阴影开关";
            this.shadowCheckBox.UseVisualStyleBackColor = true;
            this.shadowCheckBox.CheckedChanged += new System.EventHandler(this.shadowCheckBox_CheckedChanged);
            // 
            // shadowColorButton
            // 
            this.shadowColorButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(40)))), ((int)(((byte)(87)))), ((int)(((byte)(247)))));
            this.shadowColorButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(18)))), ((int)(((byte)(69)))), ((int)(((byte)(243)))));
            this.shadowColorButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(72)))), ((int)(((byte)(105)))), ((int)(((byte)(218)))));
            this.shadowColorButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.shadowColorButton.ForeColor = System.Drawing.Color.White;
            this.shadowColorButton.Location = new System.Drawing.Point(20, 70);
            this.shadowColorButton.Name = "shadowColorButton";
            this.shadowColorButton.Size = new System.Drawing.Size(380, 51);
            this.shadowColorButton.TabIndex = 2;
            this.shadowColorButton.Text = "更改阴影颜色";
            this.shadowColorButton.UseVisualStyleBackColor = false;
            this.shadowColorButton.Click += new System.EventHandler(this.shadowColorButton_Click);
            // 
            // colorSettingsPage
            // 
            this.colorSettingsPage.Controls.Add(this.topColorButton);
            this.colorSettingsPage.Controls.Add(this.middleColorButton);
            this.colorSettingsPage.Controls.Add(this.bottomColorButton);
            this.colorSettingsPage.Controls.Add(this.label8);
            this.colorSettingsPage.Controls.Add(this.label7);
            this.colorSettingsPage.Controls.Add(this.label6);
            this.colorSettingsPage.Location = new System.Drawing.Point(8, 39);
            this.colorSettingsPage.Name = "colorSettingsPage";
            this.colorSettingsPage.Padding = new System.Windows.Forms.Padding(3);
            this.colorSettingsPage.Size = new System.Drawing.Size(427, 553);
            this.colorSettingsPage.TabIndex = 2;
            this.colorSettingsPage.Text = "颜色设置";
            this.colorSettingsPage.UseVisualStyleBackColor = true;
            // 
            // topColorButton
            // 
            this.topColorButton.BackColor = System.Drawing.Color.Black;
            this.topColorButton.Location = new System.Drawing.Point(24, 64);
            this.topColorButton.Name = "topColorButton";
            this.topColorButton.Size = new System.Drawing.Size(380, 51);
            this.topColorButton.TabIndex = 0;
            this.topColorButton.UseVisualStyleBackColor = false;
            this.topColorButton.Click += new System.EventHandler(this.topColorButton_Click);
            // 
            // middleColorButton
            // 
            this.middleColorButton.BackColor = System.Drawing.Color.White;
            this.middleColorButton.Location = new System.Drawing.Point(24, 171);
            this.middleColorButton.Name = "middleColorButton";
            this.middleColorButton.Size = new System.Drawing.Size(380, 51);
            this.middleColorButton.TabIndex = 1;
            this.middleColorButton.UseVisualStyleBackColor = false;
            this.middleColorButton.Click += new System.EventHandler(this.middleColorButton_Click);
            // 
            // bottomColorButton
            // 
            this.bottomColorButton.BackColor = System.Drawing.Color.Blue;
            this.bottomColorButton.Location = new System.Drawing.Point(24, 278);
            this.bottomColorButton.Name = "bottomColorButton";
            this.bottomColorButton.Size = new System.Drawing.Size(380, 51);
            this.bottomColorButton.TabIndex = 2;
            this.bottomColorButton.UseVisualStyleBackColor = false;
            this.bottomColorButton.Click += new System.EventHandler(this.bottomColorButton_Click);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label8.Location = new System.Drawing.Point(24, 247);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(134, 31);
            this.label8.TabIndex = 5;
            this.label8.Text = "底层颜色：";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label7.Location = new System.Drawing.Point(24, 139);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(134, 31);
            this.label7.TabIndex = 4;
            this.label7.Text = "中层颜色：";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.Location = new System.Drawing.Point(24, 31);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(134, 31);
            this.label6.TabIndex = 3;
            this.label6.Text = "顶层颜色：";
            // 
            // spacingSettingsPage
            // 
            this.spacingSettingsPage.Controls.Add(this.middleOutlineNumericUpDown);
            this.spacingSettingsPage.Controls.Add(this.bottomOutlineNumericUpDown);
            this.spacingSettingsPage.Controls.Add(this.label10);
            this.spacingSettingsPage.Controls.Add(this.label9);
            this.spacingSettingsPage.Location = new System.Drawing.Point(8, 39);
            this.spacingSettingsPage.Name = "spacingSettingsPage";
            this.spacingSettingsPage.Padding = new System.Windows.Forms.Padding(3);
            this.spacingSettingsPage.Size = new System.Drawing.Size(427, 553);
            this.spacingSettingsPage.TabIndex = 3;
            this.spacingSettingsPage.Text = "轮廓设置";
            this.spacingSettingsPage.UseVisualStyleBackColor = true;
            // 
            // middleOutlineNumericUpDown
            // 
            this.middleOutlineNumericUpDown.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.middleOutlineNumericUpDown.Location = new System.Drawing.Point(20, 79);
            this.middleOutlineNumericUpDown.Name = "middleOutlineNumericUpDown";
            this.middleOutlineNumericUpDown.Size = new System.Drawing.Size(380, 39);
            this.middleOutlineNumericUpDown.TabIndex = 0;
            this.middleOutlineNumericUpDown.Value = new decimal(new int[] {
            45,
            0,
            0,
            0});
            this.middleOutlineNumericUpDown.ValueChanged += new System.EventHandler(this.middleOutlineNumericUpDown_ValueChanged);
            // 
            // bottomOutlineNumericUpDown
            // 
            this.bottomOutlineNumericUpDown.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.bottomOutlineNumericUpDown.Location = new System.Drawing.Point(20, 182);
            this.bottomOutlineNumericUpDown.Name = "bottomOutlineNumericUpDown";
            this.bottomOutlineNumericUpDown.Size = new System.Drawing.Size(380, 39);
            this.bottomOutlineNumericUpDown.TabIndex = 1;
            this.bottomOutlineNumericUpDown.Value = new decimal(new int[] {
            55,
            0,
            0,
            0});
            this.bottomOutlineNumericUpDown.ValueChanged += new System.EventHandler(this.bottomOutlineNumericUpDown_ValueChanged);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label10.Location = new System.Drawing.Point(16, 147);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(134, 31);
            this.label10.TabIndex = 3;
            this.label10.Text = "底层轮廓：";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label9.Location = new System.Drawing.Point(16, 44);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(134, 31);
            this.label9.TabIndex = 2;
            this.label9.Text = "中层轮廓：";
            // 
            // CustomCloudTextGeneratorForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(192F, 192F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(250)))), ((int)(((byte)(255)))));
            this.ClientSize = new System.Drawing.Size(443, 600);
            this.Controls.Add(this.tabControl);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.Name = "CustomCloudTextGeneratorForm";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "自定义云朵字生成";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.CustomCloudTextGeneratorForm_Load);
            this.tabControl.ResumeLayout(false);
            this.textSettingsPage.ResumeLayout(false);
            this.textSettingsPage.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.letterSpacingTrackBar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.fontSizeNumericUpDown)).EndInit();
            this.shadowSettingsPage.ResumeLayout(false);
            this.shadowSettingsPage.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.shadowBlurTrackBar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.shadowTransparencyTrackBar)).EndInit();
            this.colorSettingsPage.ResumeLayout(false);
            this.colorSettingsPage.PerformLayout();
            this.spacingSettingsPage.ResumeLayout(false);
            this.spacingSettingsPage.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.middleOutlineNumericUpDown)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bottomOutlineNumericUpDown)).EndInit();
            this.ResumeLayout(false);

        }

        private Label label3;
        private Label label2;
        private Label label1;
        private Label label4;
        private Label label5;
        private Label label8;
        private Label label7;
        private Label label6;
        private Label label10;
        private Label label9;
    }
}
