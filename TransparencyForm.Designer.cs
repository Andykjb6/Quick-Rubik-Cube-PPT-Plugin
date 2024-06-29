using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Threading.Tasks;

namespace 课件帮PPT助手
{
    partial class TransparencyForm
    {
        private System.Windows.Forms.PictureBox pictureBox;
        private System.Windows.Forms.RadioButton horizontalRadioButton;
        private System.Windows.Forms.RadioButton verticalRadioButton;
        private System.Windows.Forms.RadioButton fullTransparencyRadioButton;
        private System.Windows.Forms.RadioButton radialTransparencyRadioButton;
        private System.Windows.Forms.RadioButton diagonalTransparencyRadioButton;
        private System.Windows.Forms.TrackBar transparencyTrackBar;
        private System.Windows.Forms.Label transparencyLabel;
        private System.Windows.Forms.ComboBox flipComboBox;
        private System.Windows.Forms.Button importButton;
        private System.Windows.Forms.Button exportButton;
        private System.Windows.Forms.Button colorOptionsButton;
        private System.Windows.Forms.Panel colorOptionsPanel;
        private System.Windows.Forms.CheckBox grayscaleCheckBox;
        private System.Windows.Forms.Button colorOverlayButton;
        private System.Windows.Forms.Button resetColorButton;

        private void InitializeComponent()
        {
            this.pictureBox = new System.Windows.Forms.PictureBox();
            this.horizontalRadioButton = new System.Windows.Forms.RadioButton();
            this.verticalRadioButton = new System.Windows.Forms.RadioButton();
            this.fullTransparencyRadioButton = new System.Windows.Forms.RadioButton();
            this.radialTransparencyRadioButton = new System.Windows.Forms.RadioButton();
            this.diagonalTransparencyRadioButton = new System.Windows.Forms.RadioButton();
            this.transparencyTrackBar = new System.Windows.Forms.TrackBar();
            this.transparencyLabel = new System.Windows.Forms.Label();
            this.flipComboBox = new System.Windows.Forms.ComboBox();
            this.importButton = new System.Windows.Forms.Button();
            this.exportButton = new System.Windows.Forms.Button();
            this.colorOptionsButton = new System.Windows.Forms.Button();
            this.colorOptionsPanel = new System.Windows.Forms.Panel();
            this.colorOverlayButton = new System.Windows.Forms.Button();
            this.grayscaleCheckBox = new System.Windows.Forms.CheckBox();
            this.resetColorButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.transparencyTrackBar)).BeginInit();
            this.colorOptionsPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // pictureBox
            // 
            this.pictureBox.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pictureBox.Location = new System.Drawing.Point(12, 12);
            this.pictureBox.Name = "pictureBox";
            this.pictureBox.Size = new System.Drawing.Size(529, 303);
            this.pictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox.TabIndex = 1;
            this.pictureBox.TabStop = false;
            // 
            // horizontalRadioButton
            // 
            this.horizontalRadioButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.horizontalRadioButton.Location = new System.Drawing.Point(20, 378);
            this.horizontalRadioButton.Name = "horizontalRadioButton";
            this.horizontalRadioButton.Size = new System.Drawing.Size(104, 33);
            this.horizontalRadioButton.TabIndex = 4;
            this.horizontalRadioButton.Text = "水平";
            this.horizontalRadioButton.CheckedChanged += new System.EventHandler(this.RadioButton_CheckedChanged);
            // 
            // verticalRadioButton
            // 
            this.verticalRadioButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.verticalRadioButton.Location = new System.Drawing.Point(127, 378);
            this.verticalRadioButton.Name = "verticalRadioButton";
            this.verticalRadioButton.Size = new System.Drawing.Size(104, 33);
            this.verticalRadioButton.TabIndex = 3;
            this.verticalRadioButton.Text = "垂直";
            this.verticalRadioButton.CheckedChanged += new System.EventHandler(this.RadioButton_CheckedChanged);
            // 
            // fullTransparencyRadioButton
            // 
            this.fullTransparencyRadioButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.fullTransparencyRadioButton.Location = new System.Drawing.Point(234, 378);
            this.fullTransparencyRadioButton.Name = "fullTransparencyRadioButton";
            this.fullTransparencyRadioButton.Size = new System.Drawing.Size(104, 33);
            this.fullTransparencyRadioButton.TabIndex = 0;
            this.fullTransparencyRadioButton.Text = "整体";
            this.fullTransparencyRadioButton.CheckedChanged += new System.EventHandler(this.RadioButton_CheckedChanged);
            // 
            // radialTransparencyRadioButton
            // 
            this.radialTransparencyRadioButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.radialTransparencyRadioButton.Location = new System.Drawing.Point(341, 378);
            this.radialTransparencyRadioButton.Name = "radialTransparencyRadioButton";
            this.radialTransparencyRadioButton.Size = new System.Drawing.Size(104, 33);
            this.radialTransparencyRadioButton.TabIndex = 2;
            this.radialTransparencyRadioButton.Text = "径向";
            this.radialTransparencyRadioButton.CheckedChanged += new System.EventHandler(this.RadioButton_CheckedChanged);
            // 
            // diagonalTransparencyRadioButton
            // 
            this.diagonalTransparencyRadioButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.diagonalTransparencyRadioButton.Location = new System.Drawing.Point(448, 378);
            this.diagonalTransparencyRadioButton.Name = "diagonalTransparencyRadioButton";
            this.diagonalTransparencyRadioButton.Size = new System.Drawing.Size(104, 33);
            this.diagonalTransparencyRadioButton.TabIndex = 1;
            this.diagonalTransparencyRadioButton.Text = "对角";
            this.diagonalTransparencyRadioButton.CheckedChanged += new System.EventHandler(this.RadioButton_CheckedChanged);
            // 
            // transparencyTrackBar
            // 
            this.transparencyTrackBar.LargeChange = 10;
            this.transparencyTrackBar.Location = new System.Drawing.Point(0, 322);
            this.transparencyTrackBar.Maximum = 100;
            this.transparencyTrackBar.Name = "transparencyTrackBar";
            this.transparencyTrackBar.Size = new System.Drawing.Size(373, 90);
            this.transparencyTrackBar.TabIndex = 6;
            this.transparencyTrackBar.TickFrequency = 5;
            this.transparencyTrackBar.Scroll += new System.EventHandler(this.TrackBar_Scroll);
            // 
            // transparencyLabel
            // 
            this.transparencyLabel.AutoSize = true;
            this.transparencyLabel.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.transparencyLabel.Location = new System.Drawing.Point(371, 330);
            this.transparencyLabel.Name = "transparencyLabel";
            this.transparencyLabel.Size = new System.Drawing.Size(28, 31);
            this.transparencyLabel.TabIndex = 5;
            this.transparencyLabel.Text = "0";
            // 
            // flipComboBox
            // 
            this.flipComboBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.flipComboBox.Items.AddRange(new object[] {
            "无翻转",
            "水平翻转",
            "垂直翻转"});
            this.flipComboBox.Location = new System.Drawing.Point(405, 327);
            this.flipComboBox.Name = "flipComboBox";
            this.flipComboBox.Size = new System.Drawing.Size(136, 39);
            this.flipComboBox.TabIndex = 11;
            this.flipComboBox.SelectedIndexChanged += new System.EventHandler(this.FlipComboBox_SelectedIndexChanged);
            // 
            // importButton
            // 
            this.importButton.BackColor = System.Drawing.Color.White;
            this.importButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(25)))), ((int)(((byte)(75)))), ((int)(((byte)(249)))));
            this.importButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(206)))), ((int)(((byte)(217)))), ((int)(((byte)(254)))));
            this.importButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(233)))), ((int)(((byte)(238)))), ((int)(((byte)(255)))));
            this.importButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.importButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.importButton.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(74)))), ((int)(((byte)(249)))));
            this.importButton.Location = new System.Drawing.Point(20, 427);
            this.importButton.Name = "importButton";
            this.importButton.Size = new System.Drawing.Size(248, 44);
            this.importButton.TabIndex = 12;
            this.importButton.Text = "所选导入";
            this.importButton.UseVisualStyleBackColor = false;
            this.importButton.Click += new System.EventHandler(this.ImportButton_Click);
            // 
            // exportButton
            // 
            this.exportButton.BackColor = System.Drawing.Color.White;
            this.exportButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(25)))), ((int)(((byte)(75)))), ((int)(((byte)(249)))));
            this.exportButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(206)))), ((int)(((byte)(217)))), ((int)(((byte)(254)))));
            this.exportButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(233)))), ((int)(((byte)(238)))), ((int)(((byte)(255)))));
            this.exportButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.exportButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.exportButton.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(74)))), ((int)(((byte)(249)))));
            this.exportButton.Location = new System.Drawing.Point(293, 427);
            this.exportButton.Name = "exportButton";
            this.exportButton.Size = new System.Drawing.Size(248, 44);
            this.exportButton.TabIndex = 13;
            this.exportButton.Text = "导出至幻灯片";
            this.exportButton.UseVisualStyleBackColor = false;
            this.exportButton.Click += new System.EventHandler(this.ExportButton_Click);
            // 
            // colorOptionsButton
            // 
            this.colorOptionsButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(25)))), ((int)(((byte)(75)))), ((int)(((byte)(249)))));
            this.colorOptionsButton.FlatAppearance.BorderSize = 0;
            this.colorOptionsButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.colorOptionsButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.colorOptionsButton.ForeColor = System.Drawing.Color.White;
            this.colorOptionsButton.Location = new System.Drawing.Point(18, 477);
            this.colorOptionsButton.Name = "colorOptionsButton";
            this.colorOptionsButton.Size = new System.Drawing.Size(523, 43);
            this.colorOptionsButton.TabIndex = 7;
            this.colorOptionsButton.Text = "颜色选项";
            this.colorOptionsButton.UseVisualStyleBackColor = false;
            this.colorOptionsButton.Click += new System.EventHandler(this.ColorOptionsButton_Click);
            // 
            // colorOptionsPanel
            // 
            this.colorOptionsPanel.Controls.Add(this.colorOverlayButton);
            this.colorOptionsPanel.Controls.Add(this.grayscaleCheckBox);
            this.colorOptionsPanel.Controls.Add(this.resetColorButton);
            this.colorOptionsPanel.Location = new System.Drawing.Point(18, 526);
            this.colorOptionsPanel.Name = "colorOptionsPanel";
            this.colorOptionsPanel.Size = new System.Drawing.Size(523, 120);
            this.colorOptionsPanel.TabIndex = 14;
            this.colorOptionsPanel.Visible = false;
            // 
            // colorOverlayButton
            // 
            this.colorOverlayButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(242)))), ((int)(((byte)(255)))));
            this.colorOverlayButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(74)))), ((int)(((byte)(249)))));
            this.colorOverlayButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(193)))), ((int)(((byte)(207)))), ((int)(((byte)(255)))));
            this.colorOverlayButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(233)))), ((int)(((byte)(238)))), ((int)(((byte)(255)))));
            this.colorOverlayButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.colorOverlayButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.colorOverlayButton.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(74)))), ((int)(((byte)(249)))));
            this.colorOverlayButton.Location = new System.Drawing.Point(219, 24);
            this.colorOverlayButton.Name = "colorOverlayButton";
            this.colorOverlayButton.Size = new System.Drawing.Size(75, 45);
            this.colorOverlayButton.TabIndex = 9;
            this.colorOverlayButton.Text = "颜色";
            this.colorOverlayButton.UseVisualStyleBackColor = false;
            this.colorOverlayButton.Click += new System.EventHandler(this.ColorOverlayButton_Click);
            // 
            // grayscaleCheckBox
            // 
            this.grayscaleCheckBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.grayscaleCheckBox.Location = new System.Drawing.Point(120, 24);
            this.grayscaleCheckBox.Name = "grayscaleCheckBox";
            this.grayscaleCheckBox.Size = new System.Drawing.Size(104, 45);
            this.grayscaleCheckBox.TabIndex = 8;
            this.grayscaleCheckBox.Text = "灰度";
            this.grayscaleCheckBox.CheckedChanged += new System.EventHandler(this.GrayscaleCheckBox_CheckedChanged);
            // 
            // resetColorButton
            // 
            this.resetColorButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(242)))), ((int)(((byte)(255)))));
            this.resetColorButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(74)))), ((int)(((byte)(249)))));
            this.resetColorButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(193)))), ((int)(((byte)(207)))), ((int)(((byte)(255)))));
            this.resetColorButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(233)))), ((int)(((byte)(238)))), ((int)(((byte)(255)))));
            this.resetColorButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.resetColorButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.resetColorButton.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(74)))), ((int)(((byte)(249)))));
            this.resetColorButton.Location = new System.Drawing.Point(314, 24);
            this.resetColorButton.Name = "resetColorButton";
            this.resetColorButton.Size = new System.Drawing.Size(75, 45);
            this.resetColorButton.TabIndex = 10;
            this.resetColorButton.Text = "重置";
            this.resetColorButton.UseVisualStyleBackColor = false;
            this.resetColorButton.Click += new System.EventHandler(this.ResetColorButton_Click);
            // 
            // TransparencyForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(192F, 192F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(553, 539);
            this.Controls.Add(this.fullTransparencyRadioButton);
            this.Controls.Add(this.colorOptionsPanel);
            this.Controls.Add(this.pictureBox);
            this.Controls.Add(this.horizontalRadioButton);
            this.Controls.Add(this.verticalRadioButton);
            this.Controls.Add(this.radialTransparencyRadioButton);
            this.Controls.Add(this.diagonalTransparencyRadioButton);
            this.Controls.Add(this.transparencyTrackBar);
            this.Controls.Add(this.transparencyLabel);
            this.Controls.Add(this.flipComboBox);
            this.Controls.Add(this.importButton);
            this.Controls.Add(this.exportButton);
            this.Controls.Add(this.colorOptionsButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "TransparencyForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "图片透明化处理";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.transparencyTrackBar)).EndInit();
            this.colorOptionsPanel.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
    }
}

