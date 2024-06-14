namespace 课件帮PPT助手
{
    partial class SettingsForm
    {
        private System.Windows.Forms.FontDialog fontDialog;
        private System.Windows.Forms.ColorDialog colorDialog;
        private System.Windows.Forms.Label timeFontLabel;
        private System.Windows.Forms.Label timeTextColorLabel;
        private System.Windows.Forms.Label countdownLabel;
        private System.Windows.Forms.Label backgroundColorLabel;
        private System.Windows.Forms.Label darkModeButtonColorLabel;
        private System.Windows.Forms.ComboBox fontComboBox;
        private System.Windows.Forms.TextBox timeTextColorBox;
        private System.Windows.Forms.RadioButton countdownRadioButton;
        private System.Windows.Forms.RadioButton stopwatchRadioButton;
        private System.Windows.Forms.TextBox backgroundColorBox;
        private System.Windows.Forms.TextBox darkModeButtonColorBox;
        private System.Windows.Forms.Button timeTextColorButton;
        private System.Windows.Forms.Button backgroundColorButton;
        private System.Windows.Forms.Button darkModeButtonColorButton;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button cancelButton;

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingsForm));
            this.fontDialog = new System.Windows.Forms.FontDialog();
            this.colorDialog = new System.Windows.Forms.ColorDialog();
            this.timeFontLabel = new System.Windows.Forms.Label();
            this.fontComboBox = new System.Windows.Forms.ComboBox();
            this.timeTextColorLabel = new System.Windows.Forms.Label();
            this.timeTextColorBox = new System.Windows.Forms.TextBox();
            this.timeTextColorButton = new System.Windows.Forms.Button();
            this.countdownLabel = new System.Windows.Forms.Label();
            this.countdownRadioButton = new System.Windows.Forms.RadioButton();
            this.stopwatchRadioButton = new System.Windows.Forms.RadioButton();
            this.backgroundColorLabel = new System.Windows.Forms.Label();
            this.backgroundColorBox = new System.Windows.Forms.TextBox();
            this.backgroundColorButton = new System.Windows.Forms.Button();
            this.darkModeButtonColorLabel = new System.Windows.Forms.Label();
            this.darkModeButtonColorBox = new System.Windows.Forms.TextBox();
            this.darkModeButtonColorButton = new System.Windows.Forms.Button();
            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // timeFontLabel
            // 
            this.timeFontLabel.AutoSize = true;
            this.timeFontLabel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.timeFontLabel.Font = new System.Drawing.Font("微软雅黑", 10.125F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.timeFontLabel.Location = new System.Drawing.Point(23, 33);
            this.timeFontLabel.Name = "timeFontLabel";
            this.timeFontLabel.Size = new System.Drawing.Size(150, 36);
            this.timeFontLabel.TabIndex = 0;
            this.timeFontLabel.Text = "时钟字体：";
            // 
            // fontComboBox
            // 
            this.fontComboBox.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.fontComboBox.Font = new System.Drawing.Font("宋体", 10F);
            this.fontComboBox.Location = new System.Drawing.Point(194, 33);
            this.fontComboBox.Name = "fontComboBox";
            this.fontComboBox.Size = new System.Drawing.Size(290, 35);
            this.fontComboBox.TabIndex = 1;
            // 
            // timeTextColorLabel
            // 
            this.timeTextColorLabel.AutoSize = true;
            this.timeTextColorLabel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.timeTextColorLabel.Font = new System.Drawing.Font("微软雅黑", 10.125F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.timeTextColorLabel.Location = new System.Drawing.Point(23, 93);
            this.timeTextColorLabel.Name = "timeTextColorLabel";
            this.timeTextColorLabel.Size = new System.Drawing.Size(150, 36);
            this.timeTextColorLabel.TabIndex = 2;
            this.timeTextColorLabel.Text = "时钟颜色：";
            // 
            // timeTextColorBox
            // 
            this.timeTextColorBox.Font = new System.Drawing.Font("宋体", 10F);
            this.timeTextColorBox.Location = new System.Drawing.Point(192, 93);
            this.timeTextColorBox.Name = "timeTextColorBox";
            this.timeTextColorBox.ReadOnly = true;
            this.timeTextColorBox.Size = new System.Drawing.Size(137, 38);
            this.timeTextColorBox.TabIndex = 3;
            // 
            // timeTextColorButton
            // 
            this.timeTextColorButton.Location = new System.Drawing.Point(354, 93);
            this.timeTextColorButton.Name = "timeTextColorButton";
            this.timeTextColorButton.Size = new System.Drawing.Size(130, 40);
            this.timeTextColorButton.TabIndex = 4;
            this.timeTextColorButton.Text = "选择颜色";
            this.timeTextColorButton.Click += new System.EventHandler(this.TimeTextColorButton_Click);
            // 
            // countdownLabel
            // 
            this.countdownLabel.AutoSize = true;
            this.countdownLabel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.countdownLabel.Font = new System.Drawing.Font("微软雅黑", 10.125F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.countdownLabel.Location = new System.Drawing.Point(23, 143);
            this.countdownLabel.Name = "countdownLabel";
            this.countdownLabel.Size = new System.Drawing.Size(150, 36);
            this.countdownLabel.TabIndex = 5;
            this.countdownLabel.Text = "计时模式：";
            // 
            // countdownRadioButton
            // 
            this.countdownRadioButton.Font = new System.Drawing.Font("宋体", 10F);
            this.countdownRadioButton.Location = new System.Drawing.Point(181, 143);
            this.countdownRadioButton.Name = "countdownRadioButton";
            this.countdownRadioButton.Size = new System.Drawing.Size(140, 40);
            this.countdownRadioButton.TabIndex = 6;
            this.countdownRadioButton.Text = "倒计时";
            // 
            // stopwatchRadioButton
            // 
            this.stopwatchRadioButton.Font = new System.Drawing.Font("宋体", 10F);
            this.stopwatchRadioButton.Location = new System.Drawing.Point(316, 143);
            this.stopwatchRadioButton.Name = "stopwatchRadioButton";
            this.stopwatchRadioButton.Size = new System.Drawing.Size(142, 40);
            this.stopwatchRadioButton.TabIndex = 7;
            this.stopwatchRadioButton.Text = "顺计时";
            // 
            // backgroundColorLabel
            // 
            this.backgroundColorLabel.AutoSize = true;
            this.backgroundColorLabel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.backgroundColorLabel.Font = new System.Drawing.Font("微软雅黑", 10.125F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.backgroundColorLabel.Location = new System.Drawing.Point(23, 193);
            this.backgroundColorLabel.Name = "backgroundColorLabel";
            this.backgroundColorLabel.Size = new System.Drawing.Size(150, 36);
            this.backgroundColorLabel.TabIndex = 8;
            this.backgroundColorLabel.Text = "背景颜色：";
            // 
            // backgroundColorBox
            // 
            this.backgroundColorBox.Font = new System.Drawing.Font("宋体", 10F);
            this.backgroundColorBox.Location = new System.Drawing.Point(192, 193);
            this.backgroundColorBox.Name = "backgroundColorBox";
            this.backgroundColorBox.ReadOnly = true;
            this.backgroundColorBox.Size = new System.Drawing.Size(137, 38);
            this.backgroundColorBox.TabIndex = 9;
            // 
            // backgroundColorButton
            // 
            this.backgroundColorButton.Location = new System.Drawing.Point(354, 193);
            this.backgroundColorButton.Name = "backgroundColorButton";
            this.backgroundColorButton.Size = new System.Drawing.Size(130, 40);
            this.backgroundColorButton.TabIndex = 10;
            this.backgroundColorButton.Text = "选择颜色";
            this.backgroundColorButton.Click += new System.EventHandler(this.BackgroundColorButton_Click);
            // 
            // darkModeButtonColorLabel
            // 
            this.darkModeButtonColorLabel.AutoSize = true;
            this.darkModeButtonColorLabel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.darkModeButtonColorLabel.Font = new System.Drawing.Font("微软雅黑", 10.125F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.darkModeButtonColorLabel.Location = new System.Drawing.Point(23, 243);
            this.darkModeButtonColorLabel.Name = "darkModeButtonColorLabel";
            this.darkModeButtonColorLabel.Size = new System.Drawing.Size(163, 36);
            this.darkModeButtonColorLabel.TabIndex = 11;
            this.darkModeButtonColorLabel.Text = "暗/浅模式：";
            // 
            // darkModeButtonColorBox
            // 
            this.darkModeButtonColorBox.Font = new System.Drawing.Font("宋体", 10F);
            this.darkModeButtonColorBox.Location = new System.Drawing.Point(192, 243);
            this.darkModeButtonColorBox.Name = "darkModeButtonColorBox";
            this.darkModeButtonColorBox.ReadOnly = true;
            this.darkModeButtonColorBox.Size = new System.Drawing.Size(137, 38);
            this.darkModeButtonColorBox.TabIndex = 12;
            // 
            // darkModeButtonColorButton
            // 
            this.darkModeButtonColorButton.Location = new System.Drawing.Point(354, 243);
            this.darkModeButtonColorButton.Name = "darkModeButtonColorButton";
            this.darkModeButtonColorButton.Size = new System.Drawing.Size(130, 40);
            this.darkModeButtonColorButton.TabIndex = 13;
            this.darkModeButtonColorButton.Text = "选择颜色";
            this.darkModeButtonColorButton.Click += new System.EventHandler(this.DarkModeButtonColorButton_Click);
            // 
            // okButton
            // 
            this.okButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(57)))), ((int)(((byte)(240)))));
            this.okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.okButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.okButton.Font = new System.Drawing.Font("微软雅黑", 10.875F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.okButton.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.okButton.Location = new System.Drawing.Point(137, 308);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(104, 50);
            this.okButton.TabIndex = 14;
            this.okButton.Text = "确定";
            this.okButton.UseVisualStyleBackColor = false;
            // 
            // cancelButton
            // 
            this.cancelButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(57)))), ((int)(((byte)(240)))));
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.cancelButton.Font = new System.Drawing.Font("微软雅黑", 10.875F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cancelButton.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.cancelButton.Location = new System.Drawing.Point(268, 308);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(104, 50);
            this.cancelButton.TabIndex = 15;
            this.cancelButton.Text = "取消";
            this.cancelButton.UseVisualStyleBackColor = false;
            // 
            // SettingsForm
            // 
            this.AcceptButton = this.okButton;
            this.CancelButton = this.cancelButton;
            this.ClientSize = new System.Drawing.Size(513, 387);
            this.Controls.Add(this.stopwatchRadioButton);
            this.Controls.Add(this.fontComboBox);
            this.Controls.Add(this.timeTextColorBox);
            this.Controls.Add(this.timeTextColorButton);
            this.Controls.Add(this.countdownRadioButton);
            this.Controls.Add(this.backgroundColorBox);
            this.Controls.Add(this.backgroundColorButton);
            this.Controls.Add(this.darkModeButtonColorBox);
            this.Controls.Add(this.darkModeButtonColorButton);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.timeFontLabel);
            this.Controls.Add(this.timeTextColorLabel);
            this.Controls.Add(this.countdownLabel);
            this.Controls.Add(this.backgroundColorLabel);
            this.Controls.Add(this.darkModeButtonColorLabel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SettingsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "设置面板";
            this.Load += new System.EventHandler(this.SettingsForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
    }
}
