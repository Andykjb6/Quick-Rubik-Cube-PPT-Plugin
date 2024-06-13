namespace 课件帮PPT助手
{
    partial class TimerForm
    {
        private System.Windows.Forms.TextBox timeTextBox;
        private System.Windows.Forms.Button startButton;
        private System.Windows.Forms.Button stopButton;
        private System.Windows.Forms.Button resetButton;
        private System.Windows.Forms.Button closeButton;
        private System.Windows.Forms.Button settingsButton;
        private System.Windows.Forms.Button darkModeButton;
        private System.Windows.Forms.Timer timer;

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TimerForm));
            this.timeTextBox = new System.Windows.Forms.TextBox();
            this.startButton = new System.Windows.Forms.Button();
            this.stopButton = new System.Windows.Forms.Button();
            this.resetButton = new System.Windows.Forms.Button();
            this.closeButton = new System.Windows.Forms.Button();
            this.settingsButton = new System.Windows.Forms.Button();
            this.darkModeButton = new System.Windows.Forms.Button();
            this.timer = new System.Windows.Forms.Timer(this.components);
            this.buttonPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.buttonPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // timeTextBox
            // 
            this.timeTextBox.BackColor = System.Drawing.Color.White;
            this.timeTextBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.timeTextBox.Font = new System.Drawing.Font("Arial", 40F, System.Drawing.FontStyle.Bold);
            this.timeTextBox.ForeColor = System.Drawing.Color.Black;
            this.timeTextBox.Location = new System.Drawing.Point(17, 30);
            this.timeTextBox.Name = "timeTextBox";
            this.timeTextBox.Size = new System.Drawing.Size(455, 123);
            this.timeTextBox.TabIndex = 0;
            this.timeTextBox.Text = "00:00:00";
            this.timeTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // startButton
            // 
            this.startButton.BackColor = System.Drawing.Color.Transparent;
            this.startButton.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("startButton.BackgroundImage")));
            this.startButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.startButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(32)))), ((int)(((byte)(59)))), ((int)(((byte)(243)))));
            this.startButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.startButton.Location = new System.Drawing.Point(3, 3);
            this.startButton.Name = "startButton";
            this.startButton.Size = new System.Drawing.Size(35, 35);
            this.startButton.TabIndex = 0;
            this.startButton.UseVisualStyleBackColor = false;
            this.startButton.Click += new System.EventHandler(this.StartButton_Click);
            // 
            // stopButton
            // 
            this.stopButton.BackColor = System.Drawing.Color.Transparent;
            this.stopButton.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("stopButton.BackgroundImage")));
            this.stopButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.stopButton.Enabled = false;
            this.stopButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(32)))), ((int)(((byte)(59)))), ((int)(((byte)(243)))));
            this.stopButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.stopButton.Location = new System.Drawing.Point(44, 3);
            this.stopButton.Name = "stopButton";
            this.stopButton.Size = new System.Drawing.Size(35, 35);
            this.stopButton.TabIndex = 1;
            this.stopButton.UseVisualStyleBackColor = false;
            this.stopButton.Click += new System.EventHandler(this.StopButton_Click);
            // 
            // resetButton
            // 
            this.resetButton.BackColor = System.Drawing.Color.Transparent;
            this.resetButton.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("resetButton.BackgroundImage")));
            this.resetButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.resetButton.Enabled = false;
            this.resetButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(32)))), ((int)(((byte)(59)))), ((int)(((byte)(243)))));
            this.resetButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.resetButton.Location = new System.Drawing.Point(85, 3);
            this.resetButton.Name = "resetButton";
            this.resetButton.Size = new System.Drawing.Size(35, 35);
            this.resetButton.TabIndex = 2;
            this.resetButton.UseVisualStyleBackColor = false;
            this.resetButton.Click += new System.EventHandler(this.ResetButton_Click);
            // 
            // closeButton
            // 
            this.closeButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.closeButton.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("closeButton.BackgroundImage")));
            this.closeButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.closeButton.FlatAppearance.BorderSize = 0;
            this.closeButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.closeButton.Location = new System.Drawing.Point(459, 3);
            this.closeButton.Name = "closeButton";
            this.closeButton.Size = new System.Drawing.Size(35, 35);
            this.closeButton.TabIndex = 2;
            this.closeButton.Click += new System.EventHandler(this.CloseButton_Click);
            // 
            // settingsButton
            // 
            this.settingsButton.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("settingsButton.BackgroundImage")));
            this.settingsButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.settingsButton.FlatAppearance.BorderSize = 0;
            this.settingsButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.settingsButton.Location = new System.Drawing.Point(0, 2);
            this.settingsButton.Name = "settingsButton";
            this.settingsButton.Size = new System.Drawing.Size(35, 35);
            this.settingsButton.TabIndex = 1;
            this.settingsButton.Click += new System.EventHandler(this.SettingsButton_Click);
            // 
            // darkModeButton
            // 
            this.darkModeButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(57)))), ((int)(((byte)(240)))));
            this.darkModeButton.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("darkModeButton.BackgroundImage")));
            this.darkModeButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.darkModeButton.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.darkModeButton.FlatAppearance.BorderSize = 0;
            this.darkModeButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.darkModeButton.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.darkModeButton.Location = new System.Drawing.Point(0, 215);
            this.darkModeButton.Name = "darkModeButton";
            this.darkModeButton.Size = new System.Drawing.Size(494, 38);
            this.darkModeButton.TabIndex = 3;
            this.darkModeButton.Text = "暗色模式";
            this.darkModeButton.UseVisualStyleBackColor = false;
            this.darkModeButton.Click += new System.EventHandler(this.DarkModeButton_Click);
            // 
            // timer
            // 
            this.timer.Interval = 1000;
            this.timer.Tick += new System.EventHandler(this.Timer_Tick);
            // 
            // buttonPanel
            // 
            this.buttonPanel.Controls.Add(this.startButton);
            this.buttonPanel.Controls.Add(this.stopButton);
            this.buttonPanel.Controls.Add(this.resetButton);
            this.buttonPanel.Location = new System.Drawing.Point(189, 159);
            this.buttonPanel.Name = "buttonPanel";
            this.buttonPanel.Size = new System.Drawing.Size(126, 40);
            this.buttonPanel.TabIndex = 4;
            // 
            // TimerForm
            // 
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(494, 253);
            this.Controls.Add(this.settingsButton);
            this.Controls.Add(this.closeButton);
            this.Controls.Add(this.darkModeButton);
            this.Controls.Add(this.buttonPanel);
            this.Controls.Add(this.timeTextBox);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "TimerForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "计时器";
            this.TopMost = true;
            this.Paint += new System.Windows.Forms.PaintEventHandler(this.TimerForm_Paint);
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.TimerForm_MouseDown);
            this.buttonPanel.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void TimerForm_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                this.Capture = false;
                System.Windows.Forms.Message m = System.Windows.Forms.Message.Create(this.Handle, 0xA1, new System.IntPtr(2), System.IntPtr.Zero);
                this.WndProc(ref m);
            }
        }

        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.FlowLayoutPanel buttonPanel;
    }
}
