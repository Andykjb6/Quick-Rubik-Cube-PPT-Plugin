namespace 课件帮PPT助手
{
    partial class SingleObjectForm
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

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SingleObjectForm));
            this.radiusLabel = new System.Windows.Forms.Label();
            this.radiusTrackBar = new System.Windows.Forms.TrackBar();
            this.initialRotationLabel = new System.Windows.Forms.Label();
            this.initialRotationUpDown = new System.Windows.Forms.NumericUpDown();
            this.finalRotationLabel = new System.Windows.Forms.Label();
            this.finalRotationUpDown = new System.Windows.Forms.NumericUpDown();
            this.sizeIncrementLabel = new System.Windows.Forms.Label();
            this.sizeIncrementTrackBar = new System.Windows.Forms.TrackBar();
            this.copyCountLabel = new System.Windows.Forms.Label();
            this.copyCountTrackBar = new System.Windows.Forms.TrackBar();
            this.resetButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.radiusTrackBar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.initialRotationUpDown)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.finalRotationUpDown)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sizeIncrementTrackBar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.copyCountTrackBar)).BeginInit();
            this.SuspendLayout();
            // 
            // radiusLabel
            // 
            this.radiusLabel.AutoSize = true;
            this.radiusLabel.Font = new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Bold);
            this.radiusLabel.Location = new System.Drawing.Point(44, 45);
            this.radiusLabel.Name = "radiusLabel";
            this.radiusLabel.Size = new System.Drawing.Size(139, 27);
            this.radiusLabel.TabIndex = 0;
            this.radiusLabel.Text = "环形半径:";
            // 
            // radiusTrackBar
            // 
            this.radiusTrackBar.Location = new System.Drawing.Point(182, 31);
            this.radiusTrackBar.Maximum = 500;
            this.radiusTrackBar.Minimum = 10;
            this.radiusTrackBar.Name = "radiusTrackBar";
            this.radiusTrackBar.Size = new System.Drawing.Size(350, 90);
            this.radiusTrackBar.TabIndex = 1;
            this.radiusTrackBar.Value = 100;
            // 
            // initialRotationLabel
            // 
            this.initialRotationLabel.AutoSize = true;
            this.initialRotationLabel.Location = new System.Drawing.Point(180, 255);
            this.initialRotationLabel.Name = "initialRotationLabel";
            this.initialRotationLabel.Size = new System.Drawing.Size(70, 24);
            this.initialRotationLabel.TabIndex = 2;
            this.initialRotationLabel.Text = "起始:";
            // 
            // initialRotationUpDown
            // 
            this.initialRotationUpDown.Location = new System.Drawing.Point(256, 244);
            this.initialRotationUpDown.Maximum = new decimal(new int[] {
            360,
            0,
            0,
            0});
            this.initialRotationUpDown.Name = "initialRotationUpDown";
            this.initialRotationUpDown.Size = new System.Drawing.Size(80, 35);
            this.initialRotationUpDown.TabIndex = 3;
            // 
            // finalRotationLabel
            // 
            this.finalRotationLabel.AutoSize = true;
            this.finalRotationLabel.Location = new System.Drawing.Point(353, 255);
            this.finalRotationLabel.Name = "finalRotationLabel";
            this.finalRotationLabel.Size = new System.Drawing.Size(70, 24);
            this.finalRotationLabel.TabIndex = 4;
            this.finalRotationLabel.Text = "终点:";
            // 
            // finalRotationUpDown
            // 
            this.finalRotationUpDown.Location = new System.Drawing.Point(441, 244);
            this.finalRotationUpDown.Maximum = new decimal(new int[] {
            360,
            0,
            0,
            0});
            this.finalRotationUpDown.Name = "finalRotationUpDown";
            this.finalRotationUpDown.Size = new System.Drawing.Size(80, 35);
            this.finalRotationUpDown.TabIndex = 5;
            this.finalRotationUpDown.ValueChanged += new System.EventHandler(this.finalRotationUpDown_ValueChanged);
            // 
            // sizeIncrementLabel
            // 
            this.sizeIncrementLabel.AutoSize = true;
            this.sizeIncrementLabel.Font = new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Bold);
            this.sizeIncrementLabel.Location = new System.Drawing.Point(44, 185);
            this.sizeIncrementLabel.Name = "sizeIncrementLabel";
            this.sizeIncrementLabel.Size = new System.Drawing.Size(139, 27);
            this.sizeIncrementLabel.TabIndex = 6;
            this.sizeIncrementLabel.Text = "螺旋递进:";
            // 
            // sizeIncrementTrackBar
            // 
            this.sizeIncrementTrackBar.LargeChange = 1;
            this.sizeIncrementTrackBar.Location = new System.Drawing.Point(182, 162);
            this.sizeIncrementTrackBar.Maximum = 50;
            this.sizeIncrementTrackBar.Minimum = -50;
            this.sizeIncrementTrackBar.Name = "sizeIncrementTrackBar";
            this.sizeIncrementTrackBar.Size = new System.Drawing.Size(350, 90);
            this.sizeIncrementTrackBar.TabIndex = 7;
            // 
            // copyCountLabel
            // 
            this.copyCountLabel.AutoSize = true;
            this.copyCountLabel.Font = new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Bold);
            this.copyCountLabel.Location = new System.Drawing.Point(44, 112);
            this.copyCountLabel.Name = "copyCountLabel";
            this.copyCountLabel.Size = new System.Drawing.Size(139, 27);
            this.copyCountLabel.TabIndex = 8;
            this.copyCountLabel.Text = "复制数量:";
            // 
            // copyCountTrackBar
            // 
            this.copyCountTrackBar.Location = new System.Drawing.Point(182, 101);
            this.copyCountTrackBar.Maximum = 50;
            this.copyCountTrackBar.Name = "copyCountTrackBar";
            this.copyCountTrackBar.Size = new System.Drawing.Size(350, 90);
            this.copyCountTrackBar.TabIndex = 9;
            // 
            // resetButton
            // 
            this.resetButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(58)))), ((int)(((byte)(238)))));
            this.resetButton.FlatAppearance.BorderSize = 0;
            this.resetButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(16)))), ((int)(((byte)(45)))), ((int)(((byte)(228)))));
            this.resetButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(91)))), ((int)(((byte)(238)))));
            this.resetButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.resetButton.Font = new System.Drawing.Font("宋体", 10.875F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.resetButton.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.resetButton.Location = new System.Drawing.Point(49, 322);
            this.resetButton.Name = "resetButton";
            this.resetButton.Size = new System.Drawing.Size(128, 53);
            this.resetButton.TabIndex = 10;
            this.resetButton.Text = "重置";
            this.resetButton.UseVisualStyleBackColor = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Bold);
            this.label1.Location = new System.Drawing.Point(45, 255);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(139, 27);
            this.label1.TabIndex = 11;
            this.label1.Text = "角度递进:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Bold);
            this.label2.ForeColor = System.Drawing.Color.Red;
            this.label2.Location = new System.Drawing.Point(339, 208);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(40, 27);
            this.label2.TabIndex = 12;
            this.label2.Text = "▲";
            // 
            // SingleObjectForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(192F, 192F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(575, 426);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.initialRotationLabel);
            this.Controls.Add(this.initialRotationUpDown);
            this.Controls.Add(this.finalRotationLabel);
            this.Controls.Add(this.finalRotationUpDown);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.sizeIncrementLabel);
            this.Controls.Add(this.sizeIncrementTrackBar);
            this.Controls.Add(this.resetButton);
            this.Controls.Add(this.radiusLabel);
            this.Controls.Add(this.copyCountLabel);
            this.Controls.Add(this.copyCountTrackBar);
            this.Controls.Add(this.radiusTrackBar);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SingleObjectForm";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "环形复制";
            this.TopMost = true;
            ((System.ComponentModel.ISupportInitialize)(this.radiusTrackBar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.initialRotationUpDown)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.finalRotationUpDown)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sizeIncrementTrackBar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.copyCountTrackBar)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.Windows.Forms.Label radiusLabel;
        private System.Windows.Forms.TrackBar radiusTrackBar;
        private System.Windows.Forms.Label initialRotationLabel;
        private System.Windows.Forms.NumericUpDown initialRotationUpDown;
        private System.Windows.Forms.Label finalRotationLabel;
        private System.Windows.Forms.NumericUpDown finalRotationUpDown;
        private System.Windows.Forms.Label sizeIncrementLabel;
        private System.Windows.Forms.TrackBar sizeIncrementTrackBar;
        private System.Windows.Forms.Label copyCountLabel;
        private System.Windows.Forms.TrackBar copyCountTrackBar;
        private System.Windows.Forms.Button resetButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}
