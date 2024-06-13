namespace 课件帮PPT助手
{
    partial class MultipleObjectsForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MultipleObjectsForm));
            this.radiusLabel = new System.Windows.Forms.Label();
            this.radiusTrackBar = new System.Windows.Forms.TrackBar();
            this.initialRotationLabel = new System.Windows.Forms.Label();
            this.initialRotationUpDown = new System.Windows.Forms.NumericUpDown();
            this.finalRotationLabel = new System.Windows.Forms.Label();
            this.finalRotationUpDown = new System.Windows.Forms.NumericUpDown();
            this.sizeIncrementLabel = new System.Windows.Forms.Label();
            this.sizeIncrementTrackBar = new System.Windows.Forms.TrackBar();
            this.resetButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.radiusTrackBar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.initialRotationUpDown)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.finalRotationUpDown)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sizeIncrementTrackBar)).BeginInit();
            this.SuspendLayout();
            // 
            // radiusLabel
            // 
            this.radiusLabel.AutoSize = true;
            this.radiusLabel.Font = new System.Drawing.Font("宋体", 11F, System.Drawing.FontStyle.Bold);
            this.radiusLabel.Location = new System.Drawing.Point(24, 27);
            this.radiusLabel.Name = "radiusLabel";
            this.radiusLabel.Size = new System.Drawing.Size(153, 30);
            this.radiusLabel.TabIndex = 0;
            this.radiusLabel.Text = "环形半径:";
            // 
            // radiusTrackBar
            // 
            this.radiusTrackBar.Location = new System.Drawing.Point(186, 20);
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
            this.initialRotationLabel.Location = new System.Drawing.Point(182, 171);
            this.initialRotationLabel.Name = "initialRotationLabel";
            this.initialRotationLabel.Size = new System.Drawing.Size(70, 24);
            this.initialRotationLabel.TabIndex = 2;
            this.initialRotationLabel.Text = "起始:";
            // 
            // initialRotationUpDown
            // 
            this.initialRotationUpDown.Location = new System.Drawing.Point(266, 169);
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
            this.finalRotationLabel.Location = new System.Drawing.Point(352, 171);
            this.finalRotationLabel.Name = "finalRotationLabel";
            this.finalRotationLabel.Size = new System.Drawing.Size(70, 24);
            this.finalRotationLabel.TabIndex = 4;
            this.finalRotationLabel.Text = "终点:";
            // 
            // finalRotationUpDown
            // 
            this.finalRotationUpDown.Location = new System.Drawing.Point(428, 167);
            this.finalRotationUpDown.Maximum = new decimal(new int[] {
            360,
            0,
            0,
            0});
            this.finalRotationUpDown.Name = "finalRotationUpDown";
            this.finalRotationUpDown.Size = new System.Drawing.Size(80, 35);
            this.finalRotationUpDown.TabIndex = 5;
            // 
            // sizeIncrementLabel
            // 
            this.sizeIncrementLabel.AutoSize = true;
            this.sizeIncrementLabel.Font = new System.Drawing.Font("宋体", 11F, System.Drawing.FontStyle.Bold);
            this.sizeIncrementLabel.Location = new System.Drawing.Point(21, 89);
            this.sizeIncrementLabel.Name = "sizeIncrementLabel";
            this.sizeIncrementLabel.Size = new System.Drawing.Size(153, 30);
            this.sizeIncrementLabel.TabIndex = 6;
            this.sizeIncrementLabel.Text = "螺旋递进:";
            // 
            // sizeIncrementTrackBar
            // 
            this.sizeIncrementTrackBar.LargeChange = 1;
            this.sizeIncrementTrackBar.Location = new System.Drawing.Point(183, 84);
            this.sizeIncrementTrackBar.Maximum = 200;
            this.sizeIncrementTrackBar.Minimum = 100;
            this.sizeIncrementTrackBar.Name = "sizeIncrementTrackBar";
            this.sizeIncrementTrackBar.Size = new System.Drawing.Size(350, 90);
            this.sizeIncrementTrackBar.TabIndex = 7;
            this.sizeIncrementTrackBar.Value = 100;
            // 
            // resetButton
            // 
            this.resetButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(41)))), ((int)(((byte)(67)))), ((int)(((byte)(232)))));
            this.resetButton.FlatAppearance.BorderSize = 0;
            this.resetButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(15)))), ((int)(((byte)(42)))), ((int)(((byte)(214)))));
            this.resetButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(84)))), ((int)(((byte)(232)))));
            this.resetButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.resetButton.Font = new System.Drawing.Font("宋体", 10.875F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.resetButton.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.resetButton.Location = new System.Drawing.Point(29, 226);
            this.resetButton.Name = "resetButton";
            this.resetButton.Size = new System.Drawing.Size(128, 55);
            this.resetButton.TabIndex = 8;
            this.resetButton.Text = "重置";
            this.resetButton.UseVisualStyleBackColor = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 11F, System.Drawing.FontStyle.Bold);
            this.label1.Location = new System.Drawing.Point(24, 163);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(153, 30);
            this.label1.TabIndex = 9;
            this.label1.Text = "角度递进:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 11F, System.Drawing.FontStyle.Bold);
            this.label2.ForeColor = System.Drawing.Color.Red;
            this.label2.Location = new System.Drawing.Point(338, 127);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(44, 30);
            this.label2.TabIndex = 10;
            this.label2.Text = "▲";
            // 
            // MultipleObjectsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(192F, 192F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(567, 316);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.initialRotationLabel);
            this.Controls.Add(this.initialRotationUpDown);
            this.Controls.Add(this.finalRotationLabel);
            this.Controls.Add(this.finalRotationUpDown);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.resetButton);
            this.Controls.Add(this.radiusLabel);
            this.Controls.Add(this.sizeIncrementLabel);
            this.Controls.Add(this.sizeIncrementTrackBar);
            this.Controls.Add(this.radiusTrackBar);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MultipleObjectsForm";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "环形分布";
            this.TopMost = true;
            ((System.ComponentModel.ISupportInitialize)(this.radiusTrackBar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.initialRotationUpDown)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.finalRotationUpDown)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sizeIncrementTrackBar)).EndInit();
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
        private System.Windows.Forms.Button resetButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}
