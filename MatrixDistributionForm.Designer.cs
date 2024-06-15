// MatrixDistributionForm.Designer.cs
namespace 课件帮PPT助手
{
    partial class MatrixDistributionForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.TrackBar totalCountTrackBar;
        private System.Windows.Forms.NumericUpDown totalCountNumericUpDown;
        private System.Windows.Forms.TrackBar horizontalCountTrackBar;
        private System.Windows.Forms.NumericUpDown horizontalCountNumericUpDown;
        private System.Windows.Forms.TrackBar rowSpacingTrackBar;
        private System.Windows.Forms.NumericUpDown rowSpacingNumericUpDown;
        private System.Windows.Forms.TrackBar columnSpacingTrackBar;
        private System.Windows.Forms.NumericUpDown columnSpacingNumericUpDown;
        private System.Windows.Forms.TrackBar scaleTrackBar;
        private System.Windows.Forms.NumericUpDown scaleNumericUpDown;

        private System.Windows.Forms.Label totalCountLabel;
        private System.Windows.Forms.Label horizontalCountLabel;
        private System.Windows.Forms.Label rowSpacingLabel;
        private System.Windows.Forms.Label columnSpacingLabel;
        private System.Windows.Forms.Label scaleLabel;

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MatrixDistributionForm));
            this.totalCountTrackBar = new System.Windows.Forms.TrackBar();
            this.totalCountNumericUpDown = new System.Windows.Forms.NumericUpDown();
            this.horizontalCountTrackBar = new System.Windows.Forms.TrackBar();
            this.horizontalCountNumericUpDown = new System.Windows.Forms.NumericUpDown();
            this.rowSpacingTrackBar = new System.Windows.Forms.TrackBar();
            this.rowSpacingNumericUpDown = new System.Windows.Forms.NumericUpDown();
            this.columnSpacingTrackBar = new System.Windows.Forms.TrackBar();
            this.columnSpacingNumericUpDown = new System.Windows.Forms.NumericUpDown();
            this.scaleTrackBar = new System.Windows.Forms.TrackBar();
            this.scaleNumericUpDown = new System.Windows.Forms.NumericUpDown();
            this.totalCountLabel = new System.Windows.Forms.Label();
            this.horizontalCountLabel = new System.Windows.Forms.Label();
            this.rowSpacingLabel = new System.Windows.Forms.Label();
            this.columnSpacingLabel = new System.Windows.Forms.Label();
            this.scaleLabel = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.totalCountTrackBar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.totalCountNumericUpDown)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.horizontalCountTrackBar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.horizontalCountNumericUpDown)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rowSpacingTrackBar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rowSpacingNumericUpDown)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.columnSpacingTrackBar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.columnSpacingNumericUpDown)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.scaleTrackBar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.scaleNumericUpDown)).BeginInit();
            this.SuspendLayout();
            // 
            // totalCountTrackBar
            // 
            this.totalCountTrackBar.Location = new System.Drawing.Point(191, 60);
            this.totalCountTrackBar.Maximum = 100;
            this.totalCountTrackBar.Minimum = 1;
            this.totalCountTrackBar.Name = "totalCountTrackBar";
            this.totalCountTrackBar.Size = new System.Drawing.Size(202, 90);
            this.totalCountTrackBar.TabIndex = 1;
            this.totalCountTrackBar.Value = 1;
            this.totalCountTrackBar.Scroll += new System.EventHandler(this.totalCountTrackBar_Scroll);
            // 
            // totalCountNumericUpDown
            // 
            this.totalCountNumericUpDown.Location = new System.Drawing.Point(418, 64);
            this.totalCountNumericUpDown.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.totalCountNumericUpDown.Name = "totalCountNumericUpDown";
            this.totalCountNumericUpDown.Size = new System.Drawing.Size(120, 35);
            this.totalCountNumericUpDown.TabIndex = 2;
            this.totalCountNumericUpDown.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.totalCountNumericUpDown.ValueChanged += new System.EventHandler(this.totalCountNumericUpDown_ValueChanged);
            // 
            // horizontalCountTrackBar
            // 
            this.horizontalCountTrackBar.Location = new System.Drawing.Point(191, 135);
            this.horizontalCountTrackBar.Maximum = 100;
            this.horizontalCountTrackBar.Minimum = 1;
            this.horizontalCountTrackBar.Name = "horizontalCountTrackBar";
            this.horizontalCountTrackBar.Size = new System.Drawing.Size(202, 90);
            this.horizontalCountTrackBar.TabIndex = 4;
            this.horizontalCountTrackBar.Value = 1;
            this.horizontalCountTrackBar.Scroll += new System.EventHandler(this.horizontalCountTrackBar_Scroll);
            // 
            // horizontalCountNumericUpDown
            // 
            this.horizontalCountNumericUpDown.Location = new System.Drawing.Point(418, 141);
            this.horizontalCountNumericUpDown.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.horizontalCountNumericUpDown.Name = "horizontalCountNumericUpDown";
            this.horizontalCountNumericUpDown.Size = new System.Drawing.Size(120, 35);
            this.horizontalCountNumericUpDown.TabIndex = 5;
            this.horizontalCountNumericUpDown.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.horizontalCountNumericUpDown.ValueChanged += new System.EventHandler(this.horizontalCountNumericUpDown_ValueChanged);
            // 
            // rowSpacingTrackBar
            // 
            this.rowSpacingTrackBar.Location = new System.Drawing.Point(191, 224);
            this.rowSpacingTrackBar.Maximum = 500;
            this.rowSpacingTrackBar.Name = "rowSpacingTrackBar";
            this.rowSpacingTrackBar.Size = new System.Drawing.Size(202, 90);
            this.rowSpacingTrackBar.TabIndex = 7;
            this.rowSpacingTrackBar.Scroll += new System.EventHandler(this.rowSpacingTrackBar_Scroll);
            // 
            // rowSpacingNumericUpDown
            // 
            this.rowSpacingNumericUpDown.Location = new System.Drawing.Point(418, 218);
            this.rowSpacingNumericUpDown.Maximum = new decimal(new int[] {
            500,
            0,
            0,
            0});
            this.rowSpacingNumericUpDown.Name = "rowSpacingNumericUpDown";
            this.rowSpacingNumericUpDown.Size = new System.Drawing.Size(120, 35);
            this.rowSpacingNumericUpDown.TabIndex = 8;
            this.rowSpacingNumericUpDown.ValueChanged += new System.EventHandler(this.rowSpacingNumericUpDown_ValueChanged);
            // 
            // columnSpacingTrackBar
            // 
            this.columnSpacingTrackBar.Location = new System.Drawing.Point(191, 295);
            this.columnSpacingTrackBar.Maximum = 500;
            this.columnSpacingTrackBar.Name = "columnSpacingTrackBar";
            this.columnSpacingTrackBar.Size = new System.Drawing.Size(202, 90);
            this.columnSpacingTrackBar.TabIndex = 10;
            this.columnSpacingTrackBar.Scroll += new System.EventHandler(this.columnSpacingTrackBar_Scroll);
            // 
            // columnSpacingNumericUpDown
            // 
            this.columnSpacingNumericUpDown.Location = new System.Drawing.Point(418, 295);
            this.columnSpacingNumericUpDown.Maximum = new decimal(new int[] {
            500,
            0,
            0,
            0});
            this.columnSpacingNumericUpDown.Name = "columnSpacingNumericUpDown";
            this.columnSpacingNumericUpDown.Size = new System.Drawing.Size(120, 35);
            this.columnSpacingNumericUpDown.TabIndex = 11;
            this.columnSpacingNumericUpDown.ValueChanged += new System.EventHandler(this.columnSpacingNumericUpDown_ValueChanged);
            // 
            // scaleTrackBar
            // 
            this.scaleTrackBar.Location = new System.Drawing.Point(191, 366);
            this.scaleTrackBar.Maximum = 200;
            this.scaleTrackBar.Name = "scaleTrackBar";
            this.scaleTrackBar.Size = new System.Drawing.Size(202, 90);
            this.scaleTrackBar.TabIndex = 13;
            this.scaleTrackBar.Value = 100;
            this.scaleTrackBar.Scroll += new System.EventHandler(this.scaleTrackBar_Scroll);
            // 
            // scaleNumericUpDown
            // 
            this.scaleNumericUpDown.Location = new System.Drawing.Point(418, 372);
            this.scaleNumericUpDown.Maximum = new decimal(new int[] {
            200,
            0,
            0,
            0});
            this.scaleNumericUpDown.Name = "scaleNumericUpDown";
            this.scaleNumericUpDown.Size = new System.Drawing.Size(120, 35);
            this.scaleNumericUpDown.TabIndex = 14;
            this.scaleNumericUpDown.Value = new decimal(new int[] {
            100,
            0,
            0,
            0});
            this.scaleNumericUpDown.ValueChanged += new System.EventHandler(this.scaleNumericUpDown_ValueChanged);
            // 
            // totalCountLabel
            // 
            this.totalCountLabel.Font = new System.Drawing.Font("宋体", 10.125F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.totalCountLabel.Location = new System.Drawing.Point(32, 68);
            this.totalCountLabel.Name = "totalCountLabel";
            this.totalCountLabel.Size = new System.Drawing.Size(154, 37);
            this.totalCountLabel.TabIndex = 0;
            this.totalCountLabel.Text = "对象数量：";
            // 
            // horizontalCountLabel
            // 
            this.horizontalCountLabel.Font = new System.Drawing.Font("宋体", 10.125F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.horizontalCountLabel.Location = new System.Drawing.Point(32, 147);
            this.horizontalCountLabel.Name = "horizontalCountLabel";
            this.horizontalCountLabel.Size = new System.Drawing.Size(154, 37);
            this.horizontalCountLabel.TabIndex = 3;
            this.horizontalCountLabel.Text = "横向数量：";
            // 
            // rowSpacingLabel
            // 
            this.rowSpacingLabel.Font = new System.Drawing.Font("宋体", 10.125F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rowSpacingLabel.Location = new System.Drawing.Point(32, 224);
            this.rowSpacingLabel.Name = "rowSpacingLabel";
            this.rowSpacingLabel.Size = new System.Drawing.Size(154, 37);
            this.rowSpacingLabel.TabIndex = 6;
            this.rowSpacingLabel.Text = "水平间距：";
            // 
            // columnSpacingLabel
            // 
            this.columnSpacingLabel.Font = new System.Drawing.Font("宋体", 10.125F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.columnSpacingLabel.Location = new System.Drawing.Point(32, 306);
            this.columnSpacingLabel.Name = "columnSpacingLabel";
            this.columnSpacingLabel.Size = new System.Drawing.Size(154, 37);
            this.columnSpacingLabel.TabIndex = 9;
            this.columnSpacingLabel.Text = "垂直间距：";
            // 
            // scaleLabel
            // 
            this.scaleLabel.Font = new System.Drawing.Font("宋体", 10.125F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.scaleLabel.Location = new System.Drawing.Point(32, 375);
            this.scaleLabel.Name = "scaleLabel";
            this.scaleLabel.Size = new System.Drawing.Size(154, 37);
            this.scaleLabel.TabIndex = 12;
            this.scaleLabel.Text = "尺寸缩放：";
            // 
            // MatrixDistributionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(192F, 192F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(584, 474);
            this.Controls.Add(this.scaleTrackBar);
            this.Controls.Add(this.totalCountNumericUpDown);
            this.Controls.Add(this.horizontalCountNumericUpDown);
            this.Controls.Add(this.rowSpacingNumericUpDown);
            this.Controls.Add(this.columnSpacingNumericUpDown);
            this.Controls.Add(this.scaleLabel);
            this.Controls.Add(this.scaleNumericUpDown);
            this.Controls.Add(this.columnSpacingTrackBar);
            this.Controls.Add(this.rowSpacingTrackBar);
            this.Controls.Add(this.columnSpacingLabel);
            this.Controls.Add(this.horizontalCountTrackBar);
            this.Controls.Add(this.rowSpacingLabel);
            this.Controls.Add(this.totalCountTrackBar);
            this.Controls.Add(this.horizontalCountLabel);
            this.Controls.Add(this.totalCountLabel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MatrixDistributionForm";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "矩阵分布";
            ((System.ComponentModel.ISupportInitialize)(this.totalCountTrackBar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.totalCountNumericUpDown)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.horizontalCountTrackBar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.horizontalCountNumericUpDown)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.rowSpacingTrackBar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.rowSpacingNumericUpDown)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.columnSpacingTrackBar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.columnSpacingNumericUpDown)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.scaleTrackBar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.scaleNumericUpDown)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        public void SetTotalCount(int count)
        {
            this.totalCountTrackBar.Value = count;
            this.totalCountNumericUpDown.Value = count;
            this.totalCountTrackBar.Enabled = false;
            this.totalCountNumericUpDown.Enabled = false;
        }

        public void EnableTotalCountAdjustment()
        {
            this.totalCountTrackBar.Enabled = true;
            this.totalCountNumericUpDown.Enabled = true;
        }
    }
}
