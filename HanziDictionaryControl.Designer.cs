namespace 课件帮PPT助手
{
    partial class HanziDictionaryControl
    {
        private System.ComponentModel.IContainer components = null;

        private System.Windows.Forms.TextBox searchTextBox;
        private System.Windows.Forms.Button searchButton;
        private System.Windows.Forms.Label hanziLabel;
        private System.Windows.Forms.Label pinyinLabel;
        private System.Windows.Forms.Label radicalLabel;
        private System.Windows.Forms.Label strokesLabel;
        private System.Windows.Forms.Label structureLabel;
        private System.Windows.Forms.Label relatedWordsLabel;
        private System.Windows.Forms.FlowLayoutPanel wordsPanel;

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
            this.searchTextBox = new System.Windows.Forms.TextBox();
            this.searchButton = new System.Windows.Forms.Button();
            this.hanziLabel = new System.Windows.Forms.Label();
            this.pinyinLabel = new System.Windows.Forms.Label();
            this.radicalLabel = new System.Windows.Forms.Label();
            this.strokesLabel = new System.Windows.Forms.Label();
            this.structureLabel = new System.Windows.Forms.Label();
            this.relatedWordsLabel = new System.Windows.Forms.Label();
            this.wordsPanel = new System.Windows.Forms.FlowLayoutPanel();

            this.SuspendLayout();

            // 
            // searchTextBox
            // 
            this.searchTextBox.Location = new System.Drawing.Point(30, 30);
            this.searchTextBox.Name = "searchTextBox";
            this.searchTextBox.Size = new System.Drawing.Size(420, 31);
            this.searchTextBox.TabIndex = 0;

            // 
            // searchButton
            // 
            this.searchButton.Location = new System.Drawing.Point(470, 30);
            this.searchButton.Name = "searchButton";
            this.searchButton.Size = new System.Drawing.Size(75, 31);
            this.searchButton.TabIndex = 1;
            this.searchButton.Text = "🔍";
            this.searchButton.UseVisualStyleBackColor = true;

            // 
            // hanziLabel
            // 
            this.hanziLabel.Font = new System.Drawing.Font("SimSun", 48F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.hanziLabel.Location = new System.Drawing.Point(30, 80);
            this.hanziLabel.Name = "hanziLabel";
            this.hanziLabel.Size = new System.Drawing.Size(200, 200);
            this.hanziLabel.TabIndex = 2;
            this.hanziLabel.Text = "冯";
            this.hanziLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;

            // 
            // pinyinLabel
            // 
            this.pinyinLabel.Location = new System.Drawing.Point(250, 80);
            this.pinyinLabel.Name = "pinyinLabel";
            this.pinyinLabel.Size = new System.Drawing.Size(200, 40);
            this.pinyinLabel.TabIndex = 3;
            this.pinyinLabel.Text = "拼音: féng";

            // 
            // radicalLabel
            // 
            this.radicalLabel.Location = new System.Drawing.Point(250, 130);
            this.radicalLabel.Name = "radicalLabel";
            this.radicalLabel.Size = new System.Drawing.Size(200, 40);
            this.radicalLabel.TabIndex = 4;
            this.radicalLabel.Text = "部首: 冫";

            // 
            // strokesLabel
            // 
            this.strokesLabel.Location = new System.Drawing.Point(250, 180);
            this.strokesLabel.Name = "strokesLabel";
            this.strokesLabel.Size = new System.Drawing.Size(200, 40);
            this.strokesLabel.TabIndex = 5;
            this.strokesLabel.Text = "笔画: 5";

            // 
            // structureLabel
            // 
            this.structureLabel.Location = new System.Drawing.Point(250, 230);
            this.structureLabel.Name = "structureLabel";
            this.structureLabel.Size = new System.Drawing.Size(200, 40);
            this.structureLabel.TabIndex = 6;
            this.structureLabel.Text = "结构: 左右";

            // 
            // relatedWordsLabel
            // 
            this.relatedWordsLabel.Location = new System.Drawing.Point(30, 300);
            this.relatedWordsLabel.Name = "relatedWordsLabel";
            this.relatedWordsLabel.Size = new System.Drawing.Size(200, 40);
            this.relatedWordsLabel.TabIndex = 7;
            this.relatedWordsLabel.Text = "相关词组:";

            // 
            // wordsPanel
            // 
            this.wordsPanel.Location = new System.Drawing.Point(30, 340);
            this.wordsPanel.Name = "wordsPanel";
            this.wordsPanel.Size = new System.Drawing.Size(515, 340);
            this.wordsPanel.TabIndex = 8;
            this.wordsPanel.WrapContents = true;

            // 
            // HanziDictionaryControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(244)))), ((int)(((byte)(244)))), ((int)(((byte)(254)))));
            this.Controls.Add(this.searchTextBox);
            this.Controls.Add(this.searchButton);
            this.Controls.Add(this.hanziLabel);
            this.Controls.Add(this.pinyinLabel);
            this.Controls.Add(this.radicalLabel);
            this.Controls.Add(this.strokesLabel);
            this.Controls.Add(this.structureLabel);
            this.Controls.Add(this.relatedWordsLabel);
            this.Controls.Add(this.wordsPanel);
            this.Name = "HanziDictionaryControl";
            this.Size = new System.Drawing.Size(579, 705);
            this.Load += new System.EventHandler(this.HanziDictionaryControl_Load);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
