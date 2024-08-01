using System.Windows.Forms;

namespace 课件帮PPT助手
{
    partial class HanziDictionaryForm
    {
        private System.ComponentModel.IContainer components = null;
        private TextBox searchTextBox;
        private Button searchButton;
        private Label hanziLabel;
        private Label pinyinLabel;
        private Label radicalLabel;
        private Label strokesLabel;
        private Label structureLabel;
        private Label relatedWordsLabel;
        private PictureBox pictureBox1;
        private FlowLayoutPanel wordsPanel;

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(HanziDictionaryForm));
            this.searchTextBox = new System.Windows.Forms.TextBox();
            this.searchButton = new System.Windows.Forms.Button();
            this.hanziLabel = new System.Windows.Forms.Label();
            this.pinyinLabel = new System.Windows.Forms.Label();
            this.radicalLabel = new System.Windows.Forms.Label();
            this.strokesLabel = new System.Windows.Forms.Label();
            this.structureLabel = new System.Windows.Forms.Label();
            this.relatedWordsLabel = new System.Windows.Forms.Label();
            this.wordsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.导出 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // searchTextBox
            // 
            this.searchTextBox.Font = new System.Drawing.Font("Arial", 14F);
            this.searchTextBox.Location = new System.Drawing.Point(25, 20);
            this.searchTextBox.Multiline = true;
            this.searchTextBox.Name = "searchTextBox";
            this.searchTextBox.Size = new System.Drawing.Size(424, 50);
            this.searchTextBox.TabIndex = 0;
            this.searchTextBox.WordWrap = false;
            // 
            // searchButton
            // 
            this.searchButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(230)))), ((int)(((byte)(236)))), ((int)(((byte)(255)))));
            this.searchButton.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("searchButton.BackgroundImage")));
            this.searchButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.searchButton.Location = new System.Drawing.Point(455, 20);
            this.searchButton.Name = "searchButton";
            this.searchButton.Size = new System.Drawing.Size(50, 50);
            this.searchButton.TabIndex = 1;
            this.searchButton.UseVisualStyleBackColor = false;
            this.searchButton.Click += new System.EventHandler(this.SearchButton_Click);
            // 
            // hanziLabel
            // 
            this.hanziLabel.BackColor = System.Drawing.Color.Transparent;
            this.hanziLabel.Font = new System.Drawing.Font("宋体", 42F);
            this.hanziLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(58)))), ((int)(((byte)(102)))), ((int)(((byte)(252)))));
            this.hanziLabel.Location = new System.Drawing.Point(35, 108);
            this.hanziLabel.Name = "hanziLabel";
            this.hanziLabel.Size = new System.Drawing.Size(127, 128);
            this.hanziLabel.TabIndex = 2;
            this.hanziLabel.Text = " ";
            this.hanziLabel.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // pinyinLabel
            // 
            this.pinyinLabel.Font = new System.Drawing.Font("宋体", 11F);
            this.pinyinLabel.Location = new System.Drawing.Point(203, 105);
            this.pinyinLabel.Name = "pinyinLabel";
            this.pinyinLabel.Size = new System.Drawing.Size(300, 30);
            this.pinyinLabel.TabIndex = 3;
            // 
            // radicalLabel
            // 
            this.radicalLabel.Font = new System.Drawing.Font("宋体", 11F);
            this.radicalLabel.Location = new System.Drawing.Point(203, 145);
            this.radicalLabel.Name = "radicalLabel";
            this.radicalLabel.Size = new System.Drawing.Size(300, 30);
            this.radicalLabel.TabIndex = 4;
            // 
            // strokesLabel
            // 
            this.strokesLabel.Font = new System.Drawing.Font("宋体", 11F);
            this.strokesLabel.Location = new System.Drawing.Point(203, 185);
            this.strokesLabel.Name = "strokesLabel";
            this.strokesLabel.Size = new System.Drawing.Size(300, 30);
            this.strokesLabel.TabIndex = 5;
            // 
            // structureLabel
            // 
            this.structureLabel.Font = new System.Drawing.Font("宋体", 11F);
            this.structureLabel.Location = new System.Drawing.Point(203, 225);
            this.structureLabel.Name = "structureLabel";
            this.structureLabel.Size = new System.Drawing.Size(300, 30);
            this.structureLabel.TabIndex = 6;
            // 
            // relatedWordsLabel
            // 
            this.relatedWordsLabel.Font = new System.Drawing.Font("Arial", 11F, System.Drawing.FontStyle.Bold);
            this.relatedWordsLabel.Location = new System.Drawing.Point(20, 283);
            this.relatedWordsLabel.Name = "relatedWordsLabel";
            this.relatedWordsLabel.Size = new System.Drawing.Size(300, 35);
            this.relatedWordsLabel.TabIndex = 7;
            this.relatedWordsLabel.Text = "相关组词：";
            // 
            // wordsPanel
            // 
            this.wordsPanel.AutoSize = true;
            this.wordsPanel.Font = new System.Drawing.Font("楷体_GB2312", 11F);
            this.wordsPanel.Location = new System.Drawing.Point(25, 323);
            this.wordsPanel.Margin = new System.Windows.Forms.Padding(5);
            this.wordsPanel.Name = "wordsPanel";
            this.wordsPanel.Size = new System.Drawing.Size(485, 256);
            this.wordsPanel.TabIndex = 8;
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pictureBox1.ErrorImage = null;
            this.pictureBox1.InitialImage = null;
            this.pictureBox1.Location = new System.Drawing.Point(25, 100);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(147, 147);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 9;
            this.pictureBox1.TabStop = false;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 23);
            this.label1.TabIndex = 0;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 167);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(0, 24);
            this.label2.TabIndex = 10;
            // 
            // label3
            // 
            this.label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label3.Location = new System.Drawing.Point(25, 100);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(147, 147);
            this.label3.TabIndex = 11;
            // 
            // 导出
            // 
            this.导出.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(74)))), ((int)(((byte)(247)))));
            this.导出.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(14)))), ((int)(((byte)(62)))), ((int)(((byte)(228)))));
            this.导出.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(65)))), ((int)(((byte)(107)))), ((int)(((byte)(254)))));
            this.导出.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.导出.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.导出.ForeColor = System.Drawing.Color.White;
            this.导出.Location = new System.Drawing.Point(25, 603);
            this.导出.Name = "导出";
            this.导出.Size = new System.Drawing.Size(180, 56);
            this.导出.TabIndex = 12;
            this.导出.Text = "导出至幻灯片";
            this.导出.UseVisualStyleBackColor = false;
            this.导出.Click += new System.EventHandler(this.导出_Click);
            // 
            // HanziDictionaryForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(192F, 192F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(245)))), ((int)(((byte)(247)))), ((int)(((byte)(255)))));
            this.ClientSize = new System.Drawing.Size(526, 719);
            this.Controls.Add(this.导出);
            this.Controls.Add(this.hanziLabel);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.searchTextBox);
            this.Controls.Add(this.searchButton);
            this.Controls.Add(this.pinyinLabel);
            this.Controls.Add(this.radicalLabel);
            this.Controls.Add(this.strokesLabel);
            this.Controls.Add(this.structureLabel);
            this.Controls.Add(this.relatedWordsLabel);
            this.Controls.Add(this.wordsPanel);
            this.Controls.Add(this.pictureBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "HanziDictionaryForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "简易字典";
            this.Load += new System.EventHandler(this.HanziDictionaryForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private Label label1;
        private Label label2;
        private Label label3;
        private Button 导出;
    }
}
