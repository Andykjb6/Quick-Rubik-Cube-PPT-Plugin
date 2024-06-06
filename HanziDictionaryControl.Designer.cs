using System;
using System.Drawing;
using System.Windows.Forms;

namespace 课件帮PPT助手
{
    public class HanziLabel : Control
    {
        public string Hanzi { get; set; }
        public Image TianziImage { get; set; }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            if (TianziImage != null)
            {
                e.Graphics.DrawImage(TianziImage, 0, 0, this.Width, this.Height);
            }

            if (!string.IsNullOrEmpty(Hanzi))
            {
                using (Font font = new Font("Arial", 36, FontStyle.Regular))
                {
                    SizeF textSize = e.Graphics.MeasureString(Hanzi, font);
                    PointF locationToDraw = new PointF();
                    locationToDraw.X = (this.Width / 2) - (textSize.Width / 2);
                    locationToDraw.Y = (this.Height / 2) - (textSize.Height / 2);

                    e.Graphics.DrawString(Hanzi, font, Brushes.Black, locationToDraw);
                }
            }
        }
    }

    partial class HanziDictionaryControl
    {
        private System.ComponentModel.IContainer components = null;
        private TextBox searchTextBox;
        private Button searchButton;
        private HanziLabel hanziLabel;
        private Label pinyinLabel;
        private Label radicalLabel;
        private Label strokesLabel;
        private Label structureLabel;
        private Label relatedWordsLabel;
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(HanziDictionaryControl));
            this.searchTextBox = new System.Windows.Forms.TextBox();
            this.searchButton = new System.Windows.Forms.Button();
            this.hanziLabel = new HanziLabel();
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
            this.searchTextBox.Font = new System.Drawing.Font("Arial", 14F);
            this.searchTextBox.Location = new System.Drawing.Point(25, 20);
            this.searchTextBox.Multiline = true;
            this.searchTextBox.Name = "searchTextBox";
            this.searchTextBox.Size = new System.Drawing.Size(424, 50);
            this.searchTextBox.TabIndex = 0;
            // 
            // searchButton
            // 
            this.searchButton.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("searchButton.BackgroundImage")));
            this.searchButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.searchButton.Location = new System.Drawing.Point(455, 20);
            this.searchButton.Name = "searchButton";
            this.searchButton.Size = new System.Drawing.Size(50, 50);
            this.searchButton.TabIndex = 1;
            this.searchButton.Click += new System.EventHandler(this.SearchButton_Click);
            // 
            // hanziLabel
            // 
            this.hanziLabel.Location = new System.Drawing.Point(25, 105);
            this.hanziLabel.Name = "hanziLabel";
            this.hanziLabel.Size = new System.Drawing.Size(147, 147);
            this.hanziLabel.TabIndex = 2;
            this.hanziLabel.TianziImage = Image.FromFile(@"C:\path\to\tianzi.png"); // 设置田字格图片路径
            // 
            // pinyinLabel
            // 
            this.pinyinLabel.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.pinyinLabel.Location = new System.Drawing.Point(203, 105);
            this.pinyinLabel.Name = "pinyinLabel";
            this.pinyinLabel.Size = new System.Drawing.Size(220, 30);
            this.pinyinLabel.TabIndex = 3;
            // 
            // radicalLabel
            // 
            this.radicalLabel.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.radicalLabel.Location = new System.Drawing.Point(203, 145);
            this.radicalLabel.Name = "radicalLabel";
            this.radicalLabel.Size = new System.Drawing.Size(220, 30);
            this.radicalLabel.TabIndex = 4;
            // 
            // strokesLabel
            // 
            this.strokesLabel.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.strokesLabel.Location = new System.Drawing.Point(203, 185);
            this.strokesLabel.Name = "strokesLabel";
            this.strokesLabel.Size = new System.Drawing.Size(220, 30);
            this.strokesLabel.TabIndex = 5;
            // 
            // structureLabel
            // 
            this.structureLabel.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.structureLabel.Location = new System.Drawing.Point(203, 225);
            this.structureLabel.Name = "structureLabel";
            this.structureLabel.Size = new System.Drawing.Size(220, 30);
            this.structureLabel.TabIndex = 6;
            // 
            // relatedWordsLabel
            // 
            this.relatedWordsLabel.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.relatedWordsLabel.Location = new System.Drawing.Point(20, 283);
            this.relatedWordsLabel.Name = "relatedWordsLabel";
            this.relatedWordsLabel.Size = new System.Drawing.Size(300, 35);
            this.relatedWordsLabel.TabIndex = 7;
            this.relatedWordsLabel.Text = "相关组词：";
            // 
            // wordsPanel
            // 
            this.wordsPanel.AutoSize = true;
            this.wordsPanel.Font = new System.Drawing.Font("宋体", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.wordsPanel.Location = new System.Drawing.Point(25, 323);
            this.wordsPanel.Margin = new System.Windows.Forms.Padding(5);
            this.wordsPanel.Name = "wordsPanel";
            this.wordsPanel.Size = new System.Drawing.Size(485, 300);
            this.wordsPanel.TabIndex = 8;
            // 
            // HanziDictionaryControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(244)))), ((int)(((byte)(244)))), ((int)(((byte)(254)))));
            this.Controls.Add(this.hanziLabel); // 确保自定义控件在其他控件之前
            this.Controls.Add(this.searchTextBox);
            this.Controls.Add(this.searchButton);
            this.Controls.Add(this.pinyinLabel);
            this.Controls.Add(this.radicalLabel);
            this.Controls.Add(this.strokesLabel);
            this.Controls.Add(this.structureLabel);
            this.Controls.Add(this.relatedWordsLabel);
            this.Controls.Add(this.wordsPanel);
            this.Name = "HanziDictionaryControl";
            this.Size = new System.Drawing.Size(526, 691);
            this.Load += new System.EventHandler(this.HanziDictionaryControl_Load);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
