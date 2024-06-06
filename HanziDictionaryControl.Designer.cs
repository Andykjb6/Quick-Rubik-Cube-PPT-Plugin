using System.Drawing;
using System.Windows.Forms;

namespace 课件帮PPT助手
{
    partial class HanziDictionaryControl
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
        private TableLayoutPanel wordsPanel;

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
            this.searchTextBox = new TextBox();
            this.searchButton = new Button();
            this.hanziLabel = new Label();
            this.pinyinLabel = new Label();
            this.radicalLabel = new Label();
            this.strokesLabel = new Label();
            this.structureLabel = new Label();
            this.relatedWordsLabel = new Label();
            this.wordsPanel = new TableLayoutPanel();

            this.SuspendLayout();

            // searchTextBox
            this.searchTextBox.Location = new System.Drawing.Point(20, 20);
            this.searchTextBox.Size = new System.Drawing.Size(300, 30);

            // searchButton
            this.searchButton.Location = new System.Drawing.Point(340, 20);
            this.searchButton.Size = new System.Drawing.Size(30, 30);
            this.searchButton.Text = "";
            this.searchButton.BackgroundImage = new Bitmap("search_icon.png"); // 确保图标存在
            this.searchButton.BackgroundImageLayout = ImageLayout.Stretch;
            this.searchButton.Click += new System.EventHandler(this.SearchButton_Click);

            // hanziLabel
            this.hanziLabel.Location = new System.Drawing.Point(20, 70);
            this.hanziLabel.Size = new System.Drawing.Size(100, 100);
            this.hanziLabel.Font = new Font("Arial", 48);
            this.hanziLabel.TextAlign = ContentAlignment.MiddleCenter;
            this.hanziLabel.BorderStyle = BorderStyle.FixedSingle;

            // pinyinLabel
            this.pinyinLabel.Location = new System.Drawing.Point(140, 70);
            this.pinyinLabel.Size = new System.Drawing.Size(200, 30);
            this.pinyinLabel.Font = new Font("Arial", 16);

            // radicalLabel
            this.radicalLabel.Location = new System.Drawing.Point(140, 110);
            this.radicalLabel.Size = new System.Drawing.Size(200, 30);
            this.radicalLabel.Font = new Font("Arial", 16);

            // strokesLabel
            this.strokesLabel.Location = new System.Drawing.Point(140, 150);
            this.strokesLabel.Size = new System.Drawing.Size(200, 30);
            this.strokesLabel.Font = new Font("Arial", 16);

            // structureLabel
            this.structureLabel.Location = new System.Drawing.Point(140, 190);
            this.structureLabel.Size = new System.Drawing.Size(200, 30);
            this.structureLabel.Font = new Font("Arial", 16);

            // relatedWordsLabel
            this.relatedWordsLabel.Location = new System.Drawing.Point(20, 230);
            this.relatedWordsLabel.Size = new System.Drawing.Size(100, 30);
            this.relatedWordsLabel.Font = new Font("Arial", 18, FontStyle.Bold);
            this.relatedWordsLabel.Text = "相关组词：";

            // wordsPanel
            this.wordsPanel.Location = new System.Drawing.Point(20, 270);
            this.wordsPanel.Size = new System.Drawing.Size(350, 300);
            this.wordsPanel.ColumnCount = 4;
            this.wordsPanel.RowCount = 4;
            this.wordsPanel.AutoSize = true;
            this.wordsPanel.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single;

            // HanziDictionaryControl
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(244, 244, 254);
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
            this.Size = new System.Drawing.Size(400, 600);
            this.Load += new System.EventHandler(this.HanziDictionaryControl_Load);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
