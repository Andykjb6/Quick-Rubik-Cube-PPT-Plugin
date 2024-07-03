using System.Windows.Forms;

namespace 课件帮PPT助手
{
    partial class AnimationForm : Form
    {
        private TabControl tabControl;
        private TabPage tabPage1;
        private TabPage tabPage2;
        private Label label1;
        private Label inputLabel;
        private TextBox textBox;
        private Label label2;
        private Button selectAllButton;
        private Button animateButton;
        private Button adjustAnimationButton;
        private Panel adjustPanel;
        private ListBox listBox;
        private Label durationLabel;
        private Button upButton;
        private Button downButton;
        private Button leftButton;
        private Button rightButton;
        private NumericUpDown multiDurationControl;

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AnimationForm));
            this.tabControl = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.textBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.inputLabel = new System.Windows.Forms.Label();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.label2 = new System.Windows.Forms.Label();
            this.selectAllButton = new System.Windows.Forms.Button();
            this.animateButton = new System.Windows.Forms.Button();
            this.adjustAnimationButton = new System.Windows.Forms.Button();
            this.adjustPanel = new System.Windows.Forms.Panel();
            this.multiDurationControl = new System.Windows.Forms.NumericUpDown();
            this.listBox = new System.Windows.Forms.ListBox();
            this.upButton = new System.Windows.Forms.Button();
            this.downButton = new System.Windows.Forms.Button();
            this.leftButton = new System.Windows.Forms.Button();
            this.rightButton = new System.Windows.Forms.Button();
            this.durationLabel = new System.Windows.Forms.Label();
            this.tabControl.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.adjustPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.multiDurationControl)).BeginInit();
            this.SuspendLayout();
            this.multiDurationControl = new NumericUpDown();
            ((System.ComponentModel.ISupportInitialize)(this.multiDurationControl)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl
            // 
            this.tabControl.Controls.Add(this.tabPage1);
            this.tabControl.Controls.Add(this.tabPage2);
            this.tabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl.Location = new System.Drawing.Point(0, 0);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(503, 463);
            this.tabControl.TabIndex = 0;
            this.tabControl.Tag = "";
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.textBox);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.inputLabel);
            this.tabPage1.Location = new System.Drawing.Point(8, 39);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Size = new System.Drawing.Size(487, 416);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "第一步";
            // 
            // textBox
            // 
            this.textBox.Font = new System.Drawing.Font("宋体", 16.125F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox.Location = new System.Drawing.Point(22, 141);
            this.textBox.Name = "textBox";
            this.textBox.Size = new System.Drawing.Size(440, 57);
            this.textBox.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(19, 36);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(470, 31);
            this.label1.TabIndex = 0;
            this.label1.Text = "提示：请按照笔画顺序依次选中所有笔画。";
            // 
            // inputLabel
            // 
            this.inputLabel.AutoSize = true;
            this.inputLabel.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.inputLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(19)))), ((int)(((byte)(84)))), ((int)(((byte)(247)))));
            this.inputLabel.Location = new System.Drawing.Point(19, 90);
            this.inputLabel.Name = "inputLabel";
            this.inputLabel.Size = new System.Drawing.Size(269, 37);
            this.inputLabel.TabIndex = 1;
            this.inputLabel.Text = "①请输入对应汉字：";
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.label2);
            this.tabPage2.Controls.Add(this.selectAllButton);
            this.tabPage2.Controls.Add(this.animateButton);
            this.tabPage2.Controls.Add(this.adjustAnimationButton);
            this.tabPage2.Controls.Add(this.adjustPanel);
            this.tabPage2.Location = new System.Drawing.Point(8, 39);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Size = new System.Drawing.Size(487, 416);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "第二步";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(10, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(422, 31);
            this.label2.TabIndex = 0;
            this.label2.Text = "提示：“智能全选”→“智能动画”。";
            // 
            // selectAllButton
            // 
            this.selectAllButton.BackColor = System.Drawing.Color.White;
            this.selectAllButton.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("selectAllButton.BackgroundImage")));
            this.selectAllButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.selectAllButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(48)))), ((int)(((byte)(48)))), ((int)(((byte)(237)))));
            this.selectAllButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(223)))), ((int)(((byte)(223)))), ((int)(((byte)(249)))));
            this.selectAllButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White;
            this.selectAllButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.selectAllButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.selectAllButton.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(48)))), ((int)(((byte)(48)))), ((int)(((byte)(237)))));
            this.selectAllButton.Location = new System.Drawing.Point(10, 65);
            this.selectAllButton.Name = "selectAllButton";
            this.selectAllButton.Size = new System.Drawing.Size(220, 45);
            this.selectAllButton.TabIndex = 1;
            this.selectAllButton.UseVisualStyleBackColor = false;
            // 
            // animateButton
            // 
            this.animateButton.BackColor = System.Drawing.Color.White;
            this.animateButton.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("animateButton.BackgroundImage")));
            this.animateButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.animateButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(48)))), ((int)(((byte)(48)))), ((int)(((byte)(237)))));
            this.animateButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(223)))), ((int)(((byte)(223)))), ((int)(((byte)(249)))));
            this.animateButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White;
            this.animateButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.animateButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.animateButton.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(48)))), ((int)(((byte)(48)))), ((int)(((byte)(237)))));
            this.animateButton.Location = new System.Drawing.Point(250, 65);
            this.animateButton.Name = "animateButton";
            this.animateButton.Size = new System.Drawing.Size(220, 45);
            this.animateButton.TabIndex = 2;
            this.animateButton.UseVisualStyleBackColor = false;
            // 
            // adjustAnimationButton
            // 
            this.adjustAnimationButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(15)))), ((int)(((byte)(81)))), ((int)(((byte)(246)))));
            this.adjustAnimationButton.FlatAppearance.BorderSize = 0;
            this.adjustAnimationButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(32)))), ((int)(((byte)(32)))), ((int)(((byte)(209)))));
            this.adjustAnimationButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(55)))), ((int)(((byte)(55)))), ((int)(((byte)(235)))));
            this.adjustAnimationButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.adjustAnimationButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.adjustAnimationButton.ForeColor = System.Drawing.Color.White;
            this.adjustAnimationButton.Location = new System.Drawing.Point(10, 120);
            this.adjustAnimationButton.Name = "adjustAnimationButton";
            this.adjustAnimationButton.Size = new System.Drawing.Size(458, 45);
            this.adjustAnimationButton.TabIndex = 3;
            this.adjustAnimationButton.Text = "▾动画调整";
            this.adjustAnimationButton.UseVisualStyleBackColor = false;
            // 
            // adjustPanel
            // 
            this.adjustPanel.Controls.Add(this.multiDurationControl);
            this.adjustPanel.Controls.Add(this.listBox);
            this.adjustPanel.Controls.Add(this.upButton);
            this.adjustPanel.Controls.Add(this.downButton);
            this.adjustPanel.Controls.Add(this.leftButton);
            this.adjustPanel.Controls.Add(this.rightButton);
            this.adjustPanel.Controls.Add(this.durationLabel);
            this.adjustPanel.Location = new System.Drawing.Point(10, 180);
            this.adjustPanel.Name = "adjustPanel";
            this.adjustPanel.Size = new System.Drawing.Size(460, 510);
            this.adjustPanel.TabIndex = 4;
            this.adjustPanel.Visible = false;
            // 
            // multiDurationControl
            // 
            this.multiDurationControl.DecimalPlaces = 2;
            this.multiDurationControl.Increment = new decimal(new int[] { 1, 0, 0, 65536 });
            this.multiDurationControl.Location = new System.Drawing.Point(270, 180);
            this.multiDurationControl.Maximum = new decimal(new int[] { 10, 0, 0, 0 });
            this.multiDurationControl.Minimum = new decimal(new int[] { 1, 0, 0, 65536 });
            this.multiDurationControl.Name = "multiDurationControl";
            this.multiDurationControl.Size = new System.Drawing.Size(120, 35);
            this.multiDurationControl.TabIndex = 6;
            this.multiDurationControl.Value = new decimal(new int[] { 5, 0, 0, 65536 });
            this.multiDurationControl.Visible = false;

            ((System.ComponentModel.ISupportInitialize)(this.multiDurationControl)).EndInit();
            this.ResumeLayout(false);
        
            // 
            // listBox
            // 
            this.listBox.ItemHeight = 24;
            this.listBox.Location = new System.Drawing.Point(10, 10);
            this.listBox.Name = "listBox";
            this.listBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.listBox.Size = new System.Drawing.Size(200, 196);
            this.listBox.TabIndex = 0;
            // 
            // upButton
            // 
            this.upButton.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.upButton.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("upButton.BackgroundImage")));
            this.upButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.upButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(39)))), ((int)(((byte)(39)))), ((int)(((byte)(226)))));
            this.upButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(221)))), ((int)(((byte)(221)))), ((int)(((byte)(249)))));
            this.upButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(252)))));
            this.upButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.upButton.Location = new System.Drawing.Point(320, 10);
            this.upButton.Name = "upButton";
            this.upButton.Size = new System.Drawing.Size(50, 50);
            this.upButton.TabIndex = 1;
            this.upButton.UseVisualStyleBackColor = false;
            // 
            // downButton
            // 
            this.downButton.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.downButton.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("downButton.BackgroundImage")));
            this.downButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.downButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(39)))), ((int)(((byte)(39)))), ((int)(((byte)(226)))));
            this.downButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(221)))), ((int)(((byte)(221)))), ((int)(((byte)(249)))));
            this.downButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(252)))));
            this.downButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.downButton.Location = new System.Drawing.Point(320, 71);
            this.downButton.Name = "downButton";
            this.downButton.Size = new System.Drawing.Size(50, 50);
            this.downButton.TabIndex = 2;
            this.downButton.UseVisualStyleBackColor = false;
            // 
            // leftButton
            // 
            this.leftButton.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.leftButton.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("leftButton.BackgroundImage")));
            this.leftButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.leftButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(39)))), ((int)(((byte)(39)))), ((int)(((byte)(226)))));
            this.leftButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(221)))), ((int)(((byte)(221)))), ((int)(((byte)(249)))));
            this.leftButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(252)))));
            this.leftButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.leftButton.Location = new System.Drawing.Point(263, 40);
            this.leftButton.Name = "leftButton";
            this.leftButton.Size = new System.Drawing.Size(50, 50);
            this.leftButton.TabIndex = 3;
            this.leftButton.UseVisualStyleBackColor = false;
            // 
            // rightButton
            // 
            this.rightButton.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.rightButton.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("rightButton.BackgroundImage")));
            this.rightButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.rightButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(39)))), ((int)(((byte)(39)))), ((int)(((byte)(226)))));
            this.rightButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(221)))), ((int)(((byte)(221)))), ((int)(((byte)(249)))));
            this.rightButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(252)))));
            this.rightButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.rightButton.Location = new System.Drawing.Point(379, 40);
            this.rightButton.Name = "rightButton";
            this.rightButton.Size = new System.Drawing.Size(50, 50);
            this.rightButton.TabIndex = 4;
            this.rightButton.UseVisualStyleBackColor = false;
            // 
            // durationLabel
            // 
            this.durationLabel.AutoSize = true;
            this.durationLabel.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.durationLabel.Location = new System.Drawing.Point(269, 132);
            this.durationLabel.Name = "durationLabel";
            this.durationLabel.Size = new System.Drawing.Size(182, 31);
            this.durationLabel.TabIndex = 5;
            this.durationLabel.Text = "动画持续时间：";
            // 
            // AnimationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(192F, 192F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(503, 463);
            this.Controls.Add(this.tabControl);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "AnimationForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "书写动画";
            this.TopMost = true;
            this.tabControl.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.adjustPanel.ResumeLayout(false);
            this.adjustPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.multiDurationControl)).EndInit();
            this.ResumeLayout(false);

        }
    }
}
