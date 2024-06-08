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

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AnimationForm));
            this.tabControl = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.label1 = new System.Windows.Forms.Label();
            this.inputLabel = new System.Windows.Forms.Label();
            this.textBox = new System.Windows.Forms.TextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.label2 = new System.Windows.Forms.Label();
            this.selectAllButton = new System.Windows.Forms.Button();
            this.animateButton = new System.Windows.Forms.Button();
            this.adjustAnimationButton = new System.Windows.Forms.Button();
            this.adjustPanel = new System.Windows.Forms.Panel();
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
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.inputLabel);
            this.tabPage1.Controls.Add(this.textBox);
            this.tabPage1.Location = new System.Drawing.Point(8, 39);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Size = new System.Drawing.Size(487, 416);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "第一步";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(19, 36);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(466, 24);
            this.label1.TabIndex = 0;
            this.label1.Text = "提示：请按照笔画顺序依次选中所有笔画。";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // inputLabel
            // 
            this.inputLabel.AutoSize = true;
            this.inputLabel.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.inputLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(19)))), ((int)(((byte)(84)))), ((int)(((byte)(247)))));
            this.inputLabel.Location = new System.Drawing.Point(19, 90);
            this.inputLabel.Name = "inputLabel";
            this.inputLabel.Size = new System.Drawing.Size(273, 28);
            this.inputLabel.TabIndex = 1;
            this.inputLabel.Text = "①请输入对应汉字：";
            // 
            // textBox
            // 
            this.textBox.Font = new System.Drawing.Font("宋体", 16.125F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox.Location = new System.Drawing.Point(22, 137);
            this.textBox.Name = "textBox";
            this.textBox.Size = new System.Drawing.Size(440, 57);
            this.textBox.TabIndex = 2;
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
            this.label2.Location = new System.Drawing.Point(10, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(418, 24);
            this.label2.TabIndex = 0;
            this.label2.Text = "提示：“智能全选”→“智能动画”。";
            // 
            // selectAllButton
            // 
            this.selectAllButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(15)))), ((int)(((byte)(81)))), ((int)(((byte)(246)))));
            this.selectAllButton.ForeColor = System.Drawing.Color.White;
            this.selectAllButton.Location = new System.Drawing.Point(10, 65);
            this.selectAllButton.Name = "selectAllButton";
            this.selectAllButton.Size = new System.Drawing.Size(220, 45);
            this.selectAllButton.TabIndex = 1;
            this.selectAllButton.Text = "②智能全选";
            this.selectAllButton.UseVisualStyleBackColor = false;
            // 
            // animateButton
            // 
            this.animateButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(15)))), ((int)(((byte)(81)))), ((int)(((byte)(246)))));
            this.animateButton.ForeColor = System.Drawing.Color.White;
            this.animateButton.Location = new System.Drawing.Point(250, 65);
            this.animateButton.Name = "animateButton";
            this.animateButton.Size = new System.Drawing.Size(220, 45);
            this.animateButton.TabIndex = 2;
            this.animateButton.Text = "③智能动画";
            this.animateButton.UseVisualStyleBackColor = false;
            // 
            // adjustAnimationButton
            // 
            this.adjustAnimationButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(15)))), ((int)(((byte)(81)))), ((int)(((byte)(246)))));
            this.adjustAnimationButton.ForeColor = System.Drawing.Color.White;
            this.adjustAnimationButton.Location = new System.Drawing.Point(10, 120);
            this.adjustAnimationButton.Name = "adjustAnimationButton";
            this.adjustAnimationButton.Size = new System.Drawing.Size(458, 45);
            this.adjustAnimationButton.TabIndex = 3;
            this.adjustAnimationButton.Text = "动画调整";
            this.adjustAnimationButton.UseVisualStyleBackColor = false;
            // 
            // adjustPanel
            // 
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
            this.upButton.Location = new System.Drawing.Point(320, 10);
            this.upButton.Name = "upButton";
            this.upButton.Size = new System.Drawing.Size(50, 50);
            this.upButton.TabIndex = 1;
            this.upButton.Text = "↑";
            // 
            // downButton
            // 
            this.downButton.Location = new System.Drawing.Point(320, 70);
            this.downButton.Name = "downButton";
            this.downButton.Size = new System.Drawing.Size(50, 50);
            this.downButton.TabIndex = 2;
            this.downButton.Text = "↓";
            // 
            // leftButton
            // 
            this.leftButton.Location = new System.Drawing.Point(270, 40);
            this.leftButton.Name = "leftButton";
            this.leftButton.Size = new System.Drawing.Size(50, 50);
            this.leftButton.TabIndex = 3;
            this.leftButton.Text = "←";
            // 
            // rightButton
            // 
            this.rightButton.Location = new System.Drawing.Point(370, 40);
            this.rightButton.Name = "rightButton";
            this.rightButton.Size = new System.Drawing.Size(50, 50);
            this.rightButton.TabIndex = 4;
            this.rightButton.Text = "→";
            // 
            // durationLabel
            // 
            this.durationLabel.AutoSize = true;
            this.durationLabel.Location = new System.Drawing.Point(270, 135);
            this.durationLabel.Name = "durationLabel";
            this.durationLabel.Size = new System.Drawing.Size(178, 24);
            this.durationLabel.TabIndex = 5;
            this.durationLabel.Text = "动画持续时间：";
            // 
            // AnimationForm
            // 
            this.ClientSize = new System.Drawing.Size(503, 463);
            this.Controls.Add(this.tabControl);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "AnimationForm";
            this.Text = "书写动画";
            this.TopMost = true;
            this.tabControl.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.adjustPanel.ResumeLayout(false);
            this.adjustPanel.PerformLayout();
            this.ResumeLayout(false);

        }
    }
}
