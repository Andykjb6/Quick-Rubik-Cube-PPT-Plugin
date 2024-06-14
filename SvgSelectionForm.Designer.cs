namespace 课件帮PPT助手
{
    partial class SvgSelectionForm
    {
        private System.Windows.Forms.ListBox svgListBox;
        private System.Windows.Forms.Button selectButton;

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SvgSelectionForm));
            this.svgListBox = new System.Windows.Forms.ListBox();
            this.selectButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // svgListBox
            // 
            this.svgListBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.svgListBox.ItemHeight = 24;
            this.svgListBox.Location = new System.Drawing.Point(0, 0);
            this.svgListBox.Name = "svgListBox";
            this.svgListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.svgListBox.Size = new System.Drawing.Size(548, 218);
            this.svgListBox.TabIndex = 0;
            // 
            // selectButton
            // 
            this.selectButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(39)))), ((int)(((byte)(96)))), ((int)(((byte)(241)))));
            this.selectButton.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.selectButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.selectButton.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.selectButton.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.selectButton.Location = new System.Drawing.Point(0, 218);
            this.selectButton.Name = "selectButton";
            this.selectButton.Size = new System.Drawing.Size(548, 54);
            this.selectButton.TabIndex = 1;
            this.selectButton.Text = "确认插入";
            this.selectButton.UseVisualStyleBackColor = false;
            this.selectButton.Click += new System.EventHandler(this.SelectButton_Click);
            // 
            // SvgSelectionForm
            // 
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(548, 272);
            this.Controls.Add(this.svgListBox);
            this.Controls.Add(this.selectButton);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SvgSelectionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "笔顺查询列表";
            this.ResumeLayout(false);

        }
    }
}
