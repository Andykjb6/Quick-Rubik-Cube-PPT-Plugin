namespace 课件帮PPT助手
{
    partial class SVGSettingsForm
    {
        private System.Windows.Forms.ListBox svgListBox;
        private System.Windows.Forms.Button selectButton;

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SVGSettingsForm));
            this.svgListBox = new System.Windows.Forms.ListBox();
            this.selectButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // svgListBox
            // 
            this.svgListBox.DisplayMember = "Key";
            this.svgListBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.svgListBox.ItemHeight = 24;
            this.svgListBox.Location = new System.Drawing.Point(0, 0);
            this.svgListBox.Name = "svgListBox";
            this.svgListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.svgListBox.Size = new System.Drawing.Size(652, 306);
            this.svgListBox.TabIndex = 0;
            this.svgListBox.ValueMember = "Value";
            // 
            // selectButton
            // 
            this.selectButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(31)))), ((int)(((byte)(91)))), ((int)(((byte)(240)))));
            this.selectButton.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.selectButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.selectButton.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.selectButton.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.selectButton.Location = new System.Drawing.Point(0, 306);
            this.selectButton.Name = "selectButton";
            this.selectButton.Size = new System.Drawing.Size(652, 56);
            this.selectButton.TabIndex = 1;
            this.selectButton.Text = "确认插入";
            this.selectButton.UseVisualStyleBackColor = false;
            this.selectButton.Click += new System.EventHandler(this.SelectButton_Click);
            // 
            // SVGSettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(192F, 192F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(652, 362);
            this.Controls.Add(this.svgListBox);
            this.Controls.Add(this.selectButton);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SVGSettingsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "字源字形查询列表";
            this.ResumeLayout(false);

        }
    }
}
