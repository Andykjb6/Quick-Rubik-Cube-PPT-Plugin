using System.Drawing;
using System.Windows.Forms;

namespace 课件帮PPT助手
{
    partial class CustomizeAnnotationForm
    {
        private TextBox symbolTextBox;
        private TextBox nameTextBox;
        private RadioButton bottomRadioButton;
        private RadioButton startEndRadioButton;
        private RadioButton endRadioButton;
        private Button saveButton;
        private Button cancelButton;

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CustomizeAnnotationForm));
            this.symbolTextBox = new System.Windows.Forms.TextBox();
            this.nameTextBox = new System.Windows.Forms.TextBox();
            this.bottomRadioButton = new System.Windows.Forms.RadioButton();
            this.startEndRadioButton = new System.Windows.Forms.RadioButton();
            this.endRadioButton = new System.Windows.Forms.RadioButton();
            this.saveButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.symbolLabel = new System.Windows.Forms.Label();
            this.nameLabel = new System.Windows.Forms.Label();
            this.positionLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // symbolTextBox
            // 
            this.symbolTextBox.Location = new System.Drawing.Point(160, 20);
            this.symbolTextBox.Name = "symbolTextBox";
            this.symbolTextBox.Size = new System.Drawing.Size(300, 35);
            this.symbolTextBox.TabIndex = 1;
            // 
            // nameTextBox
            // 
            this.nameTextBox.Location = new System.Drawing.Point(160, 68);
            this.nameTextBox.Name = "nameTextBox";
            this.nameTextBox.Size = new System.Drawing.Size(300, 35);
            this.nameTextBox.TabIndex = 3;
            // 
            // bottomRadioButton
            // 
            this.bottomRadioButton.Checked = true;
            this.bottomRadioButton.Location = new System.Drawing.Point(160, 120);
            this.bottomRadioButton.Name = "bottomRadioButton";
            this.bottomRadioButton.Size = new System.Drawing.Size(220, 40);
            this.bottomRadioButton.TabIndex = 5;
            this.bottomRadioButton.TabStop = true;
            this.bottomRadioButton.Text = "所选文本的底部";
            // 
            // startEndRadioButton
            // 
            this.startEndRadioButton.Location = new System.Drawing.Point(160, 170);
            this.startEndRadioButton.Name = "startEndRadioButton";
            this.startEndRadioButton.Size = new System.Drawing.Size(300, 40);
            this.startEndRadioButton.TabIndex = 6;
            this.startEndRadioButton.Text = "所选文本的开头和末尾";
            // 
            // endRadioButton
            // 
            this.endRadioButton.Location = new System.Drawing.Point(160, 220);
            this.endRadioButton.Name = "endRadioButton";
            this.endRadioButton.Size = new System.Drawing.Size(220, 40);
            this.endRadioButton.TabIndex = 7;
            this.endRadioButton.Text = "所选文本的末尾";
            // 
            // saveButton
            // 
            this.saveButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(84)))), ((int)(((byte)(236)))));
            this.saveButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(111)))), ((int)(((byte)(233)))));
            this.saveButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(107)))), ((int)(((byte)(149)))), ((int)(((byte)(253)))));
            this.saveButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.saveButton.Font = new System.Drawing.Font("微软雅黑", 11F, System.Drawing.FontStyle.Bold);
            this.saveButton.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.saveButton.Location = new System.Drawing.Point(131, 301);
            this.saveButton.Name = "saveButton";
            this.saveButton.Size = new System.Drawing.Size(100, 50);
            this.saveButton.TabIndex = 8;
            this.saveButton.Text = "保存";
            this.saveButton.UseVisualStyleBackColor = false;
            // 
            // cancelButton
            // 
            this.cancelButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(84)))), ((int)(((byte)(236)))));
            this.cancelButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(111)))), ((int)(((byte)(233)))));
            this.cancelButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(107)))), ((int)(((byte)(149)))), ((int)(((byte)(253)))));
            this.cancelButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cancelButton.Font = new System.Drawing.Font("微软雅黑", 11F, System.Drawing.FontStyle.Bold);
            this.cancelButton.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.cancelButton.Location = new System.Drawing.Point(261, 301);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(100, 50);
            this.cancelButton.TabIndex = 9;
            this.cancelButton.Text = "退出";
            this.cancelButton.UseVisualStyleBackColor = false;
            // 
            // symbolLabel
            // 
            this.symbolLabel.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.symbolLabel.Location = new System.Drawing.Point(20, 20);
            this.symbolLabel.Name = "symbolLabel";
            this.symbolLabel.Size = new System.Drawing.Size(140, 35);
            this.symbolLabel.TabIndex = 0;
            this.symbolLabel.Text = "标注符号：";
            // 
            // nameLabel
            // 
            this.nameLabel.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.nameLabel.Location = new System.Drawing.Point(20, 73);
            this.nameLabel.Name = "nameLabel";
            this.nameLabel.Size = new System.Drawing.Size(140, 32);
            this.nameLabel.TabIndex = 2;
            this.nameLabel.Text = "符号名称：";
            // 
            // positionLabel
            // 
            this.positionLabel.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.positionLabel.Location = new System.Drawing.Point(20, 120);
            this.positionLabel.Name = "positionLabel";
            this.positionLabel.Size = new System.Drawing.Size(140, 40);
            this.positionLabel.TabIndex = 4;
            this.positionLabel.Text = "标注位置：";
            // 
            // CustomizeAnnotationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(192F, 192F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(489, 402);
            this.Controls.Add(this.symbolLabel);
            this.Controls.Add(this.symbolTextBox);
            this.Controls.Add(this.nameLabel);
            this.Controls.Add(this.nameTextBox);
            this.Controls.Add(this.positionLabel);
            this.Controls.Add(this.bottomRadioButton);
            this.Controls.Add(this.startEndRadioButton);
            this.Controls.Add(this.endRadioButton);
            this.Controls.Add(this.saveButton);
            this.Controls.Add(this.cancelButton);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "CustomizeAnnotationForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "自定义标注符号";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private Label symbolLabel;
        private Label nameLabel;
        private Label positionLabel;
    }
}
