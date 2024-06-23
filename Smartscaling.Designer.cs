namespace 课件帮PPT助手
{
    partial class SmartScalingForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.TrackBar trackBar;
        private System.Windows.Forms.NumericUpDown numericUpDown;
        private System.Windows.Forms.Button applyButton;
        private System.Windows.Forms.Button resetButton;
        private System.Windows.Forms.CheckBox checkBoxPropertySettings;
        private System.Windows.Forms.GroupBox groupBoxShapeAttributes;
        private System.Windows.Forms.CheckBox checkBoxShadow;
        private System.Windows.Forms.CheckBox checkBoxReflection;
        private System.Windows.Forms.CheckBox checkBoxGlow;
        private System.Windows.Forms.CheckBox checkBox3D;
        private System.Windows.Forms.CheckBox checkBoxTable; // 新增复选框
        private System.Windows.Forms.GroupBox groupBoxTextAttributes;
        private System.Windows.Forms.CheckBox checkBoxText;
        private System.Windows.Forms.CheckBox checkBoxTextShadow;
        private System.Windows.Forms.CheckBox checkBoxTextReflection;
        private System.Windows.Forms.CheckBox checkBoxTextGlow;
        private System.Windows.Forms.CheckBox checkBoxText3D;
        private System.Windows.Forms.GroupBox groupBoxCenterSelection;
        private System.Windows.Forms.CheckBox checkBoxCenter;
        private System.Windows.Forms.CheckBox checkBoxTopLeft;
        private System.Windows.Forms.CheckBox checkBoxTopRight;
        private System.Windows.Forms.CheckBox checkBoxBottomLeft;
        private System.Windows.Forms.CheckBox checkBoxBottomRight;

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
            this.trackBar = new System.Windows.Forms.TrackBar();
            this.numericUpDown = new System.Windows.Forms.NumericUpDown();
            this.applyButton = new System.Windows.Forms.Button();
            this.resetButton = new System.Windows.Forms.Button();
            this.checkBoxPropertySettings = new System.Windows.Forms.CheckBox();
            this.groupBoxShapeAttributes = new System.Windows.Forms.GroupBox();
            this.checkBoxTable = new System.Windows.Forms.CheckBox();
            this.checkBox3D = new System.Windows.Forms.CheckBox();
            this.checkBoxGlow = new System.Windows.Forms.CheckBox();
            this.checkBoxReflection = new System.Windows.Forms.CheckBox();
            this.checkBoxShadow = new System.Windows.Forms.CheckBox();
            this.groupBoxTextAttributes = new System.Windows.Forms.GroupBox();
            this.checkBoxText = new System.Windows.Forms.CheckBox();
            this.checkBoxTextShadow = new System.Windows.Forms.CheckBox();
            this.checkBoxTextReflection = new System.Windows.Forms.CheckBox();
            this.checkBoxTextGlow = new System.Windows.Forms.CheckBox();
            this.checkBoxText3D = new System.Windows.Forms.CheckBox();
            this.groupBoxCenterSelection = new System.Windows.Forms.GroupBox();
            this.checkBoxCenter = new System.Windows.Forms.CheckBox();
            this.checkBoxTopLeft = new System.Windows.Forms.CheckBox();
            this.checkBoxTopRight = new System.Windows.Forms.CheckBox();
            this.checkBoxBottomLeft = new System.Windows.Forms.CheckBox();
            this.checkBoxBottomRight = new System.Windows.Forms.CheckBox();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.trackBar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown)).BeginInit();
            this.groupBoxShapeAttributes.SuspendLayout();
            this.groupBoxTextAttributes.SuspendLayout();
            this.groupBoxCenterSelection.SuspendLayout();
            this.SuspendLayout();
            // 
            // trackBar
            // 
            this.trackBar.Location = new System.Drawing.Point(12, 12);
            this.trackBar.Maximum = 200;
            this.trackBar.Minimum = 10;
            this.trackBar.Name = "trackBar";
            this.trackBar.Size = new System.Drawing.Size(351, 90);
            this.trackBar.TabIndex = 0;
            this.trackBar.Value = 100;
            this.trackBar.Scroll += new System.EventHandler(this.trackBar_Scroll);
            // 
            // numericUpDown
            // 
            this.numericUpDown.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.numericUpDown.Location = new System.Drawing.Point(369, 12);
            this.numericUpDown.Maximum = new decimal(new int[] {
            200,
            0,
            0,
            0});
            this.numericUpDown.Minimum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numericUpDown.Name = "numericUpDown";
            this.numericUpDown.Size = new System.Drawing.Size(89, 39);
            this.numericUpDown.TabIndex = 1;
            this.numericUpDown.Value = new decimal(new int[] {
            100,
            0,
            0,
            0});
            this.numericUpDown.ValueChanged += new System.EventHandler(this.numericUpDown_ValueChanged);
            // 
            // applyButton
            // 
            this.applyButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(44)))), ((int)(((byte)(68)))), ((int)(((byte)(236)))));
            this.applyButton.FlatAppearance.BorderSize = 0;
            this.applyButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(47)))), ((int)(((byte)(222)))));
            this.applyButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(100)))), ((int)(((byte)(246)))));
            this.applyButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.applyButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.applyButton.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.applyButton.Location = new System.Drawing.Point(279, 71);
            this.applyButton.Name = "applyButton";
            this.applyButton.Size = new System.Drawing.Size(84, 43);
            this.applyButton.TabIndex = 2;
            this.applyButton.Text = "应用";
            this.applyButton.UseVisualStyleBackColor = false;
            this.applyButton.Click += new System.EventHandler(this.applyButton_Click);
            // 
            // resetButton
            // 
            this.resetButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(44)))), ((int)(((byte)(68)))), ((int)(((byte)(236)))));
            this.resetButton.FlatAppearance.BorderSize = 0;
            this.resetButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(21)))), ((int)(((byte)(47)))), ((int)(((byte)(222)))));
            this.resetButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(100)))), ((int)(((byte)(246)))));
            this.resetButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.resetButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.resetButton.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.resetButton.Location = new System.Drawing.Point(374, 71);
            this.resetButton.Name = "resetButton";
            this.resetButton.Size = new System.Drawing.Size(84, 43);
            this.resetButton.TabIndex = 3;
            this.resetButton.Text = "重置";
            this.resetButton.UseVisualStyleBackColor = false;
            this.resetButton.Click += new System.EventHandler(this.resetButton_Click);
            // 
            // checkBoxPropertySettings
            // 
            this.checkBoxPropertySettings.AutoSize = true;
            this.checkBoxPropertySettings.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBoxPropertySettings.Location = new System.Drawing.Point(24, 77);
            this.checkBoxPropertySettings.Name = "checkBoxPropertySettings";
            this.checkBoxPropertySettings.Size = new System.Drawing.Size(142, 35);
            this.checkBoxPropertySettings.TabIndex = 4;
            this.checkBoxPropertySettings.Text = "属性设置";
            this.checkBoxPropertySettings.UseVisualStyleBackColor = true;
            this.checkBoxPropertySettings.CheckedChanged += new System.EventHandler(this.checkBoxPropertySettings_CheckedChanged);
            // 
            // groupBoxShapeAttributes
            // 
            this.groupBoxShapeAttributes.Controls.Add(this.checkBoxTable);
            this.groupBoxShapeAttributes.Controls.Add(this.checkBox3D);
            this.groupBoxShapeAttributes.Controls.Add(this.checkBoxGlow);
            this.groupBoxShapeAttributes.Controls.Add(this.checkBoxReflection);
            this.groupBoxShapeAttributes.Controls.Add(this.checkBoxShadow);
            this.groupBoxShapeAttributes.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBoxShapeAttributes.Location = new System.Drawing.Point(18, 134);
            this.groupBoxShapeAttributes.Name = "groupBoxShapeAttributes";
            this.groupBoxShapeAttributes.Size = new System.Drawing.Size(440, 150);
            this.groupBoxShapeAttributes.TabIndex = 5;
            this.groupBoxShapeAttributes.TabStop = false;
            this.groupBoxShapeAttributes.Text = "形状属性";
            // 
            // checkBoxTable
            // 
            this.checkBoxTable.AutoSize = true;
            this.checkBoxTable.Checked = true;
            this.checkBoxTable.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxTable.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBoxTable.Location = new System.Drawing.Point(13, 87);
            this.checkBoxTable.Name = "checkBoxTable";
            this.checkBoxTable.Size = new System.Drawing.Size(94, 35);
            this.checkBoxTable.TabIndex = 4;
            this.checkBoxTable.Text = "表格";
            this.checkBoxTable.UseVisualStyleBackColor = true;
            // 
            // checkBox3D
            // 
            this.checkBox3D.AutoSize = true;
            this.checkBox3D.Checked = true;
            this.checkBox3D.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox3D.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBox3D.Location = new System.Drawing.Point(297, 46);
            this.checkBox3D.Name = "checkBox3D";
            this.checkBox3D.Size = new System.Drawing.Size(142, 35);
            this.checkBox3D.TabIndex = 3;
            this.checkBox3D.Text = "三维格式";
            this.checkBox3D.UseVisualStyleBackColor = true;
            // 
            // checkBoxGlow
            // 
            this.checkBoxGlow.AutoSize = true;
            this.checkBoxGlow.Checked = true;
            this.checkBoxGlow.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxGlow.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBoxGlow.Location = new System.Drawing.Point(202, 46);
            this.checkBoxGlow.Name = "checkBoxGlow";
            this.checkBoxGlow.Size = new System.Drawing.Size(94, 35);
            this.checkBoxGlow.TabIndex = 2;
            this.checkBoxGlow.Text = "发光";
            this.checkBoxGlow.UseVisualStyleBackColor = true;
            // 
            // checkBoxReflection
            // 
            this.checkBoxReflection.AutoSize = true;
            this.checkBoxReflection.Checked = true;
            this.checkBoxReflection.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxReflection.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBoxReflection.Location = new System.Drawing.Point(104, 46);
            this.checkBoxReflection.Name = "checkBoxReflection";
            this.checkBoxReflection.Size = new System.Drawing.Size(94, 35);
            this.checkBoxReflection.TabIndex = 1;
            this.checkBoxReflection.Text = "映像";
            this.checkBoxReflection.UseVisualStyleBackColor = true;
            // 
            // checkBoxShadow
            // 
            this.checkBoxShadow.AutoSize = true;
            this.checkBoxShadow.Checked = true;
            this.checkBoxShadow.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxShadow.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBoxShadow.Location = new System.Drawing.Point(13, 46);
            this.checkBoxShadow.Name = "checkBoxShadow";
            this.checkBoxShadow.Size = new System.Drawing.Size(94, 35);
            this.checkBoxShadow.TabIndex = 0;
            this.checkBoxShadow.Text = "阴影";
            this.checkBoxShadow.UseVisualStyleBackColor = true;
            // 
            // groupBoxTextAttributes
            // 
            this.groupBoxTextAttributes.Controls.Add(this.checkBoxText);
            this.groupBoxTextAttributes.Controls.Add(this.checkBoxTextShadow);
            this.groupBoxTextAttributes.Controls.Add(this.checkBoxTextReflection);
            this.groupBoxTextAttributes.Controls.Add(this.checkBoxTextGlow);
            this.groupBoxTextAttributes.Controls.Add(this.checkBoxText3D);
            this.groupBoxTextAttributes.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBoxTextAttributes.Location = new System.Drawing.Point(18, 310);
            this.groupBoxTextAttributes.Name = "groupBoxTextAttributes";
            this.groupBoxTextAttributes.Size = new System.Drawing.Size(440, 140);
            this.groupBoxTextAttributes.TabIndex = 6;
            this.groupBoxTextAttributes.TabStop = false;
            this.groupBoxTextAttributes.Text = "文字属性";
            
            // 
            // checkBoxText
            // 
            this.checkBoxText.AutoSize = true;
            this.checkBoxText.Checked = true;
            this.checkBoxText.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxText.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBoxText.Location = new System.Drawing.Point(13, 44);
            this.checkBoxText.Name = "checkBoxText";
            this.checkBoxText.Size = new System.Drawing.Size(94, 35);
            this.checkBoxText.TabIndex = 0;
            this.checkBoxText.Text = "文字";
            this.checkBoxText.UseVisualStyleBackColor = true;
            // 
            // checkBoxTextShadow
            // 
            this.checkBoxTextShadow.AutoSize = true;
            this.checkBoxTextShadow.Checked = true;
            this.checkBoxTextShadow.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxTextShadow.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBoxTextShadow.Location = new System.Drawing.Point(109, 44);
            this.checkBoxTextShadow.Name = "checkBoxTextShadow";
            this.checkBoxTextShadow.Size = new System.Drawing.Size(94, 35);
            this.checkBoxTextShadow.TabIndex = 1;
            this.checkBoxTextShadow.Text = "阴影";
            this.checkBoxTextShadow.UseVisualStyleBackColor = true;
            // 
            // checkBoxTextReflection
            // 
            this.checkBoxTextReflection.AutoSize = true;
            this.checkBoxTextReflection.Checked = true;
            this.checkBoxTextReflection.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxTextReflection.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBoxTextReflection.Location = new System.Drawing.Point(210, 44);
            this.checkBoxTextReflection.Name = "checkBoxTextReflection";
            this.checkBoxTextReflection.Size = new System.Drawing.Size(94, 35);
            this.checkBoxTextReflection.TabIndex = 2;
            this.checkBoxTextReflection.Text = "映像";
            this.checkBoxTextReflection.UseVisualStyleBackColor = true;
            // 
            // checkBoxTextGlow
            // 
            this.checkBoxTextGlow.AutoSize = true;
            this.checkBoxTextGlow.Checked = true;
            this.checkBoxTextGlow.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxTextGlow.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBoxTextGlow.Location = new System.Drawing.Point(319, 44);
            this.checkBoxTextGlow.Name = "checkBoxTextGlow";
            this.checkBoxTextGlow.Size = new System.Drawing.Size(94, 35);
            this.checkBoxTextGlow.TabIndex = 3;
            this.checkBoxTextGlow.Text = "发光";
            this.checkBoxTextGlow.UseVisualStyleBackColor = true;
            // 
            // checkBoxText3D
            // 
            this.checkBoxText3D.AutoSize = true;
            this.checkBoxText3D.Checked = true;
            this.checkBoxText3D.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxText3D.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.checkBoxText3D.Location = new System.Drawing.Point(13, 90);
            this.checkBoxText3D.Name = "checkBoxText3D";
            this.checkBoxText3D.Size = new System.Drawing.Size(142, 35);
            this.checkBoxText3D.TabIndex = 4;
            this.checkBoxText3D.Text = "三维格式";
            this.checkBoxText3D.UseVisualStyleBackColor = true;
            // 
            // groupBoxCenterSelection
            // 
            this.groupBoxCenterSelection.Controls.Add(this.checkBoxCenter);
            this.groupBoxCenterSelection.Controls.Add(this.checkBoxTopLeft);
            this.groupBoxCenterSelection.Controls.Add(this.checkBoxTopRight);
            this.groupBoxCenterSelection.Controls.Add(this.checkBoxBottomLeft);
            this.groupBoxCenterSelection.Controls.Add(this.checkBoxBottomRight);
            this.groupBoxCenterSelection.Controls.Add(this.button1);
            this.groupBoxCenterSelection.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBoxCenterSelection.Location = new System.Drawing.Point(18, 462);
            this.groupBoxCenterSelection.Name = "groupBoxCenterSelection";
            this.groupBoxCenterSelection.Size = new System.Drawing.Size(440, 172);
            this.groupBoxCenterSelection.TabIndex = 7;
            this.groupBoxCenterSelection.TabStop = false;
            this.groupBoxCenterSelection.Text = "缩放中心";
            // 
            // checkBoxCenter
            // 
            this.checkBoxCenter.AutoSize = true;
            this.checkBoxCenter.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(197)))), ((int)(((byte)(205)))), ((int)(((byte)(255)))));
            this.checkBoxCenter.Checked = true;
            this.checkBoxCenter.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxCenter.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(76)))), ((int)(((byte)(244)))));
            this.checkBoxCenter.FlatAppearance.BorderSize = 2;
            this.checkBoxCenter.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(218)))), ((int)(((byte)(223)))), ((int)(((byte)(255)))));
            this.checkBoxCenter.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(14)))), ((int)(((byte)(39)))), ((int)(((byte)(207)))));
            this.checkBoxCenter.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(85)))), ((int)(((byte)(105)))), ((int)(((byte)(243)))));
            this.checkBoxCenter.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxCenter.Font = new System.Drawing.Font("微软雅黑", 10F);
            this.checkBoxCenter.Location = new System.Drawing.Point(210, 84);
            this.checkBoxCenter.Name = "checkBoxCenter";
            this.checkBoxCenter.Size = new System.Drawing.Size(23, 22);
            this.checkBoxCenter.TabIndex = 0;
            this.checkBoxCenter.UseVisualStyleBackColor = false;
            this.checkBoxCenter.CheckedChanged += new System.EventHandler(this.checkBoxCenter_CheckedChanged);
            // 
            // checkBoxTopLeft
            // 
            this.checkBoxTopLeft.AutoSize = true;
            this.checkBoxTopLeft.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(197)))), ((int)(((byte)(205)))), ((int)(((byte)(255)))));
            this.checkBoxTopLeft.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(76)))), ((int)(((byte)(244)))));
            this.checkBoxTopLeft.FlatAppearance.BorderSize = 2;
            this.checkBoxTopLeft.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(218)))), ((int)(((byte)(223)))), ((int)(((byte)(255)))));
            this.checkBoxTopLeft.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(14)))), ((int)(((byte)(39)))), ((int)(((byte)(207)))));
            this.checkBoxTopLeft.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(85)))), ((int)(((byte)(105)))), ((int)(((byte)(243)))));
            this.checkBoxTopLeft.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxTopLeft.Font = new System.Drawing.Font("微软雅黑", 10F);
            this.checkBoxTopLeft.Location = new System.Drawing.Point(156, 35);
            this.checkBoxTopLeft.Name = "checkBoxTopLeft";
            this.checkBoxTopLeft.Size = new System.Drawing.Size(23, 22);
            this.checkBoxTopLeft.TabIndex = 1;
            this.checkBoxTopLeft.UseVisualStyleBackColor = false;
            this.checkBoxTopLeft.CheckedChanged += new System.EventHandler(this.checkBoxTopLeft_CheckedChanged);
            // 
            // checkBoxTopRight
            // 
            this.checkBoxTopRight.AutoSize = true;
            this.checkBoxTopRight.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(197)))), ((int)(((byte)(205)))), ((int)(((byte)(255)))));
            this.checkBoxTopRight.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(76)))), ((int)(((byte)(244)))));
            this.checkBoxTopRight.FlatAppearance.BorderSize = 2;
            this.checkBoxTopRight.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(218)))), ((int)(((byte)(223)))), ((int)(((byte)(255)))));
            this.checkBoxTopRight.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(14)))), ((int)(((byte)(39)))), ((int)(((byte)(207)))));
            this.checkBoxTopRight.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(85)))), ((int)(((byte)(105)))), ((int)(((byte)(243)))));
            this.checkBoxTopRight.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxTopRight.Font = new System.Drawing.Font("微软雅黑", 10F);
            this.checkBoxTopRight.Location = new System.Drawing.Point(266, 35);
            this.checkBoxTopRight.Name = "checkBoxTopRight";
            this.checkBoxTopRight.Size = new System.Drawing.Size(23, 22);
            this.checkBoxTopRight.TabIndex = 2;
            this.checkBoxTopRight.UseVisualStyleBackColor = false;
            this.checkBoxTopRight.CheckedChanged += new System.EventHandler(this.checkBoxTopRight_CheckedChanged);
            // 
            // checkBoxBottomLeft
            // 
            this.checkBoxBottomLeft.AutoSize = true;
            this.checkBoxBottomLeft.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(197)))), ((int)(((byte)(205)))), ((int)(((byte)(255)))));
            this.checkBoxBottomLeft.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(76)))), ((int)(((byte)(244)))));
            this.checkBoxBottomLeft.FlatAppearance.BorderSize = 2;
            this.checkBoxBottomLeft.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(218)))), ((int)(((byte)(223)))), ((int)(((byte)(255)))));
            this.checkBoxBottomLeft.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(14)))), ((int)(((byte)(39)))), ((int)(((byte)(207)))));
            this.checkBoxBottomLeft.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(85)))), ((int)(((byte)(105)))), ((int)(((byte)(243)))));
            this.checkBoxBottomLeft.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxBottomLeft.Font = new System.Drawing.Font("微软雅黑", 10F);
            this.checkBoxBottomLeft.Location = new System.Drawing.Point(156, 133);
            this.checkBoxBottomLeft.Name = "checkBoxBottomLeft";
            this.checkBoxBottomLeft.Size = new System.Drawing.Size(23, 22);
            this.checkBoxBottomLeft.TabIndex = 3;
            this.checkBoxBottomLeft.UseVisualStyleBackColor = false;
            this.checkBoxBottomLeft.CheckedChanged += new System.EventHandler(this.checkBoxBottomLeft_CheckedChanged);
            // 
            // checkBoxBottomRight
            // 
            this.checkBoxBottomRight.AutoSize = true;
            this.checkBoxBottomRight.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(197)))), ((int)(((byte)(205)))), ((int)(((byte)(255)))));
            this.checkBoxBottomRight.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(76)))), ((int)(((byte)(244)))));
            this.checkBoxBottomRight.FlatAppearance.BorderSize = 2;
            this.checkBoxBottomRight.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(218)))), ((int)(((byte)(223)))), ((int)(((byte)(255)))));
            this.checkBoxBottomRight.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(14)))), ((int)(((byte)(39)))), ((int)(((byte)(207)))));
            this.checkBoxBottomRight.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(85)))), ((int)(((byte)(105)))), ((int)(((byte)(243)))));
            this.checkBoxBottomRight.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkBoxBottomRight.Font = new System.Drawing.Font("微软雅黑", 10F);
            this.checkBoxBottomRight.Location = new System.Drawing.Point(266, 133);
            this.checkBoxBottomRight.Name = "checkBoxBottomRight";
            this.checkBoxBottomRight.Size = new System.Drawing.Size(23, 22);
            this.checkBoxBottomRight.TabIndex = 4;
            this.checkBoxBottomRight.UseVisualStyleBackColor = false;
            this.checkBoxBottomRight.CheckedChanged += new System.EventHandler(this.checkBoxBottomRight_CheckedChanged);
            // 
            // button1
            // 
            this.button1.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(76)))), ((int)(((byte)(244)))));
            this.button1.FlatAppearance.BorderSize = 2;
            this.button1.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.button1.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Location = new System.Drawing.Point(166, 45);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(112, 101);
            this.button1.TabIndex = 5;
            this.button1.UseVisualStyleBackColor = true;
            // 
            // SmartScalingForm
            // 
            this.ClientSize = new System.Drawing.Size(482, 650);
            this.Controls.Add(this.groupBoxCenterSelection);
            this.Controls.Add(this.groupBoxTextAttributes);
            this.Controls.Add(this.groupBoxShapeAttributes);
            this.Controls.Add(this.checkBoxPropertySettings);
            this.Controls.Add(this.resetButton);
            this.Controls.Add(this.applyButton);
            this.Controls.Add(this.numericUpDown);
            this.Controls.Add(this.trackBar);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "SmartScalingForm";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "智能缩放";
            this.TopMost = true;
            ((System.ComponentModel.ISupportInitialize)(this.trackBar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown)).EndInit();
            this.groupBoxShapeAttributes.ResumeLayout(false);
            this.groupBoxShapeAttributes.PerformLayout();
            this.groupBoxTextAttributes.ResumeLayout(false);
            this.groupBoxTextAttributes.PerformLayout();
            this.groupBoxCenterSelection.ResumeLayout(false);
            this.groupBoxCenterSelection.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.Windows.Forms.Button button1;
    }
}
