using Microsoft.Office.Tools.Ribbon;
using System;

namespace 课件帮PPT助手
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl9 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl10 = this.Factory.CreateRibbonDropDownItem();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.课件帮PPT助手 = this.Factory.CreateRibbonTab();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.group10 = this.Factory.CreateRibbonGroup();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.group8 = this.Factory.CreateRibbonGroup();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.group9 = this.Factory.CreateRibbonGroup();
            this.comboBox1 = this.Factory.CreateRibbonComboBox();
            this.comboBox2 = this.Factory.CreateRibbonComboBox();
            this.关于我 = this.Factory.CreateRibbonSplitButton();
            this.检查更新 = this.Factory.CreateRibbonButton();
            this.toggleTaskPaneButton = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.笔顺图解 = this.Factory.CreateRibbonButton();
            this.生字赋格 = this.Factory.CreateRibbonButton();
            this.注音编辑 = this.Factory.CreateRibbonSplitButton();
            this.文本居中 = this.Factory.CreateRibbonButton();
            this.删列补行 = this.Factory.CreateRibbonButton();
            this.增加行宽 = this.Factory.CreateRibbonButton();
            this.奇数行行高设置 = this.Factory.CreateRibbonButton();
            this.自动补齐 = this.Factory.CreateRibbonButton();
            this.合并段落 = this.Factory.CreateRibbonButton();
            this.字号调整 = this.Factory.CreateRibbonButton();
            this.晓声通在线注音 = this.Factory.CreateRibbonSplitButton();
            this.在线注音编辑器 = this.Factory.CreateRibbonButton();
            this.文转表格 = this.Factory.CreateRibbonButton();
            this.重设表格 = this.Factory.CreateRibbonButton();
            this.常用格子 = this.Factory.CreateRibbonSplitButton();
            this.生字格子 = this.Factory.CreateRibbonButton();
            this.四线三格 = this.Factory.CreateRibbonButton();
            this.注音工具 = this.Factory.CreateRibbonSplitButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.便捷注音 = this.Factory.CreateRibbonButton();
            this.一键注音 = this.Factory.CreateRibbonButton();
            this.拓展应用 = this.Factory.CreateRibbonSplitButton();
            this.Zici = this.Factory.CreateRibbonButton();
            this.WritePinyin = this.Factory.CreateRibbonButton();
            this.多音字词填空 = this.Factory.CreateRibbonButton();
            this.分解拼音 = this.Factory.CreateRibbonButton();
            this.拼音升调 = this.Factory.CreateRibbonButton();
            this.Masking = this.Factory.CreateRibbonButton();
            this.图形分割 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.Gradientrectangle = this.Factory.CreateRibbonButton();
            this.在线工具 = this.Factory.CreateRibbonSplitButton();
            this.抠图 = this.Factory.CreateRibbonSplitButton();
            this.button19 = this.Factory.CreateRibbonButton();
            this.趣作图 = this.Factory.CreateRibbonButton();
            this.Bgsub = this.Factory.CreateRibbonButton();
            this.矢量 = this.Factory.CreateRibbonSplitButton();
            this.Tmttool = this.Factory.CreateRibbonButton();
            this.矩形拆分 = this.Factory.CreateRibbonButton();
            this.Mosaic = this.Factory.CreateRibbonButton();
            this.ApplyFilter = this.Factory.CreateRibbonButton();
            this.Expandimage = this.Factory.CreateRibbonButton();
            this.平移居中 = this.Factory.CreateRibbonButton();
            this.分布 = this.Factory.CreateRibbonMenu();
            this.沿线分布 = this.Factory.CreateRibbonButton();
            this.矩阵分布 = this.Factory.CreateRibbonButton();
            this.环形分布 = this.Factory.CreateRibbonButton();
            this.贴边对齐 = this.Factory.CreateRibbonSplitButton();
            this.button11 = this.Factory.CreateRibbonButton();
            this.button13 = this.Factory.CreateRibbonButton();
            this.选择居中 = this.Factory.CreateRibbonSplitButton();
            this.button10 = this.Factory.CreateRibbonButton();
            this.Pagecentered = this.Factory.CreateRibbonButton();
            this.筛选 = this.Factory.CreateRibbonSplitButton();
            this.Type = this.Factory.CreateRibbonButton();
            this.Selectsize = this.Factory.CreateRibbonButton();
            this.SelectedColor = this.Factory.CreateRibbonButton();
            this.Selectedline = this.Factory.CreateRibbonButton();
            this.Selectfontsize = this.Factory.CreateRibbonButton();
            this.图层 = this.Factory.CreateRibbonButton();
            this.选择增强 = this.Factory.CreateRibbonButton();
            this.智能缩放 = this.Factory.CreateRibbonButton();
            this.文本 = this.Factory.CreateRibbonSplitButton();
            this.去除边距 = this.Factory.CreateRibbonButton();
            this.首行缩进 = this.Factory.CreateRibbonButton();
            this.单字拆分 = this.Factory.CreateRibbonButton();
            this.拆分段落 = this.Factory.CreateRibbonButton();
            this.批量改字 = this.Factory.CreateRibbonButton();
            this.字词加点 = this.Factory.CreateRibbonButton();
            this.文本矢量化 = this.Factory.CreateRibbonButton();
            this.文写入形 = this.Factory.CreateRibbonButton();
            this.部首描红 = this.Factory.CreateRibbonButton();
            this.分解笔顺 = this.Factory.CreateRibbonButton();
            this.绘图 = this.Factory.CreateRibbonSplitButton();
            this.辐射连线 = this.Factory.CreateRibbonButton();
            this.层级关系 = this.Factory.CreateRibbonButton();
            this.括弧关系 = this.Factory.CreateRibbonButton();
            this.更多便捷 = this.Factory.CreateRibbonMenu();
            this.统一 = this.Factory.CreateRibbonSplitButton();
            this.统一大小 = this.Factory.CreateRibbonButton();
            this.统一格式 = this.Factory.CreateRibbonButton();
            this.统一控点 = this.Factory.CreateRibbonButton();
            this.交换 = this.Factory.CreateRibbonSplitButton();
            this.交换位置 = this.Factory.CreateRibbonButton();
            this.交换文字 = this.Factory.CreateRibbonButton();
            this.交换格式 = this.Factory.CreateRibbonButton();
            this.交换尺寸 = this.Factory.CreateRibbonButton();
            this.组合 = this.Factory.CreateRibbonSplitButton();
            this.重叠组合 = this.Factory.CreateRibbonButton();
            this.临近组合 = this.Factory.CreateRibbonButton();
            this.同色组合 = this.Factory.CreateRibbonButton();
            this.复制 = this.Factory.CreateRibbonSplitButton();
            this.左右镜像 = this.Factory.CreateRibbonButton();
            this.上下镜像 = this.Factory.CreateRibbonButton();
            this.LCopy = this.Factory.CreateRibbonButton();
            this.图片 = this.Factory.CreateRibbonSplitButton();
            this.批量换图 = this.Factory.CreateRibbonButton();
            this.形状填图 = this.Factory.CreateRibbonButton();
            this.原位转图 = this.Factory.CreateRibbonButton();
            this.原位转JPG = this.Factory.CreateRibbonButton();
            this.删除 = this.Factory.CreateRibbonSplitButton();
            this.删除动画 = this.Factory.CreateRibbonButton();
            this.清空页外 = this.Factory.CreateRibbonButton();
            this.清除备注 = this.Factory.CreateRibbonButton();
            this.清除超链接 = this.Factory.CreateRibbonButton();
            this.删除未用版式 = this.Factory.CreateRibbonButton();
            this.Timer = this.Factory.CreateRibbonButton();
            this.板贴辅助 = this.Factory.CreateRibbonButton();
            this.检测字体 = this.Factory.CreateRibbonButton();
            this.生成样机 = this.Factory.CreateRibbonButton();
            this.图形修剪 = this.Factory.CreateRibbonButton();
            this.button20 = this.Factory.CreateRibbonButton();
            this.快捷盒子 = this.Factory.CreateRibbonButton();
            this.Replaceaudio = this.Factory.CreateRibbonButton();
            this.插入矩形 = this.Factory.CreateRibbonButton();
            this.splitButton2 = this.Factory.CreateRibbonSplitButton();
            this.课件帮PPT助手.SuspendLayout();
            this.group3.SuspendLayout();
            this.group10.SuspendLayout();
            this.group2.SuspendLayout();
            this.group1.SuspendLayout();
            this.group4.SuspendLayout();
            this.group8.SuspendLayout();
            this.group6.SuspendLayout();
            this.group9.SuspendLayout();
            this.SuspendLayout();
            // 
            // 课件帮PPT助手
            // 
            this.课件帮PPT助手.Groups.Add(this.group3);
            this.课件帮PPT助手.Groups.Add(this.group10);
            this.课件帮PPT助手.Groups.Add(this.group2);
            this.课件帮PPT助手.Groups.Add(this.group1);
            this.课件帮PPT助手.Groups.Add(this.group4);
            this.课件帮PPT助手.Groups.Add(this.group8);
            this.课件帮PPT助手.Groups.Add(this.group6);
            this.课件帮PPT助手.Groups.Add(this.group9);
            this.课件帮PPT助手.Label = "快捷魔方";
            this.课件帮PPT助手.Name = "课件帮PPT助手";
            // 
            // group3
            // 
            this.group3.Items.Add(this.关于我);
            this.group3.Label = "关于我";
            this.group3.Name = "group3";
            // 
            // group10
            // 
            this.group10.Items.Add(this.toggleTaskPaneButton);
            this.group10.Name = "group10";
            // 
            // group2
            // 
            this.group2.Items.Add(this.button5);
            this.group2.Items.Add(this.笔顺图解);
            this.group2.Items.Add(this.生字赋格);
            this.group2.Items.Add(this.注音编辑);
            this.group2.Items.Add(this.常用格子);
            this.group2.Items.Add(this.注音工具);
            this.group2.Items.Add(this.拓展应用);
            this.group2.Label = "字音字形";
            this.group2.Name = "group2";
            // 
            // group1
            // 
            this.group1.Items.Add(this.Masking);
            this.group1.Items.Add(this.图形分割);
            this.group1.Items.Add(this.button6);
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.Gradientrectangle);
            this.group1.Items.Add(this.在线工具);
            this.group1.Items.Add(this.矩形拆分);
            this.group1.Items.Add(this.Mosaic);
            this.group1.Items.Add(this.ApplyFilter);
            this.group1.Items.Add(this.Expandimage);
            this.group1.Label = "图形处理";
            this.group1.Name = "group1";
            // 
            // group4
            // 
            this.group4.Items.Add(this.平移居中);
            this.group4.Items.Add(this.分布);
            this.group4.Label = "参考对齐";
            this.group4.Name = "group4";
            // 
            // group8
            // 
            this.group8.Items.Add(this.筛选);
            this.group8.Items.Add(this.选择增强);
            this.group8.Items.Add(this.智能缩放);
            this.group8.Items.Add(this.文本);
            this.group8.Items.Add(this.绘图);
            this.group8.Items.Add(this.更多便捷);
            this.group8.Label = "便捷常用";
            this.group8.Name = "group8";
            // 
            // group6
            // 
            this.group6.Items.Add(this.Timer);
            this.group6.Items.Add(this.板贴辅助);
            this.group6.Items.Add(this.检测字体);
            this.group6.Items.Add(this.生成样机);
            this.group6.Items.Add(this.图形修剪);
            this.group6.Items.Add(this.button20);
            this.group6.Items.Add(this.快捷盒子);
            this.group6.Items.Add(this.Replaceaudio);
            this.group6.Items.Add(this.插入矩形);
            this.group6.Label = "辅助";
            this.group6.Name = "group6";
            // 
            // group9
            // 
            this.group9.Items.Add(this.comboBox1);
            this.group9.Items.Add(this.comboBox2);
            this.group9.Label = "页面布局";
            this.group9.Name = "group9";
            // 
            // comboBox1
            // 
            this.comboBox1.Image = ((System.Drawing.Image)(resources.GetObject("comboBox1.Image")));
            ribbonDropDownItemImpl1.Label = "A4";
            ribbonDropDownItemImpl2.Label = "A3";
            ribbonDropDownItemImpl3.Label = "A1";
            ribbonDropDownItemImpl4.Label = "A2";
            ribbonDropDownItemImpl5.Label = "16:9";
            ribbonDropDownItemImpl6.Label = "4:3";
            ribbonDropDownItemImpl7.Label = "公众号封面";
            ribbonDropDownItemImpl8.Label = "小红书图文";
            this.comboBox1.Items.Add(ribbonDropDownItemImpl1);
            this.comboBox1.Items.Add(ribbonDropDownItemImpl2);
            this.comboBox1.Items.Add(ribbonDropDownItemImpl3);
            this.comboBox1.Items.Add(ribbonDropDownItemImpl4);
            this.comboBox1.Items.Add(ribbonDropDownItemImpl5);
            this.comboBox1.Items.Add(ribbonDropDownItemImpl6);
            this.comboBox1.Items.Add(ribbonDropDownItemImpl7);
            this.comboBox1.Items.Add(ribbonDropDownItemImpl8);
            this.comboBox1.Label = "页面尺寸";
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.ShowImage = true;
            this.comboBox1.Text = null;
            this.comboBox1.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.comboBox1_TextChanged);
            // 
            // comboBox2
            // 
            this.comboBox2.Image = ((System.Drawing.Image)(resources.GetObject("comboBox2.Image")));
            ribbonDropDownItemImpl9.Label = "纵向";
            ribbonDropDownItemImpl10.Label = "横向";
            this.comboBox2.Items.Add(ribbonDropDownItemImpl9);
            this.comboBox2.Items.Add(ribbonDropDownItemImpl10);
            this.comboBox2.Label = "页面方向";
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.ShowImage = true;
            this.comboBox2.Text = null;
            this.comboBox2.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.comboBox2_TextChanged);
            // 
            // 关于我
            // 
            this.关于我.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.关于我.Image = ((System.Drawing.Image)(resources.GetObject("关于我.Image")));
            this.关于我.Items.Add(this.检查更新);
            this.关于我.Label = "快捷魔方";
            this.关于我.Name = "关于我";
            this.关于我.SuperTip = "点击进入Andy老师的博客";
            this.关于我.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.关于我_Click);
            // 
            // 检查更新
            // 
            this.检查更新.Image = ((System.Drawing.Image)(resources.GetObject("检查更新.Image")));
            this.检查更新.Label = "检查更新";
            this.检查更新.Name = "检查更新";
            this.检查更新.ShowImage = true;
            this.检查更新.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.检查更新_Click);
            // 
            // toggleTaskPaneButton
            // 
            this.toggleTaskPaneButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toggleTaskPaneButton.Image = ((System.Drawing.Image)(resources.GetObject("toggleTaskPaneButton.Image")));
            this.toggleTaskPaneButton.Label = "学科工具";
            this.toggleTaskPaneButton.Name = "toggleTaskPaneButton";
            this.toggleTaskPaneButton.ShowImage = true;
            this.toggleTaskPaneButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleTaskPane_Click);
            // 
            // button5
            // 
            this.button5.Image = ((System.Drawing.Image)(resources.GetObject("button5.Image")));
            this.button5.Label = "云朵字形";
            this.button5.Name = "button5";
            this.button5.ScreenTip = "使用说明：";
            this.button5.ShowImage = true;
            this.button5.SuperTip = "支持自定义参数，一键生成云朵字（多重描边艺术字），支持参数动态调节更新。";
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click_1);
            // 
            // 笔顺图解
            // 
            this.笔顺图解.Image = ((System.Drawing.Image)(resources.GetObject("笔顺图解.Image")));
            this.笔顺图解.Label = "查询笔顺";
            this.笔顺图解.Name = "笔顺图解";
            this.笔顺图解.ScreenTip = "使用说明：";
            this.笔顺图解.ShowImage = true;
            this.笔顺图解.SuperTip = "输入单个汉字，查询和获取对应的SVG分解笔顺图。";
            this.笔顺图解.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.笔顺图解_Click);
            // 
            // 生字赋格
            // 
            this.生字赋格.Image = ((System.Drawing.Image)(resources.GetObject("生字赋格.Image")));
            this.生字赋格.Label = "生字赋格";
            this.生字赋格.Name = "生字赋格";
            this.生字赋格.ScreenTip = "使用说明：";
            this.生字赋格.ShowImage = true;
            this.生字赋格.SuperTip = "选中一个或多个对象，为其添加田字格。默认按照行列排列对齐。按住Ctrl键单击“生成”可强制原位添加田字格。";
            this.生字赋格.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.生字赋格_Click);
            // 
            // 注音编辑
            // 
            this.注音编辑.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.注音编辑.Image = ((System.Drawing.Image)(resources.GetObject("注音编辑.Image")));
            this.注音编辑.Items.Add(this.文本居中);
            this.注音编辑.Items.Add(this.删列补行);
            this.注音编辑.Items.Add(this.增加行宽);
            this.注音编辑.Items.Add(this.奇数行行高设置);
            this.注音编辑.Items.Add(this.自动补齐);
            this.注音编辑.Items.Add(this.合并段落);
            this.注音编辑.Items.Add(this.字号调整);
            this.注音编辑.Items.Add(this.晓声通在线注音);
            this.注音编辑.Label = "注音编辑";
            this.注音编辑.Name = "注音编辑";
            this.注音编辑.ScreenTip = "使用说明：";
            this.注音编辑.SuperTip = "单击进入“注音编辑器”。";
            this.注音编辑.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.注音编辑_Click);
            // 
            // 文本居中
            // 
            this.文本居中.Image = ((System.Drawing.Image)(resources.GetObject("文本居中.Image")));
            this.文本居中.Label = "行内居中";
            this.文本居中.Name = "文本居中";
            this.文本居中.ScreenTip = "使用说明：";
            this.文本居中.ShowImage = true;
            this.文本居中.SuperTip = "选中表格，单击可使得表格行内文本居中。";
            this.文本居中.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.文本居中_Click);
            // 
            // 删列补行
            // 
            this.删列补行.Image = ((System.Drawing.Image)(resources.GetObject("删列补行.Image")));
            this.删列补行.Label = "缩短行宽";
            this.删列补行.Name = "删列补行";
            this.删列补行.ScreenTip = "使用说明：";
            this.删列补行.ShowImage = true;
            this.删列补行.SuperTip = "默认单击，每行下移缩短一个中文字符数；按住Ctrl键单击可指定每行下移缩短若干字符数。";
            this.删列补行.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.删列补行_Click);
            // 
            // 增加行宽
            // 
            this.增加行宽.Image = ((System.Drawing.Image)(resources.GetObject("增加行宽.Image")));
            this.增加行宽.Label = "增加行宽";
            this.增加行宽.Name = "增加行宽";
            this.增加行宽.ScreenTip = "使用说明：";
            this.增加行宽.ShowImage = true;
            this.增加行宽.SuperTip = "默认单击，每行上移增加一个字符数；按住Ctrl键单击可指定每行上移增加若干字符数。";
            this.增加行宽.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.增加行宽_Click);
            // 
            // 奇数行行高设置
            // 
            this.奇数行行高设置.Image = ((System.Drawing.Image)(resources.GetObject("奇数行行高设置.Image")));
            this.奇数行行高设置.Label = "文本行距";
            this.奇数行行高设置.Name = "奇数行行高设置";
            this.奇数行行高设置.ShowImage = true;
            this.奇数行行高设置.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.奇数行行高设置_Click);
            // 
            // 自动补齐
            // 
            this.自动补齐.Image = ((System.Drawing.Image)(resources.GetObject("自动补齐.Image")));
            this.自动补齐.Label = "自动补齐";
            this.自动补齐.Name = "自动补齐";
            this.自动补齐.ScreenTip = "使用说明：";
            this.自动补齐.ShowImage = true;
            this.自动补齐.SuperTip = "若表格行开头有空白格，选中该行，可使后续内容补齐行开头空白格。";
            this.自动补齐.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.自动补齐_Click);
            // 
            // 合并段落
            // 
            this.合并段落.Image = ((System.Drawing.Image)(resources.GetObject("合并段落.Image")));
            this.合并段落.Label = "合并段落";
            this.合并段落.Name = "合并段落";
            this.合并段落.ScreenTip = "使用说明：";
            this.合并段落.ShowImage = true;
            this.合并段落.SuperTip = "选中两个表格（一个段落一个表格），可合并两者。";
            this.合并段落.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.合并段落_Click);
            // 
            // 字号调整
            // 
            this.字号调整.Image = ((System.Drawing.Image)(resources.GetObject("字号调整.Image")));
            this.字号调整.Label = "字号调整";
            this.字号调整.Name = "字号调整";
            this.字号调整.ScreenTip = "使用说明：";
            this.字号调整.ShowImage = true;
            this.字号调整.SuperTip = "选择注音布局表格，使用本功能可以分别设置拼音和汉字的字号大小。";
            this.字号调整.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.字号调整_Click);
            // 
            // 晓声通在线注音
            // 
            this.晓声通在线注音.Image = ((System.Drawing.Image)(resources.GetObject("晓声通在线注音.Image")));
            this.晓声通在线注音.Items.Add(this.在线注音编辑器);
            this.晓声通在线注音.Items.Add(this.文转表格);
            this.晓声通在线注音.Items.Add(this.重设表格);
            this.晓声通在线注音.Label = "友情链接";
            this.晓声通在线注音.Name = "晓声通在线注音";
            // 
            // 在线注音编辑器
            // 
            this.在线注音编辑器.Image = ((System.Drawing.Image)(resources.GetObject("在线注音编辑器.Image")));
            this.在线注音编辑器.Label = "晓声通在线注音编辑器";
            this.在线注音编辑器.Name = "在线注音编辑器";
            this.在线注音编辑器.ShowImage = true;
            this.在线注音编辑器.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.在线注音编辑器_Click);
            // 
            // 文转表格
            // 
            this.文转表格.Image = ((System.Drawing.Image)(resources.GetObject("文转表格.Image")));
            this.文转表格.Label = "粘贴表格（晓声通专用）";
            this.文转表格.Name = "文转表格";
            this.文转表格.ScreenTip = "使用说明：";
            this.文转表格.ShowImage = true;
            this.文转表格.SuperTip = "专为“晓声通编辑器”而设的表格导入功能。";
            this.文转表格.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.文转表格_Click);
            // 
            // 重设表格
            // 
            this.重设表格.Image = ((System.Drawing.Image)(resources.GetObject("重设表格.Image")));
            this.重设表格.Label = "重设格式（晓声通专用）";
            this.重设表格.Name = "重设表格";
            this.重设表格.ScreenTip = "使用说明：";
            this.重设表格.ShowImage = true;
            this.重设表格.SuperTip = "若直接从晓声通粘贴HTML表格到PPT中，可选中表格，可按Andy预设格式重设表格。";
            this.重设表格.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.重设表格_Click);
            // 
            // 常用格子
            // 
            this.常用格子.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.常用格子.Image = ((System.Drawing.Image)(resources.GetObject("常用格子.Image")));
            this.常用格子.Items.Add(this.生字格子);
            this.常用格子.Items.Add(this.四线三格);
            this.常用格子.Label = "常用格子";
            this.常用格子.Name = "常用格子";
            // 
            // 生字格子
            // 
            this.生字格子.Image = ((System.Drawing.Image)(resources.GetObject("生字格子.Image")));
            this.生字格子.Label = "田字格";
            this.生字格子.Name = "生字格子";
            this.生字格子.ScreenTip = "使用说明：";
            this.生字格子.ShowImage = true;
            this.生字格子.SuperTip = "用于创建矩阵田字格。";
            this.生字格子.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.生字格子_Click);
            // 
            // 四线三格
            // 
            this.四线三格.Image = ((System.Drawing.Image)(resources.GetObject("四线三格.Image")));
            this.四线三格.Label = "四线三格";
            this.四线三格.Name = "四线三格";
            this.四线三格.ShowImage = true;
            this.四线三格.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.四线三格_Click);
            // 
            // 注音工具
            // 
            this.注音工具.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.注音工具.Image = ((System.Drawing.Image)(resources.GetObject("注音工具.Image")));
            this.注音工具.Items.Add(this.button3);
            this.注音工具.Items.Add(this.便捷注音);
            this.注音工具.Items.Add(this.一键注音);
            this.注音工具.Label = "注音工具";
            this.注音工具.Name = "注音工具";
            // 
            // button3
            // 
            this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
            this.button3.Label = "拼音转换";
            this.button3.Name = "button3";
            this.button3.ScreenTip = "使用说明：";
            this.button3.ShowImage = true;
            this.button3.SuperTip = "选中带有汉字的文本框，一键转换并在其顶部插入对应的无声调拼音。（源：NPinyin）";
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // 便捷注音
            // 
            this.便捷注音.Image = ((System.Drawing.Image)(resources.GetObject("便捷注音.Image")));
            this.便捷注音.Label = "便捷注音";
            this.便捷注音.Name = "便捷注音";
            this.便捷注音.ScreenTip = "便捷注音";
            this.便捷注音.ShowImage = true;
            this.便捷注音.SuperTip = "选中无声调拼音，单击“便捷注音”，自动匹配相应的四个声调，选择正确声调的拼音进行注音即可。（源：Andy拼音库）";
            this.便捷注音.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.便捷注音_Click);
            // 
            // 一键注音
            // 
            this.一键注音.Image = ((System.Drawing.Image)(resources.GetObject("一键注音.Image")));
            this.一键注音.Label = "字典注音";
            this.一键注音.Name = "一键注音";
            this.一键注音.ScreenTip = "使用说明：";
            this.一键注音.ShowImage = true;
            this.一键注音.SuperTip = "选中文本框，则在文本框顶部注音；选中文本框内的文本，则在所选字符顶部注音。";
            this.一键注音.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.一键注音_Click);
            // 
            // 拓展应用
            // 
            this.拓展应用.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.拓展应用.Image = ((System.Drawing.Image)(resources.GetObject("拓展应用.Image")));
            this.拓展应用.Items.Add(this.Zici);
            this.拓展应用.Items.Add(this.WritePinyin);
            this.拓展应用.Items.Add(this.多音字词填空);
            this.拓展应用.Items.Add(this.分解拼音);
            this.拓展应用.Items.Add(this.拼音升调);
            this.拓展应用.Label = "拓展应用";
            this.拓展应用.Name = "拓展应用";
            // 
            // Zici
            // 
            this.Zici.Image = ((System.Drawing.Image)(resources.GetObject("Zici.Image")));
            this.Zici.Label = "看拼音写词语";
            this.Zici.Name = "Zici";
            this.Zici.ScreenTip = "使用说明：";
            this.Zici.ShowImage = true;
            this.Zici.SuperTip = "选中多个文本框（词语），使用本功能可一键创建“看拼音写词语”题目。";
            this.Zici.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Zici_Click);
            // 
            // WritePinyin
            // 
            this.WritePinyin.Image = ((System.Drawing.Image)(resources.GetObject("WritePinyin.Image")));
            this.WritePinyin.Label = "看词语写拼音";
            this.WritePinyin.Name = "WritePinyin";
            this.WritePinyin.ScreenTip = "使用说明：";
            this.WritePinyin.ShowImage = true;
            this.WritePinyin.SuperTip = "选中多个文本框（词语），使用本功能可一键创建“看词语写拼音”题目。";
            this.WritePinyin.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.WritePinyin_Click);
            // 
            // 多音字词填空
            // 
            this.多音字词填空.Image = ((System.Drawing.Image)(resources.GetObject("多音字词填空.Image")));
            this.多音字词填空.Label = "多音字词语填空";
            this.多音字词填空.Name = "多音字词填空";
            this.多音字词填空.ScreenTip = "使用说明：";
            this.多音字词填空.ShowImage = true;
            this.多音字词填空.SuperTip = "选中若干多音字（确保每个文本框只有一个多音字），单击创建多音字填空题";
            this.多音字词填空.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.多音字词填空_Click);
            // 
            // 分解拼音
            // 
            this.分解拼音.Image = ((System.Drawing.Image)(resources.GetObject("分解拼音.Image")));
            this.分解拼音.Label = "分解拼音";
            this.分解拼音.Name = "分解拼音";
            this.分解拼音.ScreenTip = "使用说明：";
            this.分解拼音.ShowImage = true;
            this.分解拼音.SuperTip = "选中拼音文本框，单击可分解其声母和韵母";
            this.分解拼音.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.分解拼音_Click);
            // 
            // 拼音升调
            // 
            this.拼音升调.Image = ((System.Drawing.Image)(resources.GetObject("拼音升调.Image")));
            this.拼音升调.Label = "拼音升调";
            this.拼音升调.Name = "拼音升调";
            this.拼音升调.ShowImage = true;
            this.拼音升调.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.拼音升调_Click);
            // 
            // Masking
            // 
            this.Masking.Image = ((System.Drawing.Image)(resources.GetObject("Masking.Image")));
            this.Masking.Label = "图片透明";
            this.Masking.Name = "Masking";
            this.Masking.ScreenTip = "使用说明：";
            this.Masking.ShowImage = true;
            this.Masking.SuperTip = "除了的基础的透明化以外，还支持不同方向的图片渐变透明处理，且支持叠加颜色，转换灰度图。";
            this.Masking.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Masking_Click_1);
            // 
            // 图形分割
            // 
            this.图形分割.Image = ((System.Drawing.Image)(resources.GetObject("图形分割.Image")));
            this.图形分割.Label = "图形分割";
            this.图形分割.Name = "图形分割";
            this.图形分割.ShowImage = true;
            this.图形分割.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.图形分割_Click);
            // 
            // button6
            // 
            this.button6.Image = ((System.Drawing.Image)(resources.GetObject("button6.Image")));
            this.button6.Label = "SVG粘贴";
            this.button6.Name = "button6";
            this.button6.ScreenTip = "使用说明：";
            this.button6.ShowImage = true;
            this.button6.SuperTip = "复制一段SVG代码到剪切板，将代码转换成SVG矢量图插入当前页幻灯片。";
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button6_Click);
            // 
            // button1
            // 
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "适应全屏";
            this.button1.Name = "button1";
            this.button1.ScreenTip = "使用说明：";
            this.button1.ShowImage = true;
            this.button1.SuperTip = "选中一张图片或形状，放大至全屏。";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // Gradientrectangle
            // 
            this.Gradientrectangle.Image = ((System.Drawing.Image)(resources.GetObject("Gradientrectangle.Image")));
            this.Gradientrectangle.Label = "渐变蒙版";
            this.Gradientrectangle.Name = "Gradientrectangle";
            this.Gradientrectangle.ScreenTip = "使用说明：";
            this.Gradientrectangle.ShowImage = true;
            this.Gradientrectangle.SuperTip = "1.在没有选中任何对象时，则一键添加与幻灯片等大的黑色透明渐变蒙版；2.如果选中了一个或多个对象时，则在该对象顶层插入一块与它等大的渐变蒙版。3.默认渐变角度为0" +
    "°，按住Ctrl键单击则插入的蒙版渐变角度为90°；按住Shfit键单间则插入的蒙版渐变角度为45°。";
            this.Gradientrectangle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Gradientrectangle_Click);
            // 
            // 在线工具
            // 
            this.在线工具.Image = ((System.Drawing.Image)(resources.GetObject("在线工具.Image")));
            this.在线工具.Items.Add(this.抠图);
            this.在线工具.Items.Add(this.矢量);
            this.在线工具.Label = "在线工具";
            this.在线工具.Name = "在线工具";
            // 
            // 抠图
            // 
            this.抠图.Image = ((System.Drawing.Image)(resources.GetObject("抠图.Image")));
            this.抠图.Items.Add(this.button19);
            this.抠图.Items.Add(this.趣作图);
            this.抠图.Items.Add(this.Bgsub);
            this.抠图.Label = "抠图";
            this.抠图.Name = "抠图";
            this.抠图.ScreenTip = "在线抠图";
            // 
            // button19
            // 
            this.button19.Image = ((System.Drawing.Image)(resources.GetObject("button19.Image")));
            this.button19.Label = "AI抠图";
            this.button19.Name = "button19";
            this.button19.ScreenTip = "使用说明：";
            this.button19.ShowImage = true;
            this.button19.SuperTip = "该在线抠图工具对于抠图处理更加精细，开发作者为B站UP主@设计学姐。如有更精细的抠图处理需求，可前往这个工具网站进行处理。";
            this.button19.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button19_Click);
            // 
            // 趣作图
            // 
            this.趣作图.Image = ((System.Drawing.Image)(resources.GetObject("趣作图.Image")));
            this.趣作图.Label = "趣作图";
            this.趣作图.Name = "趣作图";
            this.趣作图.ShowImage = true;
            this.趣作图.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Experte抠图_Click);
            // 
            // Bgsub
            // 
            this.Bgsub.Image = ((System.Drawing.Image)(resources.GetObject("Bgsub.Image")));
            this.Bgsub.Label = "Bgsub";
            this.Bgsub.Name = "Bgsub";
            this.Bgsub.ShowImage = true;
            this.Bgsub.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Bgsub_Click);
            // 
            // 矢量
            // 
            this.矢量.Image = ((System.Drawing.Image)(resources.GetObject("矢量.Image")));
            this.矢量.Items.Add(this.Tmttool);
            this.矢量.Label = "矢量";
            this.矢量.Name = "矢量";
            this.矢量.ScreenTip = "位图转矢量图";
            // 
            // Tmttool
            // 
            this.Tmttool.Image = ((System.Drawing.Image)(resources.GetObject("Tmttool.Image")));
            this.Tmttool.Label = "Tmttool";
            this.Tmttool.Name = "Tmttool";
            this.Tmttool.ShowImage = true;
            this.Tmttool.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.位图转矢量图_Click);
            // 
            // 矩形拆分
            // 
            this.矩形拆分.Label = "";
            this.矩形拆分.Name = "矩形拆分";
            // 
            // Mosaic
            // 
            this.Mosaic.Label = "";
            this.Mosaic.Name = "Mosaic";
            // 
            // ApplyFilter
            // 
            this.ApplyFilter.Label = "";
            this.ApplyFilter.Name = "ApplyFilter";
            // 
            // Expandimage
            // 
            this.Expandimage.Label = "";
            this.Expandimage.Name = "Expandimage";
            // 
            // 平移居中
            // 
            this.平移居中.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.平移居中.Image = ((System.Drawing.Image)(resources.GetObject("平移居中.Image")));
            this.平移居中.Label = "对齐增强";
            this.平移居中.Name = "平移居中";
            this.平移居中.ScreenTip = "使用说明：";
            this.平移居中.ShowImage = true;
            this.平移居中.SuperTip = "以第一个被选中的对象为参考（基准），固定其位置不变，同时将后续所选的其他对象都看作一个整体（无论数量的多少），按照既定的对齐方式对齐到参考对象中。";
            this.平移居中.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.平移居中_Click);
            // 
            // 分布
            // 
            this.分布.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.分布.Image = ((System.Drawing.Image)(resources.GetObject("分布.Image")));
            this.分布.Items.Add(this.沿线分布);
            this.分布.Items.Add(this.矩阵分布);
            this.分布.Items.Add(this.环形分布);
            this.分布.Items.Add(this.贴边对齐);
            this.分布.Items.Add(this.选择居中);
            this.分布.Label = "分布";
            this.分布.Name = "分布";
            this.分布.ShowImage = true;
            // 
            // 沿线分布
            // 
            this.沿线分布.Image = ((System.Drawing.Image)(resources.GetObject("沿线分布.Image")));
            this.沿线分布.Label = "沿线分布";
            this.沿线分布.Name = "沿线分布";
            this.沿线分布.ScreenTip = "使用说明：";
            this.沿线分布.ShowImage = true;
            this.沿线分布.SuperTip = "先选中一个包含两个或两个顶点以上的自由线条，再选中要分布到线条上的其他对象，则可以使得其他对象沿线分布。";
            this.沿线分布.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.沿线分布_Click);
            // 
            // 矩阵分布
            // 
            this.矩阵分布.Image = ((System.Drawing.Image)(resources.GetObject("矩阵分布.Image")));
            this.矩阵分布.Label = "矩阵分布";
            this.矩阵分布.Name = "矩阵分布";
            this.矩阵分布.ScreenTip = "使用说明：";
            this.矩阵分布.ShowImage = true;
            this.矩阵分布.SuperTip = "使所选对象按照矩阵排列分布。";
            this.矩阵分布.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.矩阵分布_Click);
            // 
            // 环形分布
            // 
            this.环形分布.Image = ((System.Drawing.Image)(resources.GetObject("环形分布.Image")));
            this.环形分布.Label = "环形分布";
            this.环形分布.Name = "环形分布";
            this.环形分布.ScreenTip = "使用说明：";
            this.环形分布.ShowImage = true;
            this.环形分布.SuperTip = "使所选对象按照环形排列分布。";
            this.环形分布.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.环形分布_Click);
            // 
            // 贴边对齐
            // 
            this.贴边对齐.Image = ((System.Drawing.Image)(resources.GetObject("贴边对齐.Image")));
            this.贴边对齐.Items.Add(this.button11);
            this.贴边对齐.Items.Add(this.button13);
            this.贴边对齐.Label = "贴边对齐";
            this.贴边对齐.Name = "贴边对齐";
            // 
            // button11
            // 
            this.button11.Image = ((System.Drawing.Image)(resources.GetObject("button11.Image")));
            this.button11.Label = "左右贴边";
            this.button11.Name = "button11";
            this.button11.ScreenTip = "使用说明：";
            this.button11.ShowImage = true;
            this.button11.SuperTip = "选中多个对象，以第一个被选中的对象为基准，其他对象从左到右无缝贴边对齐。";
            this.button11.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button11_Click_1);
            // 
            // button13
            // 
            this.button13.Image = ((System.Drawing.Image)(resources.GetObject("button13.Image")));
            this.button13.Label = "上下贴边";
            this.button13.Name = "button13";
            this.button13.ScreenTip = "使用说明：";
            this.button13.ShowImage = true;
            this.button13.SuperTip = "选中多个对象，以第一个被选中的对象为基准，其他所有对象从上往下无缝贴边对齐。";
            this.button13.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button13_Click);
            // 
            // 选择居中
            // 
            this.选择居中.Image = ((System.Drawing.Image)(resources.GetObject("选择居中.Image")));
            this.选择居中.Items.Add(this.button10);
            this.选择居中.Items.Add(this.Pagecentered);
            this.选择居中.Label = "选择居中";
            this.选择居中.Name = "选择居中";
            // 
            // button10
            // 
            this.button10.Image = ((System.Drawing.Image)(resources.GetObject("button10.Image")));
            this.button10.Label = "参考居中";
            this.button10.Name = "button10";
            this.button10.ScreenTip = "使用说明：";
            this.button10.ShowImage = true;
            this.button10.SuperTip = "选中两个对象，以第一个选中的对象为对齐基准（固定位置不变），两者居中对齐。";
            this.button10.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button10_Click);
            // 
            // Pagecentered
            // 
            this.Pagecentered.Image = ((System.Drawing.Image)(resources.GetObject("Pagecentered.Image")));
            this.Pagecentered.Label = "页内居中";
            this.Pagecentered.Name = "Pagecentered";
            this.Pagecentered.ScreenTip = "使用说明：";
            this.Pagecentered.ShowImage = true;
            this.Pagecentered.SuperTip = "默认单击，将所选对象整体平移到页面中心。按住Ctrl键单击，则将所选对象整体平移至水平线中部。";
            this.Pagecentered.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Pagecentered_Click);
            // 
            // 筛选
            // 
            this.筛选.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.筛选.Image = ((System.Drawing.Image)(resources.GetObject("筛选.Image")));
            this.筛选.Items.Add(this.Type);
            this.筛选.Items.Add(this.Selectsize);
            this.筛选.Items.Add(this.SelectedColor);
            this.筛选.Items.Add(this.Selectedline);
            this.筛选.Items.Add(this.Selectfontsize);
            this.筛选.Items.Add(this.图层);
            this.筛选.Label = "筛选";
            this.筛选.Name = "筛选";
            // 
            // Type
            // 
            this.Type.Image = ((System.Drawing.Image)(resources.GetObject("Type.Image")));
            this.Type.Label = "类型";
            this.Type.Name = "Type";
            this.Type.ScreenTip = "使用说明";
            this.Type.ShowImage = true;
            this.Type.SuperTip = "选中一个对象后单击“类型”，可同时选中当前幻灯片与它同类型的所有对象。";
            this.Type.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Type_Click);
            // 
            // Selectsize
            // 
            this.Selectsize.Image = ((System.Drawing.Image)(resources.GetObject("Selectsize.Image")));
            this.Selectsize.Label = "尺寸";
            this.Selectsize.Name = "Selectsize";
            this.Selectsize.ScreenTip = "使用说明：";
            this.Selectsize.ShowImage = true;
            this.Selectsize.SuperTip = "选中一个对象，单击“尺寸”，可同时选中当前幻灯片与它同尺寸（大小）的所有对象。";
            this.Selectsize.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Selectsize_Click);
            // 
            // SelectedColor
            // 
            this.SelectedColor.Image = ((System.Drawing.Image)(resources.GetObject("SelectedColor.Image")));
            this.SelectedColor.Label = "颜色";
            this.SelectedColor.Name = "SelectedColor";
            this.SelectedColor.ScreenTip = "使用说明：";
            this.SelectedColor.ShowImage = true;
            this.SelectedColor.SuperTip = "选中一个对象，单击“颜色”，可同时选中当前幻灯片与它相同颜色的所有对象。";
            this.SelectedColor.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SelectedColor_Click);
            // 
            // Selectedline
            // 
            this.Selectedline.Image = ((System.Drawing.Image)(resources.GetObject("Selectedline.Image")));
            this.Selectedline.Label = "线条";
            this.Selectedline.Name = "Selectedline";
            this.Selectedline.ScreenTip = "使用说明：";
            this.Selectedline.ShowImage = true;
            this.Selectedline.SuperTip = "选中一个对象：1.单击，则同时选中与他线条颜色的所有对象；2.按住Ctrl键单击，则同时选中当前幻灯片线条宽度与它相同的所有对象；3.按住Shift键单击，则同时" +
    "选中与它相同线条类型（如虚线、实线）的所有对象。";
            this.Selectedline.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Selectedline_Click);
            // 
            // Selectfontsize
            // 
            this.Selectfontsize.Image = ((System.Drawing.Image)(resources.GetObject("Selectfontsize.Image")));
            this.Selectfontsize.Label = "字号";
            this.Selectfontsize.Name = "Selectfontsize";
            this.Selectfontsize.ScreenTip = "使用说明：";
            this.Selectfontsize.ShowImage = true;
            this.Selectfontsize.SuperTip = "选中一个文本框或带有文字的形状，可同时选中字号相同的所有对象。";
            this.Selectfontsize.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Selectfontsize_Click);
            // 
            // 图层
            // 
            this.图层.Image = ((System.Drawing.Image)(resources.GetObject("图层.Image")));
            this.图层.Label = "图层";
            this.图层.Name = "图层";
            this.图层.ScreenTip = "使用说明：";
            this.图层.ShowImage = true;
            this.图层.SuperTip = "选中一个对象，单击同时选中与其图层前缀名相同的对象（不支持组合内筛选子对象）。";
            this.图层.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.图层_Click);
            // 
            // 选择增强
            // 
            this.选择增强.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.选择增强.Image = ((System.Drawing.Image)(resources.GetObject("选择增强.Image")));
            this.选择增强.Label = "选择增强";
            this.选择增强.Name = "选择增强";
            this.选择增强.ScreenTip = "使用说明";
            this.选择增强.ShowImage = true;
            this.选择增强.SuperTip = "在开启“选择增强”后，程序将主动按照你的选择顺序记录所有被选中的对象，第二次点击即可关闭记录，同时按顺序全选所有对象。";
            this.选择增强.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.选择增强_Click);
            // 
            // 智能缩放
            // 
            this.智能缩放.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.智能缩放.Image = ((System.Drawing.Image)(resources.GetObject("智能缩放.Image")));
            this.智能缩放.Label = "智能缩放";
            this.智能缩放.Name = "智能缩放";
            this.智能缩放.ScreenTip = "使用说明：";
            this.智能缩放.ShowImage = true;
            this.智能缩放.SuperTip = "可以对所选对象大小和属性进行等比缩放，支持更改缩放中心。";
            this.智能缩放.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.智能缩放_Click);
            // 
            // 文本
            // 
            this.文本.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.文本.Image = ((System.Drawing.Image)(resources.GetObject("文本.Image")));
            this.文本.Items.Add(this.去除边距);
            this.文本.Items.Add(this.首行缩进);
            this.文本.Items.Add(this.单字拆分);
            this.文本.Items.Add(this.拆分段落);
            this.文本.Items.Add(this.批量改字);
            this.文本.Items.Add(this.字词加点);
            this.文本.Items.Add(this.文本矢量化);
            this.文本.Items.Add(this.文写入形);
            this.文本.Items.Add(this.部首描红);
            this.文本.Items.Add(this.分解笔顺);
            this.文本.Label = "文本";
            this.文本.Name = "文本";
            this.文本.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.splitButton2_Click);
            // 
            // 去除边距
            // 
            this.去除边距.Image = ((System.Drawing.Image)(resources.GetObject("去除边距.Image")));
            this.去除边距.Label = "消除边距";
            this.去除边距.Name = "去除边距";
            this.去除边距.ScreenTip = "使用说明：";
            this.去除边距.ShowImage = true;
            this.去除边距.SuperTip = "选中文本框，单击一键消除文本框左右上下边距。";
            this.去除边距.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.去除边距_Click);
            // 
            // 首行缩进
            // 
            this.首行缩进.Image = ((System.Drawing.Image)(resources.GetObject("首行缩进.Image")));
            this.首行缩进.Label = "首行缩进";
            this.首行缩进.Name = "首行缩进";
            this.首行缩进.ScreenTip = "使用说明：";
            this.首行缩进.ShowImage = true;
            this.首行缩进.SuperTip = "选中文本框，单击缩进两个字符，再次单击取消缩进。";
            this.首行缩进.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.首行缩进_Click);
            // 
            // 单字拆分
            // 
            this.单字拆分.Image = ((System.Drawing.Image)(resources.GetObject("单字拆分.Image")));
            this.单字拆分.Label = "拆分单字";
            this.单字拆分.Name = "单字拆分";
            this.单字拆分.ScreenTip = "使用说明：";
            this.单字拆分.ShowImage = true;
            this.单字拆分.SuperTip = "选中文本框，单击将文本框内的字符进行逐个拆分。";
            this.单字拆分.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.单字拆分_Click);
            // 
            // 拆分段落
            // 
            this.拆分段落.Image = ((System.Drawing.Image)(resources.GetObject("拆分段落.Image")));
            this.拆分段落.Label = "拆分段落";
            this.拆分段落.Name = "拆分段落";
            this.拆分段落.ScreenTip = "使用说明：";
            this.拆分段落.ShowImage = true;
            this.拆分段落.SuperTip = "将多个段落的文本拆分成多个文本框，一个段落一个文本框。";
            this.拆分段落.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.拆分段落_Click);
            // 
            // 批量改字
            // 
            this.批量改字.Image = ((System.Drawing.Image)(resources.GetObject("批量改字.Image")));
            this.批量改字.Label = "批量改字";
            this.批量改字.Name = "批量改字";
            this.批量改字.ScreenTip = "使用说明：";
            this.批量改字.ShowImage = true;
            this.批量改字.SuperTip = "选中一个或多个文本框（形状），可对它们的文本进行批量修改。";
            this.批量改字.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.批量改字_Click);
            // 
            // 字词加点
            // 
            this.字词加点.Image = ((System.Drawing.Image)(resources.GetObject("字词加点.Image")));
            this.字词加点.Label = "字词加点";
            this.字词加点.Name = "字词加点";
            this.字词加点.ScreenTip = "使用说明：";
            this.字词加点.ShowImage = true;
            this.字词加点.SuperTip = "在文本框内选中需要需要加点的字词。";
            this.字词加点.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.字词加点_Click);
            // 
            // 文本矢量化
            // 
            this.文本矢量化.Image = ((System.Drawing.Image)(resources.GetObject("文本矢量化.Image")));
            this.文本矢量化.Label = "文本矢量";
            this.文本矢量化.Name = "文本矢量化";
            this.文本矢量化.ScreenTip = "使用说明：";
            this.文本矢量化.ShowImage = true;
            this.文本矢量化.SuperTip = "使用本功能，可将所选的文本转换为矢量形状。";
            this.文本矢量化.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.文本矢量化_Click);
            // 
            // 文写入形
            // 
            this.文写入形.Image = ((System.Drawing.Image)(resources.GetObject("文写入形.Image")));
            this.文写入形.Label = "文写入形";
            this.文写入形.Name = "文写入形";
            this.文写入形.ScreenTip = "使用说明：";
            this.文写入形.ShowImage = true;
            this.文写入形.SuperTip = "选中若干文本框，默认单击写入圆角矩形；按Ctrl键单击写入正圆形。";
            this.文写入形.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.文形互转_Click);
            // 
            // 部首描红
            // 
            this.部首描红.Image = ((System.Drawing.Image)(resources.GetObject("部首描红.Image")));
            this.部首描红.Label = "部首描红";
            this.部首描红.Name = "部首描红";
            this.部首描红.ScreenTip = "使用说明：";
            this.部首描红.ShowImage = true;
            this.部首描红.SuperTip = "请先使用学科工具中的”笔画拆分“对汉字笔画进行拆分，并将拆分出来的汉字笔画进行组合，然后选中该组合执行”部首描红“。";
            this.部首描红.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.部首描红_Click);
            // 
            // 分解笔顺
            // 
            this.分解笔顺.Image = ((System.Drawing.Image)(resources.GetObject("分解笔顺.Image")));
            this.分解笔顺.Label = "分解笔顺（旧版）";
            this.分解笔顺.Name = "分解笔顺";
            this.分解笔顺.ScreenTip = "使用说明：";
            this.分解笔顺.ShowImage = true;
            this.分解笔顺.SuperTip = "先把拆分出来的所有笔画进行组合，选中该组合，单击“分解笔顺”即可完成操作。";
            this.分解笔顺.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.分解笔顺_Click);
            // 
            // 绘图
            // 
            this.绘图.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.绘图.Image = ((System.Drawing.Image)(resources.GetObject("绘图.Image")));
            this.绘图.Items.Add(this.辐射连线);
            this.绘图.Items.Add(this.层级关系);
            this.绘图.Items.Add(this.括弧关系);
            this.绘图.Label = "绘图";
            this.绘图.Name = "绘图";
            // 
            // 辐射连线
            // 
            this.辐射连线.Image = ((System.Drawing.Image)(resources.GetObject("辐射连线.Image")));
            this.辐射连线.Label = "辐射（扩散）";
            this.辐射连线.Name = "辐射连线";
            this.辐射连线.ShowImage = true;
            this.辐射连线.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.辐射连线_Click);
            // 
            // 层级关系
            // 
            this.层级关系.Image = ((System.Drawing.Image)(resources.GetObject("层级关系.Image")));
            this.层级关系.Label = "层级（层次）";
            this.层级关系.Name = "层级关系";
            this.层级关系.ShowImage = true;
            this.层级关系.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.层级关系_Click);
            // 
            // 括弧关系
            // 
            this.括弧关系.Image = ((System.Drawing.Image)(resources.GetObject("括弧关系.Image")));
            this.括弧关系.Label = "括弧（总分）";
            this.括弧关系.Name = "括弧关系";
            this.括弧关系.ShowImage = true;
            this.括弧关系.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.括弧关系_Click);
            // 
            // 更多便捷
            // 
            this.更多便捷.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.更多便捷.Image = ((System.Drawing.Image)(resources.GetObject("更多便捷.Image")));
            this.更多便捷.Items.Add(this.统一);
            this.更多便捷.Items.Add(this.交换);
            this.更多便捷.Items.Add(this.组合);
            this.更多便捷.Items.Add(this.复制);
            this.更多便捷.Items.Add(this.图片);
            this.更多便捷.Items.Add(this.删除);
            this.更多便捷.Label = "便捷";
            this.更多便捷.Name = "更多便捷";
            this.更多便捷.ShowImage = true;
            // 
            // 统一
            // 
            this.统一.Image = ((System.Drawing.Image)(resources.GetObject("统一.Image")));
            this.统一.Items.Add(this.统一大小);
            this.统一.Items.Add(this.统一格式);
            this.统一.Items.Add(this.统一控点);
            this.统一.Label = "统一";
            this.统一.Name = "统一";
            // 
            // 统一大小
            // 
            this.统一大小.Image = ((System.Drawing.Image)(resources.GetObject("统一大小.Image")));
            this.统一大小.Label = "统一大小";
            this.统一大小.Name = "统一大小";
            this.统一大小.ScreenTip = "使用说明：";
            this.统一大小.ShowImage = true;
            this.统一大小.SuperTip = "使用本功能，默认单击，将以第一个被选中的对象的大小为基准，统一所选对象的大小。按住Ctrl单击，则统一高度；按住Shift单击，则统一宽度。";
            this.统一大小.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.统一大小_Click);
            // 
            // 统一格式
            // 
            this.统一格式.Image = ((System.Drawing.Image)(resources.GetObject("统一格式.Image")));
            this.统一格式.Label = "统一格式";
            this.统一格式.Name = "统一格式";
            this.统一格式.ScreenTip = "使用说明：";
            this.统一格式.ShowImage = true;
            this.统一格式.SuperTip = "使用本功能，将以第一个被选中的对象的格式为基准，统一所选对象的格式。";
            this.统一格式.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.统一格式_Click);
            // 
            // 统一控点
            // 
            this.统一控点.Image = ((System.Drawing.Image)(resources.GetObject("统一控点.Image")));
            this.统一控点.Label = "统一控点";
            this.统一控点.Name = "统一控点";
            this.统一控点.ScreenTip = "使用说明：";
            this.统一控点.ShowImage = true;
            this.统一控点.SuperTip = "使用本功能，将以第一个被选中的对象的控点为基准（如果对象存在控点），统一所选对象的控点。";
            this.统一控点.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.统一控点_Click);
            // 
            // 交换
            // 
            this.交换.Image = ((System.Drawing.Image)(resources.GetObject("交换.Image")));
            this.交换.Items.Add(this.交换位置);
            this.交换.Items.Add(this.交换文字);
            this.交换.Items.Add(this.交换格式);
            this.交换.Items.Add(this.交换尺寸);
            this.交换.Label = "交换";
            this.交换.Name = "交换";
            // 
            // 交换位置
            // 
            this.交换位置.Image = ((System.Drawing.Image)(resources.GetObject("交换位置.Image")));
            this.交换位置.Label = "交换位置";
            this.交换位置.Name = "交换位置";
            this.交换位置.ScreenTip = "使用说明：";
            this.交换位置.ShowImage = true;
            this.交换位置.SuperTip = "使用本功能，可将选中的两个对象交换彼此的位置（包括图层顺序）。";
            this.交换位置.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.交换位置_Click);
            // 
            // 交换文字
            // 
            this.交换文字.Image = ((System.Drawing.Image)(resources.GetObject("交换文字.Image")));
            this.交换文字.Label = "交换文字";
            this.交换文字.Name = "交换文字";
            this.交换文字.ScreenTip = "使用说明：";
            this.交换文字.ShowImage = true;
            this.交换文字.SuperTip = "使用本功能，可将选中的两个文本框内的文字进行交换。";
            this.交换文字.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.交换文字_Click);
            // 
            // 交换格式
            // 
            this.交换格式.Image = ((System.Drawing.Image)(resources.GetObject("交换格式.Image")));
            this.交换格式.Label = "交换格式";
            this.交换格式.Name = "交换格式";
            this.交换格式.ScreenTip = "使用说明：";
            this.交换格式.ShowImage = true;
            this.交换格式.SuperTip = "使用本功能，可将选中的两个对象的格式进行交换。";
            this.交换格式.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.交换格式_Click);
            // 
            // 交换尺寸
            // 
            this.交换尺寸.Image = ((System.Drawing.Image)(resources.GetObject("交换尺寸.Image")));
            this.交换尺寸.Label = "交换尺寸";
            this.交换尺寸.Name = "交换尺寸";
            this.交换尺寸.ScreenTip = "使用说明：";
            this.交换尺寸.ShowImage = true;
            this.交换尺寸.SuperTip = "使用本功能，可将选中的两个对象的尺寸（大小）进行交换。";
            this.交换尺寸.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.交换尺寸_Click);
            // 
            // 组合
            // 
            this.组合.Image = ((System.Drawing.Image)(resources.GetObject("组合.Image")));
            this.组合.Items.Add(this.重叠组合);
            this.组合.Items.Add(this.临近组合);
            this.组合.Items.Add(this.同色组合);
            this.组合.Label = "组合";
            this.组合.Name = "组合";
            // 
            // 重叠组合
            // 
            this.重叠组合.Image = ((System.Drawing.Image)(resources.GetObject("重叠组合.Image")));
            this.重叠组合.Label = "重叠组合";
            this.重叠组合.Name = "重叠组合";
            this.重叠组合.ScreenTip = "使用说明";
            this.重叠组合.ShowImage = true;
            this.重叠组合.SuperTip = "选中多个对象，使用本功能可将重叠的对象进行组合。";
            this.重叠组合.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.重叠组合_Click);
            // 
            // 临近组合
            // 
            this.临近组合.Image = ((System.Drawing.Image)(resources.GetObject("临近组合.Image")));
            this.临近组合.Label = "临边组合";
            this.临近组合.Name = "临近组合";
            this.临近组合.ScreenTip = "使用说明：";
            this.临近组合.ShowImage = true;
            this.临近组合.SuperTip = "选中多个对象，使用本功能可将多个相邻的对象进行组合。";
            this.临近组合.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.临近组合_Click);
            // 
            // 同色组合
            // 
            this.同色组合.Image = ((System.Drawing.Image)(resources.GetObject("同色组合.Image")));
            this.同色组合.Label = "同色组合";
            this.同色组合.Name = "同色组合";
            this.同色组合.ScreenTip = "使用说明：";
            this.同色组合.ShowImage = true;
            this.同色组合.SuperTip = "选中多个对象，使用本功能可将相同填充颜色的对象进行组合。";
            this.同色组合.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.同色组合_Click);
            // 
            // 复制
            // 
            this.复制.Image = ((System.Drawing.Image)(resources.GetObject("复制.Image")));
            this.复制.Items.Add(this.左右镜像);
            this.复制.Items.Add(this.上下镜像);
            this.复制.Items.Add(this.LCopy);
            this.复制.Label = "复制";
            this.复制.Name = "复制";
            // 
            // 左右镜像
            // 
            this.左右镜像.Image = ((System.Drawing.Image)(resources.GetObject("左右镜像.Image")));
            this.左右镜像.Label = "左右镜像";
            this.左右镜像.Name = "左右镜像";
            this.左右镜像.ScreenTip = "使用说明：";
            this.左右镜像.ShowImage = true;
            this.左右镜像.SuperTip = "选中对象，使用本功能可左右镜像复制该对象。";
            this.左右镜像.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.左右镜像_Click);
            // 
            // 上下镜像
            // 
            this.上下镜像.Image = ((System.Drawing.Image)(resources.GetObject("上下镜像.Image")));
            this.上下镜像.Label = "上下镜像";
            this.上下镜像.Name = "上下镜像";
            this.上下镜像.ScreenTip = "使用说明：";
            this.上下镜像.ShowImage = true;
            this.上下镜像.SuperTip = "选中对象，使用本功能可左右镜像复制该对象。";
            this.上下镜像.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.上下镜像_Click);
            // 
            // LCopy
            // 
            this.LCopy.Image = ((System.Drawing.Image)(resources.GetObject("LCopy.Image")));
            this.LCopy.Label = "原位复制";
            this.LCopy.Name = "LCopy";
            this.LCopy.ScreenTip = "使用说明：";
            this.LCopy.ShowImage = true;
            this.LCopy.SuperTip = "选中对象，原位复制所选对象（默认一次）。";
            this.LCopy.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LCopy_Click);
            // 
            // 图片
            // 
            this.图片.Image = ((System.Drawing.Image)(resources.GetObject("图片.Image")));
            this.图片.Items.Add(this.批量换图);
            this.图片.Items.Add(this.形状填图);
            this.图片.Items.Add(this.原位转图);
            this.图片.Items.Add(this.原位转JPG);
            this.图片.Label = "图片";
            this.图片.Name = "图片";
            // 
            // 批量换图
            // 
            this.批量换图.Image = ((System.Drawing.Image)(resources.GetObject("批量换图.Image")));
            this.批量换图.Label = "批量换图";
            this.批量换图.Name = "批量换图";
            this.批量换图.ScreenTip = "使用说明：";
            this.批量换图.ShowImage = true;
            this.批量换图.SuperTip = "使用本功能可实现批量换图，且自适应原图大小和格式。";
            this.批量换图.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.批量换图_Click);
            // 
            // 形状填图
            // 
            this.形状填图.Image = ((System.Drawing.Image)(resources.GetObject("形状填图.Image")));
            this.形状填图.Label = "形状填图";
            this.形状填图.Name = "形状填图";
            this.形状填图.ScreenTip = "使用说明：";
            this.形状填图.ShowImage = true;
            this.形状填图.SuperTip = "使用本功能，可将图片批量填充到形状中";
            this.形状填图.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.形状填图_Click);
            // 
            // 原位转图
            // 
            this.原位转图.Image = ((System.Drawing.Image)(resources.GetObject("原位转图.Image")));
            this.原位转图.Label = "原位转PNG";
            this.原位转图.Name = "原位转图";
            this.原位转图.ScreenTip = "使用说明：";
            this.原位转图.ShowImage = true;
            this.原位转图.SuperTip = "使用本功能，可将选中的对象原位转换成PNG图片。";
            this.原位转图.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.原位转图_Click);
            // 
            // 原位转JPG
            // 
            this.原位转JPG.Image = ((System.Drawing.Image)(resources.GetObject("原位转JPG.Image")));
            this.原位转JPG.Label = "原位转JPG";
            this.原位转JPG.Name = "原位转JPG";
            this.原位转JPG.ScreenTip = "使用说明：";
            this.原位转JPG.ShowImage = true;
            this.原位转JPG.SuperTip = "使用本功能，可将选中的对象原位转换成JPG图片。";
            this.原位转JPG.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.原位转JPG_Click);
            // 
            // 删除
            // 
            this.删除.Image = ((System.Drawing.Image)(resources.GetObject("删除.Image")));
            this.删除.Items.Add(this.删除动画);
            this.删除.Items.Add(this.清空页外);
            this.删除.Items.Add(this.清除备注);
            this.删除.Items.Add(this.清除超链接);
            this.删除.Items.Add(this.删除未用版式);
            this.删除.Label = "删除";
            this.删除.Name = "删除";
            // 
            // 删除动画
            // 
            this.删除动画.Image = ((System.Drawing.Image)(resources.GetObject("删除动画.Image")));
            this.删除动画.Label = "删除动画";
            this.删除动画.Name = "删除动画";
            this.删除动画.ScreenTip = "使用说明：";
            this.删除动画.ShowImage = true;
            this.删除动画.SuperTip = "选中一个或多个对象，可删除所选对象的动画；按住Crtl键单击，可删除当前页动画；按住Shift键单击，可删除全部页面的动画。";
            this.删除动画.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.删除动画_Click);
            // 
            // 清空页外
            // 
            this.清空页外.Image = ((System.Drawing.Image)(resources.GetObject("清空页外.Image")));
            this.清空页外.Label = "清空页外";
            this.清空页外.Name = "清空页外";
            this.清空页外.ScreenTip = "使用说明：";
            this.清空页外.ShowImage = true;
            this.清空页外.SuperTip = "使用本功能，可对所选幻灯片页面的页外元素进行清空（超出页面边界的元素）。";
            this.清空页外.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.清空页外_Click);
            // 
            // 清除备注
            // 
            this.清除备注.Image = ((System.Drawing.Image)(resources.GetObject("清除备注.Image")));
            this.清除备注.Label = "清除备注";
            this.清除备注.Name = "清除备注";
            this.清除备注.ScreenTip = "使用说明：";
            this.清除备注.ShowImage = true;
            this.清除备注.SuperTip = "使用本功能，可将所选页面的备注进行清除。";
            this.清除备注.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.清除备注_Click);
            // 
            // 清除超链接
            // 
            this.清除超链接.Image = ((System.Drawing.Image)(resources.GetObject("清除超链接.Image")));
            this.清除超链接.Label = "清除超链接";
            this.清除超链接.Name = "清除超链接";
            this.清除超链接.ScreenTip = "使用说明：";
            this.清除超链接.ShowImage = true;
            this.清除超链接.SuperTip = "使用本功能，可一键清除所选对象或所选页面的超链接。";
            this.清除超链接.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.清除超链接_Click);
            // 
            // 删除未用版式
            // 
            this.删除未用版式.Image = ((System.Drawing.Image)(resources.GetObject("删除未用版式.Image")));
            this.删除未用版式.Label = "删除未用版式";
            this.删除未用版式.Name = "删除未用版式";
            this.删除未用版式.ScreenTip = "使用说明：";
            this.删除未用版式.ShowImage = true;
            this.删除未用版式.SuperTip = "使用本功能，可一键删除未使用的版式，可在一定程序上缩减文件大小。";
            this.删除未用版式.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.删除未用版式_Click);
            // 
            // Timer
            // 
            this.Timer.Image = ((System.Drawing.Image)(resources.GetObject("Timer.Image")));
            this.Timer.Label = "计时器";
            this.Timer.Name = "Timer";
            this.Timer.ScreenTip = "使用说明：";
            this.Timer.ShowImage = true;
            this.Timer.SuperTip = "支持“倒计时”和“顺计时”，默认为“倒计时”。支持定义界面样式。";
            this.Timer.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Timer_Click);
            // 
            // 板贴辅助
            // 
            this.板贴辅助.Image = ((System.Drawing.Image)(resources.GetObject("板贴辅助.Image")));
            this.板贴辅助.Label = "板贴辅助";
            this.板贴辅助.Name = "板贴辅助";
            this.板贴辅助.ScreenTip = "使用说明：";
            this.板贴辅助.ShowImage = true;
            this.板贴辅助.SuperTip = "在当前页幻灯片，选中一个或多个云朵字，在弹出的对话框中输入分行文本，可批量套用样式和生成多页云朵字，以便打印。按住Ctrl键单击支持导入txt格式的分行文本。";
            this.板贴辅助.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.板贴辅助_Click);
            // 
            // 检测字体
            // 
            this.检测字体.Image = ((System.Drawing.Image)(resources.GetObject("检测字体.Image")));
            this.检测字体.Label = "检测字体";
            this.检测字体.Name = "检测字体";
            this.检测字体.ScreenTip = "使用说明：";
            this.检测字体.ShowImage = true;
            this.检测字体.SuperTip = "用于检测当前幻灯片文档文本所有字体与未在文本中使用的字体，支持导出已用字体，将字体和文档一起打包。";
            this.检测字体.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.检测字体_Click);
            // 
            // 生成样机
            // 
            this.生成样机.Image = ((System.Drawing.Image)(resources.GetObject("生成样机.Image")));
            this.生成样机.Label = "生成样机";
            this.生成样机.Name = "生成样机";
            this.生成样机.ScreenTip = "使用说明：";
            this.生成样机.ShowImage = true;
            this.生成样机.SuperTip = "选中幻灯片页面，可将所选幻灯片填充到样机中，生成展示样机。";
            this.生成样机.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.生成样机_Click);
            // 
            // 图形修剪
            // 
            this.图形修剪.Image = ((System.Drawing.Image)(resources.GetObject("图形修剪.Image")));
            this.图形修剪.Label = "一键裁边";
            this.图形修剪.Name = "图形修剪";
            this.图形修剪.ScreenTip = "使用说明：";
            this.图形修剪.ShowImage = true;
            this.图形修剪.SuperTip = "选中对象，使用本功能可将超出页面以外的部分裁剪掉。";
            this.图形修剪.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.图形修剪_Click);
            // 
            // button20
            // 
            this.button20.Image = ((System.Drawing.Image)(resources.GetObject("button20.Image")));
            this.button20.Label = "图层显隐";
            this.button20.Name = "button20";
            this.button20.ScreenTip = "使用说明：";
            this.button20.ShowImage = true;
            this.button20.SuperTip = "选择对象，切换对象图层显隐状态。";
            this.button20.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button20_Click);
            // 
            // 快捷盒子
            // 
            this.快捷盒子.Image = ((System.Drawing.Image)(resources.GetObject("快捷盒子.Image")));
            this.快捷盒子.Label = "快捷盒子";
            this.快捷盒子.Name = "快捷盒子";
            this.快捷盒子.ShowImage = true;
            this.快捷盒子.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.快捷盒子_Click);
            // 
            // Replaceaudio
            // 
            this.Replaceaudio.Image = ((System.Drawing.Image)(resources.GetObject("Replaceaudio.Image")));
            this.Replaceaudio.Label = "替换音频";
            this.Replaceaudio.Name = "Replaceaudio";
            this.Replaceaudio.ScreenTip = "使用说明：";
            this.Replaceaudio.ShowImage = true;
            this.Replaceaudio.SuperTip = "选中音频图标，可直接替换原音频，并获取原音频的部分相同属性。";
            this.Replaceaudio.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Replaceaudio_Click);
            // 
            // 插入矩形
            // 
            this.插入矩形.Image = ((System.Drawing.Image)(resources.GetObject("插入矩形.Image")));
            this.插入矩形.Label = "插入矩形";
            this.插入矩形.Name = "插入矩形";
            this.插入矩形.ScreenTip = "使用说明：";
            this.插入矩形.ShowImage = true;
            this.插入矩形.SuperTip = "无选中对象，默认单击，插入与幻灯片等大的矩形；选中对象，默认单击，在所选对象顶层插入与其等大的矩形；按Ctrl单击，则在所选对象底层插入等大的矩形。";
            this.插入矩形.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.插入矩形_Click);
            // 
            // splitButton2
            // 
            this.splitButton2.Label = "splitButton2";
            this.splitButton2.Name = "splitButton2";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.课件帮PPT助手);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.课件帮PPT助手.ResumeLayout(false);
            this.课件帮PPT助手.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group10.ResumeLayout(false);
            this.group10.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group8.ResumeLayout(false);
            this.group8.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.group9.ResumeLayout(false);
            this.group9.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab 课件帮PPT助手;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal RibbonButton button5;
        internal RibbonButton button6;
        internal RibbonGroup group3;
        internal RibbonButton button10;
        internal RibbonGroup group4;
        internal RibbonButton button11;
        internal RibbonButton button13;
        internal RibbonGroup group6;
        internal RibbonButton button19;
        internal RibbonButton 形状填图;
        internal RibbonButton Replaceaudio;
        internal RibbonButton Timer;
        internal RibbonButton Type;
        internal RibbonButton Selectsize;
        internal RibbonButton SelectedColor;
        internal RibbonButton Selectedline;
        internal RibbonButton Selectfontsize;
        internal RibbonButton Masking;
        internal RibbonGroup group9;
        public RibbonComboBox comboBox1;
        internal RibbonComboBox comboBox2;
        internal RibbonButton Gradientrectangle;
        internal RibbonButton Expandimage;
        internal RibbonGroup group8;
        internal RibbonButton Pagecentered;
        internal RibbonButton Mosaic;
        internal RibbonButton ApplyFilter;
        internal RibbonButton button20;
        internal RibbonButton toggleTaskPaneButton;
        internal RibbonGroup group10;
        internal RibbonButton 平移居中;
        internal RibbonButton 矩形拆分;
        internal RibbonButton 批量改字;
        internal RibbonButton 便捷注音;
        internal RibbonButton 笔顺图解;
        internal RibbonButton 生字格子;
        internal RibbonButton 生字赋格;
        internal RibbonSplitButton 筛选;
        internal RibbonMenu 更多便捷;
        internal RibbonSplitButton 交换;
        internal RibbonButton 交换位置;
        internal RibbonButton 交换文字;
        internal RibbonButton 交换格式;
        internal RibbonButton 交换尺寸;
        internal RibbonButton 选择增强;
        internal RibbonButton 沿线分布;
        internal RibbonMenu 分布;
        internal RibbonButton 板贴辅助;
        internal RibbonButton 去除边距;
        internal RibbonButton 首行缩进;
        internal RibbonButton 矩阵分布;
        internal RibbonButton 单字拆分;
        internal RibbonButton 拆分段落;
        internal RibbonButton 环形分布;
        internal RibbonSplitButton 在线工具;
        internal RibbonButton 趣作图;
        internal RibbonSplitButton 贴边对齐;
        internal RibbonSplitButton 选择居中;
        internal RibbonSplitButton 图片;
        internal RibbonButton 原位转图;
        internal RibbonSplitButton 统一;
        internal RibbonButton 统一大小;
        internal RibbonButton 统一格式;
        internal RibbonButton 智能缩放;
        internal RibbonSplitButton 文本;
        internal RibbonButton Tmttool;
        internal RibbonSplitButton 抠图;
        internal RibbonButton Bgsub;
        internal RibbonSplitButton 矢量;
        internal RibbonButton LCopy;
        internal RibbonButton 生成样机;
        internal RibbonButton 图形修剪;
        internal RibbonButton 四线三格;
        internal RibbonSplitButton 常用格子;
        internal RibbonButton 一键注音;
        internal RibbonSplitButton 注音工具;
        internal RibbonButton Zici;
        internal RibbonSplitButton 拓展应用;
        internal RibbonButton WritePinyin;
        internal RibbonButton 检测字体;
        internal RibbonSplitButton 组合;
        internal RibbonButton 重叠组合;
        internal RibbonButton 临近组合;
        internal RibbonButton 同色组合;
        internal RibbonButton 图形分割;
        internal RibbonButton 快捷盒子;
        internal RibbonButton 批量换图;
        internal RibbonButton 原位转JPG;
        internal RibbonButton 图层;
        internal RibbonButton 统一控点;
        internal RibbonButton 文本矢量化;
        internal RibbonSplitButton 删除;
        internal RibbonButton 删除动画;
        internal RibbonButton 清空页外;
        internal RibbonButton 清除备注;
        internal RibbonButton 清除超链接;
        internal RibbonButton 删除未用版式;
        internal RibbonButton 部首描红;
        internal RibbonButton 分解笔顺;
        internal RibbonSplitButton 关于我;
        internal RibbonButton 检查更新;
        internal RibbonButton 文本居中;
        internal RibbonButton 自动补齐;
        internal RibbonSplitButton 注音编辑;
        internal RibbonButton 删列补行;
        internal RibbonButton 合并段落;
        internal RibbonButton 重设表格;
        internal RibbonButton 文转表格;
        internal RibbonSplitButton 晓声通在线注音;
        internal RibbonButton 在线注音编辑器;
        internal RibbonButton 左右镜像;
        internal RibbonSplitButton splitButton2;
        internal RibbonSplitButton 复制;
        internal RibbonButton 上下镜像;
        internal RibbonButton 分解拼音;
        internal RibbonButton 辐射连线;
        internal RibbonButton 文写入形;
        internal RibbonSplitButton 绘图;
        internal RibbonButton 层级关系;
        internal RibbonButton 括弧关系;
        internal RibbonButton 多音字词填空;
        internal RibbonButton 字号调整;
        internal RibbonButton 奇数行行高设置;
        internal RibbonButton 字词加点;
        internal RibbonButton 拼音升调;
        internal RibbonButton 增加行宽;
        internal RibbonButton 插入矩形;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
