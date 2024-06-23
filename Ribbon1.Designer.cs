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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
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
            this.课件帮PPT助手 = this.Factory.CreateRibbonTab();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button7 = this.Factory.CreateRibbonButton();
            this.group10 = this.Factory.CreateRibbonGroup();
            this.toggleTaskPaneButton = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button5 = this.Factory.CreateRibbonButton();
            this.笔顺图解 = this.Factory.CreateRibbonButton();
            this.生字赋格 = this.Factory.CreateRibbonButton();
            this.常用格子 = this.Factory.CreateRibbonSplitButton();
            this.生字格子 = this.Factory.CreateRibbonButton();
            this.四线三格 = this.Factory.CreateRibbonButton();
            this.注音工具 = this.Factory.CreateRibbonSplitButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.便捷注音 = this.Factory.CreateRibbonButton();
            this.一键注音 = this.Factory.CreateRibbonButton();
            this.提取拼音 = this.Factory.CreateRibbonButton();
            this.拓展应用 = this.Factory.CreateRibbonSplitButton();
            this.Zici = this.Factory.CreateRibbonButton();
            this.WritePinyin = this.Factory.CreateRibbonButton();
            this.生字教学 = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.Masking = this.Factory.CreateRibbonButton();
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
            this.group4 = this.Factory.CreateRibbonGroup();
            this.对齐增强 = this.Factory.CreateRibbonMenu();
            this.平移居中 = this.Factory.CreateRibbonButton();
            this.splitButton1 = this.Factory.CreateRibbonSplitButton();
            this.分组匹配 = this.Factory.CreateRibbonButton();
            this.指定对齐 = this.Factory.CreateRibbonButton();
            this.移动对齐 = this.Factory.CreateRibbonButton();
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
            this.group8 = this.Factory.CreateRibbonGroup();
            this.筛选 = this.Factory.CreateRibbonSplitButton();
            this.Type = this.Factory.CreateRibbonButton();
            this.Selectsize = this.Factory.CreateRibbonButton();
            this.SelectedColor = this.Factory.CreateRibbonButton();
            this.Selectedline = this.Factory.CreateRibbonButton();
            this.Selectfontsize = this.Factory.CreateRibbonButton();
            this.选择增强 = this.Factory.CreateRibbonButton();
            this.智能缩放 = this.Factory.CreateRibbonButton();
            this.文本 = this.Factory.CreateRibbonSplitButton();
            this.去除边距 = this.Factory.CreateRibbonButton();
            this.首行缩进 = this.Factory.CreateRibbonButton();
            this.单字拆分 = this.Factory.CreateRibbonButton();
            this.拆分段落 = this.Factory.CreateRibbonButton();
            this.批量改字 = this.Factory.CreateRibbonButton();
            this.更多便捷 = this.Factory.CreateRibbonMenu();
            this.图片 = this.Factory.CreateRibbonSplitButton();
            this.Replaceimage = this.Factory.CreateRibbonButton();
            this.原位转图 = this.Factory.CreateRibbonButton();
            this.音频 = this.Factory.CreateRibbonSplitButton();
            this.Replaceaudio = this.Factory.CreateRibbonButton();
            this.交换 = this.Factory.CreateRibbonSplitButton();
            this.交换位置 = this.Factory.CreateRibbonButton();
            this.交换文字 = this.Factory.CreateRibbonButton();
            this.交换格式 = this.Factory.CreateRibbonButton();
            this.交换尺寸 = this.Factory.CreateRibbonButton();
            this.交换图层 = this.Factory.CreateRibbonButton();
            this.完全交换 = this.Factory.CreateRibbonButton();
            this.统一 = this.Factory.CreateRibbonSplitButton();
            this.统一大小 = this.Factory.CreateRibbonButton();
            this.统一格式 = this.Factory.CreateRibbonButton();
            this.生成样机 = this.Factory.CreateRibbonButton();
            this.图形修剪 = this.Factory.CreateRibbonButton();
            this.LCopy = this.Factory.CreateRibbonButton();
            this.button20 = this.Factory.CreateRibbonButton();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.Timer = this.Factory.CreateRibbonButton();
            this.板贴辅助 = this.Factory.CreateRibbonButton();
            this.检测字体 = this.Factory.CreateRibbonButton();
            this.group7 = this.Factory.CreateRibbonGroup();
            this.尺寸缩放 = this.Factory.CreateRibbonEditBox();
            this.批量命名 = this.Factory.CreateRibbonEditBox();
            this.原位复制 = this.Factory.CreateRibbonEditBox();
            this.group9 = this.Factory.CreateRibbonGroup();
            this.comboBox1 = this.Factory.CreateRibbonComboBox();
            this.comboBox2 = this.Factory.CreateRibbonComboBox();
            this.课件帮PPT助手.SuspendLayout();
            this.group3.SuspendLayout();
            this.group10.SuspendLayout();
            this.group2.SuspendLayout();
            this.group1.SuspendLayout();
            this.group4.SuspendLayout();
            this.group8.SuspendLayout();
            this.group6.SuspendLayout();
            this.group7.SuspendLayout();
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
            this.课件帮PPT助手.Groups.Add(this.group7);
            this.课件帮PPT助手.Groups.Add(this.group9);
            this.课件帮PPT助手.Label = "快捷魔方";
            this.课件帮PPT助手.Name = "课件帮PPT助手";
            // 
            // group3
            // 
            this.group3.Items.Add(this.button7);
            this.group3.Label = "关于我";
            this.group3.Name = "group3";
            // 
            // button7
            // 
            this.button7.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button7.Image = ((System.Drawing.Image)(resources.GetObject("button7.Image")));
            this.button7.Label = "Andy";
            this.button7.Name = "button7";
            this.button7.ScreenTip = "关于我：";
            this.button7.ShowImage = true;
            this.button7.SuperTip = "访问Andy老师创建的资源分享博客";
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button7_Click);
            // 
            // group10
            // 
            this.group10.Items.Add(this.toggleTaskPaneButton);
            this.group10.Name = "group10";
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
            // group2
            // 
            this.group2.Items.Add(this.button5);
            this.group2.Items.Add(this.笔顺图解);
            this.group2.Items.Add(this.生字赋格);
            this.group2.Items.Add(this.常用格子);
            this.group2.Items.Add(this.注音工具);
            this.group2.Items.Add(this.拓展应用);
            this.group2.Label = "字音字形";
            this.group2.Name = "group2";
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
            this.注音工具.Items.Add(this.提取拼音);
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
            this.一键注音.ShowImage = true;
            this.一键注音.SuperTip = "（源：简易字典）";
            this.一键注音.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.一键注音_Click);
            // 
            // 提取拼音
            // 
            this.提取拼音.Image = ((System.Drawing.Image)(resources.GetObject("提取拼音.Image")));
            this.提取拼音.Label = "联网注音";
            this.提取拼音.Name = "提取拼音";
            this.提取拼音.ShowImage = true;
            this.提取拼音.SuperTip = "（源：百度汉语）";
            this.提取拼音.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.提取拼音_Click);
            // 
            // 拓展应用
            // 
            this.拓展应用.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.拓展应用.Image = ((System.Drawing.Image)(resources.GetObject("拓展应用.Image")));
            this.拓展应用.Items.Add(this.Zici);
            this.拓展应用.Items.Add(this.WritePinyin);
            this.拓展应用.Items.Add(this.生字教学);
            this.拓展应用.Label = "拓展应用";
            this.拓展应用.Name = "拓展应用";
            // 
            // Zici
            // 
            this.Zici.Image = ((System.Drawing.Image)(resources.GetObject("Zici.Image")));
            this.Zici.Label = "看拼音写词语";
            this.Zici.Name = "Zici";
            this.Zici.ShowImage = true;
            this.Zici.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Zici_Click);
            // 
            // WritePinyin
            // 
            this.WritePinyin.Image = ((System.Drawing.Image)(resources.GetObject("WritePinyin.Image")));
            this.WritePinyin.Label = "看词语写拼音";
            this.WritePinyin.Name = "WritePinyin";
            this.WritePinyin.ShowImage = true;
            this.WritePinyin.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.WritePinyin_Click);
            // 
            // 生字教学
            // 
            this.生字教学.Image = ((System.Drawing.Image)(resources.GetObject("生字教学.Image")));
            this.生字教学.Label = "创建生字模板";
            this.生字教学.Name = "生字教学";
            this.生字教学.ShowImage = true;
            this.生字教学.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.生字教学_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.button2);
            this.group1.Items.Add(this.button6);
            this.group1.Items.Add(this.Masking);
            this.group1.Items.Add(this.Gradientrectangle);
            this.group1.Items.Add(this.在线工具);
            this.group1.Items.Add(this.矩形拆分);
            this.group1.Items.Add(this.Mosaic);
            this.group1.Items.Add(this.ApplyFilter);
            this.group1.Items.Add(this.Expandimage);
            this.group1.Label = "图形处理";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "全屏图形";
            this.button1.Name = "button1";
            this.button1.ScreenTip = "使用说明：";
            this.button1.ShowImage = true;
            this.button1.SuperTip = "选中一张图片或形状，放大至全屏。";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.Label = "参考裁剪";
            this.button2.Name = "button2";
            this.button2.ScreenTip = "使用说明：";
            this.button2.ShowImage = true;
            this.button2.SuperTip = "选中多张图片，以选中的第一张图片的宽高和比例为参考，统一后面选中的所有图片大小。（支持第一个被选中对象为形状）";
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
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
            // group4
            // 
            this.group4.Items.Add(this.对齐增强);
            this.group4.Items.Add(this.分布);
            this.group4.Label = "参考对齐";
            this.group4.Name = "group4";
            // 
            // 对齐增强
            // 
            this.对齐增强.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.对齐增强.Image = ((System.Drawing.Image)(resources.GetObject("对齐增强.Image")));
            this.对齐增强.Items.Add(this.平移居中);
            this.对齐增强.Items.Add(this.splitButton1);
            this.对齐增强.Items.Add(this.移动对齐);
            this.对齐增强.Label = "对齐增强";
            this.对齐增强.Name = "对齐增强";
            this.对齐增强.ShowImage = true;
            // 
            // 平移居中
            // 
            this.平移居中.Image = ((System.Drawing.Image)(resources.GetObject("平移居中.Image")));
            this.平移居中.Label = "平移居中";
            this.平移居中.Name = "平移居中";
            this.平移居中.ScreenTip = "使用说明：";
            this.平移居中.ShowImage = true;
            this.平移居中.SuperTip = "以第一个被选中的对象为参考（基准），固定其位置不变，同时将后续所选的其他对象都看作一个整体（无论数量的多少），按照既定的对齐方式对齐到参考对象中。";
            this.平移居中.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.平移居中_Click);
            // 
            // splitButton1
            // 
            this.splitButton1.Image = ((System.Drawing.Image)(resources.GetObject("splitButton1.Image")));
            this.splitButton1.Items.Add(this.分组匹配);
            this.splitButton1.Items.Add(this.指定对齐);
            this.splitButton1.Label = "匹配对齐";
            this.splitButton1.Name = "splitButton1";
            // 
            // 分组匹配
            // 
            this.分组匹配.Image = ((System.Drawing.Image)(resources.GetObject("分组匹配.Image")));
            this.分组匹配.Label = "分组匹配";
            this.分组匹配.Name = "分组匹配";
            this.分组匹配.ScreenTip = "使用说明：";
            this.分组匹配.ShowImage = true;
            this.分组匹配.SuperTip = "程序会将用户所选对象按照选择的顺序均分成两组对象，第一组叫“参考组”（固定位置不变），第二组叫“目标组”，这两组对象将进行一一匹配对齐。";
            this.分组匹配.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.匹配对齐_Click);
            // 
            // 指定对齐
            // 
            this.指定对齐.Image = ((System.Drawing.Image)(resources.GetObject("指定对齐.Image")));
            this.指定对齐.Label = "连线匹配";
            this.指定对齐.Name = "指定对齐";
            this.指定对齐.ScreenTip = "使用说明：";
            this.指定对齐.ShowImage = true;
            this.指定对齐.SuperTip = "像做连线题一样指定某两个对象就行对齐（可在连选的状态下进行）。";
            this.指定对齐.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.指定对齐_Click);
            // 
            // 移动对齐
            // 
            this.移动对齐.Image = ((System.Drawing.Image)(resources.GetObject("移动对齐.Image")));
            this.移动对齐.Label = "移动对齐";
            this.移动对齐.Name = "移动对齐";
            this.移动对齐.ShowImage = true;
            this.移动对齐.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.移动对齐_Click);
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
            this.Pagecentered.SuperTip = "将所选对象整体平移到页面中心。";
            this.Pagecentered.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Pagecentered_Click);
            // 
            // group8
            // 
            this.group8.Items.Add(this.筛选);
            this.group8.Items.Add(this.选择增强);
            this.group8.Items.Add(this.智能缩放);
            this.group8.Items.Add(this.文本);
            this.group8.Items.Add(this.更多便捷);
            this.group8.Label = "便捷常用";
            this.group8.Name = "group8";
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
            this.Selectedline.SuperTip = "选中一个对象：1.单击，则同时选中与他线条宽度相同的所有对象；2.按住Ctrl键单击，则同时选中当前幻灯片线条颜色与它相同的所有对象；3.按住Shift键单击，则" +
    "同时选中与它相同线条类型（如虚线、实线）的所有对象。";
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
            this.智能缩放.SuperTip = "可以对所选对象大小和属性进行等比缩放，且支持更改缩放中心。";
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
            this.批量改字.SuperTip = "选中一个或多个文本框，可对它们的文本进行批量修改。";
            this.批量改字.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.批量改字_Click);
            // 
            // 更多便捷
            // 
            this.更多便捷.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.更多便捷.Image = ((System.Drawing.Image)(resources.GetObject("更多便捷.Image")));
            this.更多便捷.Items.Add(this.图片);
            this.更多便捷.Items.Add(this.音频);
            this.更多便捷.Items.Add(this.交换);
            this.更多便捷.Items.Add(this.统一);
            this.更多便捷.Items.Add(this.生成样机);
            this.更多便捷.Items.Add(this.图形修剪);
            this.更多便捷.Items.Add(this.LCopy);
            this.更多便捷.Items.Add(this.button20);
            this.更多便捷.Label = "便捷";
            this.更多便捷.Name = "更多便捷";
            this.更多便捷.ShowImage = true;
            // 
            // 图片
            // 
            this.图片.Image = ((System.Drawing.Image)(resources.GetObject("图片.Image")));
            this.图片.Items.Add(this.Replaceimage);
            this.图片.Items.Add(this.原位转图);
            this.图片.Label = "图片";
            this.图片.Name = "图片";
            // 
            // Replaceimage
            // 
            this.Replaceimage.Image = ((System.Drawing.Image)(resources.GetObject("Replaceimage.Image")));
            this.Replaceimage.Label = "批量换图";
            this.Replaceimage.Name = "Replaceimage";
            this.Replaceimage.ScreenTip = "使用说明：";
            this.Replaceimage.ShowImage = true;
            this.Replaceimage.SuperTip = "1.选中多张图片，单击则一键原位批量换图；2.按住Ctrl键单击，则一键原位换图，且保持与原图尺寸相等大小。";
            this.Replaceimage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Replaceimage_Click);
            // 
            // 原位转图
            // 
            this.原位转图.Image = ((System.Drawing.Image)(resources.GetObject("原位转图.Image")));
            this.原位转图.Label = "原位转PNG";
            this.原位转图.Name = "原位转图";
            this.原位转图.ShowImage = true;
            this.原位转图.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.原位转图_Click);
            // 
            // 音频
            // 
            this.音频.Image = ((System.Drawing.Image)(resources.GetObject("音频.Image")));
            this.音频.Items.Add(this.Replaceaudio);
            this.音频.Label = "音频";
            this.音频.Name = "音频";
            // 
            // Replaceaudio
            // 
            this.Replaceaudio.Image = ((System.Drawing.Image)(resources.GetObject("Replaceaudio.Image")));
            this.Replaceaudio.Label = "音频替换";
            this.Replaceaudio.Name = "Replaceaudio";
            this.Replaceaudio.ScreenTip = "使用说明：";
            this.Replaceaudio.ShowImage = true;
            this.Replaceaudio.SuperTip = "选中音频图标，可直接替换原音频，并获取原音频的部分相同属性。";
            this.Replaceaudio.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Replaceaudio_Click);
            // 
            // 交换
            // 
            this.交换.Image = ((System.Drawing.Image)(resources.GetObject("交换.Image")));
            this.交换.Items.Add(this.交换位置);
            this.交换.Items.Add(this.交换文字);
            this.交换.Items.Add(this.交换格式);
            this.交换.Items.Add(this.交换尺寸);
            this.交换.Items.Add(this.交换图层);
            this.交换.Items.Add(this.完全交换);
            this.交换.Label = "交换";
            this.交换.Name = "交换";
            // 
            // 交换位置
            // 
            this.交换位置.Image = ((System.Drawing.Image)(resources.GetObject("交换位置.Image")));
            this.交换位置.Label = "交换位置";
            this.交换位置.Name = "交换位置";
            this.交换位置.ShowImage = true;
            this.交换位置.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.交换位置_Click);
            // 
            // 交换文字
            // 
            this.交换文字.Image = ((System.Drawing.Image)(resources.GetObject("交换文字.Image")));
            this.交换文字.Label = "交换文字";
            this.交换文字.Name = "交换文字";
            this.交换文字.ShowImage = true;
            this.交换文字.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.交换文字_Click);
            // 
            // 交换格式
            // 
            this.交换格式.Image = ((System.Drawing.Image)(resources.GetObject("交换格式.Image")));
            this.交换格式.Label = "交换格式";
            this.交换格式.Name = "交换格式";
            this.交换格式.ShowImage = true;
            this.交换格式.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.交换格式_Click);
            // 
            // 交换尺寸
            // 
            this.交换尺寸.Image = ((System.Drawing.Image)(resources.GetObject("交换尺寸.Image")));
            this.交换尺寸.Label = "交换尺寸";
            this.交换尺寸.Name = "交换尺寸";
            this.交换尺寸.ShowImage = true;
            this.交换尺寸.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.交换尺寸_Click);
            // 
            // 交换图层
            // 
            this.交换图层.Image = ((System.Drawing.Image)(resources.GetObject("交换图层.Image")));
            this.交换图层.Label = "交换图层";
            this.交换图层.Name = "交换图层";
            this.交换图层.ShowImage = true;
            this.交换图层.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.交换图层_Click);
            // 
            // 完全交换
            // 
            this.完全交换.Image = ((System.Drawing.Image)(resources.GetObject("完全交换.Image")));
            this.完全交换.Label = "完全交换";
            this.完全交换.Name = "完全交换";
            this.完全交换.ScreenTip = "暂不支持图片";
            this.完全交换.ShowImage = true;
            this.完全交换.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.完全交换_Click);
            // 
            // 统一
            // 
            this.统一.Image = ((System.Drawing.Image)(resources.GetObject("统一.Image")));
            this.统一.Items.Add(this.统一大小);
            this.统一.Items.Add(this.统一格式);
            this.统一.Label = "统一";
            this.统一.Name = "统一";
            // 
            // 统一大小
            // 
            this.统一大小.Image = ((System.Drawing.Image)(resources.GetObject("统一大小.Image")));
            this.统一大小.Label = "统一大小";
            this.统一大小.Name = "统一大小";
            this.统一大小.ShowImage = true;
            this.统一大小.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.统一大小_Click);
            // 
            // 统一格式
            // 
            this.统一格式.Image = ((System.Drawing.Image)(resources.GetObject("统一格式.Image")));
            this.统一格式.Label = "统一格式";
            this.统一格式.Name = "统一格式";
            this.统一格式.ShowImage = true;
            this.统一格式.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.统一格式_Click);
            // 
            // 生成样机
            // 
            this.生成样机.Image = ((System.Drawing.Image)(resources.GetObject("生成样机.Image")));
            this.生成样机.Label = "生成样机";
            this.生成样机.Name = "生成样机";
            this.生成样机.ShowImage = true;
            this.生成样机.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.生成样机_Click);
            // 
            // 图形修剪
            // 
            this.图形修剪.Image = ((System.Drawing.Image)(resources.GetObject("图形修剪.Image")));
            this.图形修剪.Label = "一键裁边";
            this.图形修剪.Name = "图形修剪";
            this.图形修剪.ShowImage = true;
            this.图形修剪.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.图形修剪_Click);
            // 
            // LCopy
            // 
            this.LCopy.Image = ((System.Drawing.Image)(resources.GetObject("LCopy.Image")));
            this.LCopy.Label = "原位复制";
            this.LCopy.Name = "LCopy";
            this.LCopy.ShowImage = true;
            this.LCopy.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LCopy_Click);
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
            // group6
            // 
            this.group6.Items.Add(this.Timer);
            this.group6.Items.Add(this.板贴辅助);
            this.group6.Items.Add(this.检测字体);
            this.group6.Label = "辅助";
            this.group6.Name = "group6";
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
            this.检测字体.ShowImage = true;
            this.检测字体.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.检测字体_Click);
            // 
            // group7
            // 
            this.group7.Items.Add(this.尺寸缩放);
            this.group7.Items.Add(this.批量命名);
            this.group7.Items.Add(this.原位复制);
            this.group7.Label = "批量处理";
            this.group7.Name = "group7";
            // 
            // 尺寸缩放
            // 
            this.尺寸缩放.Image = ((System.Drawing.Image)(resources.GetObject("尺寸缩放.Image")));
            this.尺寸缩放.Label = "尺寸缩放";
            this.尺寸缩放.Name = "尺寸缩放";
            this.尺寸缩放.ScreenTip = "使用说明：";
            this.尺寸缩放.ShowImage = true;
            this.尺寸缩放.SuperTip = "输入固定数值，回车，可对一个或多个所选对象进行一定比例的缩放。输入两个数值，且用逗号分隔，可对所选对象进行等差缩放。";
            this.尺寸缩放.Text = null;
            this.尺寸缩放.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.尺寸缩放_TextChanged);
            // 
            // 批量命名
            // 
            this.批量命名.Image = ((System.Drawing.Image)(resources.GetObject("批量命名.Image")));
            this.批量命名.Label = "批量命名";
            this.批量命名.Name = "批量命名";
            this.批量命名.ScreenTip = "使用说明：";
            this.批量命名.ShowImage = true;
            this.批量命名.SuperTip = "选中一个或多个对象，在此输入前缀名，回车，可按照选择的顺序对它们进行批量命名。命名规则为“前缀名+序号”。";
            this.批量命名.Text = null;
            this.批量命名.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.批量命名_TextChanged);
            // 
            // 原位复制
            // 
            this.原位复制.Image = ((System.Drawing.Image)(resources.GetObject("原位复制.Image")));
            this.原位复制.Label = "原位复制";
            this.原位复制.Name = "原位复制";
            this.原位复制.ScreenTip = "使用说明：";
            this.原位复制.ShowImage = true;
            this.原位复制.SuperTip = "输入相应数值，回车，可对所选对象进行批量原位复制。";
            this.原位复制.Tag = "";
            this.原位复制.Text = null;
            this.原位复制.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.原位复制_TextChanged);
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
            this.group7.ResumeLayout(false);
            this.group7.PerformLayout();
            this.group9.ResumeLayout(false);
            this.group9.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab 课件帮PPT助手;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal RibbonButton button5;
        internal RibbonButton button6;
        internal RibbonGroup group3;
        internal RibbonButton button7;
        internal RibbonButton button10;
        internal RibbonGroup group4;
        internal RibbonButton button11;
        internal RibbonButton button13;
        internal RibbonGroup group6;
        internal RibbonButton button19;
        internal RibbonGroup group7;
        internal RibbonButton Replaceimage;
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
        internal RibbonButton 分组匹配;
        internal RibbonButton 矩形拆分;
        internal RibbonButton 批量改字;
        internal RibbonButton 便捷注音;
        internal RibbonButton 笔顺图解;
        internal RibbonButton 生字格子;
        internal RibbonButton 生字赋格;
        internal RibbonSplitButton 筛选;
        internal RibbonMenu 更多便捷;
        internal RibbonMenu 对齐增强;
        internal RibbonEditBox 原位复制;
        internal RibbonEditBox 尺寸缩放;
        internal RibbonEditBox 批量命名;
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
        internal RibbonSplitButton 音频;
        internal RibbonButton 环形分布;
        internal RibbonSplitButton 在线工具;
        internal RibbonButton 趣作图;
        internal RibbonSplitButton 贴边对齐;
        internal RibbonSplitButton 选择居中;
        internal RibbonButton 指定对齐;
        internal RibbonSplitButton splitButton1;
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
        internal RibbonButton 完全交换;
        internal RibbonButton 交换图层;
        internal RibbonButton 图形修剪;
        internal RibbonButton 四线三格;
        internal RibbonButton 移动对齐;
        internal RibbonSplitButton 常用格子;
        internal RibbonButton 一键注音;
        internal RibbonSplitButton 注音工具;
        internal RibbonButton 提取拼音;
        internal RibbonButton Zici;
        internal RibbonSplitButton 拓展应用;
        internal RibbonButton WritePinyin;
        internal RibbonButton 检测字体;
        internal RibbonButton 生字教学;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
