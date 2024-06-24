using System;
using System.Collections.Generic;
using Microsoft.Office.Tools;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace 课件帮PPT助手
{
    public partial class ThisAddIn
    {
        private Dictionary<PowerPoint.Presentation, CustomTaskPane> customTaskPanes = new Dictionary<PowerPoint.Presentation, CustomTaskPane>();
        private bool taskPaneVisible = false; // 默认任务窗格状态为隐藏

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // 订阅 PowerPoint 的 NewPresentation 和 PresentationOpen 事件
            ((PowerPoint.EApplication_Event)this.Application).NewPresentation += Application_NewPresentation;
            ((PowerPoint.EApplication_Event)this.Application).PresentationOpen += Application_PresentationOpen;
            ((PowerPoint.EApplication_Event)this.Application).PresentationClose += Application_PresentationClose;

            // 为当前打开的所有演示文稿创建 CustomTaskPane
            foreach (PowerPoint.Presentation pres in this.Application.Presentations)
            {
                AddCustomTaskPane(pres, taskPaneVisible);
            }
        }

        private void Application_NewPresentation(PowerPoint.Presentation pres)
        {
            // 在新建文档时设置任务窗格的显隐状态
            SetTaskPaneVisibility(pres, taskPaneVisible);
        }

        private void Application_PresentationOpen(PowerPoint.Presentation pres)
        {
            // 在打开文档时设置任务窗格的显隐状态
            SetTaskPaneVisibility(pres, taskPaneVisible);
        }

        private void Application_PresentationClose(PowerPoint.Presentation pres)
        {
            // 在关闭文档时移除 CustomTaskPane，并释放资源
            if (customTaskPanes.ContainsKey(pres))
            {
                CustomTaskPane taskPane = customTaskPanes[pres];
                taskPane.VisibleChanged -= TaskPane_VisibleChanged;
                customTaskPanes.Remove(pres);
                taskPane.Dispose();
            }
        }

        private void AddCustomTaskPane(PowerPoint.Presentation pres, bool isVisible)
        {
            // 创建用户控件实例
            DesignTools designTools = new DesignTools();

            // 创建 CustomTaskPane 并将用户控件添加到其中
            CustomTaskPane taskPane = this.CustomTaskPanes.Add(designTools, "学科", pres);
            taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            taskPane.Width = 280; // 设置侧边栏的宽度
            taskPane.Visible = isVisible; // 同步任务窗格初始状态

            // 订阅任务窗格的 VisibleChanged 事件
            taskPane.VisibleChanged += TaskPane_VisibleChanged;

            // 将 CustomTaskPane 存储在字典中，以便管理
            customTaskPanes[pres] = taskPane;
        }

        private void SetTaskPaneVisibility(PowerPoint.Presentation pres, bool isVisible)
        {
            if (!customTaskPanes.ContainsKey(pres))
            {
                AddCustomTaskPane(pres, isVisible);
            }
            else
            {
                customTaskPanes[pres].VisibleChanged -= TaskPane_VisibleChanged;
                customTaskPanes[pres].Visible = isVisible;
                customTaskPanes[pres].VisibleChanged += TaskPane_VisibleChanged;
            }
        }

        private void TaskPane_VisibleChanged(object sender, EventArgs e)
        {
            var taskPane = sender as CustomTaskPane;
            if (taskPane != null)
            {
                taskPaneVisible = taskPane.Visible;
                foreach (var pane in customTaskPanes.Values)
                {
                    if (pane != taskPane)
                    {
                        pane.VisibleChanged -= TaskPane_VisibleChanged;
                        pane.Visible = taskPaneVisible;
                        pane.VisibleChanged += TaskPane_VisibleChanged;
                    }
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // 移除所有事件处理程序
            ((PowerPoint.EApplication_Event)this.Application).NewPresentation -= Application_NewPresentation;
            ((PowerPoint.EApplication_Event)this.Application).PresentationOpen -= Application_PresentationOpen;
            ((PowerPoint.EApplication_Event)this.Application).PresentationClose -= Application_PresentationClose;

            // 清理所有 CustomTaskPane
            foreach (var taskPane in customTaskPanes.Values)
            {
                taskPane.VisibleChanged -= TaskPane_VisibleChanged;
                taskPane.Dispose();
            }
            customTaskPanes.Clear();
        }

        public void ToggleTaskPaneVisibility()
        {
            taskPaneVisible = !taskPaneVisible;
            foreach (var taskPane in customTaskPanes.Values)
            {
                taskPane.VisibleChanged -= TaskPane_VisibleChanged;
                taskPane.Visible = taskPaneVisible;
                taskPane.VisibleChanged += TaskPane_VisibleChanged;
            }
        }

        public CustomTaskPane GetCustomTaskPane(PowerPoint.Presentation pres)
        {
            if (customTaskPanes.ContainsKey(pres))
            {
                return customTaskPanes[pres];
            }
            return null;
        }

        #region VSTO 生成的代码

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
