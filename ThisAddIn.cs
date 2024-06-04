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

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // 订阅 PowerPoint 的 NewPresentation 和 PresentationOpen 事件
            ((PowerPoint.EApplication_Event)this.Application).NewPresentation += new PowerPoint.EApplication_NewPresentationEventHandler(Application_NewPresentation);
            ((PowerPoint.EApplication_Event)this.Application).PresentationOpen += new PowerPoint.EApplication_PresentationOpenEventHandler(Application_PresentationOpen);
            ((PowerPoint.EApplication_Event)this.Application).PresentationClose += new PowerPoint.EApplication_PresentationCloseEventHandler(Application_PresentationClose);

            // 为当前打开的所有演示文稿创建 CustomTaskPane
            foreach (PowerPoint.Presentation pres in this.Application.Presentations)
            {
                AddCustomTaskPane(pres);
            }
        }

        private void Application_NewPresentation(PowerPoint.Presentation Pres)
        {
            // 在新建文档时创建新的 CustomTaskPane
            AddCustomTaskPane(Pres);
        }

        private void Application_PresentationOpen(PowerPoint.Presentation Pres)
        {
            // 在打开文档时创建新的 CustomTaskPane
            AddCustomTaskPane(Pres);
        }

        private void Application_PresentationClose(PowerPoint.Presentation Pres)
        {
            // 在关闭文档时移除 CustomTaskPane
            if (customTaskPanes.ContainsKey(Pres))
            {
                CustomTaskPane taskPane = customTaskPanes[Pres];
                this.CustomTaskPanes.Remove(taskPane);
                customTaskPanes.Remove(Pres);
            }
        }

        private void AddCustomTaskPane(PowerPoint.Presentation pres)
        {
            // 创建用户控件实例
            DesignTools designTools = new DesignTools();

            // 创建 CustomTaskPane 并将用户控件添加到其中
            CustomTaskPane taskPane = this.CustomTaskPanes.Add(designTools, "学科", pres);
            taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            taskPane.Width = 332; // 设置侧边栏的宽度
            taskPane.Visible = true;

            // 将 CustomTaskPane 存储在字典中，以便管理
            customTaskPanes[pres] = taskPane;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // 可在此处添加任何必要的清理代码
            ((PowerPoint.EApplication_Event)this.Application).NewPresentation -= new PowerPoint.EApplication_NewPresentationEventHandler(Application_NewPresentation);
            ((PowerPoint.EApplication_Event)this.Application).PresentationOpen -= new PowerPoint.EApplication_PresentationOpenEventHandler(Application_PresentationOpen);
            ((PowerPoint.EApplication_Event)this.Application).PresentationClose -= new PowerPoint.EApplication_PresentationCloseEventHandler(Application_PresentationClose);
        }

        public void ToggleTaskPaneVisibility()
        {
            PowerPoint.Presentation pres = this.Application.ActivePresentation;
            if (customTaskPanes.ContainsKey(pres))
            {
                CustomTaskPane taskPane = customTaskPanes[pres];
                taskPane.Visible = !taskPane.Visible;
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
