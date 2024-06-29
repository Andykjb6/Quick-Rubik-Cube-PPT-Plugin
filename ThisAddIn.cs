using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Office.Tools;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace 课件帮PPT助手
{
    public partial class ThisAddIn
    {
        private Ribbon1 ribbon;
        private Dictionary<PowerPoint.Presentation, CustomTaskPane> customTaskPanes = new Dictionary<PowerPoint.Presentation, CustomTaskPane>();
        private bool taskPaneVisible = false; // 默认任务窗格状态为隐藏

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            ((PowerPoint.EApplication_Event)this.Application).NewPresentation += Application_NewPresentation;
            ((PowerPoint.EApplication_Event)this.Application).PresentationOpen += Application_PresentationOpen;
            ((PowerPoint.EApplication_Event)this.Application).PresentationCloseFinal += Application_PresentationClose;
            ((PowerPoint.EApplication_Event)this.Application).WindowActivate += Application_WindowActivate;

            ribbon = Globals.Ribbons.Ribbon1;
            ribbon.PptApplication = this.Application;

            foreach (PowerPoint.Presentation pres in this.Application.Presentations)
            {
                AddCustomTaskPane(pres, taskPaneVisible);
            }
        }

        private void Application_NewPresentation(PowerPoint.Presentation pres)
        {
            AddCustomTaskPane(pres, taskPaneVisible);
        }

        private void Application_PresentationOpen(PowerPoint.Presentation pres)
        {
            AddCustomTaskPane(pres, taskPaneVisible);
        }

        private void Application_PresentationClose(PowerPoint.Presentation pres)
        {
            if (customTaskPanes.ContainsKey(pres))
            {
                DisposeTaskPane(pres);
                customTaskPanes.Remove(pres);
            }
        }

        private void Application_WindowActivate(PowerPoint.Presentation Pres, PowerPoint.DocumentWindow Wn)
        {
            if (customTaskPanes.ContainsKey(Pres))
            {
                SetTaskPaneVisibility(Pres, taskPaneVisible);
            }
        }

        private void AddCustomTaskPane(PowerPoint.Presentation pres, bool isVisible)
        {
            DesignTools designTools = new DesignTools();
            CustomTaskPane taskPane = null;

            try
            {
                taskPane = this.CustomTaskPanes.Add(designTools, "学科", pres);
                taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
                taskPane.Width = 280; // 设置侧边栏的宽度
                taskPane.Visible = isVisible;

                taskPane.VisibleChanged += TaskPane_VisibleChanged;
                customTaskPanes[pres] = taskPane;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error adding custom task pane: {ex.Message}");
            }
        }

        private void SetTaskPaneVisibility(PowerPoint.Presentation pres, bool isVisible)
        {
            if (customTaskPanes.ContainsKey(pres))
            {
                try
                {
                    var taskPane = customTaskPanes[pres];
                    taskPane.VisibleChanged -= TaskPane_VisibleChanged;
                    taskPane.Visible = isVisible;
                    taskPane.VisibleChanged += TaskPane_VisibleChanged;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error setting task pane visibility: {ex.Message}");
                }
            }
        }

        private void DisposeTaskPane(PowerPoint.Presentation pres)
        {
            try
            {
                var taskPane = customTaskPanes[pres];
                taskPane.VisibleChanged -= TaskPane_VisibleChanged;
                taskPane.Dispose();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error disposing task pane: {ex.Message}");
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
                        try
                        {
                            pane.VisibleChanged -= TaskPane_VisibleChanged;
                            pane.Visible = taskPaneVisible;
                            pane.VisibleChanged += TaskPane_VisibleChanged;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error setting task pane visibility: {ex.Message}");
                        }
                    }
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            ((PowerPoint.EApplication_Event)this.Application).NewPresentation -= Application_NewPresentation;
            ((PowerPoint.EApplication_Event)this.Application).PresentationOpen -= Application_PresentationOpen;
            ((PowerPoint.EApplication_Event)this.Application).PresentationCloseFinal -= Application_PresentationClose;
            ((PowerPoint.EApplication_Event)this.Application).WindowActivate -= Application_WindowActivate;

            foreach (var pres in customTaskPanes.Keys)
            {
                DisposeTaskPane(pres);
            }
            customTaskPanes.Clear();
        }

        public void ToggleTaskPaneVisibility()
        {
            taskPaneVisible = !taskPaneVisible;
            foreach (var taskPane in customTaskPanes.Values)
            {
                try
                {
                    taskPane.VisibleChanged -= TaskPane_VisibleChanged;
                    taskPane.Visible = taskPaneVisible;
                    taskPane.VisibleChanged += TaskPane_VisibleChanged;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error toggling task pane visibility: {ex.Message}");
                }
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

        private void RetryAction(Action action, int maxRetries = 3, int delayMilliseconds = 1000)
        {
            int attempt = 0;
            while (attempt < maxRetries)
            {
                try
                {
                    action();
                    break;
                }
                catch (COMException ex) when (ex.HResult == unchecked((int)0x8001010A))
                {
                    attempt++;
                    if (attempt >= maxRetries)
                    {
                        throw;
                    }
                    Thread.Sleep(delayMilliseconds);
                }
            }
        }

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
    }
}
