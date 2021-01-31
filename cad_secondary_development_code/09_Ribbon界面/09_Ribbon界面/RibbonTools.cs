using Autodesk.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace _09_Ribbon界面
{
    public static partial class RibbonTools
    {
        /// <summary>
        /// 添加Ribbon选项卡
        /// </summary>
        /// <param name="ribbonCtrl">Ribbon控制器</param>
        /// <param name="title">选项卡标题</param>
        /// <param name="ID">选项卡ID</param>
        /// <param name="isActive">是否置为当前</param>
        /// <returns>RibbonTab</returns>
        public static RibbonTab AddTab(this RibbonControl ribbonCtrl, string title, string ID, bool isActive)
        {
            RibbonTab tab = new RibbonTab();
            tab.Title = title;
            tab.Id = ID;
            ribbonCtrl.Tabs.Add(tab);
            tab.IsActive = isActive;
            return tab;
        }

        /// <summary>
        /// 添加面板
        /// </summary>
        /// <param name="tab">Ribbon选项卡</param>
        /// <param name="title">面板标题</param>
        /// <returns>RibbonPanelSource</returns>
        public static RibbonPanelSource AddPanel(this RibbonTab tab, string title)
        {
            RibbonPanelSource panelSource = new RibbonPanelSource();
            panelSource.Title = title;
            RibbonPanel ribbonPanel = new RibbonPanel();
            ribbonPanel.Source = panelSource;
            tab.Panels.Add(ribbonPanel);
            return panelSource;
        }


        /// <summary>
        /// 给面板添加下拉组合按钮
        /// </summary>
        /// <param name="panelSource"></param>
        /// <param name="text"></param>
        /// <param name="size"></param>
        /// <param name="orient"></param>
        /// <returns></returns>
        public static RibbonSplitButton AddSplitButton(this RibbonPanelSource panelSource, string text, RibbonItemSize size, Orientation orient)
        {
            RibbonSplitButton splitBtn = new RibbonSplitButton();
            splitBtn.Text = text;
            splitBtn.ShowText = true;
            splitBtn.Size = size;
            splitBtn.ShowImage = true;
            splitBtn.Orientation = orient;
            panelSource.Items.Add(splitBtn);
            return splitBtn;
        }
    }
}
