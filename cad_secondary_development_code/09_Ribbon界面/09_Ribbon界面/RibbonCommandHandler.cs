using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace _09_Ribbon界面
{
    public class RibbonCommandHandler : System.Windows.Input.ICommand
    {
        //定义用于确定此命令是否可以在其当前状态下执行的方法。
        public bool CanExecute(object parameter)
        {
            return true;
        }
        public event EventHandler CanExecuteChanged;
        // 定义在调用此命令时调用的方法。
        public void Execute(object parameter)
        {

            if (parameter is RibbonButton)
            {
                RibbonButton btn = (RibbonButton)parameter;
                if (btn.CommandParameter != null)
                {
                    Document doc = Application.DocumentManager.MdiActiveDocument;
                    doc.SendStringToExecute(btn.CommandParameter.ToString(), true, false, false);
                }
            }
        }
    }
}