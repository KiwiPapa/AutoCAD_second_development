using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Collections.Specialized;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

using Autodesk.AutoCAD.Interop;  //引用需要添加 Autodesk.AutoCAD.Interop.dll  CAD安装目录内
using System.IO;


[assembly: ExtensionApplication(typeof(MyMenuTools.AcadNetApp))] //启动时加载工具栏，注意typeof括号里的类库名

namespace MyMenuTools
{
    //添加项目类引用
    public class AcadNetApp : Autodesk.AutoCAD.Runtime.IExtensionApplication
    {
        //重写初始化
        public void Initialize()
        {
            //加载后初始化的程序放在这里 这样程序一加载DLL文件就会执行
            Document doc = Application.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\n加载程序中...........\n");
            //加载菜单栏需
            AddMenu();
        }

        //重写结束
        public void Terminate()
        {
            // do somehing to cleanup resource
        }

        //加载菜单栏
        public void AddMenu()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("添加菜单栏成功！>>>>>>>>>>>>>>");
            AcadApplication acadApp = Application.AcadApplication as AcadApplication;

            //创建建菜单栏的对象
            AcadPopupMenu myMenu = null;  

            // 创建菜单
            if (myMenu == null)
            {
                // 菜单名称
                myMenu = acadApp.MenuGroups.Item(0).Menus.Add("个人专用菜单工具(2020.07)");

                myMenu.AddMenuItem(myMenu.Count, "文字批量替换", "TBR "); //每个命令后面有一个空格，相当于咱们输入命令按空格
                myMenu.AddMenuItem(myMenu.Count, "文字批量对齐", "TBA ");
                myMenu.AddMenuItem(myMenu.Count, "输出文本", "CTE ");
                //开始加子菜单栏
                AcadPopupMenu subMenu = myMenu.AddSubMenu(myMenu.Count, "CADLinkToEXcel");  //子菜单对象
                subMenu.AddMenuItem(myMenu.Count, "Excel表格导入", "CTES ");
                subMenu.AddMenuItem(myMenu.Count, "由Excel数据更新CAD表格数据", "UCTFE ");
                subMenu.AddMenuItem(myMenu.Count, "由CAD表格数据更新Excel表格数据", "UEFCT ");
                myMenu.AddSeparator(myMenu.Count); //加入分割符号
                //结束加子菜单栏


            }

            // 菜单是否显示  看看已经显示的菜单栏里面有没有这一栏
            bool isShowed = false;  //初始化没有显示
            foreach (AcadPopupMenu menu in acadApp.MenuBar)  //遍历现有所有菜单栏
            {
                if (menu == myMenu)
                {
                    isShowed = true;
                    break;
                }
            }

            // 显示菜单 加载自定义的菜单栏
            if (!isShowed)
            {
                myMenu.InsertInMenuBar(acadApp.MenuBar.Count);
            }

        }        
    }    
}
