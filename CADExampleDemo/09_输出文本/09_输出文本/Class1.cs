using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _09_输出文本
{
    public class Class1
    {
        [CommandMethod("CTE")] //CAD启动界面
        public void CADTextExport()
        {
            myForm myfrom = new myForm();
            //  Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(myfrom); //在CAD里头显示界面  模态显示

            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(myfrom); //在CAD里头显示界面  非模态显示
                                                                                         // myfrom.sh


        }
    }
}
