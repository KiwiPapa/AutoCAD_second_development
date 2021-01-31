using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _01_环境测试_智能数据笔记
{
    public class Class1
    {
        [CommandMethod("TestEnv")] // 添加命令标识符

        public void TestEnv() 
        {
            // 声明命令行对象
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            // 向命令行输出一段文字
            ed.WriteMessage("智能数据笔记（1）：CAD二次开发环境测试！");
        }


    }
}
