// (C) Copyright 2020 by  
//
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Specialized;
using Autodesk.AutoCAD.DatabaseServices;// (Database, DBPoint, Line, Spline) 
using Autodesk.AutoCAD.Geometry;//(Point3d, Line3d, Curve3d) 
using Autodesk.AutoCAD.ApplicationServices;// (Application, Document) 
using Autodesk.AutoCAD.Runtime;// (CommandMethodAttribute, RXObject, CommandFlag) 
using Autodesk.AutoCAD.EditorInput;//(Editor, PromptXOptions, PromptXResult)
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.PlottingServices;  //打印命名空间
using System.Text.RegularExpressions;  //正则表达式
using Autodesk.AutoCAD.Interop;  //引用需要添加 Autodesk.AutoCAD.Interop.dll  CAD安装目录内
using System.IO;
using Autodesk.AutoCAD.Colors;

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(_10CAD自动绘图并保存.MyCommands))]

namespace _10CAD自动绘图并保存
{

    public class MyCommands
    {
        //获取当前的活动文档 
        private Document doc;
        //获取命令行窗口 
        private Editor ed;
        

        /// <summary>
        /// 重新定义自己当前的文档，当文档切换时使用
        /// </summary>
        /// <param name="acDoc"></param>
        public void OperatingCadInitialize(Document doc)
        {

            this.doc = doc;
            ed = doc.Editor;//当前的编辑器对象，命令行对象
        }

        
        [CommandMethod("ADAS", CommandFlags.Modal)]
        public void AutoDwgAndSave() 
        {
            //通过系统变量获取样板的路径
            string sLocalRoot = AcadApp.GetSystemVariable("LOCALROOTPREFIX") as string;  
            string sTemplatePath = sLocalRoot + "Template\\acad.dwt";


            //新建一个文档
            Document m_doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.Add(sTemplatePath);  
                                                                                                                   
            Database acCurDb = m_doc.Database;

            //初始化类 并且更新需要在哪个document生成图形
            this.OperatingCadInitialize(m_doc);  //这里一定要初始化赋值

            //操作非当前文档，建立先锁定文档                                      
            using (DocumentLock docLock = m_doc.LockDocument())
            {
                Database db = doc.Database;
                // 新建表格
                Table table = new Table();
                table.SetSize(10, 5); // 表格大小
                table.SetRowHeight(10); // 设置行高
                table.SetColumnWidth(50); // 设置列宽
                table.Columns[0].Width = 20; // 设置第一列宽度为20
                table.Position = new Point3d(100, 100, 0); // 设置插入点
                table.Cells[0, 0].TextString = "测试表格数据统计"; // 表头
                table.Cells[0, 0].TextHeight = 6; //设置文字高度
                Color color = Color.FromColorIndex(ColorMethod.ByAci, 3); // 声明颜色
                table.Cells[0, 0].BackgroundColor = color; // 设置背景颜色
                color = Color.FromColorIndex(ColorMethod.ByAci, 1);
                table.Cells[0, 0].ContentColor = color; //内容颜色
                //开启事务处理
                using (Transaction trans = acCurDb.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    btr.AppendEntity(table);
                    trans.AddNewlyCreatedDBObject(table, true);
                    trans.Commit();
                }
            }

            //判断该文件是否存在
            string strCableName = "Testdoc"; // 保存文件名
            if (!File.Exists("D:" + "\\" + strCableName + ".dwg"))
            {
                //然后另存为并且关闭
                m_doc.Database.SaveAs("D:" + "\\" + strCableName + ".dwg", DwgVersion.Current);  //另存为 这里需要命名
                m_doc.CloseAndDiscard();                                                         //关闭

            }
            else
            {
                File.Delete("D:" + "\\" + strCableName + ".dwg");
                m_doc.Database.SaveAs("D:" + "\\" + strCableName + ".dwg", DwgVersion.Current);
                m_doc.CloseAndDiscard();            
            }
        }
    }
}
