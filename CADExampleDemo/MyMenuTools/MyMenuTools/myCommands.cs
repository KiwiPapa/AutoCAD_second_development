// (C) Copyright 2020 by  
//
using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using System.Collections.Generic;
using Exception = Autodesk.AutoCAD.Runtime.Exception;
using System.Windows.Forms;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.Windows;
using OpenFileDialog = Autodesk.AutoCAD.Windows.OpenFileDialog;

[assembly: CommandClass(typeof(MyMenuTools.MyCommands))]

namespace MyMenuTools
{

    public class MyCommands
    {

        // 创建个人类OtherToosClass对象 用来调用里面的方法
        OtherToolsClass otl = new OtherToolsClass();


        [CommandMethod("TBA")]
        public void TextBatchAlign()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("\n 欢迎使用文字批量对齐程序！！！");

          

            // 选择一个基准点
            Point3d point = otl.SelectPoint("\n>>>>请选择基准点！！");
            // 创建一个 TypedValue 数组来定义过滤器条件
            //TypedValue[] typeValue = new TypedValue[1];
            TypedValue[] typeValue = new TypedValue[2];
            // 过滤条件 只选择单行文本
            typeValue.SetValue(new TypedValue(0, "TEXT"), 0);
            // 文本内容
            //typeValue.SetValue(new TypedValue((int)DxfCode.Text, "数据智能笔记A"), 1);
            // 选择集 区域手动选择方式
            SelectionSet sSet = otl.SelectSsGet("GetSelection", null, typeValue);
            // 判断选择集是否为空
            if (sSet != null)
            {
                // 如果选择集不为空 遍历选择图元对象 
                foreach (SelectedObject sSetObj in sSet)
                {
                    // 开启事务处理
                    using (Transaction trans = db.TransactionManager.StartTransaction())
                    {
                        // 单行文本对象 打开方式为写
                        DBText dbText = trans.GetObject(sSetObj.ObjectId, OpenMode.ForWrite) as DBText;
                        // 垂直方向左对齐
                        dbText.HorizontalMode = TextHorizontalMode.TextLeft;
                        // 判断对齐方式是否是左对齐 
                        if (dbText.HorizontalMode != TextHorizontalMode.TextLeft)
                        {
                            // 对齐点
                            Point3d aliPoint = dbText.AlignmentPoint;
                            ed.WriteMessage("\n" + aliPoint.ToString());
                            // 位置点
                            Point3d position = dbText.Position;
                            dbText.AlignmentPoint = new Point3d(point.X, position.Y, 0);
                        }
                        // 如果是左对齐只需要调整插入点
                        else
                        {
                            Point3d position = dbText.Position;
                            dbText.Position = new Point3d(point.X, position.Y, 0);
                        }
                        trans.Commit();
                    }
                }
            }
        }


        [CommandMethod("TBR")]
        public void TextBatchReplace()
        {

            List<KeyValuePair<string, string>> liststring = new List<KeyValuePair<string, string>>();
            // 添加需要替换和替换后的文本 
            liststring.Add(new KeyValuePair<string, string>("需要替换的文本", "替换后文本"));

            // 打开图形数据库
            Database db = HostApplicationServices.WorkingDatabase;
            // 打开事务处理
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // 以只读方式打开块表
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                // 以写的方式打开块表记录
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // 在块表记录中遍历对象
                foreach (ObjectId id in btr)
                {
                    DBObject ent = trans.GetObject(id, OpenMode.ForWrite);
                    // 判断是否是单行文字
                    if (ent is DBText)
                    {
                        DBText dbText = ent as DBText;

                        int newstring = otl.ReplaceString(dbText.TextString.Trim(), liststring);
                        // 如果正确索引到该文本字符串
                        if (newstring >= 0)
                        {
                            // 执行替换
                            dbText.TextString = dbText.TextString.Replace(liststring[newstring].Key.Trim(), liststring[newstring].Value.Trim());
                        }
                    }
                    // 如果是多行文本
                    else if (ent is MText)
                    {
                        MText mText = ent as MText;
                        int newstring = otl.ReplaceString(mText.Contents.Trim(), liststring);
                        if (newstring >= 0)
                        {
                            mText.Contents = mText.Contents.Replace(liststring[newstring].Key.Trim(), liststring[newstring].Value.Trim());
                        }
                    }
                }
                trans.Commit();
            }

        }


        //输出文本到Excel和Txt 打开窗口

        [CommandMethod("CTE")] //CAD启动界面
        public void CADTextExport()
        {
            myForm myfrom = new myForm();
            //  Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(myfrom); //在CAD里头显示界面  模态显示

            Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(myfrom); //在CAD里头显示界面  非模态显示
                                                                                         // myfrom.sh


        }



         [CommandMethod("CTES")]
        static public void CADTablebyExcelSheet()
        {
            const string dlName = "从Excel导入表格";

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // 文件打开窗口 
            OpenFileDialog ofd =
                new OpenFileDialog(
                    "选择需要链接的Excel表格文档！",
                    null,
                    "xls;xlsx",
                    "Excel链接到CAD",
                    OpenFileDialog.OpenFileDialogFlags.DoNotTransferRemoteFiles
                    );

            System.Windows.Forms.DialogResult dr = ofd.ShowDialog();

            if (dr != System.Windows.Forms.DialogResult.OK)
            {
                return;
            }

            ed.WriteMessage("\n选择到的文件为\"{0}\".", ofd.Filename);

            PromptPointResult ppr = ed.GetPoint("\n请选择表格插入点: ");

            if (ppr.Status != PromptStatus.OK)
            {
                return;
            }

            // 数据链接管理对象
            DataLinkManager dlm = db.DataLinkManager;

            // 判断数据链接是否已经存在 如果存在移除
            ObjectId dlId = dlm.GetDataLink(dlName);
            if (dlId != ObjectId.Null)
            {
                dlm.RemoveDataLink(dlId);
            }

            // 创建并添加新的数据链接

            DataLink dl = new DataLink();
            dl.DataAdapterId = "AcExcel";
            dl.Name = dlName;
            dl.Description = "Excel fun with Through the Interface";
            dl.ConnectionString = ofd.Filename;
            dl.UpdateOption |= (int)UpdateOption.AllowSourceUpdate;

            dlId = dlm.AddDataLink(dl);

            // 开启事务处理
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                trans.AddNewlyCreatedDBObject(dl, true);

                BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;


                // 新建表格对象
                Table tb = new Table();
                tb.TableStyle = db.Tablestyle;
                tb.Position = ppr.Value;
                tb.SetDataLink(0, 0, dlId, true);
                tb.GenerateLayout();

                BlockTableRecord btr = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                btr.AppendEntity(tb);
                trans.AddNewlyCreatedDBObject(tb, true);
                trans.Commit();
            }

            // 强制恢复显示表格
            ed.Regen();
        }





        /// <summary>
        /// 由Excel表格中数据来更新CAD中的table
        /// 也就是说 excel数据变化后保存 然后运行该命令来更新CAD中的表格数据
        /// </summary>


        [CommandMethod("UCTFE")]
        static public void UpdateCADTableFromExcel()
        {
            // 获取当前文档 数据库 命令行
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // 定义一个交互提示选项 并设置提示内容
            PromptEntityOptions peo = new PromptEntityOptions("\n请选择需要更新的表格：");
            peo.SetRejectMessage("\n选取实体不是表格！");
            // 选择实体类型只能是表格
            peo.AddAllowedClass(typeof(Table), false);

            // 获取实体结果返回
            PromptEntityResult per = ed.GetEntity(peo);
            // 判断选取状态
            if (per.Status != PromptStatus.OK)
                return;
            // 开启事务处理
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // 获取表格对象
                    DBObject obj = trans.GetObject(per.ObjectId, OpenMode.ForRead) as DBObject;

                    Table tb = obj as Table;

                    // 检查是否是表格

                    if (tb != null)
                    {
                        // 升级打开方式 以写的方式
                        tb.UpgradeOpen();

                        // 从Excel表格数据更新链接数据

                        ObjectId dlId = tb.GetDataLink(0, 0);
                        DataLink dl = trans.GetObject(dlId, OpenMode.ForWrite) as DataLink;

                        // 更新数据
                        dl.Update(UpdateDirection.SourceToData, UpdateOption.None);

                        // 从链接数据更新到CAD表格
                        tb.UpdateDataLink(UpdateDirection.SourceToData, UpdateOption.None);
                    }
                    trans.Commit();
                    ed.WriteMessage("\n已经成功从Excel电子表中更新CAD表格数据！");
                }

                catch (Exception ex)
                {
                    ed.WriteMessage("\nException: {0}", ex.Message);
                }
            }
        }



        /// <summary>
        /// 由CAD表格数据更新Excle中的数据
        /// </summary>

        [CommandMethod("UEFCT")]
        static public void UpdateExcelFromCADTable()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;


            PromptEntityOptions peo = new PromptEntityOptions("\n 请选择你用来更新Excel的表格：");
            peo.SetRejectMessage("\n选取实体不是表格！");

            peo.AddAllowedClass(typeof(Table), false);
            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK)
            {
                return;
            }

            // 开启事务处理
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    DBObject obj = trans.GetObject(per.ObjectId, OpenMode.ForRead) as DBObject;

                    Table tb = obj as Table;

                    if (tb != null)
                    {
                        tb.UpgradeOpen();

                        // 从CAD表格数据更新链接数据

                        tb.UpdateDataLink(UpdateDirection.DataToSource, UpdateOption.ForceFullSourceUpdate);

                        // 由链接数据更新Excel表格

                        ObjectId dlId = tb.GetDataLink(0, 0);
                        DataLink dl = trans.GetObject(dlId, OpenMode.ForWrite) as DataLink;

                        dl.Update(UpdateDirection.DataToSource, UpdateOption.ForceFullSourceUpdate);
                    }
                    trans.Commit();
                    ed.WriteMessage("\n 已经成功由CAD表格更新Excel表格数据！");
                }

                catch (Exception ex)

                {

                    ed.WriteMessage(

                      "\nException: {0}",

                      ex.Message

                    );
                }
            }
        }

    }

}
