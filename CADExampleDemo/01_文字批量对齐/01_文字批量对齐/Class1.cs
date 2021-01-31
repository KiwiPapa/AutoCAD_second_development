using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using Autodesk.AutoCAD.DatabaseServices;               // (Database, DBPoint, Line, Spline) 
using Autodesk.AutoCAD.Geometry;                       //(Point3d, Line3d, Curve3d) 
using Autodesk.AutoCAD.ApplicationServices;            // (Application, Document) 
using Autodesk.AutoCAD.Runtime;                        // (CommandMethodAttribute, RXObject, CommandFlag) 
using Autodesk.AutoCAD.EditorInput;                    //(Editor, PromptXOptions, PromptXResult)
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.Colors;


namespace _01_文字批量对齐
{
    public class Class1
    {
        [CommandMethod("TBA")]
        public void TextBatchAlign()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("\n 欢迎使用文字批量对齐程序！！！");

            // 选择一个基准点
            Point3d point = this.SelectPoint("\n>>>>请选择基准点！！");
            // 创建一个 TypedValue 数组来定义过滤器条件
            //TypedValue[] typeValue = new TypedValue[1];
            TypedValue[] typeValue = new TypedValue[2]; 
            // 过滤条件 只选择单行文本
            typeValue.SetValue(new TypedValue(0, "TEXT"), 0);
            typeValue.SetValue(new TypedValue((int)DxfCode.Text, "数据智能笔记A"), 1);
            // 选择集 区域手动选择方式
            SelectionSet sSet = this.SelectSsGet("GetSelection", null, typeValue);
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

        


        /// <summary>
        /// 选择点
        /// </summary>
        /// <param name="message">输入提示信息</param>
        /// <returns></returns>
        public Point3d SelectPoint(string message)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            PromptPointResult res;
            PromptPointOptions opts = new PromptPointOptions("");
            // 提示信息
            opts.Message = message;
            res = doc.Editor.GetPoint(opts);
            Point3d startpoint = res.Value;
            return startpoint;
        }

        /// <summary>
        /// 获取选择集
        /// </summary>
        /// <param name="selectStr">选择方式</param>
        /// <param name="point3dCollection">选择点集合</param>
        /// <param name="typedValue">过滤参数</param>
        /// <returns></returns>
        public SelectionSet SelectSsGet(string selectStr, Point3dCollection point3dCollection, TypedValue[] typedValue)
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            // 将过滤条件赋值给SelectionFilter对象
            SelectionFilter selfilter = null;
            if (typedValue != null)
            {
                selfilter = new SelectionFilter(typedValue);
            }
            // 请求在图形区域选择对象
            PromptSelectionResult psr;
            if (selectStr == "GetSelection")  // 提示用户从图形文件中选取对象
            {
                psr = ed.GetSelection(selfilter);
            }
            else if (selectStr == "SelectAll") //选择当前空间内所有未锁定及未冻结的对象
            {
                psr = ed.SelectAll(selfilter);
            }
            else if (selectStr == "SelectCrossingPolygon") //选择由给定点定义的多边形内的所有对象以及与多边形相交的对象。多边形可以是任意形状，但不能与自己交叉或接触。
            {
                psr = ed.SelectCrossingPolygon(point3dCollection, selfilter);
            }
            // 选择与选择围栏相交的所有对象。围栏选择与多边形选择类似，所不同的是围栏不是封闭的， 围栏同样不能与自己相交
            else if (selectStr == "SelectFence") 
            {
                psr = ed.SelectFence(point3dCollection, selfilter);
            }
            // 选择完全框入由点定义的多边形内的对象。多边形可以是任意形状，但不能与自己交叉或接触
            else if (selectStr == "SelectWindowPolygon")
            {
                psr = ed.SelectWindowPolygon(point3dCollection, selfilter);
            }
            else if (selectStr == "SelectCrossingWindow")  //选择由两个点定义的窗口内的对象以及与窗口相交的对象
            {
                Point3d point1 = point3dCollection[0];
                Point3d point2 = point3dCollection[1];
                psr = ed.SelectCrossingWindow(point1, point2, selfilter);
            }
            else if (selectStr == "SelectWindow") // 选择完全框入由两个点定义的矩形内的所有对象。
            {
                Point3d point1 = point3dCollection[0];
                Point3d point2 = point3dCollection[1];
                psr = ed.SelectCrossingWindow(point1, point2, selfilter);
            }
            else
            {
                return null;
            }

            // 如果提示状态OK，表示对象已选
            if (psr.Status == PromptStatus.OK)
            {
                SelectionSet sSet = psr.Value;
                ed.WriteMessage("Number of objects selected: " + sSet.Count.ToString() + "\n");// 打印选择对象数量
                return sSet;
            }
            else
            {
                // 打印选择对象数量
                ed.WriteMessage("Number of objects selected 0 \n");
                return null;
            }
        }
    }
}
