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

using System.Windows.Forms;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace _02_批量统计线段长度
{
    public class Class1
    {
        [CommandMethod("BCLL")]
        public void BatchCountLineLength()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("\n 欢迎使用批量统计线段长度程序！！！");
            // 过滤删选条件设置 过滤器
            TypedValue[] typedValues = new TypedValue[2];
            typedValues.SetValue(new TypedValue(0, "*LINE"), 0); // * 通配符 所有Line
            //typedValues.SetValue(new TypedValue(8, "Lay01"), 1); // 图层设置
            // //如果颜色是BYlayer  对应的值就是256
            typedValues.SetValue(new TypedValue((int)DxfCode.Color, "256"), 1);
            SelectionSet sSet = this.SelectSsGet("GetSelection", null, typedValues);

            double sumLen = 0;
            // 判断是否选取了对象
            if(sSet != null)
            {
                // 遍历选择集
                foreach (SelectedObject sSObj in sSet)
                {
                    // 确认返回的是合法的SelectedObject对象  
                    if(sSObj != null)
                    {
                        // 开启事务处理
                        using (Transaction trans = db.TransactionManager.StartTransaction())
                        {
                            Curve curEnt = trans.GetObject(sSObj.ObjectId, OpenMode.ForRead) as Curve;
                            // 调整文字位置点和对齐点
                            Point3d endPoint = curEnt.EndPoint;
                            // GetDisAtPoint 用于返回起点到终点的长度 传入终点坐标
                            double lineLength = curEnt.GetDistAtPoint(endPoint);
                            ed.WriteMessage("\n" + lineLength.ToString());
                            sumLen = sumLen + lineLength;
                            trans.Commit();
                        }
                    }
                }
            }
            ed.WriteMessage("\n 线段总长为： ", sumLen.ToString());
        }


        public void ProgressBarManaged()
        {
            //新建进度条类
            int pmlimit = 100;
            ProgressMeter pm = new ProgressMeter();
            pm.Start("正在执行，请稍候...");
           // Ptcadlib mycad = new Ptcadlib();
            //把进度条分成最多少粉
            pm.SetLimit(pmlimit);
            for (int i = 0; i <= pmlimit; i++)
            {
                //休眠5ms
                System.Threading.Thread.Sleep(5);
                /*
                 * 正在写程序过程中，就是把程序写在这里面
                 */
                //Point3d pt = new Point3d(i * 3, i * 3, 0);
                //mycad.CreateText(i.ToString(), pt, 2, 0, null, 256);
                //进度条增加一份
                pm.MeterProgress(); //每运行依次进度条增加一格，进度条有多少格设置在pm.SetLimit(pmlimit);处
                //Application.DoEvents();
            }
            pm.Stop();
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
