using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using Autodesk.AutoCAD.DatabaseServices;          // (Database, DBPoint, Line, Spline) 
using Autodesk.AutoCAD.Runtime;                   // (CommandMethodAttribute, RXObject, CommandFlag) 
using Autodesk.AutoCAD.EditorInput;              //(Editor, PromptXOptions, PromptXResult)
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;

using System.Linq;   //这个必须引用 系统自带
using System.IO; 

namespace _04_LINQ多段线按颜色分类
{
    public class Class1
    {
        // 打开数据库
        Database db = HostApplicationServices.WorkingDatabase;
        // 命令行窗口
        Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
        [CommandMethod("LPLS")]
        public void LinqPolyLineSort()   // 多段线分类
        {          
            // 选择集对象
            SelectionSet sSet = this.SelectSsGet("GetSelection", null, null);
            // 判断是否选取了对象
            if (sSet != null && sSet.Count > 0)
            {
                // 选择集对象Id集合
                ObjectIdCollection objIdColl = new ObjectIdCollection(sSet.GetObjectIds());

                // Linq语句查询
                /*
                 * var是c#的一个关键字，申明一般位置变量，特别适合包含LINQ查询的结果
                 * var关键字告诉编译器，让编译器查询的语句自己判断查询结果是一个什么类型（所以特别适合LINQ）
                 * 这样，查询前不需要声明 从LINQ查询返回对象的类型，编译器自己能判断出来
                */
                var querys =
                        from ObjectId n in objIdColl                    //from 指定数据源 n代表数据源里的某一项
                        where n.ObjectClass.DxfName == "LWPOLYLINE"     //where指定查询条件  LWPOLYLINE 多段线
                        select n;                                       //select指定查询结果中包含哪些元素

                // 开启事务处理
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    // 遍历查询结果
                    foreach (var id in querys)
                    {
                        Polyline pLine = trans.GetObject(id, OpenMode.ForRead) as Polyline;
                        if (pLine.ColorIndex == 5)  // 蓝色
                        {
                            ed.WriteMessage("这是蓝色多段线！");
                            
                        }
                        else if (pLine.ColorIndex == 2)  // 黄色
                        {
                            ed.WriteMessage("这是黄色多段线！");
                            
                        }
                        else if (pLine.ColorIndex == 1)  // 红色
                        {

                            ed.WriteMessage("这是红色多段线！");
                        }
                    }
                }
            }
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


        public void GetPolylineMidPoint()
        {
            
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                
                PromptEntityResult per = ed.GetEntity("请选择多段线");
                if(per.Status == PromptStatus.OK)
                {
                    DBObject obj = trans.GetObject(per.ObjectId, OpenMode.ForRead);
                    {
                        Polyline pLine = obj as Polyline;
                        // 顶点数
                        int vn = pLine.NumberOfVertices;
                        for (int i = 0; i < vn; i++)
                        {
                            Point3d point = pLine.GetPoint3dAt(i - 1);
                            double vBulge = pLine.GetBulgeAt(i);
                            if (vBulge != 0)
                            {
                                Point3d midPoint = pLine.GetPointAtParameter(i + 0.5);
                                ed.WriteMessage("\n 多段线中点为： " + midPoint.ToString());
                            }
                        }
                    }
                }
                trans.Commit();
                trans.Dispose();
            }
        }
    }
}
