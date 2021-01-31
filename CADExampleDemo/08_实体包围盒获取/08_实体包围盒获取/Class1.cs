using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _08_实体包围盒获取
{
    public class Class1
    {
        // 获取当前文档和数据库
        Document doc = Application.DocumentManager.MdiActiveDocument;
        Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
        Database db = HostApplicationServices.WorkingDatabase;

        [CommandMethod("GEG")]//获取实体包围盒
        public void GeometricExtentsGet()
        {
            ed.WriteMessage("确定实体包围盒\n");

            SelectionSet sSet = this.SelectSsGet("GetSelection", null, null);

            // 范围对象
            Extents3d extend = new Extents3d();

            // 判断选择集是否为空
            if (sSet != null)
            {
                // 遍历选择对象
                foreach (SelectedObject selObj in sSet)
                {
                    // 确认返回的是合法的SelectedObject对象  
                    if (selObj != null) //
                    {
                        //开启事务处理
                        using (Transaction trans = db.TransactionManager.StartTransaction())
                        {
                            Entity ent = trans.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;
                            // 获取多个实体合在一起的获取其总范围
                            extend.AddExtents(ent.GeometricExtents);

                            trans.Commit();
                        }
                    }
                }
                if (extend != null)
                {
                    // 绘制包围盒
                    Point3d point1 = extend.MinPoint;  // 范围最大点
                    Point3d point2 = extend.MaxPoint;  // 范围最小点
                    Point3d point11 = new Point3d(point1.X, point2.Y, 0);
                    Point3d point22 = new Point3d(point2.X, point1.Y, 0);
                    
                    this.CreatePolyline(new Point3dCollection { point11, point2, point22, point1 }, null, 256, 0, true);

                }

            }
        }


        /// <summary>
        /// 绘制多段线
        /// </summary>
        /// <param name="point3dCollection"></param>
        /// <param name="layerName"></param>
        /// <param name="colorIndex"></param>
        /// <param name="Width"></param>
        /// <param name="Closeflag"></param>
        /// <returns></returns>
        public ObjectId CreatePolyline(Point3dCollection point3dCollection, string layerName, short colorIndex, double Width, bool Closeflag)
        {
           
           
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // 以读模式打开块表 
                BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                // 以写模式打开当前块表记录 model空间
                BlockTableRecord btr = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                // 多段线创建
                Polyline pline = new Polyline();
                //遍历所有的点
                for (int index = 0; index < point3dCollection.Count; index++)
                {
                    pline.AddVertexAt(index, new Point2d(point3dCollection[index].X, point3dCollection[index].Y), 0, 0, 0);
                }
                btr.AppendEntity(pline);
                trans.AddNewlyCreatedDBObject(pline, true);
                pline.ConstantWidth = Width;
                // 添加图层
                if (layerName != null)
                {
                    CreateLayer(layerName, 0);
                    pline.Layer = layerName;
                }
                pline.ColorIndex = colorIndex;
                pline.Closed = Closeflag;
                // 提交事务处理
                trans.Commit();
                return pline.ObjectId;
            }
        }

        /// <summary>
        /// 创建图层
        /// </summary>
        /// <param name="layerName">图层名称</param>
        /// <param name="color">图层颜色</param>
        public void CreateLayer(string layerName, short color)
        {
            
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // 打开层表
                LayerTable lt = trans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                // 判断图层是否存在
                if (lt.Has(layerName) == false)
                {
                    // 层表记录
                    LayerTableRecord ltr = new LayerTableRecord();
                    // 图层颜色设置
                    ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, color);
                    // 图层名称
                    ltr.Name = layerName;

                    // 升级图层表
                    lt.UpgradeOpen();
                    // 添加新图层到图层表 记录事务
                    lt.Add(ltr);
                    trans.AddNewlyCreatedDBObject(ltr, true);
                }
                // 提交事务处理
                trans.Commit();
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
    }
}
