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

namespace _05_多段线批量等间距法线绘制
{
    public class Class1
    {
        // 获取当前活动文档
        Document doc = Application.DocumentManager.MdiActiveDocument;

        [CommandMethod("BCED")]

        public void BatchCurveEquiDistance()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("欢迎使用多段线等间距绘制法线程序\n");
            // 单选图元
            ObjectId curveId = this.SelectEntity("\n请选择一个多段线对象", typeof(Polyline));

            // 判断是否选取了多段线对象
            if (curveId != null)
            {
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    Curve curEnt = trans.GetObject(curveId, OpenMode.ForRead) as Curve;
                    // 获取多段线终点
                    Point3d endPoint = curEnt.EndPoint;
                    // 多段线长度
                    double lineLength = curEnt.GetDistAtPoint(endPoint);
                    // 间隔距离
                    double dis = 10;
                    // 第一个点距离起点距离
                    double myDis = dis;
                    while (myDis < lineLength)
                    {
                        //获取距离起点 mydis 的多段线上的一个点
                        Point3d point = curEnt.GetPointAtDist(myDis);
                        //一阶导数  即切线向量
                        Vector3d curVector = curEnt.GetFirstDerivative(point);
                        //获取切线方向的垂线向量 （从切线得到法线）
                        Vector3d curPervector = curVector.GetPerpendicularVector();
                        //向量转化成标准单位向量
                        curPervector = curPervector.GetNormal();
                        //起始点+ 单位向量*距离 就是距离起始点 距离向量长度的 一个点
                        Point3d newPoint = point.Add(curPervector * 10);
                        // 绘制直线
                        this.CreateLine(point, newPoint, 1);
                        // 间隔递增
                        myDis = myDis + dis;
                    }
                    trans.Commit();
                }
            }
        }


        /// <summary>
        /// 单选实体
        /// </summary>
        /// <param name="message"></param>
        /// <param name="enttype"></param>
        /// <returns></returns>
        public ObjectId SelectEntity(string message, Type enttype)
        {
            PromptEntityOptions opts = new PromptEntityOptions(message);
            if (enttype != null)
            {
                opts.SetRejectMessage("选择图元类型错误，请重新选择" + enttype.ToString());
                opts.AddAllowedClass(enttype, true);
            }
            PromptEntityResult res = doc.Editor.GetEntity(opts);
            return res.ObjectId;
        }


        /// <summary>
        /// 创建直线
        /// </summary>
        /// <param name="strartPoint">直线起点</param>
        /// <param name="endPoint">直线终点</param>
        /// <param name="Color">直线颜色</param>
        /// <param name="layerName">图层名称</param>
        public void CreateLine(Point3d strartPoint, Point3d endPoint, int Color)
        {

            // 获得数据库
            Database db = doc.Database;
            // 开启事务处理
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // 以读模式打开块表
                BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                // 以写模式打开当前块表记录 model空间
                BlockTableRecord btr = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                Line line = new Line(strartPoint, endPoint);
                line.ColorIndex = Color;

                // 将创建的直线对象添加到Model空间 并进行事务登记
                btr.AppendEntity(line);
                trans.AddNewlyCreatedDBObject(line, true);
                // 提交事务处理
                trans.Commit();
            }
        }
    }
}
