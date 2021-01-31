using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _02_EntityCreate
{
    public class PolyLineExam
    {
        [CommandMethod("PolyLineDemo")]
        public void PolyLineDemo()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            #region
            //Polyline pl = new Polyline();
            //Point2d p1 = new Point2d(100, 100);
            //Point2d p2 = new Point2d(200, 100);
            //Point2d p3 = new Point2d(200, 200);

            //pl.AddVertexAt(0, p1, 0, 0, 0);
            //pl.AddVertexAt(1, p2, 0, 0, 0);
            //pl.AddVertexAt(2, p3, 0, 0, 0);
            //pl.Closed = true; // 闭合
            //pl.ConstantWidth = 10; // 多段线宽度

           
            //db.AddEntityToModeSpace(pl);

            //db.AddPolyLineToModeSpace(false, 0, new Point2d(10, 10), new Point2d(50, 10), new Point2d(100, 20)); //参数： 是否闭合 线宽 两个平面点
            //db.AddPolyLineToModeSpace(true, 0, new Point2d(10, 50), new Point2d(50, 50), new Point2d(100, 70));
            #endregion

            //Polyline pLine = new Polyline();
            //Point2d p1 = new Point2d(100, 100);
            //Point2d p2 = new Point2d(100, 200);
            //pLine.AddVertexAt(0, p1, 1, 0, 0);
            //pLine.AddVertexAt(1, p2, 0, 0, 0);
            //Database db = HostApplicationServices.WorkingDatabase;
            //db.AddEntityToModeSpace(pLine);

            // 矩形绘制测试
            
            //db.AddRectToModeSpace(new Point2d(0, 0), new Point2d(100, 100));
            //db.AddRectToModeSpace(new Point2d(200, 200), new Point2d(100, 100));
            //db.AddRectToModeSpace(new Point2d(500, 500), new Point2d(100, 100));

            db.AddPolygonToModeSpace(new Point2d(100, 100), 50, 3, 90);
            db.AddPolygonToModeSpace(new Point2d(200, 100), 50, 4, 45);
            db.AddPolygonToModeSpace(new Point2d(300, 100), 50, 5, 90);
            db.AddPolygonToModeSpace(new Point2d(400, 100), 50, 6, 0);
            db.AddPolygonToModeSpace(new Point2d(500, 100), 50, 12, 0);
        }

    }
}
