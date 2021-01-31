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
    public class EllipseExam
    {
        [CommandMethod("EllipseDemo()")]
        public void EllipseDemo()
        {
            // 声明一个图形数据库
            Database db = HostApplicationServices.WorkingDatabase;

            // public Ellipse(Point3d center, Vector3d unitNormal, Vector3d majorAxis, double radiusRatio, double startAngle, double endAngle);
            // Ellipse e1 = new Ellipse(); 
            //Ellipse e2 = new Ellipse(new Point3d(100, 100, 0), Vector3d.ZAxis, new Vector3d(100, 0, 0), 0.4, 0, Math.PI); // zAxis 等价于 new Vector3d(0,0,1) 平行于y轴的法向量
            //db.AddEntityToModeSpace(e2);

            // db.AddEllipseToModeSpace(new Point3d(100, 100, 0), 200, 50, 0); //调用封装好的数据
            //db.AddEllipseToModeSpace(new Point3d(20, 20, 0), new Point3d(200, 200, 0), 60);
            db.AddEllipseToModeSpace(new Point3d(100, 100, 0), new Point3d(500, 500, 0));


        }
    }
}
