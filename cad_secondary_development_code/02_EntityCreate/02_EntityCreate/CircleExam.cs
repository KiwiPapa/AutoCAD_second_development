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
    class CircleExam
    {
        [CommandMethod("CircleDemo")]
        public void CircleDemo()
        {
            //Circle c1 = new Circle();
            //c1.Center = new Point3d(50, 50, 0);
            //c1.Radius = 50;

            //Circle c2 = new Circle(new Point3d(100, 100, 0), new Vector3d(0, 0, 1),50); // 圆心 法向量 半径
            Database db = HostApplicationServices.WorkingDatabase;
            //db.AddEntityToModeSpace(c1, c2);
            db.AddCircleModeSpace(new Point3d(100, 100, 0), 100); // 圆心半径画圆

            db.AddCircleModeSpace(new Point3d(200, 100, 0), new Point3d(300, 100, 0)); // 两点画圆

            db.AddCircleModeSpace(new Point3d(400, 100, 0), new Point3d(600, 100, 0), new Point3d(600, 200, 0)); // 三点画圆
        }
    }
}
