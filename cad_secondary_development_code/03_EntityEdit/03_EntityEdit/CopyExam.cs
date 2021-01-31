using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AcDoNetTools;

namespace _03_EntityEdit
{
    public class CopyExam
    {
        // 复制图形
        [CommandMethod("CopyDemo")]
        public void CopyDemo()
        {
            Database db = HostApplicationServices.WorkingDatabase;

            Circle c1 = new Circle(new Point3d(100, 100, 0), Vector3d.ZAxis, 100);
            Circle c2 = (Circle)c1.CopyEntity(new Point3d(100, 100, 0), new Point3d(100, 200, 0));
            db.AddEntityToModeSpace(c1);
            Circle c3 = (Circle)c1.CopyEntity(new Point3d(0, 0, 0), new Point3d(-100, 0, 0));
            db.AddEntityToModeSpace(c3);
        }

        // 旋转图形
        [CommandMethod("RotateDemo")]
        public void RotateDemo()
        {
            Database db = HostApplicationServices.WorkingDatabase;

            Line l1 = new Line(new Point3d(100, 100, 0), new Point3d(300, 100, 0));
            Line l2 = new Line(new Point3d(100, 100, 0), new Point3d(300, 100, 0));
            Line l3 = new Line(new Point3d(100, 100, 0), new Point3d(300, 100, 0));

            l1.RotateEntity(new Point3d(100, 100, 0), 30);
            db.AddEntityToModeSpace(l1, l2, l3);

            l2.RotateEntity(new Point3d(0, 0, 0), 60);

            l3.RotateEntity(new Point3d(200, 500, 0), 90);
        }

    }
}
