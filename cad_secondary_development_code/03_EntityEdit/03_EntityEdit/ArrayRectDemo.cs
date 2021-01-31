using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AcDoNetTools;
using Autodesk.AutoCAD.DatabaseServices;
using AcDoNetTools;
using Autodesk.AutoCAD.Geometry;

namespace _03_EntityEdit
{
    public class ArrayRectDemo
    {
        [CommandMethod("ArrayRectDemo1")]
        public void ArrayRectDemo1()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            //ObjectId cirId = db.AddCircleModeSpace(new Point3d(100, 100, 0), 20);

            //cirId.ArrayRectEntity(3, 5, 30, 50);

            Circle c = new Circle(new Point3d(100, 100, 0), Vector3d.ZAxis, 10);

            c.ArrayRectEntity(5, 6, -50, -50);
        }

        [CommandMethod("ArrayRectDemo1")]
        public void ArrayPolarDemo()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            //ObjectId lineId = db.AddLineToModeSpace(new Point3d(100, 100, 0), new Point3d(120, 100, 0));
            //lineId.ArrayPolarEntity(6, 360, new Point3d(100, 100, 0));

            //ObjectId lineId2 = db.AddLineToModeSpace(new Point3d(500, 100, 0), new Point3d(520, 100, 0));
            //lineId.ArrayPolarEntity(6, 360, new Point3d(500, 100, 0));

            Line l = new Line(new Point3d(100, 100, 0), new Point3d(120, 120, 0));
            l.ArrayPolarEntity(7, -270, new Point3d(110, 300, 0));

        }
    }
}
