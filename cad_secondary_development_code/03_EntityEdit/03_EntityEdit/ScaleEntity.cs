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
    public class ScaleEntity
    {
        [CommandMethod("ScaleDemo")]
        public void ScaleDemo()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            Circle c1 = new Circle(new Point3d(100, 100, 0), Vector3d.ZAxis, 50);
            db.AddEntityToModeSpace(c1);
            Circle c2 = new Circle(new Point3d(200, 100, 0), Vector3d.ZAxis, 50);
            c2.ScaleEntity(new Point3d(200, 100, 0), 0.5); // 以（200,100,0）为基点 缩放比例为0,5缩放
            db.AddEntityToModeSpace(c2);
            Circle c3 = new Circle(new Point3d(100, 100, 0), Vector3d.ZAxis, 50);
            db.AddEntityToModeSpace(c3);
            c3.ScaleEntity(new Point3d(0, 0, 0), 2);
        }


        [CommandMethod("EraseDemo")]
        public void EraseDemo()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            Circle c1 = new Circle(new Point3d(100, 100, 0), Vector3d.ZAxis, 50);
            Circle c2 = new Circle(new Point3d(200, 100, 0), Vector3d.ZAxis, 50);
            db.AddEntityToModeSpace(c1,c2);
            c2.ObjectId.EraseEntity();
        }
    }
}
