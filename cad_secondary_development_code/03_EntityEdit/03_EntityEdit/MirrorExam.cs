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
    public class MirrorExam
    {
        [CommandMethod("MirrorDemo")]
        public void MirrorDemo()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            Circle c1 = new Circle(new Point3d(100, 100, 0), Vector3d.ZAxis, 50);
            ObjectId cId =  db.AddEntityToModeSpace(c1);

            Entity Ent = cId.MirrorEntity(new Point3d(200, 100, 0), new Point3d(200, 300, 0), false);
            db.AddEntityToModeSpace(Ent);

            Circle c2 = new Circle(new Point3d(100, 100, 0), Vector3d.ZAxis, 50);
            // 删除原图形
            Entity ent2 = c2.MirrorEntity(new Point3d(200, 100, 0), new Point3d(200, 300, 0), true);
            // 不删除原图形
            Entity ent3 = c2.MirrorEntity(new Point3d(200, 100, 0), new Point3d(200, 300, 0), false);
            db.AddEntityToModeSpace(c2, ent2);

        }
    }
}
