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
    public class MoveExam
    {
        [CommandMethod("MoveDemo")]
        public void MoveDemo()
        {
            Database db = HostApplicationServices.WorkingDatabase;

            Circle c1 = new Circle(new Point3d(100, 100, 0), new Vector3d(0, 0, 1), 50);
            Circle c2 = new Circle(new Point3d(100, 100, 0), new Vector3d(0, 0, 1), 50);
            Circle c3 = new Circle(new Point3d(100, 100, 0), new Vector3d(0, 0, 1), 50);
            Point3d p1 = new Point3d(100, 100, 0);
            Point3d p2 = new Point3d(200, 300, 0);

            //c2.Center = new Point3d(c2.Center.X + p2.X - p1.X, c2.Center.Y + p2.Y - p1.Y,0);
            //db.AddEntityToModeSpace(c1, c2);

            

           
            db.AddEntityToModeSpace(c1,c2);
            c2.MoveEntity(p1, p2);
            c3.MoveEntity(p1, p2);

            db.AddEntityToModeSpace(c3);

        }
    }
}
