using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AcDoNetTools;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;

namespace _03_EntityEdit
{
    public class ColorExam
    {
        [CommandMethod("EditDemo")]
        public void EditDemo()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            //db.AddCircleModeSpace(new Point3d(100, 100, 0), 50);
            Circle c1 = new Circle(new Point3d(100, 100, 0), new Vector3d(0, 0, 1), 50);
            Circle c2 = new Circle(new Point3d(200, 100, 0), new Vector3d(0, 0, 1), 50);
            c1.ColorIndex = 1;
            c2.Color = Color.FromRgb(23, 156, 255);
            db.AddEntityToModeSpace(c1, c2);                 
        }

    }
}
