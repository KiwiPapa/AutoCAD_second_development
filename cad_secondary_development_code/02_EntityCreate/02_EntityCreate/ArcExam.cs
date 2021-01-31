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
    public class ArcExam
    {
        [CommandMethod("ArcTest")]

        public void ArcTest()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            //db.AddEntityToModeSpace(arc1, arc2, arc3);  // 已经封装好的事务处理
            Point3d startPoint = new Point3d(100, 100, 0);
            Point3d endPoint = new Point3d(200, 200, 0);
            Point3d pointOnArc = new Point3d(150, 100, 0);
            CircularArc3d cArc = new CircularArc3d(startPoint, pointOnArc, endPoint);

            double radius = cArc.Radius;
            Point3d center = cArc.Center;
            Vector3d cs = center.GetVectorTo(startPoint);
            Vector3d ce = center.GetVectorTo(endPoint);
            Vector3d xvector = new Vector3d(1, 0, 0);
            double startAngle = cs.Y > 0 ? cs.GetAngleTo(xvector) : -xvector.GetAngleTo(cs);
            double endAngle = ce.Y > 0 ? xvector.GetAngleTo(ce) : -xvector.GetAngleTo(ce);
            Arc arc = new Arc(center, radius, startAngle, endAngle);
            db.AddEntityToModeSpace(arc);
        }
    }
}
