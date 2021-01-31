using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _05_选择集
{
    public class Class1
    {
        [CommandMethod("SeleDemo")]
        public void SeleDemo()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            //PromptSelectionResult psr = ed.SelectAll(); // 选择窗口中所有图形
            // 只选择窗口中的圆形
            TypedValue[] values = new TypedValue[]
            {
                new TypedValue((int)DxfCode.Start,"Circle")
            };
            SelectionFilter filter = new SelectionFilter(values);// 过滤器
            //PromptSelectionResult psr = ed.GetSelection(filter); // 选取图形对象
            //if(psr.Status == PromptStatus.OK)
            //{
            //    SelectionSet sSet = psr.Value;
            //    this.ChangeColor(sSet);

            //}

            PromptSelectionResult psr = ed.GetSelection(filter);
            List<ObjectId> ids = new List<ObjectId>();
            if (psr.Status == PromptStatus.OK)
            {
                SelectionSet sSet = psr.Value;
                List<Point3d> points = new List<Point3d>();
                points = this.GetPoint(sSet);
                for (int i = 0; i < points.Count; i++)
                {
                    PromptSelectionResult ss1 = ed.SelectCrossingWindow(points.ElementAt(i), points.ElementAt(i));
                    ids.AddRange(ss1.Value.GetObjectIds());
                }
            }
            this.ChangeColor(ids);
        }

       
        private List<Point3d> GetPoint(SelectionSet sSet)
        {
            List<Point3d> points = new List<Point3d>();
            ObjectId[] ids = sSet.GetObjectIds();
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                for (int i = 0; i < ids.Length; i++)
                {
                    Entity ent = (Entity)ids[i].GetObject(OpenMode.ForRead);
                    Point3d center = ((Circle)ent).Center;
                    double radius = ((Circle)ent).Radius;
                    points.Add(new Point3d(center.X + radius, center.Y, center.Z));
                }
                trans.Commit();
            }
            return points;
        }
                

        /// <summary>
        /// 改变颜色
        /// </summary>
        /// <param name="sSet">选取对象</param>
        private void ChangeColor(SelectionSet sSet)
        {
            ObjectId[] ids = sSet.GetObjectIds();
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                for (int i = 0; i < ids.Length; i++)
                {
                    Entity ent = (Entity)ids[i].GetObject(OpenMode.ForWrite);
                    ent.ColorIndex = 1; // 红色
                }
                trans.Commit();
            }
        }

        private void ChangeColor(List<ObjectId> ids)
        {
            
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                for (int i = 0; i < ids.Count; i++)
                {
                    Entity ent = (Entity)ids[i].GetObject(OpenMode.ForWrite);
                    ent.ColorIndex = 3; // 红色
                }
                trans.Commit();
            }
        }
    }
}
