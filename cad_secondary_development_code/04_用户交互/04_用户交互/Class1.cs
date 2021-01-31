using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AcDoNetTools;
using Autodesk.AutoCAD.DatabaseServices;

namespace _04_用户交互
{
    public class Class1
    {
        [CommandMethod("PromptDemo")]
        public void PromptDemo()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            // 获取命令行窗口
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            //PromptPointResult ppr = ed.GetPoint("\n 请选择一个点：");
            //if (ppr.Status == PromptStatus.OK)
            //{
            //    Point3d p1 = ppr.Value;
            //    ppr = ed.GetPoint("\n 请选择第二个点：");
            //    if (ppr.Status == PromptStatus.OK)
            //    {
            //        Point3d p2 = ppr.Value;
            //        db.AddLineToModeSpace(p1, p2);
            //    }
            //}

            Point3d p1 = new Point3d(0, 0, 0);
            Point3d p2 = new Point3d();
            Point3d p3 = new Point3d();

            PromptPointOptions ppo = new PromptPointOptions("请指定第一个点：");
            ppo.AllowNone = true;
            PromptPointResult ppr = GetPoint(ppo);

            if (ppr.Status == PromptStatus.Cancel) return;

            if (ppr.Status == PromptStatus.OK) p1 = ppr.Value;
            ppo.Message = "请指定第二个点";
            ppo.BasePoint = p1;
            ppo.UseBasePoint = true;
            ppr = GetPoint(ppo);
            if (ppr.Status == PromptStatus.Cancel) return;
            if (ppr.Status == PromptStatus.None) return;
            if (ppr.Status == PromptStatus.OK) p2 = ppr.Value;
            db.AddLineToModeSpace(p1, p2);

            if (ppr.Status == PromptStatus.OK) p1 = ppr.Value;
            ppo.Message = "请指定第三个点";
            ppo.BasePoint = p1;
            ppo.UseBasePoint = true;
            ppr = GetPoint(ppo);
            if (ppr.Status == PromptStatus.Cancel) return;
            if (ppr.Status == PromptStatus.None) return;
            if (ppr.Status == PromptStatus.OK) p3 = ppr.Value;
            db.AddLineToModeSpace(p1, p3);


        }

        private PromptPointResult GetPoint(PromptPointOptions ppo)
        {

            ppo.AllowNone = true;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            return ed.GetPoint(ppo);

        }

        [CommandMethod("FangLine")]
        public void FangLine()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            // 声明一个直线的集合对象
            List<ObjectId> lineList = new List<ObjectId>();

            // 声明一个预备的对象
            Point3d pointStart = new Point3d(100, 100, 0);
            Point3d pointPre = new Point3d(100, 100, 0);

            PromptPointResult ppr = ed.GetPoint2("\n 指定第一个点：");
            if (ppr.Status == PromptStatus.Cancel) return;
            if (ppr.Status == PromptStatus.None) pointPre = pointStart;
            if (ppr.Status == PromptStatus.OK)
            {
                pointStart = ppr.Value;
                pointPre = pointStart;
            }
            // 判断循环是否继续
            bool isC = true;

            while (isC)
            {
                if (lineList.Count == 2)
                {
                    ppr = ed.GetPoint("\n 指定一下点或 [闭合(C)/放弃(U)]:", pointPre, new string[] { "C", "U" });
                }
                else
                {
                    ppr = ed.GetPoint("\n 指定下一点或[放弃(U)]", pointPre, new string[] { "U" });
                }
                ppr = ed.GetPoint("\n 指定下一点或 [放弃(U)]", pointPre, new string[] { "U" });
                Point3d pointNext; // 用于接收下一点坐标
                if (ppr.Status == PromptStatus.Cancel) return;
                if (ppr.Status == PromptStatus.None) return;
                if (ppr.Status == PromptStatus.OK)
                {
                    pointNext = ppr.Value;
                    lineList.Add(db.AddLineToModeSpace(pointPre, pointNext));
                    pointPre = pointNext;
                }
                if (ppr.Status == PromptStatus.Keyword)
                {
                    switch (ppr.StringResult)
                    {
                        case "U":
                            if (lineList.Count == 0) // 没有直线
                            {
                                pointStart = new Point3d(100, 100, 0);
                                pointPre = new Point3d(100, 100, 0);
                                ppr = ed.GetPoint2("\n 指定第一个点：");
                                if (ppr.Status == PromptStatus.Cancel) return;
                                if (ppr.Status == PromptStatus.None) pointPre = pointStart;
                                if (ppr.Status == PromptStatus.OK)
                                {
                                    pointStart = ppr.Value;
                                    pointPre = pointStart;
                                }
                            }
                            else if (lineList.Count > 0)
                            {
                                int count = lineList.Count;
                                ObjectId lId = lineList.ElementAt(count - 1);
                                pointPre = this.GetLineStartPoint(lId);
                                lineList.RemoveAt(count - 1);
                                lId.EraseEntity();

                            }
                            break;
                        case "C":
                            lineList.Add(db.AddLineToModeSpace(pointPre, pointStart));
                            isC = false;
                            break;
                    }
                }
            }
        }


        [CommandMethod("CircleTest")]
        public void CirecleTest()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            Point3d center = new Point3d();
            double radius = 0;
            PromptPointResult ppr = ed.GetPoint("\n 请指定圆心：");
            if (ppr.Status == PromptStatus.OK)
            {
                center = ppr.Value;
            }
            #region
            //PromptDistanceOptions pdo = new PromptDistanceOptions("\n 指定圆上的一个点:");
            //pdo.BasePoint = center;
            //pdo.UseBasePoint = true;
            //PromptDoubleResult pdr = ed.GetDistance(pdo);
            //if (pdr.Status == PromptStatus.OK)
            //{
            //    radius = pdr.Value;
            //}
            //db.AddCircleModeSpace(center, radius);
            #endregion

            CircleJig jCircle = new CircleJig(center);
            //PromptPointResult pr = (PromptPointResult)ed.Drag(jCircle);  // 返回接收对象
            //if (pr.Status = PromptStatus.OK)
            //{
            //    Point3d pt = pr.Value;
            //    db.AddCircleModeSpace(center, pt.GetDistanceBetweenTwoPoint(center));
            //}
            PromptResult pr = ed.Drag(jCircle);
            if (pr.Status == PromptStatus.OK)
            {
                db.AddEntityToModeSpace(jCircle.GetEntity());
            }

        }

        /// <summary>
        /// 获取直线起点坐标
        /// </summary>
        /// <param name="lineId">直线对象的ObjectId</param>
        /// <returns></returns>
        private Point3d GetLineStartPoint(ObjectId lineId)
        {
            Point3d startPoint;
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                Line line = (Line)lineId.GetObject(OpenMode.ForRead);
                startPoint = line.StartPoint;
            }
            return startPoint;
        }
    }
}
