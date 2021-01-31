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

namespace _07_采样标注
{
    public class Class1
    {
        [CommandMethod("CaiYang")]
        public void CaiYang()
        {
            // 声明数据库对象
            Database db = HostApplicationServices.WorkingDatabase;
            // 交互窗口
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            // 提示用户选择基准线和采样线
            PromptEntityResult perBase = ed.GetEntity("\n 请选择基准线：");
            if (perBase.Status != PromptStatus.OK) return; // 如果没有选择
            ObjectId baseEntityId = perBase.ObjectId;
            PromptEntityResult perCurve = ed.GetEntity("\n 请选择采样线：");
            if (perCurve.Status != PromptStatus.OK) return;
            // 获取基准线和采样线图形实体对象
            Entity baseEntity = this.GetEntity(db, perBase.ObjectId);
            Entity curve = this.GetEntity(db, perCurve.ObjectId);
            // 如果是直线对象
            if (baseEntity is Line)
            {
                Line baseLine = (Line)baseEntity; // 基准线对象
                List<Point3d> divPoints = this.GetBaseLineDivPoints(baseLine, 100.0);  // 定数等分
                //List<Point3d> divPoints2 = this.GetBaseLineDivPoints(baseLine, 100); // 定距等分
                List<Line> divLines = this.GetDivLines(divPoints.ToArray(), baseLine.Angle + Math.PI * 0.5, 100);
                divLines = this.ModifyDivLines(divLines.ToArray(), curve);
                List<DBText> dimTexts = this.GetDivDimTexts(divLines.ToArray(),5,-15);

                this.AddEntity(db, divLines.ToArray());
                this.AddEntity(db, dimTexts.ToArray());
                
            }
            else
            {
                ed.WriteMessage("\n 基准线必须为直线！");
            }


        }


        /// <summary>
        /// 获取等分线的标注文字
        /// </summary>
        /// <param name="lines">等分线数组</param>
        /// <param name="height">文字高度</param>
        /// <param name="dist">文字离等分线起点距离</param>
        /// <returns></returns>
        private List<DBText> GetDivDimTexts(Line[] lines, double height,double dist)
        {
            List<DBText> texts = new List<DBText>();
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                for (int i = 0; i < lines.Length; i++)
                {
                    DBText text = new DBText();  // 单行文本对象9
                    if (lines[i].Angle >= 0 && lines[i].Angle <= Math.PI)
                    {
                        text.TextString = string.Format("{0:N}", lines[i].Length); // 文本内容 线条长度 保留两位小数
                        text.Rotation = lines[i].Angle;  // 旋转
                        text.HorizontalMode = TextHorizontalMode.TextRight;  // 水平右对齐
                    }
                    else
                    {

                        text.TextString = "-" + string.Format("{0:N}", lines[i].Length); // 文本内容 线条长度 保留两位小数
                        text.Rotation = lines[i].Angle + Math.PI;  // 旋转
                        text.HorizontalMode = TextHorizontalMode.TextLeft;  // 水平左对齐
                    }
                    text.Position = this.PolarPoint(lines[i].StartPoint,dist,lines[i].Angle); // 文本位置
                    text.Height = height; // 文字高度
                   
                    text.VerticalMode = TextVerticalMode.TextVerticalMid; // 垂直对齐                   
                    text.AlignmentPoint = text.Position; // 对齐点
                    texts.Add(text);
                }
            }
               
            return texts;
        }

        /// <summary>
        /// 根据相交修改等分线终点坐标
        /// </summary>
        /// <param name="lines">等分线数组</param>
        /// <param name="ent">与之相交的实体</param>
        private List<Line> ModifyDivLines(Line[] lines, Entity ent)
        {
            List<Line> divLines = new List<Line>();
            for (int i = 0; i < lines.Length; i++)
            {
                Point3dCollection insertPoints = new Point3dCollection();
                lines[i].IntersectWith(ent, Intersect.ExtendThis,insertPoints, IntPtr.Zero, IntPtr.Zero);
                if (insertPoints.Count > 0)
                {
                    lines[i].EndPoint = insertPoints[0];
                    divLines.Add(lines[i]);
                }
            }
            return divLines;
        }

        /// <summary>
        /// 获取等分线对象
        /// </summary>
        /// <param name="points">等分线起点</param>
        /// <param name="angle">等分线与X轴正方向夹角</param>
        /// <param name="length">长度</param>
        /// <returns>List<Line></returns>
        private List<Line> GetDivLines(Point3d[] points, double angle, double length)
        {
            List<Line> lines = new List<Line>();
            for (int i = 0; i < points.Length; i++)
            {
                lines.Add(new Line(points[i], this.PolarPoint(points[i], length, angle)));
            }
            return lines;
        }

        /// <summary>
        /// 添加图形到图形数据库
        /// </summary>
        /// <param name="db">图形数据库</param>
        /// <param name="ents">图形对象数据</param>
        /// <returns></returns>
        private List<ObjectId> AddEntity(Database db, params Entity[] ents)
        {
            List<ObjectId> entIds = new List<ObjectId>();
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                for (int i = 0; i < ents.Length; i++)
                {
                    entIds.Add(btr.AppendEntity(ents[i]));
                    trans.AddNewlyCreatedDBObject(ents[i],true);
                }
                trans.Commit();
            }
            return entIds;
               
        }

        /// <summary>
        /// 获取直线的定数等分点
        /// </summary>
        /// <param name="line">直线对象</param>
        /// <param name="divnum">等分数量</param>
        /// <returns></returns>
        private List<Point3d> GetBaseLineDivPoints(Line line, int divnum)
        {
            return this.GetBaseLineDivPoints(line, line.Length / divnum);
        }

        /// <summary>
        /// 获取直线的定距等分点
        /// </summary>
        /// <param name="line">直线对象</param>
        /// <param name="divDistence">等分距离</param>
        private List<Point3d> GetBaseLineDivPoints(Line line, double divDistence)
        {
            List<Point3d> points = new List<Point3d>();
            int divNum = (int)(line.Length / divDistence);
            Point3d startPoint = line.StartPoint;
            double angle = line.Angle;
            for (int i = 0; i < divNum + 1; i++)
            {
                points.Add(this.PolarPoint(startPoint, divDistence * i, angle));
            }
            if (divDistence*divNum != line.Length)
            {
                points.Add(line.EndPoint);
            }
            return points;
        }


        /// <summary>
        /// 获取Polar点坐标
        /// </summary>
        /// <param name="startPoint">起点坐标</param>
        /// <param name="dist">终点到起点距离</param>
        /// <param name="angle">起点到终点的线与X轴正方向的夹角</param>
        /// <returns></returns>
        private Point3d PolarPoint(Point3d startPoint, double dist, double angle)
        {
            double X = startPoint.X + dist * Math.Cos(angle);
            double Y = startPoint.Y + dist * Math.Sin(angle);
            return new Point3d(X, Y, 0);
        }


        /// <summary>
        /// 获取图形实体对象
        /// </summary>
        /// <param name="db">图形数据库</param>
        /// <param name="entId">图形实体的ObjectId</param>
        /// <returns></returns>
        private Entity GetEntity(Database db, ObjectId entId)
        {
            Entity ent;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                ent = (Entity)entId.GetObject(OpenMode.ForRead);
            }
            return ent;
        }
    }
}
