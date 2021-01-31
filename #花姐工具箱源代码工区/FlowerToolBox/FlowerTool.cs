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
using FlowerTools;

namespace FlowerTools
{
    public class FlowerTool
    {
        [CommandMethod("rrr")]
        // 道路边线在上面
        public void Flag1()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            // 获取命令行窗口
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            PromptPointOptions ppo = new PromptPointOptions("请指定第一个点：");
            ppo.AllowNone = true;
            PromptPointResult ppr = GetPoint(ppo);
            if (ppr.Status == PromptStatus.Cancel) return;
            if (ppr.Status == PromptStatus.OK)
            {
                // p1 为点击的第1个点
                Point3d p1 = ppr.Value;
                ppo.Message = "请指定第二个点";
                ppo.BasePoint = p1;
                ppo.UseBasePoint = true;
                ppr = GetPoint(ppo);
                if (ppr.Status == PromptStatus.Cancel) return;
                if (ppr.Status == PromptStatus.None) return;
                if (ppr.Status == PromptStatus.OK)
                {
                    // p2 为点击的第2个点
                    Point3d p2 = ppr.Value;
                    // 利用p1和p2 计算p3、p4和p5
                    double valueX = p2.X - p1.X;
                    double valueY = p2.Y - p1.Y;
                    double radian = Math.Atan2(valueY, valueX);
                    double X3 = p1.X + 20 * Math.Cos(radian);
                    double Y3 = p1.Y + 20 * Math.Sin(radian);
                    Point3d p3 = new Point3d(X3, Y3, 0);

                    double X4 = p1.X + 28 * Math.Cos(radian);
                    double Y4 = p1.Y + 28 * Math.Sin(radian);
                    Point3d p4 = new Point3d(X4, Y4, 0);

                    double X5 = p1.X + 35 * Math.Cos(radian);
                    double Y5 = p1.Y + 35 * Math.Sin(radian);
                    Point3d p5 = new Point3d(X5, Y5, 0);

                    // 垂直的两条线
                    double radian2 = radian + Math.PI / 2;
                    double X6 = p3.X + 15 * Math.Cos(radian2);
                    double Y6 = p3.Y + 15 * Math.Sin(radian2);
                    Point3d p6 = new Point3d(X6, Y6, 0);

                    double X7 = p4.X + 15 * Math.Cos(radian2);
                    double Y7 = p4.Y + 15 * Math.Sin(radian2);
                    Point3d p7 = new Point3d(X7, Y7, 0);
                    // db.AddLineToModeSpace(p1, p2);
                    db.AddLineToModeSpace(p1, p5);
                    db.AddLineToModeSpace(p3, p6);
                    db.AddLineToModeSpace(p4, p7);

                    ObjectId C1 = db.AddCircleModeSpace(p1, 0.755);
                    db.HatchEnity(HatchTools.HatchPatternName.solid, 1, 45, C1);
                    ObjectId C2 = db.AddCircleModeSpace(p2, 0.755);
                    db.HatchEnity(HatchTools.HatchPatternName.solid, 1, 45, C2);

                    // 文字定位的点
                    double XX3 = p1.X + 21 * Math.Cos(radian);
                    double YY3 = p1.Y + 21 * Math.Sin(radian);
                    Point3d pp3 = new Point3d(XX3, YY3, 0);

                    double XX4 = p1.X + 29 * Math.Cos(radian);
                    double YY4 = p1.Y + 29 * Math.Sin(radian);
                    Point3d pp4 = new Point3d(XX4, YY4, 0);

                    double XX6 = pp3.X + 8 * Math.Cos(radian2);
                    double YY6 = pp3.Y + 8 * Math.Sin(radian2);
                    Point3d pp6 = new Point3d(XX6, YY6, 0);

                    double XX7 = pp4.X + 8 * Math.Cos(radian2);
                    double YY7 = pp4.Y + 8 * Math.Sin(radian2);
                    Point3d pp7 = new Point3d(XX7, YY7, 0);

                    // 通过事务添加文字
                    using (Transaction trans = db.TransactionManager.StartTransaction())
                    {
                        DBText text0 = new DBText(); // 新建单行文本对象
                        text0.Position = pp6; // 设置文本位置 
                        text0.TextString = "燃气管道"; // 设置文本内容
                        text0.Height = 3;  // 设置文本高度
                        text0.WidthFactor = 0.7;
                        
                        text0.Rotation = radian2 + Math.PI;  // 设置文本选择角度
                        text0.HorizontalMode = TextHorizontalMode.TextCenter; // 设置对齐方式
                        text0.AlignmentPoint = text0.Position; //设置对齐点
                        db.AddEntityToModeSpace(text0);
                        DBText text1 = new DBText(); // 新建单行文本对象
                        text1.Position = pp7;
                        text1.TextString = "道路边线";
                        text1.Height = 3;
                        text1.WidthFactor = 0.7;
                        text1.Rotation = radian2 + Math.PI;
                        text1.HorizontalMode = TextHorizontalMode.TextCenter;
                        text1.AlignmentPoint = text1.Position; // 设置对齐点
                        db.AddEntityToModeSpace(text1);

                        trans.Commit();
                    }
                }
            }

        }

        [CommandMethod("ggg")]
        // 燃气管道在上面
        public void Flag2()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            // 获取命令行窗口
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            PromptPointOptions ppo = new PromptPointOptions("请指定第一个点：");
            ppo.AllowNone = true;
            PromptPointResult ppr = GetPoint(ppo);
            if (ppr.Status == PromptStatus.Cancel) return;
            if (ppr.Status == PromptStatus.OK)
            {
                // p1 为点击的第1个点
                Point3d p1 = ppr.Value;
                ppo.Message = "请指定第二个点";
                ppo.BasePoint = p1;
                ppo.UseBasePoint = true;
                ppr = GetPoint(ppo);
                if (ppr.Status == PromptStatus.Cancel) return;
                if (ppr.Status == PromptStatus.None) return;
                if (ppr.Status == PromptStatus.OK)
                {
                    // p2 为点击的第2个点
                    Point3d p2 = ppr.Value;
                    // 利用p1和p2 计算p3、p4和p5
                    double valueX = p2.X - p1.X;
                    double valueY = p2.Y - p1.Y;
                    double radian = Math.Atan2(valueY, valueX);
                    double X3 = p1.X + 20 * Math.Cos(radian);
                    double Y3 = p1.Y + 20 * Math.Sin(radian);
                    Point3d p3 = new Point3d(X3, Y3, 0);

                    double X4 = p1.X + 28 * Math.Cos(radian);
                    double Y4 = p1.Y + 28 * Math.Sin(radian);
                    Point3d p4 = new Point3d(X4, Y4, 0);

                    double X5 = p1.X + 35 * Math.Cos(radian);
                    double Y5 = p1.Y + 35 * Math.Sin(radian);
                    Point3d p5 = new Point3d(X5, Y5, 0);

                    // 垂直的两条线
                    double radian2 = radian + Math.PI / 2;
                    double X6 = p3.X + 15 * Math.Cos(radian2);
                    double Y6 = p3.Y + 15 * Math.Sin(radian2);
                    Point3d p6 = new Point3d(X6, Y6, 0);

                    double X7 = p4.X + 15 * Math.Cos(radian2);
                    double Y7 = p4.Y + 15 * Math.Sin(radian2);
                    Point3d p7 = new Point3d(X7, Y7, 0);
                    // db.AddLineToModeSpace(p1, p2);
                    db.AddLineToModeSpace(p1, p5);
                    db.AddLineToModeSpace(p3, p6);
                    db.AddLineToModeSpace(p4, p7);

                    ObjectId C1 = db.AddCircleModeSpace(p1, 0.755);
                    db.HatchEnity(HatchTools.HatchPatternName.solid, 1, 45, C1);
                    ObjectId C2 = db.AddCircleModeSpace(p2, 0.755);
                    db.HatchEnity(HatchTools.HatchPatternName.solid, 1, 45, C2);

                    // 文字定位的点
                    double XX3 = p1.X + 21 * Math.Cos(radian);
                    double YY3 = p1.Y + 21 * Math.Sin(radian);
                    Point3d pp3 = new Point3d(XX3, YY3, 0);

                    double XX4 = p1.X + 29 * Math.Cos(radian);
                    double YY4 = p1.Y + 29 * Math.Sin(radian);
                    Point3d pp4 = new Point3d(XX4, YY4, 0);

                    double XX6 = pp3.X + 8 * Math.Cos(radian2);
                    double YY6 = pp3.Y + 8 * Math.Sin(radian2);
                    Point3d pp6 = new Point3d(XX6, YY6, 0);

                    double XX7 = pp4.X + 8 * Math.Cos(radian2);
                    double YY7 = pp4.Y + 8 * Math.Sin(radian2);
                    Point3d pp7 = new Point3d(XX7, YY7, 0);

                    // 通过事务添加文字
                    using (Transaction trans = db.TransactionManager.StartTransaction())
                    {
                        DBText text0 = new DBText(); // 新建单行文本对象
                        text0.Position = pp6; // 设置文本位置 
                        text0.TextString = "道路边线"; // 设置文本内容
                        text0.Height = 3;  // 设置文本高度
                        text0.WidthFactor = 0.7;
                        text0.Rotation = radian2 + Math.PI;  // 设置文本选择角度
                        text0.HorizontalMode = TextHorizontalMode.TextCenter; // 设置对齐方式
                        text0.AlignmentPoint = text0.Position; //设置对齐点
                        db.AddEntityToModeSpace(text0);

                        DBText text1 = new DBText(); // 新建单行文本对象
                        text1.Position = pp7;
                        text1.TextString = "燃气管道";
                        text1.Height = 3;
                        text1.WidthFactor = 0.7;
                        text1.Rotation = radian2 + Math.PI;
                        text1.HorizontalMode = TextHorizontalMode.TextCenter;
                        text1.AlignmentPoint = text1.Position; // 设置对齐点
                        db.AddEntityToModeSpace(text1);

                        trans.Commit();
                    }
                }
            }

        }

        [CommandMethod("ddd")]
        // 燃气管道在上面
        public void DingXiangChuanYue()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            // 获取命令行窗口
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            PromptPointOptions ppo = new PromptPointOptions("请指定第一个点：");
            ppo.AllowNone = true;
            PromptPointResult ppr = GetPoint(ppo);
            if (ppr.Status == PromptStatus.Cancel) return;
            if (ppr.Status == PromptStatus.OK)
            {
                // p1 为点击的第1个点
                Point3d p1 = ppr.Value;
                ppo.Message = "请指定第二个点";
                ppo.BasePoint = p1;
                ppo.UseBasePoint = true;
                ppr = GetPoint(ppo);
                if (ppr.Status == PromptStatus.Cancel) return;
                if (ppr.Status == PromptStatus.None) return;
                if (ppr.Status == PromptStatus.OK)
                {
                    // p2 为点击的第2个点
                    Point3d p2 = ppr.Value;
                    ppo.Message = "请指定第三个点";
                    ppo.BasePoint = p2;
                    ppo.UseBasePoint = true;
                    ppr = GetPoint(ppo);
                    if (ppr.Status == PromptStatus.Cancel) return;
                    if (ppr.Status == PromptStatus.None) return;
                    if (ppr.Status == PromptStatus.OK)
                    {
                        // p3 为点击的第3个点
                        Point3d p3 = ppr.Value;

                        double XX = p3.X - p1.X;
                        double YY = p3.Y - p1.Y;
                        double radian = Math.Atan2(YY, XX);

                        if ((radian <= (Math.PI / 2)) && ((-Math.PI / 2) <= radian))
                        {
                            // 利用p1、p2、p3 计算p4
                            double valueX = p3.X + 75;
                            double valueY = p3.Y;
                            double X4 = valueX;
                            double Y4 = valueY;
                            Point3d p4 = new Point3d(X4, Y4, 0);
                            db.AddLineToModeSpace(p3, p4);
                            db.AddLineToModeSpace(p1, p3);
                            db.AddLineToModeSpace(p2, p3);

                            // 画实心圆点
                            ObjectId C1 = db.AddCircleModeSpace(p1, 0.755);
                            db.HatchEnity(HatchTools.HatchPatternName.solid, 1, 45, C1);
                            ObjectId C2 = db.AddCircleModeSpace(p2, 0.755);
                            db.HatchEnity(HatchTools.HatchPatternName.solid, 1, 45, C2);

                            // 计算p1和p2的直线距离
                            double distance_p1_p2 = Math.Sqrt(Math.Pow(p1.X - p2.X, 2) + Math.Pow(p1.Y - p2.Y, 2));
                            distance_p1_p2 = (int)distance_p1_p2;

                            // 文字的定位
                            Point3d p5 = new Point3d(p3.X + 37.0, p3.Y + 1.0, 0);

                            // 通过事务添加文字
                            using (Transaction trans = db.TransactionManager.StartTransaction())
                            {
                                DBText text0 = new DBText(); // 新建单行文本对象
                                text0.Position = p5; // 设置文本位置 
                                text0.TextString = "定向钻穿越PE100 SDR11 dnXX-" + distance_p1_p2.ToString() + "m（平面距离）"; // 设置文本内容
                                text0.Height = 3;  // 设置文本高度
                                text0.WidthFactor = 0.7;
                                text0.HorizontalMode = TextHorizontalMode.TextCenter; // 设置对齐方式
                                text0.AlignmentPoint = text0.Position; //设置对齐点
                                db.AddEntityToModeSpace(text0);

                                trans.Commit();
                            }
                        }else
                        {
                            // 利用p1、p2、p3 计算p4
                            double valueX = p3.X - 75;
                            double valueY = p3.Y;
                            double X4 = valueX;
                            double Y4 = valueY;
                            Point3d p4 = new Point3d(X4, Y4, 0);
                            db.AddLineToModeSpace(p3, p4);
                            db.AddLineToModeSpace(p1, p3);
                            db.AddLineToModeSpace(p2, p3);

                            // 画实心圆点
                            ObjectId C1 = db.AddCircleModeSpace(p1, 0.755);
                            db.HatchEnity(HatchTools.HatchPatternName.solid, 1, 45, C1);
                            ObjectId C2 = db.AddCircleModeSpace(p2, 0.755);
                            db.HatchEnity(HatchTools.HatchPatternName.solid, 1, 45, C2);

                            // 计算p1和p2的直线距离
                            double distance_p1_p2 = Math.Sqrt(Math.Pow(p1.X - p2.X, 2) + Math.Pow(p1.Y - p2.Y, 2));
                            distance_p1_p2 = (int)distance_p1_p2;

                            // 文字的定位
                            Point3d p5 = new Point3d(p3.X - 37.0, p3.Y + 1.0, 0);

                            // 通过事务添加文字
                            using (Transaction trans = db.TransactionManager.StartTransaction())
                            {
                                DBText text0 = new DBText(); // 新建单行文本对象
                                text0.Position = p5; // 设置文本位置 
                                text0.TextString = "定向钻穿越PE100 SDR11 dnXX-" + distance_p1_p2.ToString() + "m（平面距离）"; // 设置文本内容
                                text0.Height = 3;  // 设置文本高度
                                text0.WidthFactor = 0.7;
                                text0.HorizontalMode = TextHorizontalMode.TextCenter; // 设置对齐方式
                                text0.AlignmentPoint = text0.Position; //设置对齐点
                                db.AddEntityToModeSpace(text0);

                                trans.Commit();
                            }
                        }
                            
                    }
                }
            }

        }

        private PromptPointResult GetPoint(PromptPointOptions ppo)
        {

            ppo.AllowNone = true;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            return ed.GetPoint(ppo);

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
