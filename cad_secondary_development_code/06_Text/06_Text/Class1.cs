using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AcDoNetTools;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;

namespace _06_Text
{
    public class Class1
    {
        [CommandMethod("TextDemo")]
        public void TextDemo()
        {
            Database db = HostApplicationServices.WorkingDatabase;

            // 绘制一些沿Y轴增长的直线
            Line[] line = new Line[10];
            Point3d[] pt1 = new Point3d[10];
            Point3d[] pt2 = new Point3d[10];
            for (int i = 0; i < line.Length; i++)
            {
                pt1[i] = new Point3d(50, 50 + 20 * i, 0);
                pt2[i] = new Point3d(150, 50 + 20 * i, 0);
                line[i] = new Line(pt1[i], pt2[i]);
            }
            db.AddEntityToModeSpace(line);

            
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                DBText text0 = new DBText(); // 新建单行文本对象
                text0.Position = pt1[0]; // 设置文本位置 
                text0.TextString = "你好 数据智能笔记！text0"; // 设置文本内容
                text0.Height = 5;  // 设置文本高度
                text0.Rotation = Math.PI * 0.5;  // 设置文本选择角度
                text0.IsMirroredInX = true; // 在X轴镜像
                text0.HorizontalMode = TextHorizontalMode.TextCenter; // 设置对齐方式
                text0.AlignmentPoint = text0.Position; //设置对齐点
                db.AddEntityToModeSpace(text0);

                DBText text1 = new DBText(); // 新建单行文本对象
                text1.Position = pt1[1];
                text1.TextString = "你好 数据智能笔记！text1";
                text1.Height = 10;
                text1.Rotation = Math.PI * 0.1;
                text1.IsMirroredInY = true; // 在Y轴镜像
                text1.HorizontalMode = TextHorizontalMode.TextLeft;
                //text1.AlignmentPoint = text1.Position; // 设置对齐点
                db.AddEntityToModeSpace(text1);

                trans.Commit();
            }

            

            
        }


        [CommandMethod("TextDemo2")]
        public void TextDemo2()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            DBText text = new DBText();
            text.Position = new Point3d(100, 100, 0);
            //text.TextString = "面积≈%%U100%%U㎡";
            text.TextString = "数据智能笔记" + TextTools.TextSpecialSymbol.Alpha + TextTools.TextSpecialSymbol.Tolerance + "0.006"
                + TextTools.TextSpecialSymbol.Underline + "C# CAD二次开发" + TextTools.TextSpecialSymbol.Underline;
            text.Height = 10;
            db.AddEntityToModeSpace(text);
        }


        [CommandMethod("MTextDemo")]
        public void MTextDemo()
        {
            Database db = HostApplicationServices.WorkingDatabase;

            MText mtext = new MText();  // 声明多行文本对象
            mtext.Location = new Point3d(100, 100, 0); // 设置位置
            mtext.Contents = "智能数据笔记 CAD二次开发系列笔记"; // 设置文本内容
            mtext.TextHeight = 10; //设置文本高度 
            mtext.Width = 3; // 文本框宽度
            mtext.Height = 5; // 文本框高度

            MText mtext2 = new MText();
            mtext.Location = new Point3d(100, 100, 0);
            mtext.Contents = "222智能数据笔记 \nCAD二次开发系列笔记";
            mtext.TextHeight = 20;

            db.AddEntityToModeSpace(mtext, mtext2);

        }


        [CommandMethod("MTextDemo2")]
        public void MTextDemo2()
        {
            Database db = HostApplicationServices.WorkingDatabase;

            MText mtext = new MText();
            //mtext.Contents = "\\A1;Φ20{\\H0.5x;\\SH7/P7;}";

            //mtext.Contents = string.Format("\\A1;{0}{1}\\H{2}x;\\S{3}{4}{5};{6}","Φ20","{",0.2,"数据智能笔记","#","YZK","}");

            
            mtext.Contents = TextTools.StackMtext(TextTools.TextSpecialSymbol.Diameter, 0.4, "+0.5", TextTools.MTextStackType.Tolerance, "-0.1");
            mtext.Location = new Point3d(100, 100, 0);
            mtext.TextHeight = 10;
            db.AddEntityToModeSpace(mtext);
        }


        [CommandMethod("PickMTextDemo")]
        public void PickMTextDemo()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            PromptEntityResult per = ed.GetEntity("\n 请选择多行文字");
            if (per.Status != PromptStatus.OK) return;  // 没有选择就直接返回

            Entity ent;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                ent = (Entity)per.ObjectId.GetObject(OpenMode.ForRead); 
            }
            MText mtext = (MText)ent;
        }
    }
}
