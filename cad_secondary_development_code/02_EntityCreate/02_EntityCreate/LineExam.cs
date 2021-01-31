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
    
    public class LineExam
    {
        [CommandMethod("LineTest")]
        public void LineTest()
        {
            // 获取当前使用图形所在图形数据库
            Database db = HostApplicationServices.WorkingDatabase;
            // 声明一个直线对象
            // Line line1 = new Line(new Point3d(100, 100, 0), new Point3d(200, 100, 0));
            // 调用封装的函数 将图形添加进去
            // db.AddEntityToModeSpace(line1);

            // 一次绘制多个图形
            //Line line2 = new Line(new Point3d(200, 100, 0), new Point3d(200, 200, 0));
            //Line line3 = new Line(new Point3d(200, 200, 0), new Point3d(100, 100, 0));
            //db.AddEntityToModeSpace(line1, line2, line3);

            // 测试直线封装
           
            db.AddLineToModeSpace(new Point3d(100, 100, 0), new Point3d(200, 100, 0)); // 已知起点和终点
            db.AddLineToModeSpace(new Point3d(200, 200, 0), 200, 60); // 已知起点 角度 长度
            
        }
        

    }
}
