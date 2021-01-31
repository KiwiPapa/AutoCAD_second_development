using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AcDoNetTools;

namespace _04_用户交互
{
    public class CircleJig:EntityJig
    {
        private Point3d jCenter; // 圆心
        private double jRadius;  // 半径
        public CircleJig(Point3d center)
            :base(new Circle())  // 继承父类Circle的属性
        {
            ((Circle)Entity).Center = center;  // Enitty 转化为Cirecle 对象 复制center
        }
        // 用于更新图像对象 这里更新属性时无需使用事务处理
        protected override bool Update()
        {
            if(jRadius > 0)
            {
                 ((Circle)Entity).Radius = jRadius;               
            }
           
            return true;
        }
        // 这个函数的作用是当鼠标在屏幕上移动时 就会被调用 实现这个函数时 一般是用它改变图形的属性 我们在这个类定义的属性
        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            // 声明拖拽类jig提示信息
            JigPromptPointOptions jppo = new JigPromptPointOptions("\n 请指定圆上的一个点");
            char space = (char)32;        
            jppo.Keywords.Add("U");
            jppo.Keywords.Add(space.ToString());
            jppo.UserInputControls = UserInputControls.Accept3dCoordinates;
            jppo.Cursor = CursorType.RubberBand;
            jppo.BasePoint = ((Circle)Entity).Center;
            jppo.UseBasePoint = true;
            
            // 获取拖拽时鼠标的位置状态
            PromptPointResult ppr = prompts.AcquirePoint(jppo);

            jRadius = ppr.Value.GetDistanceBetweenTwoPoint(((Circle)Entity).Center);
            return SamplerStatus.NoChange; // 继续移动 循环检测
        }


        public Entity GetEntity()
        {
            return Entity;
        }
    }
}
