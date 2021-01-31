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



namespace _06_多段线顶点坐标获取
{
    public class Class1
    {
        [CommandMethod("PLVCG")]


        public void PLineVertexCoordsGet()
        {
            //需要访问Database的操作 需首先将该文档进行锁定，操作完成后，在最后进行释放
            DocumentLock docLock = Application.DocumentManager.MdiActiveDocument.LockDocument();
            // 对话框窗口
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            // 数据库对象
            Database db = HostApplicationServices.WorkingDatabase;

            try
            {
                // 开启事务处理
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    // 选择多段线
                    PromptEntityOptions entOpts = new PromptEntityOptions("请选择多段线！");
                    PromptEntityResult entResult = ed.GetEntity(entOpts);
                    // 判断选择是否成功
                    if(entResult.Status == PromptStatus.OK)
                    {
                        //Entity ent = trans.GetObject(entResult.ObjectId, OpenMode.ForRead) as Entity;
                        //Application.ShowAlertDialog(ent.GetType().ToString));

                        // 获取多段线实体对象
                        Polyline pLine = trans.GetObject(entResult.ObjectId, OpenMode.ForRead) as Polyline;
                        //多线段是否闭合 pline.Closed
                        string isclosed = pLine.Closed.ToString();
                        //多线段起始点 pline.StartPoint
                        //多线段结束点 pline.EndPoint

                        //多段线顶点数
                        int vertexNum = pLine.NumberOfVertices;
                        Point3d point;
                        // 遍历获取多段线顶点坐标
                        for (int i = 0; i < vertexNum; i++)
                        {
                            point = pLine.GetPoint3dAt(i);
                            ed.WriteMessage("\n" + point.ToString());
                        }
                    }
                    trans.Commit();
                }
            }

            catch(Autodesk.AutoCAD.Runtime.Exception e)
            {
                Application.ShowAlertDialog(e.Message);
            }
            // 解锁文档
            docLock.Dispose();
        }
    }
}
