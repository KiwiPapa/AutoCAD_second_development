using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _13_表格操作
{
    public struct BlockData
    {
        public string blockName;
        public string layerName;
        public string X;
        public string Y;
        public string Z;
        public string ZS;
        public string XS;
    }
    public class Class1
    {
        [CommandMethod("TableDemo")]
        public void TableDemo()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            Table table = new Table();
            table.SetSize(10, 5);
            table.SetRowHeight(10); // 设置行高
            table.SetColumnWidth(50); // 设置列宽
            table.Columns[0].Width = 20; // 设置第一列宽度为20
            table.Position = new Point3d(100, 100, 0); // 设置插入点
            table.Cells[0, 0].TextString = "测试表格数据统计";
            table.Cells[0, 0].TextHeight = 6; //设置文字高度
            Color color = Color.FromColorIndex(ColorMethod.ByAci, 3); // 声明颜色
            table.Cells[0, 0].BackgroundColor = color; // 设置背景颜色
            color = Color.FromColorIndex(ColorMethod.ByAci, 1);
            table.Cells[0, 0].ContentColor = color; //内容颜色
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                btr.AppendEntity(table);
                trans.AddNewlyCreatedDBObject(table, true);
                trans.Commit();
            }
        }


        [CommandMethod("DataToTableDemo")]
        public void DataToTableDemo()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor; // 交互命令行
            TypedValue[] values = new TypedValue[]
            {
                new TypedValue((int)DxfCode.Start,"TEXT"),     // 你要选择过滤的图形对象类型
                new TypedValue((int)DxfCode.LayerName,"JZP")   // 过滤图形所在图层
            };
            SelectionFilter filter = new SelectionFilter(values);//选择过滤集
            PromptSelectionResult psr = ed.GetSelection(filter);
            if (psr.Status == PromptStatus.OK)
            {
                ObjectId[] ids = psr.Value.GetObjectIds();
                PromptPointResult ppr = ed.GetPoint("选择表格的插入点:");
                if(ppr.Status == PromptStatus.OK)
                {
                    Point3d point = ppr.Value;
                    BlockData[] data = this.GetBlockRefData(db, ids);
                    // 插入表格
                    this.SetDataToTable(db, data, point);
                }              
            }
        }



        /// <summary>
        /// 将数据以表格形式插入图形
        /// </summary>
        /// <param name="db"></param>
        /// <param name="data"></param>
        /// <param name="position"></param>
        private void SetDataToTable(Database db, BlockData[] data, Point3d position)
        {
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                TableExt table = new TableExt(data.Length, 7, position,data,"数据统计");  // 新建表格
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                btr.AppendEntity(table);
                trans.AddNewlyCreatedDBObject(table, true);
                trans.Commit();
            }
        }



        /// <summary>
        /// 获取块参照的信息
        /// </summary>
        /// <param name="db">图形数据库</param>
        /// <param name="ids"></param>
        /// <returns></returns>
        private BlockData[] GetBlockRefData(Database db, ObjectId[] ids)
        {
            BlockData[] data = new BlockData[ids.Length]; // 声明一个数组用于存放选取结果
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                for (int i = 0; i < ids.Length; i++)
                {
                    // 块名 图层 坐标 ZS XS
                    BlockReference br = (BlockReference)ids[i].GetObject(OpenMode.ForRead);
                    data[i].blockName = br.Name;
                    data[i].layerName = br.Layer;
                    data[i].X = br.Position.X.ToString();
                    data[i].Y = br.Position.Y.ToString();
                    data[i].Z = br.Position.Z.ToString();
                    // 获取 ZS XS
                    foreach (ObjectId item in br.AttributeCollection)
                    {
                        AttributeReference attRef = (AttributeReference)item.GetObject(OpenMode.ForRead);
                        if (attRef.Tag.ToString() == "ZS")
                        {
                            data[i].ZS = attRef.TextString;
                        }
                        else if(attRef.Tag.ToString() == "XS")
                        {
                            data[i].XS = attRef.TextString;
                        }
                    }
                }               
            }
            return data;
        }
    }
}
