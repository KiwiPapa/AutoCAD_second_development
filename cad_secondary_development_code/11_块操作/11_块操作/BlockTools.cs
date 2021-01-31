using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _11_块操作
{
    public static class BlockTools
    {
        /// <summary>
        /// 添加块表记录到图形数据库
        /// </summary>
        /// <param name="db">图形数据库</param>
        /// <param name="btrName">块表记录名</param>
        /// <param name="ents">图形对象</param>
        /// <returns></returns>
        public static ObjectId AddBlockTableRecord(this Database db, string btrName,List<Entity>ents)
        {
            ObjectId btrId = ObjectId.Null;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                if(!bt.Has(btrName))
                {
                    BlockTableRecord btr = new BlockTableRecord();
                    for (int i = 0; i < ents.Count; i++)
                    {
                        btr.AppendEntity(ents[i]);
                        trans.AddNewlyCreatedDBObject(ents[i], true);
                    }
                }
                btrId = bt[btrName];
                trans.Commit();
            }
            return btrId;
        }
    }
}
