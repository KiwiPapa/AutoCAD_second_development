// (C) Copyright 2020 by  
//
using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using System.Collections.Generic;

// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(_11批量修改标注文本内容.MyCommands))]

namespace _11批量修改标注文本内容
{
    public class MyCommands
    {
        [CommandMethod("CHT")]
        public void ChangeDimText()
        {
            List<KeyValuePair<string, string>> liststring = new List<KeyValuePair<string, string>>();
            // 添加需要替换和替换后的文本  这类可以把你需要替换内容对都添加进来 
            liststring.Add(new KeyValuePair<string, string>("需要替换的标注内容1", "替换后标注内容1"));

            liststring.Add(new KeyValuePair<string, string>("需要替换的标注内容2", "替换后标注内容2"));

            liststring.Add(new KeyValuePair<string, string>("需要替换的标注内容3", "替换后标注内容3"));

            // 打开图形数据库
            Database db = HostApplicationServices.WorkingDatabase;
            // 打开事务处理
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // 以只读方式打开块表
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                // 以写的方式打开块表记录
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // 在块表记录中遍历对象
                foreach (ObjectId id in btr)
                {
                    DBObject ent = trans.GetObject(id, OpenMode.ForWrite);
                    // 判断是否是标注
                    if (ent is Dimension)
                    {
                        Dimension dim = ent as Dimension;

                        int newstring = ReplaceString(dim.DimensionText.Trim(), liststring);
                        // 如果正确索引到该文本字符串
                        if (newstring >= 0)
                        {
                            // 执行替换
                            dim.DimensionText = dim.DimensionText.Replace(liststring[newstring].Key.Trim(), liststring[newstring].Value.Trim());
                        }
                    }
                }
                trans.Commit();
            }
        }


        // IndexOf() 查找字串中指定字符或字串首次出现的位置,返首索引值
        // Trim()    用于删除字符串头部及尾部出现的空格，删除的过程为从外到内，直到碰到一个非空格的字符为止

        /// <summary>
        /// 字符串查找
        /// </summary>
        /// <param name="inputstring">输入字符串</param>
        /// <param name="getalls"></param>
        /// <returns></returns>
        private int ReplaceString(string inputstring, List<KeyValuePair<string, string>> allgetstring)
        {
            int returnvalues = -1;
            for (int i = 0; i < allgetstring.Count; i++)
            {
                if (inputstring.Trim().IndexOf(allgetstring[i].Key.Trim()) >= 0)
                {
                    returnvalues = i;
                    break;
                }
            }
            return returnvalues;
        }
    }

}
