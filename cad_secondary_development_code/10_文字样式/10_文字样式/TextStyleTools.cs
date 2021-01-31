using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _10_文字样式
{
    public static partial class TextStyleTools
    {
        /// <summary>
        /// 添加文字样式
        /// </summary>
        /// <param name="db">图形数据库</param>
        /// <param name="textStyleName"><文字样式名称/param>
        public static void AddTextStyle(this Database db, string textStyleName)
        {
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // 打开文字样式表
                TextStyleTable tst = (TextStyleTable)trans.GetObject(db.TextStyleTableId, OpenMode.ForRead);
                if(!tst.Has(textStyleName))
                {
                    // 声明文字样式表记录
                    TextStyleTableRecord tstr = new TextStyleTableRecord();
                    tstr.Name = textStyleName;
                    // 把新的文字样式表记录加入文字样式表
                    tst.UpgradeOpen();
                    tst.Add(tstr);
                    //更新
                    trans.AddNewlyCreatedDBObject(tstr, true);
                    tst.DowngradeOpen();
                    // 提交事务处理
                    trans.Commit();
                }
            }
        }
    }
}
