using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _10_文字样式
{
    public class Class1
    {
        [CommandMethod("TestStyleDemo")]
        public void TextStyleDemo()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            db.AddTextStyle("测试文字样式");
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                TextStyleTable tst = (TextStyleTable)trans.GetObject(db.TextStyleTableId, OpenMode.ForRead);
                foreach (var item in tst)
                {

                }

            }
        }

    }
}
