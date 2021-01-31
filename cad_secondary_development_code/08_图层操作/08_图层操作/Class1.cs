using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _10图层操作
{
    public class Class1
    {
        [CommandMethod("LayerDemo")]
        public void LayerDemo()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            AddLayerResult alr = db.AddLayer("文字");
            db.ChangeLayerColor("文字", 1);
            //db.LockLayer("文字");
            //db.UnLockLayer("文字");
            //db.ChangleLineWeight("文字", LineWeight.LineWeight050);
            //db.SetCurrentLayer("文字");
            //List<LayerTableRecord> list = db.GetAllLayers();
            //db.DeleteLayer("图层1",true);
            db.DeleteNotUsedLayer();
        }
    }
}
