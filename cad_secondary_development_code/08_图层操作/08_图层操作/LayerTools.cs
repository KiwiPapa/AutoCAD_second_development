using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _10图层操作
{
    //添加图层的返回状态
    public enum AddLayerStatuts
    {
        AddLayerOK,
        IllegalLayerName,
        LayerNameExist
    }
    //添加图层的返回值
    public struct AddLayerResult
    {
        public AddLayerStatuts statuts;
        public string layerName;
    }
    //修改图层属性的返回状态
    public enum ChangeLayerPropertyStatus
    {
        ChangeOK,
        LayerIsNotExist

    }
    public static partial class LayerTool
    {
        /// <summary>
        /// 添加图层
        /// </summary>
        /// <param name="db">图形数据库</param>
        /// <param name="layerName">图层名</param>
        /// <returns>AddLayerResult</returns>
        public static AddLayerResult AddLayer(this Database db, string layerName)
        {
            //声明AddLayerResult类型的数据，用户返回
            AddLayerResult res = new AddLayerResult() ;
            try
            {
                SymbolUtilityServices.ValidateSymbolName(layerName, false);
            }
            catch (Exception)
            {
                res.statuts = AddLayerStatuts.IllegalLayerName;
                return res;
            }
            
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                //打开层表
                LayerTable lt = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForRead);
                //新建层表记录
                if (!lt.Has(layerName))
                {
                    
                    LayerTableRecord ltr = new LayerTableRecord();
                    //判断要创建的图层名是否已经存在,不存在则创建
                    ltr.Name = layerName;
                    //升级层表打开权限
                    lt.UpgradeOpen();
                    lt.Add(ltr);
                    //降低层表打开权限
                    lt.DowngradeOpen();
                    trans.AddNewlyCreatedDBObject(ltr, true);
                    trans.Commit();
                    res.statuts = AddLayerStatuts.AddLayerOK;
                    res.layerName = layerName;
                }
                else
                {
                    res.statuts = AddLayerStatuts.LayerNameExist;
                } 
            }
            return res;
        }


        /// <summary>
        /// 修改图层颜色
        /// </summary>
        /// <param name="db">图形数据库</param>
        /// <param name="LayerName">图层名</param>
        /// <param name="colorIndex">图层颜色</param>
        /// <returns>ChangeLayerPropertyStatus</returns>
        public static ChangeLayerPropertyStatus ChangeLayerColor(this Database db, string LayerName, short colorIndex)
        {
            ChangeLayerPropertyStatus status;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                //打开层表
                LayerTable lt = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForRead);
                //判断指定的图形名是否存在
                if (lt.Has(LayerName))
                {

                    LayerTableRecord ltr = (LayerTableRecord)lt[LayerName].GetObject(OpenMode.ForWrite);
                    ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, colorIndex);
                    trans.Commit();
                    status = ChangeLayerPropertyStatus.ChangeOK;
                }
                else
                {
                    status = ChangeLayerPropertyStatus.LayerIsNotExist;
                }
            }
            return status;
        }



        /// <summary>
        /// 锁定图层
        /// </summary>
        /// <param name="db">图形数据库</param>
        /// <param name="LayerName">图层名</param>
        /// <returns>bool</returns>
        public static bool LockLayer(this Database db, string LayerName)
        {
            bool isOk = true;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                //打开层表
                LayerTable lt = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForRead);
                //判断指定的图形名是否存在
                if (lt.Has(LayerName))
                {

                    LayerTableRecord ltr = (LayerTableRecord)lt[LayerName].GetObject(OpenMode.ForWrite);
                    ltr.IsLocked = true;
                    trans.Commit();
                }
                else
                {
                    isOk = false;
                }
            }
            return isOk;
        }



        /// <summary>
        /// 解除锁定图层
        /// </summary>
        /// <param name="db">图形数据库</param>
        /// <param name="LayerName">图层名</param>
        /// <returns>bool</returns>
        public static bool UnLockLayer(this Database db, string LayerName)
        {
            bool isOk = true;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                //打开层表
                LayerTable lt = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForRead);
                //判断指定的图形名是否存在
                if (lt.Has(LayerName))
                {

                    LayerTableRecord ltr = (LayerTableRecord)lt[LayerName].GetObject(OpenMode.ForWrite);
                    ltr.IsLocked = false;
                    trans.Commit();
                }
                else
                {
                    isOk = false;
                }
            }
            return isOk;
        }



        /// <summary>
        /// 修改图层的线宽
        /// </summary>
        /// <param name="db">图形数据库</param>
        /// <param name="LayerName">图层名</param>
        /// <param name="lineWeight">线宽</param>
        /// <returns>bool</returns>
        public static bool ChangleLineWeight(this Database db, string LayerName, LineWeight lineWeight)
        {
            bool isOk = true;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                //打开层表
                LayerTable lt = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForRead);
                //判断指定的图形名是否存在
                if (lt.Has(LayerName))
                {

                    LayerTableRecord ltr = (LayerTableRecord)lt[LayerName].GetObject(OpenMode.ForWrite);
                    ltr.LineWeight = lineWeight;
                    trans.Commit();
                }
                else
                {
                    isOk = false;
                }
            }
            return isOk;
        }



        /// <summary>
        /// 设置当前图层
        /// </summary>
        /// <param name="db">图形数据库</param>
        /// <param name="LayerName">图层名</param>
        /// <returns>bool</returns>
        public static bool SetCurrentLayer(this Database db, string LayerName)
        {
            bool isSetOk = false;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                //打开层表
                LayerTable lt = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForRead);
                //判断要设置的图层名是否存在
                if (lt.Has(LayerName))
                {
                    //获取指定图层名的ObjectId
                    ObjectId layerId = lt[LayerName];
                    //判断要设置的图形是否已是当前图层
                    if (db.Clayer != layerId)
                    {
                        db.Clayer = layerId; 
                    }
                    isSetOk = true;
                }
                trans.Commit();
            }
            return isSetOk;
        }


        /// <summary>
        /// 返回所有层表记录
        /// </summary>
        /// <param name="db">图形数据库</param>
        /// <returns>List<LayerTableRecord></returns>
        public static List<LayerTableRecord> GetAllLayers(this Database db)
        {
            List<LayerTableRecord> layerList = new List<LayerTableRecord>();
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                //打开层表
                LayerTable lt = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForRead);
                lt.GenerateUsageData();
                foreach (ObjectId item in lt)
                {
                    LayerTableRecord ltr = (LayerTableRecord)item.GetObject(OpenMode.ForRead);
                    layerList.Add(ltr);
                }
            }
            return layerList;
        }


        /// <summary>
        /// 获取所有图层名
        /// </summary>
        /// <param name="db">图形数据库</param>
        /// <returns>List<string></returns>
        public static List<string> GetAllLayersName(this Database db)
        {
            List<string> layerNamesList = new List<string>();
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                //打开层表
                LayerTable lt = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForRead);
                foreach (ObjectId item in lt)
                {
                    LayerTableRecord ltr = (LayerTableRecord)item.GetObject(OpenMode.ForRead);
                    layerNamesList.Add(ltr.Name);
                }
            }
            return layerNamesList;
        }



        /// <summary>
        /// 删除图层
        /// </summary>
        /// <param name="db">图形数据库</param>
        /// <param name="layerName">图层名</param>
        /// <returns>bool</returns>
        public static bool DeleteLayer(this Database db, string layerName)
        {

            if (layerName == "0" || layerName == "Defpoints") return false;
            bool isDeleteOK = false;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                LayerTable lt = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForRead);
                lt.GenerateUsageData(); 
                //判断要删除的图层名是否存在
                if (lt.Has(layerName))
                {
                    LayerTableRecord ltr = (LayerTableRecord)lt[layerName].GetObject(OpenMode.ForWrite);
                    if (!ltr.IsUsed && db.Clayer != lt[layerName])
                    {
                        ltr.Erase();
                        isDeleteOK = true;
                    }
                }
                else
                {
                    isDeleteOK = true;
                }
                trans.Commit();
            }
            return isDeleteOK;
        }


        /// <summary>
        /// 删除所有未使用的图层
        /// </summary>
        /// <param name="db"></param>
        public static void DeleteNotUsedLayer(this Database db)
        {
            
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                LayerTable lt = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForRead);
                lt.GenerateUsageData();
                foreach (ObjectId item in lt)
                {
                    LayerTableRecord ltr = (LayerTableRecord)item.GetObject(OpenMode.ForWrite);
                    if (!ltr.IsUsed)
                    {
                        ltr.Erase();
                    }
                }
                trans.Commit();
            }
        }

        /// <summary>
        /// 强行删除图层及图层上的所有实体对象
        /// </summary>
        /// <param name="db">图形数据库</param>
        /// <param name="layerName">图层名</param>
        /// <param name="delete"></param>
        /// <returns></returns>
        public static bool DeleteLayer(this Database db, string layerName,bool delete)
        {

            if (layerName == "0" || layerName == "Defpoints") return false;
            bool isDeleteOK = false;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                LayerTable lt = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForRead);
                lt.GenerateUsageData();
                if (lt.Has(layerName))
                {
                    LayerTableRecord ltr = (LayerTableRecord)lt[layerName].GetObject(OpenMode.ForWrite);
                    if (delete)
                    {
                        if (ltr.IsUsed)
                        {
                            ltr.deleteAllEntityInLayer();  
                        }
                        if (db.Clayer == ltr.ObjectId)
                        {
                            db.Clayer = lt["0"];
                        }
                        ltr.Erase();
                        isDeleteOK = true;
                    }
                    else
                    {
                        if (!ltr.IsUsed && db.Clayer != lt[layerName])
                        {
                            ltr.Erase();
                            isDeleteOK = true;
                        }
                    }   
                }
                else
                {
                    isDeleteOK = true;
                }
                trans.Commit();
            }
            return isDeleteOK;
        }
        

        /// <summary>
        /// 删除指定图层上的所有实体对象
        /// </summary>
        /// <param name="ltr"></param>
        public static void deleteAllEntityInLayer(this LayerTableRecord ltr)
        {
            Database db = HostApplicationServices.WorkingDatabase;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            TypedValue[] value = new TypedValue[]
            {
                new TypedValue((int)DxfCode.LayerName,ltr.Name)
            };
            SelectionFilter filter = new SelectionFilter(value);
            PromptSelectionResult psr = ed.SelectAll();
            if (psr.Status == PromptStatus.OK)
            {
                ObjectId[] ids = psr.Value.GetObjectIds();
                using(Transaction trans = db.TransactionManager.StartTransaction())
                {
                    for (int i = 0; i < ids.Length; i++)
                    {
                        Entity ent = (Entity)ids[i].GetObject(OpenMode.ForWrite);
                        ent.Erase();
                    }
                    trans.Commit();
                }
            }
        }
    }
}
