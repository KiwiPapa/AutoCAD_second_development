using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


using Autodesk.AutoCAD.DatabaseServices;// (Database, DBPoint, Line, Spline) 
using Autodesk.AutoCAD.Geometry;//(Point3d, Line3d, Curve3d) 
using Autodesk.AutoCAD.ApplicationServices;// (Application, Document) 
using Autodesk.AutoCAD.Runtime;// (CommandMethodAttribute, RXObject, CommandFlag) 
using Autodesk.AutoCAD.EditorInput;//(Editor, PromptXOptions, PromptXResult)
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;


using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;   
using NPOI.XSSF.UserModel;
using System.IO;
using System.Runtime.InteropServices;  //应用[DllImport("user32.dll")]需要


namespace _09_输出文本
{
    public partial class myForm : Form
    {
        //初始化  窗口焦点切换功能
        [DllImport("user32.dll", EntryPoint = "SetFocus")]
        public static extern int SetFocus(IntPtr hWnd);
        public myForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<string> lstAddStr = new List<string>();  //读取的结果存在这里

            Database db = HostApplicationServices.WorkingDatabase;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            SetFocus(doc.Window.Handle);       //选择完文件在切换焦点
                                               //锁定文档
            using (DocumentLock acLckDoc = doc.LockDocument())
            {
                //框选获取文字 
                //设置选择集过滤器为只选择多行文本
                TypedValue[] typeValue = new TypedValue[1];
                typeValue.SetValue(new TypedValue(0, "MTEXT"), 0);
                SelectionSet acSSet = this.SelectSsGet("GetSelection", null, typeValue);

                if (acSSet != null)
                {
                    foreach (SelectedObject selObj in acSSet)
                    {
                        // 确认返回的是合法的SelectedObject对象  
                        if (selObj != null) //
                        {
                            //开始启动事物调整文字位置点和对齐点
                            using (Transaction trans = db.TransactionManager.StartTransaction())
                            {
                                MText myent = trans.GetObject(selObj.ObjectId, OpenMode.ForWrite) as MText;
                                lstAddStr.Add(myent.Contents);  //文字对象复制
                            }
                        }
                    }
                }


                //数据接下来输出到EXCEL表格
                IWorkbook wk = null;  //新建IWorkbook对象

                string localFilePath = "D:\\TEST.xlsx";

                // 调用一个系统自带的保存文件对话框 写一个EXCEL
                SaveFileDialog saveFileDialog = new SaveFileDialog();   //新建winform自带保存文件对话框对象
                saveFileDialog.Filter = "Excel Office97-2003(*.xls)|*.xls|Excel Office2007及以上(*.xlsx)|*.xlsx";  //过滤只能存储的对象
                DialogResult result = saveFileDialog.ShowDialog();     //显示对话框
                localFilePath = saveFileDialog.FileName.ToString();
                //07版之前和之后创建方式不一样
                if (localFilePath.IndexOf(".xlsx") > 0) // 2007版
                {
                    wk = new XSSFWorkbook();   //创建表格对象07版之后
                }
                else if (localFilePath.IndexOf(".xls") > 0) // 创建表格对象 2003版本
                {
                    wk = new HSSFWorkbook();  //03版
                }
                //创建工作簿
                ISheet tb = wk.CreateSheet("输出的文本");
                for (int i = 0; i < lstAddStr.Count; i++)
                {
                    ICell cell = tb.CreateRow(i).CreateCell(0);  //单元格对象 第i行第0列  cell 单元格对象
                    cell.SetCellValue(lstAddStr[i]);//循环往单元格赋值
                }
                //创建文件
                using (FileStream fs = File.OpenWrite(localFilePath)) //打开一个xls文件，如果没有则自行创建，如果存在myxls.xls文件则在创建是不要打开该文件！
                {
                    wk.Write(fs);   //文件IO 创建EXCEL
                    MessageBox.Show("提示：创建成功！");
                    fs.Close();
                }
            }
        }

      
          

        private void button2_Click(object sender, EventArgs e)
        {
            List<string> lstAddStr = new List<string>();  //读取的结果存在这里

            Database db = HostApplicationServices.WorkingDatabase;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            SetFocus(doc.Window.Handle);       //选择完文件在切换焦点
                                               //锁定文档
            using (DocumentLock acLckDoc = doc.LockDocument())
            {
                //框选获取文字 选择集知识点
                //设置选择集过滤器为只选择单行文字
                TypedValue[] acTypValAr = new TypedValue[1];
                acTypValAr.SetValue(new TypedValue(0, "MTEXT"), 0); // 单行文本DBText
                SelectionSet acSSet = this.SelectSsGet("GetSelection", null, acTypValAr);

                if (acSSet != null)
                {
                    foreach (SelectedObject selObj in acSSet)
                    {
                        // 确认返回的是合法的SelectedObject对象  
                        if (selObj != null) //
                        {
                            //开始启动事物调整文字位置点和对齐点
                            using (Transaction trans = db.TransactionManager.StartTransaction())
                            {
                                MText myent = trans.GetObject(selObj.ObjectId, OpenMode.ForWrite) as MText;
                                lstAddStr.Add(myent.Contents);  //文字对象复制
                            }
                        }
                    }
                }
            }
     

            SaveFileDialog saveDlg = new SaveFileDialog();
            saveDlg.Title = "输出文本内容";
            saveDlg.Filter = "文本文件(*.txt)|*.txt";
          
            saveDlg.InitialDirectory = Path.GetDirectoryName(db.Filename);
            string fileName = Path.GetFileName(db.Filename);
            saveDlg.FileName = fileName.Substring(0, fileName.IndexOf('.'));
            DialogResult saveDlgRes = saveDlg.ShowDialog();
            if (saveDlgRes == DialogResult.OK)
            {
                string[] contents = new string[lstAddStr.Count];
                for (int i = 0; i < lstAddStr.Count; i++)
                {
                    contents[i] = lstAddStr[i].ToString();
                }
                File.WriteAllLines(saveDlg.FileName, contents);
            }
        }


        /// <summary>
        /// 获取选择集
        /// </summary>
        /// <param name="selectStr">选择方式</param>
        /// <param name="point3dCollection">选择点集合</param>
        /// <param name="typedValue">过滤参数</param>
        /// <returns></returns>
        public SelectionSet SelectSsGet(string selectStr, Point3dCollection point3dCollection, TypedValue[] typedValue)
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            // 将过滤条件赋值给SelectionFilter对象
            SelectionFilter selfilter = null;
            if (typedValue != null)
            {
                selfilter = new SelectionFilter(typedValue);
            }
            // 请求在图形区域选择对象
            PromptSelectionResult psr;
            if (selectStr == "GetSelection")  // 提示用户从图形文件中选取对象
            {
                psr = ed.GetSelection(selfilter);
            }
            else if (selectStr == "SelectAll") //选择当前空间内所有未锁定及未冻结的对象
            {
                psr = ed.SelectAll(selfilter);
            }
            else if (selectStr == "SelectCrossingPolygon") //选择由给定点定义的多边形内的所有对象以及与多边形相交的对象。多边形可以是任意形状，但不能与自己交叉或接触。
            {
                psr = ed.SelectCrossingPolygon(point3dCollection, selfilter);
            }
            // 选择与选择围栏相交的所有对象。围栏选择与多边形选择类似，所不同的是围栏不是封闭的， 围栏同样不能与自己相交
            else if (selectStr == "SelectFence")
            {
                psr = ed.SelectFence(point3dCollection, selfilter);
            }
            // 选择完全框入由点定义的多边形内的对象。多边形可以是任意形状，但不能与自己交叉或接触
            else if (selectStr == "SelectWindowPolygon")
            {
                psr = ed.SelectWindowPolygon(point3dCollection, selfilter);
            }
            else if (selectStr == "SelectCrossingWindow")  //选择由两个点定义的窗口内的对象以及与窗口相交的对象
            {
                Point3d point1 = point3dCollection[0];
                Point3d point2 = point3dCollection[1];
                psr = ed.SelectCrossingWindow(point1, point2, selfilter);
            }
            else if (selectStr == "SelectWindow") // 选择完全框入由两个点定义的矩形内的所有对象。
            {
                Point3d point1 = point3dCollection[0];
                Point3d point2 = point3dCollection[1];
                psr = ed.SelectCrossingWindow(point1, point2, selfilter);
            }
            else
            {
                return null;
            }

            // 如果提示状态OK，表示对象已选
            if (psr.Status == PromptStatus.OK)
            {
                SelectionSet sSet = psr.Value;
                ed.WriteMessage("Number of objects selected: " + sSet.Count.ToString() + "\n");// 打印选择对象数量
                return sSet;
            }
            else
            {
                // 打印选择对象数量
                ed.WriteMessage("Number of objects selected 0 \n");
                return null;
            }
        }

    }
}
