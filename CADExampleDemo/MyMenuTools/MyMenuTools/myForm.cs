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
using Autodesk.AutoCAD.EditorInput;//(Editor, PromptXOptions, PromptXResult)
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Runtime.InteropServices;  //应用[DllImport("user32.dll")]需要

namespace MyMenuTools
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

                // 创建个人类OtherToosClass对象 用来调用里面的方法
                OtherToolsClass otl = new OtherToolsClass();
                SelectionSet acSSet = otl.SelectSsGet("GetSelection", null, typeValue);

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

                // 创建个人类OtherToosClass对象 用来调用里面的方法
                OtherToolsClass otl = new OtherToolsClass();
                SelectionSet acSSet = otl.SelectSsGet("GetSelection", null, acTypValAr);

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
    }
}
