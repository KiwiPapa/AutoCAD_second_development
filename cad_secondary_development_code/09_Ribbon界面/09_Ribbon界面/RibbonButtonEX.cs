using Autodesk.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media.Imaging;


namespace _09_Ribbon界面
{
    public class RibbonButtonEX : RibbonButton
    {
        //正常显示的图片
        private string imgFileName = "";
        public string ImgFileName
        {
            get { return imgFileName; }
            set { imgFileName = value; }
        }
        //鼠标进入时的图片
        private string imgHoverFileName = "";
        public string ImgHoverFileName
        {
            get { return imgHoverFileName; }
            set { imgHoverFileName = value; }
        }

        public RibbonButtonEX(string name, RibbonItemSize size, Orientation orient, string cmd)
            : base()
        {
            this.Name = name;//按钮的名称
            this.Text = name;
            this.ShowText = true; //显示文字
            this.MouseEntered += this_MouseEntered;
            this.MouseLeft += this_MouseLeft;
            this.Size = size; //按钮尺寸
            this.Orientation = orient; //按钮排列方式
            this.CommandHandler = new RibbonCommandHandler(); //给按钮关联命令
            this.CommandParameter = cmd + " ";
            this.ShowImage = true; //显示图片
        }
        public void SetImg(string imgFileName)
        {
            Uri uri = new Uri(imgFileName);
            BitmapImage bitmapImge = new BitmapImage(uri);
            this.Image = bitmapImge; //按钮图片
            this.LargeImage = bitmapImge; //按钮大图片
            this.imgFileName = imgFileName;
        }
        /// <summary>
        /// 鼠标离开事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void this_MouseLeft(object sender, EventArgs e)
        {
            if (this.ImgFileName != "")
            {
                RibbonButton btn = (RibbonButton)sender;
                string imgFileName = this.ImgFileName;
                Uri uri = new Uri(imgFileName);
                BitmapImage bitmapImge = new BitmapImage(uri);
                btn.Image = bitmapImge; //按钮图片
                btn.LargeImage = bitmapImge; //按钮大图片
            }

        }
        /// <summary>
        /// 鼠标进入事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void this_MouseEntered(object sender, EventArgs e)
        {
            if (this.ImgHoverFileName != "")
            {
                RibbonButton btn = (RibbonButton)sender;
                string imgFileName = this.ImgHoverFileName;
                Uri uri = new Uri(imgFileName);
                BitmapImage bitmapImge = new BitmapImage(uri);
                btn.Image = bitmapImge; //按钮图片
                btn.LargeImage = bitmapImge; //按钮大图片
            }

        }
    }
}