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
    public static class RibbonButtonInfos
    {
        //直线按钮
        private static RibbonButtonEX lineBtn;
        public static RibbonButtonEX LineBtn
        {
            get
            {
                lineBtn = new RibbonButtonEX("直线", RibbonItemSize.Large, Orientation.Vertical, "Line");
                lineBtn.SetImg(CurPaht.curPaht + "Images\\Line.PNG");//设置按钮图片
                //添加提示对象
                RibbonToolTip toolTip = new RibbonToolTip();
                toolTip.Title = "直线";
                toolTip.Content = "创建直线段";
                toolTip.Command = "LINE";
                toolTip.ExpandedContent = "是用LINE命令，可以创建一些列连续的直线段。每条线段都是可以单独进行编辑的直线对象。";
                string imgToolTipFileName = CurPaht.curPaht + "Images\\LineTooTip.PNG";
                Uri toolTipUri = new Uri(imgToolTipFileName);
                BitmapImage toolTipBitmapImge = new BitmapImage(toolTipUri);
                toolTip.ExpandedImage = toolTipBitmapImge;
                lineBtn.ToolTip = toolTip;
                //鼠标进入时的图片
                lineBtn.ImgHoverFileName = CurPaht.curPaht + "Images\\LineHover.PNG";
                return lineBtn;
            }
        }
        //多段线按钮
        private static RibbonButtonEX polylineBtn;
        public static RibbonButtonEX PolylineBtn
        {
            get
            {
                polylineBtn = new RibbonButtonEX("多段线", RibbonItemSize.Large, Orientation.Vertical, "Pline");
                polylineBtn.SetImg(CurPaht.curPaht + "Images\\Polyline.PNG");//设置按钮图片
                //添加提示对象
                RibbonToolTip toolTip = new RibbonToolTip();
                toolTip.Title = "多段线";
                toolTip.Content = "创建二维多段线";
                toolTip.Command = "PLINE";
                toolTip.ExpandedContent = "二维多段线是作为单个平面对象创建的相互连接的线段序列。可以创建直线段、圆弧段或者两者的组合线段。";
                string imgToolTipFileName = CurPaht.curPaht + "Images\\PolylineToolTip.PNG";
                Uri toolTipUri = new Uri(imgToolTipFileName);
                BitmapImage toolTipBitmapImge = new BitmapImage(toolTipUri);
                toolTip.ExpandedImage = toolTipBitmapImge;
                polylineBtn.ToolTip = toolTip;
                //鼠标进入时的图片
                polylineBtn.ImgHoverFileName = CurPaht.curPaht + "Images\\PolylineHover.PNG";
                return polylineBtn;
            }
        }
        //圆心半径按钮
        private static RibbonButtonEX circleCRBtn;
        public static RibbonButtonEX CircleCRBtn
        {
            get
            {
                circleCRBtn = new RibbonButtonEX("圆心,半径", RibbonItemSize.Large, Orientation.Horizontal, "Circle");
                circleCRBtn.SetImg(CurPaht.curPaht + "Images\\CircleCR.PNG");//设置按钮图片
                circleCRBtn.ShowText = false;
                //添加提示对象
                RibbonToolTip toolTip = new RibbonToolTip();
                toolTip.Title = "圆心,半径";
                toolTip.Content = "用圆心和半径创建圆";
                toolTip.Command = "CIRCLE";
                toolTip.ExpandedContent = "用圆心和半径创建圆。\n\n示例:";
                string imgToolTipFileName = CurPaht.curPaht + "Images\\CircleCDHover.PNG";
                Uri toolTipUri = new Uri(imgToolTipFileName);
                BitmapImage toolTipBitmapImge = new BitmapImage(toolTipUri);
                toolTip.ExpandedImage = toolTipBitmapImge;
                circleCRBtn.ToolTip = toolTip;
                //鼠标进入时的图片
                circleCRBtn.ImgHoverFileName = CurPaht.curPaht + "Images\\CircleToolTip.PNG";
                return circleCRBtn;
            }
        }
    }
}