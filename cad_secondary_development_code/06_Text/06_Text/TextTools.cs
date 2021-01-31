using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _06_Text
{
    /// <summary>
    /// 特殊字符
    /// </summary>
     public static partial class TextTools
    {
       public struct TextSpecialSymbol
        {
            // 符号16进制编码
            public static readonly string Degree = @"\U+00B0"; // 角度(°)
            public static readonly string Tolerance = @"\U+00B1"; // 公差(±）
            public static readonly string Diameter = @"\U+00D8";  // 直径(∅)
            public static readonly string Angle = @"\U+2220"; // 角度（∠）
            public static readonly string AlmostEqual = @"\U+2248"; // 约等于(≈）|
            public static readonly string LineBoundary = @"\U+E100"; // 边界线
            public static readonly string LineCenter = @"\U+2104"; // 中心线
            public static readonly string Delta = @"\U+0349"; // 增量(△)
            public static readonly string ElecttricalPhase = @"\U+0278"; //电相位（φ)
            public static readonly string LineFlow = @"\U+E101";  // 流线
            public static readonly string Identity = @"\U+2261";  // 标识
            public static readonly string InitialLength = @"\U+E200";// 初始长度
            public static readonly string LineMonument = @"\U+E102"; // 界碑线
            public static readonly string Notequal = @"\U+2260";  // 不等于（≠）
            public static readonly string Ohm = @"\U+2126";  // 欧姆
            public static readonly string Omega = @"\U+03A9"; // 欧米伽(Ω)
            public static readonly string LinePlate = @"\U+0214A"; // 地界线
            public static readonly string Subscript2 = @"\U+2082"; // 下标2
            public static readonly string Square = @"\U+00B2"; // 平方
            public static readonly string Cube = @"\U+00B3"; // 立方
            public static readonly string Overline = @"%%o"; // 单行文字上划线
            public static readonly string Underline = @"%%u";  // 单行文字下划线
            public static readonly string Alpha = @"\U+03B1";  // 希腊字母（α）
            public static readonly string Belta = @"\U+03B2";  // 希腊字母（β）
            public static readonly string Gamma = @"\U+03B3";  // 希腊字母（γ）
            public static readonly string Theta = @"\U+03B8";  // 希腊字母（θ）
            public static readonly string SeteelBar1 = @"\U+0082";  // 一级钢筋符号
            public static readonly string SeteelBar2 = @"\U+0083";  // 二级钢筋符号
            public static readonly string SeteelBar3 = @"\U+0084";  // 三级钢筋符号
            public static readonly string SeteelBa41 = @"\U+0085";  // 四级钢筋符号            
        }

        /// <summary>
        /// 堆叠符号
        /// </summary>
        public struct MTextStackType
        {
            public static readonly string Horizental = "/"; // 水平堆叠
            public static readonly string Italic = "=";     // 斜分堆叠
            public static readonly string Tolerance = "^";  // 容差堆叠
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="text"></param>
        /// <param name="scaleFactor"></param>
        /// <param name="topText"></param>
        /// <param name="stackType"></param>
        /// <param name="bottomText"></param>
        /// <returns></returns>
        public static string StackMtext(string text, double scaleFactor, string topText, string stackType,string bottomText)
        {
            return string.Format("\\A1;{0}{1}\\H{2}x;\\S{3}{4}{5};{6}", text, "{", scaleFactor, topText, stackType, bottomText, "}");
        }
    }


   
}
