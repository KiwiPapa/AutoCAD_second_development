using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FlowerTools
{
    public static class PromptTools
    {
        /// <summary>
        /// 获取点
        /// </summary>
        /// <param name="ed">命令行对象</param>
        /// <param name="promptStr">提示信息</param>
        /// <returns> PromptPointResult</returns>
        public static PromptPointResult GetPoint2(this Editor ed, string promptStr)
        {
            // 声明一个获取点的提示类
            PromptPointOptions ppo = new PromptPointOptions(promptStr);
            // 使回车和空格键有效
            ppo.AllowNone = true;
            return ed.GetPoint(ppo);
        }


        /// <summary>
        /// 获取点或关键字
        /// </summary>
        /// <param name="ed">命令行</param>
        /// <param name="promptStr">提示信息</param>
        /// <param name="pointBase">基准点</param>
        /// <param name="keyWord">关键字</param>
        /// <returns>PromptPointResult</returns>
        public static PromptPointResult GetPoint(this Editor ed, string promptStr, Point3d pointBase, params string[] keyWord)
        {

            // 声明一个获取点的提示类
            PromptPointOptions ppo = new PromptPointOptions(promptStr);
            // 使回车和空格键有效
            ppo.AllowNone = true;
            // 添加字符是的相应的字符按键有效
            for (int i = 0; i < keyWord.Length; i++)
            {
                ppo.Keywords.Add(keyWord[i]);
            }
            // 取消系统自动关键字显示
            ppo.AppendKeywordsToMessage = false;
            // 设置基准点
            ppo.BasePoint = pointBase;
            ppo.UseBasePoint = true;
            return ed.GetPoint(ppo);
        }
    }
}
