using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _11_块操作
{
    /// <summary>
    /// 第一个块
    /// </summary>
    public static class MyBlockTableRecords
    {
        private static string block1Name = "Block";
        public static string Block1Name { get => block1Name; set => block1Name = value; }

        private static List<Entity> block1Ents = new List<Entity>();
        public static List<Entity> Block1Ents
        {
            get => block1Ents;
            set => block1Ents = value;
        }

        private static ObjectId block1Id = ObjectId.Null;       
        public static ObjectId Block1Id { get => block1Id; set => block1Id = value; }
        
    }
}
