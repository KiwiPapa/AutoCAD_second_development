using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _13_表格操作
{
    public class TableExt:Table
    {
        public TableExt(int rows, int columns, Point3d position, BlockData[] data, string header)
        {
            this.SetSize(rows + 1, columns);
            this.Position = position;
            this.Cells[0, 0].TextString = header;
            this.SetColumnWidth(100);
            for (int i = 0; i < rows + 1; i++)
            {
                this.Cells[i, 0].TextString = data[i-1].blockName;
                this.Cells[i, 1].TextString = data[i-1].layerName;
                this.Cells[i, 2].TextString = data[i-1].X;
                this.Cells[i, 3].TextString = data[i-1].Y;
                this.Cells[i, 4].TextString = data[i-1].Z;
                this.Cells[i, 5].TextString = data[i-1].ZS;
                this.Cells[i, 6].TextString = data[i-1].XS;

            }
        }
    }
}
