using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace HiProtobuf.Lib
{
    internal class ExcelInfo
    {
        public string Name { get; }
        public int RowCount { get; }//行

        public int ColCount { get; }//列

        public Range Range { get; }

        public ExcelInfo(string name, int row, int col, Range range)
        {
            this.Name = name;
            this.RowCount = row;
            this.ColCount = col;
            this.Range = range;
        }
    }
}
