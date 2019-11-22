using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HelperLibrary.ExcelOpenXml
{
    public class IfFormat
    {
        public bool IfCondition { get; set; }
        public CellColor ColorCondition { get; set; }

        public IfFormat(bool cond, CellColor color)
        {
            IfCondition = cond;
            ColorCondition = color;
        }
    }
}
