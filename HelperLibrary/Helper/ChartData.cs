using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HelperLibrary
{
    public class ChartData
    {
        public string Title { get; set; }
        public string CoreColumn { get; set; }

        public int Height { get; set; }
        public int Width { get; set; }

        public int Columm { get; set; }
        public int Row { get; set; }
        public Dictionary<string, string> Data { get; set; }
        public List<int> CoresNumbers { get; set; }

        public int From { get; set; }
        public int To { get; set; }

        public string GlobalTitle { get; set; }
    }
}
