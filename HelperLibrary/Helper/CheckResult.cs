using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HelperLibrary.Helper
{
    public enum CheckResultEnum
    {
        Error, Warning, OK
    }
    public class CheckResult
    {
        public string StringResult { get; set; }
        public CheckResultEnum EnumResult { get; set; }

        public CheckResult()
        {
            StringResult = "";
        }
    }
}
