using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace HelperLibrary
{
    public static class SupportApplication
    {
        public static string ExecutableName
        {
            get
            {
                return System.Reflection.Assembly.GetExecutingAssembly().Location;
            }
        }

        public static string StartupPath
        {
            get
            {
                return System.IO.Path.GetDirectoryName(ExecutableName);
            }
        }
    }
}
