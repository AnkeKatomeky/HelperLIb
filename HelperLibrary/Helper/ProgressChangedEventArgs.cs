using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HelperLibrary
{
    public class ProgressChangedEventArgs : EventArgs
    {
        public ProgressChangedEventArgs(double percentage, string message)
        {
            ProgressPercentage = percentage;
            Message = message;
        }

        public double ProgressPercentage { get; set; }

        public string Message { get; set; }
    }
}
