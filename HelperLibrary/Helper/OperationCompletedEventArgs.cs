using HelperLibrary.Helper;
using ReportGenerator.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HelperLibrary
{
    public class OperationCompletedEventArgs : EventArgs
    {
        public CheckResult CheckResult { get; set; }
        public string ErrorMessage { get; set; }
        public string WarningMessage { get; set; }
        public List<InputReport> InputCheck { get; set; }


        public OperationCompletedEventArgs(string errorMessage, string warningMessage)
        {
            ErrorMessage = errorMessage;
            WarningMessage = warningMessage;
        }
        public OperationCompletedEventArgs(CheckResult result)
        {
            CheckResult = result;
        }
        public OperationCompletedEventArgs(string errorMessage, List<InputReport> inputCheck)
        {
            InputCheck = inputCheck;
            ErrorMessage = errorMessage;
        }
        public OperationCompletedEventArgs(string errorMessage)
        {
            ErrorMessage = errorMessage;
        }

        public bool Success
        {
            get
            {
                return ErrorMessage == null;
            }
        }
    }
}
