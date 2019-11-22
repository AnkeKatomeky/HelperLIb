using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGenerator.Model
{
    public enum InputState
    {
        //MainDivision
        CS,
        ATP,
        UBD,
        SMOG,
        SC,
        FCT,
        BaseData,
        UploadDiv,
        SL,
        ItemMaster,
        //Forcast
        MainForecast,
        Demand,
        Sellin,
        ForecastAchievement,
        //Import
        Inbound,
        Invoice,
        Upload,
        VATMatrix
    }

    public class InputReport
    {
        public string UpdateDelay { get; set; }
        public string Name { get; set; }
        public string Status { get; set; }
        public bool IsHeader { get; set; }
        public List<InputState> State { get; set; }


        public string MSG
        {
            get
            {
                return Status.Split('|')[0];
            }
        }

        public string ClearStatus
        {
            get
            {
                if (Status.Contains("|"))
                {
                    return Status.Split('|')[1];
                }
                else
                {
                    return Status;
                }
            }
        }
    }
}
