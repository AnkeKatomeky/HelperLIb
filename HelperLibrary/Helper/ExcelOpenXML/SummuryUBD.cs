using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGenerator.Model
{
    public class SummuryUBD
    {
        public string cellRef { get; set; }

        public string Business { get; set; }
        public string BUDetailed { get; set; }

        public double ExpiredD { get; set; }
        public double ExRemain0_50D { get; set; }
        public double ExRemain51_70D { get; set; }
        public double ExRemain71_100D { get; set; }
        public double NoExpireD { get; set; }
        public double TotalD { get; set; }

        public double ExpiredCount { get; set; }
        public double ExRemain0_30Count { get; set; }
        public double ExRemain31_60Count { get; set; }
        public double ExRemain61_120Count { get; set; }
        public double ExRemain121_360Count { get; set; }
        public double ExRemainMore360Count { get; set; }
        public double NoExpireCount { get; set; }
        public double NoForSale_RA { get; set; }
        public double NoForSale_Damage { get; set; }
        public double NoForSale_Other { get; set; }
        public double TotalCount { get; set; }

        public double DemoLBM { get; set; }
        public double FeildActionLBM { get; set; }
        public double ExpireLBM { get; set; }
        public double OutOfOrderLBM { get; set; }
        public double BlockLBM { get; set; }
        public double OtherLBM { get; set; }
        public double DamageLBM { get; set; }
        public double TotalSubstandartLBM
        {
            get
            {
                return DemoLBM +
                       FeildActionLBM +
                       ExpireLBM +
                       OutOfOrderLBM +
                       BlockLBM +
                       DamageLBM +
                       OtherLBM;
            }
        }

        public double ClinicalStorageLBM { get; set; }
        public double InOfficeLBM { get; set; }
        public double ProbationLBM { get; set; }
        public double TotalRegLBM
        {
            get
            {
                return ClinicalStorageLBM + InOfficeLBM + ProbationLBM;
            }
        }

    }
}
