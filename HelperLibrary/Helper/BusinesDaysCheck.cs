using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HelperLibrary.Helper
{
    public static class BusinesDaysCheck
    {
        /// <summary>
        /// Calculates number of business days, taking into account:
        ///  - weekends (Saturdays and Sundays)
        ///  - bank holidays in the middle of the week
        /// </summary>
        /// <param name="firstDay">First day in the time interval</param>
        /// <param name="lastDay">Last day in the time interval</param>
        /// <param name="bankHolidays">List of bank holidays excluding weekends</param>
        /// <returns>Number of business days during the 'span'</returns>
        public static int BusinessDaysUntil(DateTime firstDay, DateTime lastDay)
        {
            firstDay = firstDay.Date;
            lastDay = lastDay.Date;
            if (firstDay > lastDay)
                throw new ArgumentException("Incorrect last day " + lastDay);

            TimeSpan span = lastDay - firstDay;
            int businessDays = span.Days + 1;
            int fullWeekCount = businessDays / 7;
            // find out if there are weekends during the time exceedng the full weeks
            if (businessDays > fullWeekCount * 7)
            {
                // we are here to find out if there is a 1-day or 2-days weekend
                // in the time interval remaining after subtracting the complete weeks
                int firstDayOfWeek = firstDay.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)firstDay.DayOfWeek;
                int lastDayOfWeek = lastDay.DayOfWeek == DayOfWeek.Sunday ? 7 : (int)lastDay.DayOfWeek;
                if (lastDayOfWeek < firstDayOfWeek)
                    lastDayOfWeek += 7;
                if (firstDayOfWeek <= 6)
                {
                    if (lastDayOfWeek >= 7)// Both Saturday and Sunday are in the remaining time interval
                        businessDays -= 2;
                    else if (lastDayOfWeek >= 6)// Only Saturday is in the remaining time interval
                        businessDays -= 1;
                }
                else if (firstDayOfWeek <= 7 && lastDayOfWeek >= 7)// Only Sunday is in the remaining time interval
                    businessDays -= 1;
            }

            // subtract the weekends during the full weeks in the interval
            businessDays -= fullWeekCount + fullWeekCount;

            return businessDays;
        }

        public static DateTime AddBusinessDays(DateTime date, int days)
        {
            DateTime dateTime = date;

            for (int i = 0; i < days; i++)
            {
                if (dateTime.DayOfWeek == DayOfWeek.Saturday)
                {
                    dateTime = dateTime.AddDays(2);
                }
                if (dateTime.DayOfWeek == DayOfWeek.Sunday)
                {
                    dateTime = dateTime.AddDays(1);
                }
                dateTime = dateTime.AddDays(1);
            }
            if (dateTime.DayOfWeek == DayOfWeek.Saturday)
            {
                dateTime = dateTime.AddDays(2);
            }
            if (dateTime.DayOfWeek == DayOfWeek.Sunday)
            {
                dateTime = dateTime.AddDays(1);
            }
            return dateTime;
        }

        public static DateTime AddCorrectionDays(DateTime date, int days)
        {
            DateTime dateTime = date.AddDays(days);

            if (days > 0)
            {
                if (dateTime.DayOfWeek == DayOfWeek.Saturday)
                {
                    dateTime = dateTime.AddDays(2);
                }
                if (dateTime.DayOfWeek == DayOfWeek.Sunday)
                {
                    dateTime = dateTime.AddDays(1);
                }
            }
            else
            {
                if (dateTime.DayOfWeek == DayOfWeek.Saturday)
                {
                    dateTime = dateTime.AddDays(-1);
                }
                if (dateTime.DayOfWeek == DayOfWeek.Sunday)
                {
                    dateTime = dateTime.AddDays(-2);
                }
            }
            
            return dateTime;
        }
    }
}
