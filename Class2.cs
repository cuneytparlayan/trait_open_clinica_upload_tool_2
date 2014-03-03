using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OCDataImporter
{
    public static class DateUtilities
    {
        public static bool CheckDay(string day, string month)
        {
            try
            {
                int theDay = System.Convert.ToInt16(day);
                if (month == "01" || month == "03" || month == "05" || month == "07" || month == "08" || month == "10" || month == "12")
                {
                    if (theDay < 1 || theDay > 31) return (false);
                }
                if (month == "04" || month == "06" || month == "09" || month == "11")
                {
                    if (theDay < 1 || theDay > 30) return (false);
                }
                if (month == "02")
                {
                    if (theDay < 1 || theDay > 29) return (false);
                }
            }
            catch (Exception ex)
            {
                return (false);
            }
            return (true);
        }
    }
}
