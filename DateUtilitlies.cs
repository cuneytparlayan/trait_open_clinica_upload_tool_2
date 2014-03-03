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


        public static bool CheckYear(string year)
        {
            try
            {
                int theYear = System.Convert.ToInt16(year);
                DateTime dt = DateTime.Now;
                if (theYear < 1700 || theYear > (System.Convert.ToInt16(dt.Year) + 1)) return false;
                else return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public static string Get_maand(string inp)
        {
            inp = inp.ToUpper();
            switch (inp)
            {
                case "JAN":
                case "01":
                case "1":
                    return ("01");
                case "FEB":
                case "02":
                case "2":
                    return ("02");
                case "MAR":
                case "MRT":
                case "03":
                case "3":
                    return ("03");
                case "APR":
                case "04":
                case "4":
                    return ("04");
                case "MAY":
                case "05":
                case "MEI":
                case "5":
                    return ("05");
                case "JUN":
                case "06":
                case "6":
                    return ("06");
                case "JUL":
                case "07":
                case "7":
                    return ("07");
                case "AUG":
                case "08":
                case "8":
                    return ("08");
                case "SEP":
                case "09":
                case "9":
                    return ("09");
                case "OCT":
                case "OKT":
                case "10":
                    return ("10");
                case "NOV":
                case "11":
                    return ("11");
                case "DEC":
                case "12":
                    return ("12");
                default:
                    return ("Error: month:" + inp + " not found");
            }
        }

        public static string ConvertToODMFormat(string theSource, String selectedDateFormat)
        {
            if (theSource.Trim() == "") return ("");  // 1.0f .Trim = 2.0.4
            if (selectedDateFormat == "--select--") return (theSource);
            theSource = theSource.Replace('/', '-');
            if (selectedDateFormat == "day-month-year")
            {            
                try
                {
                    string[] splitd = theSource.Split('-');
                    string mon = DateUtilities.Get_maand(splitd[1]);
                    if (mon.StartsWith("Error")) return (mon);
                    string day = splitd[0];
                    if (day.Length == 1) day = "0" + day;
                    if (DateUtilities.CheckYear(splitd[2]) && DateUtilities.CheckDay(day, mon)) return (splitd[2] + "-" + mon + "-" + day);
                    else return ("Error: Wrong day or year");
                }
                catch (Exception ex)
                {
                    return ("Error: date is not in day-month-yyyy format " + ex);
                }
            }
            if (selectedDateFormat == "month-day-year")
            {                
                try
                {
                    string[] splitd = theSource.Split('-');
                    string mon = DateUtilities.Get_maand(splitd[0]);
                    if (mon.StartsWith("Error")) return (mon);
                    string day = splitd[1];
                    if (day.Length == 1) day = "0" + day;
                    if (DateUtilities.CheckYear(splitd[2]) && DateUtilities.CheckDay(day, mon)) return (splitd[2] + "-" + mon + "-" + day);
                    else return ("Error: Wrong day or year");
                }
                catch (Exception ex)
                {
                    return ("Error: date is not in month-day-yyyy format " + ex);
                }
            }
            if (selectedDateFormat == "year-month-day")
            {            
                try
                {
                    string[] splitd = theSource.Split('-');
                    string mon = DateUtilities.Get_maand(splitd[1]);
                    if (mon.StartsWith("Error")) return (mon);
                    string day = splitd[2];
                    if (day.Length == 1) day = "0" + day;
                    if (DateUtilities.CheckYear(splitd[0]) && DateUtilities.CheckDay(day, mon)) return (splitd[0] + "-" + mon + "-" + day);
                    else return ("Error: Wrong day or year");
                }
                catch (Exception ex)
                {
                    return ("Error: date is not in yyyy-month-day format " + ex);
                }
            }
            return ("Error: unknown date format ");
        }        
    }
}
