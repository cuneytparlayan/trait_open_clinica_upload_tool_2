using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace OCDataImporter
{
    public static class Utilities
    {

        /// <summary>
        /// Checks if a repeating event is correctly formated according to the format:
        /// ITEMNAME_E{N}_G{N}. Throws a OCDataImporterException if the format is not correct
        /// </summary>
        /// <param name="theName">The item name</param>
        /// <returns>the value of the repeat</returns>
        /// 

        public static string GetOID(string val)
        {
            val = val.Replace("    ", "^");
            string[] vals = val.Split('^');
            if (vals.Length == 2) return vals[1];
            else return val;
        }

        public static string Get_SE_RepeatingKey_FromStudyDataColumn(string theName)
        {
            // Format: Adverse_Event_E1_G3      We must get the "1" here. It may not exist, as in:  SurgeryPlan_E_G2 or, for any other reason. -> Then error!   
            string errtext = "Error while getting STUDYEVENT Repeating Key: Cant resolve the DataItemColumnName " + theName + ". The proper name should look like 'DataItem_E2_G3 Where E2 means Event repeating key = 2 and G3 means Group repeating key = 3. \r\nASSUMING 1 AS REPEATING KEY...If this is not what you meant, the generated files ARE INCOMPLETE AND CAN NOT BE USED. See documentation for more.";
            string[] split = theName.Split('_');
            int lengtOfName = split.Length;
            int waarde = -1;
            try
            {
                if (lengtOfName >= 2 && split[lengtOfName - 1].Length > 1 && split[lengtOfName - 1].StartsWith("E"))
                {
                    waarde = System.Convert.ToInt16(split[lengtOfName - 1].Substring(1));
                    if (waarde > 0) return (waarde.ToString());
                }
                else if (lengtOfName >= 3 && split[lengtOfName - 2].Length > 1 && split[lengtOfName - 2].StartsWith("E"))
                {
                    waarde = System.Convert.ToInt16(split[lengtOfName - 2].Substring(1));
                    if (waarde > 0) return (waarde.ToString());
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(errtext + "    " + ex.Message, "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                //exit_error(errtext + "    " + ex.Message);
                throw new OCDataImporterException(errtext + "   " + ex.Message);
            }
            // MessageBox.Show(errtext, "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            // exit_error(errtext);  Version 2.0.1 -> Assume 1 if the repeating key cant be determined. (and dont exit) The error message tells the user about the caveats.
            return ("1");
        }

        /// <summary>
        /// Checks if a repeating group is correctly formated according to the format:
        /// ITEMNAME_E{N}_G{N}. Throws a OCDataImporterException if the format is not correct
        /// </summary>
        /// <param name="theName">The item name</param>
        /// <returns>the value of the repeat</returns>
        public static string Get_GR_RepeatingKey_FromStudyDataColumn(string theName)
        {
            // Format: Adverse_Event_E1_G3      We must get the "3" here; G=group, C= WAS also group, not anymore as of 2.0.2 to avoid confusion
            string errtext = "Error while getting GROUP Repeating Key: Cant resolve the DataItemColumnName " + theName + ". The proper name should look like 'DataItem_E2_G3 Where E2 means StudyEvent repeating key = 2 and G3 means Group repeating key = 3. \r\nExiting...The generated files ARE INCOMPLETE AND CAN NOT BE USED";
            string[] split = theName.Split('_');
            int lengtOfName = split.Length;
            int waarde = -1;
            try
            {
                if (lengtOfName >= 2 && split[lengtOfName - 1].Length > 1 && (split[lengtOfName - 1].StartsWith("G")))
                {
                    waarde = System.Convert.ToInt16(split[lengtOfName - 1].Substring(1));
                    if (waarde > 0) return (waarde.ToString());
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(errtext + "    " + ex.Message, "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                //exit_error(errtext + "    " + ex.Message);
                throw new OCDataImporterException(errtext + "   " + ex.Message);
            }
            //MessageBox.Show(errtext, "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            //exit_error(errtext);
            return ("N");  // this means: Cant determine but maybe going to get from row: the check will be done in DoWork
        }

        public static string FillTildes(string var, int len)
        {
            if (var.Length < len)
            {
                int aantalnullen = len - var.Length;
                string voornullen = "";
                for (int ii = 0; ii < aantalnullen; ii++) voornullen += "~";
                return (var + voornullen);
            }
            return (var);
        }
                     

        public static bool IsNumber(String s)
        {
            bool value = true;
            foreach (Char c in s.ToCharArray())
            {
                value = value && Char.IsDigit(c);
            }

            return value;
        }
        
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

        public static string ConvertToODMPartial(string theSource)
        {
            if (theSource.Trim() == "") return ("");
            theSource = theSource.Replace('/', '-');
            try
            {
                string[] splitd = theSource.Split('-');
                if (splitd[0].Length == 4 && (splitd[1].Length == 3 || splitd[1].Length == 2))   // yyyy-mmm or yyyy-mm
                {
                    string mon = Utilities.Get_maand(splitd[1]);
                    if (mon.StartsWith("Error")) return (mon);
                    else
                    {
                        if (Utilities.CheckYear(splitd[0])) return (splitd[0] + "-" + mon);
                        else return ("Error: Wrong Year in partial date");
                    }
                }
                if (splitd[1].Length == 4 && (splitd[0].Length == 3 || splitd[0].Length == 2))   // mmm-yyyy or mm-yyyy
                {
                    string mon = Utilities.Get_maand(splitd[0]);
                    if (mon.StartsWith("Error")) return (mon);
                    else
                    {
                        if (Utilities.CheckYear(splitd[1])) return (splitd[1] + "-" + mon);
                        else return ("Error: Wrong Year in partial date");
                    }
                }
            }
            catch (Exception ex)
            {
                return ("Error: Unrecognised partial date and exception: " + ex.Message);
            }
            return ("Error: Unrecognised partial date");
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
                    string mon = Utilities.Get_maand(splitd[1]);
                    if (mon.StartsWith("Error")) return (mon);
                    string day = splitd[0];
                    if (day.Length == 1) day = "0" + day;
                    if (Utilities.CheckYear(splitd[2]) && Utilities.CheckDay(day, mon)) return (splitd[2] + "-" + mon + "-" + day);
                    else return ("Error: Wrong day or year");
                }
                catch (Exception ex)
                {
                    return ("Error: date is not in day-month-yyyy format ");
                }
            }
            if (selectedDateFormat == "month-day-year")
            {                
                try
                {
                    string[] splitd = theSource.Split('-');
                    string mon = Utilities.Get_maand(splitd[0]);
                    if (mon.StartsWith("Error")) return (mon);
                    string day = splitd[1];
                    if (day.Length == 1) day = "0" + day;
                    if (Utilities.CheckYear(splitd[2]) && Utilities.CheckDay(day, mon)) return (splitd[2] + "-" + mon + "-" + day);
                    else return ("Error: Wrong day or year");
                }
                catch (Exception ex)
                {
                    return ("Error: date is not in month-day-yyyy format ");
                }
            }
            if (selectedDateFormat == "year-month-day")
            {            
                try
                {
                    string[] splitd = theSource.Split('-');
                    string mon = Utilities.Get_maand(splitd[1]);
                    if (mon.StartsWith("Error")) return (mon);
                    string day = splitd[2];
                    if (day.Length == 1) day = "0" + day;
                    if (Utilities.CheckYear(splitd[0]) && Utilities.CheckDay(day, mon)) return (splitd[0] + "-" + mon + "-" + day);
                    else return ("Error: Wrong day or year");
                }
                catch (Exception ex)
                {
                    return ("Error: date is not in yyyy-month-day format ");
                }
            }
            return ("Error: unknown date format ");
        }        
    }
}
