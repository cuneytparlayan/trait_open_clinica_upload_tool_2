using System;
using System.Collections;
using System.Globalization;
using System.IO;
using System.Xml;

namespace OCDataImporter
{
    /// <summary>
    /// Class responsible for providing the study (meta) data for the conversion to ODM and for handling any 
    /// validation problems
    /// </summary>
    class StudyMetaDataValidator
    {

        /*
         * TODO ask Cuneyt: 
         * - removed the DEBUG section which writes some extra info to the file \\OCDataImporter_verification.txt. Is is OK?
         * - removed the try/catch of an exception and the logging of the message in the BuildVerificationArrays method -> add a catch when calling this method 
         */

        private ArrayList ItemGroupDefs = new ArrayList();
        private ArrayList ItemDefForm = new ArrayList();
        private ArrayList CodeList = new ArrayList();
        private ArrayList MSList = new ArrayList();
        private ArrayList RCList = new ArrayList();
        private ArrayList SCOList = new ArrayList();
        private ArrayList AllValuesInOneRow = new ArrayList();
        private ArrayList Hiddens = new ArrayList();


        public String repeatingGroups {get; set;}
        public String repeatingEvents { get; set;}
        
        
        private WarningLog warningLog;
        private ConversionSettings conversionSettings;
        

        public StudyMetaDataValidator(WarningLog warningLog, ConversionSettings conversionSettings) 
        {        
            this.warningLog = warningLog;
            this.conversionSettings = conversionSettings;
        }

        public void reset()
        {
            AllValuesInOneRow.Clear();
            Hiddens.Clear();
            repeatingGroups = "";
            repeatingEvents = "";
        }

        public void BuildVerificationArrays(String pathToMetaDataFile, String pathToVerificationFile, Boolean debugMode)  // 3.0 Data verification structures 
        {
            //*** ItemGroupDefs: CRF, Group contains which items. Item_OID,OPT = optional, Item_OID,MAN = mandatory. Separated by ~
            //F_COCO_V10^IG_COCO_UNGROUPED^I_COCO_LABELLINK,MAN~I_COCO_DATEINTAKE,OPT~I_COCO_DATECOMMENT,OPT~I_COCO_STUDYEXPLANATION,OPT~I_COCO_INFORMEDCONSENTPATIENT,OPT~I_COCO_DATESIGNATUREPATIENT,OPT~I_COCO_INFORMEDCONSENTINVESTIGATOR,OPT~I_COCO_DATESIGNATUREINVESTIGATOR,OPT~I_COCO_STOOLCONSENT,OPT~I_COCO_FITCONSENT,OPT~I_COCO_BLOODCONSENT,OPT~I_COCO_COLONOSCOPYDATE,OPT~I_COCO_STOOLGIVEN,OPT~I_COCO_STOOLCOMMENT,OPT~I_COCO_STOOLDELIVERY,OPT~I_COCO_PICKUPDATE,OPT~I_COCO_PICKUPCOMMENT,OPT~I_COCO_PICKUPCONFIRMATION,OPT~I_COCO_NOPICKUPCOMMENT,OPT~I_COCO_FITDONE,OPT~I_COCO_FITCOMMENT,OPT~I_COCO_FITEIKENBARCODE,OPT~I_COCO_FITONCOBARCODE,OPT
            //F_COCOS_BLOOD__V10^IG_COCOS_UNGROUPED^I_COCOS_BLOOD_THINNER,OPT~I_COCOS_BLOOD_THINNER_SPECIFIED,OPT~I_COCOS_BLOOD_THINNER_SPECIFIED_OTH,OPT~I_COCOS_ASCAL_DOSE,OPT~I_COCOS_PLAVIX_DOSE,OPT~I_COCOS_ACENOFENPRO_DOSE,OPT~I_COCOS_BLOOD_THINNER_OTHER_DOSE,OPT
            //
            // *** ItemDefForm: Item-data attributes in which form and show/hide situation
            //I_MMRES_RES_BEHAN_LIJN^F_MMRESPONSE_10^SHOW^CL_6283^integer^2^999^RES_Behan_lijn
            //I_MMRES_RES_KUUR_TYPE^F_MMRESPONSE_10^SHOW^CL_6284^integer^2^999^RES_Kuur_type
            //
            //*** CodeList: Codelist ID and values separated by ~
            //CL_75^1~0
            //CL_79^NA~1~0
            //CL_10495^1~2~3~4~5~6~7~8~9~-1
            //CL_10497^14~1~2~3~4~5~6~7~8~13~11~12~10~41~42~67~-1
            //
            //*** RCList: Range Check
            //I_TEST_ITEM4^text^GT^10
            //I_TEST_ITEM5^text^GE^1
            //I_TEST_ITEM5^text^LE^5
            //
            //*** MSList
            //MSL_18^1~2~3
            //MSL_25^1
            //
            //*** SCOList
            // ItemOID^ItemName^ControlItemName^ValueToControl
            //I_COCOS_BLOOD_THINNER_SPECIFIED^Blood_thinner_specified^Blood_thinner^1
            //I_COCOS_PLAVIX_DOSE^Plavix_dose^Blood_thinner_specified^2
            //*** Sites
            //S_PHAROSMM
            //S_MM01
            //S_MM02
            //*** Forms
            //MM-diagnose1 - 1.0    F_MMDIAGNOSE1_10
            //MM-Comorbiditeiten - 1.0    F_MMCOMORBIDIT_10
            //
            //*** Groups
            //IG_MMCOM_UNGROUPED    IG_MMCOM_UNGROUPED
            //Comorb    IG_MMCOM_COMORB
            ItemGroupDefs.Clear();
            ItemDefForm.Clear();
            CodeList.Clear();
            RCList.Clear();
            MSList.Clear();
            SCOList.Clear();
            conversionSettings.Sites.Clear();
            conversionSettings.Forms.Clear();
            conversionSettings.Groups.Clear();
            conversionSettings.DOY = false;
            int linenr = 0;

            try
              {
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.ConformanceLevel = ConformanceLevel.Fragment;
                settings.IgnoreWhitespace = true;
                settings.IgnoreComments = true;
                XmlReader reader = XmlReader.Create(pathToMetaDataFile, settings);
                bool res = false;
                res = reader.Read();
                while (res)
                {
                    linenr++;
                    if (reader.Name == "Study")
                    {
                        // build site array 4.3.1
                        if (reader.AttributeCount > 0) conversionSettings.Sites.Add(reader.GetAttribute(0));
                    }
                    if (reader.Name == "FormDef")
                    {
                        // build form array 4.3.1
                        if (reader.AttributeCount > 1) conversionSettings.Forms.Add(reader.GetAttribute(1) + "    " + reader.GetAttribute(0));
                    }
                    if (reader.Name == "ItemGroupDef")
                    {
                        // build group array 4.3.1
                        if (reader.AttributeCount > 1) conversionSettings.Groups.Add(reader.GetAttribute(1) + "    " + reader.GetAttribute(0));
                    }

                    // TODO add methods which perform the parsing of the node to make the code more legible. E.g. 
                    // if (reader.Name == "ItemGroupDef") { parseItemGroupDef(reader) }

                    if (reader.Name == "ItemGroupDef")
                    {
                        if (reader.AttributeCount > 0)
                        {
                            string GrLine = reader.GetAttribute(0) + "^";
                            string GrFLine = "";
                            string myGRPS = reader.ReadInnerXml();
                            foreach (string ss in myGRPS.Split('>'))
                            {
                                ss.Trim();
                                if (ss.Contains("<ItemRef "))
                                {
                                    string myIO = "";
                                    myIO = ss.Substring(18);
                                    int stoppunt = myIO.IndexOf('"');
                                    myIO = myIO.Substring(0, stoppunt);
                                    if (ss.Contains("Mandatory=\"Yes\"")) GrLine += myIO + ",MAN" + "~";
                                    else GrLine += myIO + ",OPT" + "~";
                                }
                                if (ss.Contains("PresentInForm FormOID="))
                                {
                                    GrFLine = ss + "~";
                                    GrFLine = GrFLine.Substring(0, GrFLine.Length - 1);
                                    GrFLine = GrFLine.Replace("<OpenClinica:PresentInForm FormOID=", "");
                                    GrFLine = GrFLine.Replace("ShowGroup=\"Yes\"", "");
                                    GrFLine = GrFLine.Replace("ShowGroup=\"No\"", "");
                                    ItemGroupDefs.Add(GrFLine.Replace("\"", "").Trim() + "^" + GrLine.Substring(0, GrLine.Length - 1));
                                }
                            }
                        }
                    }
                    else if (reader.Name == "ItemDef")
                    {
                        if (reader.AttributeCount > 0)
                        {
                            string ItLine = "";
                            string theType = reader.GetAttribute(2);

                            if (theType.ToLower() == "date") ItLine = "date^999^999";
                            else
                            {
                                if (Utilities.IsNumber(reader.GetAttribute(4)) && Utilities.IsNumber(reader.GetAttribute(3))) ItLine = theType + "^" + reader.GetAttribute(3) + "^" + reader.GetAttribute(4); // should be float
                                else // integer or text
                                {
                                    if (Utilities.IsNumber(reader.GetAttribute(3))) ItLine = theType + "^" + reader.GetAttribute(3) + "^999";
                                    else ItLine = theType + "^999^999";  // we should never come here
                                }
                            }

                            ItLine += "^" + reader.GetAttribute(1); // paste name after item info; will be needed for matching (version=3.03)
                            string myItem = reader.GetAttribute(0);
                            string ItFLine1 = reader.GetAttribute(0) + "^";
                            string ItCLine = "";
                            string ItMSLine = "";
                            string SCOLine = myItem + "^" + reader.GetAttribute(1) + "^";
                            string ItRLine = myItem + "^" + theType + "^";
                            bool RCdone = false;
                            string myGRPS = reader.ReadInnerXml();
                            foreach (string ss in myGRPS.Split('>'))
                            {
                                ss.Trim();
                                if (ss.Contains("ItemPresentInForm FormOID="))
                                {
                                    int start = ss.IndexOf("ItemPresentInForm FormOID=") + "ItemPresentInForm FormOID=".Length + 1;
                                    string myF = ss.Substring(start);
                                    int stoppunt = myF.IndexOf('"');
                                    myF = myF.Substring(0, stoppunt);
                                    string ItFLine = ItFLine1 + myF + "^";
                                    if (ss.Contains("ShowItem=\"No\"")) ItFLine += "HIDE^";
                                    else ItFLine += "SHOW^";
                                    if (ItCLine != "") ItFLine = ItFLine + ItCLine;
                                    else if (ItMSLine != "") ItFLine = ItFLine + ItMSLine;
                                    else ItFLine = ItFLine + "NOCODE";
                                    ItemDefForm.Add(ItFLine + "^" + ItLine);
                                }
                                if (ss.Contains("CodeListRef CodeListOID="))
                                {
                                    int start = ss.IndexOf("CodeListRef CodeListOID=") + "CodeListRef CodeListOID=".Length + 1;
                                    string myF = ss.Substring(start);
                                    int stoppunt = myF.IndexOf('"');
                                    ItCLine = myF.Substring(0, stoppunt);
                                }
                                if (ss.Contains("</OpenClinica:ControlItemName")) SCOLine += ss.Substring(0, ss.IndexOf('<')) + "^";
                                if (ss.Contains("</OpenClinica:OptionValue"))
                                {
                                    SCOLine += ss.Substring(0, ss.IndexOf('<'));
                                    SCOList.Add(SCOLine);
                                }
                                if (ss.Contains("OpenClinica:MultiSelectListRef MultiSelectListID="))
                                {
                                    int start = ss.IndexOf("OpenClinica:MultiSelectListRef MultiSelectListID=") + "OpenClinica:MultiSelectListRef MultiSelectListID=".Length + 1;
                                    string myF = ss.Substring(start);
                                    int stoppunt = myF.IndexOf('"');
                                    ItMSLine = "*" + myF.Substring(0, stoppunt);
                                }
                                if (ss.Contains("<RangeCheck Comparator="))
                                {
                                    if (!RCdone) // do this only once, as there can be more than 1 range checks
                                    {
                                        string rangestring = myGRPS.Replace("<RangeCheck Comparator=", "^");
                                        string[] ranges = rangestring.Split('^');
                                        foreach (string one in ranges)
                                        {
                                            if (one.Length < 3) continue;
                                            string myRCC = one.Substring(1, 2);
                                            int start = one.IndexOf("<CheckValue>") + "<CheckValue>".Length;
                                            int stop = one.IndexOf("</CheckValue>");
                                            if (stop < 0) continue;  // 3.01 -> There can be other sub xml items than CheckValue!
                                            string myRCVal = one.Substring(start, stop - start);
                                            ItRLine += myRCC + "^" + myRCVal;
                                            RCList.Add(ItRLine);
                                            ItRLine = myItem + "^" + theType + "^";
                                            RCdone = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else if (reader.Name == "CodeList")
                    {
                        if (reader.AttributeCount > 0)
                        {
                            string CLine = reader.GetAttribute(0) + "^";
                            string myGRPS = reader.ReadInnerXml();
                            foreach (string ss in myGRPS.Split('>'))
                            {
                                ss.Trim();
                                if (ss.Contains("CodeListItem CodedValue="))
                                {
                                    int start = ss.IndexOf("CodeListItem CodedValue=") + "CodeListItem CodedValue=".Length + 1;
                                    string myC = ss.Substring(start);
                                    int stoppunt = myC.IndexOf('"');
                                    myC = myC.Substring(0, stoppunt);
                                    CLine += myC + "~";
                                }
                            }
                            CLine = CLine.Substring(0, CLine.Length - 1);
                            CodeList.Add(CLine);
                        }
                    }
                    else if (reader.Name == "OpenClinica:StudyParameterListRef")
                    {
                        if (reader.AttributeCount > 1)
                        {
                            if (reader.GetAttribute(0) == "SPL_collectDob" && reader.GetAttribute(1) == "2") conversionSettings.DOY = true;
                            string myGRPS = reader.ReadInnerXml();
                        }
                    }
                    else if (reader.Name == "OpenClinica:MultiSelectList")
                    {
                        if (reader.AttributeCount > 0)
                        {
                            string CLine = reader.GetAttribute(0) + "^";
                            string myGRPS = reader.ReadInnerXml();
                            foreach (string ss in myGRPS.Split('>'))
                            {
                                ss.Trim();
                                if (ss.Contains("OpenClinica:MultiSelectListItem CodedOptionValue="))
                                {
                                    int start = ss.IndexOf("OpenClinica:MultiSelectListItem CodedOptionValue=") + "OpenClinica:MultiSelectListItem CodedOptionValue=".Length + 1;
                                    string myC = ss.Substring(start);
                                    int stoppunt = myC.IndexOf('"');
                                    myC = myC.Substring(0, stoppunt);
                                    CLine += myC + "~";
                                }
                            }
                            CLine = CLine.Substring(0, CLine.Length - 1);
                            MSList.Add(CLine);
                        }
                    }                    
                    else
                    {
                        res = reader.Read();
                    }
                }
    
            }
                
            catch (Exception ex)
            {
//                MessageBox.Show("Can't get XML definitions for verification data from the metafile; unexpected exception:", "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
  //              exit_error("Metadatafile Line=" + linenr.ToString() + " -> " + ex.Message);
                throw new OCDataImporterException("Can't get XML definitions for verification data from the metafile; unexpected exception at line: " + linenr.ToString(), ex);
            }

            if (debugMode)
            {
                using (StreamWriter swlog = new StreamWriter(pathToVerificationFile))
                {
                    swlog.WriteLine("*** ItemGroupDefs");
                    foreach (string ss in ItemGroupDefs) swlog.WriteLine(ss);
                    swlog.WriteLine("");
                    swlog.WriteLine("*** ItemDefForm");
                    foreach (string ss in ItemDefForm) swlog.WriteLine(ss);
                    swlog.WriteLine("");
                    swlog.WriteLine("*** CodeList");
                    foreach (string ss in CodeList) swlog.WriteLine(ss);
                    swlog.WriteLine("*** RCList");
                    foreach (string ss in RCList) swlog.WriteLine(ss);
                    swlog.WriteLine("");
                    swlog.WriteLine("*** MSList");
                    foreach (string ss in MSList) swlog.WriteLine(ss);
                    swlog.WriteLine("");
                    swlog.WriteLine("*** SCOList");
                    foreach (string ss in SCOList) swlog.WriteLine(ss);
                    swlog.WriteLine("*** Sites");
                    foreach (string ss in conversionSettings.Sites) swlog.WriteLine(ss);
                    swlog.WriteLine("*** Forms");
                    foreach (string ss in conversionSettings.Forms) swlog.WriteLine(ss);
                    swlog.WriteLine("*** Groups");
                    foreach (string ss in conversionSettings.Groups) swlog.WriteLine(ss);
                    if (conversionSettings.DOY) swlog.WriteLine("*** Only year of birth accepted for subjects = true");
                }
            }             
        }


        public string ValidateItem(string key, string FormOID, string ItemOID, string ItemVal, int linenr, String dateFormat)
        {
            // check integer or float, has only digits
            // check length = itemval.Length
            // check PDATE format 
            // check if SHOW and MANDATORY and the value is missing
            // check if value within range of code list
            int man = 0; // 0=unknown, 1=mandatory, 2=optional
            int show = 0; // 0=unknown, 1=show, 2, hide
            string thecodelist = "";
            string ittype = "";
            int itlen = -1;
            int itdeclen = -1;
            ItemVal = ItemVal.Trim();
            AllValuesInOneRow.Add(ItemOID + "^" + ItemVal);
            string fixedwarning = "Line " + linenr.ToString() + ", Subject= " + key + ", CRF= " + FormOID + ", ItemOID= " + ItemOID + ", Item Value= " + ItemVal + ": ";
            foreach (string igds in ItemGroupDefs)
            {
                string[] igd = igds.Split('^');
                if (igd[0] == FormOID)
                {
                    string[] fields = igd[2].Split('~');
                    foreach (string one in fields)
                    {
                        if (one == ItemOID + ",OPT") man = 2;
                        if (one == ItemOID + ",MAN") man = 1;
                    }
                }
                if (man > 0) break;
            }
            foreach (string idfs in ItemDefForm)
            {
                string[] idf = idfs.Split('^');
                if (idf[0] == ItemOID && idf[1] == FormOID)
                {
                    if (idf[2] == "SHOW") show = 1;
                    if (idf[2] == "HIDE") show = 2;
                    thecodelist = idf[3];
                    ittype = idf[4];
                    itlen = System.Convert.ToInt16(idf[5]);
                    itdeclen = System.Convert.ToInt16(idf[6]);
                    break;
                }
            }
            if (man == 0 || itlen == -1) warningLog.appendMessage(fixedwarning + "*** Wrong XML definitions: Item not found in XML. Please inform the ICT. Thank you in advance.");
            if (show == 1 && man == 1 && ItemVal == "") warningLog.appendMessage(fixedwarning + "Item is mandatory but has no value");
            if (show == 2 && man == 1 && ItemVal == "") Hiddens.Add(ItemOID);
            string ConvertedDate = "";
            if (ittype == "date")
            {
                if (dateFormat == "--select--") warningLog.appendMessage(fixedwarning + "Item is date but no date format is selected in parameters");
                ConvertedDate = Utilities.ConvertToODMFormat(ItemVal, dateFormat);

            }
            if (ittype == "partialDate") ConvertedDate = Utilities.ConvertToODMPartial(ItemVal);
            if (ConvertedDate.StartsWith("Error")) warningLog.appendMessage(fixedwarning + ConvertedDate);
            else if (ConvertedDate != "") return (ConvertedDate);

            if (thecodelist != "NOCODE" && ItemVal != "")
            {
                if (thecodelist.StartsWith("*") == false) // single selection code list
                {
                    bool found = false;
                    string theVals = "";
                    foreach (string codes in CodeList)
                    {
                        string[] cd = codes.Split('^');
                        if (cd[0] == thecodelist)
                        {
                            string[] cdl = cd[1].Split('~');
                            theVals = cd[1].Replace('~', ',');
                            foreach (string one in cdl)
                            {
                                if (one.Trim() == ItemVal.Trim()) found = true;
                            }
                            if (found == true) break;
                        }
                    }
                    if (found == false) warningLog.appendMessage(fixedwarning + "Value not in code list: " + theVals);
                }
                else  // MSList (multiple selection)
                {
                    thecodelist = thecodelist.Substring(1);
                    string theVals = "";
                    int found = 0;
                    foreach (string codes in MSList)
                    {
                        string[] cd = codes.Split('^');
                        if (cd[0] == thecodelist)
                        {
                            string[] cdl = cd[1].Split('~');
                            theVals = cd[1].Replace('~', ',');
                            string[] selvals = ItemVal.Trim().Split(',');
                            foreach (string one in cdl)
                            {
                                foreach (string one1 in selvals)
                                {
                                    if (one.Trim() == one1.Trim()) found++;
                                }
                            }
                            if (found < selvals.Length) warningLog.appendMessage(fixedwarning + "(at least one of) value(s) not in multiple selection list: " + theVals);
                        }
                    }
                }
            }
            else
            {
                if (ItemVal.Length > itlen) warningLog.appendMessage(fixedwarning + "Item value exceeds required width = " + itlen.ToString());
                if (ittype == "integer" || ittype == "float")
                {
                    string theval = ItemVal.Replace(".", "").Replace(",", "").Replace("-", "");
                    if (ittype == "float")
                    {
                        if (Utilities.IsNumber(theval) == false) warningLog.appendMessage(fixedwarning + "Item type is real but contains non numeric characters: " + ItemVal);
                        string thefloatval = ItemVal.Replace(",", ".");
                        if (thefloatval.IndexOf('.') > 0)
                        {
                            int stdec = thefloatval.LastIndexOf('.');
                            string thedec = thefloatval.Substring(stdec + 1);
                            if (thedec.Length > itdeclen) warningLog.appendMessage(fixedwarning + "Item contains more numbers than allowed after the decimal point: " + ItemVal + " (allowed = " + itdeclen.ToString() + " numbers)");
                        }
                    }
                    else
                    {
                        if (Utilities.IsNumber(ItemVal.Replace("-", "")) == false) warningLog.appendMessage(fixedwarning + "Item type is integer but contains non integer characters: " + ItemVal);
                    }
                }
            }
            string huidige = "";
            try
            {
                foreach (string one in RCList)
                {
                    huidige = one;
                    string[] RCS = one.Split('^');
                    String RCVal = RCS[3].Trim();
                    float fItemVal = 0;
                    float fRCVal = 0;
                    if ((RCS[0] == ItemOID) && ItemVal != "")
                    {
                        if (RCS[1] == "float")
                        {
                            fItemVal = Single.Parse(ItemVal, CultureInfo.InvariantCulture);
                            fRCVal = Single.Parse(RCVal, CultureInfo.InvariantCulture);
                        }
                        if ((RCS[2] == "GT" || RCS[2] == "LT" || RCS[2] == "NE") && (ItemVal == RCVal)) warningLog.appendMessage(fixedwarning + "Range Check fail: Value " + ItemVal + " does not satisfy " + RCS[2] + " " + RCVal);
                        if ((RCS[2] == "EQ") && (ItemVal != RCVal)) warningLog.appendMessage(fixedwarning + "Range Check fail: Value " + ItemVal + " should be equal to " + RCVal);
                        if (RCS[2] == "GT" || RCS[2] == "GE")
                        {
                            if (RCS[1] == "text") if (string.Compare(ItemVal, RCVal) < 0) warningLog.appendMessage(fixedwarning + "Range Check fail: Value " + ItemVal + " is not " + RCS[2] + " " + RCVal);
                            if (RCS[1] == "integer") if (System.Convert.ToInt32(ItemVal) < System.Convert.ToInt32(RCVal)) warningLog.appendMessage(fixedwarning + "Range Check fail: Value " + ItemVal + " is not " + RCS[2] + " " + RCVal);
                            if (RCS[1] == "float") if (fItemVal < fRCVal) warningLog.appendMessage(fixedwarning + "Range Check fail: Value " + ItemVal + " is not " + RCS[2] + " " + RCVal);
                        }
                        if (RCS[2] == "LT" || RCS[2] == "LE")
                        {
                            if (RCS[1] == "text") {
                                if (string.Compare(RCVal, ItemVal) < 0)
                                {
                                    warningLog.appendMessage(fixedwarning + "Range Check fail: Value " + ItemVal + " is not " + RCS[2] + " " + RCVal);
                                }
                            }

                            if (RCS[1] == "integer") {
                                if (System.Convert.ToInt32(RCVal) < System.Convert.ToInt32(ItemVal))
                                {
                                    warningLog.appendMessage(fixedwarning + "Range Check fail: Value " + ItemVal + " is not " + RCS[2] + " " + RCVal);
                                }
                            }
                            if (RCS[1] == "float")
                            {
                                if (fRCVal < fItemVal) {
                                    warningLog.appendMessage(fixedwarning + "Range Check fail: Value " + ItemVal + " is not " + RCS[2] + " " + RCVal);
                                    }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                warningLog.appendMessage(fixedwarning + "Range check ignored: " + huidige.Replace('^', ',') + ". " + ex.Message);
            }
            return (ItemVal);
        }

        public string GetGroupFromItemCRF(string ItemOID, string FormOID)
        {
            foreach (string idfs in ItemGroupDefs)
            {
                string[] idf = idfs.Split('^');
                string[] labelinfo = idf[2].Split('~');
                foreach (string labels in labelinfo)
                {
                    if (labels.StartsWith(ItemOID + ",") && idf[0] == FormOID) return (idf[1]);
                }
            }
            return ("NOTFOUND1");
        }

        public string GetItemNameFromItemOID(string ItemOID, string FormOID)
        {
            foreach (string idfs in ItemDefForm)
            {
                string[] idf = idfs.Split('^');
                if (idf[0] == ItemOID && idf[1] == FormOID) return (idf[7]);
            }
            return (ItemOID);
        }

        public string GetItemOIDFromItemName(string ItemName, string FormOID)
        {
            foreach (string idfs in ItemDefForm)
            {
                string[] idf = idfs.Split('^');
                if (idf[7].ToLower() == ItemName && idf[1] == FormOID) return (idf[0]);
            }
            foreach (string idfs in ItemDefForm)
            {
                string[] idf = idfs.Split('^');
                if (idf[7].ToLower() == ItemName) return ("ANOTHERCRF");
            }
            return ("NOTFOUND2");
        }

        public void validateHiddens(String formOID, int linecount, String studySubjectID)
        {
            string contItem = "";
            string contVal = "";
            foreach (string hd in Hiddens)
            {
                foreach (string scos in SCOList)
                {
                    string[] sco = scos.Split('^');
                    if (sco[0] == hd)
                    {
                        contItem = sco[2];
                        contVal = sco[3];
                        string ItemToControl = GetItemOIDFromItemName(contItem.ToLower(), formOID);
                        foreach (string alls in AllValuesInOneRow)
                        {
                            string[] all = alls.Split('^');
                            if (all[0] == ItemToControl && all[1] == contVal) warningLog.appendMessage("Line " + linecount.ToString() + ", Subject= " + studySubjectID + ", CRF= " + formOID + ", ItemOID= " + hd + ", Item Value is null while item is mandatory");
                        }
                    }
                }
            }
        }
    }
}
