using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Security;


namespace OCDataImporter
{   
    class Converter
    {

        private const string INSERT_1A = "INSERT INTO subject(status_id, gender, unique_identifier, date_created, owner_id, dob_collected, date_of_birth)";
        private const string INSERT_1 = "INSERT INTO subject(status_id, gender, unique_identifier, date_created, owner_id, dob_collected)";
        private const string INSERT_2 = "INSERT INTO study_subject(label, study_id, status_id, enrollment_date, date_created, date_updated, owner_id,  oc_oid, subject_id)";
        private const string INSERT_3 = "INSERT INTO study_event(study_event_definition_id, study_subject_id, location, sample_ordinal, date_start, owner_id, status_id, date_created, subject_event_status_id, start_time_flag, end_time_flag)";
        private static string LINE_SEPARATOR = System.Environment.NewLine;
               
        
        private ConversionSettings conversionSettings;
        private StudyMetaDataValidator studyMetaDataValidator;
        private ArrayList InsertSubject = new ArrayList();
        private IViewUpdater viewUpdater;
        private ArrayList Sites = new ArrayList();
        private bool DOY = false;
        private bool labelOCoidExists;
        private ArrayList labelOID;
        

        public int numberOfOutputFiles { get; private set; }

        
        private WarningLog warningLog;    

        public Converter(ConversionSettings conversionSettings, StudyMetaDataValidator studyMetaDataValidator, WarningLog warningLog, IViewUpdater viewUpdater, bool labelOCoidExists, ArrayList labelOID)
        {
            this.conversionSettings = conversionSettings;
            this.studyMetaDataValidator = studyMetaDataValidator;
            this.viewUpdater = viewUpdater;        
            this.warningLog = warningLog;
            this.labelOCoidExists = labelOCoidExists;
            this.labelOID = labelOID;
        }


        public void DoWork(DataGrid dataGrid)
        {
            
            string theDate = DateTime.Now.ToString("yyyy-MM-dd");
            string theWrittenSE = "";
            string theWrittenGR = "";
            string[] theHeaders;
            int event_index_row = -1;
            int group_index_row = -1;
            int site_index_row = -1;
            string usedStudyOID = conversionSettings.studyOID;            
            if (dataGrid.isValid() == false)
            {
                return;
            }


            int DGIndexOfOCItem = DataGrid.DGIndexOfOCItem;
            int DGIndexOfDataItem = DataGrid.DGIndexOfDataItem;

            // initialize files
            int maximalNumberOfLinesOutputFile = conversionSettings.outFMaxLines;
            if (maximalNumberOfLinesOutputFile == 0)
            {
                maximalNumberOfLinesOutputFile = 99999;
            }
            int numberOfLinesOutputFile = 0;
            String outputFileBaseName = conversionSettings.workdir + "\\DataImport_";
            numberOfOutputFiles = 1;
            String outputFilePath = outputFileBaseName + numberOfOutputFiles + ".xml";
            OutputFile odmOutputFile = new OutputFile(outputFilePath);
            masterHeader(odmOutputFile);

            String fileNameINSF = conversionSettings.workdir + "\\Inserts.sql";
            OutputFile insertsSQLFile = new OutputFile(fileNameINSF);

            String fileNameINSFR = conversionSettings.workdir + "\\Inserts_ONLY_STUDY_EVENTS.sql";
            OutputFile insertsOnlyEventsSQLFile = new OutputFile(fileNameINSFR);

            String fileNameDELF = conversionSettings.workdir + "\\Deletes.sql";
            OutputFile deletesSQLFile = new OutputFile(fileNameDELF);

            String fileNameDELFR = conversionSettings.workdir + "\\Deletes_ONLY_STUDY_EVENTS.sql";
            OutputFile deletesOnlyEventsSQLFile = new OutputFile(fileNameDELFR);

            List<String> subjectIDList = new List<String>();
            subjectIDList.Clear();
            try
            {

                String[] linesArray = File.ReadAllLines(conversionSettings.pathToInputFile);
                int linecount = 0;
                foreach (String line in linesArray)
                {

                    // line = line.Trim();  // 1.1b
                    if (line.Length == 0) continue;

                    studyMetaDataValidator.reset();

                    linecount++;
                    if (linecount == 1)
                    {
                        theHeaders = line.Split(dataGrid.delimiter); // get row event and group indexes, if defined that way.
                        for (int i = 0; i < theHeaders.Length; i++)
                        {
                            if (theHeaders[i].ToUpper() == "EVENT_INDEX") event_index_row = i;    // 2.2 Input file format allows EVENT_INDEX and GROUP_INDEX to define repeating items in rows
                            if (theHeaders[i].ToUpper() == "GROUP_INDEX") group_index_row = i;
                            if (theHeaders[i].ToUpper() == "SITE_OID") site_index_row = i;
                        }
                        continue; // skip first line; contains only headers
                    }
                    int mySepCount = 1;
                    viewUpdater.updateProgressbarStep(line.Length);
                    for (int i = 0; i < line.Length; i++) if (line[i] == dataGrid.delimiter) mySepCount++;
                    if (dataGrid.sepcount != mySepCount)
                    {
                        string errtext = "Input data file format incorrect at line = " + linecount.ToString() + " Expecting: " + conversionSettings.sepCount.ToString() + "; found: " + mySepCount.ToString() + "  items; this is the faulty line: " + line;
                        warningLog.appendMessage(errtext);
                        continue;
                    }
                    string[] split = line.Split(dataGrid.delimiter);
                    String subjectID = split[dataGrid.subjectIDIndex];
                    subjectID = subjectID.Trim();
                    if (conversionSettings.checkForDuplicateSubjects)
                    {
                        if (subjectIDList.Contains(subjectID))
                        {
                            string errtext = "Duplicate subjectID " + subjectID + " at line = " + linecount.ToString();
                            warningLog.appendMessage(errtext);
                        }
                    }
                    subjectIDList.Add(subjectID);
                    viewUpdater.performProgressbarStep();

                    //  Handle first DG line
                    int DGFirstLine = 0;
                    String lineToSplit = dataGrid.dataGridView.Rows[DGFirstLine].Cells[DGIndexOfOCItem].Value.ToString();
                    string[] ocparts = dataGrid.dataGridView.Rows[DGFirstLine].Cells[DGIndexOfOCItem].Value.ToString().Split('.');
                    String TheStudyEventOID = ocparts[0];
                    String TheFormOID = "";
                    String TheItemGroupDef = "";
                    String TheItemId = "";
                    string theSERK = "1";
                    string theGRRK = "NOT";
                    string theStudyDataColumn = "";
                    if (TheStudyEventOID == "none")
                    {
                        DGFirstLine++;
                        try // fix 2.0.3 -> If no items are matched; can't increase DGFirstLine, causes index out of range 
                        {
                            ocparts = dataGrid.dataGridView.Rows[DGFirstLine].Cells[DGIndexOfOCItem].Value.ToString().Split('.');
                            TheFormOID = ocparts[1];
                            TheItemGroupDef = ocparts[2];
                            TheItemId = ocparts[3];
                            if (TheItemGroupDef.IndexOf('-') > 0) TheItemGroupDef = TheItemGroupDef.Substring(0, TheItemGroupDef.IndexOf('-'));
                            TheStudyEventOID = ocparts[0];
                            theStudyDataColumn = dataGrid.dataGridView.Rows[DGFirstLine].Cells[DGIndexOfDataItem].Value.ToString();
                        }
                        catch (Exception ex)
                        {
                            TheFormOID = "none";
                            TheItemGroupDef = "none";
                            TheStudyEventOID = "none";
                            theStudyDataColumn = "none";
                            TheItemId = "none";
                        }
                    }
                    else  // 2.1 fix on above fix
                    {
                        TheFormOID = ocparts[1];
                        TheItemGroupDef = ocparts[2];
                        if (TheItemGroupDef.IndexOf('-') > 0) TheItemGroupDef = TheItemGroupDef.Substring(0, TheItemGroupDef.IndexOf('-'));
                        TheItemId = ocparts[3];
                        theStudyDataColumn = dataGrid.dataGridView.Rows[DGFirstLine].Cells[DGIndexOfDataItem].Value.ToString();
                    }

                    if (TheStudyEventOID.Contains("*"))
                    {
                        string[] pp = TheStudyEventOID.Split('*');
                        theWrittenSE = pp[0];
                        theSERK = pp[1];
                    }
                    else
                    {
                        theWrittenSE = TheStudyEventOID;
                        theSERK = "1";
                        // 2.0.5 use the SE selected item to get the SE repeating index  
                        if (dataGrid.startDateItem.Contains("_E"))
                        {
                            theSERK = dataGrid.startDateItem.Substring(dataGrid.startDateItem.IndexOf("_E") + 2);
                        }
                    }
                    if (TheItemGroupDef.Contains("*"))
                    {
                        string[] pp = TheItemGroupDef.Split('*');
                        theWrittenGR = pp[0];
                        theGRRK = pp[1];
                    }
                    else
                    {
                        theGRRK = "NOT";
                        theWrittenGR = TheItemGroupDef.Substring(1);   // get rid of the added A to get this above in grid
                    }
                    // 2.0.5 Do NOT print formdata with no items 
                    string theXMLEvent = "";

                    if (event_index_row != -1) theSERK = split[event_index_row];
                    if (group_index_row != -1) theGRRK = split[group_index_row];

                    if (site_index_row != -1 && split[site_index_row].Trim() != "")
                    {
                        bool found = false;
                        foreach (string one in Sites)
                        {
                            if (one.Trim() == split[site_index_row].Trim())
                            {
                                found = true;
                            }
                        }
                        if (!found)
                        {
                            string errtext = "Specified site not found in meta file: " + split[site_index_row] + " at line = " + linecount.ToString();
                            warningLog.appendMessage(errtext);
                        }
                        usedStudyOID = split[site_index_row];
                    }
                    else usedStudyOID = conversionSettings.studyOID;
                    if (site_index_row != -1 && split[site_index_row].Trim() == "")
                    {
                        string errtext = "No site specified under SITE_OID; assuming STUDY_OID, at line: " + linecount.ToString();
                        warningLog.appendMessage(errtext);
                    }


                    string SStheKEY = "SS_" + subjectID;
                    
                    if (labelOCoidExists) // 2.1.1 there is a conversion file from label to oid; get the SSid from that file.
                    {
                        foreach (string one in labelOID)
                        {
                            if (one.StartsWith(subjectID + "^")) SStheKEY = one.Substring(one.IndexOf('^') + 1);
                        }
                    }
                    
                    string theXMLForm = "";
                    odmOutputFile.Append("    <SubjectData SubjectKey=\"" + SStheKEY + "\">");
                    theXMLEvent += "        <StudyEventData StudyEventOID=\"" + theWrittenSE + "\" StudyEventRepeatKey=\"" + CheckRepeatKey(theSERK, linecount) + "\">" + LINE_SEPARATOR;
                    theXMLForm += "            <FormData FormOID=\"" + TheFormOID + "\">" + LINE_SEPARATOR;
                    if (theGRRK == "NOT") theXMLForm += "                <ItemGroupData ItemGroupOID=\"" + theWrittenGR + "\" TransactionType=\"Insert\" >" + LINE_SEPARATOR;
                    else theXMLForm += "                <ItemGroupData ItemGroupOID=\"" + theWrittenGR + "\" ItemGroupRepeatKey=\"" + CheckRepeatKey(theGRRK, linecount) + "\" TransactionType=\"Insert\" >" + LINE_SEPARATOR;
                    int indexOfItem = 0;
                    string itemval = "";
                    bool datapresentform = false;
                    bool datapresentevent = false;
                    if (TheItemId != "none")
                    {
                        indexOfItem = dataGrid.GetIndexOfItem(dataGrid.dataGridView.Rows[DGFirstLine].Cells[DGIndexOfOCItem].Value.ToString());
                        check_index_of_item(linecount, indexOfItem);
                        itemval = split[indexOfItem];
                        itemval = studyMetaDataValidator.ValidateItem(subjectID, TheFormOID, TheItemId, itemval, linecount, conversionSettings.dateFormat);
                        if (itemval != "")
                        {
                            theXMLForm += "                    <ItemData ItemOID=\"" + TheItemId + "\" Value=\"" + SecurityElement.Escape(itemval) + "\" />" + LINE_SEPARATOR;
                            datapresentform = true;
                        }
                    }
                    // Now handle the rest of the DG
                    for (int i = DGFirstLine + 1; i < dataGrid.dataGridView.RowCount; i++)
                    {
                        if (dataGrid.dataGridView.Rows[i].IsNewRow == false)
                        {
                            string[] nwdingen = dataGrid.dataGridView.Rows[i].Cells[DGIndexOfOCItem].Value.ToString().Split('.');
                            theStudyDataColumn = dataGrid.dataGridView.Rows[i].Cells[DGIndexOfDataItem].Value.ToString();
                            if (TheStudyEventOID == nwdingen[0] && TheFormOID == nwdingen[1] && TheItemGroupDef == nwdingen[2])
                            {
                                if (nwdingen[3] != "none")
                                {
                                    indexOfItem = dataGrid.GetIndexOfItem(dataGrid.dataGridView.Rows[i].Cells[DGIndexOfOCItem].Value.ToString());
                                    check_index_of_item(linecount, indexOfItem);
                                    itemval = split[indexOfItem];
                                    itemval = studyMetaDataValidator.ValidateItem(subjectID, TheFormOID, nwdingen[3], itemval, linecount, conversionSettings.dateFormat);
                                    if (itemval != "")
                                    {
                                        theXMLForm += "                    <ItemData ItemOID=\"" + nwdingen[3] + "\" Value=\"" + SecurityElement.Escape(itemval) + "\" />" + LINE_SEPARATOR;
                                        datapresentform = true;
                                    }
                                }
                            }
                            else if (TheStudyEventOID == nwdingen[0] && TheFormOID == nwdingen[1])
                            {
                                TheItemGroupDef = nwdingen[2];
                                theXMLForm += "                </ItemGroupData>" + LINE_SEPARATOR;
                                if (TheItemGroupDef.Contains("*"))
                                {
                                    string[] pp = TheItemGroupDef.Split('*');
                                    theWrittenGR = pp[0];
                                    theGRRK = pp[1];
                                }
                                else
                                {
                                    theGRRK = "NOT";
                                    theWrittenGR = TheItemGroupDef.Substring(1);
                                }
                                if (group_index_row != -1) theGRRK = split[group_index_row];
                                if (theGRRK == "NOT") theXMLForm += "                <ItemGroupData ItemGroupOID=\"" + theWrittenGR + "\" TransactionType=\"Insert\" >" + LINE_SEPARATOR;
                                else theXMLForm += "                <ItemGroupData ItemGroupOID=\"" + theWrittenGR + "\" ItemGroupRepeatKey=\"" + CheckRepeatKey(theGRRK, linecount) + "\" TransactionType=\"Insert\" >" + LINE_SEPARATOR;
                                if (nwdingen[3] != "none")
                                {
                                    indexOfItem = dataGrid.GetIndexOfItem(dataGrid.dataGridView.Rows[i].Cells[DGIndexOfOCItem].Value.ToString());
                                    check_index_of_item(linecount, indexOfItem);
                                    itemval = split[indexOfItem];
                                    itemval = studyMetaDataValidator.ValidateItem(subjectID, TheFormOID, nwdingen[3], itemval, linecount, conversionSettings.dateFormat);
                                    if (itemval != "")
                                    {
                                        theXMLForm += "                    <ItemData ItemOID=\"" + nwdingen[3] + "\" Value=\"" + SecurityElement.Escape(itemval) + "\" />" + LINE_SEPARATOR;
                                        datapresentform = true;
                                    }
                                }
                            }
                            else if (TheStudyEventOID == nwdingen[0])
                            {
                                TheFormOID = nwdingen[1];
                                TheItemGroupDef = nwdingen[2];
                                theXMLForm += "                </ItemGroupData>" + LINE_SEPARATOR;
                                theXMLForm += "            </FormData>" + LINE_SEPARATOR;
                                if (datapresentform)
                                {
                                    datapresentevent = true;
                                    theXMLEvent += theXMLForm;

                                    datapresentform = false;
                                }
                                theXMLForm = "";
                                theXMLForm += "            <FormData FormOID=\"" + TheFormOID + "\">" + LINE_SEPARATOR;
                                if (TheItemGroupDef.Contains("*"))
                                {
                                    string[] pp = TheItemGroupDef.Split('*');
                                    theWrittenGR = pp[0];
                                    theGRRK = pp[1];
                                }
                                else
                                {
                                    theGRRK = "NOT";
                                    theWrittenGR = TheItemGroupDef.Substring(1);
                                }
                                if (group_index_row != -1) theGRRK = split[group_index_row];
                                // fix 2.1.2 -> if (theGRRK == "NOT") was not there and ItemGroupRepeatKey="NOT" was generated!
                                if (theGRRK == "NOT") theXMLForm += "                <ItemGroupData ItemGroupOID=\"" + theWrittenGR + "\" TransactionType=\"Insert\" >" + LINE_SEPARATOR;
                                else theXMLForm += "                <ItemGroupData ItemGroupOID=\"" + theWrittenGR + "\" ItemGroupRepeatKey=\"" + CheckRepeatKey(theGRRK, linecount) + "\" TransactionType=\"Insert\" >" + LINE_SEPARATOR;
                                if (nwdingen[3] != "none")
                                {
                                    indexOfItem = dataGrid.GetIndexOfItem(dataGrid.dataGridView.Rows[i].Cells[DGIndexOfOCItem].Value.ToString());
                                    check_index_of_item(linecount, indexOfItem);
                                    itemval = split[indexOfItem];
                                    itemval = studyMetaDataValidator.ValidateItem(subjectID, TheFormOID, nwdingen[3], itemval, linecount, conversionSettings.dateFormat);
                                    if (itemval != "")
                                    {
                                        theXMLForm += "                    <ItemData ItemOID=\"" + nwdingen[3] + "\" Value=\"" + SecurityElement.Escape(itemval) + "\" />" + LINE_SEPARATOR;
                                        datapresentform = true;
                                    }
                                }
                            }
                            else
                            {
                                TheStudyEventOID = nwdingen[0];
                                if (TheStudyEventOID.Contains("*"))
                                {
                                    string[] pp = TheStudyEventOID.Split('*');
                                    theWrittenSE = pp[0];
                                    theSERK = pp[1];
                                }
                                else
                                {
                                    theSERK = "1";
                                    theWrittenSE = TheStudyEventOID;
                                }
                                if (event_index_row != -1) theSERK = split[event_index_row];
                                theXMLForm += "                </ItemGroupData>" + LINE_SEPARATOR;
                                theXMLForm += "            </FormData>" + LINE_SEPARATOR;
                                if (datapresentform)
                                {
                                    datapresentevent = true;
                                    theXMLEvent += theXMLForm;

                                    datapresentform = false;
                                }
                                theXMLForm = "";
                                theXMLEvent += "        </StudyEventData>";
                                if (datapresentevent)
                                {
                                    odmOutputFile.Append(theXMLEvent);
                                    datapresentevent = false;
                                }
                                theXMLEvent = "";
                                if (conversionSettings.selectedEventRepeating == "Yes")
                                {
                                    if (event_index_row != -1) theSERK = split[event_index_row];
                                    else theSERK = Utilities.Get_SE_RepeatingKey_FromStudyDataColumn(theStudyDataColumn);
                                }
                                theXMLEvent += "        <StudyEventData StudyEventOID=\"" + theWrittenSE + "\" StudyEventRepeatKey=\"" + CheckRepeatKey(theSERK, linecount) + "\">" + LINE_SEPARATOR;
                                TheFormOID = nwdingen[1];
                                theXMLForm += "            <FormData FormOID=\"" + TheFormOID + "\">" + LINE_SEPARATOR;
                                TheItemGroupDef = nwdingen[2];
                                if (TheItemGroupDef.Contains("*"))
                                {
                                    string[] pp = TheItemGroupDef.Split('*');
                                    theWrittenGR = pp[0];
                                    theGRRK = pp[1];
                                }
                                else
                                {
                                    theGRRK = "NOT";
                                    theWrittenGR = TheItemGroupDef.Substring(1);
                                }
                                if (group_index_row != -1) theGRRK = split[group_index_row];
                                if (theGRRK == "NOT") theXMLForm += "                <ItemGroupData ItemGroupOID=\"" + theWrittenGR + "\" TransactionType=\"Insert\" >" + LINE_SEPARATOR;
                                else theXMLForm += "                <ItemGroupData ItemGroupOID=\"" + theWrittenGR + "\" ItemGroupRepeatKey=\"" + CheckRepeatKey(theGRRK, linecount) + "\" TransactionType=\"Insert\" >" + LINE_SEPARATOR;
                                if (nwdingen[3] != "none")
                                {
                                    indexOfItem = dataGrid.GetIndexOfItem(dataGrid.dataGridView.Rows[i].Cells[DGIndexOfOCItem].Value.ToString());
                                    check_index_of_item(linecount, indexOfItem);
                                    itemval = split[indexOfItem];
                                    itemval = studyMetaDataValidator.ValidateItem(subjectID, TheFormOID, nwdingen[3], itemval, linecount, conversionSettings.dateFormat);
                                    if (itemval != "")
                                    {
                                        theXMLForm += "                    <ItemData ItemOID=\"" + nwdingen[3] + "\" Value=\"" + SecurityElement.Escape(itemval) + "\" />" + LINE_SEPARATOR;
                                        datapresentform = true;
                                    }
                                }
                            }
                        }
                    }
                    string SubSex = conversionSettings.defaultSubjectSex;
                    if (dataGrid.sexIndex >= 0)
                    {
                        SubSex = split[dataGrid.sexIndex];
                        SubSex = SubSex.Trim();
                    }
                    if (SubSex == conversionSettings.SUBJECTSEX_M) SubSex = "m";  // 1.0f
                    if (SubSex == conversionSettings.SUBJECTSEX_F) SubSex = "f";
                    if (SubSex != "f" && SubSex != "m" && SubSex != "") // 1.1e  Allow nothing to be entered for sex; it is not always mandatory.
                    {
                        string errtext = "Subject sex can be only 'f' or 'm'. You have '" + SubSex + "' at line " + linecount.ToString() + ". Index: " + dataGrid.sexIndex + ".";
                        warningLog.appendMessage(errtext);
                    }
                    theXMLForm += "                </ItemGroupData>" + LINE_SEPARATOR;
                    theXMLForm += "            </FormData>" + LINE_SEPARATOR;
                    if (datapresentform)
                    {
                        datapresentevent = true;
                        theXMLEvent += theXMLForm;
                        datapresentform = false;
                    }
                    theXMLForm = "";
                    theXMLEvent += "        </StudyEventData>";
                    if (datapresentevent)
                    {
                        odmOutputFile.Append(theXMLEvent);

                        datapresentevent = false;
                    }
                    theXMLEvent = "";
                    odmOutputFile.Append("    </SubjectData>");

                    numberOfLinesOutputFile = numberOfLinesOutputFile + 1;
                    if (numberOfLinesOutputFile >= maximalNumberOfLinesOutputFile)  // 1.0f
                    {
                        masterClose(odmOutputFile);
                        numberOfLinesOutputFile = 0;
                        numberOfOutputFiles = numberOfOutputFiles + 1;
                        outputFilePath = outputFileBaseName + numberOfOutputFiles + ".xml";
                        odmOutputFile.Close();
                        odmOutputFile = new OutputFile(outputFilePath);
                        masterHeader(odmOutputFile);
                    }
                    // generate insert statements
                    int theSERKInt = System.Convert.ToInt16(theSERK);
                    string theDOB = "";

                    //if (DOBIndex >= 0)
                    //{
                    //    if (DOY && split[DOBIndex].Trim().Length == 4) theDOB = split[DOBIndex].Trim();
                    //    else theDOB = ConvertToODMFormat(split[DOBIndex]);
                    //}


                    if (dataGrid.DOBIndex >= 0)
                    {
                        if (DOY && split[dataGrid.DOBIndex].Trim().Length == 4) theDOB = split[dataGrid.DOBIndex].Trim() + "-01-01";
                        else theDOB = theDOB = Utilities.ConvertToODMFormat(split[dataGrid.DOBIndex], conversionSettings.dateFormat);
                    }
                    string theSTD = "";
                    if (dataGrid.STDIndex >= 0)
                    {
                        theSTD = Utilities.ConvertToODMFormat(split[dataGrid.STDIndex], conversionSettings.dateFormat); // This is needed for non repeating events
                    }
                    string thePID = "";
                    if (dataGrid.PIDIndex < 0)
                    {
                        thePID = subjectID;
                    }
                    else
                    {
                        thePID = split[dataGrid.PIDIndex];
                    }

                    if (theDOB.StartsWith("Error") || theDOB == "" || dataGrid.DOBIndex < 0)
                    {
                        if (theDOB == "" || dataGrid.DOBIndex < 0)
                        {
                            if (!IsDuplicatePID(thePID))
                            {
                                insertsSQLFile.Append(INSERT_1);
                                insertsSQLFile.Append("    VALUES (1, '" + SubSex + "', '" + thePID + "', '" + theDate + "', 1, '1');");
                            }
                        }
                        else
                        {
                            string errtext = "Invalid subject birth date '" + theDOB + "' at line " + linecount.ToString() + ". Index: " + dataGrid.DOBIndex + ". ";
                            warningLog.appendMessage(errtext);
                        }
                    }
                    else
                    {
                        if (!IsDuplicatePID(thePID))
                        {
                            insertsSQLFile.Append(INSERT_1A);
                            insertsSQLFile.Append("    VALUES (1, '" + SubSex + "', '" + thePID + "', '" + theDate + "', 1, '1', '" + theDOB + "');");
                        }
                    }
                    insertsSQLFile.Append(INSERT_2);
                    // if there is no PID, use the key (unique_identifier) to fill the field label of study_subject. 
                    insertsSQLFile.Append("    VALUES ('" + subjectID + "', (SELECT study_id FROM study WHERE oc_oid = '" + usedStudyOID + "'),");
                    insertsSQLFile.Append("            1, '" + theDate + "', '" + theDate + "', '" + theDate + "', 1, '" + SStheKEY + "', (SELECT subject_id FROM subject where unique_identifier = '" + thePID + "'));");
                    if (theWrittenSE == "none" && conversionSettings.selectedStudyEvent != "-- select --") // 4.1 eliminate -- select --
                    {
                        theWrittenSE = Utilities.GetOID(conversionSettings.selectedStudyEvent); // 2.0.5 Use the selected SE to determine current SE, as there is no CRF data 
                    }
                    if (theWrittenSE != "none")
                    {
                        int startindex = 1;
                        // 2.0.7 
                        for (int say = startindex; say <= theSERKInt; say++)
                        {
                            // date_start of study event -> get it from file. If blank, don't create the event. OC will reject any data reletad to an event without a start date. 
                            string part1 = "";
                            string part2 = "_E" + say.ToString();
                            if (dataGrid.eventStartDates.Contains(part2)) // there is a date; get it
                            {
                                part1 = dataGrid.eventStartDates.Substring(dataGrid.eventStartDates.IndexOf(part2) + 1);
                                part1 = part1.Substring(part1.IndexOf("^") + 1);
                                part1 = part1.Substring(0, part1.IndexOf("$"));  // part1 is the index of date
                                theSTD = Utilities.ConvertToODMFormat(split[System.Convert.ToInt16(part1)], conversionSettings.dateFormat);
                                if (theSTD.StartsWith("Error"))
                                {
                                    string errtext = "Invalid start date '" + theSTD + "' at line " + linecount.ToString() + ". Index: " + dataGrid.STDIndex + ". ";
                                    warningLog.appendMessage(errtext);
                                }
                                else
                                {
                                    if (theSTD != "")
                                    {
                                        insertsSQLFile.Append(INSERT_3);
                                        insertsSQLFile.Append("    VALUES ((SELECT study_event_definition_id FROM study_event_definition WHERE oc_oid = '" + theWrittenSE + "'),");
                                        insertsOnlyEventsSQLFile.Append(INSERT_3);
                                        insertsOnlyEventsSQLFile.Append("    VALUES ((SELECT study_event_definition_id FROM study_event_definition WHERE oc_oid = '" + theWrittenSE + "'),");
                                        insertsSQLFile.Append("	    (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "'),'" + conversionSettings.defaultLocation + "', " + say.ToString() + ", '" + theSTD + " 12:00:00', 1, 1, '" + theDate + "', 1, '0', '0');");
                                        insertsOnlyEventsSQLFile.Append("	     (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "'),'" + conversionSettings.defaultLocation + "', " + say.ToString() + ", '" + theSTD + " 12:00:00', 1, 1, '" + theDate + "', 1, '0', '0');");
                                    }
                                }
                            }
                            else  // no repeating Events; use either todays date as STD or pick it from data file  2.0.9
                            {
                                if (theSTD != "") // there is a date specified in the data file
                                {
                                    insertsSQLFile.Append(INSERT_3);
                                    insertsSQLFile.Append("    VALUES ((SELECT study_event_definition_id FROM study_event_definition WHERE oc_oid = '" + theWrittenSE + "'),");
                                    insertsOnlyEventsSQLFile.Append(INSERT_3);
                                    insertsOnlyEventsSQLFile.Append("    VALUES ((SELECT study_event_definition_id FROM study_event_definition WHERE oc_oid = '" + theWrittenSE + "'),");
                                    insertsSQLFile.Append("	    (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "'),'" + conversionSettings.defaultLocation + "', " + say.ToString() + ", '" + theSTD + " 12:00:00', 1, 1, '" + theDate + "', 1, '0', '0');");
                                    insertsOnlyEventsSQLFile.Append("        (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "'),'" + conversionSettings.defaultLocation + "', " + say.ToString() + ", '" + theSTD + " 12:00:00', 1, 1, '" + theDate + "', 1, '0', '0');");
                                }
                                else  // no date specified in data file
                                {
                                    if (conversionSettings.useTodaysDateIfNoEventDate)  // 2.1.3 Generate inserts for events without dates using todays date, only if user wants to, otherwise 
                                    // TODO ask Cuneyt what otherwise 
                                    {
                                        insertsSQLFile.Append(INSERT_3);
                                        insertsSQLFile.Append("    VALUES ((SELECT study_event_definition_id FROM study_event_definition WHERE oc_oid = '" + theWrittenSE + "'),");
                                        insertsOnlyEventsSQLFile.Append(INSERT_3);
                                        insertsOnlyEventsSQLFile.Append("    VALUES ((SELECT study_event_definition_id FROM study_event_definition WHERE oc_oid = '" + theWrittenSE + "'),");
                                        insertsSQLFile.Append("	    (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "'),'" + conversionSettings.defaultLocation + "', " + say.ToString() + ", '" + theDate + " 12:00:00', 1, 1, '" + theDate + "', 1, '0', '0');");
                                        insertsOnlyEventsSQLFile.Append("        (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "'),'" + conversionSettings.defaultLocation + "', " + say.ToString() + ", '" + theDate + " 12:00:00', 1, 1, '" + theDate + "', 1, '0', '0');");
                                    }
                                }
                            }
                        }
                        deletesOnlyEventsSQLFile.Append("DELETE FROM study_event where study_subject_id = (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "');");
                    }
                    deletesSQLFile.Append("DELETE FROM study_event where study_subject_id = (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "');");
                    deletesSQLFile.Append("DELETE FROM study_subject where oc_oid = '" + SStheKEY + "';");
                    deletesSQLFile.Append("DELETE FROM subject where unique_identifier = '" + thePID + "';");
                    // Control hidden values and mandatory values

                    studyMetaDataValidator.validateHiddens(TheFormOID, linecount, SStheKEY);
                }
                masterClose(odmOutputFile);
            }
            catch (Exception ex)
            {
                string errtext = "Exception while reading data file: " + ex;
                if (errtext.Contains("ThreadAbortException") == false)
                {
                    throw new OCDataImporterException(errtext, ex);
                }
            }
            finally
            {
                deletesOnlyEventsSQLFile.Close();
                deletesSQLFile.Close();
                insertsOnlyEventsSQLFile.Close();
                insertsSQLFile.Close();
                odmOutputFile.Close();
            }
        }
       

        public void check_index_of_item(int linecount, int myioi)
        {
            if (myioi < 0)
            {
                string errtext = " Wrong index at: " + linecount.ToString() + ". Exiting...The generated files ARE INCOMPLETE AND CAN NOT BE USED";
                throw new OCDataImporterException(errtext);
                //MessageBox.Show(errtext, "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                //exit_error(errtext);
            }
        }            


        private void masterHeader(OutputFile outputFile)
        {
            String timeStamp = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");
            outputFile.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            outputFile.Append("<ODM xmlns=\"http://www.cdisc.org/ns/odm/v1.3\"");
            outputFile.Append("xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"");
            outputFile.Append("xsi:schemaLocation=\"http://www.cdisc.org/ns/odm/v1.3 ODM1-3.xsd\"");
            outputFile.Append("ODMVersion=\"1.3\" FileOID=\"1D20080412202420\" FileType=\"Snapshot\"");
            outputFile.Append("Description=\"Dataset ODM\" CreationDateTime=\"" + timeStamp + "\" >");
            outputFile.Append("<ClinicalData StudyOID=\"" + conversionSettings.studyOID + "\" MetaDataVersionOID=\"v1.0.0\">");
        }

        private void masterClose(OutputFile outputFile)
        {
            outputFile.Append("</ClinicalData>");
            outputFile.Append("</ODM>");
        }

        private bool IsDuplicatePID(string thePIDtoCheck)
        {
            bool found = false;
            foreach (string one in InsertSubject)
            {
                if (one == thePIDtoCheck)
                {
                    found = true;
                    break;
                }
            }
            if (!found)
            {
                InsertSubject.Add(thePIDtoCheck);
            }
            return (found);
        }

        private string CheckRepeatKey(string rk, int line)
        {
            if (Utilities.IsNumber(rk)) return (rk);
            warningLog.appendMessage("Event and/or Group repeat index can't be determined at line: " + line.ToString());
            return ("ERROR: " + rk);
        }

    }

}
