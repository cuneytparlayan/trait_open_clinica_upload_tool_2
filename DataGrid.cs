
using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace OCDataImporter
{
    /// <summary>
    /// Class for a data grid which contains the column definitions of the data to convert to CDISC ODM. Can be considered
    /// as a model class which werks with a DataGridView.
    /// </summary>
    class DataGrid
    {

        // Places in Data Grid
        public const int DGIndexOfDataItem = 0;
        public const int DGIndexOfOCItem = 1;
        public const int DGIndexOfKey = 2;
        public const int DGIndexOfDate = 3;
        public const int DGIndexOfSex = 4;
        public const int DGIndexOfPID = 5;
        public const int DGIndexOfDOB = 6;
        public const int DGIndexOfSTD = 7;
        private const char TAB = '\u0009';

        public int subjectIDIndex { get; private set; }
        public int sexIndex { get; private set; }
        public int PIDIndex { get; private set; }
        public int DOBIndex { get; private set; }
        public int STDIndex { get; private set; }
        public String sexItem { get; private set; }
        public String eventStartDates { get; private set; }
        public char delimiter { get; private set; }
        public int sepcount { get; private set; }


        private int maxSE = 0;
        private int maxCRF = 0;
        private int maxGR = 0;
        
        /// <summary>
        /// The length of an input line. Can be used to tell a progressbar how must characters it must process.
        /// </summary>
        private int linelen;
        
        private string PIDItem = "";
        private string DOBItem = "";
        
        public string startDateItem {get; set;}
                      

        private DataGridView dataGridView;
        

        private ArrayList dataFileItems = new ArrayList();
        private StudyMetaDataValidator studyMetaDataValidator;
        private IViewUpdater viewUpdater;
        private ConversionSettings conversionSettings;

        public DataGrid(ConversionSettings conversionSettings, StudyMetaDataValidator studyMetaDataValidator, DataGridView dataGridView, IViewUpdater viewUpdater)
        {
            this.dataGridView = dataGridView;            
            this.conversionSettings = conversionSettings;
            this.studyMetaDataValidator = studyMetaDataValidator;
            this.viewUpdater = viewUpdater;
            reset();   
        }


        public void reset()
        {
            dataFileItems = new ArrayList();
            subjectIDIndex = 0;
            sexIndex = 0;
            PIDIndex = 0;
            DOBIndex = 0;
            STDIndex = 0;
            startDateItem = "";
            eventStartDates = "";
            delimiter = ';';
            sepcount = 1;
        }

        public bool GetDataFileItemsFromInput()
        {
            // Find out how many data items are present per line and build array of data item names for using in data grid
            dataFileItems.Clear();
            sepcount = 1;
            try
            {
                using (StreamReader sr = new StreamReader(conversionSettings.pathToInputFile))
                {
                    String line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        line = line.Trim();  // 1.1b
                        if (line.Length == 0) continue;
                        linelen = line.Length;
                        if (line.IndexOf(TAB) > 0) delimiter = TAB;
                        if (line.IndexOf(';') > 0) delimiter = ';';

                        for (int i = 0; i < line.Length; i++) if (line[i] == delimiter) sepcount++;
                        string[] spfirst = line.Split(delimiter);
                        foreach (string one in spfirst) dataFileItems.Add(one);
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return (false);
            }
            return (true);
        }

        /// <summary>
        /// Performs checks on the data grid. Returns <code>false</code> and displays an error message if a problem 
        /// is found.
        /// </summary>        
        public bool isValid()
        {
            ArrayList SortableDG = new ArrayList();
            for (int i = 0; i < dataGridView.RowCount; i++)
            {
                if (dataGridView.Rows[i].IsNewRow == false)
                {
                    string[] mi = dataGridView.Rows[i].Cells[DGIndexOfOCItem].Value.ToString().Split('.');
                    if (mi.Length == 1 && (mi[0] == "none" || mi[0].StartsWith("Use link button"))) continue;
                    if (mi[0].Length > maxSE) maxSE = mi[0].Length;
                    if (mi[1].Length > maxCRF) maxCRF = mi[1].Length;
                    if (mi[2].Length > maxGR) maxGR = mi[2].Length;
                }
            }
            int sexCount = 0;
            int PIDCount = 0;
            int DOBCount = 0;
            int STDCount = 0;            
            PIDIndex = -1;
            sexIndex = -1;
            DOBIndex = -1;
            STDIndex = -1;
            sexItem = "";
            PIDItem = "";
            DOBItem = "";
            
            for (int i = 0; i < dataGridView.RowCount; i++)
            {
                if (dataGridView.Rows[i].IsNewRow == false)
                {
                    if (dataGridView.Rows[i].Cells[DGIndexOfSex].Value.ToString() == "True")
                    {
                        sexCount++;
                        for (int j = 0; j < dataFileItems.Count; j++)
                        {
                            if (dataFileItems[j].ToString() == dataGridView.Rows[i].Cells[DGIndexOfDataItem].Value.ToString())
                            {
                                sexIndex = j;
                                sexItem = dataGridView.Rows[i].Cells[DGIndexOfDataItem].Value.ToString();
                            }
                        }
                    }
                    if (dataGridView.Rows[i].Cells[DGIndexOfPID].Value.ToString() == "True")
                    {
                        PIDCount++;
                        for (int j = 0; j < dataFileItems.Count; j++)
                        {
                            if (dataFileItems[j].ToString() == dataGridView.Rows[i].Cells[DGIndexOfDataItem].Value.ToString())
                            {
                                PIDIndex = j;
                                PIDItem = dataGridView.Rows[i].Cells[DGIndexOfDataItem].Value.ToString();
                            }
                        }
                    }
                    if (dataGridView.Rows[i].Cells[DGIndexOfDOB].Value.ToString() == "True")
                    {
                        DOBCount++;
                        for (int j = 0; j < dataFileItems.Count; j++)
                        {
                            if (dataFileItems[j].ToString() == dataGridView.Rows[i].Cells[DGIndexOfDataItem].Value.ToString())
                            {
                                DOBIndex = j;
                                DOBItem = dataGridView.Rows[i].Cells[DGIndexOfDataItem].Value.ToString();
                            }
                        }
                    }
                    if (dataGridView.Rows[i].Cells[DGIndexOfSTD].Value.ToString() == "True")
                    {
                        // 2.0.7 put the STDIndex and STDItems in an ARRAY as it is possible that in a repeating event more START dates are given!
                        STDCount++;
                        for (int j = 0; j < dataFileItems.Count; j++)
                        {
                            if (dataFileItems[j].ToString() == dataGridView.Rows[i].Cells[DGIndexOfDataItem].Value.ToString())
                            {
                                STDIndex = j;
                                startDateItem = dataGridView.Rows[i].Cells[DGIndexOfDataItem].Value.ToString();
                                eventStartDates += startDateItem + "^" + STDIndex.ToString() + "$";
                            }
                        }
                    }

                    string[] mi = dataGridView.Rows[i].Cells[DGIndexOfOCItem].Value.ToString().Split('.');
                    if (mi.Length == 1 && (mi[0] == "none" || mi[0].StartsWith("Use link button")))
                    {
                        if (dataGridView.Rows[i].Cells[DGIndexOfKey].Value.ToString() == "True")
                        {
                            SortableDG.Add("none" + "^" + dataGridView.Rows[i].Cells[DGIndexOfDataItem].Value.ToString() + "^" + dataGridView.Rows[i].Cells[DGIndexOfKey].Value.ToString() + "^" + dataGridView.Rows[i].Cells[DGIndexOfDate].Value.ToString() + "^" + dataGridView.Rows[i].Cells[DGIndexOfSex].Value.ToString() + "^" + dataGridView.Rows[i].Cells[DGIndexOfPID].Value.ToString() + "^" + dataGridView.Rows[i].Cells[DGIndexOfDOB].Value.ToString() + "^" + dataGridView.Rows[i].Cells[DGIndexOfSTD].Value.ToString());
                        }
                        else continue;
                    }
                    else
                    {
                        string theDataItem = dataGridView.Rows[i].Cells[DGIndexOfDataItem].Value.ToString();
                        string ev_rep = "";
                        string gr_rep = "";
                        if (studyMetaDataValidator.repeatingEvents.Contains(mi[0])) ev_rep = Utilities.Get_SE_RepeatingKey_FromStudyDataColumn(theDataItem);
                        if (studyMetaDataValidator.repeatingGroups.Contains(mi[2])) gr_rep = Utilities.Get_GR_RepeatingKey_FromStudyDataColumn(theDataItem);
                        mi[0] = Utilities.FillTildes(mi[0], maxSE);
                        mi[1] = Utilities.FillTildes(mi[1], maxCRF);
                        mi[2] = Utilities.FillTildes(mi[2], maxGR);
                        if (ev_rep != "") mi[0] = mi[0] + "*" + ev_rep;
                        if (gr_rep != "") mi[2] = mi[2] + "*" + gr_rep;
                        else mi[2] = "A" + mi[2];  // This is due to an OC bug; UNGROUPED items must come before the grouped items in the grid.
                        SortableDG.Add(mi[0] + "." + mi[1] + "." + mi[2] + "." + mi[3] + "^" + theDataItem + "^" + dataGridView.Rows[i].Cells[DGIndexOfKey].Value.ToString() + "^" + dataGridView.Rows[i].Cells[DGIndexOfDate].Value.ToString() + "^" + dataGridView.Rows[i].Cells[DGIndexOfSex].Value.ToString() + "^" + dataGridView.Rows[i].Cells[DGIndexOfPID].Value.ToString() + "^" + dataGridView.Rows[i].Cells[DGIndexOfDOB].Value.ToString() + "^" + dataGridView.Rows[i].Cells[DGIndexOfSTD].Value.ToString());
                    }
                }
            }
            int keyCount = 0;
            for (int i = 0; i < SortableDG.Count; i++)
            {
                string[] tr = SortableDG[i].ToString().Split('^');
                if (tr[2] == "True")
                {
                    keyCount++;
                    for (int j = 0; j < dataFileItems.Count; j++) if (dataFileItems[j].ToString() == tr[1]) subjectIDIndex = j;
                }

            }
            if (keyCount != 1)
            {
                MessageBox.Show("Please select (only) one field as STUDY SUBJECT ID by using check box; You have " + keyCount.ToString() + " selected.", "OCDataImporter");
                return false;
            }
            if (sexCount > 1)
            {
                MessageBox.Show("Please select at most one field as STUDY SUBJECT SEX by using check box; You have " + sexCount.ToString() + " selected.", "OCDataImporter");
                return false;
            }
            if (PIDCount > 1)
            {
                MessageBox.Show("Please select at most one field as PERSON ID by using check box; You have " + PIDCount.ToString() + " selected.", "OCDataImporter");
                return false;
            }
            if (DOBCount > 1)
            {
                MessageBox.Show("Please select at most one field as SUBJECT DATE OF BIRTH by using check box; You have " + DOBCount.ToString() + " selected.", "OCDataImporter");
                return false;
            }
            SortableDG.Sort();
            dataGridView.Rows.Clear();
            for (int i = 0; i < SortableDG.Count; i++)
            {
                string[] fnparts;
                fnparts = new string[dataGridView.ColumnCount];
                string[] tr = SortableDG[i].ToString().Split('^');
                fnparts[DGIndexOfDataItem] = tr[1];
                fnparts[DGIndexOfKey] = tr[2];
                fnparts[DGIndexOfDate] = tr[3];
                fnparts[DGIndexOfSex] = tr[4];
                fnparts[DGIndexOfPID] = tr[5];
                fnparts[DGIndexOfDOB] = tr[6];
                fnparts[DGIndexOfSTD] = tr[7];
                fnparts[DGIndexOfOCItem] = tr[0].ToString().Replace("~", "");
                dataGridView.Rows.Add(fnparts);
            }
            return true;
        }

        public void BuildDG(bool matchcolumns)
        {
            string key = "False";
            string dat = "False";
            string sex = "False";
            string pid = "False";
            string dob = "False";
            string std = "False";
            viewUpdater.resetText();
            string[] fnparts;
            fnparts = new string[dataGridView.ColumnCount];
            if (conversionSettings.selectedStudyEvent != "-- select --" && conversionSettings.selectedCRF != "-- select --")
            {
                int matched = 0;
                bool nomatch = false;
                if (!matchcolumns) dataGridView.Rows.Clear();
                for (int i = 0; i < dataFileItems.Count; i++)
                {
                    if (dataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("subjectid")) key = "True";
                    else key = "False";
                    if (dataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("subjectsex") ||
                        dataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("gender") ||
                        dataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("geslacht")) sex = "True";
                    else sex = "False";
                    if (dataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("personid")) pid = "True";
                    else pid = "False";
                    if (dataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("subjectdateofbirth") ||
                        dataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("subjectdob") ||
                        dataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("dateofbirth") ||
                        dataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("geboortedatum")) dob = "True";
                    else dob = "False";
                    if (dataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("subjectstartdate")) std = "True";
                    else std = "False";
                    fnparts[DGIndexOfDataItem] = dataFileItems[i].ToString();
                    fnparts[DGIndexOfKey] = key;
                    fnparts[DGIndexOfDate] = dat;
                    fnparts[DGIndexOfSex] = sex;
                    fnparts[DGIndexOfPID] = pid;
                    fnparts[DGIndexOfDOB] = dob;
                    fnparts[DGIndexOfSTD] = std;
                    
                    fnparts[DGIndexOfOCItem] = "none";

                    int startGroup = 0;
                    int startEvent = 0;
                    // track _Gx or _Ex in column names and exclude those extentions from column matches. Fixed: 2.0.3
                    for (int jj = 0; jj < 10; jj++)
                    {
                        startGroup = dataFileItems[i].ToString().IndexOf("_G" + jj.ToString());
                        if (startGroup > 0) break;
                    }
                    for (int jj = 0; jj < 10; jj++)
                    {
                        startEvent = dataFileItems[i].ToString().IndexOf("_E" + jj.ToString());
                        if (startEvent > 0) break;
                    }
                    string theItem = dataFileItems[i].ToString().ToLower();
                    if (startEvent > 0) theItem = theItem.Substring(0, startEvent);
                    else if (startGroup > 0) theItem = theItem.Substring(0, startGroup);

                    string theItemOID = studyMetaDataValidator.GetItemOIDFromItemName(theItem, conversionSettings.selectedCRF);
                    if (theItemOID != "NOTFOUND2" && theItemOID != "ANOTHERCRF") // 3.03
                    {
                        fnparts[DGIndexOfOCItem] = conversionSettings.selectedStudyEvent + "." + conversionSettings.selectedCRF + "." + studyMetaDataValidator.GetGroupFromItemCRF(theItemOID, conversionSettings.selectedCRF) + "." + theItemOID;
                        matched++;
                    }
                    else
                    {
                        if (theItemOID == "NOTFOUND2" && key != "True" && sex != "True" && pid != "True" && dob != "True" && std != "True" && theItem != "event_index" && theItem != "group_index")
                        {
                            viewUpdater.appendText("No match for: " + theItem + System.Environment.NewLine);
                            nomatch = true;
                        }
                    }

                    if (!matchcolumns) dataGridView.Rows.Add(fnparts);
                    else    // version 2.0.2
                    {
                        for (int n = 0; n < dataGridView.RowCount; n++)
                        {
                            if (dataGridView.Rows[n].IsNewRow == false)
                            {
                                string gridDataItem = dataGridView.Rows[n].Cells[DGIndexOfDataItem].Value.ToString();
                                string gridOCItem = dataGridView.Rows[n].Cells[DGIndexOfOCItem].Value.ToString();
                                if (gridDataItem == fnparts[DGIndexOfDataItem])
                                {
                                    if (gridOCItem == "none" || gridOCItem.StartsWith("Use link button"))
                                    {
                                        dataGridView.Rows[n].Cells[DGIndexOfOCItem].Value = fnparts[DGIndexOfOCItem];
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
                if (nomatch)
                {
                    MessageBox.Show("Not all Items in the selected CRF could be matched. For the list of UNMATCHED Items, see the progress textbox below. You can match those items by using the comboboxes above. Control the matched items too, as the matching can not be 100% correct!", "OCDataImporter");
                }
                else
                {
                    MessageBox.Show("All Items in the selected CRF could be matched. Control the matched items as the matching can not be 100% correct!", "OCDataImporter");
                }
            }
            else
            {
                dataGridView.Rows.Clear();
                for (int i = 0; i < dataFileItems.Count; i++)
                {
                    if (dataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("subjectid")) key = "True";
                    else key = "False";
                    if (dataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("subjectsex") ||
                        dataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("gender") ||
                        dataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("geslacht")) sex = "True";
                    else sex = "False";
                    if (dataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("personid")) pid = "True";
                    else pid = "False";
                    if (dataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("subjectdateofbirth") ||
                        dataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("subjectdob") ||
                        dataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("dateofbirth") ||
                        dataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("geboortedatum")) dob = "True";
                    else dob = "False";
                    if (dataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("subjectstartdate")) std = "True";
                    else std = "False";
                    fnparts[DGIndexOfDataItem] = dataFileItems[i].ToString();
                    fnparts[DGIndexOfKey] = key;
                    fnparts[DGIndexOfDate] = dat;
                    fnparts[DGIndexOfSex] = sex;
                    fnparts[DGIndexOfPID] = pid;
                    fnparts[DGIndexOfDOB] = dob;
                    fnparts[DGIndexOfSTD] = std;
                    fnparts[DGIndexOfOCItem] = "Use link button 'CopyTarget' to fill this cell with the selected target item";
                    dataGridView.Rows.Add(fnparts);
                }
            }
        }

        public int GetIndexOfItem(string item)
        {
            
            for (int i = 0; i < dataGridView.RowCount; i++)
            {
                if (dataGridView.Rows[i].IsNewRow == false)
                {
                    if (dataGridView.Rows[i].Cells[DataGrid.DGIndexOfOCItem].Value.ToString() == "none") continue;
                    if (dataGridView.Rows[i].Cells[DataGrid.DGIndexOfOCItem].Value.ToString() == item)
                    {
                        string myDataItem = dataGridView.Rows[i].Cells[DataGrid.DGIndexOfDataItem].Value.ToString();
                        for (int j = 0; j < dataFileItems.Count; j++) if (dataFileItems[j].ToString() == myDataItem) return (j);
                    }
                }
            }
            return (-1);
        } 
    }
}