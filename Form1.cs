using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using System.Globalization;
using System.Diagnostics;
using System.Collections;
using System.Xml;
using System.Security.AccessControl;

//1.0	    21/05/2010	    Initial version	                    C. Parlayan, J.A.M. Beliën
//1.1	    21/04/2012	    Production version	                C. Parlayan
//2.0       08/11/2012      Production 2.0                      C. Parlayan
//2.0.1     01/02/2013      Added Repeating events and groups   C. Parlayan
//2.0.2	    26/03/2013      The GRID now keeps the previously 
//                          matched items when a new CRF or 
//                          Group is chosen for more matching.  C. Parlayan
//2.0.3	    01/04/2013	    Bug solved in matching with 
//                          repeated items.	                    C. Parlayan
//2.0.4	    01/05/2013	    Remove white space when matching	C. Parlayan
//2.0.5	    21/07/2013	    Do not print form data with no 
//                          items in XML as this causes error 
//                          in OpenClinica upload.	            C. Parlayan, S. de Ridder
//2.0.6	    26/08/2013	    Added “Limit number of characters 
//                          to match” to make matching easier	C. Parlayan
//2.0.7/8/9	29/08/2013	    Bugs introduced in 2.0.5 solved.	C. Parlayan, S. de Ridder
//2.1.1	    09/09/2013	    Introduced label-oid file.	        C. Parlayan, J. Rousseau, R. Voorham
//2.1.2	    01/11/2013	    Bug ItemGroupRepeatKey solved.	    C. Parlayan, R. Voorham
//2.1.3	    18/11/2013	    Possibility to not generate 
//                          events if startdate is blank.	    C. Parlayan, S. de Ridder
//2.1.4	    01/11/2013	    XML escaping.	                    C. Parlayan, J. Rousseau
//2.2       30/11/2013      Input file allows EVENT_INDEX 
//                          and GROUP_INDEX to be defined 
//                          to accept repeating events/items
//                          in rows                             C. Parlayan, J. Rousseau
//3.0       20/12/2013      Type and range validations,
//                          better messaging                    C. Parlayan
//3.01      20-01-2014      XML parsing error while RangeCheck  J. Rousseau, C. Parlayan
//3.02      20-01-2014      Warnings to status window           J. Rousseau, C. Parlayan
//3.03      22-01-2014      matching now uses "ItemName"  
//                          instead of ItemOID                  C. Parlayan, S. de Ridder
//3.03      24-01-2014      matching is done for all groups  
//                          in a CRF in one go                  C. Parlayan, S. de Ridder
//4.0       29-01-2014      Better UI                           C. Parlayan
//4.01      07-02-2014      Show/hide Subj related columns
//                          bugfixes on validation              C. Parlayan, J. Rousseau, S. de Ridder
//4.1       10-02-2014      Auto detection and show/hide         
//                          subject related columns             C. Parlayan, J. Rousseau, S. de Ridder
//4.1       10-02-2014      Added Multiple select to         
//                          validations                         C. Parlayan, J. Rousseau
//4.1       10-02-2014      Leading zeroes in times and           
//                          subjectid's in log records          C. Parlayan, J. Rousseau
//4.1       10-02-2014      fix: "-- select --" in insert            
//                          statements                          C. Parlayan, S. de Ridder
//4.2       17-02-2014      Show/hidden warn ONLY if show and 
//                          has no value                        C. Parlayan, S. de Ridder
//4.2.1     19-02-2014      Ignore range check if other 
//                          problems detected with data         C. Parlayan, S. de Ridder
//4.3       21-02-2014      SE/GR repeating index in rows 
//                          bugs fixed                          C. Parlayan, J. Rousseau
//4.3       21-02-2014      Better error messages if input
//                          file is open by other programs      C. Parlayan, J. Rousseau

namespace OCDataImporter
{
    public partial class Form1 : Form
    {
        
        public bool DEBUGMODE = true;
        public bool labelOCoidExists = false; // 2.1.1 If labelOCoid file exists, get the oid from that file, instead of 'SS_label'...
        public string dmpfilename = "";
        public string dmpprm = "";
        Thread MyThread;
        FileStream fpipdf;
        string Mynewline = System.Environment.NewLine;
        static public string workdir = "";
        static public string input_oc = "";
        static public string input_data = "";
        static public string input_oid = "";
        ArrayList Items = new ArrayList();
        ArrayList DataFileItems = new ArrayList();
        ArrayList LabelOID = new ArrayList();
        ArrayList ReplacePairStrings = new ArrayList();
        ArrayList SortableDG = new ArrayList();
        ArrayList InsertSubject = new ArrayList();
        ArrayList ItemGroupDefs = new ArrayList();
        ArrayList ItemDefForm = new ArrayList();
        ArrayList CodeList = new ArrayList();
        ArrayList MSList = new ArrayList();
        ArrayList RCList = new ArrayList();
        ArrayList SCOList = new ArrayList();
        ArrayList Warnings = new ArrayList();
        ArrayList AllValuesInOneRow = new ArrayList();
        ArrayList Hiddens = new ArrayList();
        string TheStudyOID = "";
        string TheStudyEventOID = "";
        string EventStartDates = "";
        string TheItemId = "";
        string TheFormOID = "";
        string TheItemGroupDef = "";
        char RepSeparator = ';';
        ArrayList InsertKeys = new ArrayList();
        FileStream fpoDIM;
        FileStream fpoINS;
        FileStream fpoLOG;
        FileStream fpoINSR;
        FileStream fpoDEL;
        FileStream fpoDELR;
        string INSF = "";
        string LOG = "";
        string INSFR = "";
        string DIMF = "";
        string DELF = "";
        string DELFR = "";
        char Delimiter = ';';
        char tab = '\u0009';
        int sepcount = 1;
        int PROGBARSIZE = 0;
        int WARCOUNT = 0;
        int keyIndex = 0;
        int sexIndex = 0;
        int PIDIndex = 0;
        int DOBIndex = 0;
        int STDIndex = 0;
        string sexItem = "";
        string PIDItem = "";
        string DOBItem = "";
        string STDItem = "";
        int linelen = 0;
        // Places in Data Grid
        private const int DGIndexOfDataItem = 0;
        private const int DGIndexOfOCItem = 1;
        private const int DGIndexOfKey = 2;
        private const int DGIndexOfDate = 3;
        private const int DGIndexOfSex = 4;
        private const int DGIndexOfPID = 5;
        private const int DGIndexOfDOB = 6;
        private const int DGIndexOfSTD = 7;
        string SUBJECTSEX_M = "";
        string SUBJECTSEX_F = "";
        string repeating_groups = "";
        string repeating_events = "";
        // ODM file splitting
        string OUTF = "";
        string OUTFBASIS = "";
        int OUTFLINECOUNTER = 0;
        int OUTFFILECOUNTER = 1;
        int OUTFMAXLINES = 0;
        string selectedEventRepeating = "No";
        
        static public string insert1a = "INSERT INTO subject(status_id, gender, unique_identifier, date_created, owner_id, dob_collected, date_of_birth)";
        static public string insert1 = "INSERT INTO subject(status_id, gender, unique_identifier, date_created, owner_id, dob_collected)";
        static public string insert2 = "INSERT INTO study_subject(label, study_id, status_id, enrollment_date, date_created, date_updated, owner_id,  oc_oid, subject_id)";
        static public string insert3 = "INSERT INTO study_event(study_event_definition_id, study_subject_id, location, sample_ordinal, date_start, owner_id, status_id, date_created, subject_event_status_id, start_time_flag, end_time_flag)";
        
        public Form1()
        {
            InitializeComponent();
            Menu = new MainMenu();
            if (DEBUGMODE) MessageBox.Show("D E B U G M O D E - !!!!!", "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            Menu.MenuItems.Add("&File");
            Menu.MenuItems[0].MenuItems.Add("&Open", new EventHandler(MenuFileOpenOnClick));
            Menu.MenuItems[0].MenuItems.Add("&Exit", new EventHandler(MenuFileExitOnClick));
            Menu.MenuItems.Add("&Help");
            Menu.MenuItems[1].MenuItems.Add("&User Manual", new EventHandler(MenuHelpHowToOnClick));
            Menu.MenuItems[1].MenuItems.Add("&About", new EventHandler(MenuHelpAboutOnClick));
            dataGridView1.RowsDefaultCellStyle.BackColor = Color.Bisque;
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige;
            textBoxLocation.Text = "Amsterdam"; // Our default location
            radioButtonUseTD.Checked = true;  // this should be added to dmpprm, in later releases.
            textBoxReplace.Text = "// Enter replace couples here\r\n// Everthing after a // (like this line) is a comment\r\n// Format is: STUDYDATACOLUMN-ITEM;old;new. Example:\r\n\r\n//    MYDATACOL1;Amsterdam;A'dam\r\n//    MYDATACOL2;1;M\r\n\r\n// Use <null> to replace with null:\r\n// MYDATACOL3;.00;<null>\r\n\r\n// To change the separator, use SEPARATOR command: \r\n//    SEPARATOR=$\r\n//    MYDATACOL3$.00$<null>\r\n\r\n// Use ALL to apply changes to all fields:\r\n//    ALL;.00;<null>\r\n\r\n// To concatenate strings: Use a + sign either at the beginning or at the end, depending on whether you want to concatenate before or after.\r\n//    MYDATACOL1;+thestringtoaddAFTER   (this will add the string at the end of it's original value)  or\r\n//    MYDATACOL1;thestringtoaddBEFORE+   (this will add the string at the beginning of it's original value)";
            try
            {
                fpipdf = new FileStream("OCDataImporter.pdf", FileMode.Open, FileAccess.Read);
            }
            catch (Exception exx)
            {
                MessageBox.Show("Problem opening user manual. Message = " + exx.Message, "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            workdir = Directory.GetCurrentDirectory();
            StateReadFiles();
        }

        public bool Get_DataFileItems_FrpmInput(string theInputFile)
        {
            // Find out how many data items are present per line and build array of data item names for using in data grid
            DataFileItems.Clear();
            sepcount = 1;
            try
            {
                using (StreamReader sr = new StreamReader(theInputFile))
                {
                    String line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        line = line.Trim();  // 1.1b
                        if (line.Length == 0) continue;
                        linelen = line.Length;
                        if (line.IndexOf(tab) > 0) Delimiter = tab;
                        if (line.IndexOf(';') > 0) Delimiter = ';';

                        for (int i = 0; i < line.Length; i++) if (line[i] == Delimiter) sepcount++;
                        string[] spfirst = line.Split(Delimiter);
                        foreach (string one in spfirst) DataFileItems.Add(one);
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

        public void Get_label_oid(string theInputFile)
        {
            // Get the abnormal oid's from file.
            LabelOID.Clear();
            try
            {
                using (StreamReader sr = new StreamReader(theInputFile))
                {
                    String line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        line = line.Trim();
                        if (line.Length == 0) continue;
                        string[] theAbnormal = line.Split(Delimiter);
                        LabelOID.Add(theAbnormal[0].Trim() + "^" + theAbnormal[1].Trim());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can't open selected data file " + input_oid + ", can't continue. Delimiter = " + Delimiter + ", Items per line should be 2 (label and oid) ", "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                exit_error(ex.ToString());
                return;
            }
        }

        public void Get_work_file()
        {
            MyFileDialog mfd = new MyFileDialog(workdir);
            if (mfd.fntxt == "" || mfd.fnxml == "") return;
            textBoxInput.Text = mfd.fnxml + ";" + mfd.fntxt + ";" + mfd.fnoid;
            if (mfd.fnoid != "") labelOCoidExists = true;
        }

        void MenuFileOpenOnClick(object obj, EventArgs ea)
        {
            Get_work_file();
        }

        void MenuFileExitOnClick(object obj, EventArgs ea)
        {
            fpipdf.Close();
            Close();
        }
        void MenuHelpHowToOnClick(object obj, EventArgs ea)
        {
            try
            {
                Process myProcess = new Process();
                myProcess.StartInfo.FileName = "acrord32.exe";
                myProcess.StartInfo.Arguments += '"' + fpipdf.Name + '"';
                myProcess.Start();
            }
            catch (Exception e)
            {
                MessageBox.Show(" Failed to start Acrobat reader: \n\n" + e);
            }
        }

        void MenuHelpAboutOnClick(object obj, EventArgs ea)
        {
            MessageBox.Show("OCDataImporter Version 4.x Made by: C. Parlayan, VU Medical Center, Dept. of Pathology, Amsterdam, The Netherlands - 2010-2014", Text);
        }
        private string FillTildes(string var, int len)
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

        public void FillTheParams()
        {
            try
            {
                using (StreamReader sr = new StreamReader(dmpprm))
                {
                    String line;
                    bool first = true;
                    while ((line = sr.ReadLine()) != null)
                    {
                        line.Trim();
                        if (first)
                        {
                            string[] split = line.Split('~');
                            comboBoxDateFormat.SelectedItem = split[0];
                            comboBoxSex.SelectedItem = split[1];
                            textBoxSubjectSexM.Text = split[2];
                            textBoxSubjectSexF.Text = split[3];
                            textBoxMaxLines.Text = split[4];
                            textBoxReplace.Text = split[6];
                            textBoxLocation.Text = split[5];
                            first = false;
                        }
                        else
                        {
                            textBoxReplace.Text += Mynewline + line;
                        }
                    }
                    return;
                }
            }
            catch (Exception ex)
            {
                textBoxOutput.Text += Mynewline + "No parameters file found." + Mynewline + ex.Message;
                return;
            }
        }

        public bool FillTheGrid()
        {
            try
            {
                using (StreamReader sr = new StreamReader(dmpfilename))
                {
                    DialogResult dlgResult = MessageBox.Show("Do you want to load your previous grid?", "Open Clinica Data Importer", MessageBoxButtons.YesNo);
                    if (dlgResult == DialogResult.Yes)
                    {
                        String line;
                        // NaamPatient~SE_JAARLIJK.F_RAPA_3.IG_RAPA_UNGROUPED.I_RAPA_NAAMPATIENT~False~False~False~False~False~False~False~CopyTarget
                        // PatientNummer~SE_JAARLIJK.F_RAPA_3.IG_RAPA_UNGROUPED.I_RAPA_PATIENTNUMMER~True~False~False~False~False~False~False~CopyTarget
                        // Geslacht~SE_JAARLIJK.F_RAPA_3.IG_RAPA_UNGROUPED.I_RAPA_GESLACHT~False~False~True~False~False~False~False~CopyTarget
                        // etc...
                        while ((line = sr.ReadLine()) != null)
                        {
                            line.Trim();
                            if (line.Length == 0) continue;
                            string[] split = line.Split('~');
                            dataGridView1.Rows.Add(split);
                        }
                        FillTheParams();
                        return true;
                    }
                    return false;
                }
            }
            catch (Exception ex)
            {
                textBoxOutput.Text += Mynewline + "No dmp file found." + Mynewline + ex.Message + Mynewline;
                return false;
            }
        }

        private void dumpTheGrid()
        {
            string oneline = "";
            try
            {
                using (StreamWriter outDMP = new StreamWriter(dmpfilename))
                {
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        if (dataGridView1.Rows[i].IsNewRow) break;
                        oneline = "";
                        for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                        {
                            if (oneline != "") oneline += "~";
                            oneline += dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                        outDMP.WriteLine(oneline);
                    }
                }
                using (StreamWriter outPR = new StreamWriter(dmpprm))
                {
                    outPR.WriteLine(comboBoxDateFormat.SelectedItem + "~" + comboBoxSex.SelectedItem + "~" + textBoxSubjectSexM.Text + "~" + textBoxSubjectSexF.Text + "~" + textBoxMaxLines.Text + "~" + textBoxLocation.Text + "~" + textBoxReplace.Text);
                }
            }
            catch (Exception ex)
            {
                if (ex.ToString().Contains("NullReferenceException")) oneline = "none";
                else
                {
                    MessageBox.Show("Can't generate grid dump file (see log file for details - Do you have enough permissions to write in target folder?)", "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    exit_error(ex.ToString());
                }
            }
        }

        private void masterHeader()
        {            
            String timeStamp = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");
            AppendToFile(DIMF, "<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            AppendToFile(DIMF, "<ODM xmlns=\"http://www.cdisc.org/ns/odm/v1.3\"");
            AppendToFile(DIMF, "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"");
            AppendToFile(DIMF, "xsi:schemaLocation=\"http://www.cdisc.org/ns/odm/v1.3 ODM1-3.xsd\"");
            AppendToFile(DIMF, "ODMVersion=\"1.3\" FileOID=\"1D20080412202420\" FileType=\"Snapshot\"");
            AppendToFile(DIMF, "Description=\"Dataset ODM\" CreationDateTime=\"" + timeStamp + "\" >");
            AppendToFile(DIMF, "<ClinicalData StudyOID=\"" + TheStudyOID + "\" MetaDataVersionOID=\"v1.0.0\">");
        }

        private void buttonStartConversion_Click_1(object sender, EventArgs e)
        {
            SUBJECTSEX_M = textBoxSubjectSexM.Text;
            SUBJECTSEX_F = textBoxSubjectSexF.Text;
            if (IsNumber(textBoxMaxLines.Text) == false)
            {
                MessageBox.Show("Split factor must be a number, 0 means no splitting.", "OCDataImporter");
                return;
            }
            OUTFMAXLINES = System.Convert.ToInt32(textBoxMaxLines.Text);
            BuildRepeatingGroupString();
            BuildRepeatingEventString();
            if (OUTFMAXLINES == 0) OUTFMAXLINES = 99999;
            OUTFBASIS = workdir + "\\DataImport_";
            OUTF = OUTFBASIS + OUTFFILECOUNTER + ".xml";
            // delete DataImport*.xml 
            string[] txtList = Directory.GetFiles(workdir, "DataImport_*.xml");
            if (txtList.Length > 0)
            {
                if (MessageBox.Show("DataImport_* files will be overwritten. Do you want to delete the old files?", "Confirm delete old DataImport*.xml files", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
                {
                    foreach (string f in txtList) File.Delete(f);
                }
            }

            dumpTheGrid();
            int maxSE = 0;
            int maxCRF = 0;
            int maxGR = 0;
            SortableDG.Clear();
            fpoDEL = new FileStream(workdir + "\\Deletes.sql", FileMode.Create, FileAccess.Write); // 1.0f (was deletes.xml - typo error)
            StreamWriter swDELw = new StreamWriter(fpoDEL);
            fpoDIM = new FileStream(OUTF, FileMode.Create, FileAccess.Write);
            StreamWriter swDIMw = new StreamWriter(fpoDIM);
            fpoINS = new FileStream(workdir + "\\Inserts.sql", FileMode.Create, FileAccess.Write);
            StreamWriter swINSw = new StreamWriter(fpoINS);
            fpoLOG = new FileStream(workdir + "\\OCDataImporter_warnings.txt", FileMode.Create, FileAccess.Write);
            StreamWriter swLOG = new StreamWriter(fpoLOG);
            fpoINSR = new FileStream(workdir + "\\Inserts_ONLY_STUDY_EVENTS.sql", FileMode.Create, FileAccess.Write);
            StreamWriter swINSwR = new StreamWriter(fpoINSR);
            fpoDELR = new FileStream(workdir + "\\Deletes_ONLY_STUDY_EVENTS.sql", FileMode.Create, FileAccess.Write);
            StreamWriter swDELwR = new StreamWriter(fpoDELR);
            fpoDIM.Close();
            fpoINS.Close();
            fpoINSR.Close();
            fpoDEL.Close();
            fpoDELR.Close();
            fpoLOG.Close();
            INSF = workdir + "\\Inserts.sql";
            LOG = workdir + "\\OCDataImporter_warnings.txt"; 
            INSFR = workdir + "\\Inserts_ONLY_STUDY_EVENTS.sql";
            DIMF = OUTF;
            DELF = workdir + "\\Deletes.sql";
            DELFR = workdir + "\\Deletes_ONLY_STUDY_EVENTS.sql";
            masterHeader();
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].IsNewRow == false)
                {
                    string[] mi = dataGridView1.Rows[i].Cells[DGIndexOfOCItem].Value.ToString().Split('.');
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
            STDItem = "";
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].IsNewRow == false)
                {
                    if (dataGridView1.Rows[i].Cells[DGIndexOfSex].Value.ToString() == "True")
                    {
                        sexCount++;
                        for (int j = 0; j < DataFileItems.Count; j++)
                        {
                            if (DataFileItems[j].ToString() == dataGridView1.Rows[i].Cells[DGIndexOfDataItem].Value.ToString())
                            {
                                sexIndex = j;
                                sexItem = dataGridView1.Rows[i].Cells[DGIndexOfDataItem].Value.ToString();
                            }
                        }
                    }
                    if (dataGridView1.Rows[i].Cells[DGIndexOfPID].Value.ToString() == "True")
                    {
                        PIDCount++;
                        for (int j = 0; j < DataFileItems.Count; j++)
                        {
                            if (DataFileItems[j].ToString() == dataGridView1.Rows[i].Cells[DGIndexOfDataItem].Value.ToString())
                            {
                                PIDIndex = j;
                                PIDItem = dataGridView1.Rows[i].Cells[DGIndexOfDataItem].Value.ToString();
                            }
                        }
                    }
                    if (dataGridView1.Rows[i].Cells[DGIndexOfDOB].Value.ToString() == "True")
                    {
                        DOBCount++;
                        for (int j = 0; j < DataFileItems.Count; j++)
                        {
                            if (DataFileItems[j].ToString() == dataGridView1.Rows[i].Cells[DGIndexOfDataItem].Value.ToString())
                            {
                                DOBIndex = j;
                                DOBItem = dataGridView1.Rows[i].Cells[DGIndexOfDataItem].Value.ToString();
                            }
                        }
                    }
                    if (dataGridView1.Rows[i].Cells[DGIndexOfSTD].Value.ToString() == "True")
                    {
                        // 2.0.7 put the STDIndex and STDItems in an ARRAY as it is possible that in a repeating event more START dates are given!
                        STDCount++;
                        for (int j = 0; j < DataFileItems.Count; j++)
                        {
                            if (DataFileItems[j].ToString() == dataGridView1.Rows[i].Cells[DGIndexOfDataItem].Value.ToString())
                            {
                                STDIndex = j;
                                STDItem = dataGridView1.Rows[i].Cells[DGIndexOfDataItem].Value.ToString();
                                EventStartDates += STDItem + "^" + STDIndex.ToString() + "$";
                            }
                        }
                    }
                                        
                    string[] mi = dataGridView1.Rows[i].Cells[DGIndexOfOCItem].Value.ToString().Split('.');
                    if (mi.Length == 1 && (mi[0] == "none" || mi[0].StartsWith("Use link button")))
                    {
                        if (dataGridView1.Rows[i].Cells[DGIndexOfKey].Value.ToString() == "True")
                        {
                            SortableDG.Add("none" + "^" + dataGridView1.Rows[i].Cells[DGIndexOfDataItem].Value.ToString() + "^" + dataGridView1.Rows[i].Cells[DGIndexOfKey].Value.ToString() + "^" + dataGridView1.Rows[i].Cells[DGIndexOfDate].Value.ToString() + "^" + dataGridView1.Rows[i].Cells[DGIndexOfSex].Value.ToString() + "^" + dataGridView1.Rows[i].Cells[DGIndexOfPID].Value.ToString() + "^" + dataGridView1.Rows[i].Cells[DGIndexOfDOB].Value.ToString() + "^" + dataGridView1.Rows[i].Cells[DGIndexOfSTD].Value.ToString());
                        }
                        else continue;
                    }
                    else
                    {
                        string theDataItem = dataGridView1.Rows[i].Cells[DGIndexOfDataItem].Value.ToString();
                        string ev_rep = "";
                        string gr_rep = "";
                        if (repeating_events.Contains(mi[0])) ev_rep = Get_SE_RepeatingKey_FromStudyDataColumn(theDataItem);
                        if (repeating_groups.Contains(mi[2])) gr_rep = Get_GR_RepeatingKey_FromStudyDataColumn(theDataItem);
                        mi[0] = FillTildes(mi[0], maxSE);
                        mi[1] = FillTildes(mi[1], maxCRF);
                        mi[2] = FillTildes(mi[2], maxGR);
                        if (ev_rep != "") mi[0] = mi[0] + "*" + ev_rep;
                        if (gr_rep != "") mi[2] = mi[2] + "*" + gr_rep;
                        else mi[2] = "A" + mi[2];  // This is due to an OC bug; UNGROUPED items must come before the grouped items in the grid.
                        SortableDG.Add(mi[0] + "." + mi[1] + "." + mi[2] + "." + mi[3] + "^" + theDataItem + "^" + dataGridView1.Rows[i].Cells[DGIndexOfKey].Value.ToString() + "^" + dataGridView1.Rows[i].Cells[DGIndexOfDate].Value.ToString() + "^" + dataGridView1.Rows[i].Cells[DGIndexOfSex].Value.ToString() + "^" + dataGridView1.Rows[i].Cells[DGIndexOfPID].Value.ToString() + "^" + dataGridView1.Rows[i].Cells[DGIndexOfDOB].Value.ToString() + "^" + dataGridView1.Rows[i].Cells[DGIndexOfSTD].Value.ToString());
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
                    for (int j = 0; j < DataFileItems.Count; j++) if (DataFileItems[j].ToString() == tr[1]) keyIndex = j;
                }

            }

            if (keyCount != 1)
            {
                MessageBox.Show("Please select (only) one field as STUDY SUBJECT ID by using check box; You have " + keyCount.ToString() + " selected.", "OCDataImporter");
                return;
            }
            if (sexCount > 1)
            {
                MessageBox.Show("Please select at most one field as STUDY SUBJECT SEX by using check box; You have " + sexCount.ToString() + " selected.", "OCDataImporter");
                return;
            }
            if (PIDCount > 1)
            {
                MessageBox.Show("Please select at most one field as PERSON ID by using check box; You have " + PIDCount.ToString() + " selected.", "OCDataImporter");
                return;
            }
            if (DOBCount > 1)
            {
                MessageBox.Show("Please select at most one field as SUBJECT DATE OF BIRTH by using check box; You have " + DOBCount.ToString() + " selected.", "OCDataImporter");
                return;
            }
            SortableDG.Sort();
            dataGridView1.Rows.Clear();
            for (int i = 0; i < SortableDG.Count; i++)
            {
                string[] fnparts;
                fnparts = new string[dataGridView1.ColumnCount];
                string[] tr = SortableDG[i].ToString().Split('^');
                fnparts[DGIndexOfDataItem] = tr[1];
                fnparts[DGIndexOfKey] = tr[2];
                fnparts[DGIndexOfDate] = tr[3];
                fnparts[DGIndexOfSex] = tr[4];
                fnparts[DGIndexOfPID] = tr[5];
                fnparts[DGIndexOfDOB] = tr[6];
                fnparts[DGIndexOfSTD] = tr[7];
                fnparts[DGIndexOfOCItem] = tr[0].ToString().Replace("~", "");
                dataGridView1.Rows.Add(fnparts);
            }

            string Rline = textBoxReplace.Text;
            StringReader sr = new StringReader(Rline);
            ReplacePairStrings.Clear();
            while ((Rline = sr.ReadLine()) != null)
            {
                if (Rline == "" || Rline.StartsWith("//")) continue;
                if (Rline.StartsWith("SEPARATOR="))
                {
                    RepSeparator = Rline[10];
                    continue;
                }
                ReplacePairStrings.Add(Rline);
            }

            if (input_data.Contains("\\") == false) input_data = workdir + "\\" + input_data;
            if (Delimiter != tab) textBoxOutput.Text += "\r\nData file is: " + input_data + ", delimited by: " + Delimiter + " Number of items per line: " + sepcount + "\r\n";
            else textBoxOutput.Text += "\r\nData file is: " + input_data + ", delimited by: tab, Number of items per line: " + sepcount + "\r\n";
            textBoxOutput.Text += "Started in directory " + workdir + ". This may take several minutes...\r\n";

            buttonStartConversion.Enabled = false;
            buttonExit.Enabled = false;
            buttonCancel.Enabled = true;
            buttonCancel.BackColor = System.Drawing.Color.LightGreen;
            progressBar1.Value = 0;
            this.Cursor = Cursors.AppStarting;

            if (DEBUGMODE) DoWork();
            else
            {
                MyThread = new Thread(new ThreadStart(DoWork));
                MyThread.IsBackground = true;
                MyThread.Start();
            }
        }

        public string getTodaysDate()
        {
            DateTime dt = DateTime.Now;
            return dt.ToString("yyyy-MM-dd");
        }

        public void DoWork()
        {
            string theDate = getTodaysDate();
            string theWrittenSE = "";
            string theWrittenGR = "";
            string [] theHeaders;
            int event_index_row = -1;
            int group_index_row = -1;
            InsertKeys.Clear();
            try
            {
                using (StreamReader sr = new StreamReader(input_data))
                {
                    String line;
                    int linecount = 0;
                    while ((line = sr.ReadLine()) != null)
                    {
                        // line = line.Trim();  // 1.1b
                        if (line.Length == 0) continue;
                        AllValuesInOneRow.Clear();
                        Hiddens.Clear();
                        linecount++;
                        if (linecount == 1)
                        {
                            theHeaders = line.Split(Delimiter); // get row event and group indexes, if defined that way.
                            for (int i = 0; i < theHeaders.Length; i++)
                            {
                                if (theHeaders[i].ToUpper() == "EVENT_INDEX") event_index_row = i;    // 2.2 Input file format allows EVENT_INDEX and GROUP_INDEX to define repeating items in rows
                                if (theHeaders[i].ToUpper() == "GROUP_INDEX") group_index_row = i;
                            }
                            continue; // skip first line; contains only headers
                        }
                        int mySepCount = 1;
                        progressBar1.Step = line.Length;
                        for (int i = 0; i < line.Length; i++) if (line[i] == Delimiter) mySepCount++;
                        if (sepcount != mySepCount)
                        {
                            string errtext = "Input data file format incorrect at line = " + linecount.ToString() + " Expecting: " + sepcount.ToString() + "; found: " + mySepCount.ToString() + "  items; this is the faulty line: " + line; 
                            append_warning(errtext);
                            continue;
                        }
                        string[] split = line.Split(Delimiter);

                        if (checkBoxDup.Checked) // duplicate key check
                        {
                            foreach (string one in InsertKeys)
                            {
                                if (one == split[keyIndex])
                                {
                                    string errtext = "Duplicate key " + split[keyIndex] + " at line = " + linecount.ToString();
                                    append_warning(errtext);
                                }
                            }
                        }
                        InsertKeys.Add(split[keyIndex]);
                        progressBar1.PerformStep();

                        //  Handle first DG line
                        int DGFirstLine = 0;
                        string[] ocparts = dataGridView1.Rows[DGFirstLine].Cells[DGIndexOfOCItem].Value.ToString().Split('.');
                        TheStudyEventOID = ocparts[0];
                        string theSERK = "1";
                        string theGRRK = "NOT";
                        string theStudyDataColumn = "";
                        if (TheStudyEventOID == "none")
                        {
                            DGFirstLine++;
                            try // fix 2.0.3 -> If no items are matched; can't increase DGFirstLine, causes index out of range 
                            {
                                ocparts = dataGridView1.Rows[DGFirstLine].Cells[DGIndexOfOCItem].Value.ToString().Split('.');
                                TheFormOID = ocparts[1];
                                TheItemGroupDef = ocparts[2];
                                if (TheItemGroupDef.IndexOf('-') > 0) TheItemGroupDef = TheItemGroupDef.Substring(0, TheItemGroupDef.IndexOf('-'));
                                TheItemId = ocparts[3];
                                TheStudyEventOID = ocparts[0];
                                theStudyDataColumn = dataGridView1.Rows[DGFirstLine].Cells[DGIndexOfDataItem].Value.ToString();
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
                            theStudyDataColumn = dataGridView1.Rows[DGFirstLine].Cells[DGIndexOfDataItem].Value.ToString();
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
                            if (STDItem.Contains("_E"))
                            {
                                theSERK = STDItem.Substring(STDItem.IndexOf("_E") + 2);
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
                        string theKEY = split[keyIndex];
                        if (event_index_row != -1) theSERK = split[event_index_row];
                        if (group_index_row != -1) theGRRK = split[group_index_row];
                        string SStheKEY = "SS_" + theKEY;
                        if (labelOCoidExists) // 2.1.1 there is a conversion file from label to oid; get the SSid from that file.
                        {
                            foreach (string one in LabelOID)
                            {
                                if (one.StartsWith(theKEY + "^")) SStheKEY = one.Substring(one.IndexOf('^') + 1); 
                            }
                        }
                        string theXMLForm = "";
                        AppendToFile(DIMF, "    <SubjectData SubjectKey=\"" + SStheKEY + "\">");
                        theXMLEvent += "        <StudyEventData StudyEventOID=\"" + theWrittenSE + "\" StudyEventRepeatKey=\"" + CheckRK(theSERK, linecount) + "\">" + Mynewline;
                        theXMLForm += "            <FormData FormOID=\"" + TheFormOID + "\">" + Mynewline;
                        if (theGRRK == "NOT") theXMLForm += "                <ItemGroupData ItemGroupOID=\"" + theWrittenGR + "\" TransactionType=\"Insert\" >" + Mynewline;
                        else theXMLForm += "                <ItemGroupData ItemGroupOID=\"" + theWrittenGR + "\" ItemGroupRepeatKey=\"" + CheckRK(theGRRK, linecount) + "\" TransactionType=\"Insert\" >" + Mynewline;
                        int indexOfItem = 0;
                        string itemval = "";
                        bool datapresentform = false;
                        bool datapresentevent = false;
                        if (TheItemId != "none")
                        {
                            indexOfItem = GetIndexOfItem(dataGridView1.Rows[DGFirstLine].Cells[DGIndexOfOCItem].Value.ToString());
                            check_index_of_item(linecount, indexOfItem);
                            itemval = split[indexOfItem];
                            itemval = Repl(dataGridView1.Rows[0].Cells[DGIndexOfDataItem].Value.ToString(), itemval);
                            itemval = ValidateItem(theKEY, TheFormOID, TheItemId, itemval, linecount);
                            if (itemval != "")
                            {
                                theXMLForm += "                    <ItemData ItemOID=\"" + TheItemId + "\" Value=\"" + Escape(itemval) + "\" />" + Mynewline;
                                datapresentform = true;
                            }
                        }
                        // Now handle the rest of the DG
                        for (int i = DGFirstLine + 1; i < dataGridView1.RowCount; i++)
                        {
                            if (dataGridView1.Rows[i].IsNewRow == false)
                            {
                                string[] nwdingen = dataGridView1.Rows[i].Cells[DGIndexOfOCItem].Value.ToString().Split('.');
                                theStudyDataColumn = dataGridView1.Rows[i].Cells[DGIndexOfDataItem].Value.ToString();
                                if (TheStudyEventOID == nwdingen[0] && TheFormOID == nwdingen[1] && TheItemGroupDef == nwdingen[2])
                                {
                                    if (nwdingen[3] != "none")
                                    {
                                        indexOfItem = GetIndexOfItem(dataGridView1.Rows[i].Cells[DGIndexOfOCItem].Value.ToString());
                                        check_index_of_item(linecount, indexOfItem);
                                        itemval = split[indexOfItem];
                                        itemval = Repl(dataGridView1.Rows[i].Cells[DGIndexOfDataItem].Value.ToString(), itemval);
                                        itemval = ValidateItem(theKEY, TheFormOID, nwdingen[3], itemval, linecount);
                                        if (itemval != "")
                                        {
                                            theXMLForm += "                    <ItemData ItemOID=\"" + nwdingen[3] + "\" Value=\"" + Escape(itemval) + "\" />" + Mynewline;
                                            datapresentform = true;
                                        }
                                    }
                                }
                                else if (TheStudyEventOID == nwdingen[0] && TheFormOID == nwdingen[1]) 
                                {
                                    TheItemGroupDef = nwdingen[2];
                                    theXMLForm += "                </ItemGroupData>" + Mynewline;
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
                                    if (theGRRK == "NOT") theXMLForm += "                <ItemGroupData ItemGroupOID=\"" + theWrittenGR + "\" TransactionType=\"Insert\" >" + Mynewline;
                                    else theXMLForm += "                <ItemGroupData ItemGroupOID=\"" + theWrittenGR + "\" ItemGroupRepeatKey=\"" + CheckRK(theGRRK, linecount) + "\" TransactionType=\"Insert\" >" + Mynewline;
                                    if (nwdingen[3] != "none")
                                    {
                                        indexOfItem = GetIndexOfItem(dataGridView1.Rows[i].Cells[DGIndexOfOCItem].Value.ToString());
                                        check_index_of_item(linecount, indexOfItem);
                                        itemval = split[indexOfItem];
                                        itemval = Repl(dataGridView1.Rows[i].Cells[DGIndexOfDataItem].Value.ToString(), itemval);
                                        itemval = ValidateItem(theKEY, TheFormOID, nwdingen[3], itemval, linecount);
                                        if (itemval != "")
                                        {
                                            theXMLForm += "                    <ItemData ItemOID=\"" + nwdingen[3] + "\" Value=\"" + Escape(itemval) + "\" />" + Mynewline;
                                            datapresentform = true;
                                        }
                                    }
                                }
                                else if (TheStudyEventOID == nwdingen[0]) 
                                {
                                    TheFormOID = nwdingen[1];
                                    TheItemGroupDef = nwdingen[2];
                                    theXMLForm += "                </ItemGroupData>" + Mynewline;
                                    theXMLForm += "            </FormData>" + Mynewline;
                                    if (datapresentform)
                                    {
                                        datapresentevent = true;
                                        theXMLEvent += theXMLForm;
                                        
                                        datapresentform = false;
                                    }
                                    theXMLForm = "";
                                    theXMLForm += "            <FormData FormOID=\"" + TheFormOID + "\">" + Mynewline;
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
                                    if (theGRRK == "NOT") theXMLForm += "                <ItemGroupData ItemGroupOID=\"" + theWrittenGR + "\" TransactionType=\"Insert\" >" + Mynewline;
                                    else theXMLForm += "                <ItemGroupData ItemGroupOID=\"" + theWrittenGR + "\" ItemGroupRepeatKey=\"" + CheckRK(theGRRK, linecount) + "\" TransactionType=\"Insert\" >" + Mynewline;
                                    if (nwdingen[3] != "none")
                                    {
                                        indexOfItem = GetIndexOfItem(dataGridView1.Rows[i].Cells[DGIndexOfOCItem].Value.ToString());
                                        check_index_of_item(linecount, indexOfItem);
                                        itemval = split[indexOfItem];
                                        itemval = Repl(dataGridView1.Rows[i].Cells[DGIndexOfDataItem].Value.ToString(), itemval);
                                        itemval = ValidateItem(theKEY, TheFormOID, nwdingen[3], itemval, linecount);
                                        if (itemval != "")
                                        {
                                            theXMLForm += "                    <ItemData ItemOID=\"" + nwdingen[3] + "\" Value=\"" + Escape(itemval) + "\" />" + Mynewline;
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
                                    theXMLForm += "                </ItemGroupData>" + Mynewline;
                                    theXMLForm += "            </FormData>" + Mynewline;
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
                                        AppendToFile(DIMF, theXMLEvent);

                                        datapresentevent = false;
                                    }
                                    theXMLEvent = "";
                                    if (selectedEventRepeating == "Yes")
                                    {
                                        if (event_index_row != -1) theSERK = split[event_index_row];
                                        else theSERK = Get_SE_RepeatingKey_FromStudyDataColumn(theStudyDataColumn);
                                    }
                                    theXMLEvent += "        <StudyEventData StudyEventOID=\"" + theWrittenSE + "\" StudyEventRepeatKey=\"" + CheckRK(theSERK, linecount) + "\">" + Mynewline;
                                    TheFormOID = nwdingen[1];
                                    theXMLForm += "            <FormData FormOID=\"" + TheFormOID + "\">" + Mynewline;
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
                                    if (theGRRK == "NOT") theXMLForm += "                <ItemGroupData ItemGroupOID=\"" + theWrittenGR + "\" TransactionType=\"Insert\" >" + Mynewline;
                                    else theXMLForm += "                <ItemGroupData ItemGroupOID=\"" + theWrittenGR + "\" ItemGroupRepeatKey=\"" + CheckRK(theGRRK, linecount) + "\" TransactionType=\"Insert\" >" + Mynewline;
                                    if (nwdingen[3] != "none")
                                    {
                                        indexOfItem = GetIndexOfItem(dataGridView1.Rows[i].Cells[DGIndexOfOCItem].Value.ToString());
                                        check_index_of_item(linecount, indexOfItem);
                                        itemval = split[indexOfItem];
                                        itemval = Repl(dataGridView1.Rows[i].Cells[DGIndexOfDataItem].Value.ToString(), itemval);
                                        itemval = ValidateItem(theKEY, TheFormOID, nwdingen[3], itemval, linecount);
                                        if (itemval != "")
                                        {
                                            theXMLForm += "                    <ItemData ItemOID=\"" + nwdingen[3] + "\" Value=\"" + Escape(itemval) + "\" />" + Mynewline;
                                            datapresentform = true;
                                        }
                                    }
                                }
                            }
                        }
                        string SubSex = comboBoxSex.SelectedItem.ToString();
                        if (sexIndex >= 0)
                        {
                            SubSex = split[sexIndex];
                            SubSex = Repl(sexItem, SubSex);
                            SubSex = SubSex.Trim();
                        }
                        if (SubSex == SUBJECTSEX_M) SubSex = "m";  // 1.0f
                        if (SubSex == SUBJECTSEX_F) SubSex = "f";
                        if (SubSex != "f" && SubSex != "m" && SubSex != "") // 1.1e  Allow nothing to be entered for sex; it is not always mandatory.
                        {
                            string errtext = "Subject sex can be only 'f' or 'm'. You have '" + SubSex + "' at line " + linecount.ToString() + ". Index: " + sexIndex + ".";
                            append_warning(errtext);
                        }
                        theXMLForm += "                </ItemGroupData>" + Mynewline;
                        theXMLForm += "            </FormData>" + Mynewline;
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
                            AppendToFile(DIMF, theXMLEvent);

                            datapresentevent = false;
                        }
                        theXMLEvent = "";
                        AppendToFile(DIMF, "    </SubjectData>");

                        OUTFLINECOUNTER = OUTFLINECOUNTER + 1;
                        if (OUTFLINECOUNTER >= OUTFMAXLINES)  // 1.0f
                        {
                            masterClose();
                            fpoDIM.Close();
                            OUTFLINECOUNTER = 0;
                            OUTFFILECOUNTER = OUTFFILECOUNTER + 1;
                            OUTF = OUTFBASIS + OUTFFILECOUNTER + ".xml";
                            fpoDIM = new FileStream(OUTF, FileMode.Create, FileAccess.Write);
                            fpoDIM.Close();
                            DIMF = OUTF;
                            masterHeader();
                        }
                        // generate insert statements
                        int theSERKInt = System.Convert.ToInt16(theSERK);
                        string theDOB = "";
                        if (DOBIndex >= 0) theDOB = DateUtilities.ConvertToODMFormat(split[DOBIndex], comboBoxDateFormat.SelectedItem.ToString());
                        string theSTD = "";
                        if (STDIndex >= 0) theSTD = DateUtilities.ConvertToODMFormat(split[STDIndex], comboBoxDateFormat.SelectedItem.ToString()); // This is needed for non repeating events
                        string thePID = "";
                        if (PIDIndex < 0) thePID = theKEY;
                        else thePID = split[PIDIndex];
                        
                        if (theDOB.StartsWith("Error") || theDOB == "" || DOBIndex < 0)
                        {
                            if (theDOB == "" || DOBIndex < 0)
                            {
                                if (!IsDuplicatePID(thePID))
                                {
                                    AppendToFile(INSF, insert1);
                                    AppendToFile(INSF, "    VALUES (1, '" + SubSex + "', '" + thePID + "', '" + theDate + "', 1, '1');");
                                }
                            }
                            else
                            {
                                string errtext = "Invalid subject birth date '" + theDOB + "' at line " + linecount.ToString() + ". Index: " + DOBIndex + ". ";
                                append_warning(errtext);
                            }
                        }
                        else
                        {
                            if (!IsDuplicatePID(thePID))
                            {
                                AppendToFile(INSF, insert1a);
                                AppendToFile(INSF, "    VALUES (1, '" + SubSex + "', '" + thePID + "', '" + theDate + "', 1, '1', '" + theDOB + "');");
                            }
                        }
                        AppendToFile(INSF, insert2);
                        // if there is no PID, use the key (unique_identifier) to fill the field label of study_subject. 
                        AppendToFile(INSF, "    VALUES ('" + theKEY + "', (SELECT study_id FROM study WHERE oc_oid = '" + TheStudyOID + "'),");
                        AppendToFile(INSF, "            1, '" + theDate + "', '" + theDate + "', '" + theDate + "', 1, '" + SStheKEY + "', (SELECT subject_id FROM subject where unique_identifier = '" + thePID + "'));");
                        if (theWrittenSE == "none" && comboBoxSE.SelectedItem.ToString() != "-- select --") // 4.1 eliminate -- select --
                        {
                            theWrittenSE = comboBoxSE.SelectedItem.ToString(); // 2.0.5 Use the selected SE to determine current SE, as there is no CRF data 
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
                                if (EventStartDates.Contains(part2)) // there is a date; get it
                                {
                                    part1 = EventStartDates.Substring(EventStartDates.IndexOf(part2) + 1);
                                    part1 = part1.Substring(part1.IndexOf("^") + 1);
                                    part1 = part1.Substring(0, part1.IndexOf("$"));  // part1 is the index of date
                                    theSTD = DateUtilities.ConvertToODMFormat(split[System.Convert.ToInt16(part1)], comboBoxDateFormat.SelectedItem.ToString());
                                    if (theSTD.StartsWith("Error"))
                                    {
                                        string errtext = "Invalid start date '" + theSTD + "' at line " + linecount.ToString() + ". Index: " + STDIndex + ". ";
                                        append_warning(errtext);
                                    }
                                    else
                                    {
                                        if (theSTD != "")
                                        {
                                            AppendToFile(INSF, insert3);
                                            AppendToFile(INSF, "    VALUES ((SELECT study_event_definition_id FROM study_event_definition WHERE oc_oid = '" + theWrittenSE + "'),");
                                            AppendToFile(INSFR, insert3);
                                            AppendToFile(INSFR, "    VALUES ((SELECT study_event_definition_id FROM study_event_definition WHERE oc_oid = '" + theWrittenSE + "'),");
                                            AppendToFile(INSF, "	    (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "'),'" + textBoxLocation.Text + "', " + say.ToString() + ", '" + theSTD + " 12:00:00', 1, 1, '" + theDate + "', 3, '0', '0');");
                                            AppendToFile(INSFR, "	     (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "'),'" + textBoxLocation.Text + "', " + say.ToString() + ", '" + theSTD + " 12:00:00', 1, 1, '" + theDate + "', 3, '0', '0');");
                                        }
                                    }
                                }
                                else  // no repeating Events; use either todays date as STD or pick it from data file  2.0.9
                                {
                                    if (theSTD != "") // there is a date specified in the data file
                                    {
                                        AppendToFile(INSF, insert3);
                                        AppendToFile(INSF, "    VALUES ((SELECT study_event_definition_id FROM study_event_definition WHERE oc_oid = '" + theWrittenSE + "'),");
                                        AppendToFile(INSFR, insert3);
                                        AppendToFile(INSFR, "    VALUES ((SELECT study_event_definition_id FROM study_event_definition WHERE oc_oid = '" + theWrittenSE + "'),");
                                        AppendToFile(INSF, "	    (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "'),'" + textBoxLocation.Text + "', " + say.ToString() + ", '" + theSTD + " 12:00:00', 1, 1, '" + theDate + "', 3, '0', '0');");
                                        AppendToFile(INSFR, "        (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "'),'" + textBoxLocation.Text + "', " + say.ToString() + ", '" + theSTD + " 12:00:00', 1, 1, '" + theDate + "', 3, '0', '0');");
                                    }
                                    else  // no date specified in data file
                                    {
                                        if (radioButtonUseTD.Checked)  // 2.1.3 Generate inserts for events without dates using todays date, only if user wants to, otherwise 
                                        {
                                            AppendToFile(INSF, insert3);
                                            AppendToFile(INSF, "    VALUES ((SELECT study_event_definition_id FROM study_event_definition WHERE oc_oid = '" + theWrittenSE + "'),");
                                            AppendToFile(INSFR, insert3);
                                            AppendToFile(INSFR, "    VALUES ((SELECT study_event_definition_id FROM study_event_definition WHERE oc_oid = '" + theWrittenSE + "'),");
                                            AppendToFile(INSF, "	    (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "'),'" + textBoxLocation.Text + "', " + say.ToString() + ", '" + theDate + " 12:00:00', 1, 1, '" + theDate + "', 3, '0', '0');");
                                            AppendToFile(INSFR, "        (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "'),'" + textBoxLocation.Text + "', " + say.ToString() + ", '" + theDate + " 12:00:00', 1, 1, '" + theDate + "', 3, '0', '0');");
                                        }
                                    }
                                }
                            }
                            AppendToFile(DELFR, "DELETE FROM study_event where study_subject_id = (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "');");
                        }
                        AppendToFile(DELF, "DELETE FROM study_event where study_subject_id = (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "');");
                        AppendToFile(DELF, "DELETE FROM study_subject where oc_oid = '" + SStheKEY + "';");
                        AppendToFile(DELF, "DELETE FROM subject where unique_identifier = '" + thePID + "';");
                        // Control hidden values and mandatory values
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
                                    string ItemToControl = GetItemOIDFromItemName(contItem.ToLower(), TheFormOID);
                                    foreach (string alls in AllValuesInOneRow)
                                    {
                                        string[] all = alls.Split('^');
                                        if (all[0] == ItemToControl && all[1] == contVal) append_warning("Line " + linecount.ToString() + ", Subject= " + SStheKEY + ", CRF= " + TheFormOID + ", ItemOID= " + hd + ", Item Value is null while item is mandatory");
                                    }
                                }
                            }
                        }
                    }
                    masterClose();
                }
            }
            catch (Exception ex)
            {
                string errtext = "Exception while reading data file: " + ex;
                if (errtext.Contains("ThreadAbortException") == false)
                {
                    MessageBox.Show(errtext, "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    exit_error(ex.ToString());
                }
            }
            // End work
            buttonStartConversion.Enabled = false;
            buttonStartConversion.BackColor = SystemColors.Control;
            button_start.Enabled = false;
            buttonExit.Enabled = true;
            buttonBrowse.Enabled = false;
            buttonCancel.Enabled = false;
            buttonCancel.BackColor = SystemColors.Control;
            linkLabelBuildDG.Enabled = false;
            linkbuttonSHCols.Enabled = false;
            linkLabel1.Enabled = false;
            textBoxInput.Focus();
            this.Cursor = Cursors.Arrow;
            progressBar1.Value = PROGBARSIZE;
            if (WARCOUNT == 0)
            {
                WARCOUNT = -1;
                append_warning(DateTime.Now + " Finished successfully.");
                MessageBox.Show("Process finished successfully", "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Process ended with errors or warnings: See OCDataImporter_log.txt and/or OCDataImporter_warning.txt for details", "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBoxOutput.Text += "*** There are errors or warnings: See OCDataImporter_log.txt and/or OCDataImporter_warning.txt for details ***";
            }
            if (OUTFFILECOUNTER > 1) textBoxOutput.Text += " Total: " + OUTFFILECOUNTER.ToString() + " ODM files."; 
            textBoxOutput.SelectionStart = textBoxOutput.Text.Length;
            textBoxOutput.ScrollToCaret();
        }
        public bool IsDuplicatePID(string thePIDtoCheck)
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
        public string Get_SE_RepeatingKey_FromStudyDataColumn(string theName)
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
                MessageBox.Show(errtext + "    " + ex.Message, "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                exit_error(errtext + "    " + ex.Message);
            }
            // MessageBox.Show(errtext, "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            // exit_error(errtext);  Version 2.0.1 -> Assume 1 if the repeating key cant be determined. (and dont exit) The error message tells the user about the caveats.
            return ("1");
        }

        public string Get_GR_RepeatingKey_FromStudyDataColumn(string theName)
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
                MessageBox.Show(errtext + "    " + ex.Message, "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                exit_error(errtext + "    " + ex.Message);
            }
            //MessageBox.Show(errtext, "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            //exit_error(errtext);
            return ("N");  // this means: Cant determine but maybe going to get from row: the check will be done in DoWork
        }

        private void masterClose()
        {
            AppendToFile(DIMF, "</ClinicalData>");
            AppendToFile(DIMF, "</ODM>");
        }

        public void check_index_of_item(int linecount, int myioi)
        {
            if (myioi < 0)
            {
                string errtext = " Wrong index at: " + linecount.ToString() + ". Exiting...The generated files ARE INCOMPLETE AND CAN NOT BE USED";
                MessageBox.Show(errtext, "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                exit_error(errtext);
            }
        }

        public string Repl(string it, string val)
        {
            foreach (string one in ReplacePairStrings)
            {
                string[] parts = one.Split(RepSeparator);
                if ((parts.Length == 2 || parts.Length == 3) == false)
                {
                    string errtext = " Wrong replace couple: Expecting COLUMN;old;new or COLUMN;+text or COLUMN;text+ but found: " + one;
                    append_warning(errtext);
                    return("WRONGVALUE");
                }
                if (parts.Length == 2)
                {
                    if (parts[1].Contains("+") && (parts[0] == it || parts[0] == "ALL"))
                    {
                        if (parts[1].Length > 1)
                        {
                            if (parts[1].EndsWith("+")) val = parts[1].Substring(0, parts[1].Length - 1) + val;
                            if (parts[1].StartsWith("+")) val = val + parts[1].Substring(1);
                        }
                    }
                }
                if (parts.Length == 3)
                {
                    if (parts[2] == "<null>") parts[2] = "";
                    if (parts[0] == it || parts[0] == "ALL") val = val.Replace(parts[1], parts[2]);
                }                
            }
            return (val);
        }

        public int GetIndexOfItem(string item)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].IsNewRow == false)
                {
                    if (dataGridView1.Rows[i].Cells[DGIndexOfOCItem].Value.ToString() == "none") continue;
                    if (dataGridView1.Rows[i].Cells[DGIndexOfOCItem].Value.ToString() == item) 
                    {
                        string myDataItem = dataGridView1.Rows[i].Cells[DGIndexOfDataItem].Value.ToString();
                        for (int j = 0; j < DataFileItems.Count; j++) if (DataFileItems[j].ToString() == myDataItem) return (j);
                    }
                }
            }
            return (-1);
        }
        public void GetStudyEventDef(string path)
        {
            try
            {
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.ConformanceLevel = ConformanceLevel.Fragment;
                settings.IgnoreWhitespace = true;
                settings.IgnoreComments = true;
                XmlReader reader = XmlReader.Create(path, settings);

                while (reader.Read())
                {
                    if (reader.Name == "StudyEventDef")
                    {
                        if (reader.AttributeCount == 4)
                        {
                            string SEOID = reader.GetAttribute(0);
                            if (DEBUGMODE) textBoxOutput.Text += reader.GetAttribute(0) + ", Name = " + reader.GetAttribute(1) + ", Repeating = " + reader.GetAttribute(2) + ", Type = " + reader.GetAttribute(3) + Mynewline;
                            comboBoxSE.Items.Insert(0, SEOID);
                        }
                    }
                }
                if (comboBoxSE.Items.Count == 1)
                {
                    comboBoxSE.SelectedIndex = 0;
                    comboBoxCRF.Focus();
                }
                else
                {
                    comboBoxSE.Items.Insert(0, "-- select --");
                    comboBoxSE.SelectedIndex = 0;
                    comboBoxSE.Focus();
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show("Can't get study event definitions; please check the format of the file", "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                exit_error(ex.ToString());
            }

        }

        private void BuildRepeatingGroupString()
        {
            repeating_groups = "";
            try
            {
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.ConformanceLevel = ConformanceLevel.Fragment;
                settings.IgnoreWhitespace = true;
                settings.IgnoreComments = true;
                XmlReader reader = XmlReader.Create(input_oc, settings);
                while (reader.Read())
                {
                    if (reader.Name == "ItemGroupDef")
                    {
                        if (reader.AttributeCount > 0)
                        {
                            string SEOID = reader.GetAttribute(0);
                            if (reader.GetAttribute(2) == "Yes") repeating_groups += SEOID + " ";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can't get Group definitions; please check the format of the file", "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                exit_error(ex.ToString());
            }
        }

        private void BuildRepeatingEventString()
        {
            repeating_events = "";
            try
            {
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.ConformanceLevel = ConformanceLevel.Fragment;
                settings.IgnoreWhitespace = true;
                settings.IgnoreComments = true;
                XmlReader reader = XmlReader.Create(input_oc, settings);
                while (reader.Read())
                {
                    if (reader.Name == "StudyEventDef")
                    {
                        if (reader.AttributeCount > 0)
                        {
                            string SEOID = reader.GetAttribute(0);
                            if (reader.GetAttribute(2) == "Yes") repeating_events += SEOID + " ";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can't get Event definitions; please check the format of the meta-file", "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                exit_error(ex.ToString());
            }
        }

        private void comboBoxGR_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBoxIT.Items.Clear();
            Items.Clear();
            string myGR = comboBoxGR.SelectedItem.ToString();
            if (myGR == "-- select --")
            {
                comboBoxIT.Items.Clear(); 
                Items.Clear();
                return;
            }
            try
            {
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.ConformanceLevel = ConformanceLevel.Fragment;
                settings.IgnoreWhitespace = true;
                settings.IgnoreComments = true;
                XmlReader reader = XmlReader.Create(input_oc, settings);

                if (DEBUGMODE) textBoxOutput.Text += "Selected Group = " + myGR;
                while (reader.Read())
                {
                    if (reader.Name == "ItemGroupDef")
                    {
                        if (reader.AttributeCount > 0)
                        {
                            string SEOID = reader.GetAttribute(0);
                            if (SEOID == myGR)
                            {
                                if (DEBUGMODE) textBoxOutput.Text += " Repeating = " + reader.GetAttribute(2) + Mynewline; 
                                string myGRPS = reader.ReadInnerXml();
                                if (myGRPS.Contains("<ItemRef "))
                                {
                                    myGRPS = myGRPS.Replace("<ItemRef ", "");
                                    myGRPS = myGRPS.Replace("/>", "~");
                                    if (DEBUGMODE) textBoxOutput.Text += "Items for the selected Group = " + Mynewline;
                                    foreach (string ss in myGRPS.Split('~'))
                                    {
                                        ss.Trim();
                                        if (ss != "" && ss.Contains("ItemOID"))
                                        {
                                            string myIT = "";
                                            if (DEBUGMODE) textBoxOutput.Text += ss + Mynewline;
                                            myIT = ss.Substring(9);
                                            int stoppunt = myIT.IndexOf('"');
                                            myIT = myIT.Substring(0, stoppunt);
                                            comboBoxIT.Items.Add(myIT);
                                            Items.Add(myIT);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                comboBoxIT.Items.Insert(0, "-- select --");
                comboBoxIT.SelectedIndex = 0;
                comboBoxIT.Focus();
                textBoxOutput.SelectionStart = textBoxOutput.Text.Length;
                textBoxOutput.ScrollToCaret();
            }

            catch (Exception ex)
            {
                MessageBox.Show("Can't get Group definitions; please check the format of the file", "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                exit_error(ex.ToString());
            }
        }
        private void comboBoxCRF_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBoxGR.Items.Clear();
            Items.Clear();
            string myCRF = comboBoxCRF.SelectedItem.ToString();
            if (myCRF == "-- select --")
            {
                comboBoxGR.Items.Clear();
                comboBoxIT.Items.Clear(); 
                Items.Clear();
                return;
            }

            try
            {
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.ConformanceLevel = ConformanceLevel.Fragment;
                settings.IgnoreWhitespace = true;
                settings.IgnoreComments = true;
                XmlReader reader = XmlReader.Create(input_oc, settings);
                if (DEBUGMODE) textBoxOutput.Text += "Selected CRF = " + myCRF + Mynewline;
                while (reader.Read())
                {
                    if (reader.Name == "FormDef")
                    {
                        if (reader.AttributeCount > 0)
                        {
                            string SEOID = reader.GetAttribute(0);
                            if (SEOID == myCRF)
                            {
                                string myGRPS = reader.ReadInnerXml();
                                if (myGRPS.Contains("<ItemGroupRef "))
                                {
                                    myGRPS = myGRPS.Replace("<ItemGroupRef ", "");
                                    myGRPS = myGRPS.Replace("/>", "~");
                                    if (DEBUGMODE) textBoxOutput.Text += "Groups for the selected CRF = " + Mynewline;
                                    foreach (string ss in myGRPS.Split('~'))
                                    {
                                        ss.Trim();
                                        if (ss != "" && ss.Contains("ItemGroupOID"))
                                        {
                                            string myGRP = "";
                                            if (DEBUGMODE) textBoxOutput.Text += ss + Mynewline;
                                            myGRP = ss.Substring(14);
                                            int stoppunt = myGRP.IndexOf('"');
                                            myGRP = myGRP.Substring(0, stoppunt);
                                            comboBoxGR.Items.Add(myGRP);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (comboBoxGR.Items.Count == 1)
                {
                    comboBoxGR.SelectedIndex = 0;
                    comboBoxIT.Focus();
                }
                else
                {
                    comboBoxGR.Items.Insert(0, "-- select --");
                    comboBoxGR.SelectedIndex = 0;
                    comboBoxGR.Focus();
                }
                textBoxOutput.SelectionStart = textBoxOutput.Text.Length;
                textBoxOutput.ScrollToCaret();
            }

            catch (Exception ex)
            {
                MessageBox.Show("Can't get CRF definitions; please check the format of the file", "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                exit_error(ex.ToString());
            }


        }
        private void comboBoxSE_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBoxCRF.Items.Clear();
            Items.Clear();
            string mySE = comboBoxSE.SelectedItem.ToString();
            if (mySE == "-- select --")
            {
                comboBoxCRF.Items.Clear();
                comboBoxGR.Items.Clear();
                comboBoxIT.Items.Clear();
                Items.Clear();
                return;
            }

            try
            {
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.ConformanceLevel = ConformanceLevel.Fragment;
                settings.IgnoreWhitespace = true;
                settings.IgnoreComments = true;
                XmlReader reader = XmlReader.Create(input_oc, settings);
                if (DEBUGMODE) textBoxOutput.Text += "Selected Study Event = " + mySE;
                while (reader.Read())
                {
                    if (reader.Name == "StudyEventDef")
                    {
                        if (reader.AttributeCount == 4)
                        {
                            string SEOID = reader.GetAttribute(0);
                            if (SEOID == mySE)
                            {
                                selectedEventRepeating = reader.GetAttribute(2);
                                if (DEBUGMODE) textBoxOutput.Text += "Repeating = " + selectedEventRepeating + Mynewline;
                                string myCRFS = reader.ReadInnerXml();
                                if (myCRFS.Contains("<FormRef "))
                                {
                                    myCRFS = myCRFS.Replace("<FormRef ", "");
                                    myCRFS = myCRFS.Replace("/>", "~");
                                    if (DEBUGMODE) textBoxOutput.Text += "CRF's for the selected SE = " + Mynewline;
                                    foreach (string ss in myCRFS.Split('~'))
                                    {
                                        ss.Trim();
                                        if (ss != "" && ss.Contains("FormOID"))
                                        {
                                            string myCRF = "";
                                            if (DEBUGMODE) textBoxOutput.Text += ss + Mynewline;
                                            myCRF = ss.Substring(9);
                                            int stoppunt = myCRF.IndexOf('"');
                                            myCRF = myCRF.Substring(0, stoppunt);
                                            comboBoxCRF.Items.Add(myCRF);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (comboBoxCRF.Items.Count == 1)
                {
                    comboBoxCRF.SelectedIndex = 0;
                    comboBoxGR.Focus();
                }
                else
                {
                    comboBoxCRF.Items.Insert(0, "-- select --");
                    comboBoxCRF.SelectedIndex = 0;
                    comboBoxCRF.Focus();
                }
                textBoxOutput.SelectionStart = textBoxOutput.Text.Length;
                textBoxOutput.ScrollToCaret();
            }

            catch (Exception ex)
            {
                MessageBox.Show("Can't get study event definitions; please check the format of the file", "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                exit_error(ex.ToString());
            }

        }
        private void comboBoxIT_SelectedIndexChanged(object sender, EventArgs e)
        {
            string myIT = comboBoxIT.SelectedItem.ToString();
            if (myIT == "-- select --") return;
            try
            {
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.ConformanceLevel = ConformanceLevel.Fragment;
                settings.IgnoreWhitespace = true;
                settings.IgnoreComments = true;
                XmlReader reader = XmlReader.Create(input_oc, settings);
                if (DEBUGMODE) textBoxOutput.Text += "Selected Item = " + myIT + Mynewline;
                while (reader.Read())
                {
                    if (reader.Name == "ItemDef")
                    {
                        if (reader.AttributeCount == 6)
                        {
                            string ITID = reader.GetAttribute(0);
                            if (ITID == myIT)
                            {
                                if (DEBUGMODE) textBoxOutput.Text += "ID = " + ITID + ", Name = " + reader.GetAttribute(1) + ", Type = " + reader.GetAttribute(2) + ", Len = " + reader.GetAttribute(3) + ", Comment = " + reader.GetAttribute(5) + Mynewline;
                            }
                        }
                    }
                }
                if (comboBoxIT.SelectedItem.ToString() == "-- select--") return;
                textBoxOutput.SelectionStart = textBoxOutput.Text.Length;
                textBoxOutput.ScrollToCaret();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can't get Item definitions; please check the format of the file", "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                exit_error(ex.ToString());
            }
        }
        static void AppendToFile(string theFile, string theText)
        {
            StreamWriter SW;
            SW = File.AppendText(theFile);
            SW.WriteLine(theText);
            SW.Close();
        }       

        private string ConvertToODMPartial(string theSource)
        {
            if (theSource.Trim() == "") return ("");
            theSource = theSource.Replace('/', '-');
            try
            {
                string[] splitd = theSource.Split('-');
                if (splitd[0].Length == 4 && (splitd[1].Length == 3 || splitd[1].Length == 2))   // yyyy-mmm or yyyy-mm
                {
                    string mon = DateUtilities.Get_maand(splitd[1]);
                    if (mon.StartsWith("Error")) return (mon);
                    else
                    {
                        if (DateUtilities.CheckYear(splitd[0])) return (splitd[0] + "-" + mon);
                        else return ("Error: Wrong Year in partial date");
                    }
                }
                if (splitd[1].Length == 4 && (splitd[0].Length == 3 || splitd[0].Length == 2))   // mmm-yyyy or mm-yyyy
                {
                    string mon = DateUtilities.Get_maand(splitd[0]);
                    if (mon.StartsWith("Error")) return (mon);
                    else
                    {
                        if (DateUtilities.CheckYear(splitd[1])) return (splitd[1] + "-" + mon);
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
        
        private void exit_error(string probl)
        {   
            using (StreamWriter swlog = new StreamWriter(workdir + "\\OCDataImporter_log.txt"))
            {
                swlog.WriteLine(probl);
            }
            Close();
        }
        private void append_warning(string probl)
        {
            bool found = false;
            foreach (string one in Warnings)
            {
                if (one == probl)
                {
                    found = true;
                    break;
                }
            }
            if (!found)
            {
                Warnings.Add(probl);
                DateTime dt = DateTime.Now;
                textBoxOutput.Text += probl + Mynewline; // 3.02
                probl = dt.Day.ToString("00") + "-" + dt.Month.ToString("00") + "-" + dt.Year + " " + dt.Hour.ToString("00") + ":" + dt.Minute.ToString("00") + ":" + dt.Second.ToString("00") + "    " + probl;
                AppendToFile(LOG, probl);
                WARCOUNT++;
                labelWarningCounter.Text = "WARNINGS: " + WARCOUNT.ToString();
            }
        }

        private void buttonBrowse_Click_1(object sender, EventArgs e)
        {
            Get_work_file();
            button_start.BackColor = System.Drawing.Color.LightGreen;
            buttonBrowse.BackColor = SystemColors.Control;
        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button_start_Click(object sender, EventArgs e)
        {
            WARCOUNT = 0;
            dataGridView1.Rows.Clear();
            if (textBoxInput.Text == "" || (textBoxInput.Text.Contains(";") == false))
            {
                MessageBox.Show("Please enter or select correct input files: Either type two file names separated by a semicolon (;) or use browse button.", "OCDataImporter");
                textBoxInput.Focus();
                return;
            }
            string[] split = textBoxInput.Text.Split(';');
            if (split[0].Contains(".xml") || split[0].Contains(".XML"))
            {
                input_oc = split[0];
                input_data = split[1];
            }
            else
            {
                input_oc = split[1];
                input_data = split[0];
            }
            if (!Get_DataFileItems_FrpmInput(input_data)) return;  // 1.1b
            if (labelOCoidExists)
            {
                input_oid = split[2];
                Get_label_oid(input_oid);
            }

            dmpfilename = input_data.Substring(0, input_data.Length - 4) + "_grid.dmp";
            dmpprm = input_data.Substring(0, input_data.Length - 4) + "_parameters.dmp";
            try
            {
                FileStream fp_ioc;
                long FileSize;
                fp_ioc = new FileStream(split[1], FileMode.Open);
                FileSize = fp_ioc.Length;
                fp_ioc.Close();
                if (FileSize > 2147483647) FileSize = 2147483647;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = (int)FileSize;
                PROGBARSIZE = (int)FileSize;
                progressBar1.Step = linelen;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            workdir = "";
            if (input_oc.Contains("\\")) workdir = input_oc.Substring(0, input_oc.LastIndexOf('\\'));
            else workdir = Directory.GetCurrentDirectory();
            comboBoxSE.Items.Clear();
            comboBoxCRF.Items.Clear();
            comboBoxGR.Items.Clear();
            comboBoxIT.Items.Clear();
            Items.Clear();
          
            try
            {
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.ConformanceLevel = ConformanceLevel.Fragment;
                settings.IgnoreWhitespace = true;
                settings.IgnoreComments = true;
                XmlReader reader = XmlReader.Create(input_oc, settings);

                while (reader.Read())
                {
                    if (reader.Name == "Study")
                    {
                        TheStudyOID = reader.GetAttribute(0);
                        break;
                    }
                }
                reader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can't open selected OC meta file, can't continue. Problem is: " + System.Environment.NewLine + ex.Message, "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                exit_error(ex.ToString());
                return;
            }

            GetStudyEventDef(input_oc);

            if (comboBoxSE.Items.Count == 0)
            {
                MessageBox.Show("No study event definitions found in selected file; please check the format of the file", "OCDataImporter");
                textBoxInput.Focus();
                buttonStartConversion.Enabled = true;
                return;
            }
            if (FillTheGrid() == false) BuildDG(false);
            BuildVerificationArrays();
            StateParametres();
        }

        private void BuildDG(bool matchcolumns)
        {
            string key = "False";
            string dat = "False";
            string sex = "False";
            string pid = "False";
            string dob = "False";
            string std = "False";
            textBoxOutput.Text = "";
            string[] fnparts;
            fnparts = new string[dataGridView1.ColumnCount];
            if (comboBoxSE.SelectedItem.ToString() != "-- select --" && comboBoxCRF.SelectedItem.ToString() != "-- select --")
            {
                int matched = 0;
                bool nomatch = false;
                if (!matchcolumns) dataGridView1.Rows.Clear();
                for (int i = 0; i < DataFileItems.Count; i++)
                {
                    if (DataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("subjectid")) key = "True";
                    else key = "False";
                    if (DataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("subjectsex") ||
                        DataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("gender") ||
                        DataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("geslacht")) sex = "True";
                    else sex = "False";
                    if (DataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("personid")) pid = "True";
                    else pid = "False";
                    if (DataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("subjectdateofbirth") ||
                        DataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("subjectdob") ||
                        DataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("dateofbirth") ||
                        DataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("geboortedatum")) dob = "True";
                    else dob = "False";
                    if (DataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("subjectstartdate")) std = "True";
                    else std = "False";
                    fnparts[DGIndexOfDataItem] = DataFileItems[i].ToString();
                    fnparts[DGIndexOfKey] = key;
                    fnparts[DGIndexOfDate] = dat;
                    fnparts[DGIndexOfSex] = sex;
                    fnparts[DGIndexOfPID] = pid;
                    fnparts[DGIndexOfDOB] = dob;
                    fnparts[DGIndexOfSTD] = std;
                    string theCRF = comboBoxCRF.SelectedItem.ToString();
                    fnparts[DGIndexOfOCItem] = "none";

                    int startGroup = 0;
                    int startEvent = 0;
                    // track _Gx or _Ex in column names and exclude those extentions from column matches. Fixed: 2.0.3
                    for (int jj = 0; jj < 10; jj++)
                    {
                        startGroup = DataFileItems[i].ToString().IndexOf("_G" + jj.ToString());
                        if (startGroup > 0) break;
                    }
                    for (int jj = 0; jj < 10; jj++)
                    {
                        startEvent = DataFileItems[i].ToString().IndexOf("_E" + jj.ToString());
                        if (startEvent > 0) break;
                    }
                    string theItem = DataFileItems[i].ToString().ToLower();
                    if (startEvent > 0) theItem = theItem.Substring(0, startEvent);
                    else if (startGroup > 0) theItem = theItem.Substring(0, startGroup);

                    string theItemOID = GetItemOIDFromItemName(theItem, theCRF);
                    if (theItemOID != "NOTFOUND2" && theItemOID != "ANOTHERCRF") // 3.03
                    {
                        fnparts[DGIndexOfOCItem] = comboBoxSE.SelectedItem.ToString() + "." + comboBoxCRF.SelectedItem.ToString() + "." + GetGroupFromItemCRF(theItemOID, theCRF) + "." + theItemOID;
                        matched++;
                    }
                    else
                    {
                        if (theItemOID == "NOTFOUND2" && key != "True" && sex != "True" && pid != "True" && dob != "True" && std != "True" && theItem != "event_index" && theItem != "group_index") 
                        {
                            textBoxOutput.Text += "No match for: " + theItem + Mynewline;
                            nomatch = true;
                        }
                    }

                    if (!matchcolumns) dataGridView1.Rows.Add(fnparts);
                    else    // version 2.0.2
                    {
                        for (int n = 0; n < dataGridView1.RowCount; n++)
                        {
                            if (dataGridView1.Rows[n].IsNewRow == false)
                         
                            {
                                string gridDataItem = dataGridView1.Rows[n].Cells[DGIndexOfDataItem].Value.ToString();
                                string gridOCItem = dataGridView1.Rows[n].Cells[DGIndexOfOCItem].Value.ToString();
                                if (gridDataItem == fnparts[DGIndexOfDataItem])
                                {
                                    if (gridOCItem == "none" || gridOCItem.StartsWith("Use link button"))
                                    {
                                        dataGridView1.Rows[n].Cells[DGIndexOfOCItem].Value = fnparts[DGIndexOfOCItem];
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
                dataGridView1.Rows.Clear();
                for (int i = 0; i < DataFileItems.Count; i++)
                {
                    if (DataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("subjectid")) key = "True";
                    else key = "False";
                    if (DataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("subjectsex") || 
                        DataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("gender") ||
                        DataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("geslacht")) sex = "True";
                    else sex = "False";
                    if (DataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("personid")) pid = "True";
                    else pid = "False";
                    if (DataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("subjectdateofbirth") ||
                        DataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("subjectdob") ||
                        DataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("dateofbirth") ||
                        DataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("geboortedatum")) dob = "True";
                    else dob = "False";
                    if (DataFileItems[i].ToString().ToLower().Replace("_", "").Replace("-", "").Contains("subjectstartdate")) std = "True";
                    else std = "False";
                    fnparts[DGIndexOfDataItem] = DataFileItems[i].ToString();
                    fnparts[DGIndexOfKey] = key;
                    fnparts[DGIndexOfDate] = dat;
                    fnparts[DGIndexOfSex] = sex;
                    fnparts[DGIndexOfPID] = pid;
                    fnparts[DGIndexOfDOB] = dob;
                    fnparts[DGIndexOfSTD] = std;
                    fnparts[DGIndexOfOCItem] = "Use link button 'CopyTarget' to fill this cell with the selected target item";
                    dataGridView1.Rows.Add(fnparts);
                }
            }
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Process is not finished yet. If you cancel the process now, the generated files will be INCOMPLETE AND CAN NOT BE USED. Are you sure you want to cancel?", "OCDataImporter asks confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.No) return;
            if (MyThread != null && MyThread.IsAlive)
            {
                MyThread.Abort();
                buttonExit.Enabled = true;
                this.Cursor = Cursors.Arrow;
                MyThread = null;
            }
            BackToBegin();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex >= 0)
            {
                if (dataGridView1.Columns[e.ColumnIndex].Name == "CopyCurrentTarget")
                {
                    if (dataGridView1.ReadOnly == true)
                    {
                        MessageBox.Show("The grid is not yet enabled!", "OCDataImporter");
                    }
                    else
                    {
                        if (comboBoxSE.SelectedItem.ToString() == "-- select --" || comboBoxCRF.SelectedItem.ToString() == "-- select --" ||
                            comboBoxGR.SelectedItem.ToString() == "-- select --" || comboBoxIT.SelectedItem.ToString() == "-- select --")
                        {
                            MessageBox.Show("Please use all of the comboboxes above to define a target item. (there are still -- select --'s up there)", "OCDataImporter");
                        }
                        else
                        {
                            // this is needed to get rid of "-CRF name" in Group name, otherwise OC will complain
                            int strubbish = comboBoxGR.SelectedItem.ToString().IndexOf("-" + comboBoxCRF.SelectedItem.ToString());
                            string theGroup = comboBoxGR.SelectedItem.ToString();
                            if (strubbish > 0) theGroup = comboBoxGR.SelectedItem.ToString().Substring(0, strubbish);
                            dataGridView1.Rows[e.RowIndex].Cells[DGIndexOfOCItem].Value = comboBoxSE.SelectedItem.ToString() + "." + comboBoxCRF.SelectedItem.ToString() + "." +
                                theGroup + "." + comboBoxIT.SelectedItem.ToString();
                        }
                    }
                }
            }
        }
        private void linkLabelBuildDG_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (comboBoxSE.Items.Count > 0)
            {
                if (comboBoxSE.SelectedItem.ToString() != "-- select --") BuildDG(true);
                else MessageBox.Show("Please select a StudyEvent and a CRF before matching.", "OCDataImporter");
            }
            else MessageBox.Show("Please read input files first.", "OCDataImporter");
        }
        public string CheckRK(string rk, int line)
        {
            if (IsNumber(rk)) return (rk);
            append_warning("Event and/or Group repeat index can't be determined at line: " + line.ToString());
            return ("ERROR: " + rk);
        }
        public bool IsNumber(String s)
        {
            bool value = true;
            foreach (Char c in s.ToCharArray())
            {
                value = value && Char.IsDigit(c);
            }

            return value;
        }
        public string Escape(string strtoescape)  // 2.1.4 xml escaping
        {
            return (strtoescape.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;"));
        }

        private string GetGroupFromItemCRF(string ItemOID, string FormOID)
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
        
        private string GetItemOIDFromItemName(string ItemName, string FormOID)
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

        private void BuildVerificationArrays()  // 3.0 Data verification structures 
        {
        //*** ItemGroupDefs: CRF, Group contains which items. Item_OID,OPT = optional, Item_OID,MAN = mandatory. Separated by ~
        //F_COCO_V10^IG_COCO_UNGROUPED^I_COCO_LABELLINK,MAN~I_COCO_DATEINTAKE,OPT~I_COCO_DATECOMMENT,OPT~I_COCO_STUDYEXPLANATION,OPT~I_COCO_INFORMEDCONSENTPATIENT,OPT~I_COCO_DATESIGNATUREPATIENT,OPT~I_COCO_INFORMEDCONSENTINVESTIGATOR,OPT~I_COCO_DATESIGNATUREINVESTIGATOR,OPT~I_COCO_STOOLCONSENT,OPT~I_COCO_FITCONSENT,OPT~I_COCO_BLOODCONSENT,OPT~I_COCO_COLONOSCOPYDATE,OPT~I_COCO_STOOLGIVEN,OPT~I_COCO_STOOLCOMMENT,OPT~I_COCO_STOOLDELIVERY,OPT~I_COCO_PICKUPDATE,OPT~I_COCO_PICKUPCOMMENT,OPT~I_COCO_PICKUPCONFIRMATION,OPT~I_COCO_NOPICKUPCOMMENT,OPT~I_COCO_FITDONE,OPT~I_COCO_FITCOMMENT,OPT~I_COCO_FITEIKENBARCODE,OPT~I_COCO_FITONCOBARCODE,OPT
        //F_COCOS_BLOOD__V10^IG_COCOS_UNGROUPED^I_COCOS_BLOOD_THINNER,OPT~I_COCOS_BLOOD_THINNER_SPECIFIED,OPT~I_COCOS_BLOOD_THINNER_SPECIFIED_OTH,OPT~I_COCOS_ASCAL_DOSE,OPT~I_COCOS_PLAVIX_DOSE,OPT~I_COCOS_ACENOFENPRO_DOSE,OPT~I_COCOS_BLOOD_THINNER_OTHER_DOSE,OPT
        //
        // *** ItemDefForm: Item-data attributes in which form and show/hide situation
        //I_COCO_LABELLINK^F_COCO_V10^SHOW^CL_75^integer^1
        //I_COCO_DATECOMMENT^F_COCO_V10^SHOW^NOCODE^text^216
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

            ItemGroupDefs.Clear();
            ItemDefForm.Clear();
            CodeList.Clear();
            RCList.Clear();
            MSList.Clear();
            SCOList.Clear();
            int linenr = 0;

            try
            {
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.ConformanceLevel = ConformanceLevel.Fragment;
                settings.IgnoreWhitespace = true;
                settings.IgnoreComments = true;
                XmlReader reader = XmlReader.Create(input_oc, settings);
                bool res = false;
                res = reader.Read();
                while (res)
                {
                    linenr++;
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
                                if (IsNumber(reader.GetAttribute(4)) && IsNumber(reader.GetAttribute(3))) ItLine = theType + "^" + reader.GetAttribute(3) + "^" + reader.GetAttribute(4); // should be float
                                else // integer or text
                                {
                                    if (IsNumber(reader.GetAttribute(3))) ItLine = theType + "^" + reader.GetAttribute(3) + "^999";
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
                MessageBox.Show("Can't get XML definitions for verification data from the metafile; unexpected exception:", "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                exit_error("Metadatafile Line=" + linenr.ToString() + " -> " + ex.Message);
            }
            if (DEBUGMODE)
            {
                using (StreamWriter swlog = new StreamWriter(workdir + "\\OCDataImporter_verification.txt"))
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
                }
            }
        }
        private string ValidateItem (string key, string FormOID, string ItemOID, string ItemVal, int linenr)
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
                string [] igd = igds.Split('^');
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
            if (man == 0 || itlen == -1) append_warning(fixedwarning + "*** Wrong XML definitions: Item not found in XML. Please inform the ICT. Thank you in advance.");
            if (show == 1 && man == 1 && ItemVal == "") append_warning(fixedwarning + "Item is mandatory but has no value");
            if (show == 2 && man == 1 && ItemVal == "") Hiddens.Add(ItemOID);
            string ConvertedDate = "";
            if (ittype == "date")
            {
                if (comboBoxDateFormat.SelectedItem.ToString() == "--select--") append_warning(fixedwarning + "Item is date but no date format is selected in parameters");
                ConvertedDate = DateUtilities.ConvertToODMFormat(ItemVal, comboBoxDateFormat.SelectedItem.ToString());
                
            }
            if (ittype == "partialDate") ConvertedDate = ConvertToODMPartial(ItemVal);
            if (ConvertedDate.StartsWith("Error")) append_warning(fixedwarning + ConvertedDate);
            else if (ConvertedDate != "") return (ConvertedDate);
            if (ittype == "integer" || ittype == "float")
            {
                string theval = ItemVal.Replace(".", "").Replace(",", "").Replace("-", "");
                if (ittype == "float")
                {
                    if (IsNumber(theval) == false) append_warning(fixedwarning + "Item type is real but contains non numeric characters: " + ItemVal);
                    string thefloatval = ItemVal.Replace(",", ".");
                    if (thefloatval.IndexOf('.') > 0)
                    {
                        int stdec = thefloatval.LastIndexOf('.');
                        string thedec = thefloatval.Substring(stdec + 1);
                        if (thedec.Length > itdeclen) append_warning(fixedwarning + "Item contains more numbers than allowed after the decimal point: " + ItemVal + " (allowed = " + itdeclen.ToString() + " numbers)");
                    }
                }
                else
                {
                    if (IsNumber(ItemVal.Replace("-", "")) == false) append_warning(fixedwarning + "Item type is integer but contains non integer characters: " + ItemVal);
                }
            }
            if (ItemVal.Length > itlen) append_warning(fixedwarning + "Item value exceeds required width = " + itlen.ToString()); 
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
                    if (found == false) append_warning(fixedwarning + "Value not in code list: " + theVals);
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
                            if (found < selvals.Length) append_warning(fixedwarning + "(at least one of) value(s) not in multiple selection list: " + theVals); 
                        }
                    }          
                }
            }
            string huidige = "";
            bool comma = false;
            string uiSep = CultureInfo.CurrentUICulture.NumberFormat.NumberDecimalSeparator;
            if (uiSep.Equals(",")) comma = true;
            try
            {
                foreach (string one in RCList)
                {
                    //I_TEST_ITEM4^text^GT^10
                    //I_TEST_ITEM5^text^GE^1
                    //I_TEST_ITEM5^text^LE^5
                    huidige = one;
                    string[] RCS = one.Split('^');
                    string RCVal = RCS[3].Trim();
                    float fItemVal = 0;
                    float fRCVal = 0;
                    if ((RCS[0] == ItemOID) && ItemVal != "")
                    {
                        if (RCS[1] == "float")
                        {
                            if (comma) fItemVal = System.Convert.ToSingle(ItemVal.Replace('.', ','));
                            else fItemVal = System.Convert.ToSingle(ItemVal);
                            if (comma) fRCVal = System.Convert.ToSingle(RCVal.Replace('.', ','));
                            else fRCVal = System.Convert.ToSingle(RCVal);
                        }
                        if ((RCS[2] == "GT" || RCS[2] == "LT" || RCS[2] == "NE") && (ItemVal == RCVal)) append_warning(fixedwarning + "Range Check fail: Value " + ItemVal + " does not satisfy " + RCS[2] + " " + RCVal);
                        if ((RCS[2] == "EQ") && (ItemVal != RCVal)) append_warning(fixedwarning + "Range Check fail: Value " + ItemVal + " should be equal to " + RCVal);
                        if (RCS[2] == "GT" || RCS[2] == "GE")
                        {
                            if (RCS[1] == "text") if (string.Compare(ItemVal, RCVal) < 0) append_warning(fixedwarning + "Range Check fail: Value " + ItemVal + " is not " + RCS[2] + " " + RCVal);
                            if (RCS[1] == "integer") if (System.Convert.ToInt32(ItemVal) < System.Convert.ToInt32(RCVal)) append_warning(fixedwarning + "Range Check fail: Value " + ItemVal + " is not " + RCS[2] + " " + RCVal);
                            if (RCS[1] == "float") if (fItemVal < fRCVal) append_warning(fixedwarning + "Range Check fail: Value " + ItemVal + " is not " + RCS[2] + " " + RCVal);
                        }
                        if (RCS[2] == "LT" || RCS[2] == "LE")
                        {
                            if (RCS[1] == "text") if (string.Compare(RCVal, ItemVal) < 0) append_warning(fixedwarning + "Range Check fail: Value " + ItemVal + " is not " + RCS[2] + " " + RCVal);
                            if (RCS[1] == "integer") if (System.Convert.ToInt32(RCVal) < System.Convert.ToInt32(ItemVal)) append_warning(fixedwarning + "Range Check fail: Value " + ItemVal + " is not " + RCS[2] + " " + RCVal);
                            if (RCS[1] == "float") if (fRCVal < fItemVal) append_warning(fixedwarning + "Range Check fail: Value " + ItemVal + " is not " + RCS[2] + " " + RCVal);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                append_warning(fixedwarning + "Range check ignored: " + huidige.Replace('^', ',') + ". " + ex.Message);
            }            
            return (ItemVal);
        }
        private void EnableRead()
        {
            textBoxInput.Enabled = true;
            button_start.Enabled = true;
            buttonBrowse.Enabled = true;
            button_start.BackColor = System.Drawing.Color.LightGreen;
            buttonBrowse.BackColor = System.Drawing.Color.LightGreen;
            label16.Visible = true;
            label3.Visible = true;
            buttonBackToBegin.Enabled = false;
            buttonBackToBegin.BackColor = SystemColors.Control;
            buttonCancel.Enabled = false;
            buttonCancel.BackColor = SystemColors.Control;
            label3.ForeColor = System.Drawing.Color.Black;
            label16.ForeColor = System.Drawing.Color.Black;
        }
        private void DisableRead()
        {
            textBoxInput.Enabled = false;
            button_start.Enabled = false;
            buttonBrowse.Enabled = false;
            button_start.BackColor = SystemColors.Control;
            buttonBrowse.BackColor = SystemColors.Control;
            label16.Visible = false;
            label3.Visible = false;
        }
        private void EnableProcess()
        {
            dataGridView1.ReadOnly = false;
            dataGridView1.ForeColor = SystemColors.ControlText;
            CopyCurrentTarget.LinkColor = System.Drawing.Color.Blue;
            comboBoxCRF.Enabled = true;
            comboBoxGR.Enabled = true;
            comboBoxIT.Enabled = true;
            comboBoxSE.Enabled = true;
            linkLabel1.Enabled = true;
            linkLabelBuildDG.Enabled = true;
            linkbuttonSHCols.Enabled = true;
            buttonStartConversion.Enabled = true;
            buttonStartConversion.BackColor = System.Drawing.Color.LightGreen;
            label17.Visible = true;
            label2.ForeColor = System.Drawing.Color.Black;
            label5.ForeColor = System.Drawing.Color.Black;
            label6.ForeColor = System.Drawing.Color.Black;
            label7.ForeColor = System.Drawing.Color.Black;
        }
        private void DisableProcess()
        {
            dataGridView1.Rows.Clear();
            dataGridView1.ReadOnly = true;
            CopyCurrentTarget.LinkColor = System.Drawing.Color.Gray;
            dataGridView1.ForeColor = System.Drawing.Color.Gray;
            comboBoxCRF.Enabled = false;
            comboBoxGR.Enabled = false;
            comboBoxIT.Enabled = false;
            comboBoxSE.Enabled = false;
            linkLabel1.Enabled = false;
            linkLabelBuildDG.Enabled = false;
            linkbuttonSHCols.Enabled = false;
            buttonStartConversion.Enabled = false;
            buttonCancel.Enabled = false;
            buttonCancel.BackColor = SystemColors.Control;
            buttonStartConversion.BackColor = SystemColors.Control;
            label17.Visible = false;
            label2.ForeColor = System.Drawing.Color.LightGray;
            label5.ForeColor = System.Drawing.Color.LightGray;
            label6.ForeColor = System.Drawing.Color.LightGray;
            label7.ForeColor = System.Drawing.Color.LightGray;
        }
        private void DisableParams()
        {
            label15.Visible = false;
            textBoxReplace.Enabled = false;
            comboBoxDateFormat.Enabled = false;
            comboBoxSex.Enabled = false;
            textBoxSubjectSexM.Enabled = false;
            textBoxSubjectSexF.Enabled = false;
            radioButtonUseTD.Enabled = false;
            radioButtonNoEVT.Enabled = false;
            textBoxMaxLines.Enabled = false;
            textBoxLocation.Enabled = false;
            checkBoxDup.Enabled = false;
            buttonConfPars.BackColor = SystemColors.Control;
            buttonConfPars.Enabled = false;
            label11.ForeColor = System.Drawing.Color.LightGray;
            label4.ForeColor = System.Drawing.Color.LightGray;
            label8.ForeColor = System.Drawing.Color.LightGray;
            label10.ForeColor = System.Drawing.Color.LightGray;
            label12.ForeColor = System.Drawing.Color.LightGray;
            label13.ForeColor = System.Drawing.Color.LightGray;
            label14.ForeColor = System.Drawing.Color.LightGray;
            groupBox1.ForeColor = System.Drawing.Color.LightGray;
            checkBoxDup.ForeColor = System.Drawing.Color.LightGray;
        }
        private void EnableParams()
        {
            label15.Visible = true;
            textBoxReplace.Enabled = true;
            comboBoxDateFormat.Enabled = true;
            comboBoxSex.Enabled = true;
            textBoxSubjectSexM.Enabled = true;
            textBoxSubjectSexF.Enabled = true;
            radioButtonUseTD.Enabled = true;
            radioButtonNoEVT.Enabled = true;
            textBoxMaxLines.Enabled = true;
            textBoxLocation.Enabled = true;
            checkBoxDup.Enabled = true;
            buttonConfPars.BackColor = System.Drawing.Color.LightGreen;
            buttonConfPars.Enabled = true;
            buttonBackToBegin.BackColor = System.Drawing.Color.LightGreen;
            buttonBackToBegin.Enabled = true;
            label11.ForeColor = System.Drawing.Color.Black;
            label4.ForeColor = System.Drawing.Color.Black;
            label8.ForeColor = System.Drawing.Color.Black;
            label10.ForeColor = System.Drawing.Color.Black;
            label12.ForeColor = System.Drawing.Color.Black;
            label13.ForeColor = System.Drawing.Color.Black;
            label14.ForeColor = System.Drawing.Color.Black;
            groupBox1.ForeColor = System.Drawing.Color.Black;
            checkBoxDup.ForeColor = System.Drawing.Color.Black;
        }
        private void StateReadFiles()
        {
            // initialize program variables
            buttonExit.BackColor = System.Drawing.Color.LightGreen;
            buttonConfPars.Enabled = false;
            buttonConfPars.BackColor = SystemColors.Control;
            comboBoxDateFormat.SelectedIndex = 0;
            comboBoxSex.SelectedIndex = 0; 
            labelOCoidExists = false;
            dmpfilename = "";
            dmpprm = "";
            workdir = "";
            input_oc = "";
            input_data = "";
            input_oid = "";
            Items.Clear();
            DataFileItems.Clear();
            LabelOID.Clear();
            ReplacePairStrings.Clear();
            SortableDG.Clear();
            InsertSubject.Clear();
            ItemGroupDefs.Clear();
            ItemDefForm.Clear();
            CodeList.Clear();
            MSList.Clear();
            RCList.Clear();
            Warnings.Clear();
            TheStudyOID = "";
            TheStudyEventOID = "";
            EventStartDates = "";
            TheItemId = "";
            TheFormOID = "";
            TheItemGroupDef = "";
            RepSeparator = ';';
            InsertKeys.Clear();
            INSF = "";
            LOG = "";
            INSFR = "";
            DIMF = "";
            DELF = "";
            DELFR = "";
            Delimiter = ';';
            sepcount = 1;
            PROGBARSIZE = 0;
            WARCOUNT = 0;
            keyIndex = 0;
            sexIndex = 0;
            PIDIndex = 0;
            DOBIndex = 0;
            STDIndex = 0;
            sexItem = "";
            PIDItem = "";
            DOBItem = "";
            STDItem = "";
            linelen = 0;
            SUBJECTSEX_M = "";
            SUBJECTSEX_F = "";
            repeating_groups = "";
            repeating_events = "";
            OUTF = "";
            OUTFBASIS = "";
            OUTFLINECOUNTER = 0;
            OUTFFILECOUNTER = 1;
            OUTFMAXLINES = 0;
            selectedEventRepeating = "No";
            EnableRead();
            textBoxInput.Focus();
            DisableProcess();
            DisableParams();
        }
        private void StateParametres()
        {
            DisableRead();
            EnableParams();
        }
        private void StateProcess()
        {
            DisableParams();
            EnableProcess();
        }
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            for (int n = 0; n < dataGridView1.RowCount; n++)
            {
                if (dataGridView1.Rows[n].IsNewRow == false)
                {
                    dataGridView1.Rows[n].Cells[DGIndexOfOCItem].Value = "none";
                }
            }
        }
        private void buttonConfPars_Click(object sender, EventArgs e)
        {
            if (textBoxLocation.Text == "")
            {
                MessageBox.Show("Please enter location", "OCDataImporter");
                return;
            }
            StateProcess();
        }
        private void BackToBegin()
        {
            StateReadFiles();
            progressBar1.Value = 0;
            textBoxOutput.Text = "";
            labelWarningCounter.Text = "";
        }
        private void buttonBackToBegin_Click(object sender, EventArgs e)
        {
            BackToBegin();
        }

        private void linkbuttonSHCols_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (linkbuttonSHCols.Text.StartsWith("Hide"))
            {
                linkbuttonSHCols.Text = "Show All Columns";
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    dataGridView1.Columns[DGIndexOfDOB].Visible = false;
                    dataGridView1.Columns[DGIndexOfKey].Visible = false;
                    dataGridView1.Columns[DGIndexOfPID].Visible = false;
                    dataGridView1.Columns[DGIndexOfSex].Visible = false;
                    dataGridView1.Columns[DGIndexOfSTD].Visible = false;
                }
            }
            else
            {
                linkbuttonSHCols.Text = "Hide Subject Related Columns";
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    dataGridView1.Columns[DGIndexOfDOB].Visible = true;
                    dataGridView1.Columns[DGIndexOfKey].Visible = true;
                    dataGridView1.Columns[DGIndexOfPID].Visible = true;
                    dataGridView1.Columns[DGIndexOfSex].Visible = true;
                    dataGridView1.Columns[DGIndexOfSTD].Visible = true;
                }
            }
        }
    }
}