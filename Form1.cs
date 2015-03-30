﻿using System;
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
using System.Security;
using System.Security.AccessControl;
using System.Data.OleDb;

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
//4.3.1     02-04-2014      Replace couples textbox
//                          deleted; not needed                 C. Parlayan, J. Rousseau
//4.3.1     03-04-2014      Allow date of year
//                          instead of date of birth            C. Parlayan
// 4.3.1    03-04-2014      subject_event_status_id from 3 
//                          to 1 to trigger audit               C. Parlayan, S. de Ridder
//4.3.1     03-04-2014      SITE_OID column introduced to 
//                          allow defining site for subjects    C. Parlayan, J. Rousseau
//4.3.1     04-04-2014      Names in dropdowns's beside  
//                          OID's                               C. Parlayan, S. de Ridder
//4.3.2     14-04-2014      Fix buttons enable/disable status
//                          and colors                          C. Parlayan
//4.3.3     14-04-2014      Dont change date format when back
//                          to begin button is hit              C. Parlayan, S. de Ridder
//4.3.4     23-04-2014      Date of year must end with "-01-01"
//                          to prevent OC giving error          C. Parlayan, S. de Ridder
//4.3.5     23-04-2014      culture info added for
//                          comparing floats                    C. Parlayan, J. Rousseau
//4.3.6     01-05-2014      EVENT_INDEX and GROUP_INDEX bug
//                          resolved                            C. Parlayan
// 4.3.6    02-05-2014      Lenght and type int/float control 
//                          must take place only if value not in code list.   C.Parlayan
// 4.3.7    02-05-2014      check if EVENT_INDEX is integer 
//                          and not empty.                      C.Parlayan
// 4.3.8    05-05-2014      prevent duplicate inserts           C.Parlayan
// 4.3.9    09-05-2014      Error handling Subj.Start Date      C.Parlayan
// 4.3.10   23-07-2014      Do nothing if No is chosen for
//                          "Do you want to delete *.xml"?      C. Parlayan  
// 4.3.10   25-07-2014      Split factor = 75 instead of 0      C. Parlayan
// 4.3.10   15-08-2014      Create records in TDS database
//                          for web services                    C. Parlayan
// 4.3.12/13 20-08-2014     YearOfBirth added to TDS database   C. Parlayan, S. de Ridder
// 4.3.15    20-08-2014     Ignore duplicate; for consequtive events    C.Parlayan
// 4.3.17    21-08-2014     Group index = 1 if cant be 
//                          determined by column (group_index_row = -1) C.Parlayan
// 4.3.18    01-09-2014     When more than one events used in
//                          grid, only for one event the inserts
//                          are created. This is now fixed.     C. Parlayan
// 4.3.19/20 08-09-2014     No PID? Leave PID blank when        
//                          creating subject                    C. Parlayan
// 4.3.21    16-09-2014     Accept full dates in Pdate        
//                          but return yyyy-mm                  C. Parlayan
// 4.3.22    17-09-2014     Fix:File selection                  C.Parlayan
// 4.3.23    04-12-2014     Fix:Insert2 bug fixed               C.Parlayan
// 4.3.24    13-01-2015     Fix:Trim SE OID in combobox         C.Parlayan
// 4.3.25    2-2-2015       Empty <SubjectData> delete from xml C.Parlayan 
// 4.3.25    2-2-2015       Merge main and z versions           C.Parlayan
//                          label1 (the header) -> contains (z) version without TDS dbase
//                          else full version. 
// 4.3.26    9-3-2015       Fix: Gender not mandatory in Inserts    C.Parlayan
namespace OCDataImporter
{
    public partial class Form1 : Form
    {
        public bool DEBUGMODE = false;
        // public bool DEBUGMODE = true;
        public bool labelOCoidExists = false; // 2.1.1 If labelOCoid file exists, get the oid from that file, instead of 'SS_label'...
        public bool labelMdbExists = false;   // 4.3.10
        public bool needInserts = true;
        public bool zonder = false;
        public string dmpfilename = "";
        public string dmpprm = "";
        Thread MyThread;
        FileStream fpipdf;
        string Mynewline = System.Environment.NewLine;
        static public string workdir = "";
        static public string input_oc = "";
        static public string input_data = "";
        static public string input_mdb = "";
        static public string input_oid = "";
        static public string mdbConn = "";
        ArrayList NewSubjectsToCreate = new ArrayList();
        ArrayList NewEventsToCreate = new ArrayList();
        ArrayList Items = new ArrayList();
        ArrayList DataFileItems = new ArrayList();
        ArrayList LabelOID = new ArrayList();
        ArrayList SortableDG = new ArrayList();
        ArrayList InsertSubject = new ArrayList();
        ArrayList Insert2s = new ArrayList();
        ArrayList Insert3s = new ArrayList();
        ArrayList ItemGroupDefs = new ArrayList();
        ArrayList ItemDefForm = new ArrayList();
        ArrayList CodeList = new ArrayList();
        ArrayList MSList = new ArrayList();
        ArrayList RCList = new ArrayList();
        ArrayList SCOList = new ArrayList();
        ArrayList Warnings = new ArrayList();
        ArrayList Sites = new ArrayList();
        ArrayList Forms = new ArrayList();
        ArrayList Groups = new ArrayList();
        ArrayList AllValuesInOneRow = new ArrayList();
        ArrayList Hiddens = new ArrayList();
        ArrayList ArrDIMF = new ArrayList();
        ArrayList SEArray = new ArrayList();
        ArrayList SEArrayOnlyE = new ArrayList();
        string TheStudyOID = "";
        string TheStudyEventOID = "";
        string EventStartDates = "";
        string TheItemId = "";
        string TheFormOID = "";
        string TheItemGroupDef = "";
        ArrayList InsertKeys = new ArrayList();
        FileStream fpoDIM;
        FileStream fpoLOG;
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
        int DGIndexOfDataItem = 0;
        int DGIndexOfOCItem = 1;
        int DGIndexOfKey = 2;
        int DGIndexOfDate = 3;
        int DGIndexOfSex = 4;
        int DGIndexOfPID = 5;
        int DGIndexOfDOB = 6;
        int DGIndexOfSTD = 7;
        string SUBJECTSEX_M = "";
        string SUBJECTSEX_F = "";
        string SUBJECTSEX_NOTHING = "";
        string repeating_groups = "";
        string repeating_events = "";
        // ODM file splitting
        string OUTF = "";
        string OUTFBASIS = "";
        int OUTFLINECOUNTER = 0;
        int OUTFFILECOUNTER = 1;
        int OUTFMAXLINES = 0;
        bool DOY = false;  // Study patrameter: Year of birth used?
        string PIDR = "";  // Study patrameter: PersonID required
        string at0 = "";
        string at1 = "";
        string selectedEventRepeating = "No";
        
        static public string insert1a = "INSERT INTO subject(status_id, gender, unique_identifier, date_created, owner_id, dob_collected, date_of_birth)";
        static public string insert1 = "INSERT INTO subject(status_id, gender, unique_identifier, date_created, owner_id, dob_collected)";
        static public string insert1az = "INSERT INTO subject(status_id, unique_identifier, date_created, owner_id, dob_collected, date_of_birth)";
        static public string insert1z = "INSERT INTO subject(status_id, unique_identifier, date_created, owner_id, dob_collected)";
        static public string insert2 = "INSERT INTO study_subject(label, study_id, status_id, enrollment_date, date_created, date_updated, owner_id,  oc_oid, subject_id)";
        static public string insert3 = "INSERT INTO study_event(study_event_definition_id, study_subject_id, location, sample_ordinal, date_start, owner_id, status_id, date_created, subject_event_status_id, start_time_flag, end_time_flag)";
        
        public Form1()
        {
            InitializeComponent();
            comboBoxSE.DrawMode = DrawMode.OwnerDrawFixed;
            comboBoxSE.DrawItem += new DrawItemEventHandler(comboBoxSE_DrawItem);
            comboBoxSE.DropDownClosed += new EventHandler(comboBoxSE_DropDownClosed);
            comboBoxSE.Leave += new EventHandler(comboBoxSE_Leave);
            comboBoxCRF.DrawMode = DrawMode.OwnerDrawFixed;
            comboBoxCRF.DrawItem += new DrawItemEventHandler(comboBoxCRF_DrawItem);
            comboBoxCRF.DropDownClosed += new EventHandler(comboBoxCRF_DropDownClosed);
            comboBoxCRF.Leave += new EventHandler(comboBoxCRF_Leave);
            comboBoxGR.DrawMode = DrawMode.OwnerDrawFixed;
            comboBoxGR.DrawItem += new DrawItemEventHandler(comboBoxGR_DrawItem);
            comboBoxGR.DropDownClosed += new EventHandler(comboBoxGR_DropDownClosed);
            comboBoxGR.Leave += new EventHandler(comboBoxGR_Leave);
            comboBoxIT.DrawMode = DrawMode.OwnerDrawFixed;
            comboBoxIT.DrawItem += new DrawItemEventHandler(comboBoxIT_DrawItem);
            comboBoxIT.DropDownClosed += new EventHandler(comboBoxIT_DropDownClosed);
            comboBoxIT.Leave += new EventHandler(comboBoxIT_Leave);
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
            //if (label1.Text.Contains("z"))  // 4.3.25 version without usage TDS database. 
            //{                                 // if you want to let the program support TDS database, remove comment from the if
                zonder = true;
                label3.Text = "OCMetaData file (XML), Study Data file (TXT), label-oid file (OID) separated by a ';' or use 'Browse' button. Label-oid file is optional.";
            //}
            try
            {
                fpipdf = new FileStream("OCDataImporter.pdf", FileMode.Open, FileAccess.Read);
            }
            catch (Exception exx)
            {
                MessageBox.Show("Problem opening user manual. Message = " + exx.Message, "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            workdir = Directory.GetCurrentDirectory();
            StateReadFiles(true);
        }

        void comboBoxSE_Leave(object sender, EventArgs e)
        {
            toolTip1.Hide(comboBoxSE);
        }

        void comboBoxSE_DropDownClosed(object sender, EventArgs e)
        {
            toolTip1.Hide(comboBoxSE);
        }

        void comboBoxSE_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) { return; }
            string text = comboBoxSE.GetItemText(comboBoxSE.Items[e.Index]);
            e.DrawBackground();
            using (SolidBrush br = new SolidBrush(e.ForeColor))
            { e.Graphics.DrawString(text, e.Font, br, e.Bounds); }
            if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
            { toolTip1.Show(text, comboBoxSE, e.Bounds.Right, e.Bounds.Bottom); }
            e.DrawFocusRectangle();
        }

        void comboBoxCRF_Leave(object sender, EventArgs e)
        {
            toolTip1.Hide(comboBoxCRF);
        }

        void comboBoxCRF_DropDownClosed(object sender, EventArgs e)
        {
            toolTip1.Hide(comboBoxCRF);
        }

        void comboBoxCRF_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) { return; }
            string text = comboBoxCRF.GetItemText(comboBoxCRF.Items[e.Index]);
            e.DrawBackground();
            using (SolidBrush br = new SolidBrush(e.ForeColor))
            { e.Graphics.DrawString(text, e.Font, br, e.Bounds); }
            if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
            { toolTip1.Show(text, comboBoxCRF, e.Bounds.Right, e.Bounds.Bottom); }
            e.DrawFocusRectangle();
        }

        void comboBoxGR_Leave(object sender, EventArgs e)
        {
            toolTip1.Hide(comboBoxGR);
        }

        void comboBoxGR_DropDownClosed(object sender, EventArgs e)
        {
            toolTip1.Hide(comboBoxGR);
        }

        void comboBoxGR_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) { return; }
            string text = comboBoxGR.GetItemText(comboBoxGR.Items[e.Index]);
            e.DrawBackground();
            using (SolidBrush br = new SolidBrush(e.ForeColor))
            { e.Graphics.DrawString(text, e.Font, br, e.Bounds); }
            if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
            { toolTip1.Show(text, comboBoxGR, e.Bounds.Right, e.Bounds.Bottom); }
            e.DrawFocusRectangle();
        }

        void comboBoxIT_Leave(object sender, EventArgs e)
        {
            toolTip1.Hide(comboBoxIT);
        }

        void comboBoxIT_DropDownClosed(object sender, EventArgs e)
        {
            toolTip1.Hide(comboBoxIT);
        }

        void comboBoxIT_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) { return; }
            string text = comboBoxIT.GetItemText(comboBoxIT.Items[e.Index]);
            e.DrawBackground();
            using (SolidBrush br = new SolidBrush(e.ForeColor))
            { e.Graphics.DrawString(text, e.Font, br, e.Bounds); }
            if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
            { toolTip1.Show(text, comboBoxIT, e.Bounds.Right, e.Bounds.Bottom); }
            e.DrawFocusRectangle();
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
            if ((mfd.fnxml != "" && textBoxInput.Text.Contains(".xml")) || (mfd.fntxt != "" && textBoxInput.Text.Contains(".txt")) || (mfd.fnoid != "" && textBoxInput.Text.Contains(".oid")) || (mfd.fnmdb != "" && textBoxInput.Text.Contains(".mdb"))) textBoxInput.Text = "";
            if (mfd.fnxml != "")
            {
                if (textBoxInput.Text.Contains(".xml")) textBoxInput.Text = mfd.fnxml;
                else textBoxInput.Text += ";" + mfd.fnxml;
            }
            if (mfd.fntxt != "")
            {
                if (textBoxInput.Text.Contains(".txt")) textBoxInput.Text = mfd.fntxt;
                else textBoxInput.Text += ";" + mfd.fntxt;
            }
            if (mfd.fnoid != "")
            {
                if (textBoxInput.Text.Contains(".oid")) textBoxInput.Text = mfd.fnoid;
                else textBoxInput.Text += ";" + mfd.fnoid;
            }
            if (mfd.fnmdb != "")
            {
                if (textBoxInput.Text.Contains(".mdb")) textBoxInput.Text = mfd.fnmdb;
                else textBoxInput.Text += ";" + mfd.fnmdb;
            }
            if (textBoxInput.Text.StartsWith(";")) textBoxInput.Text = textBoxInput.Text.Substring(1);
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
            FormAbout Section = new FormAbout();
            Section.Show();
            // MessageBox.Show("OCDataImporter Version 4.x Made by: C. Parlayan, VU Medical Center, Dept. of Pathology, Amsterdam, The Netherlands - 2010-2014", Text);
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
                            textBoxLocation.Text = split[5];
                            if (split.Length > 6) comboBoxSE.SelectedItem = split[6];
                            if (split.Length > 7) comboBoxCRF.SelectedItem = split[7];
                            first = false;
                        }
                        else
                        {
                            // textBoxReplace.Text += Mynewline + line;
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
                textBoxOutput.Text += Mynewline + "No dmp file found." + Mynewline;
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
                    outPR.WriteLine(comboBoxDateFormat.SelectedItem + "~" + comboBoxSex.SelectedItem + "~" + textBoxSubjectSexM.Text + "~" + textBoxSubjectSexF.Text + "~" + textBoxMaxLines.Text + "~" + textBoxLocation.Text + "~" + comboBoxSE.SelectedItem + "~" + comboBoxCRF.SelectedItem);
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
            string today = getTodaysDate();
            ArrDIMF.Clear();
            ArrDIMF.Add("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            ArrDIMF.Add("<ODM xmlns=\"http://www.cdisc.org/ns/odm/v1.3\"");
            ArrDIMF.Add("xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"");
            ArrDIMF.Add("xsi:schemaLocation=\"http://www.cdisc.org/ns/odm/v1.3 ODM1-3.xsd\"");
            ArrDIMF.Add("ODMVersion=\"1.3\" FileOID=\"1D20080412202420\" FileType=\"Snapshot\"");
            ArrDIMF.Add("Description=\"Dataset ODM\" CreationDateTime=\"" + today + "T10:00:00\" >");
            ArrDIMF.Add("<ClinicalData StudyOID=\"" + TheStudyOID + "\" MetaDataVersionOID=\"v1.0.0\">");
        }

        private void buttonStartConversion_Click_1(object sender, EventArgs e)
        {
            SUBJECTSEX_M = textBoxSubjectSexM.Text;
            SUBJECTSEX_F = textBoxSubjectSexF.Text;
            if (IsNumber(textBoxMaxLines.Text) == false)
            {
                MessageBox.Show("Split factor must be a number, 0 means no splittig.", "OCDataImporter");
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
                if (MessageBox.Show("Previously generated SQL files and DataImport_* files has to be overwritten. Do you want to delete the old files?", "Confirm delete old DataImport*.xml files", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
                {
                    foreach (string f in txtList) File.Delete(f);
                }
                else return; // 4.3.10
            }

            dumpTheGrid();
            int maxSE = 0;
            int maxCRF = 0;
            int maxGR = 0;
            SortableDG.Clear();
            fpoLOG = new FileStream(workdir + "\\OCDataImporter_warnings.txt", FileMode.Create, FileAccess.Write);
            StreamWriter swLOG = new StreamWriter(fpoLOG);
            fpoDIM = new FileStream(OUTF, FileMode.Create, FileAccess.Write);
            StreamWriter swDIMw = new StreamWriter(fpoDIM);
            fpoDIM.Close();
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
                    try
                    {
                        string[] mi = dataGridView1.Rows[i].Cells[DGIndexOfOCItem].Value.ToString().Split('.');
                        if (mi.Length == 1 && (mi[0] == "none" || mi[0].StartsWith("Use link button"))) continue;
                        if (mi[0].Length > maxSE) maxSE = mi[0].Length;
                        if (mi[1].Length > maxCRF) maxCRF = mi[1].Length;
                        if (mi[2].Length > maxGR) maxGR = mi[2].Length;
                    }
                    catch (NullReferenceException ex) 
                    {
                        MessageBox.Show("OC Target Item of " +  dataGridView1.Rows[i].Cells[DGIndexOfDataItem].Value.ToString() + " can't be blank. Use 'none' if you do not want to match.", "OCDataImporter");
                        return;
                    }
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

            bool EvtIdxPresent = false;  // this and next for loop 4.3.6
            bool GrpIdxPresent = false;

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].IsNewRow == false)
                {
                    if (dataGridView1.Rows[i].Cells[DGIndexOfDataItem].Value.ToString() == "EVENT_INDEX") EvtIdxPresent = true;
                    if (dataGridView1.Rows[i].Cells[DGIndexOfDataItem].Value.ToString() == "GROUP_INDEX") GrpIdxPresent = true;
                }
            }

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
                        if (repeating_events.Contains(mi[0]) && !EvtIdxPresent) ev_rep = Get_SE_RepeatingKey_FromStudyDataColumn(theDataItem);
                        if (repeating_groups.Contains(mi[2]) && !GrpIdxPresent) gr_rep = Get_GR_RepeatingKey_FromStudyDataColumn(theDataItem);
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

            if (input_data.Contains("\\") == false) input_data = workdir + "\\" + input_data;
            if (input_mdb.Contains("\\") == false) input_mdb = workdir + "\\" + input_mdb;
            if (Delimiter != tab) textBoxOutput.Text += "\r\nData file is: " + input_data + ", delimited by: " + Delimiter + " Number of items per line: " + sepcount + "\r\n";
            else textBoxOutput.Text += "\r\nData file is: " + input_data + ", delimited by: tab, Number of items per line: " + sepcount + "\r\n";
            textBoxOutput.Text += "Started in directory " + workdir + ". This may take several minutes...\r\n";
            buttonBackToBegin.Enabled = false;
            buttonBackToBegin.BackColor = SystemColors.Control;
            buttonStartConversion.Enabled = false;
            buttonStartConversion.BackColor = SystemColors.Control;
            buttonExit.Enabled = false;
            buttonExit.BackColor = SystemColors.Control;
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
            string day = "";
            string month = "";
            if (dt.Day < 10) day = "0";
            if (dt.Month < 10) month = "0";
            day += dt.Day.ToString();
            month += dt.Month.ToString();
            return (dt.Year.ToString() + "-" + month + "-" + day);
        }

        public void DoWork()
        {
            ArrayList ArrINSF = new ArrayList();
            ArrayList ArrINSFR = new ArrayList();
            ArrayList ArrDELF = new ArrayList();
            ArrayList ArrDELFR = new ArrayList();
            SEArray.Clear();
            SEArrayOnlyE.Clear();
            linkLabelBuildDG.Enabled = false;
            linkbuttonSHCols.Enabled = false;
            linkLabel1.Enabled = false;
            string theDate = getTodaysDate();
            string theWrittenSE = "";
            string theWrittenGR = "";
            string [] theHeaders;
            int event_index_row = -1;
            int group_index_row = -1;
            int site_index_row = -1;
            string usedStudyOID = TheStudyOID;
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
                        if (line.Trim().Length == 0) continue; // 4.3.6
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
                                if (theHeaders[i].ToUpper() == "SITE_OID") site_index_row = i;
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
                                append_warning(errtext);
                            }
                            usedStudyOID = split[site_index_row];
                        }
                        else usedStudyOID = TheStudyOID;
                        if (site_index_row != -1 && split[site_index_row].Trim() == "")
                        {
                            string errtext = "No site specified under SITE_OID; assuming STUDY_OID, at line: " + linecount.ToString();
                            append_warning(errtext);
                        }
                        string SStheKEY = "SS_" + theKEY;
                        if (labelOCoidExists) // 2.1.1 there is a conversion file from label to oid; get the SSid from that file.
                        {
                            foreach (string one in LabelOID)
                            {
                                if (one.StartsWith(theKEY + "^")) SStheKEY = one.Substring(one.IndexOf('^') + 1); 
                            }
                        }
                        string theXMLForm = "";
                        ArrayList ArrDIMFtmp = new ArrayList();
                        ArrDIMFtmp.Add("    <SubjectData SubjectKey=\"" + SStheKEY + "\">");
                        theXMLEvent += "        <StudyEventData StudyEventOID=\"" + theWrittenSE + "\" StudyEventRepeatKey=\"" + CheckRK(theSERK, linecount) + "\">" + Mynewline;
                        IsDuplicateSE(theWrittenSE + "^" + CheckRK(theSERK, linecount));
                        IsDuplicateSEOnlyE(theWrittenSE);
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
                            itemval = ValidateItem(theKEY, TheFormOID, TheItemId, itemval, linecount);
                            if (itemval != "")
                            {
                                theXMLForm += "                    <ItemData ItemOID=\"" + TheItemId + "\" Value=\"" + SecurityElement.Escape(itemval) + "\" />" + Mynewline;
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
                                        itemval = ValidateItem(theKEY, TheFormOID, nwdingen[3], itemval, linecount);
                                        if (itemval != "")
                                        {
                                            theXMLForm += "                    <ItemData ItemOID=\"" + nwdingen[3] + "\" Value=\"" + SecurityElement.Escape(itemval) + "\" />" + Mynewline;
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
                                        itemval = ValidateItem(theKEY, TheFormOID, nwdingen[3], itemval, linecount);
                                        if (itemval != "")
                                        {
                                            theXMLForm += "                    <ItemData ItemOID=\"" + nwdingen[3] + "\" Value=\"" + SecurityElement.Escape(itemval) + "\" />" + Mynewline;
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
                                        itemval = ValidateItem(theKEY, TheFormOID, nwdingen[3], itemval, linecount);
                                        if (itemval != "")
                                        {
                                            theXMLForm += "                    <ItemData ItemOID=\"" + nwdingen[3] + "\" Value=\"" + SecurityElement.Escape(itemval) + "\" />" + Mynewline;
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
                                        ArrDIMFtmp.Add(theXMLEvent);
                                        datapresentevent = false;
                                    }
                                    theXMLEvent = "";
                                    if (selectedEventRepeating == "Yes")
                                    {
                                        if (event_index_row != -1) theSERK = split[event_index_row];
                                        else theSERK = Get_SE_RepeatingKey_FromStudyDataColumn(theStudyDataColumn);
                                    }
                                    theXMLEvent += "        <StudyEventData StudyEventOID=\"" + theWrittenSE + "\" StudyEventRepeatKey=\"" + CheckRK(theSERK, linecount) + "\">" + Mynewline;
                                    IsDuplicateSE(theWrittenSE + "^" + CheckRK(theSERK, linecount));
                                    IsDuplicateSEOnlyE(theWrittenSE);
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
                                        itemval = ValidateItem(theKEY, TheFormOID, nwdingen[3], itemval, linecount);
                                        if (itemval != "")
                                        {
                                            theXMLForm += "                    <ItemData ItemOID=\"" + nwdingen[3] + "\" Value=\"" + SecurityElement.Escape(itemval) + "\" />" + Mynewline;
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
                            SubSex = SubSex.Trim();
                        }
                        if (SubSex == SUBJECTSEX_M) SubSex = "m";  // 1.0f
                        if (SubSex == SUBJECTSEX_F) SubSex = "f";
                        if (SubSex == "nothing") SubSex = "";
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
                            ArrDIMFtmp.Add(theXMLEvent);
                            datapresentevent = false;
                        }
                        theXMLEvent = "";
                        if (ArrDIMFtmp.Count > 1) // there is more than just <StudySubject SubjectId = ... 4.3.25
                        {
                            foreach (string one in ArrDIMFtmp) ArrDIMF.Add(one);
                            ArrDIMF.Add("    </SubjectData>");
                        }
                        ArrDIMFtmp.Clear();

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
                        int theSERKInt = 0;
                        try  // 4.3.7 if empty or non-integer; crashes. 
                        {
                            theSERKInt = System.Convert.ToInt16(theSERK);
                        }
                        catch
                        {
                            string errtext = "EVENT_INDEX can't be determined at line " + linecount.ToString() + ". Index: " + theSERK + ". (must be an integer)";
                            append_warning(errtext);
                        }
                        string theDOB = "";
                        if (DOBIndex >= 0)
                        {
                            if (DOY && !labelMdbExists && split[DOBIndex].Trim().Length == 4) theDOB = split[DOBIndex].Trim() + "-01-01"; // 4.3.4 OC accepts only full dates!!
                            else  // 4.3.12
                            {
                                if (DOY)
                                {
                                    if (CheckYear(split[DOBIndex])) theDOB = split[DOBIndex];
                                    else
                                    {
                                        string errtext = "Birth year at line " + linecount.ToString() + " is wrong " + split[DOBIndex];
                                        append_warning(errtext);
                                        theDOB = split[DOBIndex];
                                    }
                                }
                                else theDOB = ConvertToODMFormat(split[DOBIndex]);
                            }
                        }
                        string theSTD = "";
                        if (STDIndex >= 0)
                        {
                            theSTD = ConvertToODMFormat(split[STDIndex]); // This is needed for non repeating events
                            if (theSTD.StartsWith("Error")) append_warning("Wrong subject start date at line " + linecount.ToString());
                        }
                        string thePID = "";
                        // 4.3.19: this doesnt work if subject Person ID is 'not used' in study paramaters: if (PIDIndex < 0) thePID = theKEY; so changed to below statement
                        if (PIDIndex >= 0)
                        {
                            thePID = split[PIDIndex];
                            if (PIDR.Contains("not")) append_warning("PersonID is configured as 'not used' but there is a personID defined in Grid. OpenClinica might reject these records.");
                        }
                        else
                        {
                            if (PIDR.Contains("not") == false) thePID = theKEY;
                        }
                        if (theDOB.StartsWith("Error") || theDOB == "" || DOBIndex < 0)
                        {
                            if (theDOB == "" || DOBIndex < 0)
                            {
                                if (!IsDuplicatePID(thePID))
                                {
                                    if (needInserts)
                                    {
                                        if (SubSex == "")
                                        {
                                            ArrINSF.Add(insert1z);
                                            string myPID = thePID;
                                            if (myPID == "") myPID = theKEY;
                                            ArrINSF.Add("    VALUES (1, '" + myPID + "', '" + theDate + "', 1, '1');");
                                        }
                                        else
                                        {
                                            ArrINSF.Add(insert1);
                                            string myPID = thePID;
                                            if (myPID == "") myPID = theKEY;
                                            ArrINSF.Add("    VALUES (1, '" + SubSex + "', '" + myPID + "', '" + theDate + "', 1, '1');");
                                        }
                                    }
                                    if (labelMdbExists) NewSubjectsToCreate.Add(theKEY + "^" + theDate + "^" + thePID + "^" + SubSex + "^"); // 4.3.10
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
                                if (needInserts)
                                {
                                    if (SubSex == "")
                                    {
                                        ArrINSF.Add(insert1az);
                                        string myPID = thePID;
                                        if (myPID == "") myPID = theKEY;
                                        ArrINSF.Add("    VALUES (1, '" + myPID + "', '" + theDate + "', 1, '1', '" + theDOB + "');");
                                    }
                                    else
                                    {
                                        ArrINSF.Add(insert1a);
                                        string myPID = thePID;
                                        if (myPID == "") myPID = theKEY;
                                        ArrINSF.Add("    VALUES (1, '" + SubSex + "', '" + myPID + "', '" + theDate + "', 1, '1', '" + theDOB + "');");
                                    }
                                }
                                if (labelMdbExists) NewSubjectsToCreate.Add(theKEY + "^" + theDate + "^" + thePID + "^" + SubSex + "^" + theDOB);    // 4.3.10
                            }
                        }
                        if (!IsDuplicateInsert2s(theKEY + "^" + usedStudyOID + "^" + theDate + "^" + SStheKEY + "^" + thePID))
                        {
                            if (needInserts)
                            {
                                ArrINSF.Add(insert2);
                                // if there is no PID, use the key (unique_identifier) to fill the field label of study_subject. 
                                string myPID = thePID;  // 4.3.23
                                if (myPID == "") myPID = theKEY; // 4.3.23
                                ArrINSF.Add("    VALUES ('" + theKEY + "', (SELECT study_id FROM study WHERE oc_oid = '" + usedStudyOID + "'),");
                                ArrINSF.Add("            1, '" + theDate + "', '" + theDate + "', '" + theDate + "', 1, '" + SStheKEY + "', (SELECT subject_id FROM subject where unique_identifier = '" + myPID + "'));");
                            }
                        }

                        if (SEArrayOnlyE.Count > 1) // in the grid there are more than one Study events referred!!  4.3.18
                        {
                            if (theSTD != "") append_warning("Data for more than one event has been entered. The defined 'Subject Start Date' will be ignored and todays date will be used instead"); 
                            theSTD = ""; // grid allows only one event date; so put it to empty and warn user about this as seen above
                            foreach (string one in SEArray)
                            {
                                string[] parts = one.Split('^');
                                if (parts[1] != "1")
                                {
                                    append_warning("Data for more than one event has been entered for subject. The event " + parts[0] + " is a repeating event, so IF YOU PLAN TO SCHEDULE EVENTS, do this in steps where only one event occurs in the grid at a time. If your events are already scheduled, ignore this warning.");
                                    break;
                                }
                            }
                            foreach (string one in SEArrayOnlyE)
                            {
                                if (!IsDuplicateInsert3s(one + "^" + SStheKEY + "^" + "1" + "^" + theSTD + "^" + theDate))
                                {
                                    if (needInserts)
                                    {
                                        ArrINSF.Add(insert3);
                                        ArrINSF.Add("    VALUES ((SELECT study_event_definition_id FROM study_event_definition WHERE oc_oid = '" + one + "'),");
                                        ArrINSF.Add("	    (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "'),'" + textBoxLocation.Text + "', " + "1" + ", '" + theDate + " 12:00:00', 1, 1, '" + theDate + "', 1, '0', '0');");
                                        ArrINSFR.Add(insert3);
                                        ArrINSFR.Add("    VALUES ((SELECT study_event_definition_id FROM study_event_definition WHERE oc_oid = '" + one + "'),");
                                        ArrINSFR.Add("	     (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "'),'" + textBoxLocation.Text + "', " + "1" + ", '" + theDate + " 12:00:00', 1, 1, '" + theDate + "', 1, '0', '0');");
                                    }
                                    if (labelMdbExists) NewEventsToCreate.Add(theKEY + "^" + one + "^" + textBoxLocation.Text + "^" + theDate);  // 4.3.10
                                }
                            }
                        }
                        else
                        {
                            if (theWrittenSE == "none" && comboBoxSE.SelectedItem.ToString() != "-- select --") // 4.1 eliminate -- select --
                            {
                                theWrittenSE = GetOID(comboBoxSE.SelectedItem.ToString()); // 2.0.5 Use the selected SE to determine current SE, as there is no CRF data 
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
                                        theSTD = ConvertToODMFormat(split[System.Convert.ToInt16(part1)]);
                                        if (theSTD.StartsWith("Error"))
                                        {
                                            string errtext = "Invalid start date '" + theSTD + "' at line " + linecount.ToString() + ". Index: " + STDIndex + ". ";
                                            append_warning(errtext);
                                        }
                                        else
                                        {
                                            if (theSTD != "")
                                            {
                                                if (!IsDuplicateInsert3s(theWrittenSE + "^" + SStheKEY + "^" + say.ToString() + "^" + theSTD + "^" + theDate) && !theSTD.StartsWith("Error"))
                                                {
                                                    if (needInserts)
                                                    {
                                                        ArrINSF.Add(insert3);
                                                        ArrINSF.Add("    VALUES ((SELECT study_event_definition_id FROM study_event_definition WHERE oc_oid = '" + theWrittenSE + "'),");
                                                        ArrINSF.Add("	    (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "'),'" + textBoxLocation.Text + "', " + say.ToString() + ", '" + theSTD + " 12:00:00', 1, 1, '" + theDate + "', 1, '0', '0');");
                                                        ArrINSFR.Add(insert3);
                                                        ArrINSFR.Add("    VALUES ((SELECT study_event_definition_id FROM study_event_definition WHERE oc_oid = '" + theWrittenSE + "'),");
                                                        ArrINSFR.Add("	     (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "'),'" + textBoxLocation.Text + "', " + say.ToString() + ", '" + theSTD + " 12:00:00', 1, 1, '" + theDate + "', 1, '0', '0');");
                                                    }
                                                    if (labelMdbExists) NewEventsToCreate.Add(theKEY + "^" + theWrittenSE + "^" + textBoxLocation.Text + "^" + theSTD);  // 4.3.10
                                                }
                                            }
                                        }
                                    }
                                    else  // no repeating Events; use either todays date as STD or pick it from data file  2.0.9
                                    {
                                        if (!IsDuplicateInsert3s(theWrittenSE + "^" + SStheKEY + "^" + say.ToString() + "^" + theSTD + "^" + theDate) && !theSTD.StartsWith("Error"))
                                        {
                                            if (theSTD != "")
                                            {
                                                if (needInserts)
                                                {
                                                    ArrINSF.Add(insert3);
                                                    ArrINSF.Add("    VALUES ((SELECT study_event_definition_id FROM study_event_definition WHERE oc_oid = '" + theWrittenSE + "'),");
                                                    ArrINSF.Add("	    (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "'),'" + textBoxLocation.Text + "', " + say.ToString() + ", '" + theSTD + " 12:00:00', 1, 1, '" + theDate + "', 1, '0', '0');");
                                                    ArrINSFR.Add(insert3);
                                                    ArrINSFR.Add("    VALUES ((SELECT study_event_definition_id FROM study_event_definition WHERE oc_oid = '" + theWrittenSE + "'),");
                                                    ArrINSFR.Add("        (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "'),'" + textBoxLocation.Text + "', " + say.ToString() + ", '" + theSTD + " 12:00:00', 1, 1, '" + theDate + "', 1, '0', '0');");
                                                }
                                                if (labelMdbExists) NewEventsToCreate.Add(theKEY + "^" + theWrittenSE + "^" + textBoxLocation.Text + "^" + theSTD);  // 4.3.10
                                            }
                                            else  // no date specified in data file
                                            {
                                                if (radioButtonUseTD.Checked)  // 2.1.3 Generate inserts for events without dates using todays date, only if user wants to, otherwise 
                                                {
                                                    if (needInserts)
                                                    {
                                                        ArrINSF.Add(insert3);
                                                        ArrINSF.Add("    VALUES ((SELECT study_event_definition_id FROM study_event_definition WHERE oc_oid = '" + theWrittenSE + "'),");
                                                        ArrINSF.Add("	    (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "'),'" + textBoxLocation.Text + "', " + say.ToString() + ", '" + theDate + " 12:00:00', 1, 1, '" + theDate + "', 1, '0', '0');");
                                                        ArrINSFR.Add(insert3);
                                                        ArrINSFR.Add("    VALUES ((SELECT study_event_definition_id FROM study_event_definition WHERE oc_oid = '" + theWrittenSE + "'),");
                                                        ArrINSFR.Add("        (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "'),'" + textBoxLocation.Text + "', " + say.ToString() + ", '" + theDate + " 12:00:00', 1, 1, '" + theDate + "', 1, '0', '0');");
                                                    }
                                                    if (labelMdbExists) NewEventsToCreate.Add(theKEY + "^" + theWrittenSE + "^" + textBoxLocation.Text + "^" + theDate);     // 4.3.10
                                                }
                                            }
                                        }
                                    }
                                }
                                if (needInserts) ArrDELFR.Add("DELETE FROM study_event where study_subject_id = (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "');");
                            }
                        }
                        if (needInserts)
                        {
                            ArrDELF.Add("DELETE FROM study_event where study_subject_id = (SELECT study_subject_id FROM study_subject WHERE oc_oid = '" + SStheKEY + "');");
                            ArrDELF.Add("DELETE FROM study_subject where oc_oid = '" + SStheKEY + "';");
                            ArrDELF.Add("DELETE FROM subject where unique_identifier = '" + thePID + "';");
                        }
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

            if (labelMdbExists)
            {
                if (DEBUGMODE)
                {
                    using (StreamWriter swlog = new StreamWriter(workdir + "\\OCDataImporter_wsdb.txt"))
                    {
                        swlog.WriteLine("*** Subjects");
                        foreach (string ss in NewSubjectsToCreate) swlog.WriteLine(ss);
                        swlog.WriteLine("");
                        swlog.WriteLine("*** Events");
                        foreach (string ss in NewEventsToCreate) swlog.WriteLine(ss);
                    }
                }
                OleDbConnection con;
                con = new OleDbConnection(mdbConn);
                textBoxOutput.Text += "Now updating WS Database - this may take some time..." + Mynewline;
                textBoxOutput.ScrollToCaret();
                string theCommando = "";
                try
                {
                    con.Open();
                    OleDbCommand del1 = new OleDbCommand("Delete from tblNewEventsToCreate", con);
                    del1.ExecuteNonQuery();
                    OleDbCommand del2 = new OleDbCommand("Delete from tblNewSubjectsToCreate", con);
                    del2.ExecuteNonQuery();
                    // BIO-20-10002_0^2014-07-25^TR-20-10002_0^f^1969-01-01
                    foreach (string one in NewSubjectsToCreate)
                    {
                        string[] prt = one.Split('^');
                        if (prt[4].Trim().Length == 4) theCommando = "insert into tblNewSubjectsToCreate (StudySubjectID, EnrollmentDate, PersonID, Gender, YearOfBirth) values ('" + prt[0] + "', '" + prt[1] + "', '" + prt[2] + "', '" + prt[3] + "', '" + prt[4] + "')"; // 4.3.13
                        else theCommando = "insert into tblNewSubjectsToCreate (StudySubjectID, EnrollmentDate, PersonID, Gender, DateOfBirth) values ('" + prt[0] + "', '" + prt[1] + "', '" + prt[2] + "', '" + prt[3] + "', '" + prt[4] + "')";
                        OleDbCommand ins1 = new OleDbCommand(theCommando, con);
                        ins1.ExecuteNonQuery();
                    }
                    // BIO-20-10002_0^SE_V3^Utrecht^2014-07-25
                    foreach (string one in NewEventsToCreate)
                    {
                        string[] prt = one.Split('^');
                        theCommando = "insert into tblNewEventsToCreate (StudySubjectID, EventOID, Location, StartDate) values ('" + prt[0] + "', '" + prt[1] + "', '" + prt[2] + "', '" + prt[3] + "')";
                        OleDbCommand ins1 = new OleDbCommand(theCommando, con);
                        ins1.ExecuteNonQuery();
                    }
                    con.Close();
                }
                catch (Exception ex)
                {
                    if (ex.Message.ToLower().Contains("duplicate") == false) // 4.3.15 Ignore duplicate; for consequtive events. Dont worry about tblNewEventsToCreate; it has a incr-key so never gets dup error.
                    {
                        MessageBox.Show("Could not update Web service Database because: " + ex.Message + " The command was: " + theCommando, "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        exit_error("Could not update Web service Database because: " + ex.Message + " The command was: " + theCommando);
                    }
                }       
            }
            buttonBackToBegin.BackColor = System.Drawing.Color.LightGreen;
            buttonBackToBegin.Enabled = true;
            buttonStartConversion.Enabled = false;
            buttonStartConversion.BackColor = SystemColors.Control;
            button_start.Enabled = false;
            buttonExit.Enabled = true;
            buttonExit.BackColor = System.Drawing.Color.LightGreen;
            buttonBrowse.Enabled = false;
            buttonCancel.Enabled = false;
            buttonCancel.BackColor = SystemColors.Control;
            textBoxInput.Focus();
            this.Cursor = Cursors.Arrow;
            progressBar1.Value = PROGBARSIZE; 
            using (StreamWriter swlog = new StreamWriter(INSF))
            {
                if (needInserts)
                {
                    swlog.WriteLine("-- " + label1.Text);  // 4.3.24
                    for (int i = 0; i < ArrINSF.Count; i++) swlog.WriteLine(ArrINSF[i]);
                }
            }
            using (StreamWriter swlog = new StreamWriter(INSFR))
            {
                if (needInserts) for (int i = 0; i < ArrINSFR.Count; i++) swlog.WriteLine(ArrINSFR[i]);
            } 
            using (StreamWriter swlog = new StreamWriter(DELF))
            {
                if (needInserts) for (int i = 0; i < ArrDELF.Count; i++) swlog.WriteLine(ArrDELF[i]);
            } 
            using (StreamWriter swlog = new StreamWriter(DELFR))
            {
                if (needInserts) for (int i = 0; i < ArrDELFR.Count; i++) swlog.WriteLine(ArrDELFR[i]);
            }
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
            if (thePIDtoCheck == "") return (false);
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

        public bool IsDuplicateInsert2s(string theStrtoCheck)
        {
            bool found = false;
            foreach (string one in Insert2s)
            {
                if (one == theStrtoCheck)
                {
                    found = true;
                    break;
                }
            }
            if (!found)
            {
                Insert2s.Add(theStrtoCheck);
            }
            return (found);
        }

        public bool IsDuplicateInsert3s(string theStrtoCheck)
        {
            bool found = false;
            foreach (string one in Insert3s)
            {
                if (one == theStrtoCheck)
                {
                    found = true;
                    break;
                }
            }
            if (!found)
            {
                Insert3s.Add(theStrtoCheck);
            }
            return (found);
        }

        public bool IsDuplicateSE(string theStrtoCheck)  // 4.3.18
        {
            bool found = false;
            foreach (string one in SEArray)
            {
                if (one == theStrtoCheck)
                {
                    found = true;
                    break;
                }
            }
            if (!found)
            {
                SEArray.Add(theStrtoCheck);
            }
            return (found);
        }
        public bool IsDuplicateSEOnlyE(string theStrtoCheck)  // 4.3.18
        {
            bool found = false;
            foreach (string one in SEArrayOnlyE)
            {
                if (one == theStrtoCheck)
                {
                    found = true;
                    break;
                }
            }
            if (!found)
            {
                SEArrayOnlyE.Add(theStrtoCheck);
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
            return ("N");  // this means: Cant determine but maybe going to get from row: the check will be done in DoWork
        }

        private void masterClose()
        {
            ArrDIMF.Add("</ClinicalData>");
            ArrDIMF.Add("</ODM>");
            using (StreamWriter swlog = new StreamWriter(DIMF))
            {
                for (int i = 0; i < ArrDIMF.Count; i++) swlog.WriteLine(ArrDIMF[i]);
            }
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

        string GetOID(string val)
        {
            val = val.Replace("    ", "^");
            string[] vals = val.Split('^');
            if (vals.Length == 2) return vals[1].Trim(); // 4.3.24
            else return val;
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
                            string SEOID =  reader.GetAttribute(1) + "    " + reader.GetAttribute(0); // 4.3.1
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
            string myGR = GetOID(comboBoxGR.SelectedItem.ToString());
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
                                            string longIT = GetItemNameFromItemOID(myIT, GetOID(comboBoxCRF.SelectedItem.ToString())) + "    "  + myIT;
                                            comboBoxIT.Items.Add(longIT);
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
            string myCRF = GetOID(comboBoxCRF.SelectedItem.ToString());
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
                                            string longGRP = myGRP;
                                            if (longGRP.Contains("UNGROUPED") == false)
                                            {
                                                foreach (string one in Groups)
                                                {
                                                    if (one.EndsWith(myGRP))
                                                    {
                                                        longGRP = one;
                                                        break;
                                                    }
                                                }
                                            }
                                            comboBoxGR.Items.Add(longGRP);
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
            string mySE = GetOID(comboBoxSE.SelectedItem.ToString());
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
                                            string longCRF = myCRF;
                                            foreach (string one in Forms)
                                            {
                                                if (one.EndsWith(myCRF))
                                                {
                                                    longCRF = one;
                                                    break;
                                                }
                                            }
                                            comboBoxCRF.Items.Add(longCRF);
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
            string myIT = GetOID(comboBoxIT.SelectedItem.ToString());
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

        private bool CheckDay(string day, string month)
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

        private bool CheckYear(string year)
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

        private string ConvertToODMPartial(string theSource)
        {
            if (theSource.Trim() == "") return ("");
            theSource = theSource.Replace('/', '-');
            if (theSource.Contains("-") == false) // probably date consists of only year. 4.3.21 
            {
                if (CheckYear(theSource)) return (theSource);
                else return ("Error: Wrong Year in partial date");
            }
            try
            {
                string[] splitd = theSource.Split('-');
                if (splitd.Length == 3) return (ConvertToODMFormat(theSource)); // full date.  4.3.21
                if (splitd[0].Length == 4 && (splitd[1].Length == 3 || splitd[1].Length == 2))   // yyyy-mmm or yyyy-mm
                {
                    string mon = Get_maand(splitd[1]);
                    if (mon.StartsWith("Error")) return (mon);
                    else
                    {
                        if (CheckYear(splitd[0])) return (splitd[0] + "-" + mon);
                        else return ("Error: Wrong Year in partial date");
                    }
                }
                if (splitd[1].Length == 4 && (splitd[0].Length == 3 || splitd[0].Length == 2))   // mmm-yyyy or mm-yyyy
                {
                    string mon = Get_maand(splitd[0]);
                    if (mon.StartsWith("Error")) return (mon);
                    else
                    {
                        if (CheckYear(splitd[1])) return (splitd[1] + "-" + mon);
                        else return ("Error: Wrong Year in partial date");
                    }
                }
            }
            catch (Exception ex)
            {
                return ("Error-EX1: Unrecognised partial date");
            }
            return ("Error-NR1: Unrecognised partial date");
        }

        private string ConvertToODMFormat(string theSource)
        {
            if (theSource.Trim() == "") return ("");  // 1.0f .Trim = 2.0.4
            if (comboBoxDateFormat.SelectedItem.ToString() == "--select--") return (theSource);
            if (comboBoxDateFormat.SelectedItem.ToString() == "day-month-year")
            {
                theSource = theSource.Replace('/', '-');
                try
                {
                    string[] splitd = theSource.Split('-');
                    string mon = Get_maand(splitd[1]);
                    if (mon.StartsWith("Error")) return (mon);
                    string day = splitd[0];
                    if (day.Length == 1) day = "0" + day;
                    if (CheckYear(splitd[2]) && CheckDay(day, mon)) return (splitd[2] + "-" + mon + "-" + day);
                    else return ("Error: Wrong day or year");
                }
                catch (Exception ex)
                {
                    return ("Error: date is not in day-month-yyyy format ");
                }
            }
            if (comboBoxDateFormat.SelectedItem.ToString() == "month-day-year")
            {
                theSource = theSource.Replace('/', '-');
                try
                {
                    string[] splitd = theSource.Split('-');
                    string mon = Get_maand(splitd[0]);
                    if (mon.StartsWith("Error")) return (mon);
                    string day = splitd[1];
                    if (day.Length == 1) day = "0" + day;
                    if (CheckYear(splitd[2]) && CheckDay(day, mon)) return (splitd[2] + "-" + mon + "-" + day);
                    else return ("Error: Wrong day or year");
                }
                catch (Exception ex)
                {
                    return ("Error: date is not in month-day-yyyy format " + ex);
                }
            }
            if (comboBoxDateFormat.SelectedItem.ToString() == "year-month-day")
            {
                theSource = theSource.Replace('/', '-');
                try
                {
                    string[] splitd = theSource.Split('-');
                    string mon = Get_maand(splitd[1]);
                    if (mon.StartsWith("Error")) return (mon);
                    string day = splitd[2];
                    if (day.Length == 1) day = "0" + day;
                    if (CheckYear(splitd[0]) && CheckDay(day, mon)) return (splitd[0] + "-" + mon + "-" + day);
                    else return ("Error: Wrong day or year");
                }
                catch (Exception ex)
                {
                    return ("Error: date is not in yyyy-month-day format " + ex);
                }
            }
            return ("Error: unknown date format ");
        }
        public string Get_maand(string inp)
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
            if (textBoxInput.Text == "")
            {
                MessageBox.Show("Please enter or select correct input files: Either type file names separated by a semicolon (;) or use browse button. Input-data and OC study metadata files are mandatory.", "OCDataImporter");
                textBoxInput.Focus();
                return;
            }
            WARCOUNT = 0;
            dataGridView1.Rows.Clear();
            string[] split = textBoxInput.Text.Split(';');
            foreach (string one in split)
            {
                if (one.Contains(".xml") || one.Contains(".XML")) input_oc = one;
                if (one.Contains(".txt") || one.Contains(".TXT")) input_data = one;
                if (one.Contains(".oid") || one.Contains(".OID"))
                {
                    input_oid = one;
                    labelOCoidExists = true;
                }
                if (one.Contains(".mdb") || one.Contains(".MDB"))
                {
                    if (zonder)
                    {
                        MessageBox.Show("This version does not support TDS's mdb database. Please download the version from OpenClinica extentions site https://community.openclinica.com/extension/ocdataimporter  if you want to use this database.", "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        textBoxInput.Focus();
                        return;
                    }
                    input_mdb = one;
                    labelMdbExists = true;
                }
            }
            if (input_oc == "" || (input_data == "")) // these are mandatory
            {
                MessageBox.Show("Please enter or select correct input files: Either type file names separated by a semicolon (;) or use browse button. Input-data and OC study metadata files are mandatory.", "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBoxInput.Focus();
                return;
            }
            if (!Get_DataFileItems_FrpmInput(input_data)) return;  // 1.1b
            if (labelOCoidExists) Get_label_oid(input_oid);
            if (labelMdbExists)
            {
                mdbConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + input_mdb;  // start with 12.0 if this doesnt work, try 4.0
                OleDbConnection con;
                con = new OleDbConnection(mdbConn);
                try
                {
                    con.Open();
                    con.Close();
                }
                catch (Exception ex)
                {
                    mdbConn = "";
                }
                if (mdbConn == "")
                {
                    mdbConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + input_mdb;
                    OleDbConnection con1;
                    con1 = new OleDbConnection(mdbConn);
                    try
                    {
                        con1.Open();
                        con1.Close();
                    }
                    catch (Exception ex)
                    {
                        mdbConn = "";
                    }
                }
                if (mdbConn == "")
                {
                    MessageBox.Show("Can't connect MDB database - WS database request will be ignored");
                    labelMdbExists = false;
                }
                else textBoxOutput.Text += "WSDB connection: " + mdbConn + Mynewline;
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

            if (labelMdbExists)
            {
                // 4.3.10: Get precentages "success" in tblNewSubjectsToCreate
                int subs = 0;
                int evts = 0;
                int totsubs = 0;
                int totevts = 0;
                OleDbConnection conn;
                conn = new OleDbConnection(mdbConn);
                try
                {
                    conn.Open();
                    OleDbCommand sel1 = new OleDbCommand("Select Response from tblNewSubjectsToCreate", conn);
                    OleDbDataReader reader = sel1.ExecuteReader();

                    while (reader.Read())
                    {
                        string one = System.Convert.ToString(reader["Response"]);
                        if (one.Contains(">Success<")) subs++;
                        totsubs++;
                    }

                    OleDbCommand sel2 = new OleDbCommand("Select Response from tblNewEventsToCreate", conn);
                    sel2.ExecuteNonQuery();
                    OleDbDataReader reader2 = sel2.ExecuteReader();

                    while (reader2.Read())
                    {
                        string one = System.Convert.ToString(reader2["Response"]);
                        if (one.Contains(">Success<")) evts++;
                        totevts++;
                    }
                    conn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Could not read Web service Database because: " + ex.Message);
                    labelMdbExists = false;
                    return;
                }

                if (totevts > 0 || totsubs > 0) // at least once requests are sent
                {
                    int percsub = 0;
                    int percevt = 0;
                    if (totsubs > 0) percsub = (subs * 100) / totsubs;
                    if (percsub > 100) percsub = 100;
                    if (totevts > 0) percevt = (evts * 100) / totevts;
                    if (percevt > 100) percevt = 100;
                    string messtext = "Requests are already sent and " + percsub.ToString() + "% of StudySubjects and " + percevt.ToString() + "% of Event-schedules are already created" + Mynewline + Mynewline + "Do you want to proceed to CRF Data XML import file generation?" + Mynewline + Mynewline + "OK: Proceed further generating Data Import XML files";
                    DialogResult result = MessageBox.Show(messtext, "OCDataImporter: Action needed", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    if (result == DialogResult.OK)
                    {
                        needInserts = false;
                        //<SOAP-ENV:Envelope xmlns:SOAP-ENV="http://schemas.xmlsoap.org/soap/envelope/"><SOAP-ENV:Header/><SOAP-ENV:Body><scheduleResponse xmlns="http://openclinica.org/ws/event/v1"><result xmlns="http://openclinica.org/ws/event/v1">Success</result><eventDefinitionOID xmlns="http://openclinica.org/ws/event/v1">SE_SE_V3</eventDefinitionOID><studySubjectOID xmlns="http://openclinica.org/ws/event/v1">SS_BIO20100_251</studySubjectOID><studyEventOrdinal xmlns="http://openclinica.org/ws/event/v1">1</studyEventOrdinal></scheduleResponse></SOAP-ENV:Body></SOAP-ENV:Envelope>
                        OleDbConnection connevt;
                        connevt = new OleDbConnection(mdbConn);
                        try
                        {
                            connevt.Open();
                            OleDbCommand sel2 = new OleDbCommand("Select StudySubjectID, Response from tblNewEventsToCreate", connevt);
                            sel2.ExecuteNonQuery();
                            OleDbDataReader reader2 = sel2.ExecuteReader();

                            while (reader2.Read())
                            {
                                string one = System.Convert.ToString(reader2["Response"]);
                                if (one.Contains(">Success<"))
                                {
                                    string first = System.Convert.ToString(reader2["StudySubjectID"]);
                                    string tmp1 = one.Substring(one.IndexOf("studySubjectOID"));
                                    string[] tmp2 = tmp1.Split('>');
                                    string second = tmp2[1].Substring(0, tmp2[1].IndexOf('<'));
                                    if (String.Equals("SS_" + first, second) == false) LabelOID.Add(first + "^" + second);
                                }
                            }
                            connevt.Close();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Could not read Web service Database because: " + ex.Message);
                            return;
                        }
                        if (LabelOID.Count > 0)
                        {
                            labelOCoidExists = true;
                            labelMdbExists = false;
                        }
                    }
                    else if (result == DialogResult.Cancel)
                    {
                        return;
                    }
                }
            }
          
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
                buttonStartConversion.BackColor = System.Drawing.Color.LightGreen;
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
                    string theCRF = GetOID(comboBoxCRF.SelectedItem.ToString());
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
                        fnparts[DGIndexOfOCItem] = GetOID(comboBoxSE.SelectedItem.ToString()) + "." + GetOID(comboBoxCRF.SelectedItem.ToString()) + "." + GetGroupFromItemCRF(theItemOID, theCRF) + "." + theItemOID;
                        matched++;
                    }
                    else
                    {
                        if (theItemOID == "NOTFOUND2" && key != "True" && sex != "True" && pid != "True" && dob != "True" && std != "True" && theItem != "event_index" && theItem != "group_index" && theItem != "site_oid") 
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
                    MessageBox.Show("Not all Items in the selected CRF could be matched. For the list of UNMATCHED Items, see the progress textbox below. You can match those items by using the comboboxes above. Please also control the matched items!", "OCDataImporter");
                }
                else
                {
                    MessageBox.Show("All Items in the selected CRF could be matched. Please check if the matching has been performed correctly!", "OCDataImporter");
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
                buttonExit.BackColor = System.Drawing.Color.LightGreen;
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
                            string theGroup = GetOID(comboBoxGR.SelectedItem.ToString());
                            int strubbish = theGroup.IndexOf("-" + GetOID(comboBoxCRF.SelectedItem.ToString()));
                            if (strubbish > 0) theGroup = theGroup.Substring(0, strubbish);
                            dataGridView1.Rows[e.RowIndex].Cells[DGIndexOfOCItem].Value = GetOID(comboBoxSE.SelectedItem.ToString()) + "." + GetOID(comboBoxCRF.SelectedItem.ToString()) + "." +
                                theGroup + "." + GetOID(comboBoxIT.SelectedItem.ToString());
                        }
                    }
                }
            }
        }
        private void linkLabelBuildDG_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (comboBoxSE.Items.Count > 0)
            {
                if (comboBoxSE.SelectedItem.ToString() != "-- select --" && comboBoxCRF.SelectedItem.ToString() != "-- select --") BuildDG(true);
                else MessageBox.Show("Please select a StudyEvent and a CRF before matching.", "OCDataImporter");
            }
            else MessageBox.Show("Please read input files first.", "OCDataImporter");
        }
        public string CheckRK(string rk, int line)
        {
            if (rk == "N")  // 4.3.17 For repeating groups; Group index = 1 if cant be determined by column (group_index_row = -1) 
            {
                append_warning("Group repeat index can't be determined at line: " + line.ToString() + ". Assumed 1");
                return ("1");
            }
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

        private string GetItemNameFromItemOID(string ItemOID, string FormOID)
        {
            foreach (string idfs in ItemDefForm)
            {
                string[] idf = idfs.Split('^');
                if (idf[0] == ItemOID && idf[1] == FormOID) return (idf[7]);
            }
            return (ItemOID);
        }

        private void BuildVerificationArrays()  // 3.0 Data verification structures 
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
            Sites.Clear();
            Forms.Clear();
            Groups.Clear();
            int linenr = 0;
            DOY = false;
            PIDR = "";

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
                    if (reader.Name == "Study")
                    {
                        // build site array 4.3.1
                        if (reader.AttributeCount > 0) Sites.Add(reader.GetAttribute(0));
                    }
                    if (reader.Name == "FormDef")
                    {
                        // build form array 4.3.1
                        if (reader.AttributeCount > 1) Forms.Add(reader.GetAttribute(1) + "    " + reader.GetAttribute(0));
                    }
                    if (reader.Name == "ItemGroupDef")
                    {
                        // build group array 4.3.1
                        if (reader.AttributeCount > 1) Groups.Add(reader.GetAttribute(1) + "    " + reader.GetAttribute(0));
                    }
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
                    else if (reader.Name == "OpenClinica:StudyParameterListRef")
                    {
                        if (reader.AttributeCount > 1)
                        {
                            at0 = reader.GetAttribute(0);
                            at1 = reader.GetAttribute(1);
                            if (reader.GetAttribute(0) == "SPL_collectDob" && reader.GetAttribute(1) == "2") DOY = true;
                            if (reader.GetAttribute(0) == "SPL_subjectPersonIdRequired") PIDR = reader.GetAttribute(1);  // 4.3.20
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
                    swlog.WriteLine("*** Sites");
                    foreach (string ss in Sites) swlog.WriteLine(ss);
                    swlog.WriteLine("*** Forms");
                    foreach (string ss in Forms) swlog.WriteLine(ss);
                    swlog.WriteLine("*** Groups");
                    foreach (string ss in Groups) swlog.WriteLine(ss);
                    if (DOY) swlog.WriteLine("*** Only year of birth accepted for subjects = true");
                    if (PIDR != "") swlog.WriteLine("*** PersonID = " + PIDR);
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
                ConvertedDate = ConvertToODMFormat(ItemVal);
                
            }
            if (ittype == "partialDate") ConvertedDate = ConvertToODMPartial(ItemVal);
            if (ConvertedDate.StartsWith("Error")) append_warning(fixedwarning + ConvertedDate);
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
            else  // 4.3.6 Lenght and type int/float control must take place only if value not in code list.
            {
                if (ItemVal.Length > itlen) append_warning(fixedwarning + "Item value exceeds required width = " + itlen.ToString());
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
            }
            string huidige = "";
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
                            fItemVal = Single.Parse(ItemVal, CultureInfo.InvariantCulture);  // 4.3.5
                            fRCVal = Single.Parse(RCVal, CultureInfo.InvariantCulture);
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
                append_warning(fixedwarning + "Range check ignored: " + huidige.Replace('^', ','));
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
        private void StateReadFiles(bool begin)
        {
            // initialize program variables
            buttonExit.BackColor = System.Drawing.Color.LightGreen;
            buttonExit.Enabled = true;
            buttonConfPars.Enabled = false;
            buttonConfPars.BackColor = SystemColors.Control;
            if (begin)
            {
                comboBoxDateFormat.SelectedIndex = 0;
                comboBoxSex.SelectedIndex = 0;
            }
            Insert2s.Clear();
            Insert3s.Clear();
            labelOCoidExists = false;
            labelMdbExists = false;
            needInserts = true;
            dmpfilename = "";
            dmpprm = "";
            workdir = "";
            input_oc = "";
            input_data = "";
            input_mdb = "";
            input_oid = "";
            Items.Clear();
            DataFileItems.Clear();
            LabelOID.Clear();
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
            NewSubjectsToCreate.Clear();
            NewEventsToCreate.Clear();
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
            StateReadFiles(false);
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