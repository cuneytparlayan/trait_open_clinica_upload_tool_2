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
/******************************************************************************************
1.0	        21/05/2010      Initial version	                    C. Parlayan, J.A.M. Beliën
1.1	        21/04/2012	    Production version	                C. Parlayan
2.0         08/11/2012      Production 2.0                      C. Parlayan
2.0.1       01/02/2013      Added Repeating events and groups   C. Parlayan
2.0.2	    26/03/2013      The GRID now keeps the previously 
                            matched items when a new CRF or 
                            Group is chosen for more matching.  C. Parlayan
2.0.3	    01/04/2013	    Bug solved in matching with 
                            repeated items.	                    C. Parlayan
2.0.4	    01/05/2013	    Remove white space when matching	C. Parlayan
2.0.5	    21/07/2013	    Do not print form data with no 
                            items in XML as this causes error 
                            in OpenClinica upload.	            C. Parlayan, S. de Ridder
2.0.6	    26/08/2013	    Added “Limit number of characters 
                            to match” to make matching easier	C. Parlayan
2.0.7/8/9	29/08/2013	    Bugs introduced in 2.0.5 solved.	C. Parlayan, S. de Ridder
2.1.1	    09/09/2013	    Introduced label-oid file.	        C. Parlayan, J. Rousseau, R. Voorham
2.1.2	    01/11/2013	    Bug ItemGroupRepeatKey solved.	    C. Parlayan, R. Voorham
2.1.3	    18/11/2013	    Possibility to not generate 
                            events if startdate is blank.	    C. Parlayan, S. de Ridder
2.1.4	    01/11/2013	    XML escaping.	                    C. Parlayan, J. Rousseau
2.2         30/11/2013      Input file allows EVENT_INDEX 
                            and GROUP_INDEX to be defined 
                            to accept repeating events/items
                            in rows                             C. Parlayan, J. Rousseau
3.0       20/12/2013        Type and range validations,
                            better messaging                    C. Parlayan
3.01      20-01-2014        XML parsing error while RangeCheck  J. Rousseau, C. Parlayan
3.02      20-01-2014        Warnings to status window           J. Rousseau, C. Parlayan
3.03      22-01-2014        matching now uses "ItemName"  
                            instead of ItemOID                  C. Parlayan, S. de Ridder
3.03      24-01-2014        matching is done for all groups  
                            in a CRF in one go                  C. Parlayan, S. de Ridder
4.0       29-01-2014        Better UI                           C. Parlayan
4.01      07-02-2014        Show/hide Subj related columns
                            bugfixes on validation              C. Parlayan, J. Rousseau, S. de Ridder
4.1       10-02-2014        Auto detection and show/hide         
                            subject related columns             C. Parlayan, J. Rousseau, S. de Ridder
4.1       10-02-2014        Added Multiple select to         
                            validations                         C. Parlayan, J. Rousseau
4.1       10-02-2014        Leading zeroes in times and           
                            subjectid's in log records          C. Parlayan, J. Rousseau
4.1       10-02-2014        fix: "-- select --" in insert            
                            statements                          C. Parlayan, S. de Ridder
4.2       17-02-2014        Show/hidden warn ONLY if show and 
                            has no value                        C. Parlayan, S. de Ridder
4.2.1     19-02-2014        Ignore range check if other 
                            problems detected with data         C. Parlayan, S. de Ridder
4.3       21-02-2014        SE/GR repeating index in rows 
                            bugs fixed                          C. Parlayan, J. Rousseau
4.3       21-02-2014        Better error messages if input
                            file is open by other programs      C. Parlayan, J. Rousseau
4.4       11-03-2014        Decoupeling of the input 
                            parameters from the view, 
                            introduction of a utilities
                            class, a OutputFile class to 
                            stream output to, and a main 
                            converter which is the controler 
                            of the conversion (MVC), various
                            fixes.
                            Project is now available on github 
                            for change history.                 C. Parlayan, J. Rousseau
4.3.1     02-04-2014      Replace couples textbox
                          deleted; not needed                 C. Parlayan, J. Rousseau
4.3.1     03-04-2014      Allow date of year
                          instead of date of birth            C. Parlayan
 4.3.1    03-04-2014      subject_event_status_id from 3 
                          to 1 to trigger audit               C. Parlayan, S. de Ridder
4.3.1     03-04-2014      SITE_OID column introduced to 
                          allow defining site for subjects    C. Parlayan, J. Rousseau
4.3.1     04-04-2014      Names in dropdowns's beside  
                          OID's                               C. Parlayan, S. de Ridder
4.3.2     14-04-2014      Fix buttons enable/disable status
                          and colors                          C. Parlayan
4.3.3     14-04-2014      Dont change date format when back
                          to begin button is hit              C. Parlayan, S. de Ridder
4.3.3     14-04-2014      Added "-01-01" after year of 
                          birth                               C. Parlayan, S. de Ridder
4.4.2     01-05-2014      EVENT_INDEX and GROUP_INDEX bug
                          resolved                            C. Parlayan
4.4.2     01-05-2014      Ignore input line if it is completely
                          empty                               C. Parlayan
                            
*******************************************************************************************/
namespace OCDataImporter
{    
    public partial class Form1 : Form, IViewUpdater
    {
        
        public const String VERSION_LABEL = "OCDataImporter Version 4.4.2";   
        
        // public const bool DEBUGMODE = true;
        public const bool DEBUGMODE = false;
        
        public bool labelOCoidExists = false; // 2.1.1 If labelOCoid file exists, get the oid from that file, instead of 'SS_label'...
        public string dmpfilename = "";
        public string dmpprm = "";
        Thread MyThread;
        FileStream fpipdf;
        string Mynewline = System.Environment.NewLine;                
        static public string input_oid = "";
        ArrayList Items = new ArrayList();
        ArrayList DataFileItems = new ArrayList();
        ArrayList LabelOID = new ArrayList();
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
        ArrayList InsertKeys = new ArrayList();
               
        
        char Delimiter = ';';
        char tab = '\u0009';
        int sepcount = 1;
        int PROGBARSIZE = 0;

        
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
        
        
        string selectedEventRepeating = "No";
        
        static public string insert1a = "INSERT INTO subject(status_id, gender, unique_identifier, date_created, owner_id, dob_collected, date_of_birth)";
        static public string insert1 = "INSERT INTO subject(status_id, gender, unique_identifier, date_created, owner_id, dob_collected)";
        static public string insert2 = "INSERT INTO study_subject(label, study_id, status_id, enrollment_date, date_created, date_updated, owner_id,  oc_oid, subject_id)";
        static public string insert3 = "INSERT INTO study_event(study_event_definition_id, study_subject_id, location, sample_ordinal, date_start, owner_id, status_id, date_created, subject_event_status_id, start_time_flag, end_time_flag)";

        
        private DataGrid dataGrid;
        private StudyMetaDataValidator studyMetaDataValidator;
        private ConversionSettings conversionSettings;
        private WarningLog warningLog;
        
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
            label1.Text = OCDataImporter.Form1.VERSION_LABEL;
            conversionSettings = new ConversionSettings();
            conversionSettings.Sites = new ArrayList();
            conversionSettings.Forms = new ArrayList();
            conversionSettings.Groups = new ArrayList(); 
            warningLog = new WarningLog();
            studyMetaDataValidator = new StudyMetaDataValidator(warningLog, conversionSettings);
            dataGrid = new DataGrid(studyMetaDataValidator, dataGridView1, this);

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
            try
            {
                fpipdf = new FileStream("OCDataImporter.pdf", FileMode.Open, FileAccess.Read);
            }
            catch (Exception exx)
            {
                MessageBox.Show("Problem opening user manual. Message = " + exx.Message, "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            conversionSettings.workdir = Directory.GetCurrentDirectory();
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
            MyFileDialog mfd = new MyFileDialog(conversionSettings.workdir);
            if (mfd.fntxt == "" || mfd.fnxml == "") return;
            textBoxInput.Text = mfd.fnxml + ";" + mfd.fntxt + ";" + mfd.fnoid;
            if (mfd.fnoid != "") labelOCoidExists = true;
        }

        void MenuFileOpenOnClick(object obj, EventArgs ea)
        {
            Get_work_file();
        }


        private void safeClose(Stream stream)
        {
            if (stream != null)
            {
                stream.Close();
            }
        }

        void MenuFileExitOnClick(object obj, EventArgs ea)
        {
            safeClose(fpipdf);            
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
            MessageBox.Show(VERSION_LABEL + " - Made by: C. Parlayan/J. Rousseau, VU Medical Center, Amsterdam, The Netherlands - 2010-2014", Text);
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
                            first = false;
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
                    outPR.WriteLine(comboBoxDateFormat.SelectedItem + "~" + comboBoxSex.SelectedItem + "~" + textBoxSubjectSexM.Text + "~" + textBoxSubjectSexF.Text + "~" + textBoxMaxLines.Text + "~" + textBoxLocation.Text + "~");
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

        /**
         * 4 methodes defined in IViewUpdater
         * 
         */
        public void updateProgressbarStep(int step)
        {
            progressBar1.Step = step;
        }


        public void performProgressbarStep()
        {
            progressBar1.PerformStep();
        }


        public void appendText(String aMessage)
        {
            textBoxOutput.Text += aMessage;
        }


        public void resetText()
        {
            textBoxOutput.Text = "";
        }

        /// <summary>
        /// Initializes the user interface elements before the conversion
        /// </summary>
        private void initializeUserInterfaceElements()
        {
            if (conversionSettings.pathToInputFile.Contains("\\") == false)
            {
                conversionSettings.pathToInputFile = conversionSettings.workdir + "\\" + conversionSettings.pathToInputFile;
            }
            if (Delimiter != tab) textBoxOutput.Text += "\r\nData file is: " + conversionSettings.pathToInputFile + ", delimited by: " + Delimiter + " Number of items per line: " + dataGrid.sepcount + "\r\n";
            else textBoxOutput.Text += "\r\nData file is: " + conversionSettings.pathToInputFile + ", delimited by: tab, Number of items per line: " + dataGrid.sepcount + "\r\n";
            textBoxOutput.Text += "Started in directory " + conversionSettings.workdir + ". This may take several minutes...\r\n";

            buttonBackToBegin.Enabled = false;
            buttonBackToBegin.BackColor = SystemColors.Control;
            buttonStartConversion.Enabled = false;
            buttonStartConversion.BackColor = SystemColors.Control;
            buttonExit.Enabled = false;
            buttonExit.BackColor = SystemColors.Control;
            buttonCancel.Enabled = true;
            buttonCancel.BackColor = System.Drawing.Color.LightGreen;
            linkLabelBuildDG.Enabled = false;
            linkbuttonSHCols.Enabled = false;
            linkLabel1.Enabled = false;
            progressBar1.Value = 0;
            this.Cursor = Cursors.AppStarting;
        }

        /// <summary>
        /// Resets the user interface elements after the conversion and update the fields with the information
        /// of the run which just finished.
        /// </summary>
        private void resetUserInterfaceElements(Converter converter)
        {
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
            int numberOfWarnings = warningLog.getWarningCount();


            textBoxOutput.Text += warningLog.ToString();

            labelWarningCounter.Text = "WARNINGS: " + numberOfWarnings.ToString();
            if (numberOfWarnings == 0)
            {                
                warningLog.appendMessage(DateTime.Now + " Finished successfully.");
                MessageBox.Show("Process finished successfully", "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Process ended with errors or warnings: See OCDataImporter_log.txt and/or OCDataImporter_warning.txt for details", "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBoxOutput.Text += "*** There are errors or warnings: See OCDataImporter_log.txt and/or OCDataImporter_warning.txt for details ***";
            }
            if (converter.numberOfOutputFiles > 1)
            {
                textBoxOutput.Text += " Total: " + converter.numberOfOutputFiles.ToString() + " ODM files.";
            }
            textBoxOutput.SelectionStart = textBoxOutput.Text.Length;
            textBoxOutput.ScrollToCaret();
        }       

        private void buttonStartConversion_Click_1(object sender, EventArgs e)
        {

            if (Utilities.IsNumber(textBoxMaxLines.Text) == false)
            {
                MessageBox.Show("Split factor must be a number, 0 means no splitting.", "OCDataImporter");
                return;
            }
            BuildRepeatingEventString();
            BuildRepeatingGroupString();
            initializeUserInterfaceElements();

            // delete DataImport*.xml
            string[] txtList = Directory.GetFiles(conversionSettings.workdir, "DataImport_*.xml");
            if (txtList.Length > 0)
            {
                if (MessageBox.Show("DataImport_* files will be overwritten. Do you want to delete the old files?", "Confirm delete old DataImport*.xml files", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
                {
                    foreach (string f in txtList) File.Delete(f);
                }
            }

            
            conversionSettings.checkForDuplicateSubjects = checkBoxDup.Checked;
            conversionSettings.useTodaysDateIfNoEventDate = radioButtonUseTD.Checked;
            conversionSettings.dateFormat = comboBoxDateFormat.SelectedItem.ToString();
            conversionSettings.defaultLocation = textBoxLocation.Text;
            conversionSettings.defaultSubjectSex = comboBoxSex.SelectedItem.ToString();
            conversionSettings.outFMaxLines = System.Convert.ToInt32(textBoxMaxLines.Text);
            conversionSettings.selectedStudyEvent = comboBoxSE.SelectedItem.ToString();

            dumpTheGrid();
            Converter converter = new Converter(conversionSettings, studyMetaDataValidator, warningLog, this, labelOCoidExists, LabelOID);
            try
            {
                converter.DoWork(dataGrid);
            }
            catch (OCDataImporterException ocdie) {
                    MessageBox.Show(ocdie.ToString(), "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    exit_error(ocdie.ToString());
            }
            String pathWarningLog = conversionSettings.workdir + "\\OCDataImporter_warnings.txt";
            warningLog.dumpToFile(pathWarningLog);
            resetUserInterfaceElements(converter);
        }

        public string getTodaysDate()
        {
            DateTime dt = DateTime.Now;
            return dt.ToString("yyyy-MM-dd");
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
                            string SEOID = reader.GetAttribute(1) + "    " + reader.GetAttribute(0); // 4.3.1
                            if (DEBUGMODE)
                            {
                                textBoxOutput.Text += reader.GetAttribute(0) + ", Name = " + reader.GetAttribute(1) + ", Repeating = " + reader.GetAttribute(2) + ", Type = " + reader.GetAttribute(3) + Mynewline;
                            }
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
            String repeating_groups = "";
            try
            {
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.ConformanceLevel = ConformanceLevel.Fragment;
                settings.IgnoreWhitespace = true;
                settings.IgnoreComments = true;
                XmlReader reader = XmlReader.Create(conversionSettings.pathToMetaDataFile, settings);
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
            studyMetaDataValidator.repeatingGroups = repeating_groups;
        }

        private void BuildRepeatingEventString()
        {
            String repeating_events = "";
            try
            {
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.ConformanceLevel = ConformanceLevel.Fragment;
                settings.IgnoreWhitespace = true;
                settings.IgnoreComments = true;
                XmlReader reader = XmlReader.Create(conversionSettings.pathToMetaDataFile, settings);
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
            studyMetaDataValidator.repeatingEvents = repeating_events;
        }

        private void comboBoxGR_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBoxIT.Items.Clear();
            Items.Clear();
            string myGR = Utilities.GetOID(comboBoxGR.SelectedItem.ToString());
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
                XmlReader reader = XmlReader.Create(conversionSettings.pathToMetaDataFile, settings);

                if (DEBUGMODE)
                {
                    textBoxOutput.Text += "Selected Group = " + myGR;
                }
                while (reader.Read())
                {
                    if (reader.Name == "ItemGroupDef")
                    {
                        if (reader.AttributeCount > 0)
                        {
                            string SEOID = reader.GetAttribute(0);
                            if (SEOID == myGR)
                            {
                                if (DEBUGMODE)
                                {
                                    textBoxOutput.Text += " Repeating = " + reader.GetAttribute(2) + Mynewline;
                                }
                                string myGRPS = reader.ReadInnerXml();
                                if (myGRPS.Contains("<ItemRef "))
                                {
                                    myGRPS = myGRPS.Replace("<ItemRef ", "");
                                    myGRPS = myGRPS.Replace("/>", "~");
                                    if (DEBUGMODE)
                                    {
                                        textBoxOutput.Text += "Items for the selected Group = " + Mynewline;
                                    }
                                    foreach (string ss in myGRPS.Split('~'))
                                    {
                                        ss.Trim();
                                        if (ss != "" && ss.Contains("ItemOID"))
                                        {
                                            string myIT = "";
                                            if (DEBUGMODE) { 
                                                textBoxOutput.Text += ss + Mynewline; 
                                            }

                                            myIT = ss.Substring(9);
                                            int stoppunt = myIT.IndexOf('"');
                                            myIT = myIT.Substring(0, stoppunt);
                                            string longIT = studyMetaDataValidator.GetItemNameFromItemOID(myIT, Utilities.GetOID(comboBoxCRF.SelectedItem.ToString())) + "    " + myIT;
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
            string myCRF = Utilities.GetOID(comboBoxCRF.SelectedItem.ToString());
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
                XmlReader reader = XmlReader.Create(conversionSettings.pathToMetaDataFile, settings);
                if (DEBUGMODE) { 
                    textBoxOutput.Text += "Selected CRF = " + myCRF + Mynewline; 
                }

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
                                    if (DEBUGMODE)
                                    {
                                        textBoxOutput.Text += "Groups for the selected CRF = " + Mynewline;
                                    }
                                    foreach (string ss in myGRPS.Split('~'))
                                    {
                                        ss.Trim();
                                        if (ss != "" && ss.Contains("ItemGroupOID"))
                                        {
                                            string myGRP = "";
                                            if (DEBUGMODE)
                                            {
                                                textBoxOutput.Text += ss + Mynewline;
                                            }
                                            myGRP = ss.Substring(14);
                                            int stoppunt = myGRP.IndexOf('"');
                                            myGRP = myGRP.Substring(0, stoppunt);
                                            // comboBoxGR.Items.Add(myGRP);
                                            string longGRP = myGRP;
                                            if (longGRP.Contains("UNGROUPED") == false)
                                            {
                                                foreach (string one in conversionSettings.Groups)
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
            string mySE = Utilities.GetOID(comboBoxSE.SelectedItem.ToString());
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
                XmlReader reader = XmlReader.Create(conversionSettings.pathToMetaDataFile, settings);
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
                                            foreach (string one in conversionSettings.Forms)
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
            string myIT = Utilities.GetOID(comboBoxIT.SelectedItem.ToString());
            if (myIT == "-- select --") return;
            try
            {
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.ConformanceLevel = ConformanceLevel.Fragment;
                settings.IgnoreWhitespace = true;
                settings.IgnoreComments = true;
                XmlReader reader = XmlReader.Create(conversionSettings.pathToMetaDataFile, settings);
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
        
        /*
        static void AppendToFile(string theFile, string theText)
        {
            StreamWriter SW;
            SW = File.AppendText(theFile);
            SW.WriteLine(theText);
            SW.Close();
        }       
        */


        /*
        private string ConvertToODMPartial(string theSource)
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
        */
          
        private void exit_error(string probl)
        {   
            using (StreamWriter swlog = new StreamWriter(conversionSettings.workdir + "\\OCDataImporter_log.txt"))
            {
                swlog.WriteLine(probl);
            }
            Close();
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
                conversionSettings.pathToMetaDataFile = split[0];
                conversionSettings.pathToInputFile = split[1];
            }
            else
            {
                conversionSettings.pathToMetaDataFile = split[1];
                conversionSettings.pathToInputFile = split[0];
            }
            if (!dataGrid.GetDataFileItemsFromInput(conversionSettings.pathToInputFile)) return;  // 1.1b
            if (labelOCoidExists)
            {
                input_oid = split[2];
                Get_label_oid(input_oid);
            }

            dmpfilename = conversionSettings.pathToInputFile.Substring(0, conversionSettings.pathToInputFile.Length - 4) + "_grid.dmp";
            dmpprm = conversionSettings.pathToInputFile.Substring(0, conversionSettings.pathToInputFile.Length - 4) + "_parameters.dmp";
            try
            {
                FileStream fp_ioc;
                long FileSize;
                fp_ioc = new FileStream(split[1], FileMode.Open);
                FileSize = fp_ioc.Length;
                safeClose(fp_ioc);
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

            conversionSettings.workdir = "";
            if (conversionSettings.pathToMetaDataFile.Contains("\\"))
            {
                conversionSettings.workdir = conversionSettings.pathToMetaDataFile.Substring(0, conversionSettings.pathToMetaDataFile.LastIndexOf('\\'));
            }
            else
            {
                conversionSettings.workdir = Directory.GetCurrentDirectory();
            }

            comboBoxSE.Items.Clear();
            comboBoxCRF.Items.Clear();
            comboBoxGR.Items.Clear();
            comboBoxIT.Items.Clear();
            Items.Clear();

            XmlReader reader = null;
            try
            {
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.ConformanceLevel = ConformanceLevel.Fragment;
                settings.IgnoreWhitespace = true;
                settings.IgnoreComments = true;
                reader = XmlReader.Create(conversionSettings.pathToMetaDataFile, settings);

                while (reader.Read())
                {
                    if (reader.Name == "Study")
                    {
                        conversionSettings.studyOID = reader.GetAttribute(0);
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can't open selected OC meta file, can't continue. Problem is: " + System.Environment.NewLine + ex.Message, "OCDataImporter", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                exit_error(ex.ToString());
                return;
            }
            finally
            {
                if (reader != null)
                {
                    reader.Close();
                }
            }


            GetStudyEventDef(conversionSettings.pathToMetaDataFile);
 
            if (comboBoxSE.Items.Count == 0) 
            {
                MessageBox.Show("No study event and CRF definitions found in selected file; please check the format of the file and verify if the metadata file corresponds to the data file", "OCDataImporter");
                textBoxInput.Focus();
                buttonStartConversion.Enabled = true;
                buttonStartConversion.BackColor = System.Drawing.Color.LightGreen;
                return;
            }


            if (FillTheGrid() == false)
            {
                conversionSettings.selectedStudyEvent = comboBoxSE.SelectedItem.ToString();
                if (comboBoxCRF.SelectedItem == null)
                {
                    conversionSettings.selectedCRF = "-- select --";
                }
                else
                {
                    conversionSettings.selectedCRF = comboBoxCRF.SelectedItem.ToString();
                }
                dataGrid.BuildDG(false, conversionSettings.selectedStudyEvent, conversionSettings.selectedCRF);
            }
            studyMetaDataValidator.BuildVerificationArrays(conversionSettings.pathToMetaDataFile, conversionSettings.workdir + "\\OCDataImporter_verification.txt", DEBUGMODE);
            StateParametres();
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
                            string theGroup = Utilities.GetOID(comboBoxGR.SelectedItem.ToString());
                            int strubbish = theGroup.IndexOf("-" + Utilities.GetOID(comboBoxCRF.SelectedItem.ToString()));
                            if (strubbish > 0) theGroup = theGroup.Substring(0, strubbish);
                            dataGridView1.Rows[e.RowIndex].Cells[DGIndexOfOCItem].Value = Utilities.GetOID(comboBoxSE.SelectedItem.ToString()) + "." + Utilities.GetOID(comboBoxCRF.SelectedItem.ToString()) + "." +
                                theGroup + "." + Utilities.GetOID(comboBoxIT.SelectedItem.ToString());
                        }
                    }
                }
            }
        }
        private void linkLabelBuildDG_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (comboBoxSE.Items.Count > 0)
            {
                if (comboBoxSE.SelectedItem.ToString() != "-- select --")
                {
                    conversionSettings.selectedStudyEvent = comboBoxSE.SelectedItem.ToString();
                    conversionSettings.selectedCRF = comboBoxCRF.SelectedItem.ToString();
                    dataGrid.BuildDG(true, conversionSettings.selectedStudyEvent, conversionSettings.selectedCRF);
                }
                else MessageBox.Show("Please select a StudyEvent and a CRF before matching.", "OCDataImporter");
            }
            else MessageBox.Show("Please read input files first.", "OCDataImporter");
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
            labelOCoidExists = false;
            dmpfilename = "";
            dmpprm = "";
            
            conversionSettings.reset();
            
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
            warningLog.reset();
            
            InsertKeys.Clear();
            
            Delimiter = tab;
            sepcount = 1;
            PROGBARSIZE = 0;
            
            
            linelen = 0;

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
            conversionSettings.SUBJECTSEX_M = textBoxSubjectSexM.Text;
            conversionSettings.SUBJECTSEX_F = textBoxSubjectSexF.Text;
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