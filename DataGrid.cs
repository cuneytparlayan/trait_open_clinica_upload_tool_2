using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OCDataImporter
{
    /// <summary>
    /// Class for a data grid which contains the column definitions of the data to convert to CDISC ODM
    /// </summary>
    class DataGrid
    {

        // Places in Data Grid
        private const int DGIndexOfDataItem = 0;
        private const int DGIndexOfOCItem = 1;
        private const int DGIndexOfKey = 2;
        private const int DGIndexOfDate = 3;
        private const int DGIndexOfSex = 4;
        private const int DGIndexOfPID = 5;
        private const int DGIndexOfDOB = 6;
        private const int DGIndexOfSTD = 7;

        int keyIndex = 0;
        int sexIndex = 0;
        int PIDIndex = 0;
        int DOBIndex = 0;
        int STDIndex = 0;        
        string sexItem = "";
        string PIDItem = "";
        string DOBItem = "";
        string STDItem = "";

        /// <summary>
        /// Performs checks on the data grid. The results can be examined by the methods:
        /// TODO fill these in
        /// </summary>        
        public void performChecks(DataGridView dataGridView1) {
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
    }
}
