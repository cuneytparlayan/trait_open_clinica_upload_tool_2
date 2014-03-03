using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OCDataImporter
{
    class InputReader
    {
        private String inputFilePath;
        private ArrayList DataFileItems = new ArrayList();

        /**
         * Constructor
         * @param name="inputFilePath" the full qualified path to the input file
         */
        public InputReader(String inputFilePath)
        {
            this.inputFilePath = inputFilePath;
            read();
        }


        private void read()
        {
            Get_DataFileItems_FromInput();
        }


        private bool Get_DataFileItems_FromInput()
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
    }
}
