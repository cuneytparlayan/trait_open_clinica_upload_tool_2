using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.IO;

namespace OCDataImporter
{
    class InputReader
    {
        public int sepcount { get; private set; }
        public char delimiter {get; private set;}
        public ArrayList dataFileItems { get; private set;}
        
        
        private String inputFilePath;
                
        private const char tab = '\u0009';        
        
        /**
         * Constructor
         * @param name="inputFilePath" the full qualified path to the input file
         */
        public InputReader(String inputFilePath)
        {
            this.inputFilePath = inputFilePath;
            delimiter = ';';
            dataFileItems = new ArrayList();
        }
        
        /// <summary>
        /// Resets the input reader
        /// </summary>
        public void reset() {
            dataFileItems.Clear();
        }

        /// <summary>
        /// Reads the input file. The lines can be retrieved by the method
        /// </summary>
        public void Get_DataFileItems_FromInput() 
        {
            // Find out how many data items are present per line and build array of data item names for using in data grid
            dataFileItems.Clear();            
            int linelen = 0;            
            using (StreamReader sr = new StreamReader(inputFilePath))
            {
                String line;
                while ((line = sr.ReadLine()) != null)
                {
                    line = line.Trim();  // 1.1b
                    if (line.Length == 0) continue;
                    linelen = line.Length;
                    if (line.IndexOf(tab) > 0) delimiter = tab;
                    if (line.IndexOf(';') > 0) delimiter = ';';

                    for (int i = 0; i < line.Length; i++) if (line[i] == delimiter) sepcount++;
                    string[] spfirst = line.Split(delimiter);
                    foreach (string one in spfirst) dataFileItems.Add(one);
                    break;
                }
            }                        
        }


        /// <summary>
        /// Creates a string with information on the input file with the delimiter and file path
        /// </summary>
        /// <param name="workdir">the work directory in which the input file is found</param>
        /// <returns>a non-empty string</returns>
        public String makeInformationString(String workdir)
        {
            String ret = "";
            if (delimiter != tab)
            {
                ret += "\r\nData file is: " + inputFilePath + ", delimited by: " + delimiter + " Number of items per line: " + sepcount + "\r\n";
            }
            else
            {
                ret += "\r\nData file is: " + inputFilePath + ", delimited by: tab, Number of items per line: " + sepcount + "\r\n";
            }            
            ret += "Started in directory " + workdir + ". This may take several minutes...\r\n";
            return ret;
        }
    }
}
