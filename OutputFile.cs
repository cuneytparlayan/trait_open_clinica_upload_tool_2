using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace OCDataImporter
{
    /// <summary>
    /// Class which takes care of the writing of the output to a file. Ensures that the output file is always written.
    /// </summary>
    class OutputFile
    {
        private String path;

        public OutputFile(String aPath)
        {
            this.path = aPath;
        }

        public void init() {
            FileStream fileStream = new FileStream(path, FileMode.Create, FileAccess.Write); // 1.0f (was deletes.xml - typo error)
            StreamWriter streamWriter = new StreamWriter(fileStream);
            if (fileStream != null)
            {
                fileStream.Close();
            }
        }


        public void Append(string theText)
        {
            // TODO ask Cuneyt if using a FileStream and leaving it open for the entire run isn't faster; this way
            // you always open and close a file handle
            StreamWriter SW;
            SW = File.AppendText(path);
            SW.WriteLine(theText);
            SW.Close();
        }       

    }
}
