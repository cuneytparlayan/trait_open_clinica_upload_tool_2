using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using System.Security;

namespace OCDataImporter
{
    public class MyFileDialog
    {
        public string fnxml;
        public string fntxt;
        public string fnoid = "";
        public string fnmdb = "";
        private static OpenFileDialog ofd;

        public MyFileDialog(string workdir)
        {
            fnxml = "";
            fntxt = "";
            ofd = new OpenFileDialog();
            if (workdir != "") ofd.InitialDirectory = workdir;
            ofd.Title = "Please select OC MetaFile (.XML), Study Data file (.TXT), label-oid cross data file (.OID) and WSDB file (*.mdb) as input";
            ofd.Filter = "OCDataImporter files (*.XML;*.TXT;*.OID;*.MDB)|*.XML;*.TXT;*.OID;*.MDB|" + "All files (*.*)|*.*";
            ofd.Multiselect = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                foreach (String file in ofd.FileNames)
                {
                    try
                    {
                        if (file.EndsWith(".xml")) fnxml = file;
                        if (file.EndsWith(".txt")) fntxt = file;
                        if (file.EndsWith(".oid")) fnoid = file;
                        if (file.EndsWith(".mdb")) fnmdb = file;
                    }
                    catch (SecurityException ex)
                    {
                        // The user lacks appropriate permissions to read files, discover paths, etc.
                        MessageBox.Show("Security error. Please contact your administrator for details.\n\n" +
                            "Error message: " + ex.Message + "\n\n" +
                            "Details (send to Support):\n\n" + ex.StackTrace);
                    }
                    catch (Exception ex)
                    {
                        // Could not load the image - probably related to Windows file system permissions.
                        MessageBox.Show("Cannot load file: " + file.Substring(file.LastIndexOf('\\'))
                            + ". You may not have permission to read the file, or " +
                            "it may be corrupt.\n\nReported error: " + ex.Message);
                    }
                }
            }
        }
    }
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
