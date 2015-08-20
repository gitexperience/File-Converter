using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Windows.Forms;
namespace FileConverter
{
    class Program
    {
        [STAThread]
        static void Main()
        {

            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
           
            object oMissing = System.Reflection.Missing.Value;
            string folder="";
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                folder = fbd.SelectedPath;
            }
            else
                Environment.Exit(0);
           
            DirectoryInfo dirInfo = new DirectoryInfo(folder);
            FileInfo[] wordFiles = dirInfo.GetFiles("*.docx");
            if (wordFiles.Length == 0)
            {
                MessageBox.Show("No Word Files Found in Folder!!", "Message");
                Environment.Exit(0);
            }
            word.Visible = false;
            word.ScreenUpdating = false;

            foreach (FileInfo wordFile in wordFiles)
            {
                Object filename = (Object)wordFile.FullName;

                Document doc = word.Documents.Open(ref filename, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                doc.Activate();

                object outputFileName = wordFile.FullName.Replace(".docx", ".pdf");
                object fileFormat = WdSaveFormat.wdFormatPDF;

                doc.SaveAs(ref outputFileName,
                    ref fileFormat, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                             
                object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
                doc = null;
            }

            ((_Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
            word = null;
            MessageBox.Show("File Successfully Converted!!","Success!!");
        }
    }
}    
      