using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Diagnostics;

namespace SendToCalibre_OutlookPlugin
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var m = e.Control.Context as Inspector;
            var mailitem = m.CurrentItem as MailItem;
            if(mailitem!=null)
            {
                foreach (Attachment item in mailitem.Attachments)
                {

                    AddFileToCalibre(item);
                }
            }
        }

        string libpath = "c:\\temp\\cal";
        string command = "add \"{0}\" --library-path {1}";
        private void AddFileToCalibre(Attachment item)
        {
            string filepath = "c:\\temp\\" + item.FileName;
            item.SaveAsFile(filepath);
            string proc = string.Format(command, filepath, libpath);
            MessageBox.Show(proc);
            Process.Start("calibredb",proc);
            
        }
    }
}
