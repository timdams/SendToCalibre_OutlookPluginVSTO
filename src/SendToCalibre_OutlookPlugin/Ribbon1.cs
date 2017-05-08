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
            if (mailitem != null)
            {
                if (mailitem.Attachments.Count > 0)
                {
                    foreach (Attachment item in mailitem.Attachments)
                    {
                        //TODO only epub!

                        AddFileToCalibre(item);
                    }
                }
                else
                {
                    MessageBox.Show("This email doesn't have any attachments...", "Whoopsie", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }




        string command = "add \"{0}\" --library-path {1}";
        private void AddFileToCalibre(Attachment item)
        {
            string libpath = Properties.Settings.Default.calibredbpath;
            if (libpath == "")
            {
                if(MessageBox.Show("No calibre db chosen! Go to preferences.", "Whoopsie", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning )==DialogResult.Yes)
                {
                    btnSettings_Click(this, null);
                }
            }
            else
            {
                string temppath = Path.GetTempPath();
                string filepath = Path.Combine(temppath, item.FileName);

                item.SaveAsFile(filepath);
                string procargs = string.Format(command, filepath, libpath);

                var proc = Process.Start("calibredb", procargs);
                proc.WaitForExit();

                File.Delete(filepath);
            }

        }

        private void btnSettings_Click(object sender, RibbonControlEventArgs e)
        {
            SelectCalibPath win = new SelectCalibPath();
            win.ShowDialog();
        }
    }
}
