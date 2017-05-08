using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;

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
                    item.SaveAsFile(@"c:\temp\" + item.FileName);
                }
            }
        }
    }
}
