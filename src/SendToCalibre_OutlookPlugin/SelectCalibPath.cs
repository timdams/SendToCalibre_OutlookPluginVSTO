using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SendToCalibre_OutlookPlugin
{
    public partial class SelectCalibPath : Form
    {
        public SelectCalibPath()
        {
            InitializeComponent();
            if (Properties.Settings.Default.calibredbpath != "")
                lblCurrentPath.Text = Properties.Settings.Default.calibredbpath;
        }

        private void btnChangePath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            if (Properties.Settings.Default.calibredbpath != "")
                dlg.SelectedPath = Properties.Settings.Default.calibredbpath;
            if(dlg.ShowDialog()== DialogResult.OK)
            {
                if (File.Exists(Path.Combine(dlg.SelectedPath, "metadata.db")))
                {
                    Properties.Settings.Default.calibredbpath = dlg.SelectedPath;
                    Properties.Settings.Default.Save();
                    lblCurrentPath.Text = dlg.SelectedPath;
                }
                else
                {
                    MessageBox.Show("There doens't appear to be a calibre db there.");
                }
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
