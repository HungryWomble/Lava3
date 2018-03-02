namespace Lava3
{
    using System;
    using System.IO;
    using System.Windows.Forms;
    using Lava3.Core;

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            txtRoot.Text = Properties.Settings.Default.RootFolder;
            
            if (string.IsNullOrEmpty(txtRoot.Text) || !Directory.Exists(txtRoot.Text))
            {
                txtRoot.Text = Common.GetDropBoxFolder();            
            }

            if (Properties.Settings.Default.MostRecentFiles != null)
            {
                foreach (var item in Properties.Settings.Default.MostRecentFiles)
                {
                    var fullPath = Path.Combine(txtRoot.Text, item);
                    if (File.Exists(fullPath))
                    {
                        cboFiles.Items.Add(item);
                    }
                }
            }
        }

    
        /// <summary>
        /// Get the file name
        /// </summary>
        /// <returns></returns>
        private string GetFileName()
        {
            string retval = null; 
            if(Common.FileExists(txtRoot.Text, cboFiles.SelectedItem.ToString()))
            {
                retval = Path.Combine(txtRoot.Text, cboFiles.SelectedItem.ToString());
            }
            return retval;
        }
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            BrowseForFile();
            WriteFileNameToSettings();
        }

        private void BrowseForFile()
        {
            string path = System.IO.Path.Combine(txtRoot.Text, cboFiles.Text);
            openFileDialog1.Title = "Please choose accounts file";
            openFileDialog1.InitialDirectory = path;
            openFileDialog1.Filter = @"Excel (2007-)|*.xlsx";
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK &&
                File.Exists(openFileDialog1.FileName))
            {
                string retval = openFileDialog1.FileName.Replace(txtRoot.Text+"\\", "");
                cboFiles.Items.Add(retval);
                cboFiles.SelectedItem = retval;
            }
            
        }

        public void WriteFileNameToSettings()
        {
            Properties.Settings.Default.RootFolder = txtRoot.Text;
            Properties.Settings.Default.MostRecentFiles = new System.Collections.Specialized.StringCollection();
            foreach (var item in cboFiles.Items)
            {
                Properties.Settings.Default.MostRecentFiles.Add(item.ToString());
            }
            Properties.Settings.Default.Save();
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            var excelFile = new ExcelFile();

            excelFile.ShowFile(GetFileName());
            this.Enabled = true;
        }

        private void btnSortCategories_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            var excelFile = new ExcelFile();
            excelFile.OpenPackage(GetFileName());

            excelFile.LoadAndUpdateCategory();
            excelFile.SaveAndClose();
            this.Enabled = true;
        }

        private void btnCreditCard_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            var excelFile = new ExcelFile();
            excelFile.OpenPackage(GetFileName());

            excelFile.LoadAndUpdateCreditCard();

            excelFile.SaveAndClose();
            this.Enabled = true;

        }

        private void btnCurrentAccount_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            var excelFile = new ExcelFile();
            excelFile.OpenPackage(GetFileName());

            excelFile.LoadAndUpdateCurrentAccount();

            excelFile.SaveAndClose();
            this.Enabled = true;
        }

        private void btnSummary_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            using (var excelFile = new ExcelFile(GetFileName()))
            {
                excelFile.LoadAndUpdateAnnualSummary();
                excelFile.SaveAndClose();
            }
            this.Enabled = true;
        }

        private void btnRoot_Click(object sender, EventArgs e)
        {
            // Show the FolderBrowserDialog.
            if(Directory.Exists(txtRoot.Text))
            {
                folderBrowserDialog1.SelectedPath = txtRoot.Text;
            }
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                txtRoot.Text = folderBrowserDialog1.SelectedPath;              
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }

}
