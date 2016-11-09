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
using Lava3.Core;

namespace Lava3
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            txtPath.Text = Properties.Settings.Default.MostRecentFile1;
        }


        private void btnBrowse_Click(object sender, EventArgs e)
        {
            string path = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments);
            txtPath.Text = BrowseForFile(path);
            WriteFileNameToSettings();
        }

        private string BrowseForFile(string path)
        {
            openFileDialog1.Title = "Please choose accounts file";
            openFileDialog1.InitialDirectory = path;
            openFileDialog1.Filter = @"Excel (2007-)|*.xlsx";
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK &&
                File.Exists(openFileDialog1.FileName))
            {
                return openFileDialog1.FileName;
            }
            else
            {
                return path;
            }
        }

        public void WriteFileNameToSettings()
        {
            Properties.Settings.Default.MostRecentFile1 = txtPath.Text;
            Properties.Settings.Default.Save();
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            var excelFile = new ExcelFile();

            excelFile.ShowFile(txtPath.Text);
            this.Enabled = true;
        }

        private void btnSortCategories_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            var excelFile = new ExcelFile();
            excelFile.OpenPackage(txtPath.Text);

            excelFile.LoadAndUpdateCategory();
            excelFile.SaveAndClose();
            this.Enabled = true;
        }

        private void btnCreditCard_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            var excelFile = new ExcelFile();
            excelFile.OpenPackage(txtPath.Text);

            excelFile.LoadAndUpdateCreditCard();

            excelFile.SaveAndClose();
            this.Enabled = true;

        }

        private void btnCurrentAccount_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            var excelFile = new ExcelFile();
            excelFile.OpenPackage(txtPath.Text);

            excelFile.LoadAndUpdateCurrentAccount();

            excelFile.SaveAndClose();
            this.Enabled = true;
        }

        private void btnSummary_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            using (var excelFile = new ExcelFile(txtPath.Text))
            {
                excelFile.LoadAndUpdateAnnualSummary();
                excelFile.SaveAndClose();
            }
            this.Enabled = true;
        }
    }

}
