namespace Lava3
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnBrowse = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btnOpenFile = new System.Windows.Forms.Button();
            this.btnSortCategories = new System.Windows.Forms.Button();
            this.btnCreditCard = new System.Windows.Forms.Button();
            this.btnCurrentAccount = new System.Windows.Forms.Button();
            this.btnSummary = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtRoot = new System.Windows.Forms.TextBox();
            this.btnRoot = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.cboFiles = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // btnBrowse
            // 
            this.btnBrowse.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBrowse.Location = new System.Drawing.Point(679, 60);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(30, 23);
            this.btnBrowse.TabIndex = 5;
            this.btnBrowse.Text = "...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(19, 63);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(23, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "File";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnOpenFile
            // 
            this.btnOpenFile.Location = new System.Drawing.Point(48, 104);
            this.btnOpenFile.Name = "btnOpenFile";
            this.btnOpenFile.Size = new System.Drawing.Size(119, 23);
            this.btnOpenFile.TabIndex = 6;
            this.btnOpenFile.Text = "Open File";
            this.btnOpenFile.UseVisualStyleBackColor = true;
            this.btnOpenFile.Click += new System.EventHandler(this.btnOpenFile_Click);
            // 
            // btnSortCategories
            // 
            this.btnSortCategories.Location = new System.Drawing.Point(48, 134);
            this.btnSortCategories.Name = "btnSortCategories";
            this.btnSortCategories.Size = new System.Drawing.Size(119, 23);
            this.btnSortCategories.TabIndex = 7;
            this.btnSortCategories.Text = "Sort Categories";
            this.btnSortCategories.UseVisualStyleBackColor = true;
            this.btnSortCategories.Click += new System.EventHandler(this.btnSortCategories_Click);
            // 
            // btnCreditCard
            // 
            this.btnCreditCard.Location = new System.Drawing.Point(48, 174);
            this.btnCreditCard.Name = "btnCreditCard";
            this.btnCreditCard.Size = new System.Drawing.Size(119, 23);
            this.btnCreditCard.TabIndex = 8;
            this.btnCreditCard.Text = "Credit Card";
            this.btnCreditCard.UseVisualStyleBackColor = true;
            this.btnCreditCard.Click += new System.EventHandler(this.btnCreditCard_Click);
            // 
            // btnCurrentAccount
            // 
            this.btnCurrentAccount.Location = new System.Drawing.Point(48, 213);
            this.btnCurrentAccount.Name = "btnCurrentAccount";
            this.btnCurrentAccount.Size = new System.Drawing.Size(119, 23);
            this.btnCurrentAccount.TabIndex = 9;
            this.btnCurrentAccount.Text = "CurrentAccount";
            this.btnCurrentAccount.UseVisualStyleBackColor = true;
            this.btnCurrentAccount.Click += new System.EventHandler(this.btnCurrentAccount_Click);
            // 
            // btnSummary
            // 
            this.btnSummary.Location = new System.Drawing.Point(48, 256);
            this.btnSummary.Name = "btnSummary";
            this.btnSummary.Size = new System.Drawing.Size(119, 23);
            this.btnSummary.TabIndex = 10;
            this.btnSummary.Text = "Summary";
            this.btnSummary.UseVisualStyleBackColor = true;
            this.btnSummary.Click += new System.EventHandler(this.btnSummary_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(18, 18);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(30, 13);
            this.label2.TabIndex = 11;
            this.label2.Text = "Root";
            // 
            // txtRoot
            // 
            this.txtRoot.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtRoot.Enabled = false;
            this.txtRoot.Location = new System.Drawing.Point(48, 18);
            this.txtRoot.Name = "txtRoot";
            this.txtRoot.Size = new System.Drawing.Size(616, 20);
            this.txtRoot.TabIndex = 12;
            // 
            // btnRoot
            // 
            this.btnRoot.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRoot.Location = new System.Drawing.Point(679, 13);
            this.btnRoot.Name = "btnRoot";
            this.btnRoot.Size = new System.Drawing.Size(30, 23);
            this.btnRoot.TabIndex = 13;
            this.btnRoot.Text = "...";
            this.btnRoot.UseVisualStyleBackColor = true;
            this.btnRoot.Click += new System.EventHandler(this.btnRoot_Click);
            // 
            // folderBrowserDialog1
            // 
            this.folderBrowserDialog1.Description = "Choose the root folder";
            // 
            // cboFiles
            // 
            this.cboFiles.FormattingEnabled = true;
            this.cboFiles.Location = new System.Drawing.Point(48, 61);
            this.cboFiles.Name = "cboFiles";
            this.cboFiles.Size = new System.Drawing.Size(616, 21);
            this.cboFiles.TabIndex = 14;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(754, 312);
            this.Controls.Add(this.cboFiles);
            this.Controls.Add(this.btnRoot);
            this.Controls.Add(this.txtRoot);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnSummary);
            this.Controls.Add(this.btnCurrentAccount);
            this.Controls.Add(this.btnCreditCard);
            this.Controls.Add(this.btnSortCategories);
            this.Controls.Add(this.btnOpenFile);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btnOpenFile;
        private System.Windows.Forms.Button btnSortCategories;
        private System.Windows.Forms.Button btnCreditCard;
        private System.Windows.Forms.Button btnCurrentAccount;
        private System.Windows.Forms.Button btnSummary;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtRoot;
        private System.Windows.Forms.Button btnRoot;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.ComboBox cboFiles;
    }
}

