
namespace Office_File_Explorer.WinForms
{
    partial class FrmBatch
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmBatch));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rdoPowerPoint = new System.Windows.Forms.RadioButton();
            this.rdoExcel = new System.Windows.Forms.RadioButton();
            this.rdoWord = new System.Windows.Forms.RadioButton();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.BtnBrowseFolder = new System.Windows.Forms.Button();
            this.tbFolderPath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.BtnCopyOutput = new System.Windows.Forms.Button();
            this.ckbSubfolders = new System.Windows.Forms.CheckBox();
            this.lstOutput = new System.Windows.Forms.ListBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.BtnResetBulletMargins = new System.Windows.Forms.Button();
            this.BtnRemoveCustomTitle = new System.Windows.Forms.Button();
            this.BtnFixComments = new System.Windows.Forms.Button();
            this.BtnRemovePII = new System.Windows.Forms.Button();
            this.BtnFixTableProps = new System.Windows.Forms.Button();
            this.BtnChangeTheme = new System.Windows.Forms.Button();
            this.BtnDeleteRequestStatus = new System.Windows.Forms.Button();
            this.BtnFixRevisions = new System.Windows.Forms.Button();
            this.BtnFixBookmarks = new System.Windows.Forms.Button();
            this.BtnFixNotesPage = new System.Windows.Forms.Button();
            this.BtnRemovePIIOnSave = new System.Windows.Forms.Button();
            this.BtnConvertStrict = new System.Windows.Forms.Button();
            this.BtnUpdateNamespaces = new System.Windows.Forms.Button();
            this.BtnSetOpenByDefault = new System.Windows.Forms.Button();
            this.BtnChangeTemplate = new System.Windows.Forms.Button();
            this.BtnFixHyperlinks = new System.Windows.Forms.Button();
            this.BtnDeleteCustomProps = new System.Windows.Forms.Button();
            this.BtnAddCustomProps = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rdoPowerPoint);
            this.groupBox1.Controls.Add(this.rdoExcel);
            this.groupBox1.Controls.Add(this.rdoWord);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(220, 64);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Application";
            // 
            // rdoPowerPoint
            // 
            this.rdoPowerPoint.AutoSize = true;
            this.rdoPowerPoint.Location = new System.Drawing.Point(121, 27);
            this.rdoPowerPoint.Name = "rdoPowerPoint";
            this.rdoPowerPoint.Size = new System.Drawing.Size(86, 19);
            this.rdoPowerPoint.TabIndex = 2;
            this.rdoPowerPoint.Text = "PowerPoint";
            this.rdoPowerPoint.UseVisualStyleBackColor = true;
            this.rdoPowerPoint.CheckedChanged += new System.EventHandler(this.RdoPowerPoint_CheckedChanged);
            // 
            // rdoExcel
            // 
            this.rdoExcel.AutoSize = true;
            this.rdoExcel.Location = new System.Drawing.Point(63, 27);
            this.rdoExcel.Name = "rdoExcel";
            this.rdoExcel.Size = new System.Drawing.Size(52, 19);
            this.rdoExcel.TabIndex = 1;
            this.rdoExcel.Text = "Excel";
            this.rdoExcel.UseVisualStyleBackColor = true;
            this.rdoExcel.CheckedChanged += new System.EventHandler(this.RdoExcel_CheckedChanged);
            // 
            // rdoWord
            // 
            this.rdoWord.AutoSize = true;
            this.rdoWord.Checked = true;
            this.rdoWord.Location = new System.Drawing.Point(3, 27);
            this.rdoWord.Name = "rdoWord";
            this.rdoWord.Size = new System.Drawing.Size(54, 19);
            this.rdoWord.TabIndex = 0;
            this.rdoWord.TabStop = true;
            this.rdoWord.Text = "Word";
            this.rdoWord.UseVisualStyleBackColor = true;
            this.rdoWord.CheckedChanged += new System.EventHandler(this.RdoWord_CheckedChanged);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.BtnBrowseFolder);
            this.groupBox2.Controls.Add(this.tbFolderPath);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Location = new System.Drawing.Point(238, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(769, 64);
            this.groupBox2.TabIndex = 0;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Folder Location";
            // 
            // BtnBrowseFolder
            // 
            this.BtnBrowseFolder.Location = new System.Drawing.Point(635, 25);
            this.BtnBrowseFolder.Name = "BtnBrowseFolder";
            this.BtnBrowseFolder.Size = new System.Drawing.Size(126, 23);
            this.BtnBrowseFolder.TabIndex = 2;
            this.BtnBrowseFolder.Text = "...Choose Folder";
            this.BtnBrowseFolder.UseVisualStyleBackColor = true;
            this.BtnBrowseFolder.Click += new System.EventHandler(this.BtnBrowseFolder_Click);
            // 
            // tbFolderPath
            // 
            this.tbFolderPath.Location = new System.Drawing.Point(46, 26);
            this.tbFolderPath.Name = "tbFolderPath";
            this.tbFolderPath.Size = new System.Drawing.Size(583, 23);
            this.tbFolderPath.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 29);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(34, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "Path:";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.BtnCopyOutput);
            this.groupBox3.Controls.Add(this.ckbSubfolders);
            this.groupBox3.Controls.Add(this.lstOutput);
            this.groupBox3.Location = new System.Drawing.Point(12, 82);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(995, 285);
            this.groupBox3.TabIndex = 0;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Files";
            // 
            // BtnCopyOutput
            // 
            this.BtnCopyOutput.Location = new System.Drawing.Point(861, 254);
            this.BtnCopyOutput.Name = "BtnCopyOutput";
            this.BtnCopyOutput.Size = new System.Drawing.Size(126, 23);
            this.BtnCopyOutput.TabIndex = 2;
            this.BtnCopyOutput.Text = "Copy Output";
            this.BtnCopyOutput.UseVisualStyleBackColor = true;
            this.BtnCopyOutput.Click += new System.EventHandler(this.BtnCopyOutput_Click);
            // 
            // ckbSubfolders
            // 
            this.ckbSubfolders.AutoSize = true;
            this.ckbSubfolders.Location = new System.Drawing.Point(6, 254);
            this.ckbSubfolders.Name = "ckbSubfolders";
            this.ckbSubfolders.Size = new System.Drawing.Size(124, 19);
            this.ckbSubfolders.TabIndex = 1;
            this.ckbSubfolders.Text = "Include Subfolders";
            this.ckbSubfolders.UseVisualStyleBackColor = true;
            this.ckbSubfolders.CheckedChanged += new System.EventHandler(this.CkbSubfolders_CheckedChanged);
            // 
            // lstOutput
            // 
            this.lstOutput.FormattingEnabled = true;
            this.lstOutput.ItemHeight = 15;
            this.lstOutput.Location = new System.Drawing.Point(3, 19);
            this.lstOutput.Name = "lstOutput";
            this.lstOutput.Size = new System.Drawing.Size(984, 229);
            this.lstOutput.TabIndex = 0;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.BtnResetBulletMargins);
            this.groupBox4.Controls.Add(this.BtnRemoveCustomTitle);
            this.groupBox4.Controls.Add(this.BtnFixComments);
            this.groupBox4.Controls.Add(this.BtnRemovePII);
            this.groupBox4.Controls.Add(this.BtnFixTableProps);
            this.groupBox4.Controls.Add(this.BtnChangeTheme);
            this.groupBox4.Controls.Add(this.BtnDeleteRequestStatus);
            this.groupBox4.Controls.Add(this.BtnFixRevisions);
            this.groupBox4.Controls.Add(this.BtnFixBookmarks);
            this.groupBox4.Controls.Add(this.BtnFixNotesPage);
            this.groupBox4.Controls.Add(this.BtnRemovePIIOnSave);
            this.groupBox4.Controls.Add(this.BtnConvertStrict);
            this.groupBox4.Controls.Add(this.BtnUpdateNamespaces);
            this.groupBox4.Controls.Add(this.BtnSetOpenByDefault);
            this.groupBox4.Controls.Add(this.BtnChangeTemplate);
            this.groupBox4.Controls.Add(this.BtnFixHyperlinks);
            this.groupBox4.Controls.Add(this.BtnDeleteCustomProps);
            this.groupBox4.Controls.Add(this.BtnAddCustomProps);
            this.groupBox4.Location = new System.Drawing.Point(12, 367);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(995, 118);
            this.groupBox4.TabIndex = 0;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Batch Commands";
            // 
            // BtnResetBulletMargins
            // 
            this.BtnResetBulletMargins.Location = new System.Drawing.Point(838, 80);
            this.BtnResetBulletMargins.Name = "BtnResetBulletMargins";
            this.BtnResetBulletMargins.Size = new System.Drawing.Size(149, 23);
            this.BtnResetBulletMargins.TabIndex = 2;
            this.BtnResetBulletMargins.Text = "Reset Bullet Tab Margins";
            this.BtnResetBulletMargins.UseVisualStyleBackColor = true;
            this.BtnResetBulletMargins.Click += new System.EventHandler(this.BtnResetBulletMargins_Click);
            // 
            // BtnRemoveCustomTitle
            // 
            this.BtnRemoveCustomTitle.Location = new System.Drawing.Point(344, 51);
            this.BtnRemoveCustomTitle.Name = "BtnRemoveCustomTitle";
            this.BtnRemoveCustomTitle.Size = new System.Drawing.Size(173, 23);
            this.BtnRemoveCustomTitle.TabIndex = 1;
            this.BtnRemoveCustomTitle.Text = "Remove Custom Title Prop";
            this.BtnRemoveCustomTitle.UseVisualStyleBackColor = true;
            this.BtnRemoveCustomTitle.Click += new System.EventHandler(this.BtnRemoveCustomTitle_Click);
            // 
            // BtnFixComments
            // 
            this.BtnFixComments.Location = new System.Drawing.Point(681, 80);
            this.BtnFixComments.Name = "BtnFixComments";
            this.BtnFixComments.Size = new System.Drawing.Size(151, 23);
            this.BtnFixComments.TabIndex = 15;
            this.BtnFixComments.Text = "Fix Corrupt Comments";
            this.BtnFixComments.UseVisualStyleBackColor = true;
            this.BtnFixComments.Click += new System.EventHandler(this.BtnFixComments_Click);
            // 
            // BtnRemovePII
            // 
            this.BtnRemovePII.Location = new System.Drawing.Point(838, 22);
            this.BtnRemovePII.Name = "BtnRemovePII";
            this.BtnRemovePII.Size = new System.Drawing.Size(149, 23);
            this.BtnRemovePII.TabIndex = 14;
            this.BtnRemovePII.Text = "Remove PII";
            this.BtnRemovePII.UseVisualStyleBackColor = true;
            this.BtnRemovePII.Click += new System.EventHandler(this.BtnRemovePII_Click);
            // 
            // BtnFixTableProps
            // 
            this.BtnFixTableProps.Location = new System.Drawing.Point(523, 80);
            this.BtnFixTableProps.Name = "BtnFixTableProps";
            this.BtnFixTableProps.Size = new System.Drawing.Size(152, 23);
            this.BtnFixTableProps.TabIndex = 13;
            this.BtnFixTableProps.Text = "Fix Table Grid Props";
            this.BtnFixTableProps.UseVisualStyleBackColor = true;
            this.BtnFixTableProps.Click += new System.EventHandler(this.BtnFixTableProps_Click);
            // 
            // BtnChangeTheme
            // 
            this.BtnChangeTheme.Location = new System.Drawing.Point(838, 51);
            this.BtnChangeTheme.Name = "BtnChangeTheme";
            this.BtnChangeTheme.Size = new System.Drawing.Size(149, 23);
            this.BtnChangeTheme.TabIndex = 12;
            this.BtnChangeTheme.Text = "Change Theme";
            this.BtnChangeTheme.UseVisualStyleBackColor = true;
            this.BtnChangeTheme.Click += new System.EventHandler(this.BtnChangeTheme_Click);
            // 
            // BtnDeleteRequestStatus
            // 
            this.BtnDeleteRequestStatus.Location = new System.Drawing.Point(681, 51);
            this.BtnDeleteRequestStatus.Name = "BtnDeleteRequestStatus";
            this.BtnDeleteRequestStatus.Size = new System.Drawing.Size(151, 23);
            this.BtnDeleteRequestStatus.TabIndex = 11;
            this.BtnDeleteRequestStatus.Text = "Delete RequestStatus";
            this.BtnDeleteRequestStatus.UseVisualStyleBackColor = true;
            this.BtnDeleteRequestStatus.Click += new System.EventHandler(this.BtnDeleteRequestStatus_Click);
            // 
            // BtnFixRevisions
            // 
            this.BtnFixRevisions.Location = new System.Drawing.Point(523, 51);
            this.BtnFixRevisions.Name = "BtnFixRevisions";
            this.BtnFixRevisions.Size = new System.Drawing.Size(152, 23);
            this.BtnFixRevisions.TabIndex = 10;
            this.BtnFixRevisions.Text = "Fix Corrupt Revisions";
            this.BtnFixRevisions.UseVisualStyleBackColor = true;
            this.BtnFixRevisions.Click += new System.EventHandler(this.BtnFixRevisions_Click);
            // 
            // BtnFixBookmarks
            // 
            this.BtnFixBookmarks.Location = new System.Drawing.Point(523, 22);
            this.BtnFixBookmarks.Name = "BtnFixBookmarks";
            this.BtnFixBookmarks.Size = new System.Drawing.Size(152, 23);
            this.BtnFixBookmarks.TabIndex = 9;
            this.BtnFixBookmarks.Text = "Fix Corrupt Bookmarks";
            this.BtnFixBookmarks.UseVisualStyleBackColor = true;
            this.BtnFixBookmarks.Click += new System.EventHandler(this.BtnFixBookmarks_Click);
            // 
            // BtnFixNotesPage
            // 
            this.BtnFixNotesPage.Location = new System.Drawing.Point(344, 80);
            this.BtnFixNotesPage.Name = "BtnFixNotesPage";
            this.BtnFixNotesPage.Size = new System.Drawing.Size(173, 23);
            this.BtnFixNotesPage.TabIndex = 8;
            this.BtnFixNotesPage.Text = "Fix Notes Page Size";
            this.BtnFixNotesPage.UseVisualStyleBackColor = true;
            this.BtnFixNotesPage.Click += new System.EventHandler(this.BtnFixNotesPage_Click);
            // 
            // BtnRemovePIIOnSave
            // 
            this.BtnRemovePIIOnSave.Location = new System.Drawing.Point(681, 22);
            this.BtnRemovePIIOnSave.Name = "BtnRemovePIIOnSave";
            this.BtnRemovePIIOnSave.Size = new System.Drawing.Size(151, 23);
            this.BtnRemovePIIOnSave.TabIndex = 7;
            this.BtnRemovePIIOnSave.Text = "Remove PII On Save";
            this.BtnRemovePIIOnSave.UseVisualStyleBackColor = true;
            this.BtnRemovePIIOnSave.Click += new System.EventHandler(this.BtnRemovePIIOnSave_Click);
            // 
            // BtnConvertStrict
            // 
            this.BtnConvertStrict.Location = new System.Drawing.Point(344, 22);
            this.BtnConvertStrict.Name = "BtnConvertStrict";
            this.BtnConvertStrict.Size = new System.Drawing.Size(173, 23);
            this.BtnConvertStrict.TabIndex = 6;
            this.BtnConvertStrict.Text = "Convert Strict To Non-Strict";
            this.BtnConvertStrict.UseVisualStyleBackColor = true;
            this.BtnConvertStrict.Click += new System.EventHandler(this.BtnConvertStrict_Click);
            // 
            // BtnUpdateNamespaces
            // 
            this.BtnUpdateNamespaces.Location = new System.Drawing.Point(153, 80);
            this.BtnUpdateNamespaces.Name = "BtnUpdateNamespaces";
            this.BtnUpdateNamespaces.Size = new System.Drawing.Size(185, 23);
            this.BtnUpdateNamespaces.TabIndex = 5;
            this.BtnUpdateNamespaces.Text = "Update Quick Part Namespaces";
            this.BtnUpdateNamespaces.UseVisualStyleBackColor = true;
            this.BtnUpdateNamespaces.Click += new System.EventHandler(this.BtnUpdateNamespaces_Click);
            // 
            // BtnSetOpenByDefault
            // 
            this.BtnSetOpenByDefault.Location = new System.Drawing.Point(153, 51);
            this.BtnSetOpenByDefault.Name = "BtnSetOpenByDefault";
            this.BtnSetOpenByDefault.Size = new System.Drawing.Size(185, 23);
            this.BtnSetOpenByDefault.TabIndex = 4;
            this.BtnSetOpenByDefault.Text = "Set OpenByDefault = False";
            this.BtnSetOpenByDefault.UseVisualStyleBackColor = true;
            this.BtnSetOpenByDefault.Click += new System.EventHandler(this.BtnSetOpenByDefault_Click);
            // 
            // BtnChangeTemplate
            // 
            this.BtnChangeTemplate.Location = new System.Drawing.Point(153, 22);
            this.BtnChangeTemplate.Name = "BtnChangeTemplate";
            this.BtnChangeTemplate.Size = new System.Drawing.Size(185, 23);
            this.BtnChangeTemplate.TabIndex = 3;
            this.BtnChangeTemplate.Text = "Change Attached Template";
            this.BtnChangeTemplate.UseVisualStyleBackColor = true;
            this.BtnChangeTemplate.Click += new System.EventHandler(this.BtnChangeTemplate_Click);
            // 
            // BtnFixHyperlinks
            // 
            this.BtnFixHyperlinks.Location = new System.Drawing.Point(6, 80);
            this.BtnFixHyperlinks.Name = "BtnFixHyperlinks";
            this.BtnFixHyperlinks.Size = new System.Drawing.Size(141, 23);
            this.BtnFixHyperlinks.TabIndex = 2;
            this.BtnFixHyperlinks.Text = "Fix Corrupt Hyperlinks";
            this.BtnFixHyperlinks.UseVisualStyleBackColor = true;
            this.BtnFixHyperlinks.Click += new System.EventHandler(this.BtnFixHyperlinks_Click);
            // 
            // BtnDeleteCustomProps
            // 
            this.BtnDeleteCustomProps.Location = new System.Drawing.Point(6, 51);
            this.BtnDeleteCustomProps.Name = "BtnDeleteCustomProps";
            this.BtnDeleteCustomProps.Size = new System.Drawing.Size(141, 23);
            this.BtnDeleteCustomProps.TabIndex = 1;
            this.BtnDeleteCustomProps.Text = "Delete Custom Props";
            this.BtnDeleteCustomProps.UseVisualStyleBackColor = true;
            this.BtnDeleteCustomProps.Click += new System.EventHandler(this.BtnDeleteCustomProps_Click);
            // 
            // BtnAddCustomProps
            // 
            this.BtnAddCustomProps.Location = new System.Drawing.Point(6, 22);
            this.BtnAddCustomProps.Name = "BtnAddCustomProps";
            this.BtnAddCustomProps.Size = new System.Drawing.Size(141, 23);
            this.BtnAddCustomProps.TabIndex = 0;
            this.BtnAddCustomProps.Text = "Add Custom Props";
            this.BtnAddCustomProps.UseVisualStyleBackColor = true;
            this.BtnAddCustomProps.Click += new System.EventHandler(this.BtnAddCustomProps_Click);
            // 
            // FrmBatch
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1017, 495);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmBatch";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Batch File Processing";
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmBatch_KeyDown);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rdoPowerPoint;
        private System.Windows.Forms.RadioButton rdoExcel;
        private System.Windows.Forms.RadioButton rdoWord;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button BtnBrowseFolder;
        private System.Windows.Forms.TextBox tbFolderPath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button BtnCopyOutput;
        private System.Windows.Forms.CheckBox ckbSubfolders;
        private System.Windows.Forms.ListBox lstOutput;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Button BtnFixComments;
        private System.Windows.Forms.Button BtnRemovePII;
        private System.Windows.Forms.Button BtnFixTableProps;
        private System.Windows.Forms.Button BtnChangeTheme;
        private System.Windows.Forms.Button BtnDeleteRequestStatus;
        private System.Windows.Forms.Button BtnFixRevisions;
        private System.Windows.Forms.Button BtnFixBookmarks;
        private System.Windows.Forms.Button BtnFixNotesPage;
        private System.Windows.Forms.Button BtnRemovePIIOnSave;
        private System.Windows.Forms.Button BtnConvertStrict;
        private System.Windows.Forms.Button BtnUpdateNamespaces;
        private System.Windows.Forms.Button BtnSetOpenByDefault;
        private System.Windows.Forms.Button BtnChangeTemplate;
        private System.Windows.Forms.Button BtnFixHyperlinks;
        private System.Windows.Forms.Button BtnDeleteCustomProps;
        private System.Windows.Forms.Button BtnAddCustomProps;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button BtnRemoveCustomTitle;
        private System.Windows.Forms.Button BtnResetBulletMargins;
    }
}