
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
            groupBox1 = new System.Windows.Forms.GroupBox();
            rdoPowerPoint = new System.Windows.Forms.RadioButton();
            rdoExcel = new System.Windows.Forms.RadioButton();
            rdoWord = new System.Windows.Forms.RadioButton();
            groupBox2 = new System.Windows.Forms.GroupBox();
            BtnBrowseFolder = new System.Windows.Forms.Button();
            tbFolderPath = new System.Windows.Forms.TextBox();
            label1 = new System.Windows.Forms.Label();
            groupBox3 = new System.Windows.Forms.GroupBox();
            BtnCopyOutput = new System.Windows.Forms.Button();
            ckbSubfolders = new System.Windows.Forms.CheckBox();
            lstOutput = new System.Windows.Forms.ListBox();
            groupBox4 = new System.Windows.Forms.GroupBox();
            BtnFixTabBehavior = new System.Windows.Forms.Button();
            BtnFixDupeCustomXml = new System.Windows.Forms.Button();
            BtnRemoveCustomXml = new System.Windows.Forms.Button();
            BtnRemoveCustomFileProps = new System.Windows.Forms.Button();
            BtnFixCorruptTcTags = new System.Windows.Forms.Button();
            BtnFixFooterSpacing = new System.Windows.Forms.Button();
            BtnCheckForDigSig = new System.Windows.Forms.Button();
            BtnResetBulletMargins = new System.Windows.Forms.Button();
            BtnRemoveCustomTitle = new System.Windows.Forms.Button();
            BtnFixComments = new System.Windows.Forms.Button();
            BtnRemovePII = new System.Windows.Forms.Button();
            BtnFixTableProps = new System.Windows.Forms.Button();
            BtnChangeTheme = new System.Windows.Forms.Button();
            BtnDeleteRequestStatus = new System.Windows.Forms.Button();
            BtnFixRevisions = new System.Windows.Forms.Button();
            BtnFixBookmarks = new System.Windows.Forms.Button();
            BtnFixNotesPage = new System.Windows.Forms.Button();
            BtnRemovePIIOnSave = new System.Windows.Forms.Button();
            BtnConvertStrict = new System.Windows.Forms.Button();
            BtnUpdateNamespaces = new System.Windows.Forms.Button();
            BtnSetOpenByDefault = new System.Windows.Forms.Button();
            BtnChangeTemplate = new System.Windows.Forms.Button();
            BtnFixHyperlinks = new System.Windows.Forms.Button();
            BtnDeleteCustomProps = new System.Windows.Forms.Button();
            BtnAddCustomProps = new System.Windows.Forms.Button();
            folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            BtnFixCommentNotes = new System.Windows.Forms.Button();
            groupBox1.SuspendLayout();
            groupBox2.SuspendLayout();
            groupBox3.SuspendLayout();
            groupBox4.SuspendLayout();
            SuspendLayout();
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(rdoPowerPoint);
            groupBox1.Controls.Add(rdoExcel);
            groupBox1.Controls.Add(rdoWord);
            groupBox1.Location = new System.Drawing.Point(12, 12);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new System.Drawing.Size(220, 64);
            groupBox1.TabIndex = 0;
            groupBox1.TabStop = false;
            groupBox1.Text = "Application";
            // 
            // rdoPowerPoint
            // 
            rdoPowerPoint.AutoSize = true;
            rdoPowerPoint.Location = new System.Drawing.Point(121, 27);
            rdoPowerPoint.Name = "rdoPowerPoint";
            rdoPowerPoint.Size = new System.Drawing.Size(86, 19);
            rdoPowerPoint.TabIndex = 2;
            rdoPowerPoint.Text = "PowerPoint";
            rdoPowerPoint.UseVisualStyleBackColor = true;
            rdoPowerPoint.CheckedChanged += RdoPowerPoint_CheckedChanged;
            // 
            // rdoExcel
            // 
            rdoExcel.AutoSize = true;
            rdoExcel.Location = new System.Drawing.Point(63, 27);
            rdoExcel.Name = "rdoExcel";
            rdoExcel.Size = new System.Drawing.Size(52, 19);
            rdoExcel.TabIndex = 1;
            rdoExcel.Text = "Excel";
            rdoExcel.UseVisualStyleBackColor = true;
            rdoExcel.CheckedChanged += RdoExcel_CheckedChanged;
            // 
            // rdoWord
            // 
            rdoWord.AutoSize = true;
            rdoWord.Checked = true;
            rdoWord.Location = new System.Drawing.Point(3, 27);
            rdoWord.Name = "rdoWord";
            rdoWord.Size = new System.Drawing.Size(54, 19);
            rdoWord.TabIndex = 0;
            rdoWord.TabStop = true;
            rdoWord.Text = "Word";
            rdoWord.UseVisualStyleBackColor = true;
            rdoWord.CheckedChanged += RdoWord_CheckedChanged;
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(BtnBrowseFolder);
            groupBox2.Controls.Add(tbFolderPath);
            groupBox2.Controls.Add(label1);
            groupBox2.Location = new System.Drawing.Point(238, 12);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new System.Drawing.Size(769, 64);
            groupBox2.TabIndex = 0;
            groupBox2.TabStop = false;
            groupBox2.Text = "Folder Location";
            // 
            // BtnBrowseFolder
            // 
            BtnBrowseFolder.Location = new System.Drawing.Point(635, 25);
            BtnBrowseFolder.Name = "BtnBrowseFolder";
            BtnBrowseFolder.Size = new System.Drawing.Size(126, 23);
            BtnBrowseFolder.TabIndex = 2;
            BtnBrowseFolder.Text = "...Choose Folder";
            BtnBrowseFolder.UseVisualStyleBackColor = true;
            BtnBrowseFolder.Click += BtnBrowseFolder_Click;
            // 
            // tbFolderPath
            // 
            tbFolderPath.Location = new System.Drawing.Point(46, 26);
            tbFolderPath.Name = "tbFolderPath";
            tbFolderPath.Size = new System.Drawing.Size(583, 23);
            tbFolderPath.TabIndex = 1;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point(6, 29);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(34, 15);
            label1.TabIndex = 0;
            label1.Text = "Path:";
            // 
            // groupBox3
            // 
            groupBox3.Controls.Add(BtnCopyOutput);
            groupBox3.Controls.Add(ckbSubfolders);
            groupBox3.Controls.Add(lstOutput);
            groupBox3.Location = new System.Drawing.Point(12, 82);
            groupBox3.Name = "groupBox3";
            groupBox3.Size = new System.Drawing.Size(1004, 285);
            groupBox3.TabIndex = 0;
            groupBox3.TabStop = false;
            groupBox3.Text = "Files";
            // 
            // BtnCopyOutput
            // 
            BtnCopyOutput.Location = new System.Drawing.Point(861, 254);
            BtnCopyOutput.Name = "BtnCopyOutput";
            BtnCopyOutput.Size = new System.Drawing.Size(126, 23);
            BtnCopyOutput.TabIndex = 2;
            BtnCopyOutput.Text = "Copy Output";
            BtnCopyOutput.UseVisualStyleBackColor = true;
            BtnCopyOutput.Click += BtnCopyOutput_Click;
            // 
            // ckbSubfolders
            // 
            ckbSubfolders.AutoSize = true;
            ckbSubfolders.Location = new System.Drawing.Point(6, 254);
            ckbSubfolders.Name = "ckbSubfolders";
            ckbSubfolders.Size = new System.Drawing.Size(124, 19);
            ckbSubfolders.TabIndex = 1;
            ckbSubfolders.Text = "Include Subfolders";
            ckbSubfolders.UseVisualStyleBackColor = true;
            ckbSubfolders.CheckedChanged += CkbSubfolders_CheckedChanged;
            // 
            // lstOutput
            // 
            lstOutput.FormattingEnabled = true;
            lstOutput.HorizontalScrollbar = true;
            lstOutput.ItemHeight = 15;
            lstOutput.Location = new System.Drawing.Point(3, 19);
            lstOutput.Name = "lstOutput";
            lstOutput.Size = new System.Drawing.Size(984, 229);
            lstOutput.TabIndex = 0;
            // 
            // groupBox4
            // 
            groupBox4.Controls.Add(BtnFixCommentNotes);
            groupBox4.Controls.Add(BtnFixTabBehavior);
            groupBox4.Controls.Add(BtnFixDupeCustomXml);
            groupBox4.Controls.Add(BtnRemoveCustomXml);
            groupBox4.Controls.Add(BtnRemoveCustomFileProps);
            groupBox4.Controls.Add(BtnFixCorruptTcTags);
            groupBox4.Controls.Add(BtnFixFooterSpacing);
            groupBox4.Controls.Add(BtnCheckForDigSig);
            groupBox4.Controls.Add(BtnResetBulletMargins);
            groupBox4.Controls.Add(BtnRemoveCustomTitle);
            groupBox4.Controls.Add(BtnFixComments);
            groupBox4.Controls.Add(BtnRemovePII);
            groupBox4.Controls.Add(BtnFixTableProps);
            groupBox4.Controls.Add(BtnChangeTheme);
            groupBox4.Controls.Add(BtnDeleteRequestStatus);
            groupBox4.Controls.Add(BtnFixRevisions);
            groupBox4.Controls.Add(BtnFixBookmarks);
            groupBox4.Controls.Add(BtnFixNotesPage);
            groupBox4.Controls.Add(BtnRemovePIIOnSave);
            groupBox4.Controls.Add(BtnConvertStrict);
            groupBox4.Controls.Add(BtnUpdateNamespaces);
            groupBox4.Controls.Add(BtnSetOpenByDefault);
            groupBox4.Controls.Add(BtnChangeTemplate);
            groupBox4.Controls.Add(BtnFixHyperlinks);
            groupBox4.Controls.Add(BtnDeleteCustomProps);
            groupBox4.Controls.Add(BtnAddCustomProps);
            groupBox4.Location = new System.Drawing.Point(12, 367);
            groupBox4.Name = "groupBox4";
            groupBox4.Size = new System.Drawing.Size(1004, 171);
            groupBox4.TabIndex = 0;
            groupBox4.TabStop = false;
            groupBox4.Text = "Batch Commands";
            // 
            // BtnFixTabBehavior
            // 
            BtnFixTabBehavior.Location = new System.Drawing.Point(6, 138);
            BtnFixTabBehavior.Name = "BtnFixTabBehavior";
            BtnFixTabBehavior.Size = new System.Drawing.Size(185, 23);
            BtnFixTabBehavior.TabIndex = 18;
            BtnFixTabBehavior.Text = "Fix Tabbing Behavior";
            BtnFixTabBehavior.UseVisualStyleBackColor = true;
            BtnFixTabBehavior.Click += BtnFixTabBehavior_Click;
            // 
            // BtnFixDupeCustomXml
            // 
            BtnFixDupeCustomXml.Location = new System.Drawing.Point(838, 109);
            BtnFixDupeCustomXml.Name = "BtnFixDupeCustomXml";
            BtnFixDupeCustomXml.Size = new System.Drawing.Size(160, 23);
            BtnFixDupeCustomXml.TabIndex = 1;
            BtnFixDupeCustomXml.Text = "Fix Duplicate Custom Xml";
            BtnFixDupeCustomXml.UseVisualStyleBackColor = true;
            BtnFixDupeCustomXml.Click += BtnFixDupeCustomXml_Click;
            // 
            // BtnRemoveCustomXml
            // 
            BtnRemoveCustomXml.Location = new System.Drawing.Point(681, 109);
            BtnRemoveCustomXml.Name = "BtnRemoveCustomXml";
            BtnRemoveCustomXml.Size = new System.Drawing.Size(151, 23);
            BtnRemoveCustomXml.TabIndex = 1;
            BtnRemoveCustomXml.Text = "Remove Custom Xml";
            BtnRemoveCustomXml.UseVisualStyleBackColor = true;
            BtnRemoveCustomXml.Click += BtnRemoveCustomXml_Click;
            // 
            // BtnRemoveCustomFileProps
            // 
            BtnRemoveCustomFileProps.Location = new System.Drawing.Point(344, 109);
            BtnRemoveCustomFileProps.Name = "BtnRemoveCustomFileProps";
            BtnRemoveCustomFileProps.Size = new System.Drawing.Size(173, 23);
            BtnRemoveCustomFileProps.TabIndex = 17;
            BtnRemoveCustomFileProps.Text = "Remove Custom File Props";
            BtnRemoveCustomFileProps.UseVisualStyleBackColor = true;
            BtnRemoveCustomFileProps.Click += BtnRemoveCustomFileProps_Click;
            // 
            // BtnFixCorruptTcTags
            // 
            BtnFixCorruptTcTags.Location = new System.Drawing.Point(523, 109);
            BtnFixCorruptTcTags.Name = "BtnFixCorruptTcTags";
            BtnFixCorruptTcTags.Size = new System.Drawing.Size(152, 23);
            BtnFixCorruptTcTags.TabIndex = 1;
            BtnFixCorruptTcTags.Text = "Fix Corrupt Table Cells";
            BtnFixCorruptTcTags.UseVisualStyleBackColor = true;
            BtnFixCorruptTcTags.Click += BtnFixCorruptTcTags_Click;
            // 
            // BtnFixFooterSpacing
            // 
            BtnFixFooterSpacing.Location = new System.Drawing.Point(197, 109);
            BtnFixFooterSpacing.Name = "BtnFixFooterSpacing";
            BtnFixFooterSpacing.Size = new System.Drawing.Size(141, 23);
            BtnFixFooterSpacing.TabIndex = 1;
            BtnFixFooterSpacing.Text = "Fix Footer Spacing";
            BtnFixFooterSpacing.UseVisualStyleBackColor = true;
            BtnFixFooterSpacing.Click += BtnFixFooterSpacing_Click;
            // 
            // BtnCheckForDigSig
            // 
            BtnCheckForDigSig.Location = new System.Drawing.Point(6, 109);
            BtnCheckForDigSig.Name = "BtnCheckForDigSig";
            BtnCheckForDigSig.Size = new System.Drawing.Size(185, 23);
            BtnCheckForDigSig.TabIndex = 16;
            BtnCheckForDigSig.Text = "Check For Digital Signature";
            BtnCheckForDigSig.UseVisualStyleBackColor = true;
            BtnCheckForDigSig.Click += BtnCheckForDigSig_Click;
            // 
            // BtnResetBulletMargins
            // 
            BtnResetBulletMargins.Location = new System.Drawing.Point(838, 80);
            BtnResetBulletMargins.Name = "BtnResetBulletMargins";
            BtnResetBulletMargins.Size = new System.Drawing.Size(160, 23);
            BtnResetBulletMargins.TabIndex = 2;
            BtnResetBulletMargins.Text = "Reset Bullet Tab Margins";
            BtnResetBulletMargins.UseVisualStyleBackColor = true;
            BtnResetBulletMargins.Click += BtnResetBulletMargins_Click;
            // 
            // BtnRemoveCustomTitle
            // 
            BtnRemoveCustomTitle.Location = new System.Drawing.Point(344, 51);
            BtnRemoveCustomTitle.Name = "BtnRemoveCustomTitle";
            BtnRemoveCustomTitle.Size = new System.Drawing.Size(173, 23);
            BtnRemoveCustomTitle.TabIndex = 1;
            BtnRemoveCustomTitle.Text = "Remove Custom Title Prop";
            BtnRemoveCustomTitle.UseVisualStyleBackColor = true;
            BtnRemoveCustomTitle.Click += BtnRemoveCustomTitle_Click;
            // 
            // BtnFixComments
            // 
            BtnFixComments.Location = new System.Drawing.Point(681, 80);
            BtnFixComments.Name = "BtnFixComments";
            BtnFixComments.Size = new System.Drawing.Size(151, 23);
            BtnFixComments.TabIndex = 15;
            BtnFixComments.Text = "Fix Corrupt Comments";
            BtnFixComments.UseVisualStyleBackColor = true;
            BtnFixComments.Click += BtnFixComments_Click;
            // 
            // BtnRemovePII
            // 
            BtnRemovePII.Location = new System.Drawing.Point(838, 22);
            BtnRemovePII.Name = "BtnRemovePII";
            BtnRemovePII.Size = new System.Drawing.Size(159, 23);
            BtnRemovePII.TabIndex = 14;
            BtnRemovePII.Text = "Remove PII";
            BtnRemovePII.UseVisualStyleBackColor = true;
            BtnRemovePII.Click += BtnRemovePII_Click;
            // 
            // BtnFixTableProps
            // 
            BtnFixTableProps.Location = new System.Drawing.Point(523, 80);
            BtnFixTableProps.Name = "BtnFixTableProps";
            BtnFixTableProps.Size = new System.Drawing.Size(152, 23);
            BtnFixTableProps.TabIndex = 13;
            BtnFixTableProps.Text = "Fix Table Grid Props";
            BtnFixTableProps.UseVisualStyleBackColor = true;
            BtnFixTableProps.Click += BtnFixTableProps_Click;
            // 
            // BtnChangeTheme
            // 
            BtnChangeTheme.Location = new System.Drawing.Point(838, 51);
            BtnChangeTheme.Name = "BtnChangeTheme";
            BtnChangeTheme.Size = new System.Drawing.Size(159, 23);
            BtnChangeTheme.TabIndex = 12;
            BtnChangeTheme.Text = "Change Theme";
            BtnChangeTheme.UseVisualStyleBackColor = true;
            BtnChangeTheme.Click += BtnChangeTheme_Click;
            // 
            // BtnDeleteRequestStatus
            // 
            BtnDeleteRequestStatus.Location = new System.Drawing.Point(681, 51);
            BtnDeleteRequestStatus.Name = "BtnDeleteRequestStatus";
            BtnDeleteRequestStatus.Size = new System.Drawing.Size(151, 23);
            BtnDeleteRequestStatus.TabIndex = 11;
            BtnDeleteRequestStatus.Text = "Delete RequestStatus";
            BtnDeleteRequestStatus.UseVisualStyleBackColor = true;
            BtnDeleteRequestStatus.Click += BtnDeleteRequestStatus_Click;
            // 
            // BtnFixRevisions
            // 
            BtnFixRevisions.Location = new System.Drawing.Point(523, 51);
            BtnFixRevisions.Name = "BtnFixRevisions";
            BtnFixRevisions.Size = new System.Drawing.Size(152, 23);
            BtnFixRevisions.TabIndex = 10;
            BtnFixRevisions.Text = "Fix Corrupt Revisions";
            BtnFixRevisions.UseVisualStyleBackColor = true;
            BtnFixRevisions.Click += BtnFixRevisions_Click;
            // 
            // BtnFixBookmarks
            // 
            BtnFixBookmarks.Location = new System.Drawing.Point(523, 22);
            BtnFixBookmarks.Name = "BtnFixBookmarks";
            BtnFixBookmarks.Size = new System.Drawing.Size(152, 23);
            BtnFixBookmarks.TabIndex = 9;
            BtnFixBookmarks.Text = "Fix Corrupt Bookmarks";
            BtnFixBookmarks.UseVisualStyleBackColor = true;
            BtnFixBookmarks.Click += BtnFixBookmarks_Click;
            // 
            // BtnFixNotesPage
            // 
            BtnFixNotesPage.Location = new System.Drawing.Point(344, 80);
            BtnFixNotesPage.Name = "BtnFixNotesPage";
            BtnFixNotesPage.Size = new System.Drawing.Size(173, 23);
            BtnFixNotesPage.TabIndex = 8;
            BtnFixNotesPage.Text = "Fix Notes Page Size";
            BtnFixNotesPage.UseVisualStyleBackColor = true;
            BtnFixNotesPage.Click += BtnFixNotesPage_Click;
            // 
            // BtnRemovePIIOnSave
            // 
            BtnRemovePIIOnSave.Location = new System.Drawing.Point(681, 22);
            BtnRemovePIIOnSave.Name = "BtnRemovePIIOnSave";
            BtnRemovePIIOnSave.Size = new System.Drawing.Size(151, 23);
            BtnRemovePIIOnSave.TabIndex = 7;
            BtnRemovePIIOnSave.Text = "Remove PII On Save";
            BtnRemovePIIOnSave.UseVisualStyleBackColor = true;
            BtnRemovePIIOnSave.Click += BtnRemovePIIOnSave_Click;
            // 
            // BtnConvertStrict
            // 
            BtnConvertStrict.Location = new System.Drawing.Point(344, 22);
            BtnConvertStrict.Name = "BtnConvertStrict";
            BtnConvertStrict.Size = new System.Drawing.Size(173, 23);
            BtnConvertStrict.TabIndex = 6;
            BtnConvertStrict.Text = "Convert Strict To Non-Strict";
            BtnConvertStrict.UseVisualStyleBackColor = true;
            BtnConvertStrict.Click += BtnConvertStrict_Click;
            // 
            // BtnUpdateNamespaces
            // 
            BtnUpdateNamespaces.Location = new System.Drawing.Point(6, 80);
            BtnUpdateNamespaces.Name = "BtnUpdateNamespaces";
            BtnUpdateNamespaces.Size = new System.Drawing.Size(185, 23);
            BtnUpdateNamespaces.TabIndex = 5;
            BtnUpdateNamespaces.Text = "Update Quick Part Namespaces";
            BtnUpdateNamespaces.UseVisualStyleBackColor = true;
            BtnUpdateNamespaces.Click += BtnUpdateNamespaces_Click;
            // 
            // BtnSetOpenByDefault
            // 
            BtnSetOpenByDefault.Location = new System.Drawing.Point(6, 51);
            BtnSetOpenByDefault.Name = "BtnSetOpenByDefault";
            BtnSetOpenByDefault.Size = new System.Drawing.Size(185, 23);
            BtnSetOpenByDefault.TabIndex = 4;
            BtnSetOpenByDefault.Text = "Set OpenByDefault = False";
            BtnSetOpenByDefault.UseVisualStyleBackColor = true;
            BtnSetOpenByDefault.Click += BtnSetOpenByDefault_Click;
            // 
            // BtnChangeTemplate
            // 
            BtnChangeTemplate.Location = new System.Drawing.Point(6, 22);
            BtnChangeTemplate.Name = "BtnChangeTemplate";
            BtnChangeTemplate.Size = new System.Drawing.Size(185, 23);
            BtnChangeTemplate.TabIndex = 3;
            BtnChangeTemplate.Text = "Change Attached Template";
            BtnChangeTemplate.UseVisualStyleBackColor = true;
            BtnChangeTemplate.Click += BtnChangeTemplate_Click;
            // 
            // BtnFixHyperlinks
            // 
            BtnFixHyperlinks.Location = new System.Drawing.Point(197, 80);
            BtnFixHyperlinks.Name = "BtnFixHyperlinks";
            BtnFixHyperlinks.Size = new System.Drawing.Size(141, 23);
            BtnFixHyperlinks.TabIndex = 2;
            BtnFixHyperlinks.Text = "Fix Corrupt Hyperlinks";
            BtnFixHyperlinks.UseVisualStyleBackColor = true;
            BtnFixHyperlinks.Click += BtnFixHyperlinks_Click;
            // 
            // BtnDeleteCustomProps
            // 
            BtnDeleteCustomProps.Location = new System.Drawing.Point(197, 51);
            BtnDeleteCustomProps.Name = "BtnDeleteCustomProps";
            BtnDeleteCustomProps.Size = new System.Drawing.Size(141, 23);
            BtnDeleteCustomProps.TabIndex = 1;
            BtnDeleteCustomProps.Text = "Delete Custom Props";
            BtnDeleteCustomProps.UseVisualStyleBackColor = true;
            BtnDeleteCustomProps.Click += BtnDeleteCustomProps_Click;
            // 
            // BtnAddCustomProps
            // 
            BtnAddCustomProps.Location = new System.Drawing.Point(197, 22);
            BtnAddCustomProps.Name = "BtnAddCustomProps";
            BtnAddCustomProps.Size = new System.Drawing.Size(141, 23);
            BtnAddCustomProps.TabIndex = 0;
            BtnAddCustomProps.Text = "Add Custom Props";
            BtnAddCustomProps.UseVisualStyleBackColor = true;
            BtnAddCustomProps.Click += BtnAddCustomProps_Click;
            // 
            // BtnFixCommentNotes
            // 
            BtnFixCommentNotes.Location = new System.Drawing.Point(197, 138);
            BtnFixCommentNotes.Name = "BtnFixCommentNotes";
            BtnFixCommentNotes.Size = new System.Drawing.Size(141, 23);
            BtnFixCommentNotes.TabIndex = 1;
            BtnFixCommentNotes.Text = "Fix Comment Notes";
            BtnFixCommentNotes.UseVisualStyleBackColor = true;
            BtnFixCommentNotes.Click += BtnFixCommentNotes_Click;
            // 
            // FrmBatch
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(1021, 550);
            Controls.Add(groupBox2);
            Controls.Add(groupBox3);
            Controls.Add(groupBox4);
            Controls.Add(groupBox1);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            Icon = (System.Drawing.Icon)resources.GetObject("$this.Icon");
            KeyPreview = true;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "FrmBatch";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            Text = "Batch File Processing";
            KeyDown += FrmBatch_KeyDown;
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            groupBox2.ResumeLayout(false);
            groupBox2.PerformLayout();
            groupBox3.ResumeLayout(false);
            groupBox3.PerformLayout();
            groupBox4.ResumeLayout(false);
            ResumeLayout(false);
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
        private System.Windows.Forms.Button BtnCheckForDigSig;
        private System.Windows.Forms.Button BtnFixFooterSpacing;
        private System.Windows.Forms.Button BtnFixCorruptTcTags;
        private System.Windows.Forms.Button BtnRemoveCustomFileProps;
        private System.Windows.Forms.Button BtnRemoveCustomXml;
        private System.Windows.Forms.Button BtnFixDupeCustomXml;
        private System.Windows.Forms.Button BtnFixTabBehavior;
        private System.Windows.Forms.Button BtnFixCommentNotes;
    }
}