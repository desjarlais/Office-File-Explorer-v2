﻿
namespace Office_File_Explorer
{
    partial class FrmMain
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmMain));
            mnuMainMenu = new System.Windows.Forms.MenuStrip();
            fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            fileToolStripMenuItemOpen = new System.Windows.Forms.ToolStripMenuItem();
            fileToolStripMenuItemOpenErrorLog = new System.Windows.Forms.ToolStripMenuItem();
            fileToolStripMenuItemOpenFileBackup = new System.Windows.Forms.ToolStripMenuItem();
            fileToolStripMenuItemClose = new System.Windows.Forms.ToolStripMenuItem();
            fileToolStripMenuItemSettings = new System.Windows.Forms.ToolStripMenuItem();
            toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            fileToolStripMenuItemExit = new System.Windows.Forms.ToolStripMenuItem();
            editToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            editToolStripMenuFindReplace = new System.Windows.Forms.ToolStripMenuItem();
            editToolStripMenuItemModifyContents = new System.Windows.Forms.ToolStripMenuItem();
            editToolStripMenuItemRemoveCustomDocProps = new System.Windows.Forms.ToolStripMenuItem();
            editToolStripMenuItemRemoveCustomXml = new System.Windows.Forms.ToolStripMenuItem();
            toolsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            batchFileProcessingToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            clipboardViewerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            base64DecoderToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            structuredStorageViewerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            excelSheetViewerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            feedbackToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(components);
            copySelectedLineToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            copyAllLinesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            toolStrip1 = new System.Windows.Forms.ToolStrip();
            toolStripButtonViewContents = new System.Windows.Forms.ToolStripButton();
            toolStripButtonFixCorruptDoc = new System.Windows.Forms.ToolStripButton();
            toolStripButtonFixDoc = new System.Windows.Forms.ToolStripButton();
            toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            toolStripButtonModify = new System.Windows.Forms.ToolStripButton();
            toolStripButtonSave = new System.Windows.Forms.ToolStripButton();
            toolStripButtonInsertIcon = new System.Windows.Forms.ToolStripButton();
            toolStripButtonValidateXml = new System.Windows.Forms.ToolStripButton();
            toolStripButtonGenerateCallback = new System.Windows.Forms.ToolStripButton();
            toolStripDropDownButtonInsert = new System.Windows.Forms.ToolStripDropDownButton();
            office2010CustomUIPartToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            office2007CustomUIPartToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            customOutspaceToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            customTabToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            excelCustomTabToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            repurposeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            wordGroupOnInsertTabToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            statusStrip1 = new System.Windows.Forms.StatusStrip();
            toolStripStatusLabel3 = new System.Windows.Forms.ToolStripStatusLabel();
            toolStripStatusLabelDocType = new System.Windows.Forms.ToolStripStatusLabel();
            toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            toolStripStatusLabelFilePath = new System.Windows.Forms.ToolStripStatusLabel();
            splitContainer1 = new System.Windows.Forms.SplitContainer();
            tvFiles = new System.Windows.Forms.TreeView();
            tvImageList = new System.Windows.Forms.ImageList(components);
            rtbDisplay = new System.Windows.Forms.RichTextBox();
            mnuMainMenu.SuspendLayout();
            contextMenuStrip1.SuspendLayout();
            toolStrip1.SuspendLayout();
            statusStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)splitContainer1).BeginInit();
            splitContainer1.Panel1.SuspendLayout();
            splitContainer1.Panel2.SuspendLayout();
            splitContainer1.SuspendLayout();
            SuspendLayout();
            // 
            // mnuMainMenu
            // 
            mnuMainMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { fileToolStripMenuItem, editToolStripMenuItem, toolsToolStripMenuItem, helpToolStripMenuItem });
            mnuMainMenu.Location = new System.Drawing.Point(0, 0);
            mnuMainMenu.Name = "mnuMainMenu";
            mnuMainMenu.Size = new System.Drawing.Size(924, 24);
            mnuMainMenu.TabIndex = 0;
            mnuMainMenu.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { fileToolStripMenuItemOpen, fileToolStripMenuItemOpenErrorLog, fileToolStripMenuItemOpenFileBackup, fileToolStripMenuItemClose, toolStripSeparator1, fileToolStripMenuItemSettings, fileToolStripMenuItemExit });
            fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            fileToolStripMenuItem.Text = "&File";
            // 
            // fileToolStripMenuItemOpen
            // 
            fileToolStripMenuItemOpen.Image = Properties.Resources.OpenFile;
            fileToolStripMenuItemOpen.Name = "fileToolStripMenuItemOpen";
            fileToolStripMenuItemOpen.Size = new System.Drawing.Size(202, 22);
            fileToolStripMenuItemOpen.Text = "O&pen";
            fileToolStripMenuItemOpen.Click += OpenToolStripMenuItem_Click;
            // 
            // fileToolStripMenuItemOpenErrorLog
            // 
            fileToolStripMenuItemOpenErrorLog.Image = Properties.Resources.ErrorSummary;
            fileToolStripMenuItemOpenErrorLog.Name = "fileToolStripMenuItemOpenErrorLog";
            fileToolStripMenuItemOpenErrorLog.Size = new System.Drawing.Size(202, 22);
            fileToolStripMenuItemOpenErrorLog.Text = "Open Error Log";
            fileToolStripMenuItemOpenErrorLog.Click += OpenErrorLogToolStripMenuItem_Click;
            // 
            // fileToolStripMenuItemOpenFileBackup
            // 
            fileToolStripMenuItemOpenFileBackup.Image = (System.Drawing.Image)resources.GetObject("fileToolStripMenuItemOpenFileBackup.Image");
            fileToolStripMenuItemOpenFileBackup.Name = "fileToolStripMenuItemOpenFileBackup";
            fileToolStripMenuItemOpenFileBackup.Size = new System.Drawing.Size(202, 22);
            fileToolStripMenuItemOpenFileBackup.Text = "Open File Backup Folder";
            fileToolStripMenuItemOpenFileBackup.Click += openFileBackupFolderToolStripMenuItem_Click;
            // 
            // fileToolStripMenuItemClose
            // 
            fileToolStripMenuItemClose.Enabled = false;
            fileToolStripMenuItemClose.Image = Properties.Resources.CloseDocument;
            fileToolStripMenuItemClose.Name = "fileToolStripMenuItemClose";
            fileToolStripMenuItemClose.Size = new System.Drawing.Size(202, 22);
            fileToolStripMenuItemClose.Text = "&Close";
            fileToolStripMenuItemClose.Click += fileToolStripMenuItemClose_Click;
            // 
            // fileToolStripMenuItemSettings
            // 
            fileToolStripMenuItemSettings.Image = Properties.Resources.Settings;
            fileToolStripMenuItemSettings.Name = "fileToolStripMenuItemSettings";
            fileToolStripMenuItemSettings.Size = new System.Drawing.Size(202, 22);
            fileToolStripMenuItemSettings.Text = "&Settings";
            fileToolStripMenuItemSettings.Click += SettingsToolStripMenuItem_Click;
            // 
            // toolStripSeparator1
            // 
            toolStripSeparator1.Name = "toolStripSeparator1";
            toolStripSeparator1.Size = new System.Drawing.Size(199, 6);
            // 
            // fileToolStripMenuItemExit
            // 
            fileToolStripMenuItemExit.Name = "fileToolStripMenuItemExit";
            fileToolStripMenuItemExit.Size = new System.Drawing.Size(202, 22);
            fileToolStripMenuItemExit.Text = "E&xit";
            fileToolStripMenuItemExit.Click += ExitToolStripMenuItem_Click;
            // 
            // editToolStripMenuItem
            // 
            editToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { editToolStripMenuFindReplace, editToolStripMenuItemModifyContents, editToolStripMenuItemRemoveCustomDocProps, editToolStripMenuItemRemoveCustomXml });
            editToolStripMenuItem.Name = "editToolStripMenuItem";
            editToolStripMenuItem.Size = new System.Drawing.Size(39, 20);
            editToolStripMenuItem.Text = "&Edit";
            // 
            // editToolStripMenuFindReplace
            // 
            editToolStripMenuFindReplace.Enabled = false;
            editToolStripMenuFindReplace.Image = Properties.Resources.FindInFile;
            editToolStripMenuFindReplace.Name = "editToolStripMenuFindReplace";
            editToolStripMenuFindReplace.Size = new System.Drawing.Size(277, 22);
            editToolStripMenuFindReplace.Text = "Find and Replace";
            editToolStripMenuFindReplace.Click += editToolStripMenuFindReplace_Click;
            // 
            // editToolStripMenuItemModifyContents
            // 
            editToolStripMenuItemModifyContents.Enabled = false;
            editToolStripMenuItemModifyContents.Image = Properties.Resources.Edit;
            editToolStripMenuItemModifyContents.Name = "editToolStripMenuItemModifyContents";
            editToolStripMenuItemModifyContents.Size = new System.Drawing.Size(277, 22);
            editToolStripMenuItemModifyContents.Text = "File Contents";
            editToolStripMenuItemModifyContents.Click += editToolStripMenuItemModifyContents_Click;
            // 
            // editToolStripMenuItemRemoveCustomDocProps
            // 
            editToolStripMenuItemRemoveCustomDocProps.Enabled = false;
            editToolStripMenuItemRemoveCustomDocProps.Image = (System.Drawing.Image)resources.GetObject("editToolStripMenuItemRemoveCustomDocProps.Image");
            editToolStripMenuItemRemoveCustomDocProps.Name = "editToolStripMenuItemRemoveCustomDocProps";
            editToolStripMenuItemRemoveCustomDocProps.Size = new System.Drawing.Size(277, 22);
            editToolStripMenuItemRemoveCustomDocProps.Text = "Remove Custom Document Properties";
            editToolStripMenuItemRemoveCustomDocProps.Click += editToolStripMenuItemRemoveCustomDocProps_Click;
            // 
            // editToolStripMenuItemRemoveCustomXml
            // 
            editToolStripMenuItemRemoveCustomXml.Enabled = false;
            editToolStripMenuItemRemoveCustomXml.Image = Properties.Resources.DeleteTag;
            editToolStripMenuItemRemoveCustomXml.Name = "editToolStripMenuItemRemoveCustomXml";
            editToolStripMenuItemRemoveCustomXml.Size = new System.Drawing.Size(277, 22);
            editToolStripMenuItemRemoveCustomXml.Text = "Remove Custom Xml";
            editToolStripMenuItemRemoveCustomXml.Click += editToolStripMenuItemRemoveCustomXml_Click;
            // 
            // toolsToolStripMenuItem
            // 
            toolsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { batchFileProcessingToolStripMenuItem, clipboardViewerToolStripMenuItem, base64DecoderToolStripMenuItem, structuredStorageViewerToolStripMenuItem, excelSheetViewerToolStripMenuItem });
            toolsToolStripMenuItem.Name = "toolsToolStripMenuItem";
            toolsToolStripMenuItem.Size = new System.Drawing.Size(46, 20);
            toolsToolStripMenuItem.Text = "&Tools";
            // 
            // batchFileProcessingToolStripMenuItem
            // 
            batchFileProcessingToolStripMenuItem.Image = Properties.Resources.OpenDocumentGroup;
            batchFileProcessingToolStripMenuItem.Name = "batchFileProcessingToolStripMenuItem";
            batchFileProcessingToolStripMenuItem.Size = new System.Drawing.Size(210, 22);
            batchFileProcessingToolStripMenuItem.Text = "Batch File Processing";
            batchFileProcessingToolStripMenuItem.Click += BatchFileProcessingToolStripMenuItem_Click;
            // 
            // clipboardViewerToolStripMenuItem
            // 
            clipboardViewerToolStripMenuItem.Image = Properties.Resources.Copy;
            clipboardViewerToolStripMenuItem.Name = "clipboardViewerToolStripMenuItem";
            clipboardViewerToolStripMenuItem.Size = new System.Drawing.Size(210, 22);
            clipboardViewerToolStripMenuItem.Text = "Clipboard Viewer";
            clipboardViewerToolStripMenuItem.Click += ClipboardViewerToolStripMenuItem_Click;
            // 
            // base64DecoderToolStripMenuItem
            // 
            base64DecoderToolStripMenuItem.Image = Properties.Resources.Strings_7959_0;
            base64DecoderToolStripMenuItem.Name = "base64DecoderToolStripMenuItem";
            base64DecoderToolStripMenuItem.Size = new System.Drawing.Size(210, 22);
            base64DecoderToolStripMenuItem.Text = "Base64 Decoder";
            base64DecoderToolStripMenuItem.Click += Base64DecoderToolStripMenuItem_Click;
            // 
            // structuredStorageViewerToolStripMenuItem
            // 
            structuredStorageViewerToolStripMenuItem.Enabled = false;
            structuredStorageViewerToolStripMenuItem.Image = Properties.Resources.BinaryFile;
            structuredStorageViewerToolStripMenuItem.Name = "structuredStorageViewerToolStripMenuItem";
            structuredStorageViewerToolStripMenuItem.Size = new System.Drawing.Size(210, 22);
            structuredStorageViewerToolStripMenuItem.Text = "Structured Storage Viewer";
            structuredStorageViewerToolStripMenuItem.Click += structuredStorageViewerToolStripMenuItem_Click;
            // 
            // excelSheetViewerToolStripMenuItem
            // 
            excelSheetViewerToolStripMenuItem.Enabled = false;
            excelSheetViewerToolStripMenuItem.Image = (System.Drawing.Image)resources.GetObject("excelSheetViewerToolStripMenuItem.Image");
            excelSheetViewerToolStripMenuItem.Name = "excelSheetViewerToolStripMenuItem";
            excelSheetViewerToolStripMenuItem.Size = new System.Drawing.Size(210, 22);
            excelSheetViewerToolStripMenuItem.Text = "Excel Sheet Viewer";
            excelSheetViewerToolStripMenuItem.Click += excelSheetViewerToolStripMenuItem_Click;
            // 
            // helpToolStripMenuItem
            // 
            helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { aboutToolStripMenuItem, feedbackToolStripMenuItem });
            helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            helpToolStripMenuItem.Text = "&Help";
            // 
            // aboutToolStripMenuItem
            // 
            aboutToolStripMenuItem.Image = Properties.Resources.AboutBox;
            aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            aboutToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            aboutToolStripMenuItem.Text = "About";
            aboutToolStripMenuItem.Click += AboutToolStripMenuItem_Click;
            // 
            // feedbackToolStripMenuItem
            // 
            feedbackToolStripMenuItem.Image = Properties.Resources.Feedback;
            feedbackToolStripMenuItem.Name = "feedbackToolStripMenuItem";
            feedbackToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            feedbackToolStripMenuItem.Text = "Feedback";
            feedbackToolStripMenuItem.Click += FeedbackToolStripMenuItem_Click;
            // 
            // contextMenuStrip1
            // 
            contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { copySelectedLineToolStripMenuItem, copyAllLinesToolStripMenuItem });
            contextMenuStrip1.Name = "contextMenuStrip1";
            contextMenuStrip1.Size = new System.Drawing.Size(175, 48);
            // 
            // copySelectedLineToolStripMenuItem
            // 
            copySelectedLineToolStripMenuItem.Name = "copySelectedLineToolStripMenuItem";
            copySelectedLineToolStripMenuItem.Size = new System.Drawing.Size(174, 22);
            copySelectedLineToolStripMenuItem.Text = "Copy Selected Line";
            copySelectedLineToolStripMenuItem.Click += CopySelectedLineToolStripMenuItem_Click;
            // 
            // copyAllLinesToolStripMenuItem
            // 
            copyAllLinesToolStripMenuItem.Name = "copyAllLinesToolStripMenuItem";
            copyAllLinesToolStripMenuItem.Size = new System.Drawing.Size(174, 22);
            copyAllLinesToolStripMenuItem.Text = "Copy All Lines";
            copyAllLinesToolStripMenuItem.Click += CopyAllLinesToolStripMenuItem_Click;
            // 
            // toolStrip1
            // 
            toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { toolStripButtonViewContents, toolStripButtonFixCorruptDoc, toolStripButtonFixDoc, toolStripSeparator2, toolStripButtonModify, toolStripButtonSave, toolStripButtonInsertIcon, toolStripButtonValidateXml, toolStripButtonGenerateCallback, toolStripDropDownButtonInsert });
            toolStrip1.Location = new System.Drawing.Point(0, 24);
            toolStrip1.Name = "toolStrip1";
            toolStrip1.Size = new System.Drawing.Size(924, 25);
            toolStrip1.TabIndex = 3;
            toolStrip1.Text = "toolStrip1";
            // 
            // toolStripButtonViewContents
            // 
            toolStripButtonViewContents.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            toolStripButtonViewContents.Enabled = false;
            toolStripButtonViewContents.Image = (System.Drawing.Image)resources.GetObject("toolStripButtonViewContents.Image");
            toolStripButtonViewContents.ImageTransparentColor = System.Drawing.Color.Magenta;
            toolStripButtonViewContents.Name = "toolStripButtonViewContents";
            toolStripButtonViewContents.Size = new System.Drawing.Size(87, 22);
            toolStripButtonViewContents.Text = "View Contents";
            toolStripButtonViewContents.Click += toolStripButtonViewContents_Click;
            // 
            // toolStripButtonFixCorruptDoc
            // 
            toolStripButtonFixCorruptDoc.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            toolStripButtonFixCorruptDoc.Enabled = false;
            toolStripButtonFixCorruptDoc.Image = (System.Drawing.Image)resources.GetObject("toolStripButtonFixCorruptDoc.Image");
            toolStripButtonFixCorruptDoc.ImageTransparentColor = System.Drawing.Color.Magenta;
            toolStripButtonFixCorruptDoc.Name = "toolStripButtonFixCorruptDoc";
            toolStripButtonFixCorruptDoc.Size = new System.Drawing.Size(129, 22);
            toolStripButtonFixCorruptDoc.Text = "Fix Corrupt Document";
            toolStripButtonFixCorruptDoc.Click += toolStripButtonFixCorruptDoc_Click;
            // 
            // toolStripButtonFixDoc
            // 
            toolStripButtonFixDoc.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            toolStripButtonFixDoc.Enabled = false;
            toolStripButtonFixDoc.Image = (System.Drawing.Image)resources.GetObject("toolStripButtonFixDoc.Image");
            toolStripButtonFixDoc.ImageTransparentColor = System.Drawing.Color.Magenta;
            toolStripButtonFixDoc.Name = "toolStripButtonFixDoc";
            toolStripButtonFixDoc.Size = new System.Drawing.Size(85, 22);
            toolStripButtonFixDoc.Text = "Fix Document";
            toolStripButtonFixDoc.Click += toolStripButtonFixDoc_Click;
            // 
            // toolStripSeparator2
            // 
            toolStripSeparator2.Name = "toolStripSeparator2";
            toolStripSeparator2.Size = new System.Drawing.Size(6, 25);
            // 
            // toolStripButtonModify
            // 
            toolStripButtonModify.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            toolStripButtonModify.Enabled = false;
            toolStripButtonModify.Image = Properties.Resources.ModifyPropertyTrivial;
            toolStripButtonModify.ImageTransparentColor = System.Drawing.Color.Magenta;
            toolStripButtonModify.Name = "toolStripButtonModify";
            toolStripButtonModify.Size = new System.Drawing.Size(23, 22);
            toolStripButtonModify.Text = "Modify Xml";
            toolStripButtonModify.ToolTipText = "Modify Xml";
            toolStripButtonModify.Click += toolStripButtonModify_Click;
            // 
            // toolStripButtonSave
            // 
            toolStripButtonSave.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            toolStripButtonSave.Enabled = false;
            toolStripButtonSave.Image = Properties.Resources.Save;
            toolStripButtonSave.ImageTransparentColor = System.Drawing.Color.Magenta;
            toolStripButtonSave.Name = "toolStripButtonSave";
            toolStripButtonSave.Size = new System.Drawing.Size(23, 22);
            toolStripButtonSave.Text = "Save Xml";
            toolStripButtonSave.Click += toolStripButtonSave_Click;
            // 
            // toolStripButtonInsertIcon
            // 
            toolStripButtonInsertIcon.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            toolStripButtonInsertIcon.Enabled = false;
            toolStripButtonInsertIcon.Image = (System.Drawing.Image)resources.GetObject("toolStripButtonInsertIcon.Image");
            toolStripButtonInsertIcon.ImageTransparentColor = System.Drawing.Color.Magenta;
            toolStripButtonInsertIcon.Name = "toolStripButtonInsertIcon";
            toolStripButtonInsertIcon.Size = new System.Drawing.Size(23, 22);
            toolStripButtonInsertIcon.Text = "Insert Icon";
            toolStripButtonInsertIcon.Click += toolStripButtonInsertIcon_Click;
            // 
            // toolStripButtonValidateXml
            // 
            toolStripButtonValidateXml.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            toolStripButtonValidateXml.Enabled = false;
            toolStripButtonValidateXml.Image = (System.Drawing.Image)resources.GetObject("toolStripButtonValidateXml.Image");
            toolStripButtonValidateXml.ImageTransparentColor = System.Drawing.Color.Magenta;
            toolStripButtonValidateXml.Name = "toolStripButtonValidateXml";
            toolStripButtonValidateXml.Size = new System.Drawing.Size(23, 22);
            toolStripButtonValidateXml.Text = "Validate Xml";
            toolStripButtonValidateXml.Click += toolStripButtonValidateXml_Click;
            // 
            // toolStripButtonGenerateCallback
            // 
            toolStripButtonGenerateCallback.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            toolStripButtonGenerateCallback.Enabled = false;
            toolStripButtonGenerateCallback.Image = (System.Drawing.Image)resources.GetObject("toolStripButtonGenerateCallback.Image");
            toolStripButtonGenerateCallback.ImageTransparentColor = System.Drawing.Color.Magenta;
            toolStripButtonGenerateCallback.Name = "toolStripButtonGenerateCallback";
            toolStripButtonGenerateCallback.Size = new System.Drawing.Size(23, 22);
            toolStripButtonGenerateCallback.Text = "Generate Callbacks";
            toolStripButtonGenerateCallback.Click += toolStripButtonGenerateCallback_Click;
            // 
            // toolStripDropDownButtonInsert
            // 
            toolStripDropDownButtonInsert.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            toolStripDropDownButtonInsert.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { office2010CustomUIPartToolStripMenuItem, office2007CustomUIPartToolStripMenuItem, toolStripMenuItem1 });
            toolStripDropDownButtonInsert.Enabled = false;
            toolStripDropDownButtonInsert.Image = (System.Drawing.Image)resources.GetObject("toolStripDropDownButtonInsert.Image");
            toolStripDropDownButtonInsert.ImageTransparentColor = System.Drawing.Color.Magenta;
            toolStripDropDownButtonInsert.Name = "toolStripDropDownButtonInsert";
            toolStripDropDownButtonInsert.Size = new System.Drawing.Size(49, 22);
            toolStripDropDownButtonInsert.Text = "Insert";
            toolStripDropDownButtonInsert.ToolTipText = "Insert Custom UI";
            // 
            // office2010CustomUIPartToolStripMenuItem
            // 
            office2010CustomUIPartToolStripMenuItem.Image = (System.Drawing.Image)resources.GetObject("office2010CustomUIPartToolStripMenuItem.Image");
            office2010CustomUIPartToolStripMenuItem.Name = "office2010CustomUIPartToolStripMenuItem";
            office2010CustomUIPartToolStripMenuItem.Size = new System.Drawing.Size(216, 22);
            office2010CustomUIPartToolStripMenuItem.Text = "Office 2010 Custom UI Part";
            office2010CustomUIPartToolStripMenuItem.Click += office2010CustomUIPartToolStripMenuItem_Click;
            // 
            // office2007CustomUIPartToolStripMenuItem
            // 
            office2007CustomUIPartToolStripMenuItem.Image = Properties.Resources._90_904004_pixel_ms_office_word_2007_logo;
            office2007CustomUIPartToolStripMenuItem.Name = "office2007CustomUIPartToolStripMenuItem";
            office2007CustomUIPartToolStripMenuItem.Size = new System.Drawing.Size(216, 22);
            office2007CustomUIPartToolStripMenuItem.Text = "Office 2007 Custom UI Part";
            office2007CustomUIPartToolStripMenuItem.Click += office2007CustomUIPartToolStripMenuItem_Click;
            // 
            // toolStripMenuItem1
            // 
            toolStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { customOutspaceToolStripMenuItem, customTabToolStripMenuItem, excelCustomTabToolStripMenuItem, repurposeToolStripMenuItem, wordGroupOnInsertTabToolStripMenuItem });
            toolStripMenuItem1.Name = "toolStripMenuItem1";
            toolStripMenuItem1.Size = new System.Drawing.Size(216, 22);
            toolStripMenuItem1.Text = "Sample XML";
            // 
            // customOutspaceToolStripMenuItem
            // 
            customOutspaceToolStripMenuItem.Name = "customOutspaceToolStripMenuItem";
            customOutspaceToolStripMenuItem.Size = new System.Drawing.Size(217, 22);
            customOutspaceToolStripMenuItem.Text = "Custom Outspace";
            customOutspaceToolStripMenuItem.Click += customOutspaceToolStripMenuItem_Click;
            // 
            // customTabToolStripMenuItem
            // 
            customTabToolStripMenuItem.Name = "customTabToolStripMenuItem";
            customTabToolStripMenuItem.Size = new System.Drawing.Size(217, 22);
            customTabToolStripMenuItem.Text = "Custom Tab";
            customTabToolStripMenuItem.Click += customTabToolStripMenuItem_Click;
            // 
            // excelCustomTabToolStripMenuItem
            // 
            excelCustomTabToolStripMenuItem.Name = "excelCustomTabToolStripMenuItem";
            excelCustomTabToolStripMenuItem.Size = new System.Drawing.Size(217, 22);
            excelCustomTabToolStripMenuItem.Text = "Excel - Custom Tab";
            excelCustomTabToolStripMenuItem.Click += excelCustomTabToolStripMenuItem_Click;
            // 
            // repurposeToolStripMenuItem
            // 
            repurposeToolStripMenuItem.Name = "repurposeToolStripMenuItem";
            repurposeToolStripMenuItem.Size = new System.Drawing.Size(217, 22);
            repurposeToolStripMenuItem.Text = "Repurpose";
            repurposeToolStripMenuItem.Click += repurposeToolStripMenuItem_Click;
            // 
            // wordGroupOnInsertTabToolStripMenuItem
            // 
            wordGroupOnInsertTabToolStripMenuItem.Name = "wordGroupOnInsertTabToolStripMenuItem";
            wordGroupOnInsertTabToolStripMenuItem.Size = new System.Drawing.Size(217, 22);
            wordGroupOnInsertTabToolStripMenuItem.Text = "Word - Group on Insert Tab";
            wordGroupOnInsertTabToolStripMenuItem.Click += wordGroupOnInsertTabToolStripMenuItem_Click;
            // 
            // statusStrip1
            // 
            statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { toolStripStatusLabel3, toolStripStatusLabelDocType, toolStripStatusLabel1, toolStripStatusLabelFilePath });
            statusStrip1.Location = new System.Drawing.Point(0, 541);
            statusStrip1.Name = "statusStrip1";
            statusStrip1.Size = new System.Drawing.Size(924, 22);
            statusStrip1.TabIndex = 4;
            statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel3
            // 
            toolStripStatusLabel3.Name = "toolStripStatusLabel3";
            toolStripStatusLabel3.Size = new System.Drawing.Size(55, 17);
            toolStripStatusLabel3.Text = "File Type:";
            // 
            // toolStripStatusLabelDocType
            // 
            toolStripStatusLabelDocType.Name = "toolStripStatusLabelDocType";
            toolStripStatusLabelDocType.Size = new System.Drawing.Size(22, 17);
            toolStripStatusLabelDocType.Text = "---";
            // 
            // toolStripStatusLabel1
            // 
            toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            toolStripStatusLabel1.Size = new System.Drawing.Size(61, 17);
            toolStripStatusLabel1.Text = "| File Path:";
            // 
            // toolStripStatusLabelFilePath
            // 
            toolStripStatusLabelFilePath.Name = "toolStripStatusLabelFilePath";
            toolStripStatusLabelFilePath.Size = new System.Drawing.Size(22, 17);
            toolStripStatusLabelFilePath.Text = "---";
            // 
            // splitContainer1
            // 
            splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            splitContainer1.Location = new System.Drawing.Point(0, 49);
            splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            splitContainer1.Panel1.Controls.Add(tvFiles);
            // 
            // splitContainer1.Panel2
            // 
            splitContainer1.Panel2.Controls.Add(rtbDisplay);
            splitContainer1.Size = new System.Drawing.Size(924, 492);
            splitContainer1.SplitterDistance = 308;
            splitContainer1.TabIndex = 5;
            // 
            // tvFiles
            // 
            tvFiles.Dock = System.Windows.Forms.DockStyle.Fill;
            tvFiles.ImageIndex = 0;
            tvFiles.ImageList = tvImageList;
            tvFiles.Location = new System.Drawing.Point(0, 0);
            tvFiles.Name = "tvFiles";
            tvFiles.SelectedImageIndex = 0;
            tvFiles.Size = new System.Drawing.Size(308, 492);
            tvFiles.TabIndex = 0;
            tvFiles.AfterSelect += tvFiles_AfterSelect;
            // 
            // tvImageList
            // 
            tvImageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
            tvImageList.ImageStream = (System.Windows.Forms.ImageListStreamer)resources.GetObject("tvImageList.ImageStream");
            tvImageList.TransparentColor = System.Drawing.Color.Transparent;
            tvImageList.Images.SetKeyName(0, "worddoc.bmp");
            tvImageList.Images.SetKeyName(1, "pptpre.bmp");
            tvImageList.Images.SetKeyName(2, "excelwkb.bmp");
            tvImageList.Images.SetKeyName(3, "xml.png");
            tvImageList.Images.SetKeyName(4, "insertPicture.png");
            tvImageList.Images.SetKeyName(5, "BinaryFile.png");
            tvImageList.Images.SetKeyName(6, "folder.png");
            tvImageList.Images.SetKeyName(7, "file.png");
            // 
            // rtbDisplay
            // 
            rtbDisplay.ContextMenuStrip = contextMenuStrip1;
            rtbDisplay.Dock = System.Windows.Forms.DockStyle.Fill;
            rtbDisplay.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            rtbDisplay.Location = new System.Drawing.Point(0, 0);
            rtbDisplay.Name = "rtbDisplay";
            rtbDisplay.ReadOnly = true;
            rtbDisplay.Size = new System.Drawing.Size(612, 492);
            rtbDisplay.TabIndex = 0;
            rtbDisplay.Text = "";
            // 
            // FrmMain
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(924, 563);
            Controls.Add(splitContainer1);
            Controls.Add(statusStrip1);
            Controls.Add(toolStrip1);
            Controls.Add(mnuMainMenu);
            Icon = (System.Drawing.Icon)resources.GetObject("$this.Icon");
            MainMenuStrip = mnuMainMenu;
            MinimumSize = new System.Drawing.Size(920, 588);
            Name = "FrmMain";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text = "Office File Explorer";
            FormClosing += FrmMain_FormClosing;
            mnuMainMenu.ResumeLayout(false);
            mnuMainMenu.PerformLayout();
            contextMenuStrip1.ResumeLayout(false);
            toolStrip1.ResumeLayout(false);
            toolStrip1.PerformLayout();
            statusStrip1.ResumeLayout(false);
            statusStrip1.PerformLayout();
            splitContainer1.Panel1.ResumeLayout(false);
            splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)splitContainer1).EndInit();
            splitContainer1.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private System.Windows.Forms.MenuStrip mnuMainMenu;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItemOpen;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItemSettings;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItemExit;
        private System.Windows.Forms.ToolStripMenuItem batchFileProcessingToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem clipboardViewerToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem feedbackToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItemOpenErrorLog;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem copySelectedLineToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem copyAllLinesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem base64DecoderToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItemOpenFileBackup;
        private System.Windows.Forms.ToolStripMenuItem structuredStorageViewerToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem excelSheetViewerToolStripMenuItem;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton toolStripButtonViewContents;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelFilePath;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel3;
        private System.Windows.Forms.ToolStripButton toolStripButtonFixCorruptDoc;
        private System.Windows.Forms.ToolStripButton toolStripButtonFixDoc;
        private System.Windows.Forms.ToolStripButton toolStripButtonModify;
        private System.Windows.Forms.ToolStripButton toolStripButtonSave;
        private System.Windows.Forms.ToolStripButton toolStripButtonInsertIcon;
        private System.Windows.Forms.ToolStripButton toolStripButtonValidateXml;
        private System.Windows.Forms.ToolStripMenuItem editToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem editToolStripMenuFindReplace;
        private System.Windows.Forms.ToolStripMenuItem editToolStripMenuItemModifyContents;
        private System.Windows.Forms.ToolStripMenuItem editToolStripMenuItemRemoveCustomDocProps;
        private System.Windows.Forms.ToolStripMenuItem editToolStripMenuItemRemoveCustomXml;
        private System.Windows.Forms.ToolStripDropDownButton toolStripDropDownButtonInsert;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.TreeView tvFiles;
        private System.Windows.Forms.RichTextBox rtbDisplay;
        private System.Windows.Forms.ToolStripButton toolStripButtonGenerateCallback;
        private System.Windows.Forms.ToolStripMenuItem office2010CustomUIPartToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem office2007CustomUIPartToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem customOutspaceToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem customTabToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem excelCustomTabToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem repurposeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem wordGroupOnInsertTabToolStripMenuItem;
        private System.Windows.Forms.ImageList tvImageList;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelDocType;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItemClose;
    }
}

