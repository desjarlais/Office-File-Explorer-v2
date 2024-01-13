
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
            toolStripMenuItemMRU = new System.Windows.Forms.ToolStripMenuItem();
            mruToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            mruToolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            mruToolStripMenuItem3 = new System.Windows.Forms.ToolStripMenuItem();
            mruToolStripMenuItem4 = new System.Windows.Forms.ToolStripMenuItem();
            mruToolStripMenuItem5 = new System.Windows.Forms.ToolStripMenuItem();
            mruToolStripMenuItem6 = new System.Windows.Forms.ToolStripMenuItem();
            mruToolStripMenuItem7 = new System.Windows.Forms.ToolStripMenuItem();
            mruToolStripMenuItem8 = new System.Windows.Forms.ToolStripMenuItem();
            mruToolStripMenuItem9 = new System.Windows.Forms.ToolStripMenuItem();
            fileToolStripMenuItemClose = new System.Windows.Forms.ToolStripMenuItem();
            toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            fileToolStripMenuItemSettings = new System.Windows.Forms.ToolStripMenuItem();
            fileToolStripMenuItemExit = new System.Windows.Forms.ToolStripMenuItem();
            editToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            editToolStripMenuFindReplace = new System.Windows.Forms.ToolStripMenuItem();
            editToolStripMenuItemModifyContents = new System.Windows.Forms.ToolStripMenuItem();
            editToolStripMenuItemRemoveCustomDocProps = new System.Windows.Forms.ToolStripMenuItem();
            editToolStripMenuItemRemoveCustomXml = new System.Windows.Forms.ToolStripMenuItem();
            wordDocumentRevisionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            toolsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            batchFileProcessingToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            clipboardViewerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            base64DecoderToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            excelSheetViewerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            feedbackToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            openErrorLogToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            contextMenuRichTextBox = new System.Windows.Forms.ContextMenuStrip(components);
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
            toolStripButtonFixXml = new System.Windows.Forms.ToolStripButton();
            toolStripButtonValidateXml = new System.Windows.Forms.ToolStripButton();
            toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            statusStrip1 = new System.Windows.Forms.StatusStrip();
            toolStripStatusLabel3 = new System.Windows.Forms.ToolStripStatusLabel();
            toolStripStatusLabelDocType = new System.Windows.Forms.ToolStripStatusLabel();
            toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            toolStripStatusLabelFilePath = new System.Windows.Forms.ToolStripStatusLabel();
            splitContainer1 = new System.Windows.Forms.SplitContainer();
            TvFiles = new System.Windows.Forms.TreeView();
            contextMenuTreeView = new System.Windows.Forms.ContextMenuStrip(components);
            viewPartPropertiesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            deletePartToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            extractPartToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            tvImageList = new System.Windows.Forms.ImageList(components);
            scintilla1 = new ScintillaNET.Scintilla();
            mnuMainMenu.SuspendLayout();
            contextMenuRichTextBox.SuspendLayout();
            toolStrip1.SuspendLayout();
            statusStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)splitContainer1).BeginInit();
            splitContainer1.Panel1.SuspendLayout();
            splitContainer1.Panel2.SuspendLayout();
            splitContainer1.SuspendLayout();
            contextMenuTreeView.SuspendLayout();
            SuspendLayout();
            // 
            // mnuMainMenu
            // 
            mnuMainMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { fileToolStripMenuItem, editToolStripMenuItem, toolsToolStripMenuItem, helpToolStripMenuItem });
            mnuMainMenu.Location = new System.Drawing.Point(0, 0);
            mnuMainMenu.Name = "mnuMainMenu";
            mnuMainMenu.Size = new System.Drawing.Size(1051, 24);
            mnuMainMenu.TabIndex = 0;
            mnuMainMenu.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { fileToolStripMenuItemOpen, toolStripMenuItemMRU, fileToolStripMenuItemClose, toolStripSeparator1, fileToolStripMenuItemSettings, fileToolStripMenuItemExit });
            fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            fileToolStripMenuItem.Text = "&File";
            // 
            // fileToolStripMenuItemOpen
            // 
            fileToolStripMenuItemOpen.Image = Properties.Resources.OpenFile;
            fileToolStripMenuItemOpen.Name = "fileToolStripMenuItemOpen";
            fileToolStripMenuItemOpen.Size = new System.Drawing.Size(136, 22);
            fileToolStripMenuItemOpen.Text = "O&pen";
            fileToolStripMenuItemOpen.Click += OpenToolStripMenuItem_Click;
            // 
            // toolStripMenuItemMRU
            // 
            toolStripMenuItemMRU.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { mruToolStripMenuItem1, mruToolStripMenuItem2, mruToolStripMenuItem3, mruToolStripMenuItem4, mruToolStripMenuItem5, mruToolStripMenuItem6, mruToolStripMenuItem7, mruToolStripMenuItem8, mruToolStripMenuItem9 });
            toolStripMenuItemMRU.Image = (System.Drawing.Image)resources.GetObject("toolStripMenuItemMRU.Image");
            toolStripMenuItemMRU.Name = "toolStripMenuItemMRU";
            toolStripMenuItemMRU.Size = new System.Drawing.Size(136, 22);
            toolStripMenuItemMRU.Text = "Recent Files";
            // 
            // mruToolStripMenuItem1
            // 
            mruToolStripMenuItem1.Name = "mruToolStripMenuItem1";
            mruToolStripMenuItem1.Size = new System.Drawing.Size(108, 22);
            mruToolStripMenuItem1.Text = "empty";
            mruToolStripMenuItem1.Click += MruToolStripMenuItem1_Click;
            // 
            // mruToolStripMenuItem2
            // 
            mruToolStripMenuItem2.Name = "mruToolStripMenuItem2";
            mruToolStripMenuItem2.Size = new System.Drawing.Size(108, 22);
            mruToolStripMenuItem2.Text = "empty";
            mruToolStripMenuItem2.Click += MruToolStripMenuItem2_Click;
            // 
            // mruToolStripMenuItem3
            // 
            mruToolStripMenuItem3.Name = "mruToolStripMenuItem3";
            mruToolStripMenuItem3.Size = new System.Drawing.Size(108, 22);
            mruToolStripMenuItem3.Text = "empty";
            mruToolStripMenuItem3.Click += MruToolStripMenuItem3_Click;
            // 
            // mruToolStripMenuItem4
            // 
            mruToolStripMenuItem4.Name = "mruToolStripMenuItem4";
            mruToolStripMenuItem4.Size = new System.Drawing.Size(108, 22);
            mruToolStripMenuItem4.Text = "empty";
            mruToolStripMenuItem4.Click += MruToolStripMenuItem4_Click;
            // 
            // mruToolStripMenuItem5
            // 
            mruToolStripMenuItem5.Name = "mruToolStripMenuItem5";
            mruToolStripMenuItem5.Size = new System.Drawing.Size(108, 22);
            mruToolStripMenuItem5.Text = "empty";
            mruToolStripMenuItem5.Click += MruToolStripMenuItem5_Click;
            // 
            // mruToolStripMenuItem6
            // 
            mruToolStripMenuItem6.Name = "mruToolStripMenuItem6";
            mruToolStripMenuItem6.Size = new System.Drawing.Size(108, 22);
            mruToolStripMenuItem6.Text = "empty";
            mruToolStripMenuItem6.Click += MruToolStripMenuItem6_Click;
            // 
            // mruToolStripMenuItem7
            // 
            mruToolStripMenuItem7.Name = "mruToolStripMenuItem7";
            mruToolStripMenuItem7.Size = new System.Drawing.Size(108, 22);
            mruToolStripMenuItem7.Text = "empty";
            mruToolStripMenuItem7.Click += MruToolStripMenuItem7_Click;
            // 
            // mruToolStripMenuItem8
            // 
            mruToolStripMenuItem8.Name = "mruToolStripMenuItem8";
            mruToolStripMenuItem8.Size = new System.Drawing.Size(108, 22);
            mruToolStripMenuItem8.Text = "empty";
            mruToolStripMenuItem8.Click += MruToolStripMenuItem8_Click;
            // 
            // mruToolStripMenuItem9
            // 
            mruToolStripMenuItem9.Name = "mruToolStripMenuItem9";
            mruToolStripMenuItem9.Size = new System.Drawing.Size(108, 22);
            mruToolStripMenuItem9.Text = "empty";
            mruToolStripMenuItem9.Click += MruToolStripMenuItem9_Click;
            // 
            // fileToolStripMenuItemClose
            // 
            fileToolStripMenuItemClose.Enabled = false;
            fileToolStripMenuItemClose.Image = Properties.Resources.CloseDocument;
            fileToolStripMenuItemClose.Name = "fileToolStripMenuItemClose";
            fileToolStripMenuItemClose.Size = new System.Drawing.Size(136, 22);
            fileToolStripMenuItemClose.Text = "&Close";
            fileToolStripMenuItemClose.Click += FileToolStripMenuItemClose_Click;
            // 
            // toolStripSeparator1
            // 
            toolStripSeparator1.Name = "toolStripSeparator1";
            toolStripSeparator1.Size = new System.Drawing.Size(133, 6);
            // 
            // fileToolStripMenuItemSettings
            // 
            fileToolStripMenuItemSettings.Image = Properties.Resources.Settings;
            fileToolStripMenuItemSettings.Name = "fileToolStripMenuItemSettings";
            fileToolStripMenuItemSettings.Size = new System.Drawing.Size(136, 22);
            fileToolStripMenuItemSettings.Text = "&Settings";
            fileToolStripMenuItemSettings.Click += SettingsToolStripMenuItem_Click;
            // 
            // fileToolStripMenuItemExit
            // 
            fileToolStripMenuItemExit.Name = "fileToolStripMenuItemExit";
            fileToolStripMenuItemExit.Size = new System.Drawing.Size(136, 22);
            fileToolStripMenuItemExit.Text = "E&xit";
            fileToolStripMenuItemExit.Click += ExitToolStripMenuItem_Click;
            // 
            // editToolStripMenuItem
            // 
            editToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { editToolStripMenuFindReplace, editToolStripMenuItemModifyContents, editToolStripMenuItemRemoveCustomDocProps, editToolStripMenuItemRemoveCustomXml, wordDocumentRevisionsToolStripMenuItem });
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
            editToolStripMenuFindReplace.Text = "Search and Replace";
            editToolStripMenuFindReplace.Click += EditToolStripMenuFindReplace_Click;
            // 
            // editToolStripMenuItemModifyContents
            // 
            editToolStripMenuItemModifyContents.Enabled = false;
            editToolStripMenuItemModifyContents.Image = Properties.Resources.Edit;
            editToolStripMenuItemModifyContents.Name = "editToolStripMenuItemModifyContents";
            editToolStripMenuItemModifyContents.Size = new System.Drawing.Size(277, 22);
            editToolStripMenuItemModifyContents.Text = "File Contents";
            editToolStripMenuItemModifyContents.Click += EditToolStripMenuItemModifyContents_Click;
            // 
            // editToolStripMenuItemRemoveCustomDocProps
            // 
            editToolStripMenuItemRemoveCustomDocProps.Enabled = false;
            editToolStripMenuItemRemoveCustomDocProps.Image = (System.Drawing.Image)resources.GetObject("editToolStripMenuItemRemoveCustomDocProps.Image");
            editToolStripMenuItemRemoveCustomDocProps.Name = "editToolStripMenuItemRemoveCustomDocProps";
            editToolStripMenuItemRemoveCustomDocProps.Size = new System.Drawing.Size(277, 22);
            editToolStripMenuItemRemoveCustomDocProps.Text = "Remove Custom Document Properties";
            editToolStripMenuItemRemoveCustomDocProps.Click += EditToolStripMenuItemRemoveCustomDocProps_Click;
            // 
            // editToolStripMenuItemRemoveCustomXml
            // 
            editToolStripMenuItemRemoveCustomXml.Enabled = false;
            editToolStripMenuItemRemoveCustomXml.Image = Properties.Resources.DeleteTag;
            editToolStripMenuItemRemoveCustomXml.Name = "editToolStripMenuItemRemoveCustomXml";
            editToolStripMenuItemRemoveCustomXml.Size = new System.Drawing.Size(277, 22);
            editToolStripMenuItemRemoveCustomXml.Text = "Remove Custom Xml";
            editToolStripMenuItemRemoveCustomXml.Click += EditToolStripMenuItemRemoveCustomXml_Click;
            // 
            // wordDocumentRevisionsToolStripMenuItem
            // 
            wordDocumentRevisionsToolStripMenuItem.Enabled = false;
            wordDocumentRevisionsToolStripMenuItem.Image = Properties.Resources.word;
            wordDocumentRevisionsToolStripMenuItem.Name = "wordDocumentRevisionsToolStripMenuItem";
            wordDocumentRevisionsToolStripMenuItem.Size = new System.Drawing.Size(277, 22);
            wordDocumentRevisionsToolStripMenuItem.Text = "Word Document Revisions";
            wordDocumentRevisionsToolStripMenuItem.Click += WordDocumentRevisionsToolStripMenuItem_Click;
            // 
            // toolsToolStripMenuItem
            // 
            toolsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { batchFileProcessingToolStripMenuItem, clipboardViewerToolStripMenuItem, base64DecoderToolStripMenuItem, excelSheetViewerToolStripMenuItem });
            toolsToolStripMenuItem.Name = "toolsToolStripMenuItem";
            toolsToolStripMenuItem.Size = new System.Drawing.Size(46, 20);
            toolsToolStripMenuItem.Text = "&Tools";
            // 
            // batchFileProcessingToolStripMenuItem
            // 
            batchFileProcessingToolStripMenuItem.Image = Properties.Resources.OpenDocumentGroup;
            batchFileProcessingToolStripMenuItem.Name = "batchFileProcessingToolStripMenuItem";
            batchFileProcessingToolStripMenuItem.Size = new System.Drawing.Size(185, 22);
            batchFileProcessingToolStripMenuItem.Text = "Batch File Processing";
            batchFileProcessingToolStripMenuItem.Click += BatchFileProcessingToolStripMenuItem_Click;
            // 
            // clipboardViewerToolStripMenuItem
            // 
            clipboardViewerToolStripMenuItem.Image = Properties.Resources.Copy;
            clipboardViewerToolStripMenuItem.Name = "clipboardViewerToolStripMenuItem";
            clipboardViewerToolStripMenuItem.Size = new System.Drawing.Size(185, 22);
            clipboardViewerToolStripMenuItem.Text = "Clipboard Viewer";
            clipboardViewerToolStripMenuItem.Click += ClipboardViewerToolStripMenuItem_Click;
            // 
            // base64DecoderToolStripMenuItem
            // 
            base64DecoderToolStripMenuItem.Image = Properties.Resources.Strings_7959_0;
            base64DecoderToolStripMenuItem.Name = "base64DecoderToolStripMenuItem";
            base64DecoderToolStripMenuItem.Size = new System.Drawing.Size(185, 22);
            base64DecoderToolStripMenuItem.Text = "Base64 Decoder";
            base64DecoderToolStripMenuItem.Click += Base64DecoderToolStripMenuItem_Click;
            // 
            // excelSheetViewerToolStripMenuItem
            // 
            excelSheetViewerToolStripMenuItem.Enabled = false;
            excelSheetViewerToolStripMenuItem.Image = (System.Drawing.Image)resources.GetObject("excelSheetViewerToolStripMenuItem.Image");
            excelSheetViewerToolStripMenuItem.Name = "excelSheetViewerToolStripMenuItem";
            excelSheetViewerToolStripMenuItem.Size = new System.Drawing.Size(185, 22);
            excelSheetViewerToolStripMenuItem.Text = "Excel Sheet Viewer";
            excelSheetViewerToolStripMenuItem.Click += excelSheetViewerToolStripMenuItem_Click;
            // 
            // helpToolStripMenuItem
            // 
            helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { aboutToolStripMenuItem, feedbackToolStripMenuItem, openErrorLogToolStripMenuItem });
            helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            helpToolStripMenuItem.Text = "&Help";
            // 
            // aboutToolStripMenuItem
            // 
            aboutToolStripMenuItem.Image = Properties.Resources.AboutBox;
            aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            aboutToolStripMenuItem.Size = new System.Drawing.Size(154, 22);
            aboutToolStripMenuItem.Text = "About";
            aboutToolStripMenuItem.Click += AboutToolStripMenuItem_Click;
            // 
            // feedbackToolStripMenuItem
            // 
            feedbackToolStripMenuItem.Image = Properties.Resources.Feedback;
            feedbackToolStripMenuItem.Name = "feedbackToolStripMenuItem";
            feedbackToolStripMenuItem.Size = new System.Drawing.Size(154, 22);
            feedbackToolStripMenuItem.Text = "Feedback";
            feedbackToolStripMenuItem.Click += FeedbackToolStripMenuItem_Click;
            // 
            // openErrorLogToolStripMenuItem
            // 
            openErrorLogToolStripMenuItem.Image = Properties.Resources.ErrorSummary;
            openErrorLogToolStripMenuItem.Name = "openErrorLogToolStripMenuItem";
            openErrorLogToolStripMenuItem.Size = new System.Drawing.Size(154, 22);
            openErrorLogToolStripMenuItem.Text = "Open Error Log";
            openErrorLogToolStripMenuItem.Click += OpenErrorLogToolStripMenuItem_Click_1;
            // 
            // contextMenuRichTextBox
            // 
            contextMenuRichTextBox.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { copySelectedLineToolStripMenuItem, copyAllLinesToolStripMenuItem });
            contextMenuRichTextBox.Name = "contextMenuStrip1";
            contextMenuRichTextBox.Size = new System.Drawing.Size(175, 48);
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
            toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { toolStripButtonViewContents, toolStripButtonFixCorruptDoc, toolStripButtonFixDoc, toolStripSeparator2, toolStripButtonModify, toolStripButtonSave, toolStripButtonInsertIcon, toolStripButtonGenerateCallback, toolStripDropDownButtonInsert, toolStripButtonFixXml, toolStripButtonValidateXml, toolStripSeparator3 });
            toolStrip1.Location = new System.Drawing.Point(0, 24);
            toolStrip1.Name = "toolStrip1";
            toolStrip1.Size = new System.Drawing.Size(1051, 25);
            toolStrip1.TabIndex = 3;
            toolStrip1.Text = "toolStrip1";
            // 
            // toolStripButtonViewContents
            // 
            toolStripButtonViewContents.Enabled = false;
            toolStripButtonViewContents.Image = (System.Drawing.Image)resources.GetObject("toolStripButtonViewContents.Image");
            toolStripButtonViewContents.ImageTransparentColor = System.Drawing.Color.Magenta;
            toolStripButtonViewContents.Name = "toolStripButtonViewContents";
            toolStripButtonViewContents.Size = new System.Drawing.Size(103, 22);
            toolStripButtonViewContents.Text = "View Contents";
            toolStripButtonViewContents.Click += ToolStripButtonViewContents_Click;
            // 
            // toolStripButtonFixCorruptDoc
            // 
            toolStripButtonFixCorruptDoc.Enabled = false;
            toolStripButtonFixCorruptDoc.Image = (System.Drawing.Image)resources.GetObject("toolStripButtonFixCorruptDoc.Image");
            toolStripButtonFixCorruptDoc.ImageTransparentColor = System.Drawing.Color.Magenta;
            toolStripButtonFixCorruptDoc.Name = "toolStripButtonFixCorruptDoc";
            toolStripButtonFixCorruptDoc.Size = new System.Drawing.Size(145, 22);
            toolStripButtonFixCorruptDoc.Text = "Fix Corrupt Document";
            toolStripButtonFixCorruptDoc.Click += ToolStripButtonFixCorruptDoc_Click;
            // 
            // toolStripButtonFixDoc
            // 
            toolStripButtonFixDoc.Enabled = false;
            toolStripButtonFixDoc.Image = (System.Drawing.Image)resources.GetObject("toolStripButtonFixDoc.Image");
            toolStripButtonFixDoc.ImageTransparentColor = System.Drawing.Color.Magenta;
            toolStripButtonFixDoc.Name = "toolStripButtonFixDoc";
            toolStripButtonFixDoc.Size = new System.Drawing.Size(101, 22);
            toolStripButtonFixDoc.Text = "Fix Document";
            toolStripButtonFixDoc.Click += ToolStripButtonFixDoc_Click;
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
            toolStripButtonModify.Click += ToolStripButtonModify_Click;
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
            toolStripButtonSave.Click += ToolStripButtonSave_Click;
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
            toolStripButtonInsertIcon.Click += ToolStripButtonInsertIcon_Click;
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
            toolStripButtonGenerateCallback.Click += ToolStripButtonGenerateCallback_Click;
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
            office2010CustomUIPartToolStripMenuItem.Click += Office2010CustomUIPartToolStripMenuItem_Click;
            // 
            // office2007CustomUIPartToolStripMenuItem
            // 
            office2007CustomUIPartToolStripMenuItem.Image = Properties.Resources._90_904004_pixel_ms_office_word_2007_logo;
            office2007CustomUIPartToolStripMenuItem.Name = "office2007CustomUIPartToolStripMenuItem";
            office2007CustomUIPartToolStripMenuItem.Size = new System.Drawing.Size(216, 22);
            office2007CustomUIPartToolStripMenuItem.Text = "Office 2007 Custom UI Part";
            office2007CustomUIPartToolStripMenuItem.Click += Office2007CustomUIPartToolStripMenuItem_Click;
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
            customOutspaceToolStripMenuItem.Click += CustomOutspaceToolStripMenuItem_Click;
            // 
            // customTabToolStripMenuItem
            // 
            customTabToolStripMenuItem.Name = "customTabToolStripMenuItem";
            customTabToolStripMenuItem.Size = new System.Drawing.Size(217, 22);
            customTabToolStripMenuItem.Text = "Custom Tab";
            customTabToolStripMenuItem.Click += CustomTabToolStripMenuItem_Click;
            // 
            // excelCustomTabToolStripMenuItem
            // 
            excelCustomTabToolStripMenuItem.Name = "excelCustomTabToolStripMenuItem";
            excelCustomTabToolStripMenuItem.Size = new System.Drawing.Size(217, 22);
            excelCustomTabToolStripMenuItem.Text = "Excel - Custom Tab";
            excelCustomTabToolStripMenuItem.Click += ExcelCustomTabToolStripMenuItem_Click;
            // 
            // repurposeToolStripMenuItem
            // 
            repurposeToolStripMenuItem.Name = "repurposeToolStripMenuItem";
            repurposeToolStripMenuItem.Size = new System.Drawing.Size(217, 22);
            repurposeToolStripMenuItem.Text = "Repurpose";
            repurposeToolStripMenuItem.Click += RepurposeToolStripMenuItem_Click;
            // 
            // wordGroupOnInsertTabToolStripMenuItem
            // 
            wordGroupOnInsertTabToolStripMenuItem.Name = "wordGroupOnInsertTabToolStripMenuItem";
            wordGroupOnInsertTabToolStripMenuItem.Size = new System.Drawing.Size(217, 22);
            wordGroupOnInsertTabToolStripMenuItem.Text = "Word - Group on Insert Tab";
            wordGroupOnInsertTabToolStripMenuItem.Click += WordGroupOnInsertTabToolStripMenuItem_Click;
            // 
            // toolStripButtonFixXml
            // 
            toolStripButtonFixXml.Enabled = false;
            toolStripButtonFixXml.Image = Properties.Resources.XmlFile;
            toolStripButtonFixXml.ImageTransparentColor = System.Drawing.Color.Magenta;
            toolStripButtonFixXml.Name = "toolStripButtonFixXml";
            toolStripButtonFixXml.Size = new System.Drawing.Size(66, 22);
            toolStripButtonFixXml.Text = "Fix Xml";
            toolStripButtonFixXml.Click += ToolStripButtonFixXml_Click;
            // 
            // toolStripButtonValidateXml
            // 
            toolStripButtonValidateXml.Enabled = false;
            toolStripButtonValidateXml.Image = (System.Drawing.Image)resources.GetObject("toolStripButtonValidateXml.Image");
            toolStripButtonValidateXml.ImageTransparentColor = System.Drawing.Color.Magenta;
            toolStripButtonValidateXml.Name = "toolStripButtonValidateXml";
            toolStripButtonValidateXml.Size = new System.Drawing.Size(92, 22);
            toolStripButtonValidateXml.Text = "Validate Xml";
            toolStripButtonValidateXml.Click += ToolStripButtonValidateXml_Click;
            // 
            // toolStripSeparator3
            // 
            toolStripSeparator3.Name = "toolStripSeparator3";
            toolStripSeparator3.Size = new System.Drawing.Size(6, 25);
            // 
            // statusStrip1
            // 
            statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { toolStripStatusLabel3, toolStripStatusLabelDocType, toolStripStatusLabel1, toolStripStatusLabelFilePath });
            statusStrip1.Location = new System.Drawing.Point(0, 555);
            statusStrip1.Name = "statusStrip1";
            statusStrip1.Size = new System.Drawing.Size(1051, 22);
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
            splitContainer1.Panel1.Controls.Add(TvFiles);
            // 
            // splitContainer1.Panel2
            // 
            splitContainer1.Panel2.Controls.Add(scintilla1);
            splitContainer1.Size = new System.Drawing.Size(1051, 506);
            splitContainer1.SplitterDistance = 349;
            splitContainer1.TabIndex = 5;
            // 
            // TvFiles
            // 
            TvFiles.ContextMenuStrip = contextMenuTreeView;
            TvFiles.Dock = System.Windows.Forms.DockStyle.Fill;
            TvFiles.ImageIndex = 0;
            TvFiles.ImageList = tvImageList;
            TvFiles.Location = new System.Drawing.Point(0, 0);
            TvFiles.Name = "TvFiles";
            TvFiles.SelectedImageIndex = 0;
            TvFiles.Size = new System.Drawing.Size(349, 506);
            TvFiles.TabIndex = 0;
            TvFiles.AfterSelect += TvFiles_AfterSelect;
            TvFiles.NodeMouseClick += TvFiles_NodeMouseClick;
            TvFiles.KeyUp += TvFiles_KeyUp;
            // 
            // contextMenuTreeView
            // 
            contextMenuTreeView.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { viewPartPropertiesToolStripMenuItem, deletePartToolStripMenuItem, extractPartToolStripMenuItem });
            contextMenuTreeView.Name = "contextMenuTreeView";
            contextMenuTreeView.Size = new System.Drawing.Size(180, 70);
            // 
            // viewPartPropertiesToolStripMenuItem
            // 
            viewPartPropertiesToolStripMenuItem.Name = "viewPartPropertiesToolStripMenuItem";
            viewPartPropertiesToolStripMenuItem.Size = new System.Drawing.Size(179, 22);
            viewPartPropertiesToolStripMenuItem.Text = "View Part Properties";
            viewPartPropertiesToolStripMenuItem.Click += ViewPartPropertiesToolStripMenuItem_Click;
            // 
            // deletePartToolStripMenuItem
            // 
            deletePartToolStripMenuItem.Enabled = false;
            deletePartToolStripMenuItem.Name = "deletePartToolStripMenuItem";
            deletePartToolStripMenuItem.Size = new System.Drawing.Size(179, 22);
            deletePartToolStripMenuItem.Text = "Delete Part";
            // 
            // extractPartToolStripMenuItem
            // 
            extractPartToolStripMenuItem.Enabled = false;
            extractPartToolStripMenuItem.Name = "extractPartToolStripMenuItem";
            extractPartToolStripMenuItem.Size = new System.Drawing.Size(179, 22);
            extractPartToolStripMenuItem.Text = "Extract Part";
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
            // scintilla1
            // 
            scintilla1.AutoCMaxHeight = 9;
            scintilla1.BiDirectionality = ScintillaNET.BiDirectionalDisplayType.Disabled;
            scintilla1.CaretLineVisible = true;
            scintilla1.Dock = System.Windows.Forms.DockStyle.Fill;
            scintilla1.LexerName = null;
            scintilla1.Location = new System.Drawing.Point(0, 0);
            scintilla1.Name = "scintilla1";
            scintilla1.ScrollWidth = 49;
            scintilla1.Size = new System.Drawing.Size(698, 506);
            scintilla1.TabIndents = true;
            scintilla1.TabIndex = 0;
            scintilla1.UseRightToLeftReadingLayout = false;
            scintilla1.WrapMode = ScintillaNET.WrapMode.None;
            // 
            // FrmMain
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(1051, 577);
            Controls.Add(splitContainer1);
            Controls.Add(statusStrip1);
            Controls.Add(toolStrip1);
            Controls.Add(mnuMainMenu);
            Icon = (System.Drawing.Icon)resources.GetObject("$this.Icon");
            MainMenuStrip = mnuMainMenu;
            MinimumSize = new System.Drawing.Size(920, 588);
            Name = "FrmMain";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text = "Office File Explorer v2";
            FormClosing += FrmMain_FormClosing;
            mnuMainMenu.ResumeLayout(false);
            mnuMainMenu.PerformLayout();
            contextMenuRichTextBox.ResumeLayout(false);
            toolStrip1.ResumeLayout(false);
            toolStrip1.PerformLayout();
            statusStrip1.ResumeLayout(false);
            statusStrip1.PerformLayout();
            splitContainer1.Panel1.ResumeLayout(false);
            splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)splitContainer1).EndInit();
            splitContainer1.ResumeLayout(false);
            contextMenuTreeView.ResumeLayout(false);
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
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ContextMenuStrip contextMenuRichTextBox;
        private System.Windows.Forms.ToolStripMenuItem copySelectedLineToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem copyAllLinesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem base64DecoderToolStripMenuItem;
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
        private System.Windows.Forms.TreeView TvFiles;
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
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemMRU;
        private System.Windows.Forms.ToolStripMenuItem mruToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem mruToolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem mruToolStripMenuItem3;
        private System.Windows.Forms.ToolStripMenuItem mruToolStripMenuItem4;
        private System.Windows.Forms.ToolStripMenuItem mruToolStripMenuItem5;
        private System.Windows.Forms.ToolStripMenuItem mruToolStripMenuItem6;
        private System.Windows.Forms.ToolStripMenuItem mruToolStripMenuItem7;
        private System.Windows.Forms.ToolStripMenuItem mruToolStripMenuItem8;
        private System.Windows.Forms.ToolStripMenuItem mruToolStripMenuItem9;
        private System.Windows.Forms.ToolStripMenuItem openErrorLogToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem wordDocumentRevisionsToolStripMenuItem;
        private System.Windows.Forms.ContextMenuStrip contextMenuTreeView;
        private System.Windows.Forms.ToolStripMenuItem viewPartPropertiesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem deletePartToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem extractPartToolStripMenuItem;
        private System.Windows.Forms.ToolStripButton toolStripButtonFixXml;
        private ScintillaNET.Scintilla scintilla1;
    }
}

