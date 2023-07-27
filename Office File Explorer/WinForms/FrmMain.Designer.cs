
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
            openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            openErrorLogToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            openFileBackupFolderToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            settingsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            toolsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            batchFileProcessingToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            clipboardViewerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            base64DecoderToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            structuredStorageViewerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            openXmlPartViewerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            excelSheetViewerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            feedbackToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            groupBox1 = new System.Windows.Forms.GroupBox();
            BtnRemoveCustomXmlParts = new System.Windows.Forms.Button();
            BtnRemoveCustomFileProps = new System.Windows.Forms.Button();
            BtnValidateDoc = new System.Windows.Forms.Button();
            BtnFixCorruptDoc = new System.Windows.Forms.Button();
            BtnDocProps = new System.Windows.Forms.Button();
            BtnViewImages = new System.Windows.Forms.Button();
            lblFileType = new System.Windows.Forms.Label();
            label2 = new System.Windows.Forms.Label();
            BtnSearchAndReplace = new System.Windows.Forms.Button();
            BtnFixDocument = new System.Windows.Forms.Button();
            lblFilePath = new System.Windows.Forms.Label();
            BtnModifyContent = new System.Windows.Forms.Button();
            label1 = new System.Windows.Forms.Label();
            BtnViewContents = new System.Windows.Forms.Button();
            groupBox2 = new System.Windows.Forms.GroupBox();
            LstDisplay = new System.Windows.Forms.ListBox();
            contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(components);
            copySelectedLineToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            copyAllLinesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            mnuMainMenu.SuspendLayout();
            groupBox1.SuspendLayout();
            groupBox2.SuspendLayout();
            contextMenuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // mnuMainMenu
            // 
            mnuMainMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { fileToolStripMenuItem, toolsToolStripMenuItem, helpToolStripMenuItem });
            mnuMainMenu.Location = new System.Drawing.Point(0, 0);
            mnuMainMenu.Name = "mnuMainMenu";
            mnuMainMenu.Size = new System.Drawing.Size(924, 24);
            mnuMainMenu.TabIndex = 0;
            mnuMainMenu.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { openToolStripMenuItem, openErrorLogToolStripMenuItem, openFileBackupFolderToolStripMenuItem, settingsToolStripMenuItem, toolStripSeparator1, exitToolStripMenuItem });
            fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            fileToolStripMenuItem.Text = "&File";
            // 
            // openToolStripMenuItem
            // 
            openToolStripMenuItem.Image = Properties.Resources.OpenFile;
            openToolStripMenuItem.Name = "openToolStripMenuItem";
            openToolStripMenuItem.Size = new System.Drawing.Size(202, 22);
            openToolStripMenuItem.Text = "O&pen";
            openToolStripMenuItem.Click += OpenToolStripMenuItem_Click;
            // 
            // openErrorLogToolStripMenuItem
            // 
            openErrorLogToolStripMenuItem.Image = Properties.Resources.ErrorSummary;
            openErrorLogToolStripMenuItem.Name = "openErrorLogToolStripMenuItem";
            openErrorLogToolStripMenuItem.Size = new System.Drawing.Size(202, 22);
            openErrorLogToolStripMenuItem.Text = "Open Error Log";
            openErrorLogToolStripMenuItem.Click += OpenErrorLogToolStripMenuItem_Click;
            // 
            // openFileBackupFolderToolStripMenuItem
            // 
            openFileBackupFolderToolStripMenuItem.Image = (System.Drawing.Image)resources.GetObject("openFileBackupFolderToolStripMenuItem.Image");
            openFileBackupFolderToolStripMenuItem.Name = "openFileBackupFolderToolStripMenuItem";
            openFileBackupFolderToolStripMenuItem.Size = new System.Drawing.Size(202, 22);
            openFileBackupFolderToolStripMenuItem.Text = "Open File Backup Folder";
            openFileBackupFolderToolStripMenuItem.Click += openFileBackupFolderToolStripMenuItem_Click;
            // 
            // settingsToolStripMenuItem
            // 
            settingsToolStripMenuItem.Image = Properties.Resources.Settings;
            settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
            settingsToolStripMenuItem.Size = new System.Drawing.Size(202, 22);
            settingsToolStripMenuItem.Text = "&Settings";
            settingsToolStripMenuItem.Click += SettingsToolStripMenuItem_Click;
            // 
            // toolStripSeparator1
            // 
            toolStripSeparator1.Name = "toolStripSeparator1";
            toolStripSeparator1.Size = new System.Drawing.Size(199, 6);
            // 
            // exitToolStripMenuItem
            // 
            exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            exitToolStripMenuItem.Size = new System.Drawing.Size(202, 22);
            exitToolStripMenuItem.Text = "E&xit";
            exitToolStripMenuItem.Click += ExitToolStripMenuItem_Click;
            // 
            // toolsToolStripMenuItem
            // 
            toolsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { batchFileProcessingToolStripMenuItem, clipboardViewerToolStripMenuItem, base64DecoderToolStripMenuItem, structuredStorageViewerToolStripMenuItem, openXmlPartViewerToolStripMenuItem, excelSheetViewerToolStripMenuItem });
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
            // openXmlPartViewerToolStripMenuItem
            // 
            openXmlPartViewerToolStripMenuItem.Enabled = false;
            openXmlPartViewerToolStripMenuItem.Image = Properties.Resources.XmlFile;
            openXmlPartViewerToolStripMenuItem.Name = "openXmlPartViewerToolStripMenuItem";
            openXmlPartViewerToolStripMenuItem.Size = new System.Drawing.Size(210, 22);
            openXmlPartViewerToolStripMenuItem.Text = "Open Xml Part Viewer";
            openXmlPartViewerToolStripMenuItem.Click += openXmlPartViewerToolStripMenuItem_Click;
            // 
            // excelSheetViewerToolStripMenuItem
            // 
            excelSheetViewerToolStripMenuItem.Enabled = false;
            excelSheetViewerToolStripMenuItem.Image = Properties.Resources.excel;
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
            // groupBox1
            // 
            groupBox1.Controls.Add(BtnRemoveCustomXmlParts);
            groupBox1.Controls.Add(BtnRemoveCustomFileProps);
            groupBox1.Controls.Add(BtnValidateDoc);
            groupBox1.Controls.Add(BtnFixCorruptDoc);
            groupBox1.Controls.Add(BtnDocProps);
            groupBox1.Controls.Add(BtnViewImages);
            groupBox1.Controls.Add(lblFileType);
            groupBox1.Controls.Add(label2);
            groupBox1.Controls.Add(BtnSearchAndReplace);
            groupBox1.Controls.Add(BtnFixDocument);
            groupBox1.Controls.Add(lblFilePath);
            groupBox1.Controls.Add(BtnModifyContent);
            groupBox1.Controls.Add(label1);
            groupBox1.Controls.Add(BtnViewContents);
            groupBox1.Location = new System.Drawing.Point(3, 27);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new System.Drawing.Size(924, 136);
            groupBox1.TabIndex = 2;
            groupBox1.TabStop = false;
            groupBox1.Text = "Document Information";
            // 
            // BtnRemoveCustomXmlParts
            // 
            BtnRemoveCustomXmlParts.Location = new System.Drawing.Point(398, 107);
            BtnRemoveCustomXmlParts.Name = "BtnRemoveCustomXmlParts";
            BtnRemoveCustomXmlParts.Size = new System.Drawing.Size(134, 23);
            BtnRemoveCustomXmlParts.TabIndex = 13;
            BtnRemoveCustomXmlParts.Text = "Remove Custom Xml";
            BtnRemoveCustomXmlParts.UseVisualStyleBackColor = true;
            BtnRemoveCustomXmlParts.Click += BtnRemoveCustomXmlParts_Click;
            // 
            // BtnRemoveCustomFileProps
            // 
            BtnRemoveCustomFileProps.Location = new System.Drawing.Point(538, 107);
            BtnRemoveCustomFileProps.Name = "BtnRemoveCustomFileProps";
            BtnRemoveCustomFileProps.Size = new System.Drawing.Size(157, 23);
            BtnRemoveCustomFileProps.TabIndex = 12;
            BtnRemoveCustomFileProps.Text = "Remove Custom File Props";
            BtnRemoveCustomFileProps.UseVisualStyleBackColor = true;
            BtnRemoveCustomFileProps.Click += BtnRemoveCustomFileProps_Click;
            // 
            // BtnValidateDoc
            // 
            BtnValidateDoc.Location = new System.Drawing.Point(538, 78);
            BtnValidateDoc.Name = "BtnValidateDoc";
            BtnValidateDoc.Size = new System.Drawing.Size(135, 23);
            BtnValidateDoc.TabIndex = 10;
            BtnValidateDoc.Text = "Validate Document";
            BtnValidateDoc.UseVisualStyleBackColor = true;
            BtnValidateDoc.Click += BtnValidateDoc_Click;
            // 
            // BtnFixCorruptDoc
            // 
            BtnFixCorruptDoc.Location = new System.Drawing.Point(233, 107);
            BtnFixCorruptDoc.Name = "BtnFixCorruptDoc";
            BtnFixCorruptDoc.Size = new System.Drawing.Size(159, 23);
            BtnFixCorruptDoc.TabIndex = 9;
            BtnFixCorruptDoc.Text = "Fix Corrupt Document";
            BtnFixCorruptDoc.UseVisualStyleBackColor = true;
            BtnFixCorruptDoc.Click += BtnFixCorruptDoc_Click;
            // 
            // BtnDocProps
            // 
            BtnDocProps.Location = new System.Drawing.Point(232, 78);
            BtnDocProps.Name = "BtnDocProps";
            BtnDocProps.Size = new System.Drawing.Size(168, 23);
            BtnDocProps.TabIndex = 7;
            BtnDocProps.Text = "View Document Properties";
            BtnDocProps.UseVisualStyleBackColor = true;
            BtnDocProps.Click += BtnDocProps_Click;
            // 
            // BtnViewImages
            // 
            BtnViewImages.Location = new System.Drawing.Point(119, 78);
            BtnViewImages.Name = "BtnViewImages";
            BtnViewImages.Size = new System.Drawing.Size(107, 23);
            BtnViewImages.TabIndex = 6;
            BtnViewImages.Text = "View Images";
            BtnViewImages.UseVisualStyleBackColor = true;
            BtnViewImages.Click += BtnViewImages_Click;
            // 
            // lblFileType
            // 
            lblFileType.AutoSize = true;
            lblFileType.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            lblFileType.Location = new System.Drawing.Point(46, 46);
            lblFileType.Name = "lblFileType";
            lblFileType.Size = new System.Drawing.Size(2, 17);
            lblFileType.TabIndex = 3;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new System.Drawing.Point(3, 46);
            label2.Name = "label2";
            label2.Size = new System.Drawing.Size(37, 15);
            label2.TabIndex = 2;
            label2.Text = "Type: ";
            // 
            // BtnSearchAndReplace
            // 
            BtnSearchAndReplace.Location = new System.Drawing.Point(679, 78);
            BtnSearchAndReplace.Name = "BtnSearchAndReplace";
            BtnSearchAndReplace.Size = new System.Drawing.Size(133, 23);
            BtnSearchAndReplace.TabIndex = 1;
            BtnSearchAndReplace.Text = "Search and Replace";
            BtnSearchAndReplace.UseVisualStyleBackColor = true;
            BtnSearchAndReplace.Click += BtnSearchAndReplace_Click;
            // 
            // BtnFixDocument
            // 
            BtnFixDocument.Location = new System.Drawing.Point(119, 107);
            BtnFixDocument.Name = "BtnFixDocument";
            BtnFixDocument.Size = new System.Drawing.Size(108, 23);
            BtnFixDocument.TabIndex = 3;
            BtnFixDocument.Text = "Fix Document";
            BtnFixDocument.UseVisualStyleBackColor = true;
            BtnFixDocument.Click += BtnFixDocument_Click;
            // 
            // lblFilePath
            // 
            lblFilePath.AutoSize = true;
            lblFilePath.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            lblFilePath.Location = new System.Drawing.Point(46, 19);
            lblFilePath.Name = "lblFilePath";
            lblFilePath.Size = new System.Drawing.Size(2, 17);
            lblFilePath.TabIndex = 1;
            // 
            // BtnModifyContent
            // 
            BtnModifyContent.Location = new System.Drawing.Point(6, 107);
            BtnModifyContent.Name = "BtnModifyContent";
            BtnModifyContent.Size = new System.Drawing.Size(107, 23);
            BtnModifyContent.TabIndex = 2;
            BtnModifyContent.Text = "Modify Contents";
            BtnModifyContent.UseVisualStyleBackColor = true;
            BtnModifyContent.Click += BtnModifyContent_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point(3, 19);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(28, 15);
            label1.TabIndex = 0;
            label1.Text = "File:";
            // 
            // BtnViewContents
            // 
            BtnViewContents.Location = new System.Drawing.Point(6, 78);
            BtnViewContents.Name = "BtnViewContents";
            BtnViewContents.Size = new System.Drawing.Size(107, 23);
            BtnViewContents.TabIndex = 0;
            BtnViewContents.Text = "View Contents";
            BtnViewContents.UseVisualStyleBackColor = true;
            BtnViewContents.Click += BtnViewContents_Click;
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(LstDisplay);
            groupBox2.Location = new System.Drawing.Point(0, 169);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new System.Drawing.Size(924, 389);
            groupBox2.TabIndex = 3;
            groupBox2.TabStop = false;
            groupBox2.Text = "Output";
            // 
            // LstDisplay
            // 
            LstDisplay.ContextMenuStrip = contextMenuStrip1;
            LstDisplay.Dock = System.Windows.Forms.DockStyle.Fill;
            LstDisplay.FormattingEnabled = true;
            LstDisplay.HorizontalScrollbar = true;
            LstDisplay.ItemHeight = 15;
            LstDisplay.Location = new System.Drawing.Point(3, 19);
            LstDisplay.Name = "LstDisplay";
            LstDisplay.Size = new System.Drawing.Size(918, 367);
            LstDisplay.TabIndex = 0;
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
            // FrmMain
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(924, 563);
            Controls.Add(groupBox2);
            Controls.Add(groupBox1);
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
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            groupBox2.ResumeLayout(false);
            contextMenuStrip1.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private System.Windows.Forms.MenuStrip mnuMainMenu;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label lblFileType;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblFilePath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ListBox LstDisplay;
        private System.Windows.Forms.Button BtnFixDocument;
        private System.Windows.Forms.Button BtnModifyContent;
        private System.Windows.Forms.Button BtnSearchAndReplace;
        private System.Windows.Forms.Button BtnViewContents;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem settingsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem batchFileProcessingToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem clipboardViewerToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem feedbackToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openErrorLogToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.Button BtnDocProps;
        private System.Windows.Forms.Button BtnViewImages;
        private System.Windows.Forms.Button BtnFixCorruptDoc;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem copySelectedLineToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem copyAllLinesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem base64DecoderToolStripMenuItem;
        private System.Windows.Forms.Button BtnValidateDoc;
        private System.Windows.Forms.ToolStripMenuItem openFileBackupFolderToolStripMenuItem;
        private System.Windows.Forms.Button BtnRemoveCustomFileProps;
        private System.Windows.Forms.Button BtnRemoveCustomXmlParts;
        private System.Windows.Forms.ToolStripMenuItem structuredStorageViewerToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openXmlPartViewerToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem excelSheetViewerToolStripMenuItem;
    }
}

