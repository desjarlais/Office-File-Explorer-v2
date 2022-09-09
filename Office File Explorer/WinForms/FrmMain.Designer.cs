
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmMain));
            this.mnuMainMenu = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openErrorLogToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openFileBackupFolderToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.settingsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.batchFileProcessingToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.clipboardViewerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.base64DecoderToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.feedbackToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.BtnRemoveCustomXmlParts = new System.Windows.Forms.Button();
            this.BtnRemoveCustomFileProps = new System.Windows.Forms.Button();
            this.BtnExcelSheetViewer = new System.Windows.Forms.Button();
            this.BtnValidateDoc = new System.Windows.Forms.Button();
            this.BtnViewCustomUI = new System.Windows.Forms.Button();
            this.BtnFixCorruptDoc = new System.Windows.Forms.Button();
            this.BtnCustomXml = new System.Windows.Forms.Button();
            this.BtnDocProps = new System.Windows.Forms.Button();
            this.BtnViewImages = new System.Windows.Forms.Button();
            this.lblFileType = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.BtnSearchAndReplace = new System.Windows.Forms.Button();
            this.BtnFixDocument = new System.Windows.Forms.Button();
            this.lblFilePath = new System.Windows.Forms.Label();
            this.BtnModifyContent = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.BtnViewContents = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.LstDisplay = new System.Windows.Forms.ListBox();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.copySelectedLineToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.copyAllLinesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.mnuMainMenu.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // mnuMainMenu
            // 
            this.mnuMainMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.toolsToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.mnuMainMenu.Location = new System.Drawing.Point(0, 0);
            this.mnuMainMenu.Name = "mnuMainMenu";
            this.mnuMainMenu.Size = new System.Drawing.Size(924, 24);
            this.mnuMainMenu.TabIndex = 0;
            this.mnuMainMenu.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            this.openErrorLogToolStripMenuItem,
            this.openFileBackupFolderToolStripMenuItem,
            this.settingsToolStripMenuItem,
            this.toolStripSeparator1,
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "&File";
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("openToolStripMenuItem.Image")));
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.Size = new System.Drawing.Size(202, 22);
            this.openToolStripMenuItem.Text = "O&pen";
            this.openToolStripMenuItem.Click += new System.EventHandler(this.OpenToolStripMenuItem_Click);
            // 
            // openErrorLogToolStripMenuItem
            // 
            this.openErrorLogToolStripMenuItem.Image = global::Office_File_Explorer.Properties.Resources.ErrorSummary_16x;
            this.openErrorLogToolStripMenuItem.Name = "openErrorLogToolStripMenuItem";
            this.openErrorLogToolStripMenuItem.Size = new System.Drawing.Size(202, 22);
            this.openErrorLogToolStripMenuItem.Text = "Open Error Log";
            this.openErrorLogToolStripMenuItem.Click += new System.EventHandler(this.OpenErrorLogToolStripMenuItem_Click);
            // 
            // openFileBackupFolderToolStripMenuItem
            // 
            this.openFileBackupFolderToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("openFileBackupFolderToolStripMenuItem.Image")));
            this.openFileBackupFolderToolStripMenuItem.Name = "openFileBackupFolderToolStripMenuItem";
            this.openFileBackupFolderToolStripMenuItem.Size = new System.Drawing.Size(202, 22);
            this.openFileBackupFolderToolStripMenuItem.Text = "Open File Backup Folder";
            this.openFileBackupFolderToolStripMenuItem.Click += new System.EventHandler(this.openFileBackupFolderToolStripMenuItem_Click);
            // 
            // settingsToolStripMenuItem
            // 
            this.settingsToolStripMenuItem.Image = global::Office_File_Explorer.Properties.Resources.Settings_16x;
            this.settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
            this.settingsToolStripMenuItem.Size = new System.Drawing.Size(202, 22);
            this.settingsToolStripMenuItem.Text = "&Settings";
            this.settingsToolStripMenuItem.Click += new System.EventHandler(this.SettingsToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(199, 6);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(202, 22);
            this.exitToolStripMenuItem.Text = "E&xit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.ExitToolStripMenuItem_Click);
            // 
            // toolsToolStripMenuItem
            // 
            this.toolsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.batchFileProcessingToolStripMenuItem,
            this.clipboardViewerToolStripMenuItem,
            this.base64DecoderToolStripMenuItem});
            this.toolsToolStripMenuItem.Name = "toolsToolStripMenuItem";
            this.toolsToolStripMenuItem.Size = new System.Drawing.Size(46, 20);
            this.toolsToolStripMenuItem.Text = "&Tools";
            // 
            // batchFileProcessingToolStripMenuItem
            // 
            this.batchFileProcessingToolStripMenuItem.Image = global::Office_File_Explorer.Properties.Resources.BatchFile_16x;
            this.batchFileProcessingToolStripMenuItem.Name = "batchFileProcessingToolStripMenuItem";
            this.batchFileProcessingToolStripMenuItem.Size = new System.Drawing.Size(185, 22);
            this.batchFileProcessingToolStripMenuItem.Text = "Batch File Processing";
            this.batchFileProcessingToolStripMenuItem.Click += new System.EventHandler(this.BatchFileProcessingToolStripMenuItem_Click);
            // 
            // clipboardViewerToolStripMenuItem
            // 
            this.clipboardViewerToolStripMenuItem.Image = global::Office_File_Explorer.Properties.Resources.ASX_Copy_blue_16x;
            this.clipboardViewerToolStripMenuItem.Name = "clipboardViewerToolStripMenuItem";
            this.clipboardViewerToolStripMenuItem.Size = new System.Drawing.Size(185, 22);
            this.clipboardViewerToolStripMenuItem.Text = "Clipboard Viewer";
            this.clipboardViewerToolStripMenuItem.Click += new System.EventHandler(this.ClipboardViewerToolStripMenuItem_Click);
            // 
            // base64DecoderToolStripMenuItem
            // 
            this.base64DecoderToolStripMenuItem.Image = global::Office_File_Explorer.Properties.Resources.Strings_7959_0;
            this.base64DecoderToolStripMenuItem.Name = "base64DecoderToolStripMenuItem";
            this.base64DecoderToolStripMenuItem.Size = new System.Drawing.Size(185, 22);
            this.base64DecoderToolStripMenuItem.Text = "Base64 Decoder";
            this.base64DecoderToolStripMenuItem.Click += new System.EventHandler(this.Base64DecoderToolStripMenuItem_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aboutToolStripMenuItem,
            this.feedbackToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.helpToolStripMenuItem.Text = "&Help";
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Image = global::Office_File_Explorer.Properties.Resources.Dialog_16x;
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.aboutToolStripMenuItem.Text = "About";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.AboutToolStripMenuItem_Click);
            // 
            // feedbackToolStripMenuItem
            // 
            this.feedbackToolStripMenuItem.Image = global::Office_File_Explorer.Properties.Resources.FeedbackBubble_16x;
            this.feedbackToolStripMenuItem.Name = "feedbackToolStripMenuItem";
            this.feedbackToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.feedbackToolStripMenuItem.Text = "Feedback";
            this.feedbackToolStripMenuItem.Click += new System.EventHandler(this.FeedbackToolStripMenuItem_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.BtnRemoveCustomXmlParts);
            this.groupBox1.Controls.Add(this.BtnRemoveCustomFileProps);
            this.groupBox1.Controls.Add(this.BtnExcelSheetViewer);
            this.groupBox1.Controls.Add(this.BtnValidateDoc);
            this.groupBox1.Controls.Add(this.BtnViewCustomUI);
            this.groupBox1.Controls.Add(this.BtnFixCorruptDoc);
            this.groupBox1.Controls.Add(this.BtnCustomXml);
            this.groupBox1.Controls.Add(this.BtnDocProps);
            this.groupBox1.Controls.Add(this.BtnViewImages);
            this.groupBox1.Controls.Add(this.lblFileType);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.BtnSearchAndReplace);
            this.groupBox1.Controls.Add(this.BtnFixDocument);
            this.groupBox1.Controls.Add(this.lblFilePath);
            this.groupBox1.Controls.Add(this.BtnModifyContent);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.BtnViewContents);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox1.Location = new System.Drawing.Point(0, 24);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(924, 136);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Document Information";
            // 
            // BtnRemoveCustomXmlParts
            // 
            this.BtnRemoveCustomXmlParts.Location = new System.Drawing.Point(525, 107);
            this.BtnRemoveCustomXmlParts.Name = "BtnRemoveCustomXmlParts";
            this.BtnRemoveCustomXmlParts.Size = new System.Drawing.Size(143, 23);
            this.BtnRemoveCustomXmlParts.TabIndex = 13;
            this.BtnRemoveCustomXmlParts.Text = "Remove Custom Xml";
            this.BtnRemoveCustomXmlParts.UseVisualStyleBackColor = true;
            this.BtnRemoveCustomXmlParts.Click += new System.EventHandler(this.BtnRemoveCustomXmlParts_Click);
            // 
            // BtnRemoveCustomFileProps
            // 
            this.BtnRemoveCustomFileProps.Location = new System.Drawing.Point(674, 107);
            this.BtnRemoveCustomFileProps.Name = "BtnRemoveCustomFileProps";
            this.BtnRemoveCustomFileProps.Size = new System.Drawing.Size(175, 23);
            this.BtnRemoveCustomFileProps.TabIndex = 12;
            this.BtnRemoveCustomFileProps.Text = "Remove Custom File Props";
            this.BtnRemoveCustomFileProps.UseVisualStyleBackColor = true;
            this.BtnRemoveCustomFileProps.Click += new System.EventHandler(this.BtnRemoveCustomFileProps_Click);
            // 
            // BtnExcelSheetViewer
            // 
            this.BtnExcelSheetViewer.Location = new System.Drawing.Point(653, 78);
            this.BtnExcelSheetViewer.Name = "BtnExcelSheetViewer";
            this.BtnExcelSheetViewer.Size = new System.Drawing.Size(126, 23);
            this.BtnExcelSheetViewer.TabIndex = 11;
            this.BtnExcelSheetViewer.Text = "Excel Sheet Viewer";
            this.BtnExcelSheetViewer.UseVisualStyleBackColor = true;
            this.BtnExcelSheetViewer.Click += new System.EventHandler(this.BtnExcelSheetViewer_Click);
            // 
            // BtnValidateDoc
            // 
            this.BtnValidateDoc.Location = new System.Drawing.Point(785, 78);
            this.BtnValidateDoc.Name = "BtnValidateDoc";
            this.BtnValidateDoc.Size = new System.Drawing.Size(135, 23);
            this.BtnValidateDoc.TabIndex = 10;
            this.BtnValidateDoc.Text = "Validate Document";
            this.BtnValidateDoc.UseVisualStyleBackColor = true;
            this.BtnValidateDoc.Click += new System.EventHandler(this.BtnValidateDoc_Click);
            // 
            // BtnViewCustomUI
            // 
            this.BtnViewCustomUI.Location = new System.Drawing.Point(538, 78);
            this.BtnViewCustomUI.Name = "BtnViewCustomUI";
            this.BtnViewCustomUI.Size = new System.Drawing.Size(109, 23);
            this.BtnViewCustomUI.TabIndex = 4;
            this.BtnViewCustomUI.Text = "View Custom UI";
            this.BtnViewCustomUI.UseVisualStyleBackColor = true;
            this.BtnViewCustomUI.Click += new System.EventHandler(this.BtnViewCustomUI_Click);
            // 
            // BtnFixCorruptDoc
            // 
            this.BtnFixCorruptDoc.Location = new System.Drawing.Point(233, 107);
            this.BtnFixCorruptDoc.Name = "BtnFixCorruptDoc";
            this.BtnFixCorruptDoc.Size = new System.Drawing.Size(147, 23);
            this.BtnFixCorruptDoc.TabIndex = 9;
            this.BtnFixCorruptDoc.Text = "Fix Corrupt Document";
            this.BtnFixCorruptDoc.UseVisualStyleBackColor = true;
            this.BtnFixCorruptDoc.Click += new System.EventHandler(this.BtnFixCorruptDoc_Click);
            // 
            // BtnCustomXml
            // 
            this.BtnCustomXml.Location = new System.Drawing.Point(406, 78);
            this.BtnCustomXml.Name = "BtnCustomXml";
            this.BtnCustomXml.Size = new System.Drawing.Size(126, 23);
            this.BtnCustomXml.TabIndex = 8;
            this.BtnCustomXml.Text = "View Custom Xml";
            this.BtnCustomXml.UseVisualStyleBackColor = true;
            this.BtnCustomXml.Click += new System.EventHandler(this.BtnCustomXml_Click);
            // 
            // BtnDocProps
            // 
            this.BtnDocProps.Location = new System.Drawing.Point(232, 78);
            this.BtnDocProps.Name = "BtnDocProps";
            this.BtnDocProps.Size = new System.Drawing.Size(168, 23);
            this.BtnDocProps.TabIndex = 7;
            this.BtnDocProps.Text = "View Document Properties";
            this.BtnDocProps.UseVisualStyleBackColor = true;
            this.BtnDocProps.Click += new System.EventHandler(this.BtnDocProps_Click);
            // 
            // BtnViewImages
            // 
            this.BtnViewImages.Location = new System.Drawing.Point(119, 78);
            this.BtnViewImages.Name = "BtnViewImages";
            this.BtnViewImages.Size = new System.Drawing.Size(107, 23);
            this.BtnViewImages.TabIndex = 6;
            this.BtnViewImages.Text = "View Images";
            this.BtnViewImages.UseVisualStyleBackColor = true;
            this.BtnViewImages.Click += new System.EventHandler(this.BtnViewImages_Click);
            // 
            // lblFileType
            // 
            this.lblFileType.AutoSize = true;
            this.lblFileType.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblFileType.Location = new System.Drawing.Point(46, 46);
            this.lblFileType.Name = "lblFileType";
            this.lblFileType.Size = new System.Drawing.Size(2, 17);
            this.lblFileType.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 46);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(37, 15);
            this.label2.TabIndex = 2;
            this.label2.Text = "Type: ";
            // 
            // BtnSearchAndReplace
            // 
            this.BtnSearchAndReplace.Location = new System.Drawing.Point(386, 107);
            this.BtnSearchAndReplace.Name = "BtnSearchAndReplace";
            this.BtnSearchAndReplace.Size = new System.Drawing.Size(133, 23);
            this.BtnSearchAndReplace.TabIndex = 1;
            this.BtnSearchAndReplace.Text = "Search and Replace";
            this.BtnSearchAndReplace.UseVisualStyleBackColor = true;
            this.BtnSearchAndReplace.Click += new System.EventHandler(this.BtnSearchAndReplace_Click);
            // 
            // BtnFixDocument
            // 
            this.BtnFixDocument.Location = new System.Drawing.Point(119, 107);
            this.BtnFixDocument.Name = "BtnFixDocument";
            this.BtnFixDocument.Size = new System.Drawing.Size(108, 23);
            this.BtnFixDocument.TabIndex = 3;
            this.BtnFixDocument.Text = "Fix Document";
            this.BtnFixDocument.UseVisualStyleBackColor = true;
            this.BtnFixDocument.Click += new System.EventHandler(this.BtnFixDocument_Click);
            // 
            // lblFilePath
            // 
            this.lblFilePath.AutoSize = true;
            this.lblFilePath.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblFilePath.Location = new System.Drawing.Point(46, 19);
            this.lblFilePath.Name = "lblFilePath";
            this.lblFilePath.Size = new System.Drawing.Size(2, 17);
            this.lblFilePath.TabIndex = 1;
            // 
            // BtnModifyContent
            // 
            this.BtnModifyContent.Location = new System.Drawing.Point(6, 107);
            this.BtnModifyContent.Name = "BtnModifyContent";
            this.BtnModifyContent.Size = new System.Drawing.Size(107, 23);
            this.BtnModifyContent.TabIndex = 2;
            this.BtnModifyContent.Text = "Modify Contents";
            this.BtnModifyContent.UseVisualStyleBackColor = true;
            this.BtnModifyContent.Click += new System.EventHandler(this.BtnModifyContent_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(28, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "File:";
            // 
            // BtnViewContents
            // 
            this.BtnViewContents.Location = new System.Drawing.Point(6, 78);
            this.BtnViewContents.Name = "BtnViewContents";
            this.BtnViewContents.Size = new System.Drawing.Size(107, 23);
            this.BtnViewContents.TabIndex = 0;
            this.BtnViewContents.Text = "View Contents";
            this.BtnViewContents.UseVisualStyleBackColor = true;
            this.BtnViewContents.Click += new System.EventHandler(this.BtnViewContents_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.LstDisplay);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Location = new System.Drawing.Point(0, 160);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(924, 389);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Output";
            // 
            // LstDisplay
            // 
            this.LstDisplay.ContextMenuStrip = this.contextMenuStrip1;
            this.LstDisplay.Dock = System.Windows.Forms.DockStyle.Fill;
            this.LstDisplay.FormattingEnabled = true;
            this.LstDisplay.HorizontalScrollbar = true;
            this.LstDisplay.ItemHeight = 15;
            this.LstDisplay.Location = new System.Drawing.Point(3, 19);
            this.LstDisplay.Name = "LstDisplay";
            this.LstDisplay.Size = new System.Drawing.Size(918, 367);
            this.LstDisplay.TabIndex = 0;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.copySelectedLineToolStripMenuItem,
            this.copyAllLinesToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(175, 48);
            // 
            // copySelectedLineToolStripMenuItem
            // 
            this.copySelectedLineToolStripMenuItem.Name = "copySelectedLineToolStripMenuItem";
            this.copySelectedLineToolStripMenuItem.Size = new System.Drawing.Size(174, 22);
            this.copySelectedLineToolStripMenuItem.Text = "Copy Selected Line";
            this.copySelectedLineToolStripMenuItem.Click += new System.EventHandler(this.CopySelectedLineToolStripMenuItem_Click);
            // 
            // copyAllLinesToolStripMenuItem
            // 
            this.copyAllLinesToolStripMenuItem.Name = "copyAllLinesToolStripMenuItem";
            this.copyAllLinesToolStripMenuItem.Size = new System.Drawing.Size(174, 22);
            this.copyAllLinesToolStripMenuItem.Text = "Copy All Lines";
            this.copyAllLinesToolStripMenuItem.Click += new System.EventHandler(this.CopyAllLinesToolStripMenuItem_Click);
            // 
            // imageList1
            // 
            this.imageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "excel.png");
            this.imageList1.Images.SetKeyName(1, "powerpoint.png");
            this.imageList1.Images.SetKeyName(2, "word.png");
            this.imageList1.Images.SetKeyName(3, "XMLFile_789_32.ico");
            // 
            // FrmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(924, 549);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.mnuMainMenu);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.mnuMainMenu;
            this.MinimumSize = new System.Drawing.Size(920, 588);
            this.Name = "FrmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Office File Explorer";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmMain_FormClosing);
            this.mnuMainMenu.ResumeLayout(false);
            this.mnuMainMenu.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

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
        private System.Windows.Forms.Button BtnCustomXml;
        private System.Windows.Forms.Button BtnDocProps;
        private System.Windows.Forms.Button BtnViewImages;
        private System.Windows.Forms.Button BtnFixCorruptDoc;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem copySelectedLineToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem copyAllLinesToolStripMenuItem;
        private System.Windows.Forms.Button BtnViewCustomUI;
        private System.Windows.Forms.ToolStripMenuItem base64DecoderToolStripMenuItem;
        public System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Button BtnValidateDoc;
        private System.Windows.Forms.ToolStripMenuItem openFileBackupFolderToolStripMenuItem;
        private System.Windows.Forms.Button BtnExcelSheetViewer;
        private System.Windows.Forms.Button BtnRemoveCustomFileProps;
        private System.Windows.Forms.Button BtnRemoveCustomXmlParts;
    }
}

