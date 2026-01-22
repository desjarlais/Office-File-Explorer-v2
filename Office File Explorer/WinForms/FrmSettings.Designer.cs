
namespace Office_File_Explorer.WinForms
{
    partial class FrmSettings
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmSettings));
            groupBox1 = new System.Windows.Forms.GroupBox();
            ckbDeleteOnlyCommentBookmarks = new System.Windows.Forms.CheckBox();
            ckbFixGroupedShapes = new System.Windows.Forms.CheckBox();
            ckbListRsids = new System.Windows.Forms.CheckBox();
            ckbRemoveFallbackTags = new System.Windows.Forms.CheckBox();
            groupBox2 = new System.Windows.Forms.GroupBox();
            ckbResetIndentLevels = new System.Windows.Forms.CheckBox();
            ckbRemoveCustDataTags = new System.Windows.Forms.CheckBox();
            ckbResetNotes = new System.Windows.Forms.CheckBox();
            groupBox3 = new System.Windows.Forms.GroupBox();
            ckbCleanInvalidXml = new System.Windows.Forms.CheckBox();
            ckbOutlookMsgAsRtf = new System.Windows.Forms.CheckBox();
            ckbZipItemCorrupt = new System.Windows.Forms.CheckBox();
            ckbDeleteOnExit = new System.Windows.Forms.CheckBox();
            groupBox4 = new System.Windows.Forms.GroupBox();
            rdoSAX = new System.Windows.Forms.RadioButton();
            rdoDOM = new System.Windows.Forms.RadioButton();
            BtnOk = new System.Windows.Forms.Button();
            BtnCancel = new System.Windows.Forms.Button();
            rdoUseCCGuid = new System.Windows.Forms.RadioButton();
            rdoUseSPGuid = new System.Windows.Forms.RadioButton();
            rdoUserSelectedCC = new System.Windows.Forms.RadioButton();
            groupBox6 = new System.Windows.Forms.GroupBox();
            ckbViewInUseStylesOnly = new System.Windows.Forms.CheckBox();
            groupBox1.SuspendLayout();
            groupBox2.SuspendLayout();
            groupBox3.SuspendLayout();
            groupBox4.SuspendLayout();
            groupBox6.SuspendLayout();
            SuspendLayout();
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(ckbDeleteOnlyCommentBookmarks);
            groupBox1.Controls.Add(ckbFixGroupedShapes);
            groupBox1.Controls.Add(ckbListRsids);
            groupBox1.Controls.Add(ckbRemoveFallbackTags);
            groupBox1.Location = new System.Drawing.Point(17, 20);
            groupBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            groupBox1.Name = "groupBox1";
            groupBox1.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            groupBox1.Size = new System.Drawing.Size(336, 202);
            groupBox1.TabIndex = 0;
            groupBox1.TabStop = false;
            groupBox1.Text = "Word Options";
            // 
            // ckbDeleteOnlyCommentBookmarks
            // 
            ckbDeleteOnlyCommentBookmarks.AutoSize = true;
            ckbDeleteOnlyCommentBookmarks.Location = new System.Drawing.Point(9, 152);
            ckbDeleteOnlyCommentBookmarks.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            ckbDeleteOnlyCommentBookmarks.Name = "ckbDeleteOnlyCommentBookmarks";
            ckbDeleteOnlyCommentBookmarks.Size = new System.Drawing.Size(336, 29);
            ckbDeleteOnlyCommentBookmarks.TabIndex = 4;
            ckbDeleteOnlyCommentBookmarks.Text = "Delete Only Bookmarks In Comments";
            ckbDeleteOnlyCommentBookmarks.UseVisualStyleBackColor = true;
            // 
            // ckbFixGroupedShapes
            // 
            ckbFixGroupedShapes.AutoSize = true;
            ckbFixGroupedShapes.Location = new System.Drawing.Point(9, 110);
            ckbFixGroupedShapes.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            ckbFixGroupedShapes.Name = "ckbFixGroupedShapes";
            ckbFixGroupedShapes.Size = new System.Drawing.Size(196, 29);
            ckbFixGroupedShapes.TabIndex = 2;
            ckbFixGroupedShapes.Text = "Fix Grouped Shapes";
            ckbFixGroupedShapes.UseVisualStyleBackColor = true;
            // 
            // ckbListRsids
            // 
            ckbListRsids.AutoSize = true;
            ckbListRsids.Location = new System.Drawing.Point(9, 68);
            ckbListRsids.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            ckbListRsids.Name = "ckbListRsids";
            ckbListRsids.Size = new System.Drawing.Size(241, 29);
            ckbListRsids.TabIndex = 1;
            ckbListRsids.Text = "List Rsids With Doc Props";
            ckbListRsids.UseVisualStyleBackColor = true;
            // 
            // ckbRemoveFallbackTags
            // 
            ckbRemoveFallbackTags.AutoSize = true;
            ckbRemoveFallbackTags.Location = new System.Drawing.Point(9, 32);
            ckbRemoveFallbackTags.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            ckbRemoveFallbackTags.Name = "ckbRemoveFallbackTags";
            ckbRemoveFallbackTags.Size = new System.Drawing.Size(234, 29);
            ckbRemoveFallbackTags.TabIndex = 0;
            ckbRemoveFallbackTags.Text = "Remove All Fallback Tags";
            ckbRemoveFallbackTags.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(ckbResetIndentLevels);
            groupBox2.Controls.Add(ckbRemoveCustDataTags);
            groupBox2.Controls.Add(ckbResetNotes);
            groupBox2.Location = new System.Drawing.Point(361, 20);
            groupBox2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            groupBox2.Name = "groupBox2";
            groupBox2.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            groupBox2.Size = new System.Drawing.Size(340, 153);
            groupBox2.TabIndex = 0;
            groupBox2.TabStop = false;
            groupBox2.Text = "PowerPoint Options";
            // 
            // ckbResetIndentLevels
            // 
            ckbResetIndentLevels.AutoSize = true;
            ckbResetIndentLevels.Location = new System.Drawing.Point(11, 112);
            ckbResetIndentLevels.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            ckbResetIndentLevels.Name = "ckbResetIndentLevels";
            ckbResetIndentLevels.Size = new System.Drawing.Size(188, 29);
            ckbResetIndentLevels.TabIndex = 7;
            ckbResetIndentLevels.Text = "Reset Indent Levels";
            ckbResetIndentLevels.UseVisualStyleBackColor = true;
            // 
            // ckbRemoveCustDataTags
            // 
            ckbRemoveCustDataTags.AutoSize = true;
            ckbRemoveCustDataTags.Location = new System.Drawing.Point(11, 73);
            ckbRemoveCustDataTags.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            ckbRemoveCustDataTags.Name = "ckbRemoveCustDataTags";
            ckbRemoveCustDataTags.Size = new System.Drawing.Size(251, 29);
            ckbRemoveCustDataTags.TabIndex = 6;
            ckbRemoveCustDataTags.Text = "Remove Custom Data Tags";
            ckbRemoveCustDataTags.UseVisualStyleBackColor = true;
            // 
            // ckbResetNotes
            // 
            ckbResetNotes.AutoSize = true;
            ckbResetNotes.Location = new System.Drawing.Point(11, 32);
            ckbResetNotes.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            ckbResetNotes.Name = "ckbResetNotes";
            ckbResetNotes.Size = new System.Drawing.Size(329, 29);
            ckbResetNotes.TabIndex = 0;
            ckbResetNotes.Text = "Reset Notes Slides and Notes Master";
            ckbResetNotes.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            groupBox3.Controls.Add(ckbViewInUseStylesOnly);
            groupBox3.Controls.Add(ckbCleanInvalidXml);
            groupBox3.Controls.Add(ckbOutlookMsgAsRtf);
            groupBox3.Controls.Add(ckbZipItemCorrupt);
            groupBox3.Controls.Add(ckbDeleteOnExit);
            groupBox3.Location = new System.Drawing.Point(361, 183);
            groupBox3.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            groupBox3.Name = "groupBox3";
            groupBox3.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            groupBox3.Size = new System.Drawing.Size(340, 298);
            groupBox3.TabIndex = 0;
            groupBox3.TabStop = false;
            groupBox3.Text = "App Settings";
            // 
            // ckbCleanInvalidXml
            // 
            ckbCleanInvalidXml.AutoSize = true;
            ckbCleanInvalidXml.Location = new System.Drawing.Point(9, 170);
            ckbCleanInvalidXml.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            ckbCleanInvalidXml.Name = "ckbCleanInvalidXml";
            ckbCleanInvalidXml.Size = new System.Drawing.Size(282, 29);
            ckbCleanInvalidXml.TabIndex = 6;
            ckbCleanInvalidXml.Text = "Remove Invalid Xml Characters";
            ckbCleanInvalidXml.UseVisualStyleBackColor = true;
            // 
            // ckbOutlookMsgAsRtf
            // 
            ckbOutlookMsgAsRtf.AutoSize = true;
            ckbOutlookMsgAsRtf.Location = new System.Drawing.Point(9, 128);
            ckbOutlookMsgAsRtf.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            ckbOutlookMsgAsRtf.Name = "ckbOutlookMsgAsRtf";
            ckbOutlookMsgAsRtf.Size = new System.Drawing.Size(297, 29);
            ckbOutlookMsgAsRtf.TabIndex = 5;
            ckbOutlookMsgAsRtf.Text = "Display Outlook Msg Files in RTF";
            ckbOutlookMsgAsRtf.UseVisualStyleBackColor = true;
            // 
            // ckbZipItemCorrupt
            // 
            ckbZipItemCorrupt.AutoSize = true;
            ckbZipItemCorrupt.Location = new System.Drawing.Point(9, 83);
            ckbZipItemCorrupt.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            ckbZipItemCorrupt.Name = "ckbZipItemCorrupt";
            ckbZipItemCorrupt.Size = new System.Drawing.Size(284, 29);
            ckbZipItemCorrupt.TabIndex = 3;
            ckbZipItemCorrupt.Text = "Check Zip Corruption On Open";
            ckbZipItemCorrupt.UseVisualStyleBackColor = true;
            // 
            // ckbDeleteOnExit
            // 
            ckbDeleteOnExit.AutoSize = true;
            ckbDeleteOnExit.Location = new System.Drawing.Point(9, 37);
            ckbDeleteOnExit.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            ckbDeleteOnExit.Name = "ckbDeleteOnExit";
            ckbDeleteOnExit.Size = new System.Drawing.Size(250, 29);
            ckbDeleteOnExit.TabIndex = 0;
            ckbDeleteOnExit.Text = "Delete Copied Files On Exit";
            ckbDeleteOnExit.UseVisualStyleBackColor = true;
            // 
            // groupBox4
            // 
            groupBox4.Controls.Add(rdoSAX);
            groupBox4.Controls.Add(rdoDOM);
            groupBox4.Location = new System.Drawing.Point(17, 413);
            groupBox4.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            groupBox4.Name = "groupBox4";
            groupBox4.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            groupBox4.Size = new System.Drawing.Size(336, 125);
            groupBox4.TabIndex = 0;
            groupBox4.TabStop = false;
            groupBox4.Text = "Excel Parse Options";
            // 
            // rdoSAX
            // 
            rdoSAX.AutoSize = true;
            rdoSAX.Location = new System.Drawing.Point(9, 37);
            rdoSAX.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            rdoSAX.Name = "rdoSAX";
            rdoSAX.Size = new System.Drawing.Size(120, 29);
            rdoSAX.TabIndex = 3;
            rdoSAX.TabStop = true;
            rdoSAX.Text = "SAX Styles";
            rdoSAX.UseVisualStyleBackColor = true;
            // 
            // rdoDOM
            // 
            rdoDOM.AutoSize = true;
            rdoDOM.Location = new System.Drawing.Point(9, 78);
            rdoDOM.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            rdoDOM.Name = "rdoDOM";
            rdoDOM.Size = new System.Drawing.Size(130, 29);
            rdoDOM.TabIndex = 4;
            rdoDOM.TabStop = true;
            rdoDOM.Text = "DOM Styles";
            rdoDOM.UseVisualStyleBackColor = true;
            // 
            // BtnOk
            // 
            BtnOk.Location = new System.Drawing.Point(481, 508);
            BtnOk.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            BtnOk.Name = "BtnOk";
            BtnOk.Size = new System.Drawing.Size(107, 38);
            BtnOk.TabIndex = 1;
            BtnOk.Text = "OK";
            BtnOk.UseVisualStyleBackColor = true;
            BtnOk.Click += BtnOk_Click;
            // 
            // BtnCancel
            // 
            BtnCancel.Location = new System.Drawing.Point(597, 508);
            BtnCancel.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            BtnCancel.Name = "BtnCancel";
            BtnCancel.Size = new System.Drawing.Size(107, 38);
            BtnCancel.TabIndex = 2;
            BtnCancel.Text = "Cancel";
            BtnCancel.UseVisualStyleBackColor = true;
            BtnCancel.Click += BtnCancel_Click;
            // 
            // rdoUseCCGuid
            // 
            rdoUseCCGuid.AutoSize = true;
            rdoUseCCGuid.Location = new System.Drawing.Point(9, 78);
            rdoUseCCGuid.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            rdoUseCCGuid.Name = "rdoUseCCGuid";
            rdoUseCCGuid.Size = new System.Drawing.Size(240, 29);
            rdoUseCCGuid.TabIndex = 4;
            rdoUseCCGuid.TabStop = true;
            rdoUseCCGuid.Text = "Use Content Control Guid";
            rdoUseCCGuid.UseVisualStyleBackColor = true;
            // 
            // rdoUseSPGuid
            // 
            rdoUseSPGuid.AutoSize = true;
            rdoUseSPGuid.Location = new System.Drawing.Point(9, 37);
            rdoUseSPGuid.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            rdoUseSPGuid.Name = "rdoUseSPGuid";
            rdoUseSPGuid.Size = new System.Drawing.Size(300, 29);
            rdoUseSPGuid.TabIndex = 3;
            rdoUseSPGuid.TabStop = true;
            rdoUseSPGuid.Text = "Use SharePoint Custom Xml Guid";
            rdoUseSPGuid.UseVisualStyleBackColor = true;
            // 
            // rdoUserSelectedCC
            // 
            rdoUserSelectedCC.AutoSize = true;
            rdoUserSelectedCC.Location = new System.Drawing.Point(9, 120);
            rdoUserSelectedCC.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            rdoUserSelectedCC.Name = "rdoUserSelectedCC";
            rdoUserSelectedCC.Size = new System.Drawing.Size(322, 29);
            rdoUserSelectedCC.TabIndex = 5;
            rdoUserSelectedCC.TabStop = true;
            rdoUserSelectedCC.Text = "Use User Selected Custom Xml Guid";
            rdoUserSelectedCC.UseVisualStyleBackColor = true;
            // 
            // groupBox6
            // 
            groupBox6.Controls.Add(rdoUserSelectedCC);
            groupBox6.Controls.Add(rdoUseSPGuid);
            groupBox6.Controls.Add(rdoUseCCGuid);
            groupBox6.Location = new System.Drawing.Point(17, 232);
            groupBox6.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            groupBox6.Name = "groupBox6";
            groupBox6.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            groupBox6.Size = new System.Drawing.Size(336, 172);
            groupBox6.TabIndex = 5;
            groupBox6.TabStop = false;
            groupBox6.Text = "Fix Content Control Prefix Mappings";
            // 
            // ckbViewInUseStylesOnly
            // 
            ckbViewInUseStylesOnly.AutoSize = true;
            ckbViewInUseStylesOnly.Location = new System.Drawing.Point(10, 216);
            ckbViewInUseStylesOnly.Name = "ckbViewInUseStylesOnly";
            ckbViewInUseStylesOnly.Size = new System.Drawing.Size(221, 29);
            ckbViewInUseStylesOnly.TabIndex = 7;
            ckbViewInUseStylesOnly.Text = "View In Use Styles Only";
            ckbViewInUseStylesOnly.UseVisualStyleBackColor = true;
            // 
            // FrmSettings
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(10F, 25F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(719, 567);
            Controls.Add(groupBox6);
            Controls.Add(BtnCancel);
            Controls.Add(BtnOk);
            Controls.Add(groupBox2);
            Controls.Add(groupBox3);
            Controls.Add(groupBox4);
            Controls.Add(groupBox1);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            Icon = (System.Drawing.Icon)resources.GetObject("$this.Icon");
            KeyPreview = true;
            Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "FrmSettings";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            Text = "Settings";
            KeyDown += FrmSettings_KeyDown;
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            groupBox2.ResumeLayout(false);
            groupBox2.PerformLayout();
            groupBox3.ResumeLayout(false);
            groupBox3.PerformLayout();
            groupBox4.ResumeLayout(false);
            groupBox4.PerformLayout();
            groupBox6.ResumeLayout(false);
            groupBox6.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox ckbFixGroupedShapes;
        private System.Windows.Forms.CheckBox ckbListRsids;
        private System.Windows.Forms.CheckBox ckbRemoveFallbackTags;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox ckbResetNotes;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.CheckBox ckbDeleteOnExit;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Button BtnOk;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.RadioButton rdoSAX;
        private System.Windows.Forms.RadioButton rdoDOM;
        private System.Windows.Forms.CheckBox ckbZipItemCorrupt;
        private System.Windows.Forms.CheckBox ckbDeleteOnlyCommentBookmarks;
        private System.Windows.Forms.CheckBox ckbRemoveCustDataTags;
        private System.Windows.Forms.RadioButton rdoUseCCGuid;
        private System.Windows.Forms.RadioButton rdoUseSPGuid;
        private System.Windows.Forms.RadioButton rdoUserSelectedCC;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.CheckBox ckbResetIndentLevels;
        private System.Windows.Forms.CheckBox ckbOutlookMsgAsRtf;
        private System.Windows.Forms.CheckBox ckbCleanInvalidXml;
        private System.Windows.Forms.CheckBox ckbViewInUseStylesOnly;
    }
}