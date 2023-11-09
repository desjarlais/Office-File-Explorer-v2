
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
            ckbDisableAutoXmlColorFormatting = new System.Windows.Forms.CheckBox();
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
            groupBox1.Location = new System.Drawing.Point(12, 12);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new System.Drawing.Size(235, 173);
            groupBox1.TabIndex = 0;
            groupBox1.TabStop = false;
            groupBox1.Text = "Word Corrupt Document";
            // 
            // ckbDeleteOnlyCommentBookmarks
            // 
            ckbDeleteOnlyCommentBookmarks.AutoSize = true;
            ckbDeleteOnlyCommentBookmarks.Location = new System.Drawing.Point(6, 91);
            ckbDeleteOnlyCommentBookmarks.Name = "ckbDeleteOnlyCommentBookmarks";
            ckbDeleteOnlyCommentBookmarks.Size = new System.Drawing.Size(224, 19);
            ckbDeleteOnlyCommentBookmarks.TabIndex = 4;
            ckbDeleteOnlyCommentBookmarks.Text = "Delete Only Bookmarks In Comments";
            ckbDeleteOnlyCommentBookmarks.UseVisualStyleBackColor = true;
            // 
            // ckbFixGroupedShapes
            // 
            ckbFixGroupedShapes.AutoSize = true;
            ckbFixGroupedShapes.Location = new System.Drawing.Point(6, 66);
            ckbFixGroupedShapes.Name = "ckbFixGroupedShapes";
            ckbFixGroupedShapes.Size = new System.Drawing.Size(130, 19);
            ckbFixGroupedShapes.TabIndex = 2;
            ckbFixGroupedShapes.Text = "Fix Grouped Shapes";
            ckbFixGroupedShapes.UseVisualStyleBackColor = true;
            // 
            // ckbListRsids
            // 
            ckbListRsids.AutoSize = true;
            ckbListRsids.Location = new System.Drawing.Point(6, 41);
            ckbListRsids.Name = "ckbListRsids";
            ckbListRsids.Size = new System.Drawing.Size(159, 19);
            ckbListRsids.TabIndex = 1;
            ckbListRsids.Text = "List Rsids With Doc Props";
            ckbListRsids.UseVisualStyleBackColor = true;
            // 
            // ckbRemoveFallbackTags
            // 
            ckbRemoveFallbackTags.AutoSize = true;
            ckbRemoveFallbackTags.Location = new System.Drawing.Point(6, 19);
            ckbRemoveFallbackTags.Name = "ckbRemoveFallbackTags";
            ckbRemoveFallbackTags.Size = new System.Drawing.Size(158, 19);
            ckbRemoveFallbackTags.TabIndex = 0;
            ckbRemoveFallbackTags.Text = "Remove All Fallback Tags";
            ckbRemoveFallbackTags.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(ckbResetIndentLevels);
            groupBox2.Controls.Add(ckbRemoveCustDataTags);
            groupBox2.Controls.Add(ckbResetNotes);
            groupBox2.Location = new System.Drawing.Point(253, 12);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new System.Drawing.Size(238, 92);
            groupBox2.TabIndex = 0;
            groupBox2.TabStop = false;
            groupBox2.Text = "PowerPoint Options";
            // 
            // ckbResetIndentLevels
            // 
            ckbResetIndentLevels.AutoSize = true;
            ckbResetIndentLevels.Location = new System.Drawing.Point(8, 67);
            ckbResetIndentLevels.Name = "ckbResetIndentLevels";
            ckbResetIndentLevels.Size = new System.Drawing.Size(126, 19);
            ckbResetIndentLevels.TabIndex = 7;
            ckbResetIndentLevels.Text = "Reset Indent Levels";
            ckbResetIndentLevels.UseVisualStyleBackColor = true;
            // 
            // ckbRemoveCustDataTags
            // 
            ckbRemoveCustDataTags.AutoSize = true;
            ckbRemoveCustDataTags.Location = new System.Drawing.Point(8, 44);
            ckbRemoveCustDataTags.Name = "ckbRemoveCustDataTags";
            ckbRemoveCustDataTags.Size = new System.Drawing.Size(167, 19);
            ckbRemoveCustDataTags.TabIndex = 6;
            ckbRemoveCustDataTags.Text = "Remove Custom Data Tags";
            ckbRemoveCustDataTags.UseVisualStyleBackColor = true;
            // 
            // ckbResetNotes
            // 
            ckbResetNotes.AutoSize = true;
            ckbResetNotes.Location = new System.Drawing.Point(8, 19);
            ckbResetNotes.Name = "ckbResetNotes";
            ckbResetNotes.Size = new System.Drawing.Size(217, 19);
            ckbResetNotes.TabIndex = 0;
            ckbResetNotes.Text = "Reset Notes Slides and Notes Master";
            ckbResetNotes.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            groupBox3.Controls.Add(ckbDisableAutoXmlColorFormatting);
            groupBox3.Controls.Add(ckbZipItemCorrupt);
            groupBox3.Controls.Add(ckbDeleteOnExit);
            groupBox3.Location = new System.Drawing.Point(253, 191);
            groupBox3.Name = "groupBox3";
            groupBox3.Size = new System.Drawing.Size(238, 103);
            groupBox3.TabIndex = 0;
            groupBox3.TabStop = false;
            groupBox3.Text = "App Settings";
            // 
            // ckbDisableAutoXmlColorFormatting
            // 
            ckbDisableAutoXmlColorFormatting.AutoSize = true;
            ckbDisableAutoXmlColorFormatting.Location = new System.Drawing.Point(6, 78);
            ckbDisableAutoXmlColorFormatting.Name = "ckbDisableAutoXmlColorFormatting";
            ckbDisableAutoXmlColorFormatting.Size = new System.Drawing.Size(182, 19);
            ckbDisableAutoXmlColorFormatting.TabIndex = 4;
            ckbDisableAutoXmlColorFormatting.Text = "Disable Xml Color Formatting";
            ckbDisableAutoXmlColorFormatting.UseVisualStyleBackColor = true;
            // 
            // ckbZipItemCorrupt
            // 
            ckbZipItemCorrupt.AutoSize = true;
            ckbZipItemCorrupt.Location = new System.Drawing.Point(6, 50);
            ckbZipItemCorrupt.Name = "ckbZipItemCorrupt";
            ckbZipItemCorrupt.Size = new System.Drawing.Size(191, 19);
            ckbZipItemCorrupt.TabIndex = 3;
            ckbZipItemCorrupt.Text = "Check Zip Corruption On Open";
            ckbZipItemCorrupt.UseVisualStyleBackColor = true;
            // 
            // ckbDeleteOnExit
            // 
            ckbDeleteOnExit.AutoSize = true;
            ckbDeleteOnExit.Location = new System.Drawing.Point(6, 22);
            ckbDeleteOnExit.Name = "ckbDeleteOnExit";
            ckbDeleteOnExit.Size = new System.Drawing.Size(167, 19);
            ckbDeleteOnExit.TabIndex = 0;
            ckbDeleteOnExit.Text = "Delete Copied Files On Exit";
            ckbDeleteOnExit.UseVisualStyleBackColor = true;
            // 
            // groupBox4
            // 
            groupBox4.Controls.Add(rdoSAX);
            groupBox4.Controls.Add(rdoDOM);
            groupBox4.Location = new System.Drawing.Point(253, 110);
            groupBox4.Name = "groupBox4";
            groupBox4.Size = new System.Drawing.Size(238, 75);
            groupBox4.TabIndex = 0;
            groupBox4.TabStop = false;
            groupBox4.Text = "Excel Cell Value Option";
            // 
            // rdoSAX
            // 
            rdoSAX.AutoSize = true;
            rdoSAX.Location = new System.Drawing.Point(6, 22);
            rdoSAX.Name = "rdoSAX";
            rdoSAX.Size = new System.Drawing.Size(79, 19);
            rdoSAX.TabIndex = 3;
            rdoSAX.TabStop = true;
            rdoSAX.Text = "SAX Styles";
            rdoSAX.UseVisualStyleBackColor = true;
            // 
            // rdoDOM
            // 
            rdoDOM.AutoSize = true;
            rdoDOM.Location = new System.Drawing.Point(6, 47);
            rdoDOM.Name = "rdoDOM";
            rdoDOM.Size = new System.Drawing.Size(86, 19);
            rdoDOM.TabIndex = 4;
            rdoDOM.TabStop = true;
            rdoDOM.Text = "DOM Styles";
            rdoDOM.UseVisualStyleBackColor = true;
            // 
            // BtnOk
            // 
            BtnOk.Location = new System.Drawing.Point(337, 305);
            BtnOk.Name = "BtnOk";
            BtnOk.Size = new System.Drawing.Size(75, 23);
            BtnOk.TabIndex = 1;
            BtnOk.Text = "OK";
            BtnOk.UseVisualStyleBackColor = true;
            BtnOk.Click += BtnOk_Click;
            // 
            // BtnCancel
            // 
            BtnCancel.Location = new System.Drawing.Point(418, 305);
            BtnCancel.Name = "BtnCancel";
            BtnCancel.Size = new System.Drawing.Size(75, 23);
            BtnCancel.TabIndex = 2;
            BtnCancel.Text = "Cancel";
            BtnCancel.UseVisualStyleBackColor = true;
            BtnCancel.Click += BtnCancel_Click;
            // 
            // rdoUseCCGuid
            // 
            rdoUseCCGuid.AutoSize = true;
            rdoUseCCGuid.Location = new System.Drawing.Point(6, 47);
            rdoUseCCGuid.Name = "rdoUseCCGuid";
            rdoUseCCGuid.Size = new System.Drawing.Size(161, 19);
            rdoUseCCGuid.TabIndex = 4;
            rdoUseCCGuid.TabStop = true;
            rdoUseCCGuid.Text = "Use Content Control Guid";
            rdoUseCCGuid.UseVisualStyleBackColor = true;
            // 
            // rdoUseSPGuid
            // 
            rdoUseSPGuid.AutoSize = true;
            rdoUseSPGuid.Location = new System.Drawing.Point(6, 22);
            rdoUseSPGuid.Name = "rdoUseSPGuid";
            rdoUseSPGuid.Size = new System.Drawing.Size(201, 19);
            rdoUseSPGuid.TabIndex = 3;
            rdoUseSPGuid.TabStop = true;
            rdoUseSPGuid.Text = "Use SharePoint Custom Xml Guid";
            rdoUseSPGuid.UseVisualStyleBackColor = true;
            // 
            // rdoUserSelectedCC
            // 
            rdoUserSelectedCC.AutoSize = true;
            rdoUserSelectedCC.Location = new System.Drawing.Point(6, 72);
            rdoUserSelectedCC.Name = "rdoUserSelectedCC";
            rdoUserSelectedCC.Size = new System.Drawing.Size(214, 19);
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
            groupBox6.Location = new System.Drawing.Point(12, 191);
            groupBox6.Name = "groupBox6";
            groupBox6.Size = new System.Drawing.Size(235, 103);
            groupBox6.TabIndex = 5;
            groupBox6.TabStop = false;
            groupBox6.Text = "Fix Content Control Prefix Mappings";
            // 
            // FrmSettings
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(503, 340);
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
        private System.Windows.Forms.CheckBox ckbDisableAutoXmlColorFormatting;
        private System.Windows.Forms.CheckBox ckbResetIndentLevels;
    }
}