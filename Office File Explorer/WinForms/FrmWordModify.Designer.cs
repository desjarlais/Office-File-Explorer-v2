
namespace Office_File_Explorer.WinForms
{
    partial class FrmWordModify
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmWordModify));
            groupBox1 = new System.Windows.Forms.GroupBox();
            RdoRemoveDuplicateAuthors = new System.Windows.Forms.RadioButton();
            rdoDelBookmarks = new System.Windows.Forms.RadioButton();
            rdoUpdateNamespaces = new System.Windows.Forms.RadioButton();
            rdoRemoveCustomTitleProp = new System.Windows.Forms.RadioButton();
            rdoRemovePII = new System.Windows.Forms.RadioButton();
            rdoAcceptRevisions = new System.Windows.Forms.RadioButton();
            rdoChangeDefTemp = new System.Windows.Forms.RadioButton();
            rdoSetPrint = new System.Windows.Forms.RadioButton();
            rdoConvertDocmToDocx = new System.Windows.Forms.RadioButton();
            rdoDelOrphanStyles = new System.Windows.Forms.RadioButton();
            rdoDelOrhpanLT = new System.Windows.Forms.RadioButton();
            rdoDelEndnotes = new System.Windows.Forms.RadioButton();
            rdoDelFootnotes = new System.Windows.Forms.RadioButton();
            rdoDelHiddenText = new System.Windows.Forms.RadioButton();
            rdoDelComments = new System.Windows.Forms.RadioButton();
            rdoDelPgBreaks = new System.Windows.Forms.RadioButton();
            rdoDelHF = new System.Windows.Forms.RadioButton();
            BtnOk = new System.Windows.Forms.Button();
            BtnCancel = new System.Windows.Forms.Button();
            rdoDeleteDupeSPCustomXml = new System.Windows.Forms.RadioButton();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(rdoDeleteDupeSPCustomXml);
            groupBox1.Controls.Add(RdoRemoveDuplicateAuthors);
            groupBox1.Controls.Add(rdoDelBookmarks);
            groupBox1.Controls.Add(rdoUpdateNamespaces);
            groupBox1.Controls.Add(rdoRemoveCustomTitleProp);
            groupBox1.Controls.Add(rdoRemovePII);
            groupBox1.Controls.Add(rdoAcceptRevisions);
            groupBox1.Controls.Add(rdoChangeDefTemp);
            groupBox1.Controls.Add(rdoSetPrint);
            groupBox1.Controls.Add(rdoConvertDocmToDocx);
            groupBox1.Controls.Add(rdoDelOrphanStyles);
            groupBox1.Controls.Add(rdoDelOrhpanLT);
            groupBox1.Controls.Add(rdoDelEndnotes);
            groupBox1.Controls.Add(rdoDelFootnotes);
            groupBox1.Controls.Add(rdoDelHiddenText);
            groupBox1.Controls.Add(rdoDelComments);
            groupBox1.Controls.Add(rdoDelPgBreaks);
            groupBox1.Controls.Add(rdoDelHF);
            groupBox1.Location = new System.Drawing.Point(12, 0);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new System.Drawing.Size(414, 248);
            groupBox1.TabIndex = 0;
            groupBox1.TabStop = false;
            groupBox1.Text = "Select Content to Modify";
            // 
            // RdoRemoveDuplicateAuthors
            // 
            RdoRemoveDuplicateAuthors.AutoSize = true;
            RdoRemoveDuplicateAuthors.Location = new System.Drawing.Point(11, 219);
            RdoRemoveDuplicateAuthors.Name = "RdoRemoveDuplicateAuthors";
            RdoRemoveDuplicateAuthors.Size = new System.Drawing.Size(156, 19);
            RdoRemoveDuplicateAuthors.TabIndex = 13;
            RdoRemoveDuplicateAuthors.TabStop = true;
            RdoRemoveDuplicateAuthors.Text = "Delete Duplicate Authors";
            RdoRemoveDuplicateAuthors.UseVisualStyleBackColor = true;
            // 
            // rdoDelBookmarks
            // 
            rdoDelBookmarks.AutoSize = true;
            rdoDelBookmarks.Location = new System.Drawing.Point(11, 194);
            rdoDelBookmarks.Name = "rdoDelBookmarks";
            rdoDelBookmarks.Size = new System.Drawing.Size(120, 19);
            rdoDelBookmarks.TabIndex = 2;
            rdoDelBookmarks.TabStop = true;
            rdoDelBookmarks.Text = "Delete Bookmarks";
            rdoDelBookmarks.UseVisualStyleBackColor = true;
            // 
            // rdoUpdateNamespaces
            // 
            rdoUpdateNamespaces.AutoSize = true;
            rdoUpdateNamespaces.Location = new System.Drawing.Point(172, 194);
            rdoUpdateNamespaces.Name = "rdoUpdateNamespaces";
            rdoUpdateNamespaces.Size = new System.Drawing.Size(191, 19);
            rdoUpdateNamespaces.TabIndex = 12;
            rdoUpdateNamespaces.TabStop = true;
            rdoUpdateNamespaces.Text = "Update Quick Part Namespaces";
            rdoUpdateNamespaces.UseVisualStyleBackColor = true;
            // 
            // rdoRemoveCustomTitleProp
            // 
            rdoRemoveCustomTitleProp.AutoSize = true;
            rdoRemoveCustomTitleProp.Location = new System.Drawing.Point(172, 169);
            rdoRemoveCustomTitleProp.Name = "rdoRemoveCustomTitleProp";
            rdoRemoveCustomTitleProp.Size = new System.Drawing.Size(157, 19);
            rdoRemoveCustomTitleProp.TabIndex = 2;
            rdoRemoveCustomTitleProp.TabStop = true;
            rdoRemoveCustomTitleProp.Text = "Delete Custom Title Prop";
            rdoRemoveCustomTitleProp.UseVisualStyleBackColor = true;
            // 
            // rdoRemovePII
            // 
            rdoRemovePII.AutoSize = true;
            rdoRemovePII.Location = new System.Drawing.Point(11, 19);
            rdoRemovePII.Name = "rdoRemovePII";
            rdoRemovePII.Size = new System.Drawing.Size(84, 19);
            rdoRemovePII.TabIndex = 11;
            rdoRemovePII.TabStop = true;
            rdoRemovePII.Text = "Remove PII";
            rdoRemovePII.UseVisualStyleBackColor = true;
            // 
            // rdoAcceptRevisions
            // 
            rdoAcceptRevisions.AutoSize = true;
            rdoAcceptRevisions.Location = new System.Drawing.Point(172, 119);
            rdoAcceptRevisions.Name = "rdoAcceptRevisions";
            rdoAcceptRevisions.Size = new System.Drawing.Size(131, 19);
            rdoAcceptRevisions.TabIndex = 10;
            rdoAcceptRevisions.TabStop = true;
            rdoAcceptRevisions.Text = "Accept All Revisions";
            rdoAcceptRevisions.UseVisualStyleBackColor = true;
            // 
            // rdoChangeDefTemp
            // 
            rdoChangeDefTemp.AutoSize = true;
            rdoChangeDefTemp.Location = new System.Drawing.Point(172, 94);
            rdoChangeDefTemp.Name = "rdoChangeDefTemp";
            rdoChangeDefTemp.Size = new System.Drawing.Size(159, 19);
            rdoChangeDefTemp.TabIndex = 9;
            rdoChangeDefTemp.TabStop = true;
            rdoChangeDefTemp.Text = "Change Default Template";
            rdoChangeDefTemp.UseVisualStyleBackColor = true;
            // 
            // rdoSetPrint
            // 
            rdoSetPrint.AutoSize = true;
            rdoSetPrint.Location = new System.Drawing.Point(172, 69);
            rdoSetPrint.Name = "rdoSetPrint";
            rdoSetPrint.Size = new System.Drawing.Size(132, 19);
            rdoSetPrint.TabIndex = 2;
            rdoSetPrint.TabStop = true;
            rdoSetPrint.Text = "Set Print Orientation";
            rdoSetPrint.UseVisualStyleBackColor = true;
            // 
            // rdoConvertDocmToDocx
            // 
            rdoConvertDocmToDocx.AutoSize = true;
            rdoConvertDocmToDocx.Location = new System.Drawing.Point(172, 44);
            rdoConvertDocmToDocx.Name = "rdoConvertDocmToDocx";
            rdoConvertDocmToDocx.Size = new System.Drawing.Size(147, 19);
            rdoConvertDocmToDocx.TabIndex = 8;
            rdoConvertDocmToDocx.TabStop = true;
            rdoConvertDocmToDocx.Text = "Convert Docm To Docx";
            rdoConvertDocmToDocx.UseVisualStyleBackColor = true;
            // 
            // rdoDelOrphanStyles
            // 
            rdoDelOrphanStyles.AutoSize = true;
            rdoDelOrphanStyles.Location = new System.Drawing.Point(11, 169);
            rdoDelOrphanStyles.Name = "rdoDelOrphanStyles";
            rdoDelOrphanStyles.Size = new System.Drawing.Size(134, 19);
            rdoDelOrphanStyles.TabIndex = 7;
            rdoDelOrphanStyles.TabStop = true;
            rdoDelOrphanStyles.Text = "Delete Unused Styles";
            rdoDelOrphanStyles.UseVisualStyleBackColor = true;
            // 
            // rdoDelOrhpanLT
            // 
            rdoDelOrhpanLT.AutoSize = true;
            rdoDelOrhpanLT.Location = new System.Drawing.Point(172, 19);
            rdoDelOrhpanLT.Name = "rdoDelOrhpanLT";
            rdoDelOrhpanLT.Size = new System.Drawing.Size(179, 19);
            rdoDelOrhpanLT.TabIndex = 6;
            rdoDelOrhpanLT.TabStop = true;
            rdoDelOrhpanLT.Text = "Delete Unused List Templates";
            rdoDelOrhpanLT.UseVisualStyleBackColor = true;
            // 
            // rdoDelEndnotes
            // 
            rdoDelEndnotes.AutoSize = true;
            rdoDelEndnotes.Location = new System.Drawing.Point(11, 144);
            rdoDelEndnotes.Name = "rdoDelEndnotes";
            rdoDelEndnotes.Size = new System.Drawing.Size(110, 19);
            rdoDelEndnotes.TabIndex = 5;
            rdoDelEndnotes.TabStop = true;
            rdoDelEndnotes.Text = "Delete Endnotes";
            rdoDelEndnotes.UseVisualStyleBackColor = true;
            // 
            // rdoDelFootnotes
            // 
            rdoDelFootnotes.AutoSize = true;
            rdoDelFootnotes.Location = new System.Drawing.Point(11, 119);
            rdoDelFootnotes.Name = "rdoDelFootnotes";
            rdoDelFootnotes.Size = new System.Drawing.Size(114, 19);
            rdoDelFootnotes.TabIndex = 4;
            rdoDelFootnotes.TabStop = true;
            rdoDelFootnotes.Text = "Delete Footnotes";
            rdoDelFootnotes.UseVisualStyleBackColor = true;
            // 
            // rdoDelHiddenText
            // 
            rdoDelHiddenText.AutoSize = true;
            rdoDelHiddenText.Location = new System.Drawing.Point(11, 94);
            rdoDelHiddenText.Name = "rdoDelHiddenText";
            rdoDelHiddenText.Size = new System.Drawing.Size(124, 19);
            rdoDelHiddenText.TabIndex = 3;
            rdoDelHiddenText.TabStop = true;
            rdoDelHiddenText.Text = "Delete Hidden Text";
            rdoDelHiddenText.UseVisualStyleBackColor = true;
            // 
            // rdoDelComments
            // 
            rdoDelComments.AutoSize = true;
            rdoDelComments.Location = new System.Drawing.Point(11, 69);
            rdoDelComments.Name = "rdoDelComments";
            rdoDelComments.Size = new System.Drawing.Size(120, 19);
            rdoDelComments.TabIndex = 2;
            rdoDelComments.TabStop = true;
            rdoDelComments.Text = "Delete Comments";
            rdoDelComments.UseVisualStyleBackColor = true;
            // 
            // rdoDelPgBreaks
            // 
            rdoDelPgBreaks.AutoSize = true;
            rdoDelPgBreaks.Location = new System.Drawing.Point(11, 44);
            rdoDelPgBreaks.Name = "rdoDelPgBreaks";
            rdoDelPgBreaks.Size = new System.Drawing.Size(124, 19);
            rdoDelPgBreaks.TabIndex = 1;
            rdoDelPgBreaks.TabStop = true;
            rdoDelPgBreaks.Text = "Delete Page Breaks";
            rdoDelPgBreaks.UseVisualStyleBackColor = true;
            // 
            // rdoDelHF
            // 
            rdoDelHF.AutoSize = true;
            rdoDelHF.Location = new System.Drawing.Point(172, 144);
            rdoDelHF.Name = "rdoDelHF";
            rdoDelHF.Size = new System.Drawing.Size(169, 19);
            rdoDelHF.TabIndex = 0;
            rdoDelHF.TabStop = true;
            rdoDelHF.Text = "Delete Headers and Footers";
            rdoDelHF.UseVisualStyleBackColor = true;
            // 
            // BtnOk
            // 
            BtnOk.Location = new System.Drawing.Point(270, 250);
            BtnOk.Name = "BtnOk";
            BtnOk.Size = new System.Drawing.Size(75, 23);
            BtnOk.TabIndex = 0;
            BtnOk.Text = "OK";
            BtnOk.UseVisualStyleBackColor = true;
            BtnOk.Click += BtnOk_Click;
            // 
            // BtnCancel
            // 
            BtnCancel.Location = new System.Drawing.Point(351, 250);
            BtnCancel.Name = "BtnCancel";
            BtnCancel.Size = new System.Drawing.Size(75, 23);
            BtnCancel.TabIndex = 1;
            BtnCancel.Text = "Cancel";
            BtnCancel.UseVisualStyleBackColor = true;
            BtnCancel.Click += BtnCancel_Click;
            // 
            // rdoDeleteDupeSPCustomXml
            // 
            rdoDeleteDupeSPCustomXml.AutoSize = true;
            rdoDeleteDupeSPCustomXml.Location = new System.Drawing.Point(172, 219);
            rdoDeleteDupeSPCustomXml.Name = "rdoDeleteDupeSPCustomXml";
            rdoDeleteDupeSPCustomXml.Size = new System.Drawing.Size(240, 19);
            rdoDeleteDupeSPCustomXml.TabIndex = 2;
            rdoDeleteDupeSPCustomXml.TabStop = true;
            rdoDeleteDupeSPCustomXml.Text = "Delete Duplicate SharePoint Custom Xml";
            rdoDeleteDupeSPCustomXml.UseVisualStyleBackColor = true;
            // 
            // FrmWordModify
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(432, 280);
            Controls.Add(BtnOk);
            Controls.Add(BtnCancel);
            Controls.Add(groupBox1);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            Icon = (System.Drawing.Icon)resources.GetObject("$this.Icon");
            KeyPreview = true;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "FrmWordModify";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            Text = "Modify Word Content";
            KeyDown += FrmWordModify_KeyDown;
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button BtnOk;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.RadioButton rdoConvertDocmToDocx;
        private System.Windows.Forms.RadioButton rdoDelOrphanStyles;
        private System.Windows.Forms.RadioButton rdoDelOrhpanLT;
        private System.Windows.Forms.RadioButton rdoDelEndnotes;
        private System.Windows.Forms.RadioButton rdoDelFootnotes;
        private System.Windows.Forms.RadioButton rdoDelHiddenText;
        private System.Windows.Forms.RadioButton rdoDelComments;
        private System.Windows.Forms.RadioButton rdoDelPgBreaks;
        private System.Windows.Forms.RadioButton rdoDelHF;
        private System.Windows.Forms.RadioButton rdoAcceptRevisions;
        private System.Windows.Forms.RadioButton rdoChangeDefTemp;
        private System.Windows.Forms.RadioButton rdoSetPrint;
        private System.Windows.Forms.RadioButton rdoRemovePII;
        private System.Windows.Forms.RadioButton rdoRemoveCustomTitleProp;
        private System.Windows.Forms.RadioButton rdoUpdateNamespaces;
        private System.Windows.Forms.RadioButton rdoDelBookmarks;
        private System.Windows.Forms.RadioButton RdoRemoveDuplicateAuthors;
        private System.Windows.Forms.RadioButton rdoDeleteDupeSPCustomXml;
    }
}