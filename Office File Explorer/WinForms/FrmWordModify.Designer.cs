
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rdoRemoveCustomTitleProp = new System.Windows.Forms.RadioButton();
            this.rdoRemovePII = new System.Windows.Forms.RadioButton();
            this.rdoAcceptRevisions = new System.Windows.Forms.RadioButton();
            this.rdoChangeDefTemp = new System.Windows.Forms.RadioButton();
            this.rdoSetPrint = new System.Windows.Forms.RadioButton();
            this.rdoConvertDocmToDocx = new System.Windows.Forms.RadioButton();
            this.rdoDelOrphanStyles = new System.Windows.Forms.RadioButton();
            this.rdoDelOrhpanLT = new System.Windows.Forms.RadioButton();
            this.rdoDelEndnotes = new System.Windows.Forms.RadioButton();
            this.rdoDelFootnotes = new System.Windows.Forms.RadioButton();
            this.rdoDelHiddenText = new System.Windows.Forms.RadioButton();
            this.rdoDelComments = new System.Windows.Forms.RadioButton();
            this.rdoDelPgBreaks = new System.Windows.Forms.RadioButton();
            this.rdoDelHF = new System.Windows.Forms.RadioButton();
            this.BtnOk = new System.Windows.Forms.Button();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.rdoUpdateNamespaces = new System.Windows.Forms.RadioButton();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rdoUpdateNamespaces);
            this.groupBox1.Controls.Add(this.rdoRemoveCustomTitleProp);
            this.groupBox1.Controls.Add(this.rdoRemovePII);
            this.groupBox1.Controls.Add(this.rdoAcceptRevisions);
            this.groupBox1.Controls.Add(this.rdoChangeDefTemp);
            this.groupBox1.Controls.Add(this.rdoSetPrint);
            this.groupBox1.Controls.Add(this.rdoConvertDocmToDocx);
            this.groupBox1.Controls.Add(this.rdoDelOrphanStyles);
            this.groupBox1.Controls.Add(this.rdoDelOrhpanLT);
            this.groupBox1.Controls.Add(this.rdoDelEndnotes);
            this.groupBox1.Controls.Add(this.rdoDelFootnotes);
            this.groupBox1.Controls.Add(this.rdoDelHiddenText);
            this.groupBox1.Controls.Add(this.rdoDelComments);
            this.groupBox1.Controls.Add(this.rdoDelPgBreaks);
            this.groupBox1.Controls.Add(this.rdoDelHF);
            this.groupBox1.Location = new System.Drawing.Point(12, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(389, 226);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Select Content to Modify";
            // 
            // rdoRemoveCustomTitleProp
            // 
            this.rdoRemoveCustomTitleProp.AutoSize = true;
            this.rdoRemoveCustomTitleProp.Location = new System.Drawing.Point(211, 169);
            this.rdoRemoveCustomTitleProp.Name = "rdoRemoveCustomTitleProp";
            this.rdoRemoveCustomTitleProp.Size = new System.Drawing.Size(156, 19);
            this.rdoRemoveCustomTitleProp.TabIndex = 2;
            this.rdoRemoveCustomTitleProp.TabStop = true;
            this.rdoRemoveCustomTitleProp.Text = "Delete Custom Title Prop";
            this.rdoRemoveCustomTitleProp.UseVisualStyleBackColor = true;
            // 
            // rdoRemovePII
            // 
            this.rdoRemovePII.AutoSize = true;
            this.rdoRemovePII.Location = new System.Drawing.Point(211, 144);
            this.rdoRemovePII.Name = "rdoRemovePII";
            this.rdoRemovePII.Size = new System.Drawing.Size(84, 19);
            this.rdoRemovePII.TabIndex = 11;
            this.rdoRemovePII.TabStop = true;
            this.rdoRemovePII.Text = "Remove PII";
            this.rdoRemovePII.UseVisualStyleBackColor = true;
            // 
            // rdoAcceptRevisions
            // 
            this.rdoAcceptRevisions.AutoSize = true;
            this.rdoAcceptRevisions.Location = new System.Drawing.Point(211, 119);
            this.rdoAcceptRevisions.Name = "rdoAcceptRevisions";
            this.rdoAcceptRevisions.Size = new System.Drawing.Size(131, 19);
            this.rdoAcceptRevisions.TabIndex = 10;
            this.rdoAcceptRevisions.TabStop = true;
            this.rdoAcceptRevisions.Text = "Accept All Revisions";
            this.rdoAcceptRevisions.UseVisualStyleBackColor = true;
            // 
            // rdoChangeDefTemp
            // 
            this.rdoChangeDefTemp.AutoSize = true;
            this.rdoChangeDefTemp.Location = new System.Drawing.Point(211, 94);
            this.rdoChangeDefTemp.Name = "rdoChangeDefTemp";
            this.rdoChangeDefTemp.Size = new System.Drawing.Size(158, 19);
            this.rdoChangeDefTemp.TabIndex = 9;
            this.rdoChangeDefTemp.TabStop = true;
            this.rdoChangeDefTemp.Text = "Change Default Template";
            this.rdoChangeDefTemp.UseVisualStyleBackColor = true;
            // 
            // rdoSetPrint
            // 
            this.rdoSetPrint.AutoSize = true;
            this.rdoSetPrint.Location = new System.Drawing.Point(211, 69);
            this.rdoSetPrint.Name = "rdoSetPrint";
            this.rdoSetPrint.Size = new System.Drawing.Size(132, 19);
            this.rdoSetPrint.TabIndex = 2;
            this.rdoSetPrint.TabStop = true;
            this.rdoSetPrint.Text = "Set Print Orientation";
            this.rdoSetPrint.UseVisualStyleBackColor = true;
            // 
            // rdoConvertDocmToDocx
            // 
            this.rdoConvertDocmToDocx.AutoSize = true;
            this.rdoConvertDocmToDocx.Location = new System.Drawing.Point(211, 44);
            this.rdoConvertDocmToDocx.Name = "rdoConvertDocmToDocx";
            this.rdoConvertDocmToDocx.Size = new System.Drawing.Size(147, 19);
            this.rdoConvertDocmToDocx.TabIndex = 8;
            this.rdoConvertDocmToDocx.TabStop = true;
            this.rdoConvertDocmToDocx.Text = "Convert Docm To Docx";
            this.rdoConvertDocmToDocx.UseVisualStyleBackColor = true;
            // 
            // rdoDelOrphanStyles
            // 
            this.rdoDelOrphanStyles.AutoSize = true;
            this.rdoDelOrphanStyles.Location = new System.Drawing.Point(11, 169);
            this.rdoDelOrphanStyles.Name = "rdoDelOrphanStyles";
            this.rdoDelOrphanStyles.Size = new System.Drawing.Size(134, 19);
            this.rdoDelOrphanStyles.TabIndex = 7;
            this.rdoDelOrphanStyles.TabStop = true;
            this.rdoDelOrphanStyles.Text = "Delete Unused Styles";
            this.rdoDelOrphanStyles.UseVisualStyleBackColor = true;
            // 
            // rdoDelOrhpanLT
            // 
            this.rdoDelOrhpanLT.AutoSize = true;
            this.rdoDelOrhpanLT.Location = new System.Drawing.Point(211, 19);
            this.rdoDelOrhpanLT.Name = "rdoDelOrhpanLT";
            this.rdoDelOrhpanLT.Size = new System.Drawing.Size(178, 19);
            this.rdoDelOrhpanLT.TabIndex = 6;
            this.rdoDelOrhpanLT.TabStop = true;
            this.rdoDelOrhpanLT.Text = "Delete Unused List Templates";
            this.rdoDelOrhpanLT.UseVisualStyleBackColor = true;
            // 
            // rdoDelEndnotes
            // 
            this.rdoDelEndnotes.AutoSize = true;
            this.rdoDelEndnotes.Location = new System.Drawing.Point(11, 144);
            this.rdoDelEndnotes.Name = "rdoDelEndnotes";
            this.rdoDelEndnotes.Size = new System.Drawing.Size(110, 19);
            this.rdoDelEndnotes.TabIndex = 5;
            this.rdoDelEndnotes.TabStop = true;
            this.rdoDelEndnotes.Text = "Delete Endnotes";
            this.rdoDelEndnotes.UseVisualStyleBackColor = true;
            // 
            // rdoDelFootnotes
            // 
            this.rdoDelFootnotes.AutoSize = true;
            this.rdoDelFootnotes.Location = new System.Drawing.Point(11, 119);
            this.rdoDelFootnotes.Name = "rdoDelFootnotes";
            this.rdoDelFootnotes.Size = new System.Drawing.Size(114, 19);
            this.rdoDelFootnotes.TabIndex = 4;
            this.rdoDelFootnotes.TabStop = true;
            this.rdoDelFootnotes.Text = "Delete Footnotes";
            this.rdoDelFootnotes.UseVisualStyleBackColor = true;
            // 
            // rdoDelHiddenText
            // 
            this.rdoDelHiddenText.AutoSize = true;
            this.rdoDelHiddenText.Location = new System.Drawing.Point(11, 94);
            this.rdoDelHiddenText.Name = "rdoDelHiddenText";
            this.rdoDelHiddenText.Size = new System.Drawing.Size(124, 19);
            this.rdoDelHiddenText.TabIndex = 3;
            this.rdoDelHiddenText.TabStop = true;
            this.rdoDelHiddenText.Text = "Delete Hidden Text";
            this.rdoDelHiddenText.UseVisualStyleBackColor = true;
            // 
            // rdoDelComments
            // 
            this.rdoDelComments.AutoSize = true;
            this.rdoDelComments.Location = new System.Drawing.Point(11, 69);
            this.rdoDelComments.Name = "rdoDelComments";
            this.rdoDelComments.Size = new System.Drawing.Size(120, 19);
            this.rdoDelComments.TabIndex = 2;
            this.rdoDelComments.TabStop = true;
            this.rdoDelComments.Text = "Delete Comments";
            this.rdoDelComments.UseVisualStyleBackColor = true;
            // 
            // rdoDelPgBreaks
            // 
            this.rdoDelPgBreaks.AutoSize = true;
            this.rdoDelPgBreaks.Location = new System.Drawing.Point(11, 44);
            this.rdoDelPgBreaks.Name = "rdoDelPgBreaks";
            this.rdoDelPgBreaks.Size = new System.Drawing.Size(124, 19);
            this.rdoDelPgBreaks.TabIndex = 1;
            this.rdoDelPgBreaks.TabStop = true;
            this.rdoDelPgBreaks.Text = "Delete Page Breaks";
            this.rdoDelPgBreaks.UseVisualStyleBackColor = true;
            // 
            // rdoDelHF
            // 
            this.rdoDelHF.AutoSize = true;
            this.rdoDelHF.Location = new System.Drawing.Point(11, 19);
            this.rdoDelHF.Name = "rdoDelHF";
            this.rdoDelHF.Size = new System.Drawing.Size(169, 19);
            this.rdoDelHF.TabIndex = 0;
            this.rdoDelHF.TabStop = true;
            this.rdoDelHF.Text = "Delete Headers and Footers";
            this.rdoDelHF.UseVisualStyleBackColor = true;
            // 
            // BtnOk
            // 
            this.BtnOk.Location = new System.Drawing.Point(245, 232);
            this.BtnOk.Name = "BtnOk";
            this.BtnOk.Size = new System.Drawing.Size(75, 23);
            this.BtnOk.TabIndex = 0;
            this.BtnOk.Text = "Ok";
            this.BtnOk.UseVisualStyleBackColor = true;
            this.BtnOk.Click += new System.EventHandler(this.BtnOk_Click);
            // 
            // BtnCancel
            // 
            this.BtnCancel.Location = new System.Drawing.Point(326, 232);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(75, 23);
            this.BtnCancel.TabIndex = 1;
            this.BtnCancel.Text = "Cancel";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // rdoUpdateNamespaces
            // 
            this.rdoUpdateNamespaces.AutoSize = true;
            this.rdoUpdateNamespaces.Location = new System.Drawing.Point(11, 194);
            this.rdoUpdateNamespaces.Name = "rdoUpdateNamespaces";
            this.rdoUpdateNamespaces.Size = new System.Drawing.Size(191, 19);
            this.rdoUpdateNamespaces.TabIndex = 12;
            this.rdoUpdateNamespaces.TabStop = true;
            this.rdoUpdateNamespaces.Text = "Update Quick Part Namespaces";
            this.rdoUpdateNamespaces.UseVisualStyleBackColor = true;
            // 
            // FrmWordModify
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(409, 267);
            this.Controls.Add(this.BtnOk);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmWordModify";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Modify Word Content";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

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
    }
}