
namespace Office_File_Explorer.WinForms
{
    partial class FrmExcelModify
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmExcelModify));
            groupBox1 = new System.Windows.Forms.GroupBox();
            rdoConvertShareLinkToCanonicalLink = new System.Windows.Forms.RadioButton();
            rdoDelLink = new System.Windows.Forms.RadioButton();
            rdoDeleteSheet = new System.Windows.Forms.RadioButton();
            rdoConvertStrict = new System.Windows.Forms.RadioButton();
            rdoConvertToXlsm = new System.Windows.Forms.RadioButton();
            rdoDelEmbedLinks = new System.Windows.Forms.RadioButton();
            rdoDelComments = new System.Windows.Forms.RadioButton();
            rdoDelLinks = new System.Windows.Forms.RadioButton();
            BtnOk = new System.Windows.Forms.Button();
            BtnCancel = new System.Windows.Forms.Button();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(rdoConvertShareLinkToCanonicalLink);
            groupBox1.Controls.Add(rdoDelLink);
            groupBox1.Controls.Add(rdoDeleteSheet);
            groupBox1.Controls.Add(rdoConvertStrict);
            groupBox1.Controls.Add(rdoConvertToXlsm);
            groupBox1.Controls.Add(rdoDelEmbedLinks);
            groupBox1.Controls.Add(rdoDelComments);
            groupBox1.Controls.Add(rdoDelLinks);
            groupBox1.Location = new System.Drawing.Point(12, 3);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new System.Drawing.Size(270, 236);
            groupBox1.TabIndex = 0;
            groupBox1.TabStop = false;
            groupBox1.Text = "Select Content To Modify";
            // 
            // rdoConvertShareLinkToCanonicalLink
            // 
            rdoConvertShareLinkToCanonicalLink.AutoSize = true;
            rdoConvertShareLinkToCanonicalLink.Location = new System.Drawing.Point(6, 197);
            rdoConvertShareLinkToCanonicalLink.Name = "rdoConvertShareLinkToCanonicalLink";
            rdoConvertShareLinkToCanonicalLink.Size = new System.Drawing.Size(255, 19);
            rdoConvertShareLinkToCanonicalLink.TabIndex = 7;
            rdoConvertShareLinkToCanonicalLink.TabStop = true;
            rdoConvertShareLinkToCanonicalLink.Text = "Convert SharePoint Share Link To Canonical";
            rdoConvertShareLinkToCanonicalLink.UseVisualStyleBackColor = true;
            // 
            // rdoDelLink
            // 
            rdoDelLink.AutoSize = true;
            rdoDelLink.Location = new System.Drawing.Point(6, 22);
            rdoDelLink.Name = "rdoDelLink";
            rdoDelLink.Size = new System.Drawing.Size(167, 19);
            rdoDelLink.TabIndex = 6;
            rdoDelLink.TabStop = true;
            rdoDelLink.Text = "Delete Individual Hyperlink";
            rdoDelLink.UseVisualStyleBackColor = true;
            // 
            // rdoDeleteSheet
            // 
            rdoDeleteSheet.AutoSize = true;
            rdoDeleteSheet.Location = new System.Drawing.Point(6, 97);
            rdoDeleteSheet.Name = "rdoDeleteSheet";
            rdoDeleteSheet.Size = new System.Drawing.Size(90, 19);
            rdoDeleteSheet.TabIndex = 5;
            rdoDeleteSheet.TabStop = true;
            rdoDeleteSheet.Text = "Delete Sheet";
            rdoDeleteSheet.UseVisualStyleBackColor = true;
            // 
            // rdoConvertStrict
            // 
            rdoConvertStrict.AutoSize = true;
            rdoConvertStrict.Location = new System.Drawing.Point(6, 172);
            rdoConvertStrict.Name = "rdoConvertStrict";
            rdoConvertStrict.Size = new System.Drawing.Size(136, 19);
            rdoConvertStrict.TabIndex = 4;
            rdoConvertStrict.TabStop = true;
            rdoConvertStrict.Text = "Convert Strict To Xlsx";
            rdoConvertStrict.UseVisualStyleBackColor = true;
            // 
            // rdoConvertToXlsm
            // 
            rdoConvertToXlsm.AutoSize = true;
            rdoConvertToXlsm.Location = new System.Drawing.Point(6, 147);
            rdoConvertToXlsm.Name = "rdoConvertToXlsm";
            rdoConvertToXlsm.Size = new System.Drawing.Size(135, 19);
            rdoConvertToXlsm.TabIndex = 3;
            rdoConvertToXlsm.TabStop = true;
            rdoConvertToXlsm.Text = "Convert Xlsm To Xlsx";
            rdoConvertToXlsm.UseVisualStyleBackColor = true;
            // 
            // rdoDelEmbedLinks
            // 
            rdoDelEmbedLinks.AutoSize = true;
            rdoDelEmbedLinks.Location = new System.Drawing.Point(6, 122);
            rdoDelEmbedLinks.Name = "rdoDelEmbedLinks";
            rdoDelEmbedLinks.Size = new System.Drawing.Size(148, 19);
            rdoDelEmbedLinks.TabIndex = 2;
            rdoDelEmbedLinks.TabStop = true;
            rdoDelEmbedLinks.Text = "Delete Embedded Links";
            rdoDelEmbedLinks.UseVisualStyleBackColor = true;
            // 
            // rdoDelComments
            // 
            rdoDelComments.AutoSize = true;
            rdoDelComments.Location = new System.Drawing.Point(6, 72);
            rdoDelComments.Name = "rdoDelComments";
            rdoDelComments.Size = new System.Drawing.Size(120, 19);
            rdoDelComments.TabIndex = 1;
            rdoDelComments.TabStop = true;
            rdoDelComments.Text = "Delete Comments";
            rdoDelComments.UseVisualStyleBackColor = true;
            // 
            // rdoDelLinks
            // 
            rdoDelLinks.AutoSize = true;
            rdoDelLinks.Location = new System.Drawing.Point(6, 47);
            rdoDelLinks.Name = "rdoDelLinks";
            rdoDelLinks.Size = new System.Drawing.Size(134, 19);
            rdoDelLinks.TabIndex = 0;
            rdoDelLinks.TabStop = true;
            rdoDelLinks.Text = "Delete All Hyperlinks";
            rdoDelLinks.UseVisualStyleBackColor = true;
            // 
            // BtnOk
            // 
            BtnOk.Location = new System.Drawing.Point(126, 245);
            BtnOk.Name = "BtnOk";
            BtnOk.Size = new System.Drawing.Size(75, 23);
            BtnOk.TabIndex = 5;
            BtnOk.Text = "OK";
            BtnOk.UseVisualStyleBackColor = true;
            BtnOk.Click += BtnOk_Click;
            // 
            // BtnCancel
            // 
            BtnCancel.Location = new System.Drawing.Point(207, 245);
            BtnCancel.Name = "BtnCancel";
            BtnCancel.Size = new System.Drawing.Size(75, 23);
            BtnCancel.TabIndex = 6;
            BtnCancel.Text = "Cancel";
            BtnCancel.UseVisualStyleBackColor = true;
            BtnCancel.Click += BtnCancel_Click;
            // 
            // FrmExcelModify
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(294, 286);
            Controls.Add(BtnOk);
            Controls.Add(BtnCancel);
            Controls.Add(groupBox1);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            Icon = (System.Drawing.Icon)resources.GetObject("$this.Icon");
            KeyPreview = true;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "FrmExcelModify";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            Text = "Modify Excel Content";
            KeyDown += FrmExcelModify_KeyDown;
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rdoConvertStrict;
        private System.Windows.Forms.RadioButton rdoConvertToXlsm;
        private System.Windows.Forms.RadioButton rdoDelEmbedLinks;
        private System.Windows.Forms.RadioButton rdoDelComments;
        private System.Windows.Forms.RadioButton rdoDelLinks;
        private System.Windows.Forms.Button BtnOk;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.RadioButton rdoDeleteSheet;
        private System.Windows.Forms.RadioButton rdoDelLink;
        private System.Windows.Forms.RadioButton rdoConvertShareLinkToCanonicalLink;
    }
}