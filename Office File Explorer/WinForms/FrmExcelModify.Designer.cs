
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rdoDeleteSheet = new System.Windows.Forms.RadioButton();
            this.rdoConvertStrict = new System.Windows.Forms.RadioButton();
            this.rdoConvertToXlsm = new System.Windows.Forms.RadioButton();
            this.rdoDelEmbedLinks = new System.Windows.Forms.RadioButton();
            this.rdoDelComments = new System.Windows.Forms.RadioButton();
            this.rdoDelLinks = new System.Windows.Forms.RadioButton();
            this.BtnOk = new System.Windows.Forms.Button();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rdoDeleteSheet);
            this.groupBox1.Controls.Add(this.rdoConvertStrict);
            this.groupBox1.Controls.Add(this.rdoConvertToXlsm);
            this.groupBox1.Controls.Add(this.rdoDelEmbedLinks);
            this.groupBox1.Controls.Add(this.rdoDelComments);
            this.groupBox1.Controls.Add(this.rdoDelLinks);
            this.groupBox1.Location = new System.Drawing.Point(12, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(270, 183);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Select Content To Modify";
            // 
            // rdoDeleteSheet
            // 
            this.rdoDeleteSheet.AutoSize = true;
            this.rdoDeleteSheet.Location = new System.Drawing.Point(6, 72);
            this.rdoDeleteSheet.Name = "rdoDeleteSheet";
            this.rdoDeleteSheet.Size = new System.Drawing.Size(90, 19);
            this.rdoDeleteSheet.TabIndex = 5;
            this.rdoDeleteSheet.TabStop = true;
            this.rdoDeleteSheet.Text = "Delete Sheet";
            this.rdoDeleteSheet.UseVisualStyleBackColor = true;
            // 
            // rdoConvertStrict
            // 
            this.rdoConvertStrict.AutoSize = true;
            this.rdoConvertStrict.Location = new System.Drawing.Point(6, 147);
            this.rdoConvertStrict.Name = "rdoConvertStrict";
            this.rdoConvertStrict.Size = new System.Drawing.Size(136, 19);
            this.rdoConvertStrict.TabIndex = 4;
            this.rdoConvertStrict.TabStop = true;
            this.rdoConvertStrict.Text = "Convert Strict To Xlsx";
            this.rdoConvertStrict.UseVisualStyleBackColor = true;
            // 
            // rdoConvertToXlsm
            // 
            this.rdoConvertToXlsm.AutoSize = true;
            this.rdoConvertToXlsm.Location = new System.Drawing.Point(6, 122);
            this.rdoConvertToXlsm.Name = "rdoConvertToXlsm";
            this.rdoConvertToXlsm.Size = new System.Drawing.Size(135, 19);
            this.rdoConvertToXlsm.TabIndex = 3;
            this.rdoConvertToXlsm.TabStop = true;
            this.rdoConvertToXlsm.Text = "Convert Xlsm To Xlsx";
            this.rdoConvertToXlsm.UseVisualStyleBackColor = true;
            // 
            // rdoDelEmbedLinks
            // 
            this.rdoDelEmbedLinks.AutoSize = true;
            this.rdoDelEmbedLinks.Location = new System.Drawing.Point(6, 97);
            this.rdoDelEmbedLinks.Name = "rdoDelEmbedLinks";
            this.rdoDelEmbedLinks.Size = new System.Drawing.Size(148, 19);
            this.rdoDelEmbedLinks.TabIndex = 2;
            this.rdoDelEmbedLinks.TabStop = true;
            this.rdoDelEmbedLinks.Text = "Delete Embedded Links";
            this.rdoDelEmbedLinks.UseVisualStyleBackColor = true;
            // 
            // rdoDelComments
            // 
            this.rdoDelComments.AutoSize = true;
            this.rdoDelComments.Location = new System.Drawing.Point(6, 47);
            this.rdoDelComments.Name = "rdoDelComments";
            this.rdoDelComments.Size = new System.Drawing.Size(120, 19);
            this.rdoDelComments.TabIndex = 1;
            this.rdoDelComments.TabStop = true;
            this.rdoDelComments.Text = "Delete Comments";
            this.rdoDelComments.UseVisualStyleBackColor = true;
            // 
            // rdoDelLinks
            // 
            this.rdoDelLinks.AutoSize = true;
            this.rdoDelLinks.Location = new System.Drawing.Point(6, 22);
            this.rdoDelLinks.Name = "rdoDelLinks";
            this.rdoDelLinks.Size = new System.Drawing.Size(88, 19);
            this.rdoDelLinks.TabIndex = 0;
            this.rdoDelLinks.TabStop = true;
            this.rdoDelLinks.Text = "Delete Links";
            this.rdoDelLinks.UseVisualStyleBackColor = true;
            // 
            // BtnOk
            // 
            this.BtnOk.Location = new System.Drawing.Point(125, 192);
            this.BtnOk.Name = "BtnOk";
            this.BtnOk.Size = new System.Drawing.Size(75, 23);
            this.BtnOk.TabIndex = 5;
            this.BtnOk.Text = "Ok";
            this.BtnOk.UseVisualStyleBackColor = true;
            this.BtnOk.Click += new System.EventHandler(this.BtnOk_Click);
            // 
            // BtnCancel
            // 
            this.BtnCancel.Location = new System.Drawing.Point(206, 192);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(75, 23);
            this.BtnCancel.TabIndex = 6;
            this.BtnCancel.Text = "Cancel";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // FrmExcelModify
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(294, 225);
            this.Controls.Add(this.BtnOk);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmExcelModify";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Modify Excel Content";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

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
    }
}