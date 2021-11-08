
namespace Office_File_Explorer.WinForms
{
    partial class FrmFixDocument
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmFixDocument));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rdoFixCommentHyperlinksW = new System.Windows.Forms.RadioButton();
            this.rdoFixHyperlinksW = new System.Windows.Forms.RadioButton();
            this.rdoFixCommentsW = new System.Windows.Forms.RadioButton();
            this.rdoFixTablePropsW = new System.Windows.Forms.RadioButton();
            this.rdoFixListTemplatesW = new System.Windows.Forms.RadioButton();
            this.rdoFixEndnotesW = new System.Windows.Forms.RadioButton();
            this.rdoFixRevisionsW = new System.Windows.Forms.RadioButton();
            this.rdoFixBookmarksW = new System.Windows.Forms.RadioButton();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.rdoFixNotesPageSizeCustomP = new System.Windows.Forms.RadioButton();
            this.rdoFixNotesPageSizeP = new System.Windows.Forms.RadioButton();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.rdoFixStrictX = new System.Windows.Forms.RadioButton();
            this.BtnOk = new System.Windows.Forms.Button();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rdoFixCommentHyperlinksW);
            this.groupBox1.Controls.Add(this.rdoFixHyperlinksW);
            this.groupBox1.Controls.Add(this.rdoFixCommentsW);
            this.groupBox1.Controls.Add(this.rdoFixTablePropsW);
            this.groupBox1.Controls.Add(this.rdoFixListTemplatesW);
            this.groupBox1.Controls.Add(this.rdoFixEndnotesW);
            this.groupBox1.Controls.Add(this.rdoFixRevisionsW);
            this.groupBox1.Controls.Add(this.rdoFixBookmarksW);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(200, 231);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Word Document";
            // 
            // rdoFixCommentHyperlinksW
            // 
            this.rdoFixCommentHyperlinksW.AutoSize = true;
            this.rdoFixCommentHyperlinksW.Enabled = false;
            this.rdoFixCommentHyperlinksW.Location = new System.Drawing.Point(6, 197);
            this.rdoFixCommentHyperlinksW.Name = "rdoFixCommentHyperlinksW";
            this.rdoFixCommentHyperlinksW.Size = new System.Drawing.Size(156, 19);
            this.rdoFixCommentHyperlinksW.TabIndex = 7;
            this.rdoFixCommentHyperlinksW.TabStop = true;
            this.rdoFixCommentHyperlinksW.Text = "Fix Comment Hyperlinks";
            this.rdoFixCommentHyperlinksW.UseVisualStyleBackColor = true;
            // 
            // rdoFixHyperlinksW
            // 
            this.rdoFixHyperlinksW.AutoSize = true;
            this.rdoFixHyperlinksW.Enabled = false;
            this.rdoFixHyperlinksW.Location = new System.Drawing.Point(6, 172);
            this.rdoFixHyperlinksW.Name = "rdoFixHyperlinksW";
            this.rdoFixHyperlinksW.Size = new System.Drawing.Size(99, 19);
            this.rdoFixHyperlinksW.TabIndex = 6;
            this.rdoFixHyperlinksW.TabStop = true;
            this.rdoFixHyperlinksW.Text = "Fix Hyperlinks";
            this.rdoFixHyperlinksW.UseVisualStyleBackColor = true;
            // 
            // rdoFixCommentsW
            // 
            this.rdoFixCommentsW.AutoSize = true;
            this.rdoFixCommentsW.Enabled = false;
            this.rdoFixCommentsW.Location = new System.Drawing.Point(6, 147);
            this.rdoFixCommentsW.Name = "rdoFixCommentsW";
            this.rdoFixCommentsW.Size = new System.Drawing.Size(102, 19);
            this.rdoFixCommentsW.TabIndex = 5;
            this.rdoFixCommentsW.TabStop = true;
            this.rdoFixCommentsW.Text = "Fix Comments";
            this.rdoFixCommentsW.UseVisualStyleBackColor = true;
            // 
            // rdoFixTablePropsW
            // 
            this.rdoFixTablePropsW.AutoSize = true;
            this.rdoFixTablePropsW.Enabled = false;
            this.rdoFixTablePropsW.Location = new System.Drawing.Point(6, 122);
            this.rdoFixTablePropsW.Name = "rdoFixTablePropsW";
            this.rdoFixTablePropsW.Size = new System.Drawing.Size(126, 19);
            this.rdoFixTablePropsW.TabIndex = 4;
            this.rdoFixTablePropsW.TabStop = true;
            this.rdoFixTablePropsW.Text = "Fix Table Properties";
            this.rdoFixTablePropsW.UseVisualStyleBackColor = true;
            // 
            // rdoFixListTemplatesW
            // 
            this.rdoFixListTemplatesW.AutoSize = true;
            this.rdoFixListTemplatesW.Enabled = false;
            this.rdoFixListTemplatesW.Location = new System.Drawing.Point(6, 97);
            this.rdoFixListTemplatesW.Name = "rdoFixListTemplatesW";
            this.rdoFixListTemplatesW.Size = new System.Drawing.Size(117, 19);
            this.rdoFixListTemplatesW.TabIndex = 3;
            this.rdoFixListTemplatesW.TabStop = true;
            this.rdoFixListTemplatesW.Text = "Fix List Templates";
            this.rdoFixListTemplatesW.UseVisualStyleBackColor = true;
            // 
            // rdoFixEndnotesW
            // 
            this.rdoFixEndnotesW.AutoSize = true;
            this.rdoFixEndnotesW.Enabled = false;
            this.rdoFixEndnotesW.Location = new System.Drawing.Point(6, 72);
            this.rdoFixEndnotesW.Name = "rdoFixEndnotesW";
            this.rdoFixEndnotesW.Size = new System.Drawing.Size(92, 19);
            this.rdoFixEndnotesW.TabIndex = 2;
            this.rdoFixEndnotesW.TabStop = true;
            this.rdoFixEndnotesW.Text = "Fix Endnotes";
            this.rdoFixEndnotesW.UseVisualStyleBackColor = true;
            // 
            // rdoFixRevisionsW
            // 
            this.rdoFixRevisionsW.AutoSize = true;
            this.rdoFixRevisionsW.Enabled = false;
            this.rdoFixRevisionsW.Location = new System.Drawing.Point(6, 47);
            this.rdoFixRevisionsW.Name = "rdoFixRevisionsW";
            this.rdoFixRevisionsW.Size = new System.Drawing.Size(92, 19);
            this.rdoFixRevisionsW.TabIndex = 1;
            this.rdoFixRevisionsW.TabStop = true;
            this.rdoFixRevisionsW.Text = "Fix Revisions";
            this.rdoFixRevisionsW.UseVisualStyleBackColor = true;
            // 
            // rdoFixBookmarksW
            // 
            this.rdoFixBookmarksW.AutoSize = true;
            this.rdoFixBookmarksW.Enabled = false;
            this.rdoFixBookmarksW.Location = new System.Drawing.Point(6, 22);
            this.rdoFixBookmarksW.Name = "rdoFixBookmarksW";
            this.rdoFixBookmarksW.Size = new System.Drawing.Size(102, 19);
            this.rdoFixBookmarksW.TabIndex = 0;
            this.rdoFixBookmarksW.TabStop = true;
            this.rdoFixBookmarksW.Text = "Fix Bookmarks";
            this.rdoFixBookmarksW.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.rdoFixNotesPageSizeCustomP);
            this.groupBox2.Controls.Add(this.rdoFixNotesPageSizeP);
            this.groupBox2.Location = new System.Drawing.Point(218, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(238, 230);
            this.groupBox2.TabIndex = 0;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "PowerPoint Document";
            // 
            // rdoFixNotesPageSizeCustomP
            // 
            this.rdoFixNotesPageSizeCustomP.AutoSize = true;
            this.rdoFixNotesPageSizeCustomP.Enabled = false;
            this.rdoFixNotesPageSizeCustomP.Location = new System.Drawing.Point(12, 47);
            this.rdoFixNotesPageSizeCustomP.Name = "rdoFixNotesPageSizeCustomP";
            this.rdoFixNotesPageSizeCustomP.Size = new System.Drawing.Size(179, 19);
            this.rdoFixNotesPageSizeCustomP.TabIndex = 1;
            this.rdoFixNotesPageSizeCustomP.TabStop = true;
            this.rdoFixNotesPageSizeCustomP.Text = "Fix Notes Page Size (Custom)";
            this.rdoFixNotesPageSizeCustomP.UseVisualStyleBackColor = true;
            // 
            // rdoFixNotesPageSizeP
            // 
            this.rdoFixNotesPageSizeP.AutoSize = true;
            this.rdoFixNotesPageSizeP.Enabled = false;
            this.rdoFixNotesPageSizeP.Location = new System.Drawing.Point(12, 22);
            this.rdoFixNotesPageSizeP.Name = "rdoFixNotesPageSizeP";
            this.rdoFixNotesPageSizeP.Size = new System.Drawing.Size(126, 19);
            this.rdoFixNotesPageSizeP.TabIndex = 0;
            this.rdoFixNotesPageSizeP.TabStop = true;
            this.rdoFixNotesPageSizeP.Text = "Fix Notes Page Size";
            this.rdoFixNotesPageSizeP.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.rdoFixStrictX);
            this.groupBox3.Location = new System.Drawing.Point(462, 12);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(221, 230);
            this.groupBox3.TabIndex = 0;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Excel Document";
            // 
            // rdoFixStrictX
            // 
            this.rdoFixStrictX.AutoSize = true;
            this.rdoFixStrictX.Enabled = false;
            this.rdoFixStrictX.Location = new System.Drawing.Point(6, 22);
            this.rdoFixStrictX.Name = "rdoFixStrictX";
            this.rdoFixStrictX.Size = new System.Drawing.Size(139, 19);
            this.rdoFixStrictX.TabIndex = 0;
            this.rdoFixStrictX.TabStop = true;
            this.rdoFixStrictX.Text = "Remove Strict Format";
            this.rdoFixStrictX.UseVisualStyleBackColor = true;
            // 
            // BtnOk
            // 
            this.BtnOk.Location = new System.Drawing.Point(527, 248);
            this.BtnOk.Name = "BtnOk";
            this.BtnOk.Size = new System.Drawing.Size(75, 23);
            this.BtnOk.TabIndex = 1;
            this.BtnOk.Text = "Ok";
            this.BtnOk.UseVisualStyleBackColor = true;
            this.BtnOk.Click += new System.EventHandler(this.BtnOk_Click);
            // 
            // BtnCancel
            // 
            this.BtnCancel.Location = new System.Drawing.Point(608, 248);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(75, 23);
            this.BtnCancel.TabIndex = 2;
            this.BtnCancel.Text = "Cancel";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // FrmFixDocument
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(695, 282);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.BtnOk);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmFixDocument";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Fix Corrupt Document";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rdoFixCommentHyperlinksW;
        private System.Windows.Forms.RadioButton rdoFixHyperlinksW;
        private System.Windows.Forms.RadioButton rdoFixCommentsW;
        private System.Windows.Forms.RadioButton rdoFixTablePropsW;
        private System.Windows.Forms.RadioButton rdoFixListTemplatesW;
        private System.Windows.Forms.RadioButton rdoFixEndnotesW;
        private System.Windows.Forms.RadioButton rdoFixRevisionsW;
        private System.Windows.Forms.RadioButton rdoFixBookmarksW;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton rdoFixNotesPageSizeCustomP;
        private System.Windows.Forms.RadioButton rdoFixNotesPageSizeP;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.RadioButton rdoFixStrictX;
        private System.Windows.Forms.Button BtnOk;
        private System.Windows.Forms.Button BtnCancel;
    }
}