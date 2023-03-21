
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
            groupBox1 = new System.Windows.Forms.GroupBox();
            rdoFixTableCellTags = new System.Windows.Forms.RadioButton();
            rdoFixListStyles = new System.Windows.Forms.RadioButton();
            rdoFixDataDescriptorW = new System.Windows.Forms.RadioButton();
            rdoFixMathAccentsW = new System.Windows.Forms.RadioButton();
            rdoTryAllFixesW = new System.Windows.Forms.RadioButton();
            rdoFixContentControlsW = new System.Windows.Forms.RadioButton();
            rdoFixCommentHyperlinksW = new System.Windows.Forms.RadioButton();
            rdoFixHyperlinksW = new System.Windows.Forms.RadioButton();
            rdoFixCommentsW = new System.Windows.Forms.RadioButton();
            rdoFixTablePropsW = new System.Windows.Forms.RadioButton();
            rdoFixListTemplatesW = new System.Windows.Forms.RadioButton();
            rdoFixEndnotesW = new System.Windows.Forms.RadioButton();
            rdoFixRevisionsW = new System.Windows.Forms.RadioButton();
            rdoFixBookmarksW = new System.Windows.Forms.RadioButton();
            groupBox2 = new System.Windows.Forms.GroupBox();
            rdoFixNotesPageSizeCustomP = new System.Windows.Forms.RadioButton();
            rdoFixNotesPageSizeP = new System.Windows.Forms.RadioButton();
            groupBox3 = new System.Windows.Forms.GroupBox();
            rdoFixStrictX = new System.Windows.Forms.RadioButton();
            BtnOk = new System.Windows.Forms.Button();
            BtnCancel = new System.Windows.Forms.Button();
            groupBox1.SuspendLayout();
            groupBox2.SuspendLayout();
            groupBox3.SuspendLayout();
            SuspendLayout();
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(rdoFixTableCellTags);
            groupBox1.Controls.Add(rdoFixListStyles);
            groupBox1.Controls.Add(rdoFixDataDescriptorW);
            groupBox1.Controls.Add(rdoFixMathAccentsW);
            groupBox1.Controls.Add(rdoTryAllFixesW);
            groupBox1.Controls.Add(rdoFixContentControlsW);
            groupBox1.Controls.Add(rdoFixCommentHyperlinksW);
            groupBox1.Controls.Add(rdoFixHyperlinksW);
            groupBox1.Controls.Add(rdoFixCommentsW);
            groupBox1.Controls.Add(rdoFixTablePropsW);
            groupBox1.Controls.Add(rdoFixListTemplatesW);
            groupBox1.Controls.Add(rdoFixEndnotesW);
            groupBox1.Controls.Add(rdoFixRevisionsW);
            groupBox1.Controls.Add(rdoFixBookmarksW);
            groupBox1.Location = new System.Drawing.Point(12, 12);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new System.Drawing.Size(547, 195);
            groupBox1.TabIndex = 0;
            groupBox1.TabStop = false;
            groupBox1.Text = "Word Document";
            // 
            // rdoFixTableCellTags
            // 
            rdoFixTableCellTags.AutoSize = true;
            rdoFixTableCellTags.Location = new System.Drawing.Point(193, 147);
            rdoFixTableCellTags.Name = "rdoFixTableCellTags";
            rdoFixTableCellTags.Size = new System.Drawing.Size(142, 19);
            rdoFixTableCellTags.TabIndex = 3;
            rdoFixTableCellTags.TabStop = true;
            rdoFixTableCellTags.Text = "Fix Corrupt Table Cells";
            rdoFixTableCellTags.UseVisualStyleBackColor = true;
            // 
            // rdoFixListStyles
            // 
            rdoFixListStyles.AutoSize = true;
            rdoFixListStyles.Location = new System.Drawing.Point(6, 147);
            rdoFixListStyles.Name = "rdoFixListStyles";
            rdoFixListStyles.Size = new System.Drawing.Size(94, 19);
            rdoFixListStyles.TabIndex = 3;
            rdoFixListStyles.TabStop = true;
            rdoFixListStyles.Text = "Fix List Styles";
            rdoFixListStyles.UseVisualStyleBackColor = true;
            // 
            // rdoFixDataDescriptorW
            // 
            rdoFixDataDescriptorW.AutoSize = true;
            rdoFixDataDescriptorW.Enabled = false;
            rdoFixDataDescriptorW.Location = new System.Drawing.Point(6, 170);
            rdoFixDataDescriptorW.Name = "rdoFixDataDescriptorW";
            rdoFixDataDescriptorW.Size = new System.Drawing.Size(131, 19);
            rdoFixDataDescriptorW.TabIndex = 3;
            rdoFixDataDescriptorW.TabStop = true;
            rdoFixDataDescriptorW.Text = "Fix Corrupt Zip Item";
            rdoFixDataDescriptorW.UseVisualStyleBackColor = true;
            // 
            // rdoFixMathAccentsW
            // 
            rdoFixMathAccentsW.AutoSize = true;
            rdoFixMathAccentsW.Enabled = false;
            rdoFixMathAccentsW.Location = new System.Drawing.Point(193, 72);
            rdoFixMathAccentsW.Name = "rdoFixMathAccentsW";
            rdoFixMathAccentsW.Size = new System.Drawing.Size(170, 19);
            rdoFixMathAccentsW.TabIndex = 3;
            rdoFixMathAccentsW.TabStop = true;
            rdoFixMathAccentsW.Text = "Fix Corrupt Math Equations";
            rdoFixMathAccentsW.UseVisualStyleBackColor = true;
            // 
            // rdoTryAllFixesW
            // 
            rdoTryAllFixesW.AutoSize = true;
            rdoTryAllFixesW.Enabled = false;
            rdoTryAllFixesW.Location = new System.Drawing.Point(396, 22);
            rdoTryAllFixesW.Name = "rdoTryAllFixesW";
            rdoTryAllFixesW.Size = new System.Drawing.Size(86, 19);
            rdoTryAllFixesW.TabIndex = 9;
            rdoTryAllFixesW.TabStop = true;
            rdoTryAllFixesW.Text = "Try All Fixes";
            rdoTryAllFixesW.UseVisualStyleBackColor = true;
            // 
            // rdoFixContentControlsW
            // 
            rdoFixContentControlsW.AutoSize = true;
            rdoFixContentControlsW.Enabled = false;
            rdoFixContentControlsW.Location = new System.Drawing.Point(193, 47);
            rdoFixContentControlsW.Name = "rdoFixContentControlsW";
            rdoFixContentControlsW.Size = new System.Drawing.Size(134, 19);
            rdoFixContentControlsW.TabIndex = 8;
            rdoFixContentControlsW.TabStop = true;
            rdoFixContentControlsW.Text = "Fix Content Controls";
            rdoFixContentControlsW.UseVisualStyleBackColor = true;
            // 
            // rdoFixCommentHyperlinksW
            // 
            rdoFixCommentHyperlinksW.AutoSize = true;
            rdoFixCommentHyperlinksW.Enabled = false;
            rdoFixCommentHyperlinksW.Location = new System.Drawing.Point(193, 22);
            rdoFixCommentHyperlinksW.Name = "rdoFixCommentHyperlinksW";
            rdoFixCommentHyperlinksW.Size = new System.Drawing.Size(156, 19);
            rdoFixCommentHyperlinksW.TabIndex = 7;
            rdoFixCommentHyperlinksW.TabStop = true;
            rdoFixCommentHyperlinksW.Text = "Fix Comment Hyperlinks";
            rdoFixCommentHyperlinksW.UseVisualStyleBackColor = true;
            // 
            // rdoFixHyperlinksW
            // 
            rdoFixHyperlinksW.AutoSize = true;
            rdoFixHyperlinksW.Enabled = false;
            rdoFixHyperlinksW.Location = new System.Drawing.Point(193, 122);
            rdoFixHyperlinksW.Name = "rdoFixHyperlinksW";
            rdoFixHyperlinksW.Size = new System.Drawing.Size(99, 19);
            rdoFixHyperlinksW.TabIndex = 6;
            rdoFixHyperlinksW.TabStop = true;
            rdoFixHyperlinksW.Text = "Fix Hyperlinks";
            rdoFixHyperlinksW.UseVisualStyleBackColor = true;
            // 
            // rdoFixCommentsW
            // 
            rdoFixCommentsW.AutoSize = true;
            rdoFixCommentsW.Enabled = false;
            rdoFixCommentsW.Location = new System.Drawing.Point(6, 122);
            rdoFixCommentsW.Name = "rdoFixCommentsW";
            rdoFixCommentsW.Size = new System.Drawing.Size(102, 19);
            rdoFixCommentsW.TabIndex = 5;
            rdoFixCommentsW.TabStop = true;
            rdoFixCommentsW.Text = "Fix Comments";
            rdoFixCommentsW.UseVisualStyleBackColor = true;
            // 
            // rdoFixTablePropsW
            // 
            rdoFixTablePropsW.AutoSize = true;
            rdoFixTablePropsW.Enabled = false;
            rdoFixTablePropsW.Location = new System.Drawing.Point(193, 97);
            rdoFixTablePropsW.Name = "rdoFixTablePropsW";
            rdoFixTablePropsW.Size = new System.Drawing.Size(119, 19);
            rdoFixTablePropsW.TabIndex = 4;
            rdoFixTablePropsW.TabStop = true;
            rdoFixTablePropsW.Text = "Fix Corrupt Tables";
            rdoFixTablePropsW.UseVisualStyleBackColor = true;
            // 
            // rdoFixListTemplatesW
            // 
            rdoFixListTemplatesW.AutoSize = true;
            rdoFixListTemplatesW.Enabled = false;
            rdoFixListTemplatesW.Location = new System.Drawing.Point(6, 97);
            rdoFixListTemplatesW.Name = "rdoFixListTemplatesW";
            rdoFixListTemplatesW.Size = new System.Drawing.Size(117, 19);
            rdoFixListTemplatesW.TabIndex = 3;
            rdoFixListTemplatesW.TabStop = true;
            rdoFixListTemplatesW.Text = "Fix List Templates";
            rdoFixListTemplatesW.UseVisualStyleBackColor = true;
            // 
            // rdoFixEndnotesW
            // 
            rdoFixEndnotesW.AutoSize = true;
            rdoFixEndnotesW.Enabled = false;
            rdoFixEndnotesW.Location = new System.Drawing.Point(6, 72);
            rdoFixEndnotesW.Name = "rdoFixEndnotesW";
            rdoFixEndnotesW.Size = new System.Drawing.Size(92, 19);
            rdoFixEndnotesW.TabIndex = 2;
            rdoFixEndnotesW.TabStop = true;
            rdoFixEndnotesW.Text = "Fix Endnotes";
            rdoFixEndnotesW.UseVisualStyleBackColor = true;
            // 
            // rdoFixRevisionsW
            // 
            rdoFixRevisionsW.AutoSize = true;
            rdoFixRevisionsW.Enabled = false;
            rdoFixRevisionsW.Location = new System.Drawing.Point(6, 47);
            rdoFixRevisionsW.Name = "rdoFixRevisionsW";
            rdoFixRevisionsW.Size = new System.Drawing.Size(92, 19);
            rdoFixRevisionsW.TabIndex = 1;
            rdoFixRevisionsW.TabStop = true;
            rdoFixRevisionsW.Text = "Fix Revisions";
            rdoFixRevisionsW.UseVisualStyleBackColor = true;
            // 
            // rdoFixBookmarksW
            // 
            rdoFixBookmarksW.AutoSize = true;
            rdoFixBookmarksW.Enabled = false;
            rdoFixBookmarksW.Location = new System.Drawing.Point(6, 22);
            rdoFixBookmarksW.Name = "rdoFixBookmarksW";
            rdoFixBookmarksW.Size = new System.Drawing.Size(102, 19);
            rdoFixBookmarksW.TabIndex = 0;
            rdoFixBookmarksW.TabStop = true;
            rdoFixBookmarksW.Text = "Fix Bookmarks";
            rdoFixBookmarksW.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(rdoFixNotesPageSizeCustomP);
            groupBox2.Controls.Add(rdoFixNotesPageSizeP);
            groupBox2.Location = new System.Drawing.Point(565, 97);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new System.Drawing.Size(201, 81);
            groupBox2.TabIndex = 0;
            groupBox2.TabStop = false;
            groupBox2.Text = "PowerPoint Document";
            // 
            // rdoFixNotesPageSizeCustomP
            // 
            rdoFixNotesPageSizeCustomP.AutoSize = true;
            rdoFixNotesPageSizeCustomP.Enabled = false;
            rdoFixNotesPageSizeCustomP.Location = new System.Drawing.Point(12, 47);
            rdoFixNotesPageSizeCustomP.Name = "rdoFixNotesPageSizeCustomP";
            rdoFixNotesPageSizeCustomP.Size = new System.Drawing.Size(179, 19);
            rdoFixNotesPageSizeCustomP.TabIndex = 1;
            rdoFixNotesPageSizeCustomP.TabStop = true;
            rdoFixNotesPageSizeCustomP.Text = "Fix Notes Page Size (Custom)";
            rdoFixNotesPageSizeCustomP.UseVisualStyleBackColor = true;
            // 
            // rdoFixNotesPageSizeP
            // 
            rdoFixNotesPageSizeP.AutoSize = true;
            rdoFixNotesPageSizeP.Enabled = false;
            rdoFixNotesPageSizeP.Location = new System.Drawing.Point(12, 22);
            rdoFixNotesPageSizeP.Name = "rdoFixNotesPageSizeP";
            rdoFixNotesPageSizeP.Size = new System.Drawing.Size(126, 19);
            rdoFixNotesPageSizeP.TabIndex = 0;
            rdoFixNotesPageSizeP.TabStop = true;
            rdoFixNotesPageSizeP.Text = "Fix Notes Page Size";
            rdoFixNotesPageSizeP.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            groupBox3.Controls.Add(rdoFixStrictX);
            groupBox3.Location = new System.Drawing.Point(565, 12);
            groupBox3.Name = "groupBox3";
            groupBox3.Size = new System.Drawing.Size(201, 79);
            groupBox3.TabIndex = 0;
            groupBox3.TabStop = false;
            groupBox3.Text = "Excel Document";
            // 
            // rdoFixStrictX
            // 
            rdoFixStrictX.AutoSize = true;
            rdoFixStrictX.Enabled = false;
            rdoFixStrictX.Location = new System.Drawing.Point(6, 22);
            rdoFixStrictX.Name = "rdoFixStrictX";
            rdoFixStrictX.Size = new System.Drawing.Size(139, 19);
            rdoFixStrictX.TabIndex = 0;
            rdoFixStrictX.TabStop = true;
            rdoFixStrictX.Text = "Remove Strict Format";
            rdoFixStrictX.UseVisualStyleBackColor = true;
            // 
            // BtnOk
            // 
            BtnOk.Location = new System.Drawing.Point(593, 184);
            BtnOk.Name = "BtnOk";
            BtnOk.Size = new System.Drawing.Size(75, 23);
            BtnOk.TabIndex = 1;
            BtnOk.Text = "OK";
            BtnOk.UseVisualStyleBackColor = true;
            BtnOk.Click += BtnOk_Click;
            // 
            // BtnCancel
            // 
            BtnCancel.Location = new System.Drawing.Point(674, 184);
            BtnCancel.Name = "BtnCancel";
            BtnCancel.Size = new System.Drawing.Size(90, 23);
            BtnCancel.TabIndex = 2;
            BtnCancel.Text = "Cancel";
            BtnCancel.UseVisualStyleBackColor = true;
            BtnCancel.Click += BtnCancel_Click;
            // 
            // FrmFixDocument
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(777, 214);
            Controls.Add(BtnCancel);
            Controls.Add(BtnOk);
            Controls.Add(groupBox3);
            Controls.Add(groupBox2);
            Controls.Add(groupBox1);
            Icon = (System.Drawing.Icon)resources.GetObject("$this.Icon");
            KeyPreview = true;
            Name = "FrmFixDocument";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            Text = "Fix Corrupt Document";
            KeyDown += FrmFixDocument_KeyDown;
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            groupBox2.ResumeLayout(false);
            groupBox2.PerformLayout();
            groupBox3.ResumeLayout(false);
            groupBox3.PerformLayout();
            ResumeLayout(false);
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
        private System.Windows.Forms.RadioButton rdoFixContentControlsW;
        private System.Windows.Forms.RadioButton rdoFixMathAccentsW;
        private System.Windows.Forms.RadioButton rdoTryAllFixesW;
        private System.Windows.Forms.RadioButton rdoFixDataDescriptorW;
        private System.Windows.Forms.RadioButton rdoFixListStyles;
        private System.Windows.Forms.RadioButton rdoFixTableCellTags;
    }
}