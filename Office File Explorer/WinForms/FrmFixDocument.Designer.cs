
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
            rdoFixTextboxes = new System.Windows.Forms.RadioButton();
            rdoFixListStyles = new System.Windows.Forms.RadioButton();
            rdoFixDataDescriptorW = new System.Windows.Forms.RadioButton();
            rdoFixMathAccentsW = new System.Windows.Forms.RadioButton();
            rdoTryAllFixesW = new System.Windows.Forms.RadioButton();
            rdoFixContentControlsW = new System.Windows.Forms.RadioButton();
            rdoFixHyperlinksW = new System.Windows.Forms.RadioButton();
            rdoFixCommentsW = new System.Windows.Forms.RadioButton();
            rdoFixCorruptTables = new System.Windows.Forms.RadioButton();
            rdoFixListTemplatesW = new System.Windows.Forms.RadioButton();
            rdoFixEndnotesW = new System.Windows.Forms.RadioButton();
            rdoFixRevisionsW = new System.Windows.Forms.RadioButton();
            rdoFixBookmarksW = new System.Windows.Forms.RadioButton();
            groupBox2 = new System.Windows.Forms.GroupBox();
            rdoFixDataTags = new System.Windows.Forms.RadioButton();
            rdoResetBulletMargins = new System.Windows.Forms.RadioButton();
            rdoFixNotesPageSizeCustomP = new System.Windows.Forms.RadioButton();
            rdoFixNotesPageSizeP = new System.Windows.Forms.RadioButton();
            groupBox3 = new System.Windows.Forms.GroupBox();
            rdoFixCorruptDrawingsXL = new System.Windows.Forms.RadioButton();
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
            groupBox1.Controls.Add(rdoFixTextboxes);
            groupBox1.Controls.Add(rdoFixListStyles);
            groupBox1.Controls.Add(rdoFixDataDescriptorW);
            groupBox1.Controls.Add(rdoFixMathAccentsW);
            groupBox1.Controls.Add(rdoTryAllFixesW);
            groupBox1.Controls.Add(rdoFixContentControlsW);
            groupBox1.Controls.Add(rdoFixHyperlinksW);
            groupBox1.Controls.Add(rdoFixCommentsW);
            groupBox1.Controls.Add(rdoFixCorruptTables);
            groupBox1.Controls.Add(rdoFixListTemplatesW);
            groupBox1.Controls.Add(rdoFixEndnotesW);
            groupBox1.Controls.Add(rdoFixRevisionsW);
            groupBox1.Controls.Add(rdoFixBookmarksW);
            groupBox1.Location = new System.Drawing.Point(12, 12);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new System.Drawing.Size(330, 223);
            groupBox1.TabIndex = 0;
            groupBox1.TabStop = false;
            groupBox1.Text = "Word Document";
            // 
            // rdoFixTextboxes
            // 
            rdoFixTextboxes.AutoSize = true;
            rdoFixTextboxes.Location = new System.Drawing.Point(6, 169);
            rdoFixTextboxes.Name = "rdoFixTextboxes";
            rdoFixTextboxes.Size = new System.Drawing.Size(95, 19);
            rdoFixTextboxes.TabIndex = 10;
            rdoFixTextboxes.TabStop = true;
            rdoFixTextboxes.Text = "Fix Textboxes";
            rdoFixTextboxes.UseVisualStyleBackColor = true;
            // 
            // rdoFixListStyles
            // 
            rdoFixListStyles.AutoSize = true;
            rdoFixListStyles.Enabled = false;
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
            rdoFixDataDescriptorW.Location = new System.Drawing.Point(143, 122);
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
            rdoFixMathAccentsW.Location = new System.Drawing.Point(143, 47);
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
            rdoTryAllFixesW.Location = new System.Drawing.Point(143, 147);
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
            rdoFixContentControlsW.Location = new System.Drawing.Point(143, 22);
            rdoFixContentControlsW.Name = "rdoFixContentControlsW";
            rdoFixContentControlsW.Size = new System.Drawing.Size(134, 19);
            rdoFixContentControlsW.TabIndex = 8;
            rdoFixContentControlsW.TabStop = true;
            rdoFixContentControlsW.Text = "Fix Content Controls";
            rdoFixContentControlsW.UseVisualStyleBackColor = true;
            // 
            // rdoFixHyperlinksW
            // 
            rdoFixHyperlinksW.AutoSize = true;
            rdoFixHyperlinksW.Enabled = false;
            rdoFixHyperlinksW.Location = new System.Drawing.Point(143, 97);
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
            // rdoFixCorruptTables
            // 
            rdoFixCorruptTables.AutoSize = true;
            rdoFixCorruptTables.Enabled = false;
            rdoFixCorruptTables.Location = new System.Drawing.Point(143, 72);
            rdoFixCorruptTables.Name = "rdoFixCorruptTables";
            rdoFixCorruptTables.Size = new System.Drawing.Size(119, 19);
            rdoFixCorruptTables.TabIndex = 4;
            rdoFixCorruptTables.TabStop = true;
            rdoFixCorruptTables.Text = "Fix Corrupt Tables";
            rdoFixCorruptTables.UseVisualStyleBackColor = true;
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
            groupBox2.Controls.Add(rdoFixDataTags);
            groupBox2.Controls.Add(rdoResetBulletMargins);
            groupBox2.Controls.Add(rdoFixNotesPageSizeCustomP);
            groupBox2.Controls.Add(rdoFixNotesPageSizeP);
            groupBox2.Location = new System.Drawing.Point(348, 86);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new System.Drawing.Size(201, 120);
            groupBox2.TabIndex = 0;
            groupBox2.TabStop = false;
            groupBox2.Text = "PowerPoint Document";
            // 
            // rdoFixDataTags
            // 
            rdoFixDataTags.AutoSize = true;
            rdoFixDataTags.Enabled = false;
            rdoFixDataTags.Location = new System.Drawing.Point(6, 95);
            rdoFixDataTags.Name = "rdoFixDataTags";
            rdoFixDataTags.Size = new System.Drawing.Size(159, 19);
            rdoFixDataTags.TabIndex = 3;
            rdoFixDataTags.TabStop = true;
            rdoFixDataTags.Text = "Fix Missing custData Tags";
            rdoFixDataTags.UseVisualStyleBackColor = true;
            // 
            // rdoResetBulletMargins
            // 
            rdoResetBulletMargins.AutoSize = true;
            rdoResetBulletMargins.Enabled = false;
            rdoResetBulletMargins.Location = new System.Drawing.Point(6, 72);
            rdoResetBulletMargins.Name = "rdoResetBulletMargins";
            rdoResetBulletMargins.Size = new System.Drawing.Size(132, 19);
            rdoResetBulletMargins.TabIndex = 2;
            rdoResetBulletMargins.TabStop = true;
            rdoResetBulletMargins.Text = "Reset Bullet Margins";
            rdoResetBulletMargins.UseVisualStyleBackColor = true;
            // 
            // rdoFixNotesPageSizeCustomP
            // 
            rdoFixNotesPageSizeCustomP.AutoSize = true;
            rdoFixNotesPageSizeCustomP.Enabled = false;
            rdoFixNotesPageSizeCustomP.Location = new System.Drawing.Point(6, 47);
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
            rdoFixNotesPageSizeP.Location = new System.Drawing.Point(6, 22);
            rdoFixNotesPageSizeP.Name = "rdoFixNotesPageSizeP";
            rdoFixNotesPageSizeP.Size = new System.Drawing.Size(126, 19);
            rdoFixNotesPageSizeP.TabIndex = 0;
            rdoFixNotesPageSizeP.TabStop = true;
            rdoFixNotesPageSizeP.Text = "Fix Notes Page Size";
            rdoFixNotesPageSizeP.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            groupBox3.Controls.Add(rdoFixCorruptDrawingsXL);
            groupBox3.Controls.Add(rdoFixStrictX);
            groupBox3.Location = new System.Drawing.Point(348, 12);
            groupBox3.Name = "groupBox3";
            groupBox3.Size = new System.Drawing.Size(201, 68);
            groupBox3.TabIndex = 0;
            groupBox3.TabStop = false;
            groupBox3.Text = "Excel Document";
            // 
            // rdoFixCorruptDrawingsXL
            // 
            rdoFixCorruptDrawingsXL.AutoSize = true;
            rdoFixCorruptDrawingsXL.Enabled = false;
            rdoFixCorruptDrawingsXL.Location = new System.Drawing.Point(6, 43);
            rdoFixCorruptDrawingsXL.Name = "rdoFixCorruptDrawingsXL";
            rdoFixCorruptDrawingsXL.Size = new System.Drawing.Size(136, 19);
            rdoFixCorruptDrawingsXL.TabIndex = 1;
            rdoFixCorruptDrawingsXL.TabStop = true;
            rdoFixCorruptDrawingsXL.Text = "Fix Corrupt Drawings";
            rdoFixCorruptDrawingsXL.UseVisualStyleBackColor = true;
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
            BtnOk.Location = new System.Drawing.Point(378, 212);
            BtnOk.Name = "BtnOk";
            BtnOk.Size = new System.Drawing.Size(75, 23);
            BtnOk.TabIndex = 1;
            BtnOk.Text = "OK";
            BtnOk.UseVisualStyleBackColor = true;
            BtnOk.Click += BtnOk_Click;
            // 
            // BtnCancel
            // 
            BtnCancel.Location = new System.Drawing.Point(459, 212);
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
            ClientSize = new System.Drawing.Size(561, 247);
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
        private System.Windows.Forms.RadioButton rdoFixHyperlinksW;
        private System.Windows.Forms.RadioButton rdoFixCommentsW;
        private System.Windows.Forms.RadioButton rdoFixCorruptTables;
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
        private System.Windows.Forms.RadioButton rdoResetBulletMargins;
        private System.Windows.Forms.RadioButton rdoFixDataTags;
        private System.Windows.Forms.RadioButton rdoFixCorruptDrawingsXL;
        private System.Windows.Forms.RadioButton rdoFixTextboxes;
    }
}