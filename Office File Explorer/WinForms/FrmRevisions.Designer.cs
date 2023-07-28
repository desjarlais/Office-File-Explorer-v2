
namespace Office_File_Explorer.WinForms
{
    partial class FrmRevisions
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmRevisions));
            BtnOk = new System.Windows.Forms.Button();
            BtnCancel = new System.Windows.Forms.Button();
            groupBox2 = new System.Windows.Forms.GroupBox();
            BtnAcceptChanges = new System.Windows.Forms.Button();
            lbRevisions = new System.Windows.Forms.ListBox();
            label1 = new System.Windows.Forms.Label();
            cbAuthors = new System.Windows.Forms.ComboBox();
            ckbRevisions = new System.Windows.Forms.CheckBox();
            groupBox2.SuspendLayout();
            SuspendLayout();
            // 
            // BtnOk
            // 
            BtnOk.Location = new System.Drawing.Point(356, 506);
            BtnOk.Name = "BtnOk";
            BtnOk.Size = new System.Drawing.Size(75, 23);
            BtnOk.TabIndex = 0;
            BtnOk.Text = "OK";
            BtnOk.UseVisualStyleBackColor = true;
            BtnOk.Click += BtnOk_Click;
            // 
            // BtnCancel
            // 
            BtnCancel.Location = new System.Drawing.Point(437, 506);
            BtnCancel.Name = "BtnCancel";
            BtnCancel.Size = new System.Drawing.Size(75, 23);
            BtnCancel.TabIndex = 1;
            BtnCancel.Text = "Cancel";
            BtnCancel.UseVisualStyleBackColor = true;
            BtnCancel.Click += BtnCancel_Click;
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(BtnAcceptChanges);
            groupBox2.Controls.Add(lbRevisions);
            groupBox2.Controls.Add(label1);
            groupBox2.Controls.Add(cbAuthors);
            groupBox2.Controls.Add(ckbRevisions);
            groupBox2.Location = new System.Drawing.Point(12, 12);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new System.Drawing.Size(500, 488);
            groupBox2.TabIndex = 11;
            groupBox2.TabStop = false;
            groupBox2.Text = "Revisions";
            // 
            // BtnAcceptChanges
            // 
            BtnAcceptChanges.Enabled = false;
            BtnAcceptChanges.Location = new System.Drawing.Point(6, 459);
            BtnAcceptChanges.Name = "BtnAcceptChanges";
            BtnAcceptChanges.Size = new System.Drawing.Size(115, 23);
            BtnAcceptChanges.TabIndex = 15;
            BtnAcceptChanges.Text = "Accept Changes";
            BtnAcceptChanges.UseVisualStyleBackColor = true;
            BtnAcceptChanges.Click += BtnAcceptChanges_Click;
            // 
            // lbRevisions
            // 
            lbRevisions.Enabled = false;
            lbRevisions.FormattingEnabled = true;
            lbRevisions.HorizontalScrollbar = true;
            lbRevisions.ItemHeight = 15;
            lbRevisions.Location = new System.Drawing.Point(6, 76);
            lbRevisions.Name = "lbRevisions";
            lbRevisions.Size = new System.Drawing.Size(486, 379);
            lbRevisions.TabIndex = 14;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point(6, 51);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(50, 15);
            label1.TabIndex = 13;
            label1.Text = "Author: ";
            // 
            // cbAuthors
            // 
            cbAuthors.Enabled = false;
            cbAuthors.FormattingEnabled = true;
            cbAuthors.Location = new System.Drawing.Point(62, 47);
            cbAuthors.Name = "cbAuthors";
            cbAuthors.Size = new System.Drawing.Size(430, 23);
            cbAuthors.TabIndex = 12;
            cbAuthors.SelectedIndexChanged += CbAuthors_SelectedIndexChanged;
            // 
            // ckbRevisions
            // 
            ckbRevisions.AutoSize = true;
            ckbRevisions.Location = new System.Drawing.Point(6, 22);
            ckbRevisions.Name = "ckbRevisions";
            ckbRevisions.Size = new System.Drawing.Size(115, 19);
            ckbRevisions.TabIndex = 0;
            ckbRevisions.Text = "Tracked Changes";
            ckbRevisions.UseVisualStyleBackColor = true;
            ckbRevisions.CheckedChanged += CkbRevisions_CheckedChanged;
            // 
            // FrmWordCommands
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(529, 541);
            Controls.Add(groupBox2);
            Controls.Add(BtnOk);
            Controls.Add(BtnCancel);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            Icon = (System.Drawing.Icon)resources.GetObject("$this.Icon");
            KeyPreview = true;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "FrmWordCommands";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            Text = "Word Commands ";
            KeyDown += FrmWordCommands_KeyDown;
            groupBox2.ResumeLayout(false);
            groupBox2.PerformLayout();
            ResumeLayout(false);
        }

        #endregion
        private System.Windows.Forms.Button BtnOk;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ListBox lbRevisions;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbAuthors;
        private System.Windows.Forms.CheckBox ckbRevisions;
        private System.Windows.Forms.Button BtnAcceptChanges;
    }
}