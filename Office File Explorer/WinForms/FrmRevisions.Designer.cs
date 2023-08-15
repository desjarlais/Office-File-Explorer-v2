
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
            BtnAcceptChanges = new System.Windows.Forms.Button();
            lbRevisions = new System.Windows.Forms.ListBox();
            label1 = new System.Windows.Forms.Label();
            cbAuthors = new System.Windows.Forms.ComboBox();
            SuspendLayout();
            // 
            // BtnOk
            // 
            BtnOk.Location = new System.Drawing.Point(343, 419);
            BtnOk.Name = "BtnOk";
            BtnOk.Size = new System.Drawing.Size(75, 23);
            BtnOk.TabIndex = 0;
            BtnOk.Text = "OK";
            BtnOk.UseVisualStyleBackColor = true;
            BtnOk.Click += BtnOk_Click;
            // 
            // BtnCancel
            // 
            BtnCancel.Location = new System.Drawing.Point(424, 419);
            BtnCancel.Name = "BtnCancel";
            BtnCancel.Size = new System.Drawing.Size(75, 23);
            BtnCancel.TabIndex = 1;
            BtnCancel.Text = "Cancel";
            BtnCancel.UseVisualStyleBackColor = true;
            BtnCancel.Click += BtnCancel_Click;
            // 
            // BtnAcceptChanges
            // 
            BtnAcceptChanges.Location = new System.Drawing.Point(12, 419);
            BtnAcceptChanges.Name = "BtnAcceptChanges";
            BtnAcceptChanges.Size = new System.Drawing.Size(115, 23);
            BtnAcceptChanges.TabIndex = 15;
            BtnAcceptChanges.Text = "Accept Changes";
            BtnAcceptChanges.UseVisualStyleBackColor = true;
            BtnAcceptChanges.Click += BtnAcceptChanges_Click;
            // 
            // lbRevisions
            // 
            lbRevisions.FormattingEnabled = true;
            lbRevisions.HorizontalScrollbar = true;
            lbRevisions.ItemHeight = 15;
            lbRevisions.Location = new System.Drawing.Point(12, 34);
            lbRevisions.Name = "lbRevisions";
            lbRevisions.Size = new System.Drawing.Size(486, 379);
            lbRevisions.TabIndex = 14;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point(12, 9);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(50, 15);
            label1.TabIndex = 13;
            label1.Text = "Author: ";
            // 
            // cbAuthors
            // 
            cbAuthors.FormattingEnabled = true;
            cbAuthors.Location = new System.Drawing.Point(68, 5);
            cbAuthors.Name = "cbAuthors";
            cbAuthors.Size = new System.Drawing.Size(430, 23);
            cbAuthors.TabIndex = 12;
            cbAuthors.SelectedIndexChanged += CbAuthors_SelectedIndexChanged;
            // 
            // FrmRevisions
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(508, 451);
            Controls.Add(BtnAcceptChanges);
            Controls.Add(lbRevisions);
            Controls.Add(BtnOk);
            Controls.Add(label1);
            Controls.Add(BtnCancel);
            Controls.Add(cbAuthors);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            Icon = (System.Drawing.Icon)resources.GetObject("$this.Icon");
            KeyPreview = true;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "FrmRevisions";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            Text = "Document Revisions";
            KeyDown += FrmWordCommands_KeyDown;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private System.Windows.Forms.Button BtnOk;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.ListBox lbRevisions;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbAuthors;
        private System.Windows.Forms.Button BtnAcceptChanges;
    }
}