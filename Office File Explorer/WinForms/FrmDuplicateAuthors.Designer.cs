namespace Office_File_Explorer.WinForms
{
    partial class FrmDuplicateAuthors
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
            LstAuthors = new System.Windows.Forms.ListBox();
            BtnRemoveDupes = new System.Windows.Forms.Button();
            SuspendLayout();
            // 
            // LstAuthors
            // 
            LstAuthors.FormattingEnabled = true;
            LstAuthors.ItemHeight = 15;
            LstAuthors.Location = new System.Drawing.Point(12, 14);
            LstAuthors.Name = "LstAuthors";
            LstAuthors.Size = new System.Drawing.Size(565, 154);
            LstAuthors.TabIndex = 2;
            // 
            // BtnRemoveDupes
            // 
            BtnRemoveDupes.Location = new System.Drawing.Point(445, 174);
            BtnRemoveDupes.Name = "BtnRemoveDupes";
            BtnRemoveDupes.Size = new System.Drawing.Size(132, 23);
            BtnRemoveDupes.TabIndex = 3;
            BtnRemoveDupes.Text = "Remove Duplicates";
            BtnRemoveDupes.UseVisualStyleBackColor = true;
            BtnRemoveDupes.Click += BtnRemoveDupes_Click;
            // 
            // FrmDuplicateAuthors
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(589, 206);
            Controls.Add(BtnRemoveDupes);
            Controls.Add(LstAuthors);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "FrmDuplicateAuthors";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            Text = "Delete Duplicate Authors";
            ResumeLayout(false);
        }

        #endregion
        private System.Windows.Forms.ListBox LstAuthors;
        private System.Windows.Forms.Button BtnRemoveDupes;
    }
}