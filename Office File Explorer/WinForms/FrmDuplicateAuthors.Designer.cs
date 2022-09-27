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
            this.label1 = new System.Windows.Forms.Label();
            this.LstAuthors = new System.Windows.Forms.ListBox();
            this.BtnRemoveDupes = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.LblUserId = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 6);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(87, 15);
            this.label1.TabIndex = 1;
            this.label1.Text = "Author User ID:";
            // 
            // LstAuthors
            // 
            this.LstAuthors.FormattingEnabled = true;
            this.LstAuthors.ItemHeight = 15;
            this.LstAuthors.Location = new System.Drawing.Point(12, 59);
            this.LstAuthors.Name = "LstAuthors";
            this.LstAuthors.Size = new System.Drawing.Size(565, 109);
            this.LstAuthors.TabIndex = 2;
            // 
            // BtnRemoveDupes
            // 
            this.BtnRemoveDupes.Location = new System.Drawing.Point(445, 174);
            this.BtnRemoveDupes.Name = "BtnRemoveDupes";
            this.BtnRemoveDupes.Size = new System.Drawing.Size(132, 23);
            this.BtnRemoveDupes.TabIndex = 3;
            this.BtnRemoveDupes.Text = "Remove Duplicates";
            this.BtnRemoveDupes.UseVisualStyleBackColor = true;
            this.BtnRemoveDupes.Click += new System.EventHandler(this.BtnRemoveDupes_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 41);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(120, 15);
            this.label2.TabIndex = 4;
            this.label2.Text = "Select Correct UserID:";
            // 
            // LblUserId
            // 
            this.LblUserId.AutoSize = true;
            this.LblUserId.Location = new System.Drawing.Point(105, 6);
            this.LblUserId.Name = "LblUserId";
            this.LblUserId.Size = new System.Drawing.Size(58, 15);
            this.LblUserId.TabIndex = 5;
            this.LblUserId.Text = "<user id>";
            // 
            // FrmDuplicateAuthors
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(589, 206);
            this.Controls.Add(this.LblUserId);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.BtnRemoveDupes);
            this.Controls.Add(this.LstAuthors);
            this.Controls.Add(this.label1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmDuplicateAuthors";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Delete Duplicate Authors";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox LstAuthors;
        private System.Windows.Forms.Button BtnRemoveDupes;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label LblUserId;
    }
}