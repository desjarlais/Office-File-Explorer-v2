namespace Office_File_Explorer.WinForms
{
    partial class FrmEncryptedFile
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmEncryptedFile));
            this.tvEncryptedContents = new System.Windows.Forms.TreeView();
            this.TxbOutput = new System.Windows.Forms.TextBox();
            this.BtnSave = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // tvEncryptedContents
            // 
            this.tvEncryptedContents.Location = new System.Drawing.Point(12, 12);
            this.tvEncryptedContents.Name = "tvEncryptedContents";
            this.tvEncryptedContents.Size = new System.Drawing.Size(600, 224);
            this.tvEncryptedContents.TabIndex = 0;
            this.tvEncryptedContents.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.tvEncryptedContents_NodeMouseClick);
            this.tvEncryptedContents.KeyDown += new System.Windows.Forms.KeyEventHandler(this.tvEncryptedContents_KeyDown);
            // 
            // TxbOutput
            // 
            this.TxbOutput.Location = new System.Drawing.Point(12, 242);
            this.TxbOutput.Multiline = true;
            this.TxbOutput.Name = "TxbOutput";
            this.TxbOutput.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.TxbOutput.Size = new System.Drawing.Size(598, 245);
            this.TxbOutput.TabIndex = 1;
            // 
            // BtnSave
            // 
            this.BtnSave.Location = new System.Drawing.Point(535, 493);
            this.BtnSave.Name = "BtnSave";
            this.BtnSave.Size = new System.Drawing.Size(75, 23);
            this.BtnSave.TabIndex = 2;
            this.BtnSave.Text = "Save";
            this.BtnSave.UseVisualStyleBackColor = true;
            this.BtnSave.Click += new System.EventHandler(this.BtnSave_Click);
            // 
            // FrmEncryptedFile
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(622, 524);
            this.Controls.Add(this.BtnSave);
            this.Controls.Add(this.TxbOutput);
            this.Controls.Add(this.tvEncryptedContents);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmEncryptedFile";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "File Contents";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmEncryptedFile_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TreeView tvEncryptedContents;
        private System.Windows.Forms.TextBox TxbOutput;
        private System.Windows.Forms.Button BtnSave;
    }
}