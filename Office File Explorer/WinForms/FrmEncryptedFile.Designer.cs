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
            this.BtnValidateXml = new System.Windows.Forms.Button();
            this.BtnFixXml = new System.Windows.Forms.Button();
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
            this.tvEncryptedContents.KeyUp += new System.Windows.Forms.KeyEventHandler(this.tvEncryptedContents_KeyUp);
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
            this.BtnSave.Location = new System.Drawing.Point(507, 493);
            this.BtnSave.Name = "BtnSave";
            this.BtnSave.Size = new System.Drawing.Size(103, 23);
            this.BtnSave.TabIndex = 2;
            this.BtnSave.Text = "Save Changes";
            this.BtnSave.UseVisualStyleBackColor = true;
            this.BtnSave.Click += new System.EventHandler(this.BtnSave_Click);
            // 
            // BtnValidateXml
            // 
            this.BtnValidateXml.Location = new System.Drawing.Point(393, 493);
            this.BtnValidateXml.Name = "BtnValidateXml";
            this.BtnValidateXml.Size = new System.Drawing.Size(108, 23);
            this.BtnValidateXml.TabIndex = 3;
            this.BtnValidateXml.Text = "Validate Xml";
            this.BtnValidateXml.UseVisualStyleBackColor = true;
            this.BtnValidateXml.Click += new System.EventHandler(this.BtnValidateXml_Click);
            // 
            // BtnFixXml
            // 
            this.BtnFixXml.Location = new System.Drawing.Point(12, 493);
            this.BtnFixXml.Name = "BtnFixXml";
            this.BtnFixXml.Size = new System.Drawing.Size(75, 23);
            this.BtnFixXml.TabIndex = 4;
            this.BtnFixXml.Text = "Fix Xml";
            this.BtnFixXml.UseVisualStyleBackColor = true;
            this.BtnFixXml.Click += new System.EventHandler(this.BtnFixXml_Click);
            // 
            // FrmEncryptedFile
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(622, 524);
            this.Controls.Add(this.BtnFixXml);
            this.Controls.Add(this.BtnValidateXml);
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
        private System.Windows.Forms.Button BtnValidateXml;
        private System.Windows.Forms.Button BtnFixXml;
    }
}