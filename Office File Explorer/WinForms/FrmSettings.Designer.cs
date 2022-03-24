
namespace Office_File_Explorer.WinForms
{
    partial class FrmSettings
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmSettings));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.ckbFixGroupedShapes = new System.Windows.Forms.CheckBox();
            this.ckbListRsids = new System.Windows.Forms.CheckBox();
            this.ckbRemoveFallbackTags = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.ckbResetNotes = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.ckbBackupOnOpen = new System.Windows.Forms.CheckBox();
            this.ckbZipItemCorrupt = new System.Windows.Forms.CheckBox();
            this.ckbDeleteOnExit = new System.Windows.Forms.CheckBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.rdoSAX = new System.Windows.Forms.RadioButton();
            this.rdoDOM = new System.Windows.Forms.RadioButton();
            this.BtnOk = new System.Windows.Forms.Button();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.tbxSiteURL = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.tbxClientID = new System.Windows.Forms.TextBox();
            this.tbxTenant = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.ckbFixGroupedShapes);
            this.groupBox1.Controls.Add(this.ckbListRsids);
            this.groupBox1.Controls.Add(this.ckbRemoveFallbackTags);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(287, 141);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Word Corrupt Document";
            // 
            // ckbFixGroupedShapes
            // 
            this.ckbFixGroupedShapes.AutoSize = true;
            this.ckbFixGroupedShapes.Location = new System.Drawing.Point(19, 66);
            this.ckbFixGroupedShapes.Name = "ckbFixGroupedShapes";
            this.ckbFixGroupedShapes.Size = new System.Drawing.Size(130, 19);
            this.ckbFixGroupedShapes.TabIndex = 2;
            this.ckbFixGroupedShapes.Text = "Fix Grouped Shapes";
            this.ckbFixGroupedShapes.UseVisualStyleBackColor = true;
            // 
            // ckbListRsids
            // 
            this.ckbListRsids.AutoSize = true;
            this.ckbListRsids.Location = new System.Drawing.Point(19, 44);
            this.ckbListRsids.Name = "ckbListRsids";
            this.ckbListRsids.Size = new System.Drawing.Size(159, 19);
            this.ckbListRsids.TabIndex = 1;
            this.ckbListRsids.Text = "List Rsids With Doc Props";
            this.ckbListRsids.UseVisualStyleBackColor = true;
            // 
            // ckbRemoveFallbackTags
            // 
            this.ckbRemoveFallbackTags.AutoSize = true;
            this.ckbRemoveFallbackTags.Location = new System.Drawing.Point(19, 19);
            this.ckbRemoveFallbackTags.Name = "ckbRemoveFallbackTags";
            this.ckbRemoveFallbackTags.Size = new System.Drawing.Size(158, 19);
            this.ckbRemoveFallbackTags.TabIndex = 0;
            this.ckbRemoveFallbackTags.Text = "Remove All Fallback Tags";
            this.ckbRemoveFallbackTags.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.ckbResetNotes);
            this.groupBox2.Location = new System.Drawing.Point(305, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(238, 141);
            this.groupBox2.TabIndex = 0;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "PowerPoint Options";
            // 
            // ckbResetNotes
            // 
            this.ckbResetNotes.AutoSize = true;
            this.ckbResetNotes.Location = new System.Drawing.Point(8, 19);
            this.ckbResetNotes.Name = "ckbResetNotes";
            this.ckbResetNotes.Size = new System.Drawing.Size(217, 19);
            this.ckbResetNotes.TabIndex = 0;
            this.ckbResetNotes.Text = "Reset Notes Slides and Notes Master";
            this.ckbResetNotes.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.ckbBackupOnOpen);
            this.groupBox3.Controls.Add(this.ckbZipItemCorrupt);
            this.groupBox3.Controls.Add(this.ckbDeleteOnExit);
            this.groupBox3.Location = new System.Drawing.Point(12, 159);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(287, 100);
            this.groupBox3.TabIndex = 0;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "App Settings";
            // 
            // ckbBackupOnOpen
            // 
            this.ckbBackupOnOpen.AutoSize = true;
            this.ckbBackupOnOpen.Location = new System.Drawing.Point(19, 73);
            this.ckbBackupOnOpen.Name = "ckbBackupOnOpen";
            this.ckbBackupOnOpen.Size = new System.Drawing.Size(216, 19);
            this.ckbBackupOnOpen.TabIndex = 3;
            this.ckbBackupOnOpen.Text = "Make Backup Copy Of File On Open";
            this.ckbBackupOnOpen.UseVisualStyleBackColor = true;
            // 
            // ckbZipItemCorrupt
            // 
            this.ckbZipItemCorrupt.AutoSize = true;
            this.ckbZipItemCorrupt.Location = new System.Drawing.Point(19, 48);
            this.ckbZipItemCorrupt.Name = "ckbZipItemCorrupt";
            this.ckbZipItemCorrupt.Size = new System.Drawing.Size(238, 19);
            this.ckbZipItemCorrupt.TabIndex = 3;
            this.ckbZipItemCorrupt.Text = "Check For Zip Item Corruption On Open";
            this.ckbZipItemCorrupt.UseVisualStyleBackColor = true;
            // 
            // ckbDeleteOnExit
            // 
            this.ckbDeleteOnExit.AutoSize = true;
            this.ckbDeleteOnExit.Location = new System.Drawing.Point(19, 22);
            this.ckbDeleteOnExit.Name = "ckbDeleteOnExit";
            this.ckbDeleteOnExit.Size = new System.Drawing.Size(167, 19);
            this.ckbDeleteOnExit.TabIndex = 0;
            this.ckbDeleteOnExit.Text = "Delete Copied Files On Exit";
            this.ckbDeleteOnExit.UseVisualStyleBackColor = true;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.rdoSAX);
            this.groupBox4.Controls.Add(this.rdoDOM);
            this.groupBox4.Location = new System.Drawing.Point(306, 159);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(238, 100);
            this.groupBox4.TabIndex = 0;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Excel Cell Value Option";
            // 
            // rdoSAX
            // 
            this.rdoSAX.AutoSize = true;
            this.rdoSAX.Location = new System.Drawing.Point(6, 22);
            this.rdoSAX.Name = "rdoSAX";
            this.rdoSAX.Size = new System.Drawing.Size(79, 19);
            this.rdoSAX.TabIndex = 3;
            this.rdoSAX.TabStop = true;
            this.rdoSAX.Text = "SAX Styles";
            this.rdoSAX.UseVisualStyleBackColor = true;
            // 
            // rdoDOM
            // 
            this.rdoDOM.AutoSize = true;
            this.rdoDOM.Location = new System.Drawing.Point(6, 47);
            this.rdoDOM.Name = "rdoDOM";
            this.rdoDOM.Size = new System.Drawing.Size(86, 19);
            this.rdoDOM.TabIndex = 4;
            this.rdoDOM.TabStop = true;
            this.rdoDOM.Text = "DOM Styles";
            this.rdoDOM.UseVisualStyleBackColor = true;
            // 
            // BtnOk
            // 
            this.BtnOk.Location = new System.Drawing.Point(387, 387);
            this.BtnOk.Name = "BtnOk";
            this.BtnOk.Size = new System.Drawing.Size(75, 23);
            this.BtnOk.TabIndex = 1;
            this.BtnOk.Text = "OK";
            this.BtnOk.UseVisualStyleBackColor = true;
            this.BtnOk.Click += new System.EventHandler(this.BtnOk_Click);
            // 
            // BtnCancel
            // 
            this.BtnCancel.Location = new System.Drawing.Point(468, 387);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(75, 23);
            this.BtnCancel.TabIndex = 2;
            this.BtnCancel.Text = "Cancel";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.tbxSiteURL);
            this.groupBox5.Controls.Add(this.label3);
            this.groupBox5.Controls.Add(this.tbxClientID);
            this.groupBox5.Controls.Add(this.tbxTenant);
            this.groupBox5.Controls.Add(this.label2);
            this.groupBox5.Controls.Add(this.label1);
            this.groupBox5.Location = new System.Drawing.Point(12, 265);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(531, 116);
            this.groupBox5.TabIndex = 3;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "ADAL Settings";
            // 
            // tbxSiteURL
            // 
            this.tbxSiteURL.Location = new System.Drawing.Point(61, 81);
            this.tbxSiteURL.Name = "tbxSiteURL";
            this.tbxSiteURL.Size = new System.Drawing.Size(464, 23);
            this.tbxSiteURL.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 84);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(50, 15);
            this.label3.TabIndex = 4;
            this.label3.Text = "Site URL";
            // 
            // tbxClientID
            // 
            this.tbxClientID.Location = new System.Drawing.Point(61, 25);
            this.tbxClientID.Name = "tbxClientID";
            this.tbxClientID.Size = new System.Drawing.Size(464, 23);
            this.tbxClientID.TabIndex = 3;
            // 
            // tbxTenant
            // 
            this.tbxTenant.Location = new System.Drawing.Point(61, 52);
            this.tbxTenant.Name = "tbxTenant";
            this.tbxTenant.Size = new System.Drawing.Size(464, 23);
            this.tbxTenant.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 55);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(48, 15);
            this.label2.TabIndex = 1;
            this.label2.Text = "Tenant: ";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(58, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "Client ID: ";
            // 
            // FrmSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(546, 420);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.BtnOk);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmSettings";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Settings";
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmSettings_KeyDown);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox ckbFixGroupedShapes;
        private System.Windows.Forms.CheckBox ckbListRsids;
        private System.Windows.Forms.CheckBox ckbRemoveFallbackTags;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox ckbResetNotes;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.CheckBox ckbDeleteOnExit;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Button BtnOk;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.RadioButton rdoSAX;
        private System.Windows.Forms.RadioButton rdoDOM;
        private System.Windows.Forms.CheckBox ckbZipItemCorrupt;
        private System.Windows.Forms.CheckBox ckbBackupOnOpen;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.TextBox tbxClientID;
        private System.Windows.Forms.TextBox tbxTenant;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbxSiteURL;
        private System.Windows.Forms.Label label3;
    }
}