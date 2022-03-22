
namespace Office_File_Explorer.WinForms
{
    partial class FrmWordCommands
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmWordCommands));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.ckbTables = new System.Windows.Forms.CheckBox();
            this.ckbFieldCodes = new System.Windows.Forms.CheckBox();
            this.ckbComments = new System.Windows.Forms.CheckBox();
            this.ckbBookmarks = new System.Windows.Forms.CheckBox();
            this.ckbDocProps = new System.Windows.Forms.CheckBox();
            this.ckbEndnotes = new System.Windows.Forms.CheckBox();
            this.ckbFootnotes = new System.Windows.Forms.CheckBox();
            this.ckbFonts = new System.Windows.Forms.CheckBox();
            this.ckbListTemplates = new System.Windows.Forms.CheckBox();
            this.ckbHyperlinks = new System.Windows.Forms.CheckBox();
            this.ckbContentControls = new System.Windows.Forms.CheckBox();
            this.ckbStyles = new System.Windows.Forms.CheckBox();
            this.BtnOk = new System.Windows.Forms.Button();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.lbRevisions = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cbAuthors = new System.Windows.Forms.ComboBox();
            this.ckbRevisions = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.ckbPackageParts = new System.Windows.Forms.CheckBox();
            this.ckbShapes = new System.Windows.Forms.CheckBox();
            this.ckbOleObjects = new System.Windows.Forms.CheckBox();
            this.ckbSelectAll = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.ckbTables);
            this.groupBox1.Controls.Add(this.ckbFieldCodes);
            this.groupBox1.Controls.Add(this.ckbComments);
            this.groupBox1.Controls.Add(this.ckbBookmarks);
            this.groupBox1.Controls.Add(this.ckbDocProps);
            this.groupBox1.Controls.Add(this.ckbEndnotes);
            this.groupBox1.Controls.Add(this.ckbFootnotes);
            this.groupBox1.Controls.Add(this.ckbFonts);
            this.groupBox1.Controls.Add(this.ckbListTemplates);
            this.groupBox1.Controls.Add(this.ckbHyperlinks);
            this.groupBox1.Controls.Add(this.ckbContentControls);
            this.groupBox1.Controls.Add(this.ckbStyles);
            this.groupBox1.Location = new System.Drawing.Point(9, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(212, 379);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Select Word Content To Display";
            // 
            // ckbTables
            // 
            this.ckbTables.AutoSize = true;
            this.ckbTables.Location = new System.Drawing.Point(11, 297);
            this.ckbTables.Name = "ckbTables";
            this.ckbTables.Size = new System.Drawing.Size(58, 19);
            this.ckbTables.TabIndex = 11;
            this.ckbTables.Text = "Tables";
            this.ckbTables.UseVisualStyleBackColor = true;
            // 
            // ckbFieldCodes
            // 
            this.ckbFieldCodes.AutoSize = true;
            this.ckbFieldCodes.Location = new System.Drawing.Point(11, 272);
            this.ckbFieldCodes.Name = "ckbFieldCodes";
            this.ckbFieldCodes.Size = new System.Drawing.Size(87, 19);
            this.ckbFieldCodes.TabIndex = 10;
            this.ckbFieldCodes.Text = "Field Codes";
            this.ckbFieldCodes.UseVisualStyleBackColor = true;
            // 
            // ckbComments
            // 
            this.ckbComments.AutoSize = true;
            this.ckbComments.Location = new System.Drawing.Point(11, 247);
            this.ckbComments.Name = "ckbComments";
            this.ckbComments.Size = new System.Drawing.Size(85, 19);
            this.ckbComments.TabIndex = 9;
            this.ckbComments.Text = "Comments";
            this.ckbComments.UseVisualStyleBackColor = true;
            // 
            // ckbBookmarks
            // 
            this.ckbBookmarks.AutoSize = true;
            this.ckbBookmarks.Location = new System.Drawing.Point(11, 222);
            this.ckbBookmarks.Name = "ckbBookmarks";
            this.ckbBookmarks.Size = new System.Drawing.Size(85, 19);
            this.ckbBookmarks.TabIndex = 8;
            this.ckbBookmarks.Text = "Bookmarks";
            this.ckbBookmarks.UseVisualStyleBackColor = true;
            // 
            // ckbDocProps
            // 
            this.ckbDocProps.AutoSize = true;
            this.ckbDocProps.Location = new System.Drawing.Point(11, 197);
            this.ckbDocProps.Name = "ckbDocProps";
            this.ckbDocProps.Size = new System.Drawing.Size(138, 19);
            this.ckbDocProps.TabIndex = 7;
            this.ckbDocProps.Text = "Document Properties";
            this.ckbDocProps.UseVisualStyleBackColor = true;
            // 
            // ckbEndnotes
            // 
            this.ckbEndnotes.AutoSize = true;
            this.ckbEndnotes.Location = new System.Drawing.Point(11, 172);
            this.ckbEndnotes.Name = "ckbEndnotes";
            this.ckbEndnotes.Size = new System.Drawing.Size(75, 19);
            this.ckbEndnotes.TabIndex = 6;
            this.ckbEndnotes.Text = "Endnotes";
            this.ckbEndnotes.UseVisualStyleBackColor = true;
            // 
            // ckbFootnotes
            // 
            this.ckbFootnotes.AutoSize = true;
            this.ckbFootnotes.Location = new System.Drawing.Point(11, 147);
            this.ckbFootnotes.Name = "ckbFootnotes";
            this.ckbFootnotes.Size = new System.Drawing.Size(79, 19);
            this.ckbFootnotes.TabIndex = 5;
            this.ckbFootnotes.Text = "Footnotes";
            this.ckbFootnotes.UseVisualStyleBackColor = true;
            // 
            // ckbFonts
            // 
            this.ckbFonts.AutoSize = true;
            this.ckbFonts.Location = new System.Drawing.Point(12, 122);
            this.ckbFonts.Name = "ckbFonts";
            this.ckbFonts.Size = new System.Drawing.Size(55, 19);
            this.ckbFonts.TabIndex = 2;
            this.ckbFonts.Text = "Fonts";
            this.ckbFonts.UseVisualStyleBackColor = true;
            // 
            // ckbListTemplates
            // 
            this.ckbListTemplates.AutoSize = true;
            this.ckbListTemplates.Location = new System.Drawing.Point(12, 97);
            this.ckbListTemplates.Name = "ckbListTemplates";
            this.ckbListTemplates.Size = new System.Drawing.Size(100, 19);
            this.ckbListTemplates.TabIndex = 4;
            this.ckbListTemplates.Text = "List Templates";
            this.ckbListTemplates.UseVisualStyleBackColor = true;
            // 
            // ckbHyperlinks
            // 
            this.ckbHyperlinks.AutoSize = true;
            this.ckbHyperlinks.Location = new System.Drawing.Point(12, 72);
            this.ckbHyperlinks.Name = "ckbHyperlinks";
            this.ckbHyperlinks.Size = new System.Drawing.Size(82, 19);
            this.ckbHyperlinks.TabIndex = 2;
            this.ckbHyperlinks.Text = "Hyperlinks";
            this.ckbHyperlinks.UseVisualStyleBackColor = true;
            // 
            // ckbContentControls
            // 
            this.ckbContentControls.AutoSize = true;
            this.ckbContentControls.Location = new System.Drawing.Point(12, 22);
            this.ckbContentControls.Name = "ckbContentControls";
            this.ckbContentControls.Size = new System.Drawing.Size(117, 19);
            this.ckbContentControls.TabIndex = 2;
            this.ckbContentControls.Text = "Content Controls";
            this.ckbContentControls.UseVisualStyleBackColor = true;
            // 
            // ckbStyles
            // 
            this.ckbStyles.AutoSize = true;
            this.ckbStyles.Location = new System.Drawing.Point(12, 47);
            this.ckbStyles.Name = "ckbStyles";
            this.ckbStyles.Size = new System.Drawing.Size(56, 19);
            this.ckbStyles.TabIndex = 3;
            this.ckbStyles.Text = "Styles";
            this.ckbStyles.UseVisualStyleBackColor = true;
            // 
            // BtnOk
            // 
            this.BtnOk.Location = new System.Drawing.Point(571, 506);
            this.BtnOk.Name = "BtnOk";
            this.BtnOk.Size = new System.Drawing.Size(75, 23);
            this.BtnOk.TabIndex = 0;
            this.BtnOk.Text = "OK";
            this.BtnOk.UseVisualStyleBackColor = true;
            this.BtnOk.Click += new System.EventHandler(this.BtnOk_Click);
            // 
            // BtnCancel
            // 
            this.BtnCancel.Location = new System.Drawing.Point(652, 506);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(75, 23);
            this.BtnCancel.TabIndex = 1;
            this.BtnCancel.Text = "Cancel";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.lbRevisions);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.cbAuthors);
            this.groupBox2.Controls.Add(this.ckbRevisions);
            this.groupBox2.Location = new System.Drawing.Point(227, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(500, 488);
            this.groupBox2.TabIndex = 11;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Revisions";
            // 
            // lbRevisions
            // 
            this.lbRevisions.Enabled = false;
            this.lbRevisions.FormattingEnabled = true;
            this.lbRevisions.HorizontalScrollbar = true;
            this.lbRevisions.ItemHeight = 15;
            this.lbRevisions.Location = new System.Drawing.Point(6, 76);
            this.lbRevisions.Name = "lbRevisions";
            this.lbRevisions.Size = new System.Drawing.Size(486, 394);
            this.lbRevisions.TabIndex = 14;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 51);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(50, 15);
            this.label1.TabIndex = 13;
            this.label1.Text = "Author: ";
            // 
            // cbAuthors
            // 
            this.cbAuthors.Enabled = false;
            this.cbAuthors.FormattingEnabled = true;
            this.cbAuthors.Location = new System.Drawing.Point(62, 47);
            this.cbAuthors.Name = "cbAuthors";
            this.cbAuthors.Size = new System.Drawing.Size(430, 23);
            this.cbAuthors.TabIndex = 12;
            this.cbAuthors.SelectedIndexChanged += new System.EventHandler(this.CbAuthors_SelectedIndexChanged);
            // 
            // ckbRevisions
            // 
            this.ckbRevisions.AutoSize = true;
            this.ckbRevisions.Location = new System.Drawing.Point(6, 22);
            this.ckbRevisions.Name = "ckbRevisions";
            this.ckbRevisions.Size = new System.Drawing.Size(115, 19);
            this.ckbRevisions.TabIndex = 0;
            this.ckbRevisions.Text = "Tracked Changes";
            this.ckbRevisions.UseVisualStyleBackColor = true;
            this.ckbRevisions.CheckedChanged += new System.EventHandler(this.CkbRevisions_CheckedChanged);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.ckbPackageParts);
            this.groupBox3.Controls.Add(this.ckbShapes);
            this.groupBox3.Controls.Add(this.ckbOleObjects);
            this.groupBox3.Location = new System.Drawing.Point(9, 397);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(212, 103);
            this.groupBox3.TabIndex = 12;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Select Office Content To Display";
            // 
            // ckbPackageParts
            // 
            this.ckbPackageParts.AutoSize = true;
            this.ckbPackageParts.Location = new System.Drawing.Point(12, 72);
            this.ckbPackageParts.Name = "ckbPackageParts";
            this.ckbPackageParts.Size = new System.Drawing.Size(99, 19);
            this.ckbPackageParts.TabIndex = 2;
            this.ckbPackageParts.Text = "Package Parts";
            this.ckbPackageParts.UseVisualStyleBackColor = true;
            // 
            // ckbShapes
            // 
            this.ckbShapes.AutoSize = true;
            this.ckbShapes.Location = new System.Drawing.Point(12, 47);
            this.ckbShapes.Name = "ckbShapes";
            this.ckbShapes.Size = new System.Drawing.Size(63, 19);
            this.ckbShapes.TabIndex = 1;
            this.ckbShapes.Text = "Shapes";
            this.ckbShapes.UseVisualStyleBackColor = true;
            // 
            // ckbOleObjects
            // 
            this.ckbOleObjects.AutoSize = true;
            this.ckbOleObjects.Location = new System.Drawing.Point(12, 22);
            this.ckbOleObjects.Name = "ckbOleObjects";
            this.ckbOleObjects.Size = new System.Drawing.Size(87, 19);
            this.ckbOleObjects.TabIndex = 0;
            this.ckbOleObjects.Text = "Ole Objects";
            this.ckbOleObjects.UseVisualStyleBackColor = true;
            // 
            // ckbSelectAll
            // 
            this.ckbSelectAll.AutoSize = true;
            this.ckbSelectAll.Location = new System.Drawing.Point(9, 506);
            this.ckbSelectAll.Name = "ckbSelectAll";
            this.ckbSelectAll.Size = new System.Drawing.Size(119, 19);
            this.ckbSelectAll.TabIndex = 13;
            this.ckbSelectAll.Text = "Select All Options";
            this.ckbSelectAll.UseVisualStyleBackColor = true;
            this.ckbSelectAll.CheckedChanged += new System.EventHandler(this.CkbSelectAll_CheckedChanged);
            // 
            // FrmWordCommands
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(738, 541);
            this.Controls.Add(this.ckbSelectAll);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.BtnOk);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmWordCommands";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Word Commands ";
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmWordCommands_KeyDown);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button BtnOk;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.CheckBox ckbContentControls;
        private System.Windows.Forms.CheckBox ckbStyles;
        private System.Windows.Forms.CheckBox ckbListTemplates;
        private System.Windows.Forms.CheckBox ckbHyperlinks;
        private System.Windows.Forms.CheckBox ckbFieldCodes;
        private System.Windows.Forms.CheckBox ckbComments;
        private System.Windows.Forms.CheckBox ckbBookmarks;
        private System.Windows.Forms.CheckBox ckbDocProps;
        private System.Windows.Forms.CheckBox ckbEndnotes;
        private System.Windows.Forms.CheckBox ckbFootnotes;
        private System.Windows.Forms.CheckBox ckbFonts;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ListBox lbRevisions;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbAuthors;
        private System.Windows.Forms.CheckBox ckbRevisions;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.CheckBox ckbPackageParts;
        private System.Windows.Forms.CheckBox ckbShapes;
        private System.Windows.Forms.CheckBox ckbOleObjects;
        private System.Windows.Forms.CheckBox ckbSelectAll;
        private System.Windows.Forms.CheckBox ckbTables;
    }
}