
namespace Office_File_Explorer.WinForms
{
    partial class FrmPPTCommands
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmPPTCommands));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.ckbSlideTransitions = new System.Windows.Forms.CheckBox();
            this.ckbSlideText = new System.Windows.Forms.CheckBox();
            this.ckbComments = new System.Windows.Forms.CheckBox();
            this.ckbSlideTitles = new System.Windows.Forms.CheckBox();
            this.ckbHyperlinks = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.ckbPackageParts = new System.Windows.Forms.CheckBox();
            this.ckbShapes = new System.Windows.Forms.CheckBox();
            this.ckbOleObjects = new System.Windows.Forms.CheckBox();
            this.BtnOk = new System.Windows.Forms.Button();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.ckbSelectAll = new System.Windows.Forms.CheckBox();
            this.ckbListFonts = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.ckbListFonts);
            this.groupBox1.Controls.Add(this.ckbSlideTransitions);
            this.groupBox1.Controls.Add(this.ckbSlideText);
            this.groupBox1.Controls.Add(this.ckbComments);
            this.groupBox1.Controls.Add(this.ckbSlideTitles);
            this.groupBox1.Controls.Add(this.ckbHyperlinks);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(236, 179);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Select PowerPoint Content To Display";
            // 
            // ckbSlideTransitions
            // 
            this.ckbSlideTransitions.AutoSize = true;
            this.ckbSlideTransitions.Location = new System.Drawing.Point(19, 122);
            this.ckbSlideTransitions.Name = "ckbSlideTransitions";
            this.ckbSlideTransitions.Size = new System.Drawing.Size(110, 19);
            this.ckbSlideTransitions.TabIndex = 4;
            this.ckbSlideTransitions.Text = "Slide Transitions";
            this.ckbSlideTransitions.UseVisualStyleBackColor = true;
            // 
            // ckbSlideText
            // 
            this.ckbSlideText.AutoSize = true;
            this.ckbSlideText.Location = new System.Drawing.Point(19, 97);
            this.ckbSlideText.Name = "ckbSlideText";
            this.ckbSlideText.Size = new System.Drawing.Size(75, 19);
            this.ckbSlideText.TabIndex = 3;
            this.ckbSlideText.Text = "Slide Text";
            this.ckbSlideText.UseVisualStyleBackColor = true;
            // 
            // ckbComments
            // 
            this.ckbComments.AutoSize = true;
            this.ckbComments.Location = new System.Drawing.Point(19, 72);
            this.ckbComments.Name = "ckbComments";
            this.ckbComments.Size = new System.Drawing.Size(85, 19);
            this.ckbComments.TabIndex = 2;
            this.ckbComments.Text = "Comments";
            this.ckbComments.UseVisualStyleBackColor = true;
            // 
            // ckbSlideTitles
            // 
            this.ckbSlideTitles.AutoSize = true;
            this.ckbSlideTitles.Location = new System.Drawing.Point(19, 47);
            this.ckbSlideTitles.Name = "ckbSlideTitles";
            this.ckbSlideTitles.Size = new System.Drawing.Size(81, 19);
            this.ckbSlideTitles.TabIndex = 1;
            this.ckbSlideTitles.Text = "Slide Titles";
            this.ckbSlideTitles.UseVisualStyleBackColor = true;
            // 
            // ckbHyperlinks
            // 
            this.ckbHyperlinks.AutoSize = true;
            this.ckbHyperlinks.Location = new System.Drawing.Point(19, 22);
            this.ckbHyperlinks.Name = "ckbHyperlinks";
            this.ckbHyperlinks.Size = new System.Drawing.Size(82, 19);
            this.ckbHyperlinks.TabIndex = 0;
            this.ckbHyperlinks.Text = "Hyperlinks";
            this.ckbHyperlinks.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.ckbPackageParts);
            this.groupBox2.Controls.Add(this.ckbShapes);
            this.groupBox2.Controls.Add(this.ckbOleObjects);
            this.groupBox2.Location = new System.Drawing.Point(254, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(207, 179);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Select Office Content To Display";
            // 
            // ckbPackageParts
            // 
            this.ckbPackageParts.AutoSize = true;
            this.ckbPackageParts.Location = new System.Drawing.Point(27, 72);
            this.ckbPackageParts.Name = "ckbPackageParts";
            this.ckbPackageParts.Size = new System.Drawing.Size(99, 19);
            this.ckbPackageParts.TabIndex = 2;
            this.ckbPackageParts.Text = "Package Parts";
            this.ckbPackageParts.UseVisualStyleBackColor = true;
            // 
            // ckbShapes
            // 
            this.ckbShapes.AutoSize = true;
            this.ckbShapes.Location = new System.Drawing.Point(27, 47);
            this.ckbShapes.Name = "ckbShapes";
            this.ckbShapes.Size = new System.Drawing.Size(63, 19);
            this.ckbShapes.TabIndex = 1;
            this.ckbShapes.Text = "Shapes";
            this.ckbShapes.UseVisualStyleBackColor = true;
            // 
            // ckbOleObjects
            // 
            this.ckbOleObjects.AutoSize = true;
            this.ckbOleObjects.Location = new System.Drawing.Point(27, 22);
            this.ckbOleObjects.Name = "ckbOleObjects";
            this.ckbOleObjects.Size = new System.Drawing.Size(87, 19);
            this.ckbOleObjects.TabIndex = 0;
            this.ckbOleObjects.Text = "Ole Objects";
            this.ckbOleObjects.UseVisualStyleBackColor = true;
            // 
            // BtnOk
            // 
            this.BtnOk.Location = new System.Drawing.Point(305, 197);
            this.BtnOk.Name = "BtnOk";
            this.BtnOk.Size = new System.Drawing.Size(75, 23);
            this.BtnOk.TabIndex = 5;
            this.BtnOk.Text = "OK";
            this.BtnOk.UseVisualStyleBackColor = true;
            this.BtnOk.Click += new System.EventHandler(this.BtnOk_Click);
            // 
            // BtnCancel
            // 
            this.BtnCancel.Location = new System.Drawing.Point(386, 197);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(75, 23);
            this.BtnCancel.TabIndex = 6;
            this.BtnCancel.Text = "Cancel";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // ckbSelectAll
            // 
            this.ckbSelectAll.AutoSize = true;
            this.ckbSelectAll.Location = new System.Drawing.Point(12, 197);
            this.ckbSelectAll.Name = "ckbSelectAll";
            this.ckbSelectAll.Size = new System.Drawing.Size(119, 19);
            this.ckbSelectAll.TabIndex = 7;
            this.ckbSelectAll.Text = "Select All Options";
            this.ckbSelectAll.UseVisualStyleBackColor = true;
            this.ckbSelectAll.CheckedChanged += new System.EventHandler(this.CkbSelectAll_CheckedChanged);
            // 
            // ckbListFonts
            // 
            this.ckbListFonts.AutoSize = true;
            this.ckbListFonts.Location = new System.Drawing.Point(19, 147);
            this.ckbListFonts.Name = "ckbListFonts";
            this.ckbListFonts.Size = new System.Drawing.Size(76, 19);
            this.ckbListFonts.TabIndex = 8;
            this.ckbListFonts.Text = "List Fonts";
            this.ckbListFonts.UseVisualStyleBackColor = true;
            // 
            // FrmPPTCommands
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(477, 228);
            this.Controls.Add(this.ckbSelectAll);
            this.Controls.Add(this.BtnOk);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmPPTCommands";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "PowerPoint Commands";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox ckbSlideTransitions;
        private System.Windows.Forms.CheckBox ckbSlideText;
        private System.Windows.Forms.CheckBox ckbComments;
        private System.Windows.Forms.CheckBox ckbSlideTitles;
        private System.Windows.Forms.CheckBox ckbHyperlinks;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox ckbPackageParts;
        private System.Windows.Forms.CheckBox ckbShapes;
        private System.Windows.Forms.CheckBox ckbOleObjects;
        private System.Windows.Forms.Button BtnOk;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.CheckBox ckbSelectAll;
        private System.Windows.Forms.CheckBox ckbListFonts;
    }
}