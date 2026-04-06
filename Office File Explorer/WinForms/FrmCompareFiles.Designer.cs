namespace Office_File_Explorer.WinForms
{
    partial class FrmCompareFiles
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmCompareFiles));
            tvLeft = new System.Windows.Forms.TreeView();
            tvRight = new System.Windows.Forms.TreeView();
            scintillaDiffControl1 = new ScintillaDiff.ScintillaDiffControl();
            BtnFileLeft = new System.Windows.Forms.Button();
            BtnFileRight = new System.Windows.Forms.Button();
            lblDiffCount = new System.Windows.Forms.Label();
            groupBox1 = new System.Windows.Forms.GroupBox();
            splitContainer1 = new System.Windows.Forms.SplitContainer();
            panelTop = new System.Windows.Forms.Panel();
            groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)splitContainer1).BeginInit();
            splitContainer1.Panel1.SuspendLayout();
            splitContainer1.Panel2.SuspendLayout();
            splitContainer1.SuspendLayout();
            panelTop.SuspendLayout();
            SuspendLayout();
            // 
            // tvLeft
            // 
            tvLeft.Dock = System.Windows.Forms.DockStyle.Fill;
            tvLeft.Location = new System.Drawing.Point(0, 0);
            tvLeft.Name = "tvLeft";
            tvLeft.Size = new System.Drawing.Size(658, 536);
            tvLeft.TabIndex = 0;
            tvLeft.AfterSelect += TvLeft_AfterSelect;
            // 
            // tvRight
            // 
            tvRight.Dock = System.Windows.Forms.DockStyle.Fill;
            tvRight.Location = new System.Drawing.Point(0, 0);
            tvRight.Name = "tvRight";
            tvRight.Size = new System.Drawing.Size(664, 536);
            tvRight.TabIndex = 1;
            tvRight.AfterSelect += TvRight_AfterSelect;
            // 
            // scintillaDiffControl1
            // 
            scintillaDiffControl1.AddedCharacterSymbol = '+';
            scintillaDiffControl1.CharacterComparison = false;
            scintillaDiffControl1.CharacterComparisonMarkAddRemove = false;
            scintillaDiffControl1.DiffColorAdded = System.Drawing.Color.FromArgb(212, 242, 196);
            scintillaDiffControl1.DiffColorChangeBackground = System.Drawing.Color.FromArgb(252, 255, 140);
            scintillaDiffControl1.DiffColorCharAdded = System.Drawing.Color.FromArgb(154, 234, 111);
            scintillaDiffControl1.DiffColorCharDeleted = System.Drawing.Color.FromArgb(225, 125, 125);
            scintillaDiffControl1.DiffColorDeleted = System.Drawing.Color.FromArgb(255, 178, 178);
            scintillaDiffControl1.DiffStyle = ScintillaDiff.ScintillaDiffStyles.DiffStyle.DiffSideBySide;
            scintillaDiffControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            scintillaDiffControl1.ImageRowAdded = (System.Drawing.Bitmap)resources.GetObject("scintillaDiffControl1.ImageRowAdded");
            scintillaDiffControl1.ImageRowAddedScintillaIndex = 28;
            scintillaDiffControl1.ImageRowDeleted = (System.Drawing.Bitmap)resources.GetObject("scintillaDiffControl1.ImageRowDeleted");
            scintillaDiffControl1.ImageRowDeletedScintillaIndex = 29;
            scintillaDiffControl1.ImageRowDiff = (System.Drawing.Bitmap)resources.GetObject("scintillaDiffControl1.ImageRowDiff");
            scintillaDiffControl1.ImageRowDiffScintillaIndex = 31;
            scintillaDiffControl1.ImageRowOk = (System.Drawing.Bitmap)resources.GetObject("scintillaDiffControl1.ImageRowOk");
            scintillaDiffControl1.ImageRowOkScintillaIndex = 30;
            scintillaDiffControl1.IsEntireLineHighlighted = true;
            scintillaDiffControl1.Location = new System.Drawing.Point(3, 563);
            scintillaDiffControl1.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            scintillaDiffControl1.MarkColorIndexModifiedBackground = 31;
            scintillaDiffControl1.MarkColorIndexRemovedOrAdded = 30;
            scintillaDiffControl1.Name = "scintillaDiffControl1";
            scintillaDiffControl1.RemovedCharacterSymbol = '-';
            scintillaDiffControl1.Size = new System.Drawing.Size(1326, 752);
            scintillaDiffControl1.TabIndex = 2;
            scintillaDiffControl1.TextLeft = "";
            scintillaDiffControl1.TextRight = "";
            scintillaDiffControl1.UseRowOkSign = false;
            // 
            // BtnFileLeft
            // 
            BtnFileLeft.Dock = System.Windows.Forms.DockStyle.Left;
            BtnFileLeft.Location = new System.Drawing.Point(0, 0);
            BtnFileLeft.Name = "BtnFileLeft";
            BtnFileLeft.Size = new System.Drawing.Size(118, 44);
            BtnFileLeft.TabIndex = 5;
            BtnFileLeft.Text = "Open File";
            BtnFileLeft.UseVisualStyleBackColor = true;
            BtnFileLeft.Click += BtnFileLeft_Click;
            // 
            // BtnFileRight
            // 
            BtnFileRight.Dock = System.Windows.Forms.DockStyle.Right;
            BtnFileRight.Location = new System.Drawing.Point(1220, 0);
            BtnFileRight.Name = "BtnFileRight";
            BtnFileRight.Size = new System.Drawing.Size(112, 44);
            BtnFileRight.TabIndex = 6;
            BtnFileRight.Text = "Open File";
            BtnFileRight.UseVisualStyleBackColor = true;
            BtnFileRight.Click += BtnFileRight_Click;
            // 
            // lblDiffCount
            // 
            lblDiffCount.Dock = System.Windows.Forms.DockStyle.Fill;
            lblDiffCount.Name = "lblDiffCount";
            lblDiffCount.TabIndex = 9;
            lblDiffCount.Text = "";
            lblDiffCount.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(scintillaDiffControl1);
            groupBox1.Controls.Add(splitContainer1);
            groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            groupBox1.Location = new System.Drawing.Point(0, 44);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new System.Drawing.Size(1332, 1318);
            groupBox1.TabIndex = 7;
            groupBox1.TabStop = false;
            groupBox1.Text = "Files";
            // 
            // splitContainer1
            // 
            splitContainer1.Dock = System.Windows.Forms.DockStyle.Top;
            splitContainer1.Location = new System.Drawing.Point(3, 27);
            splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            splitContainer1.Panel1.Controls.Add(tvLeft);
            // 
            // splitContainer1.Panel2
            // 
            splitContainer1.Panel2.Controls.Add(tvRight);
            splitContainer1.Size = new System.Drawing.Size(1326, 536);
            splitContainer1.SplitterDistance = 658;
            splitContainer1.TabIndex = 3;
            // 
            // panelTop
            // 
            panelTop.Controls.Add(lblDiffCount);
            panelTop.Controls.Add(BtnFileRight);
            panelTop.Controls.Add(BtnFileLeft);
            panelTop.Dock = System.Windows.Forms.DockStyle.Top;
            panelTop.Location = new System.Drawing.Point(0, 0);
            panelTop.Name = "panelTop";
            panelTop.Size = new System.Drawing.Size(1332, 44);
            panelTop.TabIndex = 8;
            // 
            // FrmCompareFiles
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(10F, 25F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(1332, 1362);
            Controls.Add(groupBox1);
            Controls.Add(panelTop);
            Name = "FrmCompareFiles";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            Text = "Compare Files";
            groupBox1.ResumeLayout(false);
            splitContainer1.Panel1.ResumeLayout(false);
            splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)splitContainer1).EndInit();
            splitContainer1.ResumeLayout(false);
            panelTop.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.TreeView tvLeft;
        private System.Windows.Forms.TreeView tvRight;
        private ScintillaDiff.ScintillaDiffControl scintillaDiffControl1;
        private System.Windows.Forms.Button BtnFileLeft;
        private System.Windows.Forms.Button BtnFileRight;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.Panel panelTop;
        private System.Windows.Forms.Label lblDiffCount;
    }
}