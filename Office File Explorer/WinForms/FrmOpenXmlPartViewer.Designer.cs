namespace Office_File_Explorer.WinForms
{
    partial class FrmOpenXmlPartViewer
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
            components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmOpenXmlPartViewer));
            toolStrip1 = new System.Windows.Forms.ToolStrip();
            toolStripButtonModifyXml = new System.Windows.Forms.ToolStripButton();
            toolStripButtonSave = new System.Windows.Forms.ToolStripButton();
            toolStripButtonInsertIcon = new System.Windows.Forms.ToolStripButton();
            toolStripButtonValidateXml = new System.Windows.Forms.ToolStripButton();
            toolStripButtonGenerateCallbacks = new System.Windows.Forms.ToolStripButton();
            toolStripDropDownButton1 = new System.Windows.Forms.ToolStripDropDownButton();
            toolStripMenuInsertO14CustomUI = new System.Windows.Forms.ToolStripMenuItem();
            toolStripMenuInsertO12CustomUIPart = new System.Windows.Forms.ToolStripMenuItem();
            toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            customOutspaceToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            customTabToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            excelCustomTabToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            repurposeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            wordGroupOnInsertTabToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            splitContainer1 = new System.Windows.Forms.SplitContainer();
            treeView1 = new System.Windows.Forms.TreeView();
            tvImageList = new System.Windows.Forms.ImageList(components);
            rtbPartContents = new System.Windows.Forms.RichTextBox();
            toolStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)splitContainer1).BeginInit();
            splitContainer1.Panel1.SuspendLayout();
            splitContainer1.Panel2.SuspendLayout();
            splitContainer1.SuspendLayout();
            SuspendLayout();
            // 
            // toolStrip1
            // 
            toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { toolStripButtonModifyXml, toolStripButtonSave, toolStripButtonInsertIcon, toolStripButtonValidateXml, toolStripButtonGenerateCallbacks, toolStripDropDownButton1 });
            toolStrip1.Location = new System.Drawing.Point(0, 0);
            toolStrip1.Name = "toolStrip1";
            toolStrip1.Size = new System.Drawing.Size(800, 25);
            toolStrip1.TabIndex = 0;
            toolStrip1.Text = "toolStrip1";
            // 
            // toolStripButtonModifyXml
            // 
            toolStripButtonModifyXml.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            toolStripButtonModifyXml.Image = Properties.Resources.ModifyPropertyTrivial;
            toolStripButtonModifyXml.ImageTransparentColor = System.Drawing.Color.Magenta;
            toolStripButtonModifyXml.Name = "toolStripButtonModifyXml";
            toolStripButtonModifyXml.Size = new System.Drawing.Size(23, 22);
            toolStripButtonModifyXml.Text = "Modify Xml";
            toolStripButtonModifyXml.Click += toolStripButtonModifyXml_Click;
            // 
            // toolStripButtonSave
            // 
            toolStripButtonSave.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            toolStripButtonSave.Enabled = false;
            toolStripButtonSave.Image = Properties.Resources.Save;
            toolStripButtonSave.ImageTransparentColor = System.Drawing.Color.Magenta;
            toolStripButtonSave.Name = "toolStripButtonSave";
            toolStripButtonSave.Size = new System.Drawing.Size(23, 22);
            toolStripButtonSave.Text = "toolStripButton1";
            toolStripButtonSave.ToolTipText = "Save Document Part";
            toolStripButtonSave.Click += toolStripButtonSave_Click;
            // 
            // toolStripButtonInsertIcon
            // 
            toolStripButtonInsertIcon.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            toolStripButtonInsertIcon.Enabled = false;
            toolStripButtonInsertIcon.Image = (System.Drawing.Image)resources.GetObject("toolStripButtonInsertIcon.Image");
            toolStripButtonInsertIcon.ImageTransparentColor = System.Drawing.Color.Magenta;
            toolStripButtonInsertIcon.Name = "toolStripButtonInsertIcon";
            toolStripButtonInsertIcon.Size = new System.Drawing.Size(23, 22);
            toolStripButtonInsertIcon.Text = "toolStripButton1";
            toolStripButtonInsertIcon.ToolTipText = "Insert Icon";
            toolStripButtonInsertIcon.Click += toolStripButtonInsertIcon_Click;
            // 
            // toolStripButtonValidateXml
            // 
            toolStripButtonValidateXml.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            toolStripButtonValidateXml.Enabled = false;
            toolStripButtonValidateXml.Image = (System.Drawing.Image)resources.GetObject("toolStripButtonValidateXml.Image");
            toolStripButtonValidateXml.ImageTransparentColor = System.Drawing.Color.Magenta;
            toolStripButtonValidateXml.Name = "toolStripButtonValidateXml";
            toolStripButtonValidateXml.Size = new System.Drawing.Size(23, 22);
            toolStripButtonValidateXml.Text = "toolStripButton2";
            toolStripButtonValidateXml.ToolTipText = "Validate Xml";
            toolStripButtonValidateXml.Click += toolStripButtonValidateXml_Click;
            // 
            // toolStripButtonGenerateCallbacks
            // 
            toolStripButtonGenerateCallbacks.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            toolStripButtonGenerateCallbacks.Enabled = false;
            toolStripButtonGenerateCallbacks.Image = (System.Drawing.Image)resources.GetObject("toolStripButtonGenerateCallbacks.Image");
            toolStripButtonGenerateCallbacks.ImageTransparentColor = System.Drawing.Color.Magenta;
            toolStripButtonGenerateCallbacks.Name = "toolStripButtonGenerateCallbacks";
            toolStripButtonGenerateCallbacks.Size = new System.Drawing.Size(23, 22);
            toolStripButtonGenerateCallbacks.Text = "toolStripButton3";
            toolStripButtonGenerateCallbacks.ToolTipText = "Generate Callback";
            toolStripButtonGenerateCallbacks.Click += toolStripButtonGenerateCallbacks_Click;
            // 
            // toolStripDropDownButton1
            // 
            toolStripDropDownButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            toolStripDropDownButton1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { toolStripMenuInsertO14CustomUI, toolStripMenuInsertO12CustomUIPart, toolStripMenuItem1 });
            toolStripDropDownButton1.Image = (System.Drawing.Image)resources.GetObject("toolStripDropDownButton1.Image");
            toolStripDropDownButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            toolStripDropDownButton1.Name = "toolStripDropDownButton1";
            toolStripDropDownButton1.Size = new System.Drawing.Size(49, 22);
            toolStripDropDownButton1.Text = "Insert";
            // 
            // toolStripMenuInsertO14CustomUI
            // 
            toolStripMenuInsertO14CustomUI.Image = (System.Drawing.Image)resources.GetObject("toolStripMenuInsertO14CustomUI.Image");
            toolStripMenuInsertO14CustomUI.Name = "toolStripMenuInsertO14CustomUI";
            toolStripMenuInsertO14CustomUI.Size = new System.Drawing.Size(216, 22);
            toolStripMenuInsertO14CustomUI.Text = "Office 2010 Custom UI Part";
            toolStripMenuInsertO14CustomUI.Click += toolStripMenuInsertO14CustomUI_Click;
            // 
            // toolStripMenuInsertO12CustomUIPart
            // 
            toolStripMenuInsertO12CustomUIPart.Image = (System.Drawing.Image)resources.GetObject("toolStripMenuInsertO12CustomUIPart.Image");
            toolStripMenuInsertO12CustomUIPart.Name = "toolStripMenuInsertO12CustomUIPart";
            toolStripMenuInsertO12CustomUIPart.Size = new System.Drawing.Size(216, 22);
            toolStripMenuInsertO12CustomUIPart.Text = "Office 2007 Custom UI Part";
            toolStripMenuInsertO12CustomUIPart.Click += toolStripMenuInsertO12CustomUIPart_Click;
            // 
            // toolStripMenuItem1
            // 
            toolStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { customOutspaceToolStripMenuItem, customTabToolStripMenuItem, excelCustomTabToolStripMenuItem, repurposeToolStripMenuItem, wordGroupOnInsertTabToolStripMenuItem });
            toolStripMenuItem1.Name = "toolStripMenuItem1";
            toolStripMenuItem1.Size = new System.Drawing.Size(216, 22);
            toolStripMenuItem1.Text = "Sample XML";
            // 
            // customOutspaceToolStripMenuItem
            // 
            customOutspaceToolStripMenuItem.Image = (System.Drawing.Image)resources.GetObject("customOutspaceToolStripMenuItem.Image");
            customOutspaceToolStripMenuItem.Name = "customOutspaceToolStripMenuItem";
            customOutspaceToolStripMenuItem.Size = new System.Drawing.Size(217, 22);
            customOutspaceToolStripMenuItem.Text = "Custom Outspace";
            customOutspaceToolStripMenuItem.Click += xmlToolStripMenuItem_Click;
            // 
            // customTabToolStripMenuItem
            // 
            customTabToolStripMenuItem.Image = Properties.Resources.XmlFile;
            customTabToolStripMenuItem.Name = "customTabToolStripMenuItem";
            customTabToolStripMenuItem.Size = new System.Drawing.Size(217, 22);
            customTabToolStripMenuItem.Text = "Custom Tab";
            customTabToolStripMenuItem.Click += customTabToolStripMenuItem_Click;
            // 
            // excelCustomTabToolStripMenuItem
            // 
            excelCustomTabToolStripMenuItem.Image = Properties.Resources.XmlFile;
            excelCustomTabToolStripMenuItem.Name = "excelCustomTabToolStripMenuItem";
            excelCustomTabToolStripMenuItem.Size = new System.Drawing.Size(217, 22);
            excelCustomTabToolStripMenuItem.Text = "Excel - Custom Tab";
            excelCustomTabToolStripMenuItem.Click += excelCustomTabToolStripMenuItem_Click;
            // 
            // repurposeToolStripMenuItem
            // 
            repurposeToolStripMenuItem.Image = Properties.Resources.XmlFile;
            repurposeToolStripMenuItem.Name = "repurposeToolStripMenuItem";
            repurposeToolStripMenuItem.Size = new System.Drawing.Size(217, 22);
            repurposeToolStripMenuItem.Text = "Repurpose";
            repurposeToolStripMenuItem.Click += repurposeToolStripMenuItem_Click;
            // 
            // wordGroupOnInsertTabToolStripMenuItem
            // 
            wordGroupOnInsertTabToolStripMenuItem.Image = Properties.Resources.XmlFile;
            wordGroupOnInsertTabToolStripMenuItem.Name = "wordGroupOnInsertTabToolStripMenuItem";
            wordGroupOnInsertTabToolStripMenuItem.Size = new System.Drawing.Size(217, 22);
            wordGroupOnInsertTabToolStripMenuItem.Text = "Word - Group on Insert Tab";
            wordGroupOnInsertTabToolStripMenuItem.Click += wordGroupOnInsertTabToolStripMenuItem_Click;
            // 
            // splitContainer1
            // 
            splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            splitContainer1.Location = new System.Drawing.Point(0, 25);
            splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            splitContainer1.Panel1.Controls.Add(treeView1);
            // 
            // splitContainer1.Panel2
            // 
            splitContainer1.Panel2.Controls.Add(rtbPartContents);
            splitContainer1.Size = new System.Drawing.Size(800, 425);
            splitContainer1.SplitterDistance = 266;
            splitContainer1.TabIndex = 1;
            // 
            // treeView1
            // 
            treeView1.Dock = System.Windows.Forms.DockStyle.Fill;
            treeView1.ImageIndex = 0;
            treeView1.ImageList = tvImageList;
            treeView1.Location = new System.Drawing.Point(0, 0);
            treeView1.Name = "treeView1";
            treeView1.SelectedImageIndex = 0;
            treeView1.Size = new System.Drawing.Size(266, 425);
            treeView1.TabIndex = 0;
            treeView1.AfterSelect += treeView1_AfterSelect;
            // 
            // tvImageList
            // 
            tvImageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit;
            tvImageList.ImageStream = (System.Windows.Forms.ImageListStreamer)resources.GetObject("tvImageList.ImageStream");
            tvImageList.TransparentColor = System.Drawing.Color.Transparent;
            tvImageList.Images.SetKeyName(0, "worddoc.bmp");
            tvImageList.Images.SetKeyName(1, "pptpre.bmp");
            tvImageList.Images.SetKeyName(2, "excelwkb.bmp");
            tvImageList.Images.SetKeyName(3, "xml.png");
            tvImageList.Images.SetKeyName(4, "insertPicture.png");
            tvImageList.Images.SetKeyName(5, "BinaryFile.png");
            tvImageList.Images.SetKeyName(6, "folder.png");
            tvImageList.Images.SetKeyName(7, "textfile icon.png");
            // 
            // rtbPartContents
            // 
            rtbPartContents.Dock = System.Windows.Forms.DockStyle.Fill;
            rtbPartContents.Location = new System.Drawing.Point(0, 0);
            rtbPartContents.Name = "rtbPartContents";
            rtbPartContents.ReadOnly = true;
            rtbPartContents.Size = new System.Drawing.Size(530, 425);
            rtbPartContents.TabIndex = 0;
            rtbPartContents.Text = "";
            // 
            // FrmOpenXmlPartViewer
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(800, 450);
            Controls.Add(splitContainer1);
            Controls.Add(toolStrip1);
            Icon = (System.Drawing.Icon)resources.GetObject("$this.Icon");
            Name = "FrmOpenXmlPartViewer";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            Text = "Open Xml Part Viewer";
            FormClosing += FrmOpenXmlPartViewer_FormClosing;
            KeyDown += FrmOpenXmlPartViewer_KeyDown;
            toolStrip1.ResumeLayout(false);
            toolStrip1.PerformLayout();
            splitContainer1.Panel1.ResumeLayout(false);
            splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)splitContainer1).EndInit();
            splitContainer1.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton toolStripButtonModifyXml;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.RichTextBox rtbPartContents;
        private System.Windows.Forms.ToolStripButton toolStripButtonSave;
        private System.Windows.Forms.ToolStripButton toolStripButtonInsertIcon;
        private System.Windows.Forms.ToolStripButton toolStripButtonValidateXml;
        private System.Windows.Forms.ToolStripButton toolStripButtonGenerateCallbacks;
        private System.Windows.Forms.ToolStripDropDownButton toolStripDropDownButton1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuInsertO14CustomUI;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuInsertO12CustomUIPart;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem customOutspaceToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem customTabToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem excelCustomTabToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem repurposeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem wordGroupOnInsertTabToolStripMenuItem;
        private System.Windows.Forms.ImageList tvImageList;
    }
}