
namespace Office_File_Explorer.WinForms
{
    partial class FrmClipboardViewer
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmClipboardViewer));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.clipboardToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.refreshToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ownerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.clearToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveAsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.autoRefreshToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.showRichTextToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.showMemoryInHexToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.showPicturesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lbClipFormats = new System.Windows.Forms.ListBox();
            this.pbClipData = new System.Windows.Forms.PictureBox();
            this.rtbClipData = new System.Windows.Forms.RichTextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbClipData)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.clipboardToolStripMenuItem,
            this.viewToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(800, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // clipboardToolStripMenuItem
            // 
            this.clipboardToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.refreshToolStripMenuItem,
            this.ownerToolStripMenuItem,
            this.clearToolStripMenuItem,
            this.saveAsToolStripMenuItem});
            this.clipboardToolStripMenuItem.Name = "clipboardToolStripMenuItem";
            this.clipboardToolStripMenuItem.Size = new System.Drawing.Size(71, 20);
            this.clipboardToolStripMenuItem.Text = "Clipboard";
            // 
            // refreshToolStripMenuItem
            // 
            this.refreshToolStripMenuItem.Image = global::Office_File_Explorer.Properties.Resources.Exit_16x1;
            this.refreshToolStripMenuItem.Name = "refreshToolStripMenuItem";
            this.refreshToolStripMenuItem.Size = new System.Drawing.Size(114, 22);
            this.refreshToolStripMenuItem.Text = "Refresh";
            this.refreshToolStripMenuItem.Click += new System.EventHandler(this.RefreshToolStripMenuItem_Click);
            // 
            // ownerToolStripMenuItem
            // 
            this.ownerToolStripMenuItem.Image = global::Office_File_Explorer.Properties.Resources.FontDialogControl_16x;
            this.ownerToolStripMenuItem.Name = "ownerToolStripMenuItem";
            this.ownerToolStripMenuItem.Size = new System.Drawing.Size(114, 22);
            this.ownerToolStripMenuItem.Text = "Owner";
            this.ownerToolStripMenuItem.Click += new System.EventHandler(this.OwnerToolStripMenuItem_Click);
            // 
            // clearToolStripMenuItem
            // 
            this.clearToolStripMenuItem.Image = global::Office_File_Explorer.Properties.Resources.TableMissing_8931_32;
            this.clearToolStripMenuItem.Name = "clearToolStripMenuItem";
            this.clearToolStripMenuItem.Size = new System.Drawing.Size(114, 22);
            this.clearToolStripMenuItem.Text = "Clear";
            this.clearToolStripMenuItem.Click += new System.EventHandler(this.ClearToolStripMenuItem_Click);
            // 
            // saveAsToolStripMenuItem
            // 
            this.saveAsToolStripMenuItem.Image = global::Office_File_Explorer.Properties.Resources.SaveAs_16x;
            this.saveAsToolStripMenuItem.Name = "saveAsToolStripMenuItem";
            this.saveAsToolStripMenuItem.Size = new System.Drawing.Size(114, 22);
            this.saveAsToolStripMenuItem.Text = "Save As";
            this.saveAsToolStripMenuItem.Click += new System.EventHandler(this.SaveAsToolStripMenuItem_Click);
            // 
            // viewToolStripMenuItem
            // 
            this.viewToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.autoRefreshToolStripMenuItem,
            this.showRichTextToolStripMenuItem,
            this.showMemoryInHexToolStripMenuItem,
            this.showPicturesToolStripMenuItem});
            this.viewToolStripMenuItem.Name = "viewToolStripMenuItem";
            this.viewToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.viewToolStripMenuItem.Text = "View";
            // 
            // autoRefreshToolStripMenuItem
            // 
            this.autoRefreshToolStripMenuItem.Image = global::Office_File_Explorer.Properties.Resources.RefreshServer_16x;
            this.autoRefreshToolStripMenuItem.Name = "autoRefreshToolStripMenuItem";
            this.autoRefreshToolStripMenuItem.Size = new System.Drawing.Size(188, 22);
            this.autoRefreshToolStripMenuItem.Text = "Auto Refresh";
            this.autoRefreshToolStripMenuItem.Click += new System.EventHandler(this.AutoRefreshToolStripMenuItem_Click);
            // 
            // showRichTextToolStripMenuItem
            // 
            this.showRichTextToolStripMenuItem.Image = global::Office_File_Explorer.Properties.Resources.RichTextBox_16x;
            this.showRichTextToolStripMenuItem.Name = "showRichTextToolStripMenuItem";
            this.showRichTextToolStripMenuItem.Size = new System.Drawing.Size(188, 22);
            this.showRichTextToolStripMenuItem.Text = "Show Rich Text";
            this.showRichTextToolStripMenuItem.Click += new System.EventHandler(this.ShowRichTextToolStripMenuItem_Click);
            // 
            // showMemoryInHexToolStripMenuItem
            // 
            this.showMemoryInHexToolStripMenuItem.Image = global::Office_File_Explorer.Properties.Resources.Memory_16x;
            this.showMemoryInHexToolStripMenuItem.Name = "showMemoryInHexToolStripMenuItem";
            this.showMemoryInHexToolStripMenuItem.Size = new System.Drawing.Size(188, 22);
            this.showMemoryInHexToolStripMenuItem.Text = "Show Memory In Hex";
            this.showMemoryInHexToolStripMenuItem.Click += new System.EventHandler(this.ShowMemoryInHexToolStripMenuItem_Click);
            // 
            // showPicturesToolStripMenuItem
            // 
            this.showPicturesToolStripMenuItem.Image = global::Office_File_Explorer.Properties.Resources.ImageIcon_16x;
            this.showPicturesToolStripMenuItem.Name = "showPicturesToolStripMenuItem";
            this.showPicturesToolStripMenuItem.Size = new System.Drawing.Size(188, 22);
            this.showPicturesToolStripMenuItem.Text = "Show Pictures";
            this.showPicturesToolStripMenuItem.Click += new System.EventHandler(this.ShowPicturesToolStripMenuItem_Click);
            // 
            // lbClipFormats
            // 
            this.lbClipFormats.Dock = System.Windows.Forms.DockStyle.Left;
            this.lbClipFormats.FormattingEnabled = true;
            this.lbClipFormats.ItemHeight = 15;
            this.lbClipFormats.Location = new System.Drawing.Point(0, 24);
            this.lbClipFormats.Name = "lbClipFormats";
            this.lbClipFormats.Size = new System.Drawing.Size(276, 530);
            this.lbClipFormats.TabIndex = 1;
            this.lbClipFormats.SelectedIndexChanged += new System.EventHandler(this.LbClipFormats_SelectedIndexChanged);
            // 
            // pbClipData
            // 
            this.pbClipData.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pbClipData.Location = new System.Drawing.Point(3, 19);
            this.pbClipData.Name = "pbClipData";
            this.pbClipData.Size = new System.Drawing.Size(512, 176);
            this.pbClipData.TabIndex = 2;
            this.pbClipData.TabStop = false;
            // 
            // rtbClipData
            // 
            this.rtbClipData.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rtbClipData.Location = new System.Drawing.Point(3, 19);
            this.rtbClipData.Name = "rtbClipData";
            this.rtbClipData.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.ForcedBoth;
            this.rtbClipData.Size = new System.Drawing.Size(512, 288);
            this.rtbClipData.TabIndex = 3;
            this.rtbClipData.Text = "";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.groupBox3);
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(276, 24);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(524, 530);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Clipboard Data";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.pbClipData);
            this.groupBox3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox3.Location = new System.Drawing.Point(3, 329);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(518, 198);
            this.groupBox3.TabIndex = 0;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Image";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.rtbClipData);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox2.Location = new System.Drawing.Point(3, 19);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(518, 310);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Text";
            // 
            // FrmClipboardViewer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 554);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.lbClipFormats);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmClipboardViewer";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Clipboard Viewer";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FrmClipboardViewer_FormClosed);
            this.Shown += new System.EventHandler(this.FrmClipboardViewer_Shown);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmClipboardViewer_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbClipData)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem clipboardToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem refreshToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ownerToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem clearToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveAsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem autoRefreshToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem showRichTextToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem showMemoryInHexToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem showPicturesToolStripMenuItem;
        private System.Windows.Forms.ListBox lbClipFormats;
        private System.Windows.Forms.PictureBox pbClipData;
        private System.Windows.Forms.RichTextBox rtbClipData;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox2;
    }
}