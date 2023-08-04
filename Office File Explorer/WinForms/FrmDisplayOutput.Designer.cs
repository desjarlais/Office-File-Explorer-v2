namespace Office_File_Explorer.WinForms
{
    partial class FrmDisplayOutput
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
            splitContainer1 = new System.Windows.Forms.SplitContainer();
            rtbRTFContent = new System.Windows.Forms.RichTextBox();
            pictureBox1 = new System.Windows.Forms.PictureBox();
            contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(components);
            copySelectedTextToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            copyAllTextToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)splitContainer1).BeginInit();
            splitContainer1.Panel1.SuspendLayout();
            splitContainer1.Panel2.SuspendLayout();
            splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).BeginInit();
            contextMenuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // splitContainer1
            // 
            splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            splitContainer1.Location = new System.Drawing.Point(0, 0);
            splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            splitContainer1.Panel1.Controls.Add(rtbRTFContent);
            // 
            // splitContainer1.Panel2
            // 
            splitContainer1.Panel2.Controls.Add(pictureBox1);
            splitContainer1.Size = new System.Drawing.Size(769, 552);
            splitContainer1.SplitterDistance = 371;
            splitContainer1.TabIndex = 0;
            // 
            // rtbRTFContent
            // 
            rtbRTFContent.ContextMenuStrip = contextMenuStrip1;
            rtbRTFContent.Dock = System.Windows.Forms.DockStyle.Fill;
            rtbRTFContent.Location = new System.Drawing.Point(0, 0);
            rtbRTFContent.Name = "rtbRTFContent";
            rtbRTFContent.Size = new System.Drawing.Size(371, 552);
            rtbRTFContent.TabIndex = 0;
            rtbRTFContent.Text = "";
            // 
            // pictureBox1
            // 
            pictureBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            pictureBox1.Location = new System.Drawing.Point(0, 0);
            pictureBox1.Name = "pictureBox1";
            pictureBox1.Size = new System.Drawing.Size(394, 552);
            pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            pictureBox1.TabIndex = 0;
            pictureBox1.TabStop = false;
            // 
            // contextMenuStrip1
            // 
            contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { copySelectedTextToolStripMenuItem, copyAllTextToolStripMenuItem });
            contextMenuStrip1.Name = "contextMenuStrip1";
            contextMenuStrip1.Size = new System.Drawing.Size(181, 70);
            // 
            // copySelectedTextToolStripMenuItem
            // 
            copySelectedTextToolStripMenuItem.Name = "copySelectedTextToolStripMenuItem";
            copySelectedTextToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            copySelectedTextToolStripMenuItem.Text = "Copy Selected Text";
            copySelectedTextToolStripMenuItem.Click += copySelectedTextToolStripMenuItem_Click;
            // 
            // copyAllTextToolStripMenuItem
            // 
            copyAllTextToolStripMenuItem.Name = "copyAllTextToolStripMenuItem";
            copyAllTextToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            copyAllTextToolStripMenuItem.Text = "Copy All Text";
            copyAllTextToolStripMenuItem.Click += copyAllTextToolStripMenuItem_Click;
            // 
            // FrmDisplayOutput
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(769, 552);
            Controls.Add(splitContainer1);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "FrmDisplayOutput";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text = "Sample Form";
            splitContainer1.Panel1.ResumeLayout(false);
            splitContainer1.Panel2.ResumeLayout(false);
            splitContainer1.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)splitContainer1).EndInit();
            splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)pictureBox1).EndInit();
            contextMenuStrip1.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.RichTextBox rtbRTFContent;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem copySelectedTextToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem copyAllTextToolStripMenuItem;
    }
}