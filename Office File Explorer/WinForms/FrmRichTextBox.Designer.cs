namespace Office_File_Explorer.WinForms
{
    partial class FrmRichTextBox
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
            rtbOutput = new System.Windows.Forms.RichTextBox();
            SuspendLayout();
            // 
            // rtbOutput
            // 
            rtbOutput.Dock = System.Windows.Forms.DockStyle.Fill;
            rtbOutput.Location = new System.Drawing.Point(0, 0);
            rtbOutput.Name = "rtbOutput";
            rtbOutput.Size = new System.Drawing.Size(782, 556);
            rtbOutput.TabIndex = 0;
            rtbOutput.Text = "";
            // 
            // FrmRichTextBox
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(782, 556);
            Controls.Add(rtbOutput);
            Name = "FrmRichTextBox";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            Text = "RTF Content";
            ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.RichTextBox rtbOutput;
    }
}