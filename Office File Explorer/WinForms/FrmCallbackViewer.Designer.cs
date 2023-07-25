namespace Office_File_Explorer.WinForms
{
    partial class FrmCallbackViewer
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
            rtbCallbacks = new System.Windows.Forms.RichTextBox();
            SuspendLayout();
            // 
            // rtbCallbacks
            // 
            rtbCallbacks.Location = new System.Drawing.Point(12, 12);
            rtbCallbacks.Name = "rtbCallbacks";
            rtbCallbacks.Size = new System.Drawing.Size(534, 315);
            rtbCallbacks.TabIndex = 0;
            rtbCallbacks.Text = "";
            // 
            // FrmCallbackViewer
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(556, 334);
            Controls.Add(rtbCallbacks);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "FrmCallbackViewer";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            Text = "Callback Viewer";
            ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.RichTextBox rtbCallbacks;
    }
}