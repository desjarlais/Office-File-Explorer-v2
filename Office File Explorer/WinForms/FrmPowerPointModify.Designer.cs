
namespace Office_File_Explorer.WinForms
{
    partial class FrmPowerPointModify
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmPowerPointModify));
            groupBox1 = new System.Windows.Forms.GroupBox();
            rdoCustomNotesPageReset = new System.Windows.Forms.RadioButton();
            rdoResetNotesPageSize = new System.Windows.Forms.RadioButton();
            rdoDeleteComments = new System.Windows.Forms.RadioButton();
            rdoRemovePII = new System.Windows.Forms.RadioButton();
            rdoConvertPptmToPptx = new System.Windows.Forms.RadioButton();
            rdoMoveSlide = new System.Windows.Forms.RadioButton();
            BtnOk = new System.Windows.Forms.Button();
            BtnCancel = new System.Windows.Forms.Button();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(rdoCustomNotesPageReset);
            groupBox1.Controls.Add(rdoResetNotesPageSize);
            groupBox1.Controls.Add(rdoDeleteComments);
            groupBox1.Controls.Add(rdoRemovePII);
            groupBox1.Controls.Add(rdoConvertPptmToPptx);
            groupBox1.Controls.Add(rdoMoveSlide);
            groupBox1.Location = new System.Drawing.Point(12, 12);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new System.Drawing.Size(315, 181);
            groupBox1.TabIndex = 0;
            groupBox1.TabStop = false;
            groupBox1.Text = "Select Content To Modify";
            // 
            // rdoCustomNotesPageReset
            // 
            rdoCustomNotesPageReset.AutoSize = true;
            rdoCustomNotesPageReset.Location = new System.Drawing.Point(11, 148);
            rdoCustomNotesPageReset.Name = "rdoCustomNotesPageReset";
            rdoCustomNotesPageReset.Size = new System.Drawing.Size(184, 19);
            rdoCustomNotesPageReset.TabIndex = 7;
            rdoCustomNotesPageReset.TabStop = true;
            rdoCustomNotesPageReset.Text = "Custom Notes Page Size Reset";
            rdoCustomNotesPageReset.UseVisualStyleBackColor = true;
            // 
            // rdoResetNotesPageSize
            // 
            rdoResetNotesPageSize.AutoSize = true;
            rdoResetNotesPageSize.Location = new System.Drawing.Point(11, 123);
            rdoResetNotesPageSize.Name = "rdoResetNotesPageSize";
            rdoResetNotesPageSize.Size = new System.Drawing.Size(139, 19);
            rdoResetNotesPageSize.TabIndex = 6;
            rdoResetNotesPageSize.TabStop = true;
            rdoResetNotesPageSize.Text = "Reset Notes Page Size";
            rdoResetNotesPageSize.UseVisualStyleBackColor = true;
            // 
            // rdoDeleteComments
            // 
            rdoDeleteComments.AutoSize = true;
            rdoDeleteComments.Location = new System.Drawing.Point(11, 98);
            rdoDeleteComments.Name = "rdoDeleteComments";
            rdoDeleteComments.Size = new System.Drawing.Size(120, 19);
            rdoDeleteComments.TabIndex = 5;
            rdoDeleteComments.TabStop = true;
            rdoDeleteComments.Text = "Delete Comments";
            rdoDeleteComments.UseVisualStyleBackColor = true;
            // 
            // rdoRemovePII
            // 
            rdoRemovePII.AutoSize = true;
            rdoRemovePII.Location = new System.Drawing.Point(11, 72);
            rdoRemovePII.Name = "rdoRemovePII";
            rdoRemovePII.Size = new System.Drawing.Size(130, 19);
            rdoRemovePII.TabIndex = 2;
            rdoRemovePII.TabStop = true;
            rdoRemovePII.Text = "Remove PII On Save";
            rdoRemovePII.UseVisualStyleBackColor = true;
            // 
            // rdoConvertPptmToPptx
            // 
            rdoConvertPptmToPptx.AutoSize = true;
            rdoConvertPptmToPptx.Location = new System.Drawing.Point(11, 47);
            rdoConvertPptmToPptx.Name = "rdoConvertPptmToPptx";
            rdoConvertPptmToPptx.Size = new System.Drawing.Size(141, 19);
            rdoConvertPptmToPptx.TabIndex = 1;
            rdoConvertPptmToPptx.TabStop = true;
            rdoConvertPptmToPptx.Text = "Convert Pptm To Pptx";
            rdoConvertPptmToPptx.UseVisualStyleBackColor = true;
            // 
            // rdoMoveSlide
            // 
            rdoMoveSlide.AutoSize = true;
            rdoMoveSlide.Location = new System.Drawing.Point(11, 22);
            rdoMoveSlide.Name = "rdoMoveSlide";
            rdoMoveSlide.Size = new System.Drawing.Size(83, 19);
            rdoMoveSlide.TabIndex = 0;
            rdoMoveSlide.TabStop = true;
            rdoMoveSlide.Text = "Move Slide";
            rdoMoveSlide.UseVisualStyleBackColor = true;
            // 
            // BtnOk
            // 
            BtnOk.Location = new System.Drawing.Point(171, 199);
            BtnOk.Name = "BtnOk";
            BtnOk.Size = new System.Drawing.Size(75, 23);
            BtnOk.TabIndex = 3;
            BtnOk.Text = "OK";
            BtnOk.UseVisualStyleBackColor = true;
            BtnOk.Click += BtnOk_Click;
            // 
            // BtnCancel
            // 
            BtnCancel.Location = new System.Drawing.Point(252, 199);
            BtnCancel.Name = "BtnCancel";
            BtnCancel.Size = new System.Drawing.Size(75, 23);
            BtnCancel.TabIndex = 4;
            BtnCancel.Text = "Cancel";
            BtnCancel.UseVisualStyleBackColor = true;
            BtnCancel.Click += BtnCancel_Click;
            // 
            // FrmPowerPointModify
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(339, 232);
            Controls.Add(BtnOk);
            Controls.Add(BtnCancel);
            Controls.Add(groupBox1);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            Icon = (System.Drawing.Icon)resources.GetObject("$this.Icon");
            KeyPreview = true;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "FrmPowerPointModify";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            Text = "Modify PowerPoint Content";
            KeyDown += FrmPowerPointModify_KeyDown;
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rdoRemovePII;
        private System.Windows.Forms.RadioButton rdoConvertPptmToPptx;
        private System.Windows.Forms.RadioButton rdoMoveSlide;
        private System.Windows.Forms.Button BtnOk;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.RadioButton rdoDeleteComments;
        private System.Windows.Forms.RadioButton rdoCustomNotesPageReset;
        private System.Windows.Forms.RadioButton rdoResetNotesPageSize;
    }
}