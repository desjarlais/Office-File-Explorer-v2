
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rdoRemovePII = new System.Windows.Forms.RadioButton();
            this.rdoConvertPptmToPptx = new System.Windows.Forms.RadioButton();
            this.rdoMoveSlide = new System.Windows.Forms.RadioButton();
            this.BtnOk = new System.Windows.Forms.Button();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.rdoDeleteComments = new System.Windows.Forms.RadioButton();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rdoDeleteComments);
            this.groupBox1.Controls.Add(this.rdoRemovePII);
            this.groupBox1.Controls.Add(this.rdoConvertPptmToPptx);
            this.groupBox1.Controls.Add(this.rdoMoveSlide);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(315, 133);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Select Content To Modify";
            // 
            // rdoRemovePII
            // 
            this.rdoRemovePII.AutoSize = true;
            this.rdoRemovePII.Location = new System.Drawing.Point(11, 72);
            this.rdoRemovePII.Name = "rdoRemovePII";
            this.rdoRemovePII.Size = new System.Drawing.Size(130, 19);
            this.rdoRemovePII.TabIndex = 2;
            this.rdoRemovePII.TabStop = true;
            this.rdoRemovePII.Text = "Remove PII On Save";
            this.rdoRemovePII.UseVisualStyleBackColor = true;
            // 
            // rdoConvertPptmToPptx
            // 
            this.rdoConvertPptmToPptx.AutoSize = true;
            this.rdoConvertPptmToPptx.Location = new System.Drawing.Point(11, 47);
            this.rdoConvertPptmToPptx.Name = "rdoConvertPptmToPptx";
            this.rdoConvertPptmToPptx.Size = new System.Drawing.Size(141, 19);
            this.rdoConvertPptmToPptx.TabIndex = 1;
            this.rdoConvertPptmToPptx.TabStop = true;
            this.rdoConvertPptmToPptx.Text = "Convert Pptm To Pptx";
            this.rdoConvertPptmToPptx.UseVisualStyleBackColor = true;
            // 
            // rdoMoveSlide
            // 
            this.rdoMoveSlide.AutoSize = true;
            this.rdoMoveSlide.Location = new System.Drawing.Point(11, 22);
            this.rdoMoveSlide.Name = "rdoMoveSlide";
            this.rdoMoveSlide.Size = new System.Drawing.Size(83, 19);
            this.rdoMoveSlide.TabIndex = 0;
            this.rdoMoveSlide.TabStop = true;
            this.rdoMoveSlide.Text = "Move Slide";
            this.rdoMoveSlide.UseVisualStyleBackColor = true;
            // 
            // BtnOk
            // 
            this.BtnOk.Location = new System.Drawing.Point(171, 151);
            this.BtnOk.Name = "BtnOk";
            this.BtnOk.Size = new System.Drawing.Size(75, 23);
            this.BtnOk.TabIndex = 3;
            this.BtnOk.Text = "Ok";
            this.BtnOk.UseVisualStyleBackColor = true;
            this.BtnOk.Click += new System.EventHandler(this.BtnOk_Click);
            // 
            // BtnCancel
            // 
            this.BtnCancel.Location = new System.Drawing.Point(252, 151);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(75, 23);
            this.BtnCancel.TabIndex = 4;
            this.BtnCancel.Text = "Cancel";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // rdoDeleteComments
            // 
            this.rdoDeleteComments.AutoSize = true;
            this.rdoDeleteComments.Location = new System.Drawing.Point(11, 98);
            this.rdoDeleteComments.Name = "rdoDeleteComments";
            this.rdoDeleteComments.Size = new System.Drawing.Size(120, 19);
            this.rdoDeleteComments.TabIndex = 5;
            this.rdoDeleteComments.TabStop = true;
            this.rdoDeleteComments.Text = "Delete Comments";
            this.rdoDeleteComments.UseVisualStyleBackColor = true;
            // 
            // FrmPowerPointModify
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(339, 186);
            this.Controls.Add(this.BtnOk);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmPowerPointModify";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Modify PowerPoint Content";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rdoRemovePII;
        private System.Windows.Forms.RadioButton rdoConvertPptmToPptx;
        private System.Windows.Forms.RadioButton rdoMoveSlide;
        private System.Windows.Forms.Button BtnOk;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.RadioButton rdoDeleteComments;
    }
}