
namespace Office_File_Explorer.WinForms
{
    partial class FrmDeleteCustomProps
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmDeleteCustomProps));
            this.lbProps = new System.Windows.Forms.ListBox();
            this.BtnDeleteProp = new System.Windows.Forms.Button();
            this.BtnOk = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lbProps
            // 
            this.lbProps.FormattingEnabled = true;
            this.lbProps.ItemHeight = 15;
            this.lbProps.Location = new System.Drawing.Point(12, 12);
            this.lbProps.Name = "lbProps";
            this.lbProps.Size = new System.Drawing.Size(412, 349);
            this.lbProps.TabIndex = 0;
            // 
            // BtnDeleteProp
            // 
            this.BtnDeleteProp.Location = new System.Drawing.Point(12, 367);
            this.BtnDeleteProp.Name = "BtnDeleteProp";
            this.BtnDeleteProp.Size = new System.Drawing.Size(112, 23);
            this.BtnDeleteProp.TabIndex = 1;
            this.BtnDeleteProp.Text = "Delete Property";
            this.BtnDeleteProp.UseVisualStyleBackColor = true;
            this.BtnDeleteProp.Click += new System.EventHandler(this.BtnDeleteProp_Click);
            // 
            // BtnOk
            // 
            this.BtnOk.Location = new System.Drawing.Point(349, 367);
            this.BtnOk.Name = "BtnOk";
            this.BtnOk.Size = new System.Drawing.Size(75, 23);
            this.BtnOk.TabIndex = 2;
            this.BtnOk.Text = "Ok";
            this.BtnOk.UseVisualStyleBackColor = true;
            this.BtnOk.Click += new System.EventHandler(this.BtnOk_Click);
            // 
            // FrmDeleteCustomProps
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(436, 394);
            this.Controls.Add(this.BtnOk);
            this.Controls.Add(this.BtnDeleteProp);
            this.Controls.Add(this.lbProps);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmDeleteCustomProps";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Delete Custom Props";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox lbProps;
        private System.Windows.Forms.Button BtnDeleteProp;
        private System.Windows.Forms.Button BtnOk;
    }
}