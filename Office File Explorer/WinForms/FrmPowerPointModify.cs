using Office_File_Explorer.Helpers;

using System;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmPowerPointModify : Form
    {
        public AppUtilities.PowerPointModifyCmds pptModCmd = new AppUtilities.PowerPointModifyCmds();

        public FrmPowerPointModify()
        {
            InitializeComponent();
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            if (rdoMoveSlide.Checked)
            {
                pptModCmd = AppUtilities.PowerPointModifyCmds.MoveSlide;
            }

            if (rdoConvertPptmToPptx.Checked)
            {
                pptModCmd = AppUtilities.PowerPointModifyCmds.ConvertPptmToPptx;
            }

            if (rdoRemovePII.Checked)
            {
                pptModCmd = AppUtilities.PowerPointModifyCmds.RemovePIIOnSave;
            }

            if (rdoDeleteComments.Checked)
            {
                pptModCmd = AppUtilities.PowerPointModifyCmds.DelComments;
            }

            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
