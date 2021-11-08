using Office_File_Explorer.Helpers;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmWordModify : Form
    {
        public AppUtilities.WordModifyCmds wdModCmd = new AppUtilities.WordModifyCmds();

        public FrmWordModify()
        {
            InitializeComponent();
        }

        private void BtnOk_Click(object sender, System.EventArgs e)
        {
            if (rdoDelHF.Checked)
            {
                wdModCmd = AppUtilities.WordModifyCmds.DelHF;
            }

            if (rdoAcceptRevisions.Checked)
            {
                wdModCmd = AppUtilities.WordModifyCmds.AcceptRevisions;
            }

            if (rdoChangeDefTemp.Checked)
            {
                wdModCmd = AppUtilities.WordModifyCmds.ChangeDefaultTemplate;
            }

            if (rdoConvertDocmToDocx.Checked)
            {
                wdModCmd = AppUtilities.WordModifyCmds.ConvertDocmToDocx;
            }

            if (rdoDelComments.Checked)
            {
                wdModCmd = AppUtilities.WordModifyCmds.DelComments;
            }

            if (rdoDelEndnotes.Checked)
            {
                wdModCmd = AppUtilities.WordModifyCmds.DelEndnotes;
            }

            if (rdoDelFootnotes.Checked)
            {
                wdModCmd = AppUtilities.WordModifyCmds.DelFootnotes;
            }

            if (rdoDelHiddenText.Checked)
            {
                wdModCmd = AppUtilities.WordModifyCmds.DelHiddenTxt;
            }

            if (rdoDelOrhpanLT.Checked)
            {
                wdModCmd = AppUtilities.WordModifyCmds.DelOrphanLT;
            }

            if (rdoDelOrphanStyles.Checked)
            {
                wdModCmd = AppUtilities.WordModifyCmds.DelOrphanStyles;
            }

            if (rdoDelPgBreaks.Checked)
            {
                wdModCmd = AppUtilities.WordModifyCmds.DelPgBrk;
            }

            if (rdoSetPrint.Checked)
            {
                wdModCmd = AppUtilities.WordModifyCmds.SetPrintOrientation;
            }

            if (rdoRemovePII.Checked)
            {
                wdModCmd = AppUtilities.WordModifyCmds.RemovePII;
            }

            Close();
        }

        private void BtnCancel_Click(object sender, System.EventArgs e)
        {
            Close();
        }
    }
}
