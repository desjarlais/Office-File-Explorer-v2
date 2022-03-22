using Office_File_Explorer.Helpers;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmPPTCommands : Form
    {
        public AppUtilities.PowerPointViewCmds pptCmds = new AppUtilities.PowerPointViewCmds();
        public AppUtilities.OfficeViewCmds offCmds = new AppUtilities.OfficeViewCmds();

        public FrmPPTCommands()
        {
            InitializeComponent();
        }

        private void BtnOk_Click(object sender, System.EventArgs e)
        {
            // check PowerPoint features
            if (ckbHyperlinks.Checked)
            {
                pptCmds |= AppUtilities.PowerPointViewCmds.Hyperlinks;
            }
            else
            {
                pptCmds &= ~AppUtilities.PowerPointViewCmds.Hyperlinks;
            }

            if (ckbSlideTitles.Checked)
            {
                pptCmds |= AppUtilities.PowerPointViewCmds.SlideTitles;
            }
            else
            {
                pptCmds &= ~AppUtilities.PowerPointViewCmds.SlideTitles;
            }

            if (ckbSlideTransitions.Checked)
            {
                pptCmds |= AppUtilities.PowerPointViewCmds.SlideTransitions;
            }
            else
            {
                pptCmds &= ~AppUtilities.PowerPointViewCmds.SlideTransitions;
            }

            if (ckbSlideText.Checked)
            {
                pptCmds |= AppUtilities.PowerPointViewCmds.SlideText;
            }
            else
            {
                pptCmds &= ~AppUtilities.PowerPointViewCmds.SlideText;
            }

            if (ckbComments.Checked)
            {
                pptCmds |= AppUtilities.PowerPointViewCmds.Comments;
            }
            else
            {
                pptCmds &= ~AppUtilities.PowerPointViewCmds.Comments;
            }

            if (ckbListFonts.Checked)
            {
                pptCmds |= AppUtilities.PowerPointViewCmds.Fonts;
            }
            else
            {
                pptCmds &= ~AppUtilities.PowerPointViewCmds.Fonts;
            }

            // check Office features
            if (ckbShapes.Checked)
            {
                offCmds |= AppUtilities.OfficeViewCmds.Shapes;
            }
            else
            {
                offCmds &= ~AppUtilities.OfficeViewCmds.Shapes;
            }

            if (ckbOleObjects.Checked)
            {
                offCmds |= AppUtilities.OfficeViewCmds.OleObjects;
            }
            else
            {
                offCmds &= ~AppUtilities.OfficeViewCmds.OleObjects;
            }

            if (ckbPackageParts.Checked)
            {
                offCmds |= AppUtilities.OfficeViewCmds.PackageParts;
            }
            else
            {
                offCmds &= ~AppUtilities.OfficeViewCmds.PackageParts;
            }

            Close();
        }

        private void BtnCancel_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void CkbSelectAll_CheckedChanged(object sender, System.EventArgs e)
        {
            if (ckbSelectAll.Checked)
            {
                ckbShapes.Checked = true;
                ckbHyperlinks.Checked = true;
                ckbComments.Checked = true;
                ckbOleObjects.Checked = true;
                ckbPackageParts.Checked = true;
                ckbSlideText.Checked = true;
                ckbSlideTitles.Checked = true;
                ckbSlideTransitions.Checked = true;
            }
            else
            {
                ckbShapes.Checked = false;
                ckbHyperlinks.Checked = false;
                ckbComments.Checked = false;
                ckbOleObjects.Checked = false;
                ckbPackageParts.Checked = false;
                ckbSlideText.Checked = false;
                ckbSlideTitles.Checked = false;
                ckbSlideTransitions.Checked = false;
            }
        }

        private void FrmPPTCommands_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                Close();
            }
        }
    }
}
