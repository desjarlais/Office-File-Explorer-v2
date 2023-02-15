using System;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmSettings : Form
    {
        public FrmSettings()
        {
            InitializeComponent();

            // populate checkboxes from settings
            if (Properties.Settings.Default.RemoveFallback == true)
            {
                ckbRemoveFallbackTags.Checked = true;
            }

            if (Properties.Settings.Default.ListRsids == true)
            {
                ckbListRsids.Checked = true;
            }

            if (Properties.Settings.Default.FixGroupedShapes == true)
            {
                ckbFixGroupedShapes.Checked = true;
            }

            if (Properties.Settings.Default.ResetNotesMaster == true)
            {
                ckbResetNotes.Checked = true;
            }

            if (Properties.Settings.Default.DeleteCopiesOnExit == true)
            {
                ckbDeleteOnExit.Checked = true;
            }

            if (Properties.Settings.Default.ListCellValuesSax == true)
            {
                rdoSAX.Checked = true;
            }
            else
            {
                rdoDOM.Checked = true;
            }

            if (Properties.Settings.Default.CheckZipItemCorrupt == true)
            {
                ckbZipItemCorrupt.Checked = true;
            }

            if (Properties.Settings.Default.BackupOnOpen == true)
            {
                ckbBackupOnOpen.Checked = true;
            }

            if (Properties.Settings.Default.DeleteOnlyCommentBookmarks == true)
            {
                ckbDeleteOnlyCommentBookmarks.Checked = true;
            }

            if (Properties.Settings.Default.RemoveCustDataTags == true)
            {
                ckbRemoveCustDataTags.Checked = true;
            }

            if (Properties.Settings.Default.UseContentControlGuid == true)
            {
                rdoUseCCGuid.Checked = true;
            }
            else if (Properties.Settings.Default.UseSharePointGuid == true)
            {
                rdoUseSPGuid.Checked = true;
            }
            else
            {
                rdoUserSelectedCC.Checked = true;
            }
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.RemoveFallback = ckbRemoveFallbackTags.Checked;
            Properties.Settings.Default.ListRsids = ckbListRsids.Checked;
            Properties.Settings.Default.FixGroupedShapes = ckbFixGroupedShapes.Checked;
            Properties.Settings.Default.ResetNotesMaster = ckbResetNotes.Checked;
            Properties.Settings.Default.DeleteCopiesOnExit = ckbDeleteOnExit.Checked;
            Properties.Settings.Default.CheckZipItemCorrupt = ckbZipItemCorrupt.Checked;
            Properties.Settings.Default.BackupOnOpen = ckbBackupOnOpen.Checked;
            Properties.Settings.Default.DeleteOnlyCommentBookmarks = ckbDeleteOnlyCommentBookmarks.Checked;
            Properties.Settings.Default.RemoveCustDataTags = ckbRemoveCustDataTags.Checked;

            if (rdoUseCCGuid.Checked)
            {
                Properties.Settings.Default.UseContentControlGuid = true;
            }
            else
            {
                Properties.Settings.Default.UseContentControlGuid = false;
            }
            
            if (rdoUserSelectedCC.Checked)
            {
                Properties.Settings.Default.UseUserSelectedCCGuid = true;
            }
            else
            {
                Properties.Settings.Default.UseUserSelectedCCGuid = false;
            }

            if (rdoUseSPGuid.Checked)
            {
                Properties.Settings.Default.UseSharePointGuid = true;
            }
            else
            {
                Properties.Settings.Default.UseSharePointGuid = false;
            }

            if (rdoSAX.Checked)
            {
                Properties.Settings.Default.ListCellValuesSax = true;
            }
            else
            {
                Properties.Settings.Default.ListCellValuesSax = false;
            }

            Properties.Settings.Default.Save();
            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void FrmSettings_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { Close(); }
        }
    }
}
