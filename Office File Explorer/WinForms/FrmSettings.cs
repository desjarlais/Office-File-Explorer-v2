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
            Properties.Settings.Default.ClientID = tbxClientID.Text;
            Properties.Settings.Default.MySite = tbxSiteURL.Text;

            if (rdoSAX.Checked == true)
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
            if (e.KeyCode == Keys.Escape)
            {
                Close();
            }
        }
    }
}
