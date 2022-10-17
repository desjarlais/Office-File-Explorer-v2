using System.Collections.Generic;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmSPCustomXmlGuids : Form
    {
        public string newGuid = string.Empty;

        public FrmSPCustomXmlGuids(List<string> guidsFromFile)
        {
            InitializeComponent();

            foreach (string guid in guidsFromFile)
            {
                cbxSPGuids.Items.Add(guid);
            }
        }

        private void BtnOK_Click(object sender, System.EventArgs e)
        {
            newGuid = cbxSPGuids.SelectedItem.ToString();
            Close();
        }

        private void BtnCancel_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void FrmSPCustomXmlGuids_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { Close(); }
        }
    }
}
