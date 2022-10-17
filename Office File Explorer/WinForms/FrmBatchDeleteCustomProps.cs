using Office_File_Explorer.Helpers;
using System;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmBatchDeleteCustomProps : Form
    {
        public string PropName { get; set; }

        public FrmBatchDeleteCustomProps()
        {
            InitializeComponent();
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            PropName = tbPropName.Text;
            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            PropName = Strings.wCancel;
            Close();
        }

        private void FrmBatchDeleteCustomProps_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { Close(); }
        }
    }
}
