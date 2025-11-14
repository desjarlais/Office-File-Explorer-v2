using Office_File_Explorer.Helpers;
using System;
using System.Windows.Forms;
using System.ComponentModel;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmBatchDeleteCustomProps : Form
    {
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
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
