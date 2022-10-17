using System;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmSearchReplace : Form
    {
        public FrmSearchReplace()
        {
            InitializeComponent();
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            if (Owner is FrmMain f)
            {
                f.FindTextProperty = tbFind.Text;
                f.ReplaceTextProperty = tbReplace.Text;
            }
            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void FrmSearchReplace_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { Close(); }
        }
    }
}
