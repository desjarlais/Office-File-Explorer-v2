using Office_File_Explorer.Helpers;
using System;
using System.IO.Packaging;
using System.Linq;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmExcelDelLink : Form
    {
        string fPath = string.Empty;
        public bool fHasLinks = false;

        public FrmExcelDelLink(Package pkg, string path)
        {
            InitializeComponent();
            fPath = path;

            if (Excel.GetLinks(pkg, false).Any())
            {
                foreach (string s in Excel.GetLinks(pkg, false))
                {
                    cboLinks.Items.Add(s);
                    fHasLinks = true;
                }

                cboLinks.SelectedIndex = 0;
            }
            else
            {
                MessageBox.Show("Document does not contain any links.");
                Close();
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            Excel.RemoveLink(fPath, cboLinks.Text);
            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void BtnCancel_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { Close(); }
        }
    }
}
