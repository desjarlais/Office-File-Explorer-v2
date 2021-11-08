using Office_File_Explorer.Helpers;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmChangeDefaultTemplate : Form
    {
        public FrmChangeDefaultTemplate(string templatePath)
        {
            InitializeComponent();
            tbOldPath.Text = templatePath;
        }

        public FrmChangeDefaultTemplate()
        {
            InitializeComponent();
            tbOldPath.Text = string.Empty;
        }

        private void BtnCancel_Click(object sender, System.EventArgs e)
        {
            if (Owner is FrmBatch f)
            {
                f.DefaultTemplate = "Cancel";
            }
            else if (Owner is FrmMain fm)
            {
                fm.DefaultTemplate = "Cancel";
            }
            Close();
        }

        private void BtnOk_Click(object sender, System.EventArgs e)
        {
            if (tbNewPath.Text.Length > 0)
            {
                if (Owner is FrmBatch f)
                {
                    if (tbNewPath.Text != "Normal")
                    {
                        f.DefaultTemplate = FileUtilities.ConvertFilePathToUri(tbNewPath.Text);
                    }
                    else
                    {
                        f.DefaultTemplate = "Normal";
                    }
                }
                else if (Owner is FrmMain fm)
                {
                    if (tbNewPath.Text != "Normal")
                    {
                        fm.DefaultTemplate = FileUtilities.ConvertFilePathToUri(tbNewPath.Text);
                    }
                    else
                    {
                        fm.DefaultTemplate = "Normal";
                    }
                }
            }

            Close();
        }
    }
}
