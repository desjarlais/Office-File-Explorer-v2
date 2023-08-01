using DocumentFormat.OpenXml.Packaging;
using Office_File_Explorer.Helpers;
using System;
using System.IO.Packaging;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmMoveSlide : Form
    {
        string fName;
        int slideCount;

        public FrmMoveSlide(string fPath)
        {
            InitializeComponent();
            fName = fPath;

            slideCount = PowerPoint.RetrieveNumberOfSlides(fPath);

            // for ui-only, use non zero based values
            for (int i = 0; i < slideCount; i++)
            {
                cboFrom.Items.Add(i + 1);
                cboTo.Items.Add(i + 1);
            }

            // since we have items to list, pre-select the first one
            cboFrom.SelectedIndex = 0;
            cboTo.SelectedIndex = 0;
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Open(fName, true))
            {
                PowerPoint.MoveSlide(presentationDocument, cboFrom.SelectedIndex, cboTo.SelectedIndex);
            }
            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void FrmMoveSlide_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { Close(); }
        }
    }
}
