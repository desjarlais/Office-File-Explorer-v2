using Office_File_Explorer.Helpers;
using System;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmFixDocument : Form
    {
        public bool isFileFixed = false;
        public bool tryAllFixes = false;
        public string corruptionChecked = string.Empty;
        public string filePath, fileType;

        public FrmFixDocument(string fPath, string fType)
        {
            InitializeComponent();
            filePath = fPath;
            fileType = fType;

            EnableUI(fileType);
        }

        public void EnableUI(string type)
        {
            if (type == Strings.oAppWord)
            {
                rdoFixBookmarksW.Enabled = true;
                rdoFixCommentHyperlinksW.Enabled = true;
                rdoFixEndnotesW.Enabled = true;
                rdoFixListTemplatesW.Enabled = true;
                rdoFixCommentsW.Enabled = true;
                rdoFixRevisionsW.Enabled = true;
                rdoFixHyperlinksW.Enabled = true;
                rdoFixTablePropsW.Enabled = true;
                rdoFixContentControlsW.Enabled = true;
                rdoTryAllFixesW.Enabled = true;
                rdoFixDataDescriptorW.Enabled = true;
            }
            else if (type == Strings.oAppExcel)
            {
                rdoFixStrictX.Enabled = true;
            }
            else
            {
                rdoFixNotesPageSizeCustomP.Enabled = true;
                rdoFixNotesPageSizeP.Enabled = true;
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        public void SetCorruptionChecked(string fixType)
        {
            if (tryAllFixes == true)
            {
                corruptionChecked = Strings.wAllFixes;
            }
            else
            {
                corruptionChecked = fixType;
            }
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            if (rdoTryAllFixesW.Checked)
            {
                tryAllFixes = true;
            }

            if (rdoFixBookmarksW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wBookmarks);
                if (WordFixes.RemoveMissingBookmarkTags(filePath) == true || WordFixes.RemovePlainTextCcFromBookmark(filePath) == true)
                {
                    isFileFixed = true;
                }
            }

            if (rdoFixRevisionsW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wRevisions);
                if (WordFixes.FixRevisions(filePath) == true || WordFixes.FixDeleteRevision(filePath) == true)
                {
                    isFileFixed = true;
                }
            }
            
            if (rdoFixEndnotesW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wEndnotes);
                if (WordFixes.FixEndnotes(filePath) == true)
                {
                    isFileFixed = true;
                }
            }

            if (rdoFixListTemplatesW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wListTemplates);
                if (WordFixes.FixListTemplates(filePath) == true)
                {
                    isFileFixed = true;
                }
            }

            if (rdoFixTablePropsW.Checked || tryAllFixes == true)
            {
                if (WordFixes.FixTableGridProps(filePath) == true)
                {
                    SetCorruptionChecked(Strings.wTableProps);
                    isFileFixed = true;
                }
            }

            if (rdoFixCommentsW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wComments);
                if (WordFixes.FixMissingCommentRefs(filePath) == true || WordFixes.FixShapeInComment(filePath) == true)
                {
                    isFileFixed = true;
                }
            }

            if (rdoFixCommentHyperlinksW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wFieldCodes);
                if (WordFixes.FixCommentFieldCodes(filePath) == true)
                {
                    isFileFixed = true;
                }
            }

            if (rdoFixHyperlinksW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wHyperlinks);
                if (WordFixes.FixHyperlinks(filePath) == true)
                {
                    isFileFixed = true;
                }
            }

            if (rdoFixContentControlsW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wContentControls);
                if (WordFixes.FixContentControls(filePath) == true)
                {
                    isFileFixed = true;
                }
            }

            if (rdoFixMathAccentsW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wMathAccents);
                if (WordFixes.FixMathAccents(filePath) == true)
                {
                    isFileFixed = true;
                }
            }

            if (rdoFixDataDescriptorW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wZipItem);
                if (WordFixes.FixDataDescriptor(filePath) == true)
                {
                    isFileFixed = true;
                }
            }

            // isFileFixed should be set, now close the form
            Close();
        }
    }
}
