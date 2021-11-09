using Office_File_Explorer.Helpers;
using System;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmFixDocument : Form
    {
        public bool isFileFixed = false;
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
            corruptionChecked = "Cancel";
            Close();
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            if (rdoFixBookmarksW.Checked)
            {
                corruptionChecked = Strings.wBookmarks;
                if (WordFixes.RemoveMissingBookmarkTags(filePath) == true || WordFixes.RemovePlainTextCcFromBookmark(filePath) == true)
                {
                    isFileFixed = true;
                }
            }

            if (rdoFixRevisionsW.Checked)
            {
                corruptionChecked = Strings.wRevisions;
                if (WordFixes.FixRevisions(filePath) == true || WordFixes.FixDeleteRevision(filePath) == true)
                {
                    isFileFixed = true;
                }
            }
            
            if (rdoFixEndnotesW.Checked)
            {
                corruptionChecked = Strings.wEndnotes;
                if (WordFixes.FixEndnotes(filePath) == true)
                {
                    isFileFixed = true;
                }
            }

            if (rdoFixListTemplatesW.Checked)
            {
                corruptionChecked = Strings.wListTemplates;
                if (WordFixes.FixListTemplates(filePath) == true)
                {
                    isFileFixed = true;
                }
            }

            if (rdoFixTablePropsW.Checked)
            {
                if (WordFixes.FixTableGridProps(filePath) == true)
                {
                    corruptionChecked = Strings.wTableProps;
                    isFileFixed = true;
                }
            }

            if (rdoFixCommentsW.Checked)
            {
                corruptionChecked = Strings.wComments;
                if (WordFixes.FixMissingCommentRefs(filePath) == true || WordFixes.FixShapeInComment(filePath) == true)
                {
                    isFileFixed = true;
                }
            }

            if (rdoFixCommentHyperlinksW.Checked)
            {
                corruptionChecked = Strings.wFieldCodes;
                if (WordFixes.FixCommentFieldCodes(filePath) == true)
                {
                    isFileFixed = true;
                }
            }

            if (rdoFixHyperlinksW.Checked)
            {
                corruptionChecked = Strings.wHyperlinks;
                if (WordFixes.FixHyperlinks(filePath) == true)
                {
                    isFileFixed = true;
                }
            }

            if (rdoFixContentControlsW.Checked)
            {
                corruptionChecked = Strings.wContentControls;
                if (WordFixes.FixContentControls(filePath) == true)
                {
                    isFileFixed = true;
                }
            }

            // isFileFixed should be set, now close the form
            Close();
        }
    }
}
