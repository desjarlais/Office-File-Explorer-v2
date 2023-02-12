using Office_File_Explorer.Helpers;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmFixDocument : Form
    {
        public bool isFileFixed = false;
        public bool tryAllFixes = false;
        public string corruptionChecked = string.Empty;
        public string filePath, fileType;
        public List<string> featureFixed = new List<string>();

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
                rdoFixMathAccentsW.Enabled = true;
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

        private void FrmFixDocument_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { Close(); }
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            if (rdoTryAllFixesW.Checked)
            {
                tryAllFixes = true;
            }

            if (rdoFixListStyles.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wListStyles);
                if (WordFixes.FixListStyles(filePath))
                {
                    isFileFixed = true;
                }
            }

            if (rdoFixBookmarksW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wBookmarks);
                if (WordFixes.RemoveMissingBookmarkTags(filePath))
                {
                    isFileFixed = true;
                }

                if (WordFixes.RemovePlainTextCcFromBookmark(filePath))
                {
                    isFileFixed = true;
                }
                
                if (WordFixes.FixBookmarkTagInSdtContent(filePath))
                {
                    isFileFixed = true;
                }

                if (isFileFixed)
                {
                    featureFixed.Add("Bookmarks Fixed");
                }
            }

            if (rdoFixRevisionsW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wRevisions);
                if (WordFixes.FixRevisions(filePath) == true || WordFixes.FixDeleteRevision(filePath) == true)
                {
                    isFileFixed = true;
                    featureFixed.Add("Revisions Fixed");
                }
            }
            
            if (rdoFixEndnotesW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wEndnotes);
                if (WordFixes.FixEndnotes(filePath) == true)
                {
                    isFileFixed = true;
                    featureFixed.Add("Endnotes Fixed");
                }
            }

            if (rdoFixListTemplatesW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wListTemplates);
                if (WordFixes.FixListTemplates(filePath) == true)
                {
                    isFileFixed = true;
                    featureFixed.Add("List Templates Fixed");
                }
            }

            if (rdoFixTablePropsW.Checked || tryAllFixes == true)
            {
                if (WordFixes.FixTableGridProps(filePath) == true || WordFixes.FixTableRowCorruption(filePath) == true)
                {
                    SetCorruptionChecked(Strings.wTableProps);
                    isFileFixed = true;
                    featureFixed.Add("Table Corruption Fixed");
                }
            }

            if (rdoFixCommentsW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wComments);
                if (WordFixes.FixMissingCommentRefs(filePath) == true || WordFixes.FixShapeInComment(filePath) == true)
                {
                    isFileFixed = true;
                    featureFixed.Add("Comments Fixed");
                }
            }

            if (rdoFixCommentHyperlinksW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wFieldCodes);
                if (WordFixes.FixCommentFieldCodes(filePath) == true)
                {
                    isFileFixed = true;
                    featureFixed.Add("Comment Hyperlinks Fixed");
                }
            }

            if (rdoFixHyperlinksW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wHyperlinks);
                if (WordFixes.FixHyperlinks(filePath) == true)
                {
                    isFileFixed = true;
                    featureFixed.Add("Hyperlinks Fixed");
                }
            }

            if (rdoFixContentControlsW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wContentControls);
                if (WordFixes.FixContentControls(filePath) == true)
                {
                    isFileFixed = true;
                    featureFixed.Add("Content Controls Fixed");
                }
            }

            if (rdoFixMathAccentsW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wMathAccents);
                if (WordFixes.FixMathAccents(filePath) == true)
                {
                    isFileFixed = true;
                    featureFixed.Add("Math Accents Fixed");
                }
            }

            if (rdoFixDataDescriptorW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wZipItem);
                if (WordFixes.FixDataDescriptor(filePath) == true)
                {
                    isFileFixed = true;
                    featureFixed.Add("Corrupt Zip Item Fixed");
                }
            }

            if (rdoFixTableCellTags.Checked || tryAllFixes == true) 
            {
                SetCorruptionChecked(Strings.wTableCell);
                if (WordFixes.FixCorruptTableCellTags(filePath) == true)
                {
                    isFileFixed = true;
                    featureFixed.Add("Table Cell Corruption Fixed");
                }
            }

            Close();
        }
    }
}
