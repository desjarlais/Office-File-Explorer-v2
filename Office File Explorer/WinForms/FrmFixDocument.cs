using DocumentFormat.OpenXml.Packaging;
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
                rdoFixEndnotesW.Enabled = true;
                rdoFixListTemplatesW.Enabled = true;
                rdoFixCommentsW.Enabled = true;
                rdoFixRevisionsW.Enabled = true;
                rdoFixHyperlinksW.Enabled = true;
                rdoFixCorruptTables.Enabled = true;
                rdoFixContentControlsW.Enabled = true;
                rdoTryAllFixesW.Enabled = true;
                rdoFixDataDescriptorW.Enabled = true;
                rdoFixMathAccentsW.Enabled = true;
                rdoFixListStyles.Enabled = true;
            }
            else if (type == Strings.oAppExcel)
            {
                rdoFixStrictX.Enabled = true;
            }
            else
            {
                rdoFixNotesPageSizeCustomP.Enabled = true;
                rdoFixNotesPageSizeP.Enabled = true;
                rdoResetBulletMargins.Enabled = true;
                rdoFixDataTags.Enabled = true;
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            corruptionChecked = Strings.wCancel;
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
                    featureFixed.Add(Strings.wBookmarks + Strings.wFixedWithSpace);
                }
            }

            if (rdoFixRevisionsW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wRevisions);
                if (WordFixes.FixRevisions(filePath) == true || WordFixes.FixDeleteRevision(filePath) == true)
                {
                    isFileFixed = true;
                    featureFixed.Add(Strings.wRevisions + Strings.wFixedWithSpace);
                }
            }

            if (rdoFixEndnotesW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wEndnotes);
                if (WordFixes.FixEndnotes(filePath) == true)
                {
                    isFileFixed = true;
                    featureFixed.Add(Strings.wEndnotes + Strings.wFixedWithSpace);
                }
            }

            if (rdoFixListTemplatesW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wListTemplates);
                if (WordFixes.FixListTemplates(filePath) == true)
                {
                    isFileFixed = true;
                    featureFixed.Add(Strings.wListTemplates + Strings.wFixedWithSpace);
                }
            }

            if (rdoFixCorruptTables.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wTableProps);
                if (WordFixes.FixTableGridProps(filePath) == true)
                {
                    isFileFixed = true;
                    featureFixed.Add(Strings.wTableProps + Strings.wFixedWithSpace);
                }

                if (WordFixes.FixTableRowCorruption(filePath) == true)
                {
                    isFileFixed = true;
                    featureFixed.Add(Strings.wTableProps + Strings.wFixedWithSpace);
                }

                if (WordFixes.FixCorruptTableCellTags(filePath) == true)
                {
                    isFileFixed = true;
                    featureFixed.Add(Strings.wTableCell + Strings.wFixedWithSpace);
                }

                if (WordFixes.FixGridSpan(filePath) == true)
                {
                    isFileFixed = true;
                    featureFixed.Add(Strings.wTableProps + Strings.wFixedWithSpace);
                }
            }

            if (rdoFixCommentsW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wComments);
                if (WordFixes.FixMissingCommentRefs(filePath) == true || WordFixes.FixShapeInComment(filePath) == true)
                {
                    isFileFixed = true;
                    featureFixed.Add(Strings.wComments + Strings.wFixedWithSpace);
                }

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
                    featureFixed.Add(Strings.wHyperlinks + Strings.wFixedWithSpace);
                }
            }

            if (rdoFixContentControlsW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wContentControls);
                if (WordFixes.FixContentControls(filePath) == true)
                {
                    isFileFixed = true;
                    featureFixed.Add(Strings.wContentControls + Strings.wFixedWithSpace);
                }
            }

            if (rdoFixMathAccentsW.Checked || tryAllFixes == true)
            {
                SetCorruptionChecked(Strings.wMathAccents);
                if (WordFixes.FixMathAccents(filePath) == true)
                {
                    isFileFixed = true;
                    featureFixed.Add(Strings.wMathAccents + Strings.wFixedWithSpace);
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

            if (rdoResetBulletMargins.Checked)
            {
                SetCorruptionChecked(Strings.pptResetBulletMargins);
                using (PresentationDocument document = PresentationDocument.Open(filePath, true))
                {
                    PowerPointFixes.ResetBulletMargins(document);
                    isFileFixed = true;
                    featureFixed.Add(Strings.pptResetBulletMargins);
                    FileUtilities.WriteToLog(Strings.fLogFilePath, filePath + Strings.wArrow + Strings.pptResetBulletMargins);
                }
            }

            if (rdoFixDataTags.Checked)
            {
                SetCorruptionChecked(Strings.pptCustDataTags);
                if (PowerPointFixes.FixMissingRelIds(filePath))
                {
                    isFileFixed = true;
                    featureFixed.Add(Strings.pptCustDataTags);
                    FileUtilities.WriteToLog(Strings.fLogFilePath, filePath + Strings.wArrow + Strings.pptCustDataTags);
                }
            }

            if (rdoFixNotesPageSizeP.Checked)
            {
                SetCorruptionChecked(Strings.pptNotesSizeReset);
                if (PowerPointFixes.ResetNotesPageSize(filePath))
                {
                    isFileFixed = true;
                    featureFixed.Add(Strings.pptNotesSizeReset);
                    FileUtilities.WriteToLog(Strings.fLogFilePath, filePath + Strings.wArrow + Strings.pptNotesSizeReset);
                }
            }

            if (rdoFixNotesPageSizeCustomP.Checked)
            {
                SetCorruptionChecked(Strings.pptNotesSizeReset);
                if (PowerPointFixes.CustomResetNotesPageSize(filePath))
                {
                    isFileFixed = true;
                    featureFixed.Add(Strings.pptNotesSizeReset);
                    FileUtilities.WriteToLog(Strings.fLogFilePath, filePath + Strings.wArrow + Strings.pptNotesSizeReset);
                }
            }

            if (rdoFixCorruptDrawingsXL.Checked)
            {
                SetCorruptionChecked("Corrupt VmlDrawing");
                if (Excel.RemoveCorruptClientDataObjects(filePath))
                {
                    isFileFixed = true;
                    featureFixed.Add("Corrupt VmlDrawing");
                    FileUtilities.WriteToLog(Strings.fLogFilePath, filePath + Strings.wArrow + "Corrupt VmlDrawing");
                }
            }

            Close();
        }
    }
}
