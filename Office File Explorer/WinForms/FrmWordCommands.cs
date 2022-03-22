using DocumentFormat.OpenXml.Packaging;
using Office_File_Explorer.Helpers;
using System;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Office2013.Word;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmWordCommands : Form
    {
        public AppUtilities.WordViewCmds wdCmds = new AppUtilities.WordViewCmds();
        public AppUtilities.OfficeViewCmds offCmds = new AppUtilities.OfficeViewCmds();
        string filePath;

        public FrmWordCommands(string fPath)
        {
            InitializeComponent();
            filePath = fPath;
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            // check Word features
            if (ckbContentControls.Checked)
            {
                wdCmds |= AppUtilities.WordViewCmds.ContentControls;
            }
            else
            {
                wdCmds &= ~AppUtilities.WordViewCmds.ContentControls;
            }

            if (ckbStyles.Checked)
            {
                wdCmds |= AppUtilities.WordViewCmds.Styles;
            }
            else
            {
                wdCmds &= ~AppUtilities.WordViewCmds.Styles;
            }

            if (ckbHyperlinks.Checked)
            {
                wdCmds |= AppUtilities.WordViewCmds.Hyperlinks;
            }
            else
            {
                wdCmds &= ~AppUtilities.WordViewCmds.Hyperlinks;
            }

            if (ckbListTemplates.Checked)
            {
                wdCmds |= AppUtilities.WordViewCmds.ListTemplates;
            }
            else
            {
                wdCmds &= ~AppUtilities.WordViewCmds.ListTemplates;
            }

            if (ckbFonts.Checked)
            {
                wdCmds |= AppUtilities.WordViewCmds.Fonts;
            }
            else
            {
                wdCmds &= ~AppUtilities.WordViewCmds.Fonts;
            }

            if (ckbFootnotes.Checked)
            {
                wdCmds |= AppUtilities.WordViewCmds.Footnotes;
            }
            else
            {
                wdCmds &= ~AppUtilities.WordViewCmds.Footnotes;
            }

            if (ckbEndnotes.Checked)
            {
                wdCmds |= AppUtilities.WordViewCmds.Endnotes;
            }
            else
            {
                wdCmds &= ~AppUtilities.WordViewCmds.Endnotes;
            }

            if (ckbDocProps.Checked)
            {
                wdCmds |= AppUtilities.WordViewCmds.DocumentProperties;
            }
            else
            {
                wdCmds &= ~AppUtilities.WordViewCmds.DocumentProperties;
            }

            if (ckbBookmarks.Checked)
            {
                wdCmds |= AppUtilities.WordViewCmds.Bookmarks;
            }
            else
            {
                wdCmds &= ~AppUtilities.WordViewCmds.Bookmarks;
            }

            if (ckbComments.Checked)
            {
                wdCmds |= AppUtilities.WordViewCmds.Comments;
            }
            else
            {
                wdCmds &= ~AppUtilities.WordViewCmds.Comments;
            }

            if (ckbFieldCodes.Checked)
            {
                wdCmds |= AppUtilities.WordViewCmds.FieldCodes;
            }
            else
            {
                wdCmds &= ~AppUtilities.WordViewCmds.FieldCodes;
            }

            if (ckbTables.Checked)
            {
                wdCmds |= AppUtilities.WordViewCmds.Tables;
            }
            else
            {
                wdCmds &= ~AppUtilities.WordViewCmds.Tables;
            }

            // check Office features
            if (ckbShapes.Checked)
            {
                offCmds |= AppUtilities.OfficeViewCmds.Shapes;
            }
            else
            {
                offCmds &= ~AppUtilities.OfficeViewCmds.Shapes;
            }

            if (ckbOleObjects.Checked)
            {
                offCmds |= AppUtilities.OfficeViewCmds.OleObjects;
            }
            else
            {
                offCmds &= ~AppUtilities.OfficeViewCmds.OleObjects;
            }

            if (ckbPackageParts.Checked)
            {
                offCmds |= AppUtilities.OfficeViewCmds.PackageParts;
            }
            else
            {
                offCmds &= ~AppUtilities.OfficeViewCmds.PackageParts;
            }

            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void CkbRevisions_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                if (cbAuthors.Enabled == true)
                {
                    cbAuthors.Enabled = false;
                    lbRevisions.Enabled = false;
                    lbRevisions.Items.Clear();
                }
                else
                {
                    cbAuthors.Enabled = true;
                    lbRevisions.Enabled = true;

                    using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
                    {
                        int count = 0;

                        // check the peoplepart and list those authors
                        WordprocessingPeoplePart peoplePart = doc.MainDocumentPart.WordprocessingPeoplePart;
                        if (peoplePart != null)
                        {
                            foreach (Person person in peoplePart.People)
                            {
                                count++;
                                PresenceInfo pi = person.PresenceInfo;
                                cbAuthors.Items.Add(person.Author.Value);
                            }
                        }

                        List<string> tempAuthors = Word.GetAllAuthors(doc.MainDocumentPart.Document);

                        // sometimes there are authors in a file but they don't exist in people.xml
                        if (tempAuthors.Count > 0)
                        {
                            // if the people part count is the same as GetAllAuthors, they must be the same authors
                            if (count == tempAuthors.Count)
                            {
                                return;
                            }

                            // if the count is not the same, display those authors
                            foreach (string s in tempAuthors)
                            {
                                count++;
                                cbAuthors.Items.Add(s);
                                //lbRevisions.Items.Add(count + ". User Name = " + s);
                            }
                        }

                        // if the count is 0 at this point, no authors exist
                        if (count == 0)
                        {
                            cbAuthors.Enabled = false;
                            lbRevisions.Enabled = false;
                        }
                        else
                        {
                            cbAuthors.SelectedIndex = 0;
                            cbAuthors.Items.Add("* All Authors *");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void CbAuthors_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                int revCount = 0;

                List<string> authorList = new List<string>();

                using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, false))
                {
                    Document doc = document.MainDocumentPart.Document;
                    var paragraphChanged = doc.Descendants<ParagraphPropertiesChange>().ToList();
                    var runChanged = doc.Descendants<RunPropertiesChange>().ToList();
                    var deleted = doc.Descendants<DeletedRun>().ToList();
                    var deletedParagraph = doc.Descendants<Deleted>().ToList();
                    var inserted = doc.Descendants<InsertedRun>().ToList();

                    lbRevisions.Items.Clear();

                    if (cbAuthors.SelectedItem.ToString() == "* All Authors *")
                    {
                        foreach (string s in cbAuthors.Items)
                        {
                            if (s == "* All Authors *")
                            {
                                continue;
                            }
                            else
                            {
                                var tempParagraphChanged = paragraphChanged.Where(item => item.Author == s).ToList();
                                var tempRunChanged = runChanged.Where(item => item.Author == s).ToList();
                                var tempDeleted = deleted.Where(item => item.Author == s).ToList();
                                var tempInserted = inserted.Where(item => item.Author == s).ToList();
                                var tempDeletedParagraph = deletedParagraph.Where(item => item.Author == s).ToList();

                                if ((tempParagraphChanged.Count + tempRunChanged.Count + tempDeleted.Count + tempInserted.Count + tempDeletedParagraph.Count) == 0)
                                {
                                    lbRevisions.Items.Add(s + " has no changes.");
                                    continue;
                                }

                                foreach (var item in tempParagraphChanged)
                                {
                                    revCount++;
                                    lbRevisions.Items.Add(revCount + Strings.wPeriod + s + " : Paragraph Changed ");
                                }

                                foreach (var item in tempDeletedParagraph)
                                {
                                    revCount++;
                                    lbRevisions.Items.Add(revCount + Strings.wPeriod + s + " : Paragraph Deleted ");
                                }

                                foreach (var item in tempRunChanged)
                                {
                                    revCount++;
                                    lbRevisions.Items.Add(revCount + Strings.wPeriod + s + " :  Run Changed = " + item.InnerText);
                                }

                                foreach (var item in tempDeleted)
                                {
                                    revCount++;
                                    lbRevisions.Items.Add(revCount + Strings.wPeriod + s + " :  Deletion = " + item.InnerText);
                                }

                                foreach (var item in tempInserted)
                                {
                                    if (item.Parent != null)
                                    {
                                        var textRuns = item.Elements<Run>().ToList();
                                        var parent = item.Parent;

                                        foreach (var textRun in textRuns)
                                        {
                                            revCount++;
                                            lbRevisions.Items.Add(revCount + Strings.wPeriod + s + " :  Insertion = " + textRun.InnerText);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        // list the selected authors revisions
                        if (!string.IsNullOrEmpty(cbAuthors.SelectedItem.ToString()))
                        {
                            paragraphChanged = paragraphChanged.Where(item => item.Author == cbAuthors.SelectedItem.ToString()).ToList();
                            runChanged = runChanged.Where(item => item.Author == cbAuthors.SelectedItem.ToString()).ToList();
                            deleted = deleted.Where(item => item.Author == cbAuthors.SelectedItem.ToString()).ToList();
                            inserted = inserted.Where(item => item.Author == cbAuthors.SelectedItem.ToString()).ToList();
                            deletedParagraph = deletedParagraph.Where(item => item.Author == cbAuthors.SelectedItem.ToString()).ToList();

                            if ((paragraphChanged.Count + runChanged.Count + deleted.Count + inserted.Count + deletedParagraph.Count) == 0)
                            {
                                lbRevisions.Items.Add("* Author has no changes *");
                                return;
                            }
                        }
                        else
                        {
                            lbRevisions.Items.Add("** There are no revisions in this document **");
                            return;
                        }

                        foreach (var item in paragraphChanged)
                        {
                            revCount++;
                            lbRevisions.Items.Add(revCount + ". Paragraph Changed ");
                        }

                        foreach (var item in deletedParagraph)
                        {
                            revCount++;
                            lbRevisions.Items.Add(revCount + ". Paragraph Deleted ");
                        }

                        foreach (var item in runChanged)
                        {
                            revCount++;
                            lbRevisions.Items.Add(revCount + ". Run Changed = " + item.InnerText);
                        }

                        foreach (var item in deleted)
                        {
                            revCount++;
                            lbRevisions.Items.Add(revCount + ". Deletion = " + item.InnerText);
                        }

                        foreach (var item in inserted)
                        {
                            if (item.Parent != null)
                            {
                                var textRuns = item.Elements<Run>().ToList();
                                var parent = item.Parent;

                                foreach (var textRun in textRuns)
                                {
                                    revCount++;
                                    lbRevisions.Items.Add(revCount + ". Insertion = " + textRun.InnerText);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void CkbSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            if (ckbSelectAll.Checked)
            {
                ckbShapes.Checked =  true;
                ckbStyles.Checked = true;
                ckbPackageParts.Checked = true;
                ckbOleObjects.Checked = true;
                ckbListTemplates.Checked = true;
                ckbHyperlinks.Checked = true;
                ckbFootnotes.Checked = true;
                ckbFonts.Checked = true;
                ckbFieldCodes.Checked = true;
                ckbEndnotes.Checked = true;
                ckbDocProps.Checked = true;
                ckbContentControls.Checked = true;
                ckbComments.Checked = true;
                ckbBookmarks.Checked = true;
                ckbTables.Checked = true;
            }
            else
            {
                ckbShapes.Checked = false;
                ckbStyles.Checked = false;
                ckbPackageParts.Checked = false;
                ckbOleObjects.Checked = false;
                ckbListTemplates.Checked = false;
                ckbHyperlinks.Checked = false;
                ckbFootnotes.Checked = false;
                ckbFonts.Checked = false;
                ckbFieldCodes.Checked = false;
                ckbEndnotes.Checked = false;
                ckbDocProps.Checked = false;
                ckbContentControls.Checked = false;
                ckbComments.Checked = false;
                ckbBookmarks.Checked = false;
                ckbTables.Checked = false;
            }
        }

        private void FrmWordCommands_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                Close();
            }
        }
    }
}
