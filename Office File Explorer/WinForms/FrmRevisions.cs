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
    public partial class FrmRevisions : Form
    {
        string filePath;

        public FrmRevisions(string fPath)
        {
            InitializeComponent();
            filePath = fPath;
            LoadRevisions();

            if (cbAuthors.SelectedIndex == -1)
            {
                MessageBox.Show("Document does not contain revisions.", "Document Revisions", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                cbAuthors.SelectedIndex = 0;
                LoadAuthorChanges();
            }
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        public void LoadRevisions()
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
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
                        if (count != tempAuthors.Count)
                        {
                            foreach (string s in tempAuthors)
                            {
                                count++;
                                bool authorFound = false;

                                foreach (string a in cbAuthors.Items)
                                {
                                    if (a == s)
                                    {
                                        authorFound = true;
                                    }
                                }

                                if (authorFound == false)
                                {
                                    cbAuthors.Items.Add(s);
                                }
                            }
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
                        cbAuthors.Items.Add(Strings.wAllAuthors);
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

        public void LoadAuthorChanges()
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                int revCount = 0;
                int revCountHeader = 0;
                int revCountFooter = 0;

                List<string> authorList = new List<string>();

                using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, false))
                {
                    Document doc = document.MainDocumentPart.Document;

                    List<ParagraphPropertiesChange> paragraphChanged = doc.Descendants<ParagraphPropertiesChange>().ToList();
                    List<RunPropertiesChange> runChanged = doc.Descendants<RunPropertiesChange>().ToList();
                    List<DeletedRun> deleted = doc.Descendants<DeletedRun>().ToList();
                    List<Deleted> deletedParagraph = doc.Descendants<Deleted>().ToList();
                    List<InsertedRun> inserted = doc.Descendants<InsertedRun>().ToList();

                    List<ParagraphPropertiesChange> hdrParagraphChanged = null;
                    List<RunPropertiesChange> hdrRunChanged = null;
                    List<DeletedRun> hdrDeleted = null;
                    List<Deleted> hdrDeletedParagraph = null;
                    List<InsertedRun> hdrInserted = null;

                    List<ParagraphPropertiesChange> ftrParagraphChanged = null;
                    List<RunPropertiesChange> ftrRunChanged = null;
                    List<DeletedRun> ftrDeleted = null;
                    List<Deleted> ftrDeletedParagraph = null;
                    List<InsertedRun> ftrInserted = null;

                    foreach (HeaderPart hp in doc.MainDocumentPart.HeaderParts)
                    {
                        hdrParagraphChanged = hp.Header.Descendants<ParagraphPropertiesChange>().ToList();
                        hdrRunChanged = hp.Header.Descendants<RunPropertiesChange>().ToList();
                        hdrDeleted = hp.Header.Descendants<DeletedRun>().ToList();
                        hdrDeletedParagraph = hp.Header.Descendants<Deleted>().ToList();
                        hdrInserted = hp.Header.Descendants<InsertedRun>().ToList();
                        revCountHeader = hdrParagraphChanged.Count + hdrRunChanged.Count + hdrDeleted.Count + hdrDeletedParagraph.Count + hdrInserted.Count;
                    }

                    foreach (FooterPart fp in doc.MainDocumentPart.FooterParts)
                    {
                        ftrParagraphChanged = fp.Footer.Descendants<ParagraphPropertiesChange>().ToList();
                        ftrRunChanged = fp.Footer.Descendants<RunPropertiesChange>().ToList();
                        ftrDeleted = fp.Footer.Descendants<DeletedRun>().ToList();
                        ftrDeletedParagraph = fp.Footer.Descendants<Deleted>().ToList();
                        ftrInserted = fp.Footer.Descendants<InsertedRun>().ToList();
                        revCountFooter = ftrParagraphChanged.Count + ftrRunChanged.Count + ftrDeleted.Count + ftrDeletedParagraph.Count + ftrInserted.Count;
                    }

                    lbRevisions.Items.Clear();

                    if (cbAuthors.SelectedItem.ToString() == Strings.wAllAuthors)
                    {
                        int totalRevCount = 0;
                        foreach (string s in cbAuthors.Items)
                        {
                            if (s == Strings.wAllAuthors)
                            {
                                continue;
                            }
                            else
                            {
                                // main changes
                                var tempParagraphChanged = paragraphChanged.Where(item => item.Author == s).ToList();
                                var tempRunChanged = runChanged.Where(item => item.Author == s).ToList();
                                var tempDeleted = deleted.Where(item => item.Author == s).ToList();
                                var tempInserted = inserted.Where(item => item.Author == s).ToList();
                                var tempDeletedParagraph = deletedParagraph.Where(item => item.Author == s).ToList();

                                //revCount = tempParagraphChanged.Count + tempRunChanged.Count + tempDeleted.Count + tempInserted.Count + tempDeletedParagraph.Count;
                                revCount = 0;
                                revCountHeader = 0;
                                revCountFooter = 0;

                                // look through the main changes
                                foreach (var item in tempParagraphChanged)
                                {
                                    revCount++;
                                    totalRevCount++;
                                    lbRevisions.Items.Add(totalRevCount + Strings.wPeriod + s + Strings.wParaFormatChange);
                                }

                                foreach (var item in tempDeletedParagraph)
                                {
                                    revCount++;
                                    totalRevCount++;
                                    lbRevisions.Items.Add(totalRevCount + Strings.wPeriod + s + Strings.wParaDeleted);
                                }

                                foreach (var item in tempRunChanged)
                                {
                                    revCount++;
                                    totalRevCount++;
                                    lbRevisions.Items.Add(totalRevCount + Strings.wPeriod + s + Strings.wRunFormatChange);
                                }

                                foreach (var item in tempDeleted)
                                {
                                    revCount++;
                                    totalRevCount++;
                                    lbRevisions.Items.Add(totalRevCount + Strings.wPeriod + s + Strings.wDeletion + item.InnerText);
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
                                            totalRevCount++;
                                            lbRevisions.Items.Add(totalRevCount + Strings.wPeriod + s + Strings.wInsertion + textRun.InnerText);
                                        }
                                    }
                                }

                                if (revCountHeader > 0)
                                {
                                    var tempHdrParagraphChanged = hdrParagraphChanged.Where(item => item.Author == s).ToList();
                                    var tempHdrRunChanged = hdrRunChanged.Where(item => item.Author == s).ToList();
                                    var tempHdrDeleted = hdrDeleted.Where(item => item.Author == s).ToList();
                                    var tempHdrInserted = hdrInserted.Where(item => item.Author == s).ToList();
                                    var tempHdrDeletedParagraph = hdrDeletedParagraph.Where(item => item.Author == s).ToList();

                                    revCountHeader = tempHdrParagraphChanged.Count + tempHdrRunChanged.Count + tempHdrDeleted.Count + tempHdrInserted.Count + tempHdrDeletedParagraph.Count;

                                    // look at the header changes
                                    foreach (var item in tempHdrParagraphChanged)
                                    {
                                        revCount++;
                                        totalRevCount++;
                                        lbRevisions.Items.Add(totalRevCount + Strings.wPeriod + s + Strings.wParaFormatChange);
                                    }

                                    foreach (var item in tempHdrDeletedParagraph)
                                    {
                                        revCount++;
                                        totalRevCount++;
                                        lbRevisions.Items.Add(totalRevCount + Strings.wPeriod + s + Strings.wParaDeleted);
                                    }

                                    foreach (var item in tempHdrRunChanged)
                                    {
                                        revCount++;
                                        totalRevCount++;
                                        lbRevisions.Items.Add(totalRevCount + Strings.wPeriod + s + Strings.wRunFormatChange);
                                    }

                                    foreach (var item in tempHdrDeleted)
                                    {
                                        revCount++;
                                        totalRevCount++;
                                        lbRevisions.Items.Add(totalRevCount + Strings.wPeriod + s + Strings.wDeletion + item.InnerText);
                                    }

                                    foreach (var item in tempHdrInserted)
                                    {
                                        if (item.Parent != null)
                                        {
                                            var textRuns = item.Elements<Run>().ToList();
                                            var parent = item.Parent;

                                            foreach (var textRun in textRuns)
                                            {
                                                revCount++;
                                                totalRevCount++;
                                                lbRevisions.Items.Add(totalRevCount + Strings.wPeriod + s + Strings.wInsertion + textRun.InnerText);
                                            }
                                        }
                                    }
                                }

                                if (revCountFooter > 0)
                                {
                                    var tempFtrParagraphChanged = ftrParagraphChanged.Where(item => item.Author == s).ToList();
                                    var tempFtrRunChanged = ftrRunChanged.Where(item => item.Author == s).ToList();
                                    var tempFtrDeleted = ftrDeleted.Where(item => item.Author == s).ToList();
                                    var tempFtrInserted = ftrInserted.Where(item => item.Author == s).ToList();
                                    var tempFtrDeletedParagraph = ftrDeletedParagraph.Where(item => item.Author == s).ToList();

                                    // look at the footer changes
                                    foreach (var item in tempFtrParagraphChanged)
                                    {
                                        revCount++;
                                        totalRevCount++;
                                        lbRevisions.Items.Add(totalRevCount + Strings.wPeriod + s + Strings.wParaFormatChange);
                                    }

                                    foreach (var item in tempFtrDeletedParagraph)
                                    {
                                        revCount++;
                                        totalRevCount++;
                                        lbRevisions.Items.Add(totalRevCount + Strings.wPeriod + s + Strings.wParaDeleted);
                                    }

                                    foreach (var item in tempFtrRunChanged)
                                    {
                                        revCount++;
                                        totalRevCount++;
                                        lbRevisions.Items.Add(totalRevCount + Strings.wPeriod + s + Strings.wRunFormatChange);
                                    }

                                    foreach (var item in tempFtrDeleted)
                                    {
                                        revCount++;
                                        totalRevCount++;
                                        lbRevisions.Items.Add(totalRevCount + Strings.wPeriod + s + Strings.wDeletion + item.InnerText);
                                    }

                                    foreach (var item in tempFtrInserted)
                                    {
                                        if (item.Parent != null)
                                        {
                                            var textRuns = item.Elements<Run>().ToList();
                                            var parent = item.Parent;

                                            foreach (var textRun in textRuns)
                                            {
                                                revCount++;
                                                totalRevCount++;
                                                lbRevisions.Items.Add(totalRevCount + Strings.wPeriod + s + Strings.wInsertion + textRun.InnerText);
                                            }
                                        }
                                    }
                                }
                            }

                            if (revCount + revCountHeader + revCountFooter == 0)
                            {
                                lbRevisions.Items.Add(s + " has no tracked changes in the document.");
                            }
                        }
                    }
                    else
                    {
                        // list the selected authors revisions
                        if (!string.IsNullOrEmpty(cbAuthors.SelectedItem.ToString()))
                        {
                            var tempParagraphChanged = paragraphChanged.Where(item => item.Author == cbAuthors.SelectedItem.ToString()).ToList();
                            var tempRunChanged = runChanged.Where(item => item.Author == cbAuthors.SelectedItem.ToString()).ToList();
                            var tempDeleted = deleted.Where(item => item.Author == cbAuthors.SelectedItem.ToString()).ToList();
                            var tempInserted = inserted.Where(item => item.Author == cbAuthors.SelectedItem.ToString()).ToList();
                            var tempDeletedParagraph = deletedParagraph.Where(item => item.Author == cbAuthors.SelectedItem.ToString()).ToList();

                            // look through the main changes
                            foreach (var item in tempParagraphChanged)
                            {
                                revCount++;
                                lbRevisions.Items.Add(revCount + Strings.wPeriod + cbAuthors.SelectedItem.ToString() + Strings.wParaFormatChange);
                            }

                            foreach (var item in tempDeletedParagraph)
                            {
                                revCount++;
                                lbRevisions.Items.Add(revCount + Strings.wPeriod + cbAuthors.SelectedItem.ToString() + Strings.wParaDeleted);
                            }

                            foreach (var item in tempRunChanged)
                            {
                                revCount++;
                                lbRevisions.Items.Add(revCount + Strings.wPeriod + cbAuthors.SelectedItem.ToString() + Strings.wRunFormatChange);
                            }

                            foreach (var item in tempDeleted)
                            {
                                revCount++;
                                lbRevisions.Items.Add(revCount + Strings.wPeriod + cbAuthors.SelectedItem.ToString() + Strings.wDeletion + item.InnerText);
                            }

                            foreach (var item in tempInserted)
                            {
                                if (item.Parent is not null)
                                {
                                    var textRuns = item.Elements<Run>().ToList();
                                    var parent = item.Parent;

                                    foreach (var textRun in textRuns)
                                    {
                                        revCount++;
                                        lbRevisions.Items.Add(revCount + Strings.wPeriod + cbAuthors.SelectedItem.ToString() + Strings.wInsertion + textRun.InnerText);
                                    }
                                }
                            }

                            if (revCountHeader > 0)
                            {
                                var tempHdrParagraphChanged = hdrParagraphChanged.Where(item => item.Author == cbAuthors.SelectedItem.ToString()).ToList();
                                var tempHdrRunChanged = hdrRunChanged.Where(item => item.Author == cbAuthors.SelectedItem.ToString()).ToList();
                                var tempHdrDeleted = hdrDeleted.Where(item => item.Author == cbAuthors.SelectedItem.ToString()).ToList();
                                var tempHdrInserted = hdrInserted.Where(item => item.Author == cbAuthors.SelectedItem.ToString()).ToList();
                                var tempHdrDeletedParagraph = hdrDeletedParagraph.Where(item => item.Author == cbAuthors.SelectedItem.ToString()).ToList();

                                // look at the header changes
                                foreach (var item in tempHdrParagraphChanged)
                                {
                                    revCount++;
                                    lbRevisions.Items.Add(revCount + Strings.wPeriod + cbAuthors.SelectedItem.ToString() + Strings.wParaFormatChange);
                                }

                                foreach (var item in tempHdrDeletedParagraph)
                                {
                                    revCount++;
                                    lbRevisions.Items.Add(revCount + Strings.wPeriod + cbAuthors.SelectedItem.ToString() + Strings.wParaDeleted);
                                }

                                foreach (var item in tempHdrRunChanged)
                                {
                                    revCount++;
                                    lbRevisions.Items.Add(revCount + Strings.wPeriod + cbAuthors.SelectedItem.ToString() + Strings.wRunFormatChange);
                                }

                                foreach (var item in tempHdrDeleted)
                                {
                                    revCount++;
                                    lbRevisions.Items.Add(revCount + Strings.wPeriod + cbAuthors.SelectedItem.ToString() + Strings.wDeletion + item.InnerText);
                                }

                                foreach (var item in tempHdrInserted)
                                {
                                    if (item.Parent != null)
                                    {
                                        var textRuns = item.Elements<Run>().ToList();
                                        var parent = item.Parent;

                                        foreach (var textRun in textRuns)
                                        {
                                            revCount++;
                                            lbRevisions.Items.Add(revCount + Strings.wPeriod + cbAuthors.SelectedItem.ToString() + Strings.wInsertion + textRun.InnerText);
                                        }
                                    }
                                }
                            }

                            if (revCountFooter > 0)
                            {
                                var tempFtrParagraphChanged = ftrParagraphChanged.Where(item => item.Author == cbAuthors.SelectedItem.ToString()).ToList();
                                var tempFtrRunChanged = ftrRunChanged.Where(item => item.Author == cbAuthors.SelectedItem.ToString()).ToList();
                                var tempFtrDeleted = ftrDeleted.Where(item => item.Author == cbAuthors.SelectedItem.ToString()).ToList();
                                var tempFtrInserted = ftrInserted.Where(item => item.Author == cbAuthors.SelectedItem.ToString()).ToList();
                                var tempFtrDeletedParagraph = ftrDeletedParagraph.Where(item => item.Author == cbAuthors.SelectedItem.ToString()).ToList();

                                // look at the footer changes
                                foreach (var item in tempFtrParagraphChanged)
                                {
                                    revCount++;
                                    lbRevisions.Items.Add(revCount + Strings.wPeriod + cbAuthors.SelectedItem.ToString() + Strings.wParaFormatChange);
                                }

                                foreach (var item in tempFtrDeletedParagraph)
                                {
                                    revCount++;
                                    lbRevisions.Items.Add(revCount + Strings.wPeriod + cbAuthors.SelectedItem.ToString() + Strings.wParaDeleted);
                                }

                                foreach (var item in tempFtrRunChanged)
                                {
                                    revCount++;
                                    lbRevisions.Items.Add(revCount + Strings.wPeriod + cbAuthors.SelectedItem.ToString() + Strings.wRunFormatChange);
                                }

                                foreach (var item in tempFtrDeleted)
                                {
                                    revCount++;
                                    lbRevisions.Items.Add(revCount + Strings.wPeriod + cbAuthors.SelectedItem.ToString() + Strings.wDeletion + item.InnerText);
                                }

                                foreach (var item in tempFtrInserted)
                                {
                                    if (item.Parent != null)
                                    {
                                        var textRuns = item.Elements<Run>().ToList();
                                        var parent = item.Parent;

                                        foreach (var textRun in textRuns)
                                        {
                                            revCount++;
                                            lbRevisions.Items.Add(revCount + Strings.wPeriod + cbAuthors.SelectedItem.ToString() + Strings.wInsertion + textRun.InnerText);
                                        }
                                    }
                                }
                            }

                            if (revCount + revCountFooter + revCountHeader == 0)
                            {
                                lbRevisions.Items.Add(cbAuthors.SelectedItem.ToString() + " has no changes.");
                                return;
                            }
                        }
                        else
                        {
                            lbRevisions.Items.Add("** There are no revisions in this document **");
                            return;
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
            LoadAuthorChanges();
        }

        private void FrmWordCommands_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { Close(); }
        }

        private void BtnAcceptChanges_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
                {
                    if (Word.AcceptTrackedChanges(document.MainDocumentPart.Document, cbAuthors.Text))
                    {
                        MessageBox.Show(cbAuthors.Text + " changes accepted.");
                        cbAuthors.Items.Clear();
                    }
                    else
                    {
                        MessageBox.Show(cbAuthors.Text + " had no changes.");
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
    }
}
