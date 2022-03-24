// .NET refs
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Text;

// open xml sdk refs
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// shortcut namespace refs
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using DocumentFormat.OpenXml.CustomProperties;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using Comment = DocumentFormat.OpenXml.Wordprocessing.Comment;
using Font = DocumentFormat.OpenXml.Wordprocessing.Font;
using Hyperlink = DocumentFormat.OpenXml.Wordprocessing.Hyperlink;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableStyle = DocumentFormat.OpenXml.Wordprocessing.TableStyle;

namespace Office_File_Explorer.Helpers
{
    public static class Word
    {
        public static bool fSuccess;

        public static bool RemoveCustomTitleProp(string docName)
        {
            fSuccess = false;

            using (WordprocessingDocument document = WordprocessingDocument.Open(docName, true))
            {
                foreach (CustomDocumentProperty cdp in document.CustomFilePropertiesPart.RootElement)
                {
                    if (cdp.Name.ToString().ToLower() == "title")
                    {
                        cdp.Remove();
                        fSuccess = true;
                    }
                }
            }

            return fSuccess;
        }

        /// <summary>
        /// accept track changes for specific author
        /// check the body as well as headers and footers
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="author"></param>
        /// <returns></returns>
        public static bool AcceptTrackedChanges(Document doc, string author)
        {
            fSuccess = false;
            int revCount = 0;
            int revCountHeader = 0;
            int revCountFooter = 0;

            List<ParagraphPropertiesChange> paragraphChanged = doc.Descendants<ParagraphPropertiesChange>().ToList();
            List<RunPropertiesChange> runChanged = doc.Descendants<RunPropertiesChange>().ToList();
            List<DeletedRun> deleted = doc.Descendants<DeletedRun>().ToList();
            List<Deleted> deletedParagraph = doc.Descendants<Deleted>().ToList();
            List<InsertedRun> inserted = doc.Descendants<InsertedRun>().ToList();

            var tempParagraphChanged = paragraphChanged.Where(item => item.Author == author).ToList();
            var tempRunChanged = runChanged.Where(item => item.Author == author).ToList();
            var tempDeleted = deleted.Where(item => item.Author == author).ToList();
            var tempInserted = inserted.Where(item => item.Author == author).ToList();
            var tempDeletedParagraph = deletedParagraph.Where(item => item.Author == author).ToList();
            revCount = tempParagraphChanged.Count + tempRunChanged.Count + tempDeleted.Count + tempDeletedParagraph.Count + tempInserted.Count;

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

            if (revCount > 0)
            {
                foreach (var item in paragraphChanged)
                {
                    item.Remove();
                    fSuccess = true;
                }

                foreach (var item in deletedParagraph)
                {
                    item.Remove();
                    fSuccess = true;
                }

                foreach (var item in runChanged)
                {
                    item.Remove();
                    fSuccess = true;
                }

                foreach (var item in deleted)
                {
                    item.Remove();
                    fSuccess = true;
                }

                foreach (var item in inserted)
                {
                    if (item.Parent is not null)
                    {
                        var textRuns = item.Elements<Run>().ToList();
                        var parent = item.Parent;

                        foreach (var textRun in textRuns)
                        {
                            item.RemoveAttribute("rsidR", parent.NamespaceUri);
                            item.RemoveAttribute("sidRPr", parent.NamespaceUri);
                            parent.InsertBefore(textRun.CloneNode(true), item);
                        }
                        item.Remove();
                        fSuccess = true;
                    }
                }
            }

            if (revCountHeader > 0)
            {
                foreach (var item in hdrParagraphChanged)
                {
                    item.Remove();
                    fSuccess = true;
                }

                foreach (var item in hdrDeletedParagraph)
                {
                    item.Remove();
                    fSuccess = true;
                }

                foreach (var item in hdrRunChanged)
                {
                    item.Remove();
                    fSuccess = true;
                }

                foreach (var item in hdrDeleted)
                {
                    item.Remove();
                    fSuccess = true;
                }

                foreach (var item in hdrInserted)
                {
                    if (item.Parent is not null)
                    {
                        var textRuns = item.Elements<Run>().ToList();
                        var parent = item.Parent;

                        foreach (var textRun in textRuns)
                        {
                            item.RemoveAttribute("rsidR", parent.NamespaceUri);
                            item.RemoveAttribute("sidRPr", parent.NamespaceUri);
                            parent.InsertBefore(textRun.CloneNode(true), item);
                        }
                        item.Remove();
                        fSuccess = true;
                    }
                }
            }

            if (revCountFooter > 0)
            {
                foreach (var item in ftrParagraphChanged)
                {
                    item.Remove();
                    fSuccess = true;
                }

                foreach (var item in ftrDeletedParagraph)
                {
                    item.Remove();
                    fSuccess = true;
                }

                foreach (var item in ftrRunChanged)
                {
                    item.Remove();
                    fSuccess = true;
                }

                foreach (var item in ftrDeleted)
                {
                    item.Remove();
                    fSuccess = true;
                }

                foreach (var item in ftrInserted)
                {
                    if (item.Parent is not null)
                    {
                        var textRuns = item.Elements<Run>().ToList();
                        var parent = item.Parent;

                        foreach (var textRun in textRuns)
                        {
                            item.RemoveAttribute("rsidR", parent.NamespaceUri);
                            item.RemoveAttribute("sidRPr", parent.NamespaceUri);
                            parent.InsertBefore(textRun.CloneNode(true), item);
                        }
                        item.Remove();
                        fSuccess = true;
                    }
                }
            }

            if (fSuccess)
            {
                doc.Save();
            }
            
            return fSuccess;
        }


        /// <summary>
        /// Given a document name and an author name, accept all revisions by the specified author. 
        /// Pass an empty string for the author to accept all revisions.
        /// </summary>
        /// <param name="docName"></param>
        /// <param name="authorName"></param>
        public static List<string> AcceptRevisions(string docName, string authorName)
        {
            List<string> output = new List<string>();
            using (WordprocessingDocument document = WordprocessingDocument.Open(docName, true))
            {
                Document doc = document.MainDocumentPart.Document;                
                List<string> tempAuthors = new List<string>();
                tempAuthors = GetAllAuthors(document.MainDocumentPart.Document);
                                
                if (authorName == Strings.wAllAuthors)
                {
                    // create a temp list for each author so we can loop the changes individually and list them
                    foreach (string s in tempAuthors)
                    {
                        if (AcceptTrackedChanges(doc, s))
                        {
                            output.Add(s + " - changes accepted.");
                        }
                    }
                    doc.Save();
                }
                else
                {
                    // for single author, just loop that authors from the original list
                    if (!string.IsNullOrEmpty(authorName))
                    {
                        if (AcceptTrackedChanges(doc, authorName))
                        {
                            output.Add(authorName + " - changes accepted.");
                        }
                    }
                    doc.Save();
                }
            }

            return output;
        }

        /// <summary>
        /// Given a document name, set the print orientation for all the sections of the document.
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="newOrientation"></param>
        public static void SetPrintOrientation(string fileName, PageOrientationValues newOrientation)
        {
            using (var document = WordprocessingDocument.Open(fileName, true))
            {
                bool documentChanged = false;

                var docPart = document.MainDocumentPart;
                var sections = docPart.Document.Descendants<SectionProperties>();

                foreach (SectionProperties sectPr in sections)
                {
                    bool pageOrientationChanged = false;

                    PageSize pgSz = sectPr.Descendants<PageSize>().FirstOrDefault();
                    if (pgSz is not null)
                    {
                        // No Orient property? Create it now. Otherwise, just 
                        // set its value. Assume that the default orientation 
                        // is Portrait.
                        if (pgSz.Orient is null)
                        {
                            // Need to create the attribute. You do not need to 
                            // create the Orient property if the property does not 
                            // already exist, and you are setting it to Portrait. 
                            // That is the default value.
                            if (newOrientation != PageOrientationValues.Portrait)
                            {
                                pageOrientationChanged = true;
                                documentChanged = true;
                                pgSz.Orient = new EnumValue<PageOrientationValues>(newOrientation);
                            }
                        }
                        else
                        {
                            // The Orient property exists, but its value
                            // is different than the new value.
                            if (pgSz.Orient.Value != newOrientation)
                            {
                                pgSz.Orient.Value = newOrientation;
                                pageOrientationChanged = true;
                                documentChanged = true;
                            }
                        }

                        if (pageOrientationChanged)
                        {
                            // Changing the orientation is not enough. You must also 
                            // change the page size.
                            var width = pgSz.Width;
                            var height = pgSz.Height;
                            pgSz.Width = height;
                            pgSz.Height = width;

                            PageMargin pgMar = sectPr.Descendants<PageMargin>().FirstOrDefault();
                            if (pgMar is not null)
                            {
                                // Rotate margins. Printer settings control how far you 
                                // rotate when switching to landscape mode. Not having those
                                // settings, this code rotates 90 degrees. You could easily
                                // modify this behavior, or make it a parameter for the 
                                // procedure.
                                var top = pgMar.Top.Value;
                                var bottom = pgMar.Bottom.Value;
                                var left = pgMar.Left.Value;
                                var right = pgMar.Right.Value;

                                pgMar.Top = new Int32Value((int)left);
                                pgMar.Bottom = new Int32Value((int)right);
                                pgMar.Left = new UInt32Value((uint)Math.Max(0, bottom));
                                pgMar.Right = new UInt32Value((uint)Math.Max(0, top));
                            }
                        }
                    }
                }

                if (documentChanged)
                {
                    docPart.Document.Save();
                }
            }
        }

        public static bool RemoveBreaks(string filename)
        {
            fSuccess = false;

            // this function will remove both page and section breaks in a document
            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filename, true))
            {
                MainDocumentPart mainPart = myDoc.MainDocumentPart;

                List<Break> breaks = mainPart.Document.Descendants<Break>().ToList();

                foreach (Break b in breaks)
                {
                    b.Remove();
                }

                List<ParagraphProperties> paraProps = mainPart.Document.Descendants<ParagraphProperties>()
                .Where(pPr => IsSectionProps(pPr)).ToList();

                foreach (ParagraphProperties pPr in paraProps)
                {
                    pPr.RemoveChild<SectionProperties>(pPr.GetFirstChild<SectionProperties>());
                }

                mainPart.Document.Save();
                fSuccess = true;
            }

            return fSuccess;
        }

        /// <summary>
        /// given a numId, find the element in the Numbering.xml file with that same numId
        /// get the abstractnumid and delete the numId element
        /// then you can delete the abstractnumid
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="numId">orphaned ListTemplate to delete</param>
        public static void RemoveListTemplatesNumId(string filename, string numId)
        {
            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filename, true))
            {
                string absNumIdToRemove = string.Empty;

                var absNumsInUseList = myDoc.MainDocumentPart.NumberingDefinitionsPart.Numbering.Descendants<AbstractNum>().ToList();
                var numInstancesInUseList = myDoc.MainDocumentPart.NumberingDefinitionsPart.Numbering.Descendants<NumberingInstance>().ToList();

                // loop each numberinginstance and if the orphaned id matches, delete it and keep track of the abstractnumid
                foreach (NumberingInstance ni in numInstancesInUseList)
                {
                    if (ni.NumberID == numId)
                    {
                        absNumIdToRemove = ni.AbstractNumId.Val.ToString();
                        ni.Remove();
                        break;
                    }
                }

                // now that we have the abstractnum, loop that list and delete that one
                foreach (AbstractNum an in absNumsInUseList)
                {
                    if (an.AbstractNumberId == Int32.Parse(absNumIdToRemove))
                    {
                        an.Remove();
                        break;
                    }
                }

                myDoc.MainDocumentPart.Document.Save();
            }
        }

        static bool IsSectionProps(ParagraphProperties pPr)
        {
            SectionProperties sectPr = pPr.GetFirstChild<SectionProperties>();

            if (sectPr is null)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public static bool RemoveComments(string filename)
        {
            fSuccess = false;

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filename, true))
            {
                MainDocumentPart mainPart = myDoc.MainDocumentPart;

                //Delete the comment part, plus any other part referenced, like image parts 
                mainPart.DeletePart(mainPart.WordprocessingCommentsPart);

                //Find all elements that are associated with comments 
                IEnumerable<OpenXmlElement> elementList = mainPart.Document.Descendants()
                .Where(el => el is CommentRangeStart || el is CommentRangeEnd || el is CommentReference);

                //Delete every found element 
                foreach (OpenXmlElement e in elementList)
                {
                    e.Remove();
                }

                //Save changes 
                mainPart.Document.Save();
                fSuccess = true;
            }

            return fSuccess;
        }

        // Given a document, remove all hidden text.
        public static bool DeleteHiddenText(string docName)
        {
            fSuccess = false;

            using (WordprocessingDocument document = WordprocessingDocument.Open(docName, true))
            {
                Document doc = document.MainDocumentPart.Document;
                var hiddenItems = doc.Descendants<Vanish>().ToList();
                foreach (var item in hiddenItems)
                {
                    // Need to go up at least two levels to get to the run.
                    if ((item.Parent is not null) &&
                      (item.Parent.Parent is not null) &&
                      (item.Parent.Parent.Parent is not null))
                    {
                        var topNode = item.Parent.Parent;
                        var topParentNode = item.Parent.Parent.Parent;
                        if (topParentNode is not null)
                        {
                            topNode.Remove();
                            // No more children? Remove the parent node, as well.
                            if (!topParentNode.HasChildren)
                            {
                                topParentNode.Remove();
                            }
                        }
                    }
                }
                doc.Save();
                fSuccess = true;
            }

            return fSuccess;
        }

        public static List<string> RemoveUnusedStyles(string filePath)
        {
            List<string> lstStyles = new List<string>();
            int count = 0;

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filePath, true))
            {
                MainDocumentPart mainPart = myDoc.MainDocumentPart;
                StyleDefinitionsPart stylePart = mainPart.StyleDefinitionsPart;
                bool styleDeleted = false;

                try
                {
                    List<string> baseStyleChains = new List<string>();
                    List<string> baseStyles = new List<string>();
                    string[] words = null;

                    // first, get all the base styles
                    foreach (OpenXmlElement tempEl in stylePart.Styles.Elements())
                    {
                        if (tempEl.LocalName == "style")
                        {
                            Style tempStyle = (Style)tempEl;
                            if (tempStyle.BasedOn is null)
                            {
                                baseStyles.Add(tempStyle.StyleId);
                            }
                        }
                    }

                    // loop base styles and recursively get basedon chain
                    // this should create a string of the linked list sequence of styles
                    foreach (string sBase in baseStyles)
                    {
                        StringBuilder tempBaseStyleChain = new StringBuilder();
                        tempBaseStyleChain.Append(sBase);

                        StringBuilder baseStyleChain = GetBasedOnStyleChain(stylePart, sBase, tempBaseStyleChain);

                        if (baseStyleChain.ToString().Contains(Strings.wArrowOnly))
                        {
                            baseStyleChains.Add(baseStyleChain.ToString());
                        }
                    }

                    // now we need to parse out the style chains for each individual styleid in reverse order
                    // if the style is applied to any p, r, or t, don't delete
                    // if the style is default, nextParagraphStyle or LinkedStyle, don't delete
                    // if neither of these is true, we can delete the style
                    if (baseStyleChains.Count > 0)
                    {
                        foreach (string b in baseStyleChains)
                        {
                            bool doNotDeleteAnyInChain = false;
                            string[] separatingStrings = { Strings.wArrowOnly };
                            words = b.Split(separatingStrings, StringSplitOptions.None);

                            if (words.Length > 0)
                            {
                                foreach (string w in words.Reverse())
                                {
                                    int pWStyleCount = ParagraphsByStyleId(mainPart, w).Count();
                                    int rWStyleCount = RunsByStyleId(mainPart, w).Count();
                                    int tWStyleCount = TablesByStyleId(mainPart, w).Count();
                                    count += 1;

                                    // if the style is used in a para, run or table, don't delete
                                    if (pWStyleCount > 0 || rWStyleCount > 0 || tWStyleCount > 0)
                                    {
                                        lstStyles.Add(count + Strings.wPeriod + Strings.doNotDeleteStyle + w);
                                        doNotDeleteAnyInChain = true;
                                    }

                                    // if the style is not used, candidate for delete
                                    if (pWStyleCount == 0 && rWStyleCount == 0 && tWStyleCount == 0)
                                    {
                                        // if the previous style in the chain was true, we need to leave these alone
                                        if (doNotDeleteAnyInChain == true)
                                        {
                                            lstStyles.Add(count + Strings.wPeriod + Strings.doNotDeleteStyle + w);
                                        }
                                        else
                                        {
                                            // if we get here, the style is not applied and we can check the rest of the requirements for deletion
                                            foreach (OpenXmlElement tempEl in stylePart.Styles.Elements())
                                            {
                                                if (tempEl.LocalName == "style")
                                                {
                                                    Style tempStyle = (Style)tempEl;
                                                    if (tempStyle.StyleId == w)
                                                    {
                                                        // this is the last leg of the style still in use checks
                                                        // if default, nextpara and linked are all null, this style can be deleted
                                                        if (tempStyle.Default is null && tempStyle.NextParagraphStyle is null && tempStyle.LinkedStyle is null)
                                                        {
                                                            lstStyles.Add(count + Strings.wPeriod + Strings.deleteStyle + w);
                                                            tempEl.Remove();
                                                            styleDeleted = true;
                                                        }
                                                        else
                                                        {
                                                            lstStyles.Add(count + Strings.wPeriod + Strings.doNotDeleteStyle + w);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (NullReferenceException)
                {
                    lstStyles.Add("BtnDeleteUnusedStyles - Missing StylesWithEffects part.");
                }

                // if we deleted any style, save the file
                if (styleDeleted == true)
                {
                    myDoc.MainDocumentPart.Document.Save();
                }
            }

            return lstStyles;
        }

        /// <summary>
        /// delete the footnotes in a file
        /// </summary>
        /// <param name="docName"></param>
        /// <returns></returns>
        public static bool RemoveFootnotes(string docName)
        {
            fSuccess = false;

            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(docName, true))
            {
                FootnotesPart fnp = wdDoc.MainDocumentPart.FootnotesPart;
                if (fnp is not null)
                {
                    var footnotes = fnp.Footnotes.Elements<Footnote>();
                    var references = wdDoc.MainDocumentPart.Document.Body.Descendants<FootnoteReference>().ToArray();

                    foreach (var reference in references)
                    {
                        reference.Remove();
                    }

                    foreach (var footnote in footnotes)
                    {
                        footnote.Remove();
                    }
                }

                wdDoc.MainDocumentPart.Document.Save();
                fSuccess = true;
            }

            return fSuccess;
        }

        /// <summary>
        /// Delete endnotes from the document
        /// </summary>
        /// <param name="docName"></param>
        /// <returns></returns>
        public static bool RemoveEndnotes(string docName)
        {
            fSuccess = false;

            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(docName, true))
            {
                if (wdDoc.MainDocumentPart.GetPartsOfType<EndnotesPart>().Any())
                {
                    MainDocumentPart mainPart = wdDoc.MainDocumentPart;

                    var enr = mainPart.Document.Descendants<EndnoteReference>().ToArray();
                    foreach (var e in enr)
                    {
                        e.Remove();
                    }

                    EndnotesPart ep = mainPart.EndnotesPart;
                    ep.Endnotes = CreateDefaultEndnotes();
                    mainPart.Document.Save();
                    fSuccess = true;
                }
            }

            return fSuccess;
        }

        /// <summary>
        /// Delete headers and footers from a document.
        /// </summary>
        /// <param name="docName"></param>
        /// <returns></returns>
        public static bool RemoveHeadersFooters(string docName)
        {
            fSuccess = false;

            // Given a document name, remove all headers and footers.
            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(docName, true))
            {
                if (wdDoc.MainDocumentPart.GetPartsOfType<HeaderPart>().Any() || wdDoc.MainDocumentPart.GetPartsOfType<FooterPart>().Any())
                {
                    // Remove header and footer parts.
                    wdDoc.MainDocumentPart.DeleteParts(wdDoc.MainDocumentPart.HeaderParts);
                    wdDoc.MainDocumentPart.DeleteParts(wdDoc.MainDocumentPart.FooterParts);

                    Document doc = wdDoc.MainDocumentPart.Document;

                    // Remove references to the headers and footers.
                    var headers =
                      doc.Descendants<HeaderReference>().ToList();
                    foreach (var header in headers)
                    {
                        header.Parent.RemoveChild(header);
                    }

                    var footers = doc.Descendants<FooterReference>().ToList();
                    foreach (var footer in footers)
                    {
                        footer.Parent.RemoveChild(footer);
                    }
                    doc.Save();
                    fSuccess = true;
                }
            }

            return fSuccess;
        }

        public static List<string> LstFieldCodes(string fPath)
        {
            List<string> ltFieldCodes = new List<string>();

            using (WordprocessingDocument package = WordprocessingDocument.Open(fPath, false))
            {
                IEnumerable<Run> rList = package.MainDocumentPart.Document.Descendants<Run>();
                IEnumerable<Paragraph> pList = package.MainDocumentPart.Document.Descendants<Paragraph>();

                List<string> fieldCharList = new List<string>();
                List<string> fieldCodeList = new List<string>();

                foreach (Run r in rList)
                {
                    foreach (OpenXmlElement oxe in r.ChildElements)
                    {
                        if (oxe.LocalName == "fldChar")
                        {
                            FieldChar fc = new FieldChar();
                            fc = (FieldChar)oxe;
                            if (fc.FieldCharType == Strings.wBegin)
                            {
                                fieldCharList.Add(Strings.wBegin);
                            }
                            else if (fc.FieldCharType == Strings.wEnd)
                            {
                                fieldCharList.Add(Strings.wEnd);
                            }
                        }
                        else if (oxe.LocalName == "instrText")
                        {
                            fieldCharList.Add(oxe.InnerText);
                        }
                    }
                }

                foreach (Paragraph p in pList)
                {
                    foreach (OpenXmlElement oxe in p.ChildElements)
                    {
                        if (oxe.LocalName == "fldSimple")
                        {
                            SimpleField sf = new SimpleField();
                            sf = (SimpleField)oxe;
                            fieldCodeList.Add(sf.Instruction);
                        }
                    }
                }

                if (fieldCharList.Count == 0 && fieldCodeList.Count == 0)
                {
                    return ltFieldCodes;
                }
                else
                {
                    StringBuilder sb = new StringBuilder();
                    int fCount = 0;

                    foreach (string s in fieldCharList)
                    {
                        if (s == Strings.wBegin)
                        {
                            continue;
                        }
                        else if (s == Strings.wEnd)
                        {
                            // display the field code values
                            fCount++;
                            ltFieldCodes.Add(fCount + Strings.wPeriod + sb);
                            sb.Clear();
                        }
                        else
                        {
                            sb.Append(s);
                        }
                    }

                    foreach (string s in fieldCodeList)
                    {
                        fCount++;
                        ltFieldCodes.Add(fCount + Strings.wPeriod + s);
                    }
                }
            }

            return ltFieldCodes;
        }

        public static List<string> LstTables(string fPath)
        {
            List<string> ltTables = new List<string>();

            using (WordprocessingDocument package = WordprocessingDocument.Open(fPath, false))
            {
                IEnumerable<Table> tList = package.MainDocumentPart.Document.Descendants<Table>();
                int tableCount = 0;
                
                foreach (var t in tList)
                {
                    tableCount++;
                    bool isNested = false;

                    // check if the table is nested
                    OpenXmlElement tOxe = t;
                    if (tOxe.Parent.ToString() == Strings.dfowTableCell)
                    {
                        isNested = true;
                    }

                    // track the row/col for each table
                    int rowCount = 0;
                    int colCount = 0;

                    foreach (OpenXmlElement oxe in t.ChildElements)
                    {
                        // get the count of columns
                        if (oxe.GetType().ToString() == Strings.dfowTableGrid)
                        {
                            colCount = oxe.ChildElements.Count;
                        }

                        // track the number of rows
                        if (oxe.GetType().ToString() == Strings.dfowTableRow)
                        {
                            rowCount++;
                        }
                    }

                    if (isNested)
                    {
                        ltTables.Add("Table " + tableCount + Strings.wEqualSign + colCount + Strings.wXChar + rowCount + " (Nested Table)");
                    }
                    else
                    {
                        ltTables.Add("Table " + tableCount + Strings.wEqualSign + colCount + Strings.wXChar + rowCount);
                    }
                }
            }

            return ltTables;
        }

        public static List<string> LstComments(string fPath)
        {
            List<string> ltComments = new List<string>();

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(fPath, false))
            {
                int count = 0;

                // first list all comments in the comment part
                WordprocessingCommentsPart commentsPart = myDoc.MainDocumentPart.WordprocessingCommentsPart;
                if (commentsPart is null)
                {
                    return ltComments;
                }
                else
                {
                    foreach (Comment cm in commentsPart.Comments)
                    {
                        count++;
                        ltComments.Add(count + Strings.wPeriod + cm.Author + Strings.wColon + cm.InnerText);
                    }
                }

                // now we can check how many comment references there are and display that number, they should be the same
                IEnumerable<CommentReference> crList = myDoc.MainDocumentPart.Document.Descendants<CommentReference>();
                IEnumerable<OpenXmlUnknownElement> unknownList = myDoc.MainDocumentPart.Document.Descendants<OpenXmlUnknownElement>();
                int cRefCount = crList.Count();
                foreach (OpenXmlUnknownElement uk in unknownList)
                {
                    if (uk.LocalName == "commentReference")
                    {
                        cRefCount++;
                    }
                }

                ltComments.Add(string.Empty);
                ltComments.Add("Total Comments = " + count);
                ltComments.Add("Comment References = " + cRefCount);
            }

            return ltComments;
        }

        public static List<string> LstBookmarks(string fPath)
        {
            List<string> ltBookmarks = new List<string>();

            using (WordprocessingDocument package = WordprocessingDocument.Open(fPath, false))
            {
                IEnumerable<BookmarkStart> bkList = package.MainDocumentPart.Document.Descendants<BookmarkStart>();
                ltBookmarks.Add("** Document Bookmarks **");

                if (bkList.Any())
                {
                    int count = 1;

                    foreach (BookmarkStart bk in bkList)
                    {
                        var cElem = bk.Parent;
                        var pElem = bk.Parent;
                        bool endLoop = false;
                        string isCorruptText = string.Empty;

                        do
                        {
                            if (cElem is not null && cElem.Parent is not null && cElem.Parent.ToString().Contains(Strings.dfowSdt))
                            {
                                foreach (OpenXmlElement oxe in cElem.Parent.ChildElements)
                                {
                                    if (oxe.GetType().Name == "SdtProperties")
                                    {
                                        foreach (OpenXmlElement oxeSdtAlias in oxe)
                                        {
                                            if (oxeSdtAlias.GetType().Name == "SdtContentText")
                                            {
                                                // if the parent is a content control, bookmark is only allowed in rich text
                                                // if this is a plain text control, it is invalid
                                                isCorruptText = " <-- ## Warning ## - this bookmark is in a plain text content control which is not allowed";
                                                endLoop = true;
                                            }
                                        }
                                    }
                                }

                                // set next element
                                pElem = cElem.Parent;
                                cElem = pElem;
                            }
                            else
                            {
                                // if the next element is null, bail
                                if (cElem is null || cElem.Parent is null)
                                {
                                    endLoop = true;
                                }

                                // set next element
                                pElem = cElem.Parent;
                                cElem = pElem;

                                // if the parent is body, we can stop looping up
                                // otherwise, we can continue moving up the element chain
                                if (pElem is not null && pElem.ToString() == Strings.dfowBody)
                                {
                                    endLoop = true;
                                }
                            }
                        } while (endLoop == false);

                        ltBookmarks.Add(count + Strings.wPeriod + bk.Name + isCorruptText);
                        count++;
                    }
                }

                if (package.MainDocumentPart.WordprocessingCommentsPart is not null)
                {
                    if (package.MainDocumentPart.WordprocessingCommentsPart.Comments is not null)
                    {
                        IEnumerable<BookmarkStart> bkCommentList = package.MainDocumentPart.WordprocessingCommentsPart.Comments.Descendants<BookmarkStart>();
                        int bkCommentCount = 0;

                        if (bkCommentList.Any())
                        {
                            ltBookmarks.Add(string.Empty);
                            ltBookmarks.Add("** Comment Bookmarks ** ");

                            foreach (BookmarkStart bkc in bkCommentList)
                            {
                                bkCommentCount++;
                                ltBookmarks.Add(bkCommentCount + Strings.wPeriod + bkc.Name);
                            }
                        }
                    }
                }
            }

            return ltBookmarks;
        }

        public static string GetDocSecurity(string val)
        {
            string docSecurity;
            switch (val)
            {
                case "0":
                    docSecurity = "None";
                    break;
                case "1":
                    docSecurity = "Password Protected";
                    break;
                case "2":
                    docSecurity = "Read-Only Recommended";
                    break;
                case "4":
                    docSecurity = "Read-Only Enforced";
                    break;
                case "8":
                    docSecurity = "Locked For Annotation";
                    break;
                default:
                    docSecurity = "Unknown";
                    break;
            }

            return docSecurity;
        }

        public static List<string> LstDocProps(string fPath)
        {
            List<string> ltDocProps = new List<string>();
            List<string> compatList = new List<string>();
            List<string> settingList = new List<string>();
            List<string> mathPrList = new List<string>();
            List<string> rsidList = new List<string>();
            StringBuilder sb = new StringBuilder();

            compatList.Add(string.Empty);
            compatList.Add("---- Compatibility Settings ---- ");
            settingList.Add("---- Document Settings ---- ");

            using (WordprocessingDocument doc = WordprocessingDocument.Open(fPath, false))
            {
                DocumentSettingsPart docSettingsPart = doc.MainDocumentPart.DocumentSettingsPart;

                // add the standard file props
                ltDocProps.Add("Creator : " + doc.PackageProperties.Creator);
                ltDocProps.Add("Created : " + doc.PackageProperties.Created);
                ltDocProps.Add("Last Modified By : " + doc.PackageProperties.LastModifiedBy);
                ltDocProps.Add("Last Printed : " + doc.PackageProperties.LastPrinted);
                ltDocProps.Add("Modified : " + doc.PackageProperties.Modified);
                ltDocProps.Add("Subject : " + doc.PackageProperties.Subject);
                ltDocProps.Add("Revision : " + doc.PackageProperties.Revision);
                ltDocProps.Add("Title : " + doc.PackageProperties.Title);
                ltDocProps.Add("Version : " + doc.PackageProperties.Version);
                ltDocProps.Add("Category : " + doc.PackageProperties.Category);
                ltDocProps.Add("ContentStatus : " + doc.PackageProperties.ContentStatus);
                ltDocProps.Add("ContentType : " + doc.PackageProperties.ContentType);
                ltDocProps.Add("Description : " + doc.PackageProperties.Description);
                ltDocProps.Add("Language : " + doc.PackageProperties.Language);
                ltDocProps.Add("Identifier : " + doc.PackageProperties.Identifier);
                ltDocProps.Add("Keywords : " + doc.PackageProperties.Keywords);

                // add the extended file properties
                if (doc.ExtendedFilePropertiesPart is not null)
                {
                    if (doc.ExtendedFilePropertiesPart.Properties.Application is not null)
                    {
                        ltDocProps.Add("Application = " + doc.ExtendedFilePropertiesPart.Properties.Application.Text);
                    }

                    if (doc.ExtendedFilePropertiesPart.Properties.ApplicationVersion is not null)
                    {
                        ltDocProps.Add("Application Version = " + doc.ExtendedFilePropertiesPart.Properties.ApplicationVersion.Text);
                    }

                    if (doc.ExtendedFilePropertiesPart.Properties.Characters is not null)
                    {
                        ltDocProps.Add("Characters = " + doc.ExtendedFilePropertiesPart.Properties.Characters.Text);
                    }

                    if (doc.ExtendedFilePropertiesPart.Properties.CharactersWithSpaces is not null)
                    {
                        ltDocProps.Add("Characters With Spaces = " + doc.ExtendedFilePropertiesPart.Properties.CharactersWithSpaces.Text);
                    }

                    if (doc.ExtendedFilePropertiesPart.Properties.Company is not null)
                    {
                        ltDocProps.Add("Company = " + doc.ExtendedFilePropertiesPart.Properties.Company.Text);
                    }

                    ltDocProps.Add("Document Security = " + GetDocSecurity(doc.ExtendedFilePropertiesPart.Properties.DocumentSecurity.Text));

                    if (doc.ExtendedFilePropertiesPart.Properties.HyperlinksChanged is not null)
                    {
                        ltDocProps.Add("Hyperlinks Changed = " + doc.ExtendedFilePropertiesPart.Properties.HyperlinksChanged.Text);
                    }

                    if (doc.ExtendedFilePropertiesPart.Properties.Lines is not null)
                    {
                        ltDocProps.Add("Lines = " + doc.ExtendedFilePropertiesPart.Properties.Lines.Text);
                    }

                    if (doc.ExtendedFilePropertiesPart.Properties.LinksUpToDate is not null)
                    {
                        ltDocProps.Add("Links Up To Date = " + doc.ExtendedFilePropertiesPart.Properties.LinksUpToDate.Text);
                    }

                    if (doc.ExtendedFilePropertiesPart.Properties.Paragraphs is not null)
                    {
                        ltDocProps.Add("Paragraphs = " + doc.ExtendedFilePropertiesPart.Properties.Paragraphs.Text);
                    }

                    if (doc.ExtendedFilePropertiesPart.Properties.ScaleCrop is not null)
                    {
                        ltDocProps.Add("Scale Crop = " + doc.ExtendedFilePropertiesPart.Properties.ScaleCrop.Text);
                    }

                    if (doc.ExtendedFilePropertiesPart.Properties.SharedDocument is not null)
                    {
                        ltDocProps.Add("Shared Document = " + doc.ExtendedFilePropertiesPart.Properties.SharedDocument.Text);
                    }

                    if (doc.ExtendedFilePropertiesPart.Properties.Template is not null)
                    {
                        ltDocProps.Add("Template = " + doc.ExtendedFilePropertiesPart.Properties.Template.Text);
                    }

                    if (doc.ExtendedFilePropertiesPart.Properties.TotalTime is not null)
                    {
                        ltDocProps.Add("Total Time = " + doc.ExtendedFilePropertiesPart.Properties.TotalTime.Text);
                    }

                    if (doc.ExtendedFilePropertiesPart.Properties.Words is not null)
                    {
                        ltDocProps.Add("Words = " + Strings.wEqualSign + doc.ExtendedFilePropertiesPart.Properties.Words.Text);
                    }

                    if (doc.ExtendedFilePropertiesPart.Properties.Pages is not null)
                    {
                        ltDocProps.Add("Pages = " + doc.ExtendedFilePropertiesPart.Properties.Pages.Text);
                    }
                }

                if (docSettingsPart is not null)
                {
                    Settings settings = docSettingsPart.Settings;
                    foreach (var setting in settings)
                    {
                        if (setting.LocalName == "compat")
                        {
                            int settingCount = setting.Count();
                            int settingIndex = 0;

                            do
                            {
                                if (setting.ElementAt(settingIndex).LocalName != "compatSetting")
                                {
                                    if (setting.ElementAt(0).InnerText != string.Empty)
                                    {
                                        compatList.Add(setting.ElementAt(0).LocalName + Strings.wColon + setting.ElementAt(0).InnerText);
                                    }
                                    settingIndex++;
                                }
                                else
                                {
                                    CompatibilitySetting cs = (CompatibilitySetting)setting.ElementAt(settingIndex);
                                    if (cs.Name == "compatibilityMode")
                                    {
                                        string compatModeVersion = string.Empty;

                                        if (cs.Val == "11")
                                        {
                                            compatModeVersion = " (Word 2003)";
                                        }
                                        else if (cs.Val == "12")
                                        {
                                            compatModeVersion = " (Word 2007)";
                                        }
                                        else if (cs.Val == "14")
                                        {
                                            compatModeVersion = " (Word 2010)";
                                        }
                                        else if (cs.Val == "15")
                                        {
                                            compatModeVersion = " (Word 2013)";
                                        }
                                        else
                                        {
                                            compatModeVersion = " (Unknown Version)";
                                        }

                                        compatList.Add(cs.Name + Strings.wColon + cs.Val + compatModeVersion);
                                        settingIndex++;
                                    }
                                    else
                                    {
                                        compatList.Add(cs.Name + Strings.wColon + cs.Val);
                                        settingIndex++;
                                    }
                                }
                            } while (settingIndex < settingCount);

                            compatList.Add(string.Empty);
                        }
                        else
                        {
                            XmlDocument xDoc = new XmlDocument();
                            xDoc.LoadXml(setting.OuterXml);

                            foreach (XmlElement xe in xDoc.ChildNodes)
                            {
                                sb.Clear();
                                if (xe.Attributes.Count > 1)
                                {
                                    sb.Append(xe.Name + Strings.wColon);
                                    foreach (XmlAttribute xa in xe.Attributes)
                                    {
                                        if (!(xa.LocalName == "w" || xa.LocalName == "m" || xa.LocalName == "w14" || xa.LocalName == "w15" || xa.LocalName == "w16"))
                                        {
                                            if (!xa.Value.StartsWith("http"))
                                            {
                                                if (xa.LocalName == "val")
                                                {
                                                    sb.Append(xa.Value);
                                                }
                                                else
                                                {
                                                    sb.Append(xa.LocalName + Strings.wColon + xa.Value + Strings.wSpaceChar);
                                                }
                                            }
                                        }
                                    }

                                    settingList.Add(sb.ToString());
                                }
                                else if (xe.Name == "w:docVars")
                                {
                                    foreach (XmlNode cNode in xe.ChildNodes)
                                    {
                                        settingList.Add(cNode.LocalName + Strings.wColon + cNode.OuterXml);
                                    }
                                }
                                else if (xe.Name == "w:rsids" && Properties.Settings.Default.ListRsids == true)
                                {
                                    rsidList.Add(xe.Name);
                                    foreach (XmlNode rNode in xe.ChildNodes)
                                    {
                                        foreach (XmlAttribute xa in rNode.Attributes)
                                        {
                                            rsidList.Add(rNode.Name + Strings.wColon + xa.Value);
                                        }
                                    }
                                }
                                else if (xe.Name == "m:mathPr")
                                {
                                    foreach (XmlNode rNode in xe.ChildNodes)
                                    {
                                        foreach (XmlAttribute xa in rNode.Attributes)
                                        {
                                            mathPrList.Add(rNode.Name + Strings.wColon + xa.Value);
                                        }
                                    }
                                }
                                else
                                {
                                    settingList.Add(xe.Name + Strings.wColon + xe.InnerText);
                                }
                            }
                        }
                    }
                }
            }

            return ltDocProps.Concat(compatList).Concat(settingList).Concat(rsidList).Concat(mathPrList).ToList();
        }

        public static List<string> LstFootnotes(string fPath)
        {
            List<string> ltFootnotes = new List<string>();

            using (WordprocessingDocument doc = WordprocessingDocument.Open(fPath, false))
            {
                FootnotesPart footnotePart = doc.MainDocumentPart.FootnotesPart;
                if (footnotePart is not null)
                {
                    int count = 0;
                    foreach (Footnote fn in footnotePart.Footnotes)
                    {
                        if (fn.InnerText != string.Empty)
                        {
                            count++;
                            ltFootnotes.Add(count + Strings.wPeriod + fn.InnerText);
                        }
                    }
                }
            }

            return ltFootnotes;
        }

        public static List<string> LstEndnotes(string fPath)
        {
            List<string> ltEndnotes = new List<string>();

            using (WordprocessingDocument doc = WordprocessingDocument.Open(fPath, false))
            {
                EndnotesPart endnotePart = doc.MainDocumentPart.EndnotesPart;
                if (endnotePart is not null)
                {
                    int count = 0;
                    foreach (Endnote en in endnotePart.Endnotes)
                    {
                        if (en.InnerText != string.Empty)
                        {
                            count++;
                            ltEndnotes.Add(count + Strings.wPeriod + en.InnerText);
                        }
                    }
                }
            }

            return ltEndnotes;
        }

        public static List<string> LstFonts(string fPath)
        {
            List<string> ltFonts = new List<string>();

            int count = 0;

            using (WordprocessingDocument doc = WordprocessingDocument.Open(fPath, false))
            {
                foreach (Font ft in doc.MainDocumentPart.FontTablePart.Fonts)
                {
                    count++;
                    ltFonts.Add(count + Strings.wPeriod + ft.Name);
                }
            }

            return ltFonts;
        }

        public static List<string> LstRunFonts(string fPath)
        {
            List<string> ltRunFonts = new List<string>();

            using (WordprocessingDocument doc = WordprocessingDocument.Open(fPath, false))
            {
                // loop each paragraph and get the run props
                // display props for each run
            }

            return ltRunFonts;
        }

        public static List<string> LstListTemplates(string fPath, bool onlyReturnUnused)
        {
            List<string> ltList = new List<string>();

            // global numid lists
            List<int> oNumIdList = new List<int>();
            List<int> aNumIdList = new List<int>();
            List<int> numIdList = new List<int>();
            List<string> unusedListTemplates = new List<string>();

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(fPath, false))
            {
                MainDocumentPart mainPart = myDoc.MainDocumentPart;
                NumberingDefinitionsPart numPart = mainPart.NumberingDefinitionsPart;
                StyleDefinitionsPart stylePart = mainPart.StyleDefinitionsPart;

                // Loop each paragraph, get the NumberingId and add it to the array
                foreach (OpenXmlElement el in mainPart.Document.Descendants<Paragraph>())
                {
                    if (el.Descendants<NumberingId>().Any())
                    {
                        foreach (NumberingId pNumId in el.Descendants<NumberingId>())
                        {
                            numIdList.Add(pNumId.Val);
                        }
                    }
                }

                // Loop each header, get the NumId and add it to the array
                foreach (HeaderPart hdrPart in mainPart.HeaderParts)
                {
                    foreach (OpenXmlElement el in hdrPart.Header.Elements())
                    {
                        foreach (NumberingId hNumId in el.Descendants<NumberingId>())
                        {
                            numIdList.Add(hNumId.Val);
                        }
                    }
                }

                // Loop each footer, get the NumId and add it to the array
                foreach (FooterPart ftrPart in mainPart.FooterParts)
                {
                    foreach (OpenXmlElement el in ftrPart.Footer.Elements())
                    {
                        foreach (NumberingId fNumdId in el.Descendants<NumberingId>())
                        {
                            numIdList.Add(fNumdId.Val);
                        }
                    }
                }

                // Loop through each style in document and get NumId
                foreach (OpenXmlElement el in stylePart.Styles.Elements())
                {
                    try
                    {
                        string styleEl = el.GetAttribute("styleId", Strings.wordMainAttributeNamespace).Value;
                        int pStyle = ParagraphsByStyleName(mainPart, styleEl).Count();

                        if (pStyle > 0)
                        {
                            foreach (NumberingId sEl in el.Descendants<NumberingId>())
                            {
                                numIdList.Add(sEl.Val);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Not all style elements have a styleID, so just skip these scenarios
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnListTemplates_Click : " + ex.Message);
                    }
                }

                // remove dupes
                numIdList = numIdList.Distinct().ToList();

                // display the active document lists
                ltList.Add("Active List Templates in this document:");

                // if we don't have any active templates, just continue checking for orphaned
                if (numIdList.Count == 0)
                {
                    ltList.Clear();
                }

                // since we have lists, display them
                int count = 0;
                foreach (object item in numIdList)
                {
                    count++;
                    ltList.Add(count + Strings.wNumId + item);

                    // Word is limited to 2047 total active lists in a document
                    if (count == 2047)
                    {
                        ltList.Add("## You have too many lists in this file. Word will only display up to 2047 lists. ##");
                    }
                }

                // Loop through each AbstractNumId
                ltList.Add(string.Empty);
                ltList.Add("All List Templates in document:");
                int aCount = 0;

                if (numPart is not null)
                {
                    foreach (OpenXmlElement el in numPart.Numbering.Elements())
                    {
                        foreach (AbstractNumId aNumId in el.Descendants<AbstractNumId>())
                        {
                            string strNumId = el.GetAttribute("numId", Strings.wordMainAttributeNamespace).Value;
                            aNumIdList.Add(int.Parse(strNumId));
                            aCount++;
                            ltList.Add(aCount + Strings.wNumId + strNumId);
                        }
                    }
                }
                else
                {
                    ltList.Add(Strings.wNone);
                }

                // get the unused list templates
                oNumIdList = OrphanedListTemplates(numIdList, aNumIdList);

                ltList.Add(string.Empty);
                ltList.Add("Orphaned List Templates:");
                
                if (oNumIdList.Count > 0)
                {
                    int oCount = 0;
                    foreach (object item in oNumIdList)
                    {
                        oCount++;
                        ltList.Add(oCount + Strings.wNumId + item);

                        if (onlyReturnUnused)
                        {
                            unusedListTemplates.Add(item.ToString());
                        }
                    }
                }
                else
                {
                    ltList.Add(Strings.wNone);
                }

            }

            if (onlyReturnUnused)
            {
                return unusedListTemplates;
            }
            else
            {
                return ltList;
            }
        }

        public static List<string> LstHyperlinks(string fPath)
        {
            List<string> hlinkList = new List<string>();

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(fPath, false))
            {
                int count = 0;

                IEnumerable<Hyperlink> hLinks = myDoc.MainDocumentPart.Document.Descendants<Hyperlink>();

                // handle if no links are found
                if (!myDoc.MainDocumentPart.HyperlinkRelationships.Any() && !myDoc.MainDocumentPart.RootElement.Descendants<FieldCode>().Any() && !hLinks.Any())
                {
                    return hlinkList;
                }
                else
                {
                    // loop through regular hyperlinks
                    foreach (Hyperlink h in hLinks)
                    {
                        count++;

                        string hRelUri = null;

                        // then check for hyperlinks relationships
                        foreach (HyperlinkRelationship hRel in myDoc.MainDocumentPart.HyperlinkRelationships)
                        {
                            if (h.Id == hRel.Id)
                            {
                                hRelUri = hRel.Uri.ToString();
                            }
                        }

                        hlinkList.Add(count + Strings.wPeriod + h.InnerText + " Uri = " + hRelUri);
                    }

                    // now we need to check for field hyperlinks
                    foreach (var field in myDoc.MainDocumentPart.RootElement.Descendants<FieldCode>())
                    {
                        string fldText;
                        if (field.InnerText.StartsWith(" HYPERLINK"))
                        {
                            count++;
                            fldText = field.InnerText.Remove(0, 11);
                            hlinkList.Add(count + Strings.wPeriod + fldText);
                        }
                    }
                }
            }

            return hlinkList;
        }

        public static List<string> LstStyles(string fPath)
        {
            List<string> stylesList = new List<string>();

            XNamespace w = Strings.wordMainAttributeNamespace;
            XDocument xDoc = null;
            XDocument styleDoc = null;
            bool containStyle = false;
            bool styleInUse = false;
            int count = 0;

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(fPath, false))
            {
                MainDocumentPart mainPart = myDoc.MainDocumentPart;
                StyleDefinitionsPart stylePart = mainPart.StyleDefinitionsPart;

                stylesList.Add("# Style Summary #");

                try
                {
                    // loop the styles in style.xml
                    foreach (OpenXmlElement el in stylePart.Styles.Elements())
                    {
                        string sName = string.Empty;
                        string sType = string.Empty;

                        if (el.LocalName == "style")
                        {
                            Style s = (Style)el;
                            sName = s.StyleId;
                            sType = s.Type;
                            styleInUse = false;

                            int pStyleCount = ParagraphsByStyleName(mainPart, sName).Count();
                            if (sType == "paragraph")
                            {
                                if (pStyleCount > 0)
                                {
                                    count += 1;
                                    stylesList.Add(count + Strings.wPeriod + sName + Strings.wUsedIn + pStyleCount + " paragraphs");
                                    containStyle = true;
                                    styleInUse = true;
                                    continue;
                                }
                            }

                            int rStyleCount = RunsByStyleName(mainPart, sName).Count();
                            if (sType == "character")
                            {
                                if (rStyleCount > 0)
                                {
                                    count += 1;
                                    stylesList.Add(count + Strings.wPeriod + sName + Strings.wUsedIn + rStyleCount + " runs");
                                    containStyle = true;
                                    styleInUse = true;
                                    continue;
                                }
                            }

                            int tStyleCount = TablesByStyleName(mainPart, sName).Count();
                            if (sType == "table")
                            {
                                if (tStyleCount > 0)
                                {
                                    count += 1;
                                    stylesList.Add(count + Strings.wPeriod + sName + Strings.wUsedIn + tStyleCount + " tables");
                                    containStyle = true;
                                    styleInUse = true;
                                    continue;
                                }
                            }

                            if (styleInUse == false)
                            {
                                count += 1;
                                stylesList.Add(count + Strings.wPeriod + sName + " -> (Not Used)");
                            }

                            if (count == 4079)
                            {
                                stylesList.Add("WARNING: Max Count of Styles for a document is 4079");
                            }
                        }
                    }

                    // add latent style information
                    stylesList.Add(string.Empty);
                    stylesList.Add("# Latent Style Summary #");
                    foreach (LatentStyleExceptionInfo lex in stylePart.Styles.LatentStyles)
                    {
                        count += 1;
                        if (lex.UnhideWhenUsed is not null)
                        {
                            stylesList.Add(count + Strings.wPeriod + lex.Name + " (Hidden)");
                        }
                        else
                        {
                            stylesList.Add(count + Strings.wPeriod + lex.Name);
                        }
                    }

                }
                catch (NullReferenceException)
                {
                    stylesList.Add("** Missing StylesWithEffects part **");
                }
            }

            if (containStyle == false)
            {
                return stylesList;
            }
            else
            {
                // list the styles for paragraphs
                stylesList.Add(string.Empty);
                stylesList.Add("# List of paragraph styles #");
                count = 0;

                using (Package wdPackage = Package.Open(fPath, FileMode.Open, FileAccess.Read))
                {
                    PackageRelationship docPackageRelationship = wdPackage.GetRelationshipsByType(Strings.MainDocumentPartType).FirstOrDefault();
                    if (docPackageRelationship is not null)
                    {
                        Uri documentUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), docPackageRelationship.TargetUri);
                        PackagePart documentPart = wdPackage.GetPart(documentUri);

                        //  Load the document XML in the part into an XDocument instance.
                        xDoc = XDocument.Load(XmlReader.Create(documentPart.GetStream()));

                        //  Find the styles part. There will only be one.
                        PackageRelationship styleRelation = documentPart.GetRelationshipsByType(Strings.StyleDefsPartType).FirstOrDefault();
                        if (styleRelation is not null)
                        {
                            Uri styleUri = PackUriHelper.ResolvePartUri(documentUri, styleRelation.TargetUri);
                            PackagePart stylePart = wdPackage.GetPart(styleUri);

                            //  Load the style XML in the part into an XDocument instance.
                            styleDoc = XDocument.Load(XmlReader.Create(stylePart.GetStream()));
                        }
                    }
                }

                string defaultStyle = (string)(
                    from style in styleDoc.Root.Elements(w + "style")
                    where (string)style.Attribute(w + "type") == "paragraph" && (string)style.Attribute(w + "default") == "1"
                    select style
                ).First().Attribute(w + "styleId");

                // Find all paragraphs in the document.  
                var paragraphs =
                    from para in xDoc.Root.Element(w + "body").Descendants(w + "p")
                    let styleNode = para.Elements(w + "pPr").Elements(w + "pStyle").FirstOrDefault()
                    select new
                    {
                        ParagraphNode = para,
                        StyleName = styleNode is null ? defaultStyle : (string)styleNode.Attribute(w + "val")
                    };

                // Retrieve the text of each paragraph.  
                var paraWithText =
                    from para in paragraphs
                    select new
                    {
                        para.ParagraphNode,
                        para.StyleName,
                        Text = ParagraphText(para.ParagraphNode)
                    };

                foreach (var p in paraWithText)
                {
                    count++;
                    stylesList.Add(count + ". StyleName: " + p.StyleName + " Text: " + p.Text);
                }
            }

            return stylesList;
        }

        public static List<string> LstContentControls(string fPath)
        {
            List<string> ccList = new List<string>();

            using (WordprocessingDocument doc = WordprocessingDocument.Open(fPath, false))
            {
                int count = 0;

                foreach (var cc in doc.ContentControls())
                {
                    string ccType = string.Empty;
                    bool PropFound = false;
                    SdtProperties props = cc.Elements<SdtProperties>().FirstOrDefault();

                    // loop the properties and get the type
                    foreach (OpenXmlElement oxe in props.ChildElements)
                    {
                        if (oxe.GetType().Name == "SdtContentText")
                        {
                            ccType = "Plain Text";
                            PropFound = true;
                        }

                        if (oxe.GetType().Name == "SdtContentDropDownList")
                        {
                            ccType = "Drop Down List";
                            PropFound = true;
                        }

                        if (oxe.GetType().Name == "SdtContentDocPartList")
                        {
                            ccType = "Building Block Gallery";
                            PropFound = true;
                        }

                        if (oxe.GetType().Name == "SdtContentCheckBox")
                        {
                            ccType = "Check Box";
                            PropFound = true;
                        }

                        if (oxe.GetType().Name == "SdtContentPicture")
                        {
                            ccType = "Picture";
                            PropFound = true;
                        }

                        if (oxe.GetType().Name == "SdtContentComboBox")
                        {
                            ccType = "Combo Box";
                            PropFound = true;
                        }

                        if (oxe.GetType().Name == "SdtContentDate")
                        {
                            ccType = "Date Picker";
                            PropFound = true;
                        }

                        if (oxe.GetType().Name == "SdtRepeatedSection")
                        {
                            ccType = "Repeating Section";
                            PropFound = true;
                        }
                    }

                    // display the cc type
                    count++;
                    if (PropFound == true)
                    {
                        ccList.Add(count + Strings.wPeriod + ccType);
                    }
                    else
                    {
                        ccList.Add(count + Strings.wPeriod + "Rich Text");
                    }
                }
            }

            return ccList;
        }

        /// <summary>
        /// get a list of all authors in each story of the document
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        public static List<string> GetAllAuthors(Document doc)
        {
            bool nullAuthor = false;
            List<string> allAuthorsInDocument = new List<string>();

            var paragraphChanged = doc.Descendants<ParagraphPropertiesChange>().ToList();
            var runChanged = doc.Descendants<RunPropertiesChange>().ToList();
            var deleted = doc.Descendants<DeletedRun>().ToList();
            var deletedParagraph = doc.Descendants<Deleted>().ToList();
            var inserted = doc.Descendants<InsertedRun>().ToList();

            // loop through each revision and catalog the authors
            // some authors show up as null, check and ignore
            foreach (ParagraphPropertiesChange ppc in paragraphChanged)
            {
                if (ppc.Author is not null)
                {
                    allAuthorsInDocument.Add(ppc.Author);
                }
                else
                {
                    nullAuthor = true;
                }
            }

            foreach (RunPropertiesChange rpc in runChanged)
            {
                if (rpc.Author is not null)
                {
                    allAuthorsInDocument.Add(rpc.Author);
                }
                else
                {
                    nullAuthor = true;
                }
            }

            foreach (DeletedRun dr in deleted)
            {
                if (dr.Author is not null)
                {
                    allAuthorsInDocument.Add(dr.Author);
                }
                else
                {
                    nullAuthor = true;
                }
            }

            foreach (Deleted d in deletedParagraph)
            {
                if (d.Author is not null)
                {
                    allAuthorsInDocument.Add(d.Author);
                }
                else
                {
                    nullAuthor = true;
                }
            }

            foreach (InsertedRun ir in inserted)
            {
                if (ir.Author is not null)
                {
                    allAuthorsInDocument.Add(ir.Author);
                }
                else
                {
                    nullAuthor = true;
                }
            }

            // log if we have a null author, not sure how this happens yet
            if (nullAuthor)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "Null Author Found");
            }

            List<string> distinctAuthors = allAuthorsInDocument.Distinct().ToList();

            return distinctAuthors;
        }

        /// <summary>
        /// Set the font for a text run.
        /// </summary>
        /// <param name="fileName"></param>        
        public static void SetRunFont(string fileName)
        {
            // Open a Wordprocessing document for editing.
            using (WordprocessingDocument package = WordprocessingDocument.Open(fileName, true))
            {
                // Set the font to Arial to the first Run.
                // Use an object initializer for RunProperties and rPr.
                RunProperties rPr = new RunProperties(
                    new RunFonts()
                    {
                        Ascii = "Arial"
                    });

                Run r = package.MainDocumentPart.Document.Descendants<Run>().First();
                r.PrependChild<RunProperties>(rPr);

                // Save changes to the MainDocumentPart part.
                package.MainDocumentPart.Document.Save();
            }
        }

        // Creates an Endnotes instance and adds its children.
        public static Endnotes CreateDefaultEndnotes()
        {
            Endnotes endnotes1 = new Endnotes() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex wp14" } };
            endnotes1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            endnotes1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            endnotes1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            endnotes1.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            endnotes1.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            endnotes1.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            endnotes1.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            endnotes1.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            endnotes1.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            endnotes1.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            endnotes1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            endnotes1.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            endnotes1.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            endnotes1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            endnotes1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            endnotes1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            endnotes1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            endnotes1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            endnotes1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            endnotes1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            endnotes1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            endnotes1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            endnotes1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            endnotes1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            endnotes1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            endnotes1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            endnotes1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            endnotes1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            endnotes1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            endnotes1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            endnotes1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Endnote endnote1 = new Endnote() { Type = FootnoteEndnoteValues.Separator, Id = -1 };
            string rsidAdditionGuid = Office.CreateNewRsid();
            string rsidPropsGuid = Office.CreateNewRsid();

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = rsidAdditionGuid, RsidParagraphProperties = rsidPropsGuid, RsidRunAdditionDefault = rsidAdditionGuid, ParagraphId = Office.CreateNewRsid(), TextId = "77777777" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            paragraphProperties1.Append(spacingBetweenLines1);

            Run run1 = new Run();
            SeparatorMark separatorMark1 = new SeparatorMark();

            run1.Append(separatorMark1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            endnote1.Append(paragraph1);

            Endnote endnote2 = new Endnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = rsidAdditionGuid, RsidParagraphProperties = rsidPropsGuid, RsidRunAdditionDefault = rsidAdditionGuid, ParagraphId = Office.CreateNewRsid(), TextId = "77777777" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            paragraphProperties2.Append(spacingBetweenLines2);

            Run run2 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark1 = new ContinuationSeparatorMark();

            run2.Append(continuationSeparatorMark1);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run2);

            endnote2.Append(paragraph2);

            endnotes1.Append(endnote1);
            endnotes1.Append(endnote2);
            return endnotes1;
        }

        /// <summary>
        /// Function to check if a part is null
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="partType"></param>
        /// <returns></returns>
        public static bool IsPartNull(WordprocessingDocument doc, string partType)
        {
            try
            {
                if (partType == "DeletedRun")
                {
                    var deleted = doc.MainDocumentPart.Document.Descendants<DeletedRun>().ToList();
                }
                else if (partType == "Table")
                {
                    var tbls = doc.MainDocumentPart.Document.Descendants<Table>().ToList();
                }

                return false;
            }
            catch (Exception)
            {
                return true;
            }
        }

        /// <summary>
        /// function to create the style linked list chain
        /// </summary>
        /// <param name="sdp">style def part</param>
        /// <param name="prevStyle">name of the previous style to compare with basedon</param>
        /// <param name="currentStyleChain">string to hold the sequence of styles starting with the base style</param>
        /// <returns></returns>
        public static StringBuilder GetBasedOnStyleChain(StyleDefinitionsPart sdp, string prevStyle, StringBuilder currentStyleChain)
        {
            foreach (OpenXmlElement tempEl in sdp.Styles.Elements())
            {
                if (tempEl.LocalName == "style")
                {
                    Style tempStyle = (Style)tempEl;
                    if (tempStyle.BasedOn is not null)
                    {
                        if (tempStyle.BasedOn.Val == prevStyle)
                        {
                            currentStyleChain.Append(Strings.wArrowOnly + tempStyle.StyleId);
                            GetBasedOnStyleChain(sdp, tempStyle.StyleId, currentStyleChain);
                        }
                    }
                }
            }

            return currentStyleChain;
        }

        public static List<int> OrphanedListTemplates(List<int> usedNumIdList, List<int> docNumIdList)
        {
            var copyOfDocNumId = new List<int>(docNumIdList);

            foreach (var p in usedNumIdList)
            {
                copyOfDocNumId.Remove(p);
            }

            return copyOfDocNumId;
        }

        public static string ParagraphText(XElement e)
        {
            XNamespace w = e.Name.Namespace;
            return e
                   .Elements(w + "r")
                   .Elements(w + "t")
                   .StringConcatenate(element => (string)element);
        }

        public static string StringConcatenate(this IEnumerable<string> source)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string s in source)
                sb.Append(s);
            return sb.ToString();
        }

        public static string StringConcatenate<T>(this IEnumerable<T> source, Func<T, string> func)
        {
            StringBuilder sb = new StringBuilder();
            foreach (T item in source)
                sb.Append(func(item));
            return sb.ToString();
        }

        public static string StringConcatenate(this IEnumerable<string> source, string separator)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string s in source)
                sb.Append(s).Append(separator);
            return sb.ToString();
        }

        public static string StringConcatenate<T>(this IEnumerable<T> source, Func<T, string> func, string separator)
        {
            StringBuilder sb = new StringBuilder();
            foreach (T item in source)
                sb.Append(func(item)).Append(separator);
            return sb.ToString();
        }

        // Return true if the style id is in the document, false otherwise.
        public static bool IsStyleIdInDocument(WordprocessingDocument doc, string styleid)
        {
            // Get access to the Styles element for this document.
            Styles s = doc.MainDocumentPart.StyleDefinitionsPart.Styles;

            // Check that there are styles and how many.
            int n = s.Elements<Style>().Count();
            if (n == 0)
                return false;

            // Look for a match on styleid.
            Style style = s.Elements<Style>()
                .Where(st => (st.StyleId == styleid) && (st.Type == StyleValues.Paragraph))
                .FirstOrDefault();
            if (style is null)
                return false;

            return true;
        }

        // Return styleid that matches the styleName, or null when there's no match.
        public static string GetStyleIdFromStyleName(WordprocessingDocument doc, string styleName)
        {
            StyleDefinitionsPart stylePart = doc.MainDocumentPart.StyleDefinitionsPart;
            string styleId = stylePart.Styles.Descendants<StyleName>()
                .Where(s => s.Val.Value.Equals(styleName) &&
                    (((Style)s.Parent).Type == StyleValues.Paragraph))
                .Select(n => ((Style)n.Parent).StyleId).FirstOrDefault();
            return styleId;
        }

        public static IEnumerable<OpenXmlElement> ContentControls(this OpenXmlPart part)
        {
            return part.RootElement.Descendants().Where(e => e is SdtBlock || e is SdtRun);
        }

        public static IEnumerable<OpenXmlElement> ContentControls(this WordprocessingDocument doc)
        {
            foreach (var cc in doc.MainDocumentPart.ContentControls())
                yield return cc;
            foreach (var header in doc.MainDocumentPart.HeaderParts)
                foreach (var cc in header.ContentControls())
                    yield return cc;
            foreach (var footer in doc.MainDocumentPart.FooterParts)
                foreach (var cc in footer.ContentControls())
                    yield return cc;
            if (doc.MainDocumentPart.FootnotesPart is not null)
                foreach (var cc in doc.MainDocumentPart.FootnotesPart.ContentControls())
                    yield return cc;
            if (doc.MainDocumentPart.EndnotesPart is not null)
                foreach (var cc in doc.MainDocumentPart.EndnotesPart.ContentControls())
                    yield return cc;
        }

        public static XDocument GetXDocument(this OpenXmlPart part)
        {
            XDocument xdoc = part.Annotation<XDocument>();
            if (xdoc is not null)
            {
                return xdoc;
            }

            using (StreamReader sr = new StreamReader(part.GetStream()))
            using (XmlReader xr = XmlReader.Create(sr))
            {
                xdoc = XDocument.Load(xr);
            }

            part.AddAnnotation(xdoc);

            return xdoc;
        }

        public static bool HasPersonalInfo(WordprocessingDocument document)
        {
            // check for company name from /docProps/app.xml
            XNamespace x = Strings.OfficeExtendedProps;
            OpenXmlPart extendedFilePropertiesPart = document.ExtendedFilePropertiesPart;
            XDocument extendedFilePropertiesXDoc = extendedFilePropertiesPart.GetXDocument();
            string company = extendedFilePropertiesXDoc.Elements(x + Strings.wProperties).Elements(x + Strings.wCompany)
                .Select(e => (string)e).Aggregate(string.Empty, (s, i) => s + i);

            if (company.Length > 0)
            {
                return true;
            }

            // check for dc:creator, cp:lastModifiedBy from /docProps/core.xml
            XNamespace dc = Strings.DcElements;
            XNamespace cp = Strings.OfficeCoreProps;
            OpenXmlPart coreFilePropertiesPart = document.CoreFilePropertiesPart;
            XDocument coreFilePropertiesXDoc = coreFilePropertiesPart.GetXDocument();
            string creator = coreFilePropertiesXDoc.Elements(cp + Strings.wCoreProperties).Elements(dc + Strings.wCreator)
                .Select(e => (string)e).Aggregate(string.Empty, (s, i) => s + i);

            if (creator.Length > 0)
            {
                return true;
            }

            string lastModifiedBy = coreFilePropertiesXDoc.Elements(cp + Strings.wCoreProperties).Elements(cp + Strings.wLastModifiedBy)
                .Select(e => (string)e).Aggregate(string.Empty, (s, i) => s + i);

            if (lastModifiedBy.Length > 0)
            {
                return true;
            }

            // check for nonexistence of removePersonalInformation and removeDateAndTime
            XNamespace w = Strings.wordMainAttributeNamespace;
            OpenXmlPart documentSettingsPart = document.MainDocumentPart.DocumentSettingsPart;
            XDocument documentSettingsXDoc = documentSettingsPart.GetXDocument();
            XElement settings = documentSettingsXDoc.Root;

            if (settings.Element(w + Strings.wRemovePI) is null)
            {
                return true;
            }

            if (settings.Element(w + Strings.wRemoveDateTime) is null)
            {
                return true;
            }

            return false;
        }

        public static bool RemovePersonalInfo(WordprocessingDocument document)
        {
            bool isFixed = false;

            // remove the company name from /docProps/app.xml
            // set TotalTime to "0"
            XNamespace x = Strings.OfficeExtendedProps;
            OpenXmlPart extendedFilePropertiesPart = document.ExtendedFilePropertiesPart;
            XDocument extendedFilePropertiesXDoc = extendedFilePropertiesPart.GetXDocument();
            extendedFilePropertiesXDoc.Elements(x + Strings.wProperties).Elements(x + Strings.wCompany).Remove();
            XElement totalTime = extendedFilePropertiesXDoc.Elements(x + Strings.wProperties).Elements(x + "TotalTime").FirstOrDefault();
            if (totalTime is not null)
            {
                totalTime.Value = "0";
            }

            using (XmlWriter xw = XmlWriter.Create(extendedFilePropertiesPart.GetStream(FileMode.Create, FileAccess.Write)))
            {
                extendedFilePropertiesXDoc.Save(xw);
                isFixed = true;
            }

            // remove the values of dc:creator, cp:lastModifiedBy from /docProps/core.xml
            // set cp:revision to "1"
            XNamespace dc = Strings.DcElements;
            XNamespace cp = Strings.OfficeCoreProps;
            OpenXmlPart coreFilePropertiesPart = document.CoreFilePropertiesPart;
            XDocument coreFilePropertiesXDoc = coreFilePropertiesPart.GetXDocument();
            foreach (var textNode in coreFilePropertiesXDoc.Elements(cp + Strings.wCoreProperties)
                                                           .Elements(dc + Strings.wCreator)
                                                           .Nodes()
                                                           .OfType<XText>())
            {
                textNode.Value = string.Empty;
            }

            foreach (var textNode in coreFilePropertiesXDoc.Elements(cp + Strings.wCoreProperties)
                                                           .Elements(cp + Strings.wLastModifiedBy)
                                                           .Nodes()
                                                           .OfType<XText>())
            {
                textNode.Value = string.Empty;
            }

            XElement revision = coreFilePropertiesXDoc.Elements(cp + Strings.wCoreProperties).Elements(cp + "revision").FirstOrDefault();
            if (revision is not null)
            {
                revision.Value = "1";
            }

            using (XmlWriter xw = XmlWriter.Create(coreFilePropertiesPart.GetStream(FileMode.Create, FileAccess.Write)))
            {
                coreFilePropertiesXDoc.Save(xw);
                isFixed = true;
            }

            // add w:removePersonalInformation, w:removeDateAndTime to /word/settings.xml
            XNamespace w = Strings.wordMainAttributeNamespace;
            OpenXmlPart documentSettingsPart = document.MainDocumentPart.DocumentSettingsPart;
            XDocument documentSettingsXDoc = documentSettingsPart.GetXDocument();

            // add the new elements in the right position.  Add them after the following three elements
            // (which may or may not exist in the xml document).
            XElement settings = documentSettingsXDoc.Root;
            XElement lastOfTop3 = settings.Elements()
                .Where(e => e.Name == w + "writeProtection" ||
                    e.Name == w + "view" ||
                    e.Name == w + "zoom")
                .InDocumentOrder()
                .LastOrDefault();
            if (lastOfTop3 is null)
            {
                // none of those three exist, so add as first children of the root element
                settings.AddFirst(
                    settings.Elements(w + Strings.wRemovePI).Any() ?
                        null :
                        new XElement(w + Strings.wRemovePI),
                    settings.Elements(w + Strings.wRemoveDateTime).Any() ?
                        null :
                        new XElement(w + Strings.wRemoveDateTime)
                );
            }
            else
            {
                // one of those three exist, so add after the last one
                lastOfTop3.AddAfterSelf(
                    settings.Elements(w + Strings.wRemovePI).Any() ?
                        null :
                        new XElement(w + Strings.wRemovePI),
                    settings.Elements(w + Strings.wRemoveDateTime).Any() ?
                        null :
                        new XElement(w + Strings.wRemoveDateTime)
                );
            }
            using (XmlWriter xw = XmlWriter.Create(documentSettingsPart.GetStream(FileMode.Create, FileAccess.Write)))
            {
                documentSettingsXDoc.Save(xw);
                isFixed = true;
            }

            return isFixed;
        }

        private static string GetStyleIdFromStyleName(MainDocumentPart mainPart, string styleName)
        {
            StyleDefinitionsPart stylePart = mainPart.StyleDefinitionsPart;
            string styleId = stylePart.Styles.Descendants<StyleName>()
                .Where(s => s.Val.Value.Equals(styleName))
                .Select(n => ((Style)n.Parent).StyleId).FirstOrDefault();
            return styleId ?? styleName;
        }

        public static IEnumerable<Paragraph> ParagraphsByStyleName(this MainDocumentPart mainPart, string styleName)
        {
            string styleId = GetStyleIdFromStyleName(mainPart, styleName);
            IEnumerable<Paragraph> paraList = mainPart.Document.Descendants<Paragraph>()
                .Where(p => IsParagraphInStyle(p, styleId));
            return paraList;
        }

        public static IEnumerable<Paragraph> ParagraphsByStyleId(this MainDocumentPart mainPart, string styleId)
        {
            IEnumerable<Paragraph> paraList = mainPart.Document.Descendants<Paragraph>()
                .Where(p => IsParagraphInStyle(p, styleId));
            return paraList;
        }

        private static bool IsParagraphInStyle(Paragraph p, string styleId)
        {
            ParagraphProperties pPr = p.GetFirstChild<ParagraphProperties>();
            if (pPr is not null)
            {
                ParagraphStyleId paraStyle = pPr.ParagraphStyleId;

                if (paraStyle is not null)
                {
                    return paraStyle.Val.Value.Equals(styleId);
                }
                return false;
            }
            else if (pPr is null && styleId == "Normal")
            {
                // typically, if the pPr is null the style is Normal
                return true;
            }
            else
            {
                return false;
            }
        }

        public static IEnumerable<Run> RunsByStyleName(this MainDocumentPart mainPart, string styleName)
        {
            string styleId = GetStyleIdFromStyleName(mainPart, styleName);
            IEnumerable<Run> runList = mainPart.Document.Descendants<Run>()
                .Where(r => IsRunInStyle(r, styleId));
            return runList;
        }

        public static IEnumerable<Run> RunsByStyleId(this MainDocumentPart mainPart, string styleId)
        {
            IEnumerable<Run> runList = mainPart.Document.Descendants<Run>()
                .Where(r => IsRunInStyle(r, styleId));
            return runList;
        }

        private static bool IsRunInStyle(Run r, string styleId)
        {
            RunProperties rPr = r.GetFirstChild<RunProperties>();

            if (rPr is not null)
            {
                RunStyle runStyle = rPr.RunStyle;
                if (runStyle is not null)
                {
                    return runStyle.Val.Value.Equals(styleId);
                }
            }
            return false;
        }

        public static IEnumerable<Table> TablesByStyleName(this MainDocumentPart mainPart, string styleName)
        {
            string styleId = GetStyleIdFromStyleName(mainPart, styleName);
            IEnumerable<Table> tableList = mainPart.Document.Descendants<Table>()
                .Where(t => IsTableInStyle(t, styleId));
            return tableList;
        }

        public static IEnumerable<Table> TablesByStyleId(this MainDocumentPart mainPart, string styleId)
        {
            IEnumerable<Table> tableList = mainPart.Document.Descendants<Table>()
                .Where(t => IsTableInStyle(t, styleId));
            return tableList;
        }

        private static bool IsTableInStyle(Table tbl, string styleId)
        {
            TableProperties tblPr = tbl.GetFirstChild<TableProperties>();

            if (tblPr is not null)
            {
                TableStyle tblStyle = tblPr.TableStyle;

                if (tblStyle is not null)
                {
                    return tblStyle.Val.Value.Equals(styleId);
                }
            }
            return false;
        }
    }
}
