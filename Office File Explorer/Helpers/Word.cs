// .NET refs
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

// open xml sdk refs
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// shortcut namespace refs
using O = DocumentFormat.OpenXml;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using System.Text;

namespace Office_File_Explorer.Helpers
{
    class Word
    {
        public static bool fSuccess;

        /// <summary>
        /// Given a document name and an author name, accept all revisions by the specified author. 
        /// Pass an empty string for the author to accept all revisions.
        /// </summary>
        /// <param name="docName"></param>
        /// <param name="authorName"></param>
        public static bool AcceptAllRevisions(string docName, string authorName)
        {
            fSuccess = false;

            using (WordprocessingDocument document = WordprocessingDocument.Open(docName, true))
            {
                Document doc = document.MainDocumentPart.Document;
                var paragraphChanged = doc.Descendants<ParagraphPropertiesChange>().ToList();
                var runChanged = doc.Descendants<RunPropertiesChange>().ToList();
                var deleted = doc.Descendants<DeletedRun>().ToList();
                var deletedParagraph = doc.Descendants<Deleted>().ToList();
                var inserted = doc.Descendants<InsertedRun>().ToList();

                if (authorName == "* All Authors *")
                {
                    List<string> temp = new List<string>();
                    temp = WordExtensions.GetAllAuthors(document.MainDocumentPart.Document);

                    // create a temp list for each author so we can loop the changes individually and list them
                    foreach (string s in temp)
                    {
                        var tempParagraphChanged = paragraphChanged.Where(item => item.Author == s).ToList();
                        var tempRunChanged = runChanged.Where(item => item.Author == s).ToList();
                        var tempDeleted = deleted.Where(item => item.Author == s).ToList();
                        var tempInserted = inserted.Where(item => item.Author == s).ToList();
                        var tempDeletedParagraph = deletedParagraph.Where(item => item.Author == s).ToList();

                        foreach (var item in tempParagraphChanged)
                            item.Remove();

                        foreach (var item in tempDeletedParagraph)
                            item.Remove();

                        foreach (var item in tempRunChanged)
                            item.Remove();

                        foreach (var item in tempDeleted)
                            item.Remove();

                        foreach (var item in tempInserted)
                        {
                            if (item.Parent != null)
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
                            }
                        }
                    }
                    doc.Save();
                    fSuccess = true;
                }
                else
                {
                    // for single author, just loop that authors from the original list
                    if (!String.IsNullOrEmpty(authorName))
                    {
                        paragraphChanged = paragraphChanged.Where(item => item.Author == authorName).ToList();
                        runChanged = runChanged.Where(item => item.Author == authorName).ToList();
                        deleted = deleted.Where(item => item.Author == authorName).ToList();
                        inserted = inserted.Where(item => item.Author == authorName).ToList();
                        deletedParagraph = deletedParagraph.Where(item => item.Author == authorName).ToList();
                    }

                    foreach (var item in paragraphChanged)
                        item.Remove();

                    foreach (var item in deletedParagraph)
                        item.Remove();

                    foreach (var item in runChanged)
                        item.Remove();

                    foreach (var item in deleted)
                        item.Remove();

                    foreach (var item in inserted)
                    {
                        if (item.Parent != null)
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
                        }
                    }
                    doc.Save();
                    fSuccess = true;
                }
            }

            return fSuccess;
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
                            if (pgMar != null)
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
                    if ((item.Parent != null) &&
                      (item.Parent.Parent != null) &&
                      (item.Parent.Parent.Parent != null))
                    {
                        var topNode = item.Parent.Parent;
                        var topParentNode = item.Parent.Parent.Parent;
                        if (topParentNode != null)
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

                        StringBuilder baseStyleChain = WordExtensions.GetBasedOnStyleChain(stylePart, sBase, tempBaseStyleChain);

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

                            if (words.Count() > 0)
                            {
                                foreach (string w in words.Reverse())
                                {
                                    int pWStyleCount = WordExtensions.ParagraphsByStyleId(mainPart, w).Count();
                                    int rWStyleCount = WordExtensions.RunsByStyleId(mainPart, w).Count();
                                    int tWStyleCount = WordExtensions.TablesByStyleId(mainPart, w).Count();
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
                if (fnp != null)
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
                if (wdDoc.MainDocumentPart.GetPartsOfType<EndnotesPart>().Count() > 0)
                {
                    MainDocumentPart mainPart = wdDoc.MainDocumentPart;

                    var enr = mainPart.Document.Descendants<EndnoteReference>().ToArray();
                    foreach (var e in enr)
                    {
                        e.Remove();
                    }

                    EndnotesPart ep = mainPart.EndnotesPart;
                    ep.Endnotes = WordExtensions.CreateDefaultEndnotes();
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
                if (wdDoc.MainDocumentPart.GetPartsOfType<HeaderPart>().Count() > 0 ||
                  wdDoc.MainDocumentPart.GetPartsOfType<FooterPart>().Count() > 0)
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
            var ltFieldCodes = new List<string>();

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

        public static List<string> LstComments(string fPath)
        {
            var ltComments = new List<string>();

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
                int cRefCount = 0;

                cRefCount = crList.Count();

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
            var ltBookmarks = new List<string>();

            using (WordprocessingDocument package = WordprocessingDocument.Open(fPath, false))
            {
                IEnumerable<BookmarkStart> bkList = package.MainDocumentPart.Document.Descendants<BookmarkStart>();
                ltBookmarks.Add("** Document Bookmarks **");

                if (bkList.Count() > 0)
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
                            if (cElem != null && cElem.Parent != null && cElem.Parent.ToString().Contains(Strings.dfowSdt))
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

                if (package.MainDocumentPart.WordprocessingCommentsPart != null)
                {
                    if (package.MainDocumentPart.WordprocessingCommentsPart.Comments != null)
                    {
                        IEnumerable<BookmarkStart> bkCommentList = package.MainDocumentPart.WordprocessingCommentsPart.Comments.Descendants<BookmarkStart>();
                        int bkCommentCount = 0;

                        if (bkCommentList.Count() > 0)
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
            settingList.Add("---- Settings ---- ");

            using (WordprocessingDocument doc = WordprocessingDocument.Open(fPath, false))
            {
                DocumentSettingsPart docSettingsPart = doc.MainDocumentPart.DocumentSettingsPart;

                // get the standard file props
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
                ltDocProps.Add(string.Empty);

                // get the extended file props
                ltDocProps.Add("---- Extended File Properties ----");

                if (doc.ExtendedFilePropertiesPart != null)
                {
                    XmlDocument xmlProps = new XmlDocument();
                    xmlProps.Load(doc.ExtendedFilePropertiesPart.GetStream());
                    XmlNodeList exProps = xmlProps.GetElementsByTagName("Properties");

                    foreach (XmlNode xNode in exProps)
                    {
                        foreach (XmlElement xElement in xNode)
                        {
                            if (xElement.Name == Strings.wDocSecurity)
                            {
                                switch (xElement.InnerText)
                                {
                                    case "0":
                                        ltDocProps.Add(Strings.wDocSecurity + Strings.wColon + "None");
                                        break;
                                    case "1":
                                        ltDocProps.Add(Strings.wDocSecurity + Strings.wColon + "Password Protected");
                                        break;
                                    case "2":
                                        ltDocProps.Add(Strings.wDocSecurity + Strings.wColon + "Read-Only Recommended");
                                        break;
                                    case "4":
                                        ltDocProps.Add(Strings.wDocSecurity + Strings.wColon + "Read-Only Enforced");
                                        break;
                                    case "8":
                                        ltDocProps.Add(Strings.wDocSecurity + Strings.wColon + "Locked For Annotation");
                                        break;
                                    default:
                                        break;
                                }
                            }
                            else
                            {
                                ltDocProps.Add(xElement.Name + Strings.wColon + xElement.InnerText);
                            }
                        }
                    }
                }

                if (docSettingsPart != null)
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
                                else if (xe.Name == "w:rsids")
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
            var ltFootnotes = new List<string>();

            using (WordprocessingDocument doc = WordprocessingDocument.Open(fPath, false))
            {
                FootnotesPart footnotePart = doc.MainDocumentPart.FootnotesPart;
                if (footnotePart != null)
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
            var ltEndnotes = new List<string>();

            using (WordprocessingDocument doc = WordprocessingDocument.Open(fPath, false))
            {
                EndnotesPart endnotePart = doc.MainDocumentPart.EndnotesPart;
                if (endnotePart != null)
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
            var ltFonts = new List<string>();

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

        public static List<string> LstListTemplates(string fPath, bool onlyReturnUnused)
        {
            var ltList = new List<string>();

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
                    if (el.Descendants<NumberingId>().Count() > 0)
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
                        int pStyle = WordExtensions.ParagraphsByStyleName(mainPart, styleEl).Count();

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
                    return ltList;
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

                if (numPart != null)
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
                oNumIdList = WordExtensions.OrphanedListTemplates(numIdList, aNumIdList);

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
            var hlinkList = new List<string>();

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(fPath, false))
            {
                int count = 0;

                IEnumerable<Hyperlink> hLinks = myDoc.MainDocumentPart.Document.Descendants<O.Wordprocessing.Hyperlink>();

                // handle if no links are found
                if (myDoc.MainDocumentPart.HyperlinkRelationships.Count() == 0 && myDoc.MainDocumentPart.RootElement.Descendants<FieldCode>().Count() == 0 && hLinks.Count() == 0)
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
            var stylesList = new List<string>();

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

                            int pStyleCount = WordExtensions.ParagraphsByStyleName(mainPart, sName).Count();
                            if (sType == "paragraph")
                            {
                                if (pStyleCount > 0)
                                {
                                    count += 1;
                                    stylesList.Add(count + Strings.wPeriod + sName + " -> Used in " + pStyleCount + " paragraphs");
                                    containStyle = true;
                                    styleInUse = true;
                                    continue;
                                }
                            }

                            int rStyleCount = WordExtensions.RunsByStyleName(mainPart, sName).Count();
                            if (sType == "character")
                            {
                                if (rStyleCount > 0)
                                {
                                    count += 1;
                                    stylesList.Add(count + Strings.wPeriod + sName + " -> Used in " + rStyleCount + " runs");
                                    containStyle = true;
                                    styleInUse = true;
                                    continue;
                                }
                            }

                            int tStyleCount = WordExtensions.TablesByStyleName(mainPart, sName).Count();
                            if (sType == "table")
                            {
                                if (tStyleCount > 0)
                                {
                                    count += 1;
                                    stylesList.Add(count + Strings.wPeriod + sName + " -> Used in " + tStyleCount + " tables");
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
                        if (lex.UnhideWhenUsed != null)
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
                    if (docPackageRelationship != null)
                    {
                        Uri documentUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), docPackageRelationship.TargetUri);
                        PackagePart documentPart = wdPackage.GetPart(documentUri);

                        //  Load the document XML in the part into an XDocument instance.
                        xDoc = XDocument.Load(XmlReader.Create(documentPart.GetStream()));

                        //  Find the styles part. There will only be one.
                        PackageRelationship styleRelation = documentPart.GetRelationshipsByType(Strings.StyleDefsPartType).FirstOrDefault();
                        if (styleRelation != null)
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
                        StyleName = styleNode != null ? (string)styleNode.Attribute(w + "val") : defaultStyle
                    };

                // Retrieve the text of each paragraph.  
                var paraWithText =
                    from para in paragraphs
                    select new
                    {
                        para.ParagraphNode,
                        para.StyleName,
                        Text = WordExtensions.ParagraphText(para.ParagraphNode)
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
            var ccList = new List<string>();

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
    }
}
