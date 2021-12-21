// Open Xml SDK Refs
using DocumentFormat.OpenXml;
using OM = DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// .NET refs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;

namespace Office_File_Explorer.Helpers
{
    class WordFixes
    {
        static bool corruptionFound = false;

        public static bool FixDataDescriptor(string filePath)
        {
            corruptionFound = false;

            // TODO

            return corruptionFound;
        }

        public static bool FixMathAccents(string filePath)
        {
            corruptionFound = false;

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filePath, true))
            {
                // there is a scenario where accents are added to the subscript node, not oMath
                // accent is not allowed is subscript elements
                foreach (OM.Subscript sSub in myDoc.MainDocumentPart.Document.Descendants<OM.Subscript>())
                {
                    // loop through the subscript elements and if "acc" is found, delete it
                    // use localname because the openxml type shows unknown
                    foreach (OpenXmlElement oxe in sSub.ChildElements)
                    {
                        if (oxe.LocalName == "acc")
                        {
                            oxe.Remove();
                            corruptionFound = true;
                        }
                    }
                }

                if (corruptionFound)
                {
                    myDoc.MainDocumentPart.Document.Save();
                }
            }

            return corruptionFound;
        }

        public static bool FixContentControls(string filePath)
        {
            corruptionFound = false;
            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filePath, true))
            {
                // plain text content controls can't have any nested content controls
                foreach (var cc in myDoc.ContentControls())
                {
                    bool plainTextControl = false;
                    SdtProperties props = cc.Elements<SdtProperties>().FirstOrDefault();
                    foreach (OpenXmlElement oxe in cc.ChildElements)
                    {
                        // make sure we are a plain text control
                        foreach (OpenXmlElement oxeProp in props.ChildElements)
                        {
                            if (oxeProp.GetType().Name == "SdtContentText")
                            {
                                plainTextControl = true;
                            }
                        }

                        // if it is a plain text and it has an sdtcontentrun, we need to remove it
                        if (oxe.GetType().Name == "SdtContentRun" && plainTextControl == true)
                        {
                            foreach (OpenXmlElement oxeInner in oxe.ChildElements)
                            {
                                if (oxeInner.GetType().Name == "SdtRun")
                                {
                                    oxeInner.Remove();
                                    corruptionFound = true;
                                }
                            }
                        }
                    }
                }

                if (corruptionFound)
                {
                    myDoc.MainDocumentPart.Document.Save();
                }
            }

            return corruptionFound;
        }

        public static bool FixShapeInComment(string filePath)
        {
            corruptionFound = false;

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filePath, true))
            {
                if (myDoc.MainDocumentPart.WordprocessingCommentsPart is null)
                {
                    return corruptionFound;
                }

                // check for shape in a hyperlink, in a comment, which is not allowed
                foreach (Comment cm in myDoc.MainDocumentPart.WordprocessingCommentsPart.Comments)
                {
                    IEnumerable<Hyperlink> hLinks = cm.Descendants<Hyperlink>();

                    // does the comment have a hyperlink
                    if (hLinks.Any())
                    {
                        foreach (Hyperlink h in hLinks)
                        {
                            // get the runs for the hyperlink
                            IEnumerable<Run> runs = h.Descendants<Run>();
                            if (runs.Any())
                            {
                                // if there is an e1o in the run, delete the run
                                foreach (Run r in runs)
                                {
                                    if (r.Descendants<AlternateContent>().Any())
                                    {
                                        r.Remove();
                                        corruptionFound = true;
                                    }
                                }
                            }
                        }
                    }
                }

                if (corruptionFound)
                {
                    myDoc.MainDocumentPart.Document.Save();
                }
            }

            return corruptionFound;
        }

        /// <summary>
        /// del tags can't be nested, check for this scenario and fix the tags
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool FixDeleteRevision(string filePath)
        {
            corruptionFound = false;

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filePath, true))
            {
                var delRuns = myDoc.MainDocumentPart.Document.Descendants<DeletedRun>().ToList();
                foreach (DeletedRun del in delRuns)
                {
                    // if the first element in a deletedrun is another delete element, that is not correct
                    if (del.FirstChild.LocalName == "del")
                    {
                        // clone the del child element, remove it and then append it after the root del element
                        OpenXmlElement oxe = del.FirstChild.CloneNode(true);
                        del.FirstChild.Remove();
                        del.Append(oxe);
                        corruptionFound = true;
                    }
                }

                if (corruptionFound)
                {
                    myDoc.MainDocumentPart.Document.Save();
                }
            }            

            return corruptionFound;
        }

        public static bool FixHyperlinks(string filePath)
        {
            bool fileChanged = false;

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filePath, true))
            {
                bool isHyperlinkInBetweenSequence = false;
                IEnumerable<Paragraph> paras = myDoc.MainDocumentPart.Document.Descendants<Paragraph>();
                int pCount = paras.Count();
                int tempCount = 0;

                // need to keep looping paragraphs until we don't find a bad hlink position
                do
                {
                    isHyperlinkInBetweenSequence = false;
                    bool inBeginEndSequence = false;
                    int beginPosition = 0;
                    int endPosition = 0;
                    int elementCount = 0;
                    int beginCount = 0;
                    int prevRunPosition = 0;
                    tempCount = 0;

                    foreach (Paragraph p in paras)
                    {
                        tempCount++;
                        foreach (OpenXmlElement oxe in p.Descendants<OpenXmlElement>())
                        {
                            // keep track of previous run so we can get the right start position
                            // you could just use the first "begin" field code, but fixing it back up later is more challenging
                            // if we grab the begins root run, it makes this much easier
                            elementCount++;
                            if (oxe.GetType().Name == "Run")
                            {
                                prevRunPosition = elementCount;
                            }

                            // here we are keeping track of begin-end sequences
                            // the beginCount is there for nested begin-end scenarios
                            // you can have begin-begin-separate-end-end
                            if (oxe.GetType().Name == "FieldChar")
                            {
                                FieldChar fc = (FieldChar)oxe;
                                if (fc.FieldCharType == FieldCharValues.Begin)
                                {
                                    beginCount++;
                                    inBeginEndSequence = true;
                                    if (beginPosition == 0)
                                    {
                                        beginPosition = prevRunPosition;
                                    }
                                }

                                if (fc.FieldCharType == FieldCharValues.End)
                                {
                                    // valid sequence, reset values
                                    beginCount--;
                                    if (beginCount == 0)
                                    {
                                        inBeginEndSequence = false;
                                        beginPosition = 0;
                                    }
                                }
                            }

                            // if we are still in the middle of a begin-end sequence
                            // we can't have a hlink so we know we have a corruption
                            if (oxe.GetType().Name == "Hyperlink" && inBeginEndSequence == true)
                            {
                                // you can have a hlink in between the begin-end tags or vica versa
                                // so we are only looking for an hlink that has an end inside it with no begin
                                if (oxe.InnerXml.Contains(Strings.txtFieldCodeEnd) && !oxe.InnerXml.Contains(Strings.txtFieldCodeBegin))
                                {
                                    isHyperlinkInBetweenSequence = true;
                                    endPosition = elementCount;
                                    break;
                                }
                            }
                        }

                        if (isHyperlinkInBetweenSequence == true)
                        {
                            break;
                        }
                    }

                    // if isHyperlinkInBetween we need to loop again now that we have the bad position
                    if (isHyperlinkInBetweenSequence == true)
                    {
                        int tCount = 0;
                        bool atEndPosition = false;
                        List<OpenXmlElement> els = new List<OpenXmlElement>();
                        List<Run> runs = new List<Run>();

                        foreach (Paragraph p in paras)
                        {
                            foreach (OpenXmlElement oxe in p.Descendants<OpenXmlElement>())
                            {
                                tCount++;
                                if (tCount >= beginPosition)
                                {
                                    // once we are at the right position where we need to start moving elements around
                                    // create a list of elements
                                    els.Add(oxe);
                                    if (tCount == endPosition)
                                    {
                                        // once we are at the end of the bad sequence
                                        // reverse the list so we can prepend to the hyperlink
                                        atEndPosition = true;
                                        els.Reverse();

                                        foreach (OpenXmlElement e in els)
                                        {
                                            if (e.LocalName != "hyperlink")
                                            {
                                                // haven't quite figured out why, but I need to create a run for all non-run elements
                                                // then I can clone it and add it back to the right location in the hyperlink
                                                // then we can remove the original bad position elements
                                                if (e.LocalName != "r")
                                                {
                                                    Run r = new Run();
                                                    r.AppendChild(e.CloneNode(true));
                                                    oxe.PrependChild(r);
                                                    e.Remove();
                                                    fileChanged = true;
                                                }
                                                else
                                                {
                                                    oxe.PrependChild(e.CloneNode(false));
                                                    e.Remove();
                                                    fileChanged = true;
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            if (atEndPosition == true)
                            {
                                break;
                            }
                        }
                    }
                } while (tempCount < pCount);

                if (fileChanged)
                {
                    myDoc.MainDocumentPart.Document.Save();
                }
            }

            return fileChanged;
        }

        public static bool FixCommentFieldCodes(string filePath)
        {
            bool isFileChanged = false;

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filePath, true))
            {
                Regex emailPattern = new Regex(@"(.*?)<?(\b\S+@\S+\b)>?");
                if (myDoc.MainDocumentPart.WordprocessingCommentsPart is null)
                {
                    return isFileChanged;
                }
                
                WordprocessingCommentsPart commentsPart = myDoc.MainDocumentPart.WordprocessingCommentsPart;

                foreach (Comment cmt in commentsPart.Comments)
                {
                    IEnumerable<Paragraph> pList = cmt.Descendants<Paragraph>();
                    List<Run> rList = new List<Run>();

                    foreach (Paragraph p in pList)
                    {
                        // if the p has the mention style, it passes the first check we need to make
                        if (p.InnerXml.Contains(Strings.txtAtMentionStyle))
                        {
                            bool beginFound = false;
                            bool separateFound = false;
                            string emailAlias = string.Empty;

                            // now we need to loop each run and check the separate is missing
                            foreach (Run r in p.Descendants<Run>())
                            {
                                if (r.InnerXml.Contains(Strings.txtFieldCodeBegin))
                                {
                                    beginFound = true;
                                }

                                if (r.InnerXml.Contains(Strings.txtFieldCodeSeparate))
                                {
                                    separateFound = true;
                                }

                                // hold onto the mailto so we at least have something to use for the mention text
                                if (beginFound == true && r.InnerXml.Contains("mailto"))
                                {
                                    var groups = emailPattern.Match(r.InnerXml.ToString()).Groups;

                                    // trim the beginning of the mailto
                                    emailAlias = groups[2].Value.Remove(0, 7);

                                    // remove the domain and add @ to the beginning
                                    int atIndex = emailAlias.IndexOf('@');
                                    emailAlias = "@" + emailAlias.Remove(atIndex);
                                }

                                // once we get to the end, if we haven't found a separate, we need to add it back
                                if (r.InnerXml.Contains(Strings.txtFieldCodeEnd))
                                {
                                    if (r.InnerXml.Contains(Strings.txtAtMentionStyle) && separateFound == false)
                                    {
                                        // first, remove all children since we are in the area we need to change
                                        // add separate to the existing run
                                        r.RemoveAllChildren();
                                        RunProperties rPr = new RunProperties();
                                        RunStyle rs = new RunStyle();
                                        rs.Val = "Mention";
                                        rPr.Append(rs);
                                        r.Append(rPr);
                                        FieldChar fcs = new FieldChar();
                                        fcs.FieldCharType = FieldCharValues.Separate;
                                        r.Append(fcs);

                                        // add a new text run with the mailto alias
                                        Run rNewEnd = new Run();
                                        FieldChar fce = new FieldChar();
                                        fce.FieldCharType = FieldCharValues.End;
                                        rNewEnd.Append(fce);
                                        r.InsertAfterSelf(rNewEnd);

                                        // add a new end run back to close it out
                                        Run rNewText = new Run();
                                        Text t = new Text();
                                        t.Text = emailAlias;
                                        rNewText.Append(t);
                                        r.InsertAfterSelf(rNewText);

                                        isFileChanged = true;
                                    }

                                    // reset logic criteria
                                    beginFound = false;
                                    separateFound = false;
                                    emailAlias = string.Empty;
                                }
                            }
                        }
                    }
                }

                if (isFileChanged)
                {
                    myDoc.MainDocumentPart.Document.Save();
                }
            }

            return isFileChanged;
        }

        public static bool FixMissingCommentRefs(string filePath)
        {
            bool saveFile = false;

            using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
            {
                WordprocessingCommentsPart commentsPart = document.MainDocumentPart.WordprocessingCommentsPart;
                IEnumerable<OpenXmlUnknownElement> unknownList = document.MainDocumentPart.Document.Descendants<OpenXmlUnknownElement>();
                IEnumerable<CommentReference> commentRefs = document.MainDocumentPart.Document.Descendants<CommentReference>();

                bool cRefIdExists = false;

                if (commentsPart is null && commentRefs.Count() > 0)
                {
                    // if there are comment refs but no comments.xml, remove refs
                    foreach (CommentReference cr in commentRefs)
                    {
                        cr.Remove();
                        saveFile = true;
                    }
                }
                else if (commentsPart is null && commentRefs.Count() == 0)
                {
                    // for some reason these dangling refs are considered unknown types, not commentrefs
                    // convert to an openxmlelement then type it to a commentref to get the id
                    foreach (OpenXmlUnknownElement uk in unknownList)
                    {
                        if (uk.LocalName == "commentReference")
                        {
                            // so far I only see the id in the outerxml
                            XmlDocument xDoc = new XmlDocument();
                            xDoc.LoadXml(uk.OuterXml);

                            // traverse the outerxml until we get to the id
                            if (xDoc.ChildNodes.Count > 0)
                            {
                                foreach (XmlNode xNode in xDoc.ChildNodes)
                                {
                                    if (xNode.Attributes.Count > 0)
                                    {
                                        foreach (XmlAttribute xa in xNode.Attributes)
                                        {
                                            if (xa.LocalName == "id")
                                            {
                                                // now that we have the id number, we can use it to compare with the comment part
                                                // if the id exists in commentref but not the commentpart, it can be deleted
                                                foreach (Comment cm in commentsPart.Comments)
                                                {
                                                    int cId = Convert.ToInt32(cm.Id);
                                                    int cRefId = Convert.ToInt32(xa.Value);

                                                    if (cId == cRefId)
                                                    {
                                                        cRefIdExists = true;
                                                    }
                                                }

                                                if (cRefIdExists == false)
                                                {
                                                    uk.Remove();
                                                    saveFile = true;
                                                }
                                                else
                                                {
                                                    cRefIdExists = false;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }


                if (saveFile)
                {
                    document.MainDocumentPart.Document.Save();
                }
            }

            return saveFile;
        }

        public static bool FixTableGridProps(string filePath)
        {
            bool tblModified = false;

            using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
            {
                // global document variables
                OpenXmlElement tgClone = null;

                if (Word.IsPartNull(document, "Table") == false)
                {
                    // get the list of tables in the document
                    List<Table> tbls = document.MainDocumentPart.Document.Descendants<Table>().ToList();

                    foreach (Table tbl in tbls)
                    {
                        // you can have only one tblGrid per table, including nested tables
                        // it needs to be before any row elements so sequence is
                        // 1. check if the tblGrid element is before any table row
                        // 2. check for multiple tblGrid elements
                        bool tRowFound = false;
                        bool tGridBeforeRowFound = false;
                        int tGridCount = 0;

                        foreach (OpenXmlElement oxe in tbl.Elements())
                        {
                            // flag if we found a trow, once we find 1, the rest do not matter
                            if (oxe.GetType().Name == "TableRow")
                            {
                                tRowFound = true;
                            }

                            // when we get to a tablegrid, we have a few things to check
                            // 1. have we found a table row previously
                            // 2. only one table grid can exist in the table, if there are multiple, delete the extras
                            if (oxe.GetType().Name == "TableGrid")
                            {
                                // increment the tg counter
                                tGridCount++;

                                // if we have a table row and no table grid has been found yet, we need to save out this table grid
                                // then move it in front of the table row later
                                if (tRowFound == true && tGridCount == 1)
                                {
                                    tGridBeforeRowFound = true;
                                    tgClone = oxe.CloneNode(true);
                                    oxe.Remove();
                                }

                                // if we have multiple table grids, delete the extras
                                if (tGridCount > 1)
                                {
                                    oxe.Remove();
                                    tblModified = true;
                                }
                            }
                        }

                        // if we had a table grid before a row was found, move it before the first row in the table
                        if (tGridBeforeRowFound == true)
                        {
                            tbl.InsertBefore(tgClone, tbl.GetFirstChild<TableRow>());
                            tblModified = true;
                        }
                    }
                }

                // save the file if we modified the table
                if (tblModified == true)
                {
                    document.MainDocumentPart.Document.Save();
                }
            }

            return tblModified;
        }

        public static bool FixRevisions(string filePath)
        {
            bool isFixed = false;

            using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
            {
                Document doc = document.MainDocumentPart.Document;
                var deleted = doc.Descendants<DeletedRun>().ToList();

                // loop each DeletedRun
                foreach (DeletedRun dr in deleted)
                {
                    foreach (OpenXmlElement oxedr in dr)
                    {
                        // if we have a run, we need to look for Text tags
                        if (oxedr.GetType().ToString() == Strings.dfowRun)
                        {
                            Run r = (Run)oxedr;
                            foreach (OpenXmlElement oxe in oxedr.ChildElements)
                            {
                                // you can't have a Text tag inside a DeletedRun
                                if (oxe.GetType().ToString() == Strings.dfowText)
                                {
                                    // create a DeletedText object so we can replace it with the Text tag
                                    DeletedText dt = new DeletedText();

                                    // check for attributes
                                    if (oxe.HasAttributes)
                                    {
                                        if (oxe.GetAttributes().Count > 0)
                                        {
                                            dt.SetAttributes(oxe.GetAttributes());
                                        }
                                    }

                                    // set the text value
                                    dt.Text = oxe.InnerText;

                                    // replace the Text with new DeletedText
                                    r.ReplaceChild(dt, oxe);
                                    isFixed = true;
                                }
                            }
                        }
                    }
                }

                // now save the file if we have changes
                if (isFixed == true)
                {
                    doc.Save();
                }
            }

            return isFixed;
        }

        public static bool FixListTemplates(string filePath)
        {
            bool orphanedListTemplatesFound = false;

            NumberingDetails bulletMultiLevelNumberingValues = new NumberingDetails();
            NumberingDetails bulletSingleLevelNumberingValues = new NumberingDetails();

            List<int> bulletMultiLevelNumIdsInUse = new List<int>();
            List<int> bulletSingleLevelNumIdsInUse = new List<int>();

            using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
            {
                // if the numbering part does not exist, nothing to check
                if (document.MainDocumentPart.NumberingDefinitionsPart is null)
                {
                    return orphanedListTemplatesFound;
                }

                // get the list of numId's and AbstractNum's in numbering.xml
                var absNumsInUseList = document.MainDocumentPart.NumberingDefinitionsPart.Numbering.Descendants<AbstractNum>().ToList();
                var numInstancesInUseList = document.MainDocumentPart.NumberingDefinitionsPart.Numbering.Descendants<NumberingInstance>().ToList();

                bool bulletSingleLevelFound = false;
                bool bulletMultiLevelFound = false;

                foreach (AbstractNum an in absNumsInUseList)
                {
                    foreach (NumberingInstance ni in numInstancesInUseList)
                    {
                        // if the abstractnum and numId match, they are the same listtemplate
                        if (ni.AbstractNumId.Val == an.AbstractNumberId.Value)
                        {
                            // get the level count
                            var lvlNumberingList = an.Descendants<Level>().ToList();

                            // since we have the list template, find out if it is a bullet
                            foreach (OpenXmlElement anChild in an)
                            {
                                if (anChild.GetType().ToString() == Strings.dfowLevel)
                                {
                                    Level lvl = (Level)anChild;

                                    // try to catch different "types" of numberingformat
                                    // for now, I'm only checking for a single and multi-level bullets
                                    if (lvl.NumberingFormat.Val == "bullet" && lvlNumberingList.Count > 1 && lvl.LevelIndex == 0)
                                    {
                                        // if level is > 1, this is a multi level list
                                        bulletMultiLevelNumIdsInUse.Add(ni.NumberID);

                                        if (bulletMultiLevelFound == false)
                                        {
                                            bulletMultiLevelNumberingValues.AbsNumId = ni.AbstractNumId.Val;
                                            bulletMultiLevelNumberingValues.NumFormat = "bulletMultiLevel";
                                            bulletMultiLevelNumberingValues.NumId = ni.NumberID;
                                            bulletMultiLevelFound = true;
                                        }
                                    }
                                    else if (lvl.NumberingFormat.Val == "bullet" && lvlNumberingList.Count == 1 && lvl.LevelIndex == 0)
                                    {
                                        // if level = 1, this is a single level list
                                        bulletSingleLevelNumIdsInUse.Add(ni.NumberID);

                                        if (bulletSingleLevelFound == false)
                                        {
                                            bulletSingleLevelNumberingValues.AbsNumId = ni.AbstractNumId.Val;
                                            bulletSingleLevelNumberingValues.NumFormat = "bulletSingle";
                                            bulletSingleLevelNumberingValues.NumId = ni.NumberID;
                                            bulletSingleLevelFound = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // now that we have bullet numids to use, we can apply it to each paragraph
                MainDocumentPart mainPart = document.MainDocumentPart;
                StyleDefinitionsPart stylePart = mainPart.StyleDefinitionsPart;

                foreach (OpenXmlElement el in mainPart.Document.Descendants<Paragraph>())
                {
                    if (el.Descendants<NumberingId>().Count() > 0)
                    {
                        foreach (NumberingId pNumId in el.Descendants<NumberingId>())
                        {
                            foreach (var o in bulletMultiLevelNumIdsInUse)
                            {
                                if (o == pNumId.Val)
                                {
                                    pNumId.Val = bulletMultiLevelNumberingValues.NumId;
                                }
                            }

                            foreach (var o in bulletSingleLevelNumIdsInUse)
                            {
                                if (o == pNumId.Val)
                                {
                                    pNumId.Val = bulletSingleLevelNumberingValues.NumId;
                                }
                            }
                        }
                    }
                }

                foreach (HeaderPart hdrPart in mainPart.HeaderParts)
                {
                    foreach (OpenXmlElement el in hdrPart.Header.Elements())
                    {
                        foreach (NumberingId hNumId in el.Descendants<NumberingId>())
                        {
                            foreach (var o in bulletMultiLevelNumIdsInUse)
                            {
                                if (o == hNumId.Val)
                                {
                                    hNumId.Val = bulletMultiLevelNumberingValues.NumId;
                                }
                            }

                            foreach (var o in bulletSingleLevelNumIdsInUse)
                            {
                                if (o == hNumId.Val)
                                {
                                    hNumId.Val = bulletSingleLevelNumberingValues.NumId;
                                }
                            }
                        }
                    }
                }

                foreach (FooterPart ftrPart in mainPart.FooterParts)
                {
                    foreach (OpenXmlElement el in ftrPart.Footer.Elements())
                    {
                        foreach (NumberingId fNumId in el.Descendants<NumberingId>())
                        {
                            foreach (var o in bulletMultiLevelNumIdsInUse)
                            {
                                if (o == fNumId.Val)
                                {
                                    fNumId.Val = bulletMultiLevelNumberingValues.NumId;
                                }
                            }

                            foreach (var o in bulletSingleLevelNumIdsInUse)
                            {
                                if (o == fNumId.Val)
                                {
                                    fNumId.Val = bulletSingleLevelNumberingValues.NumId;
                                }
                            }
                        }
                    }
                }

                foreach (OpenXmlElement el in stylePart.Styles.Elements())
                {
                    if (el.GetType().ToString() == Strings.dfowStyle)
                    {
                        string styleEl = el.GetAttribute("styleId", Strings.wordMainAttributeNamespace).Value;
                        int pStyle = Word.ParagraphsByStyleName(mainPart, styleEl).Count();

                        if (pStyle > 0)
                        {
                            foreach (NumberingId sEl in el.Descendants<NumberingId>())
                            {
                                foreach (var o in bulletMultiLevelNumIdsInUse)
                                {
                                    if (o == sEl.Val)
                                    {
                                        sEl.Val = bulletMultiLevelNumberingValues.NumId;
                                    }
                                }

                                foreach (var o in bulletSingleLevelNumIdsInUse)
                                {
                                    if (o == sEl.Val)
                                    {
                                        sEl.Val = bulletSingleLevelNumberingValues.NumId;
                                    }
                                }
                            }
                        }
                    }
                }

                document.MainDocumentPart.Document.Save();
                return orphanedListTemplatesFound;
            }
        }

        public static bool FixEndnotes(string filePath)
        {
            bool corruptEndnotesFound = false;

            using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
            {
                if (document.MainDocumentPart.EndnotesPart != null)
                {
                    Endnotes ens = document.MainDocumentPart.EndnotesPart.Endnotes;

                    foreach (Endnote en in ens)
                    {
                        // get the paragraph list from the endnote, if it has more than 1000 runs of content
                        // delete it...need to find a way to check for dupes
                        // for now just deleting all but the first paragraph run
                        var paraList = en.Descendants<Paragraph>().ToList();
                        foreach (var p in paraList)
                        {
                            var rList = p.Descendants<Run>().ToList();
                            if (rList.Count > 1000)
                            {
                                int count = 0;
                                foreach (var r in rList)
                                {
                                    if (count > 0)
                                    {
                                        r.Remove();
                                        corruptEndnotesFound = true;
                                    }
                                    count++;
                                }
                            }
                        }
                    }
                }

                if (corruptEndnotesFound == true)
                {
                    document.MainDocumentPart.Document.Save();
                }
            }

            return corruptEndnotesFound;
        }

        /// <summary>
        /// Sometimes bookmarks are added and the start/end tag is missing
        /// This function will try to find those orphan tags and remove them
        /// </summary>
        /// <param name="filename">file to be scanned</param>
        /// <returns>true for successful removal and false if none are found</returns>
        public static bool RemoveMissingBookmarkTags(string filename)
        {
            bool isFixed = false;

            try
            {
                using (WordprocessingDocument package = WordprocessingDocument.Open(filename, true))
                {
                    if (package.MainDocumentPart.WordprocessingCommentsPart is null) { return false; }
                    if (package.MainDocumentPart.WordprocessingCommentsPart.Comments is null) { return false; }

                    IEnumerable<BookmarkStart> bkStartList = package.MainDocumentPart.WordprocessingCommentsPart.Comments.Descendants<BookmarkStart>();
                    IEnumerable<BookmarkEnd> bkEndList = package.MainDocumentPart.WordprocessingCommentsPart.Comments.Descendants<BookmarkEnd>();

                    // create temp lists so we can loop and remove any that exist in both lists
                    // if we have a start and end, the bookmark is valid and we can remove the rest
                    List<string> bkStartTagIds = new List<string>();
                    List<string> bkEndTagIds = new List<string>();

                    // check each start and find if there is a matching end tag id
                    foreach (BookmarkStart bks in bkStartList)
                    {
                        foreach (BookmarkEnd bke in bkEndList)
                        {
                            if (bke.Id.ToString() == bks.Id.ToString())
                            {
                                bkStartTagIds.Add(bke.Id);
                            }
                        }
                    }

                    // now we can check if there is a end tag with a matching start tag id
                    foreach (BookmarkEnd bke in bkEndList)
                    {
                        foreach (BookmarkStart bks in bkStartList)
                        {
                            if (bks.Id.ToString() == bke.Id.ToString())
                            {
                                bkEndTagIds.Add(bks.Id);
                            }
                        }
                    }

                    // now that we know all the id's that match, we can loop again and remove id's that are not in the lists
                    // first check orphaned start tags
                    bool startTagFound = false;

                    foreach (BookmarkStart bks in bkStartList)
                    {
                        foreach (object o in bkEndTagIds)
                        {
                            // if the end tag matches and we can ignore doing anything
                            if (o.ToString() == bks.Id.ToString())
                            {
                                startTagFound = true;
                            }
                        }

                        // if we get here and no match was found, it is orphaned and we can delete
                        if (startTagFound == false)
                        {
                            bks.Remove();
                            isFixed = true;
                        }
                        else
                        {
                            // reset the value for the next start tag check
                            startTagFound = false;
                        }
                    }

                    // do the same check for end tags
                    bool endTagFound = false;

                    foreach (BookmarkEnd bke in bkEndList)
                    {
                        foreach (object o in bkStartTagIds)
                        {
                            if (o.ToString() == bke.Id.ToString())
                            {
                                endTagFound = true;
                            }
                        }

                        if (endTagFound == false)
                        {
                            bke.Remove();
                            isFixed = true;
                        }
                        else
                        {
                            endTagFound = false;
                        }
                    }

                    if (isFixed)
                    {
                        package.MainDocumentPart.Document.Save();
                    }
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "RemoveMissingBookmarkTags Error: " + ex.Message);
                return false;
            }

            return isFixed;
        }

        // <summary>
        /// look for bookmark tags in a plain cc
        /// this is not allowed and those need to be removed
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public static bool RemovePlainTextCcFromBookmark(string filename)
        {
            bool isFixed = false;

            try
            {
                using (WordprocessingDocument package = WordprocessingDocument.Open(filename, true))
                {
                    IEnumerable<BookmarkStart> bkStartList = package.MainDocumentPart.Document.Descendants<BookmarkStart>();
                    IEnumerable<BookmarkEnd> bkEndList = package.MainDocumentPart.Document.Descendants<BookmarkEnd>();
                    List<string> removedBookmarkIds = new List<string>();

                    if (bkStartList.Count() > 0)
                    {
                        foreach (BookmarkStart bk in bkStartList)
                        {
                            var cElem = bk.Parent;
                            var pElem = bk.Parent;
                            bool endLoop = false;

                            do
                            {
                                // first check if we are a content control
                                if (cElem.Parent != null && cElem.Parent.ToString().Contains(Strings.dfowSdt))
                                {
                                    foreach (OpenXmlElement oxe in cElem.Parent.ChildElements)
                                    {
                                        // get the properties
                                        if (oxe.GetType().Name == "SdtProperties")
                                        {
                                            foreach (OpenXmlElement oxeSdtAlias in oxe)
                                            {
                                                // check for plain text
                                                if (oxeSdtAlias.GetType().Name == "SdtContentText")
                                                {
                                                    // if the parent is a plain text content control, bookmark is not allowed
                                                    // add the id to the list of bookmarks that need to be deleted
                                                    removedBookmarkIds.Add(bk.Id);
                                                    endLoop = true;
                                                }
                                            }
                                        }
                                    }

                                    // set the next element to the parent and continue moving up the element chain
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
                                    else
                                    {
                                        // set pElem to the parent so we can check for the end of the loop
                                        // set cElem to the parent also so we can continue moving up the element chain
                                        pElem = cElem.Parent;
                                        cElem = pElem;

                                        // loop should continue until we get to the body element, then we can stop looping
                                        if (pElem.ToString() == Strings.dfowBody)
                                        {
                                            endLoop = true;
                                        }
                                    }
                                }
                            } while (endLoop == false);
                        }

                        // now that we have the list of bookmark id's to be removed
                        // loop each list and delete any bookmark that has a matching id
                        foreach (var o in removedBookmarkIds)
                        {
                            foreach (BookmarkStart bkStart in bkStartList)
                            {
                                if (bkStart.Id == o)
                                {
                                    bkStart.Remove();
                                }
                            }

                            foreach (BookmarkEnd bkEnd in bkEndList)
                            {
                                if (bkEnd.Id == o)
                                {
                                    bkEnd.Remove();
                                }
                            }
                        }

                        // save the part
                        package.MainDocumentPart.Document.Save();

                        // check if there were any fixes made and update the output display
                        if (removedBookmarkIds.Count > 0)
                        {
                            isFixed = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "RemovePlainTextCcFromBookmark Error: " + ex.Message);
                return false;
            }

            return isFixed;
        }
    }
}
