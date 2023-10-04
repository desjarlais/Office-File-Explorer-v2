using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using A = DocumentFormat.OpenXml.Drawing;
using PShape = DocumentFormat.OpenXml.Presentation.Shape;
using Drawing = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;
using ShapeStyle = DocumentFormat.OpenXml.Presentation.ShapeStyle;
using ModernComment = DocumentFormat.OpenXml.Office2021.PowerPoint.Comment;
using Comment = DocumentFormat.OpenXml.Presentation.Comment;
using System.IO.Packaging;

namespace Office_File_Explorer.Helpers
{
    class PowerPoint
    {
        public static bool fSuccess;

        public static List<string> GetFonts(string fPath)
        {
            List<string> fonts = new List<string>();
            int fCount = 0;

            using (PresentationDocument pptDoc = PresentationDocument.Open(fPath, false))
            {
                // list the embedded fonts
                if (pptDoc.PresentationPart.Presentation.EmbeddedFontList is null)
                {
                    return fonts;
                }

                foreach (EmbeddedFont ef in pptDoc.PresentationPart.Presentation.EmbeddedFontList)
                {
                    fCount++;
                    if (ef.Features.IsReadOnly)
                    {
                        fonts.Add(fCount + Strings.wPeriod + ef.Font.Typeface + " || Character Set = " + AppUtilities.GetFontCharacterSet(ef.Font.CharacterSet) + " (Read-Only)");
                    }
                    else
                    {
                        fonts.Add(fCount + Strings.wPeriod + ef.Font.Typeface + " || Character Set = " + AppUtilities.GetFontCharacterSet(ef.Font.CharacterSet));
                    }
                }
            }

            return fonts;
        }

        public static bool RemoveComments(string path)
        {
            fSuccess = false;

            using (PresentationDocument pptDoc = PresentationDocument.Open(path, true))
            {
                PresentationPart pPart = pptDoc.PresentationPart;

                foreach (SlidePart sPart in pPart.SlideParts)
                {
                    SlideCommentsPart sCPart = sPart.SlideCommentsPart;
                    if (sCPart is null)
                    {
                        return fSuccess;
                    }

                    foreach (Comment cmt in sCPart.CommentList)
                    {
                        cmt.Remove();
                        fSuccess = true;
                    }
                }

                if (fSuccess)
                {
                    pptDoc.PresentationPart.Presentation.Save();
                }
            }

            return fSuccess;
        }

        /// <summary>
        /// Move a slide to a different position in the slide order in the presentation.
        /// </summary>
        /// <param name="presentationDocument"></param>
        /// <param name="from">slide index # of the source slide</param>
        /// <param name="to">slide index # of the target slide</param>
        public static void MoveSlide(PresentationDocument presentationDocument, int from, int to)
        {
            // Get the presentation part from the presentation document.
            PresentationPart presentationPart = presentationDocument.PresentationPart;

            // The slide count is not zero, so the presentation must contain slides.
            Presentation presentation = presentationPart.Presentation;
            SlideIdList slideIdList = presentation.SlideIdList;

            // Get the slide ID of the source slide.
            SlideId sourceSlide = slideIdList.ChildElements[from] as SlideId;
            SlideId targetSlide = slideIdList.ChildElements[to] as SlideId;

            // Remove the source slide from its current position.
            sourceSlide.Remove();

            // Insert the source slide at its new position after the target slide.
            // if the slide being moved is before the target position, use InsertAfter
            // otherwise, we want to use InsertBefore
            if (from < to)
            {
                slideIdList.InsertAfter(sourceSlide, targetSlide);
            }
            else
            {
                slideIdList.InsertBefore(sourceSlide, targetSlide);
            }

            // Save the modified presentation.
            presentation.Save();
        }

        /// <summary>
        /// Change the fill color of a shape, docName must have a filled shape as the first shape on the first slide.
        /// </summary>
        /// <param name="docName">path to the file</param>
        public static void SetPPTShapeColor(string docName)
        {
            using (PresentationDocument ppt = PresentationDocument.Open(docName, true))
            {
                // Get the relationship ID of the first slide.
                PresentationPart part = ppt.PresentationPart;
                OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;
                string relId = (slideIds[0] as SlideId).RelationshipId;

                // Get the slide part from the relationship ID.
                SlidePart slide = (SlidePart)part.GetPartById(relId);

                if (slide != null)
                {
                    // Get the shape tree that contains the shape to change.
                    ShapeTree tree = slide.Slide.CommonSlideData.ShapeTree;

                    // Get the first shape in the shape tree.
                    PShape shape = tree.GetFirstChild<PShape>();

                    if (shape != null)
                    {
                        // Get the style of the shape.
                        ShapeStyle style = shape.ShapeStyle;

                        // Get the fill reference.
                        FillReference fillRef = style.FillReference;

                        // Set the fill color to SchemeColor Accent 6;
                        fillRef.SchemeColor = new SchemeColor
                        {
                            Val = SchemeColorValues.Accent6
                        };

                        // Save the modified slide.
                        slide.Slide.Save();
                    }
                }
            }
        }

        /// <summary>
        /// Function to retrieve the number of slides
        /// </summary>
        /// <param name="fileName">path to the file</param>
        /// <param name="includeHidden">default is true, pass false if you don't want hidden slides counted</param>
        /// <returns></returns>
        public static int RetrieveNumberOfSlides(string fPath, bool includeHidden = true)
        {
            int slidesCount = 0;

            using (PresentationDocument doc = PresentationDocument.Open(fPath, false))
            {
                // Get the presentation part of the document.
                PresentationPart presentationPart = doc.PresentationPart;
                if (presentationPart is not null)
                {
                    if (includeHidden)
                    {
                        slidesCount = presentationPart.SlideParts.Count();
                    }
                    else
                    {
                        // Each slide can include a Show property, which if hidden 
                        // will contain the value "0". The Show property may not 
                        // exist, and most likely will not, for non-hidden slides.
                        var slides = presentationPart.SlideParts.Where((s) => (s.Slide is not null) && ((s.Slide.Show is null) ||
                            (s.Slide.Show.HasValue && s.Slide.Show.Value)));
                        slidesCount = slides.Count();
                    }
                }
            }
            return slidesCount;
        }

        public static List<string> GetHyperlinks(string fPath)
        {
            List<string> tList = new List<string>();
            
            int linkCount = 0;
            foreach (string s in GetAllExternalHyperlinksInPresentation(fPath))
            {
                linkCount++;
                tList.Add(linkCount + Strings.wPeriod + s);
            }

            return tList;
        }

        /// <summary>
        /// check for both legacy and modern comments
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static List<string> GetComments(string fPath)
        {
            List<string> tList = new List<string>();

            PresentationDocument presentationDocument = PresentationDocument.Open(fPath, false);
            PresentationPart pPart = presentationDocument.PresentationPart;
            int commentCount = 0;

            foreach (SlidePart sPart in pPart.SlideParts)
            {
                // legacy comments
                SlideCommentsPart sCPart = sPart.SlideCommentsPart;
                if (sCPart is not null)
                {
                    foreach (Comment cmt in sCPart.CommentList)
                    {
                        commentCount++;
                        tList.Add(commentCount + Strings.wPeriod + cmt.InnerText);
                    }
                }

                // modern comments
                if (sPart.commentParts is not null)
                {
                    IEnumerable<PowerPointCommentPart> modernComments = sPart.commentParts;
                    foreach (PowerPointCommentPart modernComment in modernComments)
                    {
                        foreach (ModernComment.Comment c in modernComment.CommentList)
                        {
                            string commentAuthor = string.Empty;
                            foreach (ModernComment.Author a in pPart.authorsPart.AuthorList)
                            {
                                if (a.Id == c.AuthorId)
                                {
                                    commentAuthor = a.Name;
                                }
                            }

                            commentCount++;
                            tList.Add(commentCount + Strings.wPeriod + "Author: " + commentAuthor + " Comment: " + c.InnerText);
                        }
                    }
                }
            }

            return tList;
        }

        public static List<string> GetSlideTitles(string fPath)
        {
            List<string> tList = new List<string>();
            using (PresentationDocument presentationDocument = PresentationDocument.Open(fPath, false))
            {
                int slideCount = 0;

                foreach (string s in GetSlideTitles(presentationDocument))
                {
                    slideCount++;
                    tList.Add(slideCount + Strings.wPeriod + s);
                }
            }

            return tList;
        }

        public static List<string> GetSlideText(string fPath)
        {
            List<string> tList = new List<string>();

            int sCount = RetrieveNumberOfSlides(fPath);
            if (sCount > 0)
            {
                int count = 0;

                do
                {
                    GetSlideIdAndText(out string sldText, fPath, count);
                    tList.Add("Slide " + (count + 1) + Strings.wPeriod + sldText);
                    count++;
                } while (count < sCount);
            }

            return tList;
        }

        public static List<string> GetSlideTransitions(string fPath)
        {
            List<string> tList = new List<string>();
            using (PresentationDocument ppt = PresentationDocument.Open(fPath, false))
            {
                int transitionCount = 0;
                foreach (string s in GetSlideTransitions(ppt))
                {
                    transitionCount++;
                    tList.Add(transitionCount + Strings.wPeriod + s);
                }
            }

            return tList;
        }

        // Get a list of the transitions of all the slides in the presentation.
        public static IList<string> GetSlideTransitions(PresentationDocument presentationDocument)
        {
            if (presentationDocument is null)
            {
                throw new ArgumentNullException(Strings.pptexceptionPowerPoint);
            }

            // Get a PresentationPart object from the PresentationDocument object.
            PresentationPart presentationPart = presentationDocument.PresentationPart;

            if (presentationPart != null && presentationPart.Presentation != null)
            {
                // Get a Presentation object from the PresentationPart object.
                Presentation presentation = presentationPart.Presentation;

                if (presentation.SlideIdList != null)
                {
                    List<string> transitionsList = new List<string>();

                    // Get the transition of each slide in the slide order.
                    foreach (var slideId in presentation.SlideIdList.Elements<SlideId>())
                    {
                        SlidePart slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                        string transition = string.Empty;

                        if (slidePart.Slide.Transition != null)
                        {
                            foreach (var t in slidePart.Slide.Transition)
                            {
                                transition = t.LocalName;
                            }
                        }
                        else
                        {
                            transition = "none";
                        }

                        // An empty title can also be added.
                        transitionsList.Add(transition);
                    }

                    return transitionsList;
                }
            }

            return null;
        }

        /// <summary>
        /// Get the slideId and text for that slide
        /// </summary>
        /// <param name="sldText">string returned to caller</param>
        /// <param name="docName">path to powerpoint file</param>
        /// <param name="index">slide number</param>
        public static void GetSlideIdAndText(out string sldText, string fPath, int index)
        {
            using (PresentationDocument ppt = PresentationDocument.Open(fPath, false))
            {
                // Get the relationship ID of the first slide.
                PresentationPart part = ppt.PresentationPart;
                OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;

                string relId = (slideIds[index] as SlideId).RelationshipId;

                // Get the slide part from the relationship ID.
                SlidePart slide = (SlidePart)part.GetPartById(relId);

                // Build a StringBuilder object.
                StringBuilder paragraphText = new StringBuilder();

                // Get the inner text of the slide:
                IEnumerable<A.Text> texts = slide.Slide.Descendants<A.Text>();
                foreach (A.Text text in texts)
                {
                    paragraphText.Append(text.Text);
                }
                sldText = paragraphText.ToString();
            }
        }

        public static void DeleteUnusedSlideLayoutParts(PresentationDocument ppt, List<string> usedSlideLayoutIds)
        {
            foreach (SlideMasterPart smp in ppt.PresentationPart.SlideMasterParts)
            {
                //parts.Clear();
                foreach (SlideLayoutPart slp in smp.SlideLayoutParts)
                {
                    bool sIdNotFound = true;

                    foreach (string sId in usedSlideLayoutIds)
                    {
                        if (sId == slp.Uri.ToString())
                        {
                            sIdNotFound = false;
                        }
                    }

                    if (sIdNotFound)
                    {
                        smp.DeletePart(slp);
                        ppt.Save();
                    }
                }
            }
        }

        public static List<string> GetSlideLayoutId(PresentationDocument ppt)
        {
            List<string> slideLayoutIds = new List<string>();

            // Get a PresentationPart object from the PresentationDocument object.
            PresentationPart presentationPart = ppt.PresentationPart;

            if (presentationPart != null && presentationPart.Presentation != null)
            {
                // Get a Presentation object from the PresentationPart object.
                Presentation presentation = presentationPart.Presentation;

                if (presentation.SlideIdList != null)
                {
                    // Get the title of each slide in the slide order.
                    foreach (var slideId in presentation.SlideIdList.Elements<SlideId>())
                    {
                        SlidePart slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                        SlideLayoutPart slideLayoutPart = slidePart.SlideLayoutPart;
                        slideLayoutIds.Add(slideLayoutPart.Uri.ToString());
                    }
                }
            }

            slideLayoutIds = slideLayoutIds.Distinct().ToList();

            return slideLayoutIds;
        }

        public static IList<string> GetSlideTitles(PresentationDocument presentationDocument)
        {
            if (presentationDocument is null)
            {
                throw new ArgumentNullException(Strings.pptexceptionPowerPoint);
            }

            // Get a PresentationPart object from the PresentationDocument object.
            PresentationPart presentationPart = presentationDocument.PresentationPart;

            if (presentationPart != null && presentationPart.Presentation != null)
            {
                // Get a Presentation object from the PresentationPart object.
                Presentation presentation = presentationPart.Presentation;

                if (presentation.SlideIdList != null)
                {
                    List<string> titlesList = new List<string>();

                    // Get the title of each slide in the slide order.
                    foreach (var slideId in presentation.SlideIdList.Elements<SlideId>())
                    {
                        SlidePart slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;

                        // Get the slide title.
                        string title = GetSlideTitle(slidePart);

                        // An empty title can also be added.
                        titlesList.Add(title);
                    }

                    return titlesList;
                }
            }

            return null;
        }

        // Get the title string of the slide.
        public static string GetSlideTitle(SlidePart slidePart)
        {
            if (slidePart is null)
            {
                throw new ArgumentNullException(Strings.pptexceptionPowerPoint);
            }

            // Declare a paragraph separator.
            string paragraphSeparator = null;

            if (slidePart.Slide != null)
            {
                // Find all the title shapes.
                var shapes = from shape in slidePart.Slide.Descendants<PShape>()
                             where IsTitleShape(shape)
                             select shape;

                StringBuilder paragraphText = new StringBuilder();

                foreach (var shape in shapes)
                {
                    // Get the text in each paragraph in this shape.
                    foreach (var paragraph in shape.TextBody.Descendants<Drawing.Paragraph>())
                    {
                        // Add a line break.
                        paragraphText.Append(paragraphSeparator);

                        foreach (var text in paragraph.Descendants<Drawing.Text>())
                        {
                            paragraphText.Append(text.Text);
                        }

                        paragraphSeparator = "\n";
                    }
                }

                return paragraphText.ToString();
            }

            return string.Empty;
        }

        /// <summary>
        /// Determines whether the shape is a title shape.
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        private static bool IsTitleShape(PShape shape)
        {
            var placeholderShape = shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild<PlaceholderShape>();
            if (placeholderShape != null && placeholderShape.Type != null && placeholderShape.Type.HasValue)
            {
                switch ((PlaceholderValues)placeholderShape.Type)
                {
                    // Any title shape.
                    case PlaceholderValues.Title:

                    // A centered title.
                    case PlaceholderValues.CenteredTitle:
                        return true;

                    default:
                        return false;
                }
            }
            return false;
        }

        // Returns all the external hyperlinks in the slides of a presentation.
        public static IEnumerable<String> GetAllExternalHyperlinksInPresentation(string fPath)
        {
            // Declare a list of strings.
            List<string> ret = new List<string>();

            PresentationDocument document = PresentationDocument.Open(fPath, false);

            // Iterate through all the slide parts in the presentation part.
            foreach (SlidePart slidePart in document.PresentationPart.SlideParts)
            {
                IEnumerable<HyperlinkType> links = slidePart.Slide.Descendants<HyperlinkType>();

                // Iterate through all the links in the slide part.
                foreach (HyperlinkType link in links)
                {
                    // Iterate through all the external relationships in the slide part. 
                    foreach (HyperlinkRelationship relation in slidePart.HyperlinkRelationships)
                    {
                        // If the relationship ID matches the link ID
                        if (relation.Id.Equals(link.Id))
                        {
                            // Add the URI of the external relationship to the list of strings.
                            ret.Add(relation.Uri.AbsoluteUri);
                        }
                    }
                }
            }

            // Return the list of strings.
            return ret;
        }

        // Delete comments by a specific author. Pass an empty string for the author to delete all comments, by all authors.
        public static bool DeleteComments(string fileName, string author)
        {
            bool isChanged = false;
            PresentationDocument doc = PresentationDocument.Open(fileName, true);

            // Get the authors part.
            CommentAuthorsPart authorsPart = doc.PresentationPart.GetPartsOfType<CommentAuthorsPart>().FirstOrDefault();

            if (authorsPart is null)
            {
                // There's no authors part, so just
                // fail. If no authors, there can't be any comments.
                return isChanged;
            }

            // Get the comment authors, or the specified author if supplied:
            var commentAuthors = authorsPart.CommentAuthorList.Elements<CommentAuthor>();
            if (!string.IsNullOrEmpty(author))
            {
                commentAuthors = commentAuthors.Where(e => e.Name.Value.Equals(author));
            }

            bool changed = false;
            foreach (var commentAuthor in commentAuthors.ToArray())
            {
                var authorId = commentAuthor.Id;

                // Iterate through all the slides and get the slide parts.
                foreach (var slide in doc.PresentationPart.GetPartsOfType<SlidePart>())
                {
                    // Iterate through the slide parts and find the slide comment parts.
                    var slideCommentParts = slide.GetPartsOfType<SlideCommentsPart>().ToArray();

                    foreach (var slideCommentsPart in slideCommentParts)
                    {
                        // Get the list of comments.
                        var commentList = slideCommentsPart.CommentList.Elements<Comment>().
                          Where(e => e.AuthorId.Value == authorId.Value);

                        foreach (var comment in commentList.ToArray())
                        {
                            // Delete all the comments by the specified author.
                            slideCommentsPart.CommentList.RemoveChild<Comment>(comment);
                            isChanged = true;
                        }

                        // No comments left? Delete the comments part for this slide.
                        if (slideCommentsPart.CommentList.Count() == 0)
                        {
                            slide.DeletePart(slideCommentsPart);
                        }
                        else
                        {
                            // Save the slide comments part.
                            slideCommentsPart.CommentList.Save();
                        }
                    }
                }

                // Delete the comment author from the comment authors part.
                authorsPart.CommentAuthorList.RemoveChild<CommentAuthor>(commentAuthor);

                changed = true;
            }

            // Changed will only be false if the caller requested comments
            // for a particular author, and that author has no comments.
            if (changed)
            {
                if (authorsPart.CommentAuthorList.Count() == 0)
                {
                    // No authors left, so delete the part.
                    doc.PresentationPart.DeletePart(authorsPart);
                }
                else
                {
                    // Save the comment authors part.
                    authorsPart.CommentAuthorList.Save();
                }
            }

            return isChanged;
        }

        public static int GetSlideIndexByTitle(string fileName, string slideTitle)
        {
            // Given a slide document and a slide title, retrieve the 0-based index of the 
            // first slide with a matching title. Return -1 if the title isn't found.

            // Assume that you won't find a match.
            int slideLocation = -1;

            using (var document = PresentationDocument.Open(fileName, true))
            {
                var presPart = document.PresentationPart;

                // No presentation part? Something's wrong with the document.
                if (presPart == null)
                {
                    throw new ArgumentException("fileName");
                }

                // If you're here, you know that presentationPart exists.
                var slideIdList = presPart.Presentation.SlideIdList;
                // Go through the slides in order.
                // This requires investigating the actual slide IDs, rather 
                // than just retrieving the slide parts.
                int index = 0;
                foreach (var slideId in slideIdList.Elements<SlideId>())
                {
                    SlidePart slidePart = (SlidePart)(presPart.GetPartById(slideId.RelationshipId.ToString()));

                    if (slidePart == null)
                    {
                        throw new ArgumentNullException("presentationDocument");
                    }

                    Slide theSlide = slidePart.Slide;
                    if (theSlide != null)
                    {

                        // Assume the first title shape you find contains the title.
                        var titleShape = slidePart.Slide.Descendants<A.Shape>().
                          Where(s => IsTitleShape(s)).FirstOrDefault();
                        if (titleShape != null)
                        {
                            // Compare the title, case-insensitively.
                            if (string.Compare(titleShape.InnerText, slideTitle, true) == 0)
                            {
                                slideLocation = index;
                                break;
                            }
                            else
                            {
                                index += 1;
                            }
                        }
                    }
                }
            }
            return slideLocation;
        }

        private static bool IsTitleShape(A.Shape shape)
        {
            bool isTitle = false;

            var placeholderShape = shape.NonVisualShapeProperties.NonVisualDrawingProperties.GetFirstChild<PlaceholderShape>();
            if (((placeholderShape) != null) && (((placeholderShape.Type) != null) &&
              placeholderShape.Type.HasValue))
            {
                // Any title shape
                if (placeholderShape.Type.Value == PlaceholderValues.Title)
                {
                    isTitle = true;

                }
                // A centered title.
                else if (placeholderShape.Type.Value == PlaceholderValues.CenteredTitle)
                {
                    isTitle = true;
                }
            }
            return isTitle;
        }

        // Return the number of slides, including hidden slides.
        public static int GetSlideCount(string fileName, bool includeHidden)
        {
            int slidesCount = 0;

            using (PresentationDocument doc = PresentationDocument.Open(fileName, false))
            {
                // Get the presentation part of the document.
                PresentationPart presentationPart = doc.PresentationPart;
                if (presentationPart != null)
                {
                    if (includeHidden)
                    {
                        slidesCount = presentationPart.GetPartsOfType<SlidePart>().Count();
                    }
                    else
                    {
                        // Each slide can include a Show property, which if hidden will contain the value "0".
                        // The Show property may not exist, and most likely will not, for non-hidden slides.
                        var slides = presentationPart.GetPartsOfType<SlidePart>().
                          Where((s) => (s.Slide != null) &&
                            ((s.Slide.Show == null) || (s.Slide.Show.HasValue && s.Slide.Show.Value)));
                        slidesCount = slides.Count();
                    }
                }
            }
            return slidesCount;
        }
    }
}
