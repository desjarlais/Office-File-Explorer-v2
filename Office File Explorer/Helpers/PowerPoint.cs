﻿using DocumentFormat.OpenXml.Drawing;
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
using System.Xml;

namespace Office_File_Explorer.Helpers
{
    class PowerPoint
    {
        public static bool fSuccess;

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
        public static int RetrieveNumberOfSlides(string fileName, bool includeHidden = true)
        {
            int slidesCount = 0;

            using (PresentationDocument doc = PresentationDocument.Open(fileName, false))
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
                        var slides = presentationPart.SlideParts.Where(
                            (s) => (s.Slide is not null) &&
                              ((s.Slide.Show is null) || (s.Slide.Show.HasValue &&
                              s.Slide.Show.Value)));
                        slidesCount = slides.Count();
                    }
                }
            }
            return slidesCount;
        }

        public static List<string> GetHyperlinks(string path)
        {
            List<string> tList = new List<string>();

            using (PresentationDocument document = PresentationDocument.Open(path, false))
            {
                int linkCount = 0;
                foreach (string s in GetAllExternalHyperlinksInPresentation(path))
                {
                    linkCount++;
                    tList.Add(linkCount + Strings.wPeriod + s);
                }
            }

            return tList;
        }

        /// <summary>
        /// check for both legacy and modern comments
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static List<string> GetComments(string path)
        {
            List<string> tList = new List<string>();

            using (PresentationDocument presentationDocument = PresentationDocument.Open(path, false))
            {
                PresentationPart pPart = presentationDocument.PresentationPart;
                int commentCount = 0;

                // first check legacy comments
                foreach (SlidePart sPart in pPart.SlideParts)
                {
                    SlideCommentsPart sCPart = sPart.SlideCommentsPart;
                    if (sCPart is null)
                    {
                        continue;
                    }

                    foreach (Comment cmt in sCPart.CommentList)
                    {
                        commentCount++;
                        tList.Add(commentCount + Strings.wPeriod + cmt.InnerText);
                    }
                }

                // now check for modern comments
            }

            return tList;
        }

        public static List<string> GetSlideTitles(string path)
        {
            List<string> tList = new List<string>();

            using (PresentationDocument presentationDocument = PresentationDocument.Open(path, false))
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

        public static List<string> GetSlideText(string path)
        {
            List<string> tList = new List<string>();

            int sCount = RetrieveNumberOfSlides(path);
            if (sCount > 0)
            {
                int count = 0;

                do
                {
                    GetSlideIdAndText(out string sldText, path, count);
                    tList.Add("Slide " + (count + 1) + Strings.wPeriod + sldText);
                    count++;
                } while (count < sCount);
            }

            return tList;
        }

        public static List<string> GetSlideTransitions(string path)
        {
            List<string> tList = new List<string>();

            using (PresentationDocument ppt = PresentationDocument.Open(path, false))
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
        public static void GetSlideIdAndText(out string sldText, string docName, int index)
        {
            using (PresentationDocument ppt = PresentationDocument.Open(docName, false))
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
        public static IEnumerable<String> GetAllExternalHyperlinksInPresentation(string fileName)
        {
            // Declare a list of strings.
            List<string> ret = new List<string>();

            // Open the presentation file as read-only.
            using (PresentationDocument document = PresentationDocument.Open(fileName, false))
            {
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
            }

            // Return the list of strings.
            return ret;
        }
    }
}