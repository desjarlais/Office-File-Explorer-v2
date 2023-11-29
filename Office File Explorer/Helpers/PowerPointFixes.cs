using System.Collections.Generic;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using PShape = DocumentFormat.OpenXml.Presentation.Shape;
using NonVisualDrawingProperties = DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties;
using TextBody = DocumentFormat.OpenXml.Presentation.TextBody;
using System.Linq;
using System.Windows.Forms;
using System;
using DocumentFormat.OpenXml;

namespace Office_File_Explorer.Helpers
{
    class PowerPointFixes
    {
        public static void CustomResetNotesPageSize(string filePath)
        {
            using (PresentationDocument document = PresentationDocument.Open(filePath, true))
            {
                NotesSlides nsh = GetNotesPageSizesFromFile();

                if (nsh.pNotesSz.Cx == 0)
                {
                    return;
                }

                // Get the presentation part of document
                PresentationPart presentationPart = document.PresentationPart;

                if (presentationPart != null)
                {
                    Presentation p = presentationPart.Presentation;

                    // Step 1 : Resize the presentation notesz prop
                    NotesSize defaultNotesSize = new NotesSize() { Cx = nsh.pNotesSz.Cx, Cy = nsh.pNotesSz.Cy };

                    // first reset the notes size values
                    p.NotesSize = defaultNotesSize;

                    // now save up the part
                    p.Save();

                    // Step 2 : loop the shapes in the notes master and reset their sizes
                    if (Properties.Settings.Default.ResetNotesMaster == true)
                    {
                        // we need to reset sizes in the notes master for each shape
                        ShapeTree mSt = presentationPart.NotesMasterPart.NotesMaster.CommonSlideData.ShapeTree;

                        foreach (var mShp in mSt)
                        {
                            if (mShp.ToString() == Strings.dfopShape)
                            {
                                PShape ps = (PShape)mShp;
                                NonVisualDrawingProperties nvdpr = ps.NonVisualShapeProperties.NonVisualDrawingProperties;
                                Transform2D t2d = ps.ShapeProperties.Transform2D;

                                if (nvdpr.Name.ToString().Contains(Strings.pptHeaderPlaceholder))
                                {
                                    t2d.Offset.X = nsh.t2dHeader.OffsetX;
                                    t2d.Offset.Y = nsh.t2dHeader.OffsetY;
                                    t2d.Extents.Cx = nsh.t2dHeader.ExtentsCx;
                                    t2d.Extents.Cy = nsh.t2dHeader.ExtentsCy;
                                }

                                if (nvdpr.Name.ToString().Contains(Strings.pptDatePlaceholder))
                                {
                                    t2d.Offset.X = nsh.t2dDate.OffsetX;
                                    t2d.Offset.Y = nsh.t2dDate.OffsetY;
                                    t2d.Extents.Cx = nsh.t2dDate.ExtentsCx;
                                    t2d.Extents.Cy = nsh.t2dDate.ExtentsCy;
                                }

                                if (nvdpr.Name.ToString().Contains(Strings.pptSlideImagePlaceholder))
                                {
                                    t2d.Offset.X = nsh.t2dSlideImage.OffsetX;
                                    t2d.Offset.Y = nsh.t2dSlideImage.OffsetY;
                                    t2d.Extents.Cx = nsh.t2dSlideImage.ExtentsCx;
                                    t2d.Extents.Cy = nsh.t2dSlideImage.ExtentsCy;
                                }

                                if (nvdpr.Name.ToString().Contains(Strings.pptNotesPlaceholder))
                                {
                                    t2d.Offset.X = nsh.t2dNotes.OffsetX;
                                    t2d.Offset.Y = nsh.t2dNotes.OffsetY;
                                    t2d.Extents.Cx = nsh.t2dNotes.ExtentsCx;
                                    t2d.Extents.Cy = nsh.t2dNotes.ExtentsCy;
                                }

                                if (nvdpr.Name.ToString().Contains(Strings.pptFooterPlaceholder))
                                {
                                    t2d.Offset.X = nsh.t2dFooter.OffsetX;
                                    t2d.Offset.Y = nsh.t2dFooter.OffsetY;
                                    t2d.Extents.Cx = nsh.t2dFooter.ExtentsCx;
                                    t2d.Extents.Cy = nsh.t2dFooter.ExtentsCy;
                                }

                                if (nvdpr.Name.ToString().Contains(Strings.pptSlideNumberPlaceholder))
                                {
                                    t2d.Offset.X = nsh.t2dSlideNumber.OffsetX;
                                    t2d.Offset.Y = nsh.t2dSlideNumber.OffsetY;
                                    t2d.Extents.Cx = nsh.t2dSlideNumber.ExtentsCx;
                                    t2d.Extents.Cy = nsh.t2dSlideNumber.ExtentsCy;
                                }

                                if (nvdpr.Name == Strings.pptPicture)
                                {
                                    t2d.Offset.X = nsh.t2dPicture.OffsetX;
                                    t2d.Offset.Y = nsh.t2dPicture.OffsetY;
                                    t2d.Extents.Cx = nsh.t2dPicture.ExtentsCx;
                                    t2d.Extents.Cy = nsh.t2dPicture.ExtentsCy;
                                }
                            }
                        }

                        // Step 3 : we need to delete the size values for placeholders on each notes slide
                        foreach (var slideId in p.SlideIdList.Elements<SlideId>())
                        {
                            SlidePart slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                            ShapeTree st = slidePart.NotesSlidePart.NotesSlide.CommonSlideData.ShapeTree;
                            List<RunProperties> rpList = slidePart.NotesSlidePart.NotesSlide.Descendants<RunProperties>().ToList();

                            foreach (var s in st)
                            {
                                // we only want to make changes to the shapes
                                if (s.ToString() == Strings.dfopShape)
                                {
                                    PShape ps = (PShape)s;
                                    NonVisualDrawingProperties nvdpr = ps.NonVisualShapeProperties.NonVisualDrawingProperties;
                                    Transform2D t2d = ps.ShapeProperties.Transform2D;

                                    if (t2d is null)
                                    {
                                        A.Transform2D t2dn = new Transform2D();
                                        A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
                                        A.Extents extents1 = new A.Extents() { Cx = 0L, Cy = 0L };
                                        t2d = t2dn;
                                        t2d.Offset = offset1;
                                        t2d.Extents = extents1;
                                    }

                                    if (nvdpr.Name.ToString().Contains(Strings.pptHeaderPlaceholder))
                                    {
                                        t2d.Offset.X = nsh.t2dHeader.OffsetX;
                                        t2d.Offset.Y = nsh.t2dHeader.OffsetY;
                                        t2d.Extents.Cx = nsh.t2dHeader.ExtentsCx;
                                        t2d.Extents.Cy = nsh.t2dHeader.ExtentsCy;
                                    }

                                    if (nvdpr.Name.ToString().Contains(Strings.pptDatePlaceholder))
                                    {
                                        t2d.Offset.X = nsh.t2dDate.OffsetX;
                                        t2d.Offset.Y = nsh.t2dDate.OffsetY;
                                        t2d.Extents.Cx = nsh.t2dDate.ExtentsCx;
                                        t2d.Extents.Cy = nsh.t2dDate.ExtentsCy;
                                    }

                                    if (nvdpr.Name.ToString().Contains(Strings.pptSlideImagePlaceholder))
                                    {
                                        t2d.Offset.X = nsh.t2dSlideImage.OffsetX;
                                        t2d.Offset.Y = nsh.t2dSlideImage.OffsetY;
                                        t2d.Extents.Cx = nsh.t2dSlideImage.ExtentsCx;
                                        t2d.Extents.Cy = nsh.t2dSlideImage.ExtentsCy;
                                    }

                                    if (nvdpr.Name.ToString().Contains(Strings.pptNotesPlaceholder))
                                    {
                                        t2d.Offset.X = nsh.t2dNotes.OffsetX;
                                        t2d.Offset.Y = nsh.t2dNotes.OffsetY;
                                        t2d.Extents.Cx = nsh.t2dNotes.ExtentsCx;
                                        t2d.Extents.Cy = nsh.t2dNotes.ExtentsCy;
                                    }

                                    if (nvdpr.Name.ToString().Contains(Strings.pptFooterPlaceholder))
                                    {
                                        t2d.Offset.X = nsh.t2dFooter.OffsetX;
                                        t2d.Offset.Y = nsh.t2dFooter.OffsetY;
                                        t2d.Extents.Cx = nsh.t2dFooter.ExtentsCx;
                                        t2d.Extents.Cy = nsh.t2dFooter.ExtentsCy;
                                    }

                                    if (nvdpr.Name.ToString().Contains(Strings.pptSlideNumberPlaceholder))
                                    {
                                        t2d.Offset.X = nsh.t2dSlideNumber.OffsetX;
                                        t2d.Offset.Y = nsh.t2dSlideNumber.OffsetY;
                                        t2d.Extents.Cx = nsh.t2dSlideNumber.ExtentsCx;
                                        t2d.Extents.Cy = nsh.t2dSlideNumber.ExtentsCy;
                                    }
                                }
                                else if (s.ToString() == Strings.dfopPresentationPicture)
                                {
                                    DocumentFormat.OpenXml.Presentation.Picture pic = (DocumentFormat.OpenXml.Presentation.Picture)s;
                                    Transform2D t2d = pic.ShapeProperties.Transform2D;

                                    // there are times when pictures get moved with the rest of the slide objects, need to reset those back
                                    if (t2d is null)
                                    {
                                        t2d.Offset.X = nsh.t2dPicture.OffsetX;
                                        t2d.Offset.Y = nsh.t2dPicture.OffsetY;
                                        t2d.Extents.Cx = nsh.t2dPicture.ExtentsCx;
                                        t2d.Extents.Cy = nsh.t2dPicture.ExtentsCy;
                                    }
                                    else
                                    {
                                        t2d.Offset.X = 217831L;
                                        t2d.Offset.Y = 4470109L;
                                        t2d.Extents.Cx = 3249763L;
                                        t2d.Extents.Cy = 2795946L;
                                    }
                                }
                            }

                            foreach (RunProperties r in rpList)
                            {
                                r.FontSize = 1200;
                            }
                        }
                    }

                    p.Save();
                }
            }
        }

        public static NotesSlides GetNotesPageSizesFromFile()
        {
            NotesSlides nsh = new NotesSlides();

            OpenFileDialog fDialog = new OpenFileDialog
            {
                Title = "Select PowerPoint File.",
                Filter = "PowerPoint | *.pptx",
                RestoreDirectory = true,
                InitialDirectory = @"%userprofile%"
            };

            if (fDialog.ShowDialog() == DialogResult.OK)
            {
                using (PresentationDocument document = PresentationDocument.Open(fDialog.FileName, false))
                {
                    nsh.pNotesSz.Cx = document.PresentationPart.Presentation.NotesSize.Cx;
                    nsh.pNotesSz.Cy = document.PresentationPart.Presentation.NotesSize.Cy;

                    ShapeTree mSt = document.PresentationPart.NotesMasterPart.NotesMaster.CommonSlideData.ShapeTree;

                    foreach (var mShp in mSt)
                    {
                        if (mShp.ToString() == Strings.dfopShape)
                        {
                            PShape ps = (PShape)mShp;
                            NonVisualDrawingProperties nvdpr = ps.NonVisualShapeProperties.NonVisualDrawingProperties;
                            Transform2D t2d = ps.ShapeProperties.Transform2D;

                            if (nvdpr.Name == Strings.pptHeaderPlaceholder1)
                            {
                                nsh.t2dHeader.OffsetX = t2d.Offset.X;
                                nsh.t2dHeader.OffsetY = t2d.Offset.Y;
                                nsh.t2dHeader.ExtentsCx = t2d.Extents.Cx;
                                nsh.t2dHeader.ExtentsCy = t2d.Extents.Cy;
                            }

                            if (nvdpr.Name == Strings.pptDatePlaceholder2)
                            {
                                nsh.t2dDate.OffsetX = t2d.Offset.X;
                                nsh.t2dDate.OffsetY = t2d.Offset.Y;
                                nsh.t2dDate.ExtentsCx = t2d.Extents.Cx;
                                nsh.t2dDate.ExtentsCy = t2d.Extents.Cy;
                            }

                            if (nvdpr.Name == Strings.pptSlideImagePlaceholder3)
                            {
                                nsh.t2dSlideImage.OffsetX = t2d.Offset.X;
                                nsh.t2dSlideImage.OffsetY = t2d.Offset.Y;
                                nsh.t2dSlideImage.ExtentsCx = t2d.Extents.Cx;
                                nsh.t2dSlideImage.ExtentsCy = t2d.Extents.Cy;
                            }

                            if (nvdpr.Name == Strings.pptNotesPlaceholder4)
                            {
                                nsh.t2dNotes.OffsetX = t2d.Offset.X;
                                nsh.t2dNotes.OffsetY = t2d.Offset.Y;
                                nsh.t2dNotes.ExtentsCx = t2d.Extents.Cx;
                                nsh.t2dNotes.ExtentsCy = t2d.Extents.Cy;
                            }

                            if (nvdpr.Name == Strings.pptFooterPlaceholder5)
                            {
                                nsh.t2dFooter.OffsetX = t2d.Offset.X;
                                nsh.t2dFooter.OffsetY = t2d.Offset.Y;
                                nsh.t2dFooter.ExtentsCx = t2d.Extents.Cx;
                                nsh.t2dFooter.ExtentsCy = t2d.Extents.Cy;
                            }

                            if (nvdpr.Name == Strings.pptSlideNumberPlaceholder6)
                            {
                                nsh.t2dSlideNumber.OffsetX = t2d.Offset.X;
                                nsh.t2dSlideNumber.OffsetY = t2d.Offset.Y;
                                nsh.t2dSlideNumber.ExtentsCx = t2d.Extents.Cx;
                                nsh.t2dSlideNumber.ExtentsCy = t2d.Extents.Cy;
                            }

                            if (nvdpr.Name == Strings.pptPicture)
                            {
                                nsh.t2dPicture.OffsetX = t2d.Offset.X;
                                nsh.t2dPicture.OffsetY = t2d.Offset.Y;
                                nsh.t2dPicture.ExtentsCx = t2d.Extents.Cx;
                                nsh.t2dPicture.ExtentsCy = t2d.Extents.Cy;
                            }
                        }
                    }
                }
            }

            return nsh;
        }

        /// <summary>
        /// when using Bittitan to convert google docs to PowerPoint presentations, the files might have invalid margins
        /// usually, there will be a right margin set to 0 in the presentation, which causes the tab to not work since there is nowhere to go
        /// this will check for this scenario and reset the margins to default values
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool ResetDefaultParagraphProps(string filePath)
        {
            bool isFixed = false;

            using (PresentationDocument document = PresentationDocument.Open(filePath, true))
            {
                // setup the default paragraph properties and each level paragraph properties
                DefaultParagraphProperties defaultParagraphProperties1 = new DefaultParagraphProperties() { LeftMargin = 0, Level = 0, Alignment = TextAlignmentTypeValues.Left, RightToLeft = false };
                Level1ParagraphProperties level1ParagraphProperties1 = new Level1ParagraphProperties() { LeftMargin = 0, Level = 0, Alignment = TextAlignmentTypeValues.Left, RightToLeft = false };
                Level2ParagraphProperties level2ParagraphProperties1 = new Level2ParagraphProperties() { LeftMargin = 457200, Level = 1, Alignment = TextAlignmentTypeValues.Left, RightToLeft = false };
                Level3ParagraphProperties level3ParagraphProperties1 = new Level3ParagraphProperties() { LeftMargin = 914400, Level = 2, Alignment = TextAlignmentTypeValues.Left, RightToLeft = false };
                Level4ParagraphProperties level4ParagraphProperties1 = new Level4ParagraphProperties() { LeftMargin = 1371600, Level = 3, Alignment = TextAlignmentTypeValues.Left, RightToLeft = false };
                Level5ParagraphProperties level5ParagraphProperties1 = new Level5ParagraphProperties() { LeftMargin = 1828800, Level = 4, Alignment = TextAlignmentTypeValues.Left, RightToLeft = false };
                Level6ParagraphProperties level6ParagraphProperties1 = new Level6ParagraphProperties() { LeftMargin = 2286000, Level = 5, Alignment = TextAlignmentTypeValues.Left, RightToLeft = false };
                Level7ParagraphProperties level7ParagraphProperties1 = new Level7ParagraphProperties() { LeftMargin = 2743200, Level = 6, Alignment = TextAlignmentTypeValues.Left, RightToLeft = false };
                Level8ParagraphProperties level8ParagraphProperties1 = new Level8ParagraphProperties() { LeftMargin = 3200400, Level = 7, Alignment = TextAlignmentTypeValues.Left, RightToLeft = false };
                Level9ParagraphProperties level9ParagraphProperties1 = new Level9ParagraphProperties() { LeftMargin = 3657600, Level = 8, Alignment = TextAlignmentTypeValues.Left, RightToLeft = false };

                // check defaultparagraphproperties
                if (document.PresentationPart.Presentation.DefaultTextStyle.DefaultParagraphProperties.RightToLeft == false
                    && document.PresentationPart.Presentation.DefaultTextStyle.DefaultParagraphProperties.LeftMargin is null)
                {
                    if (document.PresentationPart.Presentation.DefaultTextStyle.DefaultParagraphProperties.InnerXml.Contains("marR=\"0\""))
                    {
                        document.PresentationPart.Presentation.DefaultTextStyle.DefaultParagraphProperties.InnerXml.Replace("marR=\"0\"", "marL=\"0\"");
                    }
                    else
                    {
                        document.PresentationPart.Presentation.DefaultTextStyle.DefaultParagraphProperties = defaultParagraphProperties1;
                    }
                    
                    isFixed = true;
                }

                // check each level paragraphproperties
                if (document.PresentationPart.Presentation.DefaultTextStyle.Level1ParagraphProperties.RightToLeft == false
                    && document.PresentationPart.Presentation.DefaultTextStyle.Level1ParagraphProperties.LeftMargin is null)
                {
                    if (document.PresentationPart.Presentation.DefaultTextStyle.Level1ParagraphProperties.InnerXml.Contains("marR=\"0\""))
                    {
                        document.PresentationPart.Presentation.DefaultTextStyle.Level1ParagraphProperties.InnerXml.Replace("marR=\"0\"", "marL=\"0\"");
                    }
                    else
                    {
                        document.PresentationPart.Presentation.DefaultTextStyle.Level1ParagraphProperties = level1ParagraphProperties1;
                    }
                    
                    isFixed = true;
                }

                if (document.PresentationPart.Presentation.DefaultTextStyle.Level2ParagraphProperties.RightToLeft == false
                    && document.PresentationPart.Presentation.DefaultTextStyle.Level2ParagraphProperties.LeftMargin is null)
                {
                    if (document.PresentationPart.Presentation.DefaultTextStyle.Level2ParagraphProperties.InnerXml.Contains("marR=\"0\""))
                    {
                        document.PresentationPart.Presentation.DefaultTextStyle.Level2ParagraphProperties.InnerXml.Replace("marR=\"0\"", "marL=\"457200\"");
                    }
                    else
                    {
                        document.PresentationPart.Presentation.DefaultTextStyle.Level2ParagraphProperties = level2ParagraphProperties1;
                    }
                    isFixed = true;
                }

                if (document.PresentationPart.Presentation.DefaultTextStyle.Level3ParagraphProperties.RightToLeft == false
                    && document.PresentationPart.Presentation.DefaultTextStyle.Level3ParagraphProperties.LeftMargin is null)
                {
                    if (document.PresentationPart.Presentation.DefaultTextStyle.Level3ParagraphProperties.InnerXml.Contains("marR=\"0\""))
                    {
                        document.PresentationPart.Presentation.DefaultTextStyle.Level3ParagraphProperties.InnerXml.Replace("marR=\"0\"", "marL=\"914400\"");
                    }
                    else
                    {
                        document.PresentationPart.Presentation.DefaultTextStyle.Level3ParagraphProperties = level3ParagraphProperties1;
                    }
                    isFixed = true;
                }

                if (document.PresentationPart.Presentation.DefaultTextStyle.Level4ParagraphProperties.RightToLeft == false
                    && document.PresentationPart.Presentation.DefaultTextStyle.Level4ParagraphProperties.LeftMargin is null)
                {
                    if (document.PresentationPart.Presentation.DefaultTextStyle.Level4ParagraphProperties.InnerXml.Contains("marR=\"0\""))
                    {
                        document.PresentationPart.Presentation.DefaultTextStyle.Level4ParagraphProperties.InnerXml.Replace("marR=\"0\"", "marL=\"1371600\"");
                    }
                    else
                    {
                        document.PresentationPart.Presentation.DefaultTextStyle.Level4ParagraphProperties = level4ParagraphProperties1;
                    }
                    isFixed = true;
                }

                if (document.PresentationPart.Presentation.DefaultTextStyle.Level5ParagraphProperties.RightToLeft == false
                    && document.PresentationPart.Presentation.DefaultTextStyle.Level5ParagraphProperties.LeftMargin is null)
                {
                    if (document.PresentationPart.Presentation.DefaultTextStyle.Level5ParagraphProperties.InnerXml.Contains("marR=\"0\""))
                    {
                        document.PresentationPart.Presentation.DefaultTextStyle.Level5ParagraphProperties.InnerXml.Replace("marR=\"0\"", "marL=\"1828800\"");
                    }
                    else
                    {
                        document.PresentationPart.Presentation.DefaultTextStyle.Level5ParagraphProperties = level5ParagraphProperties1;
                    }
                    isFixed = true;
                }

                if (document.PresentationPart.Presentation.DefaultTextStyle.Level6ParagraphProperties.RightToLeft == false
                    && document.PresentationPart.Presentation.DefaultTextStyle.Level6ParagraphProperties.LeftMargin is null)
                {
                    if (document.PresentationPart.Presentation.DefaultTextStyle.Level6ParagraphProperties.InnerXml.Contains("marR=\"0\""))
                    {
                        document.PresentationPart.Presentation.DefaultTextStyle.Level6ParagraphProperties.InnerXml.Replace("marR=\"0\"", "marL=\"2286000\"");
                    }
                    else
                    {
                        document.PresentationPart.Presentation.DefaultTextStyle.Level6ParagraphProperties = level6ParagraphProperties1;
                    }
                    isFixed = true;
                }

                if (document.PresentationPart.Presentation.DefaultTextStyle.Level7ParagraphProperties.RightToLeft == false
                    && document.PresentationPart.Presentation.DefaultTextStyle.Level7ParagraphProperties.LeftMargin is null)
                {
                    if (document.PresentationPart.Presentation.DefaultTextStyle.Level7ParagraphProperties.InnerXml.Contains("marR=\"0\""))
                    {
                        document.PresentationPart.Presentation.DefaultTextStyle.Level7ParagraphProperties.InnerXml.Replace("marR=\"0\"", "marL=\"2743200\"");
                    }
                    else
                    {
                        document.PresentationPart.Presentation.DefaultTextStyle.Level7ParagraphProperties = level7ParagraphProperties1;
                    }
                    isFixed = true;
                }

                if (document.PresentationPart.Presentation.DefaultTextStyle.Level8ParagraphProperties.RightToLeft == false
                    && document.PresentationPart.Presentation.DefaultTextStyle.Level8ParagraphProperties.LeftMargin is null)
                {
                    if (document.PresentationPart.Presentation.DefaultTextStyle.Level8ParagraphProperties.InnerXml.Contains("marR=\"0\""))
                    {
                        document.PresentationPart.Presentation.DefaultTextStyle.Level8ParagraphProperties.InnerXml.Replace("marR=\"0\"", "marL=\"3200400\"");
                    }
                    else
                    {
                        document.PresentationPart.Presentation.DefaultTextStyle.Level8ParagraphProperties = level8ParagraphProperties1;
                    }
                    isFixed = true;
                }

                if (document.PresentationPart.Presentation.DefaultTextStyle.Level9ParagraphProperties.RightToLeft == false
                    && document.PresentationPart.Presentation.DefaultTextStyle.Level9ParagraphProperties.LeftMargin is null)
                {
                    if (document.PresentationPart.Presentation.DefaultTextStyle.Level9ParagraphProperties.InnerXml.Contains("marR=\"0\""))
                    {
                        document.PresentationPart.Presentation.DefaultTextStyle.Level9ParagraphProperties.InnerXml.Replace("marR=\"0\"", "marL=\"3657600\"");
                    }
                    else
                    {
                        document.PresentationPart.Presentation.DefaultTextStyle.Level9ParagraphProperties = level9ParagraphProperties1;
                    }
                    isFixed = true;
                }
            }

            return isFixed;
        }

        /// <summary>
        /// when using Bittitan to convert google docs to PowerPoint presentations, some shapes will have incorrect indent levels
        /// this fix will check for these scenarios and reset the levels
        /// NOTE: this fix usually requires ResetDefaultParagraphProps, so call both of these fixes
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool FixMissingPlaceholder(string filePath)
        {
            bool isFixed = false;

            using (PresentationDocument document = PresentationDocument.Open(filePath, true))
            {
                // check slides 
                foreach (SlidePart sp in document.PresentationPart.SlideParts)
                {
                    // check the textbox placeholder and indent levels
                    Slide sld = sp.Slide;
                    foreach (OpenXmlElement oxe in sld.CommonSlideData.ShapeTree)
                    {
                        if (oxe.LocalName == "sp")
                        {
                            bool isTxBox = false;
                            // first loop through the shapes and see if cNvSpPr has a txBox
                            // if it does not, no checks needed
                            foreach (OpenXmlElement oxeShp in oxe.ChildElements)
                            {
                                // if the shape is not a textbox, skip it
                                // if it has a textbox and cnvpr has an id value, skip it
                                if (oxeShp.LocalName == "nvSpPr")
                                {
                                    foreach (OpenXmlElement oxeNvSpPr in oxeShp.ChildElements)
                                    {
                                        if (oxeNvSpPr.LocalName == "cNvSpPr")
                                        {
                                            if (oxeNvSpPr.OuterXml.Contains("txBox=\"1\""))
                                            {
                                                isTxBox = true;
                                            }
                                        }
                                    }
                                }
                            }

                            // if a txbody was found, continue looking for missing tags
                            if (isTxBox)
                            {
                                foreach (OpenXmlElement oxeShape in oxe.ChildElements)
                                {
                                    // textbody contains level
                                    if (oxeShape.LocalName == "txBody")
                                    {
                                        // check for bad indent levels
                                        if (Properties.Settings.Default.ResetIndentLevels == true)
                                        {
                                            // if nvpr was empty, loop paragraphs and check for levels > 0 and reset to 0
                                            foreach (OpenXmlElement oxeTxBody in oxeShape.ChildElements)
                                            {
                                                if (oxeTxBody.LocalName == "p")
                                                {
                                                    foreach (OpenXmlElement oxeTxBodyPara in oxeTxBody.ChildElements)
                                                    {
                                                        if (oxeTxBodyPara.LocalName == "pPr")
                                                        {
                                                            try
                                                            {
                                                                TextParagraphPropertiesType textParagraphProperties = (TextParagraphPropertiesType)oxeTxBodyPara;
                                                                if (textParagraphProperties.Level == 8)
                                                                {
                                                                    textParagraphProperties.Level = 0;
                                                                    isFixed = true;
                                                                }
                                                            }
                                                            catch (Exception)
                                                            {
                                                                // skip null props
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
                }

                // save out the file
                if (isFixed)
                {
                    document.Save();
                }
            }

            return isFixed;
        }

        /// <summary>
        /// some files will have missing / orphaned custData tag rels
        /// this is a corrupt file scenario and this fix will remove those custData tags
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool FixMissingRelIds(string filePath)
        {
            bool isFixed = false;

            using (PresentationDocument document = PresentationDocument.Open(filePath, true))
            {
                // first get the list of presentation parts
                IEnumerable<IdPartPair> ipp = document.PresentationPart.Parts.OfType<IdPartPair>().ToList();
                bool idFound = false;

                if (document.PresentationPart.Presentation.CustomerDataList == null)
                {
                    return isFixed;
                }

                try
                {
                CdlStart:
                    foreach (CustomerData cd in document.PresentationPart.Presentation.CustomerDataList)
                    {
                        // check each custData tag id against the presentation rels
                        // if we don't find a rel id, then the custData id is missing and we need to delete it
                        idFound = false;
                        foreach (IdPartPair ippTemp in ipp)
                        {
                            if (ippTemp.RelationshipId == cd.Id)
                            {
                                idFound = true;
                            }
                        }

                        if (!idFound)
                        {
                            cd.Remove();
                            isFixed = true;
                            document.Save();
                            goto CdlStart;
                        }
                    }
                }
                catch (Exception)
                {
                    return isFixed;
                }
            }

            return isFixed;
        }

        public static void ResetNotesPageSize(string filePath)
        {
            using (PresentationDocument document = PresentationDocument.Open(filePath, true))
            {
                // Get the presentation part of document
                PresentationPart presentationPart = document.PresentationPart;

                if (presentationPart != null)
                {
                    Presentation p = presentationPart.Presentation;

                    // Step 1 : Resize the presentation notesz prop
                    // if the notes size is already the default, no need to make any changes
                    if (p.NotesSize.Cx != 6858000 || p.NotesSize.Cy != 9144000)
                    {
                        // setup default size
                        NotesSize defaultNotesSize = new NotesSize() { Cx = 6858000L, Cy = 9144000L };

                        // first reset the notes size values
                        p.NotesSize = defaultNotesSize;

                        // now save up the part
                        p.Save();
                    }

                    // Step 2 : loop the shapes in the notes master and reset their sizes
                    // need to find a way to flag a file if the notes master and/or notes slides become corrupt
                    // hiding behind a setting checkbox for now
                    if (Properties.Settings.Default.ResetNotesMaster == true)
                    {
                        // we need to reset sizes in the notes master for each shape
                        ShapeTree mSt = presentationPart.NotesMasterPart.NotesMaster.CommonSlideData.ShapeTree;

                        foreach (var mShp in mSt)
                        {
                            if (mShp.ToString() == Strings.dfopShape)
                            {
                                PShape ps = (PShape)mShp;
                                NonVisualDrawingProperties nvdpr = ps.NonVisualShapeProperties.NonVisualDrawingProperties;
                                Transform2D t2d = ps.ShapeProperties.Transform2D;

                                // use default values
                                if (nvdpr.Name == Strings.pptHeaderPlaceholder1)
                                {
                                    t2d.Offset.X = 0L;
                                    t2d.Offset.Y = 0L;
                                    t2d.Extents.Cx = 2971800L;
                                    t2d.Extents.Cy = 458788L;
                                }

                                if (nvdpr.Name == Strings.pptDatePlaceholder2)
                                {
                                    t2d.Offset.X = 3884613L;
                                    t2d.Offset.Y = 0L;
                                    t2d.Extents.Cx = 2971800L;
                                    t2d.Extents.Cy = 458788L;
                                }

                                if (nvdpr.Name == Strings.pptSlideImagePlaceholder3)
                                {
                                    t2d.Offset.X = 685800L;
                                    t2d.Offset.Y = 1143000L;
                                    t2d.Extents.Cx = 5486400L;
                                    t2d.Extents.Cy = 3086100L;
                                }

                                if (nvdpr.Name == Strings.pptNotesPlaceholder4)
                                {
                                    t2d.Offset.X = 685800L;
                                    t2d.Offset.Y = 4400550L;
                                    t2d.Extents.Cx = 5486400L;
                                    t2d.Extents.Cy = 3600450L;
                                }

                                if (nvdpr.Name == Strings.pptFooterPlaceholder5)
                                {
                                    t2d.Offset.X = 0L;
                                    t2d.Offset.Y = 8685213L;
                                    t2d.Extents.Cx = 2971800L;
                                    t2d.Extents.Cy = 458787L;
                                }

                                if (nvdpr.Name == Strings.pptSlideNumberPlaceholder6)
                                {
                                    t2d.Offset.X = 3884613L;
                                    t2d.Offset.Y = 8685213L;
                                    t2d.Extents.Cx = 2971800L;
                                    t2d.Extents.Cy = 458787L;
                                }
                            }
                        }

                        // Step 3 : we need to delete the size values for placeholders on each notes slide
                        foreach (var slideId in p.SlideIdList.Elements<SlideId>())
                        {
                            SlidePart slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                            ShapeTree st = slidePart.NotesSlidePart.NotesSlide.CommonSlideData.ShapeTree;
                            List<RunProperties> rpList = slidePart.NotesSlidePart.NotesSlide.Descendants<RunProperties>().ToList();

                            foreach (var s in st)
                            {
                                // we only want to make changes to the shapes
                                if (s.ToString() == Strings.dfopShape)
                                {
                                    PShape ps = (PShape)s;
                                    Transform2D t2d = ps.ShapeProperties.Transform2D;
                                    TextBody tb = ps.TextBody;

                                    // if the transform exists, delete it for each shape
                                    if (t2d != null)
                                    {
                                        t2d.Remove();
                                    }

                                    // if there are drawing paragraph props, reset the margin and indent to 0
                                    if (ps.TextBody != null)
                                    {
                                        foreach (var x in tb.ChildElements)
                                        {
                                            if (x.ToString() == "DocumentFormat.OpenXml.Drawing.Paragraph")
                                            {
                                                Paragraph para = (Paragraph)x;
                                                if (para.ParagraphProperties != null)
                                                {
                                                    if (para.ParagraphProperties.LeftMargin != null)
                                                    {
                                                        para.ParagraphProperties.LeftMargin = 0;
                                                    }

                                                    if (para.ParagraphProperties.Indent != null)
                                                    {
                                                        para.ParagraphProperties.Indent = 0;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                else if (s.ToString() == Strings.dfopPresentationPicture)
                                {
                                    DocumentFormat.OpenXml.Presentation.Picture pic = (DocumentFormat.OpenXml.Presentation.Picture)s;
                                    Transform2D t2d = pic.ShapeProperties.Transform2D;

                                    // there are times when pictures get moved with the rest of the slide objects, need to reset those back
                                    if (t2d != null)
                                    {
                                        t2d.Offset.X = 217831L;
                                        t2d.Offset.Y = 4470109L;
                                        t2d.Extents.Cx = 3249763L;
                                        t2d.Extents.Cy = 2795946L;
                                    }
                                }
                            }

                            foreach (RunProperties r in rpList)
                            {
                                r.FontSize = 1200;
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Check the notes page size and reset values
        /// </summary>
        /// <param name="pDoc">oxml doc to change</param>
        public static void ChangeNotesPageSize(PresentationDocument pDoc)
        {
            // Get the presentation part of document
            PresentationPart presentationPart = pDoc.PresentationPart;

            if (presentationPart != null)
            {
                Presentation p = presentationPart.Presentation;

                // Step 1 : Resize the presentation notesz prop
                // if the notes size is already the default, no need to make any changes
                if (p.NotesSize.Cx != 6858000 || p.NotesSize.Cy != 9144000)
                {
                    // setup default size
                    NotesSize defaultNotesSize = new NotesSize() { Cx = 6858000L, Cy = 9144000L };

                    // first reset the notes size values
                    p.NotesSize = defaultNotesSize;

                    // now save up the part
                    p.Save();
                }

                // Step 2 : loop the shapes in the notes master and reset their sizes
                // need to find a way to flag a file if the notes master and/or notes slides become corrupt
                // hiding behind a setting checkbox for now
                if (Properties.Settings.Default.ResetNotesMaster == true)
                {
                    // we need to reset sizes in the notes master for each shape
                    ShapeTree mSt = presentationPart.NotesMasterPart.NotesMaster.CommonSlideData.ShapeTree;

                    foreach (var mShp in mSt)
                    {
                        if (mShp.ToString() == Strings.dfopShape)
                        {
                            PShape ps = (PShape)mShp;
                            NonVisualDrawingProperties nvdpr = ps.NonVisualShapeProperties.NonVisualDrawingProperties;
                            Transform2D t2d = ps.ShapeProperties.Transform2D;

                            // use default values
                            if (nvdpr.Name == Strings.pptHeaderPlaceholder1)
                            {
                                t2d.Offset.X = 0L;
                                t2d.Offset.Y = 0L;
                                t2d.Extents.Cx = 2971800L;
                                t2d.Extents.Cy = 458788L;
                            }

                            if (nvdpr.Name == Strings.pptDatePlaceholder2)
                            {
                                t2d.Offset.X = 3884613L;
                                t2d.Offset.Y = 0L;
                                t2d.Extents.Cx = 2971800L;
                                t2d.Extents.Cy = 458788L;
                            }

                            if (nvdpr.Name == Strings.pptSlideImagePlaceholder3)
                            {
                                t2d.Offset.X = 685800L;
                                t2d.Offset.Y = 1143000L;
                                t2d.Extents.Cx = 5486400L;
                                t2d.Extents.Cy = 3086100L;
                            }

                            if (nvdpr.Name == Strings.pptNotesPlaceholder4)
                            {
                                t2d.Offset.X = 685800L;
                                t2d.Offset.Y = 4400550L;
                                t2d.Extents.Cx = 5486400L;
                                t2d.Extents.Cy = 3600450L;
                            }

                            if (nvdpr.Name == Strings.pptFooterPlaceholder5)
                            {
                                t2d.Offset.X = 0L;
                                t2d.Offset.Y = 8685213L;
                                t2d.Extents.Cx = 2971800L;
                                t2d.Extents.Cy = 458787L;
                            }

                            if (nvdpr.Name == Strings.pptSlideNumberPlaceholder6)
                            {
                                t2d.Offset.X = 3884613L;
                                t2d.Offset.Y = 8685213L;
                                t2d.Extents.Cx = 2971800L;
                                t2d.Extents.Cy = 458787L;
                            }
                        }
                    }

                    // Step 3 : we need to delete the size values for placeholders on each notes slide
                    foreach (var slideId in p.SlideIdList.Elements<SlideId>())
                    {
                        SlidePart slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                        ShapeTree st = slidePart.NotesSlidePart.NotesSlide.CommonSlideData.ShapeTree;
                        List<RunProperties> rpList = slidePart.NotesSlidePart.NotesSlide.Descendants<RunProperties>().ToList();

                        foreach (var s in st)
                        {
                            // we only want to make changes to the shapes
                            if (s.ToString() == Strings.dfopShape)
                            {
                                PShape ps = (PShape)s;
                                Transform2D t2d = ps.ShapeProperties.Transform2D;
                                TextBody tb = ps.TextBody;

                                // if the transform exists, delete it for each shape
                                if (t2d != null)
                                {
                                    t2d.Remove();
                                }

                                // if there are drawing paragraph props, reset the margin and indent to 0
                                if (ps.TextBody != null)
                                {
                                    foreach (var x in tb.ChildElements)
                                    {
                                        if (x.ToString() == "DocumentFormat.OpenXml.Drawing.Paragraph")
                                        {
                                            Paragraph para = (Paragraph)x;
                                            if (para.ParagraphProperties != null)
                                            {
                                                if (para.ParagraphProperties.LeftMargin != null)
                                                {
                                                    para.ParagraphProperties.LeftMargin = 0;
                                                }

                                                if (para.ParagraphProperties.Indent != null)
                                                {
                                                    para.ParagraphProperties.Indent = 0;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            else if (s.ToString() == Strings.dfopPresentationPicture)
                            {
                                DocumentFormat.OpenXml.Presentation.Picture pic = (DocumentFormat.OpenXml.Presentation.Picture)s;
                                Transform2D t2d = pic.ShapeProperties.Transform2D;

                                // there are times when pictures get moved with the rest of the slide objects, need to reset those back
                                if (t2d != null)
                                {
                                    t2d.Offset.X = 217831L;
                                    t2d.Offset.Y = 4470109L;
                                    t2d.Extents.Cx = 3249763L;
                                    t2d.Extents.Cy = 2795946L;
                                }
                            }
                        }

                        foreach (RunProperties r in rpList)
                        {
                            r.FontSize = 1200;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// presentations can get bloated with master slide layouts which can cause perf issues in ppt online
        /// this can delete the unused slide layouts to help with perf problems
        /// </summary>
        /// <param name="pDoc"></param>
        public static void DeleteUnusedMasterLayouts(PresentationDocument pDoc)
        {
            // Get the presentation part of document
            PresentationPart presentationPart = pDoc.PresentationPart;

            if (presentationPart != null)
            {
                List<string> usedSlideLayoutIds = PowerPoint.GetSlideLayoutId(pDoc);

                // loop the existing slidelayouts and delete any that are not in the used list
                PowerPoint.DeleteUnusedSlideLayoutParts(pDoc, usedSlideLayoutIds);
            }
        }

        /// <summary>
        /// Reset the left margin for the paragraph properties
        /// Seen some files where all margins are set to 0
        /// this means tabs will not work for bullets/numbering
        /// well, it technically is working, if they all are 0 
        /// pressing tab has nowhere to go so it looks like it isn't working
        /// this will reset the levels to default values
        /// </summary>
        /// <param name="pDoc">oxml doc to change</param>
        public static void ResetBulletMargins(PresentationDocument pDoc)
        {
            // Get the presentation part of document
            PresentationPart presentationPart = pDoc.PresentationPart;

            if (presentationPart != null)
            {
                Presentation p = presentationPart.Presentation;
                // if all paragraph level left margins are set to 0, reset them to defaults
                if (p.DefaultTextStyle.Level2ParagraphProperties.LeftMargin == 0 &&
                    p.DefaultTextStyle.Level3ParagraphProperties.LeftMargin == 0 &&
                    p.DefaultTextStyle.Level4ParagraphProperties.LeftMargin == 0 &&
                    p.DefaultTextStyle.Level5ParagraphProperties.LeftMargin == 0 &&
                    p.DefaultTextStyle.Level6ParagraphProperties.LeftMargin == 0 &&
                    p.DefaultTextStyle.Level7ParagraphProperties.LeftMargin == 0 &&
                    p.DefaultTextStyle.Level8ParagraphProperties.LeftMargin == 0 &&
                    p.DefaultTextStyle.Level9ParagraphProperties.LeftMargin == 0)
                {
                    p.DefaultTextStyle.Level1ParagraphProperties.LeftMargin = 0;
                    p.DefaultTextStyle.Level2ParagraphProperties.LeftMargin = 457200;
                    p.DefaultTextStyle.Level3ParagraphProperties.LeftMargin = 914400;
                    p.DefaultTextStyle.Level4ParagraphProperties.LeftMargin = 1371600;
                    p.DefaultTextStyle.Level5ParagraphProperties.LeftMargin = 1828800;
                    p.DefaultTextStyle.Level6ParagraphProperties.LeftMargin = 2286000;
                    p.DefaultTextStyle.Level7ParagraphProperties.LeftMargin = 2743200;
                    p.DefaultTextStyle.Level8ParagraphProperties.LeftMargin = 3200400;
                    p.DefaultTextStyle.Level9ParagraphProperties.LeftMargin = 3657600;
                }
            }
            pDoc.Save();
        }
    }
}
