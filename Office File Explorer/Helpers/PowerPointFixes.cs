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

namespace Office_File_Explorer.Helpers
{
    class PowerPointFixes
    {
        public static bool CustomResetNotesPageSize(string filePath)
        {
            bool isFixed = false;

            using (PresentationDocument document = PresentationDocument.Open(filePath, true))
            {
                NotesSlides nsh = GetNotesPageSizesFromFile();

                if (nsh.pNotesSz.Cx == 0)
                {
                    return isFixed;
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

            return isFixed;
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

        public static bool ResetNotesPageSize(string filePath)
        {
            bool isFixed = false;

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

            return isFixed;
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
            pDoc.Save();
        }
    }
}
