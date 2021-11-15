﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;

// shortcut namespace refs
using P = DocumentFormat.OpenXml.Presentation;
using O = DocumentFormat.OpenXml;
using AO = DocumentFormat.OpenXml.Office.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using Path = System.IO.Path;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using System.Text;

namespace Office_File_Explorer.Helpers
{
    class Office
    {
        // custom document property types
        public enum PropertyTypes : int
        {
            YesNo,
            Text,
            DateTime,
            NumberInteger,
            NumberDouble
        }

        public static List<string> DisplayValidationErrorInformation(OpenXmlPackage docPackage)
        {
            OpenXmlValidator validator = new OpenXmlValidator();
            List<string> output = new List<string>();
            int count = 0;

            foreach (ValidationErrorInfo error in validator.Validate(docPackage))
            {
                count++;
                output.Add("Error " + count);
                output.Add("Description: " + error.Description);
                output.Add("Path: " + error.Path.XPath);
                output.Add("Part: " + error.Part.Uri);
                output.Add(Strings.wHeaderLine);
            }

            return output;
        }

        /// <summary>
        /// Given a document name, a property name/value, and the property type add a custom property to a document. 
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="propertyName"></param>
        /// <param name="propertyValue"></param>
        /// <param name="propertyType"></param>
        /// <param name="fileType"></param>
        /// <returns>returns the original value, if it existed</returns>
        public static string SetCustomProperty(string fileName, string propertyName, object propertyValue, PropertyTypes propertyType, string fileType)
        {
            string returnValue = null;

            var newProp = new CustomDocumentProperty();
            bool propSet = false;

            // Calculate the correct type.
            switch (propertyType)
            {
                case PropertyTypes.DateTime:

                    // Be sure you were passed a real date and if so, format in the correct way. 
                    // The date/time value passed in should represent a UTC date/time.
                    if ((propertyValue) is DateTime)
                    {
                        newProp.VTFileTime = new VTFileTime(string.Format("{0:s}Z", Convert.ToDateTime(propertyValue)));
                        propSet = true;
                    }

                    break;

                case PropertyTypes.NumberInteger:
                    if ((propertyValue) is int)
                    {
                        newProp.VTInt32 = new VTInt32(propertyValue.ToString());
                        propSet = true;
                    }

                    break;

                case PropertyTypes.NumberDouble:
                    if (propertyValue is double)
                    {
                        newProp.VTFloat = new VTFloat(propertyValue.ToString());
                        propSet = true;
                    }

                    break;

                case PropertyTypes.Text:
                    newProp.VTLPWSTR = new VTLPWSTR(propertyValue.ToString());
                    propSet = true;

                    break;

                case PropertyTypes.YesNo:
                    if (propertyValue is bool)
                    {
                        // Must be lowercase.
                        newProp.VTBool = new VTBool(Convert.ToBoolean(propertyValue).ToString().ToLower());
                        propSet = true;
                    }
                    break;
            }

            if (!propSet)
            {
                // If the code was not able to convert the property to a valid value, throw an exception.
                MessageBox.Show("The value entered does not match the specific type.  The value will be stored as text.", "Invalid Type", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                newProp.VTLPWSTR = new VTLPWSTR(propertyValue.ToString());
                propSet = true;
            }

            // Now that you have handled the parameters, start working on the document.
            newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
            newProp.Name = propertyName;

            if (fileType == "Word")
            {
                using (var document = WordprocessingDocument.Open(fileName, true))
                {
                    var customProps = document.CustomFilePropertiesPart;
                    if (customProps is null)
                    {
                        // No custom properties? Add the part, and the
                        // collection of properties now.
                        customProps = document.AddCustomFilePropertiesPart();
                        customProps.Properties = new O.CustomProperties.Properties();
                    }

                    var props = customProps.Properties;
                    if (!(props is null))
                    {
                        // This will trigger an exception if the property's Name 
                        // property is null, but if that happens, the property is damaged, 
                        // and probably should raise an exception.
                        var prop = props.Where(p => ((CustomDocumentProperty)p).Name.Value == propertyName).FirstOrDefault();

                        // Does the property exist? If so, get the return value, 
                        // and then delete the property.
                        if (prop is not null)
                        {
                            returnValue = prop.InnerText;
                            prop.Remove();
                        }

                        // Append the new property, and 
                        // fix up all the property ID values. 
                        // The PropertyId value must start at 2.
                        props.AppendChild(newProp);
                        int pid = 2;
                        foreach (CustomDocumentProperty item in props)
                        {
                            item.PropertyId = pid++;
                        }
                        props.Save();
                    }
                }
            }
            else if (fileType == "Excel")
            {
                using (var document = SpreadsheetDocument.Open(fileName, true))
                {
                    var customProps = document.CustomFilePropertiesPart;
                    if (customProps is null)
                    {
                        customProps = document.AddCustomFilePropertiesPart();
                        customProps.Properties = new O.CustomProperties.Properties();
                    }

                    var props = customProps.Properties;
                    if (props != null)
                    {
                        var prop = props.Where(p => ((CustomDocumentProperty)p).Name.Value == propertyName).FirstOrDefault();

                        if (prop != null)
                        {
                            returnValue = prop.InnerText;
                            prop.Remove();
                        }

                        props.AppendChild(newProp);
                        int pid = 2;
                        foreach (CustomDocumentProperty item in props)
                        {
                            item.PropertyId = pid++;
                        }
                        props.Save();
                    }
                }
            }
            else
            {
                using (var document = PresentationDocument.Open(fileName, true))
                {
                    var customProps = document.CustomFilePropertiesPart;
                    if (customProps is null)
                    {
                        customProps = document.AddCustomFilePropertiesPart();
                        customProps.Properties = new O.CustomProperties.Properties();
                    }

                    var props = customProps.Properties;
                    if (props != null)
                    {
                        var prop = props.Where(p => ((CustomDocumentProperty)p).Name.Value == propertyName).FirstOrDefault();

                        if (prop != null)
                        {
                            returnValue = prop.InnerText;
                            prop.Remove();
                        }

                        props.AppendChild(newProp);
                        int pid = 2;
                        foreach (CustomDocumentProperty item in props)
                        {
                            item.PropertyId = pid++;
                        }

                        props.Save();
                    }
                }
            }

            return returnValue;
        }

        /// <summary>
        /// replace the current theme with a user specified theme
        /// </summary>
        /// <param name="document">document file lcoation</param>
        /// <param name="themeFile">theme xml file location</param>
        /// <param name="app">which app is the document</param>
        public static void ReplaceTheme(string document, string themeFile, string app)
        {
            if (app == Strings.oAppWord)
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
                {
                    MainDocumentPart mainPart = wordDoc.MainDocumentPart;

                    // Delete the old document part.
                    mainPart.DeletePart(mainPart.ThemePart);

                    // Add a new document part and then add content.
                    ThemePart themePart = mainPart.AddNewPart<ThemePart>();

                    using (StreamReader streamReader = new StreamReader(themeFile))
                    using (StreamWriter streamWriter = new StreamWriter(themePart.GetStream(FileMode.Create)))
                    {
                        streamWriter.Write(streamReader.ReadToEnd());
                    }
                }
            }
            else if (app == Strings.oAppPowerPoint)
            {
                using (PresentationDocument presDoc = PresentationDocument.Open(document, true))
                {
                    PresentationPart mainPart = presDoc.PresentationPart;
                    mainPart.DeletePart(mainPart.ThemePart);
                    ThemePart themePart = mainPart.AddNewPart<ThemePart>();

                    using (StreamReader streamReader = new StreamReader(themeFile))
                    using (StreamWriter streamWriter = new StreamWriter(themePart.GetStream(FileMode.Create)))
                    {
                        streamWriter.Write(streamReader.ReadToEnd());
                    }
                }
            }
            else
            {
                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(document, true))
                {
                    WorkbookPart mainPart = excelDoc.WorkbookPart;
                    mainPart.DeletePart(mainPart.ThemePart);
                    ThemePart themePart = mainPart.AddNewPart<ThemePart>();

                    using (StreamReader streamReader = new StreamReader(themeFile))
                    using (StreamWriter streamWriter = new StreamWriter(themePart.GetStream(FileMode.Create)))
                    {
                        streamWriter.Write(streamReader.ReadToEnd());
                    }
                }
            }
        }

        /// <summary>
        /// Function to convert a macro enabled file to a non-macro enabled file
        /// </summary>
        /// <param name="fileName">file location</param>
        /// <param name="app">app type</param>
        /// <returns></returns>
        public static string ConvertMacroEnabled2NonMacroEnabled(string fileName, string app)
        {
            bool fileChanged = false;
            string newFileName = string.Empty;
            string fileExtension;

            if (app == Strings.oAppWord)
            {
                fileExtension = ".docx";
                using (WordprocessingDocument document = WordprocessingDocument.Open(fileName, true))
                {
                    // Access the main document part.
                    var docPart = document.MainDocumentPart;

                    // Look for the vbaProject part. If it is there, delete it.
                    var vbaPart = docPart.VbaProjectPart;
                    if (vbaPart != null)
                    {
                        // Delete the vbaProject part and then save the document.
                        docPart.DeletePart(vbaPart);
                        docPart.Document.Save();

                        // Change the document type to not macro-enabled.
                        document.ChangeDocumentType(WordprocessingDocumentType.Document);

                        // Track that the document has been changed.
                        fileChanged = true;
                    }
                }
            }
            else if (app == Strings.oAppPowerPoint)
            {
                fileExtension = ".pptx";
                using (PresentationDocument document = PresentationDocument.Open(fileName, true))
                {
                    var docPart = document.PresentationPart;
                    var vbaPart = docPart.VbaProjectPart;
                    if (vbaPart != null)
                    {
                        docPart.DeletePart(vbaPart);
                        docPart.Presentation.Save();
                        document.ChangeDocumentType(PresentationDocumentType.Presentation);
                        fileChanged = true;
                    }
                }
            }
            else
            {
                fileExtension = ".xlsx";
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true))
                {
                    var docPart = document.WorkbookPart;
                    var vbaPart = docPart.VbaProjectPart;
                    if (vbaPart is not null)
                    {
                        docPart.DeletePart(vbaPart);
                        docPart.Workbook.Save();
                        document.ChangeDocumentType(SpreadsheetDocumentType.Workbook);
                        fileChanged = true;
                    }
                }
            }

            // If anything goes wrong in this file handling,
            // the code will raise an exception back to the caller.
            if (fileChanged)
            {
                // Create the new filename.
                newFileName = Path.ChangeExtension(fileName, fileExtension);

                // If it already exists, it will be deleted!
                if (File.Exists(newFileName))
                {
                    File.Delete(newFileName);
                }

                // Rename the file.
                File.Move(fileName, newFileName);
            }

            return newFileName;
        }

        /// <summary>
        /// create a random guid for rsid values
        /// </summary>
        /// <returns></returns>
        public static string CreateNewRsid()
        {
            Guid g = Guid.NewGuid();
            return g.ToString();
        }

        /// <summary>
        /// add a new part that needs a relationship id
        /// </summary>
        /// <param name="document"></param>
        public static void AddNewPart(string document)
        {
            // Create a new word processing document.
            WordprocessingDocument wordDoc = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document);

            // Add the MainDocumentPart part in the new word processing document.
            var mainDocPart = wordDoc.AddNewPart<MainDocumentPart>("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "rId1");
            mainDocPart.Document = new Document();

            // Add the CustomFilePropertiesPart part in the new word processing document.
            var customFilePropPart = wordDoc.AddCustomFilePropertiesPart();
            customFilePropPart.Properties = new O.CustomProperties.Properties();

            // Add the CoreFilePropertiesPart part in the new word processing document.
            var coreFilePropPart = wordDoc.AddCoreFilePropertiesPart();
            using (XmlTextWriter writer = new XmlTextWriter(coreFilePropPart.GetStream(FileMode.Create), Encoding.UTF8))
            {
                writer.WriteRaw("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"></cp:coreProperties>");
                writer.Flush();
            }

            // Add the DigitalSignatureOriginPart part in the new word processing document.
            wordDoc.AddNewPart<DigitalSignatureOriginPart>("rId4");

            // Add the ExtendedFilePropertiesPart part in the new word processing document.
            var extendedFilePropPart = wordDoc.AddNewPart<ExtendedFilePropertiesPart>("rId5");
            extendedFilePropPart.Properties = new DocumentFormat.OpenXml.ExtendedProperties.Properties();

            // Add the ThumbnailPart part in the new word processing document.
            wordDoc.AddNewPart<ThumbnailPart>("image/jpeg", "rId6");

            wordDoc.Close();
        }

        // To add a new document part to a package.
        public static void AddNewPart(string document, string fileName)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                CustomXmlPart myXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);

                using (FileStream stream = new FileStream(fileName, FileMode.Open))
                {
                    myXmlPart.FeedData(stream);
                }
            }
        }

        public static void SearchAndReplace(string document, string find, string replace)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                Regex regexText = new Regex(find);
                docText = regexText.Replace(docText, replace);

                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }

        /// <summary>
        /// Return Word Embedded Object Count
        /// </summary>
        /// <param name="mPart"></param>
        /// <returns></returns>
        public static List<string> GetEmbeddedObjectProperties(string path, string fileType)
        {
            List<string> tList = new List<string>();

            if (fileType == Strings.oAppWord)
            {
                using (WordprocessingDocument myDoc = WordprocessingDocument.Open(path, false))
                {
                    foreach (EmbeddedObjectPart oep in myDoc.MainDocumentPart.EmbeddedObjectParts)
                    {
                        tList.Add(oep.Uri.ToString());
                    }

                    foreach (EmbeddedPackagePart epp in myDoc.MainDocumentPart.EmbeddedPackageParts)
                    {
                        tList.Add(epp.Uri.ToString());
                    }

                    foreach (EmbeddedControlPersistencePart ecp in myDoc.MainDocumentPart.EmbeddedControlPersistenceParts)
                    {
                        tList.Add(ecp.Uri.ToString());
                    }
                }
            }
            else if (fileType == Strings.oAppExcel)
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(path, false))
                {
                    foreach (WorksheetPart wp in doc.WorkbookPart.WorksheetParts)
                    {
                        foreach (EmbeddedObjectPart oep in wp.EmbeddedObjectParts)
                        {
                            tList.Add(oep.Uri.ToString());
                        }

                        foreach (EmbeddedPackagePart epp in wp.EmbeddedPackageParts)
                        {
                            tList.Add(epp.Uri.ToString());
                        }

                        foreach (EmbeddedControlPersistencePart ecp in wp.EmbeddedControlPersistenceParts)
                        {
                            tList.Add(ecp.Uri.ToString());
                        }
                    }
                }
            }
            else if (fileType == Strings.oAppPowerPoint)
            {
                using (PresentationDocument doc = PresentationDocument.Open(path, false))
                {
                    foreach (SlidePart sp in doc.PresentationPart.SlideParts)
                    {
                        foreach (EmbeddedObjectPart oep in sp.EmbeddedObjectParts)
                        {
                            tList.Add(oep.Uri.ToString());
                        }

                        foreach (EmbeddedPackagePart epp in sp.EmbeddedPackageParts)
                        {
                            tList.Add(epp.Uri.ToString());
                        }

                        foreach (EmbeddedControlPersistencePart ecp in sp.EmbeddedControlPersistenceParts)
                        {
                            tList.Add(ecp.Uri.ToString());
                        }
                    }
                }
            }

            return tList;
        }

        public static List<string> GetShapes(string path, string fileType)
        {
            List<string> tList = new List<string>();

            int count = 0;

            if (fileType == Strings.oAppWord)
            {
                // with Word, we can just run through the entire body and get the shapes
                using (WordprocessingDocument document = WordprocessingDocument.Open(path, false))
                {
                    foreach (ChartPart c in document.MainDocumentPart.ChartParts)
                    {
                        count++;
                        tList.Add(count + Strings.wPeriod + c.Uri + Strings.wArrow + Strings.wShpChart);
                    }

                    foreach (AO.Shape shape in document.MainDocumentPart.Document.Body.Descendants<AO.Shape>())
                    {
                        count++;
                        tList.Add(count + Strings.shpOfficeDrawing);
                    }

                    foreach (O.Vml.Shape shape in document.MainDocumentPart.Document.Body.Descendants<O.Vml.Shape>())
                    {
                        count++;
                        tList.Add(count + Strings.wPeriod + shape.Id + Strings.wArrow + Strings.shpVml);
                    }

                    foreach (O.Math.Shape shape in document.MainDocumentPart.Document.Body.Descendants<O.Math.Shape>())
                    {
                        count++;
                        tList.Add(count + Strings.shpMath);
                    }

                    foreach (A.Diagrams.Shape shape in document.MainDocumentPart.Document.Body.Descendants<A.Diagrams.Shape>())
                    {
                        count++;
                        tList.Add(count + Strings.shpDrawingDgm);
                    }

                    foreach (A.ChartDrawing.Shape shape in document.MainDocumentPart.Document.Body.Descendants<A.ChartDrawing.Shape>())
                    {
                        count++;
                        tList.Add(count + Strings.shpDrawingDgm);
                    }

                    foreach (A.Charts.Shape shape in document.MainDocumentPart.Document.Body.Descendants<A.Charts.Shape>())
                    {
                        count++;
                        tList.Add(count + Strings.shpChartShape);
                    }

                    foreach (A.Shape shape in document.MainDocumentPart.Document.Body.Descendants<A.Shape>())
                    {
                        count++;
                        tList.Add(count + Strings.shpShape);
                    }

                    foreach (A.Diagrams.Shape3D shape in document.MainDocumentPart.Document.Body.Descendants<A.Diagrams.Shape3D>())
                    {
                        count++;
                        tList.Add(count + Strings.shp3D);
                    }
                }
            }
            else if (fileType == Strings.oAppExcel)
            {
                // with XL, we would need to check all sheets
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(path, false))
                {
                    foreach (Sheet sheet in document.WorkbookPart.Workbook.Sheets)
                    {
                        foreach (A.Spreadsheet.Shape shape in sheet.Descendants<A.Spreadsheet.Shape>())
                        {
                            count++;
                            tList.Add(count + Strings.shpXlDraw);
                        }

                        foreach (AO.Shape shape in sheet.Descendants<AO.Shape>())
                        {
                            count++;
                            tList.Add(count + Strings.shpOfficeDrawing);
                        }

                        foreach (O.Vml.Shape shape in sheet.Descendants<O.Vml.Shape>())
                        {
                            count++;
                            tList.Add(count + Strings.wPeriod + shape.Id + Strings.wArrow + Strings.shpVml);
                        }

                        foreach (O.Math.Shape shape in sheet.Descendants<O.Math.Shape>())
                        {
                            count++;
                            tList.Add(count + Strings.shpMath);
                        }

                        foreach (A.Diagrams.Shape shape in sheet.Descendants<A.Diagrams.Shape>())
                        {
                            count++;
                            tList.Add(count + Strings.shpDrawingDgm);
                        }

                        foreach (A.ChartDrawing.Shape shape in sheet.Descendants<A.ChartDrawing.Shape>())
                        {
                            count++;
                            tList.Add(count + Strings.shpChartDraw);
                        }

                        foreach (A.Charts.Shape shape in sheet.Descendants<A.Charts.Shape>())
                        {
                            count++;
                            tList.Add(count + Strings.shpChartShape);
                        }

                        foreach (A.Shape shape in sheet.Descendants<A.Shape>())
                        {
                            count++;
                            tList.Add(count + Strings.shpShape);
                        }

                        foreach (A.Diagrams.Shape3D shape in sheet.Descendants<A.Diagrams.Shape3D>())
                        {
                            count++;
                            tList.Add(count + Strings.shp3D);
                        }
                    }
                }
            }
            else if (fileType == Strings.oAppPowerPoint)
            {
                // with PPT, we need to run through all slides
                using (PresentationDocument document = PresentationDocument.Open(path, false))
                {
                    foreach (SlidePart slidePart in document.PresentationPart.SlideParts)
                    {
                        foreach (P.Shape shape in slidePart.Slide.Descendants<P.Shape>())
                        {
                            count++;
                            foreach (OpenXmlElement child1 in shape.ChildElements)
                            {
                                if (child1.GetType().ToString() == Strings.dfopNVSP)
                                {
                                    foreach (OpenXmlElement child2 in child1.ChildElements)
                                    {
                                        if (child2.GetType().ToString() == Strings.dfopNVDP)
                                        {
                                            P.NonVisualDrawingProperties nvdp = (P.NonVisualDrawingProperties)child2;
                                            tList.Add(count + Strings.wPeriod + nvdp.Name);
                                        }
                                    }
                                }
                            }
                        }

                        foreach (AO.Shape shape in slidePart.Slide.Descendants<AO.Shape>())
                        {
                            count++;
                            tList.Add(count + Strings.shpOfficeDrawing);
                        }

                        foreach (O.Vml.Shape shape in slidePart.Slide.Descendants<O.Vml.Shape>())
                        {
                            count++;
                            tList.Add(count + Strings.wPeriod + shape.Id + Strings.wArrow + Strings.shpVml);
                        }

                        foreach (O.Math.Shape shape in slidePart.Slide.Descendants<O.Math.Shape>())
                        {
                            count++;
                            tList.Add(count + Strings.shpMath);
                        }

                        foreach (A.Diagrams.Shape shape in slidePart.Slide.Descendants<A.Diagrams.Shape>())
                        {
                            count++;
                            tList.Add(count + Strings.shpDrawingDgm);
                        }

                        foreach (A.ChartDrawing.Shape shape in slidePart.Slide.Descendants<A.ChartDrawing.Shape>())
                        {
                            count++;
                            tList.Add(count + Strings.shpChartDraw);
                        }

                        foreach (A.Charts.Shape shape in slidePart.Slide.Descendants<A.Charts.Shape>())
                        {
                            count++;
                            tList.Add(count + Strings.shpChartShape);
                        }

                        foreach (A.Shape shape in slidePart.Slide.Descendants<A.Shape>())
                        {
                            count++;
                            tList.Add(count + Strings.shpShape);
                        }

                        foreach (A.Diagrams.Shape3D shape in slidePart.Slide.Descendants<A.Diagrams.Shape3D>())
                        {
                            count++;
                            tList.Add(count + Strings.shp3D);
                        }
                    }
                }
            }

            return tList;
        }
    }
}