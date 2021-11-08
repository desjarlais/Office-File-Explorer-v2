using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace Office_File_Explorer.Helpers
{
    class Excel
    {
        public static bool fSuccess;

        public static bool RemoveComments(string path)
        {
            fSuccess = false;

            using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(path, true))
            {
                WorkbookPart wbPart = excelDoc.WorkbookPart;

                foreach (WorksheetPart wsp in wbPart.WorksheetParts)
                {
                    if (wsp.WorksheetCommentsPart != null)
                    {
                        if (wsp.WorksheetCommentsPart.Comments.Count() > 0)
                        {
                            WorksheetCommentsPart wcp = wsp.WorksheetCommentsPart;
                            foreach (Comment cmt in wcp.Comments.CommentList)
                            {
                                cmt.Remove();
                                fSuccess = true;
                            }
                        }
                    }
                }

                if (fSuccess)
                {
                    wbPart.Workbook.Save();
                }
            }

            return fSuccess;
        }

        public static bool RemoveLinks(string path)
        {
            fSuccess = false;

            using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(path, true))
            {
                WorkbookPart wbPart = excelDoc.WorkbookPart;

                if (wbPart.ExternalWorkbookParts.Count() == 0)
                {
                    return fSuccess;
                }

            DeleteLinkStart:
                foreach (ExternalWorkbookPart extWbPart in wbPart.ExternalWorkbookParts)
                {
                    foreach (ExternalRelationship er in extWbPart.ExternalRelationships)
                    {
                        extWbPart.DeleteExternalRelationship(er);

                        if (extWbPart.ExternalLink.Parent != null)
                        {
                            extWbPart.ExternalLink.Remove();
                        }

                        fSuccess = true;
                        goto DeleteLinkStart;
                    }
                }

                if (fSuccess)
                {
                    excelDoc.WorkbookPart.Workbook.Save();
                }
            }

            return fSuccess;
        }

        public static List<string> GetLinks(string path)
        {
            List<string> tList = new List<string>();

            using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(path, false))
            {
                WorkbookPart wbPart = excelDoc.WorkbookPart;
                int ExtRelCount = 0;

                foreach (ExternalWorkbookPart extWbPart in wbPart.ExternalWorkbookParts)
                {
                    ExtRelCount++;
                    ExternalRelationship extRel = extWbPart.ExternalRelationships.ElementAt(0);
                    tList.Add(ExtRelCount + Strings.wPeriod + extWbPart.ExternalRelationships.ElementAt(0).Uri);
                }
            }

            return tList;
        }

        public static List<string> GetComments(string path)
        {
            List<string> tList = new List<string>();

            using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(path, false))
            {
                WorkbookPart wbPart = excelDoc.WorkbookPart;
                int commentCount = 0;

                foreach (WorksheetPart wsp in wbPart.WorksheetParts)
                {
                    WorksheetCommentsPart wcp = wsp.WorksheetCommentsPart;
                    if (wcp != null)
                    {
                        foreach (Comment cmt in wcp.Comments.CommentList)
                        {
                            commentCount++;
                            CommentText cText = cmt.CommentText;
                            tList.Add(commentCount + Strings.wPeriod + cText.InnerText);
                        }
                    }
                }
            }

            return tList;
        }

        public static List<string> GetSheetInfo(string path)
        {
            List<string> tList = new List<string>();

            using (SpreadsheetDocument mySpreadsheet = SpreadsheetDocument.Open(path, false))
            {
                Sheets sheets = mySpreadsheet.WorkbookPart.Workbook.Sheets;

                // For each sheet, display the sheet information.
                foreach (Sheet sheet in sheets)
                {
                    foreach (OpenXmlAttribute attr in sheet.GetAttributes())
                    {
                        if (attr.LocalName == "name")
                        {
                            tList.Add(attr.LocalName + Strings.wColonBuffer + attr.Value);
                        }
                        else
                        {
                            tList.Add(Strings.wMinusSign + attr.LocalName + Strings.wColonBuffer + attr.Value);
                        }
                    }
                }
            }

            return tList;
        }

        public static List<string> GetHyperlinks(string path)
        {
            List<string> tList = new List<string>();

            using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(path, false))
            {
                int count = 0;

                foreach (WorksheetPart wsp in excelDoc.WorkbookPart.WorksheetParts)
                {
                    IEnumerable<Hyperlink> hLinks = wsp.Worksheet.Descendants<Hyperlink>();
                    foreach (Hyperlink h in hLinks)
                    {
                        count++;

                        string hRelUri = null;

                        // then check for hyperlinks relationships
                        if (wsp.HyperlinkRelationships.Count() > 0)
                        {
                            foreach (HyperlinkRelationship hRel in wsp.HyperlinkRelationships)
                            {
                                if (h.Id == hRel.Id)
                                {
                                    hRelUri = hRel.Uri.ToString();
                                    tList.Add(count + Strings.wPeriod + h.InnerText + " Uri = " + hRelUri);
                                }
                            }
                        }
                    }
                }
            }

            return tList;
        }

        public static List<string> GetFormulas(string path)
        {
            List<string> tList = new List<string>();



            return tList;
        }

        public static List<Worksheet> GetWorkSheets(string fileName, bool fileIsEditable)
        {
            List<Worksheet> returnVal = new List<Worksheet>();

            using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(fileName, fileIsEditable))
            {
                foreach (WorksheetPart wsPart in excelDoc.WorkbookPart.WorksheetParts)
                {
                    returnVal.Add(wsPart.Worksheet);
                }
            }

            return returnVal;
        }

        public static List<string> GetHyperlinks2(string path)
        {
            List<string> tList = new List<string>();

            int count = 0;

            foreach (Worksheet sht in GetWorkSheets(path, false))
            {
                foreach (var s in sht)
                {
                    if (s.LocalName == "sheetData")
                    {
                        IEnumerable<Cell> cells = sht.WorksheetPart.Worksheet.Descendants<Cell>();
                        foreach (Cell c in cells)
                        {
                            if (c.CellFormula != null)
                            {
                                count++;
                                tList.Add(count + Strings.wPeriod + c.CellReference + Strings.wEqualSign + c.CellFormula.Text);
                            }
                        }
                    }
                }
            }

            return tList;
        }

        public static List<string> GetSharedStrings(string path)
        {
            List<string> tList = new List<string>();

            int sharedStringCount = 0;

            using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(path, false))
            {
                WorkbookPart wbPart = excelDoc.WorkbookPart;
                if (wbPart.SharedStringTablePart != null)
                {
                    SharedStringTable sst = wbPart.SharedStringTablePart.SharedStringTable;
                    tList.Add("SharedString Count = " + sst.Count());
                    tList.Add("Unique Count = " + sst.UniqueCount);
                    tList.Add(string.Empty);

                    foreach (SharedStringItem ssi in sst)
                    {
                        sharedStringCount++;
                        Text ssValue = ssi.Text;
                        if (ssValue.Text != null)
                        {
                            tList.Add(sharedStringCount + Strings.wPeriod + ssValue.Text);
                        }
                    }
                }
            }

            return tList;
        }

        public static List<string> GetDefinedNames(string path)
        {
            List<string> tList = new List<string>();

            int nameCount = 0;

            using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(path, false))
            {
                WorkbookPart wbPart = excelDoc.WorkbookPart;

                // Retrieve a reference to the defined names collection.
                DefinedNames definedNames = wbPart.Workbook.DefinedNames;

                // If there are defined names, add them to the dictionary.
                if (definedNames != null)
                {
                    foreach (DefinedName dn in definedNames)
                    {
                        nameCount++;
                        tList.Add(nameCount + Strings.wPeriod + dn.Name.Value + " = " + dn.Text);
                    }
                }
            }

            return tList;
        }

        public static List<string> GetConnections(string path)
        {
            List<string> tList = new List<string>();

            using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(path, false))
            {
                WorkbookPart wbPart = excelDoc.WorkbookPart;
                ConnectionsPart cPart = wbPart.ConnectionsPart;

                if (cPart is null)
                {
                    return tList;
                }

                int cCount = 0;

                foreach (Connection c in cPart.Connections)
                {
                    cCount++;
                    if (c.DatabaseProperties.Connection != null)
                    {
                        string cn = c.DatabaseProperties.Connection;
                        string[] cArray = cn.Split(';');

                        tList.Add(cCount + ". Connection= " + c.Name);
                        foreach (var s in cArray)
                        {
                            tList.Add("    " + s);
                        }

                        if (c.ConnectionFile != null)
                        {
                            tList.Add(string.Empty);
                            tList.Add("    Connection File= " + c.ConnectionFile);

                            if (c.OlapProperties != null)
                            {
                                tList.Add("    Row Drill Count= " + c.OlapProperties.RowDrillCount);
                            }
                        }
                    }
                    else
                    {
                        tList.Add("Invalid connections.xml");
                    }
                }
            }

            return tList;
        }

        public static List<string> GetHiddenRowCols(string path)
        {
            List<string> tList = new List<string>();

            using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(path, false))
            {
                WorkbookPart wbPart = excelDoc.WorkbookPart;
                Sheets theSheets = wbPart.Workbook.Sheets;

                foreach (Sheet sheet in theSheets)
                {
                    Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().Where((s) => s.Name == sheet.Name).FirstOrDefault();

                    if (theSheet is null)
                    {
                        return tList;
                    }
                    else
                    {
                        tList.Add("Worksheet Name = " + sheet.Name);

                        // The sheet does exist.
                        WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
                        Worksheet ws = wsPart.Worksheet;
                        int rowCount = 0;
                        int colCount = 0;

                        tList.Add("##    ROWS    ##");
                        IEnumerable<Row> rows = ws.Descendants<Row>().Where((r) => r.Hidden != null && r.Hidden.Value);
                        foreach (Row row in rows)
                        {
                            rowCount++;
                            tList.Add(rowCount + Strings.wPeriod + row.InnerText);
                        }

                        if (rowCount == 0)
                        {
                            tList.Add("    None");
                        }

                        tList.Add("##    COLUMNS    ##");
                        IEnumerable<Column> cols = ws.Descendants<Column>().Where((c) => c.Hidden != null && c.Hidden.Value);
                        foreach (Column item in cols)
                        {
                            for (uint i = item.Min.Value; i <= item.Max.Value; i++)
                            {
                                colCount++;
                                tList.Add(colCount + ". Column " + i);
                            }
                        }

                        if (colCount == 0)
                        {
                            tList.Add("    None");
                        }
                    }
                    tList.Add(string.Empty);
                }
            }

            return tList;
        }

        // The DOM approach.
        // Note that the code below works only for cells that contain numeric values.
        // 
        public static List<string> ReadExcelFileDOM(string fileName)
        {
            List<string> values = new List<string>();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                string text;

                foreach (Row r in sheetData.Elements<Row>())
                {
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        if (c.CellValue != null)
                        {
                            text = c.CellValue.Text;
                            values.Add(text + Strings.wSpaceChar);
                        }
                    }
                }

                return values;
            }
        }

        // The SAX approach.
        public static List<string> ReadExcelFileSAX(string fileName)
        {
            List<string> values = new List<string>();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
                string text;

                while (reader.Read())
                {
                    if (reader.ElementType == typeof(CellValue))
                    {
                        text = reader.GetText();
                        values.Add(text + Strings.wSpaceChar);
                    }
                }

                return values;
            }
        }
    }
}
