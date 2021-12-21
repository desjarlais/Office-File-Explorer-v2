using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

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

                        string hRelUri = string.Empty;

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

        public bool DeleteSheet(string fileName, string sheetToDelete)
        {
            // Delete the specified sheet from within the specified workbook.
            // Return True if the sheet was found and deleted, False if it was not.
            // Note that this procedure might leave "orphaned" references, such as strings
            // in the shared strings table. You must take care when adding new strings, for example. 
            // The XLInsertStringIntoCell snippet handles this problem for you.

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true))
            {
                WorkbookPart wbPart = document.WorkbookPart;

                Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
                  Where(s => s.Name == sheetToDelete).FirstOrDefault();
                if (theSheet == null)
                {
                    // The specified sheet doesn't exist.
                    return false;
                }

                // Remove the sheet reference from the workbook.
                WorksheetPart worksheetPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
                theSheet.Remove();

                // Delete the worksheet part.
                wbPart.DeletePart(worksheetPart);

                // Save the workbook.
                wbPart.Workbook.Save();
            }
            return true;
        }

        // Given a document name, a worksheet name, and a cell name, get the column of the cell and return
        // the content of the first cell in that column.
        public string GetColumnHeader(string docName, string worksheetName, string cellName)
        {

            string returnValue = null;

            // Open the document as read-only.
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, false))
            {
                WorkbookPart wbPart = document.WorkbookPart;

                // Given a worksheet name, first find the Sheet that corresponds to the name.
                var sheet = wbPart.Workbook.Descendants<Sheet>().
                  Where(s => s.Name == worksheetName).FirstOrDefault();
                if (sheet == null)
                {
                    // The specified worksheet does not exist.
                    return null;
                }

                // Given the Sheet, 
                WorksheetPart worksheetPart = (WorksheetPart)(wbPart.GetPartById(sheet.Id));

                // Get the column name for the specified cell.
                string columnName = GetColumnName(cellName);

                // Get the cells in the specified column and order them by row.
                var headCell = worksheetPart.Worksheet.Descendants<Cell>().
                  Where(c => string.Compare(GetColumnName(c.CellReference.Value), columnName, true) == 0).
                  OrderBy(r => GetRowIndex(r.CellReference)).FirstOrDefault();

                if (headCell == null)
                {
                    // The specified column does not exist.
                    return null;
                }

                // If the content of the first cell is stored as a shared string, get the text of the first cell
                // from the SharedStringTablePart and return it. Otherwise, return the string value of the cell.
                if (headCell.DataType != null && headCell.DataType.Value == CellValues.SharedString)
                {
                    SharedStringTablePart sharedStringPart =
                      wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    if (sharedStringPart != null)
                    {
                        var items = sharedStringPart.SharedStringTable.Elements<SharedStringItem>();
                        returnValue = items.ElementAt(int.Parse(headCell.CellValue.Text)).InnerText;
                    }
                }
                else
                {
                    returnValue = headCell.CellValue.Text;
                }
            }
            return returnValue;
        }

        // Given a cell name, parses the specified cell to get the column name.
        private string GetColumnName(string cellName)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);
            return match.Value;
        }

        // Given a cell name, parses the specified cell to get the row index.
        private uint GetRowIndex(string cellName)
        {
            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex("\\d+");
            Match match = regex.Match(cellName);
            return uint.Parse(match.Value);
        }

        public CellFormat GetCellFormat(SpreadsheetDocument document, string sheetName, string addressName)
        {
            CellFormat theCellFormat = null;

            WorkbookPart wbPart = document.WorkbookPart;

            // Find the sheet with the supplied name, and then use that Sheet object
            // to retrieve a reference to the appropriate worksheet.
            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
              Where(s => s.Name == sheetName).FirstOrDefault();

            if (theSheet == null)
            {
                throw new ArgumentException("sheetName");
            }

            // Retrieve a reference to the worksheet part, and then use its Worksheet property to get 
            // a reference to the cell whose address matches the address you've supplied:
            WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
            Cell theCell = wsPart.Worksheet.Descendants<Cell>().
              Where(c => c.CellReference == addressName).FirstOrDefault();

            // It the cell doesn't exist, simply return a null reference:
            if (theCell != null)
            {
                // Go get the styles information.
                var styles = wbPart.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault();
                // If you can't retrieve the styles part, you're done.
                if (styles != null)
                {
                    var cf = System.Convert.ToInt32(theCell.StyleIndex.Value);
                    theCellFormat = (CellFormat)(styles.Stylesheet.CellFormats.Elements().ElementAt(cf));
                }
            }
            return theCellFormat;

        }

        public CellFormat GetCellFormat(string fileName, string sheetName, string addressName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                return GetCellFormat(document, sheetName, addressName);
            }
        }

        // Delete comments from a workbook, given an author name. 
        // Pass an empty string or null for the author name to delete all comments.
        public void DeleteCommentsByUser(string fileName, string userName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true))
            {
                WorkbookPart wbPart = document.WorkbookPart;

                var wsParts = wbPart.GetPartsOfType<WorksheetPart>();
                foreach (var ws in wsParts)
                {
                    var commentPart = ws.GetPartsOfType<WorksheetCommentsPart>().FirstOrDefault();
                    if (commentPart != null)
                    {
                        // The sheet has comments.

                        if (string.IsNullOrEmpty(userName))
                        {
                            // Delete the comments part.
                            ws.DeletePart(commentPart);
                        }
                        else
                        {
                            // Delete comments by the specific user.
                            var authors = commentPart.Comments.Authors;
                            Author author = null;
                            int authorID = 0;

                            // Get the index of the author, if the author exists:
                            int i = 0;
                            foreach (var Item in authors)
                            {
                                if (Item.InnerText == userName)
                                {
                                    author = (Author)Item;
                                    authorID = i;
                                    break;
                                }
                                else
                                {
                                    i += 1;
                                }
                            }

                            // If the supplied name had added comments, remove those comments:
                            if (author != null)
                            {
                                Comments theComments = commentPart.Comments;
                                var commentArray = theComments.CommentList.ToArray();


                                foreach (Comment comment in commentArray)
                                {
                                    if (comment.AuthorId.Value == authorID)
                                    {
                                        comment.Remove();
                                    }
                                }

                                if (theComments.CommentList.Count() > 0)
                                {
                                    // Still commments left in the list?

                                    // Remove the author from the author list.
                                    authors.RemoveChild(author);

                                    // Fix up author id values in the remaining comments.
                                    foreach (Comment comment in commentArray)
                                    {
                                        if (comment.AuthorId.Value > authorID)
                                        {
                                            comment.AuthorId.Value--;
                                        }
                                    }
                                    theComments.Save();
                                }
                                else
                                {
                                    // No more comments? Just delete the part.
                                    ws.DeletePart(commentPart);
                                }
                            }
                        }
                    }
                }
            }
        }

        // Given a reference to an Excel SpreadsheetDocument, the name of a sheet,
        // and a cell address, return a reference to the cell. Throw an ArgumentException
        // if the sheet doesn't exist, or if the cell doesn't yet exist.
        public Cell GetCellForReading(SpreadsheetDocument document, string sheetName, string address)
        {

            WorkbookPart wbPart = document.WorkbookPart;

            // Find the sheet with the supplied name, and then use that Sheet object
            // to retrieve a reference to the appropriate worksheet.
            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
              Where(s => s.Name == sheetName).FirstOrDefault();

            if (theSheet == null)
            {
                throw new ArgumentException("sheetName");
            }

            // Retrieve a reference to the worksheet part, and then use its Worksheet property to get 
            // a reference to the cell whose address matches the address you've supplied:
            WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
            Cell theCell = wsPart.Worksheet.Descendants<Cell>().
              Where(c => c.CellReference == address).FirstOrDefault();

            // If the cell doesn't exist, raise an exception to the caller:
            if (theCell == null)
            {
                throw new ArgumentException("address");
            }
            return theCell;
        }

        // Given a reference to an Excel SpreadsheetDocument, the name of a sheet,
        // and a cell address, return a reference to the cell. Throw an ArgumentException
        // if the sheet doesn't exist. If the cell doesn't exist, create it.
        public Cell GetCellForWriting(SpreadsheetDocument document, string sheetName, string address)
        {

            WorkbookPart wbPart = document.WorkbookPart;

            // Find the sheet with the supplied name, and then use that Sheet object
            // to retrieve a reference to the appropriate worksheet.
            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
              Where(s => s.Name == sheetName).FirstOrDefault();

            if (theSheet == null)
            {
                throw new ArgumentException("sheetName");
            }

            // Retrieve a reference to the worksheet part, and then use its Worksheet property to get 
            // a reference to the cell whose address matches the address you've supplied:
            WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
            Worksheet ws = wsPart.Worksheet;
            Cell theCell = ws.Descendants<Cell>().
              Where(c => c.CellReference == address).FirstOrDefault();

            // If the cell doesn't exist, create it:
            if (theCell == null)
            {
                theCell = InsertCellInWorksheet(ws, address);
            }
            return theCell;
        }
        // Given a Worksheet and an address (like "AZ254"), either return a cell reference, or 
        // create the cell reference and return it.
        private Cell InsertCellInWorksheet(Worksheet ws, string addressName)
        {

            // Use regular expressions to get the row number and column name.
            // If the parameter wasn't well formed, this code
            // will fail:
            Regex rx = new Regex("^(?<col>\\D+)(?<row>\\d+)");
            Match m = rx.Match(addressName);
            uint rowNumber = uint.Parse(m.Result("${row}"));
            string colName = m.Result("${col}");

            SheetData sheetData = ws.GetFirstChild<SheetData>();
            string cellReference = (colName + rowNumber.ToString());
            Cell theCell = null;

            // If the worksheet does not contain a row with the specified row index, insert one.
            var theRow = sheetData.Elements<Row>().
              Where(r => r.RowIndex.Value == rowNumber).FirstOrDefault();

            if (theRow == null)
            {
                theRow = new Row();
                theRow.RowIndex = rowNumber;
                sheetData.Append(theRow);
            }

            // If the cell you need already exists, return it.
            // If there is not a cell with the specified column name, insert one.  
            Cell refCell = theRow.Elements<Cell>().
              Where(c => c.CellReference.Value == cellReference).FirstOrDefault();
            if (refCell != null)
            {
                theCell = refCell;
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                foreach (Cell cell in theRow.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                theCell = new Cell();
                theCell.CellReference = cellReference;

                theRow.InsertBefore(theCell, refCell);
            }
            return theCell;
        }

        public bool InsertNumberIntoCell(string fileName, string sheetName, string addressName, int value)
        {

            // Given a file, a sheet, and a cell, insert a specified value.
            // For example: InsertNumberIntoCell("C:\Test.xlsx", "Sheet3", "C3", 14)

            // Assume failure.
            bool returnValue = false;

            // Open the document for editing.
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true))
            {
                WorkbookPart wbPart = document.WorkbookPart;

                Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
                  Where(s => s.Name == sheetName).FirstOrDefault();
                if (theSheet != null)
                {
                    Worksheet ws = ((WorksheetPart)(wbPart.GetPartById(theSheet.Id))).Worksheet;
                    Cell theCell = InsertCellInWorksheet(ws, addressName);

                    // Set the value of cell A1.
                    theCell.CellValue = new CellValue(value.ToString());
                    theCell.DataType = new EnumValue<CellValues>(CellValues.Number);

                    // Save the worksheet.
                    ws.Save();
                    returnValue = true;
                }
            }

            return returnValue;
        }

        public bool InsertStringIntoCell(string fileName, string sheetName, string addressName, string value)
        {
            // Given a file, a sheet, and a cell, insert a specified string.
            // For example: XLInsertStringIntoCell("C:\Test.xlsx", "Sheet3", "C3", "Microsoft");

            // If the string exists in the shared string table, get its index.
            // If the string doesn't exist in the shared string table, add it and get the next index.

            // Then, the remainder is the same as inserting a number, but insert the string index instead
            // of a value. Also, set the cell's t attribute to be the value "s".


            // Assume failure.
            bool returnValue = false;

            // Open the document for editing.
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true))
            {
                WorkbookPart wbPart = document.WorkbookPart;

                Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().Where((s) => s.Name == sheetName).FirstOrDefault();

                if (theSheet != null)
                {
                    Worksheet ws = ((WorksheetPart)(wbPart.GetPartById(theSheet.Id))).Worksheet;
                    Cell theCell = InsertCellInWorksheet(ws, addressName);

                    // Either retrieve the index of an existing string,
                    // or insert the string into the shared string table
                    // and get the index of the new item.
                    int stringIndex = InsertSharedStringItem(wbPart, value);

                    theCell.CellValue = new CellValue(stringIndex.ToString());
                    theCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                    // Save the worksheet.
                    ws.Save();
                    returnValue = true;
                }
            }

            return returnValue;
        }

        // Given the main workbook part, and a text value, insert the text into the shared
        // string table. Create the table if necessary. If the value already exists, return
        // its index. If it doesn't exist, insert it and return its new index.
        private int InsertSharedStringItem(WorkbookPart wbPart, string value)
        {
            // Insert a value into the shared string table, creating the table if necessary.
            // Insert the string if it's not already there.
            // Return the index of the string.

            int index = 0;
            bool found = false;
            var stringTablePart = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

            // If the shared string table is missing, something's wrong.
            // Just return the index that you found in the cell.
            // Otherwise, look up the correct text in the table.
            if (stringTablePart == null)
            {
                // Create it.
                stringTablePart = wbPart.AddNewPart<SharedStringTablePart>();
            }

            var stringTable = stringTablePart.SharedStringTable;
            if (stringTable == null)
            {
                stringTable = new SharedStringTable();
            }

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in stringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == value)
                {
                    found = true;
                    break;
                }
                index += 1;
            }

            if (!found)
            {
                stringTable.AppendChild(new SharedStringItem(new Text(value)));
                stringTable.Save();
            }

            return index;
        }
    }
}
