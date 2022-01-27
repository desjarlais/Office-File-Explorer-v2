using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Office_File_Explorer.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmSheetViewer : Form
    {
        public List<string> ss = new List<string>();
        public List<Worksheet> sheets = new List<Worksheet>();
        public string fPath;

        public FrmSheetViewer(string filePath)
        {
            InitializeComponent();
            fPath = filePath;
            ss = Excel.GetSharedStringsWithoutFormatting(fPath);
            sheets = Excel.GetWorkSheets(fPath, false);
            
            foreach (Worksheet sheet in sheets)
            {
                cboWorksheets.Items.Add(sheet);
            }

            cboWorksheets.SelectedIndex = 0;

            PopulateGridView();
        }

        public void PopulateGridView()
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fPath, false))
                {
                    // get the first workbook, add other workbooks later
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    // add a row to the grid for each row in the sheet
                    AddRows(sheetData.Elements<Row>().Count());

                    // now that we have the rows and columns, populate the actual text
                    foreach (Row r in sheetData.Elements<Row>())
                    {
                        foreach (Cell c in r.Elements<Cell>())
                        {
                            if (rdoCellValues.Checked)
                            {
                                if (c.CellValue is not null)
                                {
                                    if (c.DataType is not null && c.DataType.Value.ToString() == "SharedString")
                                    {
                                        string cellText = GetCellValueFromSharedString(c);
                                        dataGridView1.Rows[Convert.ToInt32(r.RowIndex.Value)].Cells[GetColumnNumber(c.CellReference)].Value = cellText;
                                    }
                                    else
                                    {
                                        dataGridView1.Rows[Convert.ToInt32(r.RowIndex.Value)].Cells[GetColumnNumber(c.CellReference)].Value = c.CellValue.Text;
                                    }
                                }
                                else
                                {
                                    dataGridView1.Rows[Convert.ToInt32(r.RowIndex.Value)].Cells[GetColumnNumber(c.CellReference)].Value = string.Empty;
                                }
                            }
                            else if (rdoFormulas.Checked)
                            {
                                if (c.CellFormula is not null)
                                {
                                    dataGridView1.Rows[Convert.ToInt32(r.RowIndex.Value)].Cells[GetColumnNumber(c.CellReference)].Value = c.CellFormula.Text;
                                }
                                else
                                {
                                    dataGridView1.Rows[Convert.ToInt32(r.RowIndex.Value)].Cells[GetColumnNumber(c.CellReference)].Value = string.Empty;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        public string GetCellValueFromSharedString(Cell c)
        {
            int index;
            if (int.TryParse(c.CellValue.Text, out index))
            {
                return ss[index];
            }
            else
            {
                return string.Empty;
            }
        }

        public void AddRows(int rowsToAdd)
        {
            int count = 0;
            do
            {
                dataGridView1.Rows.Add();
                count++;
            } while (count <= rowsToAdd);
        }

        public int GetColumnNumber(string cellRef)
        {
            // strip any number from the column name
            var output = Regex.Replace(cellRef, @"[\d-]", string.Empty);
            int columnName = 0;

            // convert the col name to an index number
            switch (output)
            {
                case "A": columnName = 0; break;
                case "B": columnName = 1; break;
                case "C": columnName = 2; break;
                case "D": columnName = 3; break;
                case "E": columnName = 4; break;
                case "F": columnName = 5; break;
                case "G": columnName = 6; break;
                case "H": columnName = 7; break;
                case "I": columnName = 8; break;
                case "J": columnName = 9; break;
                case "K": columnName = 10; break;
                case "L": columnName = 11; break;
                case "M": columnName = 12; break;
                case "N": columnName = 13; break;
                case "O": columnName = 14; break;
                case "P": columnName = 15; break;
                case "Q": columnName = 16; break;
                case "R": columnName = 17; break;
                case "S": columnName = 18; break;
                case "T": columnName = 19; break;
                case "U": columnName = 20; break;
                case "V": columnName = 21; break;
                case "W": columnName = 22; break;
                case "X": columnName = 23; break;
                case "Y": columnName = 24; break;
                case "Z": columnName = 25; break;
                case "AA": columnName = 26; break;
                case "AB": columnName = 27; break;
                case "AC": columnName = 28; break;
                case "AD": columnName = 29; break;
                case "AE": columnName = 30; break;
                case "AF": columnName = 31; break;
                case "AG": columnName = 32; break;
                case "AH": columnName = 33; break;
                case "AI": columnName = 34; break;
                case "AJ": columnName = 35; break;
                case "AK": columnName = 36; break;
                case "AL": columnName = 37; break;
                case "AM": columnName = 38; break;
                case "AN": columnName = 39; break;
                case "AO": columnName = 40; break;
                case "AP": columnName = 41; break;
                case "AQ": columnName = 42; break;
                case "AR": columnName = 43; break;
                case "AS": columnName = 44; break;
                case "AT": columnName = 45; break;
                case "AU": columnName = 46; break;
                case "AV": columnName = 47; break;
                case "AW": columnName = 48; break;
                case "AX": columnName = 49; break;
                case "AY": columnName = 50; break;
                case "AZ": columnName = 51; break;

            }

            return columnName;
        }

        public void ClearRows()
        {
            dataGridView1.Rows.Clear();
        }

        private void rdoCellValues_CheckedChanged(object sender, EventArgs e)
        {
            ClearRows();
            PopulateGridView();
        }

        private void rdoFormulas_CheckedChanged(object sender, EventArgs e)
        {
            ClearRows();
            PopulateGridView();
        }
    }
}
