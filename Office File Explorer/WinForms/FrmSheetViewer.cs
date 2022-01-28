using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Office_File_Explorer.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmSheetViewer : Form
    {
        public List<string> ss = new List<string>();
        public string fPath;

        public FrmSheetViewer(string filePath)
        {
            InitializeComponent();

            fPath = filePath;
            ss = Excel.GetSharedStringsWithoutFormatting(fPath);

            // populate worksheet info
            PopulateComboBox(fPath);
            cboWorksheets.SelectedIndex = 0;
            PopulateGridView();
        }

        /// <summary>
        /// populate the combo box with sheet names
        /// </summary>
        /// <param name="fPath"></param>
        public void PopulateComboBox(string fPath)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fPath, false))
            {
                foreach (Sheet s in spreadsheetDocument.WorkbookPart.Workbook.Sheets)
                {
                    cboWorksheets.Items.Add(s.Name);
                }
            }
        }

        public void PopulateGridView()
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fPath, false))
                {
                    WorkbookPart wbPart = spreadsheetDocument.WorkbookPart;
                    WorksheetPart wsPart = null;
                    SheetData sheetData = null;

                    foreach (Sheet s in wbPart.Workbook.Sheets)
                    {
                        if (s.Name == cboWorksheets.SelectedItem.ToString())
                        {
                            wsPart = (WorksheetPart)(wbPart.GetPartById(s.Id));
                            sheetData = wsPart.Worksheet.Elements<SheetData>().First();
                        }
                    }

                    // add a row to the grid for each row in the sheet
                    AddRows(sheetData.Elements<Row>().Count());

                    // populate the actual text of each cell
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

                    // adjust the column width based on cell content
                    Columns cols = wsPart.Worksheet.Elements<Columns>().First();
                    int colCount = dataGridView1.Columns.Count;
                    for (int i = 0; i < colCount; i++)
                    {
                        DataGridViewColumn dgvc = dataGridView1.Columns[i];
                        dgvc.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    }
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, ex.Message);
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
                case "BA": columnName = 52; break;
                case "BB": columnName = 53; break;
                case "BC": columnName = 54; break;
                case "BD": columnName = 55; break;
                case "BE": columnName = 56; break;
                case "BF": columnName = 57; break;
                case "BG": columnName = 58; break;
                case "BH": columnName = 59; break;
                case "BI": columnName = 60; break;
                case "BJ": columnName = 61; break;
                case "BK": columnName = 62; break;
                case "BL": columnName = 63; break;
                case "BM": columnName = 64; break;
                case "BN": columnName = 65; break;
                case "BO": columnName = 66; break;
                case "BP": columnName = 67; break;
                case "BQ": columnName = 68; break;
                case "BR": columnName = 69; break;
                case "BS": columnName = 70; break;
                case "BT": columnName = 71; break;
                case "BU": columnName = 72; break;
                case "BV": columnName = 73; break;
                case "BW": columnName = 74; break;
                case "BX": columnName = 75; break;
                case "BY": columnName = 76; break;
                case "BZ": columnName = 77; break;
                case "CA": columnName = 78; break;
                case "CB": columnName = 79; break;
                case "CC": columnName = 80; break;
                case "CD": columnName = 81; break;
                case "CE": columnName = 82; break;
                case "CF": columnName = 83; break;
                case "CG": columnName = 84; break;
                case "CH": columnName = 85; break;
                case "CI": columnName = 86; break;
                case "CJ": columnName = 87; break;
                case "CK": columnName = 88; break;
                case "CL": columnName = 89; break;
                case "CM": columnName = 90; break;
                case "CN": columnName = 91; break;
                case "CO": columnName = 92; break;
                case "CP": columnName = 93; break;
                case "CQ": columnName = 94; break;
                case "CR": columnName = 95; break;
                case "CS": columnName = 96; break;
                case "CT": columnName = 97; break;
                case "CU": columnName = 98; break;
                case "CV": columnName = 99; break;
                case "CW": columnName = 100; break;
                case "CX": columnName = 101; break;
                case "CY": columnName = 102; break;
                case "CZ": columnName = 103; break;
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

        private void cboWorksheets_SelectedIndexChanged(object sender, EventArgs e)
        {
            ClearRows();
            PopulateGridView();
        }

        private void viewStylesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            CellFormat cf = Excel.GetCellFormat(fPath, cboWorksheets.SelectedItem.ToString(), 
                dataGridView1.CurrentCell.OwningColumn.HeaderText + dataGridView1.CurrentCell.RowIndex.ToString());

            if (cf.ApplyFont)
            {
                sb.Append("Font = " + Excel.GetStyleFont(fPath, Convert.ToInt32(cf.FontId.Value)).FontName.Val);
            }

            MessageBox.Show(sb.ToString());
        }
    }
}
