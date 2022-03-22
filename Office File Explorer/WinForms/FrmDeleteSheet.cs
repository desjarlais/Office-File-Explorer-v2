using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Office_File_Explorer.Helpers;
using System;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmDeleteSheet : Form
    {
        public string sheetName;
        public string filePath;

        public FrmDeleteSheet(string fPath)
        {
            InitializeComponent();
            sheetName = string.Empty;
            filePath = fPath;

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fPath, false))
            {
                foreach (Sheet s in spreadsheetDocument.WorkbookPart.Workbook.Sheets)
                {
                    cbxSheets.Items.Add(s.Name);
                }
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (Excel.DeleteSheet(filePath, cbxSheets.SelectedItem.ToString()))
            {
                sheetName = cbxSheets.SelectedItem.ToString();
            }

            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void FrmDeleteSheet_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                Close();
            }
        }
    }
}
