using Office_File_Explorer.Helpers;
using System;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmExcelCommands : Form
    {
        public AppUtilities.ExcelViewCmds xlCmds = new AppUtilities.ExcelViewCmds();
        public AppUtilities.OfficeViewCmds offCmds = new AppUtilities.OfficeViewCmds();

        public FrmExcelCommands()
        {
            InitializeComponent();
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            // check Excel features
            // check Word features
            if (ckbLinks.Checked)
            {
                xlCmds |= AppUtilities.ExcelViewCmds.Links;
            }
            else
            {
                xlCmds &= ~AppUtilities.ExcelViewCmds.Links;
            }

            if (ckbComments.Checked)
            {
                xlCmds |= AppUtilities.ExcelViewCmds.Comments;
            }
            else
            {
                xlCmds &= ~AppUtilities.ExcelViewCmds.Comments;
            }

            if (ckbWorksheetInfo.Checked)
            {
                xlCmds |= AppUtilities.ExcelViewCmds.WorksheetInfo;
            }
            else
            {
                xlCmds &= ~AppUtilities.ExcelViewCmds.WorksheetInfo;
            }

            if (ckbHiddenRowCol.Checked)
            {
                xlCmds |= AppUtilities.ExcelViewCmds.HiddenRowsCols;
            }
            else
            {
                xlCmds &= ~AppUtilities.ExcelViewCmds.HiddenRowsCols;
            }

            if (ckbSharedStrings.Checked)
            {
                xlCmds |= AppUtilities.ExcelViewCmds.SharedStrings;
            }
            else
            {
                xlCmds &= ~AppUtilities.ExcelViewCmds.SharedStrings;
            }

            if (ckbCellValues.Checked)
            {
                xlCmds |= AppUtilities.ExcelViewCmds.CellValues;
            }
            else
            {
                xlCmds &= ~AppUtilities.ExcelViewCmds.CellValues;
            }

            if (ckbDefinedNames.Checked)
            {
                xlCmds |= AppUtilities.ExcelViewCmds.DefinedNames;
            }
            else
            {
                xlCmds &= ~AppUtilities.ExcelViewCmds.DefinedNames;
            }

            if (ckbConnections.Checked)
            {
                xlCmds |= AppUtilities.ExcelViewCmds.Connections;
            }
            else
            {
                xlCmds &= ~AppUtilities.ExcelViewCmds.Connections;
            }

            if (ckbHyperlinks.Checked)
            {
                xlCmds |= AppUtilities.ExcelViewCmds.Hyperlinks;
            }
            else
            {
                xlCmds &= ~AppUtilities.ExcelViewCmds.Hyperlinks;
            }

            // check Office features
            if (ckbShapes.Checked)
            {
                offCmds |= AppUtilities.OfficeViewCmds.Shapes;
            }
            else
            {
                offCmds &= ~AppUtilities.OfficeViewCmds.Shapes;
            }

            if (ckbOleObjects.Checked)
            {
                offCmds |= AppUtilities.OfficeViewCmds.OleObjects;
            }
            else
            {
                offCmds &= ~AppUtilities.OfficeViewCmds.OleObjects;
            }

            if (ckbPackageParts.Checked)
            {
                offCmds |= AppUtilities.OfficeViewCmds.PackageParts;
            }
            else
            {
                offCmds &= ~AppUtilities.OfficeViewCmds.PackageParts;
            }

            Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void CkbSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            if (ckbSelectAll.Checked)
            {
                ckbShapes.Checked = true;
                ckbSharedStrings.Checked = true;
                ckbWorksheetInfo.Checked = true;
                ckbShapes.Checked = true;
                ckbPackageParts.Checked = true;
                ckbOleObjects.Checked = true;
                ckbHyperlinks.Checked = true;
                ckbLinks.Checked = true;
                ckbHiddenRowCol.Checked = true;
                ckbDefinedNames.Checked = true;
                ckbConnections.Checked = true;
                ckbComments.Checked = true;
                ckbCellValues.Checked = true;
            }
            else
            {
                ckbShapes.Checked = false;
                ckbSharedStrings.Checked = false;
                ckbWorksheetInfo.Checked = false;
                ckbShapes.Checked = false;
                ckbPackageParts.Checked = false;
                ckbOleObjects.Checked = false;
                ckbHyperlinks.Checked = false;
                ckbLinks.Checked = false;
                ckbHiddenRowCol.Checked = false;
                ckbDefinedNames.Checked = false;
                ckbConnections.Checked = false;
                ckbComments.Checked = false;
                ckbCellValues.Checked = false;
            }
        }
    }
}
