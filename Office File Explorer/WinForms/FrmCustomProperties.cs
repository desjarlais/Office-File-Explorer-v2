using Office_File_Explorer.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmCustomProperties : Form
    {
        static string fName, fType;
        List<string> bFiles;
        static bool isBatch;

        public FrmCustomProperties(string filePath, string fileType)
        {
            InitializeComponent();
            fName = filePath;
            fType = fileType;
            rdoNo.Enabled = false;
            rdoYes.Enabled = false;
            tbNumber.Enabled = false;
            tbText.Enabled = false;
            dtDateTime.Enabled = false;
            isBatch = false;
        }

        public FrmCustomProperties(List<string> files, string fileType)
        {
            InitializeComponent();
            fType = fileType;
            bFiles = files;
            rdoNo.Enabled = false;
            rdoYes.Enabled = false;
            tbNumber.Enabled = false;
            tbText.Enabled = false;
            dtDateTime.Enabled = false;
            isBatch = true;
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                bool value;
                int num;
                double number;

                switch (cbType.SelectedItem)
                {
                    case "YesNo":
                        if (rdoNo.Checked)
                        {
                            value = false;
                        }
                        else
                        {
                            value = true;
                        }

                        if (isBatch == true)
                        {
                            foreach (string f in bFiles)
                            {
                                Office.SetCustomProperty(f, tbName.Text, value, Office.PropertyTypes.YesNo, fType);
                                FileUtilities.WriteToLog(Strings.fLogFilePath, f + Strings.customPropSaved);
                            }
                        }
                        else
                        {
                            Office.SetCustomProperty(fName, tbName.Text, value, Office.PropertyTypes.YesNo, fType);
                        }

                        break;
                    case "Date":
                        if (isBatch == true)
                        {
                            foreach (string f in bFiles)
                            {
                                Office.SetCustomProperty(f, tbName.Text, dtDateTime.Value, Office.PropertyTypes.DateTime, fType);
                                FileUtilities.WriteToLog(Strings.fLogFilePath, f + Strings.customPropSaved);
                            }
                        }
                        else
                        {
                            Office.SetCustomProperty(fName, tbName.Text, dtDateTime.Value, Office.PropertyTypes.DateTime, fType);
                        }

                        break;
                    case "Number":
                        if (int.TryParse(tbNumber.Text, out num))
                        {
                            if (isBatch == true)
                            {
                                foreach (string f in bFiles)
                                {
                                    Office.SetCustomProperty(f, tbName.Text, num, Office.PropertyTypes.NumberInteger, fType);
                                    FileUtilities.WriteToLog(Strings.fLogFilePath, f + Strings.customPropSaved);
                                }
                            }
                            else
                            {
                                Office.SetCustomProperty(fName, tbName.Text, num, Office.PropertyTypes.NumberInteger, fType);
                            }

                        }
                        else if (double.TryParse(tbNumber.Text, out number))
                        {
                            if (isBatch == true)
                            {
                                foreach (string f in bFiles)
                                {
                                    Office.SetCustomProperty(f, tbName.Text, number, Office.PropertyTypes.NumberDouble, fType);
                                    FileUtilities.WriteToLog(Strings.fLogFilePath, f + Strings.customPropSaved);
                                }
                            }
                            else
                            {
                                Office.SetCustomProperty(fName, tbName.Text, number, Office.PropertyTypes.NumberDouble, fType);
                            }
                        }
                        else
                        {
                            // if the value isn't an int or double, just use text format
                            MessageBox.Show("The value entered is not a valid number and will be stored as text.", "Invalid Number", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                            if (isBatch == true)
                            {
                                foreach (string f in bFiles)
                                {
                                    Office.SetCustomProperty(f, tbName.Text, tbNumber.Text, Office.PropertyTypes.Text, fType);
                                    FileUtilities.WriteToLog(Strings.fLogFilePath, f + Strings.customPropSaved);
                                }
                            }
                            else
                            {
                                Office.SetCustomProperty(fName, tbName.Text, tbNumber.Text, Office.PropertyTypes.Text, fType);
                            }
                        }
                        break;
                    default:
                        // Text is default
                        if (isBatch == true)
                        {
                            foreach (string f in bFiles)
                            {
                                Office.SetCustomProperty(f, tbName.Text, tbText.Text, Office.PropertyTypes.Text, fType);
                                FileUtilities.WriteToLog(Strings.fLogFilePath, f + Strings.customPropSaved);
                            }
                        }
                        else
                        {
                            Office.SetCustomProperty(fName, tbName.Text, tbText.Text, Office.PropertyTypes.Text, fType);
                        }

                        break;
                }

                Close();
            }
            catch (InvalidDataException ide)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "SetCustomProperty: Invalid Property Value");
                FileUtilities.WriteToLog(Strings.fLogFilePath, ide.Message);
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnOKCustomProps Error: " + ex.Message);
                Close();
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void CbType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbType.SelectedItem.ToString() == "Text")
            {
                rdoNo.Enabled = false;
                rdoYes.Enabled = false;
                tbNumber.Enabled = false;
                tbText.Enabled = true;
                dtDateTime.Enabled = false;
            }
            else if (cbType.SelectedItem.ToString() == "YesNo")
            {
                rdoNo.Enabled = true;
                rdoYes.Enabled = true;
                tbNumber.Enabled = false;
                tbText.Enabled = false;
                dtDateTime.Enabled = false;
                rdoYes.Checked = true;
            }
            else if (cbType.SelectedItem.ToString() == "Number")
            {
                rdoNo.Enabled = false;
                rdoYes.Enabled = false;
                tbNumber.Enabled = true;
                tbText.Enabled = false;
                dtDateTime.Enabled = false;
            }
            else
            {
                rdoNo.Enabled = false;
                rdoYes.Enabled = false;
                tbNumber.Enabled = false;
                tbText.Enabled = false;
                dtDateTime.Enabled = true;
            }
        }

        private void FrmCustomProperties_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                Close();
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
