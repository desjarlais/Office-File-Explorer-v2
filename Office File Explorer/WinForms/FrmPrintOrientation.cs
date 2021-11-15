﻿using System;
using System.Windows.Forms;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmPrintOrientation : Form
    {
        static string fName;
        public FrmPrintOrientation(string filePath)
        {
            InitializeComponent();
            fName = filePath;
            rdoPortrait.Checked = true;
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            if (rdoLandscape.Checked)
            {
                Helpers.Word.SetPrintOrientation(fName, DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues.Landscape);
            }
            else
            {
                Helpers.Word.SetPrintOrientation(fName, DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues.Portrait);
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}