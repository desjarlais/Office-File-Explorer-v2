﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Office2013.Word;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Office_File_Explorer.Helpers;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmDuplicateAuthors : Form
    {
        public DialogResult dr;
        string fPath;

        public FrmDuplicateAuthors(Dictionary<string, string> authors, string path)
        {
            InitializeComponent();
            fPath = path;
            foreach (var auth in authors)
            {
                LstAuthors.Items.Add("Author: " + auth.Key + Strings.chPipe + " User Id: " + auth.Value);
            }
        }

        private void BtnRemoveDupes_Click(object sender, EventArgs e)
        {
            using (WordprocessingDocument document = WordprocessingDocument.Open(fPath, true))
            {
                WordprocessingPeoplePart peoplePart = document.MainDocumentPart.WordprocessingPeoplePart;
                foreach (Person person in peoplePart.People)
                {
                    string p = LstAuthors.SelectedItem.ToString();
                    if (p.Contains("Author: " + person.Author.Value))
                    {
                        person.Remove();
                        document.Save();
                        dr = DialogResult.OK;
                        Close();
                    }
                }
            }
        }
    }
}
