using DocumentFormat.OpenXml.Packaging;
using Office_File_Explorer.Helpers;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmViewImages : Form
    {
        string appName, fName;

        private void LstImages_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (appName == Strings.oAppWord)
            {
                using (WordprocessingDocument document = WordprocessingDocument.Open(fName, false))
                {
                    foreach (ImagePart ip in document.MainDocumentPart.ImageParts)
                    {
                        if (ip.Uri.ToString() == LstImages.SelectedItem.ToString())
                        {
                            if (DisplayImage(ip) == false)
                            {
                                return;
                            }
                        }
                    }
                }
            }
            else if (appName == Strings.oAppExcel)
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fName, false))
                {
                    foreach (WorksheetPart wp in document.WorkbookPart.WorksheetParts)
                    {
                        if (wp.DrawingsPart is null)
                        {
                            return;
                        }

                        if (wp.DrawingsPart.ImageParts is not null)
                        {
                            foreach (ImagePart ip in wp.DrawingsPart.ImageParts)
                            {
                                if (ip.Uri.ToString() == LstImages.SelectedItem.ToString())
                                {
                                    if (DisplayImage(ip) == false)
                                    {
                                        return;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else if (appName == Strings.oAppPowerPoint)
            {
                using (PresentationDocument document = PresentationDocument.Open(fName, false))
                {
                    foreach (SlidePart sp in document.PresentationPart.SlideParts)
                    {
                        foreach (ImagePart ip in sp.ImageParts)
                        {
                            if (ip.Uri.ToString() == LstImages.SelectedItem.ToString())
                            {
                                if (DisplayImage(ip) == false)
                                {
                                    return;
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                return;
            }
        }

        public bool DisplayImage(ImagePart ip)
        {
            bool fSuccess = false;

            try
            {
                Stream stream = ip.GetStream();

                if (ip.Uri.ToString().EndsWith(".svg") || ip.Uri.ToString().EndsWith(".emf") || ip.Uri.ToString().EndsWith(".wmf"))
                {
                    toolStripStatusInfo.Text = "Format not supported.";
                    pbImage.Image = pbImage.ErrorImage;
                    pbImage.SizeMode = PictureBoxSizeMode.CenterImage;
                }
                else
                {
                    Image imgStream = Image.FromStream(stream);
                    pbImage.Image = imgStream;
                    
                    if (imgStream.Height > pbImage.Size.Height || imgStream.Width > pbImage.Size.Width)
                    {
                        pbImage.SizeMode = PictureBoxSizeMode.Zoom;
                    }
                    else
                    {
                        pbImage.SizeMode = PictureBoxSizeMode.CenterImage;
                    }

                    fSuccess = true;
                }

                pbImage.Visible = true;
                toolStripStatusInfo.Text = "Ready";
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "ViewImages::UnableToDisplayImage : " + ex.Message);
                toolStripStatusInfo.Text = "Error - " + ex.Message;
            }

            return fSuccess;
        }

        private void FrmViewImages_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                Close();
            }
        }

        public FrmViewImages(string path, string fType)
        {
            InitializeComponent();

            appName = fType;
            fName = path;

            if (appName == Strings.oAppWord)
            {
                using (WordprocessingDocument document = WordprocessingDocument.Open(fName, false))
                {
                    foreach (ImagePart ip in document.MainDocumentPart.ImageParts)
                    {
                        LstImages.Items.Add(ip.Uri);
                    }
                }
            }
            else if (appName == Strings.oAppExcel)
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fName, false))
                {
                    foreach (WorksheetPart wp in document.WorkbookPart.WorksheetParts)
                    {
                        if (wp.DrawingsPart is null)
                        {
                            return;
                        }

                        foreach (ImagePart ip in wp.DrawingsPart.ImageParts)
                        {
                            LstImages.Items.Add(ip.Uri);
                        }
                    }
                }
            }
            else if (appName == Strings.oAppPowerPoint)
            {
                using (PresentationDocument document = PresentationDocument.Open(fName, false))
                {
                    foreach (SlidePart sp in document.PresentationPart.SlideParts)
                    {
                        foreach (ImagePart ip in sp.ImageParts)
                        {
                            LstImages.Items.Add(ip.Uri);
                        }
                    }
                }
            }

            if (LstImages.Items.Count > 0)
            {
                LstImages.SelectedIndex = 0;
            }
            else
            {
                MessageBox.Show("Document does not contain any images.", "Images", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Load += (s, e) => Close();
                return;
            }
        }
    }
}
