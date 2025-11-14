// Open Xml SDK refs
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

// app refs
using Office_File_Explorer.Helpers;

//.NET refs
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

// namespace refs
using O = DocumentFormat.OpenXml;
using System.ComponentModel;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmBatch : Form
    {
        // form globals
        public List<string> files = new List<string>();
        public string fileType = string.Empty;
        public string fType = string.Empty;
        public bool nodeDeleted = false;
        public bool nodeChanged = false;
        public string fromChangeTemplate;
        public Package pkg;

        public FrmBatch(Package package)
        {
            InitializeComponent();
            DisableUI();
            pkg = package;
        }

        public string GetFileExtension()
        {
            if (rdoWord.Checked == true)
            {
                fileType = "*.docx";
                fType = Strings.oAppWord;
            }
            else if (rdoExcel.Checked == true)
            {
                fileType = "*.xlsx";
                fType = Strings.oAppExcel;
            }
            else if (rdoPowerPoint.Checked == true)
            {
                fileType = "*.pptx";
                fType = Strings.oAppPowerPoint;
            }

            return fileType;
        }

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string DefaultTemplate
        {
            set => fromChangeTemplate = value;
        }

        public void DisableUI()
        {
            // disable all buttons
            DisableButtons();

            // disable all radio buttons
            rdoExcel.Enabled = false;
            rdoPowerPoint.Enabled = false;
            rdoWord.Enabled = false;

            // disable checkbox
            ckbSubfolders.Enabled = false;

            lstOutput.Items.Clear();
        }

        /// <summary>
        /// Disable all buttons, needed for toggle of radio buttons
        /// </summary>
        public void DisableButtons()
        {
            // disable all buttons
            BtnFixNotesPage.Enabled = false;
            BtnChangeTheme.Enabled = false;
            BtnAddCustomProps.Enabled = false;
            BtnDeleteCustomProps.Enabled = false;
            BtnRemovePIIOnSave.Enabled = false;
            BtnRemovePII.Enabled = false;
            BtnFixBookmarks.Enabled = false;
            BtnFixRevisions.Enabled = false;
            BtnConvertStrict.Enabled = false;
            BtnFixTableProps.Enabled = false;
            BtnDeleteRequestStatus.Enabled = false;
            BtnSetOpenByDefault.Enabled = false;
            BtnChangeTemplate.Enabled = false;
            BtnFixHyperlinks.Enabled = false;
            BtnUpdateNamespaces.Enabled = false;
            BtnFixComments.Enabled = false;
            BtnRemoveCustomTitle.Enabled = false;
            BtnResetBulletMargins.Enabled = false;
            BtnCheckForDigSig.Enabled = false;
            BtnFixFooterSpacing.Enabled = false;
            BtnRemoveCustomFileProps.Enabled = false;
            BtnFixCorruptTcTags.Enabled = false;
            BtnRemoveCustomXml.Enabled = false;
            BtnFixDupeCustomXml.Enabled = false;
            BtnFixTabBehavior.Enabled = false;
            BtnFixCommentNotes.Enabled = false;
        }

        public void EnableUI()
        {
            // disable all buttons first
            DisableButtons();

            // enable buttons that work for each app
            BtnChangeTheme.Enabled = true;
            BtnAddCustomProps.Enabled = true;
            BtnRemovePII.Enabled = true;
            BtnDeleteCustomProps.Enabled = true;
            BtnDeleteRequestStatus.Enabled = true;
            BtnCheckForDigSig.Enabled = true;
            BtnRemoveCustomFileProps.Enabled = true;
            BtnRemoveCustomXml.Enabled = true;

            // enable the radio buttons
            rdoExcel.Enabled = true;
            rdoPowerPoint.Enabled = true;
            rdoWord.Enabled = true;

            // enable checkbox
            ckbSubfolders.Enabled = true;

            // now check which radio button is selected and light up appropriate buttons
            if (rdoWord.Checked == true)
            {
                BtnFixBookmarks.Enabled = true;
                BtnFixRevisions.Enabled = true;
                BtnFixTableProps.Enabled = true;
                BtnRemovePII.Enabled = true;
                BtnSetOpenByDefault.Enabled = true;
                BtnChangeTemplate.Enabled = true;
                BtnUpdateNamespaces.Enabled = true;
                BtnFixComments.Enabled = true;
                BtnRemoveCustomTitle.Enabled = true;
                BtnFixFooterSpacing.Enabled = true;
                BtnFixCorruptTcTags.Enabled = true;
                BtnFixDupeCustomXml.Enabled = true;
                btnFixContentControls.Enabled = true;
            }

            if (rdoPowerPoint.Checked == true)
            {
                BtnRemovePII.Enabled = true;
                BtnRemovePIIOnSave.Enabled = true;
                BtnFixNotesPage.Enabled = true;
                BtnResetBulletMargins.Enabled = true;
                BtnFixTabBehavior.Enabled = true;
            }

            if (rdoExcel.Checked == true)
            {
                BtnConvertStrict.Enabled = true;
                BtnFixHyperlinks.Enabled = true;
                BtnFixCommentNotes.Enabled = true;
            }
        }

        public void PopulateAndDisplayFiles()
        {
            try
            {
                lstOutput.Items.Clear();
                files.Clear();
                int fCount = 0;

                DirectoryInfo dir = new DirectoryInfo(tbFolderPath.Text);
                if (ckbSubfolders.Checked == true)
                {
                    foreach (FileInfo f in dir.GetFiles("*.*", SearchOption.AllDirectories))
                    {
                        if (f.Name.StartsWith("~"))
                        {
                            // we don't want to process temp files
                            continue;
                        }
                        else
                        {
                            if (GetFileExtension() == "*.docx")
                            {
                                if (f.Name.EndsWith(Strings.docxFileExt) || f.Name.EndsWith(Strings.docmFileExt) || f.Name.EndsWith(Strings.dotxFileExt) || f.Name.EndsWith(Strings.dotmFileExt))
                                {
                                    files.Add(f.FullName);
                                    lstOutput.Items.Add(f.FullName);
                                    fCount++;
                                }
                            }
                            else if (GetFileExtension() == "*.xlsx")
                            {
                                if (f.Name.EndsWith(Strings.xlsxFileExt) || f.Name.EndsWith(Strings.xlsmFileExt) || f.Name.EndsWith(Strings.xltxFileExt) || f.Name.EndsWith(Strings.xltmFileExt))
                                {
                                    files.Add(f.FullName);
                                    lstOutput.Items.Add(f.FullName);
                                    fCount++;
                                }
                            }
                            else if (GetFileExtension() == "*.pptx")
                            {
                                if (f.Name.EndsWith(Strings.pptxFileExt) || f.Name.EndsWith(Strings.pptmFileExt) || f.Name.EndsWith(Strings.potxFileExt) || f.Name.EndsWith(Strings.potmFileExt))
                                {
                                    files.Add(f.FullName);
                                    lstOutput.Items.Add(f.FullName);
                                    fCount++;
                                }
                            }
                        }
                    }
                }
                else
                {
                    foreach (FileInfo f in dir.GetFiles("*.*"))
                    {
                        if (f.Name.StartsWith("~"))
                        {
                            // we don't want to change temp files
                            continue;
                        }
                        else
                        {
                            // populate the list of file paths
                            if (GetFileExtension() == "*.docx")
                            {
                                if (f.Name.EndsWith(Strings.docxFileExt) || f.Name.EndsWith(Strings.docmFileExt) || f.Name.EndsWith(Strings.dotxFileExt) || f.Name.EndsWith(Strings.dotmFileExt))
                                {
                                    files.Add(f.FullName);
                                    lstOutput.Items.Add(f.FullName);
                                    fCount++;
                                }
                            }
                            else if (GetFileExtension() == "*.xlsx")
                            {
                                if (f.Name.EndsWith(Strings.xlsxFileExt) || f.Name.EndsWith(Strings.xlsmFileExt) || f.Name.EndsWith(Strings.xltxFileExt) || f.Name.EndsWith(Strings.xltmFileExt))
                                {
                                    files.Add(f.FullName);
                                    lstOutput.Items.Add(f.FullName);
                                    fCount++;
                                }
                            }
                            else if (GetFileExtension() == "*.pptx")
                            {
                                if (f.Name.EndsWith(Strings.pptxFileExt) || f.Name.EndsWith(Strings.pptmFileExt) || f.Name.EndsWith(Strings.potxFileExt) || f.Name.EndsWith(Strings.potmFileExt))
                                {
                                    files.Add(f.FullName);
                                    lstOutput.Items.Add(f.FullName);
                                    fCount++;
                                }
                            }
                        }
                    }
                }

                if (fCount == 0)
                {
                    lstOutput.Items.Add("** No Files **");
                    DisableButtons();
                }
            }
            catch (ArgumentException ae)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnPopulateAndDisplayFiles Error: " + ae.Message);
                lstOutput.Items.Add("** Invalid folder path **");
            }
            catch (DirectoryNotFoundException dnfe)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnPopulateAndDisplayFiles Error: " + dnfe.Message);
                lstOutput.Items.Add("** Invalid folder path **");
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "PopulateAndDisplayFiles Error: " + ex.Message);
            }
        }

        private void BtnBrowseFolder_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult result = folderBrowserDialog1.ShowDialog();
                if (result == DialogResult.OK)
                {
                    tbFolderPath.Text = folderBrowserDialog1.SelectedPath;
                    EnableUI();
                    PopulateAndDisplayFiles();
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnBrowseDirectory Error: " + ex.Message);
            }
        }

        private void BtnCopyOutput_Click(object sender, EventArgs e)
        {
            try
            {
                if (lstOutput.Items.Count <= 0)
                {
                    return;
                }

                StringBuilder buffer = new StringBuilder();
                foreach (object t in lstOutput.Items)
                {
                    buffer.Append(t);
                    buffer.Append('\n');
                }

                Clipboard.SetText(buffer.ToString());
            }
            catch (Exception ex)
            {
                lstOutput.Items.Add(Strings.wArrow + Strings.wErrorText + ex.Message);
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BatchProcessing Copy Error: " + ex.Message);
            }
        }

        private void BtnAddCustomProps_Click(object sender, EventArgs e)
        {
            FrmCustomProperties cFrm = new FrmCustomProperties(files, fType)
            {
                Owner = this
            };
            cFrm.ShowDialog();

            lstOutput.Items.Clear();
            lstOutput.Items.Add("** Batch Processing Finished **");
        }

        private void BtnDeleteCustomProps_Click(object sender, EventArgs e)
        {
            try
            {
                string propNameToDelete = string.Empty;
                lstOutput.Items.Clear();

                if (fType == Strings.oAppWord)
                {
                    using (var fm = new FrmBatchDeleteCustomProps())
                    {
                        fm.ShowDialog();
                        propNameToDelete = fm.PropName;
                    }

                    foreach (string f in files)
                    {
                        try
                        {
                            using (WordprocessingDocument document = WordprocessingDocument.Open(f, true))
                            {
                                if (propNameToDelete == Strings.wCancel)
                                {
                                    return;
                                }
                                else
                                {
                                    if (document.CustomFilePropertiesPart != null)
                                    {
                                        bool customPropFound = false;

                                        foreach (CustomDocumentProperty cdp in document.CustomFilePropertiesPart.RootElement)
                                        {
                                            if (propNameToDelete == cdp.Name)
                                            {
                                                cdp.Remove();
                                                lstOutput.Items.Add(f + Strings.wColonBuffer + propNameToDelete + " deleted");
                                                customPropFound = true;
                                            }
                                        }

                                        if (customPropFound == false)
                                        {
                                            lstOutput.Items.Add(f + Strings.noProp);
                                        }
                                    }
                                    else
                                    {
                                        lstOutput.Items.Add(f + Strings.noProp);
                                    }
                                }
                            }
                        }
                        catch (Exception innerEx)
                        {
                            lstOutput.Items.Add(f + Strings.wArrow + Strings.wErrorText + innerEx.Message);
                            FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnListCustomProps Error: " + f + Strings.wColonBuffer + innerEx.Message);
                        }
                    }
                }
                else if (fType == Strings.oAppExcel)
                {
                    using (var fm = new FrmBatchDeleteCustomProps())
                    {
                        fm.ShowDialog();
                        propNameToDelete = fm.PropName;
                    }

                    foreach (string f in files)
                    {
                        try
                        {
                            using (SpreadsheetDocument document = SpreadsheetDocument.Open(f, true))
                            {
                                if (propNameToDelete == Strings.wCancel)
                                {
                                    return;
                                }
                                else
                                {
                                    if (document.CustomFilePropertiesPart != null)
                                    {
                                        foreach (CustomDocumentProperty cdp in document.CustomFilePropertiesPart.RootElement)
                                        {
                                            if (propNameToDelete == cdp.Name)
                                            {
                                                cdp.Remove();
                                                lstOutput.Items.Add(f + Strings.wColonBuffer + propNameToDelete + " deleted");
                                            }
                                            else
                                            {
                                                lstOutput.Items.Add(f + Strings.noProp);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        lstOutput.Items.Add(f + Strings.noProp);
                                    }
                                }
                            }
                        }
                        catch (Exception innerEx)
                        {
                            lstOutput.Items.Add(f + Strings.wArrow + Strings.wErrorText + innerEx.Message);
                            FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnListCustomProps Error: " + f + Strings.wColonBuffer + innerEx.Message);
                        }
                    }
                }
                else if (fType == Strings.oAppPowerPoint)
                {
                    using (var fm = new FrmBatchDeleteCustomProps())
                    {
                        fm.ShowDialog();
                        propNameToDelete = fm.PropName;
                    }

                    foreach (string f in files)
                    {
                        try
                        {
                            using (PresentationDocument document = PresentationDocument.Open(f, true))
                            {
                                if (propNameToDelete == Strings.wCancel)
                                {
                                    return;
                                }
                                else
                                {
                                    if (document.CustomFilePropertiesPart != null)
                                    {
                                        foreach (CustomDocumentProperty cdp in document.CustomFilePropertiesPart.RootElement)
                                        {
                                            if (propNameToDelete == cdp.Name)
                                            {
                                                cdp.Remove();
                                                lstOutput.Items.Add(f + Strings.wColonBuffer + propNameToDelete + " deleted");
                                            }
                                            else
                                            {
                                                lstOutput.Items.Add(f + Strings.noProp);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        lstOutput.Items.Add(f + Strings.noProp);
                                    }
                                }
                            }
                        }
                        catch (Exception innerEx)
                        {
                            lstOutput.Items.Add(f + Strings.wArrow + Strings.wErrorText + innerEx.Message);
                            FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnListCustomProps Error: " + f + Strings.wColonBuffer + innerEx.Message);
                        }
                    }
                }
                else
                {
                    return;
                }
            }
            catch (IOException ioe)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnListCustomProps Error: " + ioe.Message);
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnListCustomProps Error: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        /// <summary>
        /// this is a very specific type of fix that looks for a unique-ish url that needs to be changed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnFixHyperlinks_Click(object sender, EventArgs e)
        {
            try
            {
                lstOutput.Items.Clear();
                Cursor = Cursors.WaitCursor;

                foreach (string f in files)
                {
                    try
                    {
                        bool isFileChanged = false;

                        using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(f, true))
                        {
                        // adding a goto since changing the relationship during enumeration causes an error
                        // after making the change, I restart the loops again to look for more corrupt links
                        HLinkStart:
                            foreach (WorksheetPart wsp in excelDoc.WorkbookPart.WorksheetParts)
                            {
                                IEnumerable<O.Spreadsheet.Hyperlink> hLinks = wsp.Worksheet.Descendants<O.Spreadsheet.Hyperlink>();
                                // loop each hyperlink to get the rid
                                foreach (O.Spreadsheet.Hyperlink h in hLinks)
                                {
                                    // then check for hyperlinks relationships for the rid
                                    if (wsp.HyperlinkRelationships.Count() > 0)
                                    {
                                        foreach (HyperlinkRelationship hRel in wsp.HyperlinkRelationships)
                                        {
                                            // if the rid's match, we have the same hyperlink
                                            if (h.Id == hRel.Id)
                                            {
                                                // there is a scenario where files from OpenText appear to be damaged and the url is some temp file path
                                                // not the url path it should be
                                                string badUrl = string.Empty;
                                                string[] separatingStrings = { "livelink" };

                                                // check if the uri contains any of the known bad paths
                                                if (hRel.Uri.ToString().StartsWith("../../../"))
                                                {
                                                    badUrl = hRel.Uri.ToString().Replace("../../../", Strings.wBackslash);
                                                }
                                                else if (hRel.Uri.ToString().Contains("/AppData/Local/Microsoft/Windows/livelink/llsapi.dll/open/"))
                                                {
                                                    string[] urlParts = hRel.Uri.ToString().Split(separatingStrings, StringSplitOptions.RemoveEmptyEntries);
                                                    badUrl = hRel.Uri.ToString().Replace(urlParts[0], Strings.wBackslash);
                                                }
                                                else if (hRel.Uri.ToString().Contains("/AppData/Roaming/OpenText/"))
                                                {
                                                    string[] urlParts = hRel.Uri.ToString().Split(separatingStrings, StringSplitOptions.RemoveEmptyEntries);
                                                    badUrl = hRel.Uri.ToString().Replace(urlParts[0], Strings.wBackslash);
                                                }

                                                // if a bad path was found, start the work to replace it with the correct path
                                                if (badUrl != string.Empty)
                                                {
                                                    // loop the sharedstrings to get the correct replace value
                                                    if (excelDoc.WorkbookPart.SharedStringTablePart != null)
                                                    {
                                                        SharedStringTable sst = excelDoc.WorkbookPart.SharedStringTablePart.SharedStringTable;
                                                        foreach (SharedStringItem ssi in sst)
                                                        {
                                                            if (ssi.Text != null)
                                                            {
                                                                if (ssi.InnerText.ToString().EndsWith(badUrl))
                                                                {
                                                                    // now delete the relationship
                                                                    wsp.DeleteReferenceRelationship(h.Id);

                                                                    // now add a new relationship with the right address
                                                                    wsp.AddHyperlinkRelationship(new Uri(ssi.InnerText, UriKind.Absolute), true, h.Id);
                                                                    isFileChanged = true;
                                                                    goto HLinkStart;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            if (isFileChanged == true)
                            {
                                excelDoc.WorkbookPart.Workbook.Save();
                                lstOutput.Items.Add(f + "** Hyperlinks Fixed **");
                            }
                            else
                            {
                                lstOutput.Items.Add(f + "** No Corrupt Hyperlinks Found **");
                            }
                        }
                    }
                    catch (Exception innerEx)
                    {
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnFixExcelHyperlinks Error: " + f + Strings.wColonBuffer + innerEx.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnFixExcelHyperlinks Error: " + ex.Message);
                lstOutput.Items.Add(Strings.wErrorText + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnChangeTemplate_Click(object sender, EventArgs e)
        {
            try
            {
                lstOutput.Items.Clear();
                Cursor = Cursors.WaitCursor;

                // get the new template path from the user
                FrmChangeDefaultTemplate ctFrm = new FrmChangeDefaultTemplate()
                {
                    Owner = this
                };
                ctFrm.ShowDialog();

                foreach (string f in files)
                {
                    try
                    {
                        bool isFileChanged = false;
                        string attachedTemplateId = "rId1";
                        string filePath = string.Empty;

                        using (WordprocessingDocument document = WordprocessingDocument.Open(f, true))
                        {
                            DocumentSettingsPart dsp = document.MainDocumentPart.DocumentSettingsPart;

                            // if the external rel exists, we need to pull the rid and old uri
                            // we will be deleting this part and re-adding with the new uri
                            if (dsp.ExternalRelationships.Any())
                            {
                                // just change the attached template
                                foreach (ExternalRelationship er in dsp.ExternalRelationships)
                                {
                                    if (er.RelationshipType != null && er.RelationshipType == Strings.DocumentTemplatePartType)
                                    {
                                        // keep track of the existing rId for the template
                                        filePath = er.Uri.ToString();
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                // if the part does not exist, this is a Normal.dotm situation
                                // path out to where it should be based on default install settings
                                filePath = Strings.fNormalTemplatePath;

                                if (!File.Exists(filePath))
                                {
                                    // Normal.dotm path is not correct?
                                    FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnChangeDefaultTemplate Error: Invalid Attached Template Path - " + filePath);
                                    throw new Exception();
                                }
                            }

                            if (fromChangeTemplate == filePath || fromChangeTemplate is null || fromChangeTemplate == Strings.wCancel)
                            {
                                // file path is the same or user closed without wanting changes, do nothing
                                return;
                            }
                            else
                            {
                                filePath = fromChangeTemplate;

                                // delete the old part if it exists
                                if (dsp.ExternalRelationships.Any())
                                {
                                    dsp.DeleteExternalRelationship(attachedTemplateId);
                                    isFileChanged = true;
                                }

                                // if the template is not "Normal", add the new rel back
                                if (fromChangeTemplate != "Normal")
                                {
                                    // add back the new path
                                    Uri newFilePath = new Uri(filePath);
                                    dsp.AddExternalRelationship(Strings.DocumentTemplatePartType, newFilePath, attachedTemplateId);
                                    isFileChanged = true;
                                }
                                else
                                {
                                    // if we are changing to Normal, delete the attachtemplate id ref
                                    foreach (OpenXmlElement oe in dsp.Settings)
                                    {
                                        if (oe.ToString() == "DocumentFormat.OpenXml.Wordprocessing.AttachedTemplate")
                                        {
                                            oe.Remove();
                                            isFileChanged = true;
                                        }
                                    }
                                }
                            }

                            if (isFileChanged)
                            {
                                lstOutput.Items.Add(f + "** Attached Template Changed **");
                                document.MainDocumentPart.Document.Save();
                            }
                            else
                            {
                                lstOutput.Items.Add(f + "** No Changes Made To Attached Template **");
                            }
                        }
                    }
                    catch (Exception innerEx)
                    {
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnChangeAttachedTemplate Error: " + innerEx.Message);
                        lstOutput.Items.Add(Strings.wErrorText + innerEx.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnChangeAttachedTemplate Error: " + ex.Message);
                lstOutput.Items.Add(Strings.wErrorText + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnSetOpenByDefault_Click(object sender, EventArgs e)
        {
            try
            {
                List<CustomXmlPart> cxpList;
                lstOutput.Items.Clear();

                foreach (string f in files)
                {
                    try
                    {
                        nodeChanged = false;

                        if (FileUtilities.IsZipArchiveFile(f) == false)
                        {
                            lstOutput.Items.Add(f + " : Not A Valid Office File");
                            return;
                        }

                        using (WordprocessingDocument document = WordprocessingDocument.Open(f, true))
                        {
                            cxpList = document.MainDocumentPart.CustomXmlParts.ToList();

                            foreach (CustomXmlPart cxp in cxpList)
                            {
                                XmlDocument xDoc = new XmlDocument();
                                xDoc.Load(cxp.GetStream());

                                if (xDoc.DocumentElement.NamespaceURI == Strings.schemaCustomXsn)
                                {
                                    foreach (XmlNode xNode in xDoc.ChildNodes)
                                    {
                                        if (xNode.Name == "customXsn")
                                        {
                                            foreach (XmlNode x in xNode)
                                            {
                                                if (x.Name == "openByDefault")
                                                {
                                                    x.FirstChild.Value = "False";
                                                    using (MemoryStream xmlMS = new MemoryStream())
                                                    {
                                                        xDoc.Save(xmlMS);
                                                        xmlMS.Position = 0;
                                                        cxp.FeedData(xmlMS);
                                                    }
                                                    nodeChanged = true;
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            if (nodeChanged == true)
                            {
                                document.MainDocumentPart.Document.Save();
                                lstOutput.Items.Add(f + " : openByDefault Changed");
                            }
                            else
                            {
                                lstOutput.Items.Add(f + " : openByDefault Not Found");
                            }
                        }
                    }
                    catch (Exception innerEx)
                    {
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnDeleteOpenByDefault Error: " + f + Strings.wColonBuffer + innerEx.Message);
                    }
                }
            }
            catch (IOException ioe)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnDeleteOpenByDefault Error: " + ioe.Message);
                lstOutput.Items.Add(Strings.wErrorText + ioe.Message);
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnDeleteOpenByDefault Error: " + ex.Message);
                lstOutput.Items.Add(Strings.wErrorText + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnUpdateNamespaces_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            lstOutput.Items.Clear();

            foreach (string f in files)
            {
                if (WordFixes.FixContentControlNamespaces(f))
                {
                    lstOutput.Items.Add(f + " : Quick Part Updated");
                }
                else
                {
                    lstOutput.Items.Add(f + " : No Update Needed");
                }
            }

            Cursor = Cursors.Default;
        }

        private void BtnConvertStrict_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                lstOutput.Items.Clear();
                foreach (string f in files)
                {
                    try
                    {
                        // check if the excelcnv.exe exists
                        string excelcnvPath;

                        if (File.Exists(Strings.sameBitnessO365))
                        {
                            excelcnvPath = Strings.sameBitnessO365;
                        }
                        else if (File.Exists(Strings.x86OfficeO365))
                        {
                            excelcnvPath = Strings.x86OfficeO365;
                        }
                        else if (File.Exists(Strings.sameBitnessMSI2016))
                        {
                            excelcnvPath = Strings.sameBitnessMSI2016;
                        }
                        else if (File.Exists(Strings.x86OfficeMSI2016))
                        {
                            excelcnvPath = Strings.x86OfficeMSI2016;
                        }
                        else if (File.Exists(Strings.sameBitnessMSI2013))
                        {
                            excelcnvPath = Strings.sameBitnessMSI2013;
                        }
                        else if (File.Exists(Strings.x86OfficeMSI2013))
                        {
                            excelcnvPath = Strings.x86OfficeMSI2013;
                        }
                        else
                        {
                            excelcnvPath = string.Empty;
                        }

                        // check if the file is strict
                        bool isStrict = false;

                        using (Package package = Package.Open(f, FileMode.Open, FileAccess.Read))
                        {
                            foreach (PackagePart part in package.GetParts())
                            {
                                if (part.Uri.ToString() == "/xl/workbook.xml")
                                {
                                    try
                                    {
                                        string docText = string.Empty;
                                        using (StreamReader sr = new StreamReader(part.GetStream()))
                                        {
                                            docText = sr.ReadToEnd();
                                            if (docText.Contains(@"conformance=""strict"""))
                                            {
                                                isStrict = true;
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        FileUtilities.WriteToLog(Strings.fLogFilePath, ex.Message);
                                    }
                                }
                            }
                        }

                        if (isStrict == true && excelcnvPath != string.Empty)
                        {
                            // setup destination file path
                            string strOriginalFile = f;
                            string strOutputPath = Path.GetDirectoryName(strOriginalFile) + "\\";
                            string strFileExtension = Path.GetExtension(strOriginalFile);
                            string strOutputFileName = strOutputPath + Path.GetFileNameWithoutExtension(strOriginalFile) + Strings.wFixedFileParentheses + strFileExtension;

                            // run the command to convert the file "excelcnv.exe -nme -oice "file-path" "converted-file-path""
                            string cParams = " -nme -oice " + '"' + f + '"' + " " + '"' + strOutputFileName + '"';
                            var proc = Process.Start(excelcnvPath, cParams);
                            proc.Close();
                            lstOutput.Items.Add(f + " : Converted Successfully");
                            lstOutput.Items.Add("   File Location: " + strOutputFileName);
                        }
                        else
                        {
                            lstOutput.Items.Add(f + " : Is Not Strict Open Xml Format");
                        }
                    }
                    catch (Exception innerEx)
                    {
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnConvertStrict: " + f + Strings.wColonBuffer + innerEx.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                lstOutput.Items.Add(Strings.wErrorText + ex.Message);
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnConvertStrict: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnRemovePIIOnSave_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                lstOutput.Items.Clear();
                foreach (string f in files)
                {
                    try
                    {
                        bool isFixed = false;
                        Cursor = Cursors.WaitCursor;
                        using (PresentationDocument document = PresentationDocument.Open(f, true))
                        {
                            document.PresentationPart.Presentation.RemovePersonalInfoOnSave = false;
                            document.PresentationPart.Presentation.Save();
                        }

                        if (isFixed)
                        {
                            lstOutput.Items.Add(f + ": PII Reset");
                        }
                        else
                        {
                            lstOutput.Items.Add(f + ": PII Not Reset");
                        }
                    }
                    catch (Exception innerEx)
                    {
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnPPTResetPII: " + f + Strings.wColonBuffer + innerEx.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                lstOutput.Items.Add(Strings.wErrorText + ex.Message);
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnPPTResetPII: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnFixNotesPage_Click(object sender, EventArgs e)
        {
            lstOutput.Items.Clear();
            Cursor = Cursors.WaitCursor;

            foreach (string f in files)
            {
                try
                {
                    using (PresentationDocument document = PresentationDocument.Open(f, true))
                    {
                        PowerPointFixes.ChangeNotesPageSize(document);
                        lstOutput.Items.Add(f + Strings.wArrow + Strings.pptNotesSizeReset);
                        FileUtilities.WriteToLog(Strings.fLogFilePath, f + Strings.wArrow + Strings.pptNotesSizeReset);
                    }
                }
                catch (NullReferenceException nre)
                {
                    lstOutput.Items.Add(f + Strings.wArrow + "** Document does not contain Notes Master **");
                    FileUtilities.WriteToLog(Strings.fLogFilePath, f + Strings.wArrow + Strings.wErrorText + nre.Message);
                }
                catch (Exception ex)
                {
                    lstOutput.Items.Add(f + Strings.wArrow + Strings.wErrorText + ex.Message);
                    FileUtilities.WriteToLog(Strings.fLogFilePath, f + Strings.wArrow + Strings.wErrorText + ex.Message);
                }
                finally
                {
                    Cursor = Cursors.Default;
                }
            }
        }

        private void BtnFixBookmarks_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                lstOutput.Items.Clear();
                foreach (string f in files)
                {
                    try
                    {
                        if (WordFixes.RemoveMissingBookmarkTags(f) == true)
                        {
                            lstOutput.Items.Add(f + " : Fixed Corrupt Bookmarks");
                        }
                        else if (WordFixes.RemovePlainTextCcFromBookmark(f) == true)
                        {
                            lstOutput.Items.Add(f + " : Fixed Corrupt Bookmarks");
                        }
                        else if (WordFixes.FixBookmarkTagInSdtContent(f) == true)
                        {
                            lstOutput.Items.Add(f + " : Fixed Corrupt Bookmarks");
                        }
                        else
                        {
                            lstOutput.Items.Add(f + " : No Corrupt Bookmarks Found");
                        }
                    }
                    catch (Exception innerEx)
                    {
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnFixCorruptBookmarks: " + f + Strings.wColonBuffer + innerEx.Message);
                    }

                }
            }
            catch (Exception ex)
            {
                lstOutput.Items.Add(Strings.wErrorText + ex.Message);
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnFixCorruptBookmarks: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnFixRevisions_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                lstOutput.Items.Clear();
                foreach (string f in files)
                {
                    try
                    {
                        bool isFixed = false;

                        if (FileUtilities.IsZipArchiveFile(f) == false)
                        {
                            lstOutput.Items.Add(f + " : Not A Valid Office File");
                            return;
                        }

                        using (WordprocessingDocument document = WordprocessingDocument.Open(f, true))
                        {
                            if (Word.IsPartNull(document, "DeletedRun") == false)
                            {
                                var deleted = document.MainDocumentPart.Document.Descendants<DeletedRun>().ToList();

                                // loop each DeletedRun
                                foreach (DeletedRun dr in deleted)
                                {
                                    foreach (OpenXmlElement oxedr in dr)
                                    {
                                        // if we have a Run, we need to look for Text tags
                                        if (oxedr.GetType().ToString() == Strings.dfowRun)
                                        {
                                            O.Wordprocessing.Run r = (O.Wordprocessing.Run)oxedr;
                                            foreach (OpenXmlElement oxe in oxedr.ChildElements)
                                            {
                                                // you can't have a Text tag inside a DeletedRun
                                                if (oxe.GetType().ToString() == Strings.dfowText)
                                                {
                                                    // create a DeletedText object so we can replace it with the Text tag
                                                    DeletedText dt = new DeletedText();

                                                    // check for attributes
                                                    if (oxe.HasAttributes)
                                                    {
                                                        if (oxe.GetAttributes().Count > 0)
                                                        {
                                                            dt.SetAttributes(oxe.GetAttributes());
                                                        }
                                                    }

                                                    // set the text value
                                                    dt.Text = oxe.InnerText;

                                                    // replace the Text with new DeletedText
                                                    r.ReplaceChild(dt, oxe);
                                                    isFixed = true;
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            // now save the file if we have changes
                            if (isFixed == true)
                            {
                                document.MainDocumentPart.Document.Save();
                                lstOutput.Items.Add(f + ": Fixed Corrupt Revisions");
                            }
                            else
                            {
                                lstOutput.Items.Add(f + ": No Corrupt Revisions Found");
                            }
                        }
                    }
                    catch (Exception innerEx)
                    {
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnFixCorruptRevisions: " + f + Strings.wColonBuffer + innerEx.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                lstOutput.Items.Add(Strings.wErrorText + ex.Message);
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnFixCorruptRevisions: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnFixTableProps_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                lstOutput.Items.Clear();
                foreach (string f in files)
                {
                    try
                    {
                        if (FileUtilities.IsZipArchiveFile(f) == false)
                        {
                            lstOutput.Items.Add(f + " : Not A Valid Office File");
                            return;
                        }

                        using (WordprocessingDocument document = WordprocessingDocument.Open(f, true))
                        {
                            // "global" document variables
                            bool tblModified = false;
                            OpenXmlElement tgClone = null;

                            // get the list of tables in the document
                            if (Word.IsPartNull(document, "Table") == false)
                            {
                                List<O.Wordprocessing.Table> tbls = document.MainDocumentPart.Document.Descendants<O.Wordprocessing.Table>().ToList();

                                foreach (O.Wordprocessing.Table tbl in tbls)
                                {
                                    // you can have only one tblGrid per table, including nested tables
                                    // it needs to be before any row elements so sequence is
                                    // 1. check if the tblGrid element is before any trow
                                    // 2. check for multiple tblGrid elements
                                    bool tRowFound = false;
                                    bool tGridBeforeRowFound = false;
                                    int tGridCount = 0;

                                    foreach (OpenXmlElement oxe in tbl.Elements())
                                    {
                                        // flag if we found a table row, once we find 1, the rest do not matter
                                        if (oxe.GetType().Name == "TableRow")
                                        {
                                            tRowFound = true;
                                        }

                                        // when we get to a tablegrid, we have a few things to check
                                        // 1. have we found a table row
                                        // 2. only one table grid can exist in the table, if there are multiple, delete the extras
                                        if (oxe.GetType().Name == "TableGrid")
                                        {
                                            // increment the tg counter
                                            tGridCount++;

                                            // if we have a table row and no table grid has been found yet, we need to save out this table grid
                                            // then move it in front of the table row later
                                            if (tRowFound == true && tGridCount == 1)
                                            {
                                                tGridBeforeRowFound = true;
                                                tgClone = oxe.CloneNode(true);
                                                oxe.Remove();
                                            }

                                            // if we have multiple table grids, delete the extras
                                            if (tGridCount > 1)
                                            {
                                                oxe.Remove();
                                                tblModified = true;
                                            }
                                        }
                                    }

                                    // if we had a table grid before a row, move it before the first row
                                    if (tGridBeforeRowFound == true)
                                    {
                                        tbl.InsertBefore(tgClone, tbl.GetFirstChild<TableRow>());
                                        tblModified = true;
                                    }
                                }
                            }

                            if (tblModified == true)
                            {
                                document.MainDocumentPart.Document.Save();
                                lstOutput.Items.Add(f + " : Table Fix Completed");
                            }
                            else
                            {
                                lstOutput.Items.Add(f + " : No Corrupt Table Found");
                            }
                        }
                    }
                    catch (Exception innerEx)
                    {
                        lstOutput.Items.Add(f + Strings.wArrow + Strings.wErrorText + innerEx.Message);
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnFixTableProps: " + f + Strings.wColonBuffer + innerEx.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                lstOutput.Items.Add(Strings.wErrorText + ex.Message);
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnFixTableProps: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnChangeTheme_Click(object sender, EventArgs e)
        {
            lstOutput.Items.Clear();

            OpenFileDialog fDialog = new OpenFileDialog
            {
                Title = "Select Office Theme File.",
                Filter = "Open XML Theme File | *.xml",
                RestoreDirectory = true,
                InitialDirectory = @"%userprofile%"
            };

            if (fDialog.ShowDialog() == DialogResult.OK)
            {
                string sThemeFilePath = fDialog.FileName.ToString();

                foreach (string f in files)
                {
                    try
                    {
                        // call the replace function using the theme file provided
                        Office.ReplaceTheme(f, sThemeFilePath, fType);
                        FileUtilities.WriteToLog(Strings.fLogFilePath, f + "--> Theme Replaced.");
                        lstOutput.Items.Add(f + "--> Theme Replaced.");
                    }
                    catch (Exception ex)
                    {
                        FileUtilities.WriteToLog(Strings.fLogFilePath, f + " --> Failed to replace theme : Error = " + ex.Message);
                        lstOutput.Items.Add(f + " --> Failed to replace theme : Error = " + ex.Message);
                    }
                }
            }
            else
            {
                return;
            }
        }

        private void BtnDeleteRequestStatus_Click(object sender, EventArgs e)
        {
            try
            {
                List<CustomXmlPart> cxpList;
                lstOutput.Items.Clear();

                if (fType == Strings.oAppWord)
                {
                    foreach (string f in files)
                    {
                        try
                        {
                            nodeDeleted = false;

                            if (FileUtilities.IsZipArchiveFile(f) == false)
                            {
                                lstOutput.Items.Add(f + " : Not A Valid Office File");
                                return;
                            }

                            using (WordprocessingDocument document = WordprocessingDocument.Open(f, true))
                            {
                                cxpList = document.MainDocumentPart.CustomXmlParts.ToList();

                                foreach (CustomXmlPart cxp in cxpList)
                                {
                                    XmlDocument xDoc = new XmlDocument();
                                    xDoc.Load(cxp.GetStream());

                                    XPathNavigator navigator = xDoc.CreateNavigator();

                                    // we only check the metadata custom xml file for requeststatus xml
                                    if (xDoc.DocumentElement.NamespaceURI == Strings.schemaMetadataProperties)
                                    {
                                        // move to the node and delete it
                                        navigator.MoveToChild("properties", Strings.schemaMetadataProperties);
                                        navigator.MoveToChild("documentManagement", string.Empty);
                                        navigator.MoveToChild(Strings.wCustomXmlRequestStatus, Strings.wRequestStatusNS);

                                        // check if we actually moved to the RequestStatus node
                                        // if we didn't move there, no changes should happen, it doesn't exist
                                        if (navigator.Name == Strings.wCustomXmlRequestStatus)
                                        {
                                            // delete the node
                                            navigator.DeleteSelf();

                                            // re-write the part
                                            using (MemoryStream xmlMS = new MemoryStream())
                                            {
                                                xDoc.Save(xmlMS);
                                                xmlMS.Position = 0;
                                                cxp.FeedData(xmlMS);
                                            }

                                            // flag the part so we can save the file
                                            nodeDeleted = true;
                                        }
                                    }
                                }

                                if (nodeDeleted == true)
                                {
                                    document.MainDocumentPart.Document.Save();
                                    lstOutput.Items.Add(f + " : Request Status Removed");
                                }
                                else
                                {
                                    lstOutput.Items.Add(f + " : Request Status Not Found");
                                }
                            }
                        }
                        catch (Exception innerEx)
                        {
                            lstOutput.Items.Add(f + Strings.wArrow + Strings.wErrorText + innerEx.Message);
                            FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnDeleteRequestStatus Error: " + f + Strings.wColonBuffer + innerEx.Message);
                        }
                    }
                }
                else if (fType == Strings.oAppExcel)
                {
                    foreach (string f in files)
                    {
                        try
                        {
                            nodeDeleted = false;
                            using (SpreadsheetDocument document = SpreadsheetDocument.Open(f, true))
                            {
                                cxpList = document.WorkbookPart.CustomXmlParts.ToList();

                                foreach (CustomXmlPart cxp in cxpList)
                                {
                                    XmlDocument xDoc = new XmlDocument();
                                    xDoc.Load(cxp.GetStream());

                                    XPathNavigator navigator = xDoc.CreateNavigator();

                                    if (xDoc.DocumentElement.NamespaceURI == Strings.schemaMetadataProperties)
                                    {
                                        navigator.MoveToChild("properties", Strings.schemaMetadataProperties);
                                        navigator.MoveToChild("documentManagement", string.Empty);
                                        navigator.MoveToChild(Strings.wCustomXmlRequestStatus, Strings.wRequestStatusNS);

                                        if (navigator.Name == Strings.wCustomXmlRequestStatus)
                                        {
                                            navigator.DeleteSelf();

                                            using (MemoryStream xmlMS = new MemoryStream())
                                            {
                                                xDoc.Save(xmlMS);
                                                xmlMS.Position = 0;
                                                cxp.FeedData(xmlMS);
                                            }

                                            nodeDeleted = true;
                                        }
                                    }
                                }

                                if (nodeDeleted == true)
                                {
                                    document.WorkbookPart.Workbook.Save();
                                    lstOutput.Items.Add(f + " : Request Status Removed");
                                }
                                else
                                {
                                    lstOutput.Items.Add(f + " : Request Status Not Found");
                                }
                            }
                        }
                        catch (Exception innerEx)
                        {
                            lstOutput.Items.Add(f + Strings.wArrow + Strings.wErrorText + innerEx.Message);
                            FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnDeleteRequestStatus Error: " + f + Strings.wColonBuffer + innerEx.Message);
                        }
                    }
                }
                else if (fType == Strings.oAppPowerPoint)
                {
                    foreach (string f in files)
                    {
                        try
                        {
                            nodeDeleted = false;
                            using (PresentationDocument document = PresentationDocument.Open(f, true))
                            {
                                cxpList = document.PresentationPart.CustomXmlParts.ToList();

                                foreach (CustomXmlPart cxp in cxpList)
                                {
                                    XmlDocument xDoc = new XmlDocument();
                                    xDoc.Load(cxp.GetStream());

                                    XPathNavigator navigator = xDoc.CreateNavigator();

                                    if (xDoc.DocumentElement.NamespaceURI == Strings.schemaMetadataProperties)
                                    {
                                        navigator.MoveToChild("properties", Strings.schemaMetadataProperties);
                                        navigator.MoveToChild("documentManagement", string.Empty);
                                        navigator.MoveToChild(Strings.wCustomXmlRequestStatus, Strings.wRequestStatusNS);

                                        if (navigator.Name == Strings.wCustomXmlRequestStatus)
                                        {
                                            navigator.DeleteSelf();

                                            using (MemoryStream xmlMS = new MemoryStream())
                                            {
                                                xDoc.Save(xmlMS);
                                                xmlMS.Position = 0;
                                                cxp.FeedData(xmlMS);
                                            }

                                            nodeDeleted = true;
                                        }
                                    }
                                }

                                if (nodeDeleted == true)
                                {
                                    document.PresentationPart.Presentation.Save();
                                    lstOutput.Items.Add(f + " : Request Status Removed");
                                }
                                else
                                {
                                    lstOutput.Items.Add(f + " : Request Status Not Found");
                                }
                            }
                        }
                        catch (Exception innerEx)
                        {
                            FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnDeleteRequestStatus Error: " + f + Strings.wColonBuffer + innerEx.Message);
                        }
                    }
                }
                else
                {
                    return;
                }
            }
            catch (IOException ioe)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnDeleteRequestStatus Error: " + ioe.Message);
                lstOutput.Items.Add(Strings.wErrorText + ioe.Message);
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnDeleteRequestStatus Error: " + ex.Message);
                lstOutput.Items.Add(Strings.wErrorText + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnFixComments_Click(object sender, EventArgs e)
        {
            try
            {
                List<string> nsList = new List<string>();
                List<string> nList = new List<string>();
                lstOutput.Items.Clear();

                foreach (string f in files)
                {
                    try
                    {
                        using (WordprocessingDocument document = WordprocessingDocument.Open(f, true))
                        {
                            WordprocessingCommentsPart commentsPart = document.MainDocumentPart.WordprocessingCommentsPart;
                            IEnumerable<OpenXmlUnknownElement> unknownList = document.MainDocumentPart.Document.Descendants<OpenXmlUnknownElement>();
                            IEnumerable<CommentReference> commentRefs = document.MainDocumentPart.Document.Descendants<CommentReference>();

                            bool saveFile = false;
                            bool cRefIdExists = false;

                            // todo: check the logic here
                            if (commentsPart is null && commentRefs.Count() > 0)
                            {
                                // if there are comment refs but no comments.xml, remove refs
                                foreach (CommentReference cr in commentRefs)
                                {
                                    cr.Remove();
                                    saveFile = true;
                                }
                            }
                            else if (commentsPart is null && !commentRefs.Any())
                            {
                                // for some reason these dangling refs are considered unknown types, not commentrefs
                                // convert to an openxmlelement then type it to a commentref to get the id
                                foreach (OpenXmlUnknownElement uk in unknownList)
                                {
                                    if (uk.LocalName == "commentReference")
                                    {
                                        // so far I only see the id in the outerxml
                                        XmlDocument xDoc = new XmlDocument();
                                        xDoc.LoadXml(uk.OuterXml);

                                        // traverse the outerxml until we get to the id
                                        if (xDoc.ChildNodes.Count > 0)
                                        {
                                            foreach (XmlNode xNode in xDoc.ChildNodes)
                                            {
                                                if (xNode.Attributes.Count > 0)
                                                {
                                                    foreach (XmlAttribute xa in xNode.Attributes)
                                                    {
                                                        if (xa.LocalName == "id")
                                                        {
                                                            // now that we have the id number, we can use it to compare with the comment part
                                                            // if the id exists in commentref but not the commentpart, it can be deleted
                                                            foreach (O.Wordprocessing.Comment cm in commentsPart.Comments)
                                                            {
                                                                int cId = Convert.ToInt32(cm.Id);
                                                                int cRefId = Convert.ToInt32(xa.Value);

                                                                if (cId == cRefId)
                                                                {
                                                                    cRefIdExists = true;
                                                                }
                                                            }

                                                            if (cRefIdExists == false)
                                                            {
                                                                uk.Remove();
                                                                saveFile = true;
                                                            }
                                                            else
                                                            {
                                                                cRefIdExists = false;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }


                            if (saveFile)
                            {
                                document.MainDocumentPart.Document.Save();
                                lstOutput.Items.Add(f + ": Corrupt Comment Fixed");
                            }
                            else
                            {
                                lstOutput.Items.Add(f + ": No Corrupt Comments Found");
                            }
                        }
                    }
                    catch (Exception innerEx)
                    {
                        lstOutput.Items.Add(f + Strings.wArrow + Strings.wErrorText + innerEx.Message);
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnFixCorruptComments Error: " + f + Strings.wColonBuffer + innerEx.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnFixCorruptComments Error: " + ex.Message);
                lstOutput.Items.Add(Strings.wErrorText + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnRemovePII_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                lstOutput.Items.Clear();

                foreach (string f in files)
                {
                    try
                    {
                        using (WordprocessingDocument document = WordprocessingDocument.Open(f, true))
                        {
                            if (Word.RemovePersonalInfo(document))
                            {
                                lstOutput.Items.Add(f + " : PII removed from file.");
                                FileUtilities.WriteToLog(Strings.fLogFilePath, f + " : PII removed from file.");
                            }
                            else
                            {
                                lstOutput.Items.Add(f + " : does not contain PII.");
                                FileUtilities.WriteToLog(Strings.fLogFilePath, f + " : does not contain PII.");
                            }
                        }
                    }
                    catch (Exception innerEx)
                    {
                        lstOutput.Items.Add(f + Strings.wArrow + Strings.wErrorText + innerEx.Message);
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnRemovePII Error: " + f + Strings.wColonBuffer + innerEx.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                lstOutput.Items.Add(Strings.wArrow + Strings.wErrorText + ex.Message);
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnRemovePII Error: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void RdoWord_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoWord.Checked)
            {
                EnableUI();
                PopulateAndDisplayFiles();
            }
        }

        private void RdoExcel_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoExcel.Checked)
            {
                EnableUI();
                PopulateAndDisplayFiles();
            }
        }

        private void RdoPowerPoint_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoPowerPoint.Checked)
            {
                EnableUI();
                PopulateAndDisplayFiles();
            }
        }

        private void CkbSubfolders_CheckedChanged(object sender, EventArgs e)
        {
            EnableUI();
            PopulateAndDisplayFiles();
        }

        /// <summary>
        /// There is a scenario where SharePoint property promotion will use the "Title" for the document from custom.xml instead of core.xml
        /// Removing the custom.xml title will allow SP to not change the title of the document.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnRemoveCustomTitle_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                lstOutput.Items.Clear();

                foreach (string f in files)
                {
                    try
                    {
                        if (Word.RemoveCustomTitleProp(f))
                        {
                            lstOutput.Items.Add(f + Strings.wColonBuffer + "Custom Property 'Title' Removed From File.");
                        }
                        else
                        {
                            lstOutput.Items.Add(f + Strings.wColonBuffer + "'Title' Property Not Found.");
                        }
                    }
                    catch (Exception innerEx)
                    {
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnRemoveCustomTitle Error: " + f + Strings.wColonBuffer + innerEx.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                lstOutput.Items.Add(Strings.wArrow + Strings.wErrorText + ex.Message);
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnRemoveCustomTitle Error: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void FrmBatch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { Close(); }
        }

        private void BtnResetBulletMargins_Click(object sender, EventArgs e)
        {
            lstOutput.Items.Clear();
            Cursor = Cursors.WaitCursor;

            foreach (string f in files)
            {
                try
                {
                    using (PresentationDocument document = PresentationDocument.Open(f, true))
                    {
                        PowerPointFixes.ResetBulletMargins(document);
                        lstOutput.Items.Add(f + Strings.wArrow + Strings.pptResetBulletMargins);
                        FileUtilities.WriteToLog(Strings.fLogFilePath, f + Strings.wArrow + Strings.pptResetBulletMargins);
                    }
                }
                catch (Exception ex)
                {
                    lstOutput.Items.Add(f + Strings.wArrow + Strings.wErrorText + ex.Message);
                    FileUtilities.WriteToLog(Strings.fLogFilePath, f + Strings.wArrow + Strings.wErrorText + ex.Message);
                }
                finally
                {
                    Cursor = Cursors.Default;
                }
            }
        }

        private void BtnCheckForDigSig_Click(object sender, EventArgs e)
        {
            lstOutput.Items.Clear();
            Cursor = Cursors.WaitCursor;

            foreach (string f in files)
            {
                try
                {
                    bool hasSignature = false;
                    using (Package package = Package.Open(f, FileMode.Open, FileAccess.Read))
                    {
                        foreach (PackagePart part in package.GetParts())
                        {
                            if (part.Uri.ToString().Contains("/_xmlsignatures"))
                            {
                                hasSignature = true;
                            }
                        }
                    }

                    if (hasSignature)
                    {
                        lstOutput.Items.Add(f + " : contains signature");
                    }
                }
                catch (Exception ex)
                {
                    lstOutput.Items.Add(f + Strings.wArrow + Strings.wErrorText + ex.Message);
                    FileUtilities.WriteToLog(Strings.fLogFilePath, f + Strings.wArrow + Strings.wErrorText + ex.Message);
                }
                finally
                {
                    Cursor = Cursors.Default;
                }
            }
        }

        /// <summary>
        /// known issue where too many spaces in a footer can cause display issues
        /// in the only scenario where this was used I've seen, the spaces are there to layout the page x of y 
        /// this will check for a large amount of consecutive spaces and remove them which should prevent layout issues in Word
        /// then since the spacing I've seen is to push the page x of y to the right side of the page, set the justification to Right instead of Center
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnFixFooterSpacing_Click(object sender, EventArgs e)
        {
            lstOutput.Items.Clear();
            Cursor = Cursors.WaitCursor;

            foreach (string f in files)
            {
                try
                {
                    bool isFixed = false;
                    using (WordprocessingDocument document = WordprocessingDocument.Open(f, true))
                    {
                        foreach (FooterPart fp in document.MainDocumentPart.FooterParts)
                        {
                            // check the footer for text elements with many spaces and trim it to a single space
                            IEnumerable<Paragraph> pElements = fp.Footer.Descendants<Paragraph>().ToList();
                            foreach (Paragraph p in pElements)
                            {
                                IEnumerable<O.Wordprocessing.Text> txtElements = p.Descendants<O.Wordprocessing.Text>().ToList();
                                foreach (O.Wordprocessing.Text txtElement in txtElements)
                                {
                                    // if there are text elements with a lot of spaces, remove the text element and adjust the justification as needed
                                    int count = txtElement.Text.TakeWhile(char.IsWhiteSpace).Count();
                                    if (count > 10)
                                    {
                                        txtElement.Remove();
                                        isFixed = true;

                                        // if the justification is center, move it to the right
                                        if (p.ParagraphProperties.Justification.Val == JustificationValues.Center)
                                        {
                                            p.ParagraphProperties.Justification.Val = JustificationValues.Right;
                                        }
                                    }
                                }
                            }

                            if (isFixed)
                            {
                                document.Save();
                            }
                        }
                    }

                    if (isFixed)
                    {
                        lstOutput.Items.Add(f + " : Footer Spacing Fixed.");
                    }
                    else
                    {
                        lstOutput.Items.Add(f + " : No Footer Spacing Problem Found.");
                    }
                }
                catch (Exception ex)
                {
                    lstOutput.Items.Add(f + Strings.wArrow + Strings.wErrorText + ex.Message);
                    FileUtilities.WriteToLog(Strings.fLogFilePath, f + Strings.wArrow + Strings.wErrorText + ex.Message);
                }
                finally
                {
                    Cursor = Cursors.Default;
                }
            }
        }

        /// <summary>
        /// see details in WordFixes.FixCorruptTableCellTags and WordFixes.FixGridSpan
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnFixCorruptTcTags_Click(object sender, EventArgs e)
        {
            lstOutput.Items.Clear();
            Cursor = Cursors.WaitCursor;

            foreach (string f in files)
            {
                try
                {
                    bool isFixed = false;
                    using (WordprocessingDocument document = WordprocessingDocument.Open(f, true))
                    {
                        if (Word.IsPartNull(document, "Table") == false)
                        {
                            bool tableCellCorruptionFound = false;

                            do
                            {
                                tableCellCorruptionFound = WordFixes.IsTableCellCorruptionFound(document);
                                if (tableCellCorruptionFound == true)
                                {
                                    isFixed = true;
                                }
                            } while (tableCellCorruptionFound);
                        }

                        if (isFixed)
                        {
                            document.Save();
                        }
                    }

                    if (WordFixes.FixGridSpan(f) == true)
                    {
                        isFixed = true;
                    }

                    if (isFixed)
                    {
                        lstOutput.Items.Add(f + " : Table Cells Fixed.");
                    }
                    else
                    {
                        lstOutput.Items.Add(f + " : No Table Cell Problem Found.");
                    }
                }
                catch (Exception ex)
                {
                    lstOutput.Items.Add(f + Strings.wArrow + Strings.wErrorText + ex.Message);
                    FileUtilities.WriteToLog(Strings.fLogFilePath, f + Strings.wArrow + Strings.wErrorText + ex.Message);
                }
                finally
                {
                    Cursor = Cursors.Default;
                }
            }
        }

        private void BtnRemoveCustomFileProps_Click(object sender, EventArgs e)
        {
            lstOutput.Items.Clear();
            Cursor = Cursors.WaitCursor;

            foreach (string f in files)
            {
                try
                {
                    bool isFixed = false;
                    using (WordprocessingDocument document = WordprocessingDocument.Open(f, true))
                    {
                        document.DeletePart(document.CustomFilePropertiesPart);

                        if (isFixed)
                        {
                            document.Save();
                        }
                    }

                    if (isFixed)
                    {
                        lstOutput.Items.Add(f + " : Table Cells Fixed.");
                    }
                    else
                    {
                        lstOutput.Items.Add(f + " : No Table Cell Problem Found.");
                    }
                }
                catch (Exception ex)
                {
                    lstOutput.Items.Add(f + Strings.wArrow + Strings.wErrorText + ex.Message);
                    FileUtilities.WriteToLog(Strings.fLogFilePath, f + Strings.wArrow + Strings.wErrorText + ex.Message);
                }
                finally
                {
                    Cursor = Cursors.Default;
                }
            }
        }

        private void BtnRemoveCustomXml_Click(object sender, EventArgs e)
        {
            lstOutput.Items.Clear();
            Cursor = Cursors.WaitCursor;

            foreach (string f in files)
            {
                try
                {
                    if (rdoWord.Checked)
                    {
                        Office.RemoveCustomXmlParts(pkg, f, Strings.oAppWord);
                    }
                    else if (rdoExcel.Checked)
                    {
                        Office.RemoveCustomXmlParts(pkg, f, Strings.oAppExcel);
                    }
                    else if (rdoPowerPoint.Checked)
                    {
                        Office.RemoveCustomXmlParts(pkg, f, Strings.oAppPowerPoint);
                    }
                }
                catch (Exception ex)
                {
                    lstOutput.Items.Add(f + Strings.wArrow + Strings.wErrorText + ex.Message);
                    FileUtilities.WriteToLog(Strings.fLogFilePath, f + Strings.wArrow + Strings.wErrorText + ex.Message);
                }
                finally
                {
                    Cursor = Cursors.Default;
                }
            }
        }

        /// <summary>
        /// there are times when a duplicate documentManagement custom xml file is added to a document
        /// this causes an error "The server properties in this file cannot be displayed."
        /// usually the duplicate file has 0 child elements of docmgmt, so check for that
        /// other times, there are two versions of <documentManagement/>
        /// if either are found, remove the rel reference
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnFixDupeCustomXml_Click(object sender, EventArgs e)
        {
            lstOutput.Items.Clear();
            Cursor = Cursors.WaitCursor;

            foreach (string f in files)
            {
                try
                {
                    bool isCorrupt = false;
                    bool isFixed = false;
                    string badUri = string.Empty;
                    int docMgmtCount = 0;
                    string dupeDocMgmtUri = string.Empty;

                    using (WordprocessingDocument document = WordprocessingDocument.Open(f, true))
                    {
                        // first check if there is an empty documentManagement tag in the custom xml file
                        // the second check is if there is more than one docManagement custom xml file
                        // more to do here long term, not sure how to determine which is the "current" version
                        // for now, just removing the last referenced uri in the loop
                        // also not accounting for potentially more than 2 dupes of documentManagement
                        // currently assuming at most there are 2
                        foreach (CustomXmlPart part in document.MainDocumentPart.CustomXmlParts)
                        {
                            XDocument xDoc = part.GetXDocument();
                            string badElement = xDoc.ToString();
                            if (badElement.Contains("<documentManagement />"))
                            {
                                isCorrupt = true;
                                badUri = part.Uri.ToString();
                            }
                            else if (badElement.Contains("<documentManagement>"))
                            {
                                dupeDocMgmtUri = part.Uri.ToString();
                                docMgmtCount++;
                            }
                        }

                        if (docMgmtCount > 1)
                        {
                            isCorrupt = true;
                        }
                    }

                    // if either condition is found, pull the rel file as a zipentry and remove the node for the bad reference
                    if (isCorrupt)
                    {
                        // loop package parts and push the rels into an xdoc to parse and remove the uri 
                        using (FileStream zipToOpen = new FileStream(f, FileMode.Open, FileAccess.ReadWrite))
                        {
                            using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                            {
                                foreach (ZipArchiveEntry zae in archive.Entries)
                                {
                                    if (zae.Name == "document.xml.rels")
                                    {
                                        XmlDocument xDoc = new XmlDocument();
                                        MemoryStream ms = new MemoryStream();
                                        Stream relStream = zae.Open();
                                        relStream.CopyTo(ms);
                                        ms.Position = 0;
                                        xDoc.Load(ms);
                                        XmlNodeList xnl = xDoc.ChildNodes;
                                        foreach (XmlNode xn in xnl)
                                        {
                                            if (xn.Name == "Relationships")
                                            {
                                                XmlNodeList xnlRel = xn.ChildNodes;
                                                foreach (XmlNode xnRel in xnlRel)
                                                {
                                                    foreach (XmlAttribute xa in xnRel.Attributes)
                                                    {
                                                        if (xa.Value == ".." + badUri || xa.Value == ".." + dupeDocMgmtUri)
                                                        {
                                                            // remove the node and save the changes back to the file
                                                            xn.RemoveChild(xnRel);
                                                            relStream.SetLength(0);
                                                            xDoc.Save(relStream);
                                                            relStream.Flush();
                                                            isFixed = true;
                                                            goto corruptionFound;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                corruptionFound:
                    if (isFixed)
                    {
                        lstOutput.Items.Add(f + " : Removed Duplicate Custom Xml");
                    }
                    else
                    {
                        lstOutput.Items.Add(f + " : No Duplicate Custom Xml Found.");
                    }
                }
                catch (Exception ex)
                {
                    lstOutput.Items.Add(f + Strings.wArrow + Strings.wErrorText + ex.Message);
                    FileUtilities.WriteToLog(Strings.fLogFilePath, f + Strings.wArrow + Strings.wErrorText + ex.Message);
                }
                finally
                {
                    Cursor = Cursors.Default;
                }
            }
        }

        /// <summary>
        /// adding fix for known issue with converting google to pptx using bittitan
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnFixTabBehavior_Click(object sender, EventArgs e)
        {
            lstOutput.Items.Clear();

            foreach (string f in files)
            {
                try
                {
                    Cursor = Cursors.WaitCursor;
                    bool isFMPFixed = PowerPointFixes.FixMissingPlaceholder(f);
                    bool isRDPFixed = PowerPointFixes.ResetDefaultParagraphProps(f);

                    if (isFMPFixed || isRDPFixed)
                    {
                        lstOutput.Items.Add(f + " : Tab Behavior Fixed.");
                    }
                    else
                    {
                        lstOutput.Items.Add(f + " : No Tab Behavior Problem Found.");
                    }
                }
                catch (Exception ex)
                {
                    lstOutput.Items.Add(f + Strings.wArrow + Strings.wErrorText + ex.Message);
                    FileUtilities.WriteToLog(Strings.fLogFilePath, f + Strings.wArrow + Strings.wErrorText + ex.Message);
                }
                finally
                {
                    Cursor = Cursors.Default;
                }
            }
        }

        private void BtnFixCommentNotes_Click(object sender, EventArgs e)
        {
            lstOutput.Items.Clear();

            foreach (string f in files)
            {
                try
                {
                    Cursor = Cursors.WaitCursor;
                    bool isFixed = Excel.FixCorruptAnchorTags(f);

                    if (isFixed)
                    {
                        lstOutput.Items.Add(f + " : Comment Notes Fixed.");
                    }
                    else
                    {
                        lstOutput.Items.Add(f + " : No Large Comment Notes Found.");
                    }
                }
                catch (Exception ex)
                {
                    lstOutput.Items.Add(f + Strings.wArrow + Strings.wErrorText + ex.Message);
                    FileUtilities.WriteToLog(Strings.fLogFilePath, f + Strings.wArrow + Strings.wErrorText + ex.Message);
                }
                finally
                {
                    Cursor = Cursors.Default;
                }
            }
        }

        private void btnFixContentControls_Click(object sender, EventArgs e)
        {
            lstOutput.Items.Clear();

            foreach (string f in files)
            {
                try
                {
                    Cursor = Cursors.WaitCursor;
                    StringBuilder sb = new StringBuilder();
                    bool isFixed = WordFixes.FixContentControlPlaceholders(f);

                    if (isFixed)
                    {
                        lstOutput.Items.Add(f + " : Content Controls Fixed.");
                    }
                    else
                    {
                        lstOutput.Items.Add(f + " : No Corrupt Content Controls Found.");
                    }
                }
                catch (Exception ex)
                {
                    lstOutput.Items.Add(f + Strings.wArrow + Strings.wErrorText + ex.Message);
                    FileUtilities.WriteToLog(Strings.fLogFilePath, f + Strings.wArrow + Strings.wErrorText + ex.Message);
                }
                finally
                {
                    Cursor = Cursors.Default;
                }
            }
        }
    }
}
