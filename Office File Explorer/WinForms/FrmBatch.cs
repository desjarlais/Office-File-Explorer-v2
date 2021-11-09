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
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.XPath;

// namespace refs
using O = DocumentFormat.OpenXml;
using DataBinding = DocumentFormat.OpenXml.Wordprocessing.DataBinding;

namespace Office_File_Explorer.WinForms
{
    public partial class FrmBatch : Form
    {
        public List<string> files = new List<string>();
        public string fileType = string.Empty;
        public string fType = string.Empty;
        public bool nodeDeleted = false;
        public bool nodeChanged = false;
        public string fromChangeTemplate;

        public FrmBatch()
        {
            InitializeComponent();
            DisableUI();
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
            }

            if (rdoPowerPoint.Checked == true)
            {
                BtnRemovePII.Enabled = true;
                BtnRemovePIIOnSave.Enabled = true;
                BtnFixNotesPage.Enabled = true;
            }

            if (rdoExcel.Checked == true)
            {
                BtnConvertStrict.Enabled = true;
                BtnFixHyperlinks.Enabled = true;
            }
        }

        public void PopulateAndDisplayFiles()
        {
            try
            {
                lstOutput.Items.Clear();
                files.Clear();
                int fCount = 0;

                // get all the file paths for .docx files in the folder
                DirectoryInfo dir = new DirectoryInfo(tbFolderPath.Text);
                if (ckbSubfolders.Checked == true)
                {
                    foreach (FileInfo f in dir.GetFiles(GetFileExtension(), SearchOption.AllDirectories))
                    {
                        if (f.Name.StartsWith("~"))
                        {
                            // we don't want to change temp files
                            continue;
                        }
                        else
                        {
                            // populate the list of file paths
                            files.Add(f.FullName);
                            lstOutput.Items.Add(f.FullName);
                            fCount++;
                        }
                    }
                }
                else
                {
                    foreach (FileInfo f in dir.GetFiles(GetFileExtension()))
                    {
                        if (f.Name.StartsWith("~"))
                        {
                            // we don't want to change temp files
                            continue;
                        }
                        else
                        {
                            // populate the list of file paths
                            files.Add(f.FullName);
                            lstOutput.Items.Add(f.FullName);
                            fCount++;
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
            lstOutput.Items.Add("** Batch Processing done **");
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
                                if (propNameToDelete == "Cancel")
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
                            FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnListCustomProps Error: " + f + " : " + innerEx.Message);
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
                                if (propNameToDelete == "Cancel")
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
                            FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnListCustomProps Error: " + f + " : " + innerEx.Message);
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
                                if (propNameToDelete == "Cancel")
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
                            FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnListCustomProps Error: " + f + " : " + innerEx.Message);
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
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnFixExcelHyperlinks Error: " + f + " : " + innerEx.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnFixExcelHyperlinks Error: " + ex.Message);
                lstOutput.Items.Add("Error - " + ex.Message);
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
                                string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                                filePath = userProfile + "\\AppData\\Roaming\\Microsoft\\Templates\\Normal.dotm";

                                if (!File.Exists(filePath))
                                {
                                    // Normal.dotm path is not correct?
                                    FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnChangeDefaultTemplate Error: " + "Invalid Attached Template Path - " + filePath);
                                    throw new Exception();
                                }
                            }

                            if (fromChangeTemplate == filePath || fromChangeTemplate is null || fromChangeTemplate == "Cancel")
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
                                lstOutput.Items.Add(f + "** No Changed Made To Attached Template **");
                            }
                        }
                    }
                    catch (Exception innerEx)
                    {
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnChangeAttachedTemplate Error: " + innerEx.Message);
                        lstOutput.Items.Add("Error - " + innerEx.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnChangeAttachedTemplate Error: " + ex.Message);
                lstOutput.Items.Add("Error - " + ex.Message);
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
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnDeleteOpenByDefault Error: " + f + " : " + innerEx.Message);
                    }
                }
            }
            catch (IOException ioe)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnDeleteOpenByDefault Error: " + ioe.Message);
                lstOutput.Items.Add("Error - " + ioe.Message);
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnDeleteOpenByDefault Error: " + ex.Message);
                lstOutput.Items.Add("Error - " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnUpdateNamespaces_Click(object sender, EventArgs e)
        {
            try
            {
                lstOutput.Items.Clear();

                List<string> nsList = new List<string>();
                List<string> nList = new List<string>();
                List<string> tList = new List<string>();
                List<string> ntList = new List<string>();
                string targetNS = string.Empty;

                foreach (string f in files)
                {
                    try
                    {
                        bool fileChanged = false;
                        nsList.Clear();
                        nList.Clear();
                        tList.Clear();
                        ntList.Clear();

                        using (WordprocessingDocument document = WordprocessingDocument.Open(f, true))
                        {
                            // now that we know the namespaces, loop the controls and update their data binding prefix mappings
                            foreach (var cc in document.ContentControls())
                            {
                                string ccValue = string.Empty;
                                SdtProperties props = cc.Elements<SdtProperties>().FirstOrDefault();

                                foreach (OpenXmlElement oxe in props.ChildElements)
                                {
                                    // get the alias and update ccValue
                                    if (oxe.GetType().ToString() == Strings.dfowStdAlias)
                                    {
                                        foreach (OpenXmlAttribute oxa in oxe.GetAttributes())
                                        {
                                            ccValue = oxa.Value;
                                            tList.Add(ccValue);
                                        }
                                    }
                                }
                            }

                            // get the schemareferences
                            foreach (CustomXmlPart cxp in document.MainDocumentPart.CustomXmlParts)
                            {
                                XmlDocument xDoc = new XmlDocument();
                                xDoc.Load(cxp.GetStream());

                                if (xDoc.DocumentElement.NamespaceURI == Strings.schemaMetadataProperties)
                                {
                                    // loop through the metadata and get the uri's
                                    foreach (XmlNode xNode in xDoc.ChildNodes)
                                    {
                                        if (xNode.Name == "p:properties")
                                        {
                                            foreach (XmlAttribute a in xNode.Attributes)
                                            {
                                                nsList.Add(a.Value);
                                            }

                                            foreach (XmlNode xNode2 in xNode.ChildNodes)
                                            {
                                                if (xNode2.Name == "documentManagement")
                                                {
                                                    foreach (XmlNode xNode3 in xNode2.ChildNodes)
                                                    {
                                                        // add the node name and uri to a global list for comparing later
                                                        nsList.Add(xNode3.NamespaceURI);
                                                        nList.Add(xNode3.Name);
                                                    }
                                                }
                                            }

                                            // remove any duplicates
                                            nsList = nsList.Distinct().ToList();
                                        }
                                    }
                                }

                                if (xDoc.DocumentElement.NamespaceURI == Strings.schemaContentType)
                                {
                                    // get metadata field name map
                                    foreach (XmlNode xNode in xDoc.ChildNodes)
                                    {
                                        if (xNode.Name == "ct:contentTypeSchema")
                                        {
                                            foreach (XmlNode xNode2 in xNode.ChildNodes)
                                            {
                                                if (xNode2.InnerXml.Contains(Strings.schemaTypes))
                                                {
                                                    foreach (XmlNode xNode3 in xNode2.ChildNodes)
                                                    {
                                                        if (xNode3.InnerXml.Contains("complexType"))
                                                        {
                                                            foreach (string t in tList)
                                                            {
                                                                if (xNode3.OuterXml.Contains(string.Concat("\"", t, "\"")))
                                                                {
                                                                    string xNodeElement1 = string.Empty;
                                                                    int iNodeElement1 = xNode3.OuterXml.IndexOf(" name=") + 7;
                                                                    xNodeElement1 = xNode3.OuterXml.Substring(iNodeElement1, 32);

                                                                    int iNodeElement2 = xNode2.OuterXml.IndexOf(" targetNamespace=") + 18;
                                                                    if (targetNS == string.Empty)
                                                                    {
                                                                        targetNS = xNode2.OuterXml.Substring(iNodeElement2, 36);
                                                                    }

                                                                    ntList.Add(string.Concat(t, xNodeElement1));
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    // remove any duplicates
                                    ntList = ntList.Distinct().ToList();
                                }
                            }

                            // now that we know the namespaces, loop the controls and update their data binding prefix mappings
                            foreach (var cc in document.ContentControls())
                            {
                                string ccValue = string.Empty;
                                string ccName = string.Empty;
                                SdtProperties props = cc.Elements<SdtProperties>().FirstOrDefault();

                                foreach (OpenXmlElement oxe in props.ChildElements)
                                {
                                    // get the cc Value from the alias
                                    if (oxe.GetType().ToString() == Strings.dfowStdAlias)
                                    {
                                        foreach (OpenXmlAttribute oxa in oxe.GetAttributes())
                                        {
                                            ccValue = oxa.Value;
                                        }
                                    }
                                    // get the cc name from the tag
                                    if (oxe.GetType().ToString() == Strings.dfowTag)
                                    {
                                        foreach (OpenXmlAttribute oxa in oxe.GetAttributes())
                                        {
                                            ccName = oxa.Value;
                                        }
                                    }

                                    // now use databinding to check the prefix mappings
                                    if (oxe.GetType().ToString() == Strings.dfowDataBinding)
                                    {
                                        // create the DataBinding object
                                        DataBinding db = (DataBinding)oxe;

                                        if (ccName == string.Empty)
                                        {
                                            // parse out the element name
                                            string[] elemName = db.XPath.ToString().Split('/');
                                            string xPathName = elemName[elemName.Count() - 1];
                                            string mappingName = xPathName.Substring(4, xPathName.Length - 7);
                                        }

                                        bool foundName = false;
                                        // check if the content control prefix map name matches the metadata xml
                                        foreach (string sName in nList)
                                        {
                                            if (sName == ccName)
                                            {
                                                foundName = true;
                                                break;
                                            }
                                        }
                                        if (foundName == false)
                                        {
                                            foreach (string tName in ntList)
                                            {
                                                if (tName.Contains(ccValue))
                                                {
                                                    if ((tName.Substring(tName.IndexOf(ccValue) + ccValue.Length)).Length == 32)
                                                    {
                                                        db.XPath.Value = db.XPath.Value.Replace(ccName, tName.Substring(tName.IndexOf(ccValue) + ccValue.Length));
                                                        foundName = true;
                                                    }
                                                }
                                            }

                                        }
                                        if (foundName == true)
                                        {
                                            // parse out the namespace mapping but remove the space at the end first
                                            string[] prefixMappingNamespaces = db.PrefixMappings.Value.TrimEnd().Split(' ');
                                            string dValue = db.XPath.Value.Substring(db.XPath.Value.IndexOf("documentManagement[1]/ns") + "documentManagement[1]/ns".Length, 1);
                                            int nsIndex = 0;

                                            // go through each namespace to compare with the content controls namespace value
                                            foreach (var s in prefixMappingNamespaces)
                                            {
                                                bool sMatch = false;
                                                string xSubstring = string.Empty;

                                                // the first mapping usually doesn't have xmlns at the beginning
                                                if (s.StartsWith("xmlns:"))
                                                {
                                                    xSubstring = s.Substring(11, s.Length - 11);
                                                    xSubstring = xSubstring.Substring(0, xSubstring.Length - 1);
                                                }
                                                else
                                                {
                                                    xSubstring = s.Substring(0, s.Length);
                                                }

                                                foreach (string ns in nsList)
                                                {
                                                    if (xSubstring == ns)
                                                    {
                                                        sMatch = true;
                                                        break;
                                                    }

                                                }
                                                if (sMatch == false && nsIndex == short.Parse(dValue))
                                                {
                                                    // add the xmlns to the guid
                                                    string valToReplace = "xmlns:ns" + nsIndex + "='" + targetNS + "'";
                                                    db.PrefixMappings.Value = db.PrefixMappings.Value.Replace(s, valToReplace);
                                                    fileChanged = true;
                                                }

                                                nsIndex++;
                                            }

                                        }
                                    }
                                }
                            }

                            if (fileChanged)
                            {
                                lstOutput.Items.Add(f + "** Namespace Guid Updated **");
                                document.MainDocumentPart.Document.Save();
                            }
                            else
                            {
                                lstOutput.Items.Add(f + "** No Namespace Needed To Be Updated **");
                            }
                        }
                    }
                    catch (Exception innerEx)
                    {
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnUpdateQuickPartNamespaces Error: " + f + " : " + innerEx.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnUpdateQuickPartNamespaces Error: " + ex.Message);
                lstOutput.Items.Add("Error - " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
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
                                        string docText = null;
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
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnConvertStrict: " + f + " : " + innerEx.Message);
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
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnPPTResetPII: " + f + " : " + innerEx.Message);
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
                        if (WordFixes.RemoveMissingBookmarkTags(f) == true || WordFixes.RemovePlainTextCcFromBookmark(f) == true)
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
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnFixCorruptBookmarks: " + f + " : " + innerEx.Message);
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
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnFixCorruptRevisions: " + f + " : " + innerEx.Message);
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
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnFixTableProps: " + f + " : " + innerEx.Message);
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
                            FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnDeleteRequestStatus Error: " + f + " : " + innerEx.Message);
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
                            FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnDeleteRequestStatus Error: " + f + " : " + innerEx.Message);
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
                            FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnDeleteRequestStatus Error: " + f + " : " + innerEx.Message);
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
                lstOutput.Items.Add("Error - " + ioe.Message);
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnDeleteRequestStatus Error: " + ex.Message);
                lstOutput.Items.Add("Error - " + ex.Message);
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

                            if (commentsPart is null && commentRefs.Count() > 0)
                            {
                                // if there are comment refs but no comments.xml, remove refs
                                foreach (CommentReference cr in commentRefs)
                                {
                                    cr.Remove();
                                    saveFile = true;
                                }
                            }
                            else if (commentsPart is null && commentRefs.Count() == 0)
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
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnFixCorruptComments Error: " + f + " : " + innerEx.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnFixCorruptComments Error: " + ex.Message);
                lstOutput.Items.Add("Error - " + ex.Message);
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
                            if (Word.HasPersonalInfo(document) == true)
                            {
                                Word.RemovePersonalInfo(document);
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
                        FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnRemovePII Error: " + f + " : " + innerEx.Message);
                    }
                }
            }
            catch (Exception ex)
            {
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
            PopulateAndDisplayFiles();
        }
    }
}
