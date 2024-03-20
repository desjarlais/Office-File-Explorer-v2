// Open XML SDK refs
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

// App refs
using Office_File_Explorer.Helpers;
using Office_File_Explorer.WinForms;
using Office_File_Explorer.OpenMcdf;

// .NET refs
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using System.Drawing;
using System.Xml.Schema;
using System.Xml.Linq;

// named refs
using File = System.IO.File;
using Color = System.Drawing.Color;
using Person = DocumentFormat.OpenXml.Office2013.Word.Person;
using Application = System.Windows.Forms.Application;

// scintilla refs
using ScintillaNET;
using ScintillaNET_FindReplaceDialog;

namespace Office_File_Explorer
{
    public partial class FrmMain : Form
    {
        // global variables
        private string findText;
        private string replaceText;
        private string fromChangeTemplate;
        private string partPropContentType;
        private string partPropCompression;
        private bool isEncrypted;

        // openmcdf globals
        private FileStream fs;
        private CompoundFile cf;
        private CFStream cfStream;
        private bool isValidXml;
        private List<string> validationErrors = new List<string>();

        // corrupt doc legacy
        private static string StrCopiedFileName = string.Empty;
        private static string StrOfficeApp = string.Empty;
        private static char PrevChar = Strings.chLessThan;
        private bool IsRegularXmlTag;
        private bool IsFixed;
        private static string FixedFallback = string.Empty;
        private static string StrExtension = string.Empty;
        private static string StrDestFileName = string.Empty;
        static StringBuilder sbNodeBuffer = new StringBuilder();

        // temp files
        public static string tempFileReadOnly, tempFilePackageViewer;

        // lists
        private static List<string> corruptNodes = new List<string>();
        private static List<string> pParts = new List<string>();
        private List<string> oNumIdList = new List<string>();
        private Dictionary<string, Image> attachmentList = new Dictionary<string, Image>();

        // part viewer globals
        public List<PackagePart> pkgParts = new List<PackagePart>();

        // package is for viewing of contents only
        public Package package;

        // scintilla globals
        public FindReplace myFindReplace;

        // enums
        public enum OpenXmlInnerFileTypes { Word, Excel, PowerPoint, Outlook, XML, Image, Binary, Video, Audio, Text, Other }

        public enum LogInfoType { ClearAndAdd, TextOnly, InvalidFile, LogException, EmptyCount }

        public enum UIType { OpenWord, OpenExcel, OpenPowerPoint, OpenMsg, OpenCF, FileClose, ViewLabelInfo, ViewXml, ViewBinary, ViewImage, ViewCustomUI }

        public FrmMain()
        {
            InitializeComponent();

            this.Text = Strings.oAppTitle + Strings.wMinusSign + Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyFileVersionAttribute>().Version;

            // make sure the log file is created
            if (!File.Exists(Strings.fLogFilePath))
            {
                File.Create(Strings.fLogFilePath);
            }

            UpdateMRU();

            // scintilla inits
            scintilla1.ReadOnly = true;
            myFindReplace = new FindReplace(scintilla1);
            myFindReplace.KeyPressed += MyFindReplace_KeyPressed;
            scintilla1.KeyDown += Scintilla1_KeyDown;
        }

        #region Class Properties

        public string FindTextProperty
        {
            set => findText = value;
        }

        public string ReplaceTextProperty
        {
            set => replaceText = value;
        }

        public string DefaultTemplate
        {
            set => fromChangeTemplate = value;
        }
        #endregion

        #region Functions

        /// <summary>
        /// append strings like "Fixed" or "Copy" to the file name
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="TextToAdd"></param>
        /// <returns></returns>
        public static string AddModifiedTextToFileName(string fileName)
        {
            string dir = Path.GetDirectoryName(fileName) + "\\";
            StrExtension = Path.GetExtension(fileName);
            string newFileName = dir + Path.GetFileNameWithoutExtension(fileName) + Strings.wModified + StrExtension;
            return newFileName;
        }

        /// <summary>
        /// refresh the MRU UI
        /// </summary>
        public void UpdateMRU()
        {
            try
            {
                int index = 1;
                foreach (var f in Properties.Settings.Default.FileMRU)
                {
                    switch (index)
                    {
                        case 1: mruToolStripMenuItem1.Text = f.ToString(); break;
                        case 2: mruToolStripMenuItem2.Text = f.ToString(); break;
                        case 3: mruToolStripMenuItem3.Text = f.ToString(); break;
                        case 4: mruToolStripMenuItem4.Text = f.ToString(); break;
                        case 5: mruToolStripMenuItem5.Text = f.ToString(); break;
                        case 6: mruToolStripMenuItem6.Text = f.ToString(); break;
                        case 7: mruToolStripMenuItem7.Text = f.ToString(); break;
                        case 8: mruToolStripMenuItem8.Text = f.ToString(); break;
                        case 9: mruToolStripMenuItem9.Text = f.ToString(); break;
                    }
                    index++;
                }
            }
            catch (Exception ex)
            {
                // log the error and do not update mru
                LogInformation(LogInfoType.LogException, "UpdateMRU Error: ", ex.Message);
            }
        }

        /// <summary>
        /// loop the filemru and remove the entry
        /// </summary>
        /// <param name="fPath"></param>
        public void RemoveFileFromMRU(string fPath)
        {
            for (int i = 0; i < 9; i++)
            {
                if (fPath == Properties.Settings.Default.FileMRU[i])
                {
                    Properties.Settings.Default.FileMRU.RemoveAt(i);
                    Properties.Settings.Default.Save();
                    ClearRecentMenuItems();
                    UpdateMRU();
                    MessageBox.Show(fPath + " does not exist, removing from MRU", "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
                }
            }
        }

        /// <summary>
        /// check if the file is already in the MRU list and add it if not
        /// </summary>
        public void AddFileToMRU()
        {
            bool isFileInMru = false;
            foreach (var f in Properties.Settings.Default.FileMRU)
            {
                if (f.ToString() == toolStripStatusLabelFilePath.Text)
                {
                    isFileInMru = true;
                }
            }

            if (!isFileInMru)
            {
                Properties.Settings.Default.FileMRU.Add(toolStripStatusLabelFilePath.Text);
                if (Properties.Settings.Default.FileMRU.Count > 9)
                {
                    Properties.Settings.Default.FileMRU.RemoveAt(0);
                }
                UpdateMRU();
            }
        }

        /// <summary>
        /// tempFileReadOnly is used for the View Contents feature
        /// tempFilePackageViewer is used for the main form part viewer
        /// changes made in the part viewer are then saved back to the toolstripstatusfilepath
        /// </summary>
        public void TempFileSetup()
        {
            try
            {
                string fExtension = Path.GetExtension(toolStripStatusLabelFilePath.Text);

                tempFileReadOnly = Path.GetTempFileName().Replace(".tmp", fExtension);
                File.Copy(toolStripStatusLabelFilePath.Text, tempFileReadOnly, true);

                tempFilePackageViewer = Path.GetTempFileName().Replace(".tmp", fExtension);
                File.Copy(toolStripStatusLabelFilePath.Text, tempFilePackageViewer, true);

            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "Temp File Setup Error:");
                FileUtilities.WriteToLog(Strings.fLogFilePath, ex.Message);
            }
        }

        /// <summary>
        /// handle when user clicks File | Close
        /// </summary>
        public void FileClose()
        {
            DisableAllUI();
            package?.Close();
            pkgParts?.Clear();
            TvFiles.Nodes.Clear();
            toolStripStatusLabelFilePath.Text = Strings.wHeadingBegin;
            toolStripStatusLabelDocType.Text = Strings.wHeadingBegin;
            renderRTFToolStripMenuItem.Enabled = false;
            fileToolStripMenuItemClose.Enabled = false;
            scintilla1.ReadOnly = false;
            scintilla1.ClearAll();

            // check for encrypted file
            fs?.Close();
        }

        public void CopyAllItems()
        {
            try
            {
                if (scintilla1.Text.Length == 0) { return; }
                scintilla1.Copy();
            }
            catch (Exception ex)
            {
                LogInformation(LogInfoType.LogException, "BtnCopyOutput Error", ex.Message);
            }
        }

        public void OpenEncryptedOfficeDocument(string fileName, bool enableCommit)
        {
            try
            {
                fs = new FileStream(fileName, FileMode.Open, enableCommit ? FileAccess.ReadWrite : FileAccess.Read);

                try
                {
                    cf = new CompoundFile(fs, CFSUpdateMode.Update, CFSConfiguration.SectorRecycle | CFSConfiguration.NoValidationException | CFSConfiguration.EraseFreeSectors);

                    // populate treeview
                    TvFiles.Nodes.Clear();
                    TreeNode root = null;
                    root = TvFiles.Nodes.Add("Root Entry", "Root");
                    root.Tag = cf.RootStorage;
                    root.ImageIndex = 5;
                    root.SelectedImageIndex = 5;
                    AddNodes(root, cf.RootStorage);
                    TvFiles.ExpandAll();
                    isEncrypted = true;
                }
                catch (Exception ex)
                {
                    FileUtilities.WriteToLog(Strings.fLogFilePath, ex.Message);
                    MessageBox.Show(ex.Message, "File Load Fail", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    FileClose();
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogInfoType.LogException, "OpenEncryptedOfficeDocument Error", ex.Message);
            }
        }

        public void DisplayInvalidFileFormatError()
        {
            scintilla1.Clear();
            scintilla1.AppendText("Unable to open file, possible causes are:\r\n");
            scintilla1.AppendText(" - file corruption\r\n");
            scintilla1.AppendText(" - file encrypted\r\n");
            scintilla1.AppendText(" - file password protected\r\n");
            scintilla1.AppendText(" - binary Office Document (View file contents with Tools -> Structured Storage Viewer)\r\n");
        }

        /// <summary>
        /// Recursive addition of tree nodes foreach child of current item in the storage
        /// </summary>
        /// <param name="node">Current TreeNode</param>
        /// <param name="cfs">Current storage associated with node</param>
        private static void AddNodes(TreeNode node, CFStorage cfs)
        {
            Action<CFItem> va = delegate (CFItem target)
            {
                TreeNode temp = node.Nodes.Add(target.Name, target.Name + (target.IsStream ? " (" + target.Size + " bytes )" : string.Empty));
                temp.Tag = target;

                // set images for treeview
                if (target.IsStream)
                {
                    temp.ImageIndex = 5;
                    temp.SelectedImageIndex = 5;
                }
                else
                {
                    temp.ImageIndex = 6;
                    temp.SelectedImageIndex = 6;

                    // Recursion into the storage
                    AddNodes(temp, (CFStorage)target);
                }
            };

            //Visit NON-recursively (first level only)
            cfs.VisitEntries(va, false);
        }

        /// <summary>
        /// majority of open file logic is here
        /// </summary>
        public void OpenOfficeDocument(bool isFromMRU)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                isEncrypted = false;

                if (!isFromMRU)
                {
                    OpenFileDialog fDialog = new OpenFileDialog
                    {
                        Title = "Select Office Open Xml File.",
                        Filter = "Open XML Files | *.docx; *.dotx; *.docm; *.dotm; *.xlsx; *.xlsm; *.xlst; *.xltm; *.pptx; *.pptm; *.potx; *.potm|" +
                                 "Binary Office Documents | *.doc; *.dot; *.xls; *.xlt; *.ppt; *.pot|" +
                                 "Outlook Message Format | *.msg",
                        RestoreDirectory = true,
                        InitialDirectory = @"%userprofile%"
                    };

                    if (fDialog.ShowDialog() == DialogResult.OK)
                    {
                        toolStripStatusLabelFilePath.Text = fDialog.FileName.ToString();
                    }
                    else
                    {
                        // user cancelled dialog, disable the UI and go back to the form
                        DisableAllUI();
                        toolStripStatusLabelFilePath.Text = Strings.wHeadingBegin;
                        toolStripStatusLabelDocType.Text = Strings.wHeadingBegin;
                        return;
                    }
                }

                if (!File.Exists(toolStripStatusLabelFilePath.Text))
                {
                    LogInformation(LogInfoType.InvalidFile, Strings.fileDoesNotExist, string.Empty);
                    RemoveFileFromMRU(toolStripStatusLabelFilePath.Text);
                }
                else
                {
                    scintilla1.Clear();

                    // handle msg files
                    if (toolStripStatusLabelFilePath.Text.EndsWith(Strings.msgFileExt))
                    {
                        // setup message file
                        Stream messageStream = File.Open(toolStripStatusLabelFilePath.Text, FileMode.Open, FileAccess.Read);
                        OutlookStorage.Message message = new OutlookStorage.Message(messageStream);
                        messageStream.Close();

                        // load tree
                        TvFiles.Nodes.Clear();
                        LoadMsgToTree(message, TvFiles.Nodes.Add("MSG"));
                        TvFiles.ImageIndex = 6;
                        TvFiles.SelectedImageIndex = 6;
                        toolStripStatusLabelDocType.Text = Strings.oAppOutlook;
                        TvFiles.ExpandAll();
                        ScrollToTopOfRtb();

                        message.Dispose();
                        return;
                    }

                    // cleanup in case this is an open and no close of previous file
                    fs?.Close();

                    // handle office files
                    if (!FileUtilities.IsZipArchiveFile(toolStripStatusLabelFilePath.Text))
                    {
                        // check for encrypted files
                        if (FileUtilities.IsFileEncrypted(toolStripStatusLabelFilePath.Text))
                        {
                            OpenEncryptedOfficeDocument(toolStripStatusLabelFilePath.Text, true);
                            EnableModifyUI();
                        }
                        else
                        {
                            // if the file is not a zip or encrypted, it is not a valid office file
                            DisplayInvalidFileFormatError();
                            DisableAllUI();
                        }
                    }
                    else
                    {
                        // if the file does start with PK, check if it fails in the SDK
                        if (OpenWithSdk(toolStripStatusLabelFilePath.Text))
                        {
                            // set the file type
                            toolStripStatusLabelDocType.Text = StrOfficeApp;
                            DisableAllUI();

                            // populate the parts
                            PopulatePackageParts();

                            // check if any zip items are corrupt
                            if (Properties.Settings.Default.CheckZipItemCorrupt == true && toolStripStatusLabelDocType.Text == Strings.oAppWord)
                            {
                                if (Office.IsZippedFileCorrupt(toolStripStatusLabelFilePath.Text))
                                {
                                    scintilla1.AppendText("Warning - One of the zipped items is corrupt.");
                                }
                            }

                            // clear the previous doc if there was one and setup temp files
                            TempFileSetup();
                            TvFiles.Nodes.Clear();
                            scintilla1.Clear();
                            package?.Close();
                            pkgParts?.Clear();
                            LoadPartsIntoViewer();
                        }
                        else
                        {
                            // if it failed the SDK, disable all buttons except the fix corrupt doc button
                            DisableAllUI();
                            if (toolStripStatusLabelFilePath.Text.EndsWith(Strings.docxFileExt))
                            {
                                toolStripButtonFixCorruptDoc.Enabled = true;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogInfoType.LogException, "File Open Error: ", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void LoadMsgToTree(OutlookStorage.Message message, TreeNode messageNode)
        {
            messageNode.Text = message.Subject;
            messageNode.Nodes.Add("Subject: " + message.Subject);
            TreeNode bodyNode = messageNode.Nodes.Add("Body: (click to view)");
            bodyNode.Tag = new string[] { message.BodyText, message.BodyRTF };

            TreeNode recipientNode = messageNode.Nodes.Add("Recipients: " + message.Recipients.Count);
            foreach (OutlookStorage.Recipient recipient in message.Recipients)
            {
                recipientNode.Nodes.Add(recipient.Type + Strings.wColon + recipient.Email);
                recipientNode.Tag = new string[] { "Display Name: " + recipient.DisplayName, "Email: " + recipient.Email };
            }

            attachmentList.Clear();
            TreeNode attachmentNode = messageNode.Nodes.Add("Attachments: " + message.Attachments.Count);
            foreach (OutlookStorage.Attachment attachment in message.Attachments)
            {
                attachmentNode.Nodes.Add(attachment.Filename + Strings.wColon + attachment.Data.Length + "b");
                Stream imageSource = new MemoryStream(attachment.Data);
                Image image = Image.FromStream(imageSource);
                attachmentList.Add(attachment.Filename, image);
            }

            TreeNode subMessageNode = messageNode.Nodes.Add("Sub Messages: " + message.Messages.Count);
            foreach (OutlookStorage.Message subMessage in message.Messages)
            {
                LoadMsgToTree(subMessage, subMessageNode.Nodes.Add("MSG"));
            }
        }

        /// <summary>
        /// load the doc parts into the tree
        /// </summary>
        public void LoadPartsIntoViewer()
        {
            package = Package.Open(toolStripStatusLabelFilePath.Text, FileMode.Open, FileAccess.ReadWrite);

            TreeNode tRoot = new TreeNode();
            tRoot.Text = toolStripStatusLabelFilePath.Text;

            // update file icon
            if (StrOfficeApp == Strings.oAppWord)
            {
                TvFiles.SelectedImageIndex = 0;
                TvFiles.ImageIndex = 0;
            }
            else if (StrOfficeApp == Strings.oAppExcel)
            {
                TvFiles.SelectedImageIndex = 2;
                TvFiles.ImageIndex = 2;
            }
            else if (StrOfficeApp == Strings.oAppPowerPoint)
            {
                TvFiles.SelectedImageIndex = 1;
                TvFiles.ImageIndex = 1;
            }

            // update inner file icon, need to update both the selected and normal image index
            foreach (PackagePart part in package.GetParts())
            {
                tRoot.Nodes.Add(part.Uri.ToString());

                if (FileUtilities.GetFileType(part.Uri.ToString()) == OpenXmlInnerFileTypes.XML)
                {
                    tRoot.Nodes[tRoot.Nodes.Count - 1].ImageIndex = 3;
                    tRoot.Nodes[tRoot.Nodes.Count - 1].SelectedImageIndex = 3;
                }
                else if (FileUtilities.GetFileType(part.Uri.ToString()) == OpenXmlInnerFileTypes.Image)
                {
                    tRoot.Nodes[tRoot.Nodes.Count - 1].ImageIndex = 4;
                    tRoot.Nodes[tRoot.Nodes.Count - 1].SelectedImageIndex = 4;
                }
                else if (FileUtilities.GetFileType(part.Uri.ToString()) == OpenXmlInnerFileTypes.Word)
                {
                    tRoot.Nodes[tRoot.Nodes.Count - 1].ImageIndex = 0;
                    tRoot.Nodes[tRoot.Nodes.Count - 1].SelectedImageIndex = 0;
                }
                else if (FileUtilities.GetFileType(part.Uri.ToString()) == OpenXmlInnerFileTypes.Excel)
                {
                    tRoot.Nodes[tRoot.Nodes.Count - 1].ImageIndex = 2;
                    tRoot.Nodes[tRoot.Nodes.Count - 1].SelectedImageIndex = 2;
                }
                else if (FileUtilities.GetFileType(part.Uri.ToString()) == OpenXmlInnerFileTypes.PowerPoint)
                {
                    tRoot.Nodes[tRoot.Nodes.Count - 1].ImageIndex = 1;
                    tRoot.Nodes[tRoot.Nodes.Count - 1].SelectedImageIndex = 1;
                }
                else if (FileUtilities.GetFileType(part.Uri.ToString()) == OpenXmlInnerFileTypes.Binary)
                {
                    tRoot.Nodes[tRoot.Nodes.Count - 1].ImageIndex = 5;
                    tRoot.Nodes[tRoot.Nodes.Count - 1].SelectedImageIndex = 5;
                }
                else
                {
                    tRoot.Nodes[tRoot.Nodes.Count - 1].ImageIndex = 7;
                    tRoot.Nodes[tRoot.Nodes.Count - 1].SelectedImageIndex = 7;
                }

                pkgParts.Add(part);
            }

            TvFiles.Nodes.Add(tRoot);
            TvFiles.ExpandAll();
            DisableModifyUI();
        }

        public void LogInformation(LogInfoType type, string output, string ex)
        {
            switch (type)
            {
                case LogInfoType.ClearAndAdd:
                    scintilla1.Clear();
                    scintilla1.AppendText(output);
                    break;
                case LogInfoType.InvalidFile:
                    scintilla1.Clear();
                    scintilla1.AppendText(Strings.invalidFile);
                    break;
                case LogInfoType.LogException:
                    scintilla1.Clear();
                    scintilla1.AppendText(output + "\r\n" + ex);
                    FileUtilities.WriteToLog(Strings.fLogFilePath, output);
                    FileUtilities.WriteToLog(Strings.fLogFilePath, ex);
                    break;
                case LogInfoType.EmptyCount:
                    scintilla1.AppendText(Strings.wNone);
                    break;
                default:
                    scintilla1.AppendText(output);
                    break;
            }
        }

        /// <summary>
        /// add package part details to a global list
        /// basically this is only used for the View Contents button
        /// </summary>
        public void PopulatePackageParts()
        {
            pParts.Clear();
            using (FileStream zipToOpen = new FileStream(toolStripStatusLabelFilePath.Text, FileMode.Open, FileAccess.Read))
            {
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Read))
                {
                    foreach (ZipArchiveEntry zae in archive.Entries)
                    {
                        pParts.Add(zae.FullName + Strings.wColonBuffer + FileUtilities.SizeSuffix(zae.Length));
                    }
                    pParts.Sort();
                }
            }
        }

        /// <summary>
        /// open a file in the SDK, any failure means it is not a valid open xml file
        /// </summary>
        /// <param name="file">the path to the initial fix attempt</param>
        public bool OpenWithSdk(string file)
        {
            bool fSuccess = false;

            try
            {
                Cursor = Cursors.WaitCursor;

                string body;
                if (FileUtilities.GetAppFromFileExtension(file) == Strings.oAppWord)
                {
                    using (WordprocessingDocument document = WordprocessingDocument.Open(file, true))
                    {
                        // try to get the localname of the document.xml file, if it fails, it is not a Word file
                        StrOfficeApp = Strings.oAppWord;
                        body = document.MainDocumentPart.Document.LocalName;
                        fSuccess = true;
                    }
                }
                else if (FileUtilities.GetAppFromFileExtension(file) == Strings.oAppExcel)
                {
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(file, true))
                    {
                        // try to get the localname of the workbook.xml and file if it fails, its not an Excel file
                        StrOfficeApp = Strings.oAppExcel;
                        body = document.WorkbookPart.Workbook.LocalName;
                        fSuccess = true;
                    }
                }
                else if (FileUtilities.GetAppFromFileExtension(file) == Strings.oAppPowerPoint)
                {
                    using (PresentationDocument document = PresentationDocument.Open(file, true))
                    {
                        // try to get the presentation.xml local name, if it fails it is not a PPT file
                        StrOfficeApp = Strings.oAppPowerPoint;
                        body = document.PresentationPart.Presentation.LocalName;
                        fSuccess = true;
                    }
                }
                else
                {
                    // file is corrupt or not an Office document
                    StrOfficeApp = Strings.oAppUnknown;
                    LogInformation(LogInfoType.ClearAndAdd, "Invalid File", string.Empty);
                }
            }
            catch (InvalidOperationException ioe)
            {
                LogInformation(LogInfoType.LogException, Strings.errorOpenWithSDK, ioe.Message + Strings.wMinusSign + ioe.StackTrace);
            }
            catch (Exception ex)
            {
                // if the file failed to open in the sdk, it is invalid or corrupt and we need to stop opening
                LogInformation(LogInfoType.LogException, Strings.errorOpenWithSDK, ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }

            return fSuccess;
        }

        /// <summary>
        /// output content to the textbox
        /// </summary>
        /// <param name="output">the list of content to display</param>
        /// <param name="type">the type of content to display</param>
        public static StringBuilder DisplayListContents(List<string> output, string type)
        {
            StringBuilder sb = new StringBuilder();
            // add title text for the contents
            sb.AppendLine(Strings.wHeadingBegin + type + Strings.wHeadingEnd);

            // no content to display
            if (output.Count == 0)
            {
                sb.AppendLine(string.Empty);
                return sb;
            }

            // if we have any values, display them
            foreach (string s in output)
            {
                sb.AppendLine(Strings.wTripleSpace + s);
            }

            sb.AppendLine(string.Empty);
            return sb;
        }

        /// <summary>
        /// update the node buffer for BtnFixCorruptDoc_Click logic
        /// </summary>
        /// <param name="input"></param>
        public static void Node(char input)
        {
            sbNodeBuffer.Append(input);
        }

        /// <summary>
        /// There are two known label xml corruptions
        /// 1. missing method attribute
        /// 2. siteid missing brackets
        /// This function will fix both of these issues
        /// </summary>
        public void FixLabelInfo()
        {
            // populate validation errors
            ValidateLabelInfoXml(false);

            // check existing validation errors against known corrupt labelinfo xml scenarios
            if (validationErrors.Count > 0)
            {
                foreach (string s in validationErrors)
                {
                    // fix missing method attribute
                    if (s.Contains("The required attribute 'method' is missing"))
                    {
                        scintilla1.Text = scintilla1.Text.Replace("enabled=\"1\"", "enabled=\"1\" method=\"Standard\"");
                    }

                    // fix siteid missing brackets
                    // example siteId="11111111-1111-1111-1111-111111111111"
                    // should be siteId="{11111111-1111-1111-1111-111111111111}"
                    if (s.Contains("The 'siteId' attribute is invalid"))
                    {
                        // first we need to pull the full text of the siteId attribute
                        string[] split = Regex.Split(scintilla1.Text, @" +");
                        foreach (string sp in split)
                        {
                            if (sp.StartsWith("siteId=") && sp.Contains('{') == false && sp.Contains('}') == false)
                            {
                                string[] replace = sp.Split('"');
                                string valToReplace = replace[0] + "\"{" + replace[1] + "}\"";
                                scintilla1.Text = scintilla1.Text.Replace(sp, valToReplace);
                            }
                        }
                    }
                }
            }

            return;
        }

        /// <summary>
        /// this function loops through all nodes parsed out from Step 1 in BtnFixCorruptDoc_Click
        /// check each node and add fallback tags only to the list
        /// </summary>
        /// <param name="originalText"></param>
        public static void GetAllNodes(string originalText)
        {
            bool isFallback = false;
            var fallback = new List<string>();

            foreach (string o in corruptNodes)
            {
                if (o == Strings.txtFallbackStart)
                {
                    isFallback = true;
                }

                if (isFallback)
                {
                    fallback.Add(o);
                }

                if (o == Strings.txtFallbackEnd)
                {
                    isFallback = false;
                }
            }

            ParseOutFallbackTags(fallback, originalText);
        }

        /// <summary>
        /// we should only have a list of fallback start tags, end tags and each tag in between
        /// the idea is to combine these start/middle/end tags into a long string
        /// then they can be replaced with an empty string
        /// </summary>
        /// <param name="input"></param>
        /// <param name="originalText"></param>
        public static void ParseOutFallbackTags(List<string> input, string originalText)
        {
            var fallbackTagsAppended = new List<string>();
            StringBuilder sbFallback = new StringBuilder();

            foreach (string o in input)
            {
                switch (o.ToString())
                {
                    case Strings.txtFallbackStart:
                        sbFallback.Append(o);
                        continue;
                    case Strings.txtFallbackEnd:
                        sbFallback.Append(o);
                        fallbackTagsAppended.Add(sbFallback.ToString());
                        sbFallback.Clear();
                        continue;
                    default:
                        sbFallback.Append(o);
                        continue;
                }
            }

            sbFallback.Clear();

            // loop each item in the list and remove it from the document
            originalText = fallbackTagsAppended.Aggregate(originalText, (current, o) => current.Replace(o.ToString(), string.Empty));

            // each set of fallback tags should now be removed from the text
            // set it to the global variable so we can add it back into document.xml
            FixedFallback = originalText;
        }

        /// <summary>
        /// append strings like "Fixed" or "Copy" to the file name
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="TextToAdd"></param>
        /// <returns></returns>
        public static string AddTextToFileName(string fileName, string TextToAdd)
        {
            string dir = Path.GetDirectoryName(fileName) + "\\";
            StrExtension = Path.GetExtension(fileName);
            string newFileName = dir + Path.GetFileNameWithoutExtension(fileName) + TextToAdd + StrExtension;
            return newFileName;
        }

        public static void DeleteTempFiles()
        {
            try
            {
                if (File.Exists(tempFilePackageViewer))
                {
                    File.Delete(tempFilePackageViewer);
                }

                if (File.Exists(tempFileReadOnly))
                {
                    File.Delete(tempFileReadOnly);
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "DeleteTempFiles Error: " + ex.Message);
            }
        }

        /// <summary>
        /// cleanup before the app exits
        /// </summary>
        public void AppExitWork()
        {
            try
            {
                package?.Close();
                if (Properties.Settings.Default.DeleteCopiesOnExit == true && File.Exists(StrCopiedFileName))
                {
                    File.Delete(StrCopiedFileName);
                }

                Properties.Settings.Default.Save();
                DeleteTempFiles();
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "App Exit Error: " + ex.Message);
            }
            finally
            {
                Application.Exit();
            }
        }

        /// <summary>
        /// different scenarios call for different icons / UI to light up
        /// </summary>
        public void UpdateAppUI()
        {
            if (toolStripStatusLabelDocType.Text == Strings.oAppWord)
            {
                UpdateUI(UIType.OpenWord);
            }
            else if (toolStripStatusLabelDocType.Text == Strings.oAppExcel)
            {
                UpdateUI(UIType.OpenExcel);
            }
            else if (toolStripStatusLabelDocType.Text == Strings.oAppPowerPoint)
            {
                UpdateUI(UIType.OpenPowerPoint);
            }
            else if (toolStripStatusLabelDocType.Text == Strings.oAppOutlook)
            {
                UpdateUI(UIType.OpenMsg);
            }
            else if (isEncrypted)
            {
                UpdateUI(UIType.OpenCF);
            }
        }

        #endregion

        #region Events

        private void Scintilla1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.F)
            {
                myFindReplace.ShowFind();
                e.SuppressKeyPress = true;
            }
            else if (e.Shift && e.KeyCode == Keys.F3)
            {
                myFindReplace.Window.FindPrevious();
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.F3)
            {
                myFindReplace.Window.FindNext();
                e.SuppressKeyPress = true;
            }
            else if (e.Control && e.KeyCode == Keys.H)
            {
                myFindReplace.ShowReplace();
                e.SuppressKeyPress = true;
            }
            else if (e.Control && e.KeyCode == Keys.I)
            {
                myFindReplace.ShowIncrementalSearch();
                e.SuppressKeyPress = true;
            }
            else if (e.Control && e.KeyCode == Keys.G)
            {
                GoTo MyGoTo = new GoTo((Scintilla)sender);
                MyGoTo.ShowGoToDialog();
                e.SuppressKeyPress = true;
            }
        }

        private void MyFindReplace_KeyPressed(object sender, KeyEventArgs e)
        {
            Scintilla1_KeyDown(sender, e);
        }

        private void OpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenOfficeDocument(false);

            if (toolStripStatusLabelFilePath.Text != Strings.wHeadingBegin)
            {
                AddFileToMRU();
            }

            UpdateAppUI();
        }

        private void SettingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmSettings form = new FrmSettings();
            form.ShowDialog();
        }

        private void BatchFileProcessingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmBatch bFrm = new FrmBatch(package)
            {
                Owner = this
            };
            bFrm.ShowDialog();
        }

        private void FrmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            AppExitWork();
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AppExitWork();
        }

        private void FeedbackToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AppUtilities.PlatformSpecificProcessStart(Strings.helpLocation);
        }

        private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmAbout frm = new FrmAbout();
            frm.ShowDialog(this);
            frm.Dispose();
        }

        public List<string> CustomDocPropsList(CustomFilePropertiesPart cfp)
        {
            List<string> tempCfp = new List<string>();

            if (cfp is null)
            {
                LogInformation(LogInfoType.EmptyCount, Strings.wCustomDocProps, string.Empty);
                return tempCfp;
            }

            int count = 0;

            foreach (string v in Office.CfpList(cfp))
            {
                count++;
                tempCfp.Add(count + Strings.wPeriod + v);
            }

            if (count == 0)
            {
                LogInformation(LogInfoType.EmptyCount, Strings.wCustomDocProps, string.Empty);
            }

            return tempCfp;
        }

        private void ClipboardViewerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmClipboardViewer cFrm = new FrmClipboardViewer()
            {
                Owner = this
            };
            cFrm.ShowDialog();
        }

        private void CopySelectedLineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                scintilla1.Copy();
            }
            catch (Exception ex)
            {
                LogInformation(LogInfoType.LogException, "BtnCopyLineOutput Error", ex.Message);
            }
        }

        private void CopyAllLinesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyAllItems();
        }

        private void Base64DecoderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmBase64 b64Frm = new FrmBase64()
            {
                Owner = this
            };
            b64Frm.ShowDialog();
        }

        private void excelSheetViewerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var f = new FrmSheetViewer(tempFileReadOnly))
            {
                var result = f.ShowDialog();
            }
        }

        public void EnableScintillaAddText()
        {
            scintilla1.ClearAll();
            scintilla1.ReadOnly = false;
        }

        public virtual void MoveFormAwayFromSelection()
        {
            if (!Visible || scintilla1 == null) return;

            int pos = scintilla1.CurrentPosition;
            int x = scintilla1.PointXFromPosition(pos);
            int y = scintilla1.PointYFromPosition(pos);

            Point cursorPoint = new Point(x, y);
            Rectangle r = new Rectangle(Location, Size);

            if (scintilla1 != null)
            {
                r.Location = new Point(scintilla1.ClientRectangle.Right - Size.Width, 0);
            }

            if (r.Contains(cursorPoint))
            {
                Point newLocation;
                if (cursorPoint.Y < (Screen.PrimaryScreen.Bounds.Height / 2))
                {
                    //TODO - replace lineheight with ScintillaNET command, when added
                    int SCI_TEXTHEIGHT = 2279;
                    int lineHeight = scintilla1.DirectMessage(SCI_TEXTHEIGHT, IntPtr.Zero, IntPtr.Zero).ToInt32();
                    newLocation = new Point(r.X, cursorPoint.Y + lineHeight * 2); // Top half of the screen
                }
                else
                {
                    //TODO - replace lineheight with ScintillaNET command, when added
                    int SCI_TEXTHEIGHT = 2279;
                    int lineHeight = scintilla1.DirectMessage(SCI_TEXTHEIGHT, IntPtr.Zero, IntPtr.Zero).ToInt32();
                    newLocation = new Point(r.X, cursorPoint.Y - Height - (lineHeight * 2)); // Bottom half of the screen
                }

                Location = newLocation;
            }
            else
            {
                Location = r.Location;
            }
        }

        /// <summary>
        /// update toolbar and icons
        /// </summary>
        /// <param name="s"></param>
        public void UpdateUI(UIType type)
        {
            switch (type)
            {
                case UIType.OpenWord:
                    DisableAllUI();
                    toolStripButtonViewContents.Enabled = true;
                    toolStripButtonFixDoc.Enabled = true;
                    editToolStripMenuItemModifyContents.Enabled = true;
                    editToolStripMenuItemRemoveCustomDocProps.Enabled = true;
                    editToolStripMenuItemRemoveCustomXml.Enabled = true;
                    wordDocumentRevisionsToolStripMenuItem.Enabled = true;
                    toolStripDropDownButtonInsert.Enabled = true;
                    fileToolStripMenuItemClose.Enabled = true;
                    break;
                case UIType.OpenExcel:
                    DisableAllUI();
                    toolStripButtonViewContents.Enabled = true;
                    toolStripButtonFixDoc.Enabled = true;
                    excelSheetViewerToolStripMenuItem.Enabled = true;
                    editToolStripMenuItemModifyContents.Enabled = true;
                    editToolStripMenuItemRemoveCustomDocProps.Enabled = true;
                    editToolStripMenuItemRemoveCustomXml.Enabled = true;
                    toolStripDropDownButtonInsert.Enabled = true;
                    fileToolStripMenuItemClose.Enabled = true;
                    break;
                case UIType.OpenPowerPoint:
                    DisableAllUI();
                    toolStripButtonViewContents.Enabled = true;
                    toolStripButtonFixDoc.Enabled = true;
                    editToolStripMenuItemModifyContents.Enabled = true;
                    editToolStripMenuItemRemoveCustomDocProps.Enabled = true;
                    editToolStripMenuItemRemoveCustomXml.Enabled = true;
                    toolStripDropDownButtonInsert.Enabled = true;
                    fileToolStripMenuItemClose.Enabled = true;
                    break;
                case UIType.OpenMsg:
                    DisableAllUI();
                    fileToolStripMenuItemClose.Enabled = true;
                    break;
                case UIType.OpenCF:
                    DisableAllUI();
                    fileToolStripMenuItemClose.Enabled = true;
                    break;
                case UIType.ViewXml:
                    DisableCustomUI();
                    toolStripButtonModify.Enabled = true;
                    toolStripButtonValidateXml.Enabled = true;
                    break;
                case UIType.ViewBinary:
                case UIType.ViewImage:
                    DisableCustomUI();
                    break;
                case UIType.ViewLabelInfo:
                    DisableCustomUI();
                    toolStripButtonFixXml.Enabled = true;
                    toolStripButtonValidateXml.Enabled = true;
                    break;
                case UIType.ViewCustomUI:
                    toolStripButtonModify.Enabled = true;
                    toolStripButtonGenerateCallback.Enabled = true;
                    toolStripDropDownButtonInsert.Enabled = true;
                    toolStripButtonInsertIcon.Enabled = true;
                    toolStripButtonValidateXml.Enabled = true;
                    break;
                default:
                    break;
            }
        }

        public void DisableCustomUI()
        {
            toolStripButtonGenerateCallback.Enabled = false;
            toolStripDropDownButtonInsert.Enabled = false;
            toolStripButtonInsertIcon.Enabled = false;
        }

        public void DisableAllUI()
        {
            toolStripButtonViewContents.Enabled = false;
            toolStripButtonFixCorruptDoc.Enabled = false;
            toolStripButtonFixDoc.Enabled = false;
            editToolStripMenuFindReplace.Enabled = false;
            editToolStripMenuItemModifyContents.Enabled = false;
            editToolStripMenuItemRemoveCustomDocProps.Enabled = false;
            editToolStripMenuItemRemoveCustomXml.Enabled = false;
            wordDocumentRevisionsToolStripMenuItem.Enabled = false;
            toolStripButtonModify.Enabled = false;
            toolStripButtonSave.Enabled = false;
            toolStripButtonGenerateCallback.Enabled = false;
            toolStripDropDownButtonInsert.Enabled = false;
            toolStripButtonInsertIcon.Enabled = false;
            toolStripButtonFixXml.Enabled = false;
            toolStripButtonValidateXml.Enabled = false;
            excelSheetViewerToolStripMenuItem.Enabled = false;
            scintilla1.ReadOnly = true;
            scintilla1.BackColor = SystemColors.Control;

            if (package is null)
            {
                fileToolStripMenuItemClose.Enabled = false;
            }
        }

        public void EnableModifyUI()
        {
            scintilla1.ReadOnly = false;
            scintilla1.BackColor = SystemColors.Window;
            toolStripButtonSave.Enabled = true;
        }

        public void DisableModifyUI()
        {
            scintilla1.ReadOnly = true;
            scintilla1.BackColor = SystemColors.Control;
            toolStripButtonSave.Enabled = false;
        }

        private void ShowError(string errorText)
        {
            MessageBox.Show(this, errorText, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public void UpdateStreamDisplay(TreeNode tn)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                if (tn is not null)
                {
                    TvFiles.SelectedNode = tn;
                    // The tag property contains the underlying CFItem.
                    // CFItem target = (CFItem)n.Tag;
                    cfStream = tn.Tag as CFStream;
                    if (cfStream is not null)
                    {
                        byte[] buffer = new byte[cfStream.Size];
                        cfStream.Read(buffer, 0, buffer.Length);

                        StringBuilder sb = new StringBuilder();
                        foreach (byte b in buffer)
                        {
                            if (b != 0)
                            {
                                sb.Append(AppUtilities.ConvertByteToText(b.ToString()));
                            }
                        }

                        scintilla1.Text = sb.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "NodeMouseClick Fail", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        /// <summary>
        /// validate the labelinfo.xml file
        /// </summary>
        public void ValidatePartXml()
        {
            isValidXml = true;

            if (scintilla1.Text == null || scintilla1.Text.Length == 0)
            {
                return;
            }

            try
            {
                ValidationEventHandler eventHandler = new ValidationEventHandler(ValidationEventHandler);
                XmlSchemaSet schema = new XmlSchemaSet();
                XmlTextReader xtr = new XmlTextReader(@".\Schemas\LabelInfo.xsd");
                XmlSchema sch = XmlSchema.Read(xtr, ValidationEventHandler);
                schema.Add(sch);

                var settings = new XmlReaderSettings();
                settings.Schemas.Add(sch);
                settings.ValidationType = ValidationType.Schema;
                settings.ValidationFlags |= XmlSchemaValidationFlags.ProcessInlineSchema;
                settings.ValidationFlags |= XmlSchemaValidationFlags.ReportValidationWarnings;
                settings.ValidationEventHandler += new ValidationEventHandler(ValidationEventHandler);

                using (TextReader textReader = new StringReader(scintilla1.Text))
                {
                    XmlReader rd = XmlReader.Create(textReader, settings);
                    XDocument doc = XDocument.Load(rd);
                    doc.Validate(schema, eventHandler);
                }
            }
            catch (Exception ex)
            {
                // if there were xml validation errors, display a message with those details
                FileUtilities.WriteToLog(Strings.fLogFilePath, ex.Message);
            }

            if (isValidXml)
            {
                MessageBox.Show("Xml Valid", "Xml Validation", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// validate customui.xml
        /// </summary>
        /// <param name="showValidMessage"></param>
        /// <returns></returns>
        public bool ValidateCustomUIXml(bool showValidMessage)
        {
            if (scintilla1.Text == null || scintilla1.Text.Length == 0)
            {
                return false;
            }

            scintilla1.SuspendLayout();

            try
            {
                ValidationEventHandler eventHandler = new ValidationEventHandler(ValidationEventHandler);
                XmlTextReader xtr = new XmlTextReader(@".\Schemas\customui14.xsd");
                XmlSchema sch = XmlSchema.Read(xtr, ValidationEventHandler);
                XmlSchemaSet schema = new XmlSchemaSet();
                schema.Add(sch);

                var settings = new XmlReaderSettings();
                settings.Schemas.Add(sch);
                settings.ValidationType = ValidationType.Schema;
                settings.ValidationFlags |= XmlSchemaValidationFlags.ProcessInlineSchema;
                settings.ValidationFlags |= XmlSchemaValidationFlags.ReportValidationWarnings;
                settings.ValidationEventHandler += new ValidationEventHandler(ValidationEventHandler);

                using (TextReader textReader = new StringReader(scintilla1.Text))
                {
                    XmlReader rd = XmlReader.Create(textReader, settings);
                    XDocument doc = XDocument.Load(rd);
                    doc.Validate(schema, eventHandler);
                }

                isValidXml = false;
            }
            catch (XmlException ex)
            {
                ShowError("Invalid Xml" + "\n" + ex.Message);
                isValidXml = false;
                return false;
            }

            scintilla1.ResumeLayout();

            if (isValidXml)
            {
                if (showValidMessage)
                {
                    MessageBox.Show(this, "Valid Xml", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                return true;
            }
            return false;
        }

        /// <summary>
        /// display schema validation errors
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ValidationEventHandler(object sender, ValidationEventArgs e)
        {
            if (e.Severity == XmlSeverityType.Error)
            {
                isValidXml = false;
            }

            MessageBox.Show("Error at Line #" + e.Exception.LineNumber + " Position #" + e.Exception.LinePosition + Strings.wColonBuffer + e.Message,
                        "Xml Validation", MessageBoxButtons.OK, MessageBoxIcon.Error);
            FileUtilities.WriteToLog(Strings.fLogFilePath, "Xml Validation Error at Line #" + e.Exception.LineNumber + " Position #"
                + e.Exception.LinePosition + Strings.wColonBuffer + e.Message);
        }

        private void AddPart(XMLParts partType)
        {
            OfficePart newPart = CreateCustomUIPart(partType);
            TreeNode partNode = ConstructPartNode(newPart);
            TreeNode currentNode = TvFiles.Nodes[0];
            if (currentNode == null) return;

            TvFiles.SuspendLayout();
            currentNode.Nodes.Add(partNode);
            scintilla1.Text = string.Empty;
            TvFiles.SelectedNode = partNode;
            TvFiles.ResumeLayout();

            // refresh the treeview
            string prevPath = toolStripStatusLabelFilePath.Text;
            FileClose();
            toolStripStatusLabelFilePath.Text = prevPath;
            LoadPartsIntoViewer();
            toolStripButtonModify.Enabled = true;
            fileToolStripMenuItemClose.Enabled = true;
        }

        private static TreeNode ConstructPartNode(OfficePart part)
        {
            TreeNode node = new TreeNode(part.Name);
            node.Tag = part.PartType;
            node.ImageIndex = 3;
            node.SelectedImageIndex = 3;
            return node;
        }

        private OfficePart RetrieveCustomPart(XMLParts partType)
        {
            if (pParts == null || pParts.Count == 0) return null;

            OfficePart oPart;

            foreach (PackagePart pp in pkgParts)
            {
                if (pp.Uri.ToString().EndsWith(Strings.offCustomUI14Xml))
                {
                    return oPart = new OfficePart(pp, XMLParts.RibbonX14, Strings.CustomUI14PartRelType);
                }
                else if (pp.Uri.ToString().EndsWith(Strings.offCustomUIXml))
                {
                    return oPart = new OfficePart(pp, XMLParts.RibbonX14, Strings.CustomUIPartRelType);
                }
            }

            return null;
        }

        private OfficePart CreateCustomUIPart(XMLParts partType)
        {
            string relativePath;
            string relType;

            switch (partType)
            {
                case XMLParts.RibbonX12:
                    relativePath = "/customUI/customUI.xml";
                    relType = Strings.CustomUIPartRelType;
                    break;
                case XMLParts.RibbonX14:
                    relativePath = "/customUI/customUI14.xml";
                    relType = Strings.CustomUI14PartRelType;
                    break;
                case XMLParts.QAT12:
                    relativePath = "/customUI/qat.xml";
                    relType = Strings.QATPartRelType;
                    break;
                default:
                    return null;
            }

            Uri customUIUri = new Uri(relativePath, UriKind.Relative);
            PackageRelationship relationship = package.CreateRelationship(customUIUri, TargetMode.Internal, relType);

            OfficePart part;
            if (!package.PartExists(customUIUri))
            {
                part = new OfficePart(package.CreatePart(customUIUri, "application/xml"), partType, relationship.Id);
            }
            else
            {
                part = new OfficePart(package.GetPart(customUIUri), partType, relationship.Id);
            }

            return part;
        }

        /// <summary>
        /// open the selected file from the MRU
        /// </summary>
        /// <param name="path"></param>
        public void OpenRecentFile(string path)
        {
            if (path == string.Empty || path == Strings.wEmpty)
            {
                return;
            }
            else
            {
                toolStripStatusLabelFilePath.Text = path;
                OpenOfficeDocument(true);
                UpdateAppUI();
            }
        }

        public void ClearRecentMenuItems()
        {
            mruToolStripMenuItem1.Text = Strings.wEmpty;
            mruToolStripMenuItem2.Text = Strings.wEmpty;
            mruToolStripMenuItem3.Text = Strings.wEmpty;
            mruToolStripMenuItem4.Text = Strings.wEmpty;
            mruToolStripMenuItem5.Text = Strings.wEmpty;
            mruToolStripMenuItem6.Text = Strings.wEmpty;
            mruToolStripMenuItem7.Text = Strings.wEmpty;
            mruToolStripMenuItem8.Text = Strings.wEmpty;
            mruToolStripMenuItem9.Text = Strings.wEmpty;
        }

        /// <summary>
        /// displaying xml happens in this selection event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TvFiles_AfterSelect(object sender, TreeViewEventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                scintilla1.ReadOnly = false;

                // render msg content
                if (toolStripStatusLabelFilePath.Text.EndsWith(Strings.msgFileExt))
                {
                    string[] body = e.Node.Tag as string[];

                    if (body != null)
                    {
                        // handle body click
                        if (e.Node.Text.Contains("Body:"))
                        {
                            if (Properties.Settings.Default.MsgAsRtf)
                            {
                                scintilla1.Text = body[1];
                                renderRTFToolStripMenuItem.Enabled = true;
                            }
                            else
                            {
                                scintilla1.Text = body[0];
                            }
                        }
                    }

                    // handle attachments
                    if (e.Node.Parent is not null && e.Node.Parent.Text.Contains("Attachments:"))
                    {
                        foreach (var att in attachmentList)
                        {
                            if (e.Node.Text.Contains(att.Key))
                            {
                                using (var f = new FrmDisplayOutput(att.Value))
                                {
                                    var result = f.ShowDialog();
                                }
                            }
                        }
                    }

                    ScrollToTopOfRtb();
                    TvFiles.ExpandAll();
                    return;
                }

                if (isEncrypted || e.Node.Text.Contains("LabelInfo.xml"))
                {
                    foreach (PackagePart pp in pkgParts)
                    {
                        if (pp.Uri.ToString() == TvFiles.SelectedNode.Text)
                        {
                            Stream stream = pp.GetStream();
                            byte[] binData = FileUtilities.ReadToEnd(stream);
                            scintilla1.Text = Convert.ToHexString(binData);
                            stream.Close();
                            return;
                        }
                    }
                    FormatXmlColors();
                    toolStripButtonValidateXml.Enabled = true;
                    toolStripButtonFixXml.Enabled = true;
                    toolStripButtonModify.Enabled = true;
                }

                if (FileUtilities.GetFileType(e.Node.Text) == OpenXmlInnerFileTypes.XML)
                {
                    // customui files have additional editing options
                    if (e.Node.Text.EndsWith("customUI.xml") || e.Node.Text.EndsWith("customUI14.xml"))
                    {
                        UpdateUI(UIType.ViewCustomUI);
                    }
                    else
                    {
                        UpdateUI(UIType.ViewXml);
                    }

                    // load file contents
                    foreach (PackagePart pp in pkgParts)
                    {
                        if (pp.Uri.ToString() == TvFiles.SelectedNode.Text)
                        {
                            partPropCompression = pp.CompressionOption.ToString();
                            partPropContentType = pp.ContentType;

                            using (StreamReader sr = new StreamReader(pp.GetStream()))
                            {
                                string contents = sr.ReadToEnd();

                                // load the xml and indented/format xml
                                XmlDocument doc = new XmlDocument();
                                doc.LoadXml(contents);

                                using (MemoryStream ms = new MemoryStream())
                                {
                                    XmlWriterSettings settings;
                                    if (e.Node.Text.EndsWith("customUI.xml") || e.Node.Text.EndsWith("customUI14.xml"))
                                    {
                                        // custom ui files need to be saved without the xml declaration
                                        settings = new XmlWriterSettings
                                        {
                                            OmitXmlDeclaration = true,
                                            Indent = true,
                                            IndentChars = "  ",
                                            NewLineChars = "\r\n",
                                            NewLineHandling = NewLineHandling.Replace
                                        };
                                    }
                                    else
                                    {
                                        // all other xml files need to be saved with the utf8 xml declaration
                                        settings = new XmlWriterSettings
                                        {
                                            Encoding = new UTF8Encoding(false),
                                            Indent = true,
                                            IndentChars = "  ",
                                            NewLineChars = "\r\n",
                                            NewLineHandling = NewLineHandling.Replace
                                        };
                                    }

                                    // write out the xml to a memory stream
                                    using (XmlWriter writer = XmlWriter.Create(ms, settings))
                                    {
                                        doc.Save(writer);
                                    }
                                    contents = Encoding.UTF8.GetString(ms.ToArray());
                                }

                                TvFiles.SuspendLayout();
                                scintilla1.ReadOnly = false;
                                scintilla1.Text = contents;
                                FormatXmlColors();
                                TvFiles.ResumeLayout();
                                ScrollToTopOfRtb();
                                toolStripButtonValidateXml.Enabled = true;
                                toolStripButtonFixXml.Enabled = true;
                                toolStripButtonModify.Enabled = true;
                                toolStripButtonSave.Enabled = true;
                                return;
                            }
                        }
                    }
                }
                else if (FileUtilities.GetFileType(e.Node.Text) == OpenXmlInnerFileTypes.Image)
                {
                    // currently showing images with a form
                    // TODO find a way to keep the image in the main form
                    foreach (PackagePart pp in pkgParts)
                    {
                        if (pp.Uri.ToString() == TvFiles.SelectedNode.Text)
                        {
                            partPropCompression = pp.CompressionOption.ToString();
                            partPropContentType = pp.ContentType;

                            // need to implement non-bitmap images
                            if (pp.Uri.ToString().EndsWith(Strings.emfFileExt) || (pp.Uri.ToString().EndsWith(Strings.svgFileExt)))
                            {
                                scintilla1.Text = "No Viewer For File Type";
                                return;
                            }

                            Stream imageSource = pp.GetStream();
                            Image image = Image.FromStream(imageSource);
                            using (var f = new FrmDisplayOutput(image))
                            {
                                var result = f.ShowDialog();
                            }
                            imageSource.Close();
                            return;
                        }
                    }
                }
                else if (FileUtilities.GetFileType(e.Node.Text) == OpenXmlInnerFileTypes.Binary)
                {
                    foreach (PackagePart pp in pkgParts)
                    {
                        if (pp.Uri.ToString() == TvFiles.SelectedNode.Text)
                        {
                            partPropCompression = pp.CompressionOption.ToString();
                            partPropContentType = pp.ContentType;

                            Stream stream = pp.GetStream();
                            byte[] binData = FileUtilities.ReadToEnd(stream);
                            scintilla1.Text = Convert.ToHexString(binData);
                            stream.Close();
                            return;
                        }
                    }
                }
                else
                {
                    scintilla1.Text = "No Viewer For File Type";
                }
            }
            catch (Exception ex)
            {
                scintilla1.Text = "Error: " + ex.Message;
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        public void ScrollToTopOfRtb()
        {
            scintilla1.SelectionStart = 0;
            scintilla1.ScrollCaret();
        }

        /// <summary>
        /// format the xml tags with different colors to make it easier to read
        /// </summary>
        public void FormatXmlColors()
        {
            // Reset the styles
            scintilla1.StyleResetDefault();
            scintilla1.StyleClearAll();

            // custom settings
            scintilla1.TabIndents = true;
            scintilla1.IndentationGuides = IndentView.LookBoth;

            // Set the XML Lexer
            scintilla1.LexerName = "xml";

            // add line numbers, use line count to determine margin width
            if (scintilla1.Lines.Count < 100)
            {
                scintilla1.Margins[0].Width = 20;
            }
            else if (scintilla1.Lines.Count < 1000)
            {
                scintilla1.Margins[0].Width = 30;
            }
            else if (scintilla1.Lines.Count < 10000)
            {
                scintilla1.Margins[0].Width = 35;
            }
            else if (scintilla1.Lines.Count < 100000)
            {
                scintilla1.Margins[0].Width = 40;
            }
            else
            {
                scintilla1.Margins[0].Width = 50;
            }

            // Enable folding
            scintilla1.SetProperty("fold", "1");
            scintilla1.SetProperty("fold.compact", "1");
            scintilla1.SetProperty("fold.html", "1");

            // Use Margin 2 for fold markers
            scintilla1.Margins[2].Type = MarginType.Symbol;
            scintilla1.Margins[2].Mask = Marker.MaskFolders;
            scintilla1.Margins[2].Sensitive = true;
            scintilla1.Margins[2].Width = 20;

            // Reset folder markers
            for (int i = Marker.FolderEnd; i <= Marker.FolderOpen; i++)
            {
                scintilla1.Markers[i].SetForeColor(SystemColors.ControlLightLight);
                scintilla1.Markers[i].SetBackColor(SystemColors.ControlDark);
            }

            // Style the folder markers
            scintilla1.Markers[Marker.Folder].Symbol = MarkerSymbol.BoxPlus;
            scintilla1.Markers[Marker.Folder].SetBackColor(SystemColors.ControlText);
            scintilla1.Markers[Marker.FolderOpen].Symbol = MarkerSymbol.BoxMinus;
            scintilla1.Markers[Marker.FolderEnd].Symbol = MarkerSymbol.BoxPlusConnected;
            scintilla1.Markers[Marker.FolderEnd].SetBackColor(SystemColors.ControlText);
            scintilla1.Markers[Marker.FolderMidTail].Symbol = MarkerSymbol.TCorner;
            scintilla1.Markers[Marker.FolderOpenMid].Symbol = MarkerSymbol.BoxMinusConnected;
            scintilla1.Markers[Marker.FolderSub].Symbol = MarkerSymbol.VLine;
            scintilla1.Markers[Marker.FolderTail].Symbol = MarkerSymbol.LCorner;

            // Enable automatic folding
            scintilla1.AutomaticFold = AutomaticFold.Show | AutomaticFold.Click | AutomaticFold.Change;

            // Set the Styles
            scintilla1.Styles[Style.Xml.Attribute].ForeColor = Color.Red;
            scintilla1.Styles[Style.Xml.Entity].ForeColor = Color.Red;
            scintilla1.Styles[Style.Xml.Comment].ForeColor = Color.Green;
            scintilla1.Styles[Style.Xml.Tag].ForeColor = Color.Blue;
            scintilla1.Styles[Style.Xml.TagEnd].ForeColor = Color.Blue;
            scintilla1.Styles[Style.Xml.DoubleString].ForeColor = Color.DeepPink;
            scintilla1.Styles[Style.Xml.SingleString].ForeColor = Color.DeepPink;
            scintilla1.Styles[Style.Xml.XmlStart].ForeColor = Color.Blue;
            scintilla1.Styles[Style.Xml.TagEnd].ForeColor = Color.Blue;
        }

        private void ToolStripButtonValidateXml_Click(object sender, EventArgs e)
        {
            if (TvFiles.SelectedNode is null)
            {
                return;
            }

            if (TvFiles.SelectedNode.Text.EndsWith(Strings.offLabelInfo))
            {
                ValidatePartXml();
            }
            else if (TvFiles.SelectedNode.Text.EndsWith(Strings.offCustomUI14Xml) || TvFiles.SelectedNode.Text.EndsWith(Strings.offCustomUIXml))
            {
                ValidateCustomUIXml(true);
            }
            else if (isEncrypted || TvFiles.SelectedNode.Text.Contains("docMetadata/LabelInfo.xml"))
            {
                ValidateLabelInfoXml(true);
            }
        }

        private void ToolStripButtonGenerateCallback_Click(object sender, EventArgs e)
        {
            // if there is no callback, then there is no point in generating the callback code
            if (scintilla1.Text == null || scintilla1.Text.Length == 0)
            {
                return;
            }

            // if the xml is not valid, then there is no point in generating the callback code
            if (!ValidateCustomUIXml(false))
            {
                return;
            }

            // generate the callback code
            try
            {
                XmlDocument customUI = new XmlDocument();
                customUI.LoadXml(scintilla1.Text);
                StringBuilder callbacks = CallbackBuilder.GenerateCallback(customUI);
                callbacks.Append('}');

                // display the callbacks
                using (var f = new FrmDisplayOutput(callbacks, true))
                {
                    f.Text = "VBA Callback Code";
                    var result = f.ShowDialog();
                }

                if (callbacks == null || callbacks.Length == 0)
                {
                    MessageBox.Show(this, "No callbacks found", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Office2010CustomUIPartToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddPart(XMLParts.RibbonX14);
            scintilla1.Text = string.Empty;
            TvFiles.SelectedNode = TvFiles.Nodes[0].Nodes[0];
        }

        private void Office2007CustomUIPartToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddPart(XMLParts.RibbonX12);
            scintilla1.Text = string.Empty;
            TvFiles.SelectedNode = TvFiles.Nodes[0].Nodes[0];
        }

        private void CustomOutspaceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            scintilla1.Text = Strings.xmlCustomOutspace;
            FormatXmlColors();
            EnableModifyUI();
        }

        private void CustomTabToolStripMenuItem_Click(object sender, EventArgs e)
        {
            scintilla1.Text = Strings.xmlCustomTab;
            FormatXmlColors();
            EnableModifyUI();
        }

        private void ExcelCustomTabToolStripMenuItem_Click(object sender, EventArgs e)
        {
            scintilla1.Text = Strings.xmlExcelCustomTab;
            FormatXmlColors();
            EnableModifyUI();
        }

        private void RepurposeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            scintilla1.Text = Strings.xmlRepurpose;
            FormatXmlColors();
            EnableModifyUI();
        }

        private void WordGroupOnInsertTabToolStripMenuItem_Click(object sender, EventArgs e)
        {
            scintilla1.Text = Strings.xmlWordGroupInsertTab;
            FormatXmlColors();
            EnableModifyUI();
        }

        private void ToolStripButtonInsertIcon_Click(object sender, EventArgs e)
        {
            OpenFileDialog fDialog = new OpenFileDialog
            {
                Title = "Insert Custom Icon",
                Filter = "Supported Icons | *.ico; *.bmp; *.png; *.jpg; *.jpeg; *.tif;| All Files | *.*;",
                RestoreDirectory = true,
                InitialDirectory = @"%userprofile%"
            };

            if (fDialog.ShowDialog() == DialogResult.OK)
            {
                XMLParts partType = XMLParts.RibbonX14;
                OfficePart part = RetrieveCustomPart(partType);
                foreach (TreeNode node in TvFiles.Nodes[0].Nodes)
                {
                    if (node.Text == part.Name)
                    {
                        break;
                    }
                }

                TvFiles.SuspendLayout();

                foreach (string fileName in fDialog.FileNames)
                {
                    try
                    {
                        string id = XmlConvert.EncodeName(Path.GetFileNameWithoutExtension(fileName));
                        Image image = Image.FromStream(File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite), true, true);

                        // The file is a valid image at this point.
                        id = part.AddImage(fileName, id);
                        if (id == null) continue;

                        File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite).Close();

                        TreeNode imageNode = new TreeNode(id);
                        imageNode.ImageKey = Strings.chUnderscore + id;
                        imageNode.SelectedImageKey = imageNode.ImageKey;
                        imageNode.Tag = partType;

                        TvFiles.ImageList.Images.Add(imageNode.ImageKey, image);
                        TvFiles.SelectedNode.Nodes.Add(imageNode);
                    }
                    catch (Exception ex)
                    {
                        ShowError(ex.Message);
                        continue;
                    }
                }

                TvFiles.ResumeLayout();
            }
        }

        private void ToolStripButtonModify_Click(object sender, EventArgs e)
        {
            EnableModifyUI();
        }

        private void ToolStripButtonSave_Click(object sender, EventArgs e)
        {
            if (isEncrypted)
            {
                // write the stream changes and save
                cfStream.Write(Encoding.Default.GetBytes(scintilla1.Text), 0, 0, Encoding.Default.GetByteCount(scintilla1.Text));
                cf.Commit();

                // let the user know it worked, then close the stream and form
                MessageBox.Show("Stream changes saved.", "File Save", MessageBoxButtons.OK, MessageBoxIcon.Information);
                cf.Close();
                return;
            }

            bool isModified = false;

            foreach (PackagePart pp in pkgParts)
            {
                if (pp.Uri.ToString() == TvFiles.SelectedNode.Text)
                {
                    MemoryStream ms = new MemoryStream();
                    using (StreamWriter sw = new StreamWriter(ms))
                    {
                        sw.Write(scintilla1.Text);
                        sw.Flush();

                        ms.Position = 0;
                        Stream partStream = pp.GetStream(FileMode.OpenOrCreate, FileAccess.Write);
                        partStream.SetLength(0);
                        ms.WriteTo(partStream);
                        isModified = true;
                    }
                    break;
                }
            }

            // update ui
            DisableModifyUI();

            // if the part is modified, save changes and refresh the treeview
            if (isModified)
            {
                package.Flush();
                package.Close();
                pkgParts.Clear();
                TvFiles.Nodes.Clear();
                LoadPartsIntoViewer();
            }

            toolStripButtonSave.Enabled = false;
        }

        private void ToolStripButtonViewContents_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                scintilla1.Clear();
                StringBuilder sb = new StringBuilder();

                // display file contents based on user selection
                if (StrOfficeApp == Strings.oAppWord)
                {
                    sb.Append(DisplayListContents(Word.LstContentControls(tempFileReadOnly), Strings.wContentControls));
                    sb.Append(DisplayListContents(Word.LstTables(tempFileReadOnly), Strings.wTables));
                    sb.Append(DisplayListContents(Word.LstStyles(tempFileReadOnly), Strings.wStyles));
                    sb.Append(DisplayListContents(Word.LstHyperlinks(tempFileReadOnly), Strings.wHyperlinks));
                    sb.Append(DisplayListContents(Word.LstListTemplates(tempFileReadOnly, false), Strings.wListTemplates));
                    sb.Append(DisplayListContents(Word.LstFonts(tempFileReadOnly), Strings.wFonts));
                    sb.Append(DisplayListContents(Word.LstRunFonts(tempFileReadOnly), Strings.wRunFonts));
                    sb.Append(DisplayListContents(Word.LstFootnotes(tempFileReadOnly), Strings.wFootnotes));
                    sb.Append(DisplayListContents(Word.LstEndnotes(tempFileReadOnly), Strings.wEndnotes));
                    sb.Append(DisplayListContents(Word.LstDocProps(tempFileReadOnly), Strings.wDocProps));
                    sb.Append(DisplayListContents(Word.LstBookmarks(tempFileReadOnly), Strings.wBookmarks));
                    sb.Append(DisplayListContents(Word.LstFieldCodes(tempFileReadOnly), Strings.wFldCodes));
                    sb.Append(DisplayListContents(Word.LstFieldCodesInHeader(tempFileReadOnly), " ** Header Field Codes **"));
                    sb.Append(DisplayListContents(Word.LstFieldCodesInFooter(tempFileReadOnly), " ** Footer Field Codes **"));
                }
                else if (StrOfficeApp == Strings.oAppExcel)
                {
                    sb.Append(DisplayListContents(Excel.GetLinks(tempFileReadOnly, true), Strings.wLinks));
                    sb.Append(DisplayListContents(Excel.GetComments(tempFileReadOnly), Strings.wComments));
                    sb.Append(DisplayListContents(Excel.GetHyperlinks(tempFileReadOnly), Strings.wHyperlinks));
                    sb.Append(DisplayListContents(Excel.GetSheetInfo(tempFileReadOnly), Strings.wWorksheetInfo));
                    sb.Append(DisplayListContents(Excel.GetSharedStrings(tempFileReadOnly), Strings.wSharedStrings));
                    sb.Append(DisplayListContents(Excel.GetDefinedNames(tempFileReadOnly), Strings.wDefinedNames));
                    sb.Append(DisplayListContents(Excel.GetConnections(tempFileReadOnly), Strings.wConnections));
                    sb.Append(DisplayListContents(Excel.GetHiddenRowCols(tempFileReadOnly), Strings.wHiddenRowCol));
                }
                else if (StrOfficeApp == Strings.oAppPowerPoint)
                {
                    sb.Append(DisplayListContents(PowerPoint.GetHyperlinks(tempFileReadOnly), Strings.wHyperlinks));
                    sb.Append(DisplayListContents(PowerPoint.GetComments(tempFileReadOnly), Strings.wComments));
                    sb.Append(DisplayListContents(PowerPoint.GetSlideText(tempFileReadOnly), Strings.wSlideText));
                    sb.Append(DisplayListContents(PowerPoint.GetSlideTitles(tempFileReadOnly), Strings.wSlideText));
                    sb.Append(DisplayListContents(PowerPoint.GetSlideTransitions(tempFileReadOnly), Strings.wSlideTransitions));
                    sb.Append(DisplayListContents(PowerPoint.GetFonts(tempFileReadOnly), Strings.wFonts));
                }

                // display selected Office features

                sb.Append(DisplayListContents(Office.GetEmbeddedObjectProperties(tempFileReadOnly, toolStripStatusLabelDocType.Text), Strings.wEmbeddedObjects));
                sb.Append(DisplayListContents(Office.GetShapes(tempFileReadOnly, toolStripStatusLabelDocType.Text), Strings.wShapes));
                sb.Append(DisplayListContents(pParts, Strings.wPackageParts));
                sb.Append(DisplayListContents(Office.GetSignatures(tempFileReadOnly, toolStripStatusLabelDocType.Text), Strings.wXmlSignatures));

                // validate the file and update custom file props
                if (toolStripStatusLabelDocType.Text == Strings.oAppWord)
                {
                    using (WordprocessingDocument myDoc = WordprocessingDocument.Open(tempFileReadOnly, false))
                    {
                        sb.Append(DisplayListContents(CustomDocPropsList(myDoc.CustomFilePropertiesPart), Strings.wCustomDocProps));
                        sb.Append(DisplayListContents(Office.DisplayValidationErrorInformation(myDoc), Strings.errorValidation));
                    }
                }
                else if (toolStripStatusLabelDocType.Text == Strings.oAppExcel)
                {
                    using (SpreadsheetDocument myDoc = SpreadsheetDocument.Open(tempFileReadOnly, false))
                    {
                        sb.Append(DisplayListContents(CustomDocPropsList(myDoc.CustomFilePropertiesPart), Strings.wCustomDocProps));
                        sb.Append(DisplayListContents(Office.DisplayValidationErrorInformation(myDoc), Strings.errorValidation));
                    }
                }
                else if (toolStripStatusLabelDocType.Text == Strings.oAppPowerPoint)
                {
                    using (PresentationDocument myDoc = PresentationDocument.Open(tempFileReadOnly, false))
                    {
                        sb.Append(DisplayListContents(CustomDocPropsList(myDoc.CustomFilePropertiesPart), Strings.wCustomDocProps));
                        sb.Append(DisplayListContents(Office.DisplayValidationErrorInformation(myDoc), Strings.errorValidation));
                    }
                }

                using (var f = new FrmDisplayOutput(sb, false))
                {
                    f.Text = "File Contents";
                    var result = f.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogInfoType.LogException, "ViewContents Error:", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void ToolStripButtonFixCorruptDoc_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                StrDestFileName = AddTextToFileName(toolStripStatusLabelFilePath.Text, Strings.wFixedFileParentheses);
                bool isXmlException = false;
                string strDocText = string.Empty;
                IsFixed = false;
                EnableScintillaAddText();

                if (StrExtension == Strings.docxFileExt)
                {
                    if ((File.GetAttributes(toolStripStatusLabelFilePath.Text) & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                    {
                        scintilla1.Text = "ERROR: File is Read-Only.";
                        return;
                    }
                    else
                    {
                        File.Copy(toolStripStatusLabelFilePath.Text, StrDestFileName, true);
                    }
                }

                // bug in packaging API in .NET Core, need to break this fix into separate using blocks to get around the problem
                // 1. check for the xml corruption in document.xml
                using (Package package = Package.Open(StrDestFileName, FileMode.Open, FileAccess.Read))
                {
                    foreach (PackagePart part in package.GetParts())
                    {
                        if (part.Uri.ToString() == Strings.wdDocumentXml)
                        {
                            try
                            {
                                XmlDocument xdoc = new XmlDocument();
                                xdoc.Load(part.GetStream(FileMode.Open, FileAccess.Read));
                            }
                            catch (XmlException) // invalid xml found, try to fix the contents
                            {
                                isXmlException = true;
                            }
                        }
                    }
                }

                // 2. find any known bad sequences and create a string with those changes
                using (Package package = Package.Open(StrDestFileName, FileMode.Open, FileAccess.Read))
                {
                    if (isXmlException)
                    {
                        foreach (PackagePart part in package.GetParts())
                        {
                            if (part.Uri.ToString() == Strings.wdDocumentXml)
                            {
                                InvalidXmlTags invalid = new InvalidXmlTags();
                                string strDocTextBackup;

                                using (StreamReader sr = new StreamReader(part.GetStream(FileMode.Open, FileAccess.Read)))
                                {
                                    strDocText = sr.ReadToEnd();
                                    strDocTextBackup = strDocText;

                                    foreach (string el in invalid.InvalidTags())
                                    {
                                        foreach (Match m in Regex.Matches(strDocText, el))
                                        {
                                            switch (m.Value)
                                            {
                                                case ValidXmlTags.StrValidMcChoice1:
                                                case ValidXmlTags.StrValidMcChoice2:
                                                case ValidXmlTags.StrValidMcChoice3:
                                                    break;

                                                case InvalidXmlTags.StrInvalidVshape:
                                                    // the original strvalidvshape fixes most corruptions, but there are
                                                    // some that are within a group so I added this for those rare situations
                                                    // where the v:group closing tag needs to be included
                                                    if (Properties.Settings.Default.FixGroupedShapes == true)
                                                    {
                                                        strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidVshapegroup);
                                                        scintilla1.AppendText(Strings.invalidTag + m.Value);
                                                        scintilla1.AppendText(Strings.replacedWith + ValidXmlTags.StrValidVshapegroup);
                                                    }
                                                    else
                                                    {
                                                        strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidVshape);
                                                        scintilla1.AppendText(Strings.invalidTag + m.Value);
                                                        scintilla1.AppendText(Strings.replacedWith + ValidXmlTags.StrValidVshape);
                                                    }
                                                    break;

                                                case InvalidXmlTags.StrInvalidOmathWps:
                                                    strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidomathwps);
                                                    scintilla1.AppendText(Strings.invalidTag + m.Value);
                                                    scintilla1.AppendText(Strings.replacedWith + ValidXmlTags.StrValidomathwps);
                                                    break;

                                                case InvalidXmlTags.StrInvalidOmathWpg:
                                                    strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidomathwpg);
                                                    scintilla1.AppendText(Strings.invalidTag + m.Value);
                                                    scintilla1.AppendText(Strings.replacedWith + ValidXmlTags.StrValidomathwpg);
                                                    break;

                                                case InvalidXmlTags.StrInvalidOmathWpc:
                                                    strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidomathwpc);
                                                    scintilla1.AppendText(Strings.invalidTag + m.Value);
                                                    scintilla1.AppendText(Strings.replacedWith + ValidXmlTags.StrValidomathwpc);
                                                    break;

                                                case InvalidXmlTags.StrInvalidOmathWpi:
                                                    strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidomathwpi);
                                                    scintilla1.AppendText(Strings.invalidTag + m.Value);
                                                    scintilla1.AppendText(Strings.replacedWith + ValidXmlTags.StrValidomathwpi);
                                                    break;

                                                default:
                                                    // default catch for "strInvalidmcChoiceRegEx" and "strInvalidFallbackRegEx"
                                                    // since the exact string will never be the same and always has different trailing tags
                                                    // we need to conditionally check for specific patterns
                                                    // the first if </mc:Choice> is to catch and replace the invalid mc:Choice tags
                                                    if (m.Value.Contains(Strings.txtMcChoiceTagEnd))
                                                    {
                                                        if (m.Value.Contains("<mc:Fallback id="))
                                                        {
                                                            // secondary check for a fallback that has an attribute.
                                                            // we don't allow attributes in a fallback
                                                            strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidMcChoice4);
                                                            scintilla1.AppendText(Strings.invalidTag + m.Value);
                                                            scintilla1.AppendText(Strings.replacedWith + ValidXmlTags.StrValidMcChoice4);
                                                            break;
                                                        }

                                                        // replace mc:choice and hold onto the tag that follows
                                                        strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidMcChoice3 + m.Groups[2].Value);
                                                        scintilla1.AppendText(Strings.invalidTag + m.Value);
                                                        scintilla1.AppendText(Strings.replacedWith + ValidXmlTags.StrValidMcChoice3 + m.Groups[2].Value);
                                                        break;
                                                    }
                                                    // the second if <w:pict/> is to catch and replace the invalid mc:Fallback tags
                                                    else if (m.Value.Contains("<w:pict/>"))
                                                    {
                                                        if (m.Value.Contains(Strings.txtFallbackEnd))
                                                        {
                                                            // if the match contains the closing fallback we just need to remove the entire fallback
                                                            // this will leave the closing AC and Run tags, which should be correct
                                                            strDocText = strDocText.Replace(m.Value, string.Empty);
                                                            scintilla1.AppendText(Strings.invalidTag + m.Value);
                                                            scintilla1.AppendText(Strings.replacedWith + "Fallback tag deleted.");
                                                            break;
                                                        }

                                                        // if there is no closing fallback tag, we can replace the match with the omitFallback valid tags
                                                        // then we need to also add the trailing tag, since it's always different but needs to stay in the file
                                                        strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrOmitFallback + m.Groups[2].Value);
                                                        scintilla1.AppendText(Strings.invalidTag + m.Value);
                                                        scintilla1.AppendText(Strings.replacedWith + ValidXmlTags.StrOmitFallback + m.Groups[2].Value);
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        FileUtilities.WriteToLog(Strings.fLogFilePath, "Corrupt Doc Exception = Unknown scenario");
                                                        break;
                                                    }
                                            }
                                        }
                                    }

                                    // remove all fallback tags is a 3 step process
                                    // Step 1. start by getting a list of all nodes/values in the document.xml file
                                    // Step 2. call GetAllNodes to add each fallback tag
                                    // Step 3. call ParseOutFallbackTags to remove each fallback
                                    if (Properties.Settings.Default.RemoveFallback == true)
                                    {
                                        CharEnumerator charEnum = strDocText.GetEnumerator();
                                        while (charEnum.MoveNext())
                                        {
                                            // keep track of previous char
                                            PrevChar = charEnum.Current;

                                            // opening tag
                                            switch (charEnum.Current)
                                            {
                                                case Strings.chLessThan:
                                                    // if we haven't hit a close, but hit another '<' char
                                                    // we are not a true open tag so add it like a regular char
                                                    if (sbNodeBuffer.Length > 0)
                                                    {
                                                        corruptNodes.Add(sbNodeBuffer.ToString());
                                                        sbNodeBuffer.Clear();
                                                    }
                                                    Node(charEnum.Current);
                                                    break;

                                                case Strings.chGreaterThan:
                                                    // there are 2 ways to close out a tag
                                                    // 1. self contained tag like <w:sz w:val="28"/>
                                                    // 2. standard xml <w:t>test</w:t>
                                                    // if previous char is '/', then we are an end tag
                                                    if (PrevChar == Strings.chBackslash || IsRegularXmlTag)
                                                    {
                                                        Node(charEnum.Current);
                                                        IsRegularXmlTag = false;
                                                    }
                                                    Node(charEnum.Current);
                                                    corruptNodes.Add(sbNodeBuffer.ToString());
                                                    sbNodeBuffer.Clear();
                                                    break;

                                                default:
                                                    // this is the second xml closing style, keep track of char
                                                    if (PrevChar == Strings.chLessThan && charEnum.Current == Strings.chBackslash)
                                                    {
                                                        IsRegularXmlTag = true;
                                                    }
                                                    Node(charEnum.Current);
                                                    break;
                                            }

                                            // cleanup
                                            charEnum.Dispose();
                                        }

                                        GetAllNodes(strDocText);
                                        strDocText = FixedFallback;
                                    }

                                    // if no changes were made, no corruptions were found and we can exit
                                    if (strDocText.Equals(strDocTextBackup))
                                    {
                                        scintilla1.Text = " ## No Corruption Found  ## ";
                                        return;
                                    }
                                }
                            }
                        }
                    }
                }

                // 3. write the part with the changes into the new file
                using (Package package = Package.Open(StrDestFileName, FileMode.Open, FileAccess.ReadWrite))
                {
                    MemoryStream ms = new MemoryStream();
                    using (TextWriter tw = new StreamWriter(ms))
                    {
                        foreach (PackagePart part in package.GetParts())
                        {
                            if (part.Uri.ToString() == Strings.wdDocumentXml)
                            {
                                tw.Write(strDocText);
                                tw.Flush();

                                // write the part
                                ms.Position = 0;
                                Stream partStream = part.GetStream(FileMode.Open, FileAccess.Write);
                                partStream.SetLength(0);
                                ms.WriteTo(partStream);
                                IsFixed = true;
                            }
                        }
                    }
                }
            }
            catch (FileFormatException ffe)
            {
                DisplayInvalidFileFormatError();
                FileUtilities.WriteToLog(Strings.fLogFilePath, "Corrupt Doc Exception = " + ffe.Message);
            }
            catch (Exception ex)
            {
                scintilla1.Text = Strings.errorUnableToFixDocument + ex.Message;
                FileUtilities.WriteToLog(Strings.fLogFilePath, "Corrupt Doc Exception = " + ex.Message);
            }
            finally
            {
                // only delete destination file when there is an error
                // need to make sure the file stays when it is fixed
                if (IsFixed == false)
                {
                    // delete the copied file if it exists
                    if (File.Exists(StrDestFileName))
                    {
                        File.Delete(StrDestFileName);
                    }

                    LogInformation(LogInfoType.EmptyCount, Strings.wInvalidXml, string.Empty);
                }
                else
                {
                    // since we were able to attempt the fixes
                    // check if we can open in the sdk and confirm it was indeed fixed
                    if (OpenWithSdk(StrDestFileName))
                    {
                        scintilla1.Text = Strings.wHeaderLine + Environment.NewLine + "Fixed Document Location: " + StrDestFileName;
                    }
                    else
                    {
                        scintilla1.Text = "Unable to fix document";
                    }
                }

                // reset the globals
                IsFixed = false;
                IsRegularXmlTag = false;
                FixedFallback = string.Empty;
                StrExtension = string.Empty;
                StrDestFileName = string.Empty;
                PrevChar = Strings.chLessThan;
                Cursor = Cursors.Default;
            }
        }

        /// <summary>
        /// run through each known set of document corruptions and fix any that are found
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ToolStripButtonFixDoc_Click(object sender, EventArgs e)
        {
            bool corruptionFound = false;
            EnableScintillaAddText();

            StringBuilder sbFixes = new StringBuilder();

            if (toolStripStatusLabelDocType.Text == Strings.oAppWord)
            {
                if (WordFixes.FixListStyles(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("List Styles Fixed");
                    corruptionFound = true;
                }

                if (WordFixes.FixTextboxes(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("Corrupt Textboxes Fixed");
                    corruptionFound = true;
                }

                if (WordFixes.RemoveMissingBookmarkTags(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("Missing Bookmark Tags Fixed");
                    corruptionFound = true;
                }

                if (WordFixes.RemovePlainTextCcFromBookmark(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("Removed Corrupt Content Controls");
                    corruptionFound = true;
                }

                if (WordFixes.FixBookmarkTagInSdtContent(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("Fixed Bookmark Tags");
                    corruptionFound = true;
                }

                if (WordFixes.FixRevisions(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("Corrupt Revisions Fixed");
                    corruptionFound = true;
                }

                if (WordFixes.FixDeleteRevision(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("Corrupt Delete Revisions Fixed");
                    corruptionFound = true;
                }

                if (WordFixes.FixEndnotes(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("Corrupt Endnotes Fixed");
                    corruptionFound = true;
                }

                if (WordFixes.FixListTemplates(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("Fixed Orphaned List Templates");
                    corruptionFound = true;
                }

                if (WordFixes.FixTableGridProps(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("Corrupt Table Grid Properties Fixed");
                    corruptionFound = true;
                }

                if (WordFixes.FixTableRowCorruption(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("Corrupt Table Rows Fixed");
                    corruptionFound = true;
                }

                if (WordFixes.FixCorruptTableCellTags(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("Corrupt Table Cells Fixed");
                    corruptionFound = true;
                }

                if (WordFixes.FixGridSpan(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("Corrupt Table Grid Span Fixed");
                    corruptionFound = true;
                }

                if (WordFixes.FixMissingCommentRefs(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("Fixed Missing Comment References");
                    corruptionFound = true;
                }

                if (WordFixes.FixShapeInComment(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("Corrupt Shapes Fixed");
                    corruptionFound = true;
                }

                if (WordFixes.FixCommentFieldCodes(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("Corrupt Field Codes Fixed");
                    corruptionFound = true;
                }

                if (WordFixes.FixHyperlinks(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("Corrupt Hyperlinks Fixed");
                    corruptionFound = true;
                }

                if (WordFixes.FixContentControls(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("Corrupt Content Controls Fixed");
                    corruptionFound = true;
                }

                if (WordFixes.FixMathAccents(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("Corrupt Math Accents Fixed");
                    corruptionFound = true;
                }

                if (Properties.Settings.Default.CheckZipItemCorrupt)
                {
                    if (WordFixes.FixDataDescriptor(tempFilePackageViewer))
                    {
                        sbFixes.AppendLine("Corrupt Data Descriptor Fixed");
                        corruptionFound = true;
                    }
                }
            }
            else if (toolStripStatusLabelDocType.Text == Strings.oAppPowerPoint)
            {
                if (PowerPointFixes.FixMissingRelIds(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("Fixed Missing Relationship Ids");
                    corruptionFound = true;
                }

                if (PowerPointFixes.FixMissingPlaceholder(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("Fixed Missing Placeholders");
                    corruptionFound = true;
                }

                if (PowerPointFixes.ResetDefaultParagraphProps(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("Fixed Default Paragraph Properties");
                    corruptionFound = true;
                }
            }
            else if (toolStripStatusLabelDocType.Text == Strings.oAppExcel)
            {
                if (Excel.RemoveCorruptClientDataObjects(tempFilePackageViewer))
                {
                    sbFixes.AppendLine("Corrupt Client Data Objects Fixed");
                    corruptionFound = true;
                }
            }

            // if any corruptions were found, copy the file to a new location and display the fixes and new file path
            if (corruptionFound)
            {
                string modifiedPath = AddTextToFileName(toolStripStatusLabelFilePath.Text, " (Fixed)");
                File.Copy(tempFilePackageViewer, modifiedPath, true);
                scintilla1.Text = sbFixes.ToString();
                scintilla1.AppendText("\r\n\r\nModified File Location = " + modifiedPath);
            }
            else
            {
                scintilla1.AppendText("No Corruption Found.");
            }
            scintilla1.ReadOnly = true;
        }

        private void EditToolStripMenuFindReplace_Click(object sender, EventArgs e)
        {
            FrmSearchReplace srForm = new FrmSearchReplace()
            {
                Owner = this
            };
            srForm.ShowDialog();

            if (string.IsNullOrEmpty(findText) && string.IsNullOrEmpty(replaceText))
            {
                return;
            }

            Office.SearchAndReplace(toolStripStatusLabelFilePath.Text, findText, replaceText);
            LogInformation(LogInfoType.ClearAndAdd, "** Search and Replace Finished **", string.Empty);
        }

        private void EditToolStripMenuItemModifyContents_Click(object sender, EventArgs e)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                scintilla1.Clear();

                if (StrOfficeApp == Strings.oAppWord)
                {
                    using (var f = new FrmWordModify())
                    {
                        DialogResult result = f.ShowDialog();

                        if (result == DialogResult.Cancel)
                        {
                            return;
                        }

                        Cursor = Cursors.WaitCursor;

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.DelHF)
                        {
                            if (Word.RemoveHeadersFooters(tempFilePackageViewer) == true)
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "Headers and Footers Deleted", string.Empty);
                            }
                            else
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "Unable to delete headers and footers", string.Empty);
                            }
                        }

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.DelComments)
                        {
                            if (Word.RemoveComments(tempFilePackageViewer) == true)
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "Comments Deleted", string.Empty);
                            }
                            else
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "Unable to delete comments", string.Empty);
                            }
                        }

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.DelEndnotes)
                        {
                            if (Word.RemoveEndnotes(tempFilePackageViewer) == true)
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "Endnotes Deleted", string.Empty);
                            }
                            else
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "Unable to delete endnotes", string.Empty);
                            }
                        }

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.DelFootnotes)
                        {
                            if (Word.RemoveFootnotes(tempFilePackageViewer) == true)
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "Footnotes Deleted", string.Empty);
                            }
                            else
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "Unable to delete footnotes", string.Empty);
                            }
                        }

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.DelOrphanLT)
                        {
                            oNumIdList = Word.LstListTemplates(tempFilePackageViewer, true);
                            foreach (object orphanLT in oNumIdList)
                            {
                                Word.RemoveListTemplatesNumId(tempFilePackageViewer, orphanLT.ToString());
                            }
                            LogInformation(LogInfoType.ClearAndAdd, "Unused List Templates Removed", string.Empty);
                        }

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.DelOrphanStyles)
                        {
                            DisplayListContents(Word.RemoveUnusedStyles(tempFilePackageViewer), Strings.wStyles);
                        }

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.DelHiddenTxt)
                        {
                            if (Word.DeleteHiddenText(tempFilePackageViewer) == true)
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "Hidden Text Deleted", string.Empty);
                            }
                            else
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "Unable to delete hidden text", string.Empty);
                            }
                        }

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.DelPgBrk)
                        {
                            if (Word.RemoveBreaks(tempFilePackageViewer))
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "Page Breaks Deleted", string.Empty);
                            }
                            else
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "Unable to delete Page Breaks", string.Empty);
                            }
                        }

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.SetPrintOrientation)
                        {
                            FrmPrintOrientation pFrm = new FrmPrintOrientation(tempFilePackageViewer)
                            {
                                Owner = this
                            };
                            pFrm.ShowDialog();
                        }

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.AcceptRevisions)
                        {
                            foreach (var s in Word.AcceptRevisions(tempFilePackageViewer, Strings.allAuthors))
                            {
                                sb.AppendLine(s);
                            }
                        }

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.ChangeDefaultTemplate)
                        {
                            bool isFileChanged = false;
                            string attachedTemplateId = "rId1";
                            string filePath = string.Empty;

                            using (WordprocessingDocument document = WordprocessingDocument.Open(tempFilePackageViewer, true))
                            {
                                DocumentSettingsPart dsp = document.MainDocumentPart.DocumentSettingsPart;

                                // if the external rel exists, we need to pull the rId and old uri
                                // we will be deleting this part and re-adding with the new uri
                                if (dsp.ExternalRelationships.Any())
                                {
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
                                        LogInformation(LogInfoType.InvalidFile, "BtnChangeDefaultTemplate", "Invalid Attached Template Path - " + filePath);
                                        throw new Exception();
                                    }
                                }

                                // get the new template path from the user
                                FrmChangeDefaultTemplate ctFrm = new FrmChangeDefaultTemplate(FileUtilities.ConvertUriToFilePath(filePath))
                                {
                                    Owner = this
                                };
                                ctFrm.ShowDialog();

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

                                    // if we aren't Normal, add a new part back in with the new path
                                    if (fromChangeTemplate != "Normal")
                                    {
                                        Uri newFilePath = new Uri(filePath);
                                        dsp.AddExternalRelationship(Strings.DocumentTemplatePartType, newFilePath, attachedTemplateId);
                                        isFileChanged = true;
                                    }
                                    else
                                    {
                                        // if we are changing to Normal, delete the attachtemplate id ref
                                        foreach (OpenXmlElement oe in dsp.Settings)
                                        {
                                            if (oe.ToString() == Strings.dfowAttachedTemplate)
                                            {
                                                oe.Remove();
                                                isFileChanged = true;
                                            }
                                        }
                                    }
                                }

                                if (isFileChanged)
                                {
                                    sb.AppendLine("** Attached Template Path Changed **");
                                    document.MainDocumentPart.Document.Save();
                                }
                                else
                                {
                                    sb.AppendLine("** No Changes Made To Attached Template **");
                                }
                            }
                        }

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.ConvertDocmToDocx)
                        {
                            string fNewName = Office.ConvertMacroEnabled2NonMacroEnabled(tempFilePackageViewer, Strings.oAppWord);
                            if (fNewName != string.Empty)
                            {
                                sb.AppendLine(tempFilePackageViewer + Strings.convertedTo + fNewName);
                            }
                        }

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.RemovePII)
                        {
                            using (WordprocessingDocument document = WordprocessingDocument.Open(tempFilePackageViewer, true))
                            {
                                if ((Word.HasPersonalInfo(document) == true) && Word.RemovePersonalInfo(document) == true)
                                {
                                    LogInformation(LogInfoType.ClearAndAdd, "PII Removed from file.", string.Empty);
                                }
                                else
                                {
                                    LogInformation(LogInfoType.EmptyCount, Strings.wPII, string.Empty);
                                }
                            }
                        }

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.RemoveCustomTitleProp)
                        {
                            if (Word.RemoveCustomTitleProp(tempFilePackageViewer))
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "Custom Property 'Title' Removed From File.", string.Empty);
                            }
                            else
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "'Title' Property Not Found.", string.Empty);
                            }
                        }

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.UpdateCcNamespaceGuid)
                        {
                            if (WordFixes.FixContentControlNamespaces(tempFilePackageViewer))
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "Quick Part Namespaces Updated", string.Empty);
                            }
                            else
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "No Issues With Namespaces Found.", string.Empty);
                            }
                        }

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.DelBookmarks)
                        {
                            if (Word.RemoveBookmarks(tempFilePackageViewer))
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "Bookmarks Deleted", string.Empty);
                            }
                            else
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "No Bookmarks In Document", string.Empty);
                            }
                        }

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.DelDupeAuthors)
                        {
                            Dictionary<string, string> authors = new Dictionary<string, string>();

                            using (WordprocessingDocument document = WordprocessingDocument.Open(tempFilePackageViewer, true))
                            {
                                // check the peoplepart and list those authors
                                WordprocessingPeoplePart peoplePart = document.MainDocumentPart.WordprocessingPeoplePart;
                                if (peoplePart != null)
                                {
                                    foreach (Person person in peoplePart.People)
                                    {
                                        authors.Add(person.Author, person.PresenceInfo.UserId);
                                    }
                                }
                            }

                            using (var fDupe = new FrmDuplicateAuthors(authors, tempFilePackageViewer))
                            {
                                fDupe.ShowDialog();
                                if (fDupe.dr == DialogResult.OK)
                                {
                                    result = DialogResult.OK;
                                }
                                else
                                {
                                    result = DialogResult.Cancel;
                                }
                            }
                        }

                        if (result == DialogResult.OK)
                        {
                            string modifiedPath = AddModifiedTextToFileName(toolStripStatusLabelFilePath.Text);
                            File.Copy(tempFilePackageViewer, modifiedPath, true);
                        }
                    }
                }
                else if (StrOfficeApp == Strings.oAppExcel)
                {
                    using (var f = new FrmExcelModify())
                    {
                        DialogResult result = f.ShowDialog();

                        if (result == DialogResult.Cancel)
                        {
                            return;
                        }

                        Cursor = Cursors.WaitCursor;

                        if (f.xlModCmd == AppUtilities.ExcelModifyCmds.DelLink)
                        {
                            using (var fDelLink = new FrmExcelDelLink(tempFileReadOnly))
                            {
                                if (fDelLink.fHasLinks)
                                {
                                    fDelLink.ShowDialog();
                                    if (fDelLink.DialogResult == DialogResult.OK)
                                    {
                                        LogInformation(LogInfoType.ClearAndAdd, "Hyperlink Deleted", string.Empty);
                                    }
                                }
                            }
                        }

                        if (f.xlModCmd == AppUtilities.ExcelModifyCmds.DelLinks)
                        {
                            if (Excel.RemoveHyperlinks(tempFilePackageViewer) == true)
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "Hyperlinks Deleted", string.Empty);
                            }
                            else
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "Unable to delete hyperlinks", string.Empty);
                            }
                        }

                        if (f.xlModCmd == AppUtilities.ExcelModifyCmds.DelEmbeddedLinks)
                        {
                            if (Excel.RemoveLinks(tempFilePackageViewer) == true)
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "Embedded Links Deleted", string.Empty);
                            }
                            else
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "Unable to delete links", string.Empty);
                            }
                        }

                        if (f.xlModCmd == AppUtilities.ExcelModifyCmds.DelSheet)
                        {
                            using (var fds = new FrmDeleteSheet(package, tempFilePackageViewer))
                            {
                                fds.ShowDialog();

                                if (fds.sheetName != string.Empty)
                                {
                                    scintilla1.AppendText("Sheet: " + fds.sheetName + " Removed");
                                }
                            }
                        }

                        if (f.xlModCmd == AppUtilities.ExcelModifyCmds.DelComments)
                        {
                            if (Excel.RemoveComments(tempFilePackageViewer) == true)
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "Comments Deleted", string.Empty);
                            }
                            else
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "Unable to delete comments", string.Empty);
                            }
                        }

                        if (f.xlModCmd == AppUtilities.ExcelModifyCmds.ConvertXlsmToXlsx)
                        {
                            string fNewName = Office.ConvertMacroEnabled2NonMacroEnabled(tempFilePackageViewer, Strings.oAppExcel);
                            if (fNewName != string.Empty)
                            {
                                scintilla1.AppendText(tempFilePackageViewer + Strings.convertedTo + fNewName);
                            }
                        }

                        if (f.xlModCmd == AppUtilities.ExcelModifyCmds.ConvertStrictToXlsx)
                        {
                            try
                            {
                                Cursor = Cursors.WaitCursor;

                                // check if the excelcnv.exe exists, without it, no conversion can happen
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
                                    // if no path is found, we will be unable to convert
                                    excelcnvPath = string.Empty;
                                    scintilla1.AppendText("** Unable to convert file **");
                                    return;
                                }

                                // check if the file is strict, no changes are made to the file yet
                                bool isStrict = false;

                                using (Package package = Package.Open(tempFilePackageViewer, FileMode.Open, FileAccess.Read))
                                {
                                    foreach (PackagePart part in package.GetParts())
                                    {
                                        if (part.Uri.ToString() == Strings.xlWorkbookXml)
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
                                                FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnConvertToNonStrictFormat_Click ReadToEnd Error = " + ex.Message);
                                            }
                                        }
                                    }
                                }

                                // if the file is strict format
                                // run the command to convert it to non-strict
                                if (isStrict == true)
                                {
                                    // setup destination file path
                                    string strOriginalFile = tempFilePackageViewer;
                                    string strOutputPath = Path.GetDirectoryName(strOriginalFile) + "\\";
                                    string strFileExtension = Path.GetExtension(strOriginalFile);
                                    string strOutputFileName = strOutputPath + Path.GetFileNameWithoutExtension(strOriginalFile) + Strings.wFixedFileParentheses + strFileExtension;

                                    // run the command to convert the file "excelcnv.exe -nme -oice "strict-file-path" "converted-file-path""
                                    string cParams = " -nme -oice " + Strings.chDblQuote + tempFilePackageViewer + Strings.chDblQuote + Strings.wSpaceChar + Strings.chDblQuote + strOutputFileName + Strings.chDblQuote;
                                    var proc = Process.Start(excelcnvPath, cParams);
                                    proc.Close();
                                    scintilla1.AppendText(Strings.fileConvertSuccessful);
                                    scintilla1.AppendText("File Location: " + strOutputFileName);
                                }
                                else
                                {
                                    scintilla1.AppendText("** File Is Not Open Xml Format (Strict) **");
                                }
                            }
                            catch (Exception ex)
                            {
                                LogInformation(LogInfoType.LogException, "BtnConvertToNonStrictFormat_Click Error = ", ex.Message);
                            }
                            finally
                            {
                                Cursor = Cursors.Default;
                            }
                        }

                        if (result == DialogResult.OK)
                        {
                            string modifiedPath = AddModifiedTextToFileName(tempFilePackageViewer);
                            File.Copy(tempFilePackageViewer, modifiedPath, true);
                        }
                    }
                }
                else if (StrOfficeApp == Strings.oAppPowerPoint)
                {
                    using (var f = new FrmPowerPointModify())
                    {
                        DialogResult result = f.ShowDialog();

                        if (result == DialogResult.Cancel)
                        {
                            return;
                        }

                        Cursor = Cursors.WaitCursor;

                        if (f.pptModCmd == AppUtilities.PowerPointModifyCmds.DeleteUnusedMasterLayouts)
                        {
                            using (PresentationDocument pDoc = PresentationDocument.Open(tempFilePackageViewer, true))
                            {
                                PowerPointFixes.DeleteUnusedMasterLayouts(pDoc);
                                scintilla1.AppendText("Unused Slide Layouts Deleted");
                            }
                        }

                        if (f.pptModCmd == AppUtilities.PowerPointModifyCmds.ResetBulletMargins)
                        {
                            using (PresentationDocument pDoc = PresentationDocument.Open(tempFilePackageViewer, true))
                            {
                                PowerPointFixes.ResetBulletMargins(pDoc);
                                scintilla1.AppendText("Bullet Margins Reset");
                            }
                        }

                        if (f.pptModCmd == AppUtilities.PowerPointModifyCmds.ConvertPptmToPptx)
                        {
                            string fNewName = Office.ConvertMacroEnabled2NonMacroEnabled(tempFilePackageViewer, Strings.oAppPowerPoint);
                            if (fNewName != string.Empty)
                            {
                                scintilla1.AppendText(tempFilePackageViewer + Strings.convertedTo + fNewName);
                            }
                        }

                        if (f.pptModCmd == AppUtilities.PowerPointModifyCmds.DelComments)
                        {
                            if (PowerPoint.DeleteComments(tempFilePackageViewer, string.Empty))
                            {
                                scintilla1.AppendText("Comments Removed");
                            }
                            else
                            {
                                scintilla1.AppendText("No Comments Removed");
                            }
                        }

                        if (f.pptModCmd == AppUtilities.PowerPointModifyCmds.RemovePIIOnSave)
                        {
                            using (PresentationDocument document = PresentationDocument.Open(tempFilePackageViewer, true))
                            {
                                document.PresentationPart.Presentation.RemovePersonalInfoOnSave = false;
                                document.PresentationPart.Presentation.Save();
                                scintilla1.AppendText("Remove PII On Save Disabled");
                            }
                        }

                        if (f.pptModCmd == AppUtilities.PowerPointModifyCmds.MoveSlide)
                        {
                            FrmMoveSlide mvFrm = new FrmMoveSlide(tempFileReadOnly)
                            {
                                Owner = this
                            };
                            mvFrm.ShowDialog();
                        }

                        if (f.pptModCmd == AppUtilities.PowerPointModifyCmds.ResetNotesPageSize)
                        {
                            PowerPointFixes.ResetNotesPageSize(tempFilePackageViewer);
                            scintilla1.AppendText("Notes Page Size Reset");
                        }

                        if (f.pptModCmd == AppUtilities.PowerPointModifyCmds.CustomNotesPageReset)
                        {
                            PowerPointFixes.CustomResetNotesPageSize(tempFilePackageViewer);
                            scintilla1.AppendText("Notes Page Size Reset");
                        }

                        if (result == DialogResult.OK)
                        {
                            string modifiedPath = AddModifiedTextToFileName(tempFilePackageViewer);
                            File.Copy(tempFilePackageViewer, modifiedPath, true);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogInfoType.LogException, "ModifyContents Error:", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void EditToolStripMenuItemRemoveCustomDocProps_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                if (Office.RemoveCustomDocProperties(package, toolStripStatusLabelDocType.Text))
                {
                    LogInformation(LogInfoType.ClearAndAdd, "Custom File Properties Removed.", string.Empty);
                }
                else
                {
                    throw new Exception();
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogInfoType.LogException, "Remove Custom Doc Props Failed", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void EditToolStripMenuItemRemoveCustomXml_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                if (Office.RemoveCustomXmlParts(package, tempFilePackageViewer, toolStripStatusLabelDocType.Text))
                {
                    LogInformation(LogInfoType.ClearAndAdd, "Custom Xml Parts Removed.", string.Empty);
                }
                else
                {
                    LogInformation(LogInfoType.ClearAndAdd, "Document Does Not Contain Custom Xml.", string.Empty);
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogInfoType.LogException, "Remove Custom Xml Failed", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void FileToolStripMenuItemClose_Click(object sender, EventArgs e)
        {
            FileClose();
        }

        private void MruToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (sender is not null) { OpenRecentFile(sender.ToString()!); }
        }

        private void MruToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (sender is not null) { OpenRecentFile(sender.ToString()!); }
        }

        private void MruToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (sender is not null) { OpenRecentFile(sender.ToString()!); }
        }

        private void MruToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            if (sender is not null) { OpenRecentFile(sender.ToString()!); }
        }

        private void MruToolStripMenuItem5_Click(object sender, EventArgs e)
        {
            if (sender is not null) { OpenRecentFile(sender.ToString()!); }
        }

        private void MruToolStripMenuItem6_Click(object sender, EventArgs e)
        {
            if (sender is not null) { OpenRecentFile(sender.ToString()!); }
        }

        private void MruToolStripMenuItem7_Click(object sender, EventArgs e)
        {
            if (sender is not null) { OpenRecentFile(sender.ToString()!); }
        }

        private void MruToolStripMenuItem8_Click(object sender, EventArgs e)
        {
            if (sender is not null) { OpenRecentFile(sender.ToString()!); }
        }

        private void MruToolStripMenuItem9_Click(object sender, EventArgs e)
        {
            if (sender is not null) { OpenRecentFile(sender.ToString()!); }
        }

        private void OpenErrorLogToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            AppUtilities.PlatformSpecificProcessStart(Strings.fLogFilePath);
        }

        private void WordDocumentRevisionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmRevisions frmRev = new FrmRevisions(tempFilePackageViewer)
            {
                Owner = this
            };
            frmRev.ShowDialog();
        }

        private void ViewPartPropertiesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Compression Settings: " + partPropCompression);
            sb.AppendLine("Content Type: " + partPropContentType);
            MessageBox.Show(sb.ToString(), "Part Properties", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void TvFiles_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            TreeNode tNode = TvFiles.GetNodeAt(e.X, e.Y);
            UpdateStreamDisplay(tNode);
        }

        /// <summary>
        /// populate the validation errors
        /// </summary>
        /// <param name="displayOnly">only display ui for validate xml button</param>
        private void ValidateLabelInfoXml(bool displayOnly)
        {
            validationErrors.Clear();

            try
            {
                ValidationEventHandler eventHandler = new ValidationEventHandler(ValidationEventHandler);
                XmlSchemaSet schema = new XmlSchemaSet();
                XmlTextReader xtr = new XmlTextReader(@".\Schemas\LabelInfo.xsd");
                XmlSchema sch = XmlSchema.Read(xtr, ValidationEventHandler);
                schema.Add(sch);

                var settings = new XmlReaderSettings();
                settings.Schemas.Add(sch);
                settings.ValidationType = ValidationType.Schema;
                settings.ValidationFlags |= XmlSchemaValidationFlags.ProcessInlineSchema;
                settings.ValidationFlags |= XmlSchemaValidationFlags.ReportValidationWarnings;
                settings.ValidationEventHandler += new ValidationEventHandler(ValidationEventHandler);

                using (TextReader textReader = new StringReader(scintilla1.Text))
                {
                    XmlReader rd = XmlReader.Create(textReader, settings);
                    XDocument doc = XDocument.Load(rd);
                    doc.Validate(schema, eventHandler);
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, ex.Message);

                if (displayOnly)
                {
                    MessageBox.Show(ex.Message, "Xml Validation Errors", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                isValidXml = false;
            }
            finally
            {
                if (isValidXml && displayOnly)
                {
                    MessageBox.Show("Xml Is Valid.", "Xml Validation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    // display schema errors
                    if (validationErrors.Count > 0)
                    {
                        StringBuilder sb = new StringBuilder();
                        int errorCount = 0;
                        foreach (string s in validationErrors)
                        {
                            errorCount++;
                            sb.Append(errorCount + Strings.wPeriod + s + "\r\n");
                        }

                        if (displayOnly)
                        {
                            MessageBox.Show(sb.ToString(), "Schema Validation Errors", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }

        private void TvFiles_KeyUp(object sender, KeyEventArgs e)
        {
            if (isEncrypted)
            {
                if (e.KeyCode == Keys.Down || e.KeyCode == Keys.Up)
                {
                    TreeNode tNode = TvFiles.SelectedNode;
                    UpdateStreamDisplay(tNode);
                }
            }
        }

        private void ToolStripButtonFixXml_Click(object sender, EventArgs e)
        {
            FixLabelInfo();
        }

        private void copySelectedLineToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            try
            {
                Clipboard.SetText(scintilla1.Lines[scintilla1.CurrentLine].Text);
            }
            catch (Exception ex)
            {
                LogInformation(LogInfoType.LogException, "BtnCopySelectedLine Error", ex.Message);
            }
        }

        private void renderRTFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmRichTextBox frmRtb = new FrmRichTextBox(scintilla1.Text)
            {
                Owner = this
            };
            frmRtb.ShowDialog();
        }

        #endregion
    }
}