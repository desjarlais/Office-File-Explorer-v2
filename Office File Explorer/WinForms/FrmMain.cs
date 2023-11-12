// Open XML SDK refs
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

// App refs
using Office_File_Explorer.Helpers;
using Office_File_Explorer.WinForms;

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

using File = System.IO.File;
using Color = System.Drawing.Color;
using Person = DocumentFormat.OpenXml.Office2013.Word.Person;

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

        // openmcdf
        private FileStream fs;

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

        public bool hasXmlError;

        public enum OpenXmlInnerFileTypes
        {
            Word,
            Excel,
            PowerPoint,
            Outlook,
            XML,
            Image,
            Binary,
            Video,
            Audio,
            Text,
            Other
        }

        // enums
        public enum LogInfoType { ClearAndAdd, TextOnly, InvalidFile, LogException, EmptyCount };

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
                tempFileReadOnly = Path.GetTempFileName().Replace(".tmp", ".docx");
                File.Copy(toolStripStatusLabelFilePath.Text, tempFileReadOnly, true);

                tempFilePackageViewer = Path.GetTempFileName().Replace(".tmp", ".docx");
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
            DisableCustomUIIcons();
            DisableModifyUI();
            DisableUI();
            package?.Close();
            pkgParts?.Clear();
            tvFiles.Nodes.Clear();
            toolStripStatusLabelFilePath.Text = Strings.wHeadingBegin;
            toolStripStatusLabelDocType.Text = Strings.wHeadingBegin;
            rtbDisplay.Clear();
            fileToolStripMenuItemClose.Enabled = false;
        }

        /// <summary>
        /// disable app feature related buttons
        /// </summary>
        public void DisableUI()
        {
            toolStripButtonViewContents.Enabled = false;
            toolStripButtonFixDoc.Enabled = false;
            editToolStripMenuFindReplace.Enabled = false;
            editToolStripMenuItemModifyContents.Enabled = false;
            editToolStripMenuItemRemoveCustomDocProps.Enabled = false;
            editToolStripMenuItemRemoveCustomXml.Enabled = false;
            excelSheetViewerToolStripMenuItem.Enabled = false;
            toolStripButtonModify.Enabled = false;
            toolStripButtonValidateXml.Enabled = false;
            wordDocumentRevisionsToolStripMenuItem.Enabled = false;

            if (package is null)
            {
                fileToolStripMenuItemClose.Enabled = false;
            }
        }

        /// <summary>
        /// enable app feature related buttons
        /// </summary>
        public void EnableUI()
        {
            toolStripButtonViewContents.Enabled = true;
            toolStripButtonFixDoc.Enabled = true;
            editToolStripMenuFindReplace.Enabled = true;
            editToolStripMenuItemModifyContents.Enabled = true;
            editToolStripMenuItemRemoveCustomDocProps.Enabled = true;
            editToolStripMenuItemRemoveCustomXml.Enabled = true;
            toolStripButtonModify.Enabled = true;
            fileToolStripMenuItemClose.Enabled = true;
            toolStripButtonModify.Enabled = true;
            toolStripButtonValidateXml.Enabled = true;
        }

        public void CopyAllItems()
        {
            try
            {
                if (rtbDisplay.Text.Length == 0) { return; }
                StringBuilder buffer = new StringBuilder();
                foreach (string s in rtbDisplay.Lines)
                {
                    buffer.Append(s);
                    buffer.Append('\n');
                }

                Clipboard.SetText(buffer.ToString());
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
                FrmEncryptedFile cForm = new FrmEncryptedFile(fs, true)
                {
                    Owner = this
                };

                if (cForm.IsDisposed)
                {
                    return;
                }
                else
                {
                    cForm.ShowDialog();
                }
                
                if (fs is not null)
                {
                    fs.Close();
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogInfoType.LogException, "OpenEncryptedOfficeDocument Error", ex.Message);
            }
        }

        public void DisplayInvalidFileFormatError()
        {
            rtbDisplay.AppendText("Unable to open file, possible causes are:\r\n");
            rtbDisplay.AppendText(" - file corruption\r\n");
            rtbDisplay.AppendText(" - file encrypted\r\n");
            rtbDisplay.AppendText(" - file password protected\r\n");
            rtbDisplay.AppendText(" - binary Office Document (View file contents with Tools -> Structured Storage Viewer)\r\n");
        }

        /// <summary>
        /// move cursor to a given location of the richtextbox
        /// </summary>
        /// <param name="startLocation"></param>
        /// <param name="length"></param>
        public void MoveCursorToLocation(int startLocation, int length)
        {
            rtbDisplay.SelectionStart = startLocation;
            rtbDisplay.SelectionLength = length;
        }

        public void FindText()
        {
            if (toolStripTextBoxFind.Text == string.Empty)
            {
                return;
            }
            else
            {
                rtbDisplay.Focus();

                // if the cursor is at the end of the textbox, change start position to 0
                if (rtbDisplay.SelectionStart == rtbDisplay.Text.Length)
                {
                    MoveCursorToLocation(0, 0);
                }

                try
                {
                    int indexToText;
                    indexToText = rtbDisplay.Find(toolStripTextBoxFind.Text, rtbDisplay.SelectionStart + 1, RichTextBoxFinds.None);
                    if (indexToText >= 0)
                    {
                        MoveCursorToLocation(indexToText, toolStripTextBoxFind.Text.Length);
                    }

                    // end of the document, restart at the beginning
                    if (indexToText == -1)
                    {
                        // only move if something was found
                        if (rtbDisplay.SelectionStart != 0)
                        {
                            MoveCursorToLocation(0, 0);
                            FindText();
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogInformation(LogInfoType.LogException, "FindText Error", ex.Message);
                }
            }
        }

        /// <summary>
        /// majority of open file logic is here
        /// </summary>
        public void OpenOfficeDocument(bool isFromMRU)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

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
                        DisableUI();
                        DisableModifyUI();
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
                    rtbDisplay.Clear();

                    // handle msg files
                    if (toolStripStatusLabelFilePath.Text.EndsWith(Strings.msgFileExt))
                    {
                        Stream messageStream = File.Open(toolStripStatusLabelFilePath.Text, FileMode.Open, FileAccess.Read);
                        OutlookStorage.Message message = new OutlookStorage.Message(messageStream);
                        messageStream.Close();

                        LoadMsgToTree(message, tvFiles.Nodes.Add("MSG"));
                        tvFiles.ImageIndex = 6;
                        tvFiles.SelectedImageIndex = 6;
                        message.Dispose();
                        toolStripStatusLabelDocType.Text = Strings.oAppOutlook;
                        return;
                    }

                    // handle office files
                    if (!FileUtilities.IsZipArchiveFile(toolStripStatusLabelFilePath.Text))
                    {
                        // if the file doesn't start with PK, we can stop trying to process it
                        DisplayInvalidFileFormatError();
                        DisableUI();
                        
                        // handle encrypted files
                        structuredStorageViewerToolStripMenuItem.Enabled = true;
                    }
                    else
                    {
                        // if the file does start with PK, check if it fails in the SDK
                        if (OpenWithSdk(toolStripStatusLabelFilePath.Text))
                        {
                            // set the file type
                            toolStripStatusLabelDocType.Text = StrOfficeApp;
                            DisableUI();

                            // populate the parts
                            PopulatePackageParts();

                            // check if any zip items are corrupt
                            if (Properties.Settings.Default.CheckZipItemCorrupt == true && toolStripStatusLabelDocType.Text == Strings.oAppWord)
                            {
                                if (Office.IsZippedFileCorrupt(toolStripStatusLabelFilePath.Text))
                                {
                                    rtbDisplay.AppendText("Warning - One of the zipped items is corrupt.");
                                }
                            }

                            // clear the previous doc if there was one and setup temp files
                            TempFileSetup();
                            tvFiles.Nodes.Clear();
                            rtbDisplay.Clear();
                            package?.Close();
                            pkgParts?.Clear();
                            LoadPartsIntoViewer();
                            EnableUI();
                        }
                        else
                        {
                            // if it failed the SDK, disable all buttons except the fix corrupt doc button
                            DisableUI();
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
                recipientNode.Nodes.Add(recipient.Type + ": " + recipient.Email);
                recipientNode.Tag = new string[] { "Display Name: " + recipient.DisplayName, "Email: " + recipient.Email };
            }

            TreeNode attachmentNode = messageNode.Nodes.Add("Attachments: " + message.Attachments.Count);
            foreach (OutlookStorage.Attachment attachment in message.Attachments)
            {
                attachmentNode.Nodes.Add(attachment.Filename + ": " + attachment.Data.Length + "b");
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
        /// load the doc parts into the treeview
        /// </summary>
        public void LoadPartsIntoViewer()
        {
            // populate the treeview
            package = Package.Open(toolStripStatusLabelFilePath.Text, FileMode.Open, FileAccess.ReadWrite);

            TreeNode tRoot = new TreeNode();
            tRoot.Text = toolStripStatusLabelFilePath.Text;

            // update file icon
            if (GetFileType(toolStripStatusLabelFilePath.Text) == OpenXmlInnerFileTypes.Word)
            {
                tvFiles.SelectedImageIndex = 0;
                tvFiles.ImageIndex = 0;
            }
            else if (GetFileType(toolStripStatusLabelFilePath.Text) == OpenXmlInnerFileTypes.Excel)
            {
                tvFiles.SelectedImageIndex = 2;
                tvFiles.ImageIndex = 2;
            }
            else if (GetFileType(toolStripStatusLabelFilePath.Text) == OpenXmlInnerFileTypes.PowerPoint)
            {
                tvFiles.SelectedImageIndex = 1;
                tvFiles.ImageIndex = 1;
            }

            // update inner file icon, need to update both the selected and normal image index
            foreach (PackagePart part in package.GetParts())
            {
                tRoot.Nodes.Add(part.Uri.ToString());

                if (GetFileType(part.Uri.ToString()) == OpenXmlInnerFileTypes.XML)
                {
                    tRoot.Nodes[tRoot.Nodes.Count - 1].ImageIndex = 3;
                    tRoot.Nodes[tRoot.Nodes.Count - 1].SelectedImageIndex = 3;
                }
                else if (GetFileType(part.Uri.ToString()) == OpenXmlInnerFileTypes.Image)
                {
                    tRoot.Nodes[tRoot.Nodes.Count - 1].ImageIndex = 4;
                    tRoot.Nodes[tRoot.Nodes.Count - 1].SelectedImageIndex = 4;
                }
                else if (GetFileType(part.Uri.ToString()) == OpenXmlInnerFileTypes.Word)
                {
                    tRoot.Nodes[tRoot.Nodes.Count - 1].ImageIndex = 0;
                    tRoot.Nodes[tRoot.Nodes.Count - 1].SelectedImageIndex = 0;
                }
                else if (GetFileType(part.Uri.ToString()) == OpenXmlInnerFileTypes.Excel)
                {
                    tRoot.Nodes[tRoot.Nodes.Count - 1].ImageIndex = 2;
                    tRoot.Nodes[tRoot.Nodes.Count - 1].SelectedImageIndex = 2;
                }
                else if (GetFileType(part.Uri.ToString()) == OpenXmlInnerFileTypes.PowerPoint)
                {
                    tRoot.Nodes[tRoot.Nodes.Count - 1].ImageIndex = 1;
                    tRoot.Nodes[tRoot.Nodes.Count - 1].SelectedImageIndex = 1;
                }
                else if (GetFileType(part.Uri.ToString()) == OpenXmlInnerFileTypes.Binary)
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

            tvFiles.Nodes.Add(tRoot);
            tvFiles.ExpandAll();
            DisableModifyUI();
        }

        public OpenXmlInnerFileTypes GetFileType(string path)
        {
            switch (Path.GetExtension(path))
            {
                case ".docx":
                case ".dotx":
                case ".dotm":
                case ".docm":
                    return OpenXmlInnerFileTypes.Word;
                case ".xlsx":
                case ".xlsm":
                case ".xltm":
                case ".xltx":
                case ".xlsb":
                    return OpenXmlInnerFileTypes.Excel;
                case ".pptx":
                case ".pptm":
                case ".ppsx":
                case ".ppsm":
                case ".potx":
                case ".potm":
                    return OpenXmlInnerFileTypes.PowerPoint;
                case ".msg":
                    return OpenXmlInnerFileTypes.Outlook;
                case ".jpeg":
                case ".jpg":
                case ".bmp":
                case ".png":
                case ".gif":
                case ".emf":
                case ".wmf":
                    return OpenXmlInnerFileTypes.Image;
                case ".xml":
                case ".rels":
                    return OpenXmlInnerFileTypes.XML;
                case ".mp4":
                case ".avi":
                case ".wmv":
                case ".mov":
                    return OpenXmlInnerFileTypes.Video;
                case ".mp3":
                case ".wav":
                case ".wma":
                    return OpenXmlInnerFileTypes.Audio;
                case ".txt":
                    return OpenXmlInnerFileTypes.Text;
                case ".bin":
                case ".sigs":
                case ".odttf":
                    return OpenXmlInnerFileTypes.Binary;
                default:
                    return OpenXmlInnerFileTypes.Other;
            }
        }

        public void LogInformation(LogInfoType type, string output, string ex)
        {
            switch (type)
            {
                case LogInfoType.ClearAndAdd:
                    rtbDisplay.Clear();
                    rtbDisplay.AppendText(output);
                    break;
                case LogInfoType.InvalidFile:
                    rtbDisplay.Clear();
                    rtbDisplay.AppendText(Strings.invalidFile);
                    break;
                case LogInfoType.LogException:
                    rtbDisplay.Clear();
                    rtbDisplay.AppendText(output + "\r\n" + ex);
                    FileUtilities.WriteToLog(Strings.fLogFilePath, output);
                    FileUtilities.WriteToLog(Strings.fLogFilePath, ex);
                    break;
                case LogInfoType.EmptyCount:
                    rtbDisplay.AppendText(Strings.wNone);
                    break;
                default:
                    rtbDisplay.AppendText(output);
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
            string body = string.Empty;
            bool fSuccess = false;

            try
            {
                Cursor = Cursors.WaitCursor;
                UriRelationshipErrorHandler uriHandler = new UriRelationshipErrorHandler();

                // add opensettings to get around the fix malformed uri issue
                var openSettings = new OpenSettings()
                {
                    RelationshipErrorHandlerFactory = package => { return uriHandler; }
                };

                if (FileUtilities.GetAppFromFileExtension(file) == Strings.oAppWord)
                {
                    using (WordprocessingDocument document = WordprocessingDocument.Open(file, true, openSettings))
                    {
                        // try to get the localname of the document.xml file, if it fails, it is not a Word file
                        StrOfficeApp = Strings.oAppWord;
                        body = document.MainDocumentPart.Document.LocalName;
                        fSuccess = true;
                    }
                }
                else if (FileUtilities.GetAppFromFileExtension(file) == Strings.oAppExcel)
                {
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(file, true, openSettings))
                    {
                        // try to get the localname of the workbook.xml and file if it fails, its not an Excel file
                        StrOfficeApp = Strings.oAppExcel;
                        body = document.WorkbookPart.Workbook.LocalName;
                        fSuccess = true;
                    }
                }
                else if (FileUtilities.GetAppFromFileExtension(file) == Strings.oAppPowerPoint)
                {
                    using (PresentationDocument document = PresentationDocument.Open(file, true, openSettings))
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
        public static void AppExitWork()
        {
            try
            {
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

        #endregion

        #region Button Events

        private void OpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            structuredStorageViewerToolStripMenuItem.Enabled = false;
            EnableUI();
            EnableModifyUI();
            OpenOfficeDocument(false);

            if (toolStripStatusLabelFilePath.Text != Strings.wHeadingBegin)
            {
                AddFileToMRU();
            }

            if (toolStripStatusLabelDocType.Text == Strings.oAppExcel)
            {
                excelSheetViewerToolStripMenuItem.Enabled = true;
            }

            if (toolStripStatusLabelDocType.Text == Strings.oAppWord)
            {
                wordDocumentRevisionsToolStripMenuItem.Enabled = true;
            }
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
            package?.Close();
            AppExitWork();
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            package?.Close();
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

        private void OpenErrorLogToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AppUtilities.PlatformSpecificProcessStart(Strings.fLogFilePath);
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
                Clipboard.SetText(rtbDisplay.Lines[rtbDisplay.GetLineFromCharIndex(rtbDisplay.SelectionStart)]);
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

        private void structuredStorageViewerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenEncryptedOfficeDocument(toolStripStatusLabelFilePath.Text, true);
        }

        private void excelSheetViewerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var f = new FrmSheetViewer(tempFileReadOnly))
            {
                var result = f.ShowDialog();
            }
        }

        public void EnableModifyUI()
        {
            rtbDisplay.ReadOnly = false;
            rtbDisplay.BackColor = SystemColors.Window;
            toolStripButtonSave.Enabled = true;
            toolStripDropDownButtonInsert.Enabled = true;
            toolStripButtonReplace.Enabled = true;
        }

        public void DisableModifyUI()
        {
            rtbDisplay.ReadOnly = true;
            rtbDisplay.BackColor = SystemColors.Control;
            toolStripButtonSave.Enabled = false;
            toolStripDropDownButtonInsert.Enabled = false;
            toolStripButtonReplace.Enabled = false;
        }

        public void EnableCustomUIIcons()
        {
            toolStripButtonGenerateCallback.Enabled = true;
            toolStripDropDownButtonInsert.Enabled = true;
            toolStripButtonInsertIcon.Enabled = true;
        }

        public void DisableCustomUIIcons()
        {
            toolStripDropDownButtonInsert.Enabled = false;
            toolStripButtonSave.Enabled = false;
            toolStripButtonGenerateCallback.Enabled = false;
            toolStripButtonInsertIcon.Enabled = false;
        }

        private void ShowError(string errorText)
        {
            MessageBox.Show(this, errorText, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        static void ValidationCallback(object sender, ValidationEventArgs args)
        {
            if (args.Severity == XmlSeverityType.Warning)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "CustomUI Validation Warning:" + args.Message);
            }
            else if (args.Severity == XmlSeverityType.Error)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "CustomUI Validation Error:" + args.Message);
            }
        }

        /// <summary>
        /// validate the labelinfo.xml file
        /// </summary>
        public void ValidatePartXml()
        {
            hasXmlError = false;

            if (rtbDisplay.Text == null || rtbDisplay.Text.Length == 0)
            {
                return;
            }

            try
            {
                ValidationEventHandler eventHandler = new ValidationEventHandler(ValidationEventHandler);
                XmlSchemaSet schema = new XmlSchemaSet();
                schema.Add(string.Empty, XmlReader.Create(new StringReader(Strings.xsdMarkup)));

                var settings = new XmlReaderSettings();
                settings.Schemas.Add("http://schemas.microsoft.com/office/2020/mipLabelMetadata", XmlReader.Create(new StringReader(Strings.xsdMarkup)));
                settings.ValidationType = ValidationType.Schema;
                settings.ValidationFlags |= XmlSchemaValidationFlags.ProcessInlineSchema;
                settings.ValidationFlags |= XmlSchemaValidationFlags.ReportValidationWarnings;
                settings.ValidationEventHandler += new ValidationEventHandler(ValidationEventHandler);

                using (TextReader textReader = new StringReader(rtbDisplay.Text))
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

            if (!hasXmlError)
            {
                MessageBox.Show("Xml Valid", "Xml Validation", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// validate customui.xml
        /// </summary>
        /// <param name="showValidMessage"></param>
        /// <returns></returns>
        public bool ValidateXml(bool showValidMessage)
        {
            if (rtbDisplay.Text == null || rtbDisplay.Text.Length == 0)
            {
                return false;
            }

            rtbDisplay.SuspendLayout();

            try
            {
                XmlTextReader xtr = new XmlTextReader(@".\Schemas\customui14.xsd");
                XmlSchema schema = XmlSchema.Read(xtr, ValidationCallback);

                XmlDocument xmlDoc = new XmlDocument();

                if (schema == null)
                {
                    return false;
                }

                xmlDoc.Schemas.Add(schema);
                xmlDoc.LoadXml(rtbDisplay.Text);

                if (xmlDoc.DocumentElement.NamespaceURI.ToString() != schema.TargetNamespace)
                {
                    StringBuilder errorText = new StringBuilder();
                    errorText.Append("Unknown Namespace".Replace("|1", xmlDoc.DocumentElement.NamespaceURI.ToString()));
                    errorText.Append("\n" + "CustomUI Namespace".Replace("|1", schema.TargetNamespace));

                    ShowError(errorText.ToString());
                    return false;
                }

                hasXmlError = false;
                xmlDoc.Validate(XmlValidationEventHandler);
            }
            catch (XmlException ex)
            {
                ShowError("Invalid Xml" + "\n" + ex.Message);
                return false;
            }

            rtbDisplay.ResumeLayout();

            if (!hasXmlError)
            {
                if (showValidMessage)
                {
                    MessageBox.Show(this, "Valid Xml", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                return true;
            }
            return false;
        }

        public void LogXmlValidationError(ValidationEventArgs e)
        {
            lock (this)
            {
                hasXmlError = true;
            }

            MessageBox.Show("Error at Line #" + e.Exception.LineNumber + " Position #" + e.Exception.LinePosition + Strings.wColonBuffer + e.Message,
                        "Xml Validation", MessageBoxButtons.OK, MessageBoxIcon.Error);
            FileUtilities.WriteToLog(Strings.fLogFilePath, "Xml Validation Error at Line #" + e.Exception.LineNumber + " Position #"
                + e.Exception.LinePosition + Strings.wColonBuffer + e.Message);
        }

        private void XmlValidationEventHandler(object sender, ValidationEventArgs e)
        {
            LogXmlValidationError(e);
        }

        /// <summary>
        /// display schema validation errors
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ValidationEventHandler(object sender, ValidationEventArgs e)
        {
            switch (e.Severity)
            {
                case XmlSeverityType.Error:
                    LogXmlValidationError(e);
                    break;
                case XmlSeverityType.Warning:
                    LogXmlValidationError(e);
                    break;
            }
        }

        private void AddPart(XMLParts partType)
        {
            OfficePart newPart = CreateCustomUIPart(partType);
            TreeNode partNode = ConstructPartNode(newPart);
            TreeNode currentNode = tvFiles.Nodes[0];
            if (currentNode == null) return;

            tvFiles.SuspendLayout();
            currentNode.Nodes.Add(partNode);
            rtbDisplay.Text = string.Empty;
            tvFiles.SelectedNode = partNode;
            tvFiles.ResumeLayout();

            // refresh the treeview
            string prevPath = toolStripStatusLabelFilePath.Text;
            FileClose();
            toolStripStatusLabelFilePath.Text = prevPath;
            LoadPartsIntoViewer();
            toolStripButtonModify.Enabled = true;
            fileToolStripMenuItemClose.Enabled = true;
        }

        private TreeNode ConstructPartNode(OfficePart part)
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

            OfficePart part = null;
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
                structuredStorageViewerToolStripMenuItem.Enabled = false;
                EnableUI();
                EnableModifyUI();
                OpenOfficeDocument(true);

                if (toolStripStatusLabelDocType.Text == Strings.oAppWord)
                {
                    wordDocumentRevisionsToolStripMenuItem.Enabled = true;
                }
            }
        }

        /// <summary>
        /// used for the richtextbox to find/replace
        /// </summary>
        public void ReplaceText()
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

            rtbDisplay.SelectionStart = 0;
            rtbDisplay.SelectionLength = rtbDisplay.TextLength;
            rtbDisplay.SelectedText = rtbDisplay.SelectedText.Replace(findText, replaceText);
            rtbDisplay.SelectionStart = 0;
            rtbDisplay.SelectionLength = 0;

            if (Properties.Settings.Default.DisableXmlColorFormatting == false)
            {
                FormatXmlColors();
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
        private void tvFiles_AfterSelect(object sender, TreeViewEventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                // render msg content
                if (toolStripStatusLabelFilePath.Text.EndsWith(".msg"))
                {
                    string[] body = e.Node.Tag as string[];

                    if (body != null)
                    {
                        if (Properties.Settings.Default.MsgAsRtf)
                        {
                            rtbDisplay.Rtf = body[1];
                        }
                        else
                        {
                            rtbDisplay.Text = body[0];
                        }
                    }

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
                    tvFiles.ExpandAll();
                    return;
                }

                if (GetFileType(e.Node.Text) == OpenXmlInnerFileTypes.XML)
                {
                    // customui files have additional editing options
                    if (e.Node.Text.EndsWith("customUI.xml") || e.Node.Text.EndsWith("customUI14.xml"))
                    {
                        EnableCustomUIIcons();
                    }
                    else
                    {
                        DisableCustomUIIcons();
                    }

                    // load file contents
                    foreach (PackagePart pp in pkgParts)
                    {
                        if (pp.Uri.ToString() == tvFiles.SelectedNode.Text)
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

                                tvFiles.SuspendLayout();
                                rtbDisplay.Text = contents;
                                
                                // check for xml color setting
                                if (Properties.Settings.Default.DisableXmlColorFormatting == false)
                                {
                                    FormatXmlColors();
                                }

                                tvFiles.ResumeLayout();
                                ScrollToTopOfRtb();
                                return;
                            }
                        }
                    }
                }
                else if (GetFileType(e.Node.Text) == OpenXmlInnerFileTypes.Image)
                {
                    // currently showing images with a form
                    // TODO find a way to keep the image in the main form
                    foreach (PackagePart pp in pkgParts)
                    {
                        if (pp.Uri.ToString() == tvFiles.SelectedNode.Text)
                        {
                            partPropCompression = pp.CompressionOption.ToString();
                            partPropContentType = pp.ContentType;

                            // need to implement non-bitmap images
                            if (pp.Uri.ToString().EndsWith(".emf") || (pp.Uri.ToString().EndsWith(".svg")))
                            {
                                rtbDisplay.Text = "No Viewer For File Type";
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
                else if (GetFileType(e.Node.Text) == OpenXmlInnerFileTypes.Binary)
                {
                    foreach (PackagePart pp in pkgParts)
                    {
                        if (pp.Uri.ToString() == tvFiles.SelectedNode.Text)
                        {
                            partPropCompression = pp.CompressionOption.ToString();
                            partPropContentType = pp.ContentType;

                            Stream stream = pp.GetStream();
                            byte[] binData = FileUtilities.ReadToEnd(stream);
                            rtbDisplay.Text = Convert.ToHexString(binData);
                            stream.Close();
                            return;
                        }
                    }
                }
                else
                {
                    rtbDisplay.Text = "No Viewer For File Type";
                }
            }
            catch (Exception ex)
            {
                rtbDisplay.Text = "Error: " + ex.Message;
            }
            finally
            {
                Cursor = Cursors.Default;
                DisableModifyUI();
            }
        }

        public void ScrollToTopOfRtb()
        {
            rtbDisplay.SelectionStart = 0;
            rtbDisplay.ScrollToCaret();
        }

        /// <summary>
        /// format the xml tags with different colors to make it easier to read
        /// </summary>
        public void FormatXmlColors()
        {
            string pattern = @"</?(?<tagName>[a-zA-Z0-9_:\-]+)(\s+(?<attName>[a-zA-Z0-9_:\-]+)(?<attValue>(=""[^""]+"")?))*\s*/?>";
            Regex regExXmlColors = new Regex(pattern, RegexOptions.Compiled);
            int matchCount = regExXmlColors.Matches(rtbDisplay.Text).Count;

            // perf check, bail if we are over 15k matches
            if (matchCount > 15000)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "FormatXmlColor Match Count = " + matchCount.ToString());
                return;
            }

            foreach (Match m in regExXmlColors.Matches(rtbDisplay.Text))
            {
                rtbDisplay.Select(m.Index, m.Length);
                rtbDisplay.SelectionColor = Color.Blue;

                var tagName = m.Groups["tagName"].Value;
                rtbDisplay.Select(m.Groups["tagName"].Index, m.Groups["tagName"].Length);
                rtbDisplay.SelectionColor = Color.DarkRed;

                var attGroup = m.Groups["attName"];
                if (attGroup is not null)
                {
                    var atts = attGroup.Captures;
                    for (int i = 0; i < atts.Count; i++)
                    {
                        rtbDisplay.Select(atts[i].Index, atts[i].Length);
                        rtbDisplay.SelectionColor = Color.Red;
                    }
                }
            }
        }

        private void toolStripButtonValidateXml_Click(object sender, EventArgs e)
        {
            if (tvFiles.SelectedNode is null)
            {
                return;
            }

            if (tvFiles.SelectedNode.Text.EndsWith(Strings.offLabelInfo))
            {
                ValidatePartXml();
            }
            else if (tvFiles.SelectedNode.Text.EndsWith(Strings.offCustomUI14Xml) || tvFiles.SelectedNode.Text.EndsWith(Strings.offCustomUIXml))
            {
                ValidateXml(true);
            }
        }

        private void toolStripButtonGenerateCallback_Click(object sender, EventArgs e)
        {
            // if there is no callback , then there is no point in generating the callback code
            if (rtbDisplay.Text == null || rtbDisplay.Text.Length == 0)
            {
                return;
            }

            // if the xml is not valid, then there is no point in generating the callback code
            if (!ValidateXml(false))
            {
                return;
            }

            // generate the callback code
            try
            {
                XmlDocument customUI = new XmlDocument();
                customUI.LoadXml(rtbDisplay.Text);
                StringBuilder callbacks = CallbackBuilder.GenerateCallback(customUI);
                callbacks.Append("}");

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

        private void office2010CustomUIPartToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddPart(XMLParts.RibbonX14);
            rtbDisplay.Text = string.Empty;

        }

        private void office2007CustomUIPartToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddPart(XMLParts.RibbonX12);
            rtbDisplay.Text = string.Empty;
        }

        private void customOutspaceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rtbDisplay.Text = Strings.xmlCustomOutspace;
            FormatXmlColors();
            EnableModifyUI();
        }

        private void customTabToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rtbDisplay.Text = Strings.xmlCustomTab;
            FormatXmlColors();
            EnableModifyUI();
        }

        private void excelCustomTabToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rtbDisplay.Text = Strings.xmlExcelCustomTab;
            FormatXmlColors();
            EnableModifyUI();
        }

        private void repurposeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rtbDisplay.Text = Strings.xmlRepurpose;
            FormatXmlColors();
            EnableModifyUI();
        }

        private void wordGroupOnInsertTabToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rtbDisplay.Text = Strings.xmlWordGroupInsertTab;
            FormatXmlColors();
            EnableModifyUI();
        }

        private void toolStripButtonInsertIcon_Click(object sender, EventArgs e)
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

                TreeNode partNode = null;
                foreach (TreeNode node in tvFiles.Nodes[0].Nodes)
                {
                    if (node.Text == part.Name)
                    {
                        partNode = node;
                        break;
                    }
                }

                tvFiles.SuspendLayout();

                foreach (string fileName in fDialog.FileNames)
                {
                    try
                    {
                        string id = XmlConvert.EncodeName(Path.GetFileNameWithoutExtension(fileName));
                        Stream imageStream = File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                        Image image = Image.FromStream(imageStream, true, true);

                        // The file is a valid image at this point.
                        id = part.AddImage(fileName, id);
                        if (id == null) continue;

                        imageStream.Close();

                        TreeNode imageNode = new TreeNode(id);
                        imageNode.ImageKey = "_" + id;
                        imageNode.SelectedImageKey = imageNode.ImageKey;
                        imageNode.Tag = partType;

                        tvFiles.ImageList.Images.Add(imageNode.ImageKey, image);
                        tvFiles.SelectedNode.Nodes.Add(imageNode);
                    }
                    catch (Exception ex)
                    {
                        ShowError(ex.Message);
                        continue;
                    }
                }

                tvFiles.ResumeLayout();
            }
        }

        private void toolStripButtonModify_Click(object sender, EventArgs e)
        {
            EnableModifyUI();
        }

        private void toolStripButtonSave_Click(object sender, EventArgs e)
        {
            bool isModified = false;

            foreach (PackagePart pp in pkgParts)
            {
                if (pp.Uri.ToString() == tvFiles.SelectedNode.Text)
                {
                    MemoryStream ms = new MemoryStream();
                    using (TextWriter tw = new StreamWriter(ms))
                    {
                        tw.Write(rtbDisplay.Text);
                        tw.Flush();

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
                tvFiles.Nodes.Clear();
                LoadPartsIntoViewer();
            }
        }

        private void toolStripButtonViewContents_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                rtbDisplay.Clear();
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

        private void toolStripButtonFixCorruptDoc_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                StrDestFileName = AddTextToFileName(toolStripStatusLabelFilePath.Text, Strings.wFixedFileParentheses);
                bool isXmlException = false;
                string strDocText = string.Empty;
                IsFixed = false;
                rtbDisplay.Clear();

                if (StrExtension == Strings.docxFileExt)
                {
                    if ((File.GetAttributes(toolStripStatusLabelFilePath.Text) & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                    {
                        rtbDisplay.AppendText("ERROR: File is Read-Only.");
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

                                using (TextReader tr = new StreamReader(part.GetStream(FileMode.Open, FileAccess.Read)))
                                {
                                    strDocText = tr.ReadToEnd();
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
                                                        rtbDisplay.AppendText(Strings.invalidTag + m.Value);
                                                        rtbDisplay.AppendText(Strings.replacedWith + ValidXmlTags.StrValidVshapegroup);
                                                    }
                                                    else
                                                    {
                                                        strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidVshape);
                                                        rtbDisplay.AppendText(Strings.invalidTag + m.Value);
                                                        rtbDisplay.AppendText(Strings.replacedWith + ValidXmlTags.StrValidVshape);
                                                    }
                                                    break;

                                                case InvalidXmlTags.StrInvalidOmathWps:
                                                    strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidomathwps);
                                                    rtbDisplay.AppendText(Strings.invalidTag + m.Value);
                                                    rtbDisplay.AppendText(Strings.replacedWith + ValidXmlTags.StrValidomathwps);
                                                    break;

                                                case InvalidXmlTags.StrInvalidOmathWpg:
                                                    strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidomathwpg);
                                                    rtbDisplay.AppendText(Strings.invalidTag + m.Value);
                                                    rtbDisplay.AppendText(Strings.replacedWith + ValidXmlTags.StrValidomathwpg);
                                                    break;

                                                case InvalidXmlTags.StrInvalidOmathWpc:
                                                    strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidomathwpc);
                                                    rtbDisplay.AppendText(Strings.invalidTag + m.Value);
                                                    rtbDisplay.AppendText(Strings.replacedWith + ValidXmlTags.StrValidomathwpc);
                                                    break;

                                                case InvalidXmlTags.StrInvalidOmathWpi:
                                                    strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidomathwpi);
                                                    rtbDisplay.AppendText(Strings.invalidTag + m.Value);
                                                    rtbDisplay.AppendText(Strings.replacedWith + ValidXmlTags.StrValidomathwpi);
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
                                                            rtbDisplay.AppendText(Strings.invalidTag + m.Value);
                                                            rtbDisplay.AppendText(Strings.replacedWith + ValidXmlTags.StrValidMcChoice4);
                                                            break;
                                                        }

                                                        // replace mc:choice and hold onto the tag that follows
                                                        strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidMcChoice3 + m.Groups[2].Value);
                                                        rtbDisplay.AppendText(Strings.invalidTag + m.Value);
                                                        rtbDisplay.AppendText(Strings.replacedWith + ValidXmlTags.StrValidMcChoice3 + m.Groups[2].Value);
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
                                                            rtbDisplay.AppendText(Strings.invalidTag + m.Value);
                                                            rtbDisplay.AppendText(Strings.replacedWith + "Fallback tag deleted.");
                                                            break;
                                                        }

                                                        // if there is no closing fallback tag, we can replace the match with the omitFallback valid tags
                                                        // then we need to also add the trailing tag, since it's always different but needs to stay in the file
                                                        strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrOmitFallback + m.Groups[2].Value);
                                                        rtbDisplay.AppendText(Strings.invalidTag + m.Value);
                                                        rtbDisplay.AppendText(Strings.replacedWith + ValidXmlTags.StrOmitFallback + m.Groups[2].Value);
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
                                        rtbDisplay.AppendText(" ## No Corruption Found  ## ");
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
                rtbDisplay.Text = Strings.errorUnableToFixDocument + ex.Message;
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
                        rtbDisplay.AppendText(Strings.wHeaderLine + Environment.NewLine + "Fixed Document Location: " + StrDestFileName);
                    }
                    else
                    {
                        rtbDisplay.AppendText("Unable to fix document");
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
        private void toolStripButtonFixDoc_Click(object sender, EventArgs e)
        {
            bool corruptionFound = false;

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
                rtbDisplay.Text = sbFixes.ToString();
                rtbDisplay.AppendText("\r\n\r\nModified File Location = " + modifiedPath);
            }
            else
            {
                rtbDisplay.AppendText("No Corruption Found.");
            }
        }

        private void editToolStripMenuFindReplace_Click(object sender, EventArgs e)
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

        private void editToolStripMenuItemModifyContents_Click(object sender, EventArgs e)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                rtbDisplay.Clear();

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
                                    rtbDisplay.AppendText("Sheet: " + fds.sheetName + " Removed");
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
                                rtbDisplay.AppendText(tempFilePackageViewer + Strings.convertedTo + fNewName);
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
                                    rtbDisplay.AppendText("** Unable to convert file **");
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
                                    rtbDisplay.AppendText(Strings.fileConvertSuccessful);
                                    rtbDisplay.AppendText("File Location: " + strOutputFileName);
                                }
                                else
                                {
                                    rtbDisplay.AppendText("** File Is Not Open Xml Format (Strict) **");
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
                                rtbDisplay.AppendText("Unused Slide Layouts Deleted");
                            }
                        }

                        if (f.pptModCmd == AppUtilities.PowerPointModifyCmds.ResetBulletMargins)
                        {
                            using (PresentationDocument pDoc = PresentationDocument.Open(tempFilePackageViewer, true))
                            {
                                PowerPointFixes.ResetBulletMargins(pDoc);
                                rtbDisplay.AppendText("Bullet Margins Reset");
                            }
                        }

                        if (f.pptModCmd == AppUtilities.PowerPointModifyCmds.ConvertPptmToPptx)
                        {
                            string fNewName = Office.ConvertMacroEnabled2NonMacroEnabled(tempFilePackageViewer, Strings.oAppPowerPoint);
                            if (fNewName != string.Empty)
                            {
                                rtbDisplay.AppendText(tempFilePackageViewer + Strings.convertedTo + fNewName);
                            }
                        }

                        if (f.pptModCmd == AppUtilities.PowerPointModifyCmds.DelComments)
                        {
                            if (PowerPoint.DeleteComments(tempFilePackageViewer, string.Empty))
                            {
                                rtbDisplay.AppendText("Comments Removed");
                            }
                            else
                            {
                                rtbDisplay.AppendText("No Comments Removed");
                            }
                        }

                        if (f.pptModCmd == AppUtilities.PowerPointModifyCmds.RemovePIIOnSave)
                        {
                            using (PresentationDocument document = PresentationDocument.Open(tempFilePackageViewer, true))
                            {
                                document.PresentationPart.Presentation.RemovePersonalInfoOnSave = false;
                                document.PresentationPart.Presentation.Save();
                                rtbDisplay.AppendText("Remove PII On Save Disabled");
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
                            rtbDisplay.AppendText("Notes Page Size Reset");
                        }

                        if (f.pptModCmd == AppUtilities.PowerPointModifyCmds.CustomNotesPageReset)
                        {
                            PowerPointFixes.CustomResetNotesPageSize(tempFilePackageViewer);
                            rtbDisplay.AppendText("Notes Page Size Reset");
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

        private void editToolStripMenuItemRemoveCustomDocProps_Click(object sender, EventArgs e)
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

        private void editToolStripMenuItemRemoveCustomXml_Click(object sender, EventArgs e)
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

        private void fileToolStripMenuItemClose_Click(object sender, EventArgs e)
        {
            FileClose();
        }

        private void toolStripButtonFind_Click(object sender, EventArgs e)
        {
            if (rtbDisplay.Text.Length > 0)
            {
                FindText();
            }
        }

        private void mruToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (sender is not null) { OpenRecentFile(sender.ToString()!); }
        }

        private void toolStripButtonReplace_Click(object sender, EventArgs e)
        {
            ReplaceText();
        }

        private void mruToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (sender is not null) { OpenRecentFile(sender.ToString()!); }
        }

        private void mruToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (sender is not null) { OpenRecentFile(sender.ToString()!); }
        }

        private void mruToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            if (sender is not null) { OpenRecentFile(sender.ToString()!); }
        }

        private void mruToolStripMenuItem5_Click(object sender, EventArgs e)
        {
            if (sender is not null) { OpenRecentFile(sender.ToString()!); }
        }

        private void mruToolStripMenuItem6_Click(object sender, EventArgs e)
        {
            if (sender is not null) { OpenRecentFile(sender.ToString()!); }
        }

        private void mruToolStripMenuItem7_Click(object sender, EventArgs e)
        {
            if (sender is not null) { OpenRecentFile(sender.ToString()!); }
        }

        private void mruToolStripMenuItem8_Click(object sender, EventArgs e)
        {
            if (sender is not null) { OpenRecentFile(sender.ToString()!); }
        }

        private void mruToolStripMenuItem9_Click(object sender, EventArgs e)
        {
            if (sender is not null) { OpenRecentFile(sender.ToString()!); }
        }

        private void openErrorLogToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            AppUtilities.PlatformSpecificProcessStart(Strings.fLogFilePath);
        }

        private void wordDocumentRevisionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmRevisions frmRev = new FrmRevisions(tempFilePackageViewer)
            {
                Owner = this
            };
            frmRev.ShowDialog();
        }

        private void viewPartPropertiesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Compression Settings: " + partPropCompression);
            sb.AppendLine("Content Type: " + partPropContentType);
            MessageBox.Show(sb.ToString(), "Part Properties", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion
    }
}