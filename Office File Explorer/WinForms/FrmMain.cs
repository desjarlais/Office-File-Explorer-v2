// Open XML SDK refs
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
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

namespace Office_File_Explorer
{
    public partial class FrmMain : Form
    {
        // global variables
        private string findText;
        private string replaceText;
        private string fromChangeTemplate;
        
        private static string StrCopiedFileName = string.Empty;
        private static string StrOfficeApp = string.Empty;
        private static char PrevChar = '<';
        private bool IsRegularXmlTag;
        private bool IsFixed;
        private static string FixedFallback = string.Empty;
        private static string StrExtension = string.Empty;
        private static string StrDestFileName = string.Empty;

        // global lists
        private static List<string> corruptNodes = new List<string>();
        private static List<string> pParts = new List<string>();
        private List<string> oNumIdList = new List<string>();

        // corrupt doc xml node buffer
        static StringBuilder sbNodeBuffer = new StringBuilder();

        // enums
        public enum LogInfoType { ClearAndAdd, Append, TextOnly, InvalidFile, LogException, EmptyCount };

        public FrmMain()
        {
            InitializeComponent();

            // update title with version
            this.Text = Strings.oAppTitle + Strings.wMinusSign + Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyFileVersionAttribute>().Version;

            // make sure the log file is created
            if (!File.Exists(Strings.fLogFilePath))
            {
                File.Create(Strings.fLogFilePath);
            }

            // disable buttons
            DisableUI();
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

        public void DisableUI()
        {
            BtnViewContents.Enabled = false;
            BtnModifyContent.Enabled = false;
            BtnFixDocument.Enabled = false;
            BtnSearchAndReplace.Enabled = false;
            BtnCustomXml.Enabled = false;
            BtnDocProps.Enabled = false;
            BtnViewImages.Enabled = false;
            BtnFixCorruptDoc.Enabled = false;
            BtnViewCustomUI.Enabled = false;
            BtnValidateDoc.Enabled = false;
            BtnExcelSheetViewer.Enabled = false;
        }

        public void EnableUI()
        {
            BtnViewContents.Enabled = true;
            BtnModifyContent.Enabled = true;
            BtnFixDocument.Enabled = true;
            BtnSearchAndReplace.Enabled = true;
            BtnViewImages.Enabled = true;
            BtnCustomXml.Enabled = true;
            BtnDocProps.Enabled = true;
            BtnViewCustomUI.Enabled = true;
            BtnValidateDoc.Enabled = true;
        }

        public void CopyAllItems()
        {
            try
            {
                if (LstDisplay.Items.Count <= 0)
                {
                    return;
                }

                StringBuilder buffer = new StringBuilder();
                foreach (string s in LstDisplay.Items)
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

        public void OpenOfficeDocument()
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                OpenFileDialog fDialog = new OpenFileDialog
                {
                    Title = "Select Office Open Xml File.",
                    Filter = "Open XML Files | *.docx; *.dotx; *.docm; *.dotm; *.xlsx; *.xlsm; *.xlst; *.xltm; *.pptx; *.pptm; *.potx; *.potm",
                    RestoreDirectory = true,
                    InitialDirectory = @"%userprofile%"
                };

                if (fDialog.ShowDialog() == DialogResult.OK)
                {
                    lblFilePath.Text = fDialog.FileName.ToString();
                    if (!File.Exists(lblFilePath.Text))
                    {
                        LogInformation(LogInfoType.InvalidFile, Strings.fileDoesNotExist, string.Empty);
                    }
                    else
                    {
                        LstDisplay.Items.Clear();

                        // if the file doesn't start with PK, we can stop trying to process it
                        if (!FileUtilities.IsZipArchiveFile(lblFilePath.Text))
                        {
                            LstDisplay.Items.Add("Unable to open file, possible causes are:");
                            LstDisplay.Items.Add("  - file corruption");
                            LstDisplay.Items.Add("  - file encrypted");
                            LstDisplay.Items.Add("  - file password protected");
                            LstDisplay.Items.Add("  - not a valid Open Xml file");
                        }
                        else
                        {
                            // if the file does start with PK, check if it fails in the SDK
                            if (OpenWithSdk(lblFilePath.Text))
                            {
                                // set the file type
                                lblFileType.Text = StrOfficeApp;

                                // populate the parts
                                PopulatePackageParts();

                                // check if any zip items are corrupt
                                if (Properties.Settings.Default.CheckZipItemCorrupt == true && lblFileType.Text == Strings.oAppWord)
                                {
                                    if (Office.IsZippedFileCorrupt(lblFilePath.Text))
                                    {
                                        LstDisplay.Items.Add("Warning - One of the zipped items is corrupt.");
                                    }
                                }

                                // make a backup copy of the file and use it going forward
                                if (Properties.Settings.Default.BackupOnOpen == true)
                                {
                                    string backupFileName = AddTextToFileName(lblFilePath.Text, "(Backup)");
                                    File.Copy(lblFilePath.Text, backupFileName, true);
                                    lblFilePath.Text = backupFileName;
                                }
                            }
                            else
                            {
                                // if it failed the SDK, turn on Fix Corrupt Doc button only
                                DisableUI();
                                BtnFixCorruptDoc.Enabled = true;
                            }
                        }
                    }
                }
                else
                {
                    // user cancelled dialog, disable the UI and go back to the form
                    DisableUI();
                    lblFilePath.Text = string.Empty;
                    lblFileType.Text = string.Empty;
                    return;
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

        public void LogInformation(LogInfoType type, string output, string ex)
        {
            switch (type)
            {
                case LogInfoType.ClearAndAdd:
                    LstDisplay.Items.Clear();
                    LstDisplay.Items.Add(output);
                    break;
                case LogInfoType.Append:
                    LstDisplay.Items.Add(string.Empty);
                    LstDisplay.Items.Add(output);
                    break;
                case LogInfoType.InvalidFile:
                    LstDisplay.Items.Clear();
                    LstDisplay.Items.Add(Strings.invalidFile);
                    break;
                case LogInfoType.LogException:
                    LstDisplay.Items.Clear();
                    LstDisplay.Items.Add(output);
                    LstDisplay.Items.Add(ex);
                    FileUtilities.WriteToLog(Strings.fLogFilePath, DateTime.Now.ToString());
                    FileUtilities.WriteToLog(Strings.fLogFilePath, output);
                    FileUtilities.WriteToLog(Strings.fLogFilePath, ex);
                    break;
                case LogInfoType.EmptyCount:
                    LstDisplay.Items.Add(Strings.wNone);
                    break;
                default:
                    LstDisplay.Items.Add(output);
                    break;
            }
        }

        public void PopulatePackageParts()
        {
            pParts.Clear();
            using (FileStream zipToOpen = new FileStream(lblFilePath.Text, FileMode.Open, FileAccess.Read))
            {
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Read))
                {
                    foreach (ZipArchiveEntry zae in archive.Entries)
                    {
                        // add part to list
                        pParts.Add(zae.FullName + Strings.wColonBuffer + FileUtilities.SizeSuffix(zae.Length));
                    }
                }
            }
        }

        /// <summary>
        /// function to open the file in the SDK
        /// if the SDK fails to open the file, it is not a valid docx
        /// </summary>
        /// <param name="file">the path to the initial fix attempt</param>
        public bool OpenWithSdk(string file)
        {
            string body = string.Empty;
            bool fSuccess = false;

            try
            {
                Cursor = Cursors.WaitCursor;

                if (FileUtilities.GetAppFromFileExtension(file) == Strings.oAppWord)
                {
                    using (WordprocessingDocument document = WordprocessingDocument.Open(file, false))
                    {
                        // try to get the localname of the document.xml file, if it fails, it is not a Word file
                        StrOfficeApp = Strings.oAppWord;
                        body = document.MainDocumentPart.Document.LocalName;
                        fSuccess = true;
                    }
                }
                else if (FileUtilities.GetAppFromFileExtension(file) == Strings.oAppExcel)
                {
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(file, false))
                    {
                        // try to get the localname of the workbook.xml file if it fails, its not an Excel file
                        StrOfficeApp = Strings.oAppExcel;
                        body = document.WorkbookPart.Workbook.LocalName;
                        fSuccess = true;
                    }
                }
                else if (FileUtilities.GetAppFromFileExtension(file) == Strings.oAppPowerPoint)
                {
                    using (PresentationDocument document = PresentationDocument.Open(file, false))
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
            catch (OpenXmlPackageException ope)
            {
                // if the exception is related to invalid hyperlinks, use the FixInvalidUri method to change the file
                // once we change the copied file, we can open it in the SDK
                if (ope.ToString().Contains("Invalid Hyperlink"))
                {
                    // known issue in .NET with malformed hyperlinks causing SDK to throw during parse
                    // see UriFixHelper for more details
                    // get the path and make a new file name in the same directory
                    var StrCopyFileName = AddTextToFileName(lblFilePath.Text, Strings.wCopyFileParentheses);

                    // need a copy of the file to change the hyperlinks so we can open the modified version instead of the original
                    if (!File.Exists(StrCopyFileName))
                    {
                        File.Copy(lblFilePath.Text, StrCopyFileName);
                    }
                    else
                    {
                        StrCopyFileName = AddTextToFileName(lblFilePath.Text, Strings.wCopyFileParentheses + FileUtilities.GetRandomNumber().ToString());
                        File.Copy(lblFilePath.Text, StrCopyFileName);
                    }

                    // create the new file with the updated hyperlink
                    using (FileStream fs = new FileStream(StrCopyFileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    {
                        UriFix.FixInvalidUri(fs, brokenUri => FileUtilities.FixUri(brokenUri));
                    }

                    // now use the new file in the open logic from above
                    if (StrOfficeApp == Strings.oAppWord)
                    {
                        using (WordprocessingDocument document = WordprocessingDocument.Open(StrCopyFileName, false))
                        {
                            // try to get the localname of the document.xml file, if it fails, it is not a Word file
                            body = document.MainDocumentPart.Document.LocalName;
                            fSuccess = true;
                        }
                    }
                    else if (StrOfficeApp == Strings.oAppExcel)
                    {
                        using (SpreadsheetDocument document = SpreadsheetDocument.Open(StrCopyFileName, false))
                        {
                            // try to get the localname of the workbook.xml file if it fails, its not an Excel file
                            body = document.WorkbookPart.Workbook.LocalName;
                            fSuccess = true;
                        }
                    }
                    else if (StrOfficeApp == Strings.oAppPowerPoint)
                    {
                        using (PresentationDocument document = PresentationDocument.Open(StrCopyFileName, false))
                        {
                            // try to get the presentation.xml local name, if it fails it is not a PPT file
                            body = document.PresentationPart.Presentation.LocalName;
                            fSuccess = true;
                        }
                    }

                    // update the main form UI
                    lblFilePath.Text = StrCopyFileName;
                    StrCopiedFileName = StrCopyFileName;
                }
                else
                {
                    // unknown issue opening from .net
                    LogInformation(LogInfoType.LogException, "OpenWithSDK UriFix Error:", ope.Message);
                }
            }
            catch (InvalidOperationException ioe)
            {
                LogInformation(LogInfoType.LogException, "OpenWithSDK Error:", ioe.Message);
                LogInformation(LogInfoType.LogException, "OpenWithSDK Error:", ioe.StackTrace);
            }
            catch (Exception ex)
            {
                // if the file failed to open in the sdk, it is invalid or corrupt and we need to stop opening
                LogInformation(LogInfoType.LogException, "OpenWithSDK Error:", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }

            return fSuccess;
        }

        /// <summary>
        /// output content to the listbox
        /// </summary>
        /// <param name="output">the list of content to display</param>
        /// <param name="type">the type of content to display</param>
        public void DisplayListContents(List<string> output, string type)
        {
            // no content to display
            if (output.Count == 0)
            {
                LogInformation(LogInfoType.EmptyCount, type, string.Empty);
                LstDisplay.Items.Add(string.Empty);
                return;
            }

            // if we have any values, display them
            foreach (string s in output)
            {
                LstDisplay.Items.Add(Strings.wTripleSpace + s);
            }

            LstDisplay.Items.Add(string.Empty);
        }

        public static void Node(char input)
        {
            sbNodeBuffer.Append(input);
        }

        /// <summary>
        /// this function loops through all nodes parsed out from Step 1
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

        public static List<string> CfpList(CustomFilePropertiesPart part)
        {
            List<string> val = new List<string>();
            foreach (CustomDocumentProperty cdp in part.RootElement)
            {
                val.Add(cdp.Name + Strings.wColonBuffer + cdp.InnerText);
            }
            return val;
        }

        public static string AddTextToFileName(string fileName, string TextToAdd)
        {
            string dir = Path.GetDirectoryName(fileName) + "\\";
            StrExtension = Path.GetExtension(fileName);
            string newFileName = dir + Path.GetFileNameWithoutExtension(fileName) + TextToAdd + StrExtension;
            return newFileName;
        }

        public static void AppExitWork()
        {
            try
            {
                if (Properties.Settings.Default.DeleteCopiesOnExit == true && File.Exists(StrCopiedFileName))
                {
                    File.Delete(StrCopiedFileName);
                }

                Properties.Settings.Default.Save();
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

        private void BtnViewContents_Click(object sender, EventArgs e)
        {
            try
            {
                LstDisplay.Items.Clear();

                if (StrOfficeApp == Strings.oAppWord)
                {
                    using (var f = new FrmWordCommands(lblFilePath.Text))
                    {
                        var result = f.ShowDialog();

                        Cursor = Cursors.WaitCursor;
                        AppUtilities.WordViewCmds cmds = f.wdCmds;
                        AppUtilities.OfficeViewCmds oCmds = f.offCmds;

                        if (cmds.HasFlag(AppUtilities.WordViewCmds.ContentControls))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wContentControls + Strings.wHeadingEnd);
                            DisplayListContents(Word.LstContentControls(lblFilePath.Text), Strings.wContentControls);
                        }

                        if (cmds.HasFlag(AppUtilities.WordViewCmds.Styles))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wStyles + Strings.wHeadingEnd);
                            DisplayListContents(Word.LstStyles(lblFilePath.Text), Strings.wStyles);
                        }

                        if (cmds.HasFlag(AppUtilities.WordViewCmds.Hyperlinks))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wHyperlinks + Strings.wHeadingEnd);
                            DisplayListContents(Word.LstHyperlinks(lblFilePath.Text), Strings.wHyperlinks);
                        }

                        if (cmds.HasFlag(AppUtilities.WordViewCmds.ListTemplates))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wListTemplates + Strings.wHeadingEnd);
                            DisplayListContents(Word.LstListTemplates(lblFilePath.Text, false), Strings.wListTemplates);
                        }

                        if (cmds.HasFlag(AppUtilities.WordViewCmds.Fonts))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wFonts + Strings.wHeadingEnd);
                            DisplayListContents(Word.LstFonts(lblFilePath.Text), Strings.wFonts);
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wRunFonts + Strings.wHeadingEnd);
                            DisplayListContents(Word.LstRunFonts(lblFilePath.Text), Strings.wRunFonts);
                        }

                        if (cmds.HasFlag(AppUtilities.WordViewCmds.Footnotes))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wFootnotes + Strings.wHeadingEnd);
                            DisplayListContents(Word.LstFootnotes(lblFilePath.Text), Strings.wFootnotes);
                        }

                        if (cmds.HasFlag(AppUtilities.WordViewCmds.Endnotes))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wEndnotes + Strings.wHeadingEnd);
                            DisplayListContents(Word.LstEndnotes(lblFilePath.Text), Strings.wEndnotes);
                        }

                        if (cmds.HasFlag(AppUtilities.WordViewCmds.DocumentProperties))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wDocProps + Strings.wHeadingEnd);
                            DisplayListContents(Word.LstDocProps(lblFilePath.Text), Strings.wDocProps);
                        }

                        if (cmds.HasFlag(AppUtilities.WordViewCmds.Bookmarks))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wBookmarks + Strings.wHeadingEnd);
                            DisplayListContents(Word.LstBookmarks(lblFilePath.Text), Strings.wBookmarks);
                        }

                        if (cmds.HasFlag(AppUtilities.WordViewCmds.Comments))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wComments + Strings.wHeadingEnd);
                            DisplayListContents(Word.LstComments(lblFilePath.Text), Strings.wComments);
                        }

                        if (cmds.HasFlag(AppUtilities.WordViewCmds.FieldCodes))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wFldCodes + Strings.wHeadingEnd);
                            DisplayListContents(Word.LstFieldCodes(lblFilePath.Text), Strings.wFldCodes);
                            LstDisplay.Items.Add(" ** Header Field Codes **");
                            DisplayListContents(Word.LstFieldCodesInHeader(lblFilePath.Text), Strings.wFldCodes);
                            LstDisplay.Items.Add(" ** Footer Field Codes **");
                            DisplayListContents(Word.LstFieldCodesInFooter(lblFilePath.Text), Strings.wFldCodes);
                        }

                        if (cmds.HasFlag(AppUtilities.WordViewCmds.Tables))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wTables + Strings.wHeadingEnd);
                            DisplayListContents(Word.LstTables(lblFilePath.Text), Strings.wTables);
                        }

                        // display Office features
                        if (oCmds.HasFlag(AppUtilities.OfficeViewCmds.OleObjects))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wEmbeddedObjects + Strings.wHeadingEnd);
                            DisplayListContents(Office.GetEmbeddedObjectProperties(lblFilePath.Text, Strings.oAppWord), Strings.wEmbeddedObjects);
                        }

                        if (oCmds.HasFlag(AppUtilities.OfficeViewCmds.Shapes))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wShapes + Strings.wHeadingEnd);
                            DisplayListContents(Office.GetShapes(lblFilePath.Text, Strings.oAppWord), Strings.wShapes);
                        }

                        if (oCmds.HasFlag(AppUtilities.OfficeViewCmds.PackageParts))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wPackageParts + Strings.wHeadingEnd);
                            pParts.Sort();
                            DisplayListContents(pParts, Strings.wPackageParts);
                        }
                    }
                }
                else if (StrOfficeApp == Strings.oAppExcel)
                {
                    using (var f = new FrmExcelCommands())
                    {
                        var result = f.ShowDialog();

                        Cursor = Cursors.WaitCursor;
                        AppUtilities.ExcelViewCmds cmds = f.xlCmds;
                        AppUtilities.OfficeViewCmds oCmds = f.offCmds;

                        if (cmds.HasFlag(AppUtilities.ExcelViewCmds.Links))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wLinks + Strings.wHeadingEnd);
                            DisplayListContents(Excel.GetLinks(lblFilePath.Text, true), Strings.wLinks);
                        }

                        if (cmds.HasFlag(AppUtilities.ExcelViewCmds.Comments))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wComments + Strings.wHeadingEnd);
                            DisplayListContents(Excel.GetComments(lblFilePath.Text), Strings.wComments);
                        }

                        if (cmds.HasFlag(AppUtilities.ExcelViewCmds.Hyperlinks))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wHyperlinks + Strings.wHeadingEnd);
                            DisplayListContents(Excel.GetHyperlinks(lblFilePath.Text), Strings.wHyperlinks);
                        }

                        if (cmds.HasFlag(AppUtilities.ExcelViewCmds.WorksheetInfo))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wWorksheetInfo + Strings.wHeadingEnd);
                            DisplayListContents(Excel.GetSheetInfo(lblFilePath.Text), Strings.wWorksheetInfo);
                        }

                        if (cmds.HasFlag(AppUtilities.ExcelViewCmds.SharedStrings))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wSharedStrings + Strings.wHeadingEnd);
                            DisplayListContents(Excel.GetSharedStrings(lblFilePath.Text), Strings.wSharedStrings);
                        }

                        if (cmds.HasFlag(AppUtilities.ExcelViewCmds.DefinedNames))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wDefinedNames + Strings.wHeadingEnd);
                            DisplayListContents(Excel.GetDefinedNames(lblFilePath.Text), Strings.wDefinedNames);
                        }

                        if (cmds.HasFlag(AppUtilities.ExcelViewCmds.Connections))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wConnections + Strings.wHeadingEnd);
                            DisplayListContents(Excel.GetConnections(lblFilePath.Text), Strings.wConnections);
                        }

                        if (cmds.HasFlag(AppUtilities.ExcelViewCmds.HiddenRowsCols))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wHiddenRowCol + Strings.wHeadingEnd);
                            DisplayListContents(Excel.GetHiddenRowCols(lblFilePath.Text), Strings.wHiddenRowCol);
                        }

                        // display Office features
                        if (oCmds.HasFlag(AppUtilities.OfficeViewCmds.OleObjects))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wEmbeddedObjects + Strings.wHeadingEnd);
                            DisplayListContents(Office.GetEmbeddedObjectProperties(lblFilePath.Text, Strings.oAppExcel), Strings.wEmbeddedObjects);
                        }

                        if (oCmds.HasFlag(AppUtilities.OfficeViewCmds.Shapes))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wShapes + Strings.wHeadingEnd);
                            DisplayListContents(Office.GetShapes(lblFilePath.Text, Strings.oAppExcel), Strings.wShapes);
                        }

                        if (oCmds.HasFlag(AppUtilities.OfficeViewCmds.PackageParts))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wPackageParts + Strings.wHeadingEnd);
                            pParts.Sort();
                            DisplayListContents(pParts, Strings.wPackageParts);
                        }
                    }
                }
                else if (StrOfficeApp == Strings.oAppPowerPoint)
                {
                    using (var f = new FrmPPTCommands())
                    {
                        var result = f.ShowDialog();

                        Cursor = Cursors.WaitCursor;
                        AppUtilities.PowerPointViewCmds cmds = f.pptCmds;
                        AppUtilities.OfficeViewCmds oCmds = f.offCmds;

                        if (cmds.HasFlag(AppUtilities.PowerPointViewCmds.Hyperlinks))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wHyperlinks + Strings.wHeadingEnd);
                            DisplayListContents(PowerPoint.GetHyperlinks(lblFilePath.Text), Strings.wHyperlinks);
                        }

                        if (cmds.HasFlag(AppUtilities.PowerPointViewCmds.Comments))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wComments + Strings.wHeadingEnd);
                            DisplayListContents(PowerPoint.GetComments(lblFilePath.Text), Strings.wComments);
                        }

                        if (cmds.HasFlag(AppUtilities.PowerPointViewCmds.SlideText))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wSlideText + Strings.wHeadingEnd);
                            DisplayListContents(PowerPoint.GetSlideText(lblFilePath.Text), Strings.wSlideText);
                        }

                        if (cmds.HasFlag(AppUtilities.PowerPointViewCmds.SlideTitles))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wSlideText + Strings.wHeadingEnd);
                            DisplayListContents(PowerPoint.GetSlideTitles(lblFilePath.Text), Strings.wSlideText);
                        }

                        if (cmds.HasFlag(AppUtilities.PowerPointViewCmds.SlideTransitions))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wSlideTransitions + Strings.wHeadingEnd);
                            DisplayListContents(PowerPoint.GetSlideTransitions(lblFilePath.Text), Strings.wSlideTransitions);
                        }

                        if (cmds.HasFlag(AppUtilities.PowerPointViewCmds.Fonts))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wFonts + Strings.wHeadingEnd);
                            DisplayListContents(PowerPoint.GetFonts(lblFilePath.Text), Strings.wFonts);
                        }

                        // display Office features
                        if (oCmds.HasFlag(AppUtilities.OfficeViewCmds.OleObjects))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wEmbeddedObjects + Strings.wHeadingEnd);
                            DisplayListContents(Office.GetEmbeddedObjectProperties(lblFilePath.Text, Strings.oAppPowerPoint), Strings.wEmbeddedObjects);
                        }

                        if (oCmds.HasFlag(AppUtilities.OfficeViewCmds.Shapes))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wShapes + Strings.wHeadingEnd);
                            DisplayListContents(Office.GetShapes(lblFilePath.Text, Strings.oAppPowerPoint), Strings.wShapes);
                        }

                        if (oCmds.HasFlag(AppUtilities.OfficeViewCmds.PackageParts))
                        {
                            LstDisplay.Items.Add(Strings.wHeadingBegin + Strings.wPackageParts + Strings.wHeadingEnd);
                            pParts.Sort();
                            DisplayListContents(pParts, Strings.wPackageParts);
                        }
                    }
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

        private void OpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DisableUI();
            EnableUI();
            OpenOfficeDocument();

            if (lblFileType.Text == Strings.oAppExcel)
            {
                BtnExcelSheetViewer.Enabled = true;
            }
        }

        private void BtnSearchAndReplace_Click(object sender, EventArgs e)
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

            Office.SearchAndReplace(lblFilePath.Text, findText, replaceText);
            LstDisplay.Items.Clear();
            LstDisplay.Items.Add("** Search and Replace Finished **");
        }

        private void BtnFixDocument_Click(object sender, EventArgs e)
        {
            using (FrmFixDocument f = new FrmFixDocument(lblFilePath.Text, lblFileType.Text))
            {
                f.ShowDialog();

                if (f.isFileFixed == true)
                {
                    if (f.corruptionChecked == "All")
                    {
                        int count = 0;
                        foreach (var s in f.featureFixed)
                        {
                            count++;
                            LstDisplay.Items.Add(count + Strings.wPeriod + s);
                        }
                    }
                    else
                    {
                        LstDisplay.Items.Add("Corrupt " + f.corruptionChecked + " Found - " + "Document Fixed");
                    }
                    
                    return;
                }
                else
                {
                    // if it wasn't cancelled, no corruption was found
                    // if it was cancelled, do nothing
                    if (f.corruptionChecked != "Cancel")
                    {
                        LstDisplay.Items.Clear();
                        LstDisplay.Items.Add("No Corruption Found");
                    }
                }
            }
        }

        private void SettingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmSettings form = new FrmSettings();
            form.Show();
        }

        private void BtnModifyContent_Click(object sender, EventArgs e)
        {
            try
            {
                if (StrOfficeApp == Strings.oAppWord)
                {
                    using (var f = new FrmWordModify())
                    {
                        var result = f.ShowDialog();

                        Cursor = Cursors.WaitCursor;

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.DelHF)
                        {
                            if (Word.RemoveHeadersFooters(lblFilePath.Text) == true)
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
                            if (Word.RemoveComments(lblFilePath.Text) == true)
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
                            if (Word.RemoveEndnotes(lblFilePath.Text) == true)
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
                            if (Word.RemoveFootnotes(lblFilePath.Text) == true)
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
                            oNumIdList = Word.LstListTemplates(lblFilePath.Text, true);
                            foreach (object orphanLT in oNumIdList)
                            {
                                Word.RemoveListTemplatesNumId(lblFilePath.Text, orphanLT.ToString());
                            }
                            LogInformation(LogInfoType.ClearAndAdd, "Unused List Templates Removed", string.Empty);
                        }

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.DelOrphanStyles)
                        {
                            DisplayListContents(Word.RemoveUnusedStyles(lblFilePath.Text), "Styles");
                        }

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.DelHiddenTxt)
                        {
                            if (Word.DeleteHiddenText(lblFilePath.Text) == true)
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
                            if (Word.RemoveBreaks(lblFilePath.Text))
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
                            FrmPrintOrientation pFrm = new FrmPrintOrientation(lblFilePath.Text)
                            {
                                Owner = this
                            };
                            pFrm.ShowDialog();
                        }

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.AcceptRevisions)
                        {
                            foreach (var s in Word.AcceptRevisions(lblFilePath.Text, Strings.allAuthors))
                            {
                                LstDisplay.Items.Add(s);
                            }
                        }

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.ChangeDefaultTemplate)
                        {
                            bool isFileChanged = false;
                            string attachedTemplateId = "rId1";
                            string filePath = string.Empty;

                            using (WordprocessingDocument document = WordprocessingDocument.Open(lblFilePath.Text, true))
                            {
                                DocumentSettingsPart dsp = document.MainDocumentPart.DocumentSettingsPart;

                                // if the external rel exists, we need to pull the rid and old uri
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
                                    LstDisplay.Items.Add("** Attached Template Path Changed **");
                                    document.MainDocumentPart.Document.Save();
                                }
                                else
                                {
                                    LstDisplay.Items.Add("** No Changes Made To Attached Template **");
                                }
                            }
                        }

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.ConvertDocmToDocx)
                        {
                            string fNewName = Office.ConvertMacroEnabled2NonMacroEnabled(lblFilePath.Text, Strings.oAppWord);
                            LstDisplay.Items.Clear();
                            if (fNewName != string.Empty)
                            {
                                LstDisplay.Items.Add(lblFilePath.Text + Strings.convertedTo + fNewName);
                            }
                        }

                        if (f.wdModCmd == AppUtilities.WordModifyCmds.RemovePII)
                        {
                            using (WordprocessingDocument document = WordprocessingDocument.Open(lblFilePath.Text, true))
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
                            if (Word.RemoveCustomTitleProp(lblFilePath.Text))
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
                            if (WordFixes.FixContentControlNamespaces(lblFilePath.Text))
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
                            if (Word.RemoveBookmarks(lblFilePath.Text))
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "Bookmarks Deleted", string.Empty);
                            }
                            else
                            {
                                LogInformation(LogInfoType.ClearAndAdd, "No Bookmarks In Document", string.Empty);
                            }
                        }
                    }
                }
                else if (StrOfficeApp == Strings.oAppExcel)
                {
                    using (var f = new FrmExcelModify())
                    {
                        var result = f.ShowDialog();

                        Cursor = Cursors.WaitCursor;

                        if (f.xlModCmd == AppUtilities.ExcelModifyCmds.DelLink)
                        {
                            using (var fDelLink = new FrmExcelDelLink(lblFilePath.Text))
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
                            if (Excel.RemoveHyperlinks(lblFilePath.Text) == true)
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
                            if (Excel.RemoveLinks(lblFilePath.Text) == true)
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
                            using (var fds = new FrmDeleteSheet(lblFilePath.Text))
                            {
                                fds.ShowDialog();

                                if (fds.sheetName != string.Empty)
                                {
                                    LstDisplay.Items.Add("Sheet: " + fds.sheetName + " Removed");
                                }
                            }
                        }

                        if (f.xlModCmd == AppUtilities.ExcelModifyCmds.DelComments)
                        {
                            if (Excel.RemoveComments(lblFilePath.Text) == true)
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
                            string fNewName = Office.ConvertMacroEnabled2NonMacroEnabled(lblFilePath.Text, Strings.oAppExcel);
                            LstDisplay.Items.Clear();
                            if (fNewName != string.Empty)
                            {
                                LstDisplay.Items.Add(lblFilePath.Text + Strings.convertedTo + fNewName);
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
                                    LstDisplay.Items.Add("** Unable to convert file **");
                                    return;
                                }

                                // check if the file is strict, no changes are made to the file yet
                                bool isStrict = false;

                                using (Package package = Package.Open(lblFilePath.Text, FileMode.Open, FileAccess.Read))
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
                                    string strOriginalFile = lblFilePath.Text;
                                    string strOutputPath = Path.GetDirectoryName(strOriginalFile) + "\\";
                                    string strFileExtension = Path.GetExtension(strOriginalFile);
                                    string strOutputFileName = strOutputPath + Path.GetFileNameWithoutExtension(strOriginalFile) + Strings.wFixedFileParentheses + strFileExtension;

                                    // run the command to convert the file "excelcnv.exe -nme -oice "strict-file-path" "converted-file-path""
                                    string cParams = " -nme -oice " + Strings.dblQuote + lblFilePath.Text + Strings.dblQuote + Strings.wSpaceChar + Strings.dblQuote + strOutputFileName + Strings.dblQuote;
                                    var proc = Process.Start(excelcnvPath, cParams);
                                    proc.Close();
                                    LstDisplay.Items.Add(Strings.fileConvertSuccessful);
                                    LstDisplay.Items.Add("File Location: " + strOutputFileName);
                                }
                                else
                                {
                                    LstDisplay.Items.Add("** File Is Not Open Xml Format (Strict) **");
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
                    }
                }
                else if (StrOfficeApp == Strings.oAppPowerPoint)
                {
                    using (var f = new FrmPowerPointModify())
                    {
                        var result = f.ShowDialog();

                        Cursor = Cursors.WaitCursor;

                        if (f.pptModCmd == AppUtilities.PowerPointModifyCmds.ConvertPptmToPptx)
                        {
                            string fNewName = Office.ConvertMacroEnabled2NonMacroEnabled(lblFilePath.Text, Strings.oAppPowerPoint);
                            LstDisplay.Items.Clear();
                            if (fNewName != string.Empty)
                            {
                                LstDisplay.Items.Add(lblFilePath.Text + Strings.convertedTo + fNewName);
                            }
                        }

                        if (f.pptModCmd == AppUtilities.PowerPointModifyCmds.DelComments)
                        {
                            if (PowerPoint.DeleteComments(lblFilePath.Text, string.Empty))
                            {
                                LstDisplay.Items.Add("Comments Removed");
                            }
                            else
                            {
                                LstDisplay.Items.Add("No Comments Removed");
                            }
                        }

                        if (f.pptModCmd == AppUtilities.PowerPointModifyCmds.RemovePIIOnSave)
                        {
                            using (PresentationDocument document = PresentationDocument.Open(lblFilePath.Text, true))
                            {
                                document.PresentationPart.Presentation.RemovePersonalInfoOnSave = false;
                                document.PresentationPart.Presentation.Save();
                            }
                        }

                        if (f.pptModCmd == AppUtilities.PowerPointModifyCmds.MoveSlide)
                        {
                            FrmMoveSlide mvFrm = new FrmMoveSlide(lblFilePath.Text)
                            {
                                Owner = this
                            };
                            mvFrm.ShowDialog();
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

        private void BatchFileProcessingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmBatch bFrm = new FrmBatch()
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

        private void OpenErrorLogToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AppUtilities.PlatformSpecificProcessStart(Strings.fLogFilePath);
        }

        private void BtnViewImages_Click(object sender, EventArgs e)
        {
            FrmViewImages imgFrm = new FrmViewImages(lblFilePath.Text, lblFileType.Text)
            {
                Owner = this
            };
            imgFrm.ShowDialog();
        }

        private void BtnDocProps_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                LstDisplay.Items.Clear();

                if (lblFileType.Text == Strings.oAppWord)
                {
                    using (WordprocessingDocument document = WordprocessingDocument.Open(lblFilePath.Text, false))
                    {
                        AddCustomDocPropsToList(document.CustomFilePropertiesPart);
                    }
                }
                else if (lblFileType.Text == Strings.oAppExcel)
                {
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(lblFilePath.Text, false))
                    {
                        AddCustomDocPropsToList(document.CustomFilePropertiesPart);
                    }
                }
                else if (lblFileType.Text == Strings.oAppPowerPoint)
                {
                    using (PresentationDocument document = PresentationDocument.Open(lblFilePath.Text, false))
                    {
                        AddCustomDocPropsToList(document.CustomFilePropertiesPart);
                    }
                }
                else
                {
                    FileUtilities.WriteToLog(Strings.fLogFilePath, "BtnListCustomDocProps - unknown app");
                    return;
                }
            }
            catch (IOException ioe)
            {
                LogInformation(LogInfoType.LogException, Strings.wCustomDocProps, ioe.Message);
            }
            catch (Exception ex)
            {
                LogInformation(LogInfoType.LogException, "BtnListCustomDocProps Error: ", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        public void AddCustomDocPropsToList(CustomFilePropertiesPart cfp)
        {
            if (cfp is null)
            {
                LogInformation(LogInfoType.EmptyCount, Strings.wCustomDocProps, string.Empty);
                return;
            }

            int count = 0;

            foreach (string v in CfpList(cfp))
            {
                count++;
                LstDisplay.Items.Add(count + Strings.wPeriod + v);
            }

            if (count == 0)
            {
                LogInformation(LogInfoType.EmptyCount, Strings.wCustomDocProps, string.Empty);
            }
        }

        private void BtnCustomXml_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                using (var f = new FrmCustomXmlViewer(lblFilePath.Text, lblFileType.Text))
                {
                    var result = f.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogInfoType.LogException, "BtnViewCustomXml Error: ", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void ClipboardViewerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmClipboardViewer cFrm = new FrmClipboardViewer()
            {
                Owner = this
            };
            cFrm.ShowDialog();
        }

        private void BtnFixCorruptDoc_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                StrDestFileName = AddTextToFileName(lblFilePath.Text, "(Fixed)");
                bool isXmlException = false;
                string strDocText = string.Empty;
                IsFixed = false;

                // check if file we are about to copy exists and append a number so it is unique
                if (File.Exists(StrDestFileName))
                {
                    StrDestFileName = AddTextToFileName(StrDestFileName, FileUtilities.GetRandomNumber().ToString());
                }

                LstDisplay.Items.Clear();

                if (StrExtension == ".docx")
                {
                    if ((File.GetAttributes(lblFilePath.Text) & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                    {
                        LstDisplay.Items.Add("ERROR: File is Read-Only.");
                        return;
                    }
                    else
                    {
                        File.Copy(lblFilePath.Text, StrDestFileName);
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
                            XmlDocument xdoc = new XmlDocument();

                            try
                            {
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
                                                        LstDisplay.Items.Add(Strings.invalidTag + m.Value);
                                                        LstDisplay.Items.Add(Strings.replacedWith + ValidXmlTags.StrValidVshapegroup);
                                                    }
                                                    else
                                                    {
                                                        strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidVshape);
                                                        LstDisplay.Items.Add(Strings.invalidTag + m.Value);
                                                        LstDisplay.Items.Add(Strings.replacedWith + ValidXmlTags.StrValidVshape);
                                                    }
                                                    break;

                                                case InvalidXmlTags.StrInvalidOmathWps:
                                                    strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidomathwps);
                                                    LstDisplay.Items.Add(Strings.invalidTag + m.Value);
                                                    LstDisplay.Items.Add(Strings.replacedWith + ValidXmlTags.StrValidomathwps);
                                                    break;

                                                case InvalidXmlTags.StrInvalidOmathWpg:
                                                    strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidomathwpg);
                                                    LstDisplay.Items.Add(Strings.invalidTag + m.Value);
                                                    LstDisplay.Items.Add(Strings.replacedWith + ValidXmlTags.StrValidomathwpg);
                                                    break;

                                                case InvalidXmlTags.StrInvalidOmathWpc:
                                                    strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidomathwpc);
                                                    LstDisplay.Items.Add(Strings.invalidTag + m.Value);
                                                    LstDisplay.Items.Add(Strings.replacedWith + ValidXmlTags.StrValidomathwpc);
                                                    break;

                                                case InvalidXmlTags.StrInvalidOmathWpi:
                                                    strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidomathwpi);
                                                    LstDisplay.Items.Add(Strings.invalidTag + m.Value);
                                                    LstDisplay.Items.Add(Strings.replacedWith + ValidXmlTags.StrValidomathwpi);
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
                                                            LstDisplay.Items.Add(Strings.invalidTag + m.Value);
                                                            LstDisplay.Items.Add(Strings.replacedWith + ValidXmlTags.StrValidMcChoice4);
                                                            break;
                                                        }

                                                        // replace mc:choice and hold onto the tag that follows
                                                        strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrValidMcChoice3 + m.Groups[2].Value);
                                                        LstDisplay.Items.Add(Strings.invalidTag + m.Value);
                                                        LstDisplay.Items.Add(Strings.replacedWith + ValidXmlTags.StrValidMcChoice3 + m.Groups[2].Value);
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
                                                            LstDisplay.Items.Add(Strings.invalidTag + m.Value);
                                                            LstDisplay.Items.Add(Strings.replacedWith + "Fallback tag deleted.");
                                                            break;
                                                        }

                                                        // if there is no closing fallback tag, we can replace the match with the omitFallback valid tags
                                                        // then we need to also add the trailing tag, since it's always different but needs to stay in the file
                                                        strDocText = strDocText.Replace(m.Value, ValidXmlTags.StrOmitFallback + m.Groups[2].Value);
                                                        LstDisplay.Items.Add(Strings.invalidTag + m.Value);
                                                        LstDisplay.Items.Add(Strings.replacedWith + ValidXmlTags.StrOmitFallback + m.Groups[2].Value);
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        // leaving this open for future checks
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
                                                case '<':
                                                    // if we haven't hit a close, but hit another '<' char
                                                    // we are not a true open tag so add it like a regular char
                                                    if (sbNodeBuffer.Length > 0)
                                                    {
                                                        corruptNodes.Add(sbNodeBuffer.ToString());
                                                        sbNodeBuffer.Clear();
                                                    }
                                                    Node(charEnum.Current);
                                                    break;

                                                case '>':
                                                    // there are 2 ways to close out a tag
                                                    // 1. self contained tag like <w:sz w:val="28"/>
                                                    // 2. standard xml <w:t>test</w:t>
                                                    // if previous char is '/', then we are an end tag
                                                    if (PrevChar == '/' || IsRegularXmlTag)
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
                                                    if (PrevChar == '<' && charEnum.Current == '/')
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
                                        LstDisplay.Items.Add(" ## No Corruption Found  ## ");
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
                // list out the possible reasons for this type of exception
                LstDisplay.Items.Add(Strings.errorUnableToFixDocument);
                LstDisplay.Items.Add("   Possible Causes:");
                LstDisplay.Items.Add("      - File may be password protected");
                LstDisplay.Items.Add("      - File was renamed to the .docx extension, but is not an actual .docx file");
                LstDisplay.Items.Add("      - Error = " + ffe.Message);
            }
            catch (Exception ex)
            {
                LstDisplay.Items.Add(Strings.errorUnableToFixDocument + ex.Message);
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
                        LstDisplay.Items.Add("-------------------------------------------------------------");
                        LstDisplay.Items.Add("Fixed Document Location: " + StrDestFileName);
                    }
                    else
                    {
                        LstDisplay.Items.Add("Unable to fix document");
                    }
                }

                // reset the globals
                IsFixed = false;
                IsRegularXmlTag = false;
                FixedFallback = string.Empty;
                StrExtension = string.Empty;
                StrDestFileName = string.Empty;
                PrevChar = '<';

                Cursor = Cursors.Default;
            }
        }

        private void CopySelectedLineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (LstDisplay.Items.Count > 0)
                {
                    Clipboard.SetText(LstDisplay.SelectedItem.ToString());
                }
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

        private void BtnViewCustomUI_Click(object sender, EventArgs e)
        {
            using (var f = new FrmCustomUI(lblFilePath.Text))
            {
                var result = f.ShowDialog();
            }
        }

        private void Base64DecoderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FrmBase64 b64Frm = new FrmBase64()
            {
                Owner = this
            };
            b64Frm.ShowDialog();
        }

        private void BtnValidateDoc_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                LstDisplay.Items.Clear();

                // check for xml validation errors
                if (lblFileType.Text == Strings.oAppWord)
                {
                    using (WordprocessingDocument myDoc = WordprocessingDocument.Open(lblFilePath.Text, false))
                    {
                        DisplayListContents(Office.DisplayValidationErrorInformation(myDoc), Strings.validationErrors);
                    }
                }
                else if (lblFileType.Text == Strings.oAppExcel)
                {
                    using (SpreadsheetDocument myDoc = SpreadsheetDocument.Open(lblFilePath.Text, false))
                    {
                        DisplayListContents(Office.DisplayValidationErrorInformation(myDoc), Strings.validationErrors);
                    }
                }
                else if (lblFileType.Text == Strings.oAppPowerPoint)
                {
                    using (PresentationDocument myDoc = PresentationDocument.Open(lblFilePath.Text, false))
                    {
                        DisplayListContents(Office.DisplayValidationErrorInformation(myDoc), Strings.validationErrors);
                    }
                }
                else
                {
                    throw new Exception();
                }
            }
            catch (Exception ex)
            {
                LogInformation(LogInfoType.LogException, "BtnValidateFile_Click Error", ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void openFileBackupFolderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AppUtilities.PlatformSpecificProcessStart(Path.GetDirectoryName(Application.LocalUserAppDataPath));
        }

        private void BtnExcelSheetViewer_Click(object sender, EventArgs e)
        {
            using (var f = new FrmSheetViewer(lblFilePath.Text))
            {
                var result = f.ShowDialog();
            }
        }
    }
}
