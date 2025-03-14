﻿using System;
using System.IO;
using System.Windows.Forms;

namespace Office_File_Explorer.Helpers
{
    public class Strings
    {
        // char values
        public const char chDblQuote = '"';
        public const char chLessThan = '<';
        public const char chGreaterThan = '>';
        public const char chBackslash = '/';
        public const char chPipe = '|';
        public const char chUnderscore = '_';

        // byte values
        public const string bZero = "0";
        public const string bP = "80";
        public const string bK = "75";
        public const string b3 = "3";
        public const string b4 = "4";
        public const string bDataDescriptor = "8";

        // scintilla strings
        public const string sciErrorInRegex = "Error In Regex";
        public const string sciMatchNotFound = "Match Not Found";
        public const string sciMatchWrappedBeginningDocument = "Search continued at the beginning of the document";
        public const string sciMatchWrappedBeginningSelection = "Search continued at the top of the selection";
        public const string sciMatchWrappedEndDocument = "Search continued at the end of the document";
        public const string sciMatchWrappedEndSelection = "Search continued at the end of the selection";
        public const string sciTotalFound = "Total Found:";
        public const string sciTotalReplaced = "Total Replaced:";
        public const string sciMustBeInRange = "Must Be In Range";
        public const string sciMustBeNumeric = "Must Be Numeric";

        // Office file extensions
        public const string docxFileExt = ".docx";
        public const string docmFileExt = ".docm";
        public const string dotxFileExt = ".dotx";
        public const string dotmFileExt = ".dotm";
        public const string xlsxFileExt = ".xlsx";
        public const string xlsmFileExt = ".xlsm";
        public const string xltxFileExt = ".xltx";
        public const string xltmFileExt = ".xltm";
        public const string pptxFileExt = ".pptx";
        public const string pptmFileExt = ".pptm";
        public const string potxFileExt = ".potx";
        public const string potmFileExt = ".potm";
        public const string msgFileExt = ".msg";

        // non-Office file extensions
        public const string emfFileExt = ".emf";
        public const string svgFileExt = ".svg";

        // global app words
        public const string oAppWord = "Word";
        public const string oAppExcel = "Excel";
        public const string oAppPowerPoint = "PowerPoint";
        public const string oAppOutlook = "Outlook";
        public const string oAppUnknown = "Unknown";
        public const string oAppTitle = "Office File Explorer v2";
        public const string wCopyFileParentheses = "(Copy)";
        public const string wModified = "(Modifed)";
        public const string wCancel = "Cancel";
        public const string wEmpty = "empty";
        public const string wColonBuffer = " : ";
        public const string wErrorText = "Error: ";
        public const string wFixedFileParentheses = "(Fixed)";
        public const string wFixedWithSpace = " Fixed";
        public const string wEnd = "end";
        public const string wBegin = "begin";
        public const string wColon = ": ";
        public const string wNumId = ". numId = ";
        public const string wPeriod = ". ";
        public const string wArrow = " --> ";
        public const string wArrowOnly = "-->";
        public const string wMinusSign = " - ";
        public const string wEqualSign = " = ";
        public const string wEqualNoSpace = "=";
        public const string wSpaceChar = " ";
        public const string wXChar = " x ";
        public const string wHeaderLine = "-------------------------------------------";
        public const string wTripleSpace = "   ";
        public const string wCustomXmlRequestStatus = "RequestStatus";
        public const string wCustomXsn = "openByDefault";
        public const string wRequestStatusNS = "d264e665-9d0b-48fb-b78a-227e1d3d858d";
        public const string wShpChart = "Chart";
        public const string wProperties = "Properties";
        public const string wCoreProperties = "coreProperties";
        public const string wAllAuthors = "* All Authors *";
        public const string wXmlNsStart = "xmlns:";
        public const string wSPCustomXmlProperties = "p:properties";
        public const string wSPDocManagement = "documentManagement";
        public const string wMoreInformationHere = "- More information here:\r\n";

        public const string wCompany = "Company";
        public const string wCreator = "creator";
        public const string wLastModifiedBy = "lastModifiedBy";
        public const string wRemovePI = "removePersonalInformation";
        public const string wRemoveDateTime = "removeDateAndTime";
        public const string wDocSecurity = "DocSecurity";
        
        public const string wAuthors = "Authors";
        public const string wParaFormatChange = " : Paragraph Formatting Change";
        public const string wParaDeleted = " : Paragraph Deleted ";
        public const string wRunFormatChange = " :  Run Formatting Change";
        public const string wDeletion = " :  Deletion = ";
        public const string wInsertion = " :  Insertion = ";
        public const string wAllFixes = "All";
        public const string wComments = "Comments";
        public const string wOle = "OLE objects";
        public const string wBookmarks = "Bookmarks";
        public const string wMathAccents = "Math Formula Accents";
        public const string wZipItem = "Zip Item";
        public const string wTableCell = "Table Cells";
        public const string wFldCodes = "Field Codes";
        public const string wListTemplates = "List Templates";
        public const string wFootnotes = "Footnotes";
        public const string wEndnotes = "Endnotes";
        public const string wConnections = "Connections";
        public const string wRevisions = "Revisions";
        public const string wTableProps = "Table Properties";
        public const string wTables = "Tables";
        public const string wFieldCodes = "Field Codes";
        public const string wSharedStrings = "Shared Strings";
        public const string wStyles = "Styles";
        public const string wInvalidXml = "invalid xml";
        public const string wFonts = "Fonts";
        public const string wRunFonts = "Run Fonts";
        public const string wForumulas = "Formulas";
        public const string wHiddenRowCol = "Hidden Rows & Columns";
        public const string wWorksheetInfo = "Worksheet Information";
        public const string wDefinedNames = "Defined Names";
        public const string wCellValues = "Cell Values";
        public const string wHyperlinks = "Hyperlinks";
        public const string wLinks = "Links";
        public const string wSlides = "Slides";
        public const string wSlideText = "Slide Text";
        public const string wSlideTitle = "Slide Titles";
        public const string wShapes = "Shapes";
        public const string wListStyles = "List Styles";
        public const string wTextboxes = "Textboxes";
        public const string wContentControls = "Content Controls";
        public const string wSlideTransitions = "Slide Transitions";
        public const string wEmbeddedObjects = "Embedded Objects";
        public const string wPackageParts = "Package Parts";
        public const string wXmlSignatures = "Document Signatures";
        public const string wCustomDocProps = "Custom Document Properties";
        public const string wDocProps = "Document Properties";
        public const string wNotesMaster = "Notes Master";
        public const string wPII = "Personally Identifiable Information";
        public const string wValidationErr = "Validation Errors";
        public const string mbWarning = "Warning";
        public const string wNone = " -- none ";
        public const string wBackslash = "/";
        public const string wHeadingBegin = "--- ";
        public const string wHeadingEnd = " ---";
        public const string wUsedIn = " -> Used in ";

        // error messages
        public const string errorUnableToFixDocument = "ERROR: Unable to fix document.";
        public const string errorValidation = "Validation Errors";
        public const string errorOpenWithSDK = "OpenWithSDK Error:";

        // app messages
        public const string doNotDeleteStyle = "DO NOT DELETE + ";
        public const string deleteStyle = "DELETE + ";
        public const string nonEmptyId = "Target Id cannot be empty.";
        public const string duplicateId = "OOXML part Id <1> already exists.";
        public const string pptNotesSizeReset = "Notes Page Size Reset.";
        public const string pptResetBulletMargins = "Bullet Margins Reset";
        public const string pptCustDataTags = "Removed Missing custData Tags";
        public const string fileDoesNotExist = "** File does not exist **";
        public const string themeFileAdded = "Theme File Added.";
        public const string unableToDownloadUpdate = "Unable to download update.";
        public const string sampleSentence = "This is a sample sentence.  Enter your own text here.";
        public const string fileConvertSuccessful = "** File Converted Successfully **";
        public const string invalidTag = "Invalid Tag: ";
        public const string invalidFile = "Invalid File. Please select a valid document.";
        public const string replacedWith = "Replaced With: ";
        public const string shpOfficeDrawing = ". Office Drawing";
        public const string corruptVmlDrawing = "Corrupt Vml Drawing";
        public const string shpVml = "Vml Shape";
        public const string shpVmlRectangle = "Vml Rectangle";
        public const string shpGroupSpaces = "    ";
        public const string shpGroup = ". Vml Group";
        public const string shpGraphic = ". Graphic";
        public const string shpMath = ". Math Shape";
        public const string shpDrawingDgm = ". Drawing Diagram Shape";
        public const string shpChartDraw = ". Chart Drawing Shape";
        public const string shpChartShape = ". Chart Shape";
        public const string shpShape = ". Shape";
        public const string shp3D = ". 3D Shape";
        public const string shpXlDraw = ". Spreadsheet Drawing";
        public const string customPropSaved = ": Custom Property Saved.";
        public const string noProp = " : Property Does Not Exist";
        public const string resetNotesMasterRegKey = "If you need to also resize the notes slides enable via: \r\n\r\nFile | Settings | Reset Notes Master";
        public const string convertedTo = " converted to -> ";
        public const string wdDocumentXml = "/word/document.xml";
        public const string xlWorkbookXml = "/xl/workbook.xml";
        public const string offCustomUIXml = "customUI/customUI.xml";
        public const string offCustomUI14Xml = "customUI/customUI14.xml";
        public const string offLabelInfo = "docMetadata/LabelInfo.xml";
        public const string allAuthors = "* All Authors *";

        // file locations
        public readonly static string fLogFilePath = Path.GetDirectoryName(Application.LocalUserAppDataPath) + "\\offexp.txt";
        public readonly static string fBackupFilePath = Path.GetDirectoryName(Application.LocalUserAppDataPath);
        public readonly static string fNormalTemplatePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\AppData\\Roaming\\Microsoft\\Templates\\Normal.dotm";

        // win32 api refs
        public const string gdi32 = "gdi32.dll";
        public const string user32 = "user32.dll";

        // xml tag strings
        public const string txtFallbackStart = "<mc:Fallback>";
        public const string txtFallbackEnd = "</mc:Fallback>";
        public const string txtMcChoiceTagEnd = "</mc:Choice>";
        public const string txtAtMentionStyle = "<w:rStyle w:val=\"Mention\"";
        public const string txtFieldCodeBegin = "<w:fldChar w:fldCharType=\"begin\"";
        public const string txtFieldCodeSeparate = "<w:fldChar w:fldCharType=\"separate\"";
        public const string txtFieldCodeEnd = "<w:fldChar w:fldCharType=\"end\"";
        public const string txtFieldCodeEndFullXml = "<w:r><w:fldChar w:fldCharType=\"end\" /></w:r>";
        public const string txtRsid = "<w:rsidR=";

        // notes slide refs
        public const string pptHeaderPlaceholder = "Header Placeholder";
        public const string pptHeaderPlaceholder1 = "Header Placeholder 1";
        public const string pptDatePlaceholder = "Date Placeholder";
        public const string pptDatePlaceholder2 = "Date Placeholder 2";
        public const string pptSlideImagePlaceholder = "Slide Image Placeholder";
        public const string pptSlideImagePlaceholder3 = "Slide Image Placeholder 3";
        public const string pptNotesPlaceholder = "Notes Placeholder";
        public const string pptNotesPlaceholder4 = "Notes Placeholder 4";
        public const string pptFooterPlaceholder = "Footer Placeholder";
        public const string pptFooterPlaceholder5 = "Footer Placeholder 5";
        public const string pptSlideNumberPlaceholder = "Slide Number Placeholder";
        public const string pptSlideNumberPlaceholder6 = "Slide Number Placeholder 6";
        public const string pptexceptionPowerPoint = "presentationDocument";
        public const string pptPicture = "Picture";

        // word sdk refs (DocumentFormat.OpenXml.Wordprocessing = dfow)
        public const string dfowBody = "DocumentFormat.OpenXml.Wordprocessing.Body";
        public const string dfowSdt = "DocumentFormat.OpenXml.Wordprocessing.Sdt";
        public const string dfowSdtContent = "DocumentFormat.OpenXml.Wordprocessing.SdtContentRun";
        public const string dfowStyle = "DocumentFormat.OpenXml.Wordprocessing.Style";
        public const string dfowLevel = "DocumentFormat.OpenXml.Wordprocessing.Level";
        public const string dfowText = "DocumentFormat.OpenXml.Wordprocessing.Text";
        public const string dfowRun = "DocumentFormat.OpenXml.Wordprocessing.Run";
        public const string dfowStdAlias = "DocumentFormat.OpenXml.Wordprocessing.SdtAlias";
        public const string dfowTag = "DocumentFormat.OpenXml.Wordprocessing.Tag";
        public const string dfowDataBinding = "DocumentFormat.OpenXml.Wordprocessing.DataBinding";
        public const string dfowTableGrid = "DocumentFormat.OpenXml.Wordprocessing.TableGrid";
        public const string dfowTableCell = "DocumentFormat.OpenXml.Wordprocessing.TableCell";
        public const string dfowTableRow = "DocumentFormat.OpenXml.Wordprocessing.TableRow";
        public const string dfowAttachedTemplate = "DocumentFormat.OpenXml.Wordprocessing.AttachedTemplate";

        // powerpoint sdk refs (DocumentFormat.OpenXml.Presentation = dfop)
        public const string dfopNVSP = "DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties";
        public const string dfopNVDP = "DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties";
        public const string dfopShape = "DocumentFormat.OpenXml.Presentation.Shape";
        public const string dfopPresentationPicture = "DocumentFormat.OpenXml.Presentation.Picture";

        // excelcnv.exe paths
        public const string sameBitnessO365 = @"C:\Program Files\Microsoft Office\root\Office16\excelcnv.exe";
        public const string x86OfficeO365 = @"C:\Program Files (x86)\Microsoft Office\root\Office16\excelcnv.exe";
        public const string sameBitnessMSI2016 = @"C:\Program Files\Microsoft Office\Office16\excelcnv.exe";
        public const string x86OfficeMSI2016 = @"C:\Program Files (x86)\Microsoft Office\Office16\excelcnv.exe";
        public const string sameBitnessMSI2013 = @"C:\Program Files\Microsoft Office\Office15\excelcnv.exe";
        public const string x86OfficeMSI2013 = @"C:\Program Files (x86)\Microsoft Office\Office15\excelcnv.exe";

        // Microsoft Graph endpoints
        public const string graphAPIEndpoint = "https://graph.microsoft.com/v1.0/me";

        // schema base urls
        public const string schemaOxml2006 = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
        public const string schemaMsft2007 = "http://schemas.microsoft.com/office/2007/relationships/";
        public const string schemaMsft2006 = "http://schemas.microsoft.com/office/2006/relationships/";
        public const string schemaMetadataProperties = "http://schemas.microsoft.com/office/2006/metadata/properties";
        public const string schemaCustomXsn = "http://schemas.microsoft.com/office/2006/metadata/customXsn";
        public const string schemaContentType = "http://schemas.microsoft.com/office/2006/metadata/contentType";
        public const string schemaTypes = "http://schemas.microsoft.com/office/2006/documentManagement/types";
        public const string OfficeExtendedProps = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        public const string OfficeCoreProps = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        public const string DcElements = "http://purl.org/dc/elements/1.1/";
        public const string schemaMip = "http://schemas.microsoft.com/office/2020/mipLabelMetadata";
        public const string schemaClpRelationship = "http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels";
        public const string schemaCustomXml = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
        public const string schemaPptSlideLayout = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";

        // OfficeDocument strings
        public const string idsDuplicateId = "OOXML part Id |1 already exists.";
        public const string idsNonEmptyId = "Target Id cannot be empty";

        // misc urls
        public const string helpLocation = "https://github.com/desjarlais/Office-File-Explorer/issues";

        // Office package relationship ids
        public const string CustomUIPartRelType = schemaMsft2006 + "ui/extensibility";
        public const string CustomUI14PartRelType = schemaMsft2007 + "ui/extensibility";
        public const string QATPartRelType = schemaMsft2006 + "ui/customization";
        public const string ImagePartRelType = schemaOxml2006 + "image";

        // WordprocessingML package relationship ids
        public const string wordMainAttributeNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public const string AfPartType = schemaOxml2006 + "aFChunk";
        public const string CommentsPartType = schemaOxml2006 + "comments"; // same as Excel, PowerPoint
        public const string DocumentSettingsPartType = schemaOxml2006 + "settings";
        public const string EndnotesPartType = schemaOxml2006 + "endnotes";
        public const string FontsTablePartType = schemaOxml2006 + "fontTable";
        public const string FooterPartType = schemaOxml2006 + "footer";
        public const string FootnotesPartType = schemaOxml2006 + "footnotes";
        public const string GlossaryDocPartType = schemaOxml2006 + "glossaryDocument";
        public const string HeaderPartType = schemaOxml2006 + "header";
        public const string MainDocumentPartType = schemaOxml2006 + "officeDocument"; // same as Excel, PowerPoint
        public const string NumberingDefsPartType = schemaOxml2006 + "numbering";
        public const string StyleDefsPartType = schemaOxml2006 + "styles"; // same as Excel
        public const string WebSettingsPartType = schemaOxml2006 + "webSettings";
        public const string DocumentTemplatePartType = schemaOxml2006 + "attachedTemplate";
        public const string FramesetsPartType = schemaOxml2006 + "frame";
        public const string MasterSubDocumentsPartType = schemaOxml2006 + "subDocument";
        public const string MailMergeDataSourcePartType = schemaOxml2006 + "mailMergeSource";
        public const string MailMergeHeaderSourcePartType = schemaOxml2006 + "mailMergeHeaderSource";
        public const string XslTransformationPartType = schemaOxml2006 + "transform";

        // SpreadsheetML package relationship ids
        public const string CalcChainPartType = schemaOxml2006 + "calcChain";
        public const string ChartSheetPartType = schemaOxml2006 + "chartSheet";
        public const string ConnectionsPartType = schemaOxml2006 + "connections";
        public const string CustomPropertyPartType = schemaOxml2006 + "customProperty";
        public const string CustomXmlMappingsPartType = schemaOxml2006 + "xmlMaps";
        public const string DialogsheetPartType = schemaOxml2006 + "dialogSheet";
        public const string DrawingsPartType = schemaOxml2006 + "drawing";
        public const string ExternalWorkbookRefsPartType = schemaOxml2006 + "externalLink";
        public const string MetadataPartType = schemaOxml2006 + "sheetMetadata";
        public const string PivotTablePartType = schemaOxml2006 + "pivotTable";
        public const string PivotCacheDefPartType = schemaOxml2006 + "pivotCacheDefinition";
        public const string PivotTableCacheRecordsPartType = schemaOxml2006 + "pivotCacheRecords";
        public const string QueryTablePartType = schemaOxml2006 + "queryTable";
        public const string SharedStringsPartType = schemaOxml2006 + "sharedStrings";
        public const string SharedWorkbookRevisionHeadersPartType = schemaOxml2006 + "revisionHeaders";
        public const string SharedWorkbookRevisionLogPartType = schemaOxml2006 + "revisionLog";
        public const string SharedWorkbookUserDataPartType = schemaOxml2006 + "usernames";
        public const string SingleCellTableDefsPartType = schemaOxml2006 + "tableSingleCells";
        public const string TableDefsPartType = schemaOxml2006 + "table";
        public const string VolatileDependenciesPartType = schemaOxml2006 + "volatileDependencies";
        public const string WorksheetPartType = schemaOxml2006 + "worksheet";
        public const string ExternalWorkbooksPartType = schemaOxml2006 + "externalLinkPath";

        // PresentationML package relationship ids
        public const string CommentAuthorsPartType = schemaOxml2006 + "commentAuthors";
        public const string HandoutMasterPartType = schemaOxml2006 + "handoutMaster";
        public const string NotesMasterPartType = schemaOxml2006 + "notesMaster";
        public const string NotesSlidePartType = schemaOxml2006 + "notesSlide";
        public const string PresentationPropertiesPartType = schemaOxml2006 + "presProps";
        public const string SlidePartType = schemaOxml2006 + "slide";
        public const string SlideLayoutPartType = schemaOxml2006 + "slideLayout";
        public const string SlideMasterPartType = schemaOxml2006 + "slideMaster";
        public const string SlideSynchronizationDataPartType = schemaOxml2006 + "slideUpdateInfo";
        public const string UserDefinedTagsPartType = schemaOxml2006 + "tags";
        public const string ViewPropertiesPartType = schemaOxml2006 + "viewProps";
        public const string HtmlPublishLocationPartType = schemaOxml2006 + "htmlPubSaveAs";
        public const string SlideSynchronizationServerLocationPartType = schemaOxml2006 + "slideUpdateUrl";

        // DrawingML package relationship ids
        public const string ChartPartType = schemaOxml2006 + "chart";
        public const string ChartDrawingPartType = schemaOxml2006 + "chartUserShapes";
        public const string DiagramColorsPartType = schemaOxml2006 + "diagramColors";
        public const string DiagramDataPartType = schemaOxml2006 + "diagramData";
        public const string DiagramLayoutPartType = schemaOxml2006 + "diagramLayout";
        public const string DiagramStylePartType = schemaOxml2006 + "diagramQuickStyle";
        public const string ThemePartType = schemaOxml2006 + "theme";
        public const string ThemeOverridePartType = schemaOxml2006 + "themeOverride";
        public const string TableStylesPartType = schemaOxml2006 + "tableStyles";

        // SharedML package relationship ids
        public const string AudioPartType = schemaOxml2006 + "audio";
        public const string EmbeddedControlPartType = schemaOxml2006 + "control";
        public const string EmbeddedObjectPartType = schemaOxml2006 + "oleObject";
        public const string EmbeddedPackagePartType = schemaOxml2006 + "package";
        public const string CoreFilePropertiesPartType = schemaOxml2006 + "metadata/core-properties";
        public const string FontPartType = schemaOxml2006 + "font";
        public const string ImagePartType = schemaOxml2006 + "image";
        public const string PrinterSettingsPartType = schemaOxml2006 + "printerSettings";
        public const string ThumbnailPartType = schemaOxml2006 + "thumbnail";
        public const string VideoPartType = schemaOxml2006 + "video";
        public const string HyperlinkPartType = schemaOxml2006 + "hyperlink";

        // Xml Rtf Color Replacements
        public const string rtfString = @"{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fmodern\fprq1\fcharset0 Courier New;}}{\colortbl	;\red0\green0\blue255;\red128\green0\blue0;\red255\green0\blue0;\red0\green128\blue0;}\pard\f0\fs20 ";
        public const string rtfAttributeName = @"\cf3 ";
        public const string rtfAttributeValue = @"\cf1 ";
        public const string rtfDelimiter = @"\cf1 ";
        public const string rtfAttributeQuote = @"\cf0 ";
        public const string rtfName = @"\cf2 ";
        public const string rtfComment = @"\cf4 ";

        /* RTF Color Codes
			cf1 = black
			cf2 = red
			cf3 = green
			cf4 = brown
			cf5 = blue
			cf6 = purple
			cf7 = cyan
			cf8 = gray
			cf9 = darkGray
			cf10 = light Red
			cf11 = light green
			cf12 = yellow
			cf13 = light blue
			cf14 = indigo
			cf15 = light cyan
			cf16 = white */

        // custom ui callbacks
        public const string xmlCustomOutspace = @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
	            <backstage>
		        <tab id=""customTab"" label=""Custom"">
			    <firstColumn>
				    <taskGroup id=""customTaskGroup"" label=""Custom Task Group"">
					    <category id=""tgCategory1"" label=""Category One"">
						    <task id=""task1"" label=""Task 1"" imageMso=""FileOpen""/>
						    <task id=""task2"" label=""Task 2"" imageMso=""FileSave""/>
						    <task id=""task3"" label=""Task 3"" imageMso=""FileSaveAs""/>
					    </category>
				    </taskGroup>
			    </firstColumn>
		        </tab>
	            </backstage>
            </customUI>";

        public const string xmlCustomTab = @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
	        <ribbon startFromScratch=""false"">
		        <tabs>
			        <tab id=""customTab"" label=""Custom Tab"">
				        <group id=""customGroup"" label=""Custom Group"">
					        <button id=""customButton"" label=""Custom Button"" imageMso=""HappyFace"" size=""large"" onAction=""Callback"" />
				        </group>
			        </tab>
		        </tabs>
	        </ribbon>
        </customUI>";

        public const string xmlExcelCustomTab = @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
	        <ribbon>
		        <tabs>
			        <tab id=""customTab"" label=""Contoso"" insertAfterMso=""TabHome"">
				        <group idMso=""GroupClipboard"" />
				        <group idMso=""GroupFont"" />
				        <group id=""customGroup"" label=""Contoso Tools"">
					        <button id=""customButton1"" label=""ConBold"" size=""large"" onAction=""conBoldSub"" imageMso=""Bold"" />
					        <button id=""customButton2"" label=""ConItalic"" size=""large"" onAction=""conItalicSub"" imageMso=""Italic"" />
					        <button id=""customButton3"" label=""ConUnderline"" size=""large"" onAction=""conUnderlineSub"" imageMso=""Underline"" />
				        </group>
				        <group idMso=""GroupEnterDataAlignment"" />
				        <group idMso=""GroupEnterDataNumber"" />
				        <group idMso=""GroupQuickFormatting"" />
			        </tab>
		        </tabs>
	        </ribbon>
        </customUI>";

        public const string xmlRepurpose = @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
	        <commands>
		        <command idMso=""Bold"" enabled=""false""/>
		        <command idMso=""Save"" onAction=""MySave""/>
	        </commands>
        </customUI>";

        public const string xmlWordGroupInsertTab = @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
	        <ribbon>
		        <tabs>
			        <tab idMso=""TabInsert"">
				        <group id=""customGroup"" label=""Contoso"" insertAfterMso=""GroupIllustrations"">
					        <button id=""customButton"" label=""Document ID"" size=""large"" imageMso=""ListNumVal"" onAction=""insertDocID"" />
				        </group>
			        </tab>
		        </tabs>
	        </ribbon>
        </customUI>";
    }
}
