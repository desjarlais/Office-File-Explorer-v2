using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Office_File_Explorer.Helpers
{
    public class AppUtilities
    {
        [Flags]
        public enum WordViewCmds
        {
            None = 0,
            ContentControls = 1,
            Styles = 2,
            Hyperlinks = 4,
            ListTemplates = 8,
            Fonts = 16,
            Footnotes = 32,
            Endnotes = 64,
            DocumentProperties = 128,
            Bookmarks = 256,
            Comments = 512,
            FieldCodes = 1024,
            Tables = 2048
        }

        [Flags]
        public enum ExcelViewCmds
        {
            None = 0,
            Links = 1,
            Comments = 2,
            WorksheetInfo = 4,
            HiddenRowsCols = 8,
            SharedStrings = 16,
            Hyperlinks = 32,
            DefinedNames = 64,
            Connections = 128
        }

        [Flags]
        public enum PowerPointViewCmds
        {
            None = 0,
            Hyperlinks = 1,
            SlideTitles = 2,
            Comments = 4,
            SlideText = 8,
            SlideTransitions = 16,
            Fonts = 32
        }

        [Flags]
        public enum OfficeViewCmds
        {
            None = 0,
            OleObjects = 1,
            Shapes = 2,
            PackageParts = 4,
            CustomProperties = 8,
            CustomXml = 16,
            Images = 32
        }

        public enum WordModifyCmds
        {
            None,
            DelHF,
            DelPgBrk,
            DelComments,
            DelHiddenTxt,
            DelFootnotes,
            DelEndnotes,
            DelOrphanLT,
            DelOrphanStyles,
            SetPrintOrientation,
            ChangeDefaultTemplate,
            AcceptRevisions,
            ConvertDocmToDocx,
            RemovePII,
            RemoveCustomTitleProp,
            UpdateCcNamespaceGuid,
            DelBookmarks
        }

        public enum ExcelModifyCmds
        {
            None,
            DelLink,
            DelLinks,
            DelComments,
            DelSheet,
            DelEmbeddedLinks,
            ConvertXlsmToXlsx,
            ConvertStrictToXlsx
        }

        public enum PowerPointModifyCmds
        {
            None,
            MoveSlide,
            DelComments,
            ConvertPptmToPptx,
            RemovePIIOnSave
        }

        public static string GetFontCharacterSet(string input)
        {
            switch (input)
            {
                case "0": return "ANSI";
                case "1": return "Default";
                case "2": return "Symbol";
                case "4D": return "Macintosh";
                case "80": return "JIS";
                case "81": return "Hangul";
                case "82": return "Johab";
                case "86": return "GB-2312";
                case "88": return "Chinese Big Five";
                case "A1": return "Greek";
                case "A2": return "Turkish";
                case "A3": return "Vietnamese";
                case "B1": return "Hebrew";
                case "B2": return "Arabic";
                case "BA": return "Baltic";
                case "CC": return "Russian";
                case "DE": return "Thai";
                case "EE": return "Eastern European";
                case "FF": return "OEM";
                default: return "App-Defined";
            }
        }

        public static void PlatformSpecificProcessStart(string url)
        {
            // known issue in .NET Core https://github.com/dotnet/corefx/issues/10361
            try
            {
                Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            }
            catch
            {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    url = url.Replace("&", "^&");
                    Process.Start(new ProcessStartInfo("cmd", $"/c start {url}") { CreateNoWindow = true });
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                {
                    Process.Start("xdg-open", url);
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                {
                    Process.Start("open", url);
                }
                else
                {
                    FileUtilities.WriteToLog(Strings.fLogFilePath, "Unable to open web site.");
                }
            }
        }

        public static string AddQuotesIfRequired(string path)
        {
            return !string.IsNullOrWhiteSpace(path) ?
                path.Contains(' ') && (!path.StartsWith("\"") && !path.EndsWith("\"")) ?
                    "\"" + path + "\"" : path :
                    string.Empty;
        }
    }
}
