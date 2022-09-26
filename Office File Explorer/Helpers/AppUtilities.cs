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

        public static string ConvertByteToText(string byteCode)
        {
            switch (byteCode)
            {
                case "32": return " ";
                case "33": return "!";
                case "34": return "\"";
                case "35": return "#";
                case "36": return "$";
                case "37": return "%";
                case "38": return "&";
                case "39": return "'";
                case "40": return "(";
                case "41": return ")";
                case "42": return "*";
                case "43": return "+";
                case "44": return ",";
                case "45": return "-";
                case "46": return ".";
                case "47": return "/";
                case "48": return "0";
                case "49": return "1";
                case "50": return "2";
                case "51": return "3";
                case "52": return "4";
                case "53": return "5";
                case "54": return "6";
                case "55": return "7";
                case "56": return "8";
                case "57": return "9";
                case "58": return ":";
                case "59": return ";";
                case "60": return "<";
                case "61": return "=";
                case "62": return ">";
                case "63": return "?";
                case "64": return "@";
                case "65": return "A";
                case "66": return "B";
                case "67": return "C";
                case "68": return "D";
                case "69": return "E";
                case "70": return "F";
                case "71": return "G";
                case "72": return "H";
                case "73": return "I";
                case "74": return "J";
                case "75": return "K";
                case "76": return "L";
                case "77": return "M";
                case "78": return "N";
                case "79": return "O";
                case "80": return "P";
                case "81": return "Q";
                case "82": return "R";
                case "83": return "S";
                case "84": return "T";
                case "85": return "U";
                case "86": return "V";
                case "87": return "W";
                case "88": return "X";
                case "89": return "Y";
                case "90": return "Z";
                case "91": return "[";
                case "92": return "\\";
                case "93": return "]";
                case "94": return "^";
                case "95": return "_";
                case "96": return "`";
                case "97": return "a";
                case "98": return "b";
                case "99": return "c";
                case "100": return "d";
                case "101": return "e";
                case "102": return "f";
                case "103": return "g";
                case "104": return "h";
                case "105": return "i";
                case "106": return "j";
                case "107": return "k";
                case "108": return "l";
                case "109": return "m";
                case "110": return "n";
                case "111": return "o";
                case "112": return "p";
                case "113": return "q";
                case "114": return "r";
                case "115": return "s";
                case "116": return "t";
                case "117": return "u";
                case "118": return "v";
                case "119": return "w";
                case "120": return "x";
                case "121": return "y";
                case "122": return "z";
                case "123": return "{";
                case "124": return "|";
                case "125": return "}";
                case "126": return "~";
                case "127": return "DEL";
                default: return null;
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
                path.Contains(' ') && (!path.StartsWith("\"") && !path.EndsWith("\"")) ? "\"" + path + "\"" : path : string.Empty;
        }
    }
}
