using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Office_File_Explorer.Helpers
{
    public class AppUtilities
    {
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
            DelBookmarks,
            DelDupeAuthors
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
            RemovePIIOnSave,
            ResetNotesPageSize,
            ResetBulletMargins,
            CustomNotesPageReset
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
                case "0": return null;          // u0000
                case "1":                       // u0001 - show period "." for all non-printable characters
                case "2":                       // u0002
                case "3":                       // u0003
                case "4":                       // u0004
                case "5":                       // u0005
                case "6":                       // u0006
                case "7":                       // u0007
                case "8":                       // u0008
                case "9":                       // u0009
                case "10":                      // u000a
                case "11":                      // u000b
                case "12":                      // u000c
                case "13":                      // u000d
                case "14":                      // u000e
                case "15":                      // u000f
                case "16":                      // u0010
                case "17":                      // u0011
                case "18":                      // u0012
                case "19":                      // u0013
                case "20":                      // u0014
                case "21":                      // u0015
                case "22":                      // u0016
                case "23":                      // u0017
                case "24":                      // u0018
                case "25":                      // u0019
                case "26":                      // u001a
                case "27":                      // u001b
                case "28":                      // u001c
                case "29":                      // u001d
                case "30":                      // u001e
                case "31": return ".";          // u001f - printable characters
                case "32": return " ";          // u0020
                case "33": return "!";          // u0021
                case "34": return "\"";         // u0022
                case "35": return "#";          // u0023
                case "36": return "$";          // u0024
                case "37": return "%";          // u0025
                case "38": return "&";          // u0026
                case "39": return "'";          // u0027
                case "40": return "(";          // u0028
                case "41": return ")";          // u0029
                case "42": return "*";          // u002a
                case "43": return "+";          // u002b
                case "44": return ",";          // u002c
                case "45": return "-";          // u002d
                case "46": return ".";          // u002e
                case "47": return "/";          // u002f
                case "48": return "0";          // u0030
                case "49": return "1";          // u0031
                case "50": return "2";          // u0032
                case "51": return "3";          // u0033
                case "52": return "4";          // u0034
                case "53": return "5";          // u0035
                case "54": return "6";          // u0036
                case "55": return "7";          // u0037
                case "56": return "8";          // u0038
                case "57": return "9";          // u0039
                case "58": return ":";          // u003a
                case "59": return ";";          // u003b
                case "60": return "<";          // u003c
                case "61": return "=";          // u003d
                case "62": return ">";          // u003e
                case "63": return "?";          // u003f
                case "64": return "@";          // u0040
                case "65": return "A";          // u0041
                case "66": return "B";          // u0042
                case "67": return "C";          // u0043
                case "68": return "D";          // u0044
                case "69": return "E";          // u0045
                case "70": return "F";          // u0046
                case "71": return "G";          // u0047
                case "72": return "H";          // u0048
                case "73": return "I";          // u0049
                case "74": return "J";          // u004a
                case "75": return "K";          // u004b
                case "76": return "L";          // u004c
                case "77": return "M";          // u004d
                case "78": return "N";          // u004e
                case "79": return "O";          // u004f
                case "80": return "P";          // u0050
                case "81": return "Q";          // u0051
                case "82": return "R";          // u0052
                case "83": return "S";          // u0053
                case "84": return "T";          // u0054
                case "85": return "U";          // u0055
                case "86": return "V";          // u0056
                case "87": return "W";          // u0057
                case "88": return "X";          // u0058
                case "89": return "Y";          // u0059
                case "90": return "Z";          // u005a
                case "91": return "[";          // u005b
                case "92": return "\\";         // u005c
                case "93": return "]";          // u005d
                case "94": return "^";          // u005e
                case "95": return "_";          // u005f
                case "96": return "`";          // u0060
                case "97": return "a";          // u0061
                case "98": return "b";          // u0062
                case "99": return "c";          // u0063
                case "100": return "d";         // u0064
                case "101": return "e";         // u0065
                case "102": return "f";         // u0066
                case "103": return "g";         // u0067
                case "104": return "h";         // u0068
                case "105": return "i";         // u0069
                case "106": return "j";         // u006a
                case "107": return "k";         // u006b
                case "108": return "l";         // u006c
                case "109": return "m";         // u006d
                case "110": return "n";         // u006e
                case "111": return "o";         // u006f
                case "112": return "p";         // u0070
                case "113": return "q";         // u0071
                case "114": return "r";         // u0072
                case "115": return "s";         // u0073
                case "116": return "t";         // u0074
                case "117": return "u";         // u0075
                case "118": return "v";         // u0076
                case "119": return "w";         // u0077
                case "120": return "x";         // u0078
                case "121": return "y";         // u0079
                case "122": return "z";         // u007a
                case "123": return "{";         // u007b
                case "124": return "|";         // u007c
                case "125": return "}";         // u007d
                case "126": return "~";         // u007e
                case "127": return "DEL";       // u007f
                case "128": return "€";         // u0080 extended ascii
                case "129": return "\u0081";    // u0081
                case "130": return "‚";         // u0082
                case "131": return "ƒ";         // u0083
                case "132": return "„";         // u0084
                case "133": return "…";         // u0085
                case "134": return "†";         // u0086
                case "135": return "‡";         // u0087
                case "136": return "ˆ";         // u0088
                case "137": return "‰";         // u0089
                case "138": return "Š";         // u008a
                case "139": return "‹";         // u008b
                case "140": return "Œ";         // u008c
                case "141": return "\u008d";    // u008d
                case "142": return "Ž";         // u008e
                case "143": return "\u008f";    // u008f
                case "144": return "\u0090";    // u0090
                case "145": return "‘";         // u0091
                case "146": return "’";         // u0092
                case "147": return "“\t";       // u0093
                case "148": return "”";         // u0094
                case "149": return "•";         // u0095
                case "150": return "–";         // u0096
                case "151": return "—";         // u0097
                case "152": return "˜";         // u0098
                case "153": return "™";         // u0099
                case "154": return "š";         // u009a
                case "155": return "›";         // u009b
                case "156": return "œ";         // u009c
                case "157": return "";          // u009d
                case "158": return "ž";         // u009e
                case "159": return "Ÿ";         // u009f
                case "160": return "";          // u0100
                case "161": return "¡";         // u0101
                case "162": return "¢";         // u0102
                case "163": return "£";         // u0103
                case "164": return "¤";         // u0104
                case "165": return "¥";         // u0105
                case "166": return "¦";         // u0106
                case "167": return "§";         // u0107
                case "168": return "¨";         // u0108
                case "169": return "©";         // u0109
                case "170": return "ª";         // u010a
                case "171": return "«";         // u010b
                case "172": return "¬\t";       // u010c
                case "173": return "";          // u010d
                case "174": return "®";         // u010e
                case "175": return "¯";         // u010f
                case "176": return "°";         // u0110
                case "177": return "±";         // u0111
                case "178": return "²";         // u0112
                case "179": return "³";         // u0113
                case "180": return "´";         // u0114
                case "181": return "µ";         // u0115
                case "182": return "¶";         // u0116
                case "183": return "·";         // u0117
                case "184": return "¸";         // u0118
                case "185": return "¹";         // u0119
                case "186": return "º";         // u011a
                case "187": return "»";         // u011b
                case "188": return "¼";         // u011c
                case "189": return "½";         // u011d
                case "190": return "¾";         // u011e
                case "191": return "¿";         // u011f
                case "192": return "À";         // u0120
                case "193": return "Á";         // u0121
                case "194": return "Â";         // u0122
                case "195": return "Ã";         // u0123
                case "196": return "Ä";         // u0124
                case "197": return "Å";         // u0125
                case "198": return "Æ";         // u0126
                case "199": return "Ç";         // u0127
                case "200": return "È";         // u0128
                case "201": return "É\t";       // u0129
                case "202": return "Ê";         // u012a
                case "203": return "Ë";         // u012b
                case "204": return "Ì";         // u012c
                case "205": return "Í";         // u012d
                case "206": return "Î";         // u012e
                case "207": return "Ï";         // u012f
                case "208": return "Ð";         // u0130
                case "209": return "Ñ";         // u0131
                case "210": return "Ò";         // u0132
                case "211": return "Ó";         // u0133
                case "212": return "Ô";         // u0134
                case "213": return "Õ";         // u0135
                case "214": return "Ö";         // u0136
                case "215": return "×";         // u0137
                case "216": return "Ø";         // u0138
                case "217": return "Ù";         // u0139
                case "218": return "Ú";         // u013a
                case "219": return "Û";         // u013b
                case "220": return "Ü";         // u013c
                case "221": return "Ý";         // u013d
                case "222": return "Þ";         // u013e
                case "223": return "ß";         // u013f
                case "224": return "à";         // u0140
                case "225": return "á";         // u0141
                case "226": return "â";         // u0142
                case "227": return "ã";         // u0143
                case "228": return "ä";         // u0144
                case "229": return "å";         // u0145
                case "230": return "æ";         // u0146
                case "231": return "ç";         // u0147
                case "232": return "è";         // u0148
                case "233": return "é";         // u0149
                case "234": return "ê";         // u014a
                case "235": return "ë";         // u014b
                case "236": return "ì";         // u014c
                case "237": return "í";         // u014d
                case "238": return "î";         // u014e
                case "239": return "ï";         // u014f
                case "240": return "ð";         // u0150
                case "241": return "ñ";         // u0151
                case "242": return "ò";         // u0152
                case "243": return "ó";         // u0153
                case "244": return "ô";         // u0154
                case "245": return "õ";         // u0155
                case "246": return "ö";         // u0156
                case "247": return "÷";         // u0157
                case "248": return "ø";         // u0158
                case "249": return "ù";         // u0159
                case "250": return "ú";         // u015a
                case "251": return "û";         // u015b
                case "252": return "ü";         // u015c
                case "253": return "ý";         // u015d
                case "254": return "þ";         // u015e
                case "255": return "ÿ\t";       // u015f
                default: return "EXT-ASC";
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
