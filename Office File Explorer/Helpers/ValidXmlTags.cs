using System.Collections.Generic;

namespace Office_File_Explorer.Helpers
{
    internal class ValidXmlTags
    {
        // Valid tag replacements
        public const string StrValidMcChoice1 = "</mc:Choice><mc:Choice>";
        public const string StrValidMcChoice2 = "</mc:Choice><mc:Fallback>";
        public const string StrValidMcChoice3 = "</mc:Choice></mc:AlternateContent>";
        public const string StrValidMcChoice4 = "</mc:Choice></mc:AlternateContent></w:r>";
        public const string StrOmitFallback = "</mc:AlternateContent></w:r>";

        // Example: valid xml tag string values
        // <m:oMath><mc:AlternateContent><mc:Choice Requires="wps">
        // Example: Escape character value for RegEx matches
        // <m:oMath><mc:AlternateContent><mc:Choice Requires=\"wps\">
        public const string StrValidomathwps = "<m:oMath><mc:AlternateContent><mc:Choice Requires=\"wps\">";
        public const string StrValidomathwpg = "<m:oMath><mc:AlternateContent><mc:Choice Requires=\"wpg\">";
        public const string StrValidomathwpi = "<m:oMath><mc:AlternateContent><mc:Choice Requires=\"wpi\">";
        public const string StrValidomathwpc = "<m:oMath><mc:AlternateContent><mc:Choice Requires=\"wpc\">";
        public const string StrValidVshape = "</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent>";
        public const string StrValidVshapegroup = "</w:txbxContent></v:textbox></v:shape></v:group></w:pict></mc:Fallback></mc:AlternateContent>";

        public IEnumerable<string> ValidTags()
        {
            yield return StrValidMcChoice1;
            yield return StrValidMcChoice2;
            yield return StrValidMcChoice3;
            yield return StrValidMcChoice4;
            yield return StrValidomathwpc;
            yield return StrValidomathwpg;
            yield return StrValidomathwpi;
            yield return StrValidomathwps;
            yield return StrOmitFallback;
            yield return StrValidVshape;
            yield return StrValidVshapegroup;
        }
    }
}
