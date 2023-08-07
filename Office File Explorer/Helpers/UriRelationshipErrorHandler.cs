using System;
using DocumentFormat.OpenXml.Packaging;

namespace Office_File_Explorer.Helpers
{
    public class UriRelationshipErrorHandler : RelationshipErrorHandler
    {
        public override string Rewrite(Uri partUri, string id, string uri)
        {
            return "http://link-invalid";
        }
    }
}
