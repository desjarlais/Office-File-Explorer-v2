// Open Xml SDK Refs
using DocumentFormat.OpenXml;
using OM = DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

// .NET refs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using System.IO;
using Office_File_Explorer.WinForms;
using System.Reflection;
using System.Text;

using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using Hyperlink = DocumentFormat.OpenXml.Wordprocessing.Hyperlink;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;

namespace Office_File_Explorer.Helpers
{
    class WordFixes
    {
        static bool corruptionFound = false;

        /// <summary>
        /// table rows need at least one child element, besides table row props
        /// if a table row does not have any child elements, that row is corrupt and the doc is malformed
        /// if a table row also does not have a table cell, that row is corrupt
        /// this function will remove those bad rows
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool FixTableRowCorruption(string filePath)
        {
            corruptionFound = false;

            using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
            {
                IEnumerable<TableRow> tRows = document.MainDocumentPart.Document.Descendants<TableRow>();

                foreach (TableRow tRow in tRows)
                {
                    // table rows need at least one child, make sure each row has one
                    if (tRow.ChildElements.Count == 0)
                    {
                        tRow.Remove();
                        corruptionFound = true;
                    }

                    // if it has a row, but no column, that is corrupt and the row can also be removed
                    IEnumerable<TableCell> tCells = tRow.Descendants<TableCell>();

                    if (!tCells.Any())
                    {
                        tRow.Remove();
                        corruptionFound = true;
                    }
                }
            }

            return corruptionFound;
        }



        /// <summary>
        /// special circumstance where the content control value has the correct SharePoint guid
        /// the custom xml however is wrong, so this will compare those values and always use the content control
        /// which is essentially the exact opposite of FixContentControlNamespaces
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool FixSharePointGuidUsingContentControlGuid(string filePath)
        {
            corruptionFound = false;

            Dictionary<string,string> contentControlGuids = new Dictionary<string, string>();
            List<string> pMappingList = new List<string>();

            using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
            {
                // loop content controls and get the xpath and prefix mapping for the namespace as a list
                foreach (var cc in document.ContentControls())
                {
                    SdtProperties props = cc.Elements<SdtProperties>().FirstOrDefault();
                    string ccName = string.Empty;
                    string ccGuid = string.Empty;
                    string ccNs = string.Empty;

                    foreach (OpenXmlElement oxe in props.ChildElements)
                    {
                        // get the cc name
                        if (oxe.LocalName == "alias")
                        {
                            PropertyInfo prop = oxe.GetType().GetProperty("Val");
                            ccName = prop.GetValue(oxe).ToString();
                        }

                        // get details from the databinding tag
                        if (oxe.GetType().ToString() == Strings.dfowDataBinding)
                        {
                            foreach (OpenXmlAttribute oxa in oxe.GetAttributes())
                            {
                                // get the guids
                                if (oxa.LocalName == "prefixMappings")
                                {
                                    string[] prefixMappings = oxa.Value.Split(' ');
                                    foreach (string s in prefixMappings)
                                    {
                                        // parse out the namespace number
                                        pMappingList.Add(s);
                                    }
                                }

                                // get the ns number
                                if (oxa.LocalName == "xpath")
                                {
                                    string[] xpathVal = oxa.Value.Split('/');
                                    foreach (string s in xpathVal)
                                    {
                                        if (s.Contains(ccName))
                                        {
                                            // parse out the ns value
                                            string[] ns = s.Split(':');
                                            ccNs = ns[0];
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // loop the guids and find the guid based on ns #
                    bool nameExists = false;
                    foreach (var pm in pMappingList)
                    {
                        string index = Strings.wXmlNsStart + ccNs + "=";
                        if (pm.StartsWith(Strings.wXmlNsStart + ccNs))
                        {
                            // parse out the guid
                            ccGuid = pm.Substring(index.Length, pm.Length - index.Length);
                            foreach (var g in contentControlGuids)
                            {
                                if (g.Key == ccName)
                                {
                                    nameExists = true;
                                }
                            }
                        }
                    }

                    if (nameExists == false)
                    {
                        contentControlGuids.Add(ccName, ccGuid);
                    }
                }

                // loop custom xml and compare with content control value
                foreach (CustomXmlPart cxp in document.MainDocumentPart.CustomXmlParts)
                {
                    XmlDocument xDoc = new XmlDocument();
                    Stream sRW = cxp.GetStream(FileMode.Open, FileAccess.Read);
                    string docText = null;
                    using (StreamReader sr = new StreamReader(sRW))
                    {
                        docText = sr.ReadToEnd();
                    }
                    sRW.Close();

                    Stream sRO = cxp.GetStream(FileMode.Open, FileAccess.Read);
                    xDoc.Load(sRO);

                    if (xDoc.DocumentElement.NamespaceURI == Strings.schemaMetadataProperties)
                    {
                        // loop through the metadata and get the uri's
                        foreach (XmlNode xNode in xDoc.ChildNodes)
                        {
                            if (xNode.Name == Strings.wSPCustomXmlProperties)
                            {
                                foreach (XmlNode xNode2 in xNode.ChildNodes)
                                {
                                    if (xNode2.Name == Strings.wSPDocManagement)
                                    {
                                        foreach (XmlNode xNode3 in xNode2.ChildNodes)
                                        {
                                            foreach (XmlAttribute xa in xNode3.Attributes)
                                            {
                                                if (xa.LocalName == "xmlns")
                                                {
                                                    // loop the list and update the guid with the correct value from the content control
                                                    foreach (var ccg in contentControlGuids)
                                                    {
                                                        string nsValue = ccg.Value.Replace("'", string.Empty);
                                                        if (xa.OwnerElement.LocalName == ccg.Key)
                                                        {
                                                            if (xa.Value != nsValue)
                                                            {
                                                                string tagName = "<" + xa.OwnerElement.LocalName + " xmlns=" + '"';
                                                                Regex regexTagName = new Regex(tagName + xa.Value);
                                                                foreach (Match m in regexTagName.Matches(docText))
                                                                {
                                                                    if (!m.Value.Contains(ccg.Value))
                                                                    {
                                                                        docText = docText.Replace(m.Value, tagName + nsValue);
                                                                        corruptionFound = true;
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
                            }
                        }
                    }

                    sRO.Close();
                    if (corruptionFound)
                    {
                        using (StreamWriter sw = new StreamWriter(cxp.GetStream(FileMode.Create)))
                        {
                            sw.Write(docText);
                        }
                        document.Save();
                    }
                }
            }

            return corruptionFound;
        }

        /// <summary>
        /// Unknown scenario exists where the backend custom xml guids for columns from SharePoint are different
        /// this fix will populate a list of guids so the user can choose which one to use for the rest of the custom xml
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool FixSharePointGuidWithUserSelectedGuid(string filePath)
        {
            corruptionFound = false;

            List<string> guids = new List<string>();
            List<string> guidList = new List<string>();
            string docMgmtUri = string.Empty;

            using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
            {
                foreach (CustomXmlPart cxp in document.MainDocumentPart.CustomXmlParts)
                {
                    XmlDocument xDoc = new XmlDocument();

                    // need to load as a stream to get around a .net bug where using GetStream wasn't closing out properly
                    // this allows me to close the stream manually to avoid the exception
                    Stream stream = cxp.GetStream();
                    xDoc.Load(stream);

                    if (xDoc.DocumentElement.NamespaceURI == Strings.schemaMetadataProperties)
                    {
                        // loop through the metadata and get the uri's
                        foreach (XmlNode xNode in xDoc.ChildNodes)
                        {
                            if (xNode.Name == Strings.wSPCustomXmlProperties)
                            {
                                foreach (XmlNode xNode2 in xNode.ChildNodes)
                                {
                                    if (xNode2.Name == Strings.wSPDocManagement)
                                    {
                                        foreach (XmlNode xNode3 in xNode2.ChildNodes)
                                        {
                                            foreach (XmlAttribute xa in xNode3.Attributes)
                                            {
                                                if (xa.LocalName == "xmlns")
                                                {
                                                    guids.Add(xa.Value);
                                                    docMgmtUri = cxp.Uri.ToString();
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    stream.Close();
                    guidList = guids.Distinct().ToList();
                }
            }

            using (var f = new FrmSPCustomXmlGuids(guidList))
            {
                // ask the user for the guid to replace the existing values with
                var result = f.ShowDialog();

                if (f.newGuid != string.Empty)
                {
                    using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
                    {
                        foreach (CustomXmlPart cxp in document.MainDocumentPart.CustomXmlParts)
                        {
                            if (cxp.Uri.ToString() == docMgmtUri)
                            {
                                string docText = null;
                                using (StreamReader sr = new StreamReader(cxp.GetStream()))
                                {
                                    docText = sr.ReadToEnd();
                                }

                                // find each guid and replace it with the user selected value
                                Regex regexText = new Regex("[({]?[a-fA-F0-9]{8}[-]?([a-fA-F0-9]{4}[-]?){3}[a-fA-F0-9]{12}[})]?");
                                foreach (Match m in regexText.Matches(docText))
                                {
                                    docText = docText.Replace(m.Value, f.newGuid);
                                    corruptionFound = true;
                                }
                                
                                if (corruptionFound)
                                {
                                    using (StreamWriter sw = new StreamWriter(cxp.GetStream(FileMode.Create)))
                                    {
                                        sw.Write(docText);
                                    }
                                    document.Save();
                                }
                            }
                        }
                    }
                }
            }

            return corruptionFound;
        }

        public static bool FixContentControlGuidWithSharePointGuid(string filePath)
        {
            bool corruptionFound = false;
            Dictionary<string, string> contentControlGuids = new Dictionary<string, string>();

            try
            {
                using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
                {
                    // loop sharepoint metadata and get list of column names and guids
                    foreach (CustomXmlPart cxp in document.MainDocumentPart.CustomXmlParts)
                    {
                        XmlDocument xDoc = new XmlDocument();
                        Stream stream = cxp.GetStream();
                        xDoc.Load(stream);

                        if (xDoc.DocumentElement.NamespaceURI == Strings.schemaMetadataProperties)
                        {
                            foreach (XmlNode xNode in xDoc.ChildNodes)
                            {
                                if (xNode.Name == Strings.wSPCustomXmlProperties)
                                {
                                    foreach (XmlNode xNode2 in xNode.ChildNodes)
                                    {
                                        if (xNode2.Name == Strings.wSPDocManagement)
                                        {
                                            foreach (XmlNode xNode3 in xNode2.ChildNodes)
                                            {
                                                foreach (XmlAttribute xa in xNode3.Attributes)
                                                {
                                                    if (xa.LocalName == "xmlns")
                                                    {
                                                        contentControlGuids.Add(xNode3.Name, xa.Value);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        stream.Close();
                    }

                    // now loop through the content controls, check each data binding xpath against the guids
                    // if the guid is different, update the content control with the new guid from SharePoint
                    foreach (var cc in document.ContentControls())
                    {
                        List<string> prefixMappingList = new List<string>();
                        List<string> xPathList = new List<string>();
                        string oldXpath = string.Empty;

                        SdtProperties props = cc.Elements<SdtProperties>().FirstOrDefault();

                        foreach (OpenXmlElement oxe in props.ChildElements)
                        {
                            DataBinding db = new DataBinding();
                            bool pmFound = false;
                            bool xpathFound = false;

                            // get details from the databinding tag
                            if (oxe.GetType().ToString() == Strings.dfowDataBinding)
                            {
                                foreach (OpenXmlAttribute oxa in oxe.GetAttributes())
                                {
                                    if (oxa.LocalName == "prefixMappings")
                                    {
                                        string[] prefixMappings = oxa.Value.Split(' ');
                                        foreach (string s in prefixMappings)
                                        {
                                            prefixMappingList.Add(s);
                                        }
                                        pmFound = true;
                                    }

                                    if (oxa.LocalName == "xpath")
                                    {
                                        oldXpath = oxa.Value;
                                        string[] xpathVal = oxa.Value.Split('/');
                                        foreach (string s in xpathVal)
                                        {
                                            xPathList.Add(s);
                                        }
                                        xpathFound = true;
                                    }
                                }
                            }

                            // loop the prefixMapping guids and compare them with the xpath value
                            // based on the ns value of the xpath, get the guid from the prefixMapping
                            // if the guid is different from the dictionary, update the content control with the new guid
                            if (xpathFound && pmFound)
                            {
                                string lastXpath = xPathList.Last();
                                string nsValue = lastXpath.Split(':')[0];

                                foreach (var pm in prefixMappingList)
                                {
                                    if (pm.Contains("xmlns:" + nsValue))
                                    {
                                        string[] guidVal = pm.Split('=');
                                        string guid = guidVal[1].Replace("'", string.Empty);
                                        string ccName = lastXpath.Split(':')[1];
                                        string ccNameSubstring = ccName.Substring(0, ccName.Length - 3);

                                        // check the content control guid against the SharePoint guid
                                        foreach (KeyValuePair<string, string> entry in contentControlGuids)
                                        {
                                            if (entry.Key == ccNameSubstring)
                                            {
                                                if (entry.Value != guid)
                                                {
                                                    // this is the correct guid and needs to be replaced
                                                    corruptionFound = true;
                                                    // create the databinding object that will replace the old value
                                                    db.XPath = oldXpath;
                                                    db.PrefixMappings = db.PrefixMappings.Value.Replace(guid, entry.Value);
                                                    // remove the current databinding tag and add the new one back into the file
                                                    oxe.Remove();
                                                    props.Append(db);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (corruptionFound)
                    {
                        document.Save();
                    }
                }

                return corruptionFound;
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, ex.Message);
                return corruptionFound;
            }
        }

        /// <summary>
        /// there are SharePoint(SP) migration scenarios as well as SP Copy To style scenarios where a document will move from one doc library to another
        /// this causes the guid associated with the mapped content control/quick part to change
        /// the custom xml item.xml file will get this new guid, but the document and its content controls will not get this update
        /// this fix will update the content control to the new guid from the custom xml
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool FixContentControlNamespaces(string filePath)
        {
            corruptionFound = false;
            bool mismatchNamespaceFound;

            try
            {
                // hiding additional guid checks behind settings, if enabled, makes sure guids are correct in SP custom xml first
                // correct is a subjective term here, essentially the setting lets the user choose the context
                // whether the content control is "correct" or the value from SP is "correct" is up to the user
                // they can also use a user selected value and choose manually from existing namespaces
                // once the custom xml is updated, those later functions will just pull in the corrected guids established here
                if (Properties.Settings.Default.UseSharePointGuid == false)
                {
                    if (Properties.Settings.Default.UseContentControlGuid)
                    {
                        corruptionFound = FixSharePointGuidUsingContentControlGuid(filePath);
                    }
                    else if (Properties.Settings.Default.UseUserSelectedCCGuid)
                    {
                        corruptionFound = FixSharePointGuidWithUserSelectedGuid(filePath);
                    }
                }

                // now that we have the custom xml updated, check the content controls
                using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
                {
                    string newGuid = string.Empty;
                    string oldGuid = string.Empty;
                    string oldNs = string.Empty;
                    string oldXpath = string.Empty;
                    string oldStoreItemID = string.Empty;

                    // this is a 4 part fix
                    // 1. loop each content control and get the databinding xpaths and prefixmappings
                    // 2. get the custom xml file and loop its nodes looking for the node that matches up with the content control
                    // 3. compare the namespaces in the custom xml with the prefix mappings used in the content control
                    // 4. if the namespace is different, create a new databinding updated with the new namespace and push it back into the content control
                    foreach (var cc in document.ContentControls())
                    {
                        string ccTag = string.Empty;
                        string nsUri = string.Empty;
                        newGuid = string.Empty;
                        oldGuid = string.Empty;
                        oldNs = string.Empty;
                        oldXpath = string.Empty;
                        oldStoreItemID = string.Empty;

                        List<string> prefixMappingList = new List<string>();
                        List<string> xPathList = new List<string>();

                        mismatchNamespaceFound = false;

                        SdtProperties props = cc.Elements<SdtProperties>().FirstOrDefault();

                        foreach (OpenXmlElement oxe in props.ChildElements)
                        {
                            // get details from the databinding tag
                            if (oxe.GetType().ToString() == Strings.dfowDataBinding)
                            {
                                foreach (OpenXmlAttribute oxa in oxe.GetAttributes())
                                {
                                    if (oxa.LocalName == "prefixMappings")
                                    {
                                        string[] prefixMappings = oxa.Value.Split(' ');
                                        foreach (string s in prefixMappings)
                                        {
                                            prefixMappingList.Add(s);
                                        }
                                    }

                                    if (oxa.LocalName == "xpath")
                                    {
                                        oldXpath = oxa.Value;
                                        string[] xpathVal = oxa.Value.Split('/');
                                        foreach (string s in xpathVal)
                                        {
                                            xPathList.Add(s);
                                        }
                                    }

                                    if (oxa.LocalName == "storeItemID")
                                    {
                                        oldStoreItemID = oxa.Value;
                                    }
                                }
                            }
                        }

                        // loop the custom xml and check for the mapped values
                        foreach (CustomXmlPart cxp in document.MainDocumentPart.CustomXmlParts)
                        {
                            XmlDocument xDoc = new XmlDocument();
                            // need to load as a stream to get around a .net bug where using GetStream wasn't closing out properly
                            // this allows me to close the stream manually to avoid the exception
                            Stream stream = cxp.GetStream();
                            xDoc.Load(stream);

                            if (xDoc.DocumentElement.NamespaceURI == Strings.schemaMetadataProperties)
                            {
                                // loop through the metadata and get the uri's
                                foreach (XmlNode xNode in xDoc.ChildNodes)
                                {
                                    if (xNode.Name == Strings.wSPCustomXmlProperties)
                                    {
                                        foreach (XmlNode xNode2 in xNode.ChildNodes)
                                        {
                                            if (xNode2.Name == Strings.wSPDocManagement)
                                            {
                                                // loop each custom xml and find the name that matches the xpath from the content control
                                                // check the val of the custom xml ns with the prefixmapping ns value
                                                // if they don't match, replace the content control guid with the one form the custom xml guid
                                                foreach (XmlNode xNode3 in xNode2.ChildNodes)
                                                {
                                                    foreach (string s in xPathList)
                                                    {
                                                        if (s != string.Empty)
                                                        {
                                                            // pull the ns val out of xpath
                                                            if (s.Substring(0, 2) == "ns")
                                                            {
                                                                string[] clientNs = s.Split(':');
                                                                if (clientNs[1].StartsWith(xNode3.Name))
                                                                {
                                                                    foreach (string pm in prefixMappingList)
                                                                    {
                                                                        if (pm.StartsWith(Strings.wXmlNsStart + clientNs[0]))
                                                                        {
                                                                            string[] serverNs = pm.Split('=');
                                                                            string newServerNs = serverNs[1].Replace("'", string.Empty);
                                                                            if (newServerNs != xNode3.NamespaceURI)
                                                                            {
                                                                                // this is the correct guid and needs to be replaced
                                                                                newGuid = xNode3.NamespaceURI;
                                                                                oldNs = clientNs[0];
                                                                                oldGuid = serverNs[1];
                                                                                mismatchNamespaceFound = true;
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
                                    }
                                }
                            }

                            stream.Close();
                        }

                        // now replace the value in the content control if there was a new guid found
                        if (mismatchNamespaceFound)
                        {
                            foreach (OpenXmlElement oxe in props.ChildElements)
                            {
                                if (oxe.GetType().ToString() == Strings.dfowDataBinding)
                                {
                                    foreach (OpenXmlAttribute oxa in oxe.GetAttributes())
                                    {
                                        if (oxa.LocalName == "prefixMappings")
                                        {
                                            string[] prefixMappings = oxa.Value.Split(' ');
                                            foreach (string s in prefixMappings)
                                            {
                                                if (s.StartsWith(Strings.wXmlNsStart + oldNs))
                                                {
                                                    // prep the namespaces 
                                                    string oldNamespace = Strings.wXmlNsStart + oldNs + Strings.wEqualNoSpace + oldGuid;
                                                    string newNamespace = Strings.wXmlNsStart + oldNs + "='" + newGuid + "'";

                                                    // create the databinding object that will replace the old value
                                                    DataBinding db = new DataBinding();
                                                    db.XPath = oldXpath;
                                                    db.PrefixMappings = oxa.Value.Replace(oldNamespace, newNamespace);

                                                    if (oldStoreItemID != string.Empty)
                                                    {
                                                        db.StoreItemId = oldStoreItemID;
                                                    }

                                                    // remove the current databinding tag and add the new one back into the file
                                                    oxe.Remove();
                                                    props.Append(db);
                                                    corruptionFound = true;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (corruptionFound)
                    {
                        document.Save();
                    }
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, ex.Message);
                return corruptionFound;
            }

            return corruptionFound;
        }

        public static bool FixDataDescriptor(string filePath)
        {
            corruptionFound = false;

            if (Office.IsZippedFileCorrupt(filePath))
            {
                // change the datadescriptor back to 0
                byte[] zByte = new byte[] { 0x00 };

                // IsZippedFileCorrupt will populate a list of corrupt indexes to change
                // go through each and overwrite 0 instead of 8 and it should fix the bad zip items 
                foreach (var b in Office.corruptByteIndexes)
                {
                    using (FileStream zFile = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
                    {
                        zFile.Seek(b, SeekOrigin.Current);
                        zFile.WriteByte(zByte[0]);
                        corruptionFound = true;
                    }
                }
            }

            return corruptionFound;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool FixMathAccents(string filePath)
        {
            corruptionFound = false;

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filePath, true))
            {
                // there is a scenario where accents are added to the subscript node, not oMath
                // accent is not allowed in subscript elements
                foreach (OM.Subscript sSub in myDoc.MainDocumentPart.Document.Descendants<OM.Subscript>())
                {
                    // loop through the subscript elements and if "acc" is found, delete it
                    // use localname because the openxml type shows unknown
                    foreach (OpenXmlElement oxe in sSub.ChildElements)
                    {
                        if (oxe.LocalName == "acc")
                        {
                            oxe.Remove();
                            corruptionFound = true;
                        }
                    }
                }

                if (corruptionFound)
                {
                    myDoc.MainDocumentPart.Document.Save();
                }
            }

            return corruptionFound;
        }

        /// <summary>
        /// Flow has a corruption where the showingplaceholder tag is not removed after updating a content control
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool FixContentControlPlaceholders(string filePath)
        {
            corruptionFound = false;
            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filePath, true))
            {
                // loop content controls
                foreach (var cc in myDoc.ContentControls())
                {
                    bool containsDefaultText = false;

                    // 1. check for rsidR and rsidRPr
                    foreach (OpenXmlElement oxe in cc.ChildElements)
                    {
                        if (oxe.LocalName == "sdtContent")
                        {
                            foreach (OpenXmlElement oxeRun in oxe.ChildElements)
                            {
                                if (oxeRun.LocalName == "r")
                                {
                                    // Flow does not remove the placeholder style either, but for now, just look for the default text of a content control
                                    // if the text is different than the default, it should not be a placeholder and we can remove it
                                    if (oxeRun.InnerXml.Contains("Click or tap here to enter text."))
                                    {
                                        containsDefaultText = true;
                                    }
                                }
                            }
                        }
                    }

                    // 2. check for showingplchdr and remove if the text is default placeholder text
                    SdtProperties props = cc.Elements<SdtProperties>().FirstOrDefault();
                    if (props is not null)
                    {
                        foreach (OpenXmlElement oxe in props.ChildElements)
                        {
                            if (oxe.LocalName == "showingPlcHdr" && containsDefaultText == false)
                            {
                                oxe.Remove();
                                corruptionFound = true;
                            }
                        }
                    }
                }
                if (corruptionFound)
                {
                    myDoc.MainDocumentPart.Document.Save();
                }
            }
            return corruptionFound;
        }

        /// <summary>
        /// plain text content controls can't have any nested content controls
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool FixContentControls(string filePath, ref StringBuilder sb)
        {
            corruptionFound = false;
            string previousParentText = string.Empty;
            string currentParentText = string.Empty;

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filePath, true))
            {
                // loop content controls
                foreach (var cc in myDoc.ContentControls())
                {
                    // make sure we are a plain text control
                    bool plainTextControl = false;
                    SdtProperties props = cc.Elements<SdtProperties>().FirstOrDefault();
                    foreach (OpenXmlElement oxeProp in props.ChildElements)
                    {
                        if (oxeProp.GetType().Name == "SdtContentText")
                        {
                            plainTextControl = true;
                            previousParentText = currentParentText;
                            currentParentText = cc.Parent.InnerText;
                        }
                    }

                    // if it is a plain text and it has an sdtcontentrun, we need to remove it
                    if (plainTextControl)
                    {
                        foreach (OpenXmlElement oxe in cc.ChildElements)
                        {
                            if (oxe.GetType().Name == "SdtContentRun")
                            {
                                foreach (OpenXmlElement oxeInner in oxe.ChildElements)
                                {
                                    if (oxeInner.GetType().Name == "SdtRun")
                                    {
                                        oxeInner.Remove();
                                        corruptionFound = true;
                                        sb.AppendLine("Text before corruption = " + previousParentText);
                                        sb.AppendLine("Corrupt text that was deleted = " + currentParentText);
                                        previousParentText = string.Empty;
                                        currentParentText = string.Empty;
                                    }
                                }
                            }
                        }
                     }
                }

                if (corruptionFound)
                {
                    myDoc.MainDocumentPart.Document.Save();
                }
            }

            return corruptionFound;
        }

        /// <summary>
        /// check for shape in a hyperlink, in a comment, which is not allowed
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool FixShapeInComment(string filePath)
        {
            corruptionFound = false;

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filePath, true))
            {
                if (myDoc.MainDocumentPart.WordprocessingCommentsPart is null)
                {
                    return corruptionFound;
                }

                foreach (Comment cm in myDoc.MainDocumentPart.WordprocessingCommentsPart.Comments)
                {
                    IEnumerable<Hyperlink> hLinks = cm.Descendants<Hyperlink>();

                    // does the comment have a hyperlink
                    if (hLinks.Any())
                    {
                        foreach (Hyperlink h in hLinks)
                        {
                            // get the runs for the hyperlink
                            IEnumerable<Run> runs = h.Descendants<Run>();
                            if (runs.Any())
                            {
                                // if there is an E1o in the run, delete the run
                                foreach (Run r in runs)
                                {
                                    if (r.Descendants<AlternateContent>().Any())
                                    {
                                        r.Remove();
                                        corruptionFound = true;
                                    }
                                }
                            }
                        }
                    }
                }

                if (corruptionFound)
                {
                    myDoc.MainDocumentPart.Document.Save();
                }
            }

            return corruptionFound;
        }

        /// <summary>
        /// del tags can't be nested, check for this scenario and fix the tags
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool FixDeleteRevision(string filePath)
        {
            corruptionFound = false;

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filePath, true))
            {
                var delRuns = myDoc.MainDocumentPart.Document.Descendants<DeletedRun>().ToList();
                foreach (DeletedRun del in delRuns)
                {
                    // if the first element in a deletedrun is another delete element, that is not correct
                    if (del.FirstChild.LocalName == "del")
                    {
                        // clone the del child element, remove it and then append it after the root del element
                        OpenXmlElement oxe = del.FirstChild.CloneNode(true);
                        del.FirstChild.Remove();
                        del.Append(oxe);
                        corruptionFound = true;
                    }
                }

                if (corruptionFound)
                {
                    myDoc.MainDocumentPart.Document.Save();
                }
            }            

            return corruptionFound;
        }

        /// <summary>
        /// check for scenarios where a hyperlink is in the wrong part of a field code, which is not allowed
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool FixHyperlinks(string filePath)
        {
            bool fileChanged = false;

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filePath, true))
            {
                IEnumerable<Hyperlink> hLinks = myDoc.MainDocumentPart.Document.Descendants<Hyperlink>();
                
                if (hLinks.Any())
                {
                    for (int i = 0; i < hLinks.Count(); i++)
                    {
                        Hyperlink h = hLinks.ElementAt(i);
                        if (h.InnerXml.Contains(Strings.txtFieldCodeEnd) && !h.InnerXml.Contains(Strings.txtFieldCodeBegin))
                        {
                            h.InnerXml = h.InnerXml.Replace(Strings.txtFieldCodeEnd, string.Empty);
                            h.Parent.PrependChild(new FieldChar() { FieldCharType = FieldCharValues.End });
                            fileChanged = true;
                        }
                    }
                }

                if (myDoc.MainDocumentPart.HeaderParts is not null)
                {
                    foreach (HeaderPart hdrPart in myDoc.MainDocumentPart.HeaderParts)
                    {
                        if (hdrPart.Header is not null)
                        {
                            IEnumerable<Hyperlink> hdrLinks = hdrPart.Header.Descendants<Hyperlink>();
                            for (int i = 0; i < hdrLinks.Count(); i++)
                            {
                                Hyperlink h = hdrLinks.ElementAt(i);
                                if (h.InnerXml.Contains(Strings.txtFieldCodeEnd) && !h.InnerXml.Contains(Strings.txtFieldCodeBegin))
                                {
                                    h.InnerXml = h.InnerXml.Replace(Strings.txtFieldCodeEnd, string.Empty);
                                    h.Parent.PrependChild(new FieldChar() { FieldCharType = FieldCharValues.End });
                                    fileChanged = true;
                                }
                            }
                        }
                    }
                }

                if (myDoc.MainDocumentPart.FooterParts is not null)
                {
                    foreach (FooterPart ftrPart in myDoc.MainDocumentPart.FooterParts)
                    {
                        if (ftrPart.Footer is not null)
                        {
                            IEnumerable<Hyperlink> ftrLinks = ftrPart.Footer.Descendants<Hyperlink>();
                            for (int i = 0; i < ftrLinks.Count(); i++)
                            {
                                Hyperlink h = ftrLinks.ElementAt(i);
                                if (h.InnerXml.Contains(Strings.txtFieldCodeEnd) && !h.InnerXml.Contains(Strings.txtFieldCodeBegin))
                                {
                                    h.InnerXml = h.InnerXml.Replace(Strings.txtFieldCodeEnd, string.Empty);
                                    h.Parent.PrependChild(new FieldChar() { FieldCharType = FieldCharValues.End });
                                    fileChanged = true;
                                }
                            }
                        }
                    }
                }

                if (fileChanged)
                {
                    myDoc.MainDocumentPart.Document.Save();
                }
            }

            return fileChanged;
        }

        /// <summary>
        /// There are times when a comment is included inside a plain text content control
        /// this is not allowed
        /// </summary>
        /// <param name="filePath">File to attempt fix</param>
        /// <returns></returns>
        public static bool FixCommentRange(string filePath)
        {
            bool isFileChanged = false;

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filePath, true))
            {
                // loop content controls
                foreach (var cc in myDoc.ContentControls())
                {
                    bool plainTextControl = false;

                    SdtProperties props = cc.Elements<SdtProperties>().FirstOrDefault();
                    foreach (OpenXmlElement oxeProp in props.ChildElements)
                    {
                        if (oxeProp.GetType().Name == "SdtContentText")
                        {
                            plainTextControl = true;
                        }
                    }

                    // plaint text controls can't have comments
                    if (plainTextControl)
                    {
                        OpenXmlElement oxeCommentRangeStart = null;
                        OpenXmlElement oxeCommentRangeEnd = null;
                        OpenXmlElement oxeCommentReference = null;
                        OpenXmlElement oxeCommentRangeStartClone = null;
                        OpenXmlElement oxeCommentRangeEndClone = null;
                        OpenXmlElement oxeCommentReferenceClone = null;

                        foreach (OpenXmlElement oElement in cc.ChildElements)
                        {
                            // the problem is in the content tag
                            if (oElement.LocalName == "sdtContent")
                            {
                                // only do something if we have a comment range start or end
                                if (oElement.InnerXml.Contains("<w:commentRangeStart") || oElement.InnerXml.Contains("<w:commentRangeEnd"))
                                {
                                    foreach (OpenXmlElement el in oElement.ChildElements)
                                    {
                                        if (el.GetType().Name == "CommentRangeStart")
                                        {
                                            oxeCommentRangeStart = el;
                                            oxeCommentRangeStartClone = el.CloneNode(true);
                                        }

                                        if (el.GetType().Name == "CommentRangeEnd")
                                        {
                                            oxeCommentRangeEnd = el;
                                            oxeCommentRangeEndClone = el.CloneNode(true);
                                        }

                                        if (el.InnerXml.Contains("<w:commentReference"))
                                        {
                                            oxeCommentReference = el;
                                            oxeCommentReferenceClone = el.CloneNode(true);
                                        }
                                    }

                                    // if we have a comment range start, end and reference, remove them
                                    if (oxeCommentRangeStart != null)
                                    {
                                        oxeCommentRangeStart.Remove();
                                        isFileChanged = true;
                                    }

                                    if (oxeCommentRangeEnd != null)
                                    {
                                        oxeCommentRangeEnd.Remove();
                                        isFileChanged = true;
                                    }

                                    if (oxeCommentReference != null)
                                    {
                                        oxeCommentReference.Remove();
                                        isFileChanged = true;
                                    }
                                }
                            }
                        }

                        // move the comment range start before the content control
                        // move the comment range end and comment ref after the content control
                        cc.PrependChild(oxeCommentRangeStartClone);
                        cc.Parent.Append(oxeCommentRangeEndClone);
                        cc.Parent.Append(oxeCommentReferenceClone);
                    }
                }

                if (isFileChanged)
                {
                    myDoc.MainDocumentPart.Document.Save();
                }
            }

            return isFileChanged;
        }

        /// <summary>
        /// This looks for corrupt @mention style field codes in comments with missing Separate tags
        /// It will pull the email out and then add a new begin-separate-end sequence with the mailto info to fix the issue
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool FixCommentFieldCodes(string filePath)
        {
            bool isFileChanged = false;

            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(filePath, true))
            {
                Regex emailPattern = new Regex(@"(.*?)<?(\b\S+@\S+\b)>?");
                if (myDoc.MainDocumentPart.WordprocessingCommentsPart is null)
                {
                    return isFileChanged;
                }
                
                WordprocessingCommentsPart commentsPart = myDoc.MainDocumentPart.WordprocessingCommentsPart;

                foreach (Comment cmt in commentsPart.Comments)
                {
                    IEnumerable<Paragraph> pList = cmt.Descendants<Paragraph>();
                    List<Run> rList = new List<Run>();

                    foreach (Paragraph p in pList)
                    {
                        // if the p has the mention style, it passes the first check we need to make
                        if (p.InnerXml.Contains(Strings.txtAtMentionStyle))
                        {
                            bool beginFound = false;
                            bool separateFound = false;
                            string emailAlias = string.Empty;

                            // now we need to loop each run and check the separate is missing
                            foreach (Run r in p.Descendants<Run>())
                            {
                                if (r.InnerXml.Contains(Strings.txtFieldCodeBegin))
                                {
                                    beginFound = true;
                                }

                                if (r.InnerXml.Contains(Strings.txtFieldCodeSeparate))
                                {
                                    separateFound = true;
                                }

                                // hold onto the mailto so we at least have something to use for the mention text
                                if (beginFound == true && r.InnerXml.Contains("mailto"))
                                {
                                    var groups = emailPattern.Match(r.InnerXml.ToString()).Groups;

                                    // trim the beginning of the mailto
                                    emailAlias = groups[2].Value.Remove(0, 7);

                                    // remove the domain and add @ to the beginning
                                    int atIndex = emailAlias.IndexOf('@');
                                    emailAlias = "@" + emailAlias.Remove(atIndex);
                                }

                                // once we get to the end, if we haven't found a separate, we need to add it back
                                if (r.InnerXml.Contains(Strings.txtFieldCodeEnd))
                                {
                                    if (r.InnerXml.Contains(Strings.txtAtMentionStyle) && separateFound == false)
                                    {
                                        // first, remove all children since we are in the area we need to change
                                        // add separate to the existing run
                                        r.RemoveAllChildren();
                                        RunProperties rPr = new RunProperties();
                                        RunStyle rs = new RunStyle();
                                        rs.Val = "Mention";
                                        rPr.Append(rs);
                                        r.Append(rPr);
                                        FieldChar fcs = new FieldChar();
                                        fcs.FieldCharType = FieldCharValues.Separate;
                                        r.Append(fcs);

                                        // add a new text run with the mailto alias
                                        Run rNewEnd = new Run();
                                        FieldChar fce = new FieldChar();
                                        fce.FieldCharType = FieldCharValues.End;
                                        rNewEnd.Append(fce);
                                        r.InsertAfterSelf(rNewEnd);

                                        // add a new end run back to close it out
                                        Run rNewText = new Run();
                                        Text t = new Text();
                                        t.Text = emailAlias;
                                        rNewText.Append(t);
                                        r.InsertAfterSelf(rNewText);

                                        isFileChanged = true;
                                    }

                                    // reset logic criteria
                                    beginFound = false;
                                    separateFound = false;
                                    emailAlias = string.Empty;
                                }
                            }
                        }
                    }
                }

                if (isFileChanged)
                {
                    myDoc.MainDocumentPart.Document.Save();
                }
            }

            return isFileChanged;
        }

        /// <summary>
        /// this is used to fix orphaned comment references
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool FixMissingCommentRefs(string filePath)
        {
            bool saveFile = false;

            using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
            {
                WordprocessingCommentsPart commentsPart = document.MainDocumentPart.WordprocessingCommentsPart;
                IEnumerable<OpenXmlUnknownElement> unknownList = document.MainDocumentPart.Document.Descendants<OpenXmlUnknownElement>();
                IEnumerable<CommentReference> commentRefs = document.MainDocumentPart.Document.Descendants<CommentReference>();

                bool cRefIdExists = false;

                // if the comments part is gone, we can still fix up the orphan potentially
                // removing the comment refs if they exist
                if (commentsPart is null)
                {
                    // if there is no comments part and no commentrefs, the dangling unknownelements will have nothing to check against
                    // bail out in that case, otherwise loop the commentrefs and remove them so the file will open
                    if (commentRefs.Count() == 0)
                    {
                        return saveFile;
                    }
                    else
                    {
                        foreach (CommentReference cr in commentRefs)
                        {
                            cr.Remove();
                            saveFile = true;
                        }
                    }
                }
                else if (commentsPart is not null && commentRefs.Count() == 0)
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
                                                foreach (Comment cm in commentsPart.Comments)
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
                }
            }

            return saveFile;
        }

        /// <summary>
        /// there are times when the gridspan is a really large number, which makes the table appear invisible
        /// this function will look for that scenario and reset the gridspan to 1
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool FixGridSpan(string filePath)
        {
            bool tblModified = false;

            using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
            {
                // global document variables
                if (Word.IsPartNull(document, "Table") == false)
                {
                    // get the list of tables in the document
                    // need to account for nested tables
                    List<Table> tbls = document.MainDocumentPart.Document.Descendants<Table>().ToList();

                    foreach (Table tbl in tbls)
                    {
                        // get the column count
                        int colCount = 0;
                        foreach (OpenXmlElement oxe in tbl.ChildElements)
                        {
                            // get the count of columns
                            if (oxe.GetType().ToString() == Strings.dfowTableGrid)
                            {
                                colCount = oxe.ChildElements.Count;
                            }
                        }

                        // loop each table cell and check the gridSpan
                        List<GridSpan> gsList = tbl.Descendants<GridSpan>().ToList();
                        foreach (GridSpan gs in gsList) 
                        {
                            // if the gridspan is larger than the number of columns, reset to 1
                            if (gs.Val > colCount)
                            {
                                gs.Val = 1;
                                tblModified = true;
                            }
                        }
                    }
                }

                // save the file if we modified the table
                if (tblModified == true)
                {
                    document.MainDocumentPart.Document.Save();
                }
            }

            return tblModified;
        }

        /// <summary>
        /// tables can only have one table grid per table, this will fix that scenario
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool FixTableGridProps(string filePath)
        {
            bool tblModified = false;

            using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
            {
                // global document variables
                OpenXmlElement tgClone = null;

                if (Word.IsPartNull(document, "Table") == false)
                {
                    // get the list of tables in the document
                    List<Table> tbls = document.MainDocumentPart.Document.Descendants<Table>().ToList();

                    foreach (Table tbl in tbls)
                    {
                        // you can have only one tblGrid per table, including nested tables
                        // it needs to be before any row elements so sequence is
                        // 1. check if the tblGrid element is before any table row
                        // 2. check for multiple tblGrid elements
                        bool tRowFound = false;
                        bool tGridBeforeRowFound = false;
                        int tGridCount = 0;

                        foreach (OpenXmlElement oxe in tbl.Elements())
                        {
                            // flag if we found a trow, once we find 1, the rest do not matter
                            if (oxe.GetType().Name == "TableRow")
                            {
                                tRowFound = true;
                            }

                            // when we get to a tablegrid, we have a few things to check
                            // 1. have we found a table row previously
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

                        // if we had a table grid before a row was found, move it before the first row in the table
                        if (tGridBeforeRowFound == true)
                        {
                            tbl.InsertBefore(tgClone, tbl.GetFirstChild<TableRow>());
                            tblModified = true;
                        }
                    }
                }

                // save the file if we modified the table
                if (tblModified == true)
                {
                    document.MainDocumentPart.Document.Save();
                }
            }

            return tblModified;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool FixRevisions(string filePath)
        {
            bool isFixed = false;

            using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
            {
                DocumentFormat.OpenXml.Wordprocessing.Document doc = document.MainDocumentPart.Document;
                var deleted = doc.Descendants<DeletedRun>().ToList();

                // loop each DeletedRun
                foreach (DeletedRun dr in deleted)
                {
                    foreach (OpenXmlElement oxedr in dr)
                    {
                        // if we have a run, we need to look for Text tags
                        if (oxedr.GetType().ToString() == Strings.dfowRun)
                        {
                            Run r = (Run)oxedr;
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

                // now save the file if we have changes
                if (isFixed == true)
                {
                    doc.Save();
                }
            }

            return isFixed;
        }

        /// <summary>
        /// fix orphaned list templates in files
        /// this typically happens when you add and then remove some bullet/numbering in a document
        /// after removing the bullet for example, the list template (style) still exists in the file
        /// it will not get re-used however if you add a bullet later with the same data which orphans the style
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool FixListTemplates(string filePath)
        {
            bool orphanedListTemplatesFound = false;

            NumberingDetails bulletMultiLevelNumberingValues = new NumberingDetails();
            NumberingDetails bulletSingleLevelNumberingValues = new NumberingDetails();

            List<int> bulletMultiLevelNumIdsInUse = new List<int>();
            List<int> bulletSingleLevelNumIdsInUse = new List<int>();

            using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
            {
                // if the numbering part does not exist, nothing to check
                if (document.MainDocumentPart.NumberingDefinitionsPart is null)
                {
                    return orphanedListTemplatesFound;
                }

                // get the list of numId's and AbstractNum's in numbering.xml
                var absNumsInUseList = document.MainDocumentPart.NumberingDefinitionsPart.Numbering.Descendants<AbstractNum>().ToList();
                var numInstancesInUseList = document.MainDocumentPart.NumberingDefinitionsPart.Numbering.Descendants<NumberingInstance>().ToList();

                bool bulletSingleLevelFound = false;
                bool bulletMultiLevelFound = false;

                foreach (AbstractNum an in absNumsInUseList)
                {
                    foreach (NumberingInstance ni in numInstancesInUseList)
                    {
                        // if the abstractnum and numId match, they are the same listtemplate
                        if (ni.AbstractNumId.Val == an.AbstractNumberId.Value)
                        {
                            // get the level count
                            var lvlNumberingList = an.Descendants<Level>().ToList();

                            // since we have the list template, find out if it is a bullet
                            foreach (OpenXmlElement anChild in an)
                            {
                                if (anChild.GetType().ToString() == Strings.dfowLevel)
                                {
                                    Level lvl = (Level)anChild;

                                    // try to catch different "types" of numberingformat
                                    // for now, I'm only checking for a single and multi-level bullets
                                    if (lvl.NumberingFormat.Val == "bullet" && lvlNumberingList.Count > 1 && lvl.LevelIndex == 0)
                                    {
                                        // if level is > 1, this is a multi level list
                                        bulletMultiLevelNumIdsInUse.Add(ni.NumberID);

                                        if (bulletMultiLevelFound == false)
                                        {
                                            bulletMultiLevelNumberingValues.AbsNumId = ni.AbstractNumId.Val;
                                            bulletMultiLevelNumberingValues.NumFormat = "bulletMultiLevel";
                                            bulletMultiLevelNumberingValues.NumId = ni.NumberID;
                                            bulletMultiLevelFound = true;
                                        }
                                    }
                                    else if (lvl.NumberingFormat.Val == "bullet" && lvlNumberingList.Count == 1 && lvl.LevelIndex == 0)
                                    {
                                        // if level = 1, this is a single level list
                                        bulletSingleLevelNumIdsInUse.Add(ni.NumberID);

                                        if (bulletSingleLevelFound == false)
                                        {
                                            bulletSingleLevelNumberingValues.AbsNumId = ni.AbstractNumId.Val;
                                            bulletSingleLevelNumberingValues.NumFormat = "bulletSingle";
                                            bulletSingleLevelNumberingValues.NumId = ni.NumberID;
                                            bulletSingleLevelFound = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // now that we have bullet numids to use, we can apply it to each paragraph
                MainDocumentPart mainPart = document.MainDocumentPart;
                StyleDefinitionsPart stylePart = mainPart.StyleDefinitionsPart;

                foreach (OpenXmlElement el in mainPart.Document.Descendants<Paragraph>())
                {
                    if (el.Descendants<NumberingId>().Count() > 0)
                    {
                        foreach (NumberingId pNumId in el.Descendants<NumberingId>())
                        {
                            foreach (var o in bulletMultiLevelNumIdsInUse)
                            {
                                if (o == pNumId.Val)
                                {
                                    pNumId.Val = bulletMultiLevelNumberingValues.NumId;
                                }
                            }

                            foreach (var o in bulletSingleLevelNumIdsInUse)
                            {
                                if (o == pNumId.Val)
                                {
                                    pNumId.Val = bulletSingleLevelNumberingValues.NumId;
                                }
                            }
                        }
                    }
                }

                foreach (HeaderPart hdrPart in mainPart.HeaderParts)
                {
                    foreach (OpenXmlElement el in hdrPart.Header.Elements())
                    {
                        foreach (NumberingId hNumId in el.Descendants<NumberingId>())
                        {
                            foreach (var o in bulletMultiLevelNumIdsInUse)
                            {
                                if (o == hNumId.Val)
                                {
                                    hNumId.Val = bulletMultiLevelNumberingValues.NumId;
                                }
                            }

                            foreach (var o in bulletSingleLevelNumIdsInUse)
                            {
                                if (o == hNumId.Val)
                                {
                                    hNumId.Val = bulletSingleLevelNumberingValues.NumId;
                                }
                            }
                        }
                    }
                }

                foreach (FooterPart ftrPart in mainPart.FooterParts)
                {
                    foreach (OpenXmlElement el in ftrPart.Footer.Elements())
                    {
                        foreach (NumberingId fNumId in el.Descendants<NumberingId>())
                        {
                            foreach (var o in bulletMultiLevelNumIdsInUse)
                            {
                                if (o == fNumId.Val)
                                {
                                    fNumId.Val = bulletMultiLevelNumberingValues.NumId;
                                }
                            }

                            foreach (var o in bulletSingleLevelNumIdsInUse)
                            {
                                if (o == fNumId.Val)
                                {
                                    fNumId.Val = bulletSingleLevelNumberingValues.NumId;
                                }
                            }
                        }
                    }
                }

                foreach (OpenXmlElement el in stylePart.Styles.Elements())
                {
                    if (el.GetType().ToString() == Strings.dfowStyle)
                    {
                        string styleEl = el.GetAttribute("styleId", Strings.wordMainAttributeNamespace).Value;
                        int pStyle = Word.ParagraphsByStyleName(mainPart, styleEl).Count();

                        if (pStyle > 0)
                        {
                            foreach (NumberingId sEl in el.Descendants<NumberingId>())
                            {
                                foreach (var o in bulletMultiLevelNumIdsInUse)
                                {
                                    if (o == sEl.Val)
                                    {
                                        sEl.Val = bulletMultiLevelNumberingValues.NumId;
                                    }
                                }

                                foreach (var o in bulletSingleLevelNumIdsInUse)
                                {
                                    if (o == sEl.Val)
                                    {
                                        sEl.Val = bulletSingleLevelNumberingValues.NumId;
                                    }
                                }
                            }
                        }
                    }
                }

                document.MainDocumentPart.Document.Save();
                return orphanedListTemplatesFound;
            }
        }

        public static bool FixEndnotes(string filePath)
        {
            bool corruptEndnotesFound = false;

            using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, true))
            {
                if (document.MainDocumentPart.EndnotesPart != null)
                {
                    Endnotes ens = document.MainDocumentPart.EndnotesPart.Endnotes;

                    foreach (Endnote en in ens)
                    {
                        // get the paragraph list from the endnote, if it has more than 1000 runs of content
                        // delete it...need to find a way to check for dupes
                        // for now just deleting all but the first paragraph run
                        var paraList = en.Descendants<Paragraph>().ToList();
                        foreach (var p in paraList)
                        {
                            var rList = p.Descendants<Run>().ToList();
                            if (rList.Count > 1000)
                            {
                                int count = 0;
                                foreach (var r in rList)
                                {
                                    if (count > 0)
                                    {
                                        r.Remove();
                                        corruptEndnotesFound = true;
                                    }
                                    count++;
                                }
                            }
                        }
                    }
                }

                if (corruptEndnotesFound == true)
                {
                    document.MainDocumentPart.Document.Save();
                }
            }

            return corruptEndnotesFound;
        }

        /// <summary>
        /// there are times when documents have textboxes containing images with duplicate id's
        /// this will check for those duplicate id's and change them to a different number
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool FixTextboxes(string filePath)
        {
            bool isFixed = false;

            try
            {
                using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, true))
                {
                    IEnumerable<Drawing> drwList = package.MainDocumentPart.Document.Descendants<Drawing>();
                    List<uint> drwIdList = new List<uint>();

                    foreach (Drawing d in drwList)
                    {
                        if (d.Anchor is null)
                        {
                            continue;
                        }

                        // first get the list of id's
                        foreach (OpenXmlElement oxe in d.Anchor.ChildElements)
                        {
                            if (oxe.LocalName == "docPr")
                            {
                                Wp.DocProperties dProps = new Wp.DocProperties();
                                dProps = (Wp.DocProperties)oxe;
                                drwIdList.Add(dProps.Id);
                            }
                        }
                    }

                    restartDrawingLoop:
                    foreach (Drawing d in drwList)
                    {
                        if (d.Anchor is null)
                        {
                            continue;
                        }

                        foreach (OpenXmlElement oxe in d.Anchor.ChildElements)
                        {
                            if (oxe.LocalName == "docPr")
                            {
                                Wp.DocProperties dProps = new Wp.DocProperties();
                                dProps = (Wp.DocProperties)oxe;
                                foreach (uint i in drwIdList)
                                {
                                    if (i == dProps.Id)
                                    {
                                        // dupe id's are not allowed, change to a new value
                                        Random rnd = new Random();
                                        dProps.Id = dProps.Id + (uint)rnd.Next(1, 1000);
                                        
                                        oxe.Remove();
                                        d.Anchor.AddChild(dProps);
                                        isFixed = true;
                                        goto restartDrawingLoop;
                                    }
                                }
                            }
                        }
                    }

                    // save out the file
                    if (isFixed)
                    {
                        package.Save();
                    }
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "FixTextboxes Error: " + ex.Message);
                return isFixed;
            }

            return isFixed;
        }

        /// <summary>
        /// There are times when third parties fail to write out the vne attributes in the document.xml file correctly
        /// They add it to the ignorable attribute, but they don't write out the corresponding namespace
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool FixVneAttributes(string filePath)
        {
            bool isFixed = false;

            try
            {
                using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, true))
                {
                    // get the vne attribute list
                    OpenXmlAttribute vne = package.MainDocumentPart.Document.GetAttributes().First();

                    // check each vne attribute for the correct namespace
                    if (vne.LocalName == "Ignorable")
                    {
                        string[] vneArray = vne.Value.Split(' ');
                        List<KeyValuePair<string, string>> nslist = package.MainDocumentPart.Document.NamespaceDeclarations.ToList();
                        List<string> missingNamespaces = new List<string>();

                        foreach (string s in vneArray)
                        {
                            foreach (var v in nslist)
                            {
                                if (v.Value == s)
                                {
                                    missingNamespaces.Add(s);
                                }
                            }
                        }
                    }

                    return isFixed;
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "FixListStyles Error: " + ex.Message);
                return isFixed;
            }
        }

        /// <summary>
        /// this fix is for a known issue where list styles are somehow getting the style name changed to Normal
        /// the fix is to find those number styles and change them to a default list style
        /// </summary>
        /// <param name="filePath">file to check if a fix is needed</param>
        /// <returns></returns>
        public static bool FixListStyles(string filePath) 
        {
            bool isFixed = false;

            try
            {
                using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, true))
                {
                    if (package.MainDocumentPart.NumberingDefinitionsPart is null)
                    {
                        return isFixed;
                    }

                    // loop each number style and check for Normal as a style name
                    foreach (AbstractNum an in package.MainDocumentPart.NumberingDefinitionsPart.Numbering)
                    {
                        // loop the w:lvl elements
                        foreach (OpenXmlElement oxe in an.ChildElements) 
                        {
                            bool levelChanged = false;
                            OpenXmlElement lvlClone = oxe.CloneNode(true);
                            if (oxe.LocalName == "lvl")
                            {
                                Level l = (Level)lvlClone;

                                // some levels don't have a pStyleId, skip these elements
                                if (l.ParagraphStyleIdInLevel is null)
                                {
                                    continue;
                                }

                                // when we get to a number style that says Normal, we need to change it
                                if (l.ParagraphStyleIdInLevel.Val == "Normal")
                                {
                                    // loop the w:pPr elements, new name will be based on indent level value
                                    foreach (OpenXmlElement oxePr in lvlClone.ChildElements)
                                    {
                                        if (oxePr.LocalName == "pPr")
                                        {
                                            // look for the indent prop
                                            foreach (OpenXmlElement oxeIndent in oxePr.ChildElements)
                                            {
                                                // choose the new bullet style name depending on indent location
                                                if (oxeIndent.LocalName == "ind")
                                                {
                                                    Indentation lIndent = (Indentation)oxeIndent;
                                                    int leftIndent = Convert.ToInt32(lIndent.Left);
                                                    if (leftIndent >= 1080)
                                                    {
                                                        l.ParagraphStyleIdInLevel.Val = "ListBullet3";
                                                        isFixed = true;
                                                        levelChanged = true;
                                                    }
                                                    else if (leftIndent >= 720)
                                                    {
                                                        l.ParagraphStyleIdInLevel.Val = "ListBullet2";
                                                        isFixed = true;
                                                        levelChanged = true;
                                                    }
                                                    else
                                                    {
                                                        l.ParagraphStyleIdInLevel.Val = "ListBullet";
                                                        isFixed = true;
                                                        levelChanged = true;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                // write the level into the clone
                                if (levelChanged)
                                {
                                    oxe.Remove();
                                    an.AddChild(lvlClone);
                                }
                            }
                        }
                    }

                    // save out the file
                    if (isFixed)
                    {
                        package.Save();
                    }
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "FixListStyles Error: " + ex.Message);
                return isFixed;
            }

            return isFixed;
        }

        /// <summary>
        /// There are scenarios where a file is saved out with no block level content in a table cell tag
        /// <w:tc><w:tcPr/></tc>
        /// This is corrupt, at least 1 block level tag like a paragraph <p></p> should be in the tc tag
        /// This function will look for this condition and remove any table cell tags that have no block level tags
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public static bool FixCorruptTableCellTags(string filename)
        {
            bool isFixed = false;

            try
            {
                using (WordprocessingDocument document = WordprocessingDocument.Open(filename, true))
                {
                    if (Word.IsPartNull(document, "Table") == false)
                    {
                        bool tableCellCorruptionFound = false;

                        do
                        {
                            tableCellCorruptionFound = IsTableCellCorruptionFound(document);
                            if (tableCellCorruptionFound == true)
                            {
                                isFixed = true;
                            }
                        } while (tableCellCorruptionFound);
                    }

                    if (isFixed)
                    {
                        document.Save();
                    }
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "FixCorruptTableCellTags Error: " + ex.Message);
                return false;
            }

            return isFixed;
        }

        /// <summary>
        /// get the tablecells in the passed in doc and check for two known corruptions
        /// the first check is for "<w:tc />"
        /// the second is for "<w:tc ><w:tcPr></w:tc>"
        /// I'm essentially removing a row if any table cell is corrupt in either way
        /// TODO: when I can come up with a way to keep the row and fix the cell, I will update the function
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        public static bool IsTableCellCorruptionFound(WordprocessingDocument doc)
        {
            IEnumerable<TableCell> tcList = doc.MainDocumentPart.Document.Descendants<TableCell>().ToList();
            foreach (TableCell tc in tcList)
            {
                // no child elements is corruption, add paragraph
                if (tc.ChildElements.Count == 0)
                {
                    tc.Parent.Remove();
                    return true;
                }

                // if the tc has 1 element and it is tcPr, we must also be missing a block level ement, which is also corrupt
                if (tc.ChildElements.Count == 1 && tc.ChildElements.ElementAt(0).LocalName == "tcPr")
                {
                    tc.Parent.Remove();
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Sometimes bookmarks are added and the start/end tag is missing
        /// This function will try to find those orphan tags and remove them
        /// </summary>
        /// <param name="filename">file to be scanned</param>
        /// <returns>true for successful removal and false if none are found</returns>
        public static bool RemoveMissingBookmarkTags(string filename)
        {
            bool isFixed = false;

            try
            {
                using (WordprocessingDocument package = WordprocessingDocument.Open(filename, true))
                {
                    if (package.MainDocumentPart.WordprocessingCommentsPart is null) { return false; }
                    if (package.MainDocumentPart.WordprocessingCommentsPart.Comments is null) { return false; }

                    IEnumerable<BookmarkStart> bkStartList = package.MainDocumentPart.WordprocessingCommentsPart.Comments.Descendants<BookmarkStart>();
                    IEnumerable<BookmarkEnd> bkEndList = package.MainDocumentPart.WordprocessingCommentsPart.Comments.Descendants<BookmarkEnd>();

                    // create temp lists so we can loop and remove any that exist in both lists
                    // if we have a start and end, the bookmark is valid and we can remove the rest
                    List<string> bkStartTagIds = new List<string>();
                    List<string> bkEndTagIds = new List<string>();

                    // check each start and find if there is a matching end tag id
                    foreach (BookmarkStart bks in bkStartList)
                    {
                        foreach (BookmarkEnd bke in bkEndList)
                        {
                            if (bke.Id.ToString() == bks.Id.ToString())
                            {
                                bkStartTagIds.Add(bke.Id);
                            }
                        }
                    }

                    // now we can check if there is a end tag with a matching start tag id
                    foreach (BookmarkEnd bke in bkEndList)
                    {
                        foreach (BookmarkStart bks in bkStartList)
                        {
                            if (bks.Id.ToString() == bke.Id.ToString())
                            {
                                bkEndTagIds.Add(bks.Id);
                            }
                        }
                    }

                    // now that we know all the id's that match, we can loop again and remove id's that are not in the lists
                    // first check orphaned start tags
                    bool startTagFound = false;

                    foreach (BookmarkStart bks in bkStartList)
                    {
                        foreach (object o in bkEndTagIds)
                        {
                            // if the end tag matches and we can ignore doing anything
                            if (o.ToString() == bks.Id.ToString())
                            {
                                startTagFound = true;
                            }
                        }

                        // if we get here and no match was found, it is orphaned and we can delete
                        if (startTagFound == false)
                        {
                            bks.Remove();
                            isFixed = true;
                        }
                        else
                        {
                            // reset the value for the next start tag check
                            startTagFound = false;
                        }
                    }

                    // do the same check for end tags
                    bool endTagFound = false;

                    foreach (BookmarkEnd bke in bkEndList)
                    {
                        foreach (object o in bkStartTagIds)
                        {
                            if (o.ToString() == bke.Id.ToString())
                            {
                                endTagFound = true;
                            }
                        }

                        if (endTagFound == false)
                        {
                            bke.Remove();
                            isFixed = true;
                        }
                        else
                        {
                            endTagFound = false;
                        }
                    }

                    if (isFixed)
                    {
                        package.MainDocumentPart.Document.Save();
                    }
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "RemoveMissingBookmarkTags Error: " + ex.Message);
                return false;
            }

            return isFixed;
        }

        /// <summary>
        /// if only one of the bookmark tags exists in the content control run block, that is invalid
        /// this will check for that condition and remove the bookmark from the file
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public static bool FixBookmarkTagInSdtContent(string filename)
        {
            bool isFixed = false;

            try
            {
                using (WordprocessingDocument package = WordprocessingDocument.Open(filename, true))
                {
                    IEnumerable<BookmarkStart> bkStartList = package.MainDocumentPart.Document.Descendants<BookmarkStart>();
                    IEnumerable<BookmarkEnd> bkEndList = package.MainDocumentPart.Document.Descendants<BookmarkEnd>();
                    List<string> bkIdToRemove = new List<string>(); 

                    // loop the start tags, if the start is in a contentrun, flag it
                    foreach (BookmarkStart bks in bkStartList)
                    {
                        if (bks.Parent.ToString() == Strings.dfowSdtContent)
                        {
                            foreach (BookmarkEnd bke in bkEndList)
                            {
                                if (bke.Id == bks.Id && bke.Parent.ToString() != Strings.dfowSdtContent)
                                {
                                    bkIdToRemove.Add(bke.Id);
                                }
                            }
                        }
                    }

                    foreach (BookmarkEnd bke in bkEndList)
                    {
                        if (bke.Parent.ToString() == Strings.dfowSdtContent)
                        {
                            foreach (BookmarkStart bks in bkStartList)
                            {
                                if (bks.Id == bke.Id && bks.Parent.ToString() != Strings.dfowSdtContent)
                                {
                                    bkIdToRemove.Add(bks.Id);
                                }
                            }
                        }
                    }

                    foreach (string s in bkIdToRemove)
                    {
                        foreach (BookmarkStart bks in bkStartList)
                        {
                            if (s == bks.Id)
                            {
                                bks.Remove();
                                isFixed = true;
                            }
                        }

                        foreach (BookmarkEnd bke in bkEndList)
                        {
                            if (s == bke.Id)
                            {
                                bke.Remove();
                                isFixed = true;
                            }
                        }
                    }

                    if (isFixed)
                    {
                        package.Save();
                    }
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "FixBookmarkTagInSdtContent Error: " + ex.Message);
                return isFixed;
            }
            
            return isFixed;
        }

        // <summary>
        /// look for bookmark tags in a plain content control
        /// this is not allowed and those need to be removed
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public static bool RemovePlainTextCcFromBookmark(string filename)
        {
            bool isFixed = false;

            try
            {
                using (WordprocessingDocument package = WordprocessingDocument.Open(filename, true))
                {
                    IEnumerable<BookmarkStart> bkStartList = package.MainDocumentPart.Document.Descendants<BookmarkStart>();
                    IEnumerable<BookmarkEnd> bkEndList = package.MainDocumentPart.Document.Descendants<BookmarkEnd>();
                    List<string> removedBookmarkIds = new List<string>();

                    if (bkStartList.Count() > 0)
                    {
                        foreach (BookmarkStart bk in bkStartList)
                        {
                            var cElem = bk.Parent;
                            var pElem = bk.Parent;
                            bool endLoop = false;

                            do
                            {
                                // first check if we are a content control
                                if (cElem.Parent != null && cElem.Parent.ToString().Contains(Strings.dfowSdt))
                                {
                                    foreach (OpenXmlElement oxe in cElem.Parent.ChildElements)
                                    {
                                        // get the properties
                                        if (oxe.GetType().Name == "SdtProperties")
                                        {
                                            foreach (OpenXmlElement oxeSdtAlias in oxe)
                                            {
                                                // check for plain text
                                                if (oxeSdtAlias.GetType().Name == "SdtContentText")
                                                {
                                                    // if the parent is a plain text content control, bookmark is not allowed
                                                    // add the id to the list of bookmarks that need to be deleted
                                                    removedBookmarkIds.Add(bk.Id);
                                                    endLoop = true;
                                                }
                                            }
                                        }
                                    }

                                    // set the next element to the parent and continue moving up the element chain
                                    pElem = cElem.Parent;
                                    cElem = pElem;
                                }
                                else
                                {
                                    // if the next element is null, bail
                                    if (cElem is null || cElem.Parent is null)
                                    {
                                        endLoop = true;
                                    }
                                    else
                                    {
                                        // set pElem to the parent so we can check for the end of the loop
                                        // set cElem to the parent also so we can continue moving up the element chain
                                        pElem = cElem.Parent;
                                        cElem = pElem;

                                        // loop should continue until we get to the body element, then we can stop looping
                                        if (pElem.ToString() == Strings.dfowBody)
                                        {
                                            endLoop = true;
                                        }
                                    }
                                }
                            } while (endLoop == false);
                        }

                        // now that we have the list of bookmark id's to be removed
                        // loop each list and delete any bookmark that has a matching id
                        foreach (var o in removedBookmarkIds)
                        {
                            foreach (BookmarkStart bkStart in bkStartList)
                            {
                                if (bkStart.Id == o)
                                {
                                    bkStart.Remove();
                                }
                            }

                            foreach (BookmarkEnd bkEnd in bkEndList)
                            {
                                if (bkEnd.Id == o)
                                {
                                    bkEnd.Remove();
                                }
                            }
                        }

                        // save the part
                        package.MainDocumentPart.Document.Save();

                        // check if there were any fixes made and update the output display
                        if (removedBookmarkIds.Count > 0)
                        {
                            isFixed = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                FileUtilities.WriteToLog(Strings.fLogFilePath, "RemovePlainTextCcFromBookmark Error: " + ex.Message);
                return false;
            }

            return isFixed;
        }
    }
}
