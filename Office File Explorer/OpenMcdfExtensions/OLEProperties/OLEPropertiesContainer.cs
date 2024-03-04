using Office_File_Explorer.OpenMcdf;
using Office_File_Explorer.OpenMcdfExtensions.OLEProperties.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Office_File_Explorer.OpenMcdfExtensions.OLEProperties
{
    public class OLEPropertiesContainer
    {
        public Dictionary<uint, string> PropertyNames = null;

        public OLEPropertiesContainer UserDefinedProperties { get; private set; }

        public bool HasUserDefinedProperties { get; private set; }

        public ContainerType ContainerType { get; internal set; }

        public PropertyContext Context { get; private set; }

        private List<OLEProperty> properties = new List<OLEProperty>();
        internal CFStream cfStream;

        /*
         Property name	Property ID	PID	Type
        Codepage	PID_CODEPAGE	1	VT_I2
        Title	PID_TITLE	2	VT_LPSTR
        Subject	PID_SUBJECT	3	VT_LPSTR
        Author	PID_AUTHOR	4	VT_LPSTR
        Keywords	PID_KEYWORDS	5	VT_LPSTR
        Comments	PID_COMMENTS	6	VT_LPSTR
        Template	PID_TEMPLATE	7	VT_LPSTR
        Last Saved By	PID_LASTAUTHOR	8	VT_LPSTR
        Revision Number	PID_REVNUMBER	9	VT_LPSTR
        Last Printed	PID_LASTPRINTED	11	VT_FILETIME
        Create Time/Date	PID_CREATE_DTM	12	VT_FILETIME
        Last Save Time/Date	PID_LASTSAVE_DTM	13	VT_FILETIME
        Page Count	PID_PAGECOUNT	14	VT_I4
        Word Count	PID_WORDCOUNT	15	VT_I4
        Character Count	PID_CHARCOUNT	16	VT_I4
        Creating Application	PID_APPNAME	18	VT_LPSTR
        Security	PID_SECURITY	19	VT_I4
             */
        public class SummaryInfoProperties
        {
            public short CodePage { get; set; }
            public string Title { get; set; }
            public string Subject { get; set; }
            public string Author { get; set; }
            public string KeyWords { get; set; }
            public string Comments { get; set; }
            public string Template { get; set; }
            public string LastSavedBy { get; set; }
            public string RevisionNumber { get; set; }
            public DateTime LastPrinted { get; set; }
            public DateTime CreateTime { get; set; }
            public DateTime LastSavedTime { get; set; }
            public int PageCount { get; set; }
            public int WordCount { get; set; }
            public int CharacterCount { get; set; }
            public string CreatingApplication { get; set; }
            public int Security { get; set; }
        }

        public static OLEPropertiesContainer CreateNewSummaryInfo(SummaryInfoProperties sumInfoProps)
        {
            return null;
        }

        public OLEPropertiesContainer(int codePage, ContainerType containerType)
        {
            Context = new PropertyContext
            {
                CodePage = codePage,
                Behavior = Behavior.CaseInsensitive
            };

            ContainerType = containerType;
        }

        internal OLEPropertiesContainer(CFStream cfStream)
        {
            PropertySetStream pStream = new PropertySetStream();

            this.cfStream = cfStream;
            pStream.Read(new BinaryReader(new StreamDecorator(cfStream)));

            switch (pStream.FMTID0.ToString("B").ToUpperInvariant())
            {
                case WellKnownFMTID.FMTID_SummaryInformation:
                    ContainerType = ContainerType.SummaryInfo;
                    break;
                case WellKnownFMTID.FMTID_DocSummaryInformation:
                    ContainerType = ContainerType.DocumentSummaryInfo;
                    break;
                default:
                    ContainerType = ContainerType.AppSpecific;
                    break;
            }

            PropertyNames = (Dictionary<uint, string>)pStream.PropertySet0.Properties
                .Where(p => p.PropertyType == PropertyType.DictionaryProperty).FirstOrDefault()?.Value;

            Context = new PropertyContext()
            {
                CodePage = pStream.PropertySet0.PropertyContext.CodePage
            };

            for (int i = 0; i < pStream.PropertySet0.Properties.Count; i++)
            {
                if (pStream.PropertySet0.PropertyIdentifierAndOffsets[i].PropertyIdentifier == 0) continue;
                //if (pStream.PropertySet0.PropertyIdentifierAndOffsets[i].PropertyIdentifier == 1) continue;
                //if (pStream.PropertySet0.PropertyIdentifierAndOffsets[i].PropertyIdentifier == 0x80000000) continue;

                var p = (ITypedPropertyValue)pStream.PropertySet0.Properties[i];
                var poi = pStream.PropertySet0.PropertyIdentifierAndOffsets[i];

                var op = new OLEProperty(this)
                {
                    VTType = p.VTType,
                    PropertyIdentifier = (uint)pStream.PropertySet0.PropertyIdentifierAndOffsets[i].PropertyIdentifier,
                    Value = p.Value
                };

                properties.Add(op);
            }

            if (pStream.NumPropertySets == 2)
            {
                UserDefinedProperties = new OLEPropertiesContainer(pStream.PropertySet1.PropertyContext.CodePage, ContainerType.UserDefinedProperties);
                HasUserDefinedProperties = true;

                UserDefinedProperties.ContainerType = ContainerType.UserDefinedProperties;

                for (int i = 0; i < pStream.PropertySet1.Properties.Count; i++)
                {
                    if (pStream.PropertySet1.PropertyIdentifierAndOffsets[i].PropertyIdentifier == 0) continue;
                    //if (pStream.PropertySet1.PropertyIdentifierAndOffsets[i].PropertyIdentifier == 1) continue;
                    if (pStream.PropertySet1.PropertyIdentifierAndOffsets[i].PropertyIdentifier == 0x80000000) continue;

                    var p = (ITypedPropertyValue)pStream.PropertySet1.Properties[i];
                    var poi = pStream.PropertySet1.PropertyIdentifierAndOffsets[i];
                    var op = new OLEProperty(UserDefinedProperties);

                    op.VTType = p.VTType;
                    op.PropertyIdentifier = (uint)pStream.PropertySet1.PropertyIdentifierAndOffsets[i].PropertyIdentifier;
                    op.Value = p.Value;

                    UserDefinedProperties.properties.Add(op);
                }

                var existingPropertyNames = (Dictionary<uint, string>)pStream.PropertySet1.Properties
                    .Where(p => p.PropertyType == PropertyType.DictionaryProperty).FirstOrDefault()?.Value;

                UserDefinedProperties.PropertyNames = existingPropertyNames ?? new Dictionary<uint, string>();
            }
        }

        public IEnumerable<OLEProperty> Properties
        {
            get { return properties; }
        }

        public OLEProperty NewProperty(VTPropertyType vtPropertyType, uint propertyIdentifier, string propertyName = null)
        {
            var op = new OLEProperty(this)
            {
                VTType = vtPropertyType,
                PropertyIdentifier = propertyIdentifier
            };

            return op;
        }

        public void AddProperty(OLEProperty property)
        {
            properties.Add(property);
        }

        public void RemoveProperty(uint propertyIdentifier)
        {
            var toRemove = properties.Where(o => o.PropertyIdentifier == propertyIdentifier).FirstOrDefault();
            if (toRemove != null) properties.Remove(toRemove);
        }


        public void Save(CFStream cfStream)
        {
            Stream s = new StreamDecorator(cfStream);
            BinaryWriter bw = new BinaryWriter(s);

            PropertySetStream ps = new PropertySetStream
            {
                ByteOrder = 0xFFFE,
                Version = 0,
                SystemIdentifier = 0x00020006,
                CLSID = Guid.Empty,

                NumPropertySets = 1,

                FMTID0 = ContainerType == ContainerType.SummaryInfo ? new Guid(WellKnownFMTID.FMTID_SummaryInformation) : new Guid(WellKnownFMTID.FMTID_DocSummaryInformation),
                Offset0 = 0,

                FMTID1 = Guid.Empty,
                Offset1 = 0,

                PropertySet0 = new PropertySet
                {
                    NumProperties = (uint)Properties.Count(),
                    PropertyIdentifierAndOffsets = new List<PropertyIdentifierAndOffset>(),
                    Properties = new List<IProperty>(),
                    PropertyContext = Context
                }
            };

            foreach (var op in Properties)
            {
                ITypedPropertyValue p = PropertyFactory.Instance.NewProperty(op.VTType, Context.CodePage);
                p.Value = op.Value;
                ps.PropertySet0.Properties.Add(p);
                ps.PropertySet0.PropertyIdentifierAndOffsets.Add(new PropertyIdentifierAndOffset() { PropertyIdentifier = op.PropertyIdentifier, Offset = 0 });
            }

            ps.PropertySet0.NumProperties = (uint)Properties.Count();

            if (HasUserDefinedProperties)
            {
                ps.NumPropertySets = 2;

                ps.PropertySet1 = new PropertySet
                {
                    // Number of user defined properties, plus 1 for the name dictionary
                    NumProperties = (uint)UserDefinedProperties.Properties.Count(),
                    PropertyIdentifierAndOffsets = new List<PropertyIdentifierAndOffset>(),
                    Properties = new List<IProperty>(),
                    PropertyContext = UserDefinedProperties.Context
                };

                ps.FMTID1 = new Guid("{D5CDD502-2E9C-101B-9397-08002B2CF9AE}");
                ps.Offset1 = 0;

                // Add the dictionary containing the property names
                IDictionaryProperty dictionaryProperty = new DictionaryProperty(ps.PropertySet1.PropertyContext.CodePage)
                {
                    Value = UserDefinedProperties.PropertyNames
                };
                ps.PropertySet1.Properties.Add(dictionaryProperty);
                ps.PropertySet1.PropertyIdentifierAndOffsets.Add(new PropertyIdentifierAndOffset() { PropertyIdentifier = 0, Offset = 0 });

                foreach (var op in UserDefinedProperties.Properties)
                {
                    ITypedPropertyValue p = PropertyFactory.Instance.NewProperty(op.VTType, ps.PropertySet1.PropertyContext.CodePage);
                    p.Value = op.Value;
                    ps.PropertySet1.Properties.Add(p);
                    ps.PropertySet1.PropertyIdentifierAndOffsets.Add(new PropertyIdentifierAndOffset() { PropertyIdentifier = op.PropertyIdentifier, Offset = 0 });

                }
            }

            ps.Write(bw);
        }
    }
}
