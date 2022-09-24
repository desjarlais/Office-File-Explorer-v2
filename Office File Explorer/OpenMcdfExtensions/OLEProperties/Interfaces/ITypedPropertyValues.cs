namespace Office_File_Explorer.OpenMcdfExtensions.OLEProperties.Interfaces
{
    public interface ITypedPropertyValue : IProperty
    {
        VTPropertyType VTType
        {
            get;
            //set;
        }

        PropertyDimensions PropertyDimensions
        {
            get;
        }

        bool IsVariant
        {
            get;
        }
    }
}
