namespace Office_File_Explorer.OpenMcdfExtensions.OLEProperties.Interfaces
{
    public interface IProperty : IBinarySerializable
    {

        object Value
        {
            get;
            set;
        }

        PropertyType PropertyType
        {
            get;
        }

    }
}
