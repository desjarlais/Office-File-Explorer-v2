using System.IO;

namespace Office_File_Explorer.OpenMcdfExtensions.OLEProperties.Interfaces
{
    public interface IDictionaryProperty : IProperty
    {
        new void Read(BinaryReader br);
        new void Write(BinaryWriter bw);
    }
}
