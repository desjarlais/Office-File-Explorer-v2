using System.IO;

namespace Office_File_Explorer.OpenMcdfExtensions.OLEProperties.Interfaces
{
    public interface IBinarySerializable
    {
        void Write(BinaryWriter bw);
        void Read(BinaryReader br);
    }
}
