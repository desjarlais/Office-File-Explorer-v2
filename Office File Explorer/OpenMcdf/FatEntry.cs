using System.Diagnostics.CodeAnalysis;

namespace Office_File_Explorer.OpenMcdf
{
    /// <summary>
    /// Encapsulates an entry in the File Allocation Table (FAT).
    /// </summary>
    internal record struct FatEntry(uint Index, uint Value)
    {
        [ExcludeFromCodeCoverage]
        public override readonly string ToString() => $"#{Index}: {Value}";
    }

}
