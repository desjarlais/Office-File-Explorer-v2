using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Office_File_Explorer.OpenMcdf
{
    /// <summary>
    /// Encapsulates an entry in the File Allocation Table (FAT).
    /// </summary>
    internal record struct FatEntry(uint Index, uint Value)
    {
        public readonly bool IsFree => Value == SectorType.Free;

        [ExcludeFromCodeCoverage]
        public override readonly string ToString() => $"#{Index}: {Value}";
    }

}
