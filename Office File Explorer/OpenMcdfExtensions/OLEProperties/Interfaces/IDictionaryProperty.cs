using Microsoft.Graph.ExternalConnectors;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Office_File_Explorer.OpenMcdfExtensions.OLEProperties.Interfaces
{
    public interface IDictionaryProperty : IProperty
    {
        void Read(BinaryReader br);
        void Write(BinaryWriter bw);
    }
}
