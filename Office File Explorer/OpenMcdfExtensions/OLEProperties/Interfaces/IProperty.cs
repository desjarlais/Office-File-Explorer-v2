using Microsoft.Graph.ExternalConnectors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
