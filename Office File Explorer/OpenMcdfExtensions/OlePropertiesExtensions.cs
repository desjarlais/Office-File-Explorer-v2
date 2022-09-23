using Office_File_Explorer.OpenMcdf;
using Office_File_Explorer.OpenMcdfExtensions.OLEProperties;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Office_File_Explorer.OpenMcdfExtensions
{
    public static class OlePropertiesExtensions
    {
        /// <summary>
        /// Returns Stream as an OLE Properties container (see [MS-OLEPS] document for specifications).
        /// Stream is usually a SummaryInfo  or a DocumentSummaryInfo compound stream.
        /// Application specific properties stream are also supported.
        /// Properties can be added, removed and modified.
        /// </summary>
        /// <param name="cfStream">Compound Stream to be read as OLE properties container</param>
        /// <returns><see cref="T:OpenMcdf.Extensions.OLEProperties">Ole properties container</see> to enumerate and manipulate properties</returns>
        /// <remarks>
        /// This method currently has following limitations:
        /// - only SIMPLE Property Sets are supported;
        /// - Scalar, Vector and Variant Vector are supported;
        /// - Array properties are NOT supported;
        /// - Property Stream is re-created on save;
        /// </remarks>
        public static OLEPropertiesContainer AsOLEPropertiesContainer(this CFStream cfStream)
        {
            return new OLEPropertiesContainer(cfStream);
        }
    }
}
