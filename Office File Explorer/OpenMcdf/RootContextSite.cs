using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Office_File_Explorer.OpenMcdf
{
    /// <summary>
    /// A site for the <see cref="RootContext"/> object, to allow switching streams.
    /// </summary>
    internal class RootContextSite
    {
        RootContext? context;

        internal RootContext Context => context!;

        internal void Switch(RootContext context)
        {
            this.context = context;
        }
    }

}
