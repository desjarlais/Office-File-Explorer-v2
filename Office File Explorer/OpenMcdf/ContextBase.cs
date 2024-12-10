﻿using System;
using System.Collections.Generic;
using System.ComponentModel.Design.Serialization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Office_File_Explorer.OpenMcdf
{
    /// <summary>
    /// Supports switching the <see cref="RootContext"/> object.
    /// </summary>
    public abstract class ContextBase
    {
        internal RootContextSite ContextSite { get; }

        internal RootContext Context => ContextSite.Context;

        internal ContextBase(RootContextSite site)
        {
            ContextSite = site;
        }
    }

}
