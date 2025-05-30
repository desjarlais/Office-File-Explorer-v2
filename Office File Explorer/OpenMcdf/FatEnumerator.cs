﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace Office_File_Explorer.OpenMcdf
{
    /// <summary>
    /// Enumerates the <see cref="FatEntry"/> records in a <see cref="Fat"/>.
    /// </summary>
    internal class FatEnumerator : IEnumerator<FatEntry>
    {
        readonly Fat fat;
        bool start = true;
        uint index = uint.MaxValue;
        uint value = uint.MaxValue;

        public FatEnumerator(Fat fat)
        {
            this.fat = fat;
        }

        /// <inheritdoc/>
        public void Dispose()
        {
        }

        /// <inheritdoc/>
        public FatEntry Current
        {
            get
            {
                if (index == uint.MaxValue)
                    throw new InvalidOperationException("Enumeration has not started. Call MoveNext.");
                return new(index, value);
            }
        }

        /// <inheritdoc/>
        object IEnumerator.Current => Current;

        /// <inheritdoc/>
        public bool MoveNext()
        {
            if (start)
            {
                start = false;
                return MoveTo(0);
            }

            if (index >= SectorType.Maximum)
                return false;

            uint next = index + 1;
            return MoveTo(next);
        }

        public bool MoveTo(uint index)
        {
            ThrowHelper.ThrowIfSectorIdIsInvalid(index);

            start = false;
            if (this.index == index)
                return true;

            if (fat.TryGetValue(index, out value))
            {
                if (value < SectorType.Maximum && value >= fat.Context.SectorCount)
                    throw new FileFormatException($"FAT entry #{index} for sector {value} is beyond the end of the stream.");
                this.index = index;
                return true;
            }

            this.index = uint.MaxValue;
            return false;
        }

        public bool MoveNextFreeEntry()
        {
            while (MoveNext())
            {
                if (value is SectorType.Free)
                    return true;
            }

            return false;
        }

        /// <inheritdoc/>
        public void Reset()
        {
            start = true;
            index = uint.MaxValue;
            value = uint.MaxValue;
        }

        [ExcludeFromCodeCoverage]
        public override string ToString() => $"{Current}";
    }

}
