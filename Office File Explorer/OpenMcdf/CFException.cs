/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 * 
 * The Original Code is OpenMCDF - Compound Document Format library.
 * 
 * The Initial Developer of the Original Code is Federico Blaseotto.*/

using Office_File_Explorer.Helpers;
using System;

namespace Office_File_Explorer.OpenMcdf
{
    /// <summary>
    /// OpenMCDF base exception.
    /// </summary>
    public class CFException : Exception
    {
        public CFException() : base() { }

        public CFException(string message) : base(message, null) { FileUtilities.WriteToLog(Strings.fLogFilePath, message); }

        public CFException(string message, Exception innerException) : base(message, innerException) { FileUtilities.WriteToLog(Strings.fLogFilePath, message); }
    }

    /// <summary>
    /// Raised when a data setter/getter method is invoked
    /// on a stream or storage object after the disposal of the owner
    /// compound file object.
    /// </summary>
    public class CFDisposedException : CFException
    {
        public CFDisposedException() : base() { }

        public CFDisposedException(string message) : base(message, null) { FileUtilities.WriteToLog(Strings.fLogFilePath, message); }

        public CFDisposedException(string message, Exception innerException) : base(message, innerException) { FileUtilities.WriteToLog(Strings.fLogFilePath, message); }
    }

    /// <summary>
    /// Raised when opening a file with invalid header
    /// or not supported COM/OLE Structured storage version.
    /// </summary>
    public class CFFileFormatException : CFException
    {
        public CFFileFormatException() : base() { }

        public CFFileFormatException(string message) : base(message, null) { FileUtilities.WriteToLog(Strings.fLogFilePath, message); }

        public CFFileFormatException(string message, Exception innerException) : base(message, innerException) { FileUtilities.WriteToLog(Strings.fLogFilePath, message); }
    }

    /// <summary>
    /// Raised when a named stream or a storage object
    /// are not found in a parent storage.
    /// </summary>
    public class CFItemNotFound : CFException
    {
        public CFItemNotFound() : base("Entry not found") { }

        public CFItemNotFound(string message) : base(message, null) { FileUtilities.WriteToLog(Strings.fLogFilePath, message); }

        public CFItemNotFound(string message, Exception innerException) : base(message, innerException) { FileUtilities.WriteToLog(Strings.fLogFilePath, message); }
    }

    /// <summary>
    /// Raised when a method call is invalid for the current object state
    /// </summary>
    public class CFInvalidOperation : CFException
    {
        public CFInvalidOperation() : base() { }

        public CFInvalidOperation(string message) : base(message, null) { FileUtilities.WriteToLog(Strings.fLogFilePath, message); }

        public CFInvalidOperation(string message, Exception innerException) : base(message, innerException) { FileUtilities.WriteToLog(Strings.fLogFilePath, message); }
    }

    /// <summary>
    /// Raised when trying to add a duplicated CFItem
    /// </summary>
    /// <remarks>
    /// Items are compared by name as indicated by specs.
    /// Two items with the same name CANNOT be added within 
    /// the same storage or sub-storage. 
    /// </remarks>
    public class CFDuplicatedItemException : CFException
    {
        public CFDuplicatedItemException() : base() { }

        public CFDuplicatedItemException(string message) : base(message, null) { FileUtilities.WriteToLog(Strings.fLogFilePath, message); }

        public CFDuplicatedItemException(string message, Exception innerException) : base(message, innerException) { FileUtilities.WriteToLog(Strings.fLogFilePath, message); }
    }

    /// <summary>
    /// Raised when trying to load a Compound File with invalid, corrupted or mismatched fields (4.1 - specifications) 
    /// </summary>
    /// <remarks>
    /// This exception is NOT raised when Compound file has been opened with NO_VALIDATION_EXCEPTION option.
    /// </remarks>
    public class CFCorruptedFileException : CFException
    {
        public CFCorruptedFileException() : base() { }

        public CFCorruptedFileException(string message) : base(message, null) { FileUtilities.WriteToLog(Strings.fLogFilePath, message); }

        public CFCorruptedFileException(string message, Exception innerException) : base(message, innerException) { FileUtilities.WriteToLog(Strings.fLogFilePath, message); }
    }
}
