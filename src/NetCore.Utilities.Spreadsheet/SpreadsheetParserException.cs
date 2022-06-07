#nullable enable
using System;
using System.Runtime.Serialization;

namespace ICG.NetCore.Utilities.Spreadsheet;

/// <summary>
/// An exception when parsing an OpenXml Spreadsheet
/// </summary>
[Serializable]
public class SpreadsheetParserException : Exception
{
    /// <inheritdoc />
    protected SpreadsheetParserException(SerializationInfo info, StreamingContext context) : base(info, context)
    {
    }

    internal SpreadsheetParserException(string message) : base(message)
    {
            
    }

    internal SpreadsheetParserException(string message, Exception innerException) : base(message, innerException)
    {

    }
}