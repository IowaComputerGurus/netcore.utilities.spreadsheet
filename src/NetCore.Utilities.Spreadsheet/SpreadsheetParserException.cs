#nullable enable
using System;

namespace ICG.NetCore.Utilities.Spreadsheet;

/// <summary>
/// An exception when parsing an OpenXml Spreadsheet
/// </summary>
public class SpreadsheetParserException : Exception
{
    internal SpreadsheetParserException(string message) : base(message)
    {
            
    }

    internal SpreadsheetParserException(string message, Exception innerException) : base(message, innerException)
    {

    }
}