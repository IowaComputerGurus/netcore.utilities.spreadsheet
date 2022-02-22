using System;

namespace ICG.NetCore.Utilities.Spreadsheet;

/// <summary>
///     Using this custom attribute you are able to specify a column format for Spreadsheet exporting
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
[Obsolete("Replaced by the new SpreadsheetColumn, will be removed in version 7.0.0")]
public class SpreadsheetColumnFormatAttribute : Attribute
{
    /// <summary>
    ///     Constructor for the attribute allowing the specification of the format
    /// </summary>
    /// <param name="format">The target format</param>
    public SpreadsheetColumnFormatAttribute(string format)
    {
        Format = format;
    }

    /// <summary>
    ///     The desired format for the column
    /// </summary>
    public string Format { get; }
}