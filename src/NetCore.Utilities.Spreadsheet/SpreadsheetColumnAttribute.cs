using System;

namespace ICG.NetCore.Utilities.Spreadsheet;

/// <summary>
///     Controls how a property is mapped to a spreadsheet column
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class SpreadsheetColumnAttribute : Attribute
{
    /// <summary>
    ///     Initializes a new SpreadsheetColumn Attribute
    /// </summary>
    /// <param name="displayName">
    ///     Sets the display name of the column. If not provided, will fall back on the DisplayName
    ///     attribute.
    /// </param>
    /// <param name="width">Sets the width of the column</param>
    /// <param name="ignore">If true, the column will be excluded from the spreadsheet</param>
    /// <param name="format">Sets the format of the column data</param>
    public SpreadsheetColumnAttribute(string displayName = null, float width = 0, bool ignore = false,
        string format = null, string formula = null)
    {
        DisplayName = displayName;
        Width = width;
        Ignore = ignore;
        Format = format;
        Formula = formula;
    }

    /// <summary>
    ///     A custom name to use for the header when exported
    /// </summary>
    public string DisplayName { get; }

    /// <summary>
    ///     A custom width for the cell when exported
    /// </summary>
    /// <remarks>
    ///     If not set will default to Excel default of 10 if not auto-sized
    /// </remarks>
    public float Width { get; }

    /// <summary>
    ///     Should this column be ignored from export
    /// </summary>
    public bool Ignore { get; }

    /// <summary>
    ///     A custom format for the column. See <see cref="ColumnFormats" /> for valid values.
    /// </summary>
    public string Format { get; }

    /// <summary>
    ///     A custom formula for the column. For example SUM, MIN, MAX, etc.
    /// </summary>
    /// <remarks>
    ///     One additional data row will be added and an auto-created formula will be added to execute the formula at the
    ///     bottom with the range of the column
    /// </remarks>
    public string Formula { get; set; }
}

/// <summary>
///     Contains valid values for <see cref="SpreadsheetColumnAttribute.Format" />
/// </summary>
public static class ColumnFormats
{
    /// <summary>
    ///     Formats the column as date only
    /// </summary>
    public const string Date = "d";

    /// <summary>
    ///     Formats the column as currency
    /// </summary>
    public const string Currency = "c";

    /// <summary>
    ///     Formats the column as a number with no decimal places
    /// </summary>
    public const string Fixed0 = "f0";

    /// <summary>
    ///     Formats the column as a number with one decimal place
    /// </summary>
    public const string Fixed1 = "f1";

    /// <summary>
    ///     Formats the column as a number with two decimal places
    /// </summary>
    public const string Fixed2 = "f2";

    /// <summary>
    ///     formats the column as a number with three decimal places
    /// </summary>
    public const string Fixed3 = "f3";
}