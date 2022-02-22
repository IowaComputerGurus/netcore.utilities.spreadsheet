using System;

namespace ICG.NetCore.Utilities.Spreadsheet;

/// <summary>
///     Used to decorate columns of an object for importing from Excel
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class SpreadsheetImportColumnAttribute : Attribute
{
    /// <summary>
    ///     Default constructor with the required Column Index
    /// </summary>
    /// <param name="columnIndex">The index position of the column, 1 based</param>
    public SpreadsheetImportColumnAttribute(int columnIndex)
    {
        ColumnIndex = columnIndex;
    }

    /// <summary>
    ///     The index position of the column within the data file
    /// </summary>
    public int ColumnIndex { get; }
}