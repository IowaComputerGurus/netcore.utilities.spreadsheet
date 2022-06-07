using System.Collections.Generic;
using System.IO;

namespace ICG.NetCore.Utilities.Spreadsheet;
#nullable enable

/// <summary>
///     Generates MS Excel spreadsheets and returns the byte array representing the document
/// </summary>
public interface ISpreadsheetGenerator
{
    /// <summary>
    ///     Creates a single worksheet document using the provided configuration information
    /// </summary>
    /// <typeparam name="T">The object type for exporting</typeparam>
    /// <param name="exportConfiguration">The loaded configuration</param>
    /// <returns>A completed MS Excel file</returns>
    byte[] CreateSingleSheetSpreadsheet<T>(SpreadsheetConfiguration<T> exportConfiguration) where T : class;

    /// <inheritdoc cref="CreateSingleSheetSpreadsheet{T}(SpreadsheetConfiguration{T})"/>
    /// <param name="output">A writable stream to save the workbook to.</param>
    /// <param name="exportConfiguration">The loaded configuration</param>
    /// <returns>True if the workbook was successfully exported. False otherwise.</returns>
    bool CreateSingleSheetSpreadsheet<T>(Stream output, SpreadsheetConfiguration<T> exportConfiguration) where T : class;

    /// <summary>
    ///     Creates a workbook with multiple sheets using the provided configuration values
    /// </summary>
    /// <param name="exportSheets">The listing of sheets to include</param>
    /// <returns>A single workbook in Excel format</returns>
    byte[] CreateMultiSheetSpreadsheet(IEnumerable<ISpreadsheetConfiguration> exportSheets);

    /// <summary>
    ///     Creates a workbook with multiple sheets using the provided configuration values
    /// </summary>
    /// <param name="configuration">
    ///     A <see cref="MultisheetConfiguration"/> describing the sheets to export
    /// </param>
    /// <returns>A single workbook in Excel format</returns>
    byte[] CreateMultiSheetSpreadsheet(MultisheetConfiguration configuration);

    /// <inheritdoc cref="CreateMultiSheetSpreadsheet(IEnumerable{ISpreadsheetConfiguration})"/>
    /// <param name="output">A writable stream to save the workbook to.</param>
    /// <param name="exportSheets">The loaded configuration</param>
    /// <returns>True if the workbook was successfully exported. False otherwise.</returns>
    bool CreateMultiSheetSpreadsheet(Stream output, IEnumerable<ISpreadsheetConfiguration> exportSheets);

    /// <summary>
    ///     Creates a workbook with multiple sheets using the provided configuration values
    /// </summary>
    /// <param name="output">A writable stream to save the workbook to.</param>
    /// <param name="configuration">
    ///     A <see cref="MultisheetConfiguration"/> describing the sheets to export
    /// </param>
    /// <returns>True if the workbook was successfully exported. False otherwise.</returns>
    bool CreateMultiSheetSpreadsheet(Stream output, MultisheetConfiguration configuration);

}