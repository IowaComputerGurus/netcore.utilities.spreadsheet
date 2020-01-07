using System.Collections.Generic;

namespace ICG.NetCore.Utilities.Spreadsheet
{
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

        /// <summary>
        /// Creates a workbook with multiple sheets using the provided configuration values
        /// </summary>
        /// <param name="exportSheets">The listing of sheets to include</param>
        /// <returns>A single workbook in Excel format</returns>
        byte[] CreateMultiSheetSpreadsheet(IEnumerable<ISpreadsheetConfiguration<object>> exportSheets);
    }
}