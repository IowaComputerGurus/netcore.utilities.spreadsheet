using System.Collections.Generic;
using System.IO;

namespace ICG.NetCore.Utilities.Spreadsheet
{
    /// <summary>
    ///     Provides a function to import data from MS Excel files
    /// </summary>
    public interface ISpreadsheetParser
    {
        /// <summary>
        ///     Parses the provided document and returns a List of objects based on the input data
        /// </summary>
        /// <typeparam name="T">The type to use for importing</typeparam>
        /// <param name="fileStream">The contents of the Excel File (XLSX format)</param>
        /// <returns>The parsed information</returns>
        List<T> ParseDocument<T>(Stream fileStream) where T : new();

        /// <summary>
        /// Parses the provided document and returns a List of T objects based on the input data, using the specific worksheet number
        /// </summary>
        /// <typeparam name="T">The type to use for importing</typeparam>
        /// <param name="fileStream">The contents of the Excel File (XLSX format</param>
        /// <param name="worksheetNumber">The number for the worksheet, 1 based.</param>
        /// <returns>The parsed information</returns>
        List<T> ParseDocument<T>(Stream fileStream, int worksheetNumber) where T : new();

        /// <summary>
        /// Parses the provided document and returns a List of T objects based on the input data, using the specific worksheet number
        /// </summary>
        /// <typeparam name="T">The type to use for importing</typeparam>
        /// <param name="fileStream">The contents of the Excel File (XLSX format</param>
        /// <param name="worksheetNumber">The number for the worksheet, 1 based.</param>
        /// <param name="skipHeaderRow">If set to true will skip the first row of data as header information</param>
        /// <returns>The parsed information</returns>
        List<T> ParseDocument<T>(Stream fileStream, int worksheetNumber, bool skipHeaderRow) where T : new();
    }
}