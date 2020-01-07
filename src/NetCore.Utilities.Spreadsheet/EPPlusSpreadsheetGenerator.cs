using System;
using System.Collections.Generic;
using System.ComponentModel;
using OfficeOpenXml;

namespace ICG.NetCore.Utilities.Spreadsheet
{
    /// <summary>
    /// A concrete implementation of <see cref="ISpreadsheetGenerator"/> using the EPPlus project
    /// </summary>
    public class EPPlusSpreadsheetGenerator: ISpreadsheetGenerator
    {

        /// <inheritdoc />
        public byte[] CreateMultiSheetSpreadsheet(IEnumerable<ISpreadsheetConfiguration<object>> exportSheets)
        {
            //Validate inputs
            if (exportSheets == null)
                throw new ArgumentNullException(nameof(exportSheets));


            //Create the package and render each sheet
            using (var package = new ExcelPackage())
            {
                foreach (var exportConfiguration in exportSheets)
                {
                    var sheet = package.Workbook.Worksheets.Add(exportConfiguration.WorksheetName);
                    var currentDataRow =
                        CalculateDataHeaderRow(exportConfiguration.RenderTitle, exportConfiguration.RenderSubTitle);

                    //Render the data headers first, to establish the range of the sheet, plus column names
                    var headerNames = new List<string>();

                    //Run headers
                    var properties = TypeDescriptor.GetProperties(exportConfiguration.DataType);
                    for (var i = 0; i < properties.Count; i++)
                    {
                        sheet.Cells[currentDataRow, i + 1].Value = properties[i].DisplayName;
                        headerNames.Add(properties[i].DisplayName);

                        //Handle formats
                        if (properties[i].Attributes.Count <= 0) continue;
                        foreach (var attribute in properties[i].Attributes)
                        {
                            if (!(attribute is SpreadsheetColumnFormatAttribute detail))
                                continue;

                            sheet.Column(i + 1).Style.Numberformat.Format = GetFormatSpecifier(detail.Format);
                            break;
                        }
                    }

                    //Style the header cells
                    using (var headerRange = sheet.Cells[currentDataRow, 1, currentDataRow, headerNames.Count])
                    {
                        headerRange.Style.Font.Bold = true;
                        headerRange.Style.WrapText = true;
                    }

                    //Set title
                    if (exportConfiguration.RenderTitle)
                    {
                        //always at row 1
                        using (var titleRange = sheet.Cells[1, 1, 1, headerNames.Count])
                        {
                            titleRange.Merge = true;
                            titleRange.Value = exportConfiguration.DocumentTitle;
                            titleRange.Style.Font.Bold = true;
                            titleRange.Style.Font.Size = 14;
                        }
                    }

                    //Set Sub-Title
                    if (exportConfiguration.RenderSubTitle)
                    {
                        var subHeaderRow = currentDataRow - 1; //Just before header
                        using (var titleRange = sheet.Cells[subHeaderRow, 1, subHeaderRow, headerNames.Count])
                        {
                            titleRange.Merge = true;
                            titleRange.Value = exportConfiguration.DocumentSubTitle;
                            titleRange.Style.Font.Bold = true;
                            titleRange.Style.Font.Size = 12;
                        }
                    }

                    //Load data
                    foreach (var item in exportConfiguration.ExportData)
                    {
                        currentDataRow++; //Increment current row
                        for (var p = 0; p < properties.Count; p++)
                        {
                            sheet.Cells[currentDataRow, p + 1].Value = properties[p].GetValue(item);
                        }
                    }

                    //Auto Fit
                    sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                }

                //Return the final object
                return package.GetAsByteArray();
            }
        }

        /// <summary>
        ///     Creates a single worksheet document using the provided configuration information
        /// </summary>
        /// <typeparam name="T">The object type for exporting</typeparam>
        /// <param name="exportConfiguration">The loaded configuration</param>
        /// <returns>A completed MS Excel file</returns>
        public byte[] CreateSingleSheetSpreadsheet<T>(SpreadsheetConfiguration<T> exportConfiguration) where T : class
        {
            //Validate input
            if (string.IsNullOrWhiteSpace(exportConfiguration.WorksheetName))
                throw new ArgumentException("Worksheet name must be supplied", nameof(exportConfiguration.WorksheetName));

            if (exportConfiguration.ExportData == null)
                throw new ArgumentException("Export data must be specified", nameof(exportConfiguration.ExportData));

            if (exportConfiguration.RenderTitle && string.IsNullOrEmpty(exportConfiguration.DocumentTitle))
                throw new ArgumentException("Document Title is required when 'Render Title' is true",
                    nameof(exportConfiguration.DocumentTitle));

            if (exportConfiguration.RenderSubTitle && string.IsNullOrEmpty(exportConfiguration.DocumentSubTitle))
                throw new ArgumentException("Document Sub Title is required when 'Render Sub Title' is true",
                    nameof(exportConfiguration.DocumentSubTitle));

            //Create the package and render
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add(exportConfiguration.WorksheetName);
                var currentDataRow =
                    CalculateDataHeaderRow(exportConfiguration.RenderTitle, exportConfiguration.RenderSubTitle);

                //Render the data headers first, to establish the range of the sheet, plus column names
                var headerNames = new List<string>();

                //Run headers
                var properties = TypeDescriptor.GetProperties(typeof(T));
                for (var i = 0; i < properties.Count; i++)
                {
                    sheet.Cells[currentDataRow, i + 1].Value = properties[i].DisplayName;
                    headerNames.Add(properties[i].DisplayName);

                    //Handle formats
                    if (properties[i].Attributes.Count <= 0) continue;
                    foreach (var attribute in properties[i].Attributes)
                    {
                        if (!(attribute is SpreadsheetColumnFormatAttribute detail))
                            continue;

                        sheet.Column(i + 1).Style.Numberformat.Format = GetFormatSpecifier(detail.Format);
                        break;
                    }
                }

                //Style the header cells
                using (var headerRange = sheet.Cells[currentDataRow, 1, currentDataRow, headerNames.Count])
                {
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.WrapText = true;
                }

                //Set title
                if (exportConfiguration.RenderTitle)
                {
                    //always at row 1
                    using (var titleRange = sheet.Cells[1, 1, 1, headerNames.Count])
                    {
                        titleRange.Merge = true;
                        titleRange.Value = exportConfiguration.DocumentTitle;
                        titleRange.Style.Font.Bold = true;
                        titleRange.Style.Font.Size = 14;
                    }
                }

                //Set Sub-Title
                if (exportConfiguration.RenderSubTitle)
                {
                    var subHeaderRow = currentDataRow -1; //Just before header
                    using (var titleRange = sheet.Cells[subHeaderRow, 1, subHeaderRow, headerNames.Count])
                    {
                        titleRange.Merge = true;
                        titleRange.Value = exportConfiguration.DocumentSubTitle;
                        titleRange.Style.Font.Bold = true;
                        titleRange.Style.Font.Size = 12;
                    }
                }

                //Load data
                foreach(var item in exportConfiguration.ExportData)
                {
                    currentDataRow++; //Increment current row
                    for (var p = 0; p < properties.Count; p++)
                    {
                        sheet.Cells[currentDataRow, p + 1].Value = properties[p].GetValue(item);
                    }
                }
                
                //Auto Fit
                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();

                return package.GetAsByteArray();
            }
        }

        /// <summary>
        /// Calculates the position of the data header row
        /// </summary>
        /// <param name="showTitle">Are we displaying the title</param>
        /// <param name="showSubTitle">Awe we displaying the sub-title</param>
        /// <returns>The desired row</returns>
        public static int CalculateDataHeaderRow(bool showTitle, bool showSubTitle)
        {
            if (!showSubTitle && !showTitle)
                return 1; //Headers start at the top
            if (showSubTitle && showTitle)
                return 3; //Headers start after title & subtitle
            
            //Headers are after a header
            return 2;
        }

        /// <summary>
        /// Takes the user supplied format string and converts to known excel formats, if unknown format is given returns the user supplied format
        /// </summary>
        /// <param name="requestedFormat">The requested format</param>
        /// <returns>The format to use</returns>
        public static string GetFormatSpecifier(string requestedFormat)
        {
            switch (requestedFormat)
            {
                case "C":
                    return "\"$\"#,##0.00";
                case "D":
                    return "MM/dd/yyyy";
                default:
                    return requestedFormat;
            }
        }
    }
}
