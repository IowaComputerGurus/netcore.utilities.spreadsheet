using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;
using FontSize = DocumentFormat.OpenXml.Spreadsheet.FontSize;

namespace ICG.NetCore.Utilities.Spreadsheet
{
    /// <summary>
    /// A concrete implementation of <see cref="ISpreadsheetGenerator"/> using the EPPlus project
    /// </summary>
    public class EPPlusSpreadsheetGenerator : ISpreadsheetGenerator
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

        private enum FontStyleIndex
        {
            Default = 0,
            Header = 1,
            SubHeader = 2,
            DataHeader = 3
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
            using (var documentStream = new MemoryStream())
            {
                //Create the document
                var spreadsheetDocument =
                    SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);

                //Add the workbook
                var workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                //Add a worksheet to it
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var data = new SheetData();
                worksheetPart.Worksheet = new Worksheet(data);

                //Add the sheet to the workbook
                var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                var sheet = new Sheet
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = exportConfiguration.WorksheetName
                };
                sheets.Append(sheet);

                //Setup our styles
                var stylesPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet();
                // blank font list
                stylesPart.Stylesheet.Fonts = new Fonts();
                stylesPart.Stylesheet.Fonts.Count = 4;
                stylesPart.Stylesheet.Fonts.AppendChild(new Font()
                {
                    FontSize = new FontSize { Val = 11 }
                });
                stylesPart.Stylesheet.Fonts.AppendChild(new Font()
                {
                    Bold = new Bold(),
                    FontSize = new FontSize() { Val = 14 }
                });
                stylesPart.Stylesheet.Fonts.AppendChild(new Font()
                {
                    Bold = new Bold(),
                    FontSize = new FontSize() { Val = 12 }
                });
                stylesPart.Stylesheet.Fonts.AppendChild(new Font()
                {
                    Bold = new Bold()
                });

                // create fills
                stylesPart.Stylesheet.Fills = new Fills();

                // create a solid red fill
                var solidRed = new PatternFill() { PatternType = PatternValues.Solid };
                solidRed.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FFFF0000") }; // red fill
                solidRed.BackgroundColor = new BackgroundColor { Indexed = 64 };

                stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); // required, reserved by Excel
                stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); // required, reserved by Excel
                stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = solidRed });
                stylesPart.Stylesheet.Fills.Count = 3;

                // blank border list
                stylesPart.Stylesheet.Borders = new Borders();
                stylesPart.Stylesheet.Borders.Count = 1;
                stylesPart.Stylesheet.Borders.AppendChild(new Border());

                // blank cell format list
                stylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats();
                stylesPart.Stylesheet.CellStyleFormats.Count = 1;
                stylesPart.Stylesheet.CellStyleFormats.AppendChild(new CellFormat());

                // cell format list
                stylesPart.Stylesheet.CellFormats = new CellFormats();
                // empty one for index 0, seems to be required
                stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat()); //Default
                // cell format references style format 0, font 0, border 0, fill 2 and applies the fill
                //Header
                stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 1, BorderId = 0, FillId = 0 });//.AppendChild(new Alignment { Horizontal = HorizontalAlignmentValues.Center });
                //Sub-header
                stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat
                { FormatId = 0, FontId = 2, BorderId = 0, FillId = 0 });
                //Data-header
                stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat
                { FormatId = 0, FontId = 3, BorderId = 0, FillId = 0 });

                stylesPart.Stylesheet.CellFormats.Count = 3;

                stylesPart.Stylesheet.Save();



                //Build out our sheet information
                UInt32 currentRow = 1;
                if (exportConfiguration.RenderTitle)
                {
                    var row = new Row { RowIndex = 1 };
                    var headerCell = new Cell
                    {
                        CellReference = $"A{currentRow}",
                        CellValue = new CellValue(exportConfiguration.DocumentTitle),
                        DataType = CellValues.String,
                        StyleIndex = (int)FontStyleIndex.Header
                    };
                    row.Append(headerCell);
                    data.Append(row);
                    //Increment row
                    currentRow++;
                }

                if (exportConfiguration.RenderSubTitle)
                {
                    var row = new Row { RowIndex = currentRow };
                    var headerCell = new Cell
                    {
                        CellReference = $"A{currentRow}",
                        CellValue = new CellValue(exportConfiguration.DocumentSubTitle),
                        DataType = CellValues.String,
                        StyleIndex = (int)FontStyleIndex.SubHeader
                    };
                    row.Append(headerCell);
                    data.Append(row);
                    //Increment row
                    currentRow++;
                }

                //Run data headers
                var headerNames = new List<string>();
                //TODO: AUto Size - https://stackoverflow.com/questions/18268620/openxml-auto-size-column-width-in-excel
                //Run headers
                var properties = TypeDescriptor.GetProperties(typeof(T));
                var headerRow = new Row { RowIndex = currentRow };
                for (var i = 0; i < properties.Count; i++)
                {
                    var headerCell = new Cell
                    {
                        CellValue = new CellValue(properties[i].DisplayName),
                        DataType = CellValues.String,
                        StyleIndex = (int)FontStyleIndex.DataHeader
                    };
                    headerRow.Append(headerCell);
                    headerNames.Add(properties[i].DisplayName);


                    //Handle formats
                    //if (properties[i].Attributes.Count <= 0) continue;
                    //foreach (var attribute in properties[i].Attributes)
                    //{
                    //    if (!(attribute is SpreadsheetColumnFormatAttribute detail))
                    //        continue;

                    //    sheet.Column(i + 1).Style.Numberformat.Format = GetFormatSpecifier(detail.Format);
                    //    break;
                    //}
                }
                data.Append(headerRow);
                currentRow++;

                //Run the data
                foreach (var item in exportConfiguration.ExportData)
                {
                    var dataRow = new Row {RowIndex = currentRow};
                    for (var p = 0; p < properties.Count; p++)
                    {
                        var dataCell = new Cell
                        {
                            CellValue = new CellValue(properties[0].GetValue(item).ToString()),
                            DataType = CellValues.String
                        };
                        dataRow.Append(dataCell);
                    }

                    data.Append(dataRow);
                    currentRow++;
                }

                workbookPart.Workbook.Save();
                spreadsheetDocument.Close();

                //Return the bytearray
                documentStream.Seek(0, SeekOrigin.Begin);
                return documentStream.ToArray();




                

                ////Auto Fit
                //sheet.Cells[sheet.Dimension.Address].AutoFitColumns();

                //return package.GetAsByteArray();
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
