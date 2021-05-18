using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
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

        /// <summary>
        /// Internal helper for tracking font-style indexes
        /// </summary>
        private enum FontStyleIndex
        {
            /// <summary>
            /// Default cell text
            /// </summary>
            Default = 0,
            /// <summary>
            /// Document header formatting
            /// </summary>
            Header = 1,
            /// <summary>
            /// Document sub-header formatting
            /// </summary>
            SubHeader = 2,
            /// <summary>
            /// Headers for actual data.
            /// </summary>
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

            //Setup a memory stream to hold the generated file
            using (var documentStream = new MemoryStream())
            {
                //Create the document & overall workbook
                var spreadsheetDocument = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
                var workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                
                //Setup our styles
                var stylesPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet();
                // blank font list
                stylesPart.Stylesheet.Fonts = new Fonts {Count = 4};
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
                var solidRed = new PatternFill
                {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = new ForegroundColor {Rgb = HexBinaryValue.FromString("FFFF0000")},
                    BackgroundColor = new BackgroundColor {Indexed = 64}
                };
                // red fill

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
                
                stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat()); //Default

                //Header
                stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 1, BorderId = 0, FillId = 0 });//.AppendChild(new Alignment { Horizontal = HorizontalAlignmentValues.Center });
                //Sub-header
                stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat
                { FormatId = 0, FontId = 2, BorderId = 0, FillId = 0 });
                //Data-header
                stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat
                { FormatId = 0, FontId = 3, BorderId = 0, FillId = 0 });

                stylesPart.Stylesheet.CellFormats.Count = 4;
                stylesPart.Stylesheet.Save();


                //Build out our sheet information
                var data = new SheetData();
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
                //TODO: AUto Size - https://stackoverflow.com/questions/18268620/openxml-auto-size-column-width-in-excel
                //Run headers
                var properties = typeof(T).GetProperties();
                var headerProperties = TypeDescriptor.GetProperties(typeof(T));
                var headerRow = new Row { RowIndex = currentRow };
                foreach(PropertyDescriptor prop in headerProperties)
                {
                    var headerCell = new Cell
                    {
                        CellValue = new CellValue(prop.DisplayName),
                        DataType = CellValues.String,
                        StyleIndex = (int)FontStyleIndex.DataHeader
                    };
                    headerRow.Append(headerCell);

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
                    foreach (var prop in properties)
                    {
                        var itemValue = prop.GetValue(item);
                        var dataCell = new Cell
                        {
                            CellValue = new CellValue(itemValue?.ToString()),
                            DataType = CellValues.String
                        };
                        dataRow.Append(dataCell);
                    }

                    data.Append(dataRow);
                    currentRow++;
                }

                //Auto-size
                Columns columns = AutoSize(data);

                //Assemble the full document now with our properly sized/formatted sheet
                

                //Add a worksheet to it
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();
                worksheetPart.Worksheet.Append(columns);
                worksheetPart.Worksheet.Append(data);

                //Add the sheet to the workbook
                var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                var sheet = new Sheet
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = exportConfiguration.WorksheetName
                };
                sheets.Append(sheet);

                workbookPart.Workbook.Save();
                spreadsheetDocument.Close();

                //Return the bytearray
                documentStream.Seek(0, SeekOrigin.Begin);
                return documentStream.ToArray();
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

        private Columns AutoSize(SheetData sheetData)
        {
            var maxColWidth = GetMaxCharacterWidth(sheetData);

            Columns columns = new Columns();
            //this is the width of my font - yours may be different
            double maxWidth = 7;
            foreach (var item in maxColWidth)
            {
                //width = Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
                double width = Math.Truncate((item.Value * maxWidth + 5) / maxWidth * 256) / 256;

                //pixels=Truncate(((256 * {width} + Truncate(128/{Maximum Digit Width}))/256)*{Maximum Digit Width})
                double pixels = Math.Truncate(((256 * width + Math.Truncate(128 / maxWidth)) / 256) * maxWidth);

                //character width=Truncate(({pixels}-5)/{Maximum Digit Width} * 100+0.5)/100
                double charWidth = Math.Truncate((pixels - 5) / maxWidth * 100 + 0.5) / 100;

                Column col = new Column() { BestFit = true, Min = (UInt32)(item.Key + 1), Max = (UInt32)(item.Key + 1), CustomWidth = true, Width = (DoubleValue)width };
                columns.Append(col);
            }

            return columns;
        }


        private Dictionary<int, int> GetMaxCharacterWidth(SheetData sheetData)
        {
            //iterate over all cells getting a max char value for each column
            Dictionary<int, int> maxColWidth = new Dictionary<int, int>();
            var rows = sheetData.Elements<Row>();
            UInt32[] numberStyles = new UInt32[] { 5, 6, 7, 8 }; //styles that will add extra chars
            UInt32[] boldStyles = new UInt32[] { 1, 2, 3, 4, 6, 7, 8 }; //styles that will bold
            foreach (var r in rows)
            {
                var cells = r.Elements<Cell>().ToArray();

                //using cell index as my column
                for (int i = 0; i < cells.Length; i++)
                {
                    var cell = cells[i];
                    var cellValue = cell.CellValue == null ? string.Empty : cell.CellValue.InnerText;
                    var cellTextLength = cellValue.Length;

                    if (cell.StyleIndex != null && numberStyles.Contains(cell.StyleIndex))
                    {
                        int thousandCount = (int)Math.Truncate((double)cellTextLength / 4);

                        //add 3 for '.00' 
                        cellTextLength += (3 + thousandCount);
                    }

                    if (cell.StyleIndex != null && boldStyles.Contains(cell.StyleIndex))
                    {
                        //add an extra char for bold - not 100% acurate but good enough for what i need.
                        cellTextLength += 1;
                    }

                    if (maxColWidth.ContainsKey(i))
                    {
                        var current = maxColWidth[i];
                        if (cellTextLength > current)
                        {
                            maxColWidth[i] = cellTextLength;
                        }
                    }
                    else
                    {
                        maxColWidth.Add(i, cellTextLength);
                    }
                }
            }

            return maxColWidth;
        }
    }
}
