using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FontSize = DocumentFormat.OpenXml.Spreadsheet.FontSize;

namespace ICG.NetCore.Utilities.Spreadsheet
{
    /// <summary>
    /// A concrete implementation of <see cref="ISpreadsheetGenerator"/> using the OpenXML project
    /// </summary>
    public class OpenXmlSpreadsheetGenerator : ISpreadsheetGenerator
    {
        /// <inheritdoc />
        public byte[] CreateMultiSheetSpreadsheet(IEnumerable<ISpreadsheetConfiguration<object>> exportSheets)
        {
            //Validate inputs
            if (exportSheets == null)
                throw new ArgumentNullException(nameof(exportSheets));

            //Setup a memory stream to hold the generated file
            using (var documentStream = new MemoryStream())
            {
                //Create the document & overall workbook
                var spreadsheetDocument = SpreadsheetDocument.Create(documentStream, SpreadsheetDocumentType.Workbook);
                var workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                //Setup our styles
                var stylesPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = CreateStylesheet();
                stylesPart.Stylesheet.Save();

                //Loop through all of the sheets
                var sheetId = 1u;
                foreach (var item in exportSheets)
                {
                    var typeDetail = item.DataType;
                    var data = CreateExportSheet(item, out var columns);

                    //Add a worksheet to our document
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet();
                    worksheetPart.Worksheet.Append(columns);
                    worksheetPart.Worksheet.Append(data);

                    //Add the sheet to the workbook
                    var sheet = new Sheet
                    {
                        Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                        SheetId = sheetId,
                        Name = item.WorksheetName
                    };
                    sheets.Append(sheet);
                    sheetId++;
                }
                //Save off the file
                workbookPart.Workbook.Save();
                spreadsheetDocument.Close();

                //Return the bytearray
                documentStream.Seek(0, SeekOrigin.Begin);
                return documentStream.ToArray();
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
            DataHeader = 3,
            /// <summary>
            /// Normal font formatted for currency
            /// </summary>
            NormalCurrency = 4,
            /// <summary>
            /// Normal font formated for date
            /// </summary>
            NormalDate = 5
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
                stylesPart.Stylesheet = CreateStylesheet();
                stylesPart.Stylesheet.Save();

                var data = CreateExportSheet(exportConfiguration, out var columns);

                //Add a worksheet to our document
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

        private static Cell CellFromValue(Type t, object itemValue) => t switch
        {
            _ when t == typeof(int) => new Cell { CellValue = new CellValue((int)itemValue), DataType = CellValues.Number },
            _ when t == typeof(decimal) => new Cell { CellValue = new CellValue((decimal)itemValue), DataType = CellValues.Number },
            _ when t == typeof(double) => new Cell { CellValue = new CellValue((double)itemValue), DataType = CellValues.Number },
            _ when t == typeof(long) => new Cell { CellValue = new CellValue((decimal)itemValue), DataType = CellValues.Number }, //There is no constructor for longs
            _ when t == typeof(float) => new Cell { CellValue = new CellValue((float)itemValue), DataType = CellValues.Number },
            _ when t == typeof(DateTime) => new Cell { CellValue = new CellValue((DateTime)itemValue), DataType = CellValues.Date },
            _ when t == typeof(DateTimeOffset) => new Cell { CellValue = new CellValue((DateTimeOffset)itemValue), DataType = CellValues.Date },
            _ => new Cell { CellValue = new CellValue(itemValue.ToString()), DataType = CellValues.String },
        };

        private record OutputPropMap(PropDetail PropDetail, Column Column, List<Cell> Cells);

        private SheetData CreateExportSheet(ISpreadsheetConfiguration<object> exportConfiguration, out Columns columns)
        {
            //Build out our sheet information
            var data = new SheetData();
            UInt32 currentRow = 1;

            columns = new Columns();

            var outputMap = new Dictionary<PropDetail, OutputPropMap>();
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
            var headerProperties = TypeDiscoverer.GetProps(exportConfiguration.DataType);
            var headerRow = new Row { RowIndex = currentRow };
            foreach (var prop in headerProperties)
            {
                var column = new Column()
                {
                    Min = (uint)prop.Order,
                    Max = (uint)prop.Order,
                    BestFit = true,
                    Width = prop.Width > 0 ? prop.Width : 10
                };
                columns.Append(column);
                outputMap[prop] = new OutputPropMap(prop, column, new List<Cell>());

                var headerCell = new Cell
                {
                    CellValue = new CellValue(prop.DisplayName),
                    DataType = CellValues.String,
                    StyleIndex = (int)FontStyleIndex.DataHeader
                };
                headerRow.Append(headerCell);
            }

            data.Append(headerRow);
            currentRow++;

            //Run the data
            foreach (var item in exportConfiguration.ExportData)
            {
                var dataRow = new Row { RowIndex = currentRow };
                foreach (var prop in headerProperties)
                {
                    var itemValue = prop.Descriptor.GetValue(item);
                    var dataCell = CellFromValue(prop.Descriptor.PropertyType, itemValue);

                    if (prop.Format == "c")
                    {
                        dataCell.StyleIndex = (int)FontStyleIndex.NormalCurrency;
                    }
                    else if (prop.Format == "d") //Date
                    {
                        dataCell.StyleIndex = (int)FontStyleIndex.NormalDate;
                    }
                    outputMap[prop].Cells.Add(dataCell);
                    dataRow.Append(dataCell);
                }

                data.Append(dataRow);
                currentRow++;
            }

            //Auto-size
            //columns = AutoSize(data);

            if (exportConfiguration.AutoSizeColumns)
            {
                CalculateSizes(outputMap.Values.ToList());
            }
            return data;
        }


        /// <summary>
        /// Creates the needed stylesheet to support our styles for documents
        /// </summary>
        /// <returns></returns>
        private Stylesheet CreateStylesheet()
        {
            var styles = new Stylesheet();
            // blank font list
            styles.Fonts = new Fonts { Count = 4 };
            styles.Fonts.AppendChild(new Font()
            {
                FontSize = new FontSize { Val = 11 }
            });
            styles.Fonts.AppendChild(new Font()
            {
                Bold = new Bold(),
                FontSize = new FontSize() { Val = 14 }
            });
            styles.Fonts.AppendChild(new Font()
            {
                Bold = new Bold(),
                FontSize = new FontSize() { Val = 12 }
            });
            styles.Fonts.AppendChild(new Font()
            {
                Bold = new Bold()
            });

            // create fills
            styles.Fills = new Fills();

            // create a solid red fill
            var solidRed = new PatternFill
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FFFF0000") },
                BackgroundColor = new BackgroundColor { Indexed = 64 }
            };
            // red fill

            styles.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); // required, reserved by Excel
            styles.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); // required, reserved by Excel
            styles.Fills.AppendChild(new Fill { PatternFill = solidRed });
            styles.Fills.Count = 3;

            // blank border list
            styles.Borders = new Borders();
            styles.Borders.Count = 1;
            styles.Borders.AppendChild(new Border());

            // blank cell format list
            styles.CellStyleFormats = new CellStyleFormats();
            styles.CellStyleFormats.Count = 1;
            styles.CellStyleFormats.AppendChild(new CellFormat());

            styles.NumberingFormats = new NumberingFormats();
            styles.NumberingFormats.AppendChild(new NumberingFormat { NumberFormatId = 164, FormatCode = "\"$\"#,##0.00" });
            styles.NumberingFormats.AppendChild(new NumberingFormat { NumberFormatId = 300, FormatCode = "MM/dd/yyyy" });

            // cell format list
            styles.CellFormats = new CellFormats();

            styles.CellFormats.AppendChild(new CellFormat()); //Default

            //Header
            styles.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 1, BorderId = 0, FillId = 0 });//.AppendChild(new Alignment { Horizontal = HorizontalAlignmentValues.Center });

            //Sub-header
            styles.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 2, BorderId = 0, FillId = 0 });

            //Data-header
            styles.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 3, BorderId = 0, FillId = 0 });

            //normal-currency
            styles.CellFormats.AppendChild(new CellFormat
            {
                FormatId = 0,
                FontId = 0,
                BorderId = 0,
                FillId = 0,
                NumberFormatId = 164,
                ApplyNumberFormat = true
            });
            styles.CellFormats.AppendChild(new CellFormat
            {
                FormatId = 0,
                FontId = 0,
                BorderId = 0,
                FillId = 0,
                NumberFormatId = 300,
                ApplyNumberFormat = true
            });

            styles.CellFormats.Count = 6;
            return styles;
        }


        private static void CalculateSizes(IList<OutputPropMap> propMap)
        {
            //Adapted from - https://stackoverflow.com/questions/18268620/openxml-auto-size-column-width-in-excel
            //This is an approximation of the size needed for the largest single character in Calibri 
            double maxWidth = 7;
            foreach (var (_, col , cells) in propMap)
            {
                var rawWidth = GetMaxCharacterWidth(cells);
                //width = Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
                var width = Math.Truncate((rawWidth * maxWidth + 5) / maxWidth * 256) / 256;
                col.CustomWidth = true;
                col.Width = width;
            }
        }

        private static int GetMaxCharacterWidth(IList<Cell> cells)
        {
            //iterate over all cells getting a max char value for each column
            var maxWidth = 0;

            //TODO: Be smarter about this for our set styles
            var numberStyles = new UInt32[] { 5, 6, 7, 8 }; //styles that will add extra chars
            var boldStyles = new UInt32[] { 1, 2, 3, 4, 6, 7, 8 }; //styles that will bold

            //using cell index as my column
            foreach (var cell in cells)
            {
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

                if (cellTextLength > maxWidth)
                {
                    maxWidth = cellTextLength;
                }
            }

            return maxWidth;
        }
    }
}
